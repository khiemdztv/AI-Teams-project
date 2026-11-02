/**
 * Plain text out of .docx / .xlsx / .pptx (fork addition).
 *
 * Claude's Read handles text, images and PDF but not Office files, and the bot
 * machine has no Python or converter to fall back on. So after an attachment
 * is downloaded, its text is written beside it as `<file>.txt` for Claude to
 * Read. Office files are zip archives of XML; this reads just enough of both,
 * with no dependency.
 *
 * Text only: formatting, images and formulas are dropped (a cell shows its
 * cached value). Sizes are capped so a crafted archive cannot exhaust memory.
 */

import { inflateRawSync } from 'zlib'
import { extname } from 'path'

/** Cap on any one decompressed archive member. */
const MAX_MEMBER_BYTES = 50 * 1024 * 1024
/** Cap on the text handed to Claude. */
export const MAX_TEXT_CHARS = 2_000_000

export const OFFICE_EXTENSIONS = new Set(['docx', 'xlsx', 'pptx'])

export function isOfficeFile(path: string): boolean {
  return OFFICE_EXTENSIONS.has(extname(path).slice(1).toLowerCase())
}

/** Members of a zip archive, decompressed on demand. Zip64 is not supported. */
export function readZip(buf: Buffer): Map<string, () => Buffer> {
  const entries = new Map<string, () => Buffer>()
  // End of central directory: 22 bytes plus a comment of up to 64KB, at the end.
  let eocd = -1
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 22 - 0xffff); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) {
      eocd = i
      break
    }
  }
  if (eocd < 0) throw new Error('not a zip archive')

  const count = buf.readUInt16LE(eocd + 10)
  let p = buf.readUInt32LE(eocd + 16)
  for (let n = 0; n < count; n++) {
    if (p + 46 > buf.length || buf.readUInt32LE(p) !== 0x02014b50) throw new Error('corrupt zip directory')
    const method = buf.readUInt16LE(p + 10)
    const compressed = buf.readUInt32LE(p + 20)
    const size = buf.readUInt32LE(p + 24)
    const nameLen = buf.readUInt16LE(p + 28)
    const extraLen = buf.readUInt16LE(p + 30)
    const commentLen = buf.readUInt16LE(p + 32)
    const local = buf.readUInt32LE(p + 42)
    const name = buf.subarray(p + 46, p + 46 + nameLen).toString('utf8')
    p += 46 + nameLen + extraLen + commentLen

    entries.set(name, () => {
      if (size > MAX_MEMBER_BYTES) throw new Error(`${name} is too large to extract`)
      if (buf.readUInt32LE(local) !== 0x04034b50) throw new Error('corrupt zip entry')
      const start = local + 30 + buf.readUInt16LE(local + 26) + buf.readUInt16LE(local + 28)
      const data = buf.subarray(start, start + compressed)
      if (method === 0) return Buffer.from(data)
      if (method === 8) return inflateRawSync(data, { maxOutputLength: MAX_MEMBER_BYTES })
      throw new Error(`${name} uses an unsupported compression method`)
    })
  }
  return entries
}

export function decodeXml(text: string): string {
  return text.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (_, e: string) => {
    const k = e.toLowerCase()
    if (k === 'amp') return '&'
    if (k === 'lt') return '<'
    if (k === 'gt') return '>'
    if (k === 'quot') return '"'
    if (k === 'apos') return "'"
    const code = k.startsWith('#x') ? parseInt(k.slice(2), 16) : parseInt(k.slice(1), 10)
    return Number.isFinite(code) ? String.fromCodePoint(code) : ''
  })
}

/** Text of the `<prefix:t>` runs in a fragment, in order. */
function runs(xml: string, prefix: string): string {
  const re = new RegExp(`<${prefix}:t(?:\\s[^>]*)?>([^<]*)</${prefix}:t>|<${prefix}:tab/>|<${prefix}:br/>`, 'g')
  let out = ''
  for (const m of xml.matchAll(re)) {
    if (m[1] !== undefined) out += decodeXml(m[1])
    else out += m[0].includes('tab') ? '\t' : '\n'
  }
  return out
}

function text(member: (() => Buffer) | undefined): string {
  return member ? member().toString('utf8') : ''
}

export function docxToText(zip: Map<string, () => Buffer>): string {
  const xml = text(zip.get('word/document.xml'))
  if (!xml) throw new Error('no word/document.xml — not a Word document')
  const lines: string[] = []
  // Table cells become tab-separated, table rows and paragraphs become lines.
  for (const row of xml.split(/<\/w:tr>/)) {
    const cells = row.split(/<\/w:tc>/)
    const rendered = cells.map(cell =>
      cell
        .split(/<\/w:p>/)
        .map(p => runs(p, 'w'))
        .filter(s => s.trim())
        .join(cells.length > 1 ? ' ' : '\n'),
    )
    const joined = cells.length > 1 ? rendered.filter(Boolean).join('\t') : rendered.join('')
    if (joined.trim()) lines.push(joined)
  }
  return lines.join('\n')
}

/** "AB12" -> 27 (zero-based column index). */
function columnIndex(ref: string): number {
  let n = 0
  for (const ch of ref.replace(/\d+$/, '').toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}

export function xlsxToText(zip: Map<string, () => Buffer>): string {
  const shared: string[] = []
  for (const si of text(zip.get('xl/sharedStrings.xml')).split(/<\/si>/)) {
    if (si.includes('<si')) shared.push(plainRuns(si.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '')))
  }

  const workbook = text(zip.get('xl/workbook.xml'))
  if (!workbook) throw new Error('no xl/workbook.xml — not an Excel workbook')
  const rels = new Map<string, string>()
  for (const m of text(zip.get('xl/_rels/workbook.xml.rels')).matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /\bId="([^"]+)"/.exec(m[0])?.[1]
    const target = /\bTarget="([^"]+)"/.exec(m[0])?.[1]
    if (id && target) rels.set(id, target.replace(/^\/?(xl\/)?/, 'xl/'))
  }

  const out: string[] = []
  for (const m of workbook.matchAll(/<sheet\b[^>]*>/g)) {
    const name = decodeXml(/\bname="([^"]*)"/.exec(m[0])?.[1] ?? 'Sheet')
    const rid = /\br:id="([^"]+)"/.exec(m[0])?.[1]
    const sheet = text(zip.get(rels.get(rid ?? '') ?? ''))
    out.push(`## ${name}`)
    for (const row of sheet.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)) {
      const cells: string[] = []
      for (const c of row[1]!.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const attrs = c[1] ?? ''
        const body = c[2] ?? ''
        const ref = /\br="([A-Z]+\d+)"/.exec(attrs)?.[1]
        const type = /\bt="([^"]+)"/.exec(attrs)?.[1]
        const v = /<v>([^<]*)<\/v>/.exec(body)?.[1]
        let value = ''
        if (type === 's' && v !== undefined) value = shared[Number(v)] ?? ''
        else if (type === 'inlineStr') value = plainRuns(body)
        else if (v !== undefined) value = decodeXml(v)
        const at = ref ? columnIndex(ref) : cells.length
        while (cells.length < at) cells.push('')
        cells[at] = value.replace(/[\t\r\n]+/g, ' ')
      }
      if (cells.some(Boolean)) out.push(cells.join('\t'))
    }
    out.push('')
  }
  return out.join('\n').trim()
}

/** `<t>` runs with no namespace prefix (sharedStrings, inline strings). */
function plainRuns(xml: string): string {
  let out = ''
  for (const m of xml.matchAll(/<t(?:\s[^>]*)?>([^<]*)<\/t>/g)) out += decodeXml(m[1]!)
  return out
}

export function pptxToText(zip: Map<string, () => Buffer>): string {
  const slides = [...zip.keys()]
    .map(name => ({ name, n: Number(/^ppt\/slides\/slide(\d+)\.xml$/.exec(name)?.[1]) }))
    .filter(s => Number.isFinite(s.n))
    .sort((a, b) => a.n - b.n)
  if (!slides.length) throw new Error('no slides — not a PowerPoint file')
  return slides
    .map(({ name, n }) => {
      const body = text(zip.get(name))
        .split(/<\/a:p>/)
        .map(p => runs(p, 'a'))
        .filter(s => s.trim())
        .join('\n')
      return `## Slide ${n}\n${body}`
    })
    .join('\n\n')
}

/** Text of an Office file; throws with a short reason if it cannot be read. */
export function officeToText(bytes: Buffer, ext: string): string {
  const zip = readZip(bytes)
  const kind = ext.toLowerCase()
  const raw = kind === 'docx' ? docxToText(zip) : kind === 'xlsx' ? xlsxToText(zip) : pptxToText(zip)
  return raw.length > MAX_TEXT_CHARS ? `${raw.slice(0, MAX_TEXT_CHARS)}\n\n[… đã cắt, file quá dài]` : raw
}
