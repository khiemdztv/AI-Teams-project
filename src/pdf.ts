/**
 * Fork addition: a PDF from text and images, without giving Claude a shell.
 *
 * Claude passes structured content (title, text, image paths), never markup.
 * This module writes the HTML itself: text is escaped, and images are read
 * here, only from the allowed folders (the inbox and the export folder), and
 * embedded as data URIs. Microsoft Edge prints the page headless, with its own
 * empty profile and every host name failing to resolve, under a
 * Content-Security-Policy that allows nothing but inline styles and data:
 * images (no script, no fetch). So no text Claude passes can make the browser
 * fetch a URL or read a local file into the PDF. (Turning scripts off with
 * --blink-settings stops Edge from printing at all, hence the CSP alone.)
 */

import { existsSync, mkdtempSync, readFileSync, rmSync, statSync, writeFileSync, copyFileSync } from 'fs'
import { tmpdir } from 'os'
import { basename, join } from 'path'
import { pathToFileURL } from 'url'
import { assertWithinAnyRoot } from './attach.js'
import { sniffImageExt } from './attachments.js'

export const MAX_PDF_IMAGES = 20
export const MAX_PDF_IMAGE_BYTES = 20 * 1024 * 1024
export const MAX_PDF_TEXT_CHARS = 200_000
const PRINT_TIMEOUT_MS = 60_000

const IMAGE_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
  bmp: 'image/bmp',
}

export type PdfContent = {
  title?: string
  /** Blank lines separate paragraphs; "# " and "## " start headings; "- " starts a bullet; **bold**. */
  text?: string
  /** Absolute paths, each inside one of the allowed folders. */
  images?: string[]
}

export function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;')
}

/** Escape, then the one inline mark: **bold**. */
function inline(s: string): string {
  return escapeHtml(s).replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
}

export function textToHtml(text: string): string {
  const out: string[] = []
  for (const block of text.replace(/\r\n?/g, '\n').split(/\n\s*\n/)) {
    const lines = block.split('\n').filter(l => l.trim())
    if (!lines.length) continue
    let para: string[] = []
    let bullets: string[] = []
    const flush = () => {
      if (para.length) out.push(`<p>${para.join('<br>')}</p>`)
      if (bullets.length) out.push(`<ul>${bullets.map(b => `<li>${b}</li>`).join('')}</ul>`)
      para = []
      bullets = []
    }
    for (const line of lines) {
      const t = line.trim()
      const heading = /^(#{1,2})\s+(.*)$/.exec(t)
      if (heading) {
        flush()
        out.push(heading[1].length === 1 ? `<h2>${inline(heading[2])}</h2>` : `<h3>${inline(heading[2])}</h3>`)
      } else if (/^[-*•]\s+/.test(t)) {
        if (para.length) flush()
        bullets.push(inline(t.replace(/^[-*•]\s+/, '')))
      } else {
        if (bullets.length) flush()
        para.push(inline(line))
      }
    }
    flush()
  }
  return out.join('\n')
}

/** An image from an allowed folder as a data: URI; also used by src/video.ts. */
export function embedImage(path: string, imageRoots: string[]): string {
  assertWithinAnyRoot(path, imageRoots)
  if (statSync(path).size > MAX_PDF_IMAGE_BYTES) throw new Error(`${basename(path)} is larger than 20MB`)
  const bytes = readFileSync(path)
  const mime = IMAGE_MIME[sniffImageExt(bytes) ?? '']
  if (!mime) throw new Error(`${basename(path)} is not a PNG, JPEG, GIF, WebP or BMP image`)
  return `data:${mime};base64,${bytes.toString('base64')}`
}

/** The whole HTML page for `content`. Images are read and checked here. */
export function buildPdfHtml(content: PdfContent, imageRoots: string[]): string {
  const text = content.text ?? ''
  if (text.length > MAX_PDF_TEXT_CHARS) throw new Error(`text is longer than ${MAX_PDF_TEXT_CHARS} characters`)
  const images = content.images ?? []
  if (images.length > MAX_PDF_IMAGES) throw new Error(`at most ${MAX_PDF_IMAGES} images`)
  if (!content.title && !text.trim() && !images.length) throw new Error('nothing to put in the PDF')

  const figures = images.map(path => `<figure><img src="${embedImage(path, imageRoots)}"></figure>`)

  const title = content.title ? escapeHtml(content.title) : ''
  return `<!doctype html>
<html><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'">
<title>${title}</title>
<style>
@page { size: A4; margin: 2cm }
body { font-family: Calibri, Arial, sans-serif; font-size: 12pt; line-height: 1.4; color: #000 }
h1 { font-size: 18pt; margin: 0 0 12pt }
h2 { font-size: 14pt; margin: 14pt 0 6pt }
h3 { font-size: 12pt; margin: 12pt 0 6pt }
p { margin: 0 0 8pt }
figure { margin: 10pt 0; break-inside: avoid }
img { max-width: 100%; max-height: 24cm }
</style></head><body>
${title ? `<h1>${title}</h1>\n` : ''}${textToHtml(text)}
${figures.join('\n')}
</body></html>
`
}

/** A file name for the export folder: a bare name ending in .pdf. */
export function pdfFileName(raw: string): string {
  // Split on both separators: path.basename on Linux keeps `..\..\x.pdf` whole.
  let name = (String(raw ?? '').trim().split(/[\\/]/).pop() ?? '').replace(/[<>:"/\\|?*\x00-\x1f]/g, '_')
  if (!name || name === '.pdf' || /^\.+$/.test(name)) throw new Error('file_name is required')
  if (!/\.pdf$/i.test(name)) name += '.pdf'
  if (name.length > 120) throw new Error('file_name is too long')
  return name
}

/**
 * MSTEAMS_EDGE_PATH, else Edge where Windows installs it, else a Chromium-family
 * browser from a Linux package (all take the same headless flags). Snap
 * Chromium is not listed: its sandbox cannot read the page in the temp folder.
 */
export function findEdge(env: NodeJS.ProcessEnv = process.env): string | undefined {
  const candidates = [
    env.MSTEAMS_EDGE_PATH,
    env['ProgramFiles(x86)'] && join(env['ProgramFiles(x86)'], 'Microsoft', 'Edge', 'Application', 'msedge.exe'),
    env.ProgramFiles && join(env.ProgramFiles, 'Microsoft', 'Edge', 'Application', 'msedge.exe'),
    '/usr/bin/microsoft-edge',
    '/usr/bin/google-chrome',
    '/usr/bin/chromium',
  ]
  return candidates.find((p): p is string => !!p && existsSync(p))
}

/** Print `html` to `outPath` with headless Edge. Edge only ever writes inside a temp folder. */
export async function printHtmlToPdf(html: string, outPath: string, edge = findEdge()): Promise<void> {
  if (!edge) throw new Error('making a PDF needs Microsoft Edge or Google Chrome, which was not found')
  const work = mkdtempSync(join(tmpdir(), 'msteams-pdf-'))
  try {
    const page = join(work, 'page.html')
    const printed = join(work, 'out.pdf')
    writeFileSync(page, html, { mode: 0o600 })
    const proc = Bun.spawn(
      [
        edge,
        '--headless=new',
        '--disable-gpu',
        '--no-first-run',
        '--no-default-browser-check',
        '--disable-extensions',
        '--disable-sync',
        `--user-data-dir=${join(work, 'profile')}`,
        '--host-resolver-rules=MAP * ~NOTFOUND',
        '--no-pdf-header-footer',
        `--print-to-pdf=${printed}`,
        pathToFileURL(page).href,
      ],
      { stdin: 'ignore', stdout: 'ignore', stderr: 'ignore' },
    )
    const timer = setTimeout(() => proc.kill(), PRINT_TIMEOUT_MS)
    await proc.exited
    clearTimeout(timer)
    if (!existsSync(printed) || statSync(printed).size === 0) throw new Error('Edge did not produce a PDF')
    copyFileSync(printed, outPath)
  } finally {
    try {
      rmSync(work, { recursive: true, force: true })
    } catch {
      // Edge may still hold its profile for a moment; the temp folder is harmless.
    }
  }
}
