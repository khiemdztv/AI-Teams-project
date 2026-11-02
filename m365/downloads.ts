/**
 * Files the m365 tools fetch for the sender (email attachments, SharePoint and
 * OneDrive files): where they are saved, what Claude is told to do with them,
 * and cleanup.
 *
 * They are saved inside the export folder, because the Teams `reply` tool
 * sends only from there; one folder per source (an email, a drive item), so
 * two sources' "scan.pdf" do not overwrite each other.
 */

import { mkdirSync, readdirSync, rmdirSync, statSync, writeFileSync } from 'fs'
import { createHash } from 'crypto'
import { basename, extname, join } from 'path'
import { purgeOlderThan } from '../src/purge.js'
import { sniffImageExt } from '../src/attachments.js'
import { isOfficeFile, officeToText } from '../src/office-text.js'
import { makeImagePreview } from '../src/image-preview.js'

/** Largest file the tools save. */
export const MAX_DOWNLOAD_BYTES = 50 * 1024 * 1024
/** Images over this get a smaller copy to look at; Claude's image input stops at 5MB. */
export const MAX_VIEW_IMAGE_BYTES = Math.floor(3.5 * 1024 * 1024)

/** "8.9MB", "557KB". */
export function formatSize(bytes: number | undefined): string {
  const n = bytes ?? 0
  return n >= 1024 * 1024 ? `${(n / 1024 / 1024).toFixed(1)}MB` : `${Math.ceil(n / 1024)}KB`
}

/** A name from elsewhere as a file name: no folders, nothing Windows refuses. */
export function safeFileName(raw: string): string {
  // Not path.basename: on Windows it reads "a:b.pdf" as drive a:.
  const name = (raw.split(/[\\/]/).pop() ?? '')
    .replace(/[<>:"/\\|?*\x00-\x1f]/g, '_')
    .replace(/[. ]+$/, '')
    .slice(0, 150)
  return name && !/^\.+$/.test(name) ? name : 'attachment'
}

/**
 * Write `bytes` as `<root>\<hash of sourceKey>\<name>` and say how to use it:
 * Read for images and PDFs (a smaller copy for a large image), the .txt for
 * Word, Excel and PowerPoint, and reply's files to send it in Teams.
 */
export async function saveForUse(root: string, sourceKey: string, name: string, bytes: Buffer, notice: string): Promise<string> {
  const dir = join(root, createHash('sha256').update(sourceKey).digest('hex').slice(0, 10))
  mkdirSync(dir, { recursive: true, mode: 0o700 })
  const file = join(dir, safeFileName(name))
  writeFileSync(file, bytes, { mode: 0o600 })

  const lines = [`saved: ${file}`, `size: ${formatSize(bytes.length)}`, notice]
  const ext = extname(file).slice(1).toLowerCase()
  if (isOfficeFile(file)) {
    try {
      writeFileSync(`${file}.txt`, officeToText(bytes, ext), { mode: 0o600 })
      lines.push(`Read its text at: ${file}.txt`)
    } catch (err) {
      lines.push(`Its text could not be extracted (${err instanceof Error ? err.message : err}).`)
    }
  } else if (sniffImageExt(bytes)) {
    if (bytes.length > MAX_VIEW_IMAGE_BYTES) {
      try {
        const preview = await makeImagePreview(file, join(dir, `${basename(file, extname(file))}.preview.jpg`))
        lines.push(`Too large to view directly; look at the smaller copy with Read: ${preview}`)
      } catch (err) {
        lines.push(`Too large to view directly, and no smaller copy could be made (${err instanceof Error ? err.message : err}).`)
      }
    } else {
      lines.push(`Look at it with Read: ${file}`)
    }
  } else {
    lines.push(`Open it with Read: ${file}${ext === 'pdf' ? ' (for a PDF over 10 pages, pass pages)' : ''}`)
  }
  lines.push(`To send it in Teams, pass ${file} in reply's files.`)
  return lines.join('\n')
}

/** Empty each source folder of files older than `maxAgeMs`; remove it once empty. */
export function purgeDownloads(root: string, maxAgeMs: number, now = Date.now()): number {
  let names: string[]
  try {
    names = readdirSync(root)
  } catch {
    return 0
  }
  let removed = 0
  for (const name of names) {
    const dir = join(root, name)
    try {
      if (!statSync(dir).isDirectory()) continue
      removed += purgeOlderThan(dir, maxAgeMs, now)
      if (readdirSync(dir).length === 0) rmdirSync(dir)
    } catch {
      // Locked or gone: next start tries again.
    }
  }
  return removed
}
