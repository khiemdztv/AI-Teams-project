/**
 * Outbound file payloads.
 *
 * Teams has no single "send a file" call. Which route applies depends on the
 * conversation type, the content type and the size (OpenClaw MIT 32b2e161a5a,
 * `extensions/msteams/src/file-consent-helpers.ts:106`):
 *
 *   personal + (large || not an image)  -> FileConsentCard round trip
 *   channel/group + not an image        -> SharePoint upload via Graph
 *   anything else, i.e. a small image   -> inline `data:` URI, one call
 *
 * All three are implemented. Choosing between them is `src/outbound.ts`; this
 * module owns two narrower jobs: turning a path into bytes we are allowed to
 * send, and building the inline payload.
 *
 * Not `src/attachments.ts`: that one handles INBOUND attachments (classifying
 * and fetching what senders sent). This one only builds outbound payloads.
 */

import { closeSync, existsSync, fstatSync, openSync, readFileSync, realpathSync } from 'fs'
import { extname, basename, join, sep } from 'path'

/**
 * Teams switches to the consent flow at 4MB. We reject at the same boundary
 * rather than sending a data URI Teams would refuse.
 */
export const MAX_INLINE_IMAGE_BYTES = 4 * 1024 * 1024

/**
 * Each attachment reaches ~5.3MB in memory after base64 inflation, and every
 * one is built before any send. Without a count limit an arbitrarily long
 * `files` array holds that multiplied in memory at once — this bounds the
 * array the same way the byte cap bounds a single file.
 */
export const MAX_ATTACHMENTS = 10

/** Teams only accepts base64 data URIs for images, so this list is the gate. */
const IMAGE_MIME: Record<string, string> = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.bmp': 'image/bmp',
}

export type Attachment = {
  name: string
  contentType: string
  contentUrl: string
}

/** The image content type for this name, or undefined if it is not an image. */
export function imageMimeFor(name: string): string | undefined {
  return IMAGE_MIME[extname(name).toLowerCase()]
}

/**
 * The extension Teams uses to pick a file card's icon, lowercased and without
 * the dot. Empty for a name with no extension, which Teams tolerates.
 */
export function fileTypeOf(name: string): string {
  const ext = extname(name)
  return ext ? ext.slice(1).toLowerCase() : ''
}

/** Strip anything that could confuse Teams or a downstream client. */
export function sanitizeAttachmentName(name: string): string {
  return (
    basename(name)
      .replace(/[\x00-\x1f\x7f]/g, '')
      .replace(/[\\/]/g, '_')
      .slice(0, 200) || 'file'
  )
}

/**
 * Refuse to send the channel's own state. Adapted from the telegram plugin:
 * `.env` holds the Entra credentials and `conversations/` holds proactive-send
 * targets, so those are the one set of paths Claude has no legitimate reason to
 * hand to a chat. `inbox/` is exempt — that is where inbound downloads land, so
 * sending one back is normal.
 *
 * `pending-uploads/` is covered by the same rule, which is worth stating: a
 * snapshot of a file already offered to someone else can never be re-sent.
 */
export function assertSendable(path: string, stateDir: string): void {
  let real: string
  let stateReal: string
  try {
    real = realpathSync(path)
    stateReal = realpathSync(stateDir)
  } catch {
    // Missing file, or no state dir yet: the read below reports it properly.
    return
  }
  const inbox = join(stateReal, 'inbox')
  if (real.startsWith(stateReal + sep) && !real.startsWith(inbox + sep)) {
    throw new Error(`refusing to send channel state: ${path}`)
  }
}

/**
 * Refuse any file outside the one export folder (`MSTEAMS_SENDABLE_ROOT`).
 *
 * Fork addition. Upstream blocked only the channel's own state, so Claude could
 * attach any readable file on the machine — another client's documents, SSH
 * keys, Claude Code's own config. Both sides are resolved with the native
 * realpath, which follows symlinks and Windows junctions, and compared without
 * case on Windows, where `C:\Export` and `c:\export` are the same folder.
 * Unlike assertSendable, a path that cannot be resolved is refused, not let
 * through.
 */
export function assertWithinRoot(
  path: string,
  root: string,
  platform: NodeJS.Platform = process.platform,
): void {
  let real: string
  let rootReal: string
  try {
    rootReal = realpathSync.native(root)
  } catch {
    throw new Error(`file sending is misconfigured: export folder ${root} does not exist`)
  }
  try {
    real = realpathSync.native(path)
  } catch {
    throw new Error(`cannot attach ${basename(path)}: file not found`)
  }
  const fold = (p: string) => (platform === 'win32' ? p.toLowerCase() : p)
  const base = rootReal.endsWith(sep) ? rootReal : rootReal + sep
  if (!fold(real).startsWith(fold(base))) {
    throw new Error(`refusing to send ${basename(path)}: only files inside ${root} can be sent`)
  }
}

/**
 * Fork addition: `assertWithinRoot` against several folders; passes if any
 * one holds the file. The refusal names the folders.
 */
export function assertWithinAnyRoot(path: string, roots: string[]): void {
  for (const root of roots) {
    try {
      assertWithinRoot(path, root)
      return
    } catch {
      // try the next folder
    }
  }
  if (!existsSync(path)) throw new Error(`${basename(path)}: file not found`)
  throw new Error(`refused: ${basename(path)} is not in an allowed folder (${roots.join('; ')})`)
}

/** Sizes in an error the sender reads: "4.2MB", "17KB", "3 bytes". */
function formatBytes(n: number): string {
  if (n >= 1024 * 1024) return `${(n / 1024 / 1024).toFixed(1)}MB`
  if (n >= 1024) return `${Math.round(n / 1024)}KB`
  return `${n} bytes`
}

export type FileSnapshot = {
  bytes: Buffer
  /** Already sanitized: safe to show a client and to put in a URL path. */
  filename: string
  size: number
}

/**
 * Read a file once, into the bytes every outbound route then works from.
 *
 * One open, one read: the size that gets checked is the size of what is
 * actually sent, and a file that changes mid-delivery cannot make the checked
 * bytes and the sent bytes differ. (The path is still resolved before the open,
 * so a swap in that window is not detected — the guarantee is about the bytes
 * being consistent with themselves, not about racing a local attacker who can
 * already write to the file being offered.)
 *
 * The size is checked from the descriptor BEFORE reading, so an enormous file
 * is refused rather than pulled into memory first.
 */
export function loadSendableFile(
  path: string,
  stateDir: string,
  maxBytes: number,
  sendableRoot?: string,
): FileSnapshot {
  if (sendableRoot !== undefined) assertWithinRoot(path, sendableRoot)
  assertSendable(path, stateDir)

  const fd = openSync(path, 'r')
  try {
    const stat = fstatSync(fd)
    if (!stat.isFile()) {
      throw new Error(`cannot attach ${basename(path)}: not a regular file`)
    }
    if (stat.size > maxBytes) {
      throw new Error(
        `cannot attach ${basename(path)}: ${formatBytes(stat.size)} exceeds the ` +
          `${formatBytes(maxBytes)} limit for a Teams file — send a smaller file, or a link to it.`,
      )
    }
    return { bytes: readFileSync(fd), filename: sanitizeAttachmentName(path), size: stat.size }
  } finally {
    closeSync(fd)
  }
}

/**
 * Build an inline attachment, or throw an error worded for the sender.
 *
 * Throwing rather than returning undefined is deliberate: a silently skipped
 * attachment reads to Claude as a successful send, and the sender is left
 * waiting for a file that never arrives.
 */
export function buildImageAttachment(file: FileSnapshot): Attachment {
  const contentType = imageMimeFor(file.filename)
  if (!contentType) {
    throw new Error(
      `cannot attach ${file.filename}: Teams only accepts images inline. Anything else is routed ` +
        `through the consent or SharePoint flow before this point, so reaching here is a bug.`,
    )
  }

  if (file.size >= MAX_INLINE_IMAGE_BYTES) {
    throw new Error(
      `cannot attach ${file.filename}: ${formatBytes(file.size)} exceeds the 4MB ` +
        `inline limit. Larger images are routed through the consent or SharePoint flow before ` +
        `this point, so reaching here is a bug.`,
    )
  }

  return {
    name: file.filename,
    contentType,
    contentUrl: `data:${contentType};base64,${file.bytes.toString('base64')}`,
  }
}
