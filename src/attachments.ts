/**
 * Inbound attachments: classification, and fetching one to the inbox.
 *
 * Two things about Teams make this less obvious than it looks.
 *
 * **Every message carries a `text/html` attachment** holding the rendered
 * message body — a bare "hey" with no file arrives with one. So attachment
 * *presence* never means a file was sent, and classification has to be by
 * content type. Verified against the captured corpus; exclusion list matches
 * OpenClaw's `isAdvertisedFileAttachment` (MIT 32b2e161a5a,
 * `extensions/msteams/src/attachments/shared.ts:301`).
 *
 * **A file's `content.downloadUrl` is a credential, not an identifier.** It
 * embeds a `tempauth=` bearer token granting read access to the sender's
 * OneDrive, valid about an hour. It must never reach a log, an error message,
 * the notification meta, or disk. That is why the handle Claude receives is an
 * opaque id and the URL stays in memory, in this module, keyed by that id.
 *
 * Not `src/attach.ts`: that one builds OUTBOUND attachments to send. This one
 * only classifies and fetches what arrived.
 */

import { writeFileSync } from 'fs'
import { join, extname } from 'path'
import { createHash } from 'crypto'

/** A genuine file sent by a user, as opposed to the rendered message body. */
export const TEAMS_FILE_DOWNLOAD_INFO = 'application/vnd.microsoft.teams.file.download.info'

/** Teams accepts much larger, but an inbound file lands in the operator's home dir. */
export const MAX_DOWNLOAD_BYTES = 20 * 1024 * 1024

/** How long a download handle stays usable. The token in the URL dies around here anyway. */
export const HANDLE_TTL_MS = 60 * 60 * 1000

const IMAGE_EXT = new Set(['jpg', 'jpeg', 'png', 'gif', 'webp', 'bmp'])

/**
 * Hosts a download may reach, as suffixes (fork addition). Teams file
 * attachments are served from the tenant's SharePoint / OneDrive for Business.
 * Override with MSTEAMS_DOWNLOAD_HOSTS (comma-separated, e.g.
 * `.contoso.sharepoint.com,.contoso-my.sharepoint.com`) to pin the tenant.
 */
export const DEFAULT_DOWNLOAD_HOSTS = ['.sharepoint.com']

/** At most this many redirects, each re-checked against the allowed hosts. */
const MAX_REDIRECTS = 3

export function downloadHostsFromEnv(value: string | undefined): string[] {
  const hosts = (value ?? '')
    .split(',')
    .map(h => h.trim().toLowerCase())
    .filter(Boolean)
  return hosts.length ? hosts : DEFAULT_DOWNLOAD_HOSTS
}

/** True when `url` is https and its host ends with one of the suffixes. */
export function isAllowedDownloadUrl(url: string, suffixes: string[]): boolean {
  let parsed: URL
  try {
    parsed = new URL(url)
  } catch {
    return false
  }
  if (parsed.protocol !== 'https:') return false
  const host = parsed.hostname.toLowerCase()
  // ".contoso.sharepoint.com" matches that host and any subdomain of it; an
  // entry without the leading dot matches that exact host only.
  return suffixes.some(s => {
    const suffix = s.toLowerCase()
    return suffix.startsWith('.') ? host === suffix.slice(1) || host.endsWith(suffix) : host === suffix
  })
}

export type InboundAttachment = {
  /** Opaque handle given to the model. Never the URL. */
  id: string
  name?: string
  mime?: string
  kind: 'image' | 'file'
  fileType?: string
}

/**
 * `bot`: an image pasted into the chat (fork addition). Teams hosts it on the
 * Bot Connector at the activity's serviceUrl, and it is fetched with the bot's
 * own token — which may only ever be sent to that one host.
 */
type Handle = InboundAttachment & { url: string; at: number; auth?: 'bot'; host?: string }

/** A pasted image's `contentType` is `image/*` or `image/png`-style. */
function isHostedImage(contentType: unknown): boolean {
  return typeof contentType === 'string' && contentType.toLowerCase().startsWith('image/')
}

function hostOf(url: string): string | undefined {
  try {
    const parsed = new URL(url)
    return parsed.protocol === 'https:' ? parsed.hostname.toLowerCase() : undefined
  } catch {
    return undefined
  }
}

/** File extension from the leading bytes, for images whose type Teams leaves as `image/*`. */
export function sniffImageExt(bytes: Uint8Array): string | undefined {
  const b = bytes
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return 'png'
  if (b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return 'jpg'
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return 'gif'
  if (b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 && b[8] === 0x57 && b[9] === 0x45) return 'webp'
  if (b[0] === 0x42 && b[1] === 0x4d) return 'bmp'
  return undefined
}

/**
 * Body and card attachments are message rendering, not files. Excluding these
 * is the whole difference between "the user sent a file" and "the user spoke".
 */
export function isBodyOrCard(contentType: string | undefined): boolean {
  const t = (contentType ?? '').toLowerCase()
  return (
    t.startsWith('text/html') ||
    t.startsWith('application/vnd.microsoft.card.') ||
    t.startsWith('application/vnd.microsoft.teams.card.')
  )
}

function kindOf(fileType: string | undefined, name: string | undefined): 'image' | 'file' {
  const ext = (fileType || extname(name ?? '').replace('.', '')).toLowerCase()
  return IMAGE_EXT.has(ext) ? 'image' : 'file'
}

/**
 * In-memory only, deliberately. Persisting these would write live OneDrive
 * bearer tokens to the state dir, which is exactly the leak the fixtures
 * caught in the other direction.
 */
export class AttachmentHandles {
  private readonly handles = new Map<string, Handle>()

  constructor(
    private readonly allowedHosts: string[] = DEFAULT_DOWNLOAD_HOSTS,
    /** The bot's Bot Framework token, for pasted images. Without it they are skipped. */
    private readonly botToken?: () => Promise<string>,
  ) {}

  /** Register the files on an activity and return what the model may see. */
  register(activity: Record<string, any>, now = Date.now()): InboundAttachment[] {
    this.prune(now)
    const out: InboundAttachment[] = []
    // The serviceUrl is bound to the inbound token (its serviceurl claim is
    // checked), so it is the one host a pasted image may come from.
    const serviceHost = hostOf(String(activity.serviceUrl ?? ''))
    const seen = new Set<string>()

    for (const raw of activity.attachments ?? []) {
      if (isBodyOrCard(raw?.contentType)) continue

      if (isHostedImage(raw?.contentType) && this.botToken && serviceHost) {
        const url = String(raw?.contentUrl ?? '')
        if (hostOf(url) !== serviceHost || seen.has(url)) continue
        seen.add(url)
        const id = createHash('sha256').update(`${activity.id}:${url}`).digest('hex').slice(0, 32)
        const subtype = String(raw.contentType).split('/')[1]?.toLowerCase() ?? ''
        const fileType = IMAGE_EXT.has(subtype) ? subtype : undefined
        const attachment: InboundAttachment = { id, mime: 'image', kind: 'image', fileType }
        this.handles.set(id, { ...attachment, url, at: now, auth: 'bot', host: serviceHost })
        out.push(attachment)
        continue
      }

      if (raw?.contentType !== TEAMS_FILE_DOWNLOAD_INFO) continue

      const url = String(raw?.content?.downloadUrl ?? '')
      if (!url) continue

      const fileType = raw?.content?.fileType ? String(raw.content.fileType) : undefined
      const name = raw?.name ? String(raw.name) : undefined
      // Derived from the activity + unique id so the same file in the same
      // message always gets the same handle, without exposing either.
      const id = createHash('sha256')
        .update(`${activity.id}:${raw?.content?.uniqueId ?? url}`)
        .digest('hex')
        .slice(0, 32)

      const attachment: InboundAttachment = { id, name, mime: raw?.contentType, kind: kindOf(fileType, name), fileType }
      this.handles.set(id, { ...attachment, url, at: now })
      out.push(attachment)
    }
    return out
  }

  /**
   * Fetch to `inboxDir` and return the local path.
   *
   * Errors deliberately never include the URL — a failed fetch is the most
   * likely place for a credential to end up in a log.
   */
  async download(id: string, inboxDir: string, now = Date.now()): Promise<string> {
    this.prune(now)
    const handle = this.handles.get(id)
    if (!handle) {
      throw new Error(
        'unknown attachment id — it may have expired, or belong to a message from before this session started',
      )
    }

    const res = handle.auth === 'bot' ? await this.fetchHostedImage(handle) : await this.fetchFile(handle.url)
    if (!res.ok) throw new Error(`attachment download failed: HTTP ${res.status}`)

    const tooLarge = () =>
      new Error(`attachment is larger than the ${MAX_DOWNLOAD_BYTES / 1024 / 1024}MB limit`)

    const declared = Number(res.headers.get('content-length') ?? '0')
    if (declared > MAX_DOWNLOAD_BYTES) throw tooLarge()

    // content-length is a claim, not a fact, and the sender chooses what the
    // download URL points at — so cap what we actually read rather than
    // buffering the whole body and checking afterwards. Otherwise the limit
    // only detects an overrun after already allocating it.
    const reader = res.body?.getReader()
    if (!reader) throw new Error('attachment download failed: empty response body')
    const chunks: Uint8Array[] = []
    let total = 0
    try {
      for (;;) {
        const { done, value } = await reader.read()
        if (done) break
        total += value.byteLength
        if (total > MAX_DOWNLOAD_BYTES) throw tooLarge()
        chunks.push(value)
      }
    } finally {
      reader.releaseLock()
    }
    const bytes = Buffer.concat(chunks)

    // A pasted image usually arrives as `image/*`; name it by its bytes so Read
    // recognises it as an image.
    const fileType = handle.kind === 'image' && !handle.fileType ? sniffImageExt(bytes) : handle.fileType
    const path = join(inboxDir, safeInboxName(id, { name: handle.name, fileType }))
    writeFileSync(path, bytes, { mode: 0o600 })
    return path
  }

  /**
   * Fork addition: the sender chooses this URL, so it may only point at the
   * tenant's own Microsoft hosts, and every redirect is checked the same way.
   * Upstream fetched any URL and followed redirects anywhere.
   */
  private async fetchFile(start: string): Promise<Response> {
    let url = start
    for (let hop = 0; hop <= MAX_REDIRECTS; hop++) {
      if (!isAllowedDownloadUrl(url, this.allowedHosts)) {
        throw new Error('attachment download refused: the file is not on an allowed Microsoft host')
      }
      let res: Response
      try {
        res = await fetch(url, { redirect: 'manual' })
      } catch {
        throw new Error('attachment download failed: the file could not be reached')
      }
      const location = res.status >= 300 && res.status < 400 ? res.headers.get('location') : null
      if (!location) return res
      url = new URL(location, url).toString()
    }
    throw new Error('attachment download failed: too many redirects')
  }

  /**
   * The bot token goes to the serviceUrl host and nowhere else: the host is
   * re-checked here, and a redirect is refused rather than followed, since
   * following it would hand the token's response to wherever it points.
   */
  private async fetchHostedImage(handle: Handle): Promise<Response> {
    if (!this.botToken || !handle.host || hostOf(handle.url) !== handle.host) {
      throw new Error('attachment download refused: the image is not on the Teams service host')
    }
    let res: Response
    try {
      res = await fetch(handle.url, {
        redirect: 'manual',
        headers: { Authorization: `Bearer ${await this.botToken()}` },
      })
    } catch {
      throw new Error('attachment download failed: the image could not be reached')
    }
    if (res.status >= 300 && res.status < 400) {
      throw new Error('attachment download refused: the Teams service redirected the image')
    }
    return res
  }

  private prune(now: number): void {
    for (const [id, handle] of this.handles) {
      if (now - handle.at > HANDLE_TTL_MS) this.handles.delete(id)
    }
  }
}

/**
 * The sender controls `name`, so it never reaches the filesystem intact. The
 * handle id is ours and already filename-safe; the sender's extension is kept
 * only as far as it is alphanumeric, so a `.png` cannot smuggle a path.
 */
export function safeInboxName(id: string, attachment: { name?: string; fileType?: string }): string {
  const raw = attachment.fileType || extname(attachment.name ?? '').replace('.', '')
  const ext = raw.replace(/[^a-zA-Z0-9]/g, '').slice(0, 10)
  return ext ? `${id}.${ext}` : id
}
