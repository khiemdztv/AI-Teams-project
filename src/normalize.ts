/**
 * Teams activity -> `notifications/claude/channel` payload.
 *
 * Two contract rules that are easy to get wrong and silently lossy:
 *
 *   - `meta` keys must match /^[A-Za-z0-9_]+$/. Hyphenated keys are dropped
 *     without error, so a typo here becomes an invisible bug.
 *   - Every meta value must be a string.
 *
 * File paths go in `meta`, never in `content`. An in-content annotation like
 * "[image attached — read: /path]" is forgeable by any sender who simply types
 * that string; a meta key is not.
 */

import { normalizeConversationId, extractThreadId, type ConversationType } from './gate.js'

export type ChannelNotification = {
  content: string
  meta: Record<string, string>
}

export type Attachment = {
  id: string
  kind: string
  size?: number
  mime?: string
  name?: string
}

export type NormalizeInput = {
  activity: Record<string, any>
  /** Images we eagerly downloaded AFTER the gate passed, in arrival order. */
  imagePaths?: string[]
  /** Every real file on the activity — not just the first of each kind. */
  attachments?: Attachment[]
  /**
   * Fork addition (02/11): the sender's role from access.json, so Claude knows
   * whether to show technical detail. Set by the server after the gate, never
   * read from the message.
   */
  role?: 'admin' | 'user'
  /** Fork addition (05/11): the message carried "/force" (src/force.ts). Set by the server. */
  force?: boolean
  /** Fork addition (02/11): cloud files shared into the chat, by address. */
  sharedFiles?: { name?: string; url: string }[]
  /**
   * Fork addition (17/11): for a group-chat message, the sender's own 1:1 chat
   * with the bot, when there is one, so their private data can go there
   * instead of to everyone in the group. Set by the server.
   */
  dmConversationId?: string
}

const META_KEY = /^[A-Za-z0-9_]+$/

/** Filenames reach the model as meta; strip path separators and control chars. */
export function sanitizeFilename(name: string): string {
  return name
    .replace(/[\x00-\x1f\x7f]/g, '')
    .replace(/[\\/]/g, '_')
    .slice(0, 200)
}

function put(meta: Record<string, string>, key: string, value: unknown): void {
  if (value === undefined || value === null || value === '') return
  if (!META_KEY.test(key)) {
    // Fail loud in our own tests rather than let the harness drop it silently.
    throw new Error(`meta key "${key}" is not [A-Za-z0-9_]+ and would be dropped`)
  }
  meta[key] = String(value)
}

/**
 * Fork addition (17/11): drop the bot's own @mention ("<at>Fabrikam Agent</at>")
 * from a group message, so Claude reads the request, not the bot's name. Only
 * the mention whose entity points at the bot's id goes; a colleague's mention
 * ("nhắn cho <at>An</at>") is part of what is being asked.
 */
export function withoutBotMention(text: string, entities: unknown, botId: string | undefined): string {
  if (!botId || !Array.isArray(entities)) return text
  let out = text
  for (const e of entities as { type?: string; text?: string; mentioned?: { id?: string } }[]) {
    if (e?.type === 'mention' && e.mentioned?.id === botId && typeof e.text === 'string' && e.text) {
      out = out.split(e.text).join('')
    }
  }
  return out === text ? text : out.replace(/[ \t]{2,}/g, ' ').trim()
}

export function normalize(input: NormalizeInput): ChannelNotification {
  const a = input.activity
  const conversationType = (a.conversation?.conversationType ?? 'personal') as ConversationType
  const meta: Record<string, string> = {}

  const rawConversationId = String(a.conversation?.id ?? '')
  put(meta, 'conversation_id', normalizeConversationId(rawConversationId))
  put(meta, 'conversation_type', conversationType)
  put(meta, 'message_id', a.id)
  // Channels only, and the one value a reply must thread on — see extractThreadId.
  put(meta, 'thread_id', extractThreadId(rawConversationId, conversationType, a.replyToId))
  put(meta, 'user', a.from?.name)
  put(meta, 'user_id', a.from?.aadObjectId)
  if (input.role) put(meta, 'role', input.role)
  if (input.force) put(meta, 'force', 'true')
  if (conversationType !== 'personal') put(meta, 'dm_conversation_id', input.dmConversationId)
  // Fork (18/11): a Teams meeting's chat; meeting_transcript works here.
  if (a.channelData?.meeting?.id) put(meta, 'meeting_chat', 'true')
  put(meta, 'tenant_id', a.channelData?.tenant?.id ?? a.conversation?.tenantId)
  put(meta, 'ts', a.timestamp ?? new Date().toISOString())

  // Never observed on this tenant — Teams leaves it absent even on a genuine
  // thread reply, which is why threading uses the conversation id instead
  // (tests/fixtures.test.ts asserts the absence). Emitted only so a tenant or
  // activity type that does set it is not silently discarded; do not build on it.
  put(meta, 'reply_to_id', a.replyToId)

  const imagePaths = input.imagePaths ?? []
  const attachments = input.attachments ?? []

  // The first of each kind keeps its own key: the overwhelmingly common case is
  // one file, and a single key is what the instructions teach.
  if (imagePaths[0]) put(meta, 'image_path', imagePaths[0])
  if (imagePaths.length > 1) put(meta, 'image_paths', imagePaths.join('; '))

  const firstFile = attachments.find(f => f.kind !== 'image')
  if (firstFile) {
    put(meta, 'attachment_id', firstFile.id)
    put(meta, 'attachment_kind', firstFile.kind)
    if (firstFile.size != null) put(meta, 'attachment_size', String(firstFile.size))
    put(meta, 'attachment_mime', firstFile.mime)
    if (firstFile.name) put(meta, 'attachment_name', sanitizeFilename(firstFile.name))
  }

  // But a second file must not become unreachable. Count plus a listing that
  // carries every handle, matching discord's `attachment_count`/`attachments`
  // pair (external_plugins/discord/server.ts @ db253f26).
  if (attachments.length) {
    put(meta, 'attachment_count', String(attachments.length))
    if (attachments.length > 1) {
      put(
        meta,
        'attachments',
        attachments
          .map(f => `${sanitizeFilename(f.name ?? 'file')} (${f.mime ?? 'unknown'}) id=${f.id}`)
          .join('; '),
      )
    }
  }

  const shared = input.sharedFiles ?? []
  if (shared[0]) {
    put(meta, 'shared_file_url', shared[0].url)
    if (shared[0].name) put(meta, 'shared_file_name', sanitizeFilename(shared[0].name))
  }
  if (shared.length > 1) {
    put(meta, 'shared_files', shared.map(f => `${sanitizeFilename(f.name ?? 'file')} ${f.url}`).join('; '))
  }

  const text = withoutBotMention(typeof a.text === 'string' ? a.text : '', a.entities, a.recipient?.id)
  const named = attachments.find(f => f.name)?.name ?? shared.find(f => f.name)?.name
  const fallback = named ? `(${sanitizeFilename(named)})` : imagePaths.length ? '(image)' : ''

  return { content: text || fallback, meta }
}
