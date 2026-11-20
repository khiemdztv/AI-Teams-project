/**
 * Fork (18/11): Teams activity-feed notifications ("the bell") for people in a
 * group chat, in place of the @mention the bot cannot make.
 *
 * Graph `POST /chats/{chatId}/sendActivityNotification` with the app's own
 * token, under the chat's RSC grant TeamsActivity.Send.Chat (manifest 1.4.0).
 * The activity types and their text are declared in the manifest:
 *   taskAssigned  "{by} giao việc cho bạn: {task}"
 *   reminder      "{by} nhắc bạn: {task}"
 * Recipients must be members of that chat, matched against its member list
 * (Bot Framework getMembers), never looked up in the directory: the bot only
 * rings people who can already read the chat. Clicking the notification opens
 * the chat.
 */

const GRAPH = 'https://graph.microsoft.com/v1.0'

export const ACTIVITY_KINDS = ['taskAssigned', 'reminder'] as const
export type ActivityKind = (typeof ACTIVITY_KINDS)[number]

/** Teams cuts the preview at 150 characters; keep the task readable. */
export const MAX_TASK = 150
export const MAX_RECIPIENTS = 20

export type ChatMember = { aadObjectId?: string; name?: string; email?: string; userPrincipalName?: string }

const fold = (s: string) =>
  s.normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().trim()

const label = (m: ChatMember) => `${m.name ?? '?'}${m.email || m.userPrincipalName ? ` <${m.email || m.userPrincipalName}>` : ''}`

/** Words that mean "everyone in this chat" (accents ignored). */
const EVERYONE = new Set(['all', 'everyone', 'ca nhom', 'moi nguoi', 'tat ca'])

/**
 * The chat members `wanted` names: by email, or by a name as said (accents
 * ignored, part of the display name or the email's local part). A name that
 * matches several members, or nobody, throws with what was found. "all" (or
 * "cả nhóm", "mọi người") is every member except `exclude`, the sender.
 */
export function matchMembers(members: ChatMember[], wanted: string[], exclude?: string): ChatMember[] {
  const people = members.filter(m => m.aadObjectId)
  const chosen = new Map<string, ChatMember>()
  const skip = (exclude ?? '').toLowerCase()
  for (const entry of wanted.map(w => String(w).trim()).filter(Boolean)) {
    if (EVERYONE.has(fold(entry))) {
      for (const m of people) if (m.aadObjectId!.toLowerCase() !== skip) chosen.set(m.aadObjectId!, m)
      continue
    }
    const key = fold(entry)
    const hits = people.filter(m => {
      const mail = fold(m.email || m.userPrincipalName || '')
      return entry.includes('@') ? mail === key : fold(m.name ?? '').includes(key) || mail.split('@')[0] === key
    })
    if (!hits.length) throw new Error(`"${entry}" is not a member of this chat`)
    if (hits.length > 1) throw new Error(`"${entry}" matches several people in this chat: ${hits.map(label).join('; ')}. Ask which.`)
    chosen.set(hits[0]!.aadObjectId!, hits[0]!)
  }
  if (!chosen.size) throw new Error('people is required: who to notify, by name or email')
  if (chosen.size > MAX_RECIPIENTS) throw new Error(`at most ${MAX_RECIPIENTS} people at once`)
  return [...chosen.values()]
}

/** One Graph call's outcome, for the operator's log: no token, no task text, no names. */
export type NotifyTrace = { recipient: string; status: number; requestId: string; code?: string }

/** The bell's line, the same words as the manifest templates. */
export function notificationText(kind: ActivityKind, by: string, task: string): string {
  return kind === 'taskAssigned' ? `${by} giao việc cho bạn: ${task}` : `${by} nhắc bạn: ${task}`
}

export type NotifyDeps = {
  graphToken: () => Promise<string>
  fetchImpl?: typeof fetch
  /** Fork (18/11): called after every Graph call, success or not. */
  trace?: (t: NotifyTrace) => void
  /**
   * Fork (18/11): the org-catalog id of the Teams app (Teams admin center ->
   * Manage apps -> App ID), not the bot id. When two apps share the bot's
   * Entra id (an old custom upload next to the current one), Graph answers 204
   * yet the bell stays empty unless the notification names the right app.
   */
  teamsAppId?: string
}

export async function sendActivity(
  input: { chatId: string; recipients: ChatMember[]; kind: ActivityKind; task: string; by: string },
  deps: NotifyDeps,
): Promise<string> {
  if (!ACTIVITY_KINDS.includes(input.kind)) throw new Error(`kind must be one of: ${ACTIVITY_KINDS.join(', ')}`)
  const task = input.task.replace(/\s+/g, ' ').trim()
  if (!task) throw new Error('task is required: what the notification says')
  if (task.length > MAX_TASK) throw new Error(`task is too long (${task.length} characters, at most ${MAX_TASK}); shorten it, the details go in the chat`)
  const by = input.by.replace(/\s+/g, ' ').trim() || 'Fabrikam Agent'
  const fetchImpl = deps.fetchImpl ?? fetch
  const token = await deps.graphToken()
  const url = `${GRAPH}/chats/${encodeURIComponent(input.chatId)}/sendActivityNotification`

  const sent: string[] = []
  for (const r of input.recipients) {
    const res = await fetchImpl(url, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        topic: { source: 'entityUrl', value: `${GRAPH}/chats/${input.chatId}` },
        // Fork (18/11): the reserved systemDefault type, not the manifest's
        // taskAssigned/reminder. Those must match the app version installed
        // where the notification lands; with an old upload sharing the bot id,
        // Graph answered 204 and Teams showed nothing. systemDefault needs no
        // manifest entry, so the same words go in its free-form text.
        activityType: 'systemDefault',
        previewText: { content: task },
        recipient: { '@odata.type': 'microsoft.graph.aadUserNotificationRecipient', userId: r.aadObjectId },
        templateParameters: [{ name: 'systemDefaultText', value: notificationText(input.kind, by, task) }],
        ...(deps.teamsAppId ? { teamsAppId: deps.teamsAppId } : {}),
      }),
    })
    const requestId = res.headers.get('request-id') ?? ''
    if (!res.ok) {
      const body = (await res.json().catch(() => ({}))) as any
      const code = String(body?.error?.code ?? '')
      deps.trace?.({ recipient: String(r.aadObjectId), status: res.status, requestId, code })
      const reason =
        res.status === 403 || res.status === 401
          ? 'no permission: the chat needs the app version that asks for activity notifications (1.4.0), with consent'
          : `request failed (${res.status}${code ? ` ${code}` : ''})`
      throw new Error(`${sent.length ? `notified ${sent.join(', ')}; ` : ''}NOT notified ${label(r)} or anyone after: ${reason}`)
    }
    deps.trace?.({ recipient: String(r.aadObjectId), status: res.status, requestId })
    sent.push(label(r))
  }
  return `Sent a "${input.kind}" notification to ${sent.length} ${sent.length === 1 ? 'person' : 'people'}: ${sent.join('; ')}. It shows in their Teams Activity and opens this chat.`
}

/**
 * Fork (19/11): a Teams deep link that starts a call from whoever clicks it.
 * Graph cannot place a call as a user (its calling API calls as the bot), so
 * the bot hands the requester a button; one click and Teams calls from their
 * own account. Several people make a group call.
 */
export function callLink(emails: string[], video: boolean): string {
  const users = [...new Set(emails.map(e => e.trim().toLowerCase()).filter(e => /^[^\s@,]+@[^\s@,]+\.[^\s@,]+$/.test(e)))]
  if (!users.length) throw new Error('nobody to call: give people by name in a group chat, or by email')
  if (users.length > MAX_RECIPIENTS) throw new Error(`at most ${MAX_RECIPIENTS} people in one call`)
  return `https://teams.microsoft.com/l/call/0/0?users=${users.map(u => encodeURIComponent(u).replace(/%40/g, '@')).join(',')}${video ? '&withVideo=true' : ''}`
}

/** The email Teams calls a member by. */
export function memberEmail(m: ChatMember): string {
  return String(m.email || m.userPrincipalName || '').toLowerCase()
}
