/**
 * Permission-relay grammar and pending-request bookkeeping.
 *
 * Split out of server.ts so the security-relevant properties — what counts as a
 * verdict, who may give it, for how long, and that it fires at most once — are
 * unit-testable without a live tenant.
 */

/**
 * From the same spec the official plugins inline (anthropics/claude-cli-internal
 * `src/services/mcp/channelPermissions.ts`): five lowercase letters with 'l'
 * excluded so it cannot be misread as '1'. Case-insensitive, because phone
 * keyboards capitalize the first word.
 *
 * Anchored at both ends deliberately. A bare "yes", or "yes do it abcde", is
 * conversation and must reach the session as chat — silently swallowing it as a
 * verdict would be both a lost message and an unintended approval.
 */
export const PERMISSION_REPLY_RE = /^\s*(y|yes|n|no)\s+([a-km-z]{5})\s*$/i

/**
 * How long a prompt stays answerable. A verdict typed long after the prompt was
 * sent is more likely a mistake (or a stale message scrolled into view) than a
 * decision about what Claude is doing now.
 */
export const PERMISSION_TTL_MS = 10 * 60 * 1000

/**
 * Fork addition (02/11): a prompt raised for a non-admin's message waits this
 * long for an admin, then is denied, so the person who asked is told it cannot
 * be done now instead of waiting on someone who may not be looking.
 */
export const DELEGATED_TTL_MS = 5 * 60 * 1000

/** Caps for what the approver sees. Long enough to judge a command, short enough for a phone. */
export const MAX_DESCRIPTION_CHARS = 300
export const MAX_PREVIEW_CHARS = 1500

export type Verdict = { requestId: string; behavior: 'allow' | 'deny' }

/** Parse a message as a permission verdict, or undefined if it is just chat. */
export function parseVerdict(text: string): Verdict | undefined {
  const m = PERMISSION_REPLY_RE.exec(text)
  if (!m) return undefined
  return {
    requestId: m[2]!.toLowerCase(),
    behavior: m[1]!.toLowerCase().startsWith('y') ? 'allow' : 'deny',
  }
}

/** Strip control characters other than newline and tab, so nothing can redraw the message. */
function clean(text: string): string {
  return text.replace(/\r\n?/g, '\n').replace(/[\x00-\x08\x0b-\x1f\x7f]/g, '')
}

/** Truncate, saying how much was cut — a silent cut is how a malicious tail hides. */
function truncate(text: string, max: number): string {
  if (text.length <= max) return text
  return `${text.slice(0, max)}\n… (đã cắt ${text.length - max} ký tự)`
}

export type PermissionPrompt = {
  requestId: string
  toolName: string
  description: string
  inputPreview: string
  expiresAt: number
  /** Fork: set when the prompt goes to admins on behalf of a non-admin: their display name. */
  onBehalfOf?: string
}

/**
 * The text the approver reads. Upstream sent the tool name alone, so the CEO
 * would approve "Bash" without seeing the command; the description and input
 * preview Claude Code supplies are what make the decision an informed one.
 *
 * Sent as plain text, so the preview needs no escaping — only control
 * characters are stripped. The tool name stays on one line so it cannot forge
 * the instruction that follows it.
 */
export function formatPermissionPrompt(p: PermissionPrompt, timeZone?: string): string {
  const tool = clean(p.toolName).replace(/\s+/g, ' ').slice(0, 120)
  const description = truncate(clean(p.description).trim(), MAX_DESCRIPTION_CHARS)
  const preview = truncate(clean(p.inputPreview).trim(), MAX_PREVIEW_CHARS)
  const expires = new Date(p.expiresAt).toLocaleTimeString('vi-VN', {
    hour: '2-digit',
    minute: '2-digit',
    timeZone,
  })
  const behalf = p.onBehalfOf === undefined ? '' : clean(p.onBehalfOf).replace(/\s+/g, ' ').slice(0, 80) || 'người dùng'
  return [
    `🔐 Claude xin quyền dùng: ${tool} (mã ${p.requestId})`,
    behalf ? `Cho yêu cầu của ${behalf}. Người đó không thấy tin này; hết hạn mà chưa ai trả lời thì tự từ chối.` : '',
    description ? `\n${description}` : '',
    preview ? `\nNội dung:\n${preview}` : '',
    `\nTrả lời "y ${p.requestId}" để cho phép, hoặc "n ${p.requestId}" để từ chối. Hết hạn lúc ${expires}.`,
  ]
    .filter(Boolean)
    .join('\n')
}

/**
 * Fork addition: what an approver reads when they send anything other than a
 * well-formed verdict while Claude waits on them — typically the code with a
 * letter missing. Without it they got the generic busy notice, which reads as
 * if their approval had been ignored.
 */
export function formatPendingReminder(pending: { requestId: string; toolName: string }[]): string {
  const lines = pending.map(p => {
    const tool = clean(p.toolName).replace(/\s+/g, ' ').slice(0, 120)
    return `🔐 Claude đang chờ bạn xác nhận ${tool} (mã ${p.requestId}): trả lời "y ${p.requestId}" để cho phép, hoặc "n ${p.requestId}" để từ chối.`
  })
  return [...lines, 'Tin vừa rồi không khớp mã nên chưa được chuyển cho Claude.'].join('\n')
}

export type TakeResult = 'ok' | 'unknown' | 'expired' | 'not_approver'

type Live = { toolName: string; approvers: Set<string>; expiresAt: number }

/**
 * Fork addition: who a permission prompt goes to. Only the person whose Teams
 * message Claude is working on, in their 1:1 chat, and only while they are
 * still allowlisted; the other users neither see the prompt nor can answer it.
 * With no Teams turn in progress the request did not come from Teams, so it is
 * answered in the terminal and nobody is messaged.
 *
 * With `admins` given (02/11), a non-admin is never asked: the prompt goes to
 * the admins' 1:1 chats instead, so someone who does not read tool calls is
 * never handed one to approve. Admins must still be allowlisted.
 */
export function permissionTargets<R extends { conversationType?: string; senderId?: string }>(
  refs: R[],
  allowed: Set<string>,
  turnSenderId: string | undefined,
  admins?: Set<string>,
): R[] {
  if (!turnSenderId || !allowed.has(turnSenderId)) return []
  const askers =
    !admins || admins.has(turnSenderId) ? new Set([turnSenderId]) : new Set([...admins].filter(id => allowed.has(id)))
  return refs.filter(ref => ref.conversationType === 'personal' && !!ref.senderId && askers.has(ref.senderId))
}

/**
 * Live permission requests awaiting a verdict.
 *
 * A verdict counts only when it comes from a 1:1 chat with one of the people
 * the prompt was sent to, before it expires. Upstream accepted any sender who
 * passed the inbound gate, from any conversation — including a group chat the
 * prompt was never posted in.
 *
 * `take` is still the one-shot gate: it returns 'ok' at most once per request
 * id, so a duplicate Bot Framework redelivery, or the sender typing the same
 * answer twice, cannot emit a second verdict. A reply from the wrong person
 * leaves the request live, so the right approver can still answer it.
 */
export class PendingPermissions {
  private readonly live = new Map<string, Live>()

  add(requestId: string, toolName: string, approvers: string[], now = Date.now(), ttlMs = PERMISSION_TTL_MS): number {
    const expiresAt = now + ttlMs
    this.live.set(requestId, { toolName, approvers: new Set(approvers), expiresAt })
    return expiresAt
  }

  /** Fork addition: who the prompt went to, while it is still live. */
  approversOf(requestId: string): string[] {
    return [...(this.live.get(requestId)?.approvers ?? [])]
  }

  /**
   * Fork addition: retire a request nobody answered, so it can be denied on
   * the asker's behalf. True only if it was still live: a verdict that won the
   * race keeps its outcome, and a second call does nothing.
   */
  withdraw(requestId: string): boolean {
    return this.live.delete(requestId)
  }

  /** Claim a request on behalf of `senderId`. 'ok' only for the first valid caller. */
  take(
    requestId: string,
    senderId: string,
    conversationType: string | undefined,
    now = Date.now(),
  ): TakeResult {
    const entry = this.live.get(requestId)
    if (!entry) return 'unknown'
    if (entry.expiresAt < now) {
      this.live.delete(requestId)
      return 'expired'
    }
    if (conversationType !== 'personal' || !entry.approvers.has(senderId)) return 'not_approver'
    this.live.delete(requestId)
    return 'ok'
  }

  /**
   * Fork addition: the live requests `senderId` could answer from this
   * conversation, oldest first. Read-only — it never claims anything — and
   * held to the same rules as `take`, so a reminder only ever goes to someone
   * the prompt was already sent to.
   */
  awaiting(
    senderId: string,
    conversationType: string | undefined,
    now = Date.now(),
  ): { requestId: string; toolName: string }[] {
    if (conversationType !== 'personal') return []
    const out: { requestId: string; toolName: string }[] = []
    for (const [requestId, entry] of this.live) {
      if (entry.expiresAt >= now && entry.approvers.has(senderId)) {
        out.push({ requestId, toolName: entry.toolName })
      }
    }
    return out
  }

  get size(): number {
    return this.live.size
  }
}
