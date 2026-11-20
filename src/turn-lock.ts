/**
 * One Teams request at a time (fork addition).
 *
 * While Claude works on a message, the next one is refused with a notice
 * instead of queueing behind it, and the sender sees a status line that turns
 * from "⏳" to "✅" when the turn ends.
 *
 * The lock is a file because two processes share it: the listener writes it
 * when it hands a message to Claude, and the plugin's Stop hook
 * (`hooks/turn-done.ts`, a separate process Claude Code spawns when the turn
 * ends) deletes it. The listener notices the file is gone and closes the
 * status line. Nothing else crosses: the hook needs no credentials.
 *
 * A turn the hook never closes (an interrupt in the terminal fires no Stop
 * hook) is released after MAX_TURN_MS, so the bot cannot stay locked.
 *
 * Fork (20/11): a message that arrives mid-turn is no longer refused but
 * queued in the listener (src/turn-queue.ts) and started when this lock is
 * released. Each lock carries a `turnId` that the channel notification repeats
 * as meta turn_id, so the Stop hook can tell the Teams turn from another turn
 * that ended first (see transcriptHasTurn()).
 */

import { readFileSync, writeFileSync, rmSync, renameSync, openSync, readSync, fstatSync, closeSync } from 'fs'
import { randomBytes } from 'crypto'
import { join } from 'path'

export const TURN_LOCK_FILE = 'turn.json'

/** Longest a turn may hold the lock before new messages are accepted again. */
export const MAX_TURN_MS = 15 * 60 * 1000

export type TurnLock = {
  since: number
  conversationId: string
  /**
   * Entra object id of the person whose message started the turn, as the
   * inbound gate verified it. The m365 server reads it to act in that
   * person's mailbox: the model never gets to say whose mailbox it is.
   */
  senderId?: string
  /** Fork (02/11): their Teams display name, for a prompt sent to admins on their behalf. */
  senderName?: string
  /** Fork (02/11): their role; a 'user' gets the status lines in assistant wording, without icons. */
  role?: 'admin' | 'user'
  /**
   * Fork (05/11): set when a scheduled job (src/schedule.ts) started the turn
   * rather than a message. Such a turn is read-only: see src/scheduled-guard.ts.
   */
  scheduled?: 'morning' | 'evening'
  /**
   * Fork (13/11): the message carried /force (src/force.ts), so the sender
   * said to go ahead without confirming. Lets watch_outlook keep a task to do
   * later, when nobody will be there to confirm it.
   */
  force?: boolean
  /** The "⏳" status message, edited to "✅" when the turn ends. */
  statusMessageId?: string
  /**
   * Fork (18/11): where the status lines go, kept in memory only (set after
   * the lock is written). Absent: the turn's own conversation. null: nowhere.
   * See statusTarget().
   */
  statusConversationId?: string | null
  /**
   * Fork (20/11): 8 hex chars, also sent to Claude as meta turn_id with the
   * notification that starts the turn. Absent on locks from before 20/11.
   */
  turnId?: string
}

/** Fork (20/11): a fresh id for a turn: 8 random hex chars. */
export function newTurnId(): string {
  return randomBytes(4).toString('hex')
}

/** How much of the transcript's tail is read first when looking for a turn id. */
export const TRANSCRIPT_TAIL_BYTES = 512 * 1024
/** Never read further back than this, however long the transcript is. */
export const TRANSCRIPT_MAX_BYTES = 16 * 1024 * 1024

/**
 * Fork (20/11): whether `text` holds the channel tag attribute turn_id="<id>",
 * as the transcript (JSONL, so the quotes are usually escaped) records it.
 */
export function mentionsTurn(text: string, turnId: string): boolean {
  if (!/^[0-9a-f]{8}$/.test(turnId)) return false
  return new RegExp(`turn_id=\\\\*"${turnId}\\\\*"`).test(text)
}

/**
 * Fork (20/11): whether the transcript at `path` already holds the channel
 * message of turn `turnId`. A background subagent finishing, for one, makes the
 * idle session run a turn of its own, and its Stop must not close a Teams turn
 * whose message Claude Code is still holding back for the next turn.
 *
 * Reads the last TRANSCRIPT_TAIL_BYTES first; when the id is not there and the
 * file is longer, looks further back in chunks up to TRANSCRIPT_MAX_BYTES, so a
 * long Teams turn whose tool output pushed its own message out of the tail is
 * still recognised. The id is random, so reading more can only find the right
 * turn, never a wrong one. Returns null when the file cannot be read: the caller
 * then behaves as before 20/11.
 */
export function transcriptHasTurn(path: string, turnId: string): boolean | null {
  let fd: number
  try {
    fd = openSync(path, 'r')
  } catch {
    return null
  }
  try {
    const size = fstatSync(fd).size
    const overlap = 64
    let end = size
    let chunk = TRANSCRIPT_TAIL_BYTES
    while (end > 0 && size - end < TRANSCRIPT_MAX_BYTES) {
      const start = Math.max(0, end - chunk)
      const readEnd = Math.min(size, end + overlap)
      const buf = Buffer.alloc(readEnd - start)
      readSync(fd, buf, 0, buf.length, start)
      if (mentionsTurn(buf.toString('utf8'), turnId)) return true
      end = start
      chunk = TRANSCRIPT_TAIL_BYTES * 2
    }
    return false
  } catch {
    return null
  } finally {
    try {
      closeSync(fd)
    } catch {}
  }
}

/**
 * Fork (20/11): the Stop hook leaves the lock alone (no clearing, no nudge) when
 * the lock names a turn the transcript does not show yet. A lock without a
 * turnId, or a transcript that cannot be read, behaves as before.
 */
export function stopIsForOtherTurn(lock: TurnLock | null, transcriptPath: unknown): boolean {
  if (!lock?.turnId || typeof transcriptPath !== 'string' || !transcriptPath) return false
  return transcriptHasTurn(transcriptPath, lock.turnId) === false
}

/**
 * Fork (18/11): where "⏳/✅" and busy notices for a message go. In a group
 * chat they are internal reports, so they go to the sender's own 1:1 chat with
 * the bot, or nowhere when there is none; elsewhere, to the conversation itself.
 */
export function statusTarget(
  conversationType: string | undefined,
  conversationId: string,
  dmConversationId: string | undefined,
): string | null {
  if (conversationType !== 'groupChat') return conversationId
  return dmConversationId ?? null
}

/** The conversation `lock`'s status lines go to, or null for none. */
export function statusConversationOf(lock: TurnLock): string | null {
  return lock.statusConversationId === undefined ? lock.conversationId : lock.statusConversationId
}

export function turnLockPath(stateDir: string): string {
  return join(stateDir, TURN_LOCK_FILE)
}

export function readTurnLock(stateDir: string): TurnLock | null {
  try {
    const parsed = JSON.parse(readFileSync(turnLockPath(stateDir), 'utf8'))
    if (typeof parsed?.since === 'number' && typeof parsed?.conversationId === 'string') return parsed
  } catch {}
  return null
}

export function writeTurnLock(stateDir: string, lock: TurnLock): void {
  const target = turnLockPath(stateDir)
  const tmp = `${target}.tmp`
  writeFileSync(tmp, JSON.stringify(lock), { mode: 0o600 })
  renameSync(tmp, target)
}

/** Returns true when there was a lock to clear. */
export function clearTurnLock(stateDir: string): boolean {
  try {
    rmSync(turnLockPath(stateDir))
    return true
  } catch {
    return false
  }
}

/**
 * Fork (18/11): a turn that ends without `reply` leaves the sender with only
 * "✅ Xong": the answer was printed to the terminal, which Teams never sees.
 * The listener records here the `since` of the turn whose reply went out, and
 * the Stop hook (hooks/turn-done.ts) sends Claude back once to reply when it
 * does not match. A separate file, so turn.json is still written only once.
 */
export const REPLIED_FILE = 'turn-replied'

export const NO_REPLY_NUDGE =
  'Lượt này chưa có gì tới Teams: người nhắn chỉ thấy "✅ Xong". Văn bản in ra terminal không tới được Teams. ' +
  'Gọi công cụ reply của msteams để gửi câu trả lời (conversation_id của tin vừa nhận, hoặc dm_conversation_id cho báo cáo riêng trong chat nhóm), rồi mới kết thúc.'

export function markReplied(stateDir: string, since: number): void {
  try {
    writeFileSync(join(stateDir, REPLIED_FILE), String(since), { mode: 0o600 })
  } catch {}
}

export function repliedSince(stateDir: string): number | null {
  try {
    const n = Number(readFileSync(join(stateDir, REPLIED_FILE), 'utf8').trim())
    return Number.isFinite(n) ? n : null
  } catch {
    return null
  }
}

/**
 * Whether the Stop hook should send Claude back to reply: a Teams turn is open,
 * nothing was replied in it, and this stop is not already the one after a nudge.
 */
export function shouldNudgeReply(lock: TurnLock | null, replied: number | null, stopHookActive: boolean): boolean {
  return !!lock && !stopHookActive && replied !== lock.since
}

export function isStale(lock: TurnLock, now = Date.now()): boolean {
  return now - lock.since > MAX_TURN_MS
}

function clock(ms: number, timeZone?: string): string {
  return new Date(ms).toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit', timeZone })
}

// Fork (20/11): messages sent meanwhile are queued now, not refused.
export const STATUS_WORKING = '⏳ Đang xử lý… Tin nhắn mới gửi lúc này sẽ được xếp hàng, làm sau khi có dòng ✅ Xong.'

/** How often the "⏳" line is rewritten with the time spent so far. */
export const STATUS_TICK_MS = 60 * 1000

/**
 * Fork (02/11): the same four lines for a non-admin, worded as an assistant
 * would speak to the boss — no icons, no "Claude", no "bot", no "✅ Xong"
 * protocol to learn. Admins keep the original lines.
 */
export const USER_WORKING = 'Dạ, em đang làm yêu cầu này, xin đợi em một chút.'
export const USER_TIMED_OUT = 'Dạ, yêu cầu vừa rồi em làm lâu hơn dự kiến. Nếu chưa thấy kết quả, xin nhắn lại giúp em.'

/** The "⏳" line after `now - lock.since`, in whole minutes once past the first. */
export function formatWorking(lock: TurnLock, now = Date.now()): string {
  const minutes = Math.floor((now - lock.since) / 60_000)
  if (lock.role === 'user') {
    return minutes < 1 ? USER_WORKING : `Dạ, em vẫn đang làm (đã ${minutes} phút), xin đợi em thêm một chút.`
  }
  if (minutes < 1) return STATUS_WORKING
  return `⏳ Đang xử lý… (đã ${minutes} phút). Tin nhắn mới gửi lúc này sẽ được xếp hàng, làm sau khi có dòng ✅ Xong.`
}

/**
 * To someone whose message arrived mid-turn; `role` is theirs, not the turn's.
 * Fork (20/11): only when the waiting line is full (src/turn-queue.ts).
 */
export function formatBusyNotice(lock: TurnLock, timeZone?: string, role?: 'admin' | 'user'): string {
  if (role === 'user') {
    return 'Dạ, em đang làm dở yêu cầu trước nên chưa nhận được tin vừa rồi. Khi em báo xong, xin nhắn lại giúp em.'
  }
  return (
    `⏳ Claude vẫn đang xử lý yêu cầu trước (từ ${clock(lock.since, timeZone)}). ` +
    'Tin vừa rồi chưa được chuyển cho Claude. Chờ dòng ✅ Xong rồi nhắn lại.'
  )
}

export function formatDone(lock: TurnLock, now = Date.now()): string {
  if (lock.role === 'user') return 'Dạ, em làm xong rồi. Cần gì thêm, xin cứ nhắn em.'
  const seconds = Math.max(1, Math.round((now - lock.since) / 1000))
  const took = seconds < 60 ? `${seconds} giây` : `${Math.floor(seconds / 60)} phút ${seconds % 60} giây`
  return `✅ Xong (${took}). Có thể nhắn yêu cầu tiếp theo.`
}

export const STATUS_TIMED_OUT =
  '⚠️ Không thấy Claude báo xong sau 15 phút. Bot nhận tin mới trở lại; nếu chưa có câu trả lời, nhắn lại yêu cầu.'

export function formatTimedOut(lock: TurnLock): string {
  return lock.role === 'user' ? USER_TIMED_OUT : STATUS_TIMED_OUT
}
