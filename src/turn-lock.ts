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
 */

import { readFileSync, writeFileSync, rmSync, renameSync } from 'fs'
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
  /** The "⏳" status message, edited to "✅" when the turn ends. */
  statusMessageId?: string
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

export function isStale(lock: TurnLock, now = Date.now()): boolean {
  return now - lock.since > MAX_TURN_MS
}

function clock(ms: number, timeZone?: string): string {
  return new Date(ms).toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit', timeZone })
}

export const STATUS_WORKING = '⏳ Đang xử lý… Tin nhắn mới sẽ chưa được nhận cho tới khi có dòng ✅ Xong.'

/** How often the "⏳" line is rewritten with the time spent so far. */
export const STATUS_TICK_MS = 60 * 1000

/** The "⏳" line after `now - lock.since`, in whole minutes once past the first. */
export function formatWorking(lock: TurnLock, now = Date.now()): string {
  const minutes = Math.floor((now - lock.since) / 60_000)
  if (minutes < 1) return STATUS_WORKING
  return `⏳ Đang xử lý… (đã ${minutes} phút). Tin nhắn mới sẽ chưa được nhận cho tới khi có dòng ✅ Xong.`
}

export function formatBusyNotice(lock: TurnLock, timeZone?: string): string {
  return (
    `⏳ Claude vẫn đang xử lý yêu cầu trước (từ ${clock(lock.since, timeZone)}). ` +
    'Tin vừa rồi chưa được chuyển cho Claude. Chờ dòng ✅ Xong rồi nhắn lại.'
  )
}

export function formatDone(lock: TurnLock, now = Date.now()): string {
  const seconds = Math.max(1, Math.round((now - lock.since) / 1000))
  const took = seconds < 60 ? `${seconds} giây` : `${Math.floor(seconds / 60)} phút ${seconds % 60} giây`
  return `✅ Xong (${took}). Có thể nhắn yêu cầu tiếp theo.`
}

export const STATUS_TIMED_OUT =
  '⚠️ Không thấy Claude báo xong sau 15 phút. Bot nhận tin mới trở lại; nếu chưa có câu trả lời, nhắn lại yêu cầu.'
