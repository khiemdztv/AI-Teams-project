/**
 * /useransw: an admin sees the bot as a non-admin would (fork addition, 02/11).
 *
 * Typing "/useransw" alone in a 1:1 chat turns it on for that admin,
 * "/useransw off" turns it off; "/useransw" inside a sentence applies to that
 * message only, like a trigger. While on, their messages reach Claude as role="user" and the
 * status lines use the assistant wording, so they can check what the boss will
 * get without editing access.json and restarting. It only ever lowers a role:
 * a non-admin typing the command is just chatting.
 *
 * Kept in view-as-user.json in the state dir, so it survives a restart.
 */

import { readFileSync, writeFileSync, renameSync, mkdirSync } from 'fs'
import { join } from 'path'
import type { Role } from './gate.js'

export const VIEW_AS_FILE = 'view-as-user.json'

export type ViewAsCommand = 'on' | 'off' | 'toggle'

/** "/useransw", "/useransw on", "/useransw off" (any case, surrounding space); anything else is chat. */
export function parseViewAsCommand(text: string): ViewAsCommand | undefined {
  const m = /^\s*\/useransw(?:\s+(on|off|bật|tắt))?\s*$/i.exec(String(text ?? ''))
  if (!m) return undefined
  const arg = m[1]?.toLowerCase()
  if (!arg) return 'toggle'
  return arg === 'on' || arg === 'bật' ? 'on' : 'off'
}

export function readViewAs(stateDir: string): Set<string> {
  try {
    const parsed = JSON.parse(readFileSync(join(stateDir, VIEW_AS_FILE), 'utf8'))
    return new Set(Array.isArray(parsed) ? parsed.map(String) : [])
  } catch {
    return new Set()
  }
}

/** Apply `cmd` for `senderId`; returns whether view-as-user is now on for them. */
export function applyViewAs(stateDir: string, senderId: string, cmd: ViewAsCommand): boolean {
  const ids = readViewAs(stateDir)
  const on = cmd === 'toggle' ? !ids.has(senderId) : cmd === 'on'
  if (on) ids.add(senderId)
  else ids.delete(senderId)
  mkdirSync(stateDir, { recursive: true, mode: 0o700 })
  const target = join(stateDir, VIEW_AS_FILE)
  writeFileSync(`${target}.tmp`, JSON.stringify([...ids]) + '\n', { mode: 0o600 })
  renameSync(`${target}.tmp`, target)
  return on
}

const TRIGGER = /(^|\s)\/useransw(?=\s|$)/gi
/** The trigger with the spaces after it, so removing it leaves no gap at a line start. */
const TRIGGER_WITH_SPACE = /(^|\s)\/useransw(?:[ \t]+|(?=\s)|$)/gi

/**
 * "/useransw" inside a message with other text: that one message is answered
 * as a user, whatever the on/off setting. Returns the text with the trigger
 * removed, or undefined when there is no trigger. A message that is only the
 * command is parseViewAsCommand's, so check that first.
 */
export function stripViewAsTrigger(text: string): string | undefined {
  const raw = String(text ?? '')
  if (!raw.match(TRIGGER)) return undefined
  return raw.replace(TRIGGER_WITH_SPACE, '$1').trim()
}

/** The role a message is handled with: an admin viewing as a user is a user. */
export function effectiveRole(role: Role, senderId: string, viewAs: Set<string>): Role {
  return role === 'admin' && viewAs.has(senderId) ? 'user' : role
}

export function formatViewAsReply(on: boolean): string {
  return on
    ? 'Đã bật /useransw: từ giờ bot trả lời bạn như trả lời sếp (không chi tiết kỹ thuật, dòng trạng thái kiểu trợ lý). Tin 🔐 trong lượt của bạn vẫn tới các admin, kể cả bạn. Gõ "/useransw off" để tắt.'
    : 'Đã tắt /useransw: bot trả lời bạn như admin trở lại. Muốn một tin được trả lời như sếp thì thêm "/useransw" vào trong câu.'
}
