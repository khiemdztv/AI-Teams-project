/**
 * Technical log kept out of the chat (fork addition, 02/11).
 *
 * A non-admin is answered as an assistant would answer a busy executive: no
 * tool names, paths, error text or approval codes. What would have gone into
 * the chat goes here instead — Claude's own notes through the log_internal
 * tool, and every permission prompt raised on a non-admin's behalf with who
 * answered it. The admins read it on the bot machine:
 *
 *   tail -f ~/.claude/channels/msteams/logs/assistant.jsonl
 *
 * One JSON object per line, 0600 like everything else in the state dir. It
 * can hold client matter (a file path, a subject line), so it never leaves the
 * machine and is never sent to Teams.
 */

import { appendFileSync, mkdirSync } from 'fs'
import { join } from 'path'

export const ASSISTANT_LOG = join('logs', 'assistant.jsonl')
export const MAX_LOG_TEXT = 4000

export type AssistantLogEntry = {
  kind: 'note' | 'permission_request' | 'permission_verdict' | 'permission_expired'
  /** Entra object id of the person whose message Claude was working on. */
  sender?: string
  senderName?: string
  text: string
  [key: string]: unknown
}

export function assistantLogPath(stateDir: string): string {
  return join(stateDir, ASSISTANT_LOG)
}

/**
 * Append one entry. Control characters other than newline and tab are dropped,
 * so `tail` cannot be made to redraw the terminal; JSON escapes the newlines,
 * so one entry stays one line.
 */
export function appendAssistantLog(stateDir: string, entry: AssistantLogEntry, now = Date.now()): void {
  const text = String(entry.text ?? '')
    .replace(/[\x00-\x08\x0b-\x1f\x7f]/g, '')
    .slice(0, MAX_LOG_TEXT)
  mkdirSync(join(stateDir, 'logs'), { recursive: true, mode: 0o700 })
  appendFileSync(assistantLogPath(stateDir), JSON.stringify({ at: new Date(now).toISOString(), ...entry, text }) + '\n', {
    mode: 0o600,
  })
}
