/**
 * Watching Outlook during a Microsoft outage (fork addition, 13/11).
 *
 * A person asks the bot to say when Outlook works again. A cron job inside
 * the Claude session cannot do it: the turn it starts carries no Teams sender,
 * so the m365 server (rightly) uses no mailbox. Instead the msteams server,
 * which already knows the sender of the turn that asked, records a watch here
 * and checks it on a timer with that person's own sign-in, no Claude turn
 * needed. When mail and calendar both answer twice in a row, it messages them
 * in the chat they asked from; after WATCH_GIVE_UP_MS it says it stopped.
 *
 * outlook-watch.json in the state dir, by sender object id:
 *   { "<object id>": { "conversationId": "...", "since": 0, "okRounds": 0 } }
 * One watch per person; asking again starts it over.
 *
 * A watch may carry a task ("then", 13/11): what the person asked to have done
 * once Outlook is back, such as sending a message already drafted. Only from a
 * message that carried /force, since nobody will confirm it at that time.
 * When Outlook is back the watch becomes `ready`; the server starts a turn as
 * that person (so the m365 server knows whose mailbox) as soon as the bot is
 * free, and drops the task if the bot stays busy for TASK_WAIT_MS.
 */

import { readFileSync, writeFileSync, renameSync } from 'fs'
import { join } from 'path'

export const WATCH_FILE = 'outlook-watch.json'
export const WATCH_EVERY_MS = 3 * 60_000
export const WATCH_GIVE_UP_MS = 12 * 60 * 60_000
/** Rounds in a row with mail and calendar answering before it counts as back. */
export const OK_ROUNDS_NEEDED = 2
/** How long a task waits for the bot to be free once Outlook is back. */
export const TASK_WAIT_MS = 2 * 60 * 60_000
export const MAX_TASK_CHARS = 1000
const PROBE_TIMEOUT_MS = 30_000
const OBJECT_ID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/

export type Watch = {
  conversationId: string
  since: number
  okRounds: number
  lastCheck?: number
  last?: string
  /** The task to do once Outlook is back, as the person asked for it (with /force). */
  then?: string
  /** When Outlook came back, for a watch whose task has not started yet. */
  ready?: number
}
export type Watches = Record<string, Watch>
export type Probe = { ok: boolean; detail: string }
export type Notice = { senderId: string; conversationId: string; text: string }
export type Task = { senderId: string; conversationId: string; since: number; then: string }

export function readWatches(stateDir: string): Watches {
  try {
    const parsed = JSON.parse(readFileSync(join(stateDir, WATCH_FILE), 'utf8'))
    if (!parsed || typeof parsed !== 'object') return {}
    const out: Watches = {}
    for (const [id, w] of Object.entries(parsed as Record<string, any>)) {
      if (OBJECT_ID.test(id) && typeof w?.conversationId === 'string' && typeof w?.since === 'number') {
        out[id] = { ...w, okRounds: Number(w.okRounds) || 0 }
      }
    }
    return out
  } catch {
    return {}
  }
}

function writeWatches(stateDir: string, watches: Watches): void {
  const target = join(stateDir, WATCH_FILE)
  writeFileSync(`${target}.tmp`, JSON.stringify(watches, null, 2), { mode: 0o600 })
  renameSync(`${target}.tmp`, target)
}

function checkSender(senderId: string): string {
  const id = String(senderId ?? '').toLowerCase()
  if (!OBJECT_ID.test(id)) throw new Error('a watch needs the Teams sender of the current message')
  return id
}

/** Start (or start over) the watch for one person, reported in one conversation. */
export function addWatch(
  stateDir: string,
  senderId: string,
  conversationId: string,
  now = Date.now(),
  then?: string,
): Watch {
  const id = checkSender(senderId)
  if (!conversationId) throw new Error('a watch needs the conversation to report in')
  const task = String(then ?? '').trim()
  if (task.length > MAX_TASK_CHARS) throw new Error(`the task is longer than ${MAX_TASK_CHARS} characters`)
  const watch: Watch = { conversationId, since: now, okRounds: 0, ...(task ? { then: task } : {}) }
  writeWatches(stateDir, { ...readWatches(stateDir), [id]: watch })
  return watch
}

/** Tasks whose Outlook is back, oldest first, waiting for the bot to be free. */
export function readyTasks(stateDir: string): Task[] {
  return Object.entries(readWatches(stateDir))
    .filter(([, w]) => w.ready !== undefined && w.then)
    .sort(([, a], [, b]) => a.ready! - b.ready!)
    .map(([senderId, w]) => ({ senderId, conversationId: w.conversationId, since: w.since, then: w.then! }))
}

/**
 * Remove a task as it starts, before Claude is told, so a crash in between
 * never runs it twice. False when the watch was replaced or stopped meanwhile.
 */
export function takeTask(stateDir: string, task: Task): boolean {
  const watches = readWatches(stateDir)
  const w = watches[task.senderId]
  if (!w || w.since !== task.since || w.ready === undefined) return false
  delete watches[task.senderId]
  writeWatches(stateDir, watches)
  return true
}

/** The fixed text of the turn that does the task; the task itself is the person's request. */
export function taskPrompt(task: Task, now: number): string {
  return [
    `[Outlook đã chạy lại, việc dặn lúc ${vnHHMM(task.since)} kèm /force]`,
    `Lúc ${vnHHMM(task.since)} người này nhắn: khi Outlook chạy lại thì làm việc dưới đây, gửi luôn không cần hỏi lại. Outlook dùng lại được từ ${vnHHMM(now)}.`,
    'Làm đúng việc đó, một lần, rồi reply báo đã làm gì, gửi cho ai. Bản soạn cũ (plan_id, draft_id) đã hết hạn: soạn lại rồi gửi. Không rõ gửi cho ai hay gửi gì thì hỏi lại, đừng đoán. Đừng làm thêm việc nào khác.',
    `Việc đã dặn: ${task.then}`,
  ].join('\n')
}

export function stopWatch(stateDir: string, senderId: string): boolean {
  const id = checkSender(senderId)
  const watches = readWatches(stateDir)
  if (!watches[id]) return false
  delete watches[id]
  writeWatches(stateDir, watches)
  return true
}

/**
 * One look at the person's mail and calendar through Graph. Only status codes
 * come back: nothing read from the mailbox is kept or shown.
 */
export async function probeOutlook(
  token: () => Promise<string>,
  fetchImpl: typeof fetch = fetch,
  now = Date.now(),
): Promise<Probe> {
  let bearer: string
  try {
    bearer = await token()
  } catch (err) {
    // Sign-in itself failing is not an outage we can wait out; say why.
    return { ok: false, detail: `sign-in failed: ${err instanceof Error ? err.message.slice(0, 120) : 'error'}` }
  }
  const get = async (path: string): Promise<number> => {
    try {
      const res = await fetchImpl(`https://graph.microsoft.com/v1.0${path}`, {
        headers: { Authorization: `Bearer ${bearer}` },
        signal: AbortSignal.timeout(PROBE_TIMEOUT_MS),
      })
      await res.arrayBuffer().catch(() => undefined)
      return res.status
    } catch {
      return 0
    }
  }
  const from = new Date(now).toISOString()
  const to = new Date(now + 86_400_000).toISOString()
  const [mail, calendar] = await Promise.all([
    get('/me/mailFolders/inbox/messages?$top=1&$select=id'),
    get(`/me/calendarView?startDateTime=${from}&endDateTime=${to}&$top=1&$select=id`),
  ])
  const code = (n: number) => (n === 0 ? 'no answer' : String(n))
  return { ok: mail === 200 && calendar === 200, detail: `mail ${code(mail)}, calendar ${code(calendar)}` }
}

/** "09:42", Vietnam time. */
export function vnHHMM(now: number): string {
  return new Date(now + 7 * 3600_000).toISOString().slice(11, 16)
}

// Plain words only: the person may have role "user", who sees no technical detail.
export function formatBack(now: number): string {
  return `✅ Outlook đã dùng lại được (kiểm tra lúc ${vnHHMM(now)}): đọc được cả hộp thư và lịch. Cần gì cứ nhắn như bình thường.`
}

export function formatGiveUp(now: number, then?: string): string {
  const task = then ? ` Việc đã dặn chưa làm: ${shortTask(then)}.` : ''
  return `Outlook vẫn chưa dùng lại được sau 12 giờ theo dõi (lần thử cuối lúc ${vnHHMM(now)}). Đã dừng theo dõi; nhắn lại nếu cần theo dõi tiếp.${task}`
}

function shortTask(then: string): string {
  const t = then.replace(/\s+/g, ' ').trim()
  return `“${t.length > 200 ? `${t.slice(0, 199)}…` : t}”`
}

export function formatBackWithTask(now: number, then: string): string {
  return `✅ Outlook đã dùng lại được (kiểm tra lúc ${vnHHMM(now)}). Bắt đầu làm việc đã dặn: ${shortTask(then)}`
}

export function formatTaskDropped(then: string): string {
  return `Outlook đã dùng lại được nhưng bot bận suốt 2 giờ nên chưa làm việc đã dặn: ${shortTask(then)}. Nhắn lại nếu vẫn cần.`
}

/**
 * Check every watch once and say who to message. A watch that is back, or too
 * old, is removed. A watch added or restarted while the probes ran is kept as
 * it is: the file is re-read before writing and only the entries checked here
 * (same `since`) are updated.
 */
export async function checkWatches(
  stateDir: string,
  now: number,
  probe: (senderId: string) => Promise<Probe>,
): Promise<Notice[]> {
  const checked = readWatches(stateDir)
  const notices: Notice[] = []
  const updates = new Map<string, Watch | null>()
  for (const [senderId, w] of Object.entries(checked)) {
    // Back already, task waiting for the bot: nothing to probe.
    if (w.ready !== undefined) {
      if (now - w.ready >= TASK_WAIT_MS) {
        notices.push({ senderId, conversationId: w.conversationId, text: formatTaskDropped(w.then ?? '') })
        updates.set(senderId, null)
      }
      continue
    }
    if (now - w.since >= WATCH_GIVE_UP_MS) {
      notices.push({ senderId, conversationId: w.conversationId, text: formatGiveUp(now, w.then) })
      updates.set(senderId, null)
      continue
    }
    let result: Probe
    try {
      result = await probe(senderId)
    } catch (err) {
      result = { ok: false, detail: err instanceof Error ? err.message.slice(0, 120) : 'probe failed' }
    }
    const okRounds = result.ok ? w.okRounds + 1 : 0
    if (okRounds >= OK_ROUNDS_NEEDED && w.then) {
      // Kept until the server starts the task (takeTask).
      notices.push({ senderId, conversationId: w.conversationId, text: formatBackWithTask(now, w.then) })
      updates.set(senderId, { ...w, okRounds, lastCheck: now, last: result.detail, ready: now })
    } else if (okRounds >= OK_ROUNDS_NEEDED) {
      notices.push({ senderId, conversationId: w.conversationId, text: formatBack(now) })
      updates.set(senderId, null)
    } else {
      updates.set(senderId, { ...w, okRounds, lastCheck: now, last: result.detail })
    }
  }
  const current = readWatches(stateDir)
  for (const [senderId, next] of updates) {
    if (current[senderId]?.since !== checked[senderId]?.since) continue
    if (next) current[senderId] = next
    else delete current[senderId]
  }
  if (updates.size) writeWatches(stateDir, current)
  return notices
}
