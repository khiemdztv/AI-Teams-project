/**
 * Watching Outlook during an outage (fork addition, 13/11): who is watched,
 * when they are told, and what the probe reads.
 */

import { describe, expect, test, beforeEach } from 'bun:test'
import { mkdtempSync, writeFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import {
  addWatch,
  stopWatch,
  readWatches,
  checkWatches,
  probeOutlook,
  readyTasks,
  takeTask,
  taskPrompt,
  WATCH_FILE,
  WATCH_GIVE_UP_MS,
  TASK_WAIT_MS,
  MAX_TASK_CHARS,
  type Probe,
} from '../src/outlook-watch.js'

const KIEN = '0b7e0b0e-1111-4222-8333-944445555666'
const BOSS = '0b7e0b0e-aaaa-4bbb-8ccc-9ddddeeeefff'
const T0 = Date.parse('2026-11-13T09:05:00+07:00')
const down: Probe = { ok: false, detail: 'mail 502, calendar 502' }
const up: Probe = { ok: true, detail: 'mail 200, calendar 200' }

let dir: string
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'outlook-watch-'))
})

describe('watches', () => {
  test('one per person, kept on disk; asking again starts over; stop removes it', () => {
    addWatch(dir, KIEN.toUpperCase(), 'a:kien', T0)
    addWatch(dir, BOSS, 'a:boss', T0)
    addWatch(dir, KIEN, 'a:kien', T0 + 60_000)
    expect(readWatches(dir)).toEqual({
      [KIEN]: { conversationId: 'a:kien', since: T0 + 60_000, okRounds: 0 },
      [BOSS]: { conversationId: 'a:boss', since: T0, okRounds: 0 },
    })
    expect(stopWatch(dir, BOSS)).toBe(true)
    expect(stopWatch(dir, BOSS)).toBe(false)
    expect(Object.keys(readWatches(dir))).toEqual([KIEN])
  })

  test('no sender, no watch; a broken file is no watches', () => {
    expect(() => addWatch(dir, '', 'a:x')).toThrow(/Teams sender/)
    expect(() => addWatch(dir, '../access', 'a:x')).toThrow(/Teams sender/)
    writeFileSync(join(dir, WATCH_FILE), '{oops')
    expect(readWatches(dir)).toEqual({})
  })
})

describe('checking', () => {
  test('told only after two good rounds in a row, then the watch ends', async () => {
    addWatch(dir, KIEN, 'a:kien', T0)
    const answers = [down, up, down, up, up]
    const probe = async () => answers.shift()!
    for (let i = 1; i <= 4; i++) expect(await checkWatches(dir, T0 + i * 180_000, probe)).toEqual([])
    expect(readWatches(dir)[KIEN]).toMatchObject({ okRounds: 1, last: 'mail 200, calendar 200' })
    const notices = await checkWatches(dir, Date.parse('2026-11-13T09:42:00+07:00'), probe)
    expect(notices).toEqual([
      { senderId: KIEN, conversationId: 'a:kien', text: expect.stringContaining('Outlook đã dùng lại được (kiểm tra lúc 09:42)') },
    ])
    expect(readWatches(dir)).toEqual({})
  })

  test('gives up after 12 hours, without probing again', async () => {
    addWatch(dir, KIEN, 'a:kien', T0)
    let probed = 0
    const notices = await checkWatches(dir, T0 + WATCH_GIVE_UP_MS, async () => (probed++, down))
    expect(probed).toBe(0)
    expect(notices[0]!.text).toContain('sau 12 giờ theo dõi')
    expect(readWatches(dir)).toEqual({})
  })

  test('a probe that throws counts as down; a watch restarted meanwhile is kept as restarted', async () => {
    addWatch(dir, KIEN, 'a:kien', T0)
    const notices = await checkWatches(dir, T0 + 180_000, async () => {
      addWatch(dir, KIEN, 'a:kien-again', T0 + 120_000)
      throw new Error('not signed in')
    })
    expect(notices).toEqual([])
    expect(readWatches(dir)[KIEN]).toEqual({ conversationId: 'a:kien-again', since: T0 + 120_000, okRounds: 0 })
  })

  test('the notices say nothing technical', async () => {
    addWatch(dir, BOSS, 'a:boss', T0)
    await checkWatches(dir, T0 + 1, async () => up)
    const [notice] = await checkWatches(dir, T0 + 2, async () => up)
    expect(notice!.text).not.toMatch(/Graph|502|200|token|m365/i)
  })
})

describe('a task to do once Outlook is back (13/11)', () => {
  const TASK = 'Nhắn Teams cho Jordan (jordan@fabrikam.com): "Chị ơi, hợp đồng em gửi chiều nay nhé."'

  test('kept with the watch; too long is refused', () => {
    addWatch(dir, KIEN, 'a:kien', T0, `  ${TASK}  `)
    expect(readWatches(dir)[KIEN]!.then).toBe(TASK)
    expect(() => addWatch(dir, KIEN, 'a:kien', T0, 'x'.repeat(MAX_TASK_CHARS + 1))).toThrow(/longer than/)
  })

  test('once back, the person is told, the task waits for the bot, and is taken exactly once', async () => {
    addWatch(dir, KIEN, 'a:kien', T0, TASK)
    addWatch(dir, BOSS, 'a:boss', T0)
    await checkWatches(dir, T0 + 180_000, async () => up)
    const notices = await checkWatches(dir, T0 + 360_000, async () => up)
    expect(notices.find(n => n.senderId === KIEN)!.text).toContain('Bắt đầu làm việc đã dặn: “Nhắn Teams cho Jordan')
    expect(notices.find(n => n.senderId === BOSS)!.text).toContain('Outlook đã dùng lại được')
    // The plain watch is done; the one with a task waits, and is not probed again.
    expect(Object.keys(readWatches(dir))).toEqual([KIEN])
    let probed = 0
    expect(await checkWatches(dir, T0 + 540_000, async () => (probed++, down))).toEqual([])
    expect(probed).toBe(0)

    const [task] = readyTasks(dir)
    expect(task).toEqual({ senderId: KIEN, conversationId: 'a:kien', since: T0, then: TASK })
    expect(takeTask(dir, task!)).toBe(true)
    expect(takeTask(dir, task!)).toBe(false)
    expect(readyTasks(dir)).toEqual([])
  })

  test('a watch started over meanwhile is not taken as the old task', async () => {
    addWatch(dir, KIEN, 'a:kien', T0, TASK)
    await checkWatches(dir, T0 + 1, async () => up)
    await checkWatches(dir, T0 + 2, async () => up)
    const [task] = readyTasks(dir)
    addWatch(dir, KIEN, 'a:kien', T0 + 3, 'việc khác')
    expect(takeTask(dir, task!)).toBe(false)
    expect(readWatches(dir)[KIEN]!.then).toBe('việc khác')
  })

  test('a task the bot never gets to, or a watch that never comes back, is reported and dropped', async () => {
    addWatch(dir, KIEN, 'a:kien', T0, TASK)
    await checkWatches(dir, T0 + 1, async () => up)
    await checkWatches(dir, T0 + 2, async () => up)
    const [dropped] = await checkWatches(dir, T0 + 2 + TASK_WAIT_MS, async () => up)
    expect(dropped!.text).toContain('bot bận suốt 2 giờ nên chưa làm việc đã dặn')
    expect(readWatches(dir)).toEqual({})

    addWatch(dir, BOSS, 'a:boss', T0, TASK)
    const [gaveUp] = await checkWatches(dir, T0 + WATCH_GIVE_UP_MS, async () => down)
    expect(gaveUp!.text).toContain('Việc đã dặn chưa làm: “Nhắn Teams cho Jordan')
  })

  test('the turn that does it says when it was asked, to redo the draft, and to do nothing else', () => {
    const prompt = taskPrompt({ senderId: KIEN, conversationId: 'a:kien', since: T0, then: TASK }, T0 + 3600_000)
    expect(prompt).toContain('việc dặn lúc 09:05 kèm /force')
    expect(prompt).toContain('Outlook dùng lại được từ 10:05')
    expect(prompt).toContain('soạn lại rồi gửi')
    expect(prompt).toContain('Đừng làm thêm việc nào khác')
    expect(prompt.endsWith(`Việc đã dặn: ${TASK}`)).toBe(true)
  })
})

describe('the probe', () => {
  test('reads only status codes of the inbox and calendar, with the person\'s token', async () => {
    const seen: string[] = []
    const fetchImpl = (async (url: string, init: any) => {
      seen.push(`${init.headers.Authorization.slice(0, 7)} ${String(url).replace(/DateTime=[^&]+/g, 'DateTime=…')}`)
      return new Response('{}', { status: String(url).includes('calendarView') ? 503 : 200 })
    }) as unknown as typeof fetch
    expect(await probeOutlook(async () => 'tok', fetchImpl, T0)).toEqual({ ok: false, detail: 'mail 200, calendar 503' })
    expect(seen.sort()).toEqual([
      'Bearer  https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=…&endDateTime=…&$top=1&$select=id',
      'Bearer  https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=1&$select=id',
    ])
  })

  test('no answer and a failed sign-in are reported, not thrown', async () => {
    const offline = (async () => {
      throw new Error('ECONNRESET')
    }) as unknown as typeof fetch
    expect(await probeOutlook(async () => 'tok', offline, T0)).toEqual({ ok: false, detail: 'mail no answer, calendar no answer' })
    const noSignIn = await probeOutlook(async () => {
      throw new Error('This person has not signed in')
    }, offline, T0)
    expect(noSignIn).toEqual({ ok: false, detail: 'sign-in failed: This person has not signed in' })
  })
})
