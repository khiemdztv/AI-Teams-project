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
  WATCH_FILE,
  WATCH_GIVE_UP_MS,
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
