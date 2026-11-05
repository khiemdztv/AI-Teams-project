/**
 * Fork addition (30/10): the sender's calendar through Graph, and uploading a
 * file to their OneDrive. Graph is faked; nothing here needs a tenant.
 */

import { describe, expect, test, beforeEach } from 'bun:test'
import { mkdtempSync, mkdirSync, writeFileSync, readFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { Calendar, parseLocalTime, VN_ZONE, MAX_INVITES_PER_HOUR } from '../m365/calendar.js'
import { Files, UPLOAD_FOLDER } from '../m365/files.js'

type Call = { method: string; url: string; body: any; headers: Headers }

/** A fake Graph: routes by "METHOD path-prefix" (or a full URL), records every call. */
function fakeGraph(routes: Record<string, (call: Call) => { status?: number; json?: unknown }>) {
  const calls: Call[] = []
  const fetchImpl = async (url: string, init?: RequestInit) => {
    const call: Call = {
      method: String(init?.method ?? 'GET'),
      url,
      body: typeof init?.body === 'string' ? JSON.parse(init.body) : init?.body,
      headers: new Headers(init?.headers),
    }
    calls.push(call)
    const path = url.replace('https://graph.microsoft.com/v1.0', '')
    const key = Object.keys(routes)
      .filter(k => {
        const [m, p] = k.split(' ')
        return m === call.method && path.startsWith(p!)
      })
      .sort((a, b) => b.length - a.length)[0]
    if (!key) return new Response(JSON.stringify({ error: { code: 'NotFound', message: 'no route' } }), { status: 404 })
    const out = routes[key]!(call)
    return new Response(JSON.stringify(out.json ?? {}), { status: out.status ?? 200 })
  }
  return { calls, fetchImpl }
}

// 08:00 on 30/10/2026 in Vietnam.
const NOW = Date.parse('2026-10-30T01:00:00Z')

let dir: string
let exportDir: string
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'm365-cal-'))
  exportDir = join(dir, 'bot-export')
  mkdirSync(exportDir)
})

function calendar(fetchImpl: any) {
  return new Calendar({
    token: async () => 'ACCESS',
    internalDomains: ['fabrikam.com'],
    sendLog: join(dir, 'sent.jsonl'),
    fetchImpl,
    now: () => NOW,
  })
}

const EVENT = {
  id: 'AAMkAGI2THVSAAA=',
  subject: 'Họp vụ việc ABC',
  start: { dateTime: '2026-11-02T09:00:00.0000000', timeZone: VN_ZONE },
  end: { dateTime: '2026-11-02T10:00:00.0000000', timeZone: VN_ZONE },
  attendees: [{ emailAddress: { address: 'sam@fabrikam.com', name: 'Sam' }, status: { response: 'accepted' } }],
  onlineMeeting: { joinUrl: 'https://teams.microsoft.com/l/meetup-join/x' },
  responseStatus: { response: 'organizer' },
  showAs: 'busy',
}

describe('times', () => {
  test('Vietnam local times are read strictly', () => {
    expect(parseLocalTime('2026-11-02 9:30')).toBe('2026-11-02T09:30:00')
    expect(parseLocalTime('2026-11-02T14:05')).toBe('2026-11-02T14:05:00')
    for (const bad of ['2/11 9:30', '2026-11-02', '2026-11-02 25:00', 'tomorrow']) expect(() => parseLocalTime(bad)).toThrow(/not a time/)
  })
})

describe('reading the calendar', () => {
  test('lists events in Vietnam time with short ids and Teams links', async () => {
    const g = fakeGraph({ 'GET /me/calendarView': () => ({ json: { value: [EVENT] } }) })
    const out = await calendar(g.fetchImpl).listEvents({ since: '2026-11-02' })
    expect(out).toContain('2026-11-02 09:00–10:00 · Họp vụ việc ABC')
    expect(out).toContain('Sam (accepted)')
    expect(out).toContain('Teams link: https://teams.microsoft.com/l/meetup-join/x')
    expect(out).toMatch(/event_id: e_[0-9a-f]{10}/)
    expect(out).not.toContain(EVENT.id)
    expect(g.calls[0]!.headers.get('prefer')).toContain(VN_ZONE)
    expect(g.calls[0]!.url).toContain(encodeURIComponent('2026-11-01T17:00:00.000Z'))
  })

  test('an empty range is named in Vietnam dates', async () => {
    const g = fakeGraph({ 'GET /me/calendarView': () => ({ json: { value: [] } }) })
    expect(await calendar(g.fetchImpl).listEvents({ since: 'today', until: '2026-11-06' })).toBe(
      'No events from 2026-10-30 to 2026-11-06 (Vietnam time).',
    )
  })

  test('free times ask Graph for every attendee as required', async () => {
    const g = fakeGraph({
      'POST /me/findMeetingTimes': () => ({
        json: {
          meetingTimeSuggestions: [
            { meetingTimeSlot: { start: { dateTime: '2026-11-02T14:00:00.0000000' }, end: { dateTime: '2026-11-02T14:30:00.0000000' } } },
          ],
        },
      }),
    })
    const out = await calendar(g.fetchImpl).findTime({ attendees: ['Sam@fabrikam.com'], duration_minutes: 30 })
    expect(out).toContain('2026-11-02 14:00–14:30')
    expect(g.calls[0]!.body.attendees).toEqual([{ type: 'required', emailAddress: { address: 'sam@fabrikam.com' } }])
    expect(g.calls[0]!.body.meetingDuration).toBe('PT30M')
  })
})

describe('setting up a meeting', () => {
  const ask = {
    subject: 'Họp vụ việc ABC',
    start: '2026-11-02 09:00',
    attendees: ['sam@fabrikam.com', 'client@example.com'],
  }

  test('prepare creates and sends nothing, and flags external attendees', async () => {
    const g = fakeGraph({})
    const out = await calendar(g.fetchImpl).prepareEvent(ask)
    expect(g.calls).toHaveLength(0)
    expect(out).toContain('Nothing has been created or sent yet')
    expect(out).toContain('2026-11-02 09:00–10:00')
    expect(out).toContain('EXTERNAL attendees (outside the firm): client@example.com')
    expect(out).toMatch(/plan_id: p_[0-9a-f]{10}/)
  })

  test('create needs the attendees and start repeated exactly', async () => {
    const g = fakeGraph({ 'POST /me/events': () => ({ status: 201, json: EVENT }) })
    const cal = calendar(g.fetchImpl)
    const plan = /plan_id: (\S+)/.exec(await cal.prepareEvent(ask))![1]!
    await expect(cal.createEvent({ plan_id: plan, expect_attendees: ['sam@fabrikam.com'], expect_start: '2026-11-02 09:00' })).rejects.toThrow(/refused/)
    await expect(cal.createEvent({ plan_id: plan, expect_attendees: ask.attendees, expect_start: '2026-11-02 10:00' })).rejects.toThrow(/refused/)
    expect(g.calls).toHaveLength(0)

    const out = await cal.createEvent({ plan_id: plan, expect_attendees: [...ask.attendees].reverse(), expect_start: '2026-11-02 9:00' })
    expect(out).toContain('invitations sent')
    const body = g.calls[0]!.body
    expect(body.start).toEqual({ dateTime: '2026-11-02T09:00:00', timeZone: VN_ZONE })
    expect(body.end).toEqual({ dateTime: '2026-11-02T10:00:00', timeZone: VN_ZONE })
    expect(body.isOnlineMeeting).toBe(true)
    expect(body.attendees).toHaveLength(2)
    expect(JSON.parse(readFileSync(join(dir, 'sent.jsonl'), 'utf8')).kind).toBe('event')

    // A plan is used once.
    await expect(cal.createEvent({ plan_id: plan, expect_attendees: ask.attendees, expect_start: '2026-11-02 09:00' })).rejects.toThrow(/unknown or expired/)
  })

  test('a meeting in the past is refused', async () => {
    await expect(calendar(fakeGraph({}).fetchImpl).prepareEvent({ ...ask, start: '2026-10-29 09:00' })).rejects.toThrow(/past/)
  })

  test('invitations are rate limited', async () => {
    const lines = Array.from({ length: MAX_INVITES_PER_HOUR }, () =>
      JSON.stringify({ at: new Date(NOW - 60_000).toISOString(), kind: 'event', to: ['a@fabrikam.com'] }),
    )
    writeFileSync(join(dir, 'sent.jsonl'), lines.join('\n') + '\n')
    const g = fakeGraph({ 'POST /me/events': () => ({ status: 201, json: EVENT }) })
    const cal = calendar(g.fetchImpl)
    const plan = /plan_id: (\S+)/.exec(await cal.prepareEvent(ask))![1]!
    await expect(cal.createEvent({ plan_id: plan, expect_attendees: ask.attendees, expect_start: ask.start })).rejects.toThrow(/last hour/)
    expect(g.calls).toHaveLength(0)
  })
})

describe('uploading to the sender\'s OneDrive', () => {
  function files(fetchImpl: any) {
    return new Files({ token: async () => 'ACCESS', uploadRoots: [exportDir], fetchImpl })
  }

  test('a small file goes up in one request, renamed rather than overwritten, and the link comes back', async () => {
    const path = join(exportDir, 'Hợp đồng.docx')
    writeFileSync(path, 'docx-bytes')
    const g = fakeGraph({
      'PUT /me/drive/root:': () => ({ status: 201, json: { name: 'Hợp đồng.docx', size: 10, webUrl: 'https://fabrikam-my.sharepoint.com/personal/x/Hop.docx' } }),
    })
    const out = await files(g.fetchImpl).upload({ path })
    expect(out).toContain('link: https://fabrikam-my.sharepoint.com/personal/x/Hop.docx')
    expect(g.calls[0]!.url).toContain(`/me/drive/root:/${encodeURIComponent(UPLOAD_FOLDER)}/${encodeURIComponent('Hợp đồng.docx')}:/content`)
    expect(g.calls[0]!.url).toContain('conflictBehavior=rename')
  })

  test('a large file goes up through an upload session, without the bearer token', async () => {
    const path = join(exportDir, 'big.pdf')
    writeFileSync(path, Buffer.alloc(5 * 1024 * 1024, 1))
    const g = fakeGraph({
      'POST /me/drive/root:': () => ({ json: { uploadUrl: 'https://upload.example.sharepoint.com/session' } }),
      'PUT https://upload.example.sharepoint.com/session': () => ({ status: 201, json: { name: 'big.pdf', webUrl: 'https://fabrikam-my.sharepoint.com/big.pdf' } }),
    })
    const out = await files(g.fetchImpl).upload({ path })
    expect(out).toContain('link: https://fabrikam-my.sharepoint.com/big.pdf')
    const puts = g.calls.filter(c => c.method === 'PUT')
    expect(puts.length).toBeGreaterThan(1)
    expect(puts.every(c => c.headers.get('authorization') === null)).toBe(true)
  })

  test('only files inside the export folder can be uploaded', async () => {
    const outside = join(dir, 'secret.txt')
    writeFileSync(outside, 'x')
    const g = fakeGraph({})
    await expect(files(g.fetchImpl).upload({ path: outside })).rejects.toThrow()
    expect(g.calls).toHaveLength(0)
  })
})

describe('cancelling a meeting', () => {
  const MINE = { ...EVENT, isOrganizer: true }
  const ask = { expect_subject: 'Họp vụ việc ABC', expect_start: '2026-11-02 09:00' }

  async function listed(routes: Record<string, (call: Call) => { status?: number; json?: unknown }>) {
    const g = fakeGraph({ 'GET /me/calendarView': () => ({ json: { value: [EVENT] } }), ...routes })
    const cal = calendar(g.fetchImpl)
    const eventId = /event_id: (\S+)/.exec(await cal.listEvents({ since: '2026-11-02' }))![1]!
    return { g, cal, eventId }
  }

  test('the organizer cancels, attendees are notified, and it is logged', async () => {
    const { g, cal, eventId } = await listed({
      'GET /me/events/': () => ({ json: MINE }),
      'POST /me/events/': () => ({ status: 202 }),
    })
    const out = await cal.cancelEvent({ event_id: eventId, ...ask, comment: 'Dời lịch' })
    expect(out).toContain('sam@fabrikam.com were sent a cancellation')
    const post = g.calls.find(c => c.method === 'POST')!
    expect(post.url).toEndWith(`/me/events/${encodeURIComponent(EVENT.id)}/cancel`)
    expect(post.body).toEqual({ comment: 'Dời lịch' })
    const log = JSON.parse(readFileSync(join(dir, 'sent.jsonl'), 'utf8'))
    expect(log.kind).toBe('cancel')
    expect(log.to).toEqual(['sam@fabrikam.com'])
  })

  test('an entry with nobody invited is deleted, not cancelled', async () => {
    const { g, cal, eventId } = await listed({
      'GET /me/events/': () => ({ json: { ...MINE, attendees: [] } }),
      'DELETE /me/events/': () => ({ status: 204 }),
    })
    expect(await cal.cancelEvent({ event_id: eventId, ...ask })).toContain('nobody was invited')
    expect(g.calls.some(c => c.method === 'DELETE')).toBe(true)
    expect(g.calls.some(c => c.method === 'POST')).toBe(false)
  })

  test('subject and start must repeat the event exactly', async () => {
    const { g, cal, eventId } = await listed({ 'GET /me/events/': () => ({ json: MINE }) })
    await expect(cal.cancelEvent({ event_id: eventId, ...ask, expect_start: '2026-11-02 10:00' })).rejects.toThrow(/refused/)
    await expect(cal.cancelEvent({ event_id: eventId, ...ask, expect_subject: 'Họp khác' })).rejects.toThrow(/refused/)
    expect(g.calls.filter(c => c.method !== 'GET')).toHaveLength(0)
  })

  test("someone else's meeting cannot be cancelled", async () => {
    const theirs = { ...EVENT, isOrganizer: false, organizer: { emailAddress: { address: 'kien.do@fabrikam.com' } } }
    const { g, cal, eventId } = await listed({ 'GET /me/events/': () => ({ json: theirs }) })
    await expect(cal.cancelEvent({ event_id: eventId, ...ask })).rejects.toThrow(/only the organizer \(kien\.do@fabrikam\.com\)/)
    expect(g.calls.filter(c => c.method !== 'GET')).toHaveLength(0)
  })

  test('an unknown event_id is refused before Graph is asked', async () => {
    const g = fakeGraph({})
    await expect(calendar(g.fetchImpl).cancelEvent({ event_id: 'e_0000000000', ...ask })).rejects.toThrow(/unknown event_id/)
    expect(g.calls).toHaveLength(0)
  })
})
