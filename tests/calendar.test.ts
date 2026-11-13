/**
 * Fork addition (30/10): the sender's calendar through Graph, and uploading a
 * file to their OneDrive. Graph is faked; nothing here needs a tenant.
 */

import { describe, expect, test, beforeEach } from 'bun:test'
import { mkdtempSync, mkdirSync, writeFileSync, readFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { Calendar, parseLocalTime, parseLocalDate, VN_ZONE, MAX_INVITES_PER_HOUR } from '../m365/calendar.js'
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

describe('entry options (10/11)', () => {
  const created = () => fakeGraph({ 'POST /me/events': () => ({ status: 201, json: EVENT }) })

  test('a day out of office: all day, shown as oof, no Teams link, no invitations', async () => {
    const g = created()
    const cal = calendar(g.fetchImpl)
    const out = await cal.prepareEvent({ subject: 'Out of Office', start: '2026-11-13', all_day: true, show_as: 'oof', body: 'Tôi nghỉ' })
    expect(out).toContain('when: 2026-11-13, all day')
    expect(out).toContain('shown as: Out of office')
    expect(out).toContain('Teams meeting link: no')
    const plan = /plan_id: (\S+)/.exec(out)![1]!
    await cal.createEvent({ plan_id: plan, expect_start: '2026-11-13' })
    const body = g.calls[0]!.body
    expect(body.isAllDay).toBe(true)
    expect(body.showAs).toBe('oof')
    expect(body.start).toEqual({ dateTime: '2026-11-13T00:00:00', timeZone: VN_ZONE })
    expect(body.end).toEqual({ dateTime: '2026-11-14T00:00:00', timeZone: VN_ZONE })
    expect(body.isOnlineMeeting).toBe(false)
    expect(body.attendees).toEqual([])
  })

  test('several days end at the midnight after the last one', async () => {
    const out = await calendar(fakeGraph({}).fetchImpl).prepareEvent({ subject: 'Nghỉ phép', start: '2026-11-13', end: '2026-11-17', all_day: true })
    expect(out).toContain('when: 2026-11-13 to 2026-11-17, all day')
    await expect(calendar(fakeGraph({}).fetchImpl).prepareEvent({ subject: 'x', start: '2026-11-13', end: '2026-11-12', all_day: true })).rejects.toThrow(/before/)
    await expect(calendar(fakeGraph({}).fetchImpl).prepareEvent({ subject: 'x', start: '2026-10-29', all_day: true })).rejects.toThrow(/past/)
  })

  test('reminder, privacy, importance and repetition reach Graph', async () => {
    const g = created()
    const cal = calendar(g.fetchImpl)
    const out = await cal.prepareEvent({
      subject: 'Giao ban',
      start: '2026-11-02 08:30',
      duration_minutes: 30,
      reminder_minutes: 0,
      sensitivity: 'private',
      importance: 'high',
      repeat: 'weekdays',
      repeat_until: '2026-12-31',
    })
    expect(out).toContain('reminder: none')
    expect(out).toContain('repeats: every weekday (Mon–Fri), until 2026-12-31')
    await cal.createEvent({ plan_id: /plan_id: (\S+)/.exec(out)![1]!, expect_start: '2026-11-02 08:30' })
    const body = g.calls[0]!.body
    expect(body.isReminderOn).toBe(false)
    expect(body.sensitivity).toBe('private')
    expect(body.importance).toBe('high')
    expect(body.recurrence.pattern.daysOfWeek).toEqual(['monday', 'tuesday', 'wednesday', 'thursday', 'friday'])
    expect(body.recurrence.range).toEqual({ type: 'endDate', startDate: '2026-11-02', endDate: '2026-12-31', recurrenceTimeZone: VN_ZONE })
  })

  test('a weekly or monthly repeat follows the first day; a repeat needs an end', async () => {
    const g = created()
    const cal = calendar(g.fetchImpl)
    // 2026-11-04 is a Wednesday.
    const weekly = await cal.prepareEvent({ subject: 'x', start: '2026-11-04 09:00', repeat: 'weekly', repeat_count: 4 })
    await cal.createEvent({ plan_id: /plan_id: (\S+)/.exec(weekly)![1]!, expect_start: '2026-11-04 09:00' })
    expect(g.calls[0]!.body.recurrence.pattern.daysOfWeek).toEqual(['wednesday'])
    expect(g.calls[0]!.body.recurrence.range.numberOfOccurrences).toBe(4)
    await expect(cal.prepareEvent({ subject: 'x', start: '2026-11-04 09:00', repeat: 'monthly' })).rejects.toThrow(/repeat_until/)
    await expect(cal.prepareEvent({ subject: 'x', start: '2026-11-04 09:00', show_as: 'away' })).rejects.toThrow(/show_as must be one of/)
  })

  test('dates are read strictly', () => {
    expect(parseLocalDate('2026-11-13')).toBe('2026-11-13')
    expect(parseLocalDate('2026-11-13 09:00')).toBe('2026-11-13')
    for (const bad of ['13/11', '2026-02-30', 'friday']) expect(() => parseLocalDate(bad)).toThrow(/not a date/)
  })
})

describe('changing an existing entry (10/11)', () => {
  const MINE = { ...EVENT, isOrganizer: true, organizer: { emailAddress: { address: 'sam@fabrikam.com' } } }
  const OOF = {
    id: 'AAMkOOF=',
    subject: 'Out of Office',
    start: { dateTime: '2026-11-13T00:00:00.0000000', timeZone: VN_ZONE },
    end: { dateTime: '2026-11-14T00:00:00.0000000', timeZone: VN_ZONE },
    isAllDay: true,
    showAs: 'oof',
    attendees: [],
    isOrganizer: true,
  }

  /** A calendar whose list holds `event`, so it has a short event_id. */
  async function withEvent(event: any) {
    const g = fakeGraph({
      'GET /me/calendarView': () => ({ json: { value: [event] } }),
      'GET /me/events/': () => ({ json: event }),
      'PATCH /me/events/': call => ({ json: { ...event, ...call.body } }),
    })
    const cal = calendar(g.fetchImpl)
    const listed = await cal.listEvents({ since: '2026-11-01', until: '2026-11-30' })
    const eventId = /event_id: (\S+)/.exec(listed)![1]!
    return { g, cal, eventId }
  }

  test('moving a meeting keeps its length, shows who is told, and patches only on step 2', async () => {
    const { g, cal, eventId } = await withEvent(MINE)
    const out = await cal.prepareUpdate({ event_id: eventId, start: '2026-11-03 14:00' })
    expect(out).toContain('Nothing has been changed yet')
    expect(out).toContain('when: 2026-11-02 09:00–10:00 → 2026-11-03 14:00–15:00')
    expect(out).toContain('Saving sends an update to: sam@fabrikam.com')
    expect(g.calls.some(c => c.method === 'PATCH')).toBe(false)
    const plan = /plan_id: (\S+)/.exec(out)![1]!
    await expect(cal.updateEvent({ plan_id: plan, expect_subject: 'Họp khác' })).rejects.toThrow(/refused/)
    const done = await cal.updateEvent({ plan_id: plan, expect_subject: 'Họp vụ việc ABC' })
    expect(done).toContain('an update was sent')
    const patch = g.calls.find(c => c.method === 'PATCH')!.body
    expect(patch.start).toEqual({ dateTime: '2026-11-03T14:00:00', timeZone: VN_ZONE })
    expect(patch.end).toEqual({ dateTime: '2026-11-03T15:00:00', timeZone: VN_ZONE })
    expect(patch.subject).toBeUndefined()
    await expect(cal.updateEvent({ plan_id: plan, expect_subject: 'Họp vụ việc ABC' })).rejects.toThrow(/unknown or expired/)
  })

  test('an all-day entry moves by days and its description can be replaced', async () => {
    const { g, cal, eventId } = await withEvent(OOF)
    const out = await cal.prepareUpdate({ event_id: eventId, start: '2026-11-16', end: '2026-11-17', body: 'Nghỉ phép' })
    expect(out).toContain('when: 2026-11-13, all day → 2026-11-16 to 2026-11-17, all day')
    expect(out).toContain('Nobody else is notified.')
    await cal.updateEvent({ plan_id: /plan_id: (\S+)/.exec(out)![1]!, expect_subject: 'Out of Office' })
    const patch = g.calls.find(c => c.method === 'PATCH')!.body
    expect(patch.start).toEqual({ dateTime: '2026-11-16T00:00:00', timeZone: VN_ZONE })
    expect(patch.end).toEqual({ dateTime: '2026-11-18T00:00:00', timeZone: VN_ZONE })
    expect(patch.isAllDay).toBe(true)
    expect(patch.body).toEqual({ contentType: 'text', content: 'Nghỉ phép' })
  })

  test('attendees are added and removed against the current list', async () => {
    const { g, cal, eventId } = await withEvent(MINE)
    const out = await cal.prepareUpdate({ event_id: eventId, add_attendees: ['client@example.com'], remove_attendees: ['sam@fabrikam.com'] })
    expect(out).toContain('attendees removed (they get a cancellation): sam@fabrikam.com')
    expect(out).toContain('EXTERNAL attendees added (outside the firm): client@example.com')
    await cal.updateEvent({ plan_id: /plan_id: (\S+)/.exec(out)![1]!, expect_subject: 'Họp vụ việc ABC' })
    expect(g.calls.find(c => c.method === 'PATCH')!.body.attendees).toEqual([{ type: 'required', emailAddress: { address: 'client@example.com' } }])
    await expect(cal.prepareUpdate({ event_id: eventId, remove_attendees: ['ghost@fabrikam.com'] })).rejects.toThrow(/not invited/)
  })

  test("someone else's meeting: only show_as and reminder, and nobody is told", async () => {
    const theirs = { ...EVENT, isOrganizer: false, organizer: { emailAddress: { address: 'kien.do@fabrikam.com' } } }
    const { cal, eventId } = await withEvent(theirs)
    await expect(cal.prepareUpdate({ event_id: eventId, start: '2026-11-03 09:00' })).rejects.toThrow(/kien.do@fabrikam.com organised/)
    const out = await cal.prepareUpdate({ event_id: eventId, show_as: 'tentative', reminder_minutes: 30 })
    expect(out).toContain('shown as: Busy → Tentative')
    expect(out).toContain('reminder: 30 minutes before')
    expect(out).toContain('Nobody else is notified.')
  })

  test('nothing to change is said, not sent', async () => {
    const { cal, eventId } = await withEvent(OOF)
    await expect(cal.prepareUpdate({ event_id: eventId })).rejects.toThrow(/nothing to change/)
    await expect(cal.prepareUpdate({ event_id: eventId, show_as: 'oof' })).rejects.toThrow(/nothing would change/)
  })
})

describe('automatic replies (10/11)', () => {
  const graph = () =>
    fakeGraph({
      'GET /me/mailboxSettings/automaticRepliesSetting': () => ({ json: { status: 'disabled' } }),
      'PATCH /me/mailboxSettings': () => ({ status: 200, json: {} }),
    })

  test('prepare changes nothing; set applies whole days in Vietnam time', async () => {
    const g = graph()
    const cal = calendar(g.fetchImpl)
    const out = await cal.prepareAutoReply({ start: '2026-11-13', end: '2026-11-13', message: 'Tôi nghỉ <thứ Sáu>\nViệc gấp gọi văn phòng' })
    expect(out).toContain('Nothing has been changed yet')
    expect(out).toContain('now: automatic replies disabled')
    expect(out).toContain('from 2026-11-13 00:00 to 2026-11-14 00:00')
    expect(g.calls.map(c => c.method)).toEqual(['GET'])
    const done = await cal.setAutoReply({ plan_id: /plan_id: (\S+)/.exec(out)![1]! })
    expect(done).toContain('Automatic replies are set from 2026-11-13 00:00 to 2026-11-14 00:00')
    // The bot once guessed "Settings → Mail"; the result now says where it is.
    expect(done).toContain('Settings (gear) → Account → Automatic replies')
    expect(done).not.toContain('→ Mail')
    const setting = g.calls[1]!.body.automaticRepliesSetting
    expect(setting.status).toBe('scheduled')
    expect(setting.scheduledStartDateTime).toEqual({ dateTime: '2026-11-13T00:00:00', timeZone: VN_ZONE })
    expect(setting.scheduledEndDateTime).toEqual({ dateTime: '2026-11-14T00:00:00', timeZone: VN_ZONE })
    expect(setting.internalReplyMessage).toContain('&lt;thứ Sáu&gt;<br>')
    expect(setting.externalReplyMessage).toBe(setting.internalReplyMessage)
    expect(setting.externalAudience).toBe('all')
    expect(JSON.parse(readFileSync(join(dir, 'sent.jsonl'), 'utf8')).kind).toBe('auto_reply')
  })

  test('turning off, and the rules on what to pass', async () => {
    const g = graph()
    const cal = calendar(g.fetchImpl)
    const off = await cal.prepareAutoReply({ turn_off: true })
    expect(off).toContain('will: turn automatic replies off')
    await cal.setAutoReply({ plan_id: /plan_id: (\S+)/.exec(off)![1]! })
    expect(g.calls[1]!.body).toEqual({ automaticRepliesSetting: { status: 'disabled' } })
    await expect(cal.prepareAutoReply({})).rejects.toThrow(/message is required/)
    await expect(cal.prepareAutoReply({ message: 'x', start: '2026-11-13' })).rejects.toThrow(/both start and end/)
    await expect(cal.setAutoReply({ plan_id: 'a_nope' })).rejects.toThrow(/unknown or expired/)
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
