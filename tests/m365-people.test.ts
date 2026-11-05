/**
 * Email recipients and meeting attendees by name (fork addition, 05/11),
 * against a fake Graph. What matters: a name becomes one colleague or an
 * error to ask the sender about, never a guess; addresses (people outside the
 * firm included) are kept as given; the sender sees name and address before
 * anything goes out, and expect_* checks the looked-up address.
 */

import { describe, expect, test, beforeEach } from 'bun:test'
import { mkdtempSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { Mail, DRAFT_CATEGORY } from '../m365/mail.js'
import { Calendar } from '../m365/calendar.js'
import { People } from '../m365/people.js'

type Call = { method: string; path: string; body: any }

function fakeGraph(routes: Record<string, (call: Call) => { status?: number; json?: unknown }>) {
  const calls: Call[] = []
  const fetchImpl = async (url: string, init?: RequestInit) => {
    const call: Call = {
      method: String(init?.method ?? 'GET'),
      path: url.replace('https://graph.microsoft.com/v1.0', ''),
      body: typeof init?.body === 'string' ? JSON.parse(init.body) : undefined,
    }
    calls.push(call)
    const key = Object.keys(routes)
      .filter(k => {
        const [m, p] = k.split(' ')
        return m === call.method && call.path.startsWith(p!)
      })
      .sort((a, b) => b.length - a.length)[0]
    if (!key) return new Response(JSON.stringify({ error: { code: 'NotFound', message: 'no route' } }), { status: 404 })
    const out = routes[key]!(call)
    return out.status === 202 ? new Response(null, { status: 202 }) : new Response(JSON.stringify(out.json ?? {}), { status: out.status ?? 200 })
  }
  return { calls, fetchImpl }
}

const JIMMY = { id: 'jimmy-id', displayName: 'Jimmy Tran', mail: 'jimmy.tran@fabrikam.com' }
const JIMMY2 = { id: 'jimmy2-id', displayName: 'Jimmy Le', mail: 'jimmy.le@fabrikam.com' }
const HA = { id: 'ha-id', displayName: 'Nguyễn Thu Hà', mail: 'ha.nguyen@fabrikam.com' }
const GUEST = { id: 'guest', displayName: 'Jimmy Guest', mail: 'jimmy@gmail.com' }

/** The directory: a search returns everyone whose name or mail contains the query, accents ignored. */
function directory(people: any[]) {
  const fold = (s: string) => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase()
  return (c: Call) => {
    const q = /"displayName:([^"]*)"/.exec(decodeURIComponent(c.path))?.[1] ?? ''
    return { json: { value: people.filter(p => fold(p.displayName).includes(fold(q)) || p.mail.startsWith(q.toLowerCase())) } }
  }
}

let dir: string
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'm365-people-'))
})

function mail(fetchImpl: any) {
  return new Mail({
    token: async () => 'ACCESS',
    internalDomains: ['fabrikam.com'],
    sendLog: join(dir, 'sent.jsonl'),
    fetchImpl,
    now: () => Date.parse('2026-11-05T08:00:00Z'),
  })
}

function calendar(fetchImpl: any) {
  return new Calendar({
    token: async () => 'ACCESS',
    internalDomains: ['fabrikam.com'],
    sendLog: join(dir, 'sent.jsonl'),
    fetchImpl,
    now: () => Date.parse('2026-11-05T08:00:00Z'),
  })
}

/** A mailbox that keeps the one draft it is given, as Graph returns it. */
function mailbox(people: any[]) {
  let draft: any
  const g = fakeGraph({
    'GET /users?': directory(people),
    'POST /me/messages': c => {
      if (c.path.endsWith('/send')) return { status: 202 }
      draft = { id: 'd1', isDraft: true, bccRecipients: [], ...c.body, categories: [DRAFT_CATEGORY] }
      return { status: 201, json: draft }
    },
    'PATCH /me/messages/d1': c => {
      draft = { ...draft, ...c.body }
      return { json: draft }
    },
    'GET /me/messages/d1': () => ({ json: draft }),
  })
  return g
}

describe('a name is one colleague or a question', () => {
  const people = (found: any[]) =>
    new People(async path => (path.startsWith('/users?') ? directory(found)({ method: 'GET', path, body: undefined }).json : {}), ['fabrikam.com'])

  test('one match is that colleague', async () => {
    expect(await people([JIMMY, HA]).recipient('Jimmy')).toEqual({ address: 'jimmy.tran@fabrikam.com', name: 'Jimmy Tran' })
  })

  test('several matches are listed, never picked', async () => {
    await expect(people([JIMMY, JIMMY2]).recipient('Jimmy')).rejects.toThrow(
      /"Jimmy" matches several people; ask the sender.*Jimmy Tran <jimmy.tran@fabrikam.com>, Jimmy Le <jimmy.le@fabrikam.com>/,
    )
  })

  test('an exact name, typed without accents, wins over partial matches', async () => {
    const HA2 = { id: 'ha2', displayName: 'Hà', mail: 'ha@fabrikam.com' }
    expect((await people([HA, HA2]).recipient('ha')).address).toBe('ha@fabrikam.com')
  })

  test('nobody is an error asking for the address', async () => {
    await expect(people([JIMMY]).recipient('Jordan')).rejects.toThrow(/nobody called "Jordan".*ask the sender for the email address/)
  })

  test('a guest account is not a colleague', async () => {
    expect((await people([GUEST, JIMMY]).recipient('Jimmy')).address).toBe('jimmy.tran@fabrikam.com')
    await expect(people([GUEST]).recipient('Jimmy')).rejects.toThrow(/nobody called "Jimmy"/)
  })

  test('an address is kept as given, outside the firm too, without a lookup', async () => {
    let lookups = 0
    const p = new People(async () => {
      lookups++
      return {}
    }, ['fabrikam.com'])
    expect(await p.recipient('client@example.com')).toEqual({ address: 'client@example.com' })
    expect(await p.recipient('Ms Lan <Lan@Example.com>')).toEqual({ address: 'Lan@Example.com', name: 'Ms Lan' })
    expect(lookups).toBe(0)
  })
})

describe('email recipients by name', () => {
  test('the draft shows name and address; expect_to by name or address both send', async () => {
    const g = mailbox([JIMMY, HA])
    const m = mail(g.fetchImpl)
    const out = await m.createDraft({ to: ['Jimmy', 'client@example.com'], cc: ['Hà'], subject: 'Hợp đồng', body: 'Nội dung', signature: false })
    expect(out).toContain('to: client@example.com, Jimmy Tran <jimmy.tran@fabrikam.com>')
    expect(out).toContain('cc: Nguyễn Thu Hà <ha.nguyen@fabrikam.com>')
    expect(out).toContain('EXTERNAL recipients (outside the firm): client@example.com')
    const posted = g.calls.find(c => c.method === 'POST')!.body
    expect(posted.toRecipients).toEqual([
      { emailAddress: { address: 'jimmy.tran@fabrikam.com', name: 'Jimmy Tran' } },
      { emailAddress: { address: 'client@example.com' } },
    ])
    expect(g.calls.some(c => c.path.endsWith('/send'))).toBe(false)

    // The sender repeats names as they said them, or the addresses shown: both mean the same people.
    await expect(m.sendDraft({ draft_id: 'd1', expect_to: ['Jimmy'], expect_cc: ['Hà'], expect_subject: 'Hợp đồng' })).rejects.toThrow(/refused/)
    const sent = await m.sendDraft({ draft_id: 'd1', expect_to: ['client@example.com', 'Jimmy'], expect_cc: ['ha.nguyen@fabrikam.com'], expect_subject: 'Hợp đồng' })
    expect(sent).toContain('to: client@example.com, jimmy.tran@fabrikam.com')
  })

  test('a name matching several people saves no draft', async () => {
    const g = mailbox([JIMMY, JIMMY2])
    await expect(mail(g.fetchImpl).createDraft({ to: ['Jimmy'], subject: 's', body: 'b', signature: false })).rejects.toThrow(/matches several people/)
    expect(g.calls.filter(c => c.method !== 'GET')).toHaveLength(0)
  })

  test('changing recipients of a draft takes names too', async () => {
    const g = mailbox([JIMMY, HA])
    const m = mail(g.fetchImpl)
    await m.createDraft({ to: ['client@example.com'], subject: 's', body: 'b', signature: false })
    const out = await m.updateDraft({ draft_id: 'd1', to: ['Hà'] })
    expect(out).toContain('to: Nguyễn Thu Hà <ha.nguyen@fabrikam.com>')
  })

  test('a name in expect_to that now points at someone else does not send', async () => {
    let people = [JIMMY]
    const routes = mailbox(people)
    const m = mail(async (url: string, init?: RequestInit) => {
      if (url.includes('/users?')) return new Response(JSON.stringify(directory(people)({ method: 'GET', path: url, body: undefined }).json))
      return routes.fetchImpl(url, init)
    })
    await m.createDraft({ to: ['Jimmy'], subject: 's', body: 'b', signature: false })
    people = [JIMMY, JIMMY2]
    await expect(m.sendDraft({ draft_id: 'd1', expect_to: ['Jimmy'], expect_subject: 's' })).rejects.toThrow(/matches several people/)
    expect(routes.calls.some(c => c.path.endsWith('/send'))).toBe(false)
  })
})

describe('meeting attendees by name', () => {
  const EVENT = { id: 'e1', subject: 'Họp', start: { dateTime: '2026-11-06T09:00:00' }, end: { dateTime: '2026-11-06T10:00:00' }, webLink: 'https://outlook/e1' }

  test('the plan shows name and address; the invitation uses the looked-up address', async () => {
    const g = fakeGraph({ 'GET /users?': directory([JIMMY, HA]), 'POST /me/events': () => ({ status: 201, json: EVENT }) })
    const cal = calendar(g.fetchImpl)
    const out = await cal.prepareEvent({ subject: 'Họp', start: '2026-11-06 09:00', attendees: ['Jimmy', 'client@example.com'] })
    expect(out).toContain('attendees: client@example.com, Jimmy Tran <jimmy.tran@fabrikam.com>')
    expect(out).toContain('EXTERNAL attendees (outside the firm): client@example.com')
    expect(g.calls.filter(c => c.method === 'POST')).toHaveLength(0)

    const plan = /plan_id: (\S+)/.exec(out)![1]!
    await expect(cal.createEvent({ plan_id: plan, expect_attendees: ['Jimmy'], expect_start: '2026-11-06 09:00' })).rejects.toThrow(
      /refused.*Jimmy Tran <jimmy.tran@fabrikam.com>/,
    )
    await cal.createEvent({ plan_id: plan, expect_attendees: ['jimmy.tran@fabrikam.com', 'client@example.com'], expect_start: '2026-11-06 09:00' })
    expect(g.calls.find(c => c.path === '/me/events')!.body.attendees).toEqual([
      { type: 'required', emailAddress: { address: 'client@example.com' } },
      { type: 'required', emailAddress: { address: 'jimmy.tran@fabrikam.com', name: 'Jimmy Tran' } },
    ])
  })

  test('expect_attendees by name matches the same people', async () => {
    const g = fakeGraph({ 'GET /users?': directory([JIMMY, HA]), 'POST /me/events': () => ({ status: 201, json: EVENT }) })
    const cal = calendar(g.fetchImpl)
    const plan = /plan_id: (\S+)/.exec(await cal.prepareEvent({ subject: 'Họp', start: '2026-11-06 09:00', attendees: ['Hà', 'Jimmy'] }))![1]!
    expect(await cal.createEvent({ plan_id: plan, expect_attendees: ['Jimmy', 'Hà'], expect_start: '2026-11-06 09:00' })).toContain('invitations sent')
  })

  test('a name matching several people plans nothing; free times look names up too', async () => {
    const g = fakeGraph({
      'GET /users?': directory([JIMMY, JIMMY2, HA]),
      'POST /me/findMeetingTimes': () => ({ json: { meetingTimeSuggestions: [] } }),
    })
    const cal = calendar(g.fetchImpl)
    await expect(cal.prepareEvent({ subject: 'Họp', start: '2026-11-06 09:00', attendees: ['Jimmy'] })).rejects.toThrow(/matches several people/)
    await cal.findTime({ attendees: ['Hà'] })
    expect(g.calls.find(c => c.path === '/me/findMeetingTimes')!.body.attendees).toEqual([
      { type: 'required', emailAddress: { address: 'ha.nguyen@fabrikam.com' } },
    ])
  })
})
