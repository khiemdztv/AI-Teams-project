/**
 * Teams messages under the sender's own name (fork addition, 02/11), against a
 * fake Graph. What matters: nothing is sent at prepare time, a send needs the
 * plan and the recipients repeated, and a plan is only its own person's.
 */

import { describe, expect, test, beforeEach } from 'bun:test'
import { mkdtempSync, readFileSync, writeFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { TeamsChats, textToTeamsHtml, MAX_TEAMS_MESSAGES_PER_HOUR } from '../m365/teams.js'

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
    return new Response(JSON.stringify(out.json ?? {}), { status: out.status ?? 200 })
  }
  return { calls, fetchImpl }
}

const ME = { id: 'me-id', mail: 'sam@fabrikam.com' }
const KIEN = { id: 'kien-id', displayName: 'Đỗ Kiên', mail: 'kien.do@fabrikam.com' }
const HA = { id: 'ha-id', displayName: 'Hà', mail: 'ha@fabrikam.com' }
const sends = (g: { calls: Call[] }) => g.calls.filter(c => c.method === 'POST')

let dir: string
let actor: string
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'm365-teams-'))
  actor = 'aad-nhan'
})

function teams(fetchImpl: any, now = () => Date.parse('2026-11-02T08:00:00Z')) {
  return new TeamsChats({
    token: async () => 'ACCESS',
    internalDomains: ['fabrikam.com'],
    sendLog: join(dir, 'sent.jsonl'),
    actor: () => actor,
    fetchImpl,
    now,
  })
}

const routes = {
  'GET /me?': () => ({ json: ME }),
  'GET /users/kien.do%40fabrikam.com': () => ({ json: KIEN }),
  'GET /users/ha%40fabrikam.com': () => ({ json: HA }),
  'POST /chats': (c: Call) => (c.path === '/chats' ? { status: 201, json: { id: 'chat-new' } } : { status: 201, json: { id: 'msg' } }),
}

const planIdOf = (out: string) => /plan_id: (t_[0-9a-f]{10})/.exec(out)![1]!

describe('Teams messages as the sender', () => {
  test('text becomes escaped Teams HTML with its line breaks', () => {
    expect(textToTeamsHtml('a < b & c\nd')).toBe('a &lt; b &amp; c<br>d')
  })

  test('prepare sends nothing; send to one colleague uses their 1:1 chat', async () => {
    const g = fakeGraph(routes)
    const t = teams(g.fetchImpl)
    const out = await t.prepareMessage({ to: ['kien.do@fabrikam.com'], text: 'Anh xem giúp em\nhợp đồng nhé' })
    expect(sends(g)).toHaveLength(0)
    expect(out).toContain('Nothing has been sent')
    expect(out).toContain('expect_to: kien.do@fabrikam.com')

    const sent = await t.sendMessage({ plan_id: planIdOf(out), expect_to: ['kien.do@fabrikam.com'] })
    expect(sent).toContain('1:1 chat with Đỗ Kiên <kien.do@fabrikam.com>')
    const [create, message] = sends(g)
    expect(create!.body.chatType).toBe('oneOnOne')
    expect(create!.body.members.map((m: any) => m['user@odata.bind'])).toEqual([
      "https://graph.microsoft.com/v1.0/users('me-id')",
      "https://graph.microsoft.com/v1.0/users('kien-id')",
    ])
    expect(message!.path).toBe('/chats/chat-new/messages')
    expect(message!.body.body).toEqual({ contentType: 'html', content: 'Anh xem giúp em<br>hợp đồng nhé' })
    const log = JSON.parse(readFileSync(join(dir, 'sent.jsonl'), 'utf8').trim())
    expect(log).toMatchObject({ kind: 'teams', as: 'aad-nhan', to: ['kien.do@fabrikam.com'] })
  })

  // Each person's 1:1 chat gets an id from its other member, so the posts can be told apart.
  const perPerson = {
    ...routes,
    'POST /chats': (c: Call) =>
      c.path === '/chats'
        ? { status: 201, json: { id: `chat-${/users\('([^']+)'\)$/.exec(c.body.members.at(-1)['user@odata.bind'])![1]}` } }
        : { status: 201, json: { id: 'msg' } },
  }

  test('several people get the message separately, each in their own 1:1 chat', async () => {
    const g = fakeGraph(perPerson)
    const t = teams(g.fetchImpl)
    const out = await t.prepareMessage({ to: ['ha@fabrikam.com', 'kien.do@fabrikam.com'], text: 'Họp lúc 3h', topic: 'ignored' })
    expect(out).toContain('SEPARATELY')
    const sent = await t.sendMessage({ plan_id: planIdOf(out), expect_to: ['kien.do@fabrikam.com', 'ha@fabrikam.com'] })
    expect(sent).toContain('SEPARATELY')
    const chats = sends(g).filter(c => c.path === '/chats')
    expect(chats.map(c => c.body.chatType)).toEqual(['oneOnOne', 'oneOnOne'])
    expect(chats.every(c => c.body.members.length === 2 && !c.body.topic)).toBe(true)
    expect(sends(g).filter(c => c.path.endsWith('/messages')).map(c => c.path).sort()).toEqual([
      '/chats/chat-ha-id/messages',
      '/chats/chat-kien-id/messages',
    ])
    const log = readFileSync(join(dir, 'sent.jsonl'), 'utf8').trim().split('\n').map(l => JSON.parse(l).to)
    expect(log).toEqual([['ha@fabrikam.com'], ['kien.do@fabrikam.com']])
  })

  test('group: true makes one new group chat with the topic', async () => {
    const g = fakeGraph(routes)
    const t = teams(g.fetchImpl)
    const out = await t.prepareMessage({ to: ['ha@fabrikam.com', 'kien.do@fabrikam.com'], text: 'Họp lúc 3h', group: true, topic: 'Dự án A' })
    expect(out).toContain('ONE new group chat')
    await t.sendMessage({ plan_id: planIdOf(out), expect_to: ['kien.do@fabrikam.com', 'ha@fabrikam.com'] })
    expect(sends(g)[0]!.body).toMatchObject({ chatType: 'group', topic: 'Dự án A' })
    expect(sends(g)[0]!.body.members).toHaveLength(3)
    expect(sends(g)).toHaveLength(2)
  })

  test('a failure part-way says who got it and who did not, and the plan cannot be resent', async () => {
    const g = fakeGraph({
      ...perPerson,
      'POST /chats/chat-kien-id/messages': () => ({ status: 403, json: { error: { code: 'Forbidden', message: 'blocked' } } }),
    })
    const t = teams(g.fetchImpl)
    const id = planIdOf(await t.prepareMessage({ to: ['ha@fabrikam.com', 'kien.do@fabrikam.com'], text: 'x' }))
    const expect_to = ['ha@fabrikam.com', 'kien.do@fabrikam.com']
    await expect(t.sendMessage({ plan_id: id, expect_to })).rejects.toThrow(/sent to ha@fabrikam.com; NOT sent to kien.do@fabrikam.com/)
    await expect(t.sendMessage({ plan_id: id, expect_to })).rejects.toThrow(/unknown or expired/)
  })

  test('a name is looked up in the directory; several matches ask, an exact name wins', async () => {
    const JIMMY = { id: 'jimmy-id', displayName: 'Jimmy Tran', mail: 'jimmy.tran@fabrikam.com' }
    const JIMMY2 = { id: 'jimmy2-id', displayName: 'Jimmy Le', mail: 'jimmy.le@fabrikam.com' }
    let found: any[] = [JIMMY]
    const g = fakeGraph({ ...routes, 'GET /users?': () => ({ json: { value: found } }) })
    const t = teams(g.fetchImpl)
    const out = await t.prepareMessage({ to: ['Jimmy'], text: 'hi' })
    expect(out).toContain('1:1 chat with Jimmy Tran <jimmy.tran@fabrikam.com>')
    expect(out).toContain('expect_to: jimmy.tran@fabrikam.com')
    const search = g.calls.find(c => c.path.startsWith('/users?'))!
    expect(decodeURIComponent(search.path)).toContain('"displayName:Jimmy" OR "mail:Jimmy"')

    found = [JIMMY, JIMMY2]
    await expect(t.prepareMessage({ to: ['Jimmy'], text: 'hi' })).rejects.toThrow(
      /matches several people.*Jimmy Tran <jimmy.tran@fabrikam.com>, Jimmy Le <jimmy.le@fabrikam.com>/,
    )
    found = [KIEN, { id: 'k2', displayName: 'Kiên Nguyễn', mail: 'kien.nguyen@fabrikam.com' }]
    // Both match "Kiên"; typed without accents, the full name picks one.
    expect(await t.prepareMessage({ to: ['do kien'], text: 'x' })).toContain('kien.do@fabrikam.com')
    await expect(t.prepareMessage({ to: ['Kiên'], text: 'x' })).rejects.toThrow(/matches several people/)
    found = [{ id: 'guest', displayName: 'Jimmy Guest', mail: 'jimmy@gmail.com' }, JIMMY]
    expect(await t.prepareMessage({ to: ['Jimmy'], text: 'hi' })).toContain('jimmy.tran@fabrikam.com')
    found = []
    await expect(t.prepareMessage({ to: ['Nobody'], text: 'hi' })).rejects.toThrow(/nobody called "Nobody"/)
    expect(sends(g)).toHaveLength(0)
  })

  test('someone outside the directory cannot be messaged in a new chat', async () => {
    const g = fakeGraph(routes)
    await expect(teams(g.fetchImpl).prepareMessage({ to: ['client@example.com'], text: 'hi' })).rejects.toThrow(/not in the firm's directory/)
  })

  test('an existing chat flags guests as EXTERNAL and posts into that chat', async () => {
    const chat = {
      id: '19:abc@thread.v2',
      chatType: 'group',
      topic: 'Khách hàng B',
      members: [{ userId: 'me-id', email: ME.mail }, { userId: 'g', email: 'guest@client.com' }],
    }
    const g = fakeGraph({
      ...routes,
      'GET /me/chats': () => ({ json: { value: [chat] } }),
      'GET /chats/': () => ({ json: chat }),
      'POST /chats/': () => ({ status: 201, json: { id: 'msg' } }),
    })
    const t = teams(g.fetchImpl)
    const list = await t.listChats({})
    const chatId = /chat_id: (c_[0-9a-f]{10})/.exec(list)![1]!
    const out = await t.prepareMessage({ chat_id: chatId, text: 'Chào anh' })
    expect(out).toContain('EXTERNAL people in this chat (outside the firm): guest@client.com')
    await t.sendMessage({ plan_id: planIdOf(out), expect_to: ['guest@client.com'] })
    expect(sends(g).map(c => c.path)).toEqual([`/chats/${encodeURIComponent(chat.id)}/messages`])
  })

  test('a channel message goes to that team and channel', async () => {
    const g = fakeGraph({
      ...routes,
      'GET /me/joinedTeams': () => ({ json: { value: [{ id: 'team1', displayName: 'Fabrikam' }] } }),
      'GET /teams/team1/channels?': () => ({ json: { value: [{ id: 'ch1', displayName: 'General' }] } }),
      'GET /teams/team1/channels/ch1': () => ({ json: { displayName: 'General' } }),
      'GET /teams/team1?': () => ({ json: { displayName: 'Fabrikam' } }),
      'POST /teams/team1/channels/ch1/messages': () => ({ status: 201, json: { id: 'msg' } }),
    })
    const t = teams(g.fetchImpl)
    const channelId = /channel_id: (h_[0-9a-f]{10}) · Fabrikam \/ General/.exec(await t.listChannels())![1]!
    const out = await t.prepareMessage({ channel_id: channelId, text: 'Thông báo' })
    expect(out).toContain('expect_to: fabrikam / general')
    await t.sendMessage({ plan_id: planIdOf(out), expect_to: ['Fabrikam / General'] })
    expect(sends(g).map(c => c.path)).toEqual(['/teams/team1/channels/ch1/messages'])
  })

  test('wrong recipients, another person, or a reused plan send nothing', async () => {
    const g = fakeGraph(routes)
    const t = teams(g.fetchImpl)
    const id = planIdOf(await t.prepareMessage({ to: ['kien.do@fabrikam.com'], text: 'x' }))
    await expect(t.sendMessage({ plan_id: id, expect_to: ['ha@fabrikam.com'] })).rejects.toThrow(/must repeat/)
    actor = 'aad-someone-else'
    await expect(t.sendMessage({ plan_id: id, expect_to: ['kien.do@fabrikam.com'] })).rejects.toThrow(/unknown or expired/)
    actor = 'aad-nhan'
    await t.sendMessage({ plan_id: id, expect_to: ['kien.do@fabrikam.com'] })
    const before = sends(g).length
    await expect(t.sendMessage({ plan_id: id, expect_to: ['kien.do@fabrikam.com'] })).rejects.toThrow(/unknown or expired/)
    expect(sends(g)).toHaveLength(before)
  })

  test('exactly one target, and the hourly cap holds', async () => {
    const g = fakeGraph(routes)
    const t = teams(g.fetchImpl)
    await expect(t.prepareMessage({ text: 'x' })).rejects.toThrow(/exactly one/)
    await expect(t.prepareMessage({ to: ['kien.do@fabrikam.com'], chat_id: 'c_0000000000', text: 'x' })).rejects.toThrow(/exactly one/)
    const at = new Date(Date.parse('2026-11-02T07:30:00Z')).toISOString()
    writeFileSync(join(dir, 'sent.jsonl'), Array.from({ length: MAX_TEAMS_MESSAGES_PER_HOUR }, () => JSON.stringify({ at, kind: 'teams' })).join('\n') + '\n')
    const id = planIdOf(await t.prepareMessage({ to: ['kien.do@fabrikam.com'], text: 'x' }))
    await expect(t.sendMessage({ plan_id: id, expect_to: ['kien.do@fabrikam.com'] })).rejects.toThrow(/would pass/)
    expect(sends(g)).toHaveLength(0)
  })

  test('separate messages count one each against the cap', async () => {
    const g = fakeGraph(routes)
    const t = teams(g.fetchImpl)
    const at = new Date(Date.parse('2026-11-02T07:30:00Z')).toISOString()
    writeFileSync(join(dir, 'sent.jsonl'), Array.from({ length: MAX_TEAMS_MESSAGES_PER_HOUR - 1 }, () => JSON.stringify({ at, kind: 'teams' })).join('\n') + '\n')
    const id = planIdOf(await t.prepareMessage({ to: ['ha@fabrikam.com', 'kien.do@fabrikam.com'], text: 'x' }))
    await expect(t.sendMessage({ plan_id: id, expect_to: ['ha@fabrikam.com', 'kien.do@fabrikam.com'] })).rejects.toThrow(/would pass/)
    expect(sends(g)).toHaveLength(0)
  })
})
