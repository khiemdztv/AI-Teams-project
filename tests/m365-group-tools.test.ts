/**
 * Fork addition (18/11): the group-chat tools, against a fake Graph. Adding
 * people to a chat and sharing a file each take two steps: the plan changes
 * nothing, and only the plan_id with the same people (or "organization")
 * repeated acts. chat_id "this" is the group chat the message came from,
 * known from turn.json and never from the model.
 */

import { describe, expect, test, beforeEach } from 'bun:test'
import { mkdtempSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { TeamsChats } from '../m365/teams.js'
import { Files } from '../m365/files.js'

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
const GROUP = '19:abc@thread.v2'
const writes = (g: { calls: Call[] }) => g.calls.filter(c => c.method !== 'GET')

let dir: string
let actor: string
let current: string | undefined
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'm365-group-'))
  actor = 'aad-nhan'
  current = GROUP
})

function teams(fetchImpl: any) {
  return new TeamsChats({
    token: async () => 'ACCESS',
    internalDomains: ['fabrikam.com'],
    sendLog: join(dir, 'sent.jsonl'),
    actor: () => actor,
    currentChat: () => current,
    fetchImpl,
    now: () => Date.parse('2026-11-18T02:00:00Z'),
  })
}

const groupChat = (chatType = 'group') => ({
  id: GROUP,
  chatType,
  topic: 'Internal IT',
  members: [
    { userId: 'me-id', email: 'sam@fabrikam.com' },
    { userId: 'ha-id', email: 'ha@fabrikam.com' },
  ],
})

const routes = (chatType = 'group') => ({
  'GET /me?': () => ({ json: ME }),
  'GET /users/kien.do%40fabrikam.com': () => ({ json: KIEN }),
  'GET /users/ha%40fabrikam.com': () => ({ json: HA }),
  [`GET /me/chats/${encodeURIComponent(GROUP)}?`]: () => ({ json: groupChat(chatType) }),
  [`GET /me/chats/${encodeURIComponent(GROUP)}/messages`]: () => ({ json: { value: [] } }),
  [`POST /chats/${encodeURIComponent(GROUP)}/members`]: () => ({ status: 201, json: { id: 'm1' } }),
})

const planIdOf = (out: string) => /plan_id: ([a-z]_[0-9a-f]+)/.exec(out)![1]!

describe('chat_id "this"', () => {
  test('reads the group chat the message came from', async () => {
    const g = fakeGraph(routes())
    await teams(g.fetchImpl).readChat({ chat_id: 'this' })
    expect(g.calls.some(c => c.path.startsWith(`/me/chats/${encodeURIComponent(GROUP)}/messages`))).toBe(true)
  })

  test('is refused outside a group chat, and a raw Graph id is never accepted', async () => {
    current = undefined
    const g = fakeGraph(routes())
    await expect(teams(g.fetchImpl).readChat({ chat_id: 'this' })).rejects.toThrow(/only works for a message sent in a group chat/)
    current = GROUP
    await expect(teams(g.fetchImpl).readChat({ chat_id: GROUP })).rejects.toThrow(/c_/)
  })
})

describe('adding people to a group chat', () => {
  test('prepare adds nobody; the yes with the same people adds them with the history', async () => {
    const g = fakeGraph(routes())
    const t = teams(g.fetchImpl)
    const out = await t.prepareAddMembers({ chat_id: 'this', people: ['kien.do@fabrikam.com'] })
    expect(writes(g)).toHaveLength(0)
    expect(out).toContain('Nobody has been added yet')
    expect(out).toContain('expect_add: kien.do@fabrikam.com')

    await expect(t.addMembers({ plan_id: planIdOf(out), expect_add: ['ha@fabrikam.com'] })).rejects.toThrow(/repeat the plan/)
    expect(writes(g)).toHaveLength(0)

    const done = await t.addMembers({ plan_id: planIdOf(out), expect_add: ['kien.do@fabrikam.com'] })
    expect(done).toContain('kien.do@fabrikam.com')
    const [add] = writes(g)
    expect(add!.path).toBe(`/chats/${encodeURIComponent(GROUP)}/members`)
    expect(add!.body['user@odata.bind']).toBe("https://graph.microsoft.com/v1.0/users('kien-id')")
    expect(add!.body.visibleHistoryStartDateTime).toBe('0001-01-01T00:00:00Z')

    // Spent: the same plan cannot add anyone again.
    await expect(t.addMembers({ plan_id: planIdOf(out), expect_add: ['kien.do@fabrikam.com'] })).rejects.toThrow(/unknown or expired/)
  })

  test('someone already in the chat is skipped, and a chat with nobody new is refused', async () => {
    const g = fakeGraph(routes())
    const t = teams(g.fetchImpl)
    const out = await t.prepareAddMembers({ chat_id: 'this', people: ['kien.do@fabrikam.com', 'ha@fabrikam.com'] })
    expect(out).toContain('expect_add: kien.do@fabrikam.com')
    expect(out).toContain('already in the chat, skipped: ha@fabrikam.com')
    await expect(t.prepareAddMembers({ chat_id: 'this', people: ['ha@fabrikam.com'] })).rejects.toThrow(/nobody to add/)
  })

  test('a 1:1 chat cannot take more people', async () => {
    const g = fakeGraph(routes('oneOnOne'))
    await expect(teams(g.fetchImpl).prepareAddMembers({ chat_id: 'this', people: ['kien.do@fabrikam.com'] })).rejects.toThrow(
      /only be added to a group chat/,
    )
  })

  test("a plan is only its own sender's", async () => {
    const g = fakeGraph(routes())
    const t = teams(g.fetchImpl)
    const out = await t.prepareAddMembers({ chat_id: 'this', people: ['kien.do@fabrikam.com'] })
    actor = 'aad-someone-else'
    await expect(t.addMembers({ plan_id: planIdOf(out), expect_add: ['kien.do@fabrikam.com'] })).rejects.toThrow(/unknown or expired/)
    expect(writes(g)).toHaveLength(0)
  })
})

describe('sharing a file', () => {
  const URL_ = 'https://fabrikam-my.sharepoint.com/personal/nhan_le/Documents/huong-dan.pdf'
  const ITEM = { id: 'item-1', name: 'huong-dan.pdf', file: {}, webUrl: URL_, parentReference: { driveId: 'drive-1' } }
  const fileRoutes = {
    'GET /shares/': () => ({ json: ITEM }),
    'POST /drives/drive-1/items/item-1/createLink': () => ({ json: { link: { webUrl: 'https://fabrikam-my.sharepoint.com/:b:/g/org-link' } } }),
    'POST /drives/drive-1/items/item-1/invite': () => ({ json: { value: [{ link: { webUrl: 'https://fabrikam-my.sharepoint.com/:b:/g/people-link' } }] } }),
  }
  const files = (fetchImpl: any) =>
    new Files({ token: async () => 'ACCESS', actor: () => actor, internalDomains: ['fabrikam.com'], fetchImpl, now: () => 0 })

  test('with the whole firm: nothing until the yes, then a view link scoped to the organization', async () => {
    const g = fakeGraph(fileRoutes)
    const f = files(g.fetchImpl)
    const out = await f.prepareShare({ url: URL_, organization: true })
    expect(writes(g)).toHaveLength(0)
    expect(out).toContain('EVERYONE in the firm')
    expect(out).toContain('expect: organization')
    await expect(f.share({ plan_id: planIdOf(out), expect: 'kien.do@fabrikam.com' })).rejects.toThrow(/repeat the plan/)
    const done = await f.share({ plan_id: planIdOf(out), expect: 'organization' })
    expect(done).toContain('org-link')
    const [link] = writes(g)
    expect(link!.body).toEqual({ type: 'view', scope: 'organization' })
  })

  test('with named colleagues: read only, no email; outsiders are refused', async () => {
    const g = fakeGraph(fileRoutes)
    const f = files(g.fetchImpl)
    await expect(f.prepareShare({ url: URL_, people: ['someone@gmail.com'] })).rejects.toThrow(/outside the firm/)
    await expect(f.prepareShare({ url: URL_ })).rejects.toThrow(/exactly one/)
    const out = await f.prepareShare({ url: URL_, people: ['kien.do@fabrikam.com'] })
    await f.share({ plan_id: planIdOf(out), expect: 'kien.do@fabrikam.com' })
    const [invite] = writes(g)
    expect(invite!.body).toEqual({
      recipients: [{ email: 'kien.do@fabrikam.com' }],
      roles: ['read'],
      requireSignIn: true,
      sendInvitation: false,
    })
  })

  test("a share plan is only its own sender's", async () => {
    const g = fakeGraph(fileRoutes)
    const f = files(g.fetchImpl)
    const out = await f.prepareShare({ url: URL_, organization: true })
    actor = 'aad-someone-else'
    await expect(f.share({ plan_id: planIdOf(out), expect: 'organization' })).rejects.toThrow(/unknown or expired/)
    expect(writes(g)).toHaveLength(0)
  })
})
