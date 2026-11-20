/** Fork (20/11): the webhook router (src/router.ts), against fake sessions on ephemeral ports. */

import { afterAll, afterEach, beforeEach, describe, expect, test } from 'bun:test'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, renameSync, writeFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { ROUTER_LOG, WorkersCache, routeFor, startRouter, verdictText, type Router } from '../src/router.js'

const AN = '11111111-1111-1111-1111-111111111111'
const BINH = '22222222-2222-2222-2222-222222222222'
const OTHER = '33333333-3333-3333-3333-333333333333'

type Seen = { body: Uint8Array; headers: Headers; path: string }

/** A fake session: records what reached it and answers with `reply`. */
function fakeSession(reply: (req: Request) => Response | Promise<Response> = () => new Response(null, { status: 200 })) {
  const seen: Seen[] = []
  const server = Bun.serve({
    hostname: '127.0.0.1',
    port: 0,
    fetch: async req => {
      seen.push({ body: new Uint8Array(await req.arrayBuffer()), headers: req.headers, path: new URL(req.url).pathname })
      return reply(req)
    },
  })
  return { server, seen, port: server.port! }
}

/** A port nothing listens on. */
function deadPort(): number {
  const s = Bun.serve({ hostname: '127.0.0.1', port: 0, fetch: () => new Response() })
  const p = s.port!
  s.stop(true)
  return p
}

function writeWorkers(root: string, defaultPort: number, workers: { code: string; owner: string; port: number }[]) {
  // Replace by rename, as add-bot-user.sh does.
  const tmp = join(root, 'workers.json.tmp')
  writeFileSync(tmp, JSON.stringify({ defaultPort, workers }))
  renameSync(tmp, join(root, 'workers.json'))
}

function activity(fields: Record<string, unknown>): string {
  return JSON.stringify({ type: 'message', conversation: { id: 'a:conv-1' }, ...fields })
}

const TOKEN = 'Bearer eyJhbGciOiJSUzI1NiJ9.e30.c2lnbmF0dXJl'

async function post(router: Router, body: string | Uint8Array<ArrayBuffer>, headers: Record<string, string> = {}) {
  return fetch(`http://127.0.0.1:${router.server.port}/api/messages`, {
    method: 'POST',
    headers: { authorization: TOKEN, 'content-type': 'application/json; charset=utf-8', ...headers },
    body,
  })
}

const servers: { stop(force?: boolean): void }[] = []
afterAll(() => {
  for (const s of servers) s.stop(true)
})

let root: string
let shared: ReturnType<typeof fakeSession>
let an: ReturnType<typeof fakeSession>
let binh: ReturnType<typeof fakeSession>
let router: Router

beforeEach(() => {
  root = mkdtempSync(join(tmpdir(), 'router-'))
  shared = fakeSession()
  an = fakeSession()
  binh = fakeSession()
  servers.push(shared.server, an.server, binh.server)
  writeWorkers(root, shared.port, [
    { code: 'an', owner: AN, port: an.port },
    { code: 'binh', owner: BINH, port: binh.port },
  ])
  router = startRouter({ rootDir: root, port: 0 })
})

afterEach(() => router.stop())

describe('routing', () => {
  test('a sender with a worker goes to that worker', async () => {
    const res = await post(router, activity({ text: 'chào', from: { aadObjectId: BINH.toUpperCase() } }))
    expect(res.status).toBe(200)
    expect(binh.seen).toHaveLength(1)
    expect(an.seen).toHaveLength(0)
    expect(shared.seen).toHaveLength(0)
  })

  test('anyone else goes to the shared session, as does a body that does not parse', async () => {
    await post(router, activity({ text: 'chào', from: { aadObjectId: OTHER } }))
    await post(router, activity({ type: 'conversationUpdate' }))
    await post(router, '{not json')
    expect(shared.seen).toHaveLength(3)
    expect(an.seen.length + binh.seen.length).toBe(0)
  })

  test('a verdict with a registry entry goes to the session that asked, even from another worker\'s owner', async () => {
    mkdirSync(join(root, 'permission-routes'))
    writeFileSync(join(root, 'permission-routes', 'abcde'), `${an.port}\n`)
    await post(router, activity({ text: ' Y ABCDE ', from: { aadObjectId: BINH } }))
    expect(an.seen).toHaveLength(1)
    expect(binh.seen).toHaveLength(0)
  })

  test('a verdict without a registry entry routes like any message', async () => {
    await post(router, activity({ text: 'y abcde', from: { aadObjectId: BINH } }))
    await post(router, activity({ text: 'n qwert', from: { aadObjectId: OTHER } }))
    expect(binh.seen).toHaveLength(1)
    expect(shared.seen).toHaveLength(1)
  })

  test('a registry entry holding garbage, or a port no session uses, is ignored', async () => {
    mkdirSync(join(root, 'permission-routes'))
    writeFileSync(join(root, 'permission-routes', 'abcde'), 'not a port')
    writeFileSync(join(root, 'permission-routes', 'fghij'), '22')
    writeFileSync(join(root, 'permission-routes', 'kmnop'), `${an.port}; rm -rf /`)
    for (const code of ['abcde', 'fghij', 'kmnop']) {
      await post(router, activity({ text: `y ${code}`, from: { aadObjectId: OTHER } }))
    }
    expect(shared.seen).toHaveLength(3)
    expect(an.seen).toHaveLength(0)
  })

  test('a verdict behind the bot mention is still read as one', () => {
    expect(verdictText('<at>Contoso Agent</at> y abcde')).toContain('y abcde')
    const workers = { defaultPort: 4000, workers: [{ code: 'an', owner: AN, port: 4001 }] }
    mkdirSync(join(root, 'permission-routes'))
    writeFileSync(join(root, 'permission-routes', 'abcde'), '4001')
    expect(routeFor({ text: '<at>Bot</at>&nbsp;y abcde', from: {} }, workers, root)).toEqual({ port: 4001, via: 'verdict' })
    expect(routeFor({ text: 'y abcde please', from: { aadObjectId: OTHER } }, workers, root)).toEqual({
      port: 4000,
      via: 'default',
    })
    expect(routeFor(null, workers, root)).toEqual({ port: 4000, via: 'default' })
  })
})

describe('pass-through', () => {
  test('body bytes and the authorization and content-type headers arrive unchanged, at the same path', async () => {
    const text = 'chào  anh éế\n\t"quoted" \\u00e9 😀'
    const raw = new TextEncoder().encode(
      `{ "type":"message",\r\n  "text": ${JSON.stringify(text)}, "from":{"aadObjectId":"${AN}"} }  \n`,
    )
    await post(router, raw)
    expect(an.seen).toHaveLength(1)
    const got = an.seen[0]!
    expect(Buffer.from(got.body).equals(Buffer.from(raw))).toBe(true)
    expect(got.headers.get('authorization')).toBe(TOKEN)
    expect(got.headers.get('content-type')).toBe('application/json; charset=utf-8')
    expect(got.path).toBe('/api/messages')
  })

  test('an invoke answer comes back with its status, body and content-type', async () => {
    const answer = JSON.stringify({ status: 200, body: { statusCode: 200, type: 'application/vnd.microsoft.activity.message', value: 'ok' } })
    const s = fakeSession(() => new Response(answer, { status: 200, headers: { 'content-type': 'application/json' } }))
    servers.push(s.server)
    writeWorkers(root, s.port, [])
    router.stop()
    router = startRouter({ rootDir: root, port: 0 })
    const res = await post(router, activity({ type: 'invoke', name: 'fileConsent/invoke', from: { aadObjectId: OTHER } }))
    expect(res.status).toBe(200)
    expect(res.headers.get('content-type')).toBe('application/json')
    expect(await res.text()).toBe(answer)
  })

  test('a session\'s refusal (401) is returned as is', async () => {
    const s = fakeSession(() => new Response('unauthorized', { status: 401, headers: { 'content-type': 'text/plain' } }))
    servers.push(s.server)
    writeWorkers(root, s.port, [])
    router.stop()
    router = startRouter({ rootDir: root, port: 0 })
    const res = await post(router, activity({}), { authorization: 'Bearer forged' })
    expect(res.status).toBe(401)
    expect(await res.text()).toBe('unauthorized')
  })

  test('an oversized body is refused before it is forwarded', async () => {
    router.stop()
    router = startRouter({ rootDir: root, port: 0, maxBodyBytes: 64 })
    const res = await post(router, activity({ text: 'x'.repeat(200) }))
    expect(res.status).toBe(413)
    expect(shared.seen).toHaveLength(0)
  })
})

describe('failures and logging', () => {
  test('a dead session gets 503 and one log line with no text, ids or token', async () => {
    const dead = deadPort()
    writeWorkers(root, shared.port, [{ code: 'an', owner: AN, port: dead }])
    router.stop()
    router = startRouter({ rootDir: root, port: 0 })
    const marker = 'secret-message-text-4242'
    const res = await post(
      router,
      activity({ text: marker, conversation: { id: 'a:conv-xyz' }, from: { id: '29:user-xyz', aadObjectId: AN } }),
    )
    expect(res.status).toBe(503)
    const log = readFileSync(join(root, ROUTER_LOG), 'utf8')
    expect(log.trim().split('\n')).toHaveLength(1)
    expect(log).toMatch(new RegExp(`port=${dead} outcome=unreachable type=message`))
    for (const s of [marker, 'conv-xyz', 'user-xyz', AN, 'eyJ', 'Bearer']) expect(log).not.toContain(s)
  })

  test('a slow session times out to 503', async () => {
    const slow = fakeSession(async () => {
      await Bun.sleep(1000)
      return new Response(null)
    })
    servers.push(slow.server)
    writeWorkers(root, slow.port, [])
    router.stop()
    router = startRouter({ rootDir: root, port: 0, timeoutMs: 150 })
    const res = await post(router, activity({ type: 'evil type\n<script>' }))
    expect(res.status).toBe(503)
    const log = readFileSync(join(root, ROUTER_LOG), 'utf8')
    expect(log).toContain('outcome=timeout type=other')
    expect(log).not.toContain('script')
  })
})

describe('workers.json reload', () => {
  test('not re-read within the interval, re-read after it when the file changed', () => {
    let now = 1_000_000
    const cache = new WorkersCache(root, 5000, () => now)
    expect(cache.get().workers.map(w => w.code)).toEqual(['an', 'binh'])
    writeWorkers(root, shared.port, [{ code: 'an', owner: AN, port: an.port }])
    now += 4999
    expect(cache.get().workers).toHaveLength(2)
    now += 1
    expect(cache.get().workers.map(w => w.code)).toEqual(['an'])
  })

  test('a new worker starts receiving its owner\'s messages without a restart', async () => {
    router.stop()
    router = startRouter({ rootDir: root, port: 0, reloadMs: 0 })
    await post(router, activity({ from: { aadObjectId: OTHER } }))
    expect(shared.seen).toHaveLength(1)
    const other = fakeSession()
    servers.push(other.server)
    writeWorkers(root, shared.port, [
      { code: 'an', owner: AN, port: an.port },
      { code: 'binh', owner: BINH, port: binh.port },
      { code: 'chi', owner: OTHER, port: other.port },
    ])
    await post(router, activity({ from: { aadObjectId: OTHER } }))
    expect(other.seen).toHaveLength(1)
    expect(shared.seen).toHaveLength(1)
  })

  test('no workers.json: everything goes to the default port', () => {
    const empty = mkdtempSync(join(tmpdir(), 'router-empty-'))
    const cache = new WorkersCache(empty)
    expect(cache.get()).toEqual({ defaultPort: 4000, workers: [] })
    expect(existsSync(join(empty, 'logs'))).toBe(false)
  })
})

describe('other requests', () => {
  test('GET /healthz reports the worker count; anything else is 404', async () => {
    const base = `http://127.0.0.1:${router.server.port}`
    const h = await fetch(`${base}/healthz`)
    expect(h.status).toBe(200)
    expect(await h.json()).toEqual({ ok: true, workers: 2 })
    expect((await fetch(`${base}/api/messages`)).status).toBe(404)
    expect((await fetch(`${base}/other`, { method: 'POST', body: '{}' })).status).toBe(404)
    expect(shared.seen.length + an.seen.length + binh.seen.length).toBe(0)
  })
})
