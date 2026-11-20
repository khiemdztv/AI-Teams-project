/**
 * Fork (20/11): the router in front of the per-person sessions.
 *
 *   Teams → Azure Bot → dev tunnel → router :3978
 *     ├─ "y/n <code>" with <root>/permission-routes/<code> → the session that asked
 *     ├─ sender has a worker in workers.json               → that worker's port
 *     └─ anything else                                     → defaultPort (shared session)
 *
 * The router only forwards. It never validates, strips or logs the Bot
 * Framework JWT: each session's own SDK validates it, exactly as when Teams
 * reached it directly. It parses the body only to choose a port; the bytes go
 * on unchanged, and the session's status, body and content-type come back
 * unchanged (invoke activities, e.g. a file-consent Accept or a card action,
 * carry their answer in the response body).
 *
 * Everything in the body is unauthenticated here, so the body may only choose
 * among ports the operator configured: a worker's port, defaultPort, or a
 * permission-routes entry that names one of those two.
 *
 * Logging: one line per failed forward, with time, port, outcome and activity
 * type. Never message text, conversation or sender ids, tokens or headers.
 */

import { appendFileSync, chmodSync, mkdirSync, readFileSync, statSync } from 'fs'
import { join } from 'path'
import { MAX_BODY_BYTES } from './bun-adapter.js'
import { parseVerdict } from './permissions.js'
import { DEFAULT_SHARED_PORT, WORKERS_FILE, readWorkers, workerOf, type Workers } from './workers.js'

export const PERMISSION_ROUTES_DIR = 'permission-routes'
export const ROUTER_LOG = join('logs', 'router.log')
export const ROUTER_PID_FILE = 'router.pid'
export const WORKERS_RELOAD_MS = 5000
export const FORWARD_TIMEOUT_MS = 15_000

/** The root state dir. Fork (20/11): becomes sharedDir() once src/state.ts has it. */
export function routerRootDir(fallback: string): string {
  return process.env.MSTEAMS_ROOT_DIR ?? fallback
}

/**
 * The text a verdict is read from. server.ts reads `activity.text` as it comes
 * in a 1:1 chat (no mention there); in a group or channel the text starts with
 * the bot's <at> mention and may carry HTML. Routing only picks a session, which
 * then applies its own rules, so being lenient here costs nothing.
 */
export function verdictText(text: string): string {
  return text
    .replace(/^\s*(<at\b[^>]*>[\s\S]*?<\/at>\s*)+/i, '')
    .replace(/<[^>]*>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
}

/** Ports the body may route to: the shared session and every worker. */
function knownPorts(workers: Workers): Set<number> {
  return new Set([workers.defaultPort, ...workers.workers.map(w => w.port)])
}

/**
 * The port written by the session that relayed permission request `requestId`,
 * or undefined. The id comes from parseVerdict (five letters), so it cannot
 * leave the dir. Garbage, or a port no session listens on, is ignored.
 */
export function readRoutePort(rootDir: string, requestId: string, workers: Workers): number | undefined {
  if (!/^[a-z]{5}$/.test(requestId)) return undefined
  let raw: string
  try {
    raw = readFileSync(join(rootDir, PERMISSION_ROUTES_DIR, requestId), 'utf8').trim()
  } catch {
    return undefined
  }
  if (!/^\d{1,5}$/.test(raw)) return undefined
  const port = Number(raw)
  return knownPorts(workers).has(port) ? port : undefined
}

export type Route = { port: number; via: 'verdict' | 'worker' | 'default' }

/** Where an activity goes. `activity` is whatever JSON.parse gave, possibly not an object. */
export function routeFor(activity: unknown, workers: Workers, rootDir: string): Route {
  const a = (activity && typeof activity === 'object' ? activity : {}) as Record<string, any>
  if (typeof a.text === 'string') {
    const verdict = parseVerdict(verdictText(a.text))
    if (verdict) {
      const port = readRoutePort(rootDir, verdict.requestId, workers)
      if (port !== undefined) return { port, via: 'verdict' }
    }
  }
  const from = a.from && typeof a.from === 'object' ? a.from.aadObjectId : undefined
  const worker = typeof from === 'string' ? workerOf(workers, from) : undefined
  if (worker) return { port: worker.port, via: 'worker' }
  return { port: workers.defaultPort, via: 'default' }
}

/**
 * workers.json, re-read at most every `reloadMs`, and then only when the file
 * changed (mtime, size or inode: add-bot-user.sh replaces it by rename).
 */
export class WorkersCache {
  private workers: Workers = { defaultPort: DEFAULT_SHARED_PORT, workers: [] }
  private checkedAt = -Infinity
  private key: string | undefined

  constructor(
    private readonly rootDir: string,
    private readonly reloadMs = WORKERS_RELOAD_MS,
    private readonly now: () => number = Date.now,
    private readonly onErrors: (errors: string[]) => void = () => {},
  ) {}

  get(): Workers {
    const t = this.now()
    if (t - this.checkedAt < this.reloadMs) return this.workers
    this.checkedAt = t
    let key: string
    try {
      const s = statSync(join(this.rootDir, WORKERS_FILE))
      key = `${s.mtimeMs}:${s.size}:${s.ino}`
    } catch {
      key = 'missing'
    }
    if (key !== this.key) {
      this.key = key
      const { workers, errors } = readWorkers(this.rootDir)
      this.workers = workers
      if (errors.length) this.onErrors(errors)
    }
    return this.workers
  }
}

/** An activity type safe to log: letters only, or "other". The body is unauthenticated. */
function logType(activity: unknown): string {
  const t = activity && typeof activity === 'object' ? (activity as Record<string, unknown>).type : undefined
  return typeof t === 'string' && /^[A-Za-z]{1,32}$/.test(t) ? t : 'other'
}

/** One line to stderr and <root>/logs/router.log (dir 0700, file 0600). Never throws. */
export function logRouter(rootDir: string, fields: { port: number; outcome: string; type: string }): void {
  const line = `${new Date().toISOString()} port=${fields.port} outcome=${fields.outcome} type=${fields.type}\n`
  process.stderr.write(`msteams router: ${line}`)
  try {
    const dir = join(rootDir, 'logs')
    mkdirSync(dir, { recursive: true, mode: 0o700 })
    const file = join(rootDir, ROUTER_LOG)
    appendFileSync(file, line, { mode: 0o600 })
    if (process.platform !== 'win32') chmodSync(file, 0o600)
  } catch {}
}

const TOO_LARGE = Symbol('body too large')

/** Read the body as bytes, aborting past `max` (same reason as bun-adapter.ts: this runs before any JWT check). */
async function readBoundedBytes(req: Request, max: number): Promise<Uint8Array<ArrayBuffer>> {
  const reader = req.body?.getReader()
  if (!reader) return new Uint8Array()
  const chunks: Uint8Array[] = []
  let total = 0
  try {
    for (;;) {
      const { done, value } = await reader.read()
      if (done) break
      total += value.byteLength
      if (total > max) throw TOO_LARGE
      chunks.push(value)
    }
  } finally {
    reader.cancel().catch(() => {})
    reader.releaseLock()
  }
  return new Uint8Array(Buffer.concat(chunks))
}

export type RouterOptions = {
  rootDir: string
  /** Listen address. Default 127.0.0.1; port 0 picks a free one (tests). */
  hostname?: string
  port?: number
  /** The webhook path, served here and forwarded unchanged. */
  path?: string
  timeoutMs?: number
  reloadMs?: number
  maxBodyBytes?: number
}

export type Router = {
  server: ReturnType<typeof Bun.serve>
  workers: WorkersCache
  stop(): void
}

export function startRouter(opts: RouterOptions): Router {
  const rootDir = opts.rootDir
  const path = opts.path ?? '/api/messages'
  const timeoutMs = opts.timeoutMs ?? FORWARD_TIMEOUT_MS
  const maxBodyBytes = opts.maxBodyBytes ?? MAX_BODY_BYTES
  // Errors name codes only (never owners); report how many, not which.
  const workers = new WorkersCache(rootDir, opts.reloadMs ?? WORKERS_RELOAD_MS, Date.now, errors =>
    process.stderr.write(`msteams router: workers.json: ${errors.length} problem(s); affected entries ignored\n`),
  )

  const server = Bun.serve({
    hostname: opts.hostname ?? '127.0.0.1',
    port: opts.port ?? 3978,
    fetch: async req => {
      const url = new URL(req.url)
      if (req.method === 'GET' && url.pathname === '/healthz') {
        return Response.json({ ok: true, workers: workers.get().workers.length })
      }
      if (req.method !== 'POST' || url.pathname !== path) return new Response('not found', { status: 404 })

      if (Number(req.headers.get('content-length') ?? '0') > maxBodyBytes) {
        return new Response('payload too large', { status: 413 })
      }
      let body: Uint8Array<ArrayBuffer>
      try {
        body = await readBoundedBytes(req, maxBodyBytes)
      } catch (err) {
        if (err === TOO_LARGE) return new Response('payload too large', { status: 413 })
        return new Response('bad request', { status: 400 })
      }

      // Parse only to route. A body that does not parse still goes to the shared
      // session, whose SDK answers it as it always has.
      let activity: unknown
      try {
        activity = JSON.parse(new TextDecoder().decode(body))
      } catch {
        activity = undefined
      }
      const { port } = routeFor(activity, workers.get(), rootDir)

      const headers: Record<string, string> = {}
      for (const name of ['authorization', 'content-type']) {
        const v = req.headers.get(name)
        if (v !== null) headers[name] = v
      }
      try {
        const res = await fetch(`http://127.0.0.1:${port}${url.pathname}${url.search}`, {
          method: req.method,
          headers,
          body,
          redirect: 'manual',
          signal: AbortSignal.timeout(timeoutMs),
        })
        const out = await res.arrayBuffer()
        const resHeaders: Record<string, string> = {}
        const ct = res.headers.get('content-type')
        if (ct !== null) resHeaders['content-type'] = ct
        return new Response(out.byteLength ? out : null, { status: res.status, headers: resHeaders })
      } catch (err: any) {
        const outcome = err?.name === 'TimeoutError' || err?.name === 'AbortError' ? 'timeout' : 'unreachable'
        logRouter(rootDir, { port, outcome, type: logType(activity) })
        // 503 so Bot Framework retries; the session may be restarting.
        return new Response('session unavailable', { status: 503 })
      }
    },
  })

  return { server, workers, stop: () => server.stop(true) }
}
