/**
 * Fork (20/11): one Claude Code session per person, behind one router.
 *
 * `workers.json` sits in the root state dir (the shared one, `MSTEAMS_ROOT_DIR`,
 * default `~/.claude/channels/msteams`). It names each person's own session and
 * the local port its listener serves; everyone else goes to the shared session
 * on `defaultPort`. Read by the router (where to forward an activity) and by
 * every session (which scheduled jobs are its own). Written only on the bot
 * machine, by deploy/linux/add-bot-user.sh, never from Teams.
 *
 *   {
 *     "defaultPort": 4000,
 *     "workers": [
 *       { "code": "kien", "owner": "<Entra object id>", "port": 4001 }
 *     ]
 *   }
 *
 * No file, or a file that does not parse: no workers, every activity goes to
 * the shared session, which is how the bot ran before 20/11. Pure apart from
 * the one read in readWorkers().
 */

import { readFileSync } from 'fs'
import { join } from 'path'

export const WORKERS_FILE = 'workers.json'

export type Worker = {
  /** Short lowercase name: the systemd instance (claude-bot@<code>) and the state subdir. */
  code: string
  /** The person's Entra object id, lowercase. */
  owner: string
  /** Local port of that session's listener. */
  port: number
}

export type Workers = {
  defaultPort: number
  workers: Worker[]
}

export const DEFAULT_SHARED_PORT = 4000

const CODE = /^[a-z][a-z0-9-]{0,23}$/
const OBJECT_ID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/

function validPort(p: unknown): p is number {
  return typeof p === 'number' && Number.isInteger(p) && p >= 1024 && p <= 65535
}

/**
 * Parse and validate. Entries that are malformed, or repeat a code, owner or
 * port already taken, are dropped and reported in `errors`; the rest stand.
 */
export function parseWorkers(raw: unknown): { workers: Workers; errors: string[] } {
  const errors: string[] = []
  const obj = (raw && typeof raw === 'object' ? raw : {}) as Record<string, unknown>
  const defaultPort = validPort(obj.defaultPort) ? obj.defaultPort : DEFAULT_SHARED_PORT
  if (obj.defaultPort !== undefined && !validPort(obj.defaultPort)) errors.push('defaultPort is not a port; using 4000')
  const out: Worker[] = []
  const list = Array.isArray(obj.workers) ? obj.workers : []
  for (const [i, w] of list.entries()) {
    const e = (w && typeof w === 'object' ? w : {}) as Record<string, unknown>
    const code = String(e.code ?? '')
    const owner = String(e.owner ?? '').toLowerCase()
    if (!CODE.test(code)) {
      errors.push(`workers[${i}]: code must be lowercase letters, digits or "-"`)
      continue
    }
    if (!OBJECT_ID.test(owner)) {
      errors.push(`workers[${i}] (${code}): owner is not an Entra object id`)
      continue
    }
    if (!validPort(e.port) || e.port === defaultPort) {
      errors.push(`workers[${i}] (${code}): port must be a free port other than defaultPort`)
      continue
    }
    if (out.some(x => x.code === code || x.owner === owner || x.port === e.port)) {
      errors.push(`workers[${i}] (${code}): code, owner or port already used`)
      continue
    }
    out.push({ code, owner, port: e.port })
  }
  return { workers: { defaultPort, workers: out }, errors }
}

/** Read `workers.json` from the root state dir. Missing or unreadable: no workers. */
export function readWorkers(rootDir: string): { workers: Workers; errors: string[] } {
  let raw: unknown
  try {
    raw = JSON.parse(readFileSync(join(rootDir, WORKERS_FILE), 'utf8'))
  } catch (err: any) {
    if (err?.code === 'ENOENT') return { workers: { defaultPort: DEFAULT_SHARED_PORT, workers: [] }, errors: [] }
    return { workers: { defaultPort: DEFAULT_SHARED_PORT, workers: [] }, errors: ['workers.json does not parse; no per-person sessions'] }
  }
  return parseWorkers(raw)
}

/** The session that owns `senderId`, if that person has one. */
export function workerOf(workers: Workers, senderId: string | undefined): Worker | undefined {
  const id = String(senderId ?? '').toLowerCase()
  if (!id) return undefined
  return workers.workers.find(w => w.owner === id)
}

/**
 * Whether the session this process belongs to runs scheduled job `user`.
 * A person's own session runs only that person's jobs; the shared session
 * (no owner) runs the jobs of everyone who has no session of their own.
 */
export function ownsJob(workers: Workers, owner: string | undefined, user: string): boolean {
  const u = user.toLowerCase()
  if (owner) return u === owner.toLowerCase()
  return !workerOf(workers, u)
}
