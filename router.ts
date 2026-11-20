#!/usr/bin/env bun
/**
 * Fork (20/11): `bun router.ts` — the webhook router for one-session-per-person
 * mode. The tunnel points here (default 127.0.0.1:3978/api/messages); each
 * activity is forwarded to a session's own listener on 127.0.0.1. Logic and
 * the reasons for it live in src/router.ts.
 *
 * Reads no .env and holds no credentials: it never needs the bot's secret,
 * because it never checks or makes a token.
 *
 * Env: MSTEAMS_WEBHOOK_HOST (127.0.0.1), MSTEAMS_WEBHOOK_PORT (3978),
 * MSTEAMS_WEBHOOK_PATH (/api/messages), MSTEAMS_ROOT_DIR (root state dir,
 * default the msteams state dir).
 */

import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'fs'
import { join } from 'path'
import { ROUTER_PID_FILE, startRouter } from './src/router.js'
import { sharedDir } from './src/state.js'

const ROOT_DIR = sharedDir()
const HOST = process.env.MSTEAMS_WEBHOOK_HOST ?? '127.0.0.1'
const PORT = Number(process.env.MSTEAMS_WEBHOOK_PORT ?? 3978)
const PATH = process.env.MSTEAMS_WEBHOOK_PATH ?? '/api/messages'
const PID_FILE = join(ROOT_DIR, ROUTER_PID_FILE)

mkdirSync(ROOT_DIR, { recursive: true, mode: 0o700 })

const router = startRouter({ rootDir: ROOT_DIR, hostname: HOST, port: PORT, path: PATH })
writeFileSync(PID_FILE, String(process.pid), { mode: 0o600 })
process.stderr.write(
  `msteams router: listening on ${HOST}:${router.server.port}${PATH}, ${router.workers.get().workers.length} worker(s)\n`,
)

process.on('unhandledRejection', err => process.stderr.write(`msteams router: unhandled rejection: ${err}\n`))

let shuttingDown = false
function shutdown(): void {
  if (shuttingDown) return
  shuttingDown = true
  process.stderr.write('msteams router: shutting down\n')
  router.stop()
  // Only our own pidfile: a second router that failed to bind must not delete the live one's.
  try {
    if (parseInt(readFileSync(PID_FILE, 'utf8'), 10) === process.pid) rmSync(PID_FILE)
  } catch {}
  process.exit(0)
}
process.on('SIGTERM', shutdown)
process.on('SIGINT', shutdown)
process.on('SIGHUP', shutdown)
