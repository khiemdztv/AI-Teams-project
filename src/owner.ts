/**
 * Fork (20/11): owner mode, for a Claude Code session that belongs to one person.
 *
 * With per-person sessions behind the router (src/workers.ts), a person's own
 * session runs with `BOT_OWNER` set to their Entra object id. Such a session
 * answers only that person: the router should never send it anyone else, and
 * these checks are the belt to that brace. Its m365 tools act in the owner's
 * mailbox, so its messages must not reach anyone else's 1:1 chat either.
 *
 * `BOT_OWNER` unset (the shared session, and the bot before 20/11): every
 * function here allows everything, so behaviour is exactly as before.
 *
 * Pure decisions, plus the small permission-route file the router reads to
 * send an admin's "y <code>" back to the session that asked.
 */

import { mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'fs'
import { join } from 'path'
import { normalizeSenderId } from './gate.js'

/** `BOT_OWNER`, normalized, or undefined when this is not a person's own session. */
export function ownerFromEnv(env: Record<string, string | undefined> = process.env): string | undefined {
  const raw = env.BOT_OWNER?.trim()
  return raw ? normalizeSenderId(raw) : undefined
}

/** Inbound: whether a sender the gate allowed is still refused because the session is someone else's. */
export function refusedNotOwner(owner: string | undefined, senderId: string): boolean {
  return !!owner && normalizeSenderId(senderId) !== owner
}

type Target = { conversationType: string; senderId?: string }

/** Whether `ref` is the owner's own 1:1 chat with the bot. */
export function isOwnersChat(owner: string, ref: Target): boolean {
  return ref.conversationType === 'personal' && normalizeSenderId(String(ref.senderId ?? '')) === owner
}

/**
 * Outbound, on top of the usual gate: a reason to refuse, or null to allow.
 * - A 1:1 chat is only ever the owner's own.
 * - With no Teams turn running (a background task's wake-up turn, say), only
 *   the owner's own 1:1 chat may be addressed at all.
 * Group chats otherwise keep the usual rules.
 */
export function ownerOutboundRefusal(owner: string | undefined, ref: Target, turnActive: boolean): string | null {
  if (!owner) return null
  if (ref.conversationType === 'personal' && !isOwnersChat(owner, ref)) {
    return "this session belongs to one person: it only writes to that person's own 1:1 chat"
  }
  if (!turnActive && !isOwnersChat(owner, ref)) {
    return "with no Teams message being answered, only the owner's own 1:1 chat can be addressed"
  }
  return null
}

// ---------------------------------------------------------------------------
// Permission routes
// ---------------------------------------------------------------------------

export const PERMISSION_ROUTES_DIR = 'permission-routes'

/**
 * Whether this session must register its relayed permission prompts for the
 * router: a person's own session (`BOT_WORKER`), or any session in the
 * multi-session layout (`MSTEAMS_ROOT_DIR`). Single-session bots write nothing.
 */
export function routesPermissions(env: Record<string, string | undefined> = process.env): boolean {
  return !!env.BOT_WORKER || !!env.MSTEAMS_ROOT_DIR
}

// Request ids are five letters today; anything that could leave the dir is refused.
const SAFE_ID = /^[A-Za-z0-9_-]{1,64}$/

/** `<rootDir>/permission-routes/<requestId>` holding `port` as decimal text, 0600 in a 0700 dir. */
export function writePermissionRoute(rootDir: string, requestId: string, port: number): boolean {
  if (!SAFE_ID.test(requestId)) return false
  try {
    const dir = join(rootDir, PERMISSION_ROUTES_DIR)
    mkdirSync(dir, { recursive: true, mode: 0o700 })
    writeFileSync(join(dir, requestId), String(port), { mode: 0o600 })
    return true
  } catch {
    return false
  }
}

export function removePermissionRoute(rootDir: string, requestId: string): void {
  if (!SAFE_ID.test(requestId)) return
  try {
    rmSync(join(rootDir, PERMISSION_ROUTES_DIR, requestId))
  } catch {}
}

/** At boot: routes this session (its `port`) left behind are for prompts that no longer exist. */
export function clearOwnPermissionRoutes(rootDir: string, port: number): number {
  const dir = join(rootDir, PERMISSION_ROUTES_DIR)
  let names: string[]
  try {
    names = readdirSync(dir)
  } catch {
    return 0
  }
  let n = 0
  for (const name of names) {
    try {
      if (readFileSync(join(dir, name), 'utf8').trim() === String(port)) {
        rmSync(join(dir, name))
        n++
      }
    } catch {}
  }
  return n
}
