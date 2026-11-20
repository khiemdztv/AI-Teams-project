/**
 * Whose mailbox a tool call acts in.
 *
 * The answer comes from the msteams channel, not from the model: when a
 * Teams message starts a turn, the listener writes the sender's Entra object
 * id (as its inbound gate verified it) into turn.json, and it stays there
 * until the turn ends. The model is never asked whose mailbox to use, so a
 * prompt injection cannot pick someone else's.
 *
 * No Teams turn, no mailbox: there is deliberately no default user, so a
 * call made outside a Teams turn (or after the lock timed out) is refused
 * rather than sent from the wrong person's mailbox. The one exception is a
 * per-person session (BOT_OWNER), where there is only one possible mailbox.
 */

import { join } from 'path'
import { readTurnLock } from '../src/turn-lock.js'
import { TokenStore, defaultProtector, type Protector } from './token-store.js'
import { GraphAuth, type Fetch } from './auth.js'

const OBJECT_ID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/

export const NO_TEAMS_SENDER =
  'refused: email can only be drafted or sent while answering a Teams message, so the sender (and so the mailbox) is known'

export const NOT_THE_OWNER =
  'refused: this session works for one person only, and the message being answered is from someone else, so no mailbox is used'

export const BAD_OWNER = 'refused: BOT_OWNER is set but is not an Entra object id, so no mailbox is used'

/**
 * Fork (20/11): the owner of a per-person session (BOT_OWNER, set by the
 * worker's env file), lower-cased; undefined in the shared session.
 */
export function botOwner(raw: string | undefined = process.env.BOT_OWNER): string | undefined {
  const id = raw?.trim().toLowerCase()
  return id ? id : undefined
}

/**
 * The Entra object id of the Teams sender whose turn is running.
 *
 * Fork (20/11): in a per-person session (BOT_OWNER set) the mailbox is always
 * the owner's. A turn started by anyone else is refused (the router should
 * never send one; this is the belt). With no Teams turn (a background task,
 * or the turn Claude Code runs when one finishes) the owner is still the
 * only possible answer, so their own mailbox is used; what may be done
 * there without their message is limited by the m365 server (see
 * M365_COMPOSE_TOOLS in src/subagent-guard.ts).
 */
export function currentSender(msteamsStateDir: string, owner: string | undefined = process.env.BOT_OWNER): string {
  const id = readTurnLock(msteamsStateDir)?.senderId?.toLowerCase()
  const ownerId = botOwner(owner)
  if (ownerId) {
    if (!OBJECT_ID.test(ownerId)) throw new Error(BAD_OWNER)
    if (!id) return ownerId
    if (id !== ownerId) throw new Error(NOT_THE_OWNER)
    return id
  }
  if (!id || !OBJECT_ID.test(id)) throw new Error(NO_TEAMS_SENDER)
  return id
}

export function tokenPath(tokenDir: string, objectId: string): string {
  if (!OBJECT_ID.test(objectId)) throw new Error('not an Entra object id')
  return join(tokenDir, `${objectId}.dpapi`)
}

/** One GraphAuth per person, so each access token is cached separately. */
export class UserAuths {
  private readonly byUser = new Map<string, GraphAuth>()

  constructor(
    private readonly clientId: string,
    private readonly tenantId: string,
    private readonly tokenDir: string,
    private readonly protector: Protector = defaultProtector(),
    private readonly fetchImpl: Fetch = fetch,
  ) {}

  for(objectId: string): GraphAuth {
    let auth = this.byUser.get(objectId)
    if (!auth) {
      const store = new TokenStore(tokenPath(this.tokenDir, objectId), this.protector)
      auth = new GraphAuth(this.clientId, this.tenantId, store, this.fetchImpl)
      this.byUser.set(objectId, auth)
    }
    return auth
  }
}
