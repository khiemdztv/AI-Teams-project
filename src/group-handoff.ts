/**
 * Fork (18/11): let a person approve in their 1:1 chat what the bot then posts
 * in the group they tagged it from.
 *
 * A group chat is otherwise answered only from a turn that started in it (see
 * the outbound gate in server.ts), so a 1:1 turn can never post a person's
 * mail or calendar into a group. But the bot sends its previews to the 1:1
 * chat, and the "ok" comes back there, from a turn that may not post the
 * result in the group.
 *
 * So when someone tags the bot in a group, that group is handed to them for
 * HANDOFF_MS: a turn started by the same person in their own 1:1 chat may post
 * in that group (and only that group) until then. Nobody else's turn, no other
 * group, no scheduled job. Kept in memory: a restart forgets every handoff,
 * which only means tagging the bot in the group again.
 */

export const HANDOFF_MS = 60 * 60 * 1000

type Turnish = {
  conversationId: string
  senderId?: string
  scheduled?: unknown
}

export class GroupHandoffs {
  private readonly bySender = new Map<string, { groupId: string; until: number }>()

  /** Someone tagged the bot in `groupId`: hand it to them. The latest group wins. */
  note(senderId: string, groupId: string, now = Date.now()): void {
    if (!senderId || !groupId) return
    this.bySender.set(senderId.toLowerCase(), { groupId, until: now + HANDOFF_MS })
  }

  /**
   * Whether `turn`, which did not start in `groupId`, may post there.
   * `dmOf` is the conversation id of the turn sender's own 1:1 chat.
   */
  allows(turn: Turnish | null, groupId: string, dmOf: string | undefined, now = Date.now()): boolean {
    if (!turn || turn.scheduled || !turn.senderId) return false
    if (!dmOf || turn.conversationId !== dmOf) return false
    const handoff = this.bySender.get(turn.senderId.toLowerCase())
    if (!handoff || handoff.until <= now) return false
    return handoff.groupId === groupId
  }
}
