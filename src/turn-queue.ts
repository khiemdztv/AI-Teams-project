/**
 * Fork (20/11): Teams messages waiting for the turn in progress.
 *
 * Before 20/11 a message that arrived while Claude was busy got a "busy" notice
 * and was dropped, so the sender had to watch for "✅ Xong" and send it again.
 * Now it waits here, first come first served across everyone, and the listener
 * (server.ts) starts it when the running turn ends. The sender is told once that
 * it was received; that same status message later becomes the turn's "⏳" line.
 *
 * In memory only. The ingress queue entry (src/queue.ts) of a waiting message is
 * finished only once the message has been handed to Claude or dropped, so a
 * restart replays it from disk. Pure: no I/O, no clock.
 */

/** At most this many messages wait in total. */
export const MAX_QUEUED = 10
/** And at most this many from one sender, so one person cannot fill the line. */
export const MAX_QUEUED_PER_SENDER = 3

export type QueueOffer = { ok: true; position: number } | { ok: false; reason: 'full' | 'sender_full' }

export class TurnQueue<T> {
  private readonly items: { senderId: string; item: T }[] = []

  constructor(
    private readonly max = MAX_QUEUED,
    private readonly perSender = MAX_QUEUED_PER_SENDER,
  ) {}

  /** Add to the back of the line; `position` counts from 1. Refused over either limit. */
  offer(senderId: string, item: T): QueueOffer {
    if (this.items.length >= this.max) return { ok: false, reason: 'full' }
    if (this.items.filter(e => e.senderId === senderId).length >= this.perSender) {
      return { ok: false, reason: 'sender_full' }
    }
    this.items.push({ senderId, item })
    return { ok: true, position: this.items.length }
  }

  /** The oldest waiting message, removed from the line. */
  shift(): T | undefined {
    return this.items.shift()?.item
  }

  get size(): number {
    return this.items.length
  }
}

/**
 * What the sender is told when their message has to wait. `role` is theirs:
 * a non-admin gets assistant wording without icons, as for every status line
 * (src/turn-lock.ts).
 */
export function formatQueuedNotice(position: number, role?: 'admin' | 'user'): string {
  const base =
    role === 'user'
      ? 'Dạ, em đã nhận. Em đang làm việc trước, xong sẽ làm tin này.'
      : '⏳ Đã nhận. Em đang làm việc trước, xong sẽ làm tin này.'
  return position > 1 ? `${base} Tin này đứng thứ ${position} trong hàng chờ.` : base
}
