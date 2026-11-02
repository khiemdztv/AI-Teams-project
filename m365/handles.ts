/**
 * Short stand-ins for Graph ids (fork addition, 29/10).
 *
 * Outlook ids are ~150 characters of base64 with long repeated runs, e.g.
 * "...AAAAAAEPAAAfxX0vmQY2Rrtc0JiFLCSvAAAAAYuXAAA=". Claude copied one back
 * with a run missing and the send failed with "Id is malformed". The tools
 * therefore show a handle like "d_1a2b3c4d5e" and keep the id here. The same
 * id always gets the same handle (a hash), so a handle shown before a restart
 * can still be matched against a fresh listing.
 */

import { createHash } from 'crypto'

const MAX_ENTRIES = 20_000

export function hashOf(id: string): string {
  return createHash('sha256').update(id).digest('hex').slice(0, 10)
}

export class Handles {
  private readonly byHandle = new Map<string, string>()

  constructor(private readonly prefix: string) {}

  /** The handle for `id`, remembered so `full` can turn it back. */
  short(id: string): string {
    const handle = `${this.prefix}_${hashOf(id)}`
    this.byHandle.delete(handle)
    this.byHandle.set(handle, id)
    if (this.byHandle.size > MAX_ENTRIES) this.byHandle.delete(this.byHandle.keys().next().value!)
    return handle
  }

  /** Is `value` this kind of handle (known or not)? */
  isHandle(value: string): boolean {
    return new RegExp(`^${this.prefix}_[0-9a-f]{10}$`).test(String(value ?? '').trim())
  }

  /** Does `handle` stand for `id`? Works for handles from before a restart. */
  matches(handle: string, id: string): boolean {
    return String(handle ?? '').trim() === `${this.prefix}_${hashOf(id)}`
  }

  /** The id behind a handle, or undefined if this process never issued it. Anything else passes through as an id. */
  lookup(value: string): string | undefined {
    const v = String(value ?? '').trim()
    if (this.isHandle(v)) return this.byHandle.get(v)
    return v || undefined
  }

  /** As `lookup`, but an unknown value is an error telling Claude what to do. */
  full(value: string, what: string, listWith: string): string {
    const id = this.lookup(value)
    if (id) return id
    if (!String(value ?? '').trim()) throw new Error(`${what} is required`)
    throw new Error(`unknown ${what} "${String(value).trim()}": use the ${what} exactly as ${listWith} printed it (like "${this.prefix}_1a2b3c4d5e"), or run ${listWith} again`)
  }
}
