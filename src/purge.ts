/**
 * Age-based cleanup for `inbox/` (fork addition).
 *
 * Upstream never deleted inbound downloads, so every file the sender attached —
 * client documents included — stayed in the state dir indefinitely. Anything
 * worth keeping belongs in the matter's SharePoint folder, not here.
 */

import { readdirSync, rmSync, statSync } from 'fs'
import { join } from 'path'

/** How long an inbound download is kept. */
export const INBOX_TTL_MS = 7 * 24 * 60 * 60 * 1000

/** Delete regular files in `dir` last modified before `now - maxAgeMs`. Returns how many. */
export function purgeOlderThan(dir: string, maxAgeMs: number, now = Date.now()): number {
  let names: string[]
  try {
    names = readdirSync(dir)
  } catch {
    return 0
  }
  let removed = 0
  for (const name of names) {
    const path = join(dir, name)
    try {
      const stat = statSync(path)
      if (stat.isFile() && now - stat.mtimeMs > maxAgeMs) {
        rmSync(path, { force: true })
        removed++
      }
    } catch {
      // Gone already, or locked by another process: try again next sweep.
    }
  }
  return removed
}
