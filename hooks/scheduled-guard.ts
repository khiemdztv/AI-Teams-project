/**
 * PreToolUse hook (fork addition, 05/11): during a scheduled job, refuse every
 * tool that is not on the read-only list in src/scheduled-guard.ts, whatever
 * server or built-in it belongs to.
 *
 * Fork (20/11): also holds subagents to what they may use (they prepare, the
 * main conversation sends) and refuses a background agent in the shared
 * session, where the m365 mailbox changes with every message. The decision
 * itself is src/subagent-guard.ts (preToolUseDecision); this file only reads
 * the input and writes the answer. Otherwise it says nothing and every tool
 * goes on to the usual permission rules.
 *
 * Fails closed only inside a scheduled turn: there, input it cannot read is a
 * refusal. Anywhere else it never stands in the way, and it never throws.
 */

import { readTurnLock } from '../src/turn-lock.js'
import { stateDir } from '../src/state.js'
import { SCHEDULED_REFUSAL } from '../src/scheduled-guard.js'
import { preToolUseDecision } from '../src/subagent-guard.js'

function deny(reason: string): never {
  process.stdout.write(
    JSON.stringify({ hookSpecificOutput: { hookEventName: 'PreToolUse', permissionDecision: 'deny', permissionDecisionReason: reason } }),
  )
  process.exit(0)
}

let scheduled = false
try {
  scheduled = !!readTurnLock(stateDir())?.scheduled
} catch {}

let input: unknown = null
try {
  input = JSON.parse(await Bun.stdin.text())
} catch {}

let reason: string | null = null
try {
  reason = preToolUseDecision(input, { scheduled, owner: !!process.env.BOT_OWNER?.trim() })
} catch {
  reason = scheduled ? SCHEDULED_REFUSAL : null
}
if (reason) deny(reason)
process.exit(0)
