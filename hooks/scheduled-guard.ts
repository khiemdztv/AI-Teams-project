/**
 * PreToolUse hook (fork addition, 05/11): during a scheduled job, refuse every
 * tool that is not on the read-only list in src/scheduled-guard.ts, whatever
 * server or built-in it belongs to. Outside a scheduled turn it says nothing
 * and every tool goes on to the usual permission rules.
 *
 * Fails closed only inside a scheduled turn: there, input it cannot read is a
 * refusal. Anywhere else it never stands in the way.
 */

import { readTurnLock } from '../src/turn-lock.js'
import { stateDir } from '../src/state.js'
import { scheduledToolAllowed, SCHEDULED_REFUSAL } from '../src/scheduled-guard.js'

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
if (!scheduled) process.exit(0)

let toolName = ''
try {
  toolName = String(JSON.parse(await Bun.stdin.text())?.tool_name ?? '')
} catch {}
if (!toolName) deny(SCHEDULED_REFUSAL)
if (!scheduledToolAllowed(toolName)) deny(SCHEDULED_REFUSAL)
process.exit(0)
