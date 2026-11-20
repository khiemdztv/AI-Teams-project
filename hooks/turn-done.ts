/**
 * Stop / StopFailure hook: the Claude turn has ended, so release the turn
 * lock. The listener sees the lock file gone and marks the Teams status line
 * "✅ Xong". See src/turn-lock.ts.
 *
 * Fork (18/11): on Stop, a Teams turn that never called `reply` is sent back
 * once to reply (decision "block"), keeping the lock; the stop after that, or
 * any StopFailure, always releases it. See shouldNudgeReply().
 *
 * Fork (20/11): a turn that is not the Teams turn the lock names (its turn_id
 * is not in the transcript yet, e.g. a background task's wake-up turn ran
 * before Claude picked the message up) leaves the lock alone: no clearing, no
 * nudge. See stopIsForOtherTurn().
 *
 * Must never fail the turn: any error is swallowed and the exit code is 0.
 */

import {
  clearTurnLock,
  readTurnLock,
  repliedSince,
  shouldNudgeReply,
  stopIsForOtherTurn,
  NO_REPLY_NUDGE,
} from '../src/turn-lock.js'
import { stateDir } from '../src/state.js'

let input: { hook_event_name?: string; stop_hook_active?: boolean; transcript_path?: unknown } = {}
try {
  input = JSON.parse((await Bun.stdin.text()) || '{}')
} catch {}

try {
  if (stopIsForOtherTurn(readTurnLock(stateDir()), input.transcript_path)) process.exit(0)
} catch {}

try {
  const dir = stateDir()
  if (
    input.hook_event_name === 'Stop' &&
    shouldNudgeReply(readTurnLock(dir), repliedSince(dir), input.stop_hook_active === true)
  ) {
    process.stdout.write(JSON.stringify({ decision: 'block', reason: NO_REPLY_NUDGE }))
    process.exit(0)
  }
} catch {}

try {
  clearTurnLock(stateDir())
} catch {}
process.exit(0)
