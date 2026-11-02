/**
 * Stop / StopFailure hook: the Claude turn has ended, so release the turn
 * lock. The listener sees the lock file gone and marks the Teams status line
 * "✅ Xong". See src/turn-lock.ts.
 *
 * Must never fail the turn: any error is swallowed and the exit code is 0.
 */

import { clearTurnLock } from '../src/turn-lock.js'
import { stateDir } from '../src/state.js'

try {
  clearTurnLock(stateDir())
} catch {}
process.exit(0)
