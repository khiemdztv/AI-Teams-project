/**
 * Fork addition: one Teams request at a time. The listener claims the turn in
 * turn.json; the plugin's Stop hook, a separate process, clears it.
 */

import { describe, expect, test, beforeEach } from 'bun:test'
import { mkdtempSync, existsSync, readFileSync, writeFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import {
  readTurnLock,
  writeTurnLock,
  clearTurnLock,
  turnLockPath,
  isStale,
  formatBusyNotice,
  formatDone,
  formatWorking,
  STATUS_WORKING,
  STATUS_TIMED_OUT,
  USER_WORKING,
  formatTimedOut,
  MAX_TURN_MS,
} from '../src/turn-lock.js'

const ROOT = join(import.meta.dir, '..')

let dir: string
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'msteams-turn-'))
})

describe('turn lock file', () => {
  test('round trips, and clearing reports whether there was one', () => {
    expect(readTurnLock(dir)).toBeNull()
    writeTurnLock(dir, { since: 1000, conversationId: 'a:1' })
    expect(readTurnLock(dir)).toEqual({ since: 1000, conversationId: 'a:1' })
    expect(clearTurnLock(dir)).toBe(true)
    expect(clearTurnLock(dir)).toBe(false)
    expect(readTurnLock(dir)).toBeNull()
  })

  test('a corrupt file reads as no lock', () => {
    writeFileSync(turnLockPath(dir), '{not json')
    expect(readTurnLock(dir)).toBeNull()
  })

  test('a turn is stale only after MAX_TURN_MS', () => {
    const lock = { since: 0, conversationId: 'a:1' }
    expect(isStale(lock, MAX_TURN_MS)).toBe(false)
    expect(isStale(lock, MAX_TURN_MS + 1)).toBe(true)
  })
})

describe('status text', () => {
  test('the busy notice says when the running request started', () => {
    const since = Date.UTC(2026, 9, 27, 7, 27)
    const text = formatBusyNotice({ since, conversationId: 'a:1' }, 'Asia/Ho_Chi_Minh')
    expect(text).toContain('14:27')
    expect(text).toContain('✅ Xong')
  })

  test('done reports the duration', () => {
    const lock = { since: 0, conversationId: 'a:1' }
    expect(formatDone(lock, 29_400)).toContain('29 giây')
    expect(formatDone(lock, 125_000)).toContain('2 phút 5 giây')
    expect(formatDone(lock, 100)).toContain('1 giây')
  })

  test('the working line shows whole minutes once past the first', () => {
    const lock = { since: 0, conversationId: 'a:1' }
    expect(formatWorking(lock, 45_000)).toBe(STATUS_WORKING)
    expect(formatWorking(lock, 61_000)).toContain('đã 1 phút')
    expect(formatWorking(lock, 185_000)).toContain('đã 3 phút')
    expect(formatWorking(lock, 185_000)).toContain('✅ Xong')
  })
})

describe('Stop hook', () => {
  const runHook = () =>
    Bun.spawnSync(['bun', join(ROOT, 'hooks', 'turn-done.ts')], {
      env: { ...process.env, MSTEAMS_STATE_DIR: dir },
      stdin: new TextEncoder().encode('{"hook_event_name":"Stop"}'),
    })

  test('clears the lock and exits 0', () => {
    writeTurnLock(dir, { since: Date.now(), conversationId: 'a:1' })
    const result = runHook()
    expect(result.exitCode).toBe(0)
    expect(existsSync(turnLockPath(dir))).toBe(false)
  })

  test('with no lock it is a quiet no-op', () => {
    const result = runHook()
    expect(result.exitCode).toBe(0)
    expect(result.stderr.toString()).toBe('')
  })

  test('hooks.json runs the hook on Stop and StopFailure through the plugin root', () => {
    const config = JSON.parse(readFileSync(join(ROOT, 'hooks', 'hooks.json'), 'utf8'))
    for (const event of ['Stop', 'StopFailure']) {
      const hook = config.hooks[event][0].hooks[0]
      expect(hook.command).toBe('bun')
      expect(hook.args).toEqual(['${CLAUDE_PLUGIN_ROOT}/hooks/turn-done.ts'])
    }
  })
})

describe('status lines for a non-admin (fork, 02/11)', () => {
  const lock = { since: 0, conversationId: 'a:1', role: 'user' as const }
  const all = () => [
    formatWorking(lock, 10_000),
    formatWorking(lock, 185_000),
    formatDone(lock, 29_400),
    formatBusyNotice(lock, 'Asia/Ho_Chi_Minh', 'user'),
    formatTimedOut(lock),
  ]

  test('assistant wording, without icons, "Claude", "bot" or the ✅ protocol', () => {
    for (const text of all()) {
      expect(text.startsWith('Dạ,')).toBe(true)
      expect(text).not.toMatch(/[⏳✅⚠️]|Claude|bot|Xong/)
    }
  })

  test('the minutes still show while it works', () => {
    expect(formatWorking(lock, 10_000)).toBe(USER_WORKING)
    expect(formatWorking(lock, 185_000)).toContain('đã 3 phút')
  })

  test('the busy notice follows the role of whoever just wrote, not the turn', () => {
    const adminTurn = { since: 0, conversationId: 'a:1', role: 'admin' as const }
    expect(formatBusyNotice(adminTurn, 'Asia/Ho_Chi_Minh', 'user').startsWith('Dạ,')).toBe(true)
    expect(formatBusyNotice(lock, 'Asia/Ho_Chi_Minh', 'admin')).toContain('✅ Xong')
  })

  test('admins, and locks from before roles, keep the original lines', () => {
    for (const l of [{ since: 0, conversationId: 'a:1', role: 'admin' as const }, { since: 0, conversationId: 'a:1' }]) {
      expect(formatWorking(l, 10_000)).toBe(STATUS_WORKING)
      expect(formatDone(l, 29_400)).toContain('✅ Xong (29 giây)')
      expect(formatTimedOut(l)).toBe(STATUS_TIMED_OUT)
    }
  })
})
