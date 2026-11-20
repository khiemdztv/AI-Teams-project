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
  statusTarget,
  statusConversationOf,
  markReplied,
  repliedSince,
  shouldNudgeReply,
  NO_REPLY_NUDGE,
  newTurnId,
  mentionsTurn,
  transcriptHasTurn,
  stopIsForOtherTurn,
  TRANSCRIPT_TAIL_BYTES,
} from '../src/turn-lock.js'
import { spawnSync } from 'child_process'

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
    const since = Date.now()
    writeTurnLock(dir, { since, conversationId: 'a:1' })
    markReplied(dir, since) // a turn that replied; see "no-reply nudge" below
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

describe('status target (18/11)', () => {
  test('a 1:1 chat or channel keeps its status lines', () => {
    expect(statusTarget('personal', 'a:1', undefined)).toBe('a:1')
    expect(statusTarget('channel', '19:c', 'a:dm')).toBe('19:c')
    expect(statusTarget(undefined, 'a:1', undefined)).toBe('a:1')
  })

  test("a group chat sends them to the sender's 1:1 chat, or nowhere", () => {
    expect(statusTarget('groupChat', '19:g', 'a:dm')).toBe('a:dm')
    expect(statusTarget('groupChat', '19:g', undefined)).toBeNull()
  })

  test('the lock remembers where they went', () => {
    expect(statusConversationOf({ since: 1, conversationId: 'a:1' })).toBe('a:1')
    expect(statusConversationOf({ since: 1, conversationId: '19:g', statusConversationId: 'a:dm' })).toBe('a:dm')
    expect(statusConversationOf({ since: 1, conversationId: '19:g', statusConversationId: null })).toBeNull()
  })
})

describe('no-reply nudge (18/11)', () => {
  test('nudges once when a Teams turn ended without reply', () => {
    const lock = { since: 1000, conversationId: 'a:1' }
    expect(shouldNudgeReply(lock, null, false)).toBe(true)
    expect(shouldNudgeReply(lock, 999, false)).toBe(true) // an older turn's reply
    expect(shouldNudgeReply(lock, 1000, false)).toBe(false)
    expect(shouldNudgeReply(lock, null, true)).toBe(false) // the stop after a nudge
    expect(shouldNudgeReply(null, null, false)).toBe(false) // not a Teams turn
  })

  test('markReplied round-trips', () => {
    const dir = mkdtempSync(join(tmpdir(), 'replied-'))
    expect(repliedSince(dir)).toBeNull()
    markReplied(dir, 1234)
    expect(repliedSince(dir)).toBe(1234)
  })

  const runHook = (dir: string, input: object) =>
    spawnSync('bun', [join(ROOT, 'hooks/turn-done.ts')], {
      input: JSON.stringify(input),
      env: { ...process.env, MSTEAMS_STATE_DIR: dir },
      encoding: 'utf8',
    })

  test('the hook blocks the first stop and keeps the lock, then releases it', () => {
    const dir = mkdtempSync(join(tmpdir(), 'hook-'))
    writeTurnLock(dir, { since: 5, conversationId: 'a:1' })
    const first = runHook(dir, { hook_event_name: 'Stop', stop_hook_active: false })
    expect(first.status).toBe(0)
    expect(JSON.parse(first.stdout)).toEqual({ decision: 'block', reason: NO_REPLY_NUDGE })
    expect(existsSync(turnLockPath(dir))).toBe(true)
    const second = runHook(dir, { hook_event_name: 'Stop', stop_hook_active: true })
    expect(second.stdout).toBe('')
    expect(existsSync(turnLockPath(dir))).toBe(false)
  })

  test('the hook lets a replied turn, or a failed one, end at once', () => {
    const dir = mkdtempSync(join(tmpdir(), 'hook-'))
    writeTurnLock(dir, { since: 7, conversationId: 'a:1' })
    markReplied(dir, 7)
    expect(runHook(dir, { hook_event_name: 'Stop', stop_hook_active: false }).stdout).toBe('')
    expect(existsSync(turnLockPath(dir))).toBe(false)
    writeTurnLock(dir, { since: 8, conversationId: 'a:1' })
    expect(runHook(dir, { hook_event_name: 'StopFailure' }).stdout).toBe('')
    expect(existsSync(turnLockPath(dir))).toBe(false)
  })
})

describe('turn id (20/11)', () => {
  test('8 random hex chars', () => {
    const a = newTurnId()
    expect(a).toMatch(/^[0-9a-f]{8}$/)
    expect(newTurnId()).not.toBe(a)
  })

  test('found in the channel tag, whether or not the JSONL escaped its quotes', () => {
    expect(mentionsTurn('<channel source="msteams" turn_id="0a1b2c3d">', '0a1b2c3d')).toBe(true)
    expect(mentionsTurn(JSON.stringify({ content: '<channel turn_id="0a1b2c3d">hi</channel>' }), '0a1b2c3d')).toBe(true)
    expect(mentionsTurn('turn_id="0a1b2c3e"', '0a1b2c3d')).toBe(false)
    expect(mentionsTurn('0a1b2c3d', '0a1b2c3d')).toBe(false)
    expect(mentionsTurn('turn_id=".*"', '.*')).toBe(false) // not an id: never a pattern
  })

  test('read from the transcript tail, and further back when a long turn pushed it out', () => {
    const path = join(dir, 'transcript.jsonl')
    const line = (s: string) => JSON.stringify({ type: 'user', message: { content: s } }) + '\n'
    writeFileSync(path, line('<channel turn_id="0a1b2c3d">x</channel>') + 'y'.repeat(TRANSCRIPT_TAIL_BYTES * 3) + '\n')
    expect(transcriptHasTurn(path, '0a1b2c3d')).toBe(true)
    expect(transcriptHasTurn(path, 'ffffffff')).toBe(false)
    expect(transcriptHasTurn(join(dir, 'missing.jsonl'), '0a1b2c3d')).toBeNull()
  })

  test('the Stop hook leaves the lock only when the lock names a turn the transcript lacks', () => {
    const path = join(dir, 't.jsonl')
    writeFileSync(path, JSON.stringify({ content: '<channel turn_id="0a1b2c3d">' }) + '\n')
    const lock = { since: 1, conversationId: 'a:1', turnId: '0a1b2c3d' }
    expect(stopIsForOtherTurn(lock, path)).toBe(false)
    expect(stopIsForOtherTurn({ ...lock, turnId: 'ffffffff' }, path)).toBe(true)
    // As before 20/11: no turnId, no transcript path, or an unreadable transcript.
    expect(stopIsForOtherTurn({ since: 1, conversationId: 'a:1' }, path)).toBe(false)
    expect(stopIsForOtherTurn({ ...lock, turnId: 'ffffffff' }, undefined)).toBe(false)
    expect(stopIsForOtherTurn({ ...lock, turnId: 'ffffffff' }, join(dir, 'missing.jsonl'))).toBe(false)
    expect(stopIsForOtherTurn(null, path)).toBe(false)
  })

  const runHook = (input: object) =>
    spawnSync('bun', [join(ROOT, 'hooks/turn-done.ts')], {
      input: JSON.stringify(input),
      env: { ...process.env, MSTEAMS_STATE_DIR: dir },
      encoding: 'utf8',
    })

  test("a wake-up turn's Stop neither clears nor nudges; the Teams turn's Stop then does", () => {
    const path = join(dir, 'session.jsonl')
    writeFileSync(path, JSON.stringify({ content: 'a background task finished' }) + '\n')
    writeTurnLock(dir, { since: 9, conversationId: 'a:1', turnId: '0a1b2c3d' })
    const early = runHook({ hook_event_name: 'Stop', stop_hook_active: false, transcript_path: path })
    expect(early.status).toBe(0)
    expect(early.stdout).toBe('')
    expect(existsSync(turnLockPath(dir))).toBe(true)
    const failed = runHook({ hook_event_name: 'StopFailure', transcript_path: path })
    expect(failed.stdout).toBe('')
    expect(existsSync(turnLockPath(dir))).toBe(true)

    writeFileSync(path, JSON.stringify({ content: '<channel source="msteams" turn_id="0a1b2c3d">hi</channel>' }) + '\n', { flag: 'a' })
    markReplied(dir, 9)
    const done = runHook({ hook_event_name: 'Stop', stop_hook_active: false, transcript_path: path })
    expect(done.stdout).toBe('')
    expect(existsSync(turnLockPath(dir))).toBe(false)
  })

  test('the nudge still applies to the Teams turn itself', () => {
    const path = join(dir, 'session.jsonl')
    writeFileSync(path, JSON.stringify({ content: '<channel turn_id="0a1b2c3d">hi</channel>' }) + '\n')
    writeTurnLock(dir, { since: 10, conversationId: 'a:1', turnId: '0a1b2c3d' })
    const first = runHook({ hook_event_name: 'Stop', stop_hook_active: false, transcript_path: path })
    expect(JSON.parse(first.stdout)).toEqual({ decision: 'block', reason: NO_REPLY_NUDGE })
    expect(existsSync(turnLockPath(dir))).toBe(true)
  })
})
