/** Fork (20/11): messages that arrive mid-turn wait in line (src/turn-queue.ts). */

import { describe, expect, test } from 'bun:test'
import { TurnQueue, formatQueuedNotice, MAX_QUEUED, MAX_QUEUED_PER_SENDER } from '../src/turn-queue.js'

describe('TurnQueue', () => {
  test('first come, first served across senders, with 1-based positions', () => {
    const q = new TurnQueue<string>()
    expect(q.offer('an', 'a1')).toEqual({ ok: true, position: 1 })
    expect(q.offer('binh', 'b1')).toEqual({ ok: true, position: 2 })
    expect(q.offer('an', 'a2')).toEqual({ ok: true, position: 3 })
    expect(q.size).toBe(3)
    expect([q.shift(), q.shift(), q.shift(), q.shift()]).toEqual(['a1', 'b1', 'a2', undefined])
    expect(q.size).toBe(0)
  })

  test('at most 3 from one sender; a slot frees once one of theirs starts', () => {
    const q = new TurnQueue<number>()
    for (let i = 0; i < MAX_QUEUED_PER_SENDER; i++) expect(q.offer('an', i).ok).toBe(true)
    expect(q.offer('an', 9)).toEqual({ ok: false, reason: 'sender_full' })
    expect(q.offer('binh', 9).ok).toBe(true)
    q.shift()
    expect(q.offer('an', 10).ok).toBe(true)
  })

  test('at most 10 in total', () => {
    const q = new TurnQueue<number>()
    for (let i = 0; i < MAX_QUEUED; i++) expect(q.offer(`s${i}`, i).ok).toBe(true)
    expect(q.offer('late', 99)).toEqual({ ok: false, reason: 'full' })
    expect(q.size).toBe(MAX_QUEUED)
  })
})

describe('queued notice', () => {
  test('says it was received, and the place in line only past the first', () => {
    expect(formatQueuedNotice(1)).toBe('⏳ Đã nhận. Em đang làm việc trước, xong sẽ làm tin này.')
    expect(formatQueuedNotice(1, 'admin')).not.toContain('thứ')
    expect(formatQueuedNotice(3, 'admin')).toContain('thứ 3')
  })

  test('a non-admin gets assistant wording without icons', () => {
    for (const text of [formatQueuedNotice(1, 'user'), formatQueuedNotice(2, 'user')]) {
      expect(text.startsWith('Dạ,')).toBe(true)
      expect(text).not.toMatch(/[⏳✅⚠️]|Claude|bot|Xong/)
    }
  })
})
