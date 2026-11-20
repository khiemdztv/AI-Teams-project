import { describe, expect, test } from 'bun:test'
import { GroupHandoffs, HANDOFF_MS } from '../src/group-handoff.js'

const SAM = 'aaaa-1111'
const JIMMY = 'bbbb-2222'
const DM_NHAN = 'a:dm-nhan'
const GROUP = '19:group@thread.v2'

describe('group handoff (18/11)', () => {
  test("the tagger's own 1:1 turn may post in the group they tagged from", () => {
    const h = new GroupHandoffs()
    h.note(SAM, GROUP, 0)
    expect(h.allows({ conversationId: DM_NHAN, senderId: SAM }, GROUP, DM_NHAN, 1000)).toBe(true)
    // sender ids compare case-insensitively
    expect(h.allows({ conversationId: DM_NHAN, senderId: SAM.toUpperCase() }, GROUP, DM_NHAN, 1000)).toBe(true)
  })

  test('nobody else, no other group, no other chat, no job, no turn', () => {
    const h = new GroupHandoffs()
    h.note(SAM, GROUP, 0)
    expect(h.allows({ conversationId: 'a:dm-jimmy', senderId: JIMMY }, GROUP, 'a:dm-jimmy', 1)).toBe(false)
    expect(h.allows({ conversationId: DM_NHAN, senderId: SAM }, '19:other@thread.v2', DM_NHAN, 1)).toBe(false)
    // a turn of Nhan's that is not in his own 1:1 chat
    expect(h.allows({ conversationId: '19:third@thread.v2', senderId: SAM }, GROUP, DM_NHAN, 1)).toBe(false)
    expect(h.allows({ conversationId: DM_NHAN, senderId: SAM }, GROUP, undefined, 1)).toBe(false)
    expect(h.allows({ conversationId: DM_NHAN, senderId: SAM, scheduled: 'morning' }, GROUP, DM_NHAN, 1)).toBe(false)
    expect(h.allows({ conversationId: DM_NHAN }, GROUP, DM_NHAN, 1)).toBe(false)
    expect(h.allows(null, GROUP, DM_NHAN, 1)).toBe(false)
  })

  test('expires, and the latest tag wins', () => {
    const h = new GroupHandoffs()
    h.note(SAM, GROUP, 0)
    expect(h.allows({ conversationId: DM_NHAN, senderId: SAM }, GROUP, DM_NHAN, HANDOFF_MS - 1)).toBe(true)
    expect(h.allows({ conversationId: DM_NHAN, senderId: SAM }, GROUP, DM_NHAN, HANDOFF_MS)).toBe(false)
    h.note(SAM, GROUP, 0)
    h.note(SAM, '19:other@thread.v2', 10)
    expect(h.allows({ conversationId: DM_NHAN, senderId: SAM }, GROUP, DM_NHAN, 20)).toBe(false)
    expect(h.allows({ conversationId: DM_NHAN, senderId: SAM }, '19:other@thread.v2', DM_NHAN, 20)).toBe(true)
  })
})
