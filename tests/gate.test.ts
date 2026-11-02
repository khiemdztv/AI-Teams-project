import { describe, expect, test } from 'bun:test'
import {
  gate,
  outboundAllowed,
  mentionsBot,
  normalizeConversationId,
  normalizeSenderId,
  DEFAULT_ACCESS,
  type Access,
  type GateInput,
  type GroupPolicy,
} from '../src/gate.js'

const TENANT = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
const SENDER = '00000000-1111-2222-3333-444444444444'
const BOT = '28:00000000-aaaa-bbbb-cccc-dddddddddddd'

const access = (over: Partial<Access> = {}): Access => ({ ...DEFAULT_ACCESS, ...over })

const dm = (over: Partial<GateInput> = {}): GateInput => ({
  tenantId: TENANT,
  conversationId: 'a:1conv',
  conversationType: 'personal',
  senderAadObjectId: SENDER,
  mentionsBot: false,
  ...over,
})

describe('conversation id normalization', () => {
  test('strips the thread messageid suffix', () => {
    expect(normalizeConversationId('19:abc@thread.tacv2;messageid=1700000000000')).toBe(
      '19:abc@thread.tacv2',
    )
  })

  test('leaves a plain id untouched', () => {
    expect(normalizeConversationId('a:1conv')).toBe('a:1conv')
  })

  test('preserves case — these ids are opaque and case-sensitive', () => {
    expect(normalizeConversationId('19:AbCdEf@thread.tacv2')).toBe('19:AbCdEf@thread.tacv2')
  })

  test('a thread and its parent resolve to the same conversation', () => {
    const parent = '19:abc@thread.tacv2'
    expect(normalizeConversationId(`${parent};messageid=1`)).toBe(
      normalizeConversationId(`${parent};messageid=2`),
    )
  })
})

describe('sender id normalization', () => {
  test('lowercases and trims (AAD object ids are case-insensitive)', () => {
    expect(normalizeSenderId('  ABCD-1234  ')).toBe('abcd-1234')
  })
})

describe('tenant boundary', () => {
  test('an activity from another tenant is refused', () => {
    const v = gate(dm({ tenantId: 'ffffffff-0000-0000-0000-000000000000' }), access({ allowFrom: [SENDER] }), TENANT)
    expect(v).toEqual({ allowed: false, reason: 'tenant_mismatch' })
  })

  test('a missing tenant is refused', () => {
    const v = gate(dm({ tenantId: undefined }), access({ allowFrom: [SENDER] }), TENANT)
    expect(v).toEqual({ allowed: false, reason: 'tenant_mismatch' })
  })

  test('tenant comparison is case-insensitive', () => {
    const v = gate(dm({ tenantId: TENANT.toUpperCase() }), access({ allowFrom: [SENDER] }), TENANT)
    expect(v.allowed).toBe(true)
  })

  test('tenant is checked before sender — a foreign tenant never reports sender_not_allowed', () => {
    const v = gate(dm({ tenantId: 'other', senderAadObjectId: undefined }), access(), TENANT)
    expect(v).toEqual({ allowed: false, reason: 'tenant_mismatch' })
  })
})

describe('sender identity', () => {
  test('an allowlisted sender passes in a DM', () => {
    expect(gate(dm(), access({ allowFrom: [SENDER] }), TENANT).allowed).toBe(true)
  })

  test('an unlisted sender is refused', () => {
    expect(gate(dm(), access(), TENANT)).toEqual({ allowed: false, reason: 'sender_not_allowed' })
  })

  test('a principal with no AAD object id is refused — names never grant access', () => {
    const v = gate(dm({ senderAadObjectId: undefined }), access({ allowFrom: [SENDER] }), TENANT)
    expect(v).toEqual({ allowed: false, reason: 'no_sender_identity' })
  })

  test('allowlist matching is case-insensitive on the sender id', () => {
    const v = gate(dm({ senderAadObjectId: SENDER.toUpperCase() }), access({ allowFrom: [SENDER] }), TENANT)
    expect(v.allowed).toBe(true)
  })
})

describe('dm policy', () => {
  test('disabled refuses even an allowlisted sender', () => {
    const v = gate(dm(), access({ dmPolicy: 'disabled', allowFrom: [SENDER] }), TENANT)
    expect(v).toEqual({ allowed: false, reason: 'dm_disabled' })
  })

  test('pairing still refuses an unknown sender until they are added', () => {
    const v = gate(dm(), access({ dmPolicy: 'pairing' }), TENANT)
    expect(v).toEqual({ allowed: false, reason: 'sender_not_allowed' })
  })

  test('disabled is a global kill switch, not just a DM setting', () => {
    // Matches discord, whose ACCESS.md states it plainly: "Drop everything,
    // including allowlisted users and guild channels." An operator reaching for
    // this wants the bot off; leaving channels answering would fail in the
    // dangerous direction.
    const off = access({
      dmPolicy: 'disabled',
      allowFrom: [SENDER],
      groups: { '19:abc@thread.tacv2': { requireMention: false, allowFrom: [] } },
    })
    expect(gate(dm(), off, TENANT)).toEqual({ allowed: false, reason: 'dm_disabled' })
    expect(
      gate(
        dm({
          conversationType: 'groupChat',
          conversationId: '19:abc@thread.tacv2',
          mentionsBot: true,
        }),
        off,
        TENANT,
      ),
    ).toEqual({ allowed: false, reason: 'dm_disabled' })
  })
})

describe('group and channel gating (fork: always refused)', () => {
  const group = (over: Partial<GateInput> = {}): GateInput =>
    dm({ conversationType: 'groupChat', conversationId: '19:abc@thread.tacv2', ...over })

  // The most permissive configuration upstream allowed: opted in, no mention
  // required, empty group allowFrom (= anyone in the room).
  const wideOpen = (policy: Partial<GroupPolicy> = {}) =>
    access({
      allowFrom: [SENDER],
      groups: { '19:abc@thread.tacv2': { requireMention: false, allowFrom: [], ...policy } },
    })

  test('a group chat is refused even when opted in and mentioned', () => {
    const v = gate(group({ mentionsBot: true }), wideOpen(), TENANT)
    expect(v).toEqual({ allowed: false, reason: 'group_disabled' })
  })

  test('a channel is refused even when opted in and mentioned', () => {
    const v = gate(group({ conversationType: 'channel', mentionsBot: true }), wideOpen(), TENANT)
    expect(v).toEqual({ allowed: false, reason: 'group_disabled' })
  })

  test('an empty group allowFrom no longer admits anyone in the room', () => {
    // Upstream: opted in + empty allowFrom = any sender could drive the session.
    const v = gate(group({ mentionsBot: true, senderAadObjectId: 'someone-else' }), wideOpen(), TENANT)
    expect(v).toEqual({ allowed: false, reason: 'group_disabled' })
  })

  test('a thread reply in an opted-in channel is refused too', () => {
    const v = gate(
      group({
        conversationType: 'channel',
        conversationId: '19:abc@thread.tacv2;messageid=170000',
        mentionsBot: true,
      }),
      wideOpen(),
      TENANT,
    )
    expect(v).toEqual({ allowed: false, reason: 'group_disabled' })
  })

  test('an activity with no conversation type is refused, not assumed personal', () => {
    const v = gate(dm({ conversationType: undefined }), access({ allowFrom: [SENDER] }), TENANT)
    expect(v).toEqual({ allowed: false, reason: 'group_disabled' })
  })
})

describe('outbound gate', () => {
  const CONV = '19:abc@thread.tacv2'
  const dmRef = { conversationId: 'a:1conv', conversationType: 'personal', senderId: SENDER }
  const channelRef = { conversationId: CONV, conversationType: 'channel' }

  test('a DM to a still-allowlisted sender is allowed', () => {
    expect(outboundAllowed(dmRef, access({ allowFrom: [SENDER] })).allowed).toBe(true)
  })

  test('REVOCATION REVOKES OUTBOUND — the bug this predicate exists for', () => {
    // Before this check, holding a stored conversation reference was the entire
    // outbound gate, so `/msteams:access remove` left every outbound path
    // working indefinitely: inbound refused while proactive sends kept landing.
    expect(outboundAllowed(dmRef, access({ allowFrom: [] }))).toEqual({
      allowed: false,
      reason: 'sender_not_allowed',
    })
  })

  test('a channel is never an outbound target (fork)', () => {
    expect(outboundAllowed(channelRef, access({ groups: {} }))).toEqual({
      allowed: false,
      reason: 'group_disabled',
    })
  })

  test('an opted-in channel is refused too (fork)', () => {
    const a = access({
      allowFrom: [SENDER],
      groups: { [CONV]: { requireMention: true, allowFrom: [] } },
    })
    expect(outboundAllowed(channelRef, a)).toEqual({ allowed: false, reason: 'group_disabled' })
  })

  test('the disabled kill switch stops outbound too', () => {
    // Deliberately stricter than discord/telegram, whose outbound gates ignore
    // dmPolicy. ACCESS.md documents 'disabled' as "drop everything", and a kill
    // switch that left proactive sends flowing would contradict its own docs.
    const off = access({
      dmPolicy: 'disabled',
      allowFrom: [SENDER],
      groups: { [CONV]: { requireMention: true, allowFrom: [] } },
    })
    expect(outboundAllowed(dmRef, off)).toEqual({ allowed: false, reason: 'dm_disabled' })
    expect(outboundAllowed(channelRef, off)).toEqual({ allowed: false, reason: 'dm_disabled' })
  })

  test('a personal reference with no recorded sender fails closed', () => {
    // References written before senderId was recorded carry no identity to
    // check, so they are refused until one inbound message repopulates them.
    // Discord's DM branch falls through to its throw in the same situation.
    expect(
      outboundAllowed(
        { conversationId: 'a:1conv', conversationType: 'personal' },
        access({ allowFrom: [SENDER] }),
      ),
    ).toEqual({ allowed: false, reason: 'no_sender_identity' })
  })

  test('a thread-scoped channel id is refused like its parent (fork)', () => {
    const a = access({ groups: { [CONV]: { requireMention: true, allowFrom: [] } } })
    expect(
      outboundAllowed(
        { conversationId: `${CONV};messageid=170000`, conversationType: 'channel' },
        a,
      ),
    ).toEqual({ allowed: false, reason: 'group_disabled' })
  })

  test('the sender match is case-insensitive on both sides', () => {
    expect(
      outboundAllowed(
        { ...dmRef, senderId: SENDER.toUpperCase() },
        access({ allowFrom: [SENDER] }),
      ).allowed,
    ).toBe(true)
  })
})

describe('mentionPatterns', () => {
  // access.json persists these and the access skill offers `set mentionPatterns`,
  // so a version of mentionsBot() that ignored them would ship a documented
  // feature with no consumer — the exact "documented but absent" trap telegram's
  // ACCESS.md falls into.
  test('an operator pattern counts as addressing the bot', () => {
    expect(mentionsBot([], BOT, 'claude, what is up', ['^claude\\b'])).toBe(true)
  })

  test('patterns are case-insensitive', () => {
    expect(mentionsBot([], BOT, 'Claude, hello', ['^claude\\b'])).toBe(true)
  })

  test('text not matching any pattern is not a mention', () => {
    expect(mentionsBot([], BOT, 'talking amongst ourselves', ['^claude\\b'])).toBe(false)
  })

  test('a real mention still wins with no patterns configured', () => {
    expect(mentionsBot([{ type: 'mention', mentioned: { id: BOT } }], BOT, 'hi')).toBe(true)
  })

  test('a malformed pattern disables itself, not the whole gate', () => {
    expect(mentionsBot([], BOT, 'claude hello', ['(unclosed', '^claude\\b'])).toBe(true)
  })

  test('patterns need text to match against', () => {
    expect(mentionsBot([], BOT, undefined, ['^claude\\b'])).toBe(false)
  })
})

describe('mention detection', () => {
  test('a mention of the bot counts', () => {
    const entities = [{ type: 'mention', mentioned: { id: BOT } }]
    expect(mentionsBot(entities, BOT)).toBe(true)
  })

  test('a mention of someone else does not count', () => {
    const entities = [{ type: 'mention', mentioned: { id: '29:other' } }]
    expect(mentionsBot(entities, BOT)).toBe(false)
  })

  test('typing the bot name as plain text does not count', () => {
    expect(mentionsBot([{ type: 'clientInfo' }], BOT)).toBe(false)
    expect(mentionsBot([], BOT)).toBe(false)
  })

  test('missing or malformed entities are handled', () => {
    expect(mentionsBot(undefined, BOT)).toBe(false)
    expect(mentionsBot('not-an-array', BOT)).toBe(false)
    expect(mentionsBot([{ type: 'mention' }], BOT)).toBe(false)
  })

  test('no recipient id means no mention can be confirmed', () => {
    expect(mentionsBot([{ type: 'mention', mentioned: { id: BOT } }], undefined)).toBe(false)
  })
})
