import { test, expect } from 'bun:test'
import { mkdtempSync, readFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { describeRefusal, appendRefusal, REFUSAL_LOG } from '../src/refusal-log.js'

const TENANT = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
const BOT = '28:44444444-4444-4444-8444-444444444444'
const SENDER = '00000000-1111-2222-3333-444444444444'

const activity = () => ({
  type: 'message',
  text: '<at>Fabrikam Agent</at> secret plan for the client',
  from: { aadObjectId: SENDER, name: 'Someone' },
  recipient: { id: BOT },
  conversation: { id: '19:room@thread.v2', conversationType: 'groupChat', tenantId: TENANT },
  channelData: { tenant: { id: TENANT } },
  entities: [{ type: 'mention', text: '<at>Fabrikam Agent</at>', mentioned: { id: BOT } }],
})

test('a refusal line carries the shape of the activity and nothing that identifies it', () => {
  const line = describeRefusal(activity(), 'mention_required', TENANT, new Date('2026-11-18T01:00:00Z'))
  expect(line).toBe(
    '2026-11-18T01:00:00.000Z reason=mention_required type=groupChat activity=message tenant_match=true has_aad=true mentions=1 bot_mentioned=true recipient=28: mention_ids=28:',
  )
  for (const leak of ['secret', 'Someone', SENDER, TENANT, '44444444', 'room']) expect(line).not.toContain(leak)
})

test('a mention of someone else, or none, shows as such', () => {
  const a = { ...activity(), entities: [{ type: 'mention', mentioned: { id: '29:other' } }] }
  expect(describeRefusal(a, 'mention_required', TENANT)).toContain('mentions=1 bot_mentioned=false recipient=28: mention_ids=29:')
  expect(describeRefusal({ ...activity(), entities: undefined }, 'x', TENANT)).toContain('mentions=0 bot_mentioned=false')
})

test('lines are appended to refused.log in the state dir', () => {
  const dir = mkdtempSync(join(tmpdir(), 'msteams-refused-'))
  appendRefusal(dir, 'one')
  appendRefusal(dir, 'two')
  expect(readFileSync(join(dir, REFUSAL_LOG), 'utf8')).toBe('one\ntwo\n')
})
