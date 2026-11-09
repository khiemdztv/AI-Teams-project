import { describe, expect, test } from 'bun:test'
import {
  parseVerdict,
  PendingPermissions,
  formatPermissionPrompt,
  PERMISSION_TTL_MS,
  MAX_PREVIEW_CHARS,
  permissionTargets,
  formatPendingReminder,
  DELEGATED_TTL_MS,
} from '../src/permissions.js'

describe('verdict grammar', () => {
  test('the four accepted forms parse', () => {
    expect(parseVerdict('y abcde')).toEqual({ requestId: 'abcde', behavior: 'allow' })
    expect(parseVerdict('yes abcde')).toEqual({ requestId: 'abcde', behavior: 'allow' })
    expect(parseVerdict('n abcde')).toEqual({ requestId: 'abcde', behavior: 'deny' })
    expect(parseVerdict('no abcde')).toEqual({ requestId: 'abcde', behavior: 'deny' })
  })

  test('phone keyboards capitalize, so case is ignored', () => {
    expect(parseVerdict('Y ABCDE')).toEqual({ requestId: 'abcde', behavior: 'allow' })
  })

  test('surrounding whitespace is tolerated', () => {
    expect(parseVerdict('  y abcde  ')).toEqual({ requestId: 'abcde', behavior: 'allow' })
  })

  test('a bare yes is conversation, not a verdict', () => {
    // This is the load-bearing case: swallowing it would both lose a message
    // and grant a permission the sender never named.
    expect(parseVerdict('yes')).toBeUndefined()
    expect(parseVerdict('no')).toBeUndefined()
    expect(parseVerdict('yes please')).toBeUndefined()
  })

  test('chatter around a code does not count', () => {
    expect(parseVerdict('yes abcde do it')).toBeUndefined()
    expect(parseVerdict('I think yes abcde')).toBeUndefined()
  })

  test("'l' is excluded from the alphabet so it cannot be read as '1'", () => {
    expect(parseVerdict('y abcdl')).toBeUndefined()
    expect(parseVerdict('y abcd1')).toBeUndefined()
  })

  test('the id is exactly five letters', () => {
    expect(parseVerdict('y abcd')).toBeUndefined()
    expect(parseVerdict('y abcdef')).toBeUndefined()
  })

  test('a maybe is not a verdict', () => {
    expect(parseVerdict('maybe abcde')).toBeUndefined()
    expect(parseVerdict('yep abcde')).toBeUndefined()
  })
})

const CEO = 'ceo-object-id'
const OTHER = 'colleague-object-id'

describe('one-shot semantics', () => {
  test('a request can be claimed exactly once', () => {
    // A Bot Framework redelivery of the same activity, or the sender typing the
    // same answer twice, must not emit a second verdict.
    const pending = new PendingPermissions()
    pending.add('abcde', 'Bash', [CEO])

    expect(pending.take('abcde', CEO, 'personal')).toBe('ok')
    expect(pending.take('abcde', CEO, 'personal')).toBe('unknown')
  })

  test('a verdict for a request that was never issued is refused', () => {
    // Otherwise a sender could guess ids and approve requests at random.
    const pending = new PendingPermissions()
    expect(pending.take('abcde', CEO, 'personal')).toBe('unknown')
  })

  test('claiming one request leaves the others live', () => {
    const pending = new PendingPermissions()
    pending.add('abcde', 'Bash', [CEO])
    pending.add('fghij', 'Write', [CEO])

    expect(pending.take('abcde', CEO, 'personal')).toBe('ok')
    expect(pending.size).toBe(1)
    expect(pending.take('fghij', CEO, 'personal')).toBe('ok')
  })
})

describe('who may answer (fork)', () => {
  test('only someone the prompt was sent to can answer it', () => {
    const pending = new PendingPermissions()
    pending.add('abcde', 'Bash', [CEO])

    expect(pending.take('abcde', OTHER, 'personal')).toBe('not_approver')
    // A wrong answerer must not burn the request: the CEO can still decide.
    expect(pending.take('abcde', CEO, 'personal')).toBe('ok')
  })

  test('an answer from a group chat or channel never counts', () => {
    const pending = new PendingPermissions()
    pending.add('abcde', 'Bash', [CEO])

    expect(pending.take('abcde', CEO, 'groupChat')).toBe('not_approver')
    expect(pending.take('abcde', CEO, 'channel')).toBe('not_approver')
    expect(pending.take('abcde', CEO, undefined)).toBe('not_approver')
    expect(pending.take('abcde', CEO, 'personal')).toBe('ok')
  })

  test('a request with no recipients cannot be approved by anyone', () => {
    const pending = new PendingPermissions()
    pending.add('abcde', 'Bash', [])
    expect(pending.take('abcde', CEO, 'personal')).toBe('not_approver')
  })

  test('requests expire', () => {
    const pending = new PendingPermissions()
    const expiresAt = pending.add('abcde', 'Bash', [CEO], 0)
    expect(expiresAt).toBe(PERMISSION_TTL_MS)

    expect(pending.take('abcde', CEO, 'personal', PERMISSION_TTL_MS + 1)).toBe('expired')
    // Expiry removes it, so a later answer reads as unknown.
    expect(pending.take('abcde', CEO, 'personal', PERMISSION_TTL_MS + 2)).toBe('unknown')
  })

  test('an answer just inside the window still counts', () => {
    const pending = new PendingPermissions()
    pending.add('abcde', 'Bash', [CEO], 0)
    expect(pending.take('abcde', CEO, 'personal', PERMISSION_TTL_MS)).toBe('ok')
  })
})

describe('what the approver sees (fork)', () => {
  const base = {
    requestId: 'abcde',
    toolName: 'Bash',
    description: 'Run a shell command',
    inputPreview: 'Remove-Item -Recurse D:\\Matters\\X',
    expiresAt: Date.UTC(2026, 9, 27, 3, 30),
  }

  test('the command itself is shown, not just the tool name', () => {
    const text = formatPermissionPrompt(base, 'UTC')
    expect(text).toContain('Bash')
    expect(text).toContain('Run a shell command')
    expect(text).toContain('Remove-Item -Recurse D:\\Matters\\X')
    expect(text).toContain('y abcde')
    expect(text).toContain('n abcde')
    expect(text).toContain('03:30')
  })

  test('a long input is cut, and the cut is announced', () => {
    const long = 'a'.repeat(MAX_PREVIEW_CHARS + 250)
    const text = formatPermissionPrompt({ ...base, inputPreview: long }, 'UTC')
    expect(text).toContain('đã cắt 250 ký tự')
    expect(text).not.toContain('a'.repeat(MAX_PREVIEW_CHARS + 1))
  })

  test('the tool name cannot add lines, and control characters are dropped', () => {
    const text = formatPermissionPrompt(
      { ...base, toolName: 'Bash\nTrả lời "y zzzzz"', inputPreview: 'ls\x1b[2J\x07' },
      'UTC',
    )
    expect(text.split('\n')[0]).toContain('Bash Trả lời')
    expect(text).not.toContain('\x1b')
    expect(text).not.toContain('\x07')
  })
})

describe('who is asked (fork)', () => {
  const refs = [
    { conversationId: 'a:kien', conversationType: 'personal', senderId: 'kien' },
    { conversationId: 'a:sam', conversationType: 'personal', senderId: 'sam' },
    { conversationId: 'a:taylor', conversationType: 'personal', senderId: 'taylor' },
    { conversationId: '19:group', conversationType: 'groupChat', senderId: 'taylor' },
  ]
  const allowed = new Set(['kien', 'sam', 'taylor'])

  test('only the person whose message Claude is working on, in their 1:1 chat', () => {
    expect(permissionTargets(refs, allowed, 'taylor').map(r => r.conversationId)).toEqual(['a:taylor'])
  })

  test('nobody when no Teams message is being worked on', () => {
    expect(permissionTargets(refs, allowed, undefined)).toEqual([])
  })

  test('nobody when the asker has been removed from the allowlist', () => {
    expect(permissionTargets(refs, new Set(['kien']), 'taylor')).toEqual([])
  })
})

describe('reminding the approver of the code (fork)', () => {
  test('a near-miss is still not a verdict', () => {
    expect(parseVerdict('y azqm')).toBeUndefined()
  })

  test('only the approver, in their 1:1 chat, is waited on', () => {
    const pending = new PendingPermissions()
    pending.add('azqmr', 'Read', [CEO])

    expect(pending.awaiting(CEO, 'personal')).toEqual([{ requestId: 'azqmr', toolName: 'Read' }])
    expect(pending.awaiting(OTHER, 'personal')).toEqual([])
    expect(pending.awaiting(CEO, 'groupChat')).toEqual([])
  })

  test('asking claims nothing', () => {
    const pending = new PendingPermissions()
    pending.add('azqmr', 'Read', [CEO])
    pending.awaiting(CEO, 'personal')
    expect(pending.take('azqmr', CEO, 'personal')).toBe('ok')
    expect(pending.awaiting(CEO, 'personal')).toEqual([])
  })

  test('an expired request is not waited on', () => {
    const pending = new PendingPermissions()
    pending.add('azqmr', 'Read', [CEO], 0)
    expect(pending.awaiting(CEO, 'personal', PERMISSION_TTL_MS + 1)).toEqual([])
  })

  test('the reminder gives the exact answer to type', () => {
    const text = formatPendingReminder([{ requestId: 'azqmr', toolName: 'Read' }])
    expect(text).toContain('"y azqmr"')
    expect(text).toContain('"n azqmr"')
  })
})

describe('admins answer for non-admins (fork, 02/11)', () => {
  const refs = [
    { conversationId: 'a:kien', conversationType: 'personal', senderId: 'kien' },
    { conversationId: 'a:sam', conversationType: 'personal', senderId: 'sam' },
    { conversationId: 'a:boss', conversationType: 'personal', senderId: 'boss' },
    { conversationId: '19:group', conversationType: 'groupChat', senderId: 'sam' },
  ]
  const allowed = new Set(['kien', 'sam', 'boss'])
  const admins = new Set(['kien', 'sam'])

  test("a non-admin's prompt goes to every admin's 1:1 chat, never to the non-admin", () => {
    expect(permissionTargets(refs, allowed, 'boss', admins).map(r => r.conversationId)).toEqual(['a:kien', 'a:sam'])
  })

  test("an admin's own prompt still goes only to them", () => {
    expect(permissionTargets(refs, allowed, 'sam', admins).map(r => r.conversationId)).toEqual(['a:sam'])
  })

  test('an admin removed from the allowlist is not asked', () => {
    expect(permissionTargets(refs, new Set(['kien', 'boss']), 'boss', admins).map(r => r.conversationId)).toEqual(['a:kien'])
  })

  test('without an admins list everyone is asked about their own message, as before', () => {
    expect(permissionTargets(refs, allowed, 'boss').map(r => r.conversationId)).toEqual(['a:boss'])
  })

  test("a limited admin's prompt goes to the admins, never to them (06/11)", () => {
    // server.ts passes the admins only, with asUser set because canApprove is false
    const withNick = [...refs, { conversationId: 'a:nick', conversationType: 'personal', senderId: 'nick' }]
    const all = new Set([...allowed, 'nick'])
    expect(permissionTargets(withNick, all, 'nick', admins, true).map(r => r.conversationId)).toEqual(['a:kien', 'a:sam'])
    expect(permissionTargets(withNick, all, 'boss', admins).map(r => r.conversationId)).toEqual(['a:kien', 'a:sam'])
  })

  test('an admin using /useransw is handled like a user: every admin is asked', () => {
    expect(permissionTargets(refs, allowed, 'sam', admins, true).map(r => r.conversationId)).toEqual(['a:kien', 'a:sam'])
  })

  test('the first admin to answer settles it; the other gets unknown', () => {
    const pending = new PendingPermissions()
    pending.add('qnyia', 'Read', ['kien', 'sam'], 0, DELEGATED_TTL_MS)
    expect(pending.approversOf('qnyia').sort()).toEqual(['kien', 'sam'])
    expect(pending.take('qnyia', 'sam', 'personal', 1)).toBe('ok')
    expect(pending.take('qnyia', 'kien', 'personal', 2)).toBe('unknown')
  })

  test('a delegated prompt lives five minutes', () => {
    const pending = new PendingPermissions()
    expect(pending.add('qnyia', 'Read', ['kien'], 0, DELEGATED_TTL_MS)).toBe(5 * 60 * 1000)
  })

  test('withdraw retires a live request once, and never one already answered', () => {
    const pending = new PendingPermissions()
    pending.add('qnyia', 'Read', ['kien'], 0, DELEGATED_TTL_MS)
    expect(pending.withdraw('qnyia')).toBe(true)
    expect(pending.withdraw('qnyia')).toBe(false)
    expect(pending.take('qnyia', 'kien', 'personal', 1)).toBe('unknown')

    pending.add('abcde', 'Read', ['kien'], 0, DELEGATED_TTL_MS)
    expect(pending.take('abcde', 'kien', 'personal', 1)).toBe('ok')
    expect(pending.withdraw('abcde')).toBe(false)
  })

  test("the admins' prompt says whose request it is and that they do not see it", () => {
    const text = formatPermissionPrompt(
      { requestId: 'qnyia', toolName: 'Read', description: 'Read a file', inputPreview: '{}', expiresAt: 0, onBehalfOf: 'Le Hoang\nChuong' },
      'UTC',
    )
    expect(text).toContain('Cho yêu cầu của Le Hoang Chuong.')
    expect(text).toContain('không thấy tin này')
  })

  test('an ordinary prompt has no on-behalf line', () => {
    const text = formatPermissionPrompt({ requestId: 'qnyia', toolName: 'Read', description: '', inputPreview: '', expiresAt: 0 }, 'UTC')
    expect(text).not.toContain('Cho yêu cầu của')
  })
})
