import { describe, expect, test } from 'bun:test'
import { matchMembers, sendActivity, callLink, memberEmail, MAX_TASK } from '../src/activity-notify.js'

const MEMBERS = [
  { aadObjectId: 'an-id', name: 'Andy, Le Van An (Firm)', email: 'an.le@fabrikam.com' },
  { aadObjectId: 'binh-id', name: 'Ben, Tran Thanh Binh (Firm)', email: 'binh.tran@fabrikam.com' },
  { aadObjectId: 'chi-id', name: 'Chloe, Pham Thu Chi (Firm)', userPrincipalName: 'thuchi.pham@fabrikam.com' },
  { aadObjectId: 'dan1', name: 'Daniel, Hoàng Minh Đức', email: 'duc.hoang@fabrikam.com' },
  { aadObjectId: 'dan2', name: 'Daniel Võ', email: 'daniel.vo@fabrikam.com' },
  { name: 'Bot without id' },
]

describe('matchMembers', () => {
  test('by name as said, accents ignored, or by email; duplicates collapse', () => {
    const r = matchMembers(MEMBERS, ['ben', 'thuchi.pham@fabrikam.com', 'hoang minh duc', 'Ben'])
    expect(r.map(m => m.aadObjectId)).toEqual(['binh-id', 'chi-id', 'dan1'])
  })

  test('"all" / "cả nhóm" / "mọi người" is every member but the sender', () => {
    for (const word of ['all', 'cả nhóm', 'Mọi người']) {
      expect(matchMembers(MEMBERS, [word], 'AN-ID').map(m => m.aadObjectId)).toEqual(['binh-id', 'chi-id', 'dan1', 'dan2'])
    }
  })

  test('several matches or nobody throw; only chat members can be rung', () => {
    expect(() => matchMembers(MEMBERS, ['Daniel'])).toThrow(/matches several people/)
    expect(() => matchMembers(MEMBERS, ['ha@fabrikam.com'])).toThrow(/not a member of this chat/)
    expect(() => matchMembers(MEMBERS, ['  '])).toThrow(/people is required/)
  })
})

describe('sendActivity', () => {
  function fake(status = 204) {
    const calls: { url: string; body: any; auth: string }[] = []
    const fetchImpl = (async (url: string, init?: RequestInit) => {
      calls.push({ url, body: JSON.parse(String(init?.body)), auth: String((init?.headers as any)?.Authorization) })
      return status === 204 ? new Response(null, { status }) : Response.json({ error: { code: 'Forbidden', message: 'detail' } }, { status })
    }) as any
    return { calls, fetchImpl }
  }

  test('one notification per person, on the chat, with the declared type and parameters', async () => {
    const f = fake()
    const out = await sendActivity(
      { chatId: '19:abc@thread.v2', recipients: matchMembers(MEMBERS, ['ben', 'chloe']), kind: 'taskAssigned', task: ' Điền   thông tin hai bên, hạn thứ Sáu ', by: 'Andy, Le Van An (Firm)' },
      { graphToken: async () => 'APP', fetchImpl: f.fetchImpl },
    )
    expect(out).toContain('binh.tran@fabrikam.com')
    expect(f.calls).toHaveLength(2)
    const c = f.calls[0]!
    expect(c.url).toBe('https://graph.microsoft.com/v1.0/chats/19%3Aabc%40thread.v2/sendActivityNotification')
    expect(c.auth).toBe('Bearer APP')
    expect(c.body).toEqual({
      topic: { source: 'entityUrl', value: 'https://graph.microsoft.com/v1.0/chats/19:abc@thread.v2' },
      activityType: 'systemDefault',
      previewText: { content: 'Điền thông tin hai bên, hạn thứ Sáu' },
      recipient: { '@odata.type': 'microsoft.graph.aadUserNotificationRecipient', userId: 'binh-id' },
      templateParameters: [
        { name: 'systemDefaultText', value: 'Andy, Le Van An (Firm) giao việc cho bạn: Điền thông tin hai bên, hạn thứ Sáu' },
      ],
    })
  })

  test('traces every Graph answer with status and request-id, nothing else', async () => {
    const traces: any[] = []
    const fetchImpl = (async () => new Response(null, { status: 204, headers: { 'request-id': 'req-1' } })) as any
    await sendActivity(
      { chatId: '19:a', recipients: matchMembers(MEMBERS, ['ben']), kind: 'reminder', task: 'họp 4h', by: 'A' },
      { graphToken: async () => 'APP', fetchImpl, trace: t => traces.push(t) },
    )
    expect(traces).toEqual([{ recipient: 'binh-id', status: 204, requestId: 'req-1' }])
  })

  test('names the Teams app when given, so a duplicate upload cannot swallow it', async () => {
    const f = fake()
    const r = matchMembers(MEMBERS, ['ben'])
    await sendActivity({ chatId: '19:a', recipients: r, kind: 'reminder', task: 'x', by: 'A' }, { graphToken: async () => 'APP', fetchImpl: f.fetchImpl, teamsAppId: 'catalog-id' })
    expect(f.calls[0]!.body.teamsAppId).toBe('catalog-id')
    const g = fake()
    await sendActivity({ chatId: '19:a', recipients: r, kind: 'reminder', task: 'x', by: 'A' }, { graphToken: async () => 'APP', fetchImpl: g.fetchImpl })
    expect('teamsAppId' in g.calls[0]!.body).toBe(false)
  })

  test('refuses an unknown kind, an empty or long task, and explains a 403', async () => {
    const r = matchMembers(MEMBERS, ['ben'])
    const run = (over: object, f = fake()) =>
      sendActivity({ chatId: '19:a', recipients: r, kind: 'taskAssigned', task: 'x', by: 'A', ...over } as any, { graphToken: async () => 'APP', fetchImpl: f.fetchImpl })
    await expect(run({ kind: 'party' })).rejects.toThrow(/kind must be one of/)
    await expect(run({ task: '   ' })).rejects.toThrow(/task is required/)
    await expect(run({ task: 'x'.repeat(MAX_TASK + 1) })).rejects.toThrow(/too long/)
    await expect(run({}, fake(403))).rejects.toThrow(/no permission/)
  })
})

describe('callLink (19/11)', () => {
  test('video by default, emails lower-cased and de-duplicated, group calls joined with commas', () => {
    expect(callLink(['Binh.Tran@fabrikam.com'], true)).toBe('https://teams.microsoft.com/l/call/0/0?users=binh.tran@fabrikam.com&withVideo=true')
    expect(callLink(['a@x.com', 'b@x.com', 'A@x.com'], false)).toBe('https://teams.microsoft.com/l/call/0/0?users=a@x.com,b@x.com')
  })
  test('refuses no valid email, and takes the UPN when a member has no email', () => {
    expect(() => callLink(['Ben', ''], true)).toThrow(/nobody to call/)
    expect(memberEmail({ userPrincipalName: 'ThuChi.Pham@fabrikam.com' })).toBe('thuchi.pham@fabrikam.com')
  })
})
