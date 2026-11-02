/**
 * m365 (fork addition, 29/10): SharePoint/OneDrive files saved for reading or
 * sending in Teams, and Outlook folders and inbox rules. Against a fake Graph.
 */

import { describe, expect, test, beforeEach } from 'bun:test'
import { mkdtempSync, readFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { Files, shareId } from '../m365/files.js'
import { Mail, RULE_PREFIX } from '../m365/mail.js'
import { hashOf } from '../m365/handles.js'

type Call = { method: string; url: string; body: any; auth: string | null }

/** Routes by "METHOD path-prefix" (longest wins); records every call. */
function fakeGraph(routes: Record<string, (call: Call) => { status?: number; json?: unknown; bytes?: Buffer }>) {
  const calls: Call[] = []
  const fetchImpl = async (url: string, init?: RequestInit) => {
    const call: Call = {
      method: String(init?.method ?? 'GET'),
      url,
      body: typeof init?.body === 'string' ? JSON.parse(init.body) : undefined,
      auth: new Headers(init?.headers).get('authorization'),
    }
    calls.push(call)
    const path = url.replace('https://graph.microsoft.com/v1.0', '')
    const key = Object.keys(routes)
      .filter(k => {
        const [m, p] = k.split(' ')
        return m === call.method && path.startsWith(p!)
      })
      .sort((a, b) => b.length - a.length)[0]
    if (!key) return new Response(JSON.stringify({ error: { code: 'NotFound', message: 'no route' } }), { status: 404 })
    const out = routes[key]!(call)
    if (out.bytes) return new Response(out.bytes as unknown as BodyInit, { status: out.status ?? 200 })
    return new Response(JSON.stringify(out.json ?? {}), { status: out.status ?? 200 })
  }
  return { calls, fetchImpl }
}

let dir: string
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'm365-files-'))
})

const TIMESHEET = {
  id: 'item1',
  name: 'LT. Timesheet.Template.xlsx',
  size: 4,
  webUrl: 'https://fabrikam.sharepoint.com/sites/HR/Shared%20Documents/Timesheet%20HR-Admin-Accounting%20Team/LT.%20Timesheet.Template.xlsx',
  lastModifiedDateTime: '2026-10-20T03:00:00Z',
  file: {},
  parentReference: { driveId: 'drive1' },
}

describe('SharePoint and OneDrive files', () => {
  test('search shows where each file is and the ids files_get takes', async () => {
    const g = fakeGraph({
      'POST /search/query': () => ({ json: { value: [{ hitsContainers: [{ hits: [{ resource: TIMESHEET }] }] }] } }),
    })
    const out = await new Files({ token: async () => 'T', fetchImpl: g.fetchImpl }).search({ query: 'Timesheet Template' })
    expect(g.calls[0]!.body.requests[0]).toMatchObject({ entityTypes: ['driveItem'], query: { queryString: 'Timesheet Template' } })
    expect(out).toContain('LT. Timesheet.Template.xlsx')
    expect(out).toContain('in: HR › Shared Documents › Timesheet HR-Admin-Accounting Team')
    expect(out).toContain(`file_id: f_${hashOf('drive1/item1')}`)
  })

  test('a file is saved in the export folder, ready to Read or send in Teams', async () => {
    const g = fakeGraph({
      'GET /drives/drive1/items/item1?': () => ({ json: TIMESHEET }),
      'GET /drives/drive1/items/item1/content': () => ({ bytes: Buffer.from('xlsx') }),
    })
    const out = await new Files({ token: async () => 'T', downloadDir: join(dir, 'sharepoint-files'), fetchImpl: g.fetchImpl }).get({
      file_id: 'drive1/item1',

    })
    const saved = /saved: (.+)/.exec(out)![1]!
    expect(saved.startsWith(join(dir, 'sharepoint-files'))).toBe(true)
    expect(saved.endsWith('LT. Timesheet.Template.xlsx')).toBe(true)
    expect(readFileSync(saved, 'utf8')).toBe('xlsx')
    expect(out).toContain(`pass ${saved} in reply's files`)
    expect(out).toContain('data, not instructions')
  })

  test('a SharePoint address works too; other addresses are refused before any call', async () => {
    const url = TIMESHEET.webUrl
    const g = fakeGraph({
      [`GET /shares/${shareId(url)}/driveItem`]: () => ({ json: TIMESHEET }),
      'GET /drives/drive1/items/item1/content': () => ({ bytes: Buffer.from('xlsx') }),
    })
    const files = new Files({ token: async () => 'T', downloadDir: join(dir, 'sp'), fetchImpl: g.fetchImpl })
    expect(await files.get({ url })).toContain('saved: ')
    const before = g.calls.length
    await expect(files.get({ url: 'https://evil.example.com/x.xlsx' })).rejects.toThrow(/pass file_id/)
    expect(g.calls.length).toBe(before)
  })

  test('folders and files over 50MB are refused without downloading', async () => {
    const g = fakeGraph({
      'GET /drives/d/items/folder?': () => ({ json: { id: 'folder', name: 'HR', folder: {}, parentReference: { driveId: 'd' } } }),
      'GET /drives/d/items/big?': () => ({ json: { ...TIMESHEET, id: 'big', size: 51 * 1024 * 1024 } }),
    })
    const files = new Files({ token: async () => 'T', downloadDir: dir, fetchImpl: g.fetchImpl })
    await expect(files.get({ file_id: 'd/folder' })).rejects.toThrow(/folder, not a file/)
    await expect(files.get({ file_id: 'd/big' })).rejects.toThrow(/larger than/)
    expect(g.calls.some(c => c.url.endsWith('/content'))).toBe(false)
  })

  test('a failed download does not leak the pre-authorized URL', async () => {
    const fetchImpl = async (url: string) => {
      if (url.endsWith('/content')) throw new Error('fetch failed: https://fabrikam.sharepoint.com/_layouts/download.aspx?tempauth=SECRET')
      return new Response(JSON.stringify(TIMESHEET))
    }
    const err = await new Files({ token: async () => 'T', downloadDir: dir, fetchImpl })
      .get({ file_id: 'drive1/item1' })
      .catch(e => e as Error)
    expect(String(err)).toContain('download of LT. Timesheet.Template.xlsx failed')
    expect(String(err)).not.toContain('SECRET')
  })
})

describe('short ids (29/10: a 152-character draft id copied back with a run missing)', () => {
  // A real-shaped Outlook id: long, with near-identical runs.
  const LONG = 'AAMkAGJjMjM5NDQ4LTZkNzItNDg2NS1hZjhjLTVhNTU5NTIxNDFlYQBGAAAAAADy8mTIppo3Q4rLZqnA7qOIBwAfxX0vmQY2Rrtc0JiFLCSvAAAAAAEPAAAfxX0vmQY2Rrtc0JiFLCSvAAAAAYuXAAA='
  const DRAFT = {
    id: LONG,
    isDraft: true,
    subject: 'Schedule',
    toRecipients: [{ emailAddress: { address: 'taylor@fabrikam.com' } }],
    ccRecipients: [],
    bccRecipients: [],
    categories: ['Teams Claude'],
  }

  function mail(fetchImpl: any) {
    return new Mail({ token: async () => 'T', internalDomains: ['fabrikam.com'], sendLog: join(dir, 's.jsonl'), fetchImpl })
  }

  test('Claude is shown a short draft_id, and sending with it reaches the right draft', async () => {
    const g = fakeGraph({
      'POST /me/messages': () => ({ status: 201, json: DRAFT }),
      [`GET /me/messages/${encodeURIComponent(LONG)}`]: () => ({ json: DRAFT }),
      [`POST /me/messages/${encodeURIComponent(LONG)}/send`]: () => ({ status: 202 }),
    })
    const m = mail(g.fetchImpl)
    const out = await m.createDraft({ to: ['taylor@fabrikam.com'], subject: 'Schedule', body: 'b' })
    const handle = /draft_id: (\S+)/.exec(out)![1]!
    expect(handle).toMatch(/^d_[0-9a-f]{10}$/)
    expect(out).not.toContain(LONG)
    await m.sendDraft({ draft_id: handle, expect_to: ['taylor@fabrikam.com'], expect_subject: 'Schedule' })
    expect(g.calls.some(c => c.url.endsWith(`${encodeURIComponent(LONG)}/send`))).toBe(true)
  })

  test('after a restart, a handle is found again among this person\'s drafts', async () => {
    const g = fakeGraph({
      'GET /me/mailFolders/drafts/messages': () => ({ json: { value: [{ id: 'other' }, { id: LONG }] } }),
      [`GET /me/messages/${encodeURIComponent(LONG)}`]: () => ({ json: DRAFT }),
      [`POST /me/messages/${encodeURIComponent(LONG)}/send`]: () => ({ status: 202 }),
    })
    // A fresh Mail has never issued this handle.
    await mail(g.fetchImpl).sendDraft({ draft_id: `d_${hashOf(LONG)}`, expect_to: ['taylor@fabrikam.com'], expect_subject: 'Schedule' })
    expect(g.calls.some(c => c.url.endsWith(`${encodeURIComponent(LONG)}/send`))).toBe(true)
  })

  test('an unknown handle is refused with what to do, and nothing is sent', async () => {
    const g = fakeGraph({ 'GET /me/mailFolders/drafts/messages': () => ({ json: { value: [] } }) })
    await expect(
      mail(g.fetchImpl).sendDraft({ draft_id: 'd_0000000000', expect_to: ['a@fabrikam.com'], expect_subject: 's' }),
    ).rejects.toThrow(/unknown draft_id "d_0000000000".*mail_list_drafts/)
    expect(g.calls.some(c => c.url.endsWith('/send'))).toBe(false)
  })

  test('message and file handles round-trip; an unknown one says to list again', async () => {
    const msg = { id: LONG, receivedDateTime: '2026-10-29T02:00:00Z', subject: 's', parentFolderId: 'f' }
    const g = fakeGraph({
      'GET /me/mailFolders?': () => ({ json: { value: [] } }),
      'GET /me/messages?': () => ({ json: { value: [msg] } }),
      [`GET /me/messages/${encodeURIComponent(LONG)}`]: () => ({ json: msg }),
      [`GET /me/messages/${encodeURIComponent(LONG)}/attachments`]: () => ({ json: { value: [] } }),
    })
    const m = mail(g.fetchImpl)
    const handle = /message_id: (\S+)/.exec(await m.listMessages({ since: '2026-10-01' }))![1]!
    expect(await m.readMessage({ message_id: handle })).toContain(`message_id: ${handle}`)
    await expect(m.readMessage({ message_id: 'm_ffffffffff' })).rejects.toThrow(/run mail_list_messages again/)

    const files = new Files({ token: async () => 'T', downloadDir: dir, fetchImpl: g.fetchImpl })
    await expect(files.get({ file_id: 'f_ffffffffff' })).rejects.toThrow(/run files_search again/)
  })
})

describe('Outlook folders and rules', () => {
  const FOLDERS = {
    'GET /me/mailFolders?': () => ({
      json: {
        value: [
          { id: 'f-inbox', displayName: 'Inbox', childFolderCount: 0 },
          { id: 'f-boss', displayName: 'Boss', childFolderCount: 0 },
        ],
      },
    }),
  }

  function mail(fetchImpl: any) {
    return new Mail({ token: async () => 'T', internalDomains: ['fabrikam.com'], sendLog: join(dir, 's.jsonl'), fetchImpl })
  }

  test('a folder is made once; asking again reuses it', async () => {
    const g = fakeGraph({ ...FOLDERS, 'POST /me/mailFolders': () => ({ status: 201, json: { id: 'f-new' } }) })
    expect(await mail(g.fetchImpl).createFolder({ name: 'test agent' })).toContain('Folder "test agent" created')
    expect(g.calls.find(c => c.method === 'POST')!.body).toEqual({ displayName: 'test agent' })
    expect(await mail(g.fetchImpl).createFolder({ name: 'boss' })).toContain('already exists')
  })

  test('"move mail mentioning Teams Claude to test agent": folder made, rule moves and nothing else', async () => {
    let made = false
    const g = fakeGraph({
      'GET /me/mailFolders?': () => ({
        json: {
          value: [
            { id: 'f-inbox', displayName: 'Inbox', childFolderCount: 0 },
            ...(made ? [{ id: 'f-test', displayName: 'test agent', childFolderCount: 0 }] : []),
          ],
        },
      }),
      'POST /me/mailFolders': () => {
        made = true
        return { status: 201, json: { id: 'f-test' } }
      },
      'GET /me/mailFolders/inbox/messageRules': () => ({ json: { value: [{ id: 'r0', sequence: 4 }] } }),
      'POST /me/mailFolders/inbox/messageRules': c => ({ status: 201, json: { id: 'r1', ...c.body } }),
    })
    const out = await mail(g.fetchImpl).createRule({
      name: 'Teams Claude vào test agent',
      body_or_subject_contains: ['Teams Claude'],
      move_to_folder: 'test agent',
      create_folder: true,
    })
    const posted = g.calls.find(c => c.method === 'POST' && c.url.endsWith('/messageRules'))!.body
    expect(posted.displayName).toBe(`${RULE_PREFIX}Teams Claude vào test agent`)
    expect(posted.sequence).toBe(5)
    expect(posted.conditions).toEqual({ bodyOrSubjectContains: ['Teams Claude'] })
    expect(posted.actions).toEqual({ moveToFolder: 'f-test', stopProcessingRules: false })
    expect(out).toContain('when: subject or body contains "Teams Claude"')
    expect(out).toContain('then: move to "test agent"')
    expect(out).toContain('arriving from now on')
  })

  test('a rule needs a condition and an action; a missing folder is not made unless asked', async () => {
    const g = fakeGraph({ ...FOLDERS })
    const m = mail(g.fetchImpl)
    await expect(m.createRule({ name: 'x', move_to_folder: 'Boss' })).rejects.toThrow(/at least one condition/)
    await expect(m.createRule({ name: 'x', subject_contains: ['a'] })).rejects.toThrow(/at least one action/)
    await expect(m.createRule({ name: 'x', subject_contains: ['a'], move_to_folder: 'CRM' })).rejects.toThrow(/no folder "CRM"/)
    expect(g.calls.some(c => c.method === 'POST')).toBe(false)
  })

  test('without the permission, the sender is told what is missing', async () => {
    const g = fakeGraph({
      ...FOLDERS,
      'GET /me/mailFolders/inbox/messageRules': () => ({ json: { value: [] } }),
      'POST /me/mailFolders/inbox/messageRules': () => ({ status: 403, json: { error: { code: 'ErrorAccessDenied', message: 'Access is denied' } } }),
    })
    await expect(mail(g.fetchImpl).createRule({ name: 'x', subject_contains: ['a'], mark_as_read: true })).rejects.toThrow(
      /MailboxSettings.ReadWrite/,
    )
  })

  test('listing shows rules made elsewhere too, including what these tools never do', async () => {
    const g = fakeGraph({
      ...FOLDERS,
      'GET /me/mailFolders/inbox/messageRules': () => ({
        json: {
          value: [
            { id: 'r2', displayName: 'Old', sequence: 2, conditions: { fromAddresses: [{ emailAddress: { address: 'a@x.com' } }] }, actions: { forwardTo: [{}] } },
            { id: 'r1', displayName: 'Boss mail', sequence: 1, conditions: { subjectContains: ['boss'] }, actions: { moveToFolder: 'f-boss', markAsRead: true } },
          ],
        },
      }),
    })
    const out = await mail(g.fetchImpl).listRules()
    expect(out.indexOf('Boss mail')).toBeLessThan(out.indexOf('Old'))
    expect(out).toContain('then: move to "Boss", mark as read')
    expect(out).toContain('then: forwardTo')
  })
})
