/**
 * Fork addition (02/11): cloud files shared into the bot's chat, and a file sent
 * with a Teams message (shared with the recipients only). Graph is faked.
 */

import { describe, expect, test, beforeEach } from 'bun:test'
import { mkdtempSync, mkdirSync, writeFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { TeamsChats } from '../m365/teams.js'
import { Files } from '../m365/files.js'
import { sharedFiles, TEAMS_FILE_DOWNLOAD_INFO } from '../src/attachments.js'
import { normalize } from '../src/normalize.js'

type Call = { method: string; url: string; body: any }

function fakeGraph(routes: Record<string, (call: Call) => { status?: number; json?: unknown }>) {
  const calls: Call[] = []
  const fetchImpl = async (url: string, init?: RequestInit) => {
    const call: Call = { method: String(init?.method ?? 'GET'), url, body: typeof init?.body === 'string' ? JSON.parse(init.body) : init?.body }
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
    return new Response(JSON.stringify(out.json ?? {}), { status: out.status ?? 200 })
  }
  return { calls, fetchImpl }
}

const JIMMY = { id: 'u-jimmy', displayName: 'Jimmy, Tran Van A (Fabrikam)', mail: 'jimmy.tran@fabrikam.com' }

let dir: string
let exportDir: string
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'm365-teams-files-'))
  exportDir = join(dir, 'bot-export')
  mkdirSync(exportDir)
})

function teams(fetchImpl: any, withFiles = true) {
  return new TeamsChats({
    token: async () => 'ACCESS',
    internalDomains: ['fabrikam.com'],
    sendLog: join(dir, 'sent.jsonl'),
    ...(withFiles ? { files: new Files({ token: async () => 'ACCESS', uploadRoots: [exportDir], fetchImpl }) } : {}),
    fetchImpl,
    now: () => Date.parse('2026-11-02T08:50:00Z'),
  })
}

function routes() {
  return fakeGraph({
    'GET /me': () => ({ json: { id: 'u-me', mail: 'sam@fabrikam.com' } }),
    'GET /users/': () => ({ json: JIMMY }),
    'POST /chats/': () => ({ status: 201, json: { id: 'm1' } }),
    'POST /chats': () => ({ status: 201, json: { id: '19:chat' } }),
    'PUT /me/drive/root:': () => ({ status: 201, json: { id: 'item1', name: 'Phan-cong.pdf', size: 9, webUrl: 'https://fabrikam-my.sharepoint.com/personal/x/Phan-cong.pdf' } }),
    'POST /me/drive/items/': () => ({ json: { value: [] } }),
  })
}

describe('a file sent with a Teams message', () => {
  test('is shown in the plan, then uploaded, shared with the recipient only and linked', async () => {
    const g = routes()
    const pdf = join(exportDir, 'Phan-cong.pdf')
    writeFileSync(pdf, '%PDF-1.4 x')
    const t = teams(g.fetchImpl)
    const plan = await t.prepareMessage({ to: ['jimmy.tran@fabrikam.com'], text: 'Gửi anh bảng phân công', file_path: pdf })
    expect(plan).toContain('file: Phan-cong.pdf')
    expect(plan).toContain('shared with jimmy.tran@fabrikam.com only')
    expect(g.calls.some(c => c.method !== 'GET')).toBe(false)

    const id = /plan_id: (\S+)/.exec(plan)![1]!
    await t.sendMessage({ plan_id: id, expect_to: ['jimmy.tran@fabrikam.com'] })
    const invite = g.calls.find(c => c.url.includes('/invite'))!
    expect(invite.body).toEqual({ recipients: [{ email: 'jimmy.tran@fabrikam.com' }], roles: ['read'], requireSignIn: true, sendInvitation: false })
    const msg = g.calls.find(c => c.url.includes('/messages'))!
    expect(msg.body.body.content).toBe(
      'Gửi anh bảng phân công<br><br>📎 <a href="https://fabrikam-my.sharepoint.com/personal/x/Phan-cong.pdf">Phan-cong.pdf</a>',
    )
    // Shared before the message goes out, so the link opens on arrival.
    expect(g.calls.indexOf(invite)).toBeLessThan(g.calls.indexOf(msg))
  })

  test('must come from the export folder, and never goes to a channel', async () => {
    const g = routes()
    await expect(teams(g.fetchImpl).prepareMessage({ to: ['jimmy.tran@fabrikam.com'], text: 'x', file_path: join(dir, 'elsewhere.pdf') })).rejects.toThrow()
    await expect(teams(g.fetchImpl).prepareMessage({ channel_id: 'h1', text: 'x', file_path: join(exportDir, 'a.pdf') })).rejects.toThrow(/channel/)
    await expect(teams(g.fetchImpl, false).prepareMessage({ to: ['jimmy.tran@fabrikam.com'], text: 'x', file_path: join(exportDir, 'a.pdf') })).rejects.toThrow(/disabled/)
  })
})

describe('shared cloud files (inbound)', () => {
  const base = { id: '1', serviceUrl: 'https://smba.trafficmanager.net/apac/', conversation: { id: 'a:1', conversationType: 'personal' }, from: { aadObjectId: 'x' } }

  test('a OneDrive file attached from the cloud comes through by address', () => {
    const activity = {
      ...base,
      attachments: [
        { contentType: 'reference', contentUrl: 'https://fabrikam-my.sharepoint.com/personal/trang/Documents/Recordings/Meeting.docx', name: 'Meeting.docx' },
        { contentType: 'text/html', content: '<p>Tổng hợp lại</p>' },
      ],
    }
    const shared = sharedFiles(activity)
    expect(shared).toEqual([{ url: 'https://fabrikam-my.sharepoint.com/personal/trang/Documents/Recordings/Meeting.docx', name: 'Meeting.docx' }])
    const n = normalize({ activity: { ...activity, text: 'Tổng hợp lại' }, sharedFiles: shared })
    expect(n.meta.shared_file_url).toBe(shared[0]!.url)
    expect(n.meta.shared_file_name).toBe('Meeting.docx')
  })

  test('a download.info card with no downloadUrl falls back to its address; one with a downloadUrl is left to register', () => {
    const withUrl = { contentType: TEAMS_FILE_DOWNLOAD_INFO, name: 'a.docx', contentUrl: 'https://t.sharepoint.com/a.docx', content: { downloadUrl: 'https://t.sharepoint.com/dl?tempauth=SECRET' } }
    const without = { contentType: TEAMS_FILE_DOWNLOAD_INFO, name: 'b.docx', contentUrl: 'https://t.sharepoint.com/b.docx', content: {} }
    expect(sharedFiles({ ...base, attachments: [withUrl, without] })).toEqual([{ url: 'https://t.sharepoint.com/b.docx', name: 'b.docx' }])
  })

  test('links in the message body count; other hosts, plain http and tempauth URLs never do', () => {
    const html = [
      '<a href="https://fabrikam.sharepoint.com/:w:/s/Team/abc?e=x1&amp;at=9">Plan</a>',
      '<a href="https://evil.example.com/x">bad</a>',
      '<a href="http://fabrikam.sharepoint.com/x">http</a>',
      '<a href="https://fabrikam.sharepoint.com/dl?tempauth=TOKEN">cred</a>',
    ].join('')
    expect(sharedFiles({ ...base, attachments: [{ contentType: 'text/html', content: html }] })).toEqual([
      { url: 'https://fabrikam.sharepoint.com/:w:/s/Team/abc?e=x1&at=9' },
    ])
  })

  test('several shared files are all listed', () => {
    const shared = [{ url: 'https://t.sharepoint.com/1', name: 'one.docx' }, { url: 'https://t.sharepoint.com/2', name: 'two.xlsx' }]
    const n = normalize({ activity: { ...base, text: '' }, sharedFiles: shared })
    expect(n.meta.shared_files).toBe('one.docx https://t.sharepoint.com/1; two.xlsx https://t.sharepoint.com/2')
    expect(n.content).toBe('(one.docx)')
  })
})
