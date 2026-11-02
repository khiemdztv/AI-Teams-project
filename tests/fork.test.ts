/**
 * Tests for the Fabrikam fork's additions that have no upstream test file:
 * the export-folder rule (B1), credential redaction in the queue (B2), inbox
 * cleanup (B3) and the download host allowlist (B4).
 */

import { describe, expect, test, beforeEach } from 'bun:test'
import { mkdtempSync, mkdirSync, writeFileSync, readFileSync, readdirSync, symlinkSync, utimesSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { assertWithinRoot, loadSendableFile } from '../src/attach.js'
import { IngressQueue, redactCredentials } from '../src/queue.js'
import { purgeOlderThan, INBOX_TTL_MS } from '../src/purge.js'
import { loadEnvFile } from '../src/env.js'
import {
  AttachmentHandles,
  isAllowedDownloadUrl,
  downloadHostsFromEnv,
  DEFAULT_DOWNLOAD_HOSTS,
} from '../src/attachments.js'

const FIXTURES = join(import.meta.dir, 'fixtures')
const load = (name: string) => JSON.parse(readFileSync(join(FIXTURES, name), 'utf8')) as Record<string, any>

let root: string
let exportDir: string
let outside: string

beforeEach(() => {
  root = mkdtempSync(join(tmpdir(), 'msteams-fork-'))
  exportDir = join(root, 'export')
  outside = join(root, 'clients')
  mkdirSync(exportDir)
  mkdirSync(outside)
})

describe('B1: only the export folder is sendable', () => {
  test('a file inside the export folder passes', () => {
    const file = join(exportDir, 'draft.docx')
    writeFileSync(file, 'x')
    expect(() => assertWithinRoot(file, exportDir)).not.toThrow()
  })

  test('a file in a subfolder of the export folder passes', () => {
    mkdirSync(join(exportDir, 'matter-1'))
    const file = join(exportDir, 'matter-1', 'memo.pdf')
    writeFileSync(file, 'x')
    expect(() => assertWithinRoot(file, exportDir)).not.toThrow()
  })

  test('a file outside it is refused', () => {
    const file = join(outside, 'other-client.pdf')
    writeFileSync(file, 'x')
    expect(() => assertWithinRoot(file, exportDir)).toThrow(/only files inside/)
  })

  test('a sibling folder sharing the prefix is refused', () => {
    // "export-old" starts with "export" but is not inside it.
    mkdirSync(join(root, 'export-old'))
    const file = join(root, 'export-old', 'x.pdf')
    writeFileSync(file, 'x')
    expect(() => assertWithinRoot(file, exportDir)).toThrow(/only files inside/)
  })

  test('.. in the path does not escape', () => {
    const file = join(outside, 'secret.pdf')
    writeFileSync(file, 'x')
    expect(() => assertWithinRoot(join(exportDir, '..', 'clients', 'secret.pdf'), exportDir)).toThrow(
      /only files inside/,
    )
  })

  test('a missing file is refused, not let through', () => {
    expect(() => assertWithinRoot(join(exportDir, 'gone.pdf'), exportDir)).toThrow(/not found/)
  })

  test('a missing export folder is a configuration error', () => {
    const file = join(outside, 'x.pdf')
    writeFileSync(file, 'x')
    expect(() => assertWithinRoot(file, join(root, 'nope'))).toThrow(/does not exist/)
  })

  // A case-insensitive file system is what makes the upper-cased root resolve.
  test.skipIf(process.platform !== 'win32')('on Windows the comparison ignores case', () => {
    const file = join(exportDir, 'draft.docx')
    writeFileSync(file, 'x')
    // Passing the root in a different case must still match on win32 — the
    // realpath of the root keeps the on-disk case, so this exercises the fold.
    expect(() => assertWithinRoot(file, exportDir.toUpperCase(), 'win32')).not.toThrow()
  })

  test.skipIf(process.platform !== 'win32')('a junction inside the export folder pointing out is refused', () => {
    const secret = join(outside, 'secret.pdf')
    writeFileSync(secret, 'x')
    // Junctions need no admin rights on Windows, unlike symlinks.
    symlinkSync(outside, join(exportDir, 'link'), 'junction')
    expect(() => assertWithinRoot(join(exportDir, 'link', 'secret.pdf'), exportDir)).toThrow(
      /only files inside/,
    )
  })

  test('loadSendableFile applies the root when given one', () => {
    const file = join(outside, 'x.pdf')
    writeFileSync(file, 'x')
    const stateDir = join(root, 'state')
    mkdirSync(stateDir)
    expect(() => loadSendableFile(file, stateDir, 1024, exportDir)).toThrow(/only files inside/)
  })
})

describe('B2: the queue never stores download URLs', () => {
  const ACTIVITY = load('dm-attachment-image.json')

  test('redaction removes every downloadUrl and nothing else', () => {
    const redacted = redactCredentials(ACTIVITY) as Record<string, any>
    expect(JSON.stringify(redacted)).not.toContain('downloadUrl')
    expect(redacted.id).toBe(ACTIVITY.id)
    expect(redacted.text).toBe(ACTIVITY.text)
    expect(redacted.attachments).toHaveLength(ACTIVITY.attachments.length)
  })

  test('the live activity is left intact for this delivery', () => {
    redactCredentials(ACTIVITY)
    expect(JSON.stringify(ACTIVITY)).toContain('downloadUrl')
  })

  test('what lands on disk has no downloadUrl', () => {
    const dir = join(root, 'queue')
    mkdirSync(dir)
    new IngressQueue(dir).enqueue(ACTIVITY)
    const files = readdirSync(dir)
    expect(files).toHaveLength(1)
    expect(readFileSync(join(dir, files[0]!), 'utf8')).not.toContain('downloadUrl')
  })

  test('an SDK-style activity (class instance with methods) is accepted', () => {
    // The live handler's activity is not plain JSON: it carries values that
    // structuredClone rejects ("The object can not be cloned", seen on the
    // first real Teams message), and every inbound message was dropped. An own
    // function property reproduces the same rejection.
    const live = { ...structuredClone(ACTIVITY), reply: () => undefined } as Record<string, unknown>
    const dir = join(root, 'queue-sdk')
    mkdirSync(dir)
    expect(new IngressQueue(dir).enqueue(live)).toBe(true)
    const written = readFileSync(join(dir, readdirSync(dir)[0]!), 'utf8')
    expect(written).not.toContain('downloadUrl')
    expect(written).toContain(String(ACTIVITY.id))
  })

  test('a replayed (redacted) activity registers no file handles', () => {
    const redacted = redactCredentials(ACTIVITY)
    expect(new AttachmentHandles().register(redacted)).toHaveLength(0)
  })
})

describe('B3: inbox cleanup', () => {
  test('files older than the TTL are deleted, newer ones kept', () => {
    const inbox = join(root, 'inbox')
    mkdirSync(inbox)
    const old = join(inbox, 'old.pdf')
    const fresh = join(inbox, 'fresh.pdf')
    writeFileSync(old, 'x')
    writeFileSync(fresh, 'x')
    const now = Date.now()
    const eightDaysAgo = (now - 8 * 24 * 60 * 60 * 1000) / 1000
    utimesSync(old, eightDaysAgo, eightDaysAgo)

    expect(purgeOlderThan(inbox, INBOX_TTL_MS, now)).toBe(1)
    expect(readdirSync(inbox)).toEqual(['fresh.pdf'])
  })

  test('a missing inbox is not an error', () => {
    expect(purgeOlderThan(join(root, 'nope'), INBOX_TTL_MS)).toBe(0)
  })
})

describe('Windows .env files', () => {
  const KEYS = ['MSTEAMS_FORKTEST_A', 'MSTEAMS_FORKTEST_B']
  const load = (content: string) => {
    for (const k of KEYS) delete process.env[k]
    const file = join(root, '.env')
    writeFileSync(file, content)
    loadEnvFile(file)
    return KEYS.map(k => process.env[k])
  }

  test('CRLF line endings (Notepad) are read', () => {
    expect(load('MSTEAMS_FORKTEST_A=one\r\nMSTEAMS_FORKTEST_B=two\r\n')).toEqual(['one', 'two'])
  })

  test('a UTF-8 BOM before the first key is ignored', () => {
    expect(load('\uFEFFMSTEAMS_FORKTEST_A=one\nMSTEAMS_FORKTEST_B=two\n')).toEqual(['one', 'two'])
  })

  test('a Windows path value keeps its backslashes', () => {
    expect(load('MSTEAMS_FORKTEST_A=C:\\AI Teams project\\bot-export\r\n')[0]).toBe('C:\\AI Teams project\\bot-export')
  })
})

describe('B4: download host allowlist', () => {
  test('SharePoint and OneDrive for Business hosts pass by default', () => {
    expect(isAllowedDownloadUrl('https://contoso-my.sharepoint.com/personal/x/file', DEFAULT_DOWNLOAD_HOSTS)).toBe(true)
    expect(isAllowedDownloadUrl('https://contoso.sharepoint.com/sites/x/file', DEFAULT_DOWNLOAD_HOSTS)).toBe(true)
  })

  test('other hosts, lookalikes and plain http are refused', () => {
    for (const url of [
      'https://evil.example.com/file',
      'https://sharepoint.com.evil.example/file',
      'https://evilsharepoint.com/file',
      'http://contoso.sharepoint.com/file',
      'https://169.254.169.254/latest/meta-data',
      'file:///C:/Windows/win.ini',
      'not a url',
    ]) {
      expect(isAllowedDownloadUrl(url, DEFAULT_DOWNLOAD_HOSTS)).toBe(false)
    }
  })

  test('the tenant can be pinned through MSTEAMS_DOWNLOAD_HOSTS', () => {
    const hosts = downloadHostsFromEnv(' .fabrikam.sharepoint.com , .fabrikam-my.sharepoint.com ')
    expect(isAllowedDownloadUrl('https://fabrikam-my.sharepoint.com/x', hosts)).toBe(true)
    expect(isAllowedDownloadUrl('https://contoso-my.sharepoint.com/x', hosts)).toBe(false)
  })

  test('an empty setting falls back to the default', () => {
    expect(downloadHostsFromEnv(undefined)).toEqual(DEFAULT_DOWNLOAD_HOSTS)
    expect(downloadHostsFromEnv(' , ')).toEqual(DEFAULT_DOWNLOAD_HOSTS)
  })

  test('a download to a disallowed host is refused before any request', async () => {
    const activity = structuredClone(load('dm-attachment-image.json'))
    for (const a of activity.attachments) {
      if (a?.content) a.content.downloadUrl = 'https://evil.example.com/steal'
    }
    const handles = new AttachmentHandles()
    const id = handles.register(activity)[0]!.id
    const inbox = join(root, 'inbox')
    mkdirSync(inbox)

    const err = await handles.download(id, inbox).catch((e: Error) => e)
    expect((err as Error).message).toMatch(/not on an allowed Microsoft host/)
    expect((err as Error).message).not.toContain('evil.example.com')
    expect(readdirSync(inbox)).toHaveLength(0)
  })
})
