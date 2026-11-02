/**
 * Fork addition (02/11): the technical log kept out of a non-admin's chat.
 */

import { describe, expect, test } from 'bun:test'
import { mkdtempSync, readFileSync, statSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { appendAssistantLog, assistantLogPath, MAX_LOG_TEXT } from '../src/assistant-log.js'

describe('assistant log', () => {
  test('one JSON line per entry, with a timestamp', () => {
    const dir = mkdtempSync(join(tmpdir(), 'msteams-log-'))
    appendAssistantLog(dir, { kind: 'note', sender: 'boss', text: 'files_get: 403' }, Date.parse('2026-11-02T02:00:00Z'))
    appendAssistantLog(dir, { kind: 'permission_expired', request_id: 'qnyia', text: 'Read denied' })
    const lines = readFileSync(assistantLogPath(dir), 'utf8').trim().split('\n').map(l => JSON.parse(l))
    expect(lines).toHaveLength(2)
    expect(lines[0]).toEqual({ at: '2026-11-02T02:00:00.000Z', kind: 'note', sender: 'boss', text: 'files_get: 403' })
    expect(lines[1].request_id).toBe('qnyia')
  })

  test('the file is private, control characters are dropped and long text is capped', () => {
    const dir = mkdtempSync(join(tmpdir(), 'msteams-log-'))
    appendAssistantLog(dir, { kind: 'note', text: `a\x1b[2Jb\nc${'x'.repeat(MAX_LOG_TEXT)}` })
    const entry = JSON.parse(readFileSync(assistantLogPath(dir), 'utf8'))
    expect(entry.text.startsWith('a[2Jb\nc')).toBe(true)
    expect(entry.text.length).toBe(MAX_LOG_TEXT)
    if (process.platform !== 'win32') expect(statSync(assistantLogPath(dir)).mode & 0o777).toBe(0o600)
  })
})
