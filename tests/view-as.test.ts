/**
 * Fork addition (02/11): /useransw, an admin answered as a non-admin.
 */

import { describe, expect, test } from 'bun:test'
import { mkdtempSync, readFileSync, statSync, writeFileSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { parseViewAsCommand, applyViewAs, readViewAs, effectiveRole, stripViewAsTrigger, VIEW_AS_FILE } from '../src/view-as.js'

describe('the command', () => {
  test('bare, on and off; case and surrounding space ignored', () => {
    expect(parseViewAsCommand('/useransw')).toBe('toggle')
    expect(parseViewAsCommand('  /UserAnsw on ')).toBe('on')
    expect(parseViewAsCommand('/useransw off')).toBe('off')
    expect(parseViewAsCommand('/useransw tắt')).toBe('off')
  })

  test('anything around it is chat', () => {
    for (const text of ['/useransw please', 'hi /useransw', '/useranswer', 'useransw', '']) {
      expect(parseViewAsCommand(text)).toBeUndefined()
    }
  })
})

describe('the store', () => {
  test('toggle, on and off per admin, kept across reads', () => {
    const dir = mkdtempSync(join(tmpdir(), 'msteams-viewas-'))
    expect(applyViewAs(dir, 'sam', 'toggle')).toBe(true)
    expect(applyViewAs(dir, 'kien', 'on')).toBe(true)
    expect([...readViewAs(dir)].sort()).toEqual(['kien', 'sam'])
    expect(applyViewAs(dir, 'sam', 'toggle')).toBe(false)
    expect(applyViewAs(dir, 'kien', 'off')).toBe(false)
    expect(readViewAs(dir).size).toBe(0)
    if (process.platform !== 'win32') expect(statSync(join(dir, VIEW_AS_FILE)).mode & 0o777).toBe(0o600)
  })

  test('a missing or broken file means nobody', () => {
    const dir = mkdtempSync(join(tmpdir(), 'msteams-viewas-'))
    expect(readViewAs(dir).size).toBe(0)
    writeFileSync(join(dir, VIEW_AS_FILE), '{not json')
    expect(readViewAs(dir).size).toBe(0)
    expect(readFileSync(join(dir, VIEW_AS_FILE), 'utf8')).toBe('{not json')
  })
})

describe('the role', () => {
  test('only lowers: an admin viewing as a user is a user, a user stays a user', () => {
    const on = new Set(['sam', 'boss'])
    expect(effectiveRole('admin', 'sam', on)).toBe('user')
    expect(effectiveRole('admin', 'kien', on)).toBe('admin')
    expect(effectiveRole('user', 'boss', on)).toBe('user')
    expect(effectiveRole('user', 'boss', new Set())).toBe('user')
  })
})

describe('the trigger inside a sentence', () => {
  test('removes the word and keeps the rest, wherever it sits', () => {
    expect(stripViewAsTrigger('lịch hôm nay của tôi /useransw')).toBe('lịch hôm nay của tôi')
    expect(stripViewAsTrigger('/useransw lịch hôm nay của tôi')).toBe('lịch hôm nay của tôi')
    expect(stripViewAsTrigger('lịch /UserAnsw hôm nay')).toBe('lịch hôm nay')
    expect(stripViewAsTrigger('dòng 1\n/useransw dòng 2')).toBe('dòng 1\ndòng 2')
  })

  test('no trigger, or the word glued to something else: undefined', () => {
    for (const text of ['lịch hôm nay của tôi', 'xem/useransw', '/useranswer abc', 'a/useransw b', '']) {
      expect(stripViewAsTrigger(text)).toBeUndefined()
    }
  })

  test('every exact command is still a command, checked before the trigger', () => {
    for (const text of ['/useransw', '/useransw on', '/useransw off']) expect(parseViewAsCommand(text)).toBeDefined()
    expect(parseViewAsCommand('lịch hôm nay /useransw')).toBeUndefined()
  })
})
