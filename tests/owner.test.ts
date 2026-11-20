/** Fork (20/11): owner mode for a person's own session (src/owner.ts). */

import { describe, expect, test } from 'bun:test'
import { existsSync, mkdtempSync, readFileSync, statSync, writeFileSync, mkdirSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import {
  ownerFromEnv,
  refusedNotOwner,
  isOwnersChat,
  ownerOutboundRefusal,
  routesPermissions,
  writePermissionRoute,
  removePermissionRoute,
  clearOwnPermissionRoutes,
  PERMISSION_ROUTES_DIR,
} from '../src/owner.js'

const OWNER = '11111111-1111-1111-1111-111111111111'
const OTHER = '22222222-2222-2222-2222-222222222222'
const ownersDm = { conversationType: 'personal', senderId: OWNER }
const othersDm = { conversationType: 'personal', senderId: OTHER }
const group = { conversationType: 'groupChat', senderId: OTHER }

describe('owner from the environment', () => {
  test('normalized, and unset or blank means no owner', () => {
    expect(ownerFromEnv({ BOT_OWNER: ` ${OWNER.toUpperCase()} ` })).toBe(OWNER)
    expect(ownerFromEnv({})).toBeUndefined()
    expect(ownerFromEnv({ BOT_OWNER: '  ' })).toBeUndefined()
  })
})

describe('inbound', () => {
  test('without an owner nobody is refused here', () => {
    expect(refusedNotOwner(undefined, OTHER)).toBe(false)
  })

  test('with one, everyone but the owner is', () => {
    expect(refusedNotOwner(OWNER, OWNER.toUpperCase())).toBe(false)
    expect(refusedNotOwner(OWNER, OTHER)).toBe(true)
    expect(refusedNotOwner(OWNER, '')).toBe(true)
  })
})

describe('outbound', () => {
  test('without an owner, nothing changes, turn or not', () => {
    for (const ref of [ownersDm, othersDm, group]) {
      expect(ownerOutboundRefusal(undefined, ref, true)).toBeNull()
      expect(ownerOutboundRefusal(undefined, ref, false)).toBeNull()
    }
  })

  test("a 1:1 chat is only ever the owner's own", () => {
    expect(isOwnersChat(OWNER, ownersDm)).toBe(true)
    expect(isOwnersChat(OWNER, { conversationType: 'personal', senderId: OWNER.toUpperCase() })).toBe(true)
    expect(isOwnersChat(OWNER, { conversationType: 'personal' })).toBe(false)
    expect(isOwnersChat(OWNER, { conversationType: 'groupChat', senderId: OWNER })).toBe(false)
    expect(ownerOutboundRefusal(OWNER, ownersDm, true)).toBeNull()
    expect(ownerOutboundRefusal(OWNER, othersDm, true)).not.toBeNull()
  })

  test('during a turn group chats keep the usual rules', () => {
    expect(ownerOutboundRefusal(OWNER, group, true)).toBeNull()
  })

  test("with no turn running, only the owner's own 1:1 chat", () => {
    expect(ownerOutboundRefusal(OWNER, ownersDm, false)).toBeNull()
    expect(ownerOutboundRefusal(OWNER, group, false)).not.toBeNull()
    expect(ownerOutboundRefusal(OWNER, othersDm, false)).not.toBeNull()
  })
})

describe('permission routes', () => {
  test('written only in the multi-session layout', () => {
    expect(routesPermissions({})).toBe(false)
    expect(routesPermissions({ BOT_WORKER: 'an' })).toBe(true)
    expect(routesPermissions({ MSTEAMS_ROOT_DIR: '/root' })).toBe(true)
  })

  test('holds the port as decimal text, and goes away again', () => {
    const root = mkdtempSync(join(tmpdir(), 'routes-'))
    expect(writePermissionRoute(root, 'abcde', 4001)).toBe(true)
    const file = join(root, PERMISSION_ROUTES_DIR, 'abcde')
    expect(readFileSync(file, 'utf8')).toBe('4001')
    if (process.platform !== 'win32') {
      expect(statSync(file).mode & 0o777).toBe(0o600)
      expect(statSync(join(root, PERMISSION_ROUTES_DIR)).mode & 0o777).toBe(0o700)
    }
    removePermissionRoute(root, 'abcde')
    expect(existsSync(file)).toBe(false)
    removePermissionRoute(root, 'abcde') // twice is fine
  })

  test('an id that could leave the dir is refused', () => {
    const root = mkdtempSync(join(tmpdir(), 'routes-'))
    expect(writePermissionRoute(root, '../x', 4001)).toBe(false)
    expect(writePermissionRoute(root, '', 4001)).toBe(false)
    expect(existsSync(join(root, 'x'))).toBe(false)
  })

  test("at boot a session clears only its own port's notes", () => {
    const root = mkdtempSync(join(tmpdir(), 'routes-'))
    expect(clearOwnPermissionRoutes(root, 4001)).toBe(0) // no dir yet
    writePermissionRoute(root, 'abcde', 4001)
    writePermissionRoute(root, 'fghij', 4002)
    mkdirSync(join(root, PERMISSION_ROUTES_DIR), { recursive: true })
    writeFileSync(join(root, PERMISSION_ROUTES_DIR, 'kmnop'), '4001\n')
    expect(clearOwnPermissionRoutes(root, 4001)).toBe(2)
    expect(existsSync(join(root, PERMISSION_ROUTES_DIR, 'fghij'))).toBe(true)
  })
})
