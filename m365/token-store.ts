/**
 * Where the Microsoft Graph refresh token lives between sessions.
 *
 * The refresh token lets anyone holding it act in the user's mailbox for
 * months, so it is never stored in the clear:
 *
 * - Windows: DPAPI (CurrentUser scope): only the same Windows account on the
 *   same machine can decrypt it. DPAPI is reached through Windows PowerShell,
 *   with the data on stdin: argv is visible to every process on the machine,
 *   stdin is not.
 * - Linux and macOS: AES-256-GCM with a key generated on first use and kept in
 *   `~/.config/teams-claude-m365/token.key` (0600), outside the token folder. A copy
 *   of the token folder alone (a backup, a stray sync) is useless. Anyone who
 *   can read the bot account's home can read both, as with DPAPI and that
 *   account's login; the managed settings deny Claude reading either path.
 */

import { createCipheriv, createDecipheriv, randomBytes } from 'crypto'
import { existsSync, readFileSync, writeFileSync, renameSync, rmSync, mkdirSync } from 'fs'
import { homedir } from 'os'
import { dirname, join } from 'path'

export type Protector = {
  protect(plain: Buffer): Promise<Buffer>
  unprotect(sealed: Buffer): Promise<Buffer>
}

/** No encryption: for tests only. */
export const plainProtector: Protector = {
  protect: async plain => plain,
  unprotect: async sealed => sealed,
}

async function dpapi(op: 'Protect' | 'Unprotect', data: Buffer): Promise<Buffer> {
  const script =
    'Add-Type -AssemblyName System.Security; ' +
    '$in = [Convert]::FromBase64String([Console]::In.ReadToEnd()); ' +
    `$out = [Security.Cryptography.ProtectedData]::${op}($in, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser); ` +
    '[Console]::Out.Write([Convert]::ToBase64String($out))'
  const proc = Bun.spawn(['powershell.exe', '-NoProfile', '-NonInteractive', '-Command', script], {
    stdin: new TextEncoder().encode(data.toString('base64')),
    stdout: 'pipe',
    stderr: 'pipe',
  })
  const [out, code] = await Promise.all([new Response(proc.stdout).text(), proc.exited])
  // stderr is not echoed: on Unprotect it could name the blob it failed on.
  if (code !== 0 || !out.trim()) throw new Error(`DPAPI ${op.toLowerCase()} failed (exit ${code})`)
  return Buffer.from(out.trim(), 'base64')
}

export const dpapiProtector: Protector = {
  protect: plain => dpapi('Protect', plain),
  unprotect: sealed => dpapi('Unprotect', sealed),
}

/** Version byte, then 12-byte IV, 16-byte GCM tag, ciphertext. */
const SEALED_V1 = 1

export function defaultKeyPath(): string {
  return process.env.M365_TOKEN_KEY_FILE ?? join(homedir(), '.config', 'teams-claude-m365', 'token.key')
}

/** The 32-byte key at `path`, created (0600, in a 0700 folder) when missing. */
function loadOrCreateKey(path: string): Buffer {
  if (!existsSync(path)) {
    mkdirSync(dirname(path), { recursive: true, mode: 0o700 })
    // wx: two processes racing to create it must not overwrite each other's key.
    try {
      writeFileSync(path, randomBytes(32), { mode: 0o600, flag: 'wx' })
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code !== 'EEXIST') throw err
    }
  }
  const key = readFileSync(path)
  if (key.length !== 32) throw new Error(`token key ${path} is not 32 bytes`)
  return key
}

export function keyFileProtector(keyPath: string = defaultKeyPath()): Protector {
  return {
    protect: async plain => {
      const iv = randomBytes(12)
      const cipher = createCipheriv('aes-256-gcm', loadOrCreateKey(keyPath), iv)
      const body = Buffer.concat([cipher.update(plain), cipher.final()])
      return Buffer.concat([Buffer.from([SEALED_V1]), iv, cipher.getAuthTag(), body])
    },
    unprotect: async sealed => {
      if (sealed.length < 29 || sealed[0] !== SEALED_V1) throw new Error('stored token is not in the expected format')
      const decipher = createDecipheriv('aes-256-gcm', loadOrCreateKey(keyPath), sealed.subarray(1, 13))
      decipher.setAuthTag(sealed.subarray(13, 29))
      return Buffer.concat([decipher.update(sealed.subarray(29)), decipher.final()])
    },
  }
}

export function defaultProtector(platform = process.platform): Protector {
  return platform === 'win32' ? dpapiProtector : keyFileProtector()
}

export class TokenStore {
  constructor(
    private readonly path: string,
    private readonly protector: Protector = defaultProtector(),
  ) {}

  /** The stored refresh token, or null when the user has not signed in. */
  async load(): Promise<string | null> {
    let sealed: Buffer
    try {
      sealed = readFileSync(this.path)
    } catch {
      return null
    }
    return (await this.protector.unprotect(sealed)).toString('utf8')
  }

  /** Atomic: a crash mid-write must not leave a token that cannot be read. */
  async save(refreshToken: string): Promise<void> {
    mkdirSync(dirname(this.path), { recursive: true, mode: 0o700 })
    const sealed = await this.protector.protect(Buffer.from(refreshToken, 'utf8'))
    const tmp = `${this.path}.tmp`
    writeFileSync(tmp, sealed, { mode: 0o600 })
    renameSync(tmp, this.path)
  }

  clear(): void {
    rmSync(this.path, { force: true })
  }
}
