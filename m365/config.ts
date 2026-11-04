/**
 * Settings for the m365 server and its login CLI.
 *
 * State lives beside the msteams channel's, in its own folder:
 * `%USERPROFILE%\.claude\channels\m365` (M365_STATE_DIR overrides it).
 * Its `.env` holds no secret, only which app and tenant to sign in to and
 * the attachment folder; each person's refresh token is in `tokens\`.
 */

import { existsSync, readFileSync } from 'fs'
import { homedir } from 'os'
import { join } from 'path'
import { loadEnvFile } from '../src/env.js'
import { stateDir as msteamsStateDir } from '../src/state.js'

export function m365StateDir(): string {
  return process.env.M365_STATE_DIR ?? join(homedir(), '.claude', 'channels', 'm365')
}

export type M365Config = {
  stateDir: string
  clientId?: string
  tenantId?: string
  /** The export folder (M365_ATTACH_ROOT) and the Teams inbox. */
  attachRoots: string[]
  /** Where email attachments are saved: `<export folder>email-attachments`. */
  downloadDir?: string
  /** Where SharePoint and OneDrive files are saved: `<export folder>\sharepoint-files`. */
  filesDir?: string
  internalDomains: string[]
  /** One encrypted refresh token per person: `<object id>.dpapi` (DPAPI on Windows, AES key file elsewhere). */
  tokenDir: string
  sendLog: string
  /** `{"<object id>": ["line of their signature to leave out", ...]}`; optional. */
  signatureOmitFile: string
  /** Where the msteams channel writes turn.json (who is being answered). */
  msteamsStateDir: string
}

export function loadConfig(): M365Config {
  const stateDir = m365StateDir()
  loadEnvFile(join(stateDir, '.env'))
  return {
    stateDir,
    clientId: process.env.M365_CLIENT_ID || undefined,
    tenantId: process.env.M365_TENANT_ID || undefined,
    attachRoots: process.env.M365_ATTACH_ROOT ? [process.env.M365_ATTACH_ROOT, join(msteamsStateDir(), 'inbox')] : [],
    downloadDir: process.env.M365_ATTACH_ROOT ? join(process.env.M365_ATTACH_ROOT, 'email-attachments') : undefined,
    filesDir: process.env.M365_ATTACH_ROOT ? join(process.env.M365_ATTACH_ROOT, 'sharepoint-files') : undefined,
    internalDomains: (process.env.M365_INTERNAL_DOMAINS ?? 'fabrikam.com')
      .split(',')
      .map(d => d.trim().toLowerCase())
      .filter(Boolean),
    tokenDir: join(stateDir, 'tokens'),
    sendLog: join(stateDir, 'sent.jsonl'),
    signatureOmitFile: join(stateDir, 'signature-omit.json'),
    msteamsStateDir: msteamsStateDir(),
  }
}

/**
 * The signature lines one person wants left out of drafts, from
 * signature-omit.json. Read whenever a signature is looked up (at most hourly
 * per person), so an edit needs no restart; a missing or broken file leaves
 * the signature whole.
 */
export function signatureOmit(file: string, objectId: string): string[] {
  if (!existsSync(file)) return []
  try {
    const lines = JSON.parse(readFileSync(file, 'utf8'))?.[objectId.toLowerCase()]
    return Array.isArray(lines) ? lines.filter((l): l is string => typeof l === 'string') : []
  } catch {
    return []
  }
}
