/**
 * Delegated Microsoft Graph sign-in for the "Teams Claude - Graph" app.
 *
 * The user signs in once with a device code (`bun m365/login.ts`); after that
 * the server trades the stored refresh token for short-lived access tokens.
 * Entra rotates the refresh token on every use, so each new one is stored.
 *
 * Tokens never reach a log, an error message or the model.
 */

import type { TokenStore } from './token-store.js'

/**
 * ".default" is every delegated permission IT has consented to for the app,
 * whatever that is today. Naming them one by one would break every sign-in
 * the moment a permission is added to the app before IT grants it. Consent is
 * tenant-wide, so a newly granted permission reaches existing sign-ins
 * without anyone signing in again.
 */
export const SCOPES = 'https://graph.microsoft.com/.default offline_access'

export const NOT_SIGNED_IN =
  'This person has not signed in their mailbox for the bot. On the bot machine, in a normal PowerShell window, run ' +
  'cd "C:\\AI Teams project\\msteams-channel-fork"; bun m365/login.ts — then the person opens the link it shows and enters the code with their own account'

type TokenResponse = {
  access_token?: string
  refresh_token?: string
  expires_in?: number
  scope?: string
  error?: string
  error_description?: string
}

export type Fetch = (input: string, init?: RequestInit) => Promise<Response>

function authority(tenantId: string): string {
  return `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0`
}

/** First line only: Entra appends trace and correlation ids. */
function reason(res: TokenResponse): string {
  return `${res.error}: ${String(res.error_description ?? '').split(/\r?\n/)[0]}`
}

export class GraphAuth {
  private cached: { token: string; expiresAt: number } | null = null

  constructor(
    private readonly clientId: string,
    private readonly tenantId: string,
    private readonly store: TokenStore,
    private readonly fetchImpl: Fetch = fetch,
    private readonly now: () => number = Date.now,
  ) {}

  async accessToken(): Promise<string> {
    if (this.cached && this.cached.expiresAt > this.now()) return this.cached.token

    const refresh = await this.store.load()
    if (!refresh) throw new Error(NOT_SIGNED_IN)

    const res = (await (
      await this.fetchImpl(`${authority(this.tenantId)}/token`, {
        method: 'POST',
        body: new URLSearchParams({
          grant_type: 'refresh_token',
          client_id: this.clientId,
          refresh_token: refresh,
          scope: SCOPES,
        }),
      })
    ).json()) as TokenResponse

    if (!res.access_token) {
      if (res.error === 'invalid_grant') {
        // Revoked, expired, or the password changed: only a new sign-in helps.
        this.store.clear()
        throw new Error(`${NOT_SIGNED_IN} (the previous sign-in is no longer valid)`)
      }
      throw new Error(`Microsoft sign-in failed: ${reason(res)}`)
    }
    if (res.refresh_token) await this.store.save(res.refresh_token)
    // A minute of margin so a token is never used in its last seconds.
    this.cached = { token: res.access_token, expiresAt: this.now() + (Number(res.expires_in ?? 3600) - 60) * 1000 }
    return res.access_token
  }
}

/**
 * Interactive sign-in, for the login CLI only. `show` receives Microsoft's
 * "open this page and enter this code" text. The caller stores the refresh
 * token once it knows whose it is.
 */
export async function deviceCodeSignIn(
  clientId: string,
  tenantId: string,
  show: (message: string) => void,
  fetchImpl: Fetch = fetch,
): Promise<{ refreshToken: string; accessToken: string; scope: string }> {
  const form = (fields: Record<string, string>) => ({ method: 'POST', body: new URLSearchParams(fields) })
  const code = (await (
    await fetchImpl(`${authority(tenantId)}/devicecode`, form({ client_id: clientId, scope: SCOPES }))
  ).json()) as { device_code?: string; message?: string; interval?: number; expires_in?: number } & TokenResponse
  if (!code.device_code) throw new Error(`could not start sign-in: ${reason(code)}`)
  show(String(code.message))

  let interval = Number(code.interval ?? 5) * 1000
  const deadline = Date.now() + Number(code.expires_in ?? 900) * 1000
  while (Date.now() < deadline) {
    await Bun.sleep(interval)
    const res = (await (
      await fetchImpl(
        `${authority(tenantId)}/token`,
        form({
          grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
          client_id: clientId,
          device_code: code.device_code,
        }),
      )
    ).json()) as TokenResponse
    if (res.error === 'authorization_pending') continue
    if (res.error === 'slow_down') {
      interval += 5000
      continue
    }
    if (res.error || !res.refresh_token || !res.access_token) throw new Error(`sign-in failed: ${reason(res)}`)
    return { refreshToken: res.refresh_token, accessToken: res.access_token, scope: String(res.scope ?? '') }
  }
  throw new Error('sign-in timed out; run it again')
}
