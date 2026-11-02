/**
 * Signs a person's mailbox in for the m365 server. Run on the bot machine by
 * whoever operates it; the person signs in on their own device with the code
 * it shows, so their password never touches this machine.
 *
 *   cd "C:\AI Teams project\msteams-channel-fork"
 *   bun m365/login.ts                    sign someone in (device code)
 *   bun m365/login.ts --list             who is signed in
 *   bun m365/login.ts --logout <id>      forget one person (Entra object id)
 */

import { readdirSync } from 'fs'
import { loadConfig } from './config.js'
import { TokenStore } from './token-store.js'
import { deviceCodeSignIn } from './auth.js'
import { tokenPath } from './users.js'

const config = loadConfig()

if (process.argv.includes('--list')) {
  let files: string[] = []
  try {
    files = readdirSync(config.tokenDir).filter(f => f.endsWith('.dpapi'))
  } catch {}
  console.log(files.length ? files.map(f => f.replace('.dpapi', '')).join('\n') : 'Chưa ai đăng nhập.')
  process.exit(0)
}

const logout = process.argv.indexOf('--logout')
if (logout !== -1) {
  const id = String(process.argv[logout + 1] ?? '').toLowerCase()
  new TokenStore(tokenPath(config.tokenDir, id)).clear()
  console.log(`Đã xóa phiên đăng nhập email của ${id}.`)
  process.exit(0)
}

if (!config.clientId || !config.tenantId) {
  console.error(`Thiếu M365_CLIENT_ID hoặc M365_TENANT_ID trong ${config.stateDir}\\.env`)
  process.exit(2)
}

try {
  const signedIn = await deviceCodeSignIn(config.clientId, config.tenantId, message => {
    console.log(message)
    console.log('Người cần đăng nhập mở link trên máy của họ, nhập mã, và đăng nhập bằng tài khoản công ty của chính họ.\n')
  })
  const me = (await (
    await fetch('https://graph.microsoft.com/v1.0/me?$select=id,userPrincipalName,displayName', {
      headers: { Authorization: `Bearer ${signedIn.accessToken}` },
    })
  ).json()) as { id?: string; userPrincipalName?: string; displayName?: string }
  if (!me.id) throw new Error('could not read who signed in')
  const path = tokenPath(config.tokenDir, me.id.toLowerCase())
  await new TokenStore(path).save(signedIn.refreshToken)
  console.log(`Đã đăng nhập: ${me.displayName} <${me.userPrincipalName}> (${me.id})`)
  console.log(`Quyền: ${signedIn.scope}`)
  console.log(`Lưu mã hóa tại ${path}`)
} catch (err) {
  console.error(`Không đăng nhập được: ${err instanceof Error ? err.message : err}`)
  process.exit(1)
}
