/**
 * Fork addition (18/11): why an inbound activity was refused, on disk.
 *
 * The refusal line on stderr is invisible during a live session (mid-session
 * MCP stderr reaches no log), which left a refused group-chat tag with no
 * trace at all. This writes one line per refusal to `refused.log` in the state
 * dir: the reason and the shape of the activity, never its text, never an id
 * (the sender, conversation, tenant and bot ids are reduced to yes/no and
 * prefixes). Kept small: past MAX_BYTES the file starts over.
 */

import { appendFileSync, statSync, writeFileSync } from 'fs'
import { join } from 'path'

export const REFUSAL_LOG = 'refused.log'
const MAX_BYTES = 256 * 1024

type Mention = { type?: string; mentioned?: { id?: unknown } }

/** One line describing a refused activity. Pure: no ids or text leave it. */
export function describeRefusal(
  activity: Record<string, any>,
  reason: string,
  configuredTenantId: string,
  now = new Date(),
): string {
  const tenant = String(activity.channelData?.tenant?.id ?? activity.conversation?.tenantId ?? '')
  const recipient = String(activity.recipient?.id ?? '')
  const mentions = (Array.isArray(activity.entities) ? activity.entities : []).filter(
    (e: Mention) => e?.type === 'mention',
  ) as Mention[]
  const ids = mentions.map(m => String(m.mentioned?.id ?? ''))
  const prefix = (id: string) => (id.includes(':') ? id.slice(0, id.indexOf(':') + 1) : id ? 'bare' : 'none')
  const fields = [
    now.toISOString(),
    `reason=${reason}`,
    `type=${String(activity.conversation?.conversationType ?? 'none')}`,
    `activity=${String(activity.type ?? 'none')}`,
    `tenant_match=${!!tenant && tenant.toLowerCase() === configuredTenantId.toLowerCase()}`,
    `has_aad=${!!activity.from?.aadObjectId}`,
    `mentions=${mentions.length}`,
    `bot_mentioned=${!!recipient && ids.includes(recipient)}`,
    `recipient=${prefix(recipient)}`,
    `mention_ids=${ids.map(prefix).join(',') || '-'}`,
  ]
  return fields.join(' ')
}

/** Append `line`; never throws, since a log must not break the inbound path. */
export function appendRefusal(stateDir: string, line: string): void {
  const path = join(stateDir, REFUSAL_LOG)
  try {
    let size = 0
    try {
      size = statSync(path).size
    } catch {}
    if (size > MAX_BYTES) writeFileSync(path, '', { mode: 0o600 })
    appendFileSync(path, `${line}\n`, { mode: 0o600 })
  } catch {}
}
