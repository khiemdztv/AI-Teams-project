/**
 * Fork (18/11): fetch a Teams meeting's transcript, for a summary.
 *
 * Uses only the per-meeting RSC permissions the app asks for (manifest 1.4.0),
 * granted when someone adds the bot to the meeting chat:
 *   - OnlineMeeting.ReadBasic.Chat: the Bot Framework meeting API turns the
 *     meeting id carried by activities in the meeting chat
 *     (`channelData.meeting.id`) into the Graph meeting id and the organizer.
 *   - OnlineMeetingTranscript.Read.Chat: Graph, with the app's own token, lists
 *     and downloads that meeting's transcripts.
 * Nothing tenant-wide: a meeting whose chat does not have the bot is refused
 * by Graph.
 *
 * The transcript is turned from WebVTT into plain "[hh:mm:ss] Name: text"
 * lines (consecutive lines of one speaker merged) and written to a file in the
 * inbox, which Claude reads; it is never put in a tool result or a log.
 */

import { writeFileSync, mkdirSync } from 'fs'
import { join } from 'path'
import { createHash } from 'crypto'

const GRAPH = 'https://graph.microsoft.com/v1.0'

/** The Bot Framework meeting id seen in each meeting chat, newest wins. In memory. */
export class MeetingChats {
  private readonly byConversation = new Map<string, string>()

  /** Record the meeting id an inbound activity carries, if any. */
  note(conversationId: string, activity: { channelData?: { meeting?: { id?: unknown } } }): void {
    const id = activity.channelData?.meeting?.id
    if (conversationId && typeof id === 'string' && id) this.byConversation.set(conversationId, id)
  }

  meetingOf(conversationId: string): string | undefined {
    return this.byConversation.get(conversationId)
  }
}

export type MeetingLookup = (meetingId: string) => Promise<{
  details?: { msGraphResourceId?: string; title?: string; scheduledStartTime?: string }
  organizer?: { aadObjectId?: string }
}>

export type TranscriptDeps = {
  getMeeting: MeetingLookup
  graphToken: () => Promise<string>
  fetchImpl?: typeof fetch
  saveDir: string
  now?: () => Date
}

export type TranscriptResult = {
  path: string
  title: string
  transcripts: number
  lines: number
  speakers: string[]
}

type Transcript = { id: string; createdDateTime?: string; endDateTime?: string }

/** WebVTT -> "[hh:mm:ss] Name: text", one speaker's consecutive lines merged. */
export function vttToText(vtt: string): { text: string; lines: number; speakers: string[] } {
  const cue = /^(?:(\d+):)?(\d{1,2}):(\d{2})[.,]\d{1,3}\s+-->/
  const out: { at: string; who: string; said: string }[] = []
  const speakers = new Set<string>()
  const rows = vtt.replace(/\r/g, '').split('\n')
  for (let i = 0; i < rows.length; i++) {
    const m = cue.exec(rows[i]!.trim())
    if (!m) continue
    const h = Number(m[1] ?? 0)
    const at = `${String(h).padStart(2, '0')}:${m[2]!.padStart(2, '0')}:${m[3]}`
    const body: string[] = []
    while (i + 1 < rows.length && rows[i + 1]!.trim() !== '') body.push(rows[++i]!.trim())
    const joined = body.join(' ')
    const voice = /<v\s+([^>]+)>/.exec(joined)
    const who = (voice?.[1] ?? '').trim() || '?'
    const said = joined.replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim()
    if (!said) continue
    speakers.add(who)
    const last = out[out.length - 1]
    if (last && last.who === who) last.said += ` ${said}`
    else out.push({ at, who, said })
  }
  return { text: out.map(r => `[${r.at}] ${r.who}: ${r.said}`).join('\n'), lines: out.length, speakers: [...speakers] }
}

/** A short, human reason for a Graph or meeting API failure; never the response body. */
function why(status: number, code: string): string {
  if (status === 403 && /GraphAccessToTranscriptsDisabled/i.test(code)) {
    return 'the tenant turned off Graph access to transcripts (Teams admin setting)'
  }
  if (status === 403 || status === 401) {
    return 'no permission: the bot must be in this meeting\'s chat with the app version that asks for meeting transcripts (1.4.0), and the chat must have consented'
  }
  if (status === 404) return 'the meeting or its transcript was not found'
  return `request failed (${status}${code ? ` ${code}` : ''})`
}

export async function fetchMeetingTranscript(meetingId: string, deps: TranscriptDeps): Promise<TranscriptResult> {
  const fetchImpl = deps.fetchImpl ?? fetch
  let info: Awaited<ReturnType<MeetingLookup>>
  try {
    info = await deps.getMeeting(meetingId)
  } catch (err: any) {
    const status = Number(err?.response?.status ?? err?.status ?? 0)
    throw new Error(`could not read the meeting: ${why(status, '')}`)
  }
  const graphId = info?.details?.msGraphResourceId
  const organizer = info?.organizer?.aadObjectId
  if (!graphId || !organizer) throw new Error('could not read the meeting: no Graph meeting id or organizer returned')
  const title = info.details?.title?.trim() || 'Cuộc họp'

  const token = await deps.graphToken()
  const base = `${GRAPH}/users/${encodeURIComponent(organizer)}/onlineMeetings/${encodeURIComponent(graphId)}/transcripts`
  const get = async (url: string, accept?: string) => {
    const res = await fetchImpl(url, { headers: { Authorization: `Bearer ${token}`, ...(accept ? { Accept: accept } : {}) } })
    if (!res.ok) {
      const body = (await res.json().catch(() => ({}))) as any
      const code = String(body?.error?.innerError?.code ?? body?.error?.code ?? '')
      throw new Error(why(res.status, code))
    }
    return res
  }

  const list = ((await (await get(base)).json()) as { value?: Transcript[] }).value ?? []
  if (!list.length) {
    throw new Error('no transcript yet: transcription was not turned on in this meeting, or the meeting has not ended')
  }
  list.sort((a, b) => String(a.createdDateTime ?? '').localeCompare(String(b.createdDateTime ?? '')))

  const parts: string[] = []
  let lines = 0
  const speakers = new Set<string>()
  for (const [n, t] of list.entries()) {
    const vtt = await (await get(`${base}/${encodeURIComponent(t.id)}/content?$format=text/vtt`, 'text/vtt')).text()
    const r = vttToText(vtt)
    lines += r.lines
    r.speakers.forEach(s => speakers.add(s))
    parts.push(list.length > 1 ? `--- Phần ${n + 1}/${list.length} (${t.createdDateTime ?? '?'}) ---\n${r.text}` : r.text)
  }

  mkdirSync(deps.saveDir, { recursive: true, mode: 0o700 })
  const stamp = (deps.now?.() ?? new Date()).toISOString().slice(0, 10)
  const hash = createHash('sha256').update(meetingId).digest('hex').slice(0, 10)
  const path = join(deps.saveDir, `transcript-${stamp}-${hash}.txt`)
  const header = `# ${title}\n# Bắt đầu: ${info.details?.scheduledStartTime ?? '?'}\n# Người nói: ${[...speakers].join(', ')}\n\n`
  writeFileSync(path, header + parts.join('\n\n') + '\n', { mode: 0o600 })
  return { path, title, transcripts: list.length, lines, speakers: [...speakers] }
}
