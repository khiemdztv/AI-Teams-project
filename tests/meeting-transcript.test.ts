import { describe, expect, test } from 'bun:test'
import { mkdtempSync, readFileSync, statSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import { MeetingChats, fetchMeetingTranscript, vttToText, type MeetingLookup } from '../src/meeting-transcript.js'

const VTT = `WEBVTT

1a2b/12-0
00:00:01.120 --> 00:00:04.500
<v Le Van An>Chào mọi người,</v>

1a2b/13-0
00:00:04.600 --> 00:00:06.000
<v Le Van An>hôm nay bàn hợp đồng ABC.</v>

00:00:07.000 --> 00:00:09.250
<v Tran Thanh Binh>Em đã gửi bản V2.</v>

1:02:03.000 --> 1:02:05.000
<v Tran Thanh Binh>Hết ạ.</v>
`

describe('vttToText', () => {
  test('speaker lines, merged per speaker, with hh:mm:ss', () => {
    const r = vttToText(VTT)
    expect(r.text).toBe(
      [
        '[00:00:01] Le Van An: Chào mọi người, hôm nay bàn hợp đồng ABC.',
        '[00:00:07] Tran Thanh Binh: Em đã gửi bản V2. Hết ạ.',
      ].join('\n'),
    )
    expect(r.lines).toBe(2)
    expect(r.speakers).toEqual(['Le Van An', 'Tran Thanh Binh'])
  })

  test('no voice tag gives "?", empty cues are skipped, CRLF is fine', () => {
    const r = vttToText('WEBVTT\r\n\r\n00:00:01.000 --> 00:00:02.000\r\nxin chào\r\n\r\n00:00:03.000 --> 00:00:04.000\r\n<v A></v>\r\n')
    expect(r.text).toBe('[00:00:01] ?: xin chào')
  })
})

describe('MeetingChats', () => {
  test('remembers the meeting id of a meeting chat only', () => {
    const m = new MeetingChats()
    m.note('19:meet', { channelData: { meeting: { id: 'MCMjMSM=' } } })
    m.note('19:group', { channelData: {} })
    expect(m.meetingOf('19:meet')).toBe('MCMjMSM=')
    expect(m.meetingOf('19:group')).toBeUndefined()
  })
})

describe('fetchMeetingTranscript', () => {
  const meeting = async () => ({
    details: { msGraphResourceId: 'GRAPH-MEETING', title: 'Họp hợp đồng ABC', scheduledStartTime: '2026-11-18T03:00:00Z' },
    organizer: { aadObjectId: 'org-id' },
  })
  const base = 'https://graph.microsoft.com/v1.0/users/org-id/onlineMeetings/GRAPH-MEETING/transcripts'

  function fake(routes: Record<string, () => Response>) {
    const seen: { url: string; auth: string }[] = []
    const fetchImpl = (async (url: string, init?: RequestInit) => {
      seen.push({ url, auth: String((init?.headers as any)?.Authorization ?? '') })
      const r = routes[url]
      return r ? r() : new Response('{}', { status: 404 })
    }) as any
    return { seen, fetchImpl }
  }

  test('lists, downloads in order, writes a 0600 text file and returns no transcript text', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'transcript-'))
    const f = fake({
      [base]: () =>
        Response.json({ value: [{ id: 't2', createdDateTime: '2026-11-18T03:40:00Z' }, { id: 't1', createdDateTime: '2026-11-18T03:00:00Z' }] }),
      [`${base}/t1/content?$format=text/vtt`]: () => new Response(VTT),
      [`${base}/t2/content?$format=text/vtt`]: () => new Response('WEBVTT\n\n00:00:01.000 --> 00:00:02.000\n<v Ha>Phần hai.</v>\n'),
    })
    const r = await fetchMeetingTranscript('MCMjMSM=', {
      getMeeting: meeting,
      graphToken: async () => 'APP',
      fetchImpl: f.fetchImpl,
      saveDir: dir,
      now: () => new Date('2026-11-18T05:00:00Z'),
    })
    expect(r.transcripts).toBe(2)
    expect(r.speakers.sort()).toEqual(['Ha', 'Le Van An', 'Tran Thanh Binh'])
    expect(JSON.stringify(r)).not.toContain('hợp đồng ABC.') // the result carries the path, not the words
    const text = readFileSync(r.path, 'utf8')
    expect(text.indexOf('Phần 1/2')).toBeLessThan(text.indexOf('Phần 2/2'))
    expect(text).toContain('# Họp hợp đồng ABC')
    expect(text).toContain('[00:00:01] Ha: Phần hai.')
    // Windows keeps no POSIX mode; the bot runs on Linux, where this holds.
    if (process.platform !== 'win32') expect(statSync(r.path).mode & 0o777).toBe(0o600)
    expect(f.seen.every(s => s.auth === 'Bearer APP')).toBe(true)
  })

  test('explains the usual failures without the response body', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'transcript-'))
    const run = (routes: Record<string, () => Response>, getMeeting: MeetingLookup = meeting) =>
      fetchMeetingTranscript('m', { getMeeting, graphToken: async () => 'APP', fetchImpl: fake(routes).fetchImpl, saveDir: dir })
    await expect(run({ [base]: () => Response.json({ value: [] }) })).rejects.toThrow(/no transcript yet/)
    await expect(run({ [base]: () => Response.json({ error: { code: 'Forbidden', message: 'secret detail' } }, { status: 403 }) })).rejects.toThrow(/no permission/)
    await expect(
      run({ [base]: () => Response.json({ error: { code: 'Forbidden', innerError: { code: 'GraphAccessToTranscriptsDisabled' } } }, { status: 403 }) }),
    ).rejects.toThrow(/turned off Graph access to transcripts/)
    await expect(run({}, async () => ({ details: {} }))).rejects.toThrow(/no Graph meeting id/)
    await expect(run({}, async () => { throw Object.assign(new Error('x'), { status: 403 }) })).rejects.toThrow(/could not read the meeting: no permission/)
  })
})
