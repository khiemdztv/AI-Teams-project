import { describe, expect, test } from 'bun:test'
import { mkdtempSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import {
  ADAPTIVE_CARD_CONTENT_TYPE,
  MAX_LINK_BUTTONS,
  buildLinkCard,
  parseLinkButtons,
} from '../src/link-buttons.js'
import { deliverOutbound } from '../src/outbound.js'
import { PendingUploadStore } from '../src/pending-uploads.js'

describe('link buttons (fork, 09/11)', () => {
  test('nothing passed is no buttons', () => {
    expect(parseLinkButtons(undefined)).toEqual([])
    expect(parseLinkButtons(null)).toEqual([])
  })

  test('a link without a title is labelled with its site', () => {
    expect(parseLinkButtons([{ url: 'https://www.youtube.com/' }])).toEqual([
      { title: 'youtube.com', url: 'https://www.youtube.com/' },
    ])
    expect(parseLinkButtons([{ title: 'Mở YouTube', url: 'https://youtube.com' }])[0].title).toBe('Mở YouTube')
  })

  test('only full http(s) URLs become buttons', () => {
    for (const url of ['javascript:alert(1)', 'file:///etc/passwd', 'msteams://x', 'youtube.com', '']) {
      expect(() => parseLinkButtons([{ url }])).toThrow()
    }
    expect(() => parseLinkButtons([{ url: 'https://user:pw@example.com' }])).toThrow(/password/)
  })

  test('too many links is refused, not cut', () => {
    const links = Array.from({ length: MAX_LINK_BUTTONS + 1 }, (_, i) => ({ url: `https://e.com/${i}` }))
    expect(() => parseLinkButtons(links)).toThrow(/at most/)
    expect(() => parseLinkButtons('https://e.com')).toThrow(/array/)
  })

  test('the card holds one OpenUrl action per link', () => {
    const card = buildLinkCard([{ title: 'A', url: 'https://a.com/' }])
    expect(card.contentType).toBe(ADAPTIVE_CARD_CONTENT_TYPE)
    expect(card.content.actions).toEqual([{ type: 'Action.OpenUrl', title: 'A', url: 'https://a.com/' }])
  })

  const pending = new PendingUploadStore(mkdtempSync(join(tmpdir(), 'links-')))
  async function deliver(text: string) {
    const posted: Record<string, any>[] = []
    const result = await deliverOutbound({
      text,
      links: [{ title: 'Mở', url: 'https://youtube.com/' }],
      plan: [],
      conversationId: 'conv',
      pending,
      post: async activity => {
        posted.push(activity)
        return { id: `sent-${posted.length}` }
      },
    })
    return { posted, result }
  }

  test('the buttons ride on the last text message', async () => {
    const { posted, result } = await deliver('x'.repeat(30_000))
    expect(posted.length).toBeGreaterThan(1)
    expect(posted.slice(0, -1).every(a => !a.attachments)).toBe(true)
    expect(posted.at(-1)!.attachments[0].contentType).toBe(ADAPTIVE_CARD_CONTENT_TYPE)
    expect(result.sent).toBe(posted.length)
  })

  test('a card Teams accepts without an id still counts as sent', async () => {
    const result = await deliverOutbound({
      text: 'Mở YouTube',
      links: [{ title: 'Mở', url: 'https://youtube.com/' }],
      plan: [],
      conversationId: 'conv',
      pending,
      post: async () => ({}),
    })
    expect(result).toEqual({ sentIds: [], sent: 1, offered: [] })
  })

  test('with no text the buttons are a message of their own', async () => {
    const { posted } = await deliver('')
    expect(posted).toHaveLength(1)
    expect(posted[0].text).toBeUndefined()
    expect(posted[0].attachments[0].content.actions).toHaveLength(1)
  })
})
