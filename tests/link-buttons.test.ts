import { describe, expect, test } from 'bun:test'
import { mkdtempSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import {
  ADAPTIVE_CARD_CONTENT_TYPE,
  fitsInCard,
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

  test('the card holds one filled OpenUrl button per link, in its body', () => {
    const card = buildLinkCard([{ title: 'A', url: 'https://a.com/' }])
    expect(card.contentType).toBe(ADAPTIVE_CARD_CONTENT_TYPE)
    expect(card.content.body).toEqual([
      { type: 'ActionSet', spacing: 'None', actions: [{ type: 'Action.OpenUrl', title: 'A', url: 'https://a.com/', style: 'positive' }] },
    ])
    // Card-level actions under an empty body left a tall empty frame in Teams.
    expect((card.content as Record<string, unknown>).actions).toBeUndefined()
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

  test('the last part of the text goes inside the card, above the buttons', async () => {
    const { posted, result } = await deliver('x'.repeat(30_000))
    expect(posted.length).toBeGreaterThan(1)
    expect(posted.slice(0, -1).every(a => !a.attachments && a.text)).toBe(true)
    const last = posted.at(-1)!
    expect(last.text).toBeUndefined()
    expect(last.attachments[0].contentType).toBe(ADAPTIVE_CARD_CONTENT_TYPE)
    const [block, buttons] = last.attachments[0].content.body
    expect(block).toMatchObject({ type: 'TextBlock', wrap: true })
    expect(block.text.length).toBeGreaterThan(0)
    expect(buttons.type).toBe('ActionSet')
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

  test('text a card cannot render stays a message, with the buttons under it', async () => {
    for (const text of ['| a | b |\n|---|---|\n| 1 | 2 |', '```\ncode\n```', '# Tiêu đề', '> trích dẫn']) {
      expect(fitsInCard(text)).toBe(false)
    }
    expect(fitsInCard('**Đậm**, _nghiêng_\n- một\n- hai [link](https://e.com)')).toBe(true)
    expect(fitsInCard('x', 'plain')).toBe(false)
    const { posted } = await deliver('| a | b |\n|---|---|\n| 1 | 2 |')
    expect(posted).toHaveLength(1)
    expect(posted[0].text).toContain('| a |')
    expect(posted[0].attachments[0].content.body).toHaveLength(1)
  })

  test('with no text the buttons are a message of their own', async () => {
    const { posted } = await deliver('')
    expect(posted).toHaveLength(1)
    expect(posted[0].text).toBeUndefined()
    expect(posted[0].attachments[0].content.body[0].actions).toHaveLength(1)
  })
})
