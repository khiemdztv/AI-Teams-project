/**
 * Fork addition (09/11): "Open" buttons under a reply.
 *
 * The bot runs on its own machine, so it cannot open a browser on the
 * sender's. A Teams card button is the nearest thing: one click opens the page
 * in the sender's own browser. Teams never lets a bot open a page without that
 * click.
 *
 * Only http(s) URLs: a button is a link the sender is invited to click, so
 * nothing that runs script, reads local files or launches another app.
 */

export const ADAPTIVE_CARD_CONTENT_TYPE = 'application/vnd.microsoft.card.adaptive'

/** Teams desktop shows about six actions in a row before wrapping badly. */
export const MAX_LINK_BUTTONS = 5

const MAX_TITLE = 40

export type LinkButton = { title: string; url: string }

/**
 * Validate the `links` argument of reply. Throws with a message for Claude on
 * anything it should fix; returns [] when nothing was passed.
 */
export function parseLinkButtons(raw: unknown): LinkButton[] {
  if (raw === undefined || raw === null) return []
  if (!Array.isArray(raw)) throw new Error('links must be an array of {title, url}')
  if (raw.length > MAX_LINK_BUTTONS) {
    throw new Error(`at most ${MAX_LINK_BUTTONS} links per reply (got ${raw.length})`)
  }
  return raw.map((item, i) => {
    const urlText = String((item as { url?: unknown })?.url ?? '').trim()
    let url: URL
    try {
      url = new URL(urlText)
    } catch {
      throw new Error(`links[${i}].url is not a full URL: start it with https://`)
    }
    if (url.protocol !== 'https:' && url.protocol !== 'http:') {
      throw new Error(`links[${i}].url must be http or https, not ${url.protocol}`)
    }
    if (url.username || url.password) {
      throw new Error(`links[${i}].url must not carry a user name or password`)
    }
    const given = String((item as { title?: unknown })?.title ?? '').trim()
    const title = (given || url.hostname.replace(/^www\./, '')).slice(0, MAX_TITLE)
    return { title, url: url.href }
  })
}

/**
 * Whether `text` can go inside the card above the buttons. A card is a fixed
 * width, so buttons alone leave most of it empty (10/11); with the text in it
 * the reply is one bubble. A card's TextBlock renders only part of markdown
 * (bold, italic, lists, links), so a table, code, a heading or a quote stays
 * in its own message.
 */
export function fitsInCard(text: string, format: 'markdown' | 'plain' = 'markdown'): boolean {
  if (format === 'plain' || !text.trim()) return false
  return !/^\s*\|/m.test(text) && !text.includes('```') && !/^\s*#{1,6}\s/m.test(text) && !/^\s*>/m.test(text)
}

/**
 * An Adaptive Card with the buttons, and the reply's text above them when
 * given.
 *
 * The buttons sit in the body as an ActionSet rather than in the card's own
 * actions: with an empty body Teams still pads a body area above the action
 * row, which left a tall empty frame (10/11). "positive" is drawn as a filled
 * button in the theme's accent colour instead of a grey outline.
 */
export function buildLinkCard(links: LinkButton[], text?: string) {
  return {
    contentType: ADAPTIVE_CARD_CONTENT_TYPE,
    content: {
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.4',
      body: [
        ...(text ? [{ type: 'TextBlock', text, wrap: true }] : []),
        {
          type: 'ActionSet',
          spacing: text ? 'Medium' : 'None',
          actions: links.map(link => ({ type: 'Action.OpenUrl', title: link.title, url: link.url, style: 'positive' })),
        },
      ],
    },
  }
}
