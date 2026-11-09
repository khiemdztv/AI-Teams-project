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

/** An Adaptive Card holding only the buttons; the reply's text sits above it. */
export function buildLinkCard(links: LinkButton[]) {
  return {
    contentType: ADAPTIVE_CARD_CONTENT_TYPE,
    content: {
      type: 'AdaptiveCard',
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      version: '1.4',
      body: [],
      actions: links.map(link => ({ type: 'Action.OpenUrl', title: link.title, url: link.url })),
    },
  }
}
