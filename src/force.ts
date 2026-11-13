/**
 * /force: send without showing the draft first (fork addition, 05/11).
 *
 * Email, Teams messages under the sender's name and meetings go in two steps:
 * the bot shows what it prepared and sends only when the sender says so.
 * "/force" anywhere in a message makes that message itself the "send it":
 * Claude prepares and sends in the same turn. It applies to that one message
 * only, there is no on/off mode, so nobody forgets it is on. The word is
 * removed before Claude sees the text and travels as meta force="true", which
 * the server sets, so a message cannot fake it by typing the attribute.
 *
 * "/force" alone, after the bot showed a draft, means "send that one".
 */

// "\force" counts too, as "\useransw" does (src/view-as.ts).
const TRIGGER = /(^|\s)[/\\]force(?=\s|$)/gi
/** The trigger with the spaces after it, so removing it leaves no gap at a line start. */
const TRIGGER_WITH_SPACE = /(^|\s)[/\\]force(?:[ \t]+|(?=\s)|$)/gi

/** What a bare "/force" stands for once the word is gone. */
export const FORCE_ALONE_TEXT = 'Gửi luôn.'

/** The text with "/force" removed, or undefined when the message has none. */
export function stripForceTrigger(text: string): string | undefined {
  const raw = String(text ?? '')
  if (!raw.match(TRIGGER)) return undefined
  return raw.replace(TRIGGER_WITH_SPACE, '$1').trim() || FORCE_ALONE_TEXT
}
