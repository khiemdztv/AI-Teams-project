/**
 * What a scheduled turn may use (fork addition, 05/11). A job runs with nobody
 * watching and reads email from outside the firm, so it is read-only: it reads
 * the person's own mail, calendar, Teams and notes, and answers in their own
 * chat. Anything else is refused, by name, in three places that do not trust
 * each other: the PreToolUse hook (hooks/scheduled-guard.ts, every tool), the
 * m365 server and the msteams server.
 *
 * An allowlist, not a denylist: a tool added later is refused here until
 * someone decides it only reads.
 */

/** m365 tools that only read the person's own Microsoft 365. */
export const M365_READ_TOOLS = new Set([
  'm365_whoami',
  'mail_list_messages',
  'mail_read_message',
  'mail_list_drafts',
  'mail_list_rules',
  'calendar_list_events',
  'calendar_find_time',
  'files_search',
  'teams_list_chats',
  'teams_list_channels',
  'teams_read_chat',
  'teams_read_channel',
])

/** msteams tools a job needs: its one answer, its notes, and the developers' log. */
export const MSTEAMS_SCHEDULED_TOOLS = new Set(['reply', 'notes_list', 'log_internal'])

/** Claude Code's own tools that change nothing outside the session. */
const BUILTIN_TOOLS = new Set(['Read', 'ToolSearch', 'TodoWrite'])

/** "mcp__plugin_msteams_m365__mail_list_messages" -> { server: "m365", tool: "mail_list_messages" }. */
function mcpName(toolName: string): { server: string; tool: string } | null {
  const parts = toolName.split('__')
  if (parts.length !== 3 || parts[0] !== 'mcp') return null
  const server = parts[1]!.replace(/^plugin_msteams_/, '')
  return { server, tool: parts[2]! }
}

export function scheduledToolAllowed(toolName: string): boolean {
  if (BUILTIN_TOOLS.has(toolName)) return true
  const mcp = mcpName(toolName)
  if (!mcp) return false
  if (mcp.server === 'm365') return M365_READ_TOOLS.has(mcp.tool)
  if (mcp.server === 'msteams') return MSTEAMS_SCHEDULED_TOOLS.has(mcp.tool)
  return false
}

export const SCHEDULED_REFUSAL = 'refused: this is a scheduled job, which only reads. Tell the person in your one reply what they may want to do; they can ask for it in a message of their own.'
