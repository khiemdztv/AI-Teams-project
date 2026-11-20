/**
 * What a turn may use when nobody's message is being answered (fork addition,
 * 20/11): background tasks in a per-person session, and the subagents that
 * run them.
 *
 * A background task works while its person is away. It may read and prepare
 * (drafts, prepare_* plans, files in the export folder or the person's own
 * OneDrive), and the main conversation sends the result when the person says
 * so in a message of their own. Three places enforce it, as for scheduled
 * jobs (src/scheduled-guard.ts): the PreToolUse hook (hooks/scheduled-guard.ts,
 * every subagent call), the m365 server (no Teams turn in a per-person
 * session) and the msteams server (owner mode).
 *
 * Allowlists, not denylists: a tool added later is refused here until someone
 * decides it only reads or prepares.
 */

import { M365_READ_TOOLS, mcpName, scheduledToolAllowed, SCHEDULED_REFUSAL } from './scheduled-guard.js'

/**
 * m365 tools that change nothing anyone else can see, so a background task
 * may use them without the person's message. Each one, and why:
 */
export const M365_COMPOSE_TOOLS = new Set([
  // Save an attachment or a SharePoint/OneDrive file on this machine to read it: nothing changes in Microsoft 365.
  'mail_get_attachment',
  'files_get',
  // Put a file in the owner's own OneDrive folder; the link opens for them only, nobody else is told.
  'files_upload',
  // A draft in the owner's own mailbox: nothing leaves until mail_send_draft, which needs their message.
  'mail_create_draft',
  'mail_update_draft',
  // Step 1 of every two-step action: returns a plan (held in the m365 server's memory) and changes nothing.
  // The matching step 2 (calendar_create_event, calendar_update_event, teams_send_message, teams_add_member,
  // files_share, mail_set_auto_reply) is refused below.
  'calendar_prepare_event',
  'calendar_prepare_update',
  'teams_prepare_message',
  'teams_prepare_add_member',
  'files_prepare_share',
  'mail_prepare_auto_reply',
  // Deliberately absent: mail_create_folder and mail_create_rule (lasting changes to the mailbox the person
  // did not ask for in that moment), teams_remove_member and teams_rename_chat (one step, visible to others),
  // calendar_cancel_event, and every step 2 and send above.
])

/** Read or prepare in the person's own Microsoft 365. */
export function m365PrepareOnlyAllowed(tool: string): boolean {
  return M365_READ_TOOLS.has(tool) || M365_COMPOSE_TOOLS.has(tool)
}

export const OWNER_IDLE_REFUSAL =
  'refused: this sends, invites, shares or changes something other people see, so it runs only while answering the person\'s own message. ' +
  'Prepare it (a draft or a prepare_* step), tell them in your reply what is ready, and do it when they answer.'

/**
 * msteams tools a subagent may use: files it is given or makes, the person's
 * notes and meeting transcripts, and the developers' log. Never reply, edit,
 * react, notify or anything addressed to a conversation: the main
 * conversation sends what a subagent prepared.
 */
export const MSTEAMS_SUBAGENT_TOOLS = new Set([
  'download_attachment',
  'make_pdf',
  'make_video',
  'notes_list',
  'note_add',
  'log_internal',
  'meeting_transcript',
])

/**
 * Whether a subagent may call `toolName`. Claude Code's built-in tools (Bash,
 * Read, Write, WebSearch, Skill, …) and other MCP servers are not this
 * guard's business: only msteams and m365 can reach people.
 */
export function subagentToolAllowed(toolName: string): boolean {
  const mcp = mcpName(toolName)
  if (!mcp) return true
  if (mcp.server === 'm365') return m365PrepareOnlyAllowed(mcp.tool)
  if (mcp.server === 'msteams') return MSTEAMS_SUBAGENT_TOOLS.has(mcp.tool)
  return true
}

export const SUBAGENT_REFUSAL =
  'refused: a subagent only prepares (reads, drafts, prepare_* steps, files in the export folder); it cannot send, reply, post, invite or share. ' +
  'Finish your work and return the result: the main conversation sends it after the person agrees.'

export const BACKGROUND_REFUSAL =
  'refused: background tasks only run in a person\'s own session. In this shared session the mailbox changes with every message, ' +
  'so a background task could read the wrong person\'s. Do the work in this turn instead (without run_in_background).'

/** The tools that start a subagent: "Agent" today, "Task" in older Claude Code. */
const AGENT_TOOLS = new Set(['Agent', 'Task'])

export type GuardContext = {
  /** The turn lock says a scheduled job is running (read-only for everyone). */
  scheduled: boolean
  /** BOT_OWNER is set: this is one person's own session. */
  owner: boolean
}

/**
 * The PreToolUse decision for one hook input: a refusal reason, or null to
 * say nothing and let the usual permission rules decide. `input` is the
 * parsed hook JSON, or null when it could not be read.
 *
 * Order matters: the scheduled-turn guard runs first, exactly as before, and
 * fails closed (unreadable input is a refusal); then subagent calls
 * (`agent_id` present) are held to the subagent allowlist; then the main
 * thread may not start a background agent outside a person's own session.
 */
export function preToolUseDecision(input: unknown, ctx: GuardContext): string | null {
  const obj = input && typeof input === 'object' ? (input as Record<string, unknown>) : {}
  const toolName = typeof obj.tool_name === 'string' ? obj.tool_name : ''
  if (ctx.scheduled && (!toolName || !scheduledToolAllowed(toolName))) return SCHEDULED_REFUSAL
  if (!toolName) return null
  if (typeof obj.agent_id === 'string' && obj.agent_id) {
    return subagentToolAllowed(toolName) ? null : SUBAGENT_REFUSAL
  }
  if (AGENT_TOOLS.has(toolName) && !ctx.owner) {
    const toolInput = obj.tool_input && typeof obj.tool_input === 'object' ? (obj.tool_input as Record<string, unknown>) : {}
    if (toolInput.run_in_background === true) return BACKGROUND_REFUSAL
  }
  return null
}
