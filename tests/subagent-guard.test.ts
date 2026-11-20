/**
 * Background tasks and subagents (fork addition, 20/11): a subagent only
 * reads and prepares, a background agent runs only in a person's own session,
 * and the scheduled-turn guard still comes first.
 */

import { describe, expect, test, beforeEach } from 'bun:test'
import { mkdtempSync } from 'fs'
import { tmpdir } from 'os'
import { join } from 'path'
import {
  preToolUseDecision,
  subagentToolAllowed,
  m365PrepareOnlyAllowed,
  M365_COMPOSE_TOOLS,
  SUBAGENT_REFUSAL,
  BACKGROUND_REFUSAL,
} from '../src/subagent-guard.js'
import { SCHEDULED_REFUSAL } from '../src/scheduled-guard.js'
import { writeTurnLock } from '../src/turn-lock.js'

const ROOT = join(import.meta.dir, '..')
const OWNER = '11111111-1111-1111-1111-111111111111'
const m365 = (tool: string) => `mcp__plugin_msteams_m365__${tool}`
const msteams = (tool: string) => `mcp__plugin_msteams_msteams__${tool}`
const idle = { scheduled: false, owner: false }
const ownerIdle = { scheduled: false, owner: true }
const fromSubagent = (tool_name: string, tool_input: unknown = {}) => ({ tool_name, tool_input, agent_id: 'a1b2c3', agent_type: 'general-purpose' })

describe('what a subagent may use', () => {
  test('m365 reads and compose steps, and the msteams file, note and log tools', () => {
    for (const ok of [
      m365('mail_list_messages'),
      m365('mail_read_message'),
      m365('mail_get_attachment'),
      m365('files_get'),
      m365('files_upload'),
      m365('mail_create_draft'),
      m365('mail_update_draft'),
      m365('calendar_prepare_event'),
      m365('teams_prepare_message'),
      m365('files_prepare_share'),
      msteams('download_attachment'),
      msteams('make_pdf'),
      msteams('make_video'),
      msteams('notes_list'),
      msteams('note_add'),
      msteams('log_internal'),
      msteams('meeting_transcript'),
      'mcp__m365__mail_list_messages',
    ]) {
      expect(subagentToolAllowed(ok)).toBe(true)
    }
  })

  test('nothing that sends, replies, posts, invites, shares or changes chats', () => {
    for (const no of [
      m365('mail_send_draft'),
      m365('teams_send_message'),
      m365('teams_add_member'),
      m365('teams_remove_member'),
      m365('teams_rename_chat'),
      m365('files_share'),
      m365('calendar_create_event'),
      m365('calendar_update_event'),
      m365('calendar_cancel_event'),
      m365('mail_set_auto_reply'),
      m365('mail_create_rule'),
      m365('mail_create_folder'),
      m365('a_tool_added_later'),
      msteams('reply'),
      msteams('edit_message'),
      msteams('react'),
      msteams('notify_activity'),
      msteams('call_link'),
      msteams('note_done'),
      msteams('watch_outlook'),
      msteams('a_tool_added_later'),
    ]) {
      expect(subagentToolAllowed(no)).toBe(false)
    }
  })

  test('built-in tools and other MCP servers are not this guard\'s business', () => {
    for (const ok of ['Bash', 'Read', 'Write', 'Edit', 'Glob', 'Grep', 'WebSearch', 'WebFetch', 'Skill', 'mcp__claude_ai_Microsoft_365__search']) {
      expect(subagentToolAllowed(ok)).toBe(true)
    }
  })

  test('compose tools are prepare-only for the m365 server too, and none of them is a send step', () => {
    for (const tool of M365_COMPOSE_TOOLS) {
      expect(m365PrepareOnlyAllowed(tool)).toBe(true)
      expect(tool).not.toMatch(/send|^files_share$|^teams_add_member$|remove|rename|cancel|set_|create_event|update_event|rule|folder/)
    }
    expect(m365PrepareOnlyAllowed('mail_send_draft')).toBe(false)
  })
})

describe('the PreToolUse decision', () => {
  test('a subagent is refused a send and allowed a draft; the main thread is left alone', () => {
    expect(preToolUseDecision(fromSubagent(msteams('reply')), ownerIdle)).toBe(SUBAGENT_REFUSAL)
    expect(preToolUseDecision(fromSubagent(m365('mail_send_draft')), ownerIdle)).toBe(SUBAGENT_REFUSAL)
    expect(preToolUseDecision(fromSubagent(m365('mail_create_draft')), ownerIdle)).toBeNull()
    expect(preToolUseDecision(fromSubagent('Bash'), ownerIdle)).toBeNull()
    expect(preToolUseDecision({ tool_name: msteams('reply') }, ownerIdle)).toBeNull()
    expect(preToolUseDecision({ tool_name: m365('mail_send_draft') }, idle)).toBeNull()
  })

  test('a background agent only in a person\'s own session; a foreground one anywhere', () => {
    const bg = { tool_name: 'Agent', tool_input: { prompt: 'x', run_in_background: true } }
    expect(preToolUseDecision(bg, idle)).toBe(BACKGROUND_REFUSAL)
    expect(preToolUseDecision({ ...bg, tool_name: 'Task' }, idle)).toBe(BACKGROUND_REFUSAL)
    expect(preToolUseDecision(bg, ownerIdle)).toBeNull()
    expect(preToolUseDecision({ tool_name: 'Agent', tool_input: { prompt: 'x' } }, idle)).toBeNull()
    expect(preToolUseDecision({ tool_name: 'Agent', tool_input: { run_in_background: 'true' } }, idle)).toBeNull()
  })

  test('the scheduled guard comes first, for subagents too, and fails closed on bad input', () => {
    const scheduled = { scheduled: true, owner: true }
    expect(preToolUseDecision({ tool_name: 'Bash' }, scheduled)).toBe(SCHEDULED_REFUSAL)
    expect(preToolUseDecision(fromSubagent('Bash'), scheduled)).toBe(SCHEDULED_REFUSAL)
    expect(preToolUseDecision({ tool_name: m365('mail_list_messages') }, scheduled)).toBeNull()
    // Allowed by the scheduled list, but a subagent never replies.
    expect(preToolUseDecision(fromSubagent(msteams('reply')), scheduled)).toBe(SUBAGENT_REFUSAL)
    expect(preToolUseDecision(null, scheduled)).toBe(SCHEDULED_REFUSAL)
    expect(preToolUseDecision({}, scheduled)).toBe(SCHEDULED_REFUSAL)
  })

  test('outside a scheduled turn, input it cannot read says nothing', () => {
    expect(preToolUseDecision(null, idle)).toBeNull()
    expect(preToolUseDecision('garbage', idle)).toBeNull()
    expect(preToolUseDecision({ tool_name: 42, agent_id: 'x' }, idle)).toBeNull()
  })
})

describe('the hook process', () => {
  let dir: string
  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'subagent-guard-'))
  })

  const runHook = (stdin: string, owner?: string) => {
    const env: Record<string, string> = {}
    for (const [k, v] of Object.entries(process.env)) if (v !== undefined && k !== 'BOT_OWNER') env[k] = v
    env.MSTEAMS_STATE_DIR = dir
    if (owner) env.BOT_OWNER = owner
    return Bun.spawnSync(['bun', join(ROOT, 'hooks', 'scheduled-guard.ts')], { env, stdin: new TextEncoder().encode(stdin) })
  }
  const decisionOf = (out: ReturnType<typeof runHook>) => JSON.parse(out.stdout.toString()).hookSpecificOutput

  test('refuses a subagent\'s reply, with a reason, and exits 0', () => {
    const out = runHook(JSON.stringify({ hook_event_name: 'PreToolUse', ...fromSubagent(msteams('reply')) }), OWNER)
    expect(out.exitCode).toBe(0)
    expect(decisionOf(out)).toEqual({ hookEventName: 'PreToolUse', permissionDecision: 'deny', permissionDecisionReason: SUBAGENT_REFUSAL })
  })

  test('refuses a background agent only without BOT_OWNER', () => {
    const bg = JSON.stringify({ hook_event_name: 'PreToolUse', tool_name: 'Agent', tool_input: { prompt: 'x', run_in_background: true } })
    expect(decisionOf(runHook(bg)).permissionDecisionReason).toBe(BACKGROUND_REFUSAL)
    const owned = runHook(bg, OWNER)
    expect(owned.exitCode).toBe(0)
    expect(owned.stdout.toString()).toBe('')
  })

  test('says nothing for an allowed call or unreadable input; a scheduled lock still refuses', () => {
    for (const stdin of [JSON.stringify(fromSubagent(m365('mail_create_draft'))), 'not json', '']) {
      const out = runHook(stdin)
      expect(out.exitCode).toBe(0)
      expect(out.stdout.toString()).toBe('')
    }
    writeTurnLock(dir, { since: Date.now(), conversationId: 'a:1', senderId: OWNER, scheduled: 'morning' })
    expect(decisionOf(runHook('not json')).permissionDecisionReason).toBe(SCHEDULED_REFUSAL)
    expect(decisionOf(runHook(JSON.stringify(fromSubagent(m365('mail_create_draft'))))).permissionDecisionReason).toBe(SCHEDULED_REFUSAL)
  })
})
