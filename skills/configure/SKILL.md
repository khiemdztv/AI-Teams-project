---
name: configure
description: Report the Microsoft Teams channel's status — which credentials are set and who is allowed — without changing anything. Use when the user asks whether Teams is configured, who can reach the bot, or wants to check channel status. Never accept or save a secret here.
user-invocable: true
allowed-tools:
  - Read
  - Bash(echo *)
  - Bash(ls *)
---

# /msteams:configure — Teams channel status (read only)

Fork policy: this skill only reports. Credentials and the allowlist
are edited by the operator directly on the host, in a text editor, never through
a chat. A secret pasted into a conversation stays in its transcript.

## Resolve the state dir

```bash
echo "${MSTEAMS_STATE_DIR:-$HOME/.claude/channels/msteams}"
```

On Windows the default is `%USERPROFILE%\.claude\channels\msteams`. Report the
resolved path on its own line.

## Report

1. **Credentials** — read `<STATE_DIR>/.env` and show set / not set for
   `MSTEAMS_APP_ID`, `MSTEAMS_APP_PASSWORD`, `MSTEAMS_TENANT_ID` and
   `MSTEAMS_SENDABLE_ROOT`. Never echo a value; mask the app id beyond its first
   segment.
2. **Access** — read `<STATE_DIR>/access.json`. This fork always runs in static
   allowlist mode and accepts 1:1 chats only: show the `allowFrom` entries and
   say that groups and pairing are disabled in code.
3. **What next** — if anything is missing, tell the operator which file to edit
   and that the bot reads both files only at startup, so a restart is needed.

If the user pastes something that looks like a secret, do not save or repeat
it. Tell them to put it in `.env` themselves and to rotate it in Azure, since it
is now in this transcript.
