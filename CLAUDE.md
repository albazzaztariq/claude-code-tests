# Project: Claude Code Tests

## Session Continuity

This project uses session summaries to maintain context across Claude CLI sessions.

### How it works:
1. Before ending a session, run `/save-session` to save a summary
2. The summary is written to `SESSION_HISTORY.md` in the project root
3. On new sessions, Claude reads this file to understand previous work

### Session History Location:
- **File:** `SESSION_HISTORY.md` (project root)
- **Command:** `/save-session`

## Important Files
- `.claude/settings.local.json` - Local Claude settings with BurntToast notification hooks
- `.claude/commands/save-session.md` - Session save command

## Notes
- BurntToast notifications configured for Stop, idle_prompt, and permission_prompt hooks
- No built-in command exists to refresh settings mid-session; restart Claude CLI to apply changes
