# SystemAccessMCP Guidance For Claude Code

Use the project skill at `skills/system-access-mcp/SKILL.md` when controlling
or observing Windows desktops through SystemAccessMCP. If installing skills into
Claude Code, copy that folder to the appropriate Claude skills directory, or use
the project-local copy under `.claude/skills/system-access-mcp`.

Core rule:

- `GuestDesktop` is for normal signed-in desktop work inside the machine or VM.
- `HostHyperV` is for VM login, lock screen, Winlogon, UAC secure desktop, and
  any case where in-guest automation cannot see/control the screen.

For details, read `docs/agent-usage.md`.

