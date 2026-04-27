# SystemAccessMCP Guidance For Gemini / Antigravity

Use `docs/agent-usage.md` as the canonical guide for operating this project.
The same workflow is packaged as a reusable skill in
`skills/system-access-mcp/SKILL.md`.

Choose tools by surface:

- Use `GuestDesktop` inside the Windows machine or VM for normal signed-in
  desktop interaction.
- Use `HostHyperV` on the Hyper-V host for VM login, lock screen, Winlogon,
  UAC secure desktop, and other surfaces unavailable to the in-guest process.

Before clicking, observe the screen and window state. For VMConnect, prefer
`hyperv_console_screenshot` with `area: "client"`, then move, call
`hyperv_console_pointer_state`, and only then click.

After every mouse or keyboard action, capture or observe again before deciding
the next action.

