# SystemAccessMCP Agent Instructions

When using this repository or its MCP servers, follow
`docs/agent-usage.md` and the workflow in `skills/system-access-mcp/SKILL.md`.

Use the right server profile:

- `GuestDesktop`: use inside the Windows machine or VM for normal signed-in
  desktop work.
- `HostHyperV`: use on the Hyper-V host for VM login, lock screen, Winlogon,
  UAC secure desktop, and cases where in-guest automation cannot see or control
  the screen.

For normal desktop interaction, observe before acting:

1. Call `screen_state` or `screenshot`.
2. Use `window_hover`, `window_from_point`, or `window_list` to confirm geometry.
3. Move with `mouse_move`.
4. Verify with `cursor_state` or `window_hover`.
5. Click/type.
6. Capture or observe again.

For Hyper-V secure surfaces:

1. Call `hyperv_console_connect` if needed.
2. Call `hyperv_console_screenshot` with `area: "client"`.
3. Move with `hyperv_console_mouse_move`.
4. Verify with `hyperv_console_pointer_state`.
5. Click with `hyperv_console_mouse_click`.
6. Capture again.

Do not assume `GuestDesktop` can see or control UAC, lock screen, logon, or
Winlogon surfaces. Switch to `HostHyperV` for those.

