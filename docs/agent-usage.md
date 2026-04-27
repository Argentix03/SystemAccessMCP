# Agent Usage Guide

This guide explains how agents should choose and use SystemAccessMCP servers.

## Server Roles

Use two MCP or HTTP server instances for reliable VM automation:

- `GuestDesktop`: run inside the Windows machine or VM. Use this for normal
  signed-in desktop automation: screenshots, mouse, keyboard, cursor state,
  focused/hovered windows, hit-testing, and visible window geometry.
- `HostHyperV`: run elevated on the Hyper-V host. Use this for Hyper-V VMConnect
  access, especially when the guest is at login, lock screen, Winlogon, UAC
  secure desktop, or any state the in-guest process cannot see or control.

Shared observation tools such as `cursor_state`, `window_hover`,
`window_from_point`, `window_list`, and `screen_state` exist in both profiles.
Their meaning depends on where the server runs: in the guest they describe the
guest desktop; on the host they describe the host desktop and VMConnect window.

## Decision Tree

1. If the normal signed-in guest desktop is visible and responsive, use
   `GuestDesktop`.
2. If the guest is locked, at logon, showing UAC, or in a secure desktop, use
   `HostHyperV`.
3. If in-guest screenshots or input stop working, switch to `HostHyperV`.
4. If a click misses in VMConnect, use `hyperv_console_pointer_state` before the
   next click.

## Normal Desktop Loop

1. Call `screenshot` or `screen_state`.
2. Use `window_hover`, `window_from_point`, or `window_list` to confirm target
   geometry.
3. Move with `mouse_move`.
4. Verify with `cursor_state` or `window_hover`.
5. Click or type.
6. Capture again before deciding the next action.

## Hyper-V / Secure Surface Loop

1. Call `hyperv_console_connect` if no VMConnect window is open.
2. Capture with `hyperv_console_screenshot`, usually with `area: "client"`.
3. Choose coordinates relative to that screenshot.
4. Move with `hyperv_console_mouse_move`.
5. Call `hyperv_console_pointer_state` and verify:
   - `insideClient` is `true`
   - `relativeToClient` matches the intended screenshot coordinate
6. Click with `hyperv_console_mouse_click`.
7. Capture again.

Use `hyperv_console_ctrl_alt_delete` for guest Ctrl+Alt+Del at the lock or login
screen. VMConnect maps host Ctrl+Alt+End to guest Ctrl+Alt+Del.

## Common Gotchas

- Do not expect an in-guest `GuestDesktop` server to see UAC secure desktop,
  lock screen, or logon UI.
- Prefer `area: "client"` for VMConnect screenshots and coordinates. The
  `window` area includes VMConnect chrome and toolbars.
- The host web server must listen on a reachable address and Windows Firewall
  must allow TCP 8765 if the guest web UI proxies Hyper-V calls to the host.
- After every input action, capture or observe again. Do not chain many blind
  clicks.
- If both servers expose similar tool names, inspect the MCP server name/profile
  or configured server alias before acting.

