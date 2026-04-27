---
name: system-access-mcp
description: Use when controlling or observing Windows desktops through SystemAccessMCP, especially when choosing between an in-guest GuestDesktop server and a host-side HostHyperV server for VM login screens, lock screens, Winlogon, UAC secure desktop prompts, VMConnect screenshots/input, cursor/window observation, or avoiding mouse misclicks.
---

# SystemAccessMCP Workflow

Use the right server for the surface:

- Use `GuestDesktop` for normal signed-in desktop work inside the Windows machine
  or VM.
- Use `HostHyperV` for VM login, lock screen, Winlogon, UAC secure desktop, and
  any case where the in-guest server cannot see or control the screen.

Shared tools such as `cursor_state`, `window_hover`, `window_from_point`,
`window_list`, and `screen_state` exist in both profiles. Interpret them relative
to the server location.

## Normal Desktop

1. Call `screen_state` or `screenshot`.
2. Use `window_hover`, `window_from_point`, or `window_list` to confirm target
   geometry.
3. Move with `mouse_move`.
4. Verify with `cursor_state` or `window_hover`.
5. Click/type.
6. Capture or observe again.

## Hyper-V Secure Surfaces

1. Call `hyperv_console_connect` if VMConnect is not already open.
2. Call `hyperv_console_screenshot` with `area: "client"` unless host chrome is
   intentionally needed.
3. Choose coordinates relative to the returned screenshot.
4. Call `hyperv_console_mouse_move`.
5. Call `hyperv_console_pointer_state` before clicking. Verify `insideClient`
   and `relativeToClient`.
6. Call `hyperv_console_mouse_click`.
7. Capture again.

Use `hyperv_console_ctrl_alt_delete` for guest Ctrl+Alt+Del at lock/login.

## Switch To HostHyperV When

- UAC prompt appears.
- Guest is locked or at login.
- Winlogon or secure desktop is active.
- In-guest screenshot/input stops working.
- The task requires interacting before user login.

## Avoid

- Do not assume `GuestDesktop` can see UAC, lock screen, or logon UI.
- Do not use VMConnect `window` coordinates when `client` coordinates are
  available.
- Do not chain blind clicks. Move, verify pointer state, click, then capture.

