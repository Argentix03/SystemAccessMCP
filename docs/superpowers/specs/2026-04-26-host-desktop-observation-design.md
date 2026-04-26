# Host Desktop Observation Design

## Goal

Expose host-side desktop observation APIs so an AI client can reason about where
the mouse is, which host window is focused, what host window/control is under the
cursor, and how VMConnect console coordinates map to host screen coordinates.

## Scope

This is host-side inspection only. It does not introspect guest secure desktop
controls inside a Hyper-V VM. For VM UAC and logon surfaces, the host sees
VMConnect pixels and host window geometry; semantic guest UI inspection would
require a separate in-guest channel and is out of scope.

## API Shape

Add general tools and REST endpoints:

- `cursor_state`: current host cursor position and virtual screen bounds.
- `window_foreground`: foreground host window metadata.
- `window_hover`: host window/control under the current cursor.
- `window_from_point`: host window/control at a supplied screen coordinate.
- `window_list`: visible top-level host windows with title, process, and rect.
- `screen_state`: combined cursor, foreground, hover, virtual screen, and
  optional visible windows snapshot.

Add a Hyper-V helper:

- `hyperv_console_pointer_state`: VMConnect console metadata plus current cursor
  position relative to the VMConnect window and client rectangles.

## Data Model

Window metadata includes:

- `windowHandle`
- `title`
- `className`
- `processId`
- `processName`
- `rect`
- `clientRect`
- `isForeground`
- `containsCursor`

Hit-test responses include both:

- `hitWindow`: exact child/control handle under the point
- `rootWindow`: owning top-level window

## Usage Loop

For VM UAC or other pixel-only surfaces:

1. Capture `hyperv_console_screenshot` with `area: "client"` when possible.
2. Move toward the target.
3. Call `hyperv_console_pointer_state`.
4. Confirm the cursor is inside the expected client area and relative position.
5. Click, then capture again.

For normal host desktop apps:

1. Call `screen_state` or `window_hover`.
2. Use returned geometry to adjust mouse movement.
3. Call `window_from_point` to verify a target point before clicking.

