# SystemAccessMCP

Simple MCP and HTTP API servers that expose an interactive
screen/mouse/keyboard interface for Windows machines.

The project can run in two complementary places:

- **Inside a Windows machine or VM** with the `GuestDesktop` profile, exposing
  the current interactive desktop, screenshots, mouse and keyboard input, cursor
  position, focused/hovered windows, window hit-testing, and visible window
  geometry.
- **On a Hyper-V host** with the `HostHyperV` profile, exposing VMConnect-based
  control of guest VMs from outside the guest OS. This reaches VM login screens,
  lock screens, Winlogon surfaces, and UAC secure desktop prompts because input
  and screenshots are applied to the VM console from the host side.

The same codebase also provides:

- a stdio MCP server for AI clients
- an MCP-over-HTTP server for already-running elevated processes
- a web API and small human testing UI
- network-capable host/guest operation with explicit `-ListenAddress` and
  `-HyperVApiBaseUrl` options

A common setup is to run `GuestDesktop` inside the VM for normal desktop
automation, and run `HostHyperV` on the Hyper-V host for cases the in-guest
process cannot see or control, such as logon, lock screen, and UAC prompts. The
guest web UI can proxy Hyper-V actions to the host web server, so an agent or
human operating from inside the VM can still interact with those secure surfaces.

## Requirements

- Windows PowerShell 5.1
- Windows desktop session

No system configuration changes are required for the included `UserSession`
provider.

For Hyper-V host-side VM console access:

- Hyper-V PowerShell module (`Get-VM`)
- VMConnect (`vmconnect.exe`)
- membership in the local `Hyper-V Administrators` group, or equivalent
  administrative rights
- an unlocked interactive host desktop for screenshot/input capture

To grant the current user Hyper-V console permissions on the host, run an
elevated PowerShell or Command Prompt:

```powershell
net localgroup "Hyper-V Administrators" "$env:USERNAME" /add
```

Log out and back in after changing group membership so the user's access token
includes the new group.

## Start the Web Server

```powershell
.\Start-WebServer.ps1 -Port 8765
```

The web server supports the same operational profiles as the MCP server:

```powershell
# Inside the guest VM: normal guest desktop automation.
.\Start-WebServer.ps1 -Port 8765 -Profile GuestDesktop

# On the Hyper-V host: VMConnect access for login, UAC, and secure desktop.
.\Start-WebServer.ps1 -ListenAddress 0.0.0.0 -Port 8765 -Profile HostHyperV

# Inside the guest VM, with Hyper-V actions proxied to the host web server.
.\Start-WebServer.ps1 -Port 8765 -Profile GuestDesktop -HyperVApiBaseUrl http://HOST-IP-OR-NAME:8765
```

Use `-ListenAddress 0.0.0.0` only on a trusted/private network or firewall rule.
The web API can capture screenshots and send input.

If a guest VM or another machine must connect to the host web server, allow the
port through Windows Firewall on the host from an elevated PowerShell:

```powershell
New-NetFirewallRule `
  -DisplayName "SystemAccessMCP Web 8765" `
  -Direction Inbound `
  -Action Allow `
  -Protocol TCP `
  -LocalPort 8765 `
  -Profile Any
```

From the guest VM, verify connectivity before starting the proxied guest web
server:

```powershell
Test-NetConnection HOST-IP-OR-NAME -Port 8765
Invoke-RestMethod http://HOST-IP-OR-NAME:8765/health
```

Stop a foreground server with `Ctrl+C`. If you launched it as a hidden/background
PowerShell process, stop that process explicitly.

Open:

```text
http://127.0.0.1:8765/
```

The web UI includes desktop controls, Hyper-V controls, and an Observation panel
for cursor, foreground window, hover window, point hit-testing, visible window
listing, and combined screen state.

Available endpoints:

- `GET /health`
- `GET /screenshot`
- `GET /api/screenshot`
- `GET /api/cursor/state`
- `GET /api/window/foreground`
- `GET /api/window/hover`
- `GET /api/window/from-point?x=100&y=100`
- `GET /api/window/list?limit=100&includeUntitled=false`
- `GET /api/screen/state?includeWindows=false&windowLimit=100`
- `POST /api/mouse/move`
- `POST /api/mouse/click`
- `POST /api/keyboard/type`
- `POST /api/keyboard/key`
- `GET /api/hyperv/status`
- `GET /api/hyperv/vms?server=localhost`
- `POST /api/hyperv/vm/start`
- `POST /api/hyperv/console/connect`
- `GET /api/hyperv/console/pointer-state`
- `GET /api/hyperv/console/screenshot`
- `GET /hyperv/screenshot`
- `POST /api/hyperv/console/mouse/move`
- `POST /api/hyperv/console/mouse/click`
- `POST /api/hyperv/console/keyboard/type`
- `POST /api/hyperv/console/keyboard/key`
- `POST /api/hyperv/console/ctrl-alt-delete`

Profile behavior:

- `GuestDesktop` exposes desktop and observation endpoints. Hyper-V endpoints
  are available only when `-HyperVApiBaseUrl` is set, and are proxied to that
  host-side web server.
- `HostHyperV` exposes Hyper-V and observation endpoints. Generic desktop
  screenshot/mouse/keyboard endpoints are hidden.
- `All` exposes every endpoint for development and local testing.

Example:

```powershell
Invoke-RestMethod http://127.0.0.1:8765/health

Invoke-RestMethod http://127.0.0.1:8765/api/mouse/click `
  -Method Post `
  -ContentType application/json `
  -Body '{"x":100,"y":100,"button":"left"}'
```

## Run as an MCP Server

Configure an MCP client to launch:

```text
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Users\Tester\Desktop\SystemAccessMCP\Start-McpServer.ps1
```

Use `-Profile` to choose which tool set this instance exposes:

```powershell
.\Start-McpServer.ps1 -Profile GuestDesktop
.\Start-McpServer.ps1 -Profile HostHyperV
.\Start-McpServer.ps1 -Profile All
```

Profiles:

- `GuestDesktop`: current desktop input/screenshot plus shared cursor/window
  observation. Run this elevated inside the guest VM for normal guest desktop
  automation.
- `HostHyperV`: Hyper-V/VMConnect plus shared cursor/window observation. Run
  this elevated on the Hyper-V host for guest login, UAC, and secure desktop
  surfaces through VMConnect.
- `All`: every tool, useful for development and local manual testing.

Tools exposed by the default `All` profile:

- `system_status`
- `screenshot`
- `mouse_move`
- `mouse_click`
- `keyboard_type`
- `keyboard_key`
- `cursor_state`
- `window_foreground`
- `window_hover`
- `window_from_point`
- `window_list`
- `screen_state`
- `hyperv_status`
- `hyperv_list_vms`
- `hyperv_start_vm`
- `hyperv_console_windows`
- `hyperv_console_connect`
- `hyperv_console_pointer_state`
- `hyperv_console_screenshot`
- `hyperv_console_mouse_move`
- `hyperv_console_mouse_click`
- `hyperv_console_keyboard_type`
- `hyperv_console_keyboard_key`
- `hyperv_console_ctrl_alt_delete`

## Run as an Elevated HTTP MCP Server

Use this mode when the MCP client should not launch the server itself. For
example, Codex can run normally while you start the MCP server in an elevated
PowerShell window so Hyper-V and VMConnect actions run with the elevated token.

Start the HTTP MCP server from an elevated PowerShell:

```powershell
.\Start-McpHttpServer.ps1 -Port 8766 -Profile HostHyperV
```

For a server that another machine must reach, bind deliberately to an address on
that network:

```powershell
.\Start-McpHttpServer.ps1 -ListenAddress 0.0.0.0 -Port 8766 -Profile HostHyperV
```

Only expose this on a trusted/private network or firewall rule. These tools can
send input and capture screenshots.

Configure Codex to connect to already-running guest and host servers:

```toml
[mcp_servers."vm-desktop"]
url = "http://VM-IP-OR-NAME:8766/mcp"

[mcp_servers."hyperv-host"]
url = "http://HOST-IP-OR-NAME:8766/mcp"
```

Restart Codex after changing the MCP configuration. Keep the elevated server
windows open while Codex uses the tools.

## Tests

```powershell
.\tests\Invoke-SmokeTests.ps1
```

For host-side VM console validation, follow
[docs/manual-hyperv-tests.md](docs/manual-hyperv-tests.md).

## Security Boundary

The `UserSession` provider can only interact with the current user desktop. The
`HyperVVmConnect` provider interacts with guest secure/logon surfaces through the
host-side VM console. See
[docs/security-boundaries.md](docs/security-boundaries.md) for the reasoning and
remaining limits.
