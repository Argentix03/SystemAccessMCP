# Manual Hyper-V Tests

Run these tests on the Hyper-V host, not inside the guest VM. The MCP server must
run in an unlocked interactive host desktop because the first Hyper-V provider
captures and controls the visible `vmconnect.exe` console window.

Microsoft documents VMConnect as the tool for interacting with a Hyper-V guest
operating system and documents the command form `VMConnect.exe <ServerName>
<VMName> /edit`: https://learn.microsoft.com/en-us/windows-server/virtualization/hyper-v/virtual-machine-connection

## Host Prerequisites

- Hyper-V host or management workstation
- Hyper-V PowerShell module available (`Get-Command Get-VM`)
- VMConnect available (`Get-Command vmconnect.exe`)
- The user running MCP has permission to connect to the target VM. For a local
  Hyper-V host, add the user to `Hyper-V Administrators` from an elevated shell:

  ```powershell
  net localgroup "Hyper-V Administrators" "$env:USERNAME" /add
  ```

  Log out and back in after changing group membership.
- Host desktop is unlocked while screenshots and input are being tested

## Automated Smoke Tests

From the project root:

```powershell
.\tests\Invoke-SmokeTests.ps1
```

Expected result:

- syntax checks pass
- current desktop status and screenshot pass
- MCP tool listing includes Hyper-V tools
- web `/api/hyperv/status` passes
- Hyper-V VM listing passes on a host with the Hyper-V module and sufficient
  permissions, otherwise it is reported as skipped or returns an access error

## Web Manual Test

Start the web server on the Hyper-V host:

```powershell
.\Start-WebServer.ps1 -ListenAddress 0.0.0.0 -Port 8765 -Profile HostHyperV
```

If the guest VM cannot connect to the host on TCP 8765, allow the port through
Windows Firewall on the host from an elevated PowerShell:

```powershell
New-NetFirewallRule `
  -DisplayName "SystemAccessMCP Web 8765" `
  -Direction Inbound `
  -Action Allow `
  -Protocol TCP `
  -LocalPort 8765 `
  -Profile Any
```

From inside the guest VM, verify:

```powershell
Test-NetConnection HOST-IP-OR-NAME -Port 8765
Invoke-RestMethod http://HOST-IP-OR-NAME:8765/health
```

Open:

```text
http://HOST-IP-OR-NAME:8765/
```

To use the web UI from inside the guest VM while still routing Hyper-V actions
through the host, start a guest web server with a host API proxy:

```powershell
.\Start-WebServer.ps1 -Port 8765 -Profile GuestDesktop -HyperVApiBaseUrl http://HOST-IP-OR-NAME:8765
```

Then open `http://127.0.0.1:8765/` inside the guest. Desktop controls use the
guest session, and Hyper-V controls proxy to the host web server.

Test sequence:

1. In the Hyper-V panel, keep `Server` as `localhost` or enter the Hyper-V host
   name.
2. Click `List VMs`; verify the expected VM names appear.
3. Enter a VM name, click `Start`, then click `Connect`.
4. Set `Control Target` to `Hyper-V Console`.
5. Click `Refresh VM`; verify the screenshot shows the VMConnect window.
6. Query `/api/screen/state` or call `screen_state`; verify the response includes
   cursor, foreground window, and hover window metadata.
7. Move inside the VMConnect screenshot, then query
   `/api/hyperv/console/pointer-state` or call
   `hyperv_console_pointer_state`; verify `insideClient` is true and
   `relativeToClient` matches the expected screenshot coordinate.
8. Move/click inside the screenshot using the Mouse controls.
9. Type into a focused text box or login field using the Keyboard controls.
10. At a Windows guest lock/login screen, click `Ctrl+Alt+Del`; verify the guest
   receives the secure attention sequence through VMConnect.
11. Trigger a UAC prompt inside the guest and verify screenshot, click, and
   keyboard input still operate through the VMConnect console.

## REST Manual Test

```powershell
$base = "http://127.0.0.1:8765"
$vm = "Your VM Name"

Invoke-RestMethod "$base/api/hyperv/status"
Invoke-RestMethod "$base/api/hyperv/vms?server=localhost"

Invoke-RestMethod "$base/api/hyperv/console/connect" `
  -Method Post `
  -ContentType application/json `
  -Body (@{ server = "localhost"; vmName = $vm } | ConvertTo-Json)

Invoke-RestMethod "$base/api/hyperv/console/screenshot?server=localhost&vmName=$([uri]::EscapeDataString($vm))"

Invoke-RestMethod "$base/api/screen/state?includeWindows=true&windowLimit=25"

Invoke-RestMethod "$base/api/window/hover"

Invoke-RestMethod "$base/api/hyperv/console/pointer-state?server=localhost&vmName=$([uri]::EscapeDataString($vm))"

Invoke-RestMethod "$base/api/hyperv/console/mouse/click" `
  -Method Post `
  -ContentType application/json `
  -Body (@{ server = "localhost"; vmName = $vm; x = 120; y = 120; button = "left" } | ConvertTo-Json)

Invoke-RestMethod "$base/api/hyperv/console/keyboard/type" `
  -Method Post `
  -ContentType application/json `
  -Body (@{ server = "localhost"; vmName = $vm; text = "test input" } | ConvertTo-Json)

Invoke-RestMethod "$base/api/hyperv/console/ctrl-alt-delete" `
  -Method Post `
  -ContentType application/json `
  -Body (@{ server = "localhost"; vmName = $vm } | ConvertTo-Json)
```

## MCP Manual Test

Configure the MCP client to launch:

```text
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\Users\Tester\Desktop\SystemAccessMCP\Start-McpServer.ps1
```

Ask the AI client to run this sequence:

1. Call `hyperv_status`.
2. Call `hyperv_list_vms` with `{ "server": "localhost" }`.
3. Call `hyperv_console_connect` with the chosen `vmName`.
4. Call `hyperv_console_screenshot`.
5. Use the screenshot to choose coordinates.
6. Call `hyperv_console_mouse_move`, then
   `hyperv_console_pointer_state` to verify the cursor is inside the VMConnect
   client area and mapped to the intended relative coordinate.
7. Call `hyperv_console_mouse_click`.
8. Call `hyperv_console_keyboard_type`.
9. Call `hyperv_console_ctrl_alt_delete` at the guest lock/login screen.

For host desktop apps outside Hyper-V, use `cursor_state`, `window_hover`,
`window_from_point`, `window_foreground`, `window_list`, and `screen_state` to
inspect the current host-side window geometry before clicking.

## Known Limits

- The initial provider captures the visible VMConnect window with
  `CopyFromScreen`; the host desktop must be unlocked.
- If another window covers VMConnect at capture time, activate or reconnect the
  console and capture again.
- Coordinate input is relative to the screenshot area. The default area is the
  whole VMConnect window, including chrome/toolbars.
- A fully headless provider would need a deeper Hyper-V console/graphics channel
  integration. This first provider keeps the implementation simple and supported.
