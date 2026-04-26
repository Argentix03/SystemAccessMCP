# Security Boundaries

## What This Project Implements

The current `UserSession` provider uses normal Windows user-session APIs:

- `System.Drawing.Graphics.CopyFromScreen` for screenshots
- `user32.dll` cursor and input APIs for mouse and keyboard events

That means it works in the signed-in interactive desktop where the process is
running. It is suitable for development VMs, test automation, and AI-assisted
interactive workflows after a user session is available.

## What It Does Not Implement

This project does not include code or instructions to defeat or bypass:

- Windows secure desktop
- UAC secure prompts
- lock screen controls
- logon screen controls
- credential collection or injection

Those surfaces are intentionally protected by Windows session and desktop
isolation. Attempting to bypass them from inside the guest OS would weaken the
security boundary this software should respect.

## Supported Path For Full VM Console Access

For full coverage across boot, logon, lock screen, and UAC prompts, implement a
provider that talks to the virtualization host or remote console layer, for
example:

- Hyper-V VMConnect / enhanced session equivalent APIs
- VMware console or VNC-compatible VM console
- VirtualBox console APIs
- QEMU/SPICE/VNC
- cloud provider serial/console screenshot and input APIs where available

That model sends input to the VM console from outside the guest operating system.
It is the correct architecture for system-wide VM control because it does not
require bypassing Windows protections within the guest.

## Included Hyper-V Provider

The included `HyperVVmConnect` provider is the first implementation of that model
for Hyper-V. It runs on the Hyper-V host, opens or reuses `vmconnect.exe`, captures
the VMConnect window, and sends mouse/keyboard input to that console window.

This reaches guest logon and UAC screens because the interaction is with the
virtual machine console from the host, not with a protected Windows desktop from
inside the guest.

Current limits:

- the host desktop must be unlocked
- `vmconnect.exe` must be installed
- the Hyper-V PowerShell module is required for VM listing/start operations
- the user must have Hyper-V permissions on the host, typically through local
  `Hyper-V Administrators` group membership
- screenshots are captured from the visible VMConnect window

## Configuration Changes

No Windows configuration changes are required by the included provider.

For the Hyper-V provider, a host may require adding the user running MCP to the
local `Hyper-V Administrators` group:

```powershell
net localgroup "Hyper-V Administrators" "$env:USERNAME" /add
```

Log out and back in after changing group membership so the new access token
contains the group.

Do not enable test-signing, kernel debugging, or install input/display drivers for
this project unless you are implementing a separate, signed, explicitly scoped
driver-based provider and documenting that provider's threat model.
