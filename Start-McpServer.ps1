param(
    [ValidateSet("All", "GuestDesktop", "HostHyperV")]
    [string] $Profile = "All"
)

Set-StrictMode -Version 2.0

. "$PSScriptRoot\src\SystemAccess.Core.ps1"
. "$PSScriptRoot\src\SystemAccess.HyperV.ps1"

$ProtocolVersion = "2024-11-05"
$script:McpProfile = $Profile
$script:SharedObservationToolNames = @(
    "system_status",
    "cursor_state",
    "window_foreground",
    "window_hover",
    "window_from_point",
    "window_list",
    "screen_state"
)
$script:GuestDesktopToolNames = @(
    "screenshot",
    "mouse_move",
    "mouse_click",
    "keyboard_type",
    "keyboard_key"
) + $script:SharedObservationToolNames
$script:HostHyperVToolNames = @(
    "hyperv_status",
    "hyperv_list_vms",
    "hyperv_start_vm",
    "hyperv_console_windows",
    "hyperv_console_connect",
    "hyperv_console_pointer_state",
    "hyperv_console_screenshot",
    "hyperv_console_mouse_move",
    "hyperv_console_mouse_click",
    "hyperv_console_keyboard_type",
    "hyperv_console_keyboard_key",
    "hyperv_console_ctrl_alt_delete"
) + $script:SharedObservationToolNames

function Get-McpAllowedToolNames {
    switch ($script:McpProfile) {
        "GuestDesktop" {
            return @($script:GuestDesktopToolNames | Select-Object -Unique)
        }
        "HostHyperV" {
            return @($script:HostHyperVToolNames | Select-Object -Unique)
        }
        default {
            return @($script:GuestDesktopToolNames + $script:HostHyperVToolNames | Select-Object -Unique)
        }
    }
}

function Get-McpProfileDescription {
    switch ($script:McpProfile) {
        "GuestDesktop" {
            return "GuestDesktop controls and observes the current interactive Windows desktop from inside the machine or VM. Use it for normal signed-in desktop work. It cannot see UAC secure desktop, lock screen, logon, or Winlogon surfaces from inside the guest; use HostHyperV for those."
        }
        "HostHyperV" {
            return "HostHyperV controls Hyper-V VMConnect windows from the host. Use it for VM login screens, lock screens, Winlogon, UAC secure desktop prompts, and cases where in-guest desktop automation cannot see or interact with the screen."
        }
        default {
            return "All exposes both GuestDesktop and HostHyperV tools. Use GuestDesktop-style tools for normal desktop work and Hyper-V VMConnect tools for guest login, lock screen, Winlogon, and UAC secure desktop surfaces."
        }
    }
}

function Write-McpMessage {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Message
    )

    $json = $Message | ConvertTo-Json -Depth 20 -Compress
    [Console]::Out.WriteLine($json)
    [Console]::Out.Flush()
}

function New-McpResponse {
    param(
        [Parameter(Mandatory = $true)]
        $Id,

        [Parameter(Mandatory = $true)]
        [object] $Result
    )

    [pscustomobject]@{
        jsonrpc = "2.0"
        id = $Id
        result = $Result
    }
}

function New-McpErrorResponse {
    param(
        $Id,
        [int] $Code = -32603,
        [string] $Message = "Internal error"
    )

    [pscustomobject]@{
        jsonrpc = "2.0"
        id = $Id
        error = [pscustomobject]@{
            code = $Code
            message = $Message
        }
    }
}

function New-TextContent {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string] $Text
    )

    [pscustomobject]@{
        type = "text"
        text = $Text
    }
}

function New-JsonTextContent {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Value
    )

    New-TextContent -Text (ConvertTo-SystemAccessJson -Value $Value)
}

function Get-McpArgument {
    param(
        [object] $Arguments,
        [Parameter(Mandatory = $true)]
        [string] $Name,
        [object] $Default = $null
    )

    if ($null -ne $Arguments) {
        $matches = $Arguments.PSObject.Properties.Match($Name)
        if ($matches.Count -gt 0) {
            return $matches[0].Value
        }
    }

    return $Default
}

function Assert-McpArgument {
    param(
        [object] $Arguments,
        [Parameter(Mandatory = $true)]
        [string] $Name
    )

    $value = Get-McpArgument -Arguments $Arguments -Name $Name
    if ($null -eq $value) {
        throw "$Name is required"
    }

    return $value
}

function Get-McpTools {
    $tools = @(
        [pscustomobject]@{
            name = "system_status"
            description = "Return provider scope, current user, virtual screen metadata, and the current desktop-session boundary."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{}
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "screenshot"
            description = "Capture the current interactive desktop as PNG image content. In GuestDesktop this is the in-guest signed-in desktop and does not cover secure desktop, lock screen, logon, or UAC prompts."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{}
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "mouse_move"
            description = "Move the mouse cursor on the current user desktop."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{
                    x = [pscustomobject]@{ type = "integer"; description = "Absolute X coordinate, or delta when relative is true." }
                    y = [pscustomobject]@{ type = "integer"; description = "Absolute Y coordinate, or delta when relative is true." }
                    relative = [pscustomobject]@{ type = "boolean"; description = "Use relative movement instead of absolute coordinates."; default = $false }
                }
                required = @("x", "y")
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "mouse_click"
            description = "Click a mouse button, optionally moving to an absolute coordinate first."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{
                    x = [pscustomobject]@{ type = "integer"; description = "Optional absolute X coordinate." }
                    y = [pscustomobject]@{ type = "integer"; description = "Optional absolute Y coordinate." }
                    button = [pscustomobject]@{ type = "string"; enum = @("left", "right", "middle"); default = "left" }
                    clicks = [pscustomobject]@{ type = "integer"; minimum = 1; maximum = 10; default = 1 }
                }
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "keyboard_type"
            description = "Type Unicode text into the focused control on the current user desktop."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{
                    text = [pscustomobject]@{ type = "string" }
                }
                required = @("text")
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "keyboard_key"
            description = "Send a virtual-key press/down/up event to the current user desktop."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{
                    virtualKey = [pscustomobject]@{ type = "integer"; minimum = 1; maximum = 255; description = "Windows virtual-key code." }
                    action = [pscustomobject]@{ type = "string"; enum = @("press", "down", "up"); default = "press" }
                }
                required = @("virtualKey")
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "cursor_state"
            description = "Return current mouse cursor position and virtual screen metadata for the machine running this MCP server."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{}
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "window_foreground"
            description = "Return metadata for the current foreground window on the machine running this MCP server."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{}
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "window_hover"
            description = "Return window/control metadata for what the mouse is currently hovering over on the machine running this MCP server."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{}
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "window_from_point"
            description = "Return window/control metadata at a screen coordinate on the machine running this MCP server."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{
                    x = [pscustomobject]@{ type = "integer"; description = "Host screen X coordinate. Defaults to current cursor X." }
                    y = [pscustomobject]@{ type = "integer"; description = "Host screen Y coordinate. Defaults to current cursor Y." }
                }
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "window_list"
            description = "List visible top-level windows with titles, processes, and rectangles on the machine running this MCP server."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{
                    limit = [pscustomobject]@{ type = "integer"; minimum = 1; maximum = 500; default = 100 }
                    includeUntitled = [pscustomobject]@{ type = "boolean"; default = $false }
                }
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "screen_state"
            description = "Return a combined desktop observation snapshot with cursor, foreground window, hover window, virtual screen, and optional visible windows."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{
                    includeWindows = [pscustomobject]@{ type = "boolean"; default = $false }
                    windowLimit = [pscustomobject]@{ type = "integer"; minimum = 1; maximum = 500; default = 100 }
                }
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "hyperv_status"
            description = "Return Hyper-V/VMConnect availability and open console window metadata from the host. Use HostHyperV for guest login, lock screen, Winlogon, and UAC secure desktop surfaces."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{}
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "hyperv_list_vms"
            description = "List Hyper-V virtual machines visible from the host PowerShell module."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{
                    server = [pscustomobject]@{ type = "string"; default = "localhost" }
                }
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "hyperv_start_vm"
            description = "Start a Hyper-V virtual machine by name."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{
                    vmName = [pscustomobject]@{ type = "string" }
                    server = [pscustomobject]@{ type = "string"; default = "localhost" }
                }
                required = @("vmName")
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "hyperv_console_windows"
            description = "List open VMConnect console windows."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{
                    vmName = [pscustomobject]@{ type = "string" }
                    server = [pscustomobject]@{ type = "string" }
                }
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "hyperv_console_connect"
            description = "Open or reuse a VMConnect console window for a Hyper-V VM."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{
                    vmName = [pscustomobject]@{ type = "string" }
                    server = [pscustomobject]@{ type = "string"; default = "localhost" }
                    timeoutMs = [pscustomobject]@{ type = "integer"; default = 15000 }
                    forceNew = [pscustomobject]@{ type = "boolean"; default = $false }
                }
                required = @("vmName")
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "hyperv_console_pointer_state"
            description = "Return current host cursor position mapped relative to a VMConnect window and client area. Use after moving and before clicking to verify insideClient and relativeToClient."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{
                    vmName = [pscustomobject]@{ type = "string" }
                    server = [pscustomobject]@{ type = "string"; default = "localhost" }
                    processId = [pscustomobject]@{ type = "integer"; default = 0 }
                }
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "hyperv_console_screenshot"
            description = "Capture an open VMConnect console as PNG image content from the host. Use area='client' for guest UI coordinates, including login, lock screen, Winlogon, and UAC prompts."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{
                    vmName = [pscustomobject]@{ type = "string" }
                    server = [pscustomobject]@{ type = "string"; default = "localhost" }
                    processId = [pscustomobject]@{ type = "integer"; default = 0 }
                    area = [pscustomobject]@{ type = "string"; enum = @("window", "client"); default = "window" }
                }
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "hyperv_console_mouse_move"
            description = "Move the host mouse to a coordinate inside a VMConnect console screenshot. Prefer coordinates from a client-area screenshot."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{
                    x = [pscustomobject]@{ type = "integer" }
                    y = [pscustomobject]@{ type = "integer" }
                    vmName = [pscustomobject]@{ type = "string" }
                    server = [pscustomobject]@{ type = "string"; default = "localhost" }
                    processId = [pscustomobject]@{ type = "integer"; default = 0 }
                    area = [pscustomobject]@{ type = "string"; enum = @("window", "client"); default = "window" }
                }
                required = @("x", "y")
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "hyperv_console_mouse_click"
            description = "Click inside a VMConnect console screenshot. Prefer move plus hyperv_console_pointer_state verification before clicking."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{
                    x = [pscustomobject]@{ type = "integer" }
                    y = [pscustomobject]@{ type = "integer" }
                    button = [pscustomobject]@{ type = "string"; enum = @("left", "right", "middle"); default = "left" }
                    clicks = [pscustomobject]@{ type = "integer"; minimum = 1; maximum = 10; default = 1 }
                    vmName = [pscustomobject]@{ type = "string" }
                    server = [pscustomobject]@{ type = "string"; default = "localhost" }
                    processId = [pscustomobject]@{ type = "integer"; default = 0 }
                    area = [pscustomobject]@{ type = "string"; enum = @("window", "client"); default = "window" }
                }
                required = @("x", "y")
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "hyperv_console_keyboard_type"
            description = "Type Unicode text into the active VMConnect console, including guest login or UAC surfaces reachable through the host console."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{
                    text = [pscustomobject]@{ type = "string" }
                    vmName = [pscustomobject]@{ type = "string" }
                    server = [pscustomobject]@{ type = "string"; default = "localhost" }
                    processId = [pscustomobject]@{ type = "integer"; default = 0 }
                }
                required = @("text")
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "hyperv_console_keyboard_key"
            description = "Send a virtual-key press/down/up event to the active VMConnect console, including guest secure/logon surfaces reachable through VMConnect."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{
                    virtualKey = [pscustomobject]@{ type = "integer"; minimum = 1; maximum = 255 }
                    action = [pscustomobject]@{ type = "string"; enum = @("press", "down", "up"); default = "press" }
                    vmName = [pscustomobject]@{ type = "string" }
                    server = [pscustomobject]@{ type = "string"; default = "localhost" }
                    processId = [pscustomobject]@{ type = "integer"; default = 0 }
                }
                required = @("virtualKey")
                additionalProperties = $false
            }
        },
        [pscustomobject]@{
            name = "hyperv_console_ctrl_alt_delete"
            description = "Send Ctrl+Alt+End to VMConnect, which VMConnect maps to guest Ctrl+Alt+Del."
            inputSchema = [pscustomobject]@{
                type = "object"
                properties = [pscustomobject]@{
                    vmName = [pscustomobject]@{ type = "string" }
                    server = [pscustomobject]@{ type = "string"; default = "localhost" }
                    processId = [pscustomobject]@{ type = "integer"; default = 0 }
                }
                additionalProperties = $false
            }
        }
    )

    $allowed = @(Get-McpAllowedToolNames)
    @($tools | Where-Object { $allowed -contains $_.name })
}

function Invoke-McpTool {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Name,

        [object] $Arguments
    )

    if ($null -eq $Arguments) {
        $Arguments = [pscustomobject]@{}
    }

    $allowed = @(Get-McpAllowedToolNames)
    if ($allowed -notcontains $Name) {
        throw "Tool '$Name' is not available in MCP profile '$script:McpProfile'."
    }

    switch ($Name) {
        "system_status" {
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value (Get-SystemAccessStatus))
            }
        }
        "screenshot" {
            $shot = Get-SystemAccessScreenshot
            return [pscustomobject]@{
                content = @(
                    [pscustomobject]@{
                        type = "image"
                        data = $shot.data
                        mimeType = $shot.mimeType
                    },
                    (New-JsonTextContent -Value ([pscustomobject]@{
                        width = $shot.width
                        height = $shot.height
                        left = $shot.left
                        top = $shot.top
                        timestamp = $shot.timestamp
                    }))
                )
            }
        }
        "mouse_move" {
            $x = Assert-McpArgument -Arguments $Arguments -Name "x"
            $y = Assert-McpArgument -Arguments $Arguments -Name "y"
            $relative = Get-McpArgument -Arguments $Arguments -Name "relative" -Default $false
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value (Invoke-SystemAccessMouseMove -X ([int]$x) -Y ([int]$y) -Relative ([bool]$relative)))
            }
        }
        "mouse_click" {
            $button = [string](Get-McpArgument -Arguments $Arguments -Name "button" -Default "left")
            $clicks = [int](Get-McpArgument -Arguments $Arguments -Name "clicks" -Default 1)
            $x = Get-McpArgument -Arguments $Arguments -Name "x"
            $y = Get-McpArgument -Arguments $Arguments -Name "y"
            if ($null -ne $x -and $null -ne $y) {
                $result = Invoke-SystemAccessMouseClick -Button $button -X ([int]$x) -Y ([int]$y) -Clicks $clicks
            }
            else {
                $result = Invoke-SystemAccessMouseClick -Button $button -Clicks $clicks
            }

            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value $result)
            }
        }
        "keyboard_type" {
            $text = Get-McpArgument -Arguments $Arguments -Name "text" -Default ""
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value (Invoke-SystemAccessKeyboardType -Text ([string]$text)))
            }
        }
        "keyboard_key" {
            $virtualKey = Assert-McpArgument -Arguments $Arguments -Name "virtualKey"
            $action = [string](Get-McpArgument -Arguments $Arguments -Name "action" -Default "press")
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value (Invoke-SystemAccessKeyboardKey -VirtualKey ([int]$virtualKey) -Action $action))
            }
        }
        "cursor_state" {
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value (Get-SystemAccessCursorState))
            }
        }
        "window_foreground" {
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value (Get-SystemAccessForegroundWindow))
            }
        }
        "window_hover" {
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value (Get-SystemAccessWindowHover))
            }
        }
        "window_from_point" {
            $x = Get-McpArgument -Arguments $Arguments -Name "x"
            $y = Get-McpArgument -Arguments $Arguments -Name "y"
            if ($null -ne $x -and $null -ne $y) {
                $result = Get-SystemAccessWindowFromPoint -X ([int]$x) -Y ([int]$y)
            }
            else {
                $result = Get-SystemAccessWindowFromPoint
            }
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value $result)
            }
        }
        "window_list" {
            $limit = [int](Get-McpArgument -Arguments $Arguments -Name "limit" -Default 100)
            $includeUntitled = [bool](Get-McpArgument -Arguments $Arguments -Name "includeUntitled" -Default $false)
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value @(Get-SystemAccessWindowList -Limit $limit -IncludeUntitled $includeUntitled))
            }
        }
        "screen_state" {
            $includeWindows = [bool](Get-McpArgument -Arguments $Arguments -Name "includeWindows" -Default $false)
            $windowLimit = [int](Get-McpArgument -Arguments $Arguments -Name "windowLimit" -Default 100)
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value (Get-SystemAccessScreenState -IncludeWindows $includeWindows -WindowLimit $windowLimit))
            }
        }
        "hyperv_status" {
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value (Get-SystemAccessHyperVStatus))
            }
        }
        "hyperv_list_vms" {
            $server = [string](Get-McpArgument -Arguments $Arguments -Name "server" -Default "localhost")
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value @(Get-SystemAccessHyperVVMs -Server $server))
            }
        }
        "hyperv_start_vm" {
            $vmName = [string](Assert-McpArgument -Arguments $Arguments -Name "vmName")
            $server = [string](Get-McpArgument -Arguments $Arguments -Name "server" -Default "localhost")
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value (Start-SystemAccessHyperVVM -VMName $vmName -Server $server))
            }
        }
        "hyperv_console_windows" {
            $vmName = [string](Get-McpArgument -Arguments $Arguments -Name "vmName" -Default "")
            $server = [string](Get-McpArgument -Arguments $Arguments -Name "server" -Default "")
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value @(Get-SystemAccessHyperVConsoleWindows -VMName $vmName -Server $server))
            }
        }
        "hyperv_console_connect" {
            $vmName = [string](Assert-McpArgument -Arguments $Arguments -Name "vmName")
            $server = [string](Get-McpArgument -Arguments $Arguments -Name "server" -Default "localhost")
            $timeoutMs = [int](Get-McpArgument -Arguments $Arguments -Name "timeoutMs" -Default 15000)
            $forceNew = [bool](Get-McpArgument -Arguments $Arguments -Name "forceNew" -Default $false)
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value (Open-SystemAccessHyperVConsole -VMName $vmName -Server $server -TimeoutMs $timeoutMs -ForceNew:$forceNew))
            }
        }
        "hyperv_console_pointer_state" {
            $vmName = [string](Get-McpArgument -Arguments $Arguments -Name "vmName" -Default "")
            $server = [string](Get-McpArgument -Arguments $Arguments -Name "server" -Default "localhost")
            $processId = [int](Get-McpArgument -Arguments $Arguments -Name "processId" -Default 0)
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value (Get-SystemAccessHyperVConsolePointerState -VMName $vmName -Server $server -ProcessId $processId))
            }
        }
        "hyperv_console_screenshot" {
            $vmName = [string](Get-McpArgument -Arguments $Arguments -Name "vmName" -Default "")
            $server = [string](Get-McpArgument -Arguments $Arguments -Name "server" -Default "localhost")
            $processId = [int](Get-McpArgument -Arguments $Arguments -Name "processId" -Default 0)
            $area = [string](Get-McpArgument -Arguments $Arguments -Name "area" -Default "window")
            $shot = Get-SystemAccessHyperVConsoleScreenshot -VMName $vmName -Server $server -ProcessId $processId -Area $area
            return [pscustomobject]@{
                content = @(
                    [pscustomobject]@{
                        type = "image"
                        data = $shot.data
                        mimeType = $shot.mimeType
                    },
                    (New-JsonTextContent -Value ([pscustomobject]@{
                        width = $shot.width
                        height = $shot.height
                        left = $shot.left
                        top = $shot.top
                        area = $shot.area
                        console = $shot.console
                        timestamp = $shot.timestamp
                    }))
                )
            }
        }
        "hyperv_console_mouse_move" {
            $x = [int](Assert-McpArgument -Arguments $Arguments -Name "x")
            $y = [int](Assert-McpArgument -Arguments $Arguments -Name "y")
            $vmName = [string](Get-McpArgument -Arguments $Arguments -Name "vmName" -Default "")
            $server = [string](Get-McpArgument -Arguments $Arguments -Name "server" -Default "localhost")
            $processId = [int](Get-McpArgument -Arguments $Arguments -Name "processId" -Default 0)
            $area = [string](Get-McpArgument -Arguments $Arguments -Name "area" -Default "window")
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value (Invoke-SystemAccessHyperVConsoleMouseMove -X $x -Y $y -VMName $vmName -Server $server -ProcessId $processId -Area $area))
            }
        }
        "hyperv_console_mouse_click" {
            $x = [int](Assert-McpArgument -Arguments $Arguments -Name "x")
            $y = [int](Assert-McpArgument -Arguments $Arguments -Name "y")
            $button = [string](Get-McpArgument -Arguments $Arguments -Name "button" -Default "left")
            $clicks = [int](Get-McpArgument -Arguments $Arguments -Name "clicks" -Default 1)
            $vmName = [string](Get-McpArgument -Arguments $Arguments -Name "vmName" -Default "")
            $server = [string](Get-McpArgument -Arguments $Arguments -Name "server" -Default "localhost")
            $processId = [int](Get-McpArgument -Arguments $Arguments -Name "processId" -Default 0)
            $area = [string](Get-McpArgument -Arguments $Arguments -Name "area" -Default "window")
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value (Invoke-SystemAccessHyperVConsoleMouseClick -X $x -Y $y -Button $button -Clicks $clicks -VMName $vmName -Server $server -ProcessId $processId -Area $area))
            }
        }
        "hyperv_console_keyboard_type" {
            $text = [string](Get-McpArgument -Arguments $Arguments -Name "text" -Default "")
            $vmName = [string](Get-McpArgument -Arguments $Arguments -Name "vmName" -Default "")
            $server = [string](Get-McpArgument -Arguments $Arguments -Name "server" -Default "localhost")
            $processId = [int](Get-McpArgument -Arguments $Arguments -Name "processId" -Default 0)
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value (Invoke-SystemAccessHyperVConsoleKeyboardType -Text $text -VMName $vmName -Server $server -ProcessId $processId))
            }
        }
        "hyperv_console_keyboard_key" {
            $virtualKey = [int](Assert-McpArgument -Arguments $Arguments -Name "virtualKey")
            $action = [string](Get-McpArgument -Arguments $Arguments -Name "action" -Default "press")
            $vmName = [string](Get-McpArgument -Arguments $Arguments -Name "vmName" -Default "")
            $server = [string](Get-McpArgument -Arguments $Arguments -Name "server" -Default "localhost")
            $processId = [int](Get-McpArgument -Arguments $Arguments -Name "processId" -Default 0)
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value (Invoke-SystemAccessHyperVConsoleKeyboardKey -VirtualKey $virtualKey -Action $action -VMName $vmName -Server $server -ProcessId $processId))
            }
        }
        "hyperv_console_ctrl_alt_delete" {
            $vmName = [string](Get-McpArgument -Arguments $Arguments -Name "vmName" -Default "")
            $server = [string](Get-McpArgument -Arguments $Arguments -Name "server" -Default "localhost")
            $processId = [int](Get-McpArgument -Arguments $Arguments -Name "processId" -Default 0)
            return [pscustomobject]@{
                content = @(New-JsonTextContent -Value (Invoke-SystemAccessHyperVConsoleCtrlAltDelete -VMName $vmName -Server $server -ProcessId $processId))
            }
        }
        default {
            throw "Unknown tool: $Name"
        }
    }
}

function Invoke-McpRequest {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Request
    )

    $id = Get-McpArgument -Arguments $Request -Name "id"
    $method = [string](Get-McpArgument -Arguments $Request -Name "method" -Default "")

    switch ($method) {
        "initialize" {
            return (New-McpResponse -Id $id -Result ([pscustomobject]@{
                protocolVersion = $ProtocolVersion
                capabilities = [pscustomobject]@{
                    tools = [pscustomobject]@{
                        listChanged = $false
                    }
                }
                serverInfo = [pscustomobject]@{
                    name = "system-access-mcp"
                    version = "0.1.0"
                    profile = $script:McpProfile
                    description = (Get-McpProfileDescription)
                }
            }))
        }
        "notifications/initialized" {
            return $null
        }
        "tools/list" {
            return (New-McpResponse -Id $id -Result ([pscustomobject]@{
                tools = @(Get-McpTools)
            }))
        }
        "tools/call" {
            $name = [string]$Request.params.name
            $arguments = $Request.params.arguments
            return (New-McpResponse -Id $id -Result (Invoke-McpTool -Name $name -Arguments $arguments))
        }
        "ping" {
            return (New-McpResponse -Id $id -Result ([pscustomobject]@{}))
        }
        default {
            if ($null -ne $id) {
                return (New-McpErrorResponse -Id $id -Code -32601 -Message "Method not found: $method")
            }
            return $null
        }
    }
}

if ($env:SYSTEM_ACCESS_MCP_IMPORT_ONLY -eq "1") {
    return
}

while ($true) {
    $line = [Console]::In.ReadLine()
    if ($null -eq $line) {
        break
    }
    if ([string]::IsNullOrWhiteSpace($line)) {
        continue
    }

    $request = $null
    try {
        $request = $line | ConvertFrom-Json
        $response = Invoke-McpRequest -Request $request
        if ($null -ne $response) {
            Write-McpMessage $response
        }
    }
    catch {
        $responseId = $null
        if ($null -ne $request) {
            $responseId = Get-McpArgument -Arguments $request -Name "id"
        }

        Write-McpMessage (New-McpErrorResponse -Id $responseId -Message $_.Exception.Message)
    }
}
