param(
    [ValidateRange(1, 65535)]
    [int] $Port = 8890
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $PSScriptRoot
$Results = New-Object System.Collections.Generic.List[object]
$Failed = $false

function Add-Result {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Name,

        [Parameter(Mandatory = $true)]
        [ValidateSet("pass", "fail", "skip")]
        [string] $Status,

        [string] $Details = ""
    )

    $script:Results.Add([pscustomobject]@{
        name = $Name
        status = $Status
        details = $Details
    }) | Out-Null

    if ($Status -eq "fail") {
        $script:Failed = $true
    }
}

function Invoke-Check {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Name,

        [Parameter(Mandatory = $true)]
        [scriptblock] $ScriptBlock
    )

    try {
        $details = & $ScriptBlock
        Add-Result -Name $Name -Status pass -Details ([string]$details)
    }
    catch {
        Add-Result -Name $Name -Status fail -Details $_.Exception.Message
    }
}

function Test-PowerShellSyntax {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $errors = $null
    [System.Management.Automation.PSParser]::Tokenize((Get-Content -Raw $Path), [ref]$errors) | Out-Null
    if ($errors.Count -gt 0) {
        throw (($errors | ForEach-Object { $_.Message }) -join "; ")
    }
}

function Get-McpToolNamesForProfile {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("All", "GuestDesktop", "HostHyperV")]
        [string] $Profile
    )

    $messages = @(
        '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"smoke","version":"0"}}}',
        '{"jsonrpc":"2.0","method":"notifications/initialized"}',
        '{"jsonrpc":"2.0","id":2,"method":"tools/list"}'
    )

    $output = $messages | powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Start-McpServer.ps1 -Profile $Profile
    $responses = @($output | ForEach-Object { $_ | ConvertFrom-Json })
    @($responses | Where-Object { $_.id -eq 2 } | ForEach-Object { $_.result.tools.name })
}

function Assert-ToolPresent {
    param(
        [string[]] $Tools,
        [Parameter(Mandatory = $true)]
        [string] $Name
    )

    if ($Tools -notcontains $Name) {
        throw "$Name was not advertised"
    }
}

function Assert-ToolAbsent {
    param(
        [string[]] $Tools,
        [Parameter(Mandatory = $true)]
        [string] $Name
    )

    if ($Tools -contains $Name) {
        throw "$Name should not be advertised"
    }
}

Push-Location $Root
try {
    Invoke-Check "PowerShell syntax" {
        @(
            ".\Start-WebServer.ps1",
            ".\Start-McpServer.ps1",
            ".\Start-McpHttpServer.ps1",
            ".\src\SystemAccess.Core.ps1",
            ".\src\SystemAccess.HyperV.ps1"
        ) | ForEach-Object { Test-PowerShellSyntax -Path $_ }
        "syntax ok"
    }

    Invoke-Check "Core status" {
        . .\src\SystemAccess.Core.ps1
        $status = Get-SystemAccessStatus
        "$($status.provider) $($status.virtualScreen.width)x$($status.virtualScreen.height)"
    }

    Invoke-Check "Core desktop observation" {
        . .\src\SystemAccess.Core.ps1
        $cursor = Get-SystemAccessCursorState
        if ($null -eq $cursor.position -or $null -eq $cursor.virtualScreen) {
            throw "cursor state did not include position and virtualScreen"
        }

        $screen = Get-SystemAccessScreenState
        foreach ($propertyName in @("cursor", "foregroundWindow", "hoverWindow")) {
            if ($screen.PSObject.Properties.Match($propertyName).Count -eq 0) {
                throw "screen state did not include $propertyName"
            }
        }
        if ($null -eq $screen.cursor -or $null -eq $screen.hoverWindow) {
            throw "screen state did not include cursor and hoverWindow values"
        }

        "cursor=$($cursor.position.x),$($cursor.position.y)"
    }

    Invoke-Check "JSON serialization helpers" {
        . .\src\SystemAccess.Core.ps1
        $emptyArrayJson = ConvertTo-SystemAccessJson -Value @()
        if ($emptyArrayJson -ne "[]") {
            throw "empty array serialized as '$emptyArrayJson'"
        }
        "empty array ok"
    }

    Invoke-Check "Core screenshot" {
        . .\src\SystemAccess.Core.ps1
        $shot = Get-SystemAccessScreenshot
        if ($shot.data.Length -le 0) {
            throw "screenshot data was empty"
        }
        "$($shot.width)x$($shot.height), $($shot.data.Length) base64 chars"
    }

    Invoke-Check "Core keyboard SendInput" {
        . .\src\SystemAccess.Core.ps1
        $result = Invoke-SystemAccessKeyboardKey -VirtualKey 16
        "vk=$($result.virtualKey), ok=$($result.ok)"
    }

    Invoke-Check "Hyper-V status" {
        . .\src\SystemAccess.Core.ps1
        . .\src\SystemAccess.HyperV.ps1
        $status = Get-SystemAccessHyperVStatus
        "PowerShell=$($status.hyperVPowerShellAvailable), VMConnect=$($status.vmConnectAvailable)"
    }

    . .\src\SystemAccess.Core.ps1
    . .\src\SystemAccess.HyperV.ps1
    $hyperVStatus = Get-SystemAccessHyperVStatus
    if ($hyperVStatus.hyperVPowerShellAvailable) {
        Invoke-Check "Hyper-V list VMs" {
            $vms = @(Get-SystemAccessHyperVVMs -Server localhost)
            "$($vms.Count) VMs returned"
        }
    }
    else {
        Add-Result -Name "Hyper-V list VMs" -Status skip -Details "Hyper-V PowerShell module is not available on this machine."
    }

    Invoke-Check "MCP handshake and tools" {
        $messages = @(
            '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"smoke","version":"0"}}}',
            '{"jsonrpc":"2.0","method":"notifications/initialized"}',
            '{"jsonrpc":"2.0","id":2,"method":"tools/list"}',
            '{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"hyperv_status","arguments":{}}}'
        )

        $output = $messages | powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Start-McpServer.ps1
        $responses = @($output | ForEach-Object { $_ | ConvertFrom-Json })
        $tools = @($responses | Where-Object { $_.id -eq 2 } | ForEach-Object { $_.result.tools })
        $expectedTools = @(
            "hyperv_console_screenshot",
            "cursor_state",
            "window_foreground",
            "window_hover",
            "window_from_point",
            "window_list",
            "screen_state",
            "hyperv_console_pointer_state"
        )
        foreach ($toolName in $expectedTools) {
            if (-not ($tools | Where-Object { $_.name -eq $toolName })) {
                throw "$toolName tool was not advertised"
            }
        }

        "$($tools.Count) tools advertised"
    }

    Invoke-Check "MCP tool profiles" {
        $guestTools = @(Get-McpToolNamesForProfile -Profile GuestDesktop)
        Assert-ToolPresent -Tools $guestTools -Name "system_status"
        Assert-ToolPresent -Tools $guestTools -Name "mouse_click"
        Assert-ToolPresent -Tools $guestTools -Name "screen_state"
        Assert-ToolAbsent -Tools $guestTools -Name "hyperv_status"
        Assert-ToolAbsent -Tools $guestTools -Name "hyperv_console_screenshot"

        $hostTools = @(Get-McpToolNamesForProfile -Profile HostHyperV)
        Assert-ToolPresent -Tools $hostTools -Name "system_status"
        Assert-ToolPresent -Tools $hostTools -Name "cursor_state"
        Assert-ToolPresent -Tools $hostTools -Name "window_hover"
        Assert-ToolPresent -Tools $hostTools -Name "hyperv_status"
        Assert-ToolPresent -Tools $hostTools -Name "hyperv_console_screenshot"
        Assert-ToolAbsent -Tools $hostTools -Name "mouse_click"
        Assert-ToolAbsent -Tools $hostTools -Name "keyboard_type"

        $allTools = @(Get-McpToolNamesForProfile -Profile All)
        Assert-ToolPresent -Tools $allTools -Name "mouse_click"
        Assert-ToolPresent -Tools $allTools -Name "hyperv_status"

        "guest=$($guestTools.Count), host=$($hostTools.Count), all=$($allTools.Count)"
    }

    Invoke-Check "Web endpoint profiles" {
        $badWebBooleanCalls = @(Select-String -Path .\Start-WebServer.ps1 -Pattern "-Path [`$]path -and|-Path [`$]path -or")
        if ($badWebBooleanCalls.Count -gt 0) {
            throw "web route boolean expression is missing parentheses"
        }

        $previousImportOnly = $env:SYSTEM_ACCESS_WEB_IMPORT_ONLY
        $env:SYSTEM_ACCESS_WEB_IMPORT_ONLY = "1"
        try {
            . .\Start-WebServer.ps1 -Profile GuestDesktop
            if (-not (Test-SystemAccessWebEndpointAllowed -Path "/api/mouse/click")) {
                throw "GuestDesktop did not allow desktop mouse endpoint"
            }
            if (Test-SystemAccessWebEndpointAllowed -Path "/api/hyperv/status") {
                throw "GuestDesktop allowed local Hyper-V endpoint without proxy"
            }
            if (-not (Test-SystemAccessWebEndpointAllowed -Path "/api/window/hover")) {
                throw "GuestDesktop did not allow shared observation endpoint"
            }

            . .\Start-WebServer.ps1 -Profile GuestDesktop -HyperVApiBaseUrl "http://host.example:8765"
            if (-not (Test-SystemAccessWebEndpointAllowed -Path "/api/hyperv/status")) {
                throw "GuestDesktop did not allow proxied Hyper-V endpoint"
            }

            . .\Start-WebServer.ps1 -Profile HostHyperV
            if (Test-SystemAccessWebEndpointAllowed -Path "/api/mouse/click") {
                throw "HostHyperV allowed desktop mouse endpoint"
            }
            if (-not (Test-SystemAccessWebEndpointAllowed -Path "/api/hyperv/status")) {
                throw "HostHyperV did not allow local Hyper-V endpoint"
            }
            if (-not (Test-SystemAccessWebEndpointAllowed -Path "/api/screen/state")) {
                throw "HostHyperV did not allow shared screen state endpoint"
            }

            $html = Get-IndexHtml
            foreach ($expectedText in @(
                'id="obsCursor"',
                'id="obsForeground"',
                'id="obsHover"',
                'id="obsPoint"',
                'id="obsWindows"',
                'id="obsScreen"',
                "/api/window/hover",
                "/api/window/from-point",
                "/api/window/list",
                "/api/screen/state"
            )) {
                if ($html -notlike "*$expectedText*") {
                    throw "web UI did not include observation control '$expectedText'"
                }
            }

            "web profiles ok"
        }
        finally {
            if ($null -eq $previousImportOnly) {
                Remove-Item Env:\SYSTEM_ACCESS_WEB_IMPORT_ONLY -ErrorAction SilentlyContinue
            }
            else {
                $env:SYSTEM_ACCESS_WEB_IMPORT_ONLY = $previousImportOnly
            }
        }
    }

    Invoke-Check "HTTP MCP handshake and tools" {
        $proc = $null
        try {
            $proc = Start-Process -FilePath powershell.exe -ArgumentList @(
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                (Join-Path $Root "Start-McpHttpServer.ps1"),
                "-Port",
                [string]($Port + 1)
            ) -WindowStyle Hidden -PassThru

            $ready = $false
            for ($i = 0; $i -lt 20; $i++) {
                Start-Sleep -Milliseconds 250
                try {
                    Invoke-RestMethod -Uri "http://127.0.0.1:$($Port + 1)/health" -TimeoutSec 2 | Out-Null
                    $ready = $true
                    break
                }
                catch {
                }
            }

            if (-not $ready) {
                throw "HTTP MCP server did not become ready"
            }

            $initialize = Invoke-RestMethod -Uri "http://127.0.0.1:$($Port + 1)/mcp" `
                -Method Post `
                -ContentType "application/json" `
                -Body '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"smoke","version":"0"}}}'
            if ($initialize.result.serverInfo.name -ne "system-access-mcp") {
                throw "unexpected server name"
            }

            $toolsResponse = Invoke-RestMethod -Uri "http://127.0.0.1:$($Port + 1)/mcp" `
                -Method Post `
                -ContentType "application/json" `
                -Body '{"jsonrpc":"2.0","id":2,"method":"tools/list"}'
            $tools = @($toolsResponse.result.tools)
            $expectedTools = @(
                "hyperv_console_screenshot",
                "cursor_state",
                "window_foreground",
                "window_hover",
                "window_from_point",
                "window_list",
                "screen_state",
                "hyperv_console_pointer_state"
            )
            foreach ($toolName in $expectedTools) {
                if (-not ($tools | Where-Object { $_.name -eq $toolName })) {
                    throw "$toolName tool was not advertised"
                }
            }

            "$($tools.Count) tools advertised over HTTP"
        }
        finally {
            if ($null -ne $proc -and -not $proc.HasExited) {
                Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
            }
        }
    }

    Invoke-Check "Web server endpoints" {
        $proc = $null
        try {
            $proc = Start-Process -FilePath powershell.exe -ArgumentList @(
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                (Join-Path $Root "Start-WebServer.ps1"),
                "-Port",
                [string]$Port
            ) -WindowStyle Hidden -PassThru

            $ready = $false
            for ($i = 0; $i -lt 20; $i++) {
                Start-Sleep -Milliseconds 250
                try {
                    $health = Invoke-RestMethod -Uri "http://127.0.0.1:$Port/health" -TimeoutSec 2
                    $ready = $true
                    break
                }
                catch {
                }
            }

            if (-not $ready) {
                throw "web server did not become ready"
            }

            $hyperv = Invoke-RestMethod -Uri "http://127.0.0.1:$Port/api/hyperv/status" -TimeoutSec 2
            $screen = Invoke-RestMethod -Uri "http://127.0.0.1:$Port/api/screen/state" -TimeoutSec 2
            if ($null -eq $screen.cursor -or $null -eq $screen.hoverWindow) {
                throw "screen state endpoint did not include cursor and hoverWindow"
            }

            "health=$($health.provider), hyperv=$($hyperv.provider), cursor=$($screen.cursor.position.x),$($screen.cursor.position.y)"
        }
        finally {
            if ($null -ne $proc -and -not $proc.HasExited) {
                Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
            }
        }
    }
}
finally {
    Pop-Location
}

$Results | Format-Table -AutoSize

if ($Failed) {
    exit 1
}
