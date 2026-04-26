Set-StrictMode -Version 2.0

if (-not (Get-Command ConvertTo-SystemAccessJson -ErrorAction SilentlyContinue)) {
    . "$PSScriptRoot\SystemAccess.Core.ps1"
}

if (-not ("SystemAccess.Windowing" -as [type])) {
    Add-Type -Language CSharp -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace SystemAccess
{
    public static class Windowing
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        private const int SW_SHOWNORMAL = 1;
        private const int SW_RESTORE = 9;

        public static bool IsValidWindow(IntPtr handle)
        {
            return handle != IntPtr.Zero && IsWindow(handle);
        }

        public static void ActivateWindow(IntPtr handle)
        {
            if (!IsValidWindow(handle))
            {
                throw new InvalidOperationException("Window handle is not valid.");
            }

            if (IsIconic(handle))
            {
                ShowWindow(handle, SW_RESTORE);
            }
            else
            {
                ShowWindow(handle, SW_SHOWNORMAL);
            }

            SetForegroundWindow(handle);
        }

        public static int[] GetWindowRectangle(IntPtr handle)
        {
            RECT rect;
            if (!GetWindowRect(handle, out rect))
            {
                throw new InvalidOperationException("GetWindowRect failed.");
            }

            return new int[] { rect.Left, rect.Top, rect.Right, rect.Bottom };
        }

        public static int[] GetClientRectangleOnScreen(IntPtr handle)
        {
            RECT client;
            if (!GetClientRect(handle, out client))
            {
                throw new InvalidOperationException("GetClientRect failed.");
            }

            POINT point = new POINT();
            point.X = 0;
            point.Y = 0;
            if (!ClientToScreen(handle, ref point))
            {
                throw new InvalidOperationException("ClientToScreen failed.");
            }

            return new int[] {
                point.X,
                point.Y,
                point.X + (client.Right - client.Left),
                point.Y + (client.Bottom - client.Top)
            };
        }
    }
}
"@
}

function ConvertTo-HyperVQuotedArgument {
    param(
        [AllowEmptyString()]
        [string] $Value
    )

    return '"' + ($Value -replace '"', '\"') + '"'
}

function ConvertFrom-HyperVRectangle {
    param(
        [Parameter(Mandatory = $true)]
        [int[]] $Rectangle
    )

    $width = $Rectangle[2] - $Rectangle[0]
    $height = $Rectangle[3] - $Rectangle[1]

    [pscustomobject]@{
        left = $Rectangle[0]
        top = $Rectangle[1]
        right = $Rectangle[2]
        bottom = $Rectangle[3]
        width = $width
        height = $height
    }
}

function Get-SystemAccessHyperVVmConnectPath {
    $command = Get-Command vmconnect.exe -ErrorAction SilentlyContinue
    if ($null -ne $command) {
        return $command.Source
    }

    $candidates = @(
        "$env:WINDIR\System32\vmconnect.exe",
        "$env:WINDIR\Sysnative\vmconnect.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    return $null
}

function Test-SystemAccessHyperVPowerShell {
    return ($null -ne (Get-Command Get-VM -ErrorAction SilentlyContinue))
}

function Get-SystemAccessHyperVStatus {
    [pscustomobject]@{
        provider = "HyperVVmConnect"
        providerScope = "host-side VMConnect console windows"
        hyperVPowerShellAvailable = (Test-SystemAccessHyperVPowerShell)
        vmConnectPath = (Get-SystemAccessHyperVVmConnectPath)
        vmConnectAvailable = ($null -ne (Get-SystemAccessHyperVVmConnectPath))
        hostName = $env:COMPUTERNAME
        userName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        consoleWindows = @(Get-SystemAccessHyperVConsoleWindows)
        timestamp = (Get-Date).ToUniversalTime().ToString("o")
    }
}

function Get-SystemAccessHyperVVMs {
    param(
        [string] $Server = "localhost"
    )

    $command = Get-Command Get-VM -ErrorAction SilentlyContinue
    if ($null -eq $command) {
        throw "Hyper-V PowerShell module is not available. Install Hyper-V Management Tools on the host."
    }

    $parameters = @{}
    if (-not [string]::IsNullOrWhiteSpace($Server)) {
        $parameters["ComputerName"] = $Server
    }
    $parameters["ErrorAction"] = "Stop"

    $vms = & $command @parameters
    @($vms | ForEach-Object {
        [pscustomobject]@{
            name = $_.Name
            id = $_.Id.ToString()
            state = $_.State.ToString()
            status = $_.Status
            cpuUsage = $_.CPUUsage
            memoryAssigned = $_.MemoryAssigned
            uptime = $_.Uptime.ToString()
            generation = $_.Generation
            computerName = $_.ComputerName
        }
    })
}

function Start-SystemAccessHyperVVM {
    param(
        [Parameter(Mandatory = $true)]
        [string] $VMName,

        [string] $Server = "localhost"
    )

    $getCommand = Get-Command Get-VM -ErrorAction SilentlyContinue
    $startCommand = Get-Command Start-VM -ErrorAction SilentlyContinue
    if ($null -eq $getCommand -or $null -eq $startCommand) {
        throw "Hyper-V PowerShell module is not available. Install Hyper-V Management Tools on the host."
    }

    $parameters = @{
        Name = $VMName
    }
    if (-not [string]::IsNullOrWhiteSpace($Server)) {
        $parameters["ComputerName"] = $Server
    }

    $vm = & $getCommand @parameters
    $wasAlreadyRunning = ($vm.State.ToString() -eq "Running")
    if (-not $wasAlreadyRunning) {
        $vm = & $startCommand @parameters -PassThru -WarningAction SilentlyContinue
    }

    [pscustomobject]@{
        ok = $true
        action = "hyperv_start_vm"
        name = $vm.Name
        state = $vm.State.ToString()
        computerName = $vm.ComputerName
        alreadyRunning = $wasAlreadyRunning
    }
}

function Get-SystemAccessHyperVConsoleWindows {
    param(
        [string] $VMName = "",

        [string] $Server = ""
    )

    $processes = @(Get-Process vmconnect -ErrorAction SilentlyContinue)
    $windows = foreach ($process in $processes) {
        if ($process.MainWindowHandle -eq [IntPtr]::Zero) {
            continue
        }

        $title = $process.MainWindowTitle
        if (-not [string]::IsNullOrWhiteSpace($VMName) -and $title -notlike "*$VMName*") {
            continue
        }

        if (-not [string]::IsNullOrWhiteSpace($Server) -and $Server -ne "localhost" -and $title -notlike "*$Server*") {
            continue
        }

        $handle = [IntPtr]::new([int64]$process.MainWindowHandle)
        $rect = ConvertFrom-HyperVRectangle -Rectangle ([SystemAccess.Windowing]::GetWindowRectangle($handle))
        $startTime = $null
        try {
            $startTime = $process.StartTime.ToUniversalTime().ToString("o")
        }
        catch {
            $startTime = $null
        }

        [pscustomobject]@{
            processId = $process.Id
            windowHandle = $process.MainWindowHandle.ToInt64()
            title = $title
            rect = $rect
            startTime = $startTime
        }
    }

    @($windows | Sort-Object processId -Descending)
}

function Resolve-SystemAccessHyperVConsoleWindow {
    param(
        [string] $VMName = "",

        [string] $Server = "localhost",

        [int] $ProcessId = 0,

        [int] $TimeoutMs = 0
    )

    $started = Get-Date
    do {
        if ($ProcessId -gt 0) {
            $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
            if ($null -ne $process -and $process.MainWindowHandle -ne [IntPtr]::Zero) {
                $handle = [IntPtr]::new([int64]$process.MainWindowHandle)
                $rect = ConvertFrom-HyperVRectangle -Rectangle ([SystemAccess.Windowing]::GetWindowRectangle($handle))

                return [pscustomobject]@{
                    processId = $process.Id
                    windowHandle = $process.MainWindowHandle.ToInt64()
                    title = $process.MainWindowTitle
                    rect = $rect
                }
            }
        }
        else {
            $windows = @(Get-SystemAccessHyperVConsoleWindows -VMName $VMName -Server $Server)
            if ($windows.Count -eq 1) {
                return $windows[0]
            }
            if ($windows.Count -gt 1) {
                return $windows[0]
            }
        }

        if ($TimeoutMs -le 0) {
            break
        }

        Start-Sleep -Milliseconds 250
    }
    while (((Get-Date) - $started).TotalMilliseconds -lt $TimeoutMs)

    if ($ProcessId -gt 0) {
        throw "No VMConnect window found for process id $ProcessId."
    }

    if ([string]::IsNullOrWhiteSpace($VMName)) {
        throw "No VMConnect window found. Provide vmName or processId."
    }

    throw "No VMConnect window found for VM '$VMName' on '$Server'."
}

function Open-SystemAccessHyperVConsole {
    param(
        [Parameter(Mandatory = $true)]
        [string] $VMName,

        [string] $Server = "localhost",

        [int] $TimeoutMs = 15000,

        [switch] $ForceNew
    )

    if (-not $ForceNew) {
        $existing = @(Get-SystemAccessHyperVConsoleWindows -VMName $VMName -Server $Server)
        if ($existing.Count -gt 0) {
            $handle = [IntPtr]::new([int64]$existing[0].windowHandle)
            [SystemAccess.Windowing]::ActivateWindow($handle)
            return [pscustomobject]@{
                ok = $true
                action = "hyperv_console_connect"
                reused = $true
                console = $existing[0]
            }
        }
    }

    $vmConnectPath = Get-SystemAccessHyperVVmConnectPath
    if ($null -eq $vmConnectPath) {
        throw "vmconnect.exe is not available. Install Hyper-V Management Tools on the host."
    }

    $arguments = (ConvertTo-HyperVQuotedArgument -Value $Server) + " " + (ConvertTo-HyperVQuotedArgument -Value $VMName)
    $process = Start-Process -FilePath $vmConnectPath -ArgumentList $arguments -WindowStyle Normal -PassThru
    $window = Resolve-SystemAccessHyperVConsoleWindow -VMName $VMName -Server $Server -ProcessId $process.Id -TimeoutMs $TimeoutMs
    $handle = [IntPtr]::new([int64]$window.windowHandle)
    [SystemAccess.Windowing]::ActivateWindow($handle)

    [pscustomobject]@{
        ok = $true
        action = "hyperv_console_connect"
        reused = $false
        console = $window
    }
}

function Get-SystemAccessHyperVConsoleRect {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Window,

        [ValidateSet("window", "client")]
        [string] $Area = "window"
    )

    $handle = [IntPtr]::new([int64]$Window.windowHandle)
    if ($Area -eq "client") {
        return ConvertFrom-HyperVRectangle -Rectangle ([SystemAccess.Windowing]::GetClientRectangleOnScreen($handle))
    }

    return ConvertFrom-HyperVRectangle -Rectangle ([SystemAccess.Windowing]::GetWindowRectangle($handle))
}

function ConvertTo-SystemAccessRelativePoint {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Cursor,

        [Parameter(Mandatory = $true)]
        [object] $Rectangle
    )

    [pscustomobject]@{
        x = $Cursor.position.x - $Rectangle.left
        y = $Cursor.position.y - $Rectangle.top
    }
}

function Get-SystemAccessHyperVConsolePointerState {
    param(
        [string] $VMName = "",

        [string] $Server = "localhost",

        [int] $ProcessId = 0
    )

    $window = Resolve-SystemAccessHyperVConsoleWindow -VMName $VMName -Server $Server -ProcessId $ProcessId
    $windowRect = Get-SystemAccessHyperVConsoleRect -Window $window -Area "window"
    $clientRect = Get-SystemAccessHyperVConsoleRect -Window $window -Area "client"
    $cursor = Get-SystemAccessCursorState

    [pscustomobject]@{
        provider = "HyperVVmConnect"
        cursor = $cursor
        console = $window
        windowRect = $windowRect
        clientRect = $clientRect
        relativeToWindow = (ConvertTo-SystemAccessRelativePoint -Cursor $cursor -Rectangle $windowRect)
        relativeToClient = (ConvertTo-SystemAccessRelativePoint -Cursor $cursor -Rectangle $clientRect)
        insideWindow = (Test-SystemAccessPointInRectangle -Point $cursor.position -Rectangle $windowRect)
        insideClient = (Test-SystemAccessPointInRectangle -Point $cursor.position -Rectangle $clientRect)
        foregroundWindow = Get-SystemAccessForegroundWindow
        hoverWindow = Get-SystemAccessWindowHover
        timestamp = (Get-Date).ToUniversalTime().ToString("o")
    }
}

function Get-SystemAccessHyperVConsoleScreenshotBytes {
    param(
        [string] $VMName = "",

        [string] $Server = "localhost",

        [int] $ProcessId = 0,

        [ValidateSet("window", "client")]
        [string] $Area = "window",

        [switch] $NoActivate
    )

    $window = Resolve-SystemAccessHyperVConsoleWindow -VMName $VMName -Server $Server -ProcessId $ProcessId
    $handle = [IntPtr]::new([int64]$window.windowHandle)
    if (-not $NoActivate) {
        [SystemAccess.Windowing]::ActivateWindow($handle)
        Start-Sleep -Milliseconds 200
    }

    $rect = Get-SystemAccessHyperVConsoleRect -Window $window -Area $Area
    if ($rect.width -le 0 -or $rect.height -le 0) {
        throw "VMConnect window has an invalid capture rectangle."
    }

    $bitmap = New-Object System.Drawing.Bitmap $rect.width, $rect.height, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $stream = New-Object System.IO.MemoryStream

    try {
        $graphics.CopyFromScreen($rect.left, $rect.top, 0, 0, (New-Object System.Drawing.Size $rect.width, $rect.height))
        $bitmap.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
        [pscustomobject]@{
            bytes = $stream.ToArray()
            rect = $rect
            console = $window
            area = $Area
        }
    }
    finally {
        $stream.Dispose()
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

function Get-SystemAccessHyperVConsoleScreenshot {
    param(
        [string] $VMName = "",

        [string] $Server = "localhost",

        [int] $ProcessId = 0,

        [ValidateSet("window", "client")]
        [string] $Area = "window"
    )

    $capture = Get-SystemAccessHyperVConsoleScreenshotBytes -VMName $VMName -Server $Server -ProcessId $ProcessId -Area $Area

    [pscustomobject]@{
        mimeType = "image/png"
        data = [Convert]::ToBase64String($capture.bytes)
        width = $capture.rect.width
        height = $capture.rect.height
        left = $capture.rect.left
        top = $capture.rect.top
        area = $capture.area
        console = $capture.console
        timestamp = (Get-Date).ToUniversalTime().ToString("o")
    }
}

function Invoke-SystemAccessHyperVConsoleMouseMove {
    param(
        [Parameter(Mandatory = $true)]
        [int] $X,

        [Parameter(Mandatory = $true)]
        [int] $Y,

        [string] $VMName = "",

        [string] $Server = "localhost",

        [int] $ProcessId = 0,

        [ValidateSet("window", "client")]
        [string] $Area = "window"
    )

    $window = Resolve-SystemAccessHyperVConsoleWindow -VMName $VMName -Server $Server -ProcessId $ProcessId
    $handle = [IntPtr]::new([int64]$window.windowHandle)
    [SystemAccess.Windowing]::ActivateWindow($handle)
    Start-Sleep -Milliseconds 100

    $rect = Get-SystemAccessHyperVConsoleRect -Window $window -Area $Area
    $screenX = $rect.left + $X
    $screenY = $rect.top + $Y
    Invoke-SystemAccessMouseMove -X $screenX -Y $screenY | Out-Null

    [pscustomobject]@{
        ok = $true
        action = "hyperv_console_mouse_move"
        vmName = $VMName
        processId = $window.processId
        area = $Area
        x = $X
        y = $Y
        screenX = $screenX
        screenY = $screenY
    }
}

function Invoke-SystemAccessHyperVConsoleMouseClick {
    param(
        [Parameter(Mandatory = $true)]
        [int] $X,

        [Parameter(Mandatory = $true)]
        [int] $Y,

        [ValidateSet("left", "right", "middle")]
        [string] $Button = "left",

        [ValidateRange(1, 10)]
        [int] $Clicks = 1,

        [string] $VMName = "",

        [string] $Server = "localhost",

        [int] $ProcessId = 0,

        [ValidateSet("window", "client")]
        [string] $Area = "window"
    )

    $move = Invoke-SystemAccessHyperVConsoleMouseMove -X $X -Y $Y -VMName $VMName -Server $Server -ProcessId $ProcessId -Area $Area
    Invoke-SystemAccessMouseClick -Button $Button -Clicks $Clicks | Out-Null

    [pscustomobject]@{
        ok = $true
        action = "hyperv_console_mouse_click"
        vmName = $VMName
        processId = $move.processId
        area = $Area
        button = $Button
        clicks = $Clicks
        x = $X
        y = $Y
        screenX = $move.screenX
        screenY = $move.screenY
    }
}

function Invoke-SystemAccessHyperVConsoleKeyboardType {
    param(
        [AllowEmptyString()]
        [string] $Text = "",

        [string] $VMName = "",

        [string] $Server = "localhost",

        [int] $ProcessId = 0
    )

    $window = Resolve-SystemAccessHyperVConsoleWindow -VMName $VMName -Server $Server -ProcessId $ProcessId
    $handle = [IntPtr]::new([int64]$window.windowHandle)
    [SystemAccess.Windowing]::ActivateWindow($handle)
    Start-Sleep -Milliseconds 100
    Invoke-SystemAccessKeyboardType -Text $Text | Out-Null

    [pscustomobject]@{
        ok = $true
        action = "hyperv_console_keyboard_type"
        vmName = $VMName
        processId = $window.processId
        length = $Text.Length
    }
}

function Invoke-SystemAccessHyperVConsoleKeyboardKey {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 255)]
        [int] $VirtualKey,

        [ValidateSet("press", "down", "up")]
        [string] $Action = "press",

        [string] $VMName = "",

        [string] $Server = "localhost",

        [int] $ProcessId = 0
    )

    $window = Resolve-SystemAccessHyperVConsoleWindow -VMName $VMName -Server $Server -ProcessId $ProcessId
    $handle = [IntPtr]::new([int64]$window.windowHandle)
    [SystemAccess.Windowing]::ActivateWindow($handle)
    Start-Sleep -Milliseconds 100
    Invoke-SystemAccessKeyboardKey -VirtualKey $VirtualKey -Action $Action | Out-Null

    [pscustomobject]@{
        ok = $true
        action = "hyperv_console_keyboard_key"
        vmName = $VMName
        processId = $window.processId
        virtualKey = $VirtualKey
        keyAction = $Action
    }
}

function Invoke-SystemAccessHyperVConsoleCtrlAltDelete {
    param(
        [string] $VMName = "",

        [string] $Server = "localhost",

        [int] $ProcessId = 0
    )

    $window = Resolve-SystemAccessHyperVConsoleWindow -VMName $VMName -Server $Server -ProcessId $ProcessId
    $handle = [IntPtr]::new([int64]$window.windowHandle)
    [SystemAccess.Windowing]::ActivateWindow($handle)
    Start-Sleep -Milliseconds 100

    [SystemAccess.NativeInput]::KeyVirtual([uint16]17, $true)
    [SystemAccess.NativeInput]::KeyVirtual([uint16]18, $true)
    [SystemAccess.NativeInput]::KeyVirtual([uint16]35, $true)
    Start-Sleep -Milliseconds 35
    [SystemAccess.NativeInput]::KeyVirtual([uint16]35, $false)
    [SystemAccess.NativeInput]::KeyVirtual([uint16]18, $false)
    [SystemAccess.NativeInput]::KeyVirtual([uint16]17, $false)

    [pscustomobject]@{
        ok = $true
        action = "hyperv_console_ctrl_alt_delete"
        vmName = $VMName
        processId = $window.processId
        sent = "Ctrl+Alt+End"
    }
}
