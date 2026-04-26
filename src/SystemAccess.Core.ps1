Set-StrictMode -Version 2.0

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

if (-not ("SystemAccess.NativeInput" -as [type])) {
    Add-Type -Language CSharp -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace SystemAccess
{
    public static class NativeInput
    {
        [DllImport("user32.dll")]
        public static extern bool SetProcessDPIAware();

        [DllImport("user32.dll")]
        public static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        private const uint INPUT_KEYBOARD = 1;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const uint KEYEVENTF_UNICODE = 0x0004;

        private const uint MOUSEEVENTF_MOVE = 0x0001;
        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;
        private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
        private const uint MOUSEEVENTF_RIGHTUP = 0x0010;
        private const uint MOUSEEVENTF_MIDDLEDOWN = 0x0020;
        private const uint MOUSEEVENTF_MIDDLEUP = 0x0040;
        private const uint MOUSEEVENTF_WHEEL = 0x0800;

        [StructLayout(LayoutKind.Sequential)]
        private struct INPUT
        {
            public uint type;
            public InputUnion U;
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct InputUnion
        {
            [FieldOffset(0)]
            public KEYBDINPUT ki;

            [FieldOffset(0)]
            public MOUSEINPUT mi;

            [FieldOffset(0)]
            public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        public static void RelativeMouseMove(int dx, int dy)
        {
            mouse_event(MOUSEEVENTF_MOVE, unchecked((uint)dx), unchecked((uint)dy), 0, UIntPtr.Zero);
        }

        public static void MouseButton(string button, bool down)
        {
            uint flag;
            switch ((button ?? "left").ToLowerInvariant())
            {
                case "left":
                    flag = down ? MOUSEEVENTF_LEFTDOWN : MOUSEEVENTF_LEFTUP;
                    break;
                case "right":
                    flag = down ? MOUSEEVENTF_RIGHTDOWN : MOUSEEVENTF_RIGHTUP;
                    break;
                case "middle":
                    flag = down ? MOUSEEVENTF_MIDDLEDOWN : MOUSEEVENTF_MIDDLEUP;
                    break;
                default:
                    throw new ArgumentException("button must be left, right, or middle");
            }

            mouse_event(flag, 0, 0, 0, UIntPtr.Zero);
        }

        public static void MouseWheel(int delta)
        {
            mouse_event(MOUSEEVENTF_WHEEL, 0, 0, unchecked((uint)delta), UIntPtr.Zero);
        }

        public static void KeyVirtual(ushort virtualKey, bool down)
        {
            INPUT input = new INPUT();
            input.type = INPUT_KEYBOARD;
            input.U.ki.wVk = virtualKey;
            input.U.ki.wScan = 0;
            input.U.ki.dwFlags = down ? 0 : KEYEVENTF_KEYUP;
            input.U.ki.time = 0;
            input.U.ki.dwExtraInfo = IntPtr.Zero;
            Send(input);
        }

        public static void TypeText(string text)
        {
            if (text == null)
            {
                return;
            }

            foreach (char ch in text)
            {
                INPUT down = new INPUT();
                down.type = INPUT_KEYBOARD;
                down.U.ki.wVk = 0;
                down.U.ki.wScan = ch;
                down.U.ki.dwFlags = KEYEVENTF_UNICODE;
                down.U.ki.time = 0;
                down.U.ki.dwExtraInfo = IntPtr.Zero;

                INPUT up = down;
                up.U.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;

                Send(down);
                Send(up);
            }
        }

        private static void Send(INPUT input)
        {
            INPUT[] inputs = new INPUT[] { input };
            uint sent = SendInput(1, inputs, Marshal.SizeOf(typeof(INPUT)));
            if (sent != 1)
            {
                int error = Marshal.GetLastWin32Error();
                throw new InvalidOperationException("SendInput failed with Win32 error " + error + ".");
            }
        }
    }
}
"@
}

if (-not ("SystemAccess.Windowing" -as [type])) {
    Add-Type -Language CSharp -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

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

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(POINT point);

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll")]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder className, int count);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

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
        private const uint GA_ROOT = 2;

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

        public static int[] GetCursorPosition()
        {
            POINT point;
            if (!GetCursorPos(out point))
            {
                throw new InvalidOperationException("GetCursorPos failed.");
            }

            return new int[] { point.X, point.Y };
        }

        public static long GetForegroundWindowHandle()
        {
            return GetForegroundWindow().ToInt64();
        }

        public static long GetWindowFromPointHandle(int x, int y)
        {
            POINT point = new POINT();
            point.X = x;
            point.Y = y;
            return WindowFromPoint(point).ToInt64();
        }

        public static long GetRootWindowHandle(IntPtr handle)
        {
            if (!IsValidWindow(handle))
            {
                return 0;
            }

            return GetAncestor(handle, GA_ROOT).ToInt64();
        }

        public static string GetWindowTitle(IntPtr handle)
        {
            StringBuilder text = new StringBuilder(1024);
            GetWindowText(handle, text, text.Capacity);
            return text.ToString();
        }

        public static string GetWindowClassName(IntPtr handle)
        {
            StringBuilder text = new StringBuilder(256);
            GetClassName(handle, text, text.Capacity);
            return text.ToString();
        }

        public static int GetWindowProcessId(IntPtr handle)
        {
            uint processId;
            GetWindowThreadProcessId(handle, out processId);
            return unchecked((int)processId);
        }

        public static bool IsWindowShown(IntPtr handle)
        {
            return IsValidWindow(handle) && IsWindowVisible(handle);
        }

        public static long[] GetTopLevelWindowHandles()
        {
            List<long> handles = new List<long>();
            EnumWindows(delegate(IntPtr hWnd, IntPtr lParam)
            {
                handles.Add(hWnd.ToInt64());
                return true;
            }, IntPtr.Zero);

            return handles.ToArray();
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

[SystemAccess.NativeInput]::SetProcessDPIAware() | Out-Null

function ConvertTo-SystemAccessJson {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Value,

        [int] $Depth = 8
    )

    return (ConvertTo-Json -InputObject $Value -Depth $Depth -Compress)
}

function Get-SystemAccessVirtualScreen {
    $bounds = [System.Windows.Forms.SystemInformation]::VirtualScreen

    [pscustomobject]@{
        left = $bounds.Left
        top = $bounds.Top
        width = $bounds.Width
        height = $bounds.Height
        right = $bounds.Right
        bottom = $bounds.Bottom
    }
}

function Get-SystemAccessStatus {
    [pscustomobject]@{
        provider = "UserSession"
        providerScope = "current interactive user desktop"
        secureDesktopSupported = $false
        logonDesktopSupported = $false
        machineName = $env:COMPUTERNAME
        userName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        virtualScreen = Get-SystemAccessVirtualScreen
        timestamp = (Get-Date).ToUniversalTime().ToString("o")
    }
}

function ConvertFrom-SystemAccessRectangle {
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

function Test-SystemAccessPointInRectangle {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Point,

        [object] $Rectangle
    )

    if ($null -eq $Rectangle) {
        return $false
    }

    return (
        $Point.x -ge $Rectangle.left -and
        $Point.x -lt $Rectangle.right -and
        $Point.y -ge $Rectangle.top -and
        $Point.y -lt $Rectangle.bottom
    )
}

function Get-SystemAccessCursorPosition {
    try {
        return [SystemAccess.Windowing]::GetCursorPosition()
    }
    catch {
        $position = [System.Windows.Forms.Cursor]::Position
        return @([int]$position.X, [int]$position.Y)
    }
}

function Get-SystemAccessCursorState {
    $position = Get-SystemAccessCursorPosition
    $screen = Get-SystemAccessVirtualScreen

    [pscustomobject]@{
        position = [pscustomobject]@{
            x = $position[0]
            y = $position[1]
        }
        virtualPosition = [pscustomobject]@{
            x = $position[0] - $screen.left
            y = $position[1] - $screen.top
        }
        virtualScreen = $screen
        timestamp = (Get-Date).ToUniversalTime().ToString("o")
    }
}

function Get-SystemAccessWindowInfo {
    param(
        [Parameter(Mandatory = $true)]
        [Int64] $WindowHandle
    )

    if ($WindowHandle -eq 0) {
        return $null
    }

    $handle = [IntPtr]::new($WindowHandle)
    if (-not [SystemAccess.Windowing]::IsValidWindow($handle)) {
        return $null
    }

    $rect = $null
    try {
        $rect = ConvertFrom-SystemAccessRectangle -Rectangle ([SystemAccess.Windowing]::GetWindowRectangle($handle))
    }
    catch {
        $rect = $null
    }

    $clientRect = $null
    try {
        $clientRect = ConvertFrom-SystemAccessRectangle -Rectangle ([SystemAccess.Windowing]::GetClientRectangleOnScreen($handle))
    }
    catch {
        $clientRect = $null
    }

    $processId = [SystemAccess.Windowing]::GetWindowProcessId($handle)
    $processName = $null
    if ($processId -gt 0) {
        try {
            $processName = (Get-Process -Id $processId -ErrorAction Stop).ProcessName
        }
        catch {
            $processName = $null
        }
    }

    $cursorPosition = Get-SystemAccessCursorPosition
    $cursor = [pscustomobject]@{
        x = $cursorPosition[0]
        y = $cursorPosition[1]
    }

    [pscustomobject]@{
        windowHandle = $WindowHandle
        title = [SystemAccess.Windowing]::GetWindowTitle($handle)
        className = [SystemAccess.Windowing]::GetWindowClassName($handle)
        processId = $processId
        processName = $processName
        rect = $rect
        clientRect = $clientRect
        isVisible = [SystemAccess.Windowing]::IsWindowShown($handle)
        isForeground = ($WindowHandle -eq [SystemAccess.Windowing]::GetForegroundWindowHandle())
        containsCursor = (Test-SystemAccessPointInRectangle -Point $cursor -Rectangle $rect)
    }
}

function Get-SystemAccessForegroundWindow {
    $handle = [SystemAccess.Windowing]::GetForegroundWindowHandle()
    Get-SystemAccessWindowInfo -WindowHandle $handle
}

function Get-SystemAccessWindowFromPoint {
    param(
        [Nullable[int]] $X = $null,

        [Nullable[int]] $Y = $null
    )

    if ($null -eq $X -or $null -eq $Y) {
        $cursor = Get-SystemAccessCursorState
        $X = $cursor.position.x
        $Y = $cursor.position.y
    }

    $hitHandle = [SystemAccess.Windowing]::GetWindowFromPointHandle([int]$X, [int]$Y)
    $rootHandle = 0
    if ($hitHandle -ne 0) {
        $rootHandle = [SystemAccess.Windowing]::GetRootWindowHandle([IntPtr]::new([Int64]$hitHandle))
    }

    [pscustomobject]@{
        point = [pscustomobject]@{
            x = [int]$X
            y = [int]$Y
        }
        hitWindow = (Get-SystemAccessWindowInfo -WindowHandle ([Int64]$hitHandle))
        rootWindow = (Get-SystemAccessWindowInfo -WindowHandle ([Int64]$rootHandle))
        timestamp = (Get-Date).ToUniversalTime().ToString("o")
    }
}

function Get-SystemAccessWindowHover {
    $cursor = Get-SystemAccessCursorState
    $hit = Get-SystemAccessWindowFromPoint -X $cursor.position.x -Y $cursor.position.y

    [pscustomobject]@{
        cursor = $cursor
        point = $hit.point
        hitWindow = $hit.hitWindow
        rootWindow = $hit.rootWindow
        timestamp = $hit.timestamp
    }
}

function Get-SystemAccessWindowList {
    param(
        [ValidateRange(1, 500)]
        [int] $Limit = 100,

        [bool] $IncludeUntitled = $false
    )

    $windows = New-Object System.Collections.Generic.List[object]
    $handles = [SystemAccess.Windowing]::GetTopLevelWindowHandles()
    foreach ($handleValue in $handles) {
        if ($windows.Count -ge $Limit) {
            break
        }

        $info = Get-SystemAccessWindowInfo -WindowHandle ([Int64]$handleValue)
        if ($null -eq $info -or -not $info.isVisible) {
            continue
        }
        if ($null -eq $info.rect -or $info.rect.width -le 0 -or $info.rect.height -le 0) {
            continue
        }
        if (-not $IncludeUntitled -and [string]::IsNullOrWhiteSpace($info.title)) {
            continue
        }

        $windows.Add($info) | Out-Null
    }

    @($windows.ToArray())
}

function Get-SystemAccessScreenState {
    param(
        [bool] $IncludeWindows = $false,

        [ValidateRange(1, 500)]
        [int] $WindowLimit = 100
    )

    $state = [ordered]@{
        virtualScreen = Get-SystemAccessVirtualScreen
        cursor = Get-SystemAccessCursorState
        foregroundWindow = Get-SystemAccessForegroundWindow
        hoverWindow = Get-SystemAccessWindowHover
        timestamp = (Get-Date).ToUniversalTime().ToString("o")
    }

    if ($IncludeWindows) {
        $state["windows"] = @(Get-SystemAccessWindowList -Limit $WindowLimit)
    }

    [pscustomobject]$state
}

function Get-SystemAccessScreenshotBytes {
    $bounds = [System.Windows.Forms.SystemInformation]::VirtualScreen

    $bitmap = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $stream = New-Object System.IO.MemoryStream

    try {
        $graphics.CopyFromScreen($bounds.Left, $bounds.Top, 0, 0, $bounds.Size)
        $bitmap.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
        return $stream.ToArray()
    }
    finally {
        $stream.Dispose()
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

function Get-SystemAccessScreenshot {
    $bytes = Get-SystemAccessScreenshotBytes
    $screen = Get-SystemAccessVirtualScreen

    [pscustomobject]@{
        mimeType = "image/png"
        data = [Convert]::ToBase64String($bytes)
        width = $screen.width
        height = $screen.height
        left = $screen.left
        top = $screen.top
        timestamp = (Get-Date).ToUniversalTime().ToString("o")
    }
}

function Invoke-SystemAccessMouseMove {
    param(
        [Parameter(Mandatory = $true)]
        [int] $X,

        [Parameter(Mandatory = $true)]
        [int] $Y,

        [bool] $Relative = $false
    )

    if ($Relative) {
        [SystemAccess.NativeInput]::RelativeMouseMove($X, $Y)
    }
    else {
        [SystemAccess.NativeInput]::SetCursorPos($X, $Y) | Out-Null
    }

    [pscustomobject]@{
        ok = $true
        action = "mouse_move"
        x = $X
        y = $Y
        relative = $Relative
    }
}

function Invoke-SystemAccessMouseClick {
    param(
        [ValidateSet("left", "right", "middle")]
        [string] $Button = "left",

        [Nullable[int]] $X = $null,

        [Nullable[int]] $Y = $null,

        [ValidateRange(1, 10)]
        [int] $Clicks = 1
    )

    if ($null -ne $X -and $null -ne $Y) {
        [SystemAccess.NativeInput]::SetCursorPos([int]$X, [int]$Y) | Out-Null
    }

    for ($i = 0; $i -lt $Clicks; $i++) {
        [SystemAccess.NativeInput]::MouseButton($Button, $true)
        Start-Sleep -Milliseconds 35
        [SystemAccess.NativeInput]::MouseButton($Button, $false)

        if ($i -lt ($Clicks - 1)) {
            Start-Sleep -Milliseconds 80
        }
    }

    [pscustomobject]@{
        ok = $true
        action = "mouse_click"
        button = $Button
        x = $X
        y = $Y
        clicks = $Clicks
    }
}

function Invoke-SystemAccessMouseWheel {
    param(
        [Parameter(Mandatory = $true)]
        [int] $Delta
    )

    [SystemAccess.NativeInput]::MouseWheel($Delta)

    [pscustomobject]@{
        ok = $true
        action = "mouse_wheel"
        delta = $Delta
    }
}

function Invoke-SystemAccessKeyboardType {
    param(
        [AllowEmptyString()]
        [string] $Text = ""
    )

    [SystemAccess.NativeInput]::TypeText($Text)

    [pscustomobject]@{
        ok = $true
        action = "keyboard_type"
        length = $Text.Length
    }
}

function Invoke-SystemAccessKeyboardKey {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 255)]
        [int] $VirtualKey,

        [ValidateSet("press", "down", "up")]
        [string] $Action = "press"
    )

    switch ($Action) {
        "down" {
            [SystemAccess.NativeInput]::KeyVirtual([uint16]$VirtualKey, $true)
        }
        "up" {
            [SystemAccess.NativeInput]::KeyVirtual([uint16]$VirtualKey, $false)
        }
        default {
            [SystemAccess.NativeInput]::KeyVirtual([uint16]$VirtualKey, $true)
            Start-Sleep -Milliseconds 35
            [SystemAccess.NativeInput]::KeyVirtual([uint16]$VirtualKey, $false)
        }
    }

    [pscustomobject]@{
        ok = $true
        action = "keyboard_key"
        virtualKey = $VirtualKey
        keyAction = $Action
    }
}
