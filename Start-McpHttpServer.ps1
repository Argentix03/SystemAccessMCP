param(
    [ValidateNotNullOrEmpty()]
    [string] $ListenAddress = "127.0.0.1",

    [ValidateRange(1, 65535)]
    [int] $Port = 8766,

    [ValidatePattern("^/")]
    [string] $Path = "/mcp",

    [ValidateSet("All", "GuestDesktop", "HostHyperV")]
    [string] $Profile = "All"
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$previousImportOnly = $env:SYSTEM_ACCESS_MCP_IMPORT_ONLY
$env:SYSTEM_ACCESS_MCP_IMPORT_ONLY = "1"
try {
    . "$PSScriptRoot\Start-McpServer.ps1" -Profile $Profile
}
finally {
    if ($null -eq $previousImportOnly) {
        Remove-Item Env:\SYSTEM_ACCESS_MCP_IMPORT_ONLY -ErrorAction SilentlyContinue
    }
    else {
        $env:SYSTEM_ACCESS_MCP_IMPORT_ONLY = $previousImportOnly
    }
}

function Read-McpHttpBody {
    param(
        [Parameter(Mandatory = $true)]
        [System.Net.HttpListenerRequest] $Request
    )

    if (-not $Request.HasEntityBody) {
        throw "Request body is required."
    }

    $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
    try {
        $text = $reader.ReadToEnd()
        if ([string]::IsNullOrWhiteSpace($text)) {
            throw "Request body is required."
        }
        return ($text | ConvertFrom-Json)
    }
    finally {
        $reader.Dispose()
    }
}

function Write-McpHttpResponse {
    param(
        [Parameter(Mandatory = $true)]
        [System.Net.HttpListenerResponse] $Response,

        [int] $StatusCode = 200,

        [string] $ContentType = "application/json; charset=utf-8",

        [byte[]] $Bytes = ([byte[]]@())
    )

    $Response.StatusCode = $StatusCode
    $Response.ContentType = $ContentType
    $Response.Headers["Cache-Control"] = "no-store"
    $Response.Headers["Access-Control-Allow-Origin"] = "http://127.0.0.1:$Port"
    $Response.Headers["Access-Control-Allow-Methods"] = "POST,OPTIONS"
    $Response.Headers["Access-Control-Allow-Headers"] = "content-type,accept,mcp-session-id"
    $Response.ContentLength64 = $Bytes.LongLength
    $Response.OutputStream.Write($Bytes, 0, $Bytes.Length)
    $Response.OutputStream.Close()
}

function Write-McpHttpJson {
    param(
        [Parameter(Mandatory = $true)]
        [System.Net.HttpListenerResponse] $Response,

        [Parameter(Mandatory = $true)]
        [object] $Value,

        [int] $StatusCode = 200
    )

    $json = $Value | ConvertTo-Json -Depth 20 -Compress
    Write-McpHttpResponse -Response $Response -StatusCode $StatusCode -Bytes ([Text.Encoding]::UTF8.GetBytes($json))
}

function Invoke-McpHttpContext {
    param(
        [Parameter(Mandatory = $true)]
        [System.Net.HttpListenerContext] $Context
    )

    $request = $Context.Request
    $response = $Context.Response
    $requestPath = $request.Url.AbsolutePath.TrimEnd("/")
    $mcpPath = $Path.TrimEnd("/")

    if ([string]::IsNullOrWhiteSpace($requestPath)) {
        $requestPath = "/"
    }
    if ([string]::IsNullOrWhiteSpace($mcpPath)) {
        $mcpPath = "/"
    }

    try {
        if ($request.HttpMethod -eq "OPTIONS") {
            Write-McpHttpResponse -Response $response -StatusCode 204 -ContentType "text/plain"
            return
        }

        if ($requestPath -eq "/health" -and $request.HttpMethod -eq "GET") {
            Write-McpHttpJson -Response $response -Value ([pscustomobject]@{
                ok = $true
                transport = "http"
                path = $mcpPath
                profile = $Profile
                timestamp = (Get-Date).ToUniversalTime().ToString("o")
            })
            return
        }

        if ($requestPath -ne $mcpPath) {
            Write-McpHttpJson -Response $response -StatusCode 404 -Value (New-McpErrorResponse -Id $null -Code -32601 -Message "Not found: $requestPath")
            return
        }

        if ($request.HttpMethod -ne "POST") {
            Write-McpHttpJson -Response $response -StatusCode 405 -Value (New-McpErrorResponse -Id $null -Code -32600 -Message "POST required")
            return
        }

        $body = Read-McpHttpBody -Request $request
        if ($body -is [array]) {
            $results = @()
            foreach ($item in $body) {
                $itemResponse = Invoke-McpRequest -Request $item
                if ($null -ne $itemResponse) {
                    $results += $itemResponse
                }
            }

            if ($results.Count -eq 0) {
                Write-McpHttpResponse -Response $response -StatusCode 202 -ContentType "text/plain"
            }
            else {
                Write-McpHttpJson -Response $response -Value $results
            }
            return
        }

        $mcpResponse = Invoke-McpRequest -Request $body
        if ($null -eq $mcpResponse) {
            Write-McpHttpResponse -Response $response -StatusCode 202 -ContentType "text/plain"
        }
        else {
            Write-McpHttpJson -Response $response -Value $mcpResponse
        }
    }
    catch {
        $responseId = $null
        try {
            if ($null -ne $body) {
                $responseId = Get-McpArgument -Arguments $body -Name "id"
            }
        }
        catch {
        }
        Write-McpHttpJson -Response $response -StatusCode 400 -Value (New-McpErrorResponse -Id $responseId -Message $_.Exception.Message)
    }
}

$listener = New-Object System.Net.HttpListener
$prefixHost = $ListenAddress
if ($ListenAddress -eq "0.0.0.0") {
    $prefixHost = "+"
}
$prefix = "http://$prefixHost`:$Port/"
$listener.Prefixes.Add($prefix)
$listener.Start()

Write-Host "SystemAccessMCP HTTP MCP server listening on http://$ListenAddress`:$Port$Path with profile $Profile"
Write-Host "Press Ctrl+C to stop."

$script:StopRequested = $false
$cancelHandlerRegistered = $false
$cancelHandler = [ConsoleCancelEventHandler] {
    param($Sender, $EventArgs)

    $script:StopRequested = $true
    $EventArgs.Cancel = $true
    try {
        if ($listener.IsListening) {
            $listener.Stop()
        }
    }
    catch {
    }
}

try {
    [Console]::add_CancelKeyPress($cancelHandler)
    $cancelHandlerRegistered = $true
}
catch {
    Write-Warning "Ctrl+C handler registration failed: $($_.Exception.Message)"
}

try {
    while ($listener.IsListening -and -not $script:StopRequested) {
        $contextTask = $listener.GetContextAsync()
        while (-not $contextTask.AsyncWaitHandle.WaitOne(200)) {
            if ($script:StopRequested -or -not $listener.IsListening) {
                break
            }
        }

        if ($script:StopRequested -or -not $listener.IsListening) {
            break
        }

        try {
            Invoke-McpHttpContext -Context $contextTask.GetAwaiter().GetResult()
        }
        catch [System.Net.HttpListenerException] {
            if (-not $script:StopRequested) {
                throw
            }
        }
        catch [System.ObjectDisposedException] {
            if (-not $script:StopRequested) {
                throw
            }
        }
    }
}
finally {
    if ($cancelHandlerRegistered) {
        try {
            [Console]::remove_CancelKeyPress($cancelHandler)
        }
        catch {
        }
    }
    try {
        if ($listener.IsListening) {
            $listener.Stop()
        }
    }
    catch {
    }
    $listener.Close()
    Write-Host "SystemAccessMCP HTTP MCP server stopped."
}
