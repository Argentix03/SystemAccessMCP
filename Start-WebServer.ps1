param(
    [ValidateNotNullOrEmpty()]
    [string] $ListenAddress = "127.0.0.1",

    [ValidateRange(1, 65535)]
    [int] $Port = 8765,

    [ValidateSet("All", "GuestDesktop", "HostHyperV")]
    [string] $Profile = "All",

    [AllowEmptyString()]
    [string] $HyperVApiBaseUrl = ""
)

Set-StrictMode -Version 2.0

. "$PSScriptRoot\src\SystemAccess.Core.ps1"
. "$PSScriptRoot\src\SystemAccess.HyperV.ps1"

$script:WebProfile = $Profile
$script:HyperVApiBaseUrl = $HyperVApiBaseUrl.TrimEnd("/")

function Read-RequestBody {
    param(
        [Parameter(Mandatory = $true)]
        [System.Net.HttpListenerRequest] $Request
    )

    if (-not $Request.HasEntityBody) {
        return $null
    }

    $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
    try {
        $text = $reader.ReadToEnd()
        if ([string]::IsNullOrWhiteSpace($text)) {
            return $null
        }

        return ($text | ConvertFrom-Json)
    }
    finally {
        $reader.Dispose()
    }
}

function Get-JsonValue {
    param(
        [object] $Object,
        [Parameter(Mandatory = $true)]
        [string] $Name,
        [object] $Default = $null
    )

    if ($null -ne $Object) {
        $matches = $Object.PSObject.Properties.Match($Name)
        if ($matches.Count -gt 0) {
            return $matches[0].Value
        }
    }

    return $Default
}

function Assert-JsonValue {
    param(
        [object] $Object,
        [Parameter(Mandatory = $true)]
        [string] $Name
    )

    $value = Get-JsonValue -Object $Object -Name $Name
    if ($null -eq $value) {
        throw "$Name is required"
    }

    return $value
}

function Get-QueryValue {
    param(
        [Parameter(Mandatory = $true)]
        [System.Net.HttpListenerRequest] $Request,
        [Parameter(Mandatory = $true)]
        [string] $Name,
        [object] $Default = $null
    )

    $value = $Request.QueryString[$Name]
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $Default
    }

    return $value
}

function Write-Response {
    param(
        [Parameter(Mandatory = $true)]
        [System.Net.HttpListenerResponse] $Response,

        [int] $StatusCode = 200,

        [string] $ContentType = "application/json; charset=utf-8",

        [byte[]] $Bytes
    )

    $Response.StatusCode = $StatusCode
    $Response.ContentType = $ContentType
    $Response.Headers["Cache-Control"] = "no-store"
    $Response.Headers["Access-Control-Allow-Origin"] = "http://127.0.0.1:$Port"

    if ($null -eq $Bytes) {
        $Bytes = [byte[]]@()
    }

    $Response.ContentLength64 = $Bytes.LongLength
    $Response.OutputStream.Write($Bytes, 0, $Bytes.Length)
    $Response.OutputStream.Close()
}

function Write-Json {
    param(
        [Parameter(Mandatory = $true)]
        [System.Net.HttpListenerResponse] $Response,

        [object] $Value,

        [int] $StatusCode = 200
    )

    $json = ConvertTo-SystemAccessJson -Value $Value
    Write-Response -Response $Response -StatusCode $StatusCode -ContentType "application/json; charset=utf-8" -Bytes ([Text.Encoding]::UTF8.GetBytes($json))
}

function Write-ErrorJson {
    param(
        [Parameter(Mandatory = $true)]
        [System.Net.HttpListenerResponse] $Response,

        [int] $StatusCode,

        [string] $Message
    )

    try {
        Write-Json -Response $Response -StatusCode $StatusCode -Value ([pscustomobject]@{
            ok = $false
            error = $Message
        })
    }
    catch {
        try {
            $Response.Abort()
        }
        catch {
        }
    }
}

function Test-SystemAccessWebHyperVPath {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    return ($Path -like "/api/hyperv*" -or $Path -eq "/hyperv/screenshot")
}

function Test-SystemAccessWebDesktopPath {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    return (
        $Path -eq "/screenshot" -or
        $Path -eq "/api/screenshot" -or
        $Path -like "/api/mouse/*" -or
        $Path -like "/api/keyboard/*"
    )
}

function Test-SystemAccessWebObservationPath {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    return (
        $Path -like "/api/cursor/*" -or
        $Path -like "/api/window/*" -or
        $Path -like "/api/screen/*"
    )
}

function Test-SystemAccessWebEndpointAllowed {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    if ($Path -eq "/" -or $Path -eq "/health") {
        return $true
    }

    if (Test-SystemAccessWebObservationPath -Path $Path) {
        return $true
    }

    if (Test-SystemAccessWebDesktopPath -Path $Path) {
        return ($script:WebProfile -eq "All" -or $script:WebProfile -eq "GuestDesktop")
    }

    if (Test-SystemAccessWebHyperVPath -Path $Path) {
        if (-not [string]::IsNullOrWhiteSpace($script:HyperVApiBaseUrl)) {
            return $true
        }
        return ($script:WebProfile -eq "All" -or $script:WebProfile -eq "HostHyperV")
    }

    return $true
}

function Get-SystemAccessWebHealth {
    $status = Get-SystemAccessStatus

    [pscustomobject]@{
        ok = $true
        provider = $status.provider
        providerScope = $status.providerScope
        profile = $script:WebProfile
        hyperVApiBaseUrl = $script:HyperVApiBaseUrl
        machineName = $status.machineName
        userName = $status.userName
        virtualScreen = $status.virtualScreen
        timestamp = (Get-Date).ToUniversalTime().ToString("o")
    }
}

function Invoke-SystemAccessWebProxyContext {
    param(
        [Parameter(Mandatory = $true)]
        [System.Net.HttpListenerContext] $Context,

        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    if ([string]::IsNullOrWhiteSpace($script:HyperVApiBaseUrl)) {
        throw "Hyper-V API base URL is not configured."
    }

    $request = $Context.Request
    $response = $Context.Response
    $targetUri = $script:HyperVApiBaseUrl + $Path + $request.Url.Query
    $proxyRequest = [System.Net.HttpWebRequest]::Create($targetUri)
    $proxyRequest.Method = $request.HttpMethod
    $proxyRequest.Timeout = 30000
    $proxyRequest.ReadWriteTimeout = 30000
    if (-not [string]::IsNullOrWhiteSpace($request.ContentType)) {
        $proxyRequest.ContentType = $request.ContentType
    }
    if ($null -ne $request.AcceptTypes -and $request.AcceptTypes.Count -gt 0) {
        $proxyRequest.Accept = ($request.AcceptTypes -join ",")
    }

    if ($request.HasEntityBody) {
        $proxyRequestStream = $proxyRequest.GetRequestStream()
        try {
            $request.InputStream.CopyTo($proxyRequestStream)
        }
        finally {
            $proxyRequestStream.Dispose()
        }
    }

    $proxyResponse = $null
    try {
        $proxyResponse = $proxyRequest.GetResponse()
    }
    catch [System.Net.WebException] {
        if ($null -eq $_.Exception.Response) {
            throw
        }
        $proxyResponse = $_.Exception.Response
    }

    $stream = $proxyResponse.GetResponseStream()
    $memory = New-Object System.IO.MemoryStream
    try {
        $stream.CopyTo($memory)
        $statusCode = [int]$proxyResponse.StatusCode
        $contentType = $proxyResponse.ContentType
        if ([string]::IsNullOrWhiteSpace($contentType)) {
            $contentType = "application/octet-stream"
        }
        Write-Response -Response $response -StatusCode $statusCode -ContentType $contentType -Bytes $memory.ToArray()
    }
    finally {
        $memory.Dispose()
        $stream.Dispose()
        $proxyResponse.Dispose()
    }
}

function Get-IndexHtml {
    $html = @'
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>SystemAccessMCP</title>
  <style>
    :root {
      color-scheme: light dark;
      --bg: #f6f7f9;
      --panel: #ffffff;
      --text: #17202a;
      --muted: #5c6670;
      --border: #d8dde4;
      --accent: #1264a3;
      --accent-text: #ffffff;
      --danger: #a32020;
      font-family: Segoe UI, system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
    }
    @media (prefers-color-scheme: dark) {
      :root {
        --bg: #121416;
        --panel: #1d2126;
        --text: #f2f4f6;
        --muted: #aab2bb;
        --border: #343b44;
        --accent: #4ea1d3;
        --accent-text: #061018;
      }
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: var(--bg);
      color: var(--text);
      min-height: 100vh;
    }
    header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      padding: 14px 18px;
      border-bottom: 1px solid var(--border);
      background: var(--panel);
    }
    h1 {
      margin: 0;
      font-size: 18px;
      line-height: 1.2;
      font-weight: 650;
    }
    main {
      display: grid;
      grid-template-columns: minmax(320px, 1fr) 360px;
      gap: 16px;
      padding: 16px;
      align-items: start;
    }
    .viewer, .controls {
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 8px;
      overflow: hidden;
    }
    .viewer-toolbar, .group {
      border-bottom: 1px solid var(--border);
      padding: 12px;
    }
    .group:last-child { border-bottom: 0; }
    .viewer-toolbar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
    }
    #screen {
      display: block;
      width: 100%;
      min-height: 240px;
      background: #000;
      object-fit: contain;
    }
    h2 {
      margin: 0 0 10px;
      font-size: 14px;
      line-height: 1.2;
    }
    label {
      display: block;
      color: var(--muted);
      font-size: 12px;
      margin-bottom: 4px;
    }
    input, select, textarea, button {
      font: inherit;
      border-radius: 6px;
    }
    input, select, textarea {
      width: 100%;
      color: var(--text);
      background: transparent;
      border: 1px solid var(--border);
      padding: 8px 9px;
    }
    textarea { resize: vertical; min-height: 72px; }
    button {
      border: 1px solid var(--accent);
      background: var(--accent);
      color: var(--accent-text);
      padding: 8px 10px;
      cursor: pointer;
      min-height: 36px;
    }
    button.secondary {
      background: transparent;
      color: var(--text);
      border-color: var(--border);
    }
    button:disabled { opacity: .6; cursor: not-allowed; }
    .row {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
      margin-bottom: 8px;
    }
    .row.three { grid-template-columns: 1fr 1fr 1fr; }
    .actions {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
    }
    .status {
      color: var(--muted);
      font-size: 12px;
      line-height: 1.4;
      overflow-wrap: anywhere;
    }
    .warning {
      color: var(--danger);
      font-size: 12px;
      line-height: 1.4;
    }
    pre.status {
      margin: 8px 0 0;
      max-height: 160px;
      overflow: auto;
      white-space: pre-wrap;
    }
    @media (max-width: 900px) {
      main { grid-template-columns: 1fr; }
      .controls { order: -1; }
    }
  </style>
</head>
<body>
  <header>
    <h1>SystemAccessMCP</h1>
    <div id="status" class="status">Connecting</div>
  </header>
  <main>
    <section class="viewer">
      <div class="viewer-toolbar">
        <div class="status" id="screenMeta">No screenshot loaded</div>
        <div class="actions">
          <button class="secondary" id="refresh">Refresh</button>
        </div>
      </div>
      <img id="screen" alt="Current desktop screenshot">
    </section>
    <aside class="controls">
      <div class="group">
        <h2>Target</h2>
        <label for="target">Control Target</label>
        <select id="target">
          <option value="desktop">Current Desktop</option>
          <option value="hyperv">Hyper-V Console</option>
        </select>
        <div class="status" id="profileMeta" style="margin-top: 8px"></div>
      </div>
      <div class="group">
        <h2>Hyper-V</h2>
        <label for="hypervApiBase">Host API Base URL</label>
        <input id="hypervApiBase" readonly>
        <label for="server">Server</label>
        <input id="server" value="localhost">
        <label for="vmName" style="margin-top: 8px">VM Name</label>
        <input id="vmName" placeholder="Example: Windows 11 Dev">
        <div class="actions" style="margin-top: 8px">
          <button id="listVms" class="secondary">List VMs</button>
          <button id="startVm" class="secondary">Start</button>
          <button id="connectVm">Connect</button>
        </div>
        <div class="actions" style="margin-top: 8px">
          <button id="refreshVm" class="secondary">Refresh VM</button>
          <button id="ctrlAltDelete" class="secondary">Ctrl+Alt+Del</button>
        </div>
        <pre id="vmList" class="status"></pre>
      </div>
      <div class="group">
        <h2>Mouse</h2>
        <div class="row">
          <div><label for="x">X</label><input id="x" type="number" value="100"></div>
          <div><label for="y">Y</label><input id="y" type="number" value="100"></div>
        </div>
        <div class="row">
          <div><label for="button">Button</label><select id="button"><option>left</option><option>right</option><option>middle</option></select></div>
          <div><label for="clicks">Clicks</label><input id="clicks" type="number" min="1" max="10" value="1"></div>
        </div>
        <div class="actions">
          <button id="move">Move</button>
          <button id="click">Click</button>
        </div>
      </div>
      <div class="group">
        <h2>Observation</h2>
        <div class="row">
          <div><label for="obsWindowLimit">Window Limit</label><input id="obsWindowLimit" type="number" min="1" max="500" value="25"></div>
          <div><label for="obsIncludeUntitled">Untitled</label><input id="obsIncludeUntitled" type="checkbox"></div>
        </div>
        <div class="actions">
          <button id="obsCursor" class="secondary">Cursor</button>
          <button id="obsForeground" class="secondary">Foreground</button>
          <button id="obsHover" class="secondary">Hover</button>
          <button id="obsPoint" class="secondary">Point</button>
          <button id="obsWindows" class="secondary">Windows</button>
          <button id="obsScreen" class="secondary">Screen</button>
        </div>
        <pre id="observationOutput" class="status"></pre>
      </div>
      <div class="group">
        <h2>Keyboard</h2>
        <label for="text">Text</label>
        <textarea id="text"></textarea>
        <div class="actions" style="margin-top: 8px">
          <button id="type">Type</button>
        </div>
        <div class="row" style="margin-top: 12px">
          <div><label for="vk">Virtual Key</label><input id="vk" type="number" min="1" max="255" value="13"></div>
          <div><label for="keyAction">Action</label><select id="keyAction"><option>press</option><option>down</option><option>up</option></select></div>
        </div>
        <div class="actions">
          <button id="key">Send Key</button>
        </div>
      </div>
      <div class="group">
        <h2>Boundary</h2>
        <div class="warning">Current Desktop controls the signed-in user session. Hyper-V Console controls a VMConnect window on the host and can reach guest login and UAC surfaces while the host desktop is unlocked.</div>
      </div>
    </aside>
  </main>
  <script>
    const $ = (id) => document.getElementById(id);
    const setStatus = (text) => { $('status').textContent = text; };
    const webProfile = '__SYSTEM_ACCESS_WEB_PROFILE__';
    const configuredHyperVApiBaseUrl = '__SYSTEM_ACCESS_HYPERV_API_BASE_URL__';
    const target = () => $('target').value;
    const server = () => $('server').value || 'localhost';
    const vmName = () => $('vmName').value.trim();
    const vmQuery = () => `server=${encodeURIComponent(server())}&vmName=${encodeURIComponent(vmName())}&area=window`;
    const withTarget = (body) => target() === 'hyperv'
      ? { ...body, server: server(), vmName: vmName(), area: 'window' }
      : body;
    const route = (desktopPath, hypervPath) => target() === 'hyperv' ? hypervPath : desktopPath;
    if (webProfile === 'HostHyperV') {
      $('target').value = 'hyperv';
      $('target').querySelector('option[value="desktop"]').disabled = true;
    }
    async function json(path, body) {
      const res = await fetch(path, {
        method: body ? 'POST' : 'GET',
        headers: body ? { 'content-type': 'application/json' } : undefined,
        body: body ? JSON.stringify(body) : undefined
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || res.statusText);
      return data;
    }
    async function loadStatus() {
      const data = await json('/health');
      const remote = data.hyperVApiBaseUrl ? ` | Hyper-V proxy ${data.hyperVApiBaseUrl}` : '';
      setStatus(`${data.provider} | ${data.profile} | ${data.userName} | ${data.virtualScreen.width}x${data.virtualScreen.height}${remote}`);
      $('profileMeta').textContent = `${data.profile}${data.hyperVApiBaseUrl ? ' | Hyper-V via host API proxy' : ''}`;
      $('hypervApiBase').value = data.hyperVApiBaseUrl || 'local';
    }
    async function refresh() {
      const img = $('screen');
      let data;
      if (target() === 'hyperv') {
        if (!vmName()) throw new Error('VM name is required');
        img.src = `/hyperv/screenshot?${vmQuery()}&t=${Date.now()}`;
        data = await json(`/api/hyperv/console/screenshot?${vmQuery()}`);
      } else {
        img.src = `/screenshot?t=${Date.now()}`;
        data = await json('/api/screenshot');
      }
      $('screenMeta').textContent = `${data.width}x${data.height} captured ${new Date(data.timestamp).toLocaleTimeString()}`;
    }
    async function run(label, fn) {
      setStatus(`${label}...`);
      try {
        await fn();
        await loadStatus();
      } catch (err) {
        setStatus(err.message);
      }
    }
    function showObservation(data) {
      $('observationOutput').textContent = JSON.stringify(data, null, 2);
    }
    async function observe(label, path) {
      await run(label, async () => {
        showObservation(await json(path));
      });
    }
    $('refresh').onclick = () => run('Refreshing', refresh);
    $('move').onclick = () => run('Moving mouse', () => json(route('/api/mouse/move', '/api/hyperv/console/mouse/move'), withTarget({ x: Number($('x').value), y: Number($('y').value) })));
    $('click').onclick = () => run('Clicking mouse', () => json(route('/api/mouse/click', '/api/hyperv/console/mouse/click'), withTarget({ x: Number($('x').value), y: Number($('y').value), button: $('button').value, clicks: Number($('clicks').value) })));
    $('obsCursor').onclick = () => observe('Reading cursor', '/api/cursor/state');
    $('obsForeground').onclick = () => observe('Reading foreground window', '/api/window/foreground');
    $('obsHover').onclick = () => observe('Reading hover window', '/api/window/hover');
    $('obsPoint').onclick = () => observe('Reading point window', `/api/window/from-point?x=${encodeURIComponent(Number($('x').value))}&y=${encodeURIComponent(Number($('y').value))}`);
    $('obsWindows').onclick = () => observe('Listing windows', `/api/window/list?limit=${encodeURIComponent(Number($('obsWindowLimit').value) || 25)}&includeUntitled=${$('obsIncludeUntitled').checked}`);
    $('obsScreen').onclick = () => observe('Reading screen state', `/api/screen/state?includeWindows=true&windowLimit=${encodeURIComponent(Number($('obsWindowLimit').value) || 25)}`);
    $('type').onclick = () => run('Typing', () => json(route('/api/keyboard/type', '/api/hyperv/console/keyboard/type'), withTarget({ text: $('text').value })));
    $('key').onclick = () => run('Sending key', () => json(route('/api/keyboard/key', '/api/hyperv/console/keyboard/key'), withTarget({ virtualKey: Number($('vk').value), action: $('keyAction').value })));
    $('listVms').onclick = () => run('Listing VMs', async () => {
      const data = await json(`/api/hyperv/vms?server=${encodeURIComponent(server())}`);
      $('vmList').textContent = data.map(vm => `${vm.name} | ${vm.state}`).join('\n') || 'No VMs returned';
      if (!vmName() && data.length) $('vmName').value = data[0].name;
    });
    $('startVm').onclick = () => run('Starting VM', () => json('/api/hyperv/vm/start', { server: server(), vmName: vmName() }));
    $('connectVm').onclick = () => run('Connecting VM', () => json('/api/hyperv/console/connect', { server: server(), vmName: vmName() }));
    $('refreshVm').onclick = () => run('Refreshing VM', async () => { $('target').value = 'hyperv'; await refresh(); });
    $('ctrlAltDelete').onclick = () => run('Sending Ctrl+Alt+Del', () => json('/api/hyperv/console/ctrl-alt-delete', { server: server(), vmName: vmName() }));
    loadStatus().then(refresh).catch(err => setStatus(err.message));
  </script>
</body>
</html>
'@

    $escapedProfile = $script:WebProfile.Replace("\", "\\").Replace("'", "\'")
    $escapedHyperVApiBaseUrl = $script:HyperVApiBaseUrl.Replace("\", "\\").Replace("'", "\'")
    return $html.Replace("__SYSTEM_ACCESS_WEB_PROFILE__", $escapedProfile).Replace("__SYSTEM_ACCESS_HYPERV_API_BASE_URL__", $escapedHyperVApiBaseUrl)
}

if ($env:SYSTEM_ACCESS_WEB_IMPORT_ONLY -eq "1") {
    return
}

$listener = New-Object System.Net.HttpListener
$prefixHost = $ListenAddress
if ($ListenAddress -eq "0.0.0.0") {
    $prefixHost = "+"
}
$prefix = "http://$prefixHost`:$Port/"
$listener.Prefixes.Add($prefix)
$listener.Start()

Write-Host "SystemAccessMCP web server listening on http://$ListenAddress`:$Port/ with profile $Profile"
if (-not [string]::IsNullOrWhiteSpace($script:HyperVApiBaseUrl)) {
    Write-Host "Hyper-V API requests will proxy to $script:HyperVApiBaseUrl"
}
Write-Host "Press Ctrl+C to stop."

function Invoke-SystemAccessWebContext {
    param(
        [Parameter(Mandatory = $true)]
        [System.Net.HttpListenerContext] $Context
    )

    $request = $Context.Request
    $response = $Context.Response
    $path = $request.Url.AbsolutePath.TrimEnd("/")
    if ($path -eq "") {
        $path = "/"
    }

    try {
        if ($request.HttpMethod -eq "OPTIONS") {
            $response.Headers["Access-Control-Allow-Origin"] = "http://127.0.0.1:$Port"
            $response.Headers["Access-Control-Allow-Methods"] = "GET,POST,OPTIONS"
            $response.Headers["Access-Control-Allow-Headers"] = "content-type"
            Write-Response -Response $response -StatusCode 204 -ContentType "text/plain" -Bytes ([byte[]]@())
            return
        }

        if (-not (Test-SystemAccessWebEndpointAllowed -Path $path)) {
            Write-ErrorJson -Response $response -StatusCode 404 -Message "Endpoint '$path' is not available in web profile '$script:WebProfile'."
            return
        }

        if ((Test-SystemAccessWebHyperVPath -Path $path) -and -not [string]::IsNullOrWhiteSpace($script:HyperVApiBaseUrl)) {
            Invoke-SystemAccessWebProxyContext -Context $Context -Path $path
            return
        }

        switch ($path) {
                "/" {
                    $html = Get-IndexHtml
                    Write-Response -Response $response -ContentType "text/html; charset=utf-8" -Bytes ([Text.Encoding]::UTF8.GetBytes($html))
                }
                "/health" {
                    Write-Json -Response $response -Value (Get-SystemAccessWebHealth)
                }
                "/screenshot" {
                    Write-Response -Response $response -ContentType "image/png" -Bytes (Get-SystemAccessScreenshotBytes)
                }
                "/api/screenshot" {
                    Write-Json -Response $response -Value (Get-SystemAccessScreenshot)
                }
                "/api/cursor/state" {
                    Write-Json -Response $response -Value (Get-SystemAccessCursorState)
                }
                "/api/window/foreground" {
                    Write-Json -Response $response -Value (Get-SystemAccessForegroundWindow)
                }
                "/api/window/hover" {
                    Write-Json -Response $response -Value (Get-SystemAccessWindowHover)
                }
                "/api/window/from-point" {
                    $x = Get-QueryValue -Request $request -Name "x"
                    $y = Get-QueryValue -Request $request -Name "y"
                    if ($null -ne $x -and $null -ne $y) {
                        Write-Json -Response $response -Value (Get-SystemAccessWindowFromPoint -X ([int]$x) -Y ([int]$y))
                    }
                    else {
                        Write-Json -Response $response -Value (Get-SystemAccessWindowFromPoint)
                    }
                }
                "/api/window/list" {
                    $limit = [int](Get-QueryValue -Request $request -Name "limit" -Default 100)
                    $includeUntitledText = [string](Get-QueryValue -Request $request -Name "includeUntitled" -Default "false")
                    $includeUntitled = ($includeUntitledText -eq "1" -or $includeUntitledText -eq "true")
                    Write-Json -Response $response -Value @(Get-SystemAccessWindowList -Limit $limit -IncludeUntitled $includeUntitled)
                }
                "/api/screen/state" {
                    $includeWindowsText = [string](Get-QueryValue -Request $request -Name "includeWindows" -Default "false")
                    $includeWindows = ($includeWindowsText -eq "1" -or $includeWindowsText -eq "true")
                    $windowLimit = [int](Get-QueryValue -Request $request -Name "windowLimit" -Default 100)
                    Write-Json -Response $response -Value (Get-SystemAccessScreenState -IncludeWindows $includeWindows -WindowLimit $windowLimit)
                }
                "/api/hyperv/status" {
                    Write-Json -Response $response -Value (Get-SystemAccessHyperVStatus)
                }
                "/api/hyperv/vms" {
                    $server = [string](Get-QueryValue -Request $request -Name "server" -Default "localhost")
                    Write-Json -Response $response -Value @(Get-SystemAccessHyperVVMs -Server $server)
                }
                "/api/hyperv/vm/start" {
                    if ($request.HttpMethod -ne "POST") { throw "POST required" }
                    $body = Read-RequestBody -Request $request
                    $vmName = [string](Assert-JsonValue -Object $body -Name "vmName")
                    $server = [string](Get-JsonValue -Object $body -Name "server" -Default "localhost")
                    Write-Json -Response $response -Value (Start-SystemAccessHyperVVM -VMName $vmName -Server $server)
                }
                "/api/hyperv/console/windows" {
                    $vmName = [string](Get-QueryValue -Request $request -Name "vmName" -Default "")
                    $server = [string](Get-QueryValue -Request $request -Name "server" -Default "")
                    Write-Json -Response $response -Value @(Get-SystemAccessHyperVConsoleWindows -VMName $vmName -Server $server)
                }
                "/api/hyperv/console/connect" {
                    if ($request.HttpMethod -ne "POST") { throw "POST required" }
                    $body = Read-RequestBody -Request $request
                    $vmName = [string](Assert-JsonValue -Object $body -Name "vmName")
                    $server = [string](Get-JsonValue -Object $body -Name "server" -Default "localhost")
                    $timeoutMs = [int](Get-JsonValue -Object $body -Name "timeoutMs" -Default 15000)
                    $forceNew = [bool](Get-JsonValue -Object $body -Name "forceNew" -Default $false)
                    Write-Json -Response $response -Value (Open-SystemAccessHyperVConsole -VMName $vmName -Server $server -TimeoutMs $timeoutMs -ForceNew:$forceNew)
                }
                "/api/hyperv/console/pointer-state" {
                    $vmName = [string](Get-QueryValue -Request $request -Name "vmName" -Default "")
                    $server = [string](Get-QueryValue -Request $request -Name "server" -Default "localhost")
                    $processId = [int](Get-QueryValue -Request $request -Name "processId" -Default 0)
                    Write-Json -Response $response -Value (Get-SystemAccessHyperVConsolePointerState -VMName $vmName -Server $server -ProcessId $processId)
                }
                "/hyperv/screenshot" {
                    $vmName = [string](Get-QueryValue -Request $request -Name "vmName" -Default "")
                    $server = [string](Get-QueryValue -Request $request -Name "server" -Default "localhost")
                    $processId = [int](Get-QueryValue -Request $request -Name "processId" -Default 0)
                    $area = [string](Get-QueryValue -Request $request -Name "area" -Default "window")
                    $capture = Get-SystemAccessHyperVConsoleScreenshotBytes -VMName $vmName -Server $server -ProcessId $processId -Area $area
                    Write-Response -Response $response -ContentType "image/png" -Bytes $capture.bytes
                }
                "/api/hyperv/console/screenshot" {
                    $vmName = [string](Get-QueryValue -Request $request -Name "vmName" -Default "")
                    $server = [string](Get-QueryValue -Request $request -Name "server" -Default "localhost")
                    $processId = [int](Get-QueryValue -Request $request -Name "processId" -Default 0)
                    $area = [string](Get-QueryValue -Request $request -Name "area" -Default "window")
                    Write-Json -Response $response -Value (Get-SystemAccessHyperVConsoleScreenshot -VMName $vmName -Server $server -ProcessId $processId -Area $area)
                }
                "/api/hyperv/console/mouse/move" {
                    if ($request.HttpMethod -ne "POST") { throw "POST required" }
                    $body = Read-RequestBody -Request $request
                    $x = [int](Assert-JsonValue -Object $body -Name "x")
                    $y = [int](Assert-JsonValue -Object $body -Name "y")
                    $vmName = [string](Get-JsonValue -Object $body -Name "vmName" -Default "")
                    $server = [string](Get-JsonValue -Object $body -Name "server" -Default "localhost")
                    $processId = [int](Get-JsonValue -Object $body -Name "processId" -Default 0)
                    $area = [string](Get-JsonValue -Object $body -Name "area" -Default "window")
                    Write-Json -Response $response -Value (Invoke-SystemAccessHyperVConsoleMouseMove -X $x -Y $y -VMName $vmName -Server $server -ProcessId $processId -Area $area)
                }
                "/api/hyperv/console/mouse/click" {
                    if ($request.HttpMethod -ne "POST") { throw "POST required" }
                    $body = Read-RequestBody -Request $request
                    $x = [int](Assert-JsonValue -Object $body -Name "x")
                    $y = [int](Assert-JsonValue -Object $body -Name "y")
                    $button = [string](Get-JsonValue -Object $body -Name "button" -Default "left")
                    $clicks = [int](Get-JsonValue -Object $body -Name "clicks" -Default 1)
                    $vmName = [string](Get-JsonValue -Object $body -Name "vmName" -Default "")
                    $server = [string](Get-JsonValue -Object $body -Name "server" -Default "localhost")
                    $processId = [int](Get-JsonValue -Object $body -Name "processId" -Default 0)
                    $area = [string](Get-JsonValue -Object $body -Name "area" -Default "window")
                    Write-Json -Response $response -Value (Invoke-SystemAccessHyperVConsoleMouseClick -X $x -Y $y -Button $button -Clicks $clicks -VMName $vmName -Server $server -ProcessId $processId -Area $area)
                }
                "/api/hyperv/console/keyboard/type" {
                    if ($request.HttpMethod -ne "POST") { throw "POST required" }
                    $body = Read-RequestBody -Request $request
                    $text = [string](Get-JsonValue -Object $body -Name "text" -Default "")
                    $vmName = [string](Get-JsonValue -Object $body -Name "vmName" -Default "")
                    $server = [string](Get-JsonValue -Object $body -Name "server" -Default "localhost")
                    $processId = [int](Get-JsonValue -Object $body -Name "processId" -Default 0)
                    Write-Json -Response $response -Value (Invoke-SystemAccessHyperVConsoleKeyboardType -Text $text -VMName $vmName -Server $server -ProcessId $processId)
                }
                "/api/hyperv/console/keyboard/key" {
                    if ($request.HttpMethod -ne "POST") { throw "POST required" }
                    $body = Read-RequestBody -Request $request
                    $virtualKey = [int](Assert-JsonValue -Object $body -Name "virtualKey")
                    $action = [string](Get-JsonValue -Object $body -Name "action" -Default "press")
                    $vmName = [string](Get-JsonValue -Object $body -Name "vmName" -Default "")
                    $server = [string](Get-JsonValue -Object $body -Name "server" -Default "localhost")
                    $processId = [int](Get-JsonValue -Object $body -Name "processId" -Default 0)
                    Write-Json -Response $response -Value (Invoke-SystemAccessHyperVConsoleKeyboardKey -VirtualKey $virtualKey -Action $action -VMName $vmName -Server $server -ProcessId $processId)
                }
                "/api/hyperv/console/ctrl-alt-delete" {
                    if ($request.HttpMethod -ne "POST") { throw "POST required" }
                    $body = Read-RequestBody -Request $request
                    $vmName = [string](Get-JsonValue -Object $body -Name "vmName" -Default "")
                    $server = [string](Get-JsonValue -Object $body -Name "server" -Default "localhost")
                    $processId = [int](Get-JsonValue -Object $body -Name "processId" -Default 0)
                    Write-Json -Response $response -Value (Invoke-SystemAccessHyperVConsoleCtrlAltDelete -VMName $vmName -Server $server -ProcessId $processId)
                }
                "/api/mouse/move" {
                    if ($request.HttpMethod -ne "POST") { throw "POST required" }
                    $body = Read-RequestBody -Request $request
                    $x = Assert-JsonValue -Object $body -Name "x"
                    $y = Assert-JsonValue -Object $body -Name "y"
                    $relative = Get-JsonValue -Object $body -Name "relative" -Default $false
                    Write-Json -Response $response -Value (Invoke-SystemAccessMouseMove -X ([int]$x) -Y ([int]$y) -Relative ([bool]$relative))
                }
                "/api/mouse/click" {
                    if ($request.HttpMethod -ne "POST") { throw "POST required" }
                    $body = Read-RequestBody -Request $request
                    $button = [string](Get-JsonValue -Object $body -Name "button" -Default "left")
                    $clicks = [int](Get-JsonValue -Object $body -Name "clicks" -Default 1)
                    $x = Get-JsonValue -Object $body -Name "x"
                    $y = Get-JsonValue -Object $body -Name "y"
                    if ($null -ne $x -and $null -ne $y) {
                        Write-Json -Response $response -Value (Invoke-SystemAccessMouseClick -Button $button -X ([int]$x) -Y ([int]$y) -Clicks $clicks)
                    }
                    else {
                        Write-Json -Response $response -Value (Invoke-SystemAccessMouseClick -Button $button -Clicks $clicks)
                    }
                }
                "/api/mouse/wheel" {
                    if ($request.HttpMethod -ne "POST") { throw "POST required" }
                    $body = Read-RequestBody -Request $request
                    $delta = Assert-JsonValue -Object $body -Name "delta"
                    Write-Json -Response $response -Value (Invoke-SystemAccessMouseWheel -Delta ([int]$delta))
                }
                "/api/keyboard/type" {
                    if ($request.HttpMethod -ne "POST") { throw "POST required" }
                    $body = Read-RequestBody -Request $request
                    $text = Get-JsonValue -Object $body -Name "text" -Default ""
                    Write-Json -Response $response -Value (Invoke-SystemAccessKeyboardType -Text ([string]$text))
                }
                "/api/keyboard/key" {
                    if ($request.HttpMethod -ne "POST") { throw "POST required" }
                    $body = Read-RequestBody -Request $request
                    $virtualKey = Assert-JsonValue -Object $body -Name "virtualKey"
                    $action = [string](Get-JsonValue -Object $body -Name "action" -Default "press")
                    Write-Json -Response $response -Value (Invoke-SystemAccessKeyboardKey -VirtualKey ([int]$virtualKey) -Action $action)
                }
                default {
                    Write-ErrorJson -Response $response -StatusCode 404 -Message "Not found"
                }
        }
    }
    catch {
        Write-ErrorJson -Response $response -StatusCode 400 -Message $_.Exception.Message
    }
}

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
            Invoke-SystemAccessWebContext -Context $contextTask.GetAwaiter().GetResult()
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
    Write-Host "SystemAccessMCP web server stopped."
}
