param(
    [Parameter(Mandatory = $true)]
    [string]$UserHint,
    [string]$TeamId,
    [string]$ChannelId,

    # -Target selects a preset environment:
    #   dev  — local dev server (default)  http://127.0.0.1:8080
    #   aci  — Azure Container Instance    http://<aci-fqdn>:8080
    #   afd  — Azure Front Door (HTTPS)    https://<fd-hostname>  (recommended for production)
    [ValidateSet('dev','aci','afd')]
    [string]$Target = 'dev',

    # Override the base URL entirely (takes precedence over -Target)
    [string]$BaseUrl = ''
)

# ─── Resolve BaseUrl from -Target if not explicitly overridden ─────────────────
if (-not $BaseUrl) {
    $BaseUrl = switch ($Target) {
        'dev' { 'http://127.0.0.1:8080' }
        'aci' { 'http://msgraph-mcp-weu-27992.westeurope.azurecontainer.io:8080' }
        'afd' { 'https://ep-msgraphmcp-43613-c6dvbtfyfccmhzf8.a03.azurefd.net' }
    }
}
Write-Host "Target : $Target"
Write-Host "BaseUrl: $BaseUrl"
Write-Host ''

$ErrorActionPreference = 'Stop'

function Invoke-McpHttp {
    param(
        [hashtable]$Payload,
        [string]$McpSessionId = ''
    )

    $tempFile = Join-Path $env:TEMP ('mcp_' + [Guid]::NewGuid().ToString('N') + '.json')
    try {
        ($Payload | ConvertTo-Json -Depth 10 -Compress) | Out-File $tempFile -Encoding utf8 -NoNewline

        $headers = @('-H', 'Content-Type: application/json')
        if ($McpSessionId) {
            $headers += @('-H', "mcp-session-id: $McpSessionId")
        }

        $response = & curl.exe -s -D - @headers -X POST "$BaseUrl/mcp" -d "@$tempFile"
        if ($LASTEXITCODE -ne 0) {
            throw 'curl request to MCP endpoint failed.'
        }

        $raw = ($response -join "`n")
        $sessionMatch = [regex]::Match($raw, 'mcp-session-id:\s*([^\r\n]+)')
        $dataLine = (($raw -split "`n") | Where-Object { $_ -like 'data:*' } | Select-Object -Last 1)
        if (-not $dataLine) {
            throw 'No MCP data payload returned.'
        }

        $json = $dataLine.Substring(5).Trim() | ConvertFrom-Json -Depth 20
        [pscustomobject]@{
            McpSessionId = if ($sessionMatch.Success) { $sessionMatch.Groups[1].Value.Trim() } else { $McpSessionId }
            Body = $json
        }
    }
    finally {
        Remove-Item $tempFile -ErrorAction SilentlyContinue
    }
}

# ─── Connectivity pre-check ───────────────────────────────────────────────────
try {
    $null = Invoke-WebRequest -Uri "$BaseUrl/health" -UseBasicParsing -TimeoutSec 10
} catch {
    if ($Target -eq 'dev') {
        throw "MSGraphMCP is not reachable at $BaseUrl. Start it first with .\deploy\run-local.ps1"
    } else {
        throw "MSGraphMCP is not reachable at $BaseUrl. Verify the container is running.`n$_"
    }
}

$init = Invoke-McpHttp -Payload @{
    jsonrpc = '2.0'
    id = 1
    method = 'initialize'
    params = @{
        protocolVersion = '2024-11-05'
        capabilities = @{}
        clientInfo = @{ name = 'scope-smoke-runner'; version = '1.0' }
    }
}

$mcpSessionId = $init.McpSessionId
if (-not $mcpSessionId) {
    throw 'Failed to obtain MCP session id.'
}

$login = Invoke-McpHttp -McpSessionId $mcpSessionId -Payload @{
    jsonrpc = '2.0'
    id = 2
    method = 'tools/call'
    params = @{
        name = 'GraphInitiateLogin'
        arguments = @{ userHint = $UserHint }
    }
}

$text = $login.Body.result.content[0].text | ConvertFrom-Json -Depth 20
if ($text.status -eq 'pending') {
    Write-Host ''
    Write-Host 'Interactive sign-in required.'
    Write-Host "1. Open: $($text.verificationUrl)"
    Write-Host "2. Enter code: $($text.userCode)"
    Write-Host '3. Re-run this command after sign-in completes.'
    return
}

if ($text.status -ne 'authenticated') {
    throw "Login failed: $($text | ConvertTo-Json -Compress)"
}

$scopeSmokeBody = @{
    sessionId = $text.sessionId
}
if ($TeamId) { $scopeSmokeBody.teamId = $TeamId }
if ($ChannelId) { $scopeSmokeBody.channelId = $ChannelId }

$result = Invoke-RestMethod -Method Post -Uri "$BaseUrl/test/scope-smoke" -ContentType 'application/json' -Body ($scopeSmokeBody | ConvertTo-Json -Depth 10)
$result | ConvertTo-Json -Depth 10
