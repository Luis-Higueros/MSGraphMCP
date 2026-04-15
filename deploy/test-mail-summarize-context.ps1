param(
    [string]$UserHint,
    [string]$GraphSessionId,
    [ValidateSet('dev','aci','afd')]
    [string]$Target = 'afd',
    [string]$BaseUrl = '',
    [string]$Context = 'Summarize emails containing the exact phrase Meeting Summarized between 2026-03-19 and 2026-04-09.',
    [string]$Keywords = 'Meeting Summarized',
    [string]$Folder = '',
    [string]$Since = '2026-03-19',
    [string]$Until = '2026-04-09',
    [int]$MaxEmails = 30,
    [string]$JsonOutputPath = 'deploy/reports/mail-summarize-context-latest.json',
    [int]$TimeoutSec = 45
)

$ErrorActionPreference = 'Stop'

if (-not $BaseUrl) {
    $BaseUrl = switch ($Target) {
        'dev' { 'http://127.0.0.1:8080' }
        'aci' { 'http://msgraph-mcp-weu-27992.westeurope.azurecontainer.io:8080' }
        'afd' { 'https://ep-msgraphmcp-43613-c6dvbtfyfccmhzf8.a03.azurefd.net' }
    }
}

function Invoke-McpHttp {
    param(
        [hashtable]$Payload,
        [string]$McpSessionId = ''
    )

    $tempFile = Join-Path $env:TEMP ('mcp_' + [Guid]::NewGuid().ToString('N') + '.json')
    try {
        ($Payload | ConvertTo-Json -Depth 30 -Compress) | Out-File $tempFile -Encoding utf8 -NoNewline

        $headers = @('-H', 'Content-Type: application/json', '-H', 'Accept: application/json, text/event-stream')
        if ($McpSessionId) {
            $headers += @('-H', "mcp-session-id: $McpSessionId")
        }

        $response = & curl.exe --max-time $TimeoutSec -sS -D - @headers -X POST "$BaseUrl/mcp" -d "@$tempFile"
        if ($LASTEXITCODE -ne 0) {
            throw 'curl request failed.'
        }

        $raw = ($response -join "`n")
        $sessionMatch = [regex]::Match($raw, 'mcp-session-id:\s*([^\r\n]+)')
        $dataLine = (($raw -split "`n") | Where-Object { $_ -like 'data:*' } | Select-Object -Last 1)
        if (-not $dataLine) {
            throw "No MCP data payload returned.`nRaw response:`n$raw"
        }

        $json = $dataLine.Substring(5).Trim() | ConvertFrom-Json
        [pscustomobject]@{
            McpSessionId = if ($sessionMatch.Success) { $sessionMatch.Groups[1].Value.Trim() } else { $McpSessionId }
            Body = $json
            Raw = $raw
        }
    }
    finally {
        Remove-Item $tempFile -ErrorAction SilentlyContinue
    }
}

function Invoke-Tool {
    param(
        [string]$McpSessionId,
        [int]$Id,
        [string]$Name,
        [hashtable]$Arguments
    )

    Invoke-McpHttp -McpSessionId $McpSessionId -Payload @{
        jsonrpc = '2.0'
        id = $Id
        method = 'tools/call'
        params = @{
            name = $Name
            arguments = $Arguments
        }
    }
}

function Parse-ToolPayload {
    param([object]$Response)

    if ($null -ne $Response.Body.error) {
        return [pscustomobject]@{
            ok = $false
            payload = $null
            error = "transport_error: $($Response.Body.error.message)"
        }
    }

    $text = $null
    if ($null -ne $Response.Body.result -and $null -ne $Response.Body.result.content -and $Response.Body.result.content.Count -gt 0) {
        $text = $Response.Body.result.content[0].text
    }

    $parsed = $text
    if ($text -is [string] -and $text.Trim().StartsWith('{')) {
        try {
            $parsed = $text | ConvertFrom-Json
        }
        catch {
            $parsed = [pscustomobject]@{ rawText = $text }
        }
    }

    $isError = $false
    if ($null -ne $Response.Body.result -and $null -ne $Response.Body.result.isError) {
        $isError = [bool]$Response.Body.result.isError
    }

    $semanticError = $false
    if ($parsed -isnot [string] -and $null -ne $parsed) {
        if ($null -ne $parsed.error -and [string]::IsNullOrWhiteSpace([string]$parsed.error) -eq $false) {
            $semanticError = $true
        }
        if ($null -ne $parsed.status -and ($parsed.status -eq 'error' -or $parsed.status -eq 'search_failed')) {
            $semanticError = $true
        }
    }

    [pscustomobject]@{
        ok = (-not $isError) -and (-not $semanticError)
        payload = $parsed
        error = if ((-not $isError) -and (-not $semanticError)) { '' } elseif ($isError) { [string]$text } elseif ($parsed -isnot [string] -and $null -ne $parsed.error) { [string]$parsed.error } else { [string]$text }
    }
}

Write-Host "Target: $Target"
Write-Host "BaseUrl: $BaseUrl"
Write-Host ''

try {
    $health = Invoke-WebRequest -Uri "$BaseUrl/health" -UseBasicParsing -TimeoutSec $TimeoutSec
    Write-Host "Health: $($health.StatusCode)"
}
catch {
    throw "Health check failed for $BaseUrl. $_"
}

$init = Invoke-McpHttp -Payload @{
    jsonrpc = '2.0'
    id = 1
    method = 'initialize'
    params = @{
        protocolVersion = '2024-11-05'
        capabilities = @{}
        clientInfo = @{ name = 'mail-summarize-context-tester'; version = '1.0' }
    }
}

$mcpSessionId = $init.McpSessionId
if (-not $mcpSessionId) {
    throw 'Failed to acquire mcp-session-id from initialize response.'
}

Write-Host "MCP session: $mcpSessionId"

if ($GraphSessionId) {
    $status = Invoke-Tool -McpSessionId $mcpSessionId -Id 2 -Name 'GraphCheckLoginStatus' -Arguments @{ sessionId = $GraphSessionId }
    $statusParsed = Parse-ToolPayload -Response $status
    if (-not $statusParsed.ok -or $statusParsed.payload.status -ne 'authenticated') {
        if (-not $UserHint) {
            throw 'Provided GraphSessionId is not authenticated. Pass -UserHint for silent re-login or use a fresh Graph session.'
        }

        Write-Host 'Provided GraphSessionId is not authenticated. Attempting GraphInitiateLogin with UserHint...'
        $GraphSessionId = ''
    }
}

if (-not $GraphSessionId) {
    if (-not $UserHint) {
        throw 'Provide -GraphSessionId or -UserHint.'
    }

    $login = Invoke-Tool -McpSessionId $mcpSessionId -Id 3 -Name 'GraphInitiateLogin' -Arguments @{ userHint = $UserHint }
    $loginParsed = Parse-ToolPayload -Response $login
    if (-not $loginParsed.ok) {
        throw "GraphInitiateLogin failed: $($loginParsed.error)"
    }

    if ($loginParsed.payload.status -eq 'pending') {
        Write-Host ''
        Write-Host 'Interactive sign-in required:'
        Write-Host "1. Open: $($loginParsed.payload.verificationUrl)"
        Write-Host "2. Enter code: $($loginParsed.payload.userCode)"
        Write-Host '3. Re-run this command after sign-in completes.'
        exit 2
    }

    if ($loginParsed.payload.status -ne 'authenticated') {
        throw "Unexpected GraphInitiateLogin status: $($loginParsed.payload.status)"
    }

    $GraphSessionId = [string]$loginParsed.payload.sessionId
}

$mailSummarize = Invoke-Tool -McpSessionId $mcpSessionId -Id 4 -Name 'MailSummarize' -Arguments @{
    sessionId = $GraphSessionId
    context = $Context
    keywords = $Keywords
    folder = $Folder
    since = $Since
    until = $Until
    maxEmails = $MaxEmails
}

$mailParsed = Parse-ToolPayload -Response $mailSummarize
if (-not $mailParsed.ok) {
    throw "MailSummarize failed: $($mailParsed.error)"
}

$payload = $mailParsed.payload
$emailCount = $null
if ($payload.data -and $payload.data.count -ne $null) {
    $emailCount = [int]$payload.data.count
}

$summary = [pscustomobject]@{
    target = $Target
    baseUrl = $BaseUrl
    mcpSessionId = $mcpSessionId
    graphSessionId = $GraphSessionId
    summarizationRequest = $payload.summarizationRequest
    context = $payload.context
    emailCount = $emailCount
    outputPath = $JsonOutputPath
}

$dir = Split-Path -Path $JsonOutputPath -Parent
if ($dir -and -not (Test-Path $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

($payload | ConvertTo-Json -Depth 60) | Out-File -FilePath $JsonOutputPath -Encoding utf8

Write-Host ''
Write-Host 'MailSummarize context test succeeded.'
$summary | ConvertTo-Json -Depth 20
