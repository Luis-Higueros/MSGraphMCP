param(
    [Parameter(Mandatory = $true)]
    [string]$GraphSessionId,

    [string]$BaseUrl = 'https://ep-msgraphmcp-43613-c6dvbtfyfccmhzf8.a03.azurefd.net',
    [string]$Keywords = 'Meeting Summarized',
    [string]$Folder = '',
    [string]$Since = '',
    [string]$Until = '',
    [int]$MaxResults = 25,

    # Run a small matrix to isolate failure patterns.
    [switch]$Matrix
)

$ErrorActionPreference = 'Stop'

function Invoke-McpHttp {
    param(
        [hashtable]$Payload,
        [string]$McpSessionId = ''
    )

    $tempFile = Join-Path $env:TEMP ('mcp_' + [Guid]::NewGuid().ToString('N') + '.json')
    try {
        ($Payload | ConvertTo-Json -Depth 20 -Compress) | Out-File $tempFile -Encoding utf8 -NoNewline

        $headers = @('-H', 'Content-Type: application/json', '-H', 'Accept: application/json, text/event-stream')
        if ($McpSessionId) {
            $headers += @('-H', "mcp-session-id: $McpSessionId")
        }

        $response = & curl.exe -sS -D - @headers -X POST "$BaseUrl/mcp" -d "@$tempFile"
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

    return Invoke-McpHttp -McpSessionId $McpSessionId -Payload @{
        jsonrpc = '2.0'
        id = $Id
        method = 'tools/call'
        params = @{
            name = $Name
            arguments = $Arguments
        }
    }
}

Write-Host "BaseUrl        : $BaseUrl"
Write-Host "GraphSessionId : $GraphSessionId"
Write-Host ""

$init = Invoke-McpHttp -Payload @{
    jsonrpc = '2.0'
    id = 1
    method = 'initialize'
    params = @{
        protocolVersion = '2024-11-05'
        capabilities = @{}
        clientInfo = @{ name = 'mail-search-tester'; version = '1.0' }
    }
}

$mcpSessionId = $init.McpSessionId
if (-not $mcpSessionId) {
    throw 'Failed to acquire mcp-session-id from initialize response.'
}

Write-Host "McpSessionId   : $mcpSessionId"

$status = Invoke-Tool -McpSessionId $mcpSessionId -Id 2 -Name 'GraphCheckLoginStatus' -Arguments @{
    sessionId = $GraphSessionId
}

$statusPayload = $status.Body.result.content[0].text | ConvertFrom-Json
Write-Host "Auth status    : $($statusPayload.status)"
if ($statusPayload.status -ne 'authenticated') {
    Write-Host ''
    Write-Host 'Graph session is not authenticated. Full status payload:'
    $statusPayload | ConvertTo-Json -Depth 20
    exit 1
}

Write-Host ''

if ($Matrix) {
    $testMatrix = @(
        @{ label = 'No keywords'; args = @{ sessionId = $GraphSessionId; maxResults = $MaxResults } },
        @{ label = 'Keywords only'; args = @{ sessionId = $GraphSessionId; keywords = $Keywords; folder = $Folder; maxResults = $MaxResults } },
        @{ label = 'Date range only'; args = @{ sessionId = $GraphSessionId; folder = $Folder; since = $Since; until = $Until; maxResults = $MaxResults } },
        @{ label = 'Keywords + date range'; args = @{ sessionId = $GraphSessionId; keywords = $Keywords; folder = $Folder; since = $Since; until = $Until; maxResults = $MaxResults } }
    )

    $id = 10
    foreach ($entry in $testMatrix) {
        $toolArgs = @{}
        foreach ($k in $entry.args.Keys) {
            if ($entry.args[$k] -is [string] -and [string]::IsNullOrWhiteSpace($entry.args[$k])) { continue }
            $toolArgs[$k] = $entry.args[$k]
        }

        Write-Host "=== $($entry.label) ==="
        $result = Invoke-Tool -McpSessionId $mcpSessionId -Id $id -Name 'MailSearch' -Arguments $toolArgs
        $result.Body | ConvertTo-Json -Depth 30
        Write-Host ''
        $id++
    }

    exit 0
}

$arguments = @{
    sessionId = $GraphSessionId
    keywords = $Keywords
    folder = $Folder
    maxResults = $MaxResults
}
if (-not [string]::IsNullOrWhiteSpace($Since)) { $arguments.since = $Since }
if (-not [string]::IsNullOrWhiteSpace($Until)) { $arguments.until = $Until }

$result = Invoke-Tool -McpSessionId $mcpSessionId -Id 10 -Name 'MailSearch' -Arguments $arguments
$result.Body | ConvertTo-Json -Depth 30
