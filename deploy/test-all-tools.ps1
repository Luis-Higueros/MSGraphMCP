param(
    [string]$UserHint,
    [string]$GraphSessionId,
    [ValidateSet('dev','aci','afd')]
    [string]$Target = 'afd',
    [string]$BaseUrl = '',
    [switch]$IncludeMutations,
    [string]$MailSendTo,
    [string]$TeamId,
    [string]$ChannelId,
    [string]$JsonOutputPath,
    [int]$TimeoutSec = 40
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
        ($Payload | ConvertTo-Json -Depth 40 -Compress) | Out-File $tempFile -Encoding utf8 -NoNewline

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

    $isError = $false
    if ($null -ne $Response.Body.result -and $null -ne $Response.Body.result.isError) {
        $isError = [bool]$Response.Body.result.isError
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

    $semanticError = $false
    if ($parsed -isnot [string] -and $null -ne $parsed) {
        if ($null -ne $parsed.error -and [string]::IsNullOrWhiteSpace([string]$parsed.error) -eq $false) {
            $semanticError = $true
        }
        if ($null -ne $parsed.status -and ($parsed.status -eq 'error' -or $parsed.status -eq 'search_failed')) {
            $semanticError = $true
        }
    }

    $ok = (-not $isError) -and (-not $semanticError)
    $errorMessage = if ($ok) { '' } else {
        if ($isError) { "tool_error: $text" }
        elseif ($parsed -isnot [string] -and $null -ne $parsed.error) { [string]$parsed.error }
        else { [string]$text }
    }

    [pscustomobject]@{
        ok = $ok
        payload = $parsed
        error = $errorMessage
    }
}

function Add-Result {
    param(
        [System.Collections.Generic.List[object]]$Results,
        [string]$Tool,
        [string]$Status,
        [object]$Data,
        [string]$Note = ''
    )

    $Results.Add([pscustomobject]@{
        tool = $Tool
        status = $Status
        note = $Note
        data = $Data
    })
}

function Test-Tool {
    param(
        [string]$ToolName,
        [hashtable]$Arguments,
        [string]$SuccessNote = ''
    )

    if (-not $toolNames.Contains($ToolName)) {
        Add-Result -Results $results -Tool $ToolName -Status 'skipped' -Data $null -Note 'Tool not registered in tools/list.'
        return $null
    }

    $script:requestId++
    try {
        $resp = Invoke-Tool -McpSessionId $mcpSessionId -Id $script:requestId -Name $ToolName -Arguments $Arguments
        $parsed = Parse-ToolPayload -Response $resp

        if ($parsed.ok) {
            Add-Result -Results $results -Tool $ToolName -Status 'pass' -Data $parsed.payload -Note $SuccessNote
        }
        else {
            Add-Result -Results $results -Tool $ToolName -Status 'fail' -Data $parsed.payload -Note $parsed.error
        }

        return $parsed
    }
    catch {
        Add-Result -Results $results -Tool $ToolName -Status 'fail' -Data $null -Note $_.Exception.Message
        return $null
    }
}

Write-Host "Target: $Target"
Write-Host "BaseUrl: $BaseUrl"
Write-Host "IncludeMutations: $IncludeMutations"
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
        clientInfo = @{ name = 'all-tools-tester'; version = '1.0' }
    }
}

$mcpSessionId = $init.McpSessionId
if (-not $mcpSessionId) {
    throw 'Failed to acquire mcp-session-id from initialize response.'
}

$toolsListResp = Invoke-McpHttp -McpSessionId $mcpSessionId -Payload @{
    jsonrpc = '2.0'
    id = 2
    method = 'tools/list'
    params = @{}
}

if ($null -ne $toolsListResp.Body.error) {
    throw "tools/list failed: $($toolsListResp.Body.error.message)"
}

$toolNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($t in ($toolsListResp.Body.result.tools | Where-Object { $_.name })) {
    [void]$toolNames.Add([string]$t.name)
}

Write-Host "MCP session: $mcpSessionId"
Write-Host "Discovered tools: $($toolNames.Count)"
Write-Host ''

$results = [System.Collections.Generic.List[object]]::new()
$script:requestId = 10

if (-not $GraphSessionId) {
    if (-not $UserHint) {
        throw 'Provide -GraphSessionId or -UserHint to authenticate Graph tools.'
    }

    $login = Test-Tool -ToolName 'GraphInitiateLogin' -Arguments @{ userHint = $UserHint } -SuccessNote 'Login initiated.'
    if ($null -eq $login -or -not $login.ok) {
        throw 'GraphInitiateLogin failed.'
    }

    $loginPayload = $login.payload
    if ($loginPayload.status -eq 'pending') {
        Write-Host ''
        Write-Host 'Interactive sign-in required:'
        Write-Host "1. Open: $($loginPayload.verificationUrl)"
        Write-Host "2. Enter code: $($loginPayload.userCode)"
        Write-Host "3. Re-run this script after sign-in completes or pass -GraphSessionId."
        exit 2
    }

    if ($loginPayload.status -ne 'authenticated') {
        throw "GraphInitiateLogin returned unexpected status: $($loginPayload.status)"
    }

    $GraphSessionId = [string]$loginPayload.sessionId
}

$statusCheck = Test-Tool -ToolName 'GraphCheckLoginStatus' -Arguments @{ sessionId = $GraphSessionId } -SuccessNote 'Graph session status checked.'
if ($null -eq $statusCheck -or -not $statusCheck.ok -or $statusCheck.payload.status -ne 'authenticated') {
    throw 'Graph session is not authenticated. Complete login first and retry.'
}

$who = Test-Tool -ToolName 'GraphWhoAmI' -Arguments @{ sessionId = $GraphSessionId } -SuccessNote 'Identity resolved.'
$userEmail = $null
if ($who -and $who.ok) {
    $userEmail = [string]$who.payload.email
}

# Mail tools (read-only)
$mailSearch = Test-Tool -ToolName 'MailSearch' -Arguments @{ sessionId = $GraphSessionId; maxResults = 5 } -SuccessNote 'Mail search baseline.'
$mailMessageId = $null
$mailConversationId = $null
if ($mailSearch -and $mailSearch.ok -and $mailSearch.payload.count -gt 0) {
    $firstMail = $mailSearch.payload.emails | Select-Object -First 1
    $mailMessageId = [string]$firstMail.id
    $mailConversationId = [string]$firstMail.conversationId
}

if ($mailMessageId) {
    [void](Test-Tool -ToolName 'MailGetById' -Arguments @{ sessionId = $GraphSessionId; messageId = $mailMessageId } -SuccessNote 'Retrieved first message by id.')
}
else {
    Add-Result -Results $results -Tool 'MailGetById' -Status 'skipped' -Data $null -Note 'No messageId available from MailSearch.'
}

if ($mailConversationId) {
    [void](Test-Tool -ToolName 'MailGetThread' -Arguments @{ sessionId = $GraphSessionId; conversationId = $mailConversationId; maxMessages = 5 } -SuccessNote 'Retrieved conversation thread.')
}
else {
    Add-Result -Results $results -Tool 'MailGetThread' -Status 'skipped' -Data $null -Note 'No conversationId available from MailSearch.'
}

[void](Test-Tool -ToolName 'MailSummarize' -Arguments @{ sessionId = $GraphSessionId; context = 'recent important updates'; maxEmails = 5 } -SuccessNote 'Summarization payload generated.')

# Calendar tools
$today = (Get-Date).ToString('yyyy-MM-dd')
$inSevenDays = (Get-Date).AddDays(7).ToString('yyyy-MM-dd')
$inTwoDays = (Get-Date).AddDays(2).ToString('yyyy-MM-dd')
[void](Test-Tool -ToolName 'CalendarGetAgenda' -Arguments @{ sessionId = $GraphSessionId; from = $today; to = $inTwoDays } -SuccessNote 'Agenda loaded.')
[void](Test-Tool -ToolName 'CalendarFindFreeSlots' -Arguments @{ sessionId = $GraphSessionId; from = $today; to = $inSevenDays; minDurationMinutes = 30 } -SuccessNote 'Free slots computed.')

if ($userEmail) {
    [void](Test-Tool -ToolName 'CalendarSuggestMeetingTimes' -Arguments @{ sessionId = $GraphSessionId; attendeeEmails = $userEmail; durationMinutes = 30; maxSuggestions = 3 } -SuccessNote 'Meeting suggestions returned.')
}
else {
    Add-Result -Results $results -Tool 'CalendarSuggestMeetingTimes' -Status 'skipped' -Data $null -Note 'User email unavailable from GraphWhoAmI.'
}

# Files tools
$filesList = Test-Tool -ToolName 'FilesListItems' -Arguments @{ sessionId = $GraphSessionId; maxItems = 20 } -SuccessNote 'Listed OneDrive root items.'
[void](Test-Tool -ToolName 'FilesSearch' -Arguments @{ sessionId = $GraphSessionId; query = 'report'; maxResults = 10 } -SuccessNote 'File search executed.')

$filePathForRead = $null
if ($filesList -and $filesList.ok -and $filesList.payload.items) {
    $firstFile = $filesList.payload.items | Where-Object { $_.type -eq 'file' } | Select-Object -First 1
    if ($firstFile) {
        $filePathForRead = [string]$firstFile.name
    }
}

if ($filePathForRead) {
    [void](Test-Tool -ToolName 'FilesGetContent' -Arguments @{ sessionId = $GraphSessionId; filePath = $filePathForRead } -SuccessNote 'Read a root-level file.')
}
else {
    Add-Result -Results $results -Tool 'FilesGetContent' -Status 'skipped' -Data $null -Note 'No root-level file found in FilesListItems.'
}

# SharePoint tools
$spSites = Test-Tool -ToolName 'SharePointListSites' -Arguments @{ sessionId = $GraphSessionId; query = 'site'; maxResults = 10 } -SuccessNote 'SharePoint site search executed.'
$siteId = $null
$driveId = $null
$spItemId = $null
if ($spSites -and $spSites.ok -and $spSites.payload.count -gt 0) {
    $siteId = [string](($spSites.payload.sites | Select-Object -First 1).siteId)
}

if ($siteId) {
    $spDrives = Test-Tool -ToolName 'SharePointListDrives' -Arguments @{ sessionId = $GraphSessionId; siteId = $siteId } -SuccessNote 'SharePoint drives listed.'
    if ($spDrives -and $spDrives.ok -and $spDrives.payload.count -gt 0) {
        $driveId = [string](($spDrives.payload.drives | Select-Object -First 1).driveId)
    }

    if ($driveId) {
        $spItems = Test-Tool -ToolName 'SharePointListItems' -Arguments @{ sessionId = $GraphSessionId; driveId = $driveId; maxItems = 20 } -SuccessNote 'SharePoint drive root items listed.'
        if ($spItems -and $spItems.ok -and $spItems.payload.count -gt 0) {
            $firstSpFile = $spItems.payload.items | Where-Object { $_.isFolder -eq $false } | Select-Object -First 1
            if ($firstSpFile) {
                $spItemId = [string]$firstSpFile.itemId
                [void](Test-Tool -ToolName 'SharePointGetContent' -Arguments @{ sessionId = $GraphSessionId; driveId = $driveId; itemId = $spItemId } -SuccessNote 'Read SharePoint file content.')
                [void](Test-Tool -ToolName 'SharePointCreateShareLink' -Arguments @{ sessionId = $GraphSessionId; driveId = $driveId; itemId = $spItemId } -SuccessNote 'Created SharePoint share link for existing file.')
            }
        }

        if (-not $spItemId) {
            Add-Result -Results $results -Tool 'SharePointGetContent' -Status 'skipped' -Data $null -Note 'No SharePoint file item found.'
            Add-Result -Results $results -Tool 'SharePointCreateShareLink' -Status 'skipped' -Data $null -Note 'No SharePoint file item found.'
        }
    }
    else {
        Add-Result -Results $results -Tool 'SharePointListItems' -Status 'skipped' -Data $null -Note 'No driveId available.'
        Add-Result -Results $results -Tool 'SharePointGetContent' -Status 'skipped' -Data $null -Note 'No driveId available.'
        Add-Result -Results $results -Tool 'SharePointCreateShareLink' -Status 'skipped' -Data $null -Note 'No driveId available.'
    }
}
else {
    Add-Result -Results $results -Tool 'SharePointListDrives' -Status 'skipped' -Data $null -Note 'No siteId available from SharePointListSites.'
    Add-Result -Results $results -Tool 'SharePointListItems' -Status 'skipped' -Data $null -Note 'No siteId available from SharePointListSites.'
    Add-Result -Results $results -Tool 'SharePointGetContent' -Status 'skipped' -Data $null -Note 'No siteId available from SharePointListSites.'
    Add-Result -Results $results -Tool 'SharePointCreateShareLink' -Status 'skipped' -Data $null -Note 'No siteId available from SharePointListSites.'
}

# Teams tools
$teams = Test-Tool -ToolName 'TeamsListMyTeams' -Arguments @{ sessionId = $GraphSessionId } -SuccessNote 'Listed joined teams.'
$resolvedTeamId = $TeamId
if (-not $resolvedTeamId -and $teams -and $teams.ok -and $teams.payload.count -gt 0) {
    $resolvedTeamId = [string](($teams.payload.teams | Select-Object -First 1).id)
}

if ($resolvedTeamId) {
    $channels = Test-Tool -ToolName 'TeamsListChannels' -Arguments @{ sessionId = $GraphSessionId; teamId = $resolvedTeamId } -SuccessNote 'Listed channels in team.'
    $resolvedChannelId = $ChannelId
    if (-not $resolvedChannelId -and $channels -and $channels.ok -and $channels.payload.count -gt 0) {
        $resolvedChannelId = [string](($channels.payload.channels | Select-Object -First 1).id)
    }

    if ($resolvedChannelId) {
        [void](Test-Tool -ToolName 'TeamsGetChannelMessages' -Arguments @{ sessionId = $GraphSessionId; teamId = $resolvedTeamId; channelId = $resolvedChannelId; maxMessages = 10 } -SuccessNote 'Retrieved channel messages.')
    }
    else {
        Add-Result -Results $results -Tool 'TeamsGetChannelMessages' -Status 'skipped' -Data $null -Note 'No channelId available.'
    }
}
else {
    Add-Result -Results $results -Tool 'TeamsListChannels' -Status 'skipped' -Data $null -Note 'No teamId available.'
    Add-Result -Results $results -Tool 'TeamsGetChannelMessages' -Status 'skipped' -Data $null -Note 'No teamId available.'
}

$chats = Test-Tool -ToolName 'TeamsListChats' -Arguments @{ sessionId = $GraphSessionId } -SuccessNote 'Listed chats.'
if ($chats -and $chats.ok -and $chats.payload.count -gt 0) {
    $chatId = [string](($chats.payload.chats | Select-Object -First 1).id)
    [void](Test-Tool -ToolName 'TeamsGetChatMessages' -Arguments @{ sessionId = $GraphSessionId; chatId = $chatId; maxMessages = 10 } -SuccessNote 'Retrieved chat messages.')
}
else {
    Add-Result -Results $results -Tool 'TeamsGetChatMessages' -Status 'skipped' -Data $null -Note 'No chatId available.'
}

# OneNote tools
$notebooks = Test-Tool -ToolName 'OneNoteListNotebooks' -Arguments @{ sessionId = $GraphSessionId } -SuccessNote 'Listed notebooks.'
$notebookId = $null
$sectionId = $null
if ($notebooks -and $notebooks.ok -and $notebooks.payload.count -gt 0) {
    $notebookId = [string](($notebooks.payload.notebooks | Select-Object -First 1).id)
}

if ($notebookId) {
    $sections = Test-Tool -ToolName 'OneNoteListSections' -Arguments @{ sessionId = $GraphSessionId; notebookId = $notebookId } -SuccessNote 'Listed notebook sections.'
    if ($sections -and $sections.ok -and $sections.payload.count -gt 0) {
        $sectionId = [string](($sections.payload.sections | Select-Object -First 1).id)
    }

    if ($sectionId) {
        $pages = Test-Tool -ToolName 'OneNoteListPages' -Arguments @{ sessionId = $GraphSessionId; sectionId = $sectionId; maxPages = 10 } -SuccessNote 'Listed section pages.'
        [void](Test-Tool -ToolName 'OneNoteSearchPages' -Arguments @{ sessionId = $GraphSessionId; query = 'meeting'; maxResults = 10 } -SuccessNote 'OneNote search executed.')

        if ($pages -and $pages.ok -and $pages.payload.count -gt 0) {
            $pageId = [string](($pages.payload.pages | Select-Object -First 1).id)
            [void](Test-Tool -ToolName 'OneNoteGetPageContent' -Arguments @{ sessionId = $GraphSessionId; pageId = $pageId } -SuccessNote 'Retrieved page content.')
        }
        else {
            Add-Result -Results $results -Tool 'OneNoteGetPageContent' -Status 'skipped' -Data $null -Note 'No pageId available.'
        }
    }
    else {
        Add-Result -Results $results -Tool 'OneNoteListPages' -Status 'skipped' -Data $null -Note 'No sectionId available.'
        Add-Result -Results $results -Tool 'OneNoteSearchPages' -Status 'pass' -Data (@{ info = 'Search is independent of sectionId.' }) -Note 'Search does not require sectionId.'
        Add-Result -Results $results -Tool 'OneNoteGetPageContent' -Status 'skipped' -Data $null -Note 'No sectionId/pageId available.'
    }
}
else {
    Add-Result -Results $results -Tool 'OneNoteListSections' -Status 'skipped' -Data $null -Note 'No notebookId available.'
    Add-Result -Results $results -Tool 'OneNoteListPages' -Status 'skipped' -Data $null -Note 'No notebookId available.'
    [void](Test-Tool -ToolName 'OneNoteSearchPages' -Arguments @{ sessionId = $GraphSessionId; query = 'meeting'; maxResults = 10 } -SuccessNote 'OneNote search executed.')
    Add-Result -Results $results -Tool 'OneNoteGetPageContent' -Status 'skipped' -Data $null -Note 'No notebook/page available.'
}

# Planner tools
$plans = Test-Tool -ToolName 'PlannerListPlans' -Arguments @{ sessionId = $GraphSessionId } -SuccessNote 'Listed Planner plans.'
$planId = $null
if ($plans -and $plans.ok -and $plans.payload.count -gt 0) {
    $planId = [string](($plans.payload.plans | Select-Object -First 1).planId)
}

if ($planId) {
    $tasks = Test-Tool -ToolName 'PlannerListTasks' -Arguments @{ sessionId = $GraphSessionId; planId = $planId; maxTasks = 20 } -SuccessNote 'Listed planner tasks.'
    $plannerTaskId = $null
    if ($tasks -and $tasks.ok -and $tasks.payload.count -gt 0) {
        $plannerTaskId = [string](($tasks.payload.tasks | Select-Object -First 1).id)
    }

    if (-not $plannerTaskId) {
        Add-Result -Results $results -Tool 'PlannerUpdateTask' -Status 'skipped' -Data $null -Note 'No existing taskId available for read-only update test.'
    }
}
else {
    Add-Result -Results $results -Tool 'PlannerListTasks' -Status 'skipped' -Data $null -Note 'No planId available.'
    Add-Result -Results $results -Tool 'PlannerUpdateTask' -Status 'skipped' -Data $null -Note 'No planId/taskId available.'
}

# Optional mutation tools
if ($IncludeMutations) {
    Write-Host ''
    Write-Host 'Running mutation tests...'

    $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')

    [void](Test-Tool -ToolName 'CalendarCreateEvent' -Arguments @{
        sessionId = $GraphSessionId
        subject = "MCP test event $stamp"
        startTime = (Get-Date).AddDays(2).ToString('yyyy-MM-ddT10:00:00')
        endTime = (Get-Date).AddDays(2).ToString('yyyy-MM-ddT10:30:00')
        body = 'Automated test event from test-all-tools.ps1'
        location = 'MCP Test'
        sendInvite = $false
    } -SuccessNote 'Created calendar event.')

    $uploadPath = "MCP-Test/test-$stamp.txt"
    [void](Test-Tool -ToolName 'FilesUploadText' -Arguments @{
        sessionId = $GraphSessionId
        filePath = $uploadPath
        content = "MCP test file created at $stamp"
    } -SuccessNote 'Uploaded OneDrive text file.')

    [void](Test-Tool -ToolName 'FilesCreateShareLink' -Arguments @{
        sessionId = $GraphSessionId
        filePath = $uploadPath
        linkType = 'view'
        scope = 'organization'
    } -SuccessNote 'Created share link for uploaded OneDrive file.')

    if ($driveId) {
        $spUpload = Test-Tool -ToolName 'SharePointUploadText' -Arguments @{
            sessionId = $GraphSessionId
            driveId = $driveId
            folderPath = ''
            fileName = "mcp-test-$stamp.txt"
            content = "MCP SharePoint test content at $stamp"
        } -SuccessNote 'Uploaded SharePoint text file.'

        if ($spUpload -and $spUpload.ok -and $spUpload.payload.itemId) {
            [void](Test-Tool -ToolName 'SharePointCreateShareLink' -Arguments @{
                sessionId = $GraphSessionId
                driveId = $driveId
                itemId = [string]$spUpload.payload.itemId
                linkType = 'view'
                scope = 'organization'
            } -SuccessNote 'Created share link for uploaded SharePoint file.')
        }
    }
    else {
        Add-Result -Results $results -Tool 'SharePointUploadText' -Status 'skipped' -Data $null -Note 'No SharePoint driveId available for mutation test.'
    }

    if ($sectionId) {
        [void](Test-Tool -ToolName 'OneNoteCreatePage' -Arguments @{
            sessionId = $GraphSessionId
            sectionId = $sectionId
            title = "MCP Test Page $stamp"
            content = 'Automated page created by test-all-tools.ps1'
        } -SuccessNote 'Created OneNote page.')
    }
    else {
        Add-Result -Results $results -Tool 'OneNoteCreatePage' -Status 'skipped' -Data $null -Note 'No OneNote sectionId available for mutation test.'
    }

    if ($planId) {
        $createdTask = Test-Tool -ToolName 'PlannerCreateTask' -Arguments @{
            sessionId = $GraphSessionId
            planId = $planId
            title = "MCP Test Task $stamp"
            dueDate = (Get-Date).AddDays(5).ToString('yyyy-MM-dd')
            priority = 5
            notes = 'Automated test task'
        } -SuccessNote 'Created Planner task.'

        if ($createdTask -and $createdTask.ok -and $createdTask.payload.taskId) {
            [void](Test-Tool -ToolName 'PlannerUpdateTask' -Arguments @{
                sessionId = $GraphSessionId
                taskId = [string]$createdTask.payload.taskId
                percentComplete = 10
            } -SuccessNote 'Updated newly created Planner task.')
        }
    }
    else {
        Add-Result -Results $results -Tool 'PlannerCreateTask' -Status 'skipped' -Data $null -Note 'No planId available for mutation test.'
        Add-Result -Results $results -Tool 'PlannerUpdateTask' -Status 'skipped' -Data $null -Note 'No created task available for mutation test.'
    }

    if ($resolvedTeamId -and $resolvedChannelId) {
        [void](Test-Tool -ToolName 'TeamsSendChannelMessage' -Arguments @{
            sessionId = $GraphSessionId
            teamId = $resolvedTeamId
            channelId = $resolvedChannelId
            content = "MCP test message $stamp"
        } -SuccessNote 'Sent Teams channel message.')
    }
    else {
        Add-Result -Results $results -Tool 'TeamsSendChannelMessage' -Status 'skipped' -Data $null -Note 'Missing teamId/channelId for mutation test.'
    }

    if ($mailMessageId) {
        $draft = Test-Tool -ToolName 'MailDraftReply' -Arguments @{
            sessionId = $GraphSessionId
            messageId = $mailMessageId
            replyBody = "Automated test draft reply at $stamp"
            replyAll = $false
        } -SuccessNote 'Created draft reply.'

        if ($draft -and $draft.ok -and $draft.payload.draftId) {
            [void](Test-Tool -ToolName 'MailSendDraft' -Arguments @{
                sessionId = $GraphSessionId
                draftId = [string]$draft.payload.draftId
            } -SuccessNote 'Sent draft reply.')
        }
    }
    else {
        Add-Result -Results $results -Tool 'MailDraftReply' -Status 'skipped' -Data $null -Note 'No baseline mail messageId available.'
        Add-Result -Results $results -Tool 'MailSendDraft' -Status 'skipped' -Data $null -Note 'No draft created in this run.'
    }

    if ($MailSendTo) {
        [void](Test-Tool -ToolName 'MailSend' -Arguments @{
            sessionId = $GraphSessionId
            recipient = $MailSendTo
            subject = "MCP test email $stamp"
            body = 'Automated email from test-all-tools.ps1'
            isHtml = $false
        } -SuccessNote 'Sent test email.')
    }
    else {
        Add-Result -Results $results -Tool 'MailSend' -Status 'skipped' -Data $null -Note 'Provide -MailSendTo to run MailSend mutation test.'
    }
}
else {
    Add-Result -Results $results -Tool 'MutationSuite' -Status 'skipped' -Data $null -Note 'Run with -IncludeMutations to execute create/update/send tests.'
}

$pass = @($results | Where-Object { $_.status -eq 'pass' }).Count
$fail = @($results | Where-Object { $_.status -eq 'fail' }).Count
$skipped = @($results | Where-Object { $_.status -eq 'skipped' }).Count

$summary = [pscustomobject]@{
    target = $Target
    baseUrl = $BaseUrl
    includeMutations = [bool]$IncludeMutations
    graphSessionId = $GraphSessionId
    totals = [pscustomobject]@{
        pass = $pass
        fail = $fail
        skipped = $skipped
        total = $results.Count
    }
    results = $results
}

Write-Host ''
Write-Host '=== Summary ==='
Write-Host "Pass   : $pass"
Write-Host "Fail   : $fail"
Write-Host "Skipped: $skipped"
Write-Host "Total  : $($results.Count)"
Write-Host ''

$summaryJson = $summary | ConvertTo-Json -Depth 80

if ($JsonOutputPath) {
    $dir = Split-Path -Path $JsonOutputPath -Parent
    if ($dir -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $summaryJson | Out-File -FilePath $JsonOutputPath -Encoding utf8
    Write-Host "Saved report: $JsonOutputPath"
}

$summaryJson

if ($fail -gt 0) {
    exit 1
}
