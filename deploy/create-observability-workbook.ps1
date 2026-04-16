param(
    [string]$SubscriptionId = '380d9153-35d4-45fd-a604-fe72aaf453ae',
    [string]$ResourceGroup = 'euw-ea_teamsai-sandbox-rg',
    [string]$Location = 'westeurope',
    [string]$WorkspaceName = 'log-msgraphmcp-27992',
    [string]$AppInsightsName = 'ai-msgraphmcp-27992',
    [string]$DisplayName = 'MSGraphMCP Observability Workbook',
    [string]$WorkbookId = '9d5e08d2-6bf8-4a13-9f54-2f4ea4a8f245'
)

$ErrorActionPreference = 'Stop'

az account set --subscription $SubscriptionId | Out-Null

$workspaceId = "/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroup/providers/Microsoft.OperationalInsights/workspaces/$WorkspaceName"
$appInsightsId = "/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroup/providers/Microsoft.Insights/components/$AppInsightsName"

$items = @(
    @{
        type = 1
        content = @{ json = "## MSGraphMCP Observability`nCorrelated telemetry for MCP requests, tools, and sessions." }
        name = 'text-intro'
    },
    @{
        type = 3
        content = @{
            version = 'KqlItem/1.0'
            query = @"
AppRequests
| where TimeGenerated > ago(30m)
| project TimeGenerated, Name, ResultCode, Success, DurationMs, OperationName
| order by TimeGenerated desc
"@
            size = 0
            title = 'Live Request Stream'
            queryType = 0
            resourceType = 'microsoft.operationalinsights/workspaces'
            crossComponentResources = @($workspaceId)
            visualization = 'table'
        }
        name = 'live-requests'
    },
    @{
        type = 3
        content = @{
            version = 'KqlItem/1.0'
            query = @"
AppRequests
| where TimeGenerated > ago(24h)
| extend ToolName = tostring(Properties['mcp.tool_name'])
| extend GraphSessionId = tostring(Properties['mcp.graph_session_id'])
| extend TransportSessionId = tostring(Properties['mcp.transport_session_id'])
| summarize Calls=count(), Failed=countif(Success == false) by ToolName, GraphSessionId, TransportSessionId
| order by Calls desc
"@
            size = 0
            title = 'MCP Correlation (Tool + Session)'
            queryType = 0
            resourceType = 'microsoft.operationalinsights/workspaces'
            crossComponentResources = @($workspaceId)
            visualization = 'table'
        }
        name = 'correlation'
    },
    @{
        type = 3
        content = @{
            version = 'KqlItem/1.0'
            query = @"
AppExceptions
| where TimeGenerated > ago(24h)
| project TimeGenerated, ExceptionType, OuterMessage, OperationName, ProblemId
| order by TimeGenerated desc
"@
            size = 0
            title = 'Exceptions Timeline'
            queryType = 0
            resourceType = 'microsoft.operationalinsights/workspaces'
            crossComponentResources = @($workspaceId)
            visualization = 'table'
        }
        name = 'exceptions'
    },
    @{
        type = 3
        content = @{
            version = 'KqlItem/1.0'
            query = @"
AppDependencies
| where TimeGenerated > ago(24h)
| where Success == false
| project TimeGenerated, Name, Target, DependencyType, ResultCode, DurationMs
| order by TimeGenerated desc
"@
            size = 0
            title = 'Dependency Failures'
            queryType = 0
            resourceType = 'microsoft.operationalinsights/workspaces'
            crossComponentResources = @($workspaceId)
            visualization = 'table'
        }
        name = 'dependencies'
    }
)

$serializedData = @{
    version = 'Notebook/1.0'
    items = $items
    isLocked = $false
} | ConvertTo-Json -Depth 60 -Compress

$resourceBody = @{
    location = $Location
    kind = 'shared'
    properties = @{
        displayName = $DisplayName
        serializedData = $serializedData
        version = '1.0'
        sourceId = $appInsightsId
        category = 'workbook'
        description = 'MSGraphMCP telemetry workbook with correlation fields.'
    }
} | ConvertTo-Json -Depth 80

$tmpBody = Join-Path $env:TEMP 'msgraphmcp-workbook.json'
$resourceBody | Out-File -FilePath $tmpBody -Encoding utf8

$url = 'https://management.azure.com/subscriptions/' +
    $SubscriptionId +
    '/resourceGroups/' +
    $ResourceGroup +
    '/providers/Microsoft.Insights/workbooks/' +
    $WorkbookId +
    '?api-version=2022-04-01'

az rest --method put --url $url --body "@$tmpBody" | Out-Null
if ($LASTEXITCODE -ne 0)
{
    throw 'Failed to create/update workbook via az rest.'
}

Write-Host ''
Write-Host 'Workbook deployed/updated successfully.' -ForegroundColor Green
Write-Host "Workbook Resource ID: /subscriptions/$SubscriptionId/resourceGroups/$ResourceGroup/providers/Microsoft.Insights/workbooks/$WorkbookId"
Write-Host "Azure Portal: https://portal.azure.com/#@/resource/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroup/providers/Microsoft.Insights/workbooks/$WorkbookId/overview"
