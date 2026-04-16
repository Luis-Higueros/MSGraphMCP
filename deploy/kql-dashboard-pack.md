# MSGraphMCP KQL Dashboard Pack

Use these queries against the Log Analytics workspace linked to `ai-msgraphmcp-27992`.

## 1. Live request stream (last 30 min)

```kusto
AppRequests
| where TimeGenerated > ago(30m)
| project TimeGenerated, Name, ResultCode, Success, DurationMs, OperationName, AppRoleName
| order by TimeGenerated desc
```

## 2. Error summary by endpoint

```kusto
AppRequests
| where TimeGenerated > ago(24h)
| where Success == false or toint(ResultCode) >= 400
| summarize failures = count() by Name, ResultCode
| order by failures desc
```

## 3. Slow requests (P95)

```kusto
AppRequests
| where TimeGenerated > ago(24h)
| summarize p50=percentile(DurationMs, 50), p95=percentile(DurationMs, 95), p99=percentile(DurationMs, 99) by Name
| order by p95 desc
```

## 4. Exceptions timeline

```kusto
AppExceptions
| where TimeGenerated > ago(24h)
| project TimeGenerated, ExceptionType, OuterMessage, OperationName, ProblemId
| order by TimeGenerated desc
```

## 5. MCP request traces (initialize/tools)

```kusto
AppTraces
| where TimeGenerated > ago(24h)
| where Message has "/mcp" or Message has "initialize" or Message has "tools/call"
| project TimeGenerated, SeverityLevel, Message, OperationName
| order by TimeGenerated desc
```

## 6. Dependency failures (Graph/HTTP)

```kusto
AppDependencies
| where TimeGenerated > ago(24h)
| where Success == false
| project TimeGenerated, Name, Target, DependencyType, ResultCode, DurationMs
| order by TimeGenerated desc
```

## 7. Front Door health checks vs backend health endpoint

```kusto
AppRequests
| where TimeGenerated > ago(24h)
| where Name has "/health"
| summarize hits=count(), failures=countif(Success == false) by bin(TimeGenerated, 5m), AppRoleName
| order by TimeGenerated desc
```

## 8. Session/auth warnings from app logs

```kusto
AppTraces
| where TimeGenerated > ago(24h)
| where Message has "Session" or Message has "authenticated" or Message has "not authenticated"
| project TimeGenerated, SeverityLevel, Message, OperationName
| order by TimeGenerated desc
```

## 9. Relevance test window filter

Set the time range in query editor and use:

```kusto
union AppRequests, AppTraces, AppExceptions, AppDependencies
| where TimeGenerated > ago(2h)
| order by TimeGenerated desc
```

## 10. MCP correlation fields on requests (session/tool)

```kusto
AppRequests
| where TimeGenerated > ago(24h)
| extend ToolName = tostring(Properties["mcp.tool_name"])
| extend GraphSessionId = tostring(Properties["mcp.graph_session_id"])
| extend TransportSessionId = tostring(Properties["mcp.transport_session_id"])
| extend McpMethod = tostring(Properties["mcp.method"])
| where isnotempty(ToolName) or isnotempty(GraphSessionId) or isnotempty(TransportSessionId)
| project TimeGenerated, Name, ResultCode, Success, DurationMs, ToolName, GraphSessionId, TransportSessionId, McpMethod
| order by TimeGenerated desc
```

## 11. Filter by a specific Graph session or tool

```kusto
let targetSession = "<graph-session-id>";
let targetTool = "MailSummarize";
AppRequests
| where TimeGenerated > ago(24h)
| extend ToolName = tostring(Properties["mcp.tool_name"])
| extend GraphSessionId = tostring(Properties["mcp.graph_session_id"])
| where (targetSession == "<graph-session-id>" or GraphSessionId == targetSession)
	and (targetTool == "MailSummarize" or ToolName == targetTool)
| project TimeGenerated, Name, ResultCode, Success, DurationMs, ToolName, GraphSessionId
| order by TimeGenerated desc
```

## Suggested Workbook Tiles

- Request volume (requests count by 5m)
- Error rate (% failed requests)
- P95 latency by operation
- Top exceptions table
- Failed dependencies table
- MCP trace stream table

## Quick sanity query

```kusto
AppRequests
| where TimeGenerated > ago(10m)
| summarize count()
```

If this returns >0 after calling Relevance, telemetry ingestion is working.
