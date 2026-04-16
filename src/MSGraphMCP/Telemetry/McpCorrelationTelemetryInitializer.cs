using Microsoft.ApplicationInsights.Channel;
using Microsoft.ApplicationInsights.DataContracts;
using Microsoft.ApplicationInsights.Extensibility;
using Microsoft.AspNetCore.Http;

namespace MSGraphMCP.Telemetry;

public static class McpTelemetryKeys
{
    public const string TransportSessionId = "mcp.transport_session_id";
    public const string GraphSessionId = "mcp.graph_session_id";
    public const string ToolName = "mcp.tool_name";
    public const string Method = "mcp.method";
}

public sealed class McpCorrelationTelemetryInitializer(IHttpContextAccessor httpContextAccessor)
    : ITelemetryInitializer
{
    public void Initialize(ITelemetry telemetry)
    {
        if (telemetry is not ISupportProperties withProperties)
            return;

        var context = httpContextAccessor.HttpContext;
        if (context is null)
            return;

        AddProperty(withProperties, McpTelemetryKeys.TransportSessionId, context.Items[McpTelemetryKeys.TransportSessionId] as string);
        AddProperty(withProperties, McpTelemetryKeys.GraphSessionId, context.Items[McpTelemetryKeys.GraphSessionId] as string);
        AddProperty(withProperties, McpTelemetryKeys.ToolName, context.Items[McpTelemetryKeys.ToolName] as string);
        AddProperty(withProperties, McpTelemetryKeys.Method, context.Items[McpTelemetryKeys.Method] as string);

        if (telemetry is RequestTelemetry requestTelemetry)
        {
            AddProperty(withProperties, "mcp.path", context.Request.Path.ToString());
            requestTelemetry.Name ??= $"{context.Request.Method} {context.Request.Path}";
        }
    }

    private static void AddProperty(ISupportProperties telemetry, string key, string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return;

        if (!telemetry.Properties.ContainsKey(key))
            telemetry.Properties[key] = value;
    }
}
