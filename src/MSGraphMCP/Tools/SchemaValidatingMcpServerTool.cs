using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace MSGraphMCP.Tools;

/// <summary>
/// Cross-cutting validator for tool arguments. Wraps any MCP tool and validates
/// against the generated JSON schema before invocation.
///
/// Self-check examples (expected validation_failed):
/// - unknown field: { sessionId: "...", foo: "bar" }
/// - schema stub: { sessionId: { "type": "string" } }
/// - wrong type:   { sessionId: "...", maxResults: "10" }
/// - bad format:   { since: "04-01-2026" }
/// </summary>
public sealed class SchemaValidatingMcpServerTool(McpServerTool innerTool) : DelegatingMcpServerTool(innerTool)
{
    private static readonly Regex DateRegex = new("^\\d{4}-\\d{2}-\\d{2}$", RegexOptions.Compiled);
    private static readonly Regex DateTimeRegex = new("^\\d{4}-\\d{2}-\\d{2}T\\d{2}:\\d{2}:\\d{2}$", RegexOptions.Compiled);
    private static readonly Regex TimeRegex = new("^\\d{2}:\\d{2}$", RegexOptions.Compiled);

    public override async ValueTask<CallToolResponse> InvokeAsync(
        RequestContext<CallToolRequestParams> request,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(request);

        if (request.Params is null)
        {
            return CreateValidationError(
                toolName: ProtocolTool.Name,
                missing: [],
                unexpected: [],
                typeErrors: ["request params are missing"],
                formatErrors: []);
        }

        var schema = GetSchemaObject(ProtocolTool.InputSchema);
        var properties = schema.TryGetProperty("properties", out var propsElement) && propsElement.ValueKind == JsonValueKind.Object
            ? propsElement
            : default;

        var allowedKeys = new HashSet<string>(StringComparer.Ordinal);
        if (properties.ValueKind == JsonValueKind.Object)
        {
            foreach (var p in properties.EnumerateObject())
                allowedKeys.Add(p.Name);
        }

        var required = new HashSet<string>(StringComparer.Ordinal);
        if (schema.TryGetProperty("required", out var requiredElement) && requiredElement.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in requiredElement.EnumerateArray())
            {
                if (item.ValueKind == JsonValueKind.String)
                    required.Add(item.GetString()!);
            }
        }

        var args = request.Params.Arguments is null
            ? new Dictionary<string, JsonElement>(StringComparer.Ordinal)
            : new Dictionary<string, JsonElement>(request.Params.Arguments, StringComparer.Ordinal);

        var missing = new List<string>();
        var unexpected = new List<string>();
        var typeErrors = new List<string>();
        var formatErrors = new List<string>();

        // Reject unknown keys.
        foreach (var key in args.Keys)
        {
            if (!allowedKeys.Contains(key))
                unexpected.Add(key);
        }

        // Validate required keys.
        foreach (var key in required)
        {
            if (!args.TryGetValue(key, out var value) || value.ValueKind == JsonValueKind.Null)
            {
                missing.Add(key);
                continue;
            }

            if (value.ValueKind == JsonValueKind.String)
            {
                var str = value.GetString();
                if (string.IsNullOrWhiteSpace(str))
                    missing.Add(key);
            }
        }

        // Reject schema stub argument values and validate declared types.
        foreach (var arg in args)
        {
            var key = arg.Key;
            var value = arg.Value;

            if (IsSchemaStub(value))
                typeErrors.Add($"{key}: schema stub object is not a valid argument value");

            if (properties.ValueKind != JsonValueKind.Object || !properties.TryGetProperty(key, out var propSchema))
                continue;

            if (!propSchema.TryGetProperty("type", out var typeElement) || typeElement.ValueKind != JsonValueKind.String)
                continue;

            var expectedType = typeElement.GetString();
            if (string.IsNullOrWhiteSpace(expectedType))
                continue;

            if (!MatchesType(value, expectedType!))
            {
                typeErrors.Add($"{key}: expected {expectedType}, got {value.ValueKind}");
            }
        }

        // Apply defaults only if property default exists, key is missing, and default != null.
        if (properties.ValueKind == JsonValueKind.Object)
        {
            foreach (var prop in properties.EnumerateObject())
            {
                if (args.ContainsKey(prop.Name))
                    continue;

                if (!prop.Value.TryGetProperty("default", out var defaultValue))
                    continue;

                if (defaultValue.ValueKind == JsonValueKind.Null)
                    continue;

                args.Add(prop.Name, defaultValue.Clone());
            }
        }

        // Format checks based on argument names.
        foreach (var arg in args)
        {
            if (arg.Value.ValueKind != JsonValueKind.String)
                continue;

            var key = arg.Key;
            var text = arg.Value.GetString();
            if (string.IsNullOrWhiteSpace(text))
                continue;

            if (IsDateName(key) && !DateRegex.IsMatch(text))
                formatErrors.Add($"{key}: expected yyyy-MM-dd");
            else if (IsDateTimeName(key) && !DateTimeRegex.IsMatch(text))
                formatErrors.Add($"{key}: expected yyyy-MM-ddTHH:mm:ss");
            else if (IsTimeName(key) && !TimeRegex.IsMatch(text))
                formatErrors.Add($"{key}: expected HH:mm");
        }

        if (missing.Count > 0 || unexpected.Count > 0 || typeErrors.Count > 0 || formatErrors.Count > 0)
        {
            return CreateValidationError(request.Params.Name, missing, unexpected, typeErrors, formatErrors);
        }

        var normalizedRequest = new RequestContext<CallToolRequestParams>(request.Server)
        {
            Services = request.Services,
            Params = new CallToolRequestParams
            {
                Name = request.Params.Name,
                Arguments = args
            }
        };

        return await base.InvokeAsync(normalizedRequest, cancellationToken);
    }

    private static JsonElement GetSchemaObject(object? schema)
    {
        if (schema is JsonElement element)
            return element;

        if (schema is JsonDocument doc)
            return doc.RootElement.Clone();

        if (schema is JsonNode node)
        {
            using var nodeDoc = JsonDocument.Parse(node.ToJsonString());
            return nodeDoc.RootElement.Clone();
        }

        using var fallbackDoc = JsonDocument.Parse(JsonSerializer.Serialize(schema));
        return fallbackDoc.RootElement.Clone();
    }

    private static bool IsSchemaStub(JsonElement value)
    {
        if (value.ValueKind != JsonValueKind.Object)
            return false;

        var properties = value.EnumerateObject().ToList();
        return properties.Count == 1 && properties[0].NameEquals("type");
    }

    private static bool MatchesType(JsonElement value, string expectedType)
    {
        return expectedType switch
        {
            "string" => value.ValueKind == JsonValueKind.String,
            "integer" => value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out _),
            "boolean" => value.ValueKind == JsonValueKind.True || value.ValueKind == JsonValueKind.False,
            _ => true // Keep compatibility for types not explicitly enforced here.
        };
    }

    private static bool IsDateName(string name) =>
        name.Equals("from", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("to", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("since", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("until", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("dueDate", StringComparison.OrdinalIgnoreCase);

    private static bool IsDateTimeName(string name) =>
        name.Equals("startTime", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("endTime", StringComparison.OrdinalIgnoreCase);

    private static bool IsTimeName(string name) =>
        name.Equals("workDayStart", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("workDayEnd", StringComparison.OrdinalIgnoreCase);

    private static CallToolResponse CreateValidationError(
        string toolName,
        List<string> missing,
        List<string> unexpected,
        List<string> typeErrors,
        List<string> formatErrors)
    {
        var payload = new
        {
            status = "validation_failed",
            tool = toolName,
            message = "Tool arguments failed schema validation.",
            missing,
            unexpected,
            typeErrors,
            formatErrors
        };

        return new CallToolResponse
        {
            IsError = true,
            Content =
            [
                new Content
                {
                    Type = "text",
                    Text = JsonSerializer.Serialize(payload)
                }
            ]
        };
    }
}
