using System.Diagnostics;
using System.Text.Json;
using Microsoft.ApplicationInsights.Extensibility;
using MSGraphMCP.Auth;
using MSGraphMCP.Session;
using MSGraphMCP.Telemetry;
using MSGraphMCP.Tools;

var builder = WebApplication.CreateBuilder(args);

// ── Configuration ─────────────────────────────────────────────────────────────
builder.Configuration
    .AddJsonFile("appsettings.json", optional: false)
    .AddJsonFile($"appsettings.{builder.Environment.EnvironmentName}.json", optional: true)
    .AddEnvironmentVariables(); // Env vars override appsettings (important for ACI)

var originAllowListEnabled = builder.Configuration.GetValue<bool>("Mcp:OriginAllowListEnabled", false);
var configuredOrigins = builder.Configuration.GetSection("Mcp:AllowedOrigins").Get<string[]>() ?? [];
var envOrigins = (Environment.GetEnvironmentVariable("MCP_ALLOWED_ORIGINS") ?? string.Empty)
    .Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
var allowedOrigins = configuredOrigins
    .Concat(envOrigins)
    .Where(static origin => !string.IsNullOrWhiteSpace(origin))
    .Distinct(StringComparer.OrdinalIgnoreCase)
    .ToHashSet(StringComparer.OrdinalIgnoreCase);

// ── Logging ───────────────────────────────────────────────────────────────────
builder.Logging.ClearProviders();
builder.Logging.AddConsole();

// ── Telemetry ────────────────────────────────────────────────────────────────
builder.Services.AddApplicationInsightsTelemetry();
builder.Services.AddHttpContextAccessor();
builder.Services.AddSingleton<ITelemetryInitializer, McpCorrelationTelemetryInitializer>();

// ── Core Services ─────────────────────────────────────────────────────────────
builder.Services.AddSingleton<BlobTokenCache>();
builder.Services.AddSingleton<GraphAuthProvider>();
builder.Services.AddSingleton<SessionStore>();
builder.Services.AddHostedService(sp => sp.GetRequiredService<SessionStore>());

var validatedTools = McpToolRegistration.CreateValidatedTools(
    builder.Services,
    typeof(AuthTools),
    typeof(MailTools),
    typeof(CalendarTools),
    typeof(TeamsTools),
    typeof(FilesTools),
    typeof(OneNoteTools),
    typeof(PlannerTools),
    typeof(SharePointTools));

// ── MCP Server (HTTP/SSE transport) ──────────────────────────────────────────
builder.Services
    .AddMcpServer()
    .WithHttpTransport()
    .WithTools(validatedTools);

// ── Health check endpoint (used by ACI liveness probe) ───────────────────────
builder.Services.AddHealthChecks();

var app = builder.Build();

if (originAllowListEnabled)
{
    app.Logger.LogInformation("MCP origin allow-list enforcement enabled with {Count} allowed origin(s).", allowedOrigins.Count);
}

var aiConnectionString = builder.Configuration["APPLICATIONINSIGHTS_CONNECTION_STRING"]
    ?? builder.Configuration["ApplicationInsights:ConnectionString"];
if (string.IsNullOrWhiteSpace(aiConnectionString))
{
    app.Logger.LogWarning("Application Insights connection string is not configured. Telemetry will not be sent.");
}
else
{
    app.Logger.LogInformation("Application Insights telemetry is enabled.");
}

// ── Startup: ensure blob container exists ────────────────────────────────────
var blobCache = app.Services.GetRequiredService<BlobTokenCache>();
await blobCache.EnsureContainerAsync();

// ── Middleware ────────────────────────────────────────────────────────────────
app.UseRouting();

static bool IsMcpPath(PathString path)
{
    return path.Equals("/mcp", StringComparison.OrdinalIgnoreCase)
        || path.Equals("/mcp/", StringComparison.OrdinalIgnoreCase);
}

// Mapping MCP/ as well as MCP since some clients will add the backslash to MCP endpoint.
app.Use(async (context, next) =>
{
    if (context.Request.Path.Equals("/mcp/", StringComparison.OrdinalIgnoreCase))
    {
        context.Request.Path = "/mcp";
    }

    await next();
});

// Optional origin allow-list for MCP endpoint. Disabled by default.
app.Use(async (context, next) =>
{
    if (!IsMcpPath(context.Request.Path))
    {
        await next();
        return;
    }

    var origin = context.Request.Headers.Origin.ToString();
    if (string.IsNullOrWhiteSpace(origin))
    {
        await next();
        return;
    }

    if (!originAllowListEnabled)
    {
        await next();
        return;
    }

    if (!allowedOrigins.Contains(origin))
    {
        app.Logger.LogWarning("MCP request denied for origin {Origin}; allow-list enforcement is enabled.", origin);
        context.Response.StatusCode = StatusCodes.Status403Forbidden;
        context.Response.ContentType = "application/json";
        await context.Response.WriteAsJsonAsync(new
        {
            error = "origin_not_allowed",
            origin
        });
        return;
    }

    await next();
});

// Capture MCP correlation fields for telemetry enrichment.
app.Use(async (context, next) =>
{
    if (!IsMcpPath(context.Request.Path))
    {
        await next();
        return;
    }

    var transportSessionId = context.Request.Headers["mcp-session-id"].ToString();
    if (!string.IsNullOrWhiteSpace(transportSessionId))
    {
        context.Items[McpTelemetryKeys.TransportSessionId] = transportSessionId;
        Activity.Current?.SetTag("mcp.transport_session_id", transportSessionId);
    }

    if (HttpMethods.IsPost(context.Request.Method)
        && context.Request.ContentType?.Contains("application/json", StringComparison.OrdinalIgnoreCase) == true)
    {
        context.Request.EnableBuffering();
        try
        {
            using var payload = await JsonDocument.ParseAsync(context.Request.Body, cancellationToken: context.RequestAborted);
            if (payload.RootElement.TryGetProperty("method", out var methodEl)
                && methodEl.ValueKind == JsonValueKind.String)
            {
                var method = methodEl.GetString();
                if (!string.IsNullOrWhiteSpace(method))
                {
                    context.Items[McpTelemetryKeys.Method] = method;
                    Activity.Current?.SetTag("mcp.method", method);
                }
            }

            if (payload.RootElement.TryGetProperty("params", out var paramsEl)
                && paramsEl.ValueKind == JsonValueKind.Object)
            {
                if (paramsEl.TryGetProperty("name", out var toolEl)
                    && toolEl.ValueKind == JsonValueKind.String)
                {
                    var toolName = toolEl.GetString();
                    if (!string.IsNullOrWhiteSpace(toolName))
                    {
                        context.Items[McpTelemetryKeys.ToolName] = toolName;
                        Activity.Current?.SetTag("mcp.tool_name", toolName);
                    }
                }

                if (paramsEl.TryGetProperty("arguments", out var argsEl)
                    && argsEl.ValueKind == JsonValueKind.Object
                    && argsEl.TryGetProperty("sessionId", out var graphSessionEl)
                    && graphSessionEl.ValueKind == JsonValueKind.String)
                {
                    var graphSessionId = graphSessionEl.GetString();
                    if (!string.IsNullOrWhiteSpace(graphSessionId))
                    {
                        context.Items[McpTelemetryKeys.GraphSessionId] = graphSessionId;
                        Activity.Current?.SetTag("mcp.graph_session_id", graphSessionId);
                    }
                }
            }
        }
        catch (JsonException)
        {
            // Ignore non-JSON payloads to avoid impacting MCP execution.
        }
        finally
        {
            if (context.Request.Body.CanSeek)
                context.Request.Body.Position = 0;
        }
    }

    await next();
});

// Compatibility shim for MCP connectors that probe with GET/OPTIONS before
// sending JSON-RPC initialize over POST.
app.Use(async (context, next) =>
{
    if (IsMcpPath(context.Request.Path))
    {
        if (HttpMethods.IsGet(context.Request.Method))
        {
            context.Response.StatusCode = StatusCodes.Status200OK;
            context.Response.ContentType = "application/json";
            await context.Response.WriteAsJsonAsync(new
            {
                service = "MSGraphMCP",
                transport = "streamable-http",
                path = "/mcp",
                initialize = "POST /mcp"
            });
            return;
        }

        if (HttpMethods.IsOptions(context.Request.Method))
        {
            context.Response.StatusCode = StatusCodes.Status204NoContent;
            context.Response.Headers.Append("Allow", "GET, POST, DELETE, OPTIONS");
            return;
        }
    }

    await next();
});

// Health check at /health
app.MapHealthChecks("/health");

// MCP endpoint at /mcp
app.MapMcp("/mcp");

// Lightweight smoke endpoint to validate delegated Graph scopes for an existing session.
app.MapPost("/test/scope-smoke", async (ScopeSmokeRequest request, SessionStore sessionStore) =>
{
    if (string.IsNullOrWhiteSpace(request.SessionId))
        return Results.BadRequest(new { error = "sessionId is required." });

    var ctx = sessionStore.Get(request.SessionId);
    if (ctx is null)
        return Results.NotFound(new { error = "Session not found or expired." });
    if (!ctx.IsAuthenticated || ctx.GraphClient is null)
        return Results.BadRequest(new { error = "Session is not authenticated yet." });

    var graph = ctx.GraphClient;
    var checks = new List<object>();
    var passed = 0;
    var failed = 0;
    var skipped = 0;

    async Task RunCheck(string name, Func<Task> probe)
    {
        try
        {
            await probe();
            checks.Add(new { check = name, status = "passed" });
            passed++;
        }
        catch (Exception ex)
        {
            checks.Add(new { check = name, status = "failed", error = ex.Message });
            failed++;
        }
    }

    Task SkipCheck(string name, string reason)
    {
        checks.Add(new { check = name, status = "skipped", reason });
        skipped++;
        return Task.CompletedTask;
    }

    await RunCheck("GraphWhoAmI", async () =>
    {
        await graph.Me.GetAsync(cfg =>
            cfg.QueryParameters.Select = ["displayName", "mail", "userPrincipalName"]);
    });

    await RunCheck("MailSearch", async () =>
    {
        var messages = await graph.Me.Messages.GetAsync(cfg =>
        {
            cfg.QueryParameters.Top = 1;
            cfg.QueryParameters.Select = ["id", "subject", "receivedDateTime"];
        });
        _ = messages?.Value?.FirstOrDefault();
    });

    await RunCheck("CalendarGetAgenda", async () =>
    {
        var mailboxSettings = await graph.Me.MailboxSettings.GetAsync();
        var timezone = mailboxSettings?.TimeZone ?? "UTC";
        var from = (request.FromUtc ?? DateTimeOffset.UtcNow.Date).UtcDateTime;
        var to = (request.ToUtc ?? DateTimeOffset.UtcNow.Date.AddDays(2)).UtcDateTime;

        var events = await graph.Me.CalendarView.GetAsync(cfg =>
        {
            cfg.QueryParameters.StartDateTime = from.ToString("o");
            cfg.QueryParameters.EndDateTime = to.ToString("o");
            cfg.QueryParameters.Top = 1;
            cfg.QueryParameters.Select = ["id", "subject", "start", "end"];
        });
        _ = timezone;
        _ = events?.Value?.FirstOrDefault();
    });

    await RunCheck("FilesListItems", async () =>
    {
        var drive = await graph.Me.Drive.GetAsync(cfg => cfg.QueryParameters.Select = ["id"]);
        var driveId = drive?.Id ?? throw new InvalidOperationException("Unable to resolve user drive.");
        var root = await graph.Drives[driveId].Root.GetAsync(cfg =>
            cfg.QueryParameters.Select = ["id"]);
        var rootId = root?.Id ?? throw new InvalidOperationException("Unable to resolve drive root id.");
        var rootItems = await graph.Drives[driveId].Items[rootId].Children.GetAsync(cfg =>
        {
            cfg.QueryParameters.Top = 1;
            cfg.QueryParameters.Select = ["id", "name"];
        });
        _ = rootItems?.Value?.FirstOrDefault();
    });

    string? teamId = request.TeamId;
    await RunCheck("TeamsListMyTeams", async () =>
    {
        var teams = await graph.Me.JoinedTeams.GetAsync(cfg =>
            cfg.QueryParameters.Select = ["id", "displayName"]);
        var first = teams?.Value?.FirstOrDefault();
        teamId ??= first?.Id;
        _ = first?.DisplayName;
    });

    if (string.IsNullOrWhiteSpace(teamId))
    {
        await SkipCheck("TeamsListChannels", "No teamId provided or discovered from TeamsListMyTeams.");
        await SkipCheck("TeamsGetChannelMessages", "No teamId provided or discovered from TeamsListMyTeams.");
    }
    else
    {
        string? channelId = request.ChannelId;
        await RunCheck("TeamsListChannels", async () =>
        {
            var channels = await graph.Teams[teamId].Channels.GetAsync(cfg =>
            {
                cfg.QueryParameters.Select = ["id", "displayName"];
            });
            var first = channels?.Value?.FirstOrDefault();
            channelId ??= first?.Id;
            _ = first?.DisplayName;
        });

        if (string.IsNullOrWhiteSpace(channelId))
        {
            await SkipCheck("TeamsGetChannelMessages", "No channelId provided or discovered from TeamsListChannels.");
        }
        else
        {
            await RunCheck("TeamsGetChannelMessages", async () =>
            {
                var messages = await graph.Teams[teamId].Channels[channelId].Messages.GetAsync(cfg =>
                {
                    cfg.QueryParameters.Top = 1;
                });
                _ = messages?.Value?.FirstOrDefault();
            });
        }
    }

    await RunCheck("OneNoteListNotebooks", async () =>
    {
        var notebooks = await graph.Me.Onenote.Notebooks.GetAsync(cfg =>
        {
            cfg.QueryParameters.Top = 1;
            cfg.QueryParameters.Select = ["id", "displayName"];
        });
        _ = notebooks?.Value?.FirstOrDefault();
    });

    await RunCheck("PlannerListPlans", async () =>
    {
        var groups = await graph.Me.MemberOf.GetAsync(cfg =>
            cfg.QueryParameters.Select = ["id", "displayName"]);

        var planCount = 0;
        string? firstPlanId = null;
        foreach (var group in groups?.Value?.OfType<Microsoft.Graph.Models.Group>() ?? [])
        {
            try
            {
                var plans = await graph.Groups[group.Id].Planner.Plans.GetAsync(cfg =>
                {
                    cfg.QueryParameters.Top = 1;
                    cfg.QueryParameters.Select = ["id", "title"];
                });
                var first = plans?.Value?.FirstOrDefault();
                if (first is not null)
                {
                    planCount++;
                    firstPlanId ??= first.Id;
                }
            }
            catch
            {
                // Some groups have no planner access; continue scanning.
            }
        }

        _ = planCount;
        _ = firstPlanId;
    });

    return Results.Ok(new
    {
        sessionId = request.SessionId,
        summary = new { passed, failed, skipped },
        checks
    });
});

// Root info endpoint
app.MapGet("/", () => new
{
    service   = "MSGraphMCP",
    version   = "1.0.0",
    status    = "running",
    endpoints = new { mcp = "/mcp", health = "/health", scopeSmoke = "/test/scope-smoke" }
});

app.Logger.LogInformation("MSGraphMCP server starting on {Urls}", builder.Configuration["Urls"]);

await app.RunAsync();

public sealed class ScopeSmokeRequest
{
    public string SessionId { get; set; } = string.Empty;
    public string? TeamId { get; set; }
    public string? ChannelId { get; set; }
    public DateTimeOffset? FromUtc { get; set; }
    public DateTimeOffset? ToUtc { get; set; }
}
