using MSGraphMCP.Auth;
using MSGraphMCP.Session;
using MSGraphMCP.Tools;

var builder = WebApplication.CreateBuilder(args);

// ── Configuration ─────────────────────────────────────────────────────────────
builder.Configuration
    .AddJsonFile("appsettings.json", optional: false)
    .AddJsonFile($"appsettings.{builder.Environment.EnvironmentName}.json", optional: true)
    .AddEnvironmentVariables(); // Env vars override appsettings (important for ACI)

// ── Logging ───────────────────────────────────────────────────────────────────
builder.Logging.ClearProviders();
builder.Logging.AddConsole();

// ── Core Services ─────────────────────────────────────────────────────────────
builder.Services.AddSingleton<BlobTokenCache>();
builder.Services.AddSingleton<GraphAuthProvider>();
builder.Services.AddSingleton<SessionStore>();
builder.Services.AddHostedService(sp => sp.GetRequiredService<SessionStore>());

// ── MCP Server (HTTP/SSE transport) ──────────────────────────────────────────
builder.Services
    .AddMcpServer()
    .WithHttpTransport()
    .WithTools<AuthTools>()
    .WithTools<MailTools>()
    .WithTools<CalendarTools>()
    .WithTools<TeamsTools>()
    .WithTools<FilesTools>()
    .WithTools<OneNoteTools>()
    .WithTools<PlannerTools>();

// ── Health check endpoint (used by ACI liveness probe) ───────────────────────
builder.Services.AddHealthChecks();

var app = builder.Build();

// ── Startup: ensure blob container exists ────────────────────────────────────
var blobCache = app.Services.GetRequiredService<BlobTokenCache>();
await blobCache.EnsureContainerAsync();

// ── Middleware ────────────────────────────────────────────────────────────────
app.UseRouting();

// Health check at /health
app.MapHealthChecks("/health");

// MCP endpoint at /mcp
app.MapMcp("/mcp");

// Root info endpoint
app.MapGet("/", () => new
{
    service   = "MSGraphMCP",
    version   = "1.0.0",
    status    = "running",
    endpoints = new { mcp = "/mcp", health = "/health" }
});

app.Logger.LogInformation("MSGraphMCP server starting on {Urls}", builder.Configuration["Urls"]);

await app.RunAsync();
