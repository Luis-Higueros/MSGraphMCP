using System.ComponentModel;
using Microsoft.Graph;
using ModelContextProtocol.Server;
using MSGraphMCP.Auth;
using MSGraphMCP.Session;

namespace MSGraphMCP.Tools;

/// <summary>
/// MCP tools that manage authentication lifecycle.
/// Relevance AI workflow:
///   1. Call graph_initiate_login (with userHint)
///      → If cached token: returns { status: "authenticated", sessionId }  ← skip to step 4
///      → If not cached:   returns { status: "pending", sessionId, verificationUrl, userCode }
///   2. Show user the verificationUrl + userCode
///   3. Poll graph_check_login_status(sessionId) until status == "authenticated"
///   4. Store sessionId; pass it in X-Session-Id header for all subsequent tool calls
/// </summary>
[McpServerToolType]
public class AuthTools(
    SessionStore sessionStore,
    GraphAuthProvider authProvider,
    ILogger<AuthTools> logger)
{
    [McpServerTool]
    [Description(
        "Initiates login to Microsoft 365 via the user's account. " +
        "If the user has logged in before, authentication is silent (no device code needed). " +
        "Returns a sessionId to use in all subsequent tool calls. " +
        "If first-time login, also returns a verificationUrl and userCode for the user to complete auth.")]
    public async Task<object> GraphInitiateLogin(
        [Description(
            "The user's Microsoft 365 email address. Used to find their cached credentials " +
            "so they don't have to log in again after the first time.")]
        string userHint)
    {
        userHint = userHint.Trim().ToLowerInvariant();

        // ── Fast path: silent auth from blob cache ────────────────────────
        var graphClient = await authProvider.TrySilentAuthAsync(userHint);
        if (graphClient is not null)
        {
            var session = sessionStore.Create();
            session.UserHint    = userHint;
            session.GraphClient = graphClient;
            session.AuthenticatedAt = DateTimeOffset.UtcNow;

            logger.LogInformation("Silent auth successful for {User}, session {Id}",
                userHint, session.SessionId);

            return new
            {
                status    = "authenticated",
                sessionId = session.SessionId,
                message   = $"Welcome back! Authenticated as {userHint} (no login required)."
            };
        }

        // ── Slow path: device code flow ───────────────────────────────────
        var ctx = sessionStore.Create();
        ctx.UserHint = userHint;

        await authProvider.StartDeviceCodeFlowAsync(ctx, userHint);

        if (ctx.PendingDeviceCode is null)
        {
            return new
            {
                status    = "error",
                sessionId = ctx.SessionId,
                message   = "Failed to initiate device code flow. Check server logs."
            };
        }

        return new
        {
            status          = "pending",
            sessionId       = ctx.SessionId,
            verificationUrl = ctx.PendingDeviceCode.VerificationUrl,
            userCode        = ctx.PendingDeviceCode.UserCode,
            expiresAt       = ctx.PendingDeviceCode.ExpiresOn,
            message         = $"Go to {ctx.PendingDeviceCode.VerificationUrl} and enter code: {ctx.PendingDeviceCode.UserCode}"
        };
    }

    [McpServerTool]
    [Description(
        "Polls the authentication status for a given sessionId. " +
        "Call this every 3–5 seconds after graph_initiate_login until status is 'authenticated'. " +
        "Once authenticated, the session is ready for all other Graph tools.")]
    public object GraphCheckLoginStatus(
        [Description("The sessionId returned by graph_initiate_login.")] string sessionId)
    {
        var ctx = sessionStore.Get(sessionId);

        if (ctx is null)
            return new { status = "not_found", message = "Session not found or expired. Call graph_initiate_login again." };

        if (ctx.AuthError is not null)
            return new { status = "error", message = ctx.AuthError };

        if (ctx.IsAuthenticated)
            return new
            {
                status      = "authenticated",
                sessionId,
                userHint    = ctx.UserHint,
                authenticatedAt = ctx.AuthenticatedAt
            };

        return new
        {
            status  = "pending",
            sessionId,
            message = "User has not completed authentication yet. Keep polling."
        };
    }

    [McpServerTool]
    [Description(
        "Logs out the current session and optionally revokes the cached credentials " +
        "so that graph_initiate_login will require a new device code next time.")]
    public async Task<object> GraphLogout(
        [Description("The sessionId to terminate.")] string sessionId,
        [Description(
            "If true, also deletes the persisted token cache from Azure Blob Storage. " +
            "The user will need to complete device code flow again next time. Default: false.")]
        bool revokeCache = false)
    {
        var ctx = sessionStore.Get(sessionId);
        if (ctx is null)
            return new { status = "not_found" };

        var hint = ctx.UserHint;
        sessionStore.Remove(sessionId);

        if (revokeCache && !string.IsNullOrEmpty(hint))
        {
            await authProvider.RevokeAsync(hint);
            return new { status = "logged_out", message = $"Session ended and token cache revoked for {hint}." };
        }

        return new { status = "logged_out", message = "Session ended. Cached token retained for next silent login." };
    }

    [McpServerTool]
    [Description("Returns info about the current session and authenticated user.")]
    public async Task<object> GraphWhoAmI(
        [Description("The active sessionId.")] string sessionId)
    {
        var ctx = sessionStore.Get(sessionId);
        if (ctx is null)
            return new { status = "not_found" };
        if (!ctx.IsAuthenticated)
            return new { status = "not_authenticated" };

        var me = await ctx.GraphClient!.Me.GetAsync(cfg =>
            cfg.QueryParameters.Select = ["displayName", "mail", "userPrincipalName", "timeZone"]);

        return new
        {
            status            = "authenticated",
            displayName       = me?.DisplayName,
            email             = me?.Mail ?? me?.UserPrincipalName,
            sessionId,
            authenticatedAt   = ctx.AuthenticatedAt,
            lastAccessedAt    = ctx.LastAccessedAt
        };
    }
}
