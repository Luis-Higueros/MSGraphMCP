using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions.Authentication;
using MSGraphMCP.Session;

namespace MSGraphMCP.Auth;

/// <summary>
/// Manages MSAL authentication against Azure AD using delegated (user) permissions.
///
/// One-time login strategy:
///   1. Caller provides a userHint (email). We check blob cache.
///   2. If cache hit → AcquireTokenSilent → no device code needed (truly one-time).
///   3. If cache miss → device code flow → user authenticates once → token saved to blob.
///   4. Refresh tokens (90-day sliding) ensure silent re-auth indefinitely while active.
/// </summary>
public class GraphAuthProvider
{
    private readonly BlobTokenCache _blobCache;
    private readonly ILogger<GraphAuthProvider> _logger;

    private readonly string   _tenantId;
    private readonly string   _clientId;
    private readonly string[] _scopes;

    public GraphAuthProvider(
        IConfiguration config,
        BlobTokenCache blobCache,
        ILogger<GraphAuthProvider> logger)
    {
        _blobCache = blobCache;
        _logger    = logger;
        _tenantId  = config["AzureAd:TenantId"]  ?? throw new InvalidOperationException("AzureAd:TenantId required");
        _clientId  = config["AzureAd:ClientId"]  ?? throw new InvalidOperationException("AzureAd:ClientId required");
        _scopes    = config.GetSection("AzureAd:Scopes").Get<string[]>()
                    ?? throw new InvalidOperationException("AzureAd:Scopes required");
    }

    /// <summary>
    /// Checks if a user already has a cached token — skips device code if they do.
    /// </summary>
    public Task<bool> HasCachedTokenAsync(string userHint) =>
        _blobCache.HasCachedTokenAsync(userHint);

    /// <summary>
    /// Attempts silent authentication using a cached refresh token.
    /// Returns a ready GraphServiceClient if successful; null if device code is required.
    /// </summary>
    public async Task<GraphServiceClient?> TrySilentAuthAsync(string userHint)
    {
        if (!await _blobCache.HasCachedTokenAsync(userHint))
            return null;

        var app = BuildApp(userHint);

        try
        {
            var accounts = await app.GetAccountsAsync();
            var account  = accounts.FirstOrDefault() 
                        ?? accounts.FirstOrDefault(a =>
                            a.Username.Equals(userHint, StringComparison.OrdinalIgnoreCase));

            var result = await app.AcquireTokenSilent(_scopes, account).ExecuteAsync();
            _logger.LogInformation("Silent auth succeeded for {User}. Token expires {Expiry}",
                userHint, result.ExpiresOn);

            return BuildGraphClient(app, _scopes);
        }
        catch (MsalUiRequiredException ex)
        {
            _logger.LogWarning("Silent auth failed for {User}, device code needed: {Msg}",
                userHint, ex.Message);
            return null;
        }
    }

    /// <summary>
    /// Starts the device code flow. Returns the DeviceCodeResult immediately so the caller
    /// can return the URL and user code to the end user (via Relevance AI UI step).
    /// The actual token acquisition completes asynchronously; poll SessionContext.IsAuthenticated.
    /// </summary>
    public async Task StartDeviceCodeFlowAsync(SessionContext session, string userHint)
    {
        var app = BuildApp(userHint);
        session.MsalApp  = app;
        session.UserHint = userHint;

        // Fire-and-forget: MSAL will call back when user completes auth
        _ = app.AcquireTokenWithDeviceCode(_scopes, async deviceCode =>
        {
            session.PendingDeviceCode = deviceCode;
            _logger.LogInformation("Device code issued for {User}. Expires: {Expiry}",
                userHint, deviceCode.ExpiresOn);
            await Task.CompletedTask;
        })
        .ExecuteAsync(session.CancellationToken)
        .ContinueWith(async task =>
        {
            if (task.IsCompletedSuccessfully)
            {
                _logger.LogInformation("Device code auth completed for {User}", userHint);
                await FinalizeSessionAsync(session, task.Result);
            }
            else
            {
                var ex = task.Exception?.InnerException;
                _logger.LogError(ex, "Device code auth failed for {User}", userHint);
                session.AuthError = ex?.Message ?? "Authentication failed";
            }
        }, TaskScheduler.Default);

        // Give MSAL 800ms to populate the device code result before we return
        var deadline = DateTime.UtcNow.AddSeconds(5);
        while (session.PendingDeviceCode is null && DateTime.UtcNow < deadline)
            await Task.Delay(100);
    }

    /// <summary>
    /// Called when the user completes device code auth. Wires up the GraphClient
    /// and schedules proactive token refresh.
    /// </summary>
    private async Task FinalizeSessionAsync(SessionContext session, AuthenticationResult result)
    {
        session.GraphClient    = BuildGraphClient(session.MsalApp!, _scopes);
        session.AuthenticatedAt = DateTimeOffset.UtcNow;

        // Schedule silent refresh 5 minutes before the access token expires.
        // The refresh token in the blob cache ensures this works indefinitely.
        ScheduleRefresh(session, result.ExpiresOn);

        await Task.CompletedTask;
    }

    /// <summary>
    /// Schedules a timer to silently refresh the access token before expiry.
    /// Keeps rescheduling itself, so sessions stay alive for as long as the user is active.
    /// </summary>
    private void ScheduleRefresh(SessionContext session, DateTimeOffset expiresOn)
    {
        var refreshIn = expiresOn - DateTimeOffset.UtcNow - TimeSpan.FromMinutes(5);
        if (refreshIn < TimeSpan.Zero) refreshIn = TimeSpan.Zero;

        session.RefreshTimer?.Dispose();
        session.RefreshTimer = new Timer(async _ =>
        {
            try
            {
                var accounts = await session.MsalApp!.GetAccountsAsync();
                var result   = await session.MsalApp
                    .AcquireTokenSilent(_scopes, accounts.FirstOrDefault())
                    .ExecuteAsync();

                _logger.LogInformation("Token silently refreshed for {User}. Next expiry: {Expiry}",
                    session.UserHint, result.ExpiresOn);

                // Reschedule for the new expiry
                ScheduleRefresh(session, result.ExpiresOn);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Silent token refresh failed for {User}", session.UserHint);
            }
        }, null, refreshIn, Timeout.InfiniteTimeSpan);
    }

    /// <summary>
    /// Clears a user's cached token from blob storage (full logout / revoke).
    /// </summary>
    public Task RevokeAsync(string userHint) => _blobCache.DeleteAsync(userHint);

    // ── Private builders ─────────────────────────────────────────────────────

    private IPublicClientApplication BuildApp(string userHint)
    {
        var app = PublicClientApplicationBuilder
            .Create(_clientId)
            .WithTenantId(_tenantId)
            .WithDefaultRedirectUri()
            .Build();

        // Wire the blob cache to this MSAL app instance
        _blobCache.Register(app.UserTokenCache, userHint);

        return app;
    }

    private static GraphServiceClient BuildGraphClient(IPublicClientApplication app, string[] scopes)
    {
        return new GraphServiceClient(
            new BaseBearerTokenAuthenticationProvider(
                new MsalTokenProvider(app, scopes)));
    }
}

/// <summary>
/// Bridges MSAL's token acquisition to Graph SDK's IAccessTokenProvider interface.
/// </summary>
public class MsalTokenProvider : IAccessTokenProvider
{
    private readonly IPublicClientApplication _app;
    private readonly string[] _scopes;

    public MsalTokenProvider(IPublicClientApplication app, string[] scopes)
    {
        _app = app;
        _scopes = scopes;
    }

    public async Task<string> GetAuthorizationTokenAsync(
        Uri uri,
        Dictionary<string, object>? additionalAuthenticationContext = null,
        CancellationToken cancellationToken = default)
    {
        var accounts = await _app.GetAccountsAsync();
        var account = accounts.FirstOrDefault()
            ?? throw new InvalidOperationException("No signed-in account available for Graph token acquisition.");

        var result = await _app
            .AcquireTokenSilent(_scopes, account)
            .ExecuteAsync(cancellationToken);
        return result.AccessToken;
    }

    public AllowedHostsValidator AllowedHostsValidator { get; } = new();
}
