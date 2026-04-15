using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace MSGraphMCP.Session;

/// <summary>
/// Holds all state for a single authenticated session.
/// One SessionContext per user login in the in-memory SessionStore.
/// </summary>
public class SessionContext : IDisposable
{
    public string SessionId { get; init; } = Guid.NewGuid().ToString("N");

    /// <summary>The email hint used to key the blob token cache.</summary>
    public string UserHint { get; set; } = string.Empty;

    /// <summary>The MSAL application instance (one per session, shares blob cache).</summary>
    public IPublicClientApplication? MsalApp { get; set; }

    /// <summary>Ready-to-use Graph client. Null until auth completes.</summary>
    public GraphServiceClient? GraphClient { get; set; }

    /// <summary>Device code info to show to the end user during first-time login.</summary>
    public DeviceCodeResult? PendingDeviceCode { get; set; }

    /// <summary>Set if auth fails during device code flow.</summary>
    public string? AuthError { get; set; }

    public bool IsAuthenticated => GraphClient is not null;
    public bool IsPending       => !IsAuthenticated && AuthError is null;

    public DateTimeOffset CreatedAt      { get; init; } = DateTimeOffset.UtcNow;
    public DateTimeOffset? AuthenticatedAt { get; set; }
    public DateTimeOffset LastAccessedAt { get; set; } = DateTimeOffset.UtcNow;

    /// <summary>Proactive token refresh timer.</summary>
    public Timer? RefreshTimer { get; set; }

    /// <summary>Cancellation for in-flight device code acquisition.</summary>
    public CancellationTokenSource CancellationTokenSource { get; } = new();
    public CancellationToken CancellationToken => CancellationTokenSource.Token;

    public void Touch() => LastAccessedAt = DateTimeOffset.UtcNow;

    public void Dispose()
    {
        RefreshTimer?.Dispose();
        CancellationTokenSource.Cancel();
        CancellationTokenSource.Dispose();
    }
}
