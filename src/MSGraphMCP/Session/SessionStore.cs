using System.Collections.Concurrent;
using MSGraphMCP.Session;

namespace MSGraphMCP.Session;

/// <summary>
/// Thread-safe in-memory store for active sessions.
/// Sessions are keyed by sessionId (a GUID string).
/// Expired sessions are cleaned up by a background timer.
/// </summary>
public class SessionStore : IHostedService, IDisposable
{
    private readonly ConcurrentDictionary<string, SessionContext> _sessions = new();
    private readonly TimeSpan _sessionTimeout;
    private readonly ILogger<SessionStore> _logger;
    private Timer? _cleanupTimer;

    public SessionStore(IConfiguration config, ILogger<SessionStore> logger)
    {
        _logger = logger;
        var hours = config.GetValue("Session:TimeoutHours", 8);
        _sessionTimeout = TimeSpan.FromHours(hours);
    }

    public SessionContext Create()
    {
        var ctx = new SessionContext();
        _sessions[ctx.SessionId] = ctx;
        _logger.LogInformation("Session created: {SessionId}", ctx.SessionId);
        return ctx;
    }

    public SessionContext? Get(string sessionId)
    {
        if (!_sessions.TryGetValue(sessionId, out var ctx)) return null;

        if (ctx.LastAccessedAt.Add(_sessionTimeout) < DateTimeOffset.UtcNow)
        {
            Remove(sessionId);
            return null;
        }

        ctx.Touch();
        return ctx;
    }

    public void Remove(string sessionId)
    {
        if (_sessions.TryRemove(sessionId, out var ctx))
        {
            ctx.Dispose();
            _logger.LogInformation("Session removed: {SessionId}", sessionId);
        }
    }

    public int ActiveCount => _sessions.Count;

    // ── IHostedService: run cleanup every 15 minutes ──────────────────────

    public Task StartAsync(CancellationToken cancellationToken)
    {
        _cleanupTimer = new Timer(Cleanup, null, TimeSpan.FromMinutes(15), TimeSpan.FromMinutes(15));
        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken cancellationToken)
    {
        _cleanupTimer?.Change(Timeout.Infinite, 0);
        return Task.CompletedTask;
    }

    private void Cleanup(object? _)
    {
        var cutoff  = DateTimeOffset.UtcNow.Subtract(_sessionTimeout);
        var expired = _sessions
            .Where(kv => kv.Value.LastAccessedAt < cutoff)
            .Select(kv => kv.Key)
            .ToList();

        foreach (var id in expired)
            Remove(id);

        if (expired.Count > 0)
            _logger.LogInformation("Cleaned up {Count} expired sessions", expired.Count);
    }

    public void Dispose()
    {
        _cleanupTimer?.Dispose();
        foreach (var ctx in _sessions.Values) ctx.Dispose();
    }
}
