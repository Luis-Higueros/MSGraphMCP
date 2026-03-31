using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using Microsoft.Identity.Client;

namespace MSGraphMCP.Auth;

/// <summary>
/// Persists the MSAL token cache (including refresh tokens) to Azure Blob Storage.
/// This enables one-time login: even after container restarts, the server silently
/// re-acquires access tokens using the persisted refresh token (valid 90 days sliding).
/// Each user gets their own blob, keyed by their AAD account identifier.
/// </summary>
public class BlobTokenCache
{
    private readonly BlobContainerClient _containerClient;
    private readonly string _blobNamePrefix;
    private readonly ILogger<BlobTokenCache> _logger;
    private readonly SemaphoreSlim _lock = new(1, 1);

    public BlobTokenCache(IConfiguration config, ILogger<BlobTokenCache> logger)
    {
        _logger = logger;
        var connectionString = config["TokenCache:StorageConnectionString"]
            ?? throw new InvalidOperationException("TokenCache:StorageConnectionString is required.");
        var containerName = config["TokenCache:ContainerName"] ?? "msal-token-cache";
        _blobNamePrefix   = config["TokenCache:BlobNamePrefix"] ?? "mcp-user-";

        _containerClient = new BlobContainerClient(connectionString, containerName);
    }

    /// <summary>
    /// Ensures the blob container exists. Call once at startup.
    /// </summary>
    public async Task EnsureContainerAsync()
    {
        await _containerClient.CreateIfNotExistsAsync(PublicAccessType.None);
        _logger.LogInformation("Token cache container ready: {Container}", _containerClient.Name);
    }

    /// <summary>
    /// Registers MSAL's BeforeAccess / AfterAccess callbacks to wire up blob persistence.
    /// Call this after creating your IPublicClientApplication.
    /// </summary>
    public void Register(ITokenCache tokenCache, string cacheKey)
    {
        tokenCache.SetBeforeAccessAsync(async args =>
        {
            var bytes = await LoadAsync(cacheKey);
            if (bytes is not null)
                args.TokenCache.DeserializeMsalV3(bytes);
        });

        tokenCache.SetAfterAccessAsync(async args =>
        {
            if (args.HasStateChanged)
                await SaveAsync(cacheKey, args.TokenCache.SerializeMsalV3());
        });
    }

    /// <summary>
    /// Checks whether a cached token exists for this user. If yes, they won't need
    /// to go through device code flow again.
    /// </summary>
    public async Task<bool> HasCachedTokenAsync(string cacheKey)
    {
        try
        {
            var blob = _containerClient.GetBlobClient(BlobName(cacheKey));
            return await blob.ExistsAsync();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not check cache for key {Key}", cacheKey);
            return false;
        }
    }

    /// <summary>
    /// Deletes a user's cached token (logout / revoke).
    /// </summary>
    public async Task DeleteAsync(string cacheKey)
    {
        try
        {
            var blob = _containerClient.GetBlobClient(BlobName(cacheKey));
            await blob.DeleteIfExistsAsync();
            _logger.LogInformation("Token cache deleted for key {Key}", cacheKey);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not delete cache for key {Key}", cacheKey);
        }
    }

    // ── Private helpers ──────────────────────────────────────────────────────

    private async Task<byte[]?> LoadAsync(string cacheKey)
    {
        await _lock.WaitAsync();
        try
        {
            var blob = _containerClient.GetBlobClient(BlobName(cacheKey));
            if (!await blob.ExistsAsync()) return null;

            using var ms = new MemoryStream();
            await blob.DownloadToAsync(ms);
            _logger.LogDebug("Token cache loaded for key {Key} ({Bytes} bytes)", cacheKey, ms.Length);
            return ms.ToArray();
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to load token cache for key {Key}", cacheKey);
            return null;
        }
        finally
        {
            _lock.Release();
        }
    }

    private async Task SaveAsync(string cacheKey, byte[] data)
    {
        await _lock.WaitAsync();
        try
        {
            var blob = _containerClient.GetBlobClient(BlobName(cacheKey));
            using var ms = new MemoryStream(data);
            await blob.UploadAsync(ms, overwrite: true);
            _logger.LogDebug("Token cache saved for key {Key} ({Bytes} bytes)", cacheKey, data.Length);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to save token cache for key {Key}", cacheKey);
        }
        finally
        {
            _lock.Release();
        }
    }

    private string BlobName(string cacheKey) =>
        $"{_blobNamePrefix}{cacheKey.ToLowerInvariant().Replace("@", "_at_").Replace(".", "_")}.bin";
}
