using System.ComponentModel;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Drives.Item.Items.Item.CreateLink;
using ModelContextProtocol.Server;
using MSGraphMCP.Session;

namespace MSGraphMCP.Tools;

/// <summary>
/// MCP tools for SharePoint document library operations.
/// Provides site search, drive listing, item navigation, content access, and sharing.
/// All tools require a sessionId (from graph_initiate_login) passed in X-Session-Id header.
/// 
/// USAGE EXAMPLES (tool_args only, sessionId added by MCP client):
/// 
/// 1. SharePointListSites
///    tool_args: {"query": "Finance", "maxResults": 20}
///    Returns: {query, count, sites: [{siteId, displayName, webUrl}]}
/// 
/// 2. SharePointListDrives
///    tool_args: {"siteId": "contoso.sharepoint.com,site-id-guid,web-id-guid"}
///    Returns: {siteId, count, drives: [{driveId, name, webUrl}]}
/// 
/// 3. SharePointListItems
///    tool_args: {"driveId": "b!...", "folderPath": "Shared Documents", "maxItems": 50}
///    Returns: {driveId, folderPath, count, items: [{itemId, name, isFolder, sizeKb, lastModifiedDateTime, webUrl}]}
/// 
/// 4. SharePointGetContent
///    tool_args: {"driveId": "b!...", "itemId": "01ABCD..."}
///    Returns: {driveId, itemId, contentLength, contentText}
/// 
/// 5. SharePointUploadText
///    tool_args: {"driveId": "b!...", "folderPath": "Reports", "fileName": "summary.txt", "content": "..."}
///    Returns: {driveId, folderPath, fileName, itemId, webUrl, size}
/// 
/// 6. SharePointCreateShareLink
///    tool_args: {"driveId": "b!...", "itemId": "01ABCD...", "linkType": "edit", "scope": "organization"}
///    Returns: {driveId, itemId, linkType, scope, webUrl}
/// </summary>
[McpServerToolType]
public class SharePointTools(SessionStore sessionStore, ILogger<SharePointTools> logger)
{
    // ── TOOL 1: SharePointListSites ───────────────────────────────────────────
    
    [McpServerTool]
    [Description("Search SharePoint sites by name or keyword. Returns site IDs and metadata.")]
    public async Task<object> SharePointListSites(
        [Description("Active sessionId.")] string sessionId,
        [Description("Search query for site name.")] string query,
        [Description("Max results to return (default 20).")] int maxResults = 20)
    {
        var ctx = GetSession(sessionId);
        var graph = ctx.GraphClient!;

        try
        {
            // Use Graph SDK Sites.GetAsync() with search parameter
            var sites = await graph.Sites.GetAsync(cfg =>
            {
                cfg.QueryParameters.Search = query;
                cfg.QueryParameters.Top = Math.Clamp(maxResults, 1, 100);
                cfg.QueryParameters.Select = ["id", "displayName", "webUrl"];
            });

            var items = sites?.Value ?? [];
            if (items.Count == 0)
                return new { query, count = 0, sites = Array.Empty<object>() };

            return new
            {
                query,
                count = items.Count,
                sites = items.Select(s => new
                {
                    siteId = s.Id,
                    displayName = s.DisplayName,
                    webUrl = s.WebUrl
                })
            };
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "SharePointListSites failed for query '{Query}'", query);
            return new { error = $"SharePointListSites failed: {ex.Message}" };
        }
    }

    // ── TOOL 2: SharePointListDrives ──────────────────────────────────────────

    [McpServerTool]
    [Description("List document libraries (drives) for a given SharePoint site.")]
    public async Task<object> SharePointListDrives(
        [Description("Active sessionId.")] string sessionId,
        [Description("Site ID from SharePointListSites.")] string siteId)
    {
        var ctx = GetSession(sessionId);
        var graph = ctx.GraphClient!;

        try
        {
            var drives = await graph.Sites[siteId].Drives.GetAsync(cfg =>
                cfg.QueryParameters.Select = ["id", "name", "webUrl"]);

            var items = drives?.Value ?? [];
            if (items.Count == 0)
                return new { siteId, count = 0, drives = Array.Empty<object>() };

            return new
            {
                siteId,
                count = items.Count,
                drives = items.Select(d => new
                {
                    driveId = d.Id,
                    name = d.Name,
                    webUrl = d.WebUrl
                })
            };
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "SharePointListDrives failed for siteId '{SiteId}'", siteId);
            return new { error = $"SharePointListDrives failed: {ex.Message}" };
        }
    }

    // ── TOOL 3: SharePointListItems ───────────────────────────────────────────

    [McpServerTool]
    [Description("List files and folders in a SharePoint drive at root or nested folder path.")]
    public async Task<object> SharePointListItems(
        [Description("Active sessionId.")] string sessionId,
        [Description("Drive ID from SharePointListDrives.")] string driveId,
        [Description("Folder path inside the drive (e.g., 'Folder/Subfolder'). Omit or empty for root.")]
        string? folderPath = null,
        [Description("Max items to return (default 50).")] int maxItems = 50)
    {
        var ctx = GetSession(sessionId);
        var graph = ctx.GraphClient!;

        try
        {
            DriveItemCollectionResponse? result;
            
            if (string.IsNullOrEmpty(folderPath))
            {
                // Root folder: access via Items["root"]
                result = await graph.Drives[driveId].Items["root"].Children.GetAsync(cfg =>
                {
                    cfg.QueryParameters.Top = Math.Clamp(maxItems, 1, 100);
                    cfg.QueryParameters.Select = ["id", "name", "size", "folder", "file", "lastModifiedDateTime", "webUrl"];
                });
            }
            else
            {
                // Navigate to folder path, then get children
                result = await graph.Drives[driveId].Items["root"]
                    .ItemWithPath(folderPath).Children.GetAsync(cfg =>
                    {
                        cfg.QueryParameters.Top = Math.Clamp(maxItems, 1, 100);
                        cfg.QueryParameters.Select = ["id", "name", "size", "folder", "file", "lastModifiedDateTime", "webUrl"];
                    });
            }

            var items = result?.Value ?? [];
            if (items.Count == 0)
                return new { driveId, folderPath = folderPath ?? "", count = 0, items = Array.Empty<object>() };

            return new
            {
                driveId,
                folderPath = folderPath ?? "",
                count = items.Count,
                items = items.Select(i => new
                {
                    itemId = i.Id,
                    name = i.Name,
                    isFolder = i.Folder is not null,
                    sizeKb = i.Size.HasValue ? (long?)(i.Size.Value / 1024) : null,
                    lastModifiedDateTime = i.LastModifiedDateTime?.ToString("f"),
                    webUrl = i.WebUrl
                })
            };
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "SharePointListItems failed for driveId '{DriveId}', folderPath '{FolderPath}'", 
                driveId, folderPath ?? "(root)");
            return new { error = $"SharePointListItems failed: {ex.Message}" };
        }
    }

    // ── TOOL 4: SharePointGetContent ──────────────────────────────────────────

    [McpServerTool]
    [Description("Download text content from a SharePoint file.")]
    public async Task<object> SharePointGetContent(
        [Description("Active sessionId.")] string sessionId,
        [Description("Drive ID.")] string driveId,
        [Description("Item ID from SharePointListItems.")] string itemId)
    {
        var ctx = GetSession(sessionId);
        var graph = ctx.GraphClient!;

        try
        {
            var stream = await graph.Drives[driveId].Items[itemId].Content.GetAsync();
            
            if (stream is null)
                return new { error = "File not found or could not be read." };

            using var reader = new StreamReader(stream);
            var content = await reader.ReadToEndAsync();

            // Truncate very large files for practical reasons
            var displayContent = content.Length > 100_000 
                ? content[..100_000] + "\n... [truncated at 100,000 chars]" 
                : content;

            return new
            {
                driveId,
                itemId,
                contentLength = content.Length,
                contentText = displayContent
            };
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "SharePointGetContent failed for driveId '{DriveId}', itemId '{ItemId}'", 
                driveId, itemId);
            return new { error = $"SharePointGetContent failed: {ex.Message}" };
        }
    }

    // ── TOOL 5: SharePointUploadText ──────────────────────────────────────────

    [McpServerTool]
    [Description("Upload or overwrite a text file in a SharePoint drive.")]
    public async Task<object> SharePointUploadText(
        [Description("Active sessionId.")] string sessionId,
        [Description("Drive ID from SharePointListDrives.")] string driveId,
        [Description("Folder path inside the drive (e.g., 'Folder/Subfolder'). Leave empty for root.")]
        string? folderPath = null,
        [Description("File name, e.g., 'notes.txt'.")] string? fileName = null,
        [Description("Text content to upload.")] string? content = null)
    {
        // Validate required fields
        if (string.IsNullOrWhiteSpace(fileName))
            return new { error = "fileName is required." };
        if (content is null)
            return new { error = "content is required." };

        var ctx = GetSession(sessionId);
        var graph = ctx.GraphClient!;

        try
        {
            // Build path: folderPath/fileName, properly encoded
            var path = CombinePath(folderPath, fileName);
            
            var bytes = System.Text.Encoding.UTF8.GetBytes(content);
            using var stream = new MemoryStream(bytes);

            var item = await graph.Drives[driveId].Items["root"].ItemWithPath(path).Content.PutAsync(stream);

            return new
            {
                driveId,
                folderPath = folderPath ?? "",
                fileName,
                itemId = item?.Id,
                webUrl = item?.WebUrl,
                size = item?.Size ?? 0
            };
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "SharePointUploadText failed for driveId '{DriveId}', path '{Path}/{FileName}'", 
                driveId, folderPath ?? "(root)", fileName);
            return new { error = $"SharePointUploadText failed: {ex.Message}" };
        }
    }

    // ── TOOL 6: SharePointCreateShareLink ─────────────────────────────────────

    [McpServerTool]
    [Description("Create a shareable link for a SharePoint file.")]
    public async Task<object> SharePointCreateShareLink(
        [Description("Active sessionId.")] string sessionId,
        [Description("Drive ID.")] string driveId,
        [Description("Item ID.")] string itemId,
        [Description("Link type: 'view' (read-only) or 'edit'. Default: 'view'.")] string linkType = "view",
        [Description("Scope: 'organization' or 'anonymous'. Default: 'organization'.")] string scope = "organization")
    {
        var ctx = GetSession(sessionId);
        var graph = ctx.GraphClient!;

        try
        {
            var link = await graph.Drives[driveId].Items[itemId].CreateLink.PostAsync(
                new CreateLinkPostRequestBody
                {
                    Type = linkType,
                    Scope = scope
                });

            return new
            {
                driveId,
                itemId,
                linkType,
                scope,
                webUrl = link?.Link?.WebUrl
            };
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "SharePointCreateShareLink failed for driveId '{DriveId}', itemId '{ItemId}'", 
                driveId, itemId);
            return new { error = $"SharePointCreateShareLink failed: {ex.Message}" };
        }
    }

    // ── Helper Methods ────────────────────────────────────────────────────────

    private SessionContext GetSession(string sessionId)
    {
        var ctx = sessionStore.Get(sessionId)
            ?? throw new InvalidOperationException("Session not found. Call graph_initiate_login.");
        if (!ctx.IsAuthenticated)
            throw new InvalidOperationException("Not authenticated yet.");
        return ctx;
    }

    /// <summary>
    /// Combines a folder path and file name, handling empty paths and proper escaping.
    /// Empty folderPath results in just fileName.
    /// </summary>
    private static string CombinePath(string? folderPath, string fileName)
    {
        if (string.IsNullOrEmpty(folderPath))
            return Uri.EscapeDataString(fileName);

        var segments = folderPath.Split('/', StringSplitOptions.RemoveEmptyEntries)
            .Select(Uri.EscapeDataString)
            .Concat([Uri.EscapeDataString(fileName)]);

        return string.Join("/", segments);
    }
}
