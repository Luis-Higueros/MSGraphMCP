using System.ComponentModel;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using ModelContextProtocol.Server;
using MSGraphMCP.Session;

namespace MSGraphMCP.Tools;

[McpServerToolType]
public class FilesTools(SessionStore sessionStore, ILogger<FilesTools> logger)
{
    [McpServerTool]
    [Description("Lists files and folders in the user's OneDrive root or a specific folder.")]
    public async Task<object> FilesListItems(
        [Description("Active sessionId.")] string sessionId,
        [Description("Folder path relative to root, e.g. 'Documents/Projects'. Leave empty for root.")] string? folderPath = null,
        [Description("Max items to return (default 50).")]
        int maxItems = 50)
    {
        var ctx = GetSession(sessionId);
        var root = await GetUserDriveRootItemAsync(ctx.GraphClient!);

        DriveItemCollectionResponse? result;
        if (string.IsNullOrEmpty(folderPath))
        {
            result = await root.Children.GetAsync(cfg =>
            {
                cfg.QueryParameters.Top = maxItems;
                cfg.QueryParameters.Select = ["id", "name", "size", "folder", "file", "lastModifiedDateTime", "webUrl"];
            });
        }
        else
        {
            result = await root.ItemWithPath(folderPath).Children.GetAsync(cfg =>
            {
                cfg.QueryParameters.Top = maxItems;
                cfg.QueryParameters.Select = ["id", "name", "size", "folder", "file", "lastModifiedDateTime", "webUrl"];
            });
        }

        return new
        {
            folderPath = folderPath ?? "/",
            count = result?.Value?.Count ?? 0,
            items = result?.Value?.Select(i => new
            {
                id = i.Id,
                name = i.Name,
                type = i.Folder is not null ? "folder" : "file",
                sizeKb = i.Size.HasValue ? (long?)(i.Size.Value / 1024) : null,
                modified = i.LastModifiedDateTime?.ToString("f"),
                webUrl = i.WebUrl
            })
        };
    }

    [McpServerTool]
    [Description("Downloads the text content of a file from OneDrive. Best for text, markdown, CSV, and similar files.")]
    public async Task<object> FilesGetContent(
        [Description("Active sessionId.")] string sessionId,
        [Description("File path relative to root, e.g. 'Documents/report.txt'.")]
        string filePath)
    {
        var ctx = GetSession(sessionId);
        var root = await GetUserDriveRootItemAsync(ctx.GraphClient!);
        var stream = await root.ItemWithPath(filePath).Content.GetAsync();

        if (stream is null)
            return new { error = "File not found or could not be read." };

        using var reader = new StreamReader(stream);
        var content = await reader.ReadToEndAsync();

        return new
        {
            filePath,
            contentLength = content.Length,
            content = content.Length > 50_000 ? content[..50_000] + "\n... [truncated at 50,000 chars]" : content
        };
    }

    [McpServerTool]
    [Description("Uploads a text file to the user's OneDrive. Creates or overwrites the file at the given path.")]
    public async Task<object> FilesUploadText(
        [Description("Active sessionId.")] string sessionId,
        [Description("Destination file path, e.g. 'Documents/output.txt'.")]
        string filePath,
        [Description("Text content to write to the file.")] string content)
    {
        var ctx = GetSession(sessionId);
        var root = await GetUserDriveRootItemAsync(ctx.GraphClient!);
        var bytes = System.Text.Encoding.UTF8.GetBytes(content);
        using var stream = new MemoryStream(bytes);

        var item = await root.ItemWithPath(filePath).Content.PutAsync(stream);

        return new
        {
            status = "uploaded",
            filePath,
            fileId = item?.Id,
            sizeBytes = bytes.Length,
            webUrl = item?.WebUrl
        };
    }

    [McpServerTool]
    [Description("Searches for files across the user's OneDrive and SharePoint by name or content keywords.")]
    public async Task<object> FilesSearch(
        [Description("Active sessionId.")] string sessionId,
        [Description("Search query - matches file names and content.")] string query,
        [Description("Max results (default 20).")]
        int maxResults = 20)
    {
        var ctx = GetSession(sessionId);
        var drive = await GetUserDriveAsync(ctx.GraphClient!);
        var driveBuilder = ctx.GraphClient!.Drives[drive.Id!];

        List<DriveItem> items;
        try
        {
            var response = await driveBuilder.SearchWithQ(query).GetAsSearchWithQGetResponseAsync(cfg =>
            {
                cfg.QueryParameters.Top = maxResults;
                cfg.QueryParameters.Select = ["id", "name", "size", "lastModifiedDateTime", "webUrl", "parentReference"];
            });
            items = response?.Value?.ToList() ?? [];
        }
        catch (ODataError)
        {
            items = [];
        }

        return new
        {
            query,
            count = items.Count,
            results = items.Select(i => new
            {
                id = i.Id,
                name = i.Name,
                sizeKb = i.Size.HasValue ? (long?)(i.Size.Value / 1024) : null,
                modified = i.LastModifiedDateTime?.ToString("f"),
                path = i.ParentReference?.Path,
                webUrl = i.WebUrl
            })
        };
    }

    [McpServerTool]
    [Description("Gets a shareable link for a file in OneDrive.")]
    public async Task<object> FilesCreateShareLink(
        [Description("Active sessionId.")] string sessionId,
        [Description("File path in OneDrive, e.g. 'Documents/report.pdf'.")]
        string filePath,
        [Description("Link type: 'view' (read-only) or 'edit' (read-write). Default: 'view'.")]
        string linkType = "view",
        [Description("Scope: 'anonymous' (anyone with link) or 'organization'. Default: 'organization'.")]
        string scope = "organization")
    {
        var ctx = GetSession(sessionId);
        var root = await GetUserDriveRootItemAsync(ctx.GraphClient!);
        var link = await root.ItemWithPath(filePath)
            .CreateLink.PostAsync(new()
            {
                Type = linkType,
                Scope = scope
            });

        return new
        {
            filePath,
            shareUrl = link?.Link?.WebUrl,
            linkType,
            scope
        };
    }

    private SessionContext GetSession(string sessionId)
    {
        var ctx = sessionStore.Get(sessionId)
            ?? throw new InvalidOperationException("Session not found. Call graph_initiate_login.");
        if (!ctx.IsAuthenticated)
            throw new InvalidOperationException("Not authenticated yet.");
        return ctx;
    }

    private static async Task<Drive> GetUserDriveAsync(GraphServiceClient graph)
    {
        var drive = await graph.Me.Drive.GetAsync(cfg => cfg.QueryParameters.Select = ["id"])
            ?? throw new InvalidOperationException("Unable to resolve user's drive.");
        if (string.IsNullOrWhiteSpace(drive.Id))
            throw new InvalidOperationException("Unable to resolve user's drive id.");
        return drive;
    }

    private static async Task<Microsoft.Graph.Drives.Item.Items.Item.DriveItemItemRequestBuilder> GetUserDriveRootItemAsync(GraphServiceClient graph)
    {
        var drive = await GetUserDriveAsync(graph);
        return graph.Drives[drive.Id!].Items["root"];
    }
}
