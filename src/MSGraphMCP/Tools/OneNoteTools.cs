using System.ComponentModel;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using ModelContextProtocol.Server;
using MSGraphMCP.Session;

namespace MSGraphMCP.Tools;

[McpServerToolType]
public class OneNoteTools(SessionStore sessionStore, ILogger<OneNoteTools> logger)
{
    [McpServerTool]
    [Description("Lists all OneNote notebooks accessible by the user.")]
    public async Task<object> OneNoteListNotebooks(
        [Description("Active sessionId.")] string sessionId)
    {
        var ctx    = GetSession(sessionId);
        var result = await ctx.GraphClient!.Me.Onenote.Notebooks.GetAsync(cfg =>
            cfg.QueryParameters.Select = ["id", "displayName", "lastModifiedDateTime", "isShared"]);

        return new
        {
            count     = result?.Value?.Count ?? 0,
            notebooks = result?.Value?.Select(n => new
            {
                id       = n.Id,
                name     = n.DisplayName,
                modified = n.LastModifiedDateTime?.ToString("f"),
                isShared = n.IsShared
            })
        };
    }

    [McpServerTool]
    [Description("Lists all sections within a OneNote notebook.")]
    public async Task<object> OneNoteListSections(
        [Description("Active sessionId.")] string sessionId,
        [Description("Notebook ID from OneNoteListNotebooks.")] string notebookId)
    {
        var ctx    = GetSession(sessionId);
        var result = await ctx.GraphClient!.Me.Onenote.Notebooks[notebookId].Sections
            .GetAsync(cfg =>
                cfg.QueryParameters.Select = ["id", "displayName", "lastModifiedDateTime"]);

        return new
        {
            notebookId,
            count    = result?.Value?.Count ?? 0,
            sections = result?.Value?.Select(s => new
            {
                id       = s.Id,
                name     = s.DisplayName,
                modified = s.LastModifiedDateTime?.ToString("f")
            })
        };
    }

    [McpServerTool]
    [Description("Lists pages in a OneNote section, with title and last modified date.")]
    public async Task<object> OneNoteListPages(
        [Description("Active sessionId.")] string sessionId,
        [Description("Section ID from OneNoteListSections.")] string sectionId,
        [Description("Max pages to return (default 30).")] int maxPages = 30)
    {
        var ctx    = GetSession(sessionId);
        var result = await ctx.GraphClient!.Me.Onenote.Sections[sectionId].Pages
            .GetAsync(cfg =>
            {
                cfg.QueryParameters.Top    = maxPages;
                cfg.QueryParameters.Select = ["id", "title", "lastModifiedDateTime", "createdDateTime"];
                cfg.QueryParameters.Orderby = ["lastModifiedDateTime desc"];
            });

        return new
        {
            sectionId,
            count = result?.Value?.Count ?? 0,
            pages = result?.Value?.Select(p => new
            {
                id       = p.Id,
                title    = p.Title,
                created  = p.CreatedDateTime?.ToString("f"),
                modified = p.LastModifiedDateTime?.ToString("f")
            })
        };
    }

    [McpServerTool]
    [Description("Retrieves the plain text content of a OneNote page.")]
    public async Task<object> OneNoteGetPageContent(
        [Description("Active sessionId.")] string sessionId,
        [Description("Page ID from OneNoteListPages.")] string pageId)
    {
        var ctx    = GetSession(sessionId);
        var stream = await ctx.GraphClient!.Me.Onenote.Pages[pageId].Content.GetAsync();

        if (stream is null)
            return new { error = "Could not retrieve page content." };

        using var reader = new StreamReader(stream);
        var html    = await reader.ReadToEndAsync();
        // Strip HTML tags for clean text output
        var text = System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", " ")
            .Replace("&nbsp;", " ")
            .Replace("&amp;", "&")
            .Replace("&lt;", "<")
            .Replace("&gt;", ">")
            .Trim();

        return new
        {
            pageId,
            contentLength = text.Length,
            content       = text.Length > 50_000 ? text[..50_000] + "\n...[truncated]" : text
        };
    }

    [McpServerTool]
    [Description("Searches across all OneNote pages for a keyword or phrase.")]
    public async Task<object> OneNoteSearchPages(
        [Description("Active sessionId.")] string sessionId,
        [Description("Search query.")] string query,
        [Description("Max results (default 20).")] int maxResults = 20)
    {
        var ctx    = GetSession(sessionId);
        var result = await ctx.GraphClient!.Me.Onenote.Pages.GetAsync(cfg =>
        {
            cfg.QueryParameters.Search = query;
            cfg.QueryParameters.Top    = maxResults;
            cfg.QueryParameters.Select = ["id", "title", "lastModifiedDateTime", "parentSection"];
        });

        return new
        {
            query,
            count   = result?.Value?.Count ?? 0,
            results = result?.Value?.Select(p => new
            {
                id       = p.Id,
                title    = p.Title,
                modified = p.LastModifiedDateTime?.ToString("f"),
                section  = p.ParentSection?.DisplayName
            })
        };
    }

    [McpServerTool]
    [Description("Creates a new OneNote page in a specified section.")]
    public async Task<object> OneNoteCreatePage(
        [Description("Active sessionId.")] string sessionId,
        [Description("Section ID where the page will be created.")] string sectionId,
        [Description("Page title.")] string title,
        [Description("Page content in plain text.")] string content)
    {
        var ctx = GetSession(sessionId);

        try
        {
            var page = await ctx.GraphClient!.Me.Onenote.Sections[sectionId].Pages
                .PostAsync(new OnenotePage
                {
                    Title = title
                });

            return new
            {
                status = page is null ? "error" : "created",
                message = page is null ? "Graph returned no page object." : "Page created.",
                sectionId,
                title,
                contentLength = content.Length,
                pageId = page?.Id,
                webUrl = page?.Links?.OneNoteWebUrl?.Href
            };
        }
        catch (Exception ex)
        {
            return new
            {
                status = "error",
                message = "Failed to create OneNote page.",
                details = ex.Message,
                sectionId,
                title,
                contentLength = content.Length
            };
        }
    }

    private SessionContext GetSession(string sessionId)
    {
        var ctx = sessionStore.Get(sessionId)
            ?? throw new InvalidOperationException("Session not found. Call graph_initiate_login.");
        if (!ctx.IsAuthenticated)
            throw new InvalidOperationException("Not authenticated yet.");
        return ctx;
    }
}
