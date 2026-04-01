using System.ComponentModel;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using ModelContextProtocol.Server;
using MSGraphMCP.Session;

namespace MSGraphMCP.Tools;

[McpServerToolType]
public class TeamsTools(SessionStore sessionStore, ILogger<TeamsTools> logger)
{
    [McpServerTool]
    [Description("Lists all Microsoft Teams the authenticated user is a member of.")]
    public async Task<object> TeamsListMyTeams(
        [Description("Active sessionId.")] string sessionId)
    {
        var ctx    = GetSession(sessionId);
        var result = await ctx.GraphClient!.Me.JoinedTeams.GetAsync(cfg =>
            cfg.QueryParameters.Select = ["id", "displayName", "description", "isArchived"]);

        return new
        {
            count = result?.Value?.Count ?? 0,
            teams = result?.Value?.Select(t => new
            {
                id          = t.Id,
                name        = t.DisplayName,
                description = t.Description,
                isArchived  = t.IsArchived
            })
        };
    }

    [McpServerTool]
    [Description("Lists all channels in a given Team.")]
    public async Task<object> TeamsListChannels(
        [Description("Active sessionId.")] string sessionId,
        [Description("The Team ID from TeamsListMyTeams.")] string teamId)
    {
        var ctx    = GetSession(sessionId);
        var result = await ctx.GraphClient!.Teams[teamId].Channels.GetAsync(cfg =>
            cfg.QueryParameters.Select = ["id", "displayName", "description", "membershipType"]);

        return new
        {
            teamId,
            count    = result?.Value?.Count ?? 0,
            channels = result?.Value?.Select(c => new
            {
                id             = c.Id,
                name           = c.DisplayName,
                description    = c.Description,
                membershipType = c.MembershipType?.ToString()
            })
        };
    }

    [McpServerTool]
    [Description("Retrieves recent messages from a Teams channel, with sender and content.")]
    public async Task<object> TeamsGetChannelMessages(
        [Description("Active sessionId.")] string sessionId,
        [Description("The Team ID.")] string teamId,
        [Description("The Channel ID from TeamsListChannels.")] string channelId,
        [Description("Max messages to retrieve (default 20, max 50).")] int maxMessages = 20)
    {
        var ctx    = GetSession(sessionId);
        var result = await ctx.GraphClient!.Teams[teamId].Channels[channelId].Messages
            .GetAsync(cfg =>
            {
                cfg.QueryParameters.Top    = Math.Clamp(maxMessages, 1, 50);
            });

        return new
        {
            teamId, channelId,
            count    = result?.Value?.Count ?? 0,
            messages = result?.Value?.Select(m => new
            {
                id          = m.Id,
                from        = m.From?.User?.DisplayName,
                date        = m.CreatedDateTime?.ToString("f"),
                subject     = m.Subject,
                body        = StripHtml(m.Body?.Content),
                importance  = m.Importance?.ToString()
            })
        };
    }

    [McpServerTool]
    [Description("Sends a message to a Teams channel.")]
    public async Task<object> TeamsSendChannelMessage(
        [Description("Active sessionId.")] string sessionId,
        [Description("The Team ID.")] string teamId,
        [Description("The Channel ID.")] string channelId,
        [Description("Message content (plain text or basic HTML).")] string content,
        [Description("Optional subject for the message.")] string? subject = null)
    {
        var ctx = GetSession(sessionId);

        var message = await ctx.GraphClient!.Teams[teamId].Channels[channelId].Messages
            .PostAsync(new ChatMessage
            {
                Body    = new() { ContentType = BodyType.Html, Content = content }
            });

        return new { status = "sent", messageId = message?.Id, teamId, channelId };
    }

    [McpServerTool]
    [Description("Lists the user's Teams chats (1:1 and group chats).")]
    public async Task<object> TeamsListChats(
        [Description("Active sessionId.")] string sessionId)
    {
        var ctx    = GetSession(sessionId);
        var result = await ctx.GraphClient!.Me.Chats.GetAsync(cfg =>
            cfg.QueryParameters.Select = ["id", "topic", "chatType", "lastUpdatedDateTime"]);

        return new
        {
            count = result?.Value?.Count ?? 0,
            chats = result?.Value?.Select(c => new
            {
                id          = c.Id,
                topic       = c.Topic,
                type        = c.ChatType?.ToString(),
                lastUpdated = c.LastUpdatedDateTime?.ToString("f")
            })
        };
    }

    [McpServerTool]
    [Description("Retrieves recent messages from a Teams chat (1:1 or group).")]
    public async Task<object> TeamsGetChatMessages(
        [Description("Active sessionId.")] string sessionId,
        [Description("The Chat ID from TeamsListChats.")] string chatId,
        [Description("Max messages to retrieve (default 20).")] int maxMessages = 20)
    {
        var ctx    = GetSession(sessionId);
        var result = await ctx.GraphClient!.Me.Chats[chatId].Messages
            .GetAsync(cfg =>
            {
                cfg.QueryParameters.Top    = Math.Clamp(maxMessages, 1, 50);
                cfg.QueryParameters.Select = ["id", "body", "from", "createdDateTime"];
            });

        return new
        {
            chatId,
            count    = result?.Value?.Count ?? 0,
            messages = result?.Value?.Select(m => new
            {
                id   = m.Id,
                from = m.From?.User?.DisplayName,
                date = m.CreatedDateTime?.ToString("f"),
                body = StripHtml(m.Body?.Content)
            })
        };
    }

    private static string? StripHtml(string? html)
    {
        if (html is null) return null;
        return System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", " ")
            .Replace("&nbsp;", " ").Trim();
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
