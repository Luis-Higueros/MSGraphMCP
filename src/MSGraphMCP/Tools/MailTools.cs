using System.ComponentModel;
using System.Text;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using ModelContextProtocol.Server;
using MSGraphMCP.Session;

namespace MSGraphMCP.Tools;

[McpServerToolType]
public class MailTools(SessionStore sessionStore, ILogger<MailTools> logger)
{
    // ── Search & Summarize ────────────────────────────────────────────────────

    [McpServerTool]
    [Description(
        "Search emails by keywords, sender, date range, or subject. " +
        "Supports looking back across weeks or months. " +
        "Returns structured results with IDs that can be used in other mail tools.")]
    public async Task<object> MailSearch(
        [Description("Active sessionId.")] string sessionId,
        [Description("Keywords to search in subject and body. Supports phrases in quotes.")] string? keywords = null,
        [Description("Filter by sender email address.")] string? fromAddress = null,
        [Description("Filter by email address in To field.")] string? toAddress = null,
        [Description("Start date (ISO 8601), e.g. '2025-03-01'. Can look back up to 12 months.")] string? since = null,
        [Description("End date (ISO 8601), e.g. '2025-03-31'. Defaults to today.")] string? until = null,
        [Description("Only return emails with attachments.")] bool hasAttachments = false,
        [Description("Max results to return (default 25, max 100).")] int maxResults = 25,
        [Description("Return full body text instead of preview (slower but more complete).")] bool includeBody = false)
    {
        var ctx = GetSession(sessionId);
        var graph = ctx.GraphClient!;

        // Build OData $filter
        var filters = new List<string>();
        if (since is not null)
            filters.Add($"receivedDateTime ge {DateTime.Parse(since).ToUniversalTime():o}");
        if (until is not null)
            filters.Add($"receivedDateTime le {DateTime.Parse(until).ToUniversalTime():o}");
        if (fromAddress is not null)
            filters.Add($"from/emailAddress/address eq '{fromAddress}'");
        if (toAddress is not null)
            filters.Add($"toRecipients/any(t: t/emailAddress/address eq '{toAddress}')");
        if (hasAttachments)
            filters.Add("hasAttachments eq true");

        var selectFields = new[] { "id", "subject", "from", "toRecipients", "receivedDateTime",
                                   "bodyPreview", "hasAttachments", "conversationId", "isRead" };
        if (includeBody) selectFields = [.. selectFields, "body"];

        var messages = await graph.Me.Messages.GetAsync(cfg =>
        {
            if (keywords is not null)
                cfg.QueryParameters.Search = $"\"{keywords}\"";
            if (filters.Count > 0)
                cfg.QueryParameters.Filter = string.Join(" and ", filters);
            cfg.QueryParameters.Top     = Math.Clamp(maxResults, 1, 100);
            cfg.QueryParameters.Select  = selectFields;
            cfg.QueryParameters.Orderby = ["receivedDateTime desc"];
        });

        var items = messages?.Value ?? [];
        if (items.Count == 0)
            return new { count = 0, message = "No emails matched your search.", emails = Array.Empty<object>() };

        return new
        {
            count  = items.Count,
            emails = items.Select((m, i) => new
            {
                index          = i + 1,
                id             = m.Id,
                conversationId = m.ConversationId,
                from           = m.From?.EmailAddress?.Address,
                subject        = m.Subject,
                date           = m.ReceivedDateTime?.ToString("f"),
                isRead         = m.IsRead,
                hasAttachments = m.HasAttachments,
                preview        = m.BodyPreview,
                body           = includeBody ? m.Body?.Content : null
            })
        };
    }

    [McpServerTool]
    [Description(
        "Search emails and return a structured summary grouped by context or theme. " +
        "Ideal for 'what happened with project X last month?' queries. " +
        "Returns a data payload designed to be summarized by an LLM in the next step.")]
    public async Task<object> MailSummarize(
        [Description("Active sessionId.")] string sessionId,
        [Description("The context or question to focus the summary on, e.g. 'project Alpha budget discussions'.")] string context,
        [Description("Keywords to search for. Defaults to words extracted from context.")] string? keywords = null,
        [Description("Sender filter.")] string? fromAddress = null,
        [Description("Start date (ISO 8601).")] string? since = null,
        [Description("End date (ISO 8601).")] string? until = null,
        [Description("Max emails to include (default 30).")] int maxEmails = 30)
    {
        var searchKeywords = keywords ?? context;
        var raw = await MailSearch(sessionId, searchKeywords, fromAddress, null, since, until,
                                   false, maxEmails, false);

        // Return structured payload for LLM to summarize
        return new
        {
            summarizationRequest = true,
            context,
            instructions = $"Summarize the following emails in the context of: \"{context}\". " +
                           "Group related threads. Highlight key decisions, action items, and people involved.",
            data = raw
        };
    }

    // ── Thread & Detail ───────────────────────────────────────────────────────

    [McpServerTool]
    [Description("Retrieves all messages in an email conversation thread, in chronological order.")]
    public async Task<object> MailGetThread(
        [Description("Active sessionId.")] string sessionId,
        [Description("The conversationId from a MailSearch result.")] string conversationId,
        [Description("Max messages to return from this thread (default 20).")] int maxMessages = 20)
    {
        var ctx   = GetSession(sessionId);
        var graph = ctx.GraphClient!;

        var messages = await graph.Me.Messages.GetAsync(cfg =>
        {
            cfg.QueryParameters.Filter  = $"conversationId eq '{conversationId}'";
            cfg.QueryParameters.Top     = maxMessages;
            cfg.QueryParameters.Select  = ["id", "subject", "from", "toRecipients", "receivedDateTime", "body", "hasAttachments"];
            cfg.QueryParameters.Orderby = ["receivedDateTime asc"];
        });

        var items = messages?.Value ?? [];
        return new
        {
            conversationId,
            messageCount = items.Count,
            messages = items.Select(m => new
            {
                id      = m.Id,
                from    = m.From?.EmailAddress?.Address,
                to      = m.ToRecipients?.Select(r => r.EmailAddress?.Address),
                date    = m.ReceivedDateTime?.ToString("f"),
                subject = m.Subject,
                body    = m.Body?.Content
            })
        };
    }

    [McpServerTool]
    [Description("Retrieves the full body of a single email by its ID.")]
    public async Task<object> MailGetById(
        [Description("Active sessionId.")] string sessionId,
        [Description("The email ID from a MailSearch result.")] string messageId)
    {
        var ctx = GetSession(sessionId);
        var m   = await ctx.GraphClient!.Me.Messages[messageId].GetAsync(cfg =>
            cfg.QueryParameters.Select = ["id", "subject", "from", "toRecipients", "ccRecipients",
                                          "receivedDateTime", "body", "hasAttachments", "attachments"]);

        return new
        {
            id      = m?.Id,
            from    = m?.From?.EmailAddress?.Address,
            to      = m?.ToRecipients?.Select(r => r.EmailAddress?.Address),
            cc      = m?.CcRecipients?.Select(r => r.EmailAddress?.Address),
            subject = m?.Subject,
            date    = m?.ReceivedDateTime?.ToString("f"),
            body    = m?.Body?.Content
        };
    }

    // ── Send & Draft ──────────────────────────────────────────────────────────

    [McpServerTool]
    [Description("Sends an email on behalf of the authenticated user.")]
    public async Task<object> MailSend(
        [Description("Active sessionId.")] string sessionId,
        [Description("Recipient email address.")] string to,
        [Description("Email subject.")] string subject,
        [Description("Email body. Supports plain text or basic HTML.")] string body,
        [Description("Optional CC recipients, comma-separated.")] string? cc = null,
        [Description("If true, body is treated as HTML. Default: false (plain text).")] bool isHtml = false)
    {
        var ctx = GetSession(sessionId);

        var ccRecipients = cc?.Split(',')
            .Select(e => new Recipient { EmailAddress = new() { Address = e.Trim() } })
            .ToList();

        await ctx.GraphClient!.Me.SendMail.PostAsync(new()
        {
            Message = new()
            {
                Subject = subject,
                Body    = new() { ContentType = isHtml ? BodyType.Html : BodyType.Text, Content = body },
                ToRecipients = [new() { EmailAddress = new() { Address = to } }],
                CcRecipients = ccRecipients
            },
            SaveToSentItems = true
        });

        return new { status = "sent", to, subject };
    }

    [McpServerTool]
    [Description(
        "Drafts a reply to an existing email. Does not send — saves to Drafts folder. " +
        "Use MailSend with the draft ID to send, or review in Outlook first.")]
    public async Task<object> MailDraftReply(
        [Description("Active sessionId.")] string sessionId,
        [Description("The message ID to reply to.")] string messageId,
        [Description("The reply body text.")] string replyBody,
        [Description("If true, reply-all. Default: false (reply to sender only).")] bool replyAll = false)
    {
        var ctx   = GetSession(sessionId);
        var graph = ctx.GraphClient!;

        Message draft;
        if (replyAll)
        {
            var result = await graph.Me.Messages[messageId].CreateReplyAll.PostAsync(new()
            {
                Message = new() { Body = new() { Content = replyBody, ContentType = BodyType.Text } }
            });
            draft = result!;
        }
        else
        {
            var result = await graph.Me.Messages[messageId].CreateReply.PostAsync(new()
            {
                Message = new() { Body = new() { Content = replyBody, ContentType = BodyType.Text } }
            });
            draft = result!;
        }

        return new
        {
            status    = "draft_saved",
            draftId   = draft.Id,
            subject   = draft.Subject,
            message   = "Draft saved. Review in Outlook or call MailSendDraft to send."
        };
    }

    [McpServerTool]
    [Description("Sends a previously saved draft email by its draft message ID.")]
    public async Task<object> MailSendDraft(
        [Description("Active sessionId.")] string sessionId,
        [Description("The draft message ID returned by MailDraftReply.")] string draftId)
    {
        var ctx = GetSession(sessionId);
        await ctx.GraphClient!.Me.Messages[draftId].Send.PostAsync();
        return new { status = "sent", draftId };
    }

    // ── Helper ────────────────────────────────────────────────────────────────

    private SessionContext GetSession(string sessionId)
    {
        var ctx = sessionStore.Get(sessionId)
            ?? throw new InvalidOperationException($"Session '{sessionId}' not found or expired. Call graph_initiate_login.");
        if (!ctx.IsAuthenticated)
            throw new InvalidOperationException("Session is not authenticated yet. Poll graph_check_login_status.");
        return ctx;
    }
}
