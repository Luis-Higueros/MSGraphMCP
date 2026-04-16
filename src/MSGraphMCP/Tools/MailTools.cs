using System.ComponentModel;
using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
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
        [Description("Optional folder to search in (e.g. 'inbox', 'sentitems', 'drafts', 'archive', or a folder ID). Leave empty to search all folders.")] string? folder = null,
        [Description("Filter by sender email address.")] string? fromAddress = null,
        [Description("Filter by email address in To field.")] string? toAddress = null,
        [Description("Start date (ISO 8601), e.g. '2025-03-01'. Can look back up to 12 months.")] string? since = null,
        [Description("End date (ISO 8601), e.g. '2025-03-31'. Defaults to today.")] string? until = null,
        [Description("Only return emails with attachments.")] bool hasAttachments = false,
        [Description("Max results to return (default 25, max 100).")]
        int maxResults = 25,
        [Description("Return full body text instead of preview (slower but more complete).")]
        bool includeBody = false)
    {
        try
        {
            var ctx = GetSession(sessionId);
            var graph = ctx.GraphClient!;

            var sinceUtc = ParseDateBoundaryUtc(since, endOfDay: false);
            var untilUtc = ParseDateBoundaryUtc(until, endOfDay: true);
            var normalizedFromAddress = string.IsNullOrWhiteSpace(fromAddress) ? null : fromAddress.Trim();
            var normalizedToAddress = string.IsNullOrWhiteSpace(toAddress) ? null : toAddress.Trim();
            var normalizedKeywords = string.IsNullOrWhiteSpace(keywords) ? null : keywords.Trim();
            var resolvedFolder = ResolveMailFolder(folder);

            // Build OData $filter only for non-$search queries.
            var filters = new List<string>();
            if (sinceUtc is not null)
                filters.Add($"receivedDateTime ge {sinceUtc.Value:O}");
            if (untilUtc is not null)
                filters.Add($"receivedDateTime le {untilUtc.Value:O}");
            if (normalizedFromAddress is not null)
                filters.Add($"from/emailAddress/address eq '{EscapeODataLiteral(normalizedFromAddress)}'");
            if (normalizedToAddress is not null)
                filters.Add($"toRecipients/any(t: t/emailAddress/address eq '{EscapeODataLiteral(normalizedToAddress)}')");
            if (hasAttachments)
                filters.Add("hasAttachments eq true");

            var baseSelectFields = new[]
            {
                "id", "subject", "from", "toRecipients", "receivedDateTime",
                "bodyPreview", "hasAttachments", "conversationId", "isRead"
            };

            var requestedTop = Math.Clamp(maxResults, 1, 100);
            var isKeywordSearch = normalizedKeywords is not null;
            var selectFields = includeBody && !isKeywordSearch
                ? [.. baseSelectFields, "body"]
                : baseSelectFields;

            MessageCollectionResponse? messages;
            if (resolvedFolder is null)
            {
                messages = await graph.Me.Messages.GetAsync(cfg =>
                {
                    cfg.QueryParameters.Select = selectFields;

                    if (isKeywordSearch)
                    {
                        // $search needs eventual consistency and is not reliable when mixed with
                        // custom $filter/$orderby across tenants.
                        cfg.Headers.Add("ConsistencyLevel", "eventual");
                        cfg.QueryParameters.Search = $"\"{EscapeSearchText(normalizedKeywords!)}\"";
                        cfg.QueryParameters.Top = 100;
                        return;
                    }

                    if (filters.Count > 0)
                        cfg.QueryParameters.Filter = string.Join(" and ", filters);

                    cfg.QueryParameters.Top = requestedTop;
                    cfg.QueryParameters.Orderby = ["receivedDateTime desc"];
                });
            }
            else
            {
                messages = await graph.Me.MailFolders[resolvedFolder].Messages.GetAsync(cfg =>
                {
                    cfg.QueryParameters.Select = selectFields;

                    if (isKeywordSearch)
                    {
                        cfg.Headers.Add("ConsistencyLevel", "eventual");
                        cfg.QueryParameters.Search = $"\"{EscapeSearchText(normalizedKeywords!)}\"";
                        cfg.QueryParameters.Top = 100;
                        return;
                    }

                    if (filters.Count > 0)
                        cfg.QueryParameters.Filter = string.Join(" and ", filters);

                    cfg.QueryParameters.Top = requestedTop;
                    cfg.QueryParameters.Orderby = ["receivedDateTime desc"];
                });
            }

            IEnumerable<Message> filtered = messages?.Value ?? [];
            if (isKeywordSearch)
            {
                filtered = filtered.Where(m =>
                    MatchesDateRange(m.ReceivedDateTime, sinceUtc, untilUtc)
                    && MatchesFromAddress(m, normalizedFromAddress)
                    && MatchesToAddress(m, normalizedToAddress)
                    && (!hasAttachments || m.HasAttachments == true));

                filtered = filtered.OrderByDescending(m => m.ReceivedDateTime)
                    .Take(requestedTop);
            }

            var items = filtered.ToList();
            if (includeBody && isKeywordSearch && items.Count > 0)
            {
                await PopulateBodiesAsync(graph, items);
            }

            if (items.Count == 0)
                return new { count = 0, message = "No emails matched your search.", emails = Array.Empty<object>() };

            return new
            {
                count = items.Count,
                emails = items.Select((m, i) => new
                {
                    index = i + 1,
                    id = m.Id,
                    conversationId = m.ConversationId,
                    from = m.From?.EmailAddress?.Address,
                    subject = m.Subject,
                    date = m.ReceivedDateTime?.ToString("f"),
                    isRead = m.IsRead,
                    hasAttachments = m.HasAttachments,
                    preview = m.BodyPreview,
                    body = includeBody ? m.Body?.Content : null
                })
            };
        }
        catch (Exception ex)
        {
            logger.LogError(ex,
                "MailSearch failed. sessionId={SessionId}, keywordsPresent={KeywordsPresent}, since={Since}, until={Until}, fromAddress={FromAddress}, toAddress={ToAddress}, maxResults={MaxResults}, includeBody={IncludeBody}",
                sessionId,
                !string.IsNullOrWhiteSpace(keywords),
                since,
                until,
                fromAddress,
                toAddress,
                maxResults,
                includeBody);

            return new
            {
                status = "search_failed",
                message = "MailSearch failed. Try removing keywords or narrowing date filters, then retry.",
                error = ex.Message
            };
        }
    }

    [McpServerTool]
    [Description(
        "Search emails and return a server-generated summary grouped by thread and key points. " +
        "Ideal for 'what happened with project X last month?' queries. " +
        "Returns summarized JSON directly, without requiring per-email follow-up calls.")]
    public async Task<object> MailSummarize(
        [Description("Active sessionId.")] string sessionId,
        [Description("The context or question to focus the summary on, e.g. 'project Alpha budget discussions'.")] string context,
        [Description("Keywords to search for. Defaults to words extracted from context.")] string? keywords = null,
        [Description("Optional folder to search in (e.g. 'inbox', 'sentitems', 'drafts', 'archive', or a folder ID). Leave empty to search all folders.")] string? folder = null,
        [Description("Sender filter.")] string? fromAddress = null,
        [Description("Start date (ISO 8601).")] string? since = null,
        [Description("End date (ISO 8601).")] string? until = null,
        [Description("Max emails to include (default 30).")]
        int maxEmails = 30)
    {
        var searchKeywords = keywords ?? context;
        var raw = await MailSearch(sessionId, searchKeywords, folder, fromAddress, null, since, until,
                                   false, maxEmails, true);

        var search = ConvertToMailSearchResult(raw);
        if (!string.IsNullOrWhiteSpace(search.Status) &&
            search.Status.Equals("search_failed", StringComparison.OrdinalIgnoreCase))
        {
            return new
            {
                status = "summarize_failed",
                context,
                message = search.Message ?? "MailSummarize could not complete because MailSearch failed.",
                error = search.Error
            };
        }

        if (search.Count == 0 || search.Emails.Count == 0)
        {
            return new
            {
                status = "ok",
                summarizationRequest = false,
                context,
                query = new
                {
                    keywords = searchKeywords,
                    folder,
                    fromAddress,
                    since,
                    until,
                    maxEmails
                },
                summary = "No emails matched the requested criteria.",
                counts = new
                {
                    emails = 0,
                    threads = 0,
                    unread = 0,
                    withAttachments = 0
                },
                topSenders = Array.Empty<object>(),
                threads = Array.Empty<object>(),
                emails = Array.Empty<object>()
            };
        }

        var orderedEmails = search.Emails
            .Where(static e => !string.IsNullOrWhiteSpace(e.Id))
            .OrderBy(static e => e.Index)
            .ToList();

        var emailSummaries = orderedEmails.Select(e =>
        {
            var content = CleanTextForSummary(string.IsNullOrWhiteSpace(e.Body) ? e.Preview : e.Body);
            var shortSummary = BuildEmailSummary(e.Subject, content);
            var actionItems = ExtractActionItems(content).Take(3).ToArray();

            return new
            {
                index = e.Index,
                id = e.Id,
                conversationId = e.ConversationId,
                from = e.From,
                subject = e.Subject,
                date = e.Date,
                isRead = e.IsRead,
                hasAttachments = e.HasAttachments,
                shortSummary,
                actionItems,
                preview = BuildSnippet(content, 280)
            };
        }).ToList();

        var threadSummaries = orderedEmails
            .GroupBy(e => string.IsNullOrWhiteSpace(e.ConversationId) ? "(no-conversation-id)" : e.ConversationId!)
            .Select(g => new
            {
                conversationId = g.Key,
                messageCount = g.Count(),
                latestSubject = g.OrderBy(e => e.Index).First().Subject,
                participants = g.Select(e => e.From)
                    .Where(static f => !string.IsNullOrWhiteSpace(f))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Take(8)
                    .ToArray()
            })
            .OrderByDescending(t => t.messageCount)
            .ToList();

        var topSenders = orderedEmails
            .GroupBy(e => string.IsNullOrWhiteSpace(e.From) ? "(unknown)" : e.From!)
            .Select(g => new { sender = g.Key, count = g.Count() })
            .OrderByDescending(s => s.count)
            .Take(5)
            .ToList();

        var unreadCount = orderedEmails.Count(e => e.IsRead == false);
        var attachmentCount = orderedEmails.Count(e => e.HasAttachments == true);

        var highLevelSummary = BuildHighLevelSummary(context, orderedEmails.Count, threadSummaries.Count, unreadCount, attachmentCount, topSenders);

        return new
        {
            status = "ok",
            summarizationRequest = false,
            context,
            query = new
            {
                keywords = searchKeywords,
                folder,
                fromAddress,
                since,
                until,
                maxEmails
            },
            summary = highLevelSummary,
            counts = new
            {
                emails = orderedEmails.Count,
                threads = threadSummaries.Count,
                unread = unreadCount,
                withAttachments = attachmentCount
            },
            topSenders,
            threads = threadSummaries,
            emails = emailSummaries
        };
    }

    // ── Thread & Detail ───────────────────────────────────────────────────────

    [McpServerTool]
    [Description("Retrieves all messages in an email conversation thread, in chronological order.")]
    public async Task<object> MailGetThread(
        [Description("Active sessionId.")] string sessionId,
        [Description("The conversationId from a MailSearch result.")] string conversationId,
        [Description("Max messages to return from this thread (default 20).")]
        int maxMessages = 20)
    {
        var ctx = GetSession(sessionId);
        var graph = ctx.GraphClient!;

        var messages = await graph.Me.Messages.GetAsync(cfg =>
        {
            cfg.QueryParameters.Filter = $"conversationId eq '{conversationId}'";
            cfg.QueryParameters.Top = maxMessages;
            cfg.QueryParameters.Select = ["id", "subject", "from", "toRecipients", "receivedDateTime", "body", "hasAttachments"];
            cfg.QueryParameters.Orderby = ["receivedDateTime asc"];
        });

        var items = messages?.Value ?? [];
        return new
        {
            conversationId,
            messageCount = items.Count,
            messages = items.Select(m => new
            {
                id = m.Id,
                from = m.From?.EmailAddress?.Address,
                to = m.ToRecipients?.Select(r => r.EmailAddress?.Address),
                date = m.ReceivedDateTime?.ToString("f"),
                subject = m.Subject,
                body = m.Body?.Content
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
        var m = await ctx.GraphClient!.Me.Messages[messageId].GetAsync(cfg =>
            cfg.QueryParameters.Select = ["id", "subject", "from", "toRecipients", "ccRecipients",
                                          "receivedDateTime", "body", "hasAttachments", "attachments"]);

        return new
        {
            id = m?.Id,
            from = m?.From?.EmailAddress?.Address,
            to = m?.ToRecipients?.Select(r => r.EmailAddress?.Address),
            cc = m?.CcRecipients?.Select(r => r.EmailAddress?.Address),
            subject = m?.Subject,
            date = m?.ReceivedDateTime?.ToString("f"),
            body = m?.Body?.Content
        };
    }

    // ── Send & Draft ──────────────────────────────────────────────────────────

    [McpServerTool]
    [Description("Sends an email on behalf of the authenticated user.")]
    public async Task<object> MailSend(
        [Description("Active sessionId.")] string sessionId,
        [Description("Recipient email address.")] string recipient,
        [Description("Email subject.")] string subject,
        [Description("Email body. Supports plain text or basic HTML.")] string body,
        [Description("Optional CC recipients, comma-separated.")] string? cc = null,
        [Description("If true, body is treated as HTML. Default: false (plain text).")]
        bool isHtml = false)
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
                Body = new() { ContentType = isHtml ? BodyType.Html : BodyType.Text, Content = body },
                ToRecipients = [new() { EmailAddress = new() { Address = recipient } }],
                CcRecipients = ccRecipients
            },
            SaveToSentItems = true
        });

        return new { status = "sent", to = recipient, subject };
    }

    [McpServerTool]
    [Description(
        "Drafts a reply to an existing email. Does not send — saves to Drafts folder. " +
        "Use MailSend with the draft ID to send, or review in Outlook first.")]
    public async Task<object> MailDraftReply(
        [Description("Active sessionId.")] string sessionId,
        [Description("The message ID to reply to.")] string messageId,
        [Description("The reply body text.")] string replyBody,
        [Description("If true, reply-all. Default: false (reply to sender only).")]
        bool replyAll = false)
    {
        var ctx = GetSession(sessionId);
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
            status = "draft_saved",
            draftId = draft.Id,
            subject = draft.Subject,
            message = "Draft saved. Review in Outlook or call MailSendDraft to send."
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

    private static DateTimeOffset? ParseDateBoundaryUtc(string? value, bool endOfDay)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;

        var trimmed = value.Trim();
        if (DateOnly.TryParseExact(trimmed, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateOnly))
        {
            var time = endOfDay ? new TimeOnly(23, 59, 59) : TimeOnly.MinValue;
            return new DateTimeOffset(dateOnly.ToDateTime(time), TimeSpan.Zero);
        }

        if (DateTimeOffset.TryParse(trimmed, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var dateTimeOffset))
            return dateTimeOffset.ToUniversalTime();

        throw new ArgumentException($"Invalid date value '{value}'. Expected yyyy-MM-dd.");
    }

    private static string EscapeODataLiteral(string value) => value.Replace("'", "''", StringComparison.Ordinal);

    private static string EscapeSearchText(string value) => value.Replace("\"", "\\\"", StringComparison.Ordinal);

    private static bool MatchesDateRange(DateTimeOffset? messageDate, DateTimeOffset? sinceUtc, DateTimeOffset? untilUtc)
    {
        if (messageDate is null)
            return false;

        var date = messageDate.Value;
        if (sinceUtc is not null && date < sinceUtc.Value)
            return false;
        if (untilUtc is not null && date > untilUtc.Value)
            return false;

        return true;
    }

    private static bool MatchesFromAddress(Message message, string? expectedAddress)
    {
        if (expectedAddress is null)
            return true;

        var from = message.From?.EmailAddress?.Address;
        return from is not null && from.Equals(expectedAddress, StringComparison.OrdinalIgnoreCase);
    }

    private static bool MatchesToAddress(Message message, string? expectedAddress)
    {
        if (expectedAddress is null)
            return true;

        return message.ToRecipients?.Any(r =>
            r.EmailAddress?.Address?.Equals(expectedAddress, StringComparison.OrdinalIgnoreCase) == true) == true;
    }

    private static string? ResolveMailFolder(string? folder)
    {
        if (string.IsNullOrWhiteSpace(folder))
            return null;

        var trimmed = folder.Trim();
        return trimmed.ToLowerInvariant() switch
        {
            "inbox" => "inbox",
            "sent" => "sentitems",
            "sentitems" => "sentitems",
            "drafts" => "drafts",
            "archive" => "archive",
            "deleted" => "deleteditems",
            "deleteditems" => "deleteditems",
            "junk" => "junkemail",
            "junkemail" => "junkemail",
            _ => trimmed
        };
    }

    private static MailSearchResult ConvertToMailSearchResult(object raw)
    {
        var json = JsonSerializer.Serialize(raw);
        var parsed = JsonSerializer.Deserialize<MailSearchResult>(json, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        });

        return parsed ?? new MailSearchResult();
    }

    private static async Task PopulateBodiesAsync(GraphServiceClient graph, IList<Message> messages)
    {
        using var throttler = new SemaphoreSlim(6);
        var tasks = messages
            .Where(static message => !string.IsNullOrWhiteSpace(message.Id))
            .Select(async message =>
            {
                await throttler.WaitAsync();
                try
                {
                    var detail = await graph.Me.Messages[message.Id!].GetAsync(cfg =>
                        cfg.QueryParameters.Select = ["id", "body"]);

                    if (detail?.Body is not null)
                        message.Body = detail.Body;
                }
                finally
                {
                    throttler.Release();
                }
            });

        await Task.WhenAll(tasks);
    }

    private static string CleanTextForSummary(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        var text = Regex.Replace(value, "<[^>]+>", " ", RegexOptions.Compiled);
        text = WebUtility.HtmlDecode(text);
        text = text.Replace("\r", " ", StringComparison.Ordinal)
                   .Replace("\n", " ", StringComparison.Ordinal)
                   .Replace("\t", " ", StringComparison.Ordinal);
        text = Regex.Replace(text, "\\s+", " ", RegexOptions.Compiled).Trim();
        return text;
    }

    private static string BuildSnippet(string text, int maxChars)
    {
        if (string.IsNullOrWhiteSpace(text))
            return string.Empty;

        return text.Length <= maxChars
            ? text
            : text[..maxChars].TrimEnd() + "...";
    }

    private static string BuildEmailSummary(string? subject, string content)
    {
        var subjectPart = string.IsNullOrWhiteSpace(subject) ? "(no subject)" : subject.Trim();
        var snippet = BuildSnippet(content, 180);
        return string.IsNullOrWhiteSpace(snippet)
            ? subjectPart
            : $"{subjectPart}: {snippet}";
    }

    private static IEnumerable<string> ExtractActionItems(string content)
    {
        if (string.IsNullOrWhiteSpace(content))
            return [];

        var sentences = Regex.Split(content, "(?<=[.!?])\\s+", RegexOptions.Compiled)
            .Select(s => s.Trim())
            .Where(s => s.Length >= 15)
            .Take(40);

        var markers = new[]
        {
            "action", "todo", "follow up", "follow-up", "next step",
            "please", "need to", "required", "deadline", "by "
        };

        return sentences
            .Where(s => markers.Any(m => s.Contains(m, StringComparison.OrdinalIgnoreCase)))
            .Select(s => BuildSnippet(s, 160))
            .Distinct(StringComparer.OrdinalIgnoreCase);
    }

    private static string BuildHighLevelSummary(
        string context,
        int emailCount,
        int threadCount,
        int unreadCount,
        int attachmentCount,
        IEnumerable<object> topSenders)
    {
        var senderText = string.Join(", ", topSenders.Select(s => s?.ToString()).Where(static s => !string.IsNullOrWhiteSpace(s)).Take(3));
        var baseSummary =
            $"Found {emailCount} email(s) across {threadCount} thread(s) for context '{context}'. " +
            $"Unread: {unreadCount}. With attachments: {attachmentCount}.";

        return string.IsNullOrWhiteSpace(senderText)
            ? baseSummary
            : baseSummary + " Top senders are included in the response.";
    }

    private sealed class MailSearchResult
    {
        public string? Status { get; set; }
        public string? Message { get; set; }
        public string? Error { get; set; }
        public int Count { get; set; }
        public List<MailSearchEmail> Emails { get; set; } = [];
    }

    private sealed class MailSearchEmail
    {
        public int Index { get; set; }
        public string? Id { get; set; }
        public string? ConversationId { get; set; }
        public string? From { get; set; }
        public string? Subject { get; set; }
        public string? Date { get; set; }
        public bool? IsRead { get; set; }
        public bool? HasAttachments { get; set; }
        public string? Preview { get; set; }
        public string? Body { get; set; }
    }
}
