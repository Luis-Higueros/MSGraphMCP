using System.ComponentModel;
using Microsoft.Graph;
using Microsoft.Graph.Me.Calendar.GetSchedule;
using Microsoft.Graph.Me.FindMeetingTimes;
using Microsoft.Graph.Models;
using ModelContextProtocol.Server;
using MSGraphMCP.Session;

namespace MSGraphMCP.Tools;

[McpServerToolType]
public class CalendarTools(SessionStore sessionStore, ILogger<CalendarTools> logger)
{
    // ── Free Slot Finder ──────────────────────────────────────────────────────

    [McpServerTool]
    [Description(
        "Find all free time slots in the user's calendar within a date range. " +
        "Automatically adjusts for the user's configured timezone. " +
        "Great for 'when am I free this week for a 30-minute call?' queries.")]
    public async Task<object> CalendarFindFreeSlots(
        [Description("Active sessionId.")] string sessionId,
        [Description("Start of search window (ISO 8601), e.g. '2025-04-01'.")] string from,
        [Description("End of search window (ISO 8601), e.g. '2025-04-07'.")] string to,
        [Description("Minimum free slot duration in minutes (default 30).")] int minDurationMinutes = 30,
        [Description("Start of working day in HH:mm (default '09:00').")] string workDayStart = "09:00",
        [Description("End of working day in HH:mm (default '17:30').")] string workDayEnd = "17:30",
        [Description("Include weekends in search. Default: false.")] bool includeWeekends = false)
    {
        var ctx   = GetSession(sessionId);
        var graph = ctx.GraphClient!;

        // Get user's timezone from their mailbox settings
        var mailboxSettings = await graph.Me.MailboxSettings.GetAsync();
        var tzId = mailboxSettings?.TimeZone ?? "UTC";
        TimeZoneInfo tz;
        try { tz = TimeZoneInfo.FindSystemTimeZoneById(tzId); }
        catch { tz = TimeZoneInfo.Utc; }

        var startLocal = DateTime.Parse(from);
        var endLocal   = DateTime.Parse(to).AddDays(1).AddSeconds(-1);
        var startUtc   = TimeZoneInfo.ConvertTimeToUtc(startLocal, tz);
        var endUtc     = TimeZoneInfo.ConvertTimeToUtc(endLocal,   tz);

        // Fetch all non-free events in the window (handle paging)
        var events = new List<Event>();
        var page   = await graph.Me.CalendarView.GetAsync(cfg =>
        {
            cfg.QueryParameters.StartDateTime = startUtc.ToString("o");
            cfg.QueryParameters.EndDateTime   = endUtc.ToString("o");
            cfg.QueryParameters.Select        = ["subject", "start", "end", "showAs", "isCancelled"];
            cfg.QueryParameters.Top           = 100;
            cfg.QueryParameters.Orderby       = ["start/dateTime"];
        });

        while (page?.Value is not null)
        {
            events.AddRange(page.Value.Where(e =>
                e.IsCancelled != true &&
                e.ShowAs != FreeBusyStatus.Free &&
                e.ShowAs != FreeBusyStatus.Tentative));

            if (page.OdataNextLink is null) break;
            page = await graph.Me.CalendarView.WithUrl(page.OdataNextLink).GetAsync();
        }

        // Compute working-hours gaps per day
        var freeSlots = new List<object>();
        var wdStart   = TimeSpan.Parse(workDayStart);
        var wdEnd     = TimeSpan.Parse(workDayEnd);
        var minSpan   = TimeSpan.FromMinutes(minDurationMinutes);

        for (var day = startLocal.Date; day <= endLocal.Date; day = day.AddDays(1))
        {
            if (!includeWeekends && (day.DayOfWeek == DayOfWeek.Saturday || day.DayOfWeek == DayOfWeek.Sunday))
                continue;

            var dayStart = TimeZoneInfo.ConvertTimeToUtc(day + wdStart, tz);
            var dayEnd   = TimeZoneInfo.ConvertTimeToUtc(day + wdEnd,   tz);

            var dayEvents = events
                .Where(e =>
                {
                    var es = ParseUtc(e.Start!.DateTime!);
                    var ee = ParseUtc(e.End!.DateTime!);
                    return ee > dayStart && es < dayEnd;
                })
                .OrderBy(e => ParseUtc(e.Start!.DateTime!))
                .ToList();

            var cursor = dayStart;
            foreach (var ev in dayEvents)
            {
                var evStart = ParseUtc(ev.Start!.DateTime!);
                var evEnd   = ParseUtc(ev.End!.DateTime!);
                evStart = evStart < dayStart ? dayStart : evStart;
                evEnd   = evEnd   > dayEnd   ? dayEnd   : evEnd;

                if (evStart - cursor >= minSpan)
                    freeSlots.Add(Slot(cursor, evStart, tz));

                cursor = evEnd > cursor ? evEnd : cursor;
            }

            if (dayEnd - cursor >= minSpan)
                freeSlots.Add(Slot(cursor, dayEnd, tz));
        }

        return new
        {
            timezone   = tzId,
            slotCount  = freeSlots.Count,
            minDuration = $"{minDurationMinutes} minutes",
            slots      = freeSlots
        };
    }

    // ── Meeting Time Suggestions ──────────────────────────────────────────────

    [McpServerTool]
    [Description(
        "Uses Microsoft's built-in findMeetingTimes API to suggest optimal meeting times " +
        "across multiple attendees' calendars. Returns ranked suggestions with availability scores.")]
    public async Task<object> CalendarSuggestMeetingTimes(
        [Description("Active sessionId.")] string sessionId,
        [Description("Comma-separated attendee email addresses.")] string attendeeEmails,
        [Description("Meeting duration in minutes (default 60).")] int durationMinutes = 60,
        [Description("Start of search window (ISO 8601).")] string? from = null,
        [Description("End of search window (ISO 8601).")] string? to = null,
        [Description("Max suggestions to return (default 5).")] int maxSuggestions = 5)
    {
        var ctx   = GetSession(sessionId);
        var graph = ctx.GraphClient!;

        var mailboxSettings = await graph.Me.MailboxSettings.GetAsync();
        var tzId = mailboxSettings?.TimeZone ?? "UTC";

        var now       = DateTime.UtcNow;
        var fromDt    = from is not null ? DateTime.Parse(from).ToUniversalTime() : now;
        var toDt      = to   is not null ? DateTime.Parse(to).ToUniversalTime()   : now.AddDays(7);

        var attendees = attendeeEmails.Split(',')
            .Select(e => new AttendeeBase
            {
                Type         = AttendeeType.Required,
                EmailAddress = new() { Address = e.Trim() }
            })
            .ToList();

        var body = new FindMeetingTimesPostRequestBody
        {
            Attendees       = attendees,
            MeetingDuration = TimeSpan.FromMinutes(durationMinutes),
            MaxCandidates   = maxSuggestions,
            TimeConstraint  = new()
            {
                ActivityDomain = ActivityDomain.Work,
                TimeSlots      =
                [
                    new() {
                        Start = new() { DateTime = fromDt.ToString("o"), TimeZone = "UTC" },
                        End   = new() { DateTime = toDt.ToString("o"),   TimeZone = "UTC" }
                    }
                ]
            }
        };

        var suggestions = await graph.Me.FindMeetingTimes.PostAsync(body);

        if (suggestions?.MeetingTimeSuggestions is null || suggestions.MeetingTimeSuggestions.Count == 0)
            return new { message = "No meeting times found. Try a wider date range or fewer attendees." };

        TimeZoneInfo tz;
        try { tz = TimeZoneInfo.FindSystemTimeZoneById(tzId); }
        catch { tz = TimeZoneInfo.Utc; }

        return new
        {
            timezone    = tzId,
            suggestions = suggestions.MeetingTimeSuggestions.Select(s => new
            {
                confidence = $"{s.Confidence:P0}",
                start      = FormatLocal(s.MeetingTimeSlot?.Start?.DateTime, tz),
                end        = FormatLocal(s.MeetingTimeSlot?.End?.DateTime,   tz),
                attendeeAvailability = s.AttendeeAvailability?.Select(a => new
                {
                    attendee     = a.Attendee?.EmailAddress?.Address,
                    availability = a.Availability?.ToString()
                })
            })
        };
    }

    // ── Agenda ────────────────────────────────────────────────────────────────

    [McpServerTool]
    [Description("Returns the user's calendar agenda for a given date range, grouped by day.")]
    public async Task<object> CalendarGetAgenda(
        [Description("Active sessionId.")] string sessionId,
        [Description("Start date (ISO 8601).")] string from,
        [Description("End date (ISO 8601).")] string to)
    {
        var ctx   = GetSession(sessionId);
        var graph = ctx.GraphClient!;

        var mailboxSettings = await graph.Me.MailboxSettings.GetAsync();
        var tzId = mailboxSettings?.TimeZone ?? "UTC";
        TimeZoneInfo tz;
        try { tz = TimeZoneInfo.FindSystemTimeZoneById(tzId); }
        catch { tz = TimeZoneInfo.Utc; }

        var startUtc = DateTime.Parse(from).ToUniversalTime();
        var endUtc   = DateTime.Parse(to).AddDays(1).ToUniversalTime();

        var events = new List<Event>();
        var page   = await graph.Me.CalendarView.GetAsync(cfg =>
        {
            cfg.QueryParameters.StartDateTime = startUtc.ToString("o");
            cfg.QueryParameters.EndDateTime   = endUtc.ToString("o");
            cfg.QueryParameters.Select        = ["subject", "start", "end", "location",
                                                  "organizer", "isAllDay", "showAs", "bodyPreview"];
            cfg.QueryParameters.Top           = 200;
            cfg.QueryParameters.Orderby       = ["start/dateTime"];
        });

        while (page?.Value is not null)
        {
            events.AddRange(page.Value);
            if (page.OdataNextLink is null) break;
            page = await graph.Me.CalendarView.WithUrl(page.OdataNextLink).GetAsync();
        }

        var grouped = events
            .GroupBy(e =>
            {
                var localStart = TimeZoneInfo.ConvertTimeFromUtc(ParseUtc(e.Start!.DateTime!), tz);
                return localStart.Date;
            })
            .OrderBy(g => g.Key)
            .Select(g => new
            {
                date   = g.Key.ToString("dddd, MMMM d, yyyy"),
                events = g.Select(e => new
                {
                    subject   = e.Subject,
                    start     = FormatLocal(e.Start?.DateTime, tz),
                    end       = FormatLocal(e.End?.DateTime,   tz),
                    isAllDay  = e.IsAllDay,
                    location  = e.Location?.DisplayName,
                    organizer = e.Organizer?.EmailAddress?.Address,
                    status    = e.ShowAs?.ToString(),
                    preview   = e.BodyPreview
                })
            });

        return new { timezone = tzId, agenda = grouped };
    }

    // ── Create Event ──────────────────────────────────────────────────────────

    [McpServerTool]
    [Description("Creates a new calendar event. Times are interpreted in the user's configured timezone.")]
    public async Task<object> CalendarCreateEvent(
        [Description("Active sessionId.")] string sessionId,
        [Description("Event title.")] string subject,
        [Description("Start time (ISO 8601 local time), e.g. '2025-04-10T14:00:00'.")] string startTime,
        [Description("End time (ISO 8601 local time), e.g. '2025-04-10T15:00:00'.")] string endTime,
        [Description("Optional body / agenda text.")] string? body = null,
        [Description("Optional location.")] string? location = null,
        [Description("Optional comma-separated attendee email addresses.")] string? attendees = null,
        [Description("If true, sends meeting invites to attendees.")] bool sendInvite = true)
    {
        var ctx   = GetSession(sessionId);
        var graph = ctx.GraphClient!;

        var mailboxSettings = await graph.Me.MailboxSettings.GetAsync();
        var tzId = mailboxSettings?.TimeZone ?? "UTC";

        var newEvent = new Event
        {
            Subject = subject,
            Body    = body is not null ? new() { ContentType = BodyType.Text, Content = body } : null,
            Start   = new() { DateTime = startTime, TimeZone = tzId },
            End     = new() { DateTime = endTime,   TimeZone = tzId },
            Location = location is not null ? new() { DisplayName = location } : null,
            Attendees = attendees?
                .Split(',')
                .Select(e => new Attendee
                {
                    Type         = AttendeeType.Required,
                    EmailAddress = new() { Address = e.Trim() }
                })
                .ToList<Attendee>()
        };

        var created = await graph.Me.Events.PostAsync(newEvent);

        return new
        {
            status     = "created",
            eventId    = created?.Id,
            subject    = created?.Subject,
            start      = created?.Start?.DateTime,
            end        = created?.End?.DateTime,
            timezone   = tzId,
            webLink    = created?.WebLink
        };
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static DateTime ParseUtc(string dt) =>
        DateTime.SpecifyKind(DateTime.Parse(dt), DateTimeKind.Utc);

    private static object Slot(DateTime start, DateTime end, TimeZoneInfo tz) => new
    {
        start    = TimeZoneInfo.ConvertTimeFromUtc(start, tz).ToString("ddd MMM d, h:mm tt"),
        end      = TimeZoneInfo.ConvertTimeFromUtc(end,   tz).ToString("h:mm tt"),
        duration = $"{(end - start).TotalMinutes:0} min"
    };

    private static string? FormatLocal(string? utcStr, TimeZoneInfo tz)
    {
        if (utcStr is null) return null;
        var dt = DateTime.SpecifyKind(DateTime.Parse(utcStr), DateTimeKind.Utc);
        return TimeZoneInfo.ConvertTimeFromUtc(dt, tz).ToString("ddd MMM d, h:mm tt");
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
