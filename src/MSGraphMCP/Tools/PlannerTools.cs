using System.ComponentModel;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using ModelContextProtocol.Server;
using MSGraphMCP.Session;

namespace MSGraphMCP.Tools;

[McpServerToolType]
public class PlannerTools(SessionStore sessionStore, ILogger<PlannerTools> logger)
{
    [McpServerTool]
    [Description("Lists all Planner plans the user has access to (via their joined groups).")]
    public async Task<object> PlannerListPlans(
        [Description("Active sessionId.")] string sessionId)
    {
        var ctx    = GetSession(sessionId);
        var groups = await ctx.GraphClient!.Me.MemberOf.GetAsync(cfg =>
            cfg.QueryParameters.Select = ["id", "displayName"]);

        var plans = new List<object>();
        foreach (var group in groups?.Value?.OfType<Group>() ?? [])
        {
            try
            {
                var groupPlans = await ctx.GraphClient.Groups[group.Id].Planner.Plans
                    .GetAsync(cfg =>
                        cfg.QueryParameters.Select = ["id", "title", "createdDateTime"]);

                plans.AddRange(groupPlans?.Value?.Select(p => new
                {
                    planId    = p.Id,
                    title     = p.Title,
                    groupId   = group.Id,
                    groupName = group.DisplayName,
                    created   = p.CreatedDateTime?.ToString("f")
                }) ?? []);
            }
            catch { /* Group may not have Planner */ }
        }

        return new { count = plans.Count, plans };
    }

    [McpServerTool]
    [Description("Lists tasks in a Planner plan, optionally filtered by bucket or assignee.")]
    public async Task<object> PlannerListTasks(
        [Description("Active sessionId.")] string sessionId,
        [Description("Plan ID from PlannerListPlans.")] string planId,
        [Description("Filter to a specific bucket name (optional).")] string? bucketName = null,
        [Description("Filter to tasks assigned to a specific user display name (optional).")] string? assignedTo = null,
        [Description("Include completed tasks. Default: false.")] bool includeCompleted = false,
        [Description("Max tasks to return (default 50).")] int maxTasks = 50)
    {
        var ctx    = GetSession(sessionId);
        var result = await ctx.GraphClient!.Planner.Plans[planId].Tasks.GetAsync(cfg =>
            cfg.QueryParameters.Select = ["id", "title", "percentComplete", "dueDateTime",
                                          "assignments", "bucketId", "priority", "createdDateTime"]);

        var tasks = result?.Value ?? [];

        // Get bucket names if filtering
        Dictionary<string, string> bucketMap = [];
        if (!string.IsNullOrEmpty(bucketName) || true)
        {
            var buckets = await ctx.GraphClient.Planner.Plans[planId].Buckets.GetAsync();
            bucketMap = buckets?.Value?.ToDictionary(b => b.Id!, b => b.Name ?? "") ?? [];
        }

        var filtered = tasks
            .Where(t => includeCompleted || t.PercentComplete < 100)
            .Where(t => bucketName is null ||
                        (bucketMap.TryGetValue(t.BucketId ?? "", out var bn) &&
                         bn.Contains(bucketName, StringComparison.OrdinalIgnoreCase)))
            .Take(maxTasks)
            .Select(t => new
            {
                id             = t.Id,
                title          = t.Title,
                percentComplete = t.PercentComplete,
                priority       = PriorityLabel(t.Priority),
                due            = t.DueDateTime?.ToString("f"),
                bucket         = bucketMap.TryGetValue(t.BucketId ?? "", out var bn2) ? bn2 : null,
                created        = t.CreatedDateTime?.ToString("f"),
                assigneeCount  = t.Assignments?.AdditionalData?.Count ?? 0
            })
            .ToList();

        return new { planId, count = filtered.Count, tasks = filtered };
    }

    [McpServerTool]
    [Description("Creates a new task in a Planner plan.")]
    public async Task<object> PlannerCreateTask(
        [Description("Active sessionId.")] string sessionId,
        [Description("Plan ID from PlannerListPlans.")] string planId,
        [Description("Task title.")] string title,
        [Description("Bucket name to place the task in (optional — defaults to first bucket).")] string? bucketName = null,
        [Description("Due date (ISO 8601), e.g. '2025-04-15'.")] string? dueDate = null,
        [Description("Priority: 0=Urgent, 1=Important, 5=Medium, 9=Low. Default: 5.")] int priority = 5,
        [Description("Optional task notes/description.")] string? notes = null)
    {
        var ctx = GetSession(sessionId);

        // Resolve bucket
        string? bucketId = null;
        if (bucketName is not null)
        {
            var buckets = await ctx.GraphClient!.Planner.Plans[planId].Buckets.GetAsync();
            bucketId = buckets?.Value?
                .FirstOrDefault(b => b.Name?.Contains(bucketName, StringComparison.OrdinalIgnoreCase) == true)?.Id;
        }

        var task = new PlannerTask
        {
            PlanId   = planId,
            Title    = title,
            BucketId = bucketId,
            Priority = priority,
            DueDateTime = dueDate is not null
                ? DateTimeOffset.Parse(dueDate).ToUniversalTime()
                : null
        };

        var created = await ctx.GraphClient!.Planner.Tasks.PostAsync(task);

        // Add notes via task details if provided
        if (notes is not null && created?.Id is not null)
        {
            var details = await ctx.GraphClient.Planner.Tasks[created.Id].Details.GetAsync();
            await ctx.GraphClient.Planner.Tasks[created.Id].Details.PatchAsync(new PlannerTaskDetails
            {
                Description = notes
            });
        }

        return new
        {
            status   = "created",
            taskId   = created?.Id,
            title    = created?.Title,
            planId,
            bucketId = created?.BucketId,
            due      = created?.DueDateTime?.ToString("f")
        };
    }

    [McpServerTool]
    [Description("Updates a Planner task's completion status, priority, or due date.")]
    public async Task<object> PlannerUpdateTask(
        [Description("Active sessionId.")] string sessionId,
        [Description("Task ID from PlannerListTasks.")] string taskId,
        [Description("Completion percentage (0–100). Use 100 to mark complete.")] int? percentComplete = null,
        [Description("New priority: 0=Urgent, 1=Important, 5=Medium, 9=Low.")] int? priority = null,
        [Description("New due date (ISO 8601).")] string? dueDate = null,
        [Description("New task title.")] string? title = null)
    {
        var ctx     = GetSession(sessionId);
        var current = await ctx.GraphClient!.Planner.Tasks[taskId].GetAsync();

        if (current is null)
            return new { error = "Task not found." };

        // Graph requires the ETag for optimistic concurrency on PATCH
        var update = new PlannerTask
        {
            PercentComplete = percentComplete ?? current.PercentComplete,
            Priority        = priority ?? current.Priority,
            Title           = title ?? current.Title,
            DueDateTime     = dueDate is not null
                ? DateTimeOffset.Parse(dueDate).ToUniversalTime()
                : current.DueDateTime
        };

        await ctx.GraphClient.Planner.Tasks[taskId].PatchAsync(update, cfg =>
            cfg.Headers.Add("If-Match", current.AdditionalData.TryGetValue("@odata.etag", out var etag)
                ? etag?.ToString() ?? "*" : "*"));

        return new
        {
            status          = "updated",
            taskId,
            percentComplete = update.PercentComplete,
            priority        = PriorityLabel(update.Priority)
        };
    }

    private static string PriorityLabel(int? priority) => priority switch
    {
        0 => "Urgent",
        1 => "Important",
        5 => "Medium",
        9 => "Low",
        _ => priority?.ToString() ?? "Unknown"
    };

    private SessionContext GetSession(string sessionId)
    {
        var ctx = sessionStore.Get(sessionId)
            ?? throw new InvalidOperationException("Session not found. Call graph_initiate_login.");
        if (!ctx.IsAuthenticated)
            throw new InvalidOperationException("Not authenticated yet.");
        return ctx;
    }
}
