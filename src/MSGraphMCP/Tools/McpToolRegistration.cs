using System.Reflection;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace MSGraphMCP.Tools;

public static class McpToolRegistration
{
    public static IReadOnlyList<McpServerTool> CreateValidatedTools(
        IServiceCollection services,
        params Type[] toolTypes)
    {
        var tools = new List<McpServerTool>();

        foreach (var toolType in toolTypes)
        {
            services.AddSingleton(toolType);

            foreach (var method in GetToolMethods(toolType))
            {
                McpServerTool baseTool;
                if (method.IsStatic)
                {
                    baseTool = McpServerTool.Create(method, target: null, options: null);
                }
                else
                {
                    baseTool = McpServerTool.Create(
                        method,
                        createTargetFunc: request => (request.Services ?? throw new InvalidOperationException("Request services unavailable."))
                            .GetRequiredService(toolType),
                        options: null);
                }

                tools.Add(new SchemaValidatingMcpServerTool(baseTool));
            }
        }

        return tools;
    }

    private static IEnumerable<MethodInfo> GetToolMethods(Type toolType)
    {
        return toolType
            .GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static)
            .Where(static m => m.GetCustomAttribute<McpServerToolAttribute>() is not null);
    }
}
