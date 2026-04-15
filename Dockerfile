# ── Build stage ───────────────────────────────────────────────────────────────
FROM mcr.microsoft.com/dotnet/sdk:9.0 AS build
WORKDIR /src

COPY ["src/MSGraphMCP/MSGraphMCP.csproj", "MSGraphMCP/"]
RUN dotnet restore "MSGraphMCP/MSGraphMCP.csproj"

COPY src/MSGraphMCP/ MSGraphMCP/
WORKDIR /src/MSGraphMCP
RUN dotnet publish "MSGraphMCP.csproj" \
    -c Release \
    -o /app/publish \
    --no-restore

# ── Runtime stage ─────────────────────────────────────────────────────────────
FROM mcr.microsoft.com/dotnet/aspnet:9.0 AS runtime
WORKDIR /app

# Non-root user for security
RUN adduser --disabled-password --gecos "" appuser && chown -R appuser /app
USER appuser

COPY --from=build --chown=appuser:appuser /app/publish .

# ACI will inject these via environment variables:
# AzureAd__TenantId, AzureAd__ClientId
# TokenCache__StorageConnectionString
ENV ASPNETCORE_ENVIRONMENT=Production
ENV DOTNET_RUNNING_IN_CONTAINER=true

EXPOSE 8080
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD wget -qO- http://localhost:8080/health || exit 1

ENTRYPOINT ["dotnet", "MSGraphMCP.dll"]
