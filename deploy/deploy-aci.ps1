# =============================================================================
# deploy-aci.ps1 — Deploy MSGraphMCP to Azure Container Instances (PowerShell)
#
# Usage:
#   $env:AAD_TENANT_ID = "your-tenant-id"
#   $env:AAD_CLIENT_ID = "your-client-id"
#   .\deploy-aci.ps1
#
# Prerequisites:
#   - Azure CLI installed and logged in (az login)
#   - Docker Desktop running
# =============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ─── CONFIGURATION ────────────────────────────────────────────────────────────

$ResourceGroup    = "rg-msgraph-mcp"
$Location         = "eastus"
$AcrName          = "msgraphmcpacr"
$StorageAccount   = "msgraphmcpstore"
$StorageContainer = "msal-token-cache"
$AciName          = "msgraph-mcp"
$AciDnsLabel      = "msgraph-mcp-$(Get-Random -Maximum 9999)"
$ImageTag         = "latest"

$TenantId = $env:AAD_TENANT_ID
$ClientId = $env:AAD_CLIENT_ID

if (-not $TenantId -or -not $ClientId) {
    Write-Error "Set AAD_TENANT_ID and AAD_CLIENT_ID environment variables before running."
    exit 1
}

Write-Host "🚀 Deploying MSGraphMCP to Azure Container Instances" -ForegroundColor Cyan
Write-Host "   Resource Group : $ResourceGroup ($Location)"
Write-Host "   ACR            : $AcrName"
Write-Host "   ACI            : $AciName"
Write-Host ""

# ─── 1. RESOURCE GROUP ────────────────────────────────────────────────────────

Write-Host "📦 Creating resource group..." -ForegroundColor Yellow
az group create --name $ResourceGroup --location $Location --output none

# ─── 2. AZURE CONTAINER REGISTRY ─────────────────────────────────────────────

Write-Host "🐳 Creating Azure Container Registry..." -ForegroundColor Yellow
az acr create `
    --resource-group $ResourceGroup `
    --name $AcrName `
    --sku Basic `
    --admin-enabled true `
    --output none

$AcrLoginServer = az acr show --name $AcrName --query loginServer -o tsv
$AcrUsername    = az acr credential show --name $AcrName --query username -o tsv
$AcrPassword    = az acr credential show --name $AcrName --query "passwords[0].value" -o tsv

# ─── 3. BUILD & PUSH IMAGE ───────────────────────────────────────────────────

Write-Host "🔨 Building Docker image..." -ForegroundColor Yellow
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot  = Split-Path -Parent $ScriptDir
$FullImage = "${AcrLoginServer}/msgraphmcp:${ImageTag}"

docker build -t $FullImage $RepoRoot
if ($LASTEXITCODE -ne 0) { Write-Error "Docker build failed"; exit 1 }

Write-Host "⬆️  Pushing image to ACR..." -ForegroundColor Yellow
$AcrPassword | docker login $AcrLoginServer -u $AcrUsername --password-stdin
docker push $FullImage
if ($LASTEXITCODE -ne 0) { Write-Error "Docker push failed"; exit 1 }

# ─── 4. STORAGE ACCOUNT ──────────────────────────────────────────────────────

Write-Host "💾 Creating Storage Account for MSAL token cache..." -ForegroundColor Yellow
az storage account create `
    --name $StorageAccount `
    --resource-group $ResourceGroup `
    --location $Location `
    --sku Standard_LRS `
    --kind StorageV2 `
    --min-tls-version TLS1_2 `
    --output none

$StorageConnStr = az storage account show-connection-string `
    --name $StorageAccount `
    --resource-group $ResourceGroup `
    --query connectionString -o tsv

az storage container create `
    --name $StorageContainer `
    --connection-string $StorageConnStr `
    --output none

# ─── 5. DEPLOY CONTAINER INSTANCE ────────────────────────────────────────────

Write-Host "🚢 Deploying Azure Container Instance..." -ForegroundColor Yellow
az container create `
    --resource-group $ResourceGroup `
    --name $AciName `
    --image $FullImage `
    --registry-login-server $AcrLoginServer `
    --registry-username $AcrUsername `
    --registry-password $AcrPassword `
    --dns-name-label $AciDnsLabel `
    --ports 8080 `
    --protocol TCP `
    --os-type Linux `
    --cpu 1 `
    --memory 1.5 `
    --restart-policy Always `
    --environment-variables `
        ASPNETCORE_ENVIRONMENT=Production `
        "AzureAd__TenantId=$TenantId" `
        "AzureAd__ClientId=$ClientId" `
        "TokenCache__ContainerName=$StorageContainer" `
    --secure-environment-variables `
        "TokenCache__StorageConnectionString=$StorageConnStr" `
    --output none

# ─── 6. RESULTS ───────────────────────────────────────────────────────────────

$Fqdn = az container show `
    --resource-group $ResourceGroup `
    --name $AciName `
    --query ipAddress.fqdn -o tsv

Write-Host ""
Write-Host "✅ Deployment complete!" -ForegroundColor Green
Write-Host ""
Write-Host "   MCP Server URL : http://${Fqdn}:8080/mcp" -ForegroundColor Cyan
Write-Host "   Health Check   : http://${Fqdn}:8080/health"
Write-Host ""
Write-Host "   To stream logs:"
Write-Host "   az container logs --resource-group $ResourceGroup --name $AciName --follow"
Write-Host ""
Write-Host "   ⚠️  Add HTTPS before exposing to Relevance AI (use Azure API Management or App Gateway)"
