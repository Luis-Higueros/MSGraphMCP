# =============================================================================
# update-aci.ps1 — Rebuild image in ACR and restart the existing ACI container
#
# Run this whenever you make code changes and want to push them to Azure.
# No Docker required — image is built in Azure via az acr build.
#
# Usage (from repo root or deploy/ folder):
#   .\deploy\update-aci.ps1
#
# Prerequisites:
#   - Azure CLI installed and logged in (az login)
#   - Resources already provisioned (run deploy-aci.ps1 once to set up)
# =============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ─── LIVE INFRASTRUCTURE ─────────────────────────────────────────────────────
# These values match the deployed resources in your sandbox resource group.
# Update them if you ever provision a new environment.

$TenantId      = "425a5546-5a6e-4f1b-ab62-23d91d07d893"
$ClientId      = "ba14f7ed-4216-450f-a2ff-7a93ae92fc74"
$Subscription  = "380d9153-35d4-45fd-a604-fe72aaf453ae"
$ResourceGroup = "euw-ea_teamsai-sandbox-rg"
$AcrName       = "msgraphmcp95932"
$AciName       = "msgraph-mcp-27992"
$ImageName     = "msgraphmcp"
$ImageTag      = "latest"

# Front Door endpoint (for reference — updated automatically, no action needed)
$FrontDoorHost = "ep-msgraphmcp-43613-c6dvbtfyfccmhzf8.a03.azurefd.net"

# ─── RESOLVE REPO ROOT ────────────────────────────────────────────────────────

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot  = Split-Path -Parent $ScriptDir
if (-not (Test-Path (Join-Path $RepoRoot "Dockerfile"))) {
    # Fallback: assume we're already at repo root
    $RepoRoot = $ScriptDir
}

# ─── START ────────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "MSGraphMCP — Update deployment" -ForegroundColor Cyan
Write-Host "   Subscription : $Subscription"
Write-Host "   Resource Group: $ResourceGroup"
Write-Host "   ACR           : $AcrName"
Write-Host "   ACI           : $AciName"
Write-Host ""

az account set --subscription $Subscription | Out-Null

# ─── 1. BUILD & PUSH IMAGE IN ACR (no Docker needed) ─────────────────────────

Write-Host "Building image in ACR (cloud build)..." -ForegroundColor Yellow
az acr build `
    --registry $AcrName `
    --image "${ImageName}:${ImageTag}" `
    (Resolve-Path $RepoRoot).Path `
    --file (Join-Path $RepoRoot "Dockerfile")

Write-Host ""
Write-Host "Build complete. New image: $AcrName.azurecr.io/${ImageName}:${ImageTag}" -ForegroundColor Green

# ─── 2. RESTART ACI TO PULL NEW IMAGE ────────────────────────────────────────
# ACI does not have a native "restart" command, so we stop then start the
# container group, which forces it to pull the latest image from ACR.

Write-Host ""
Write-Host "Restarting ACI container group: $AciName..." -ForegroundColor Yellow
az container stop  --resource-group $ResourceGroup --name $AciName --output none
az container start --resource-group $ResourceGroup --name $AciName --output none

# ─── 3. WAIT FOR HEALTHY ─────────────────────────────────────────────────────

Write-Host ""
Write-Host "Waiting for container to become healthy..." -ForegroundColor Yellow

$AciFqdn = (az container show `
    --resource-group $ResourceGroup `
    --name $AciName `
    --query ipAddress.fqdn -o tsv).Trim()

$AciHealth  = "http://$AciFqdn:8080/health"
$FdHealth   = "https://$FrontDoorHost/health"
$FdMcp      = "https://$FrontDoorHost/mcp"

$healthy = $false
for ($i = 1; $i -le 30; $i++) {
    Start-Sleep -Seconds 6
    try {
        $resp = Invoke-WebRequest -Uri $AciHealth -UseBasicParsing -TimeoutSec 10
        if ($resp.Content -match "Healthy") {
            $healthy = $true
            break
        }
    } catch { }
    Write-Host "   Attempt $i/30 — waiting..." -ForegroundColor DarkGray
}

Write-Host ""
if ($healthy) {
    Write-Host "Container is healthy!" -ForegroundColor Green
} else {
    Write-Host "Container did not respond healthy within timeout." -ForegroundColor Yellow
    Write-Host "   Check logs: az container logs -g $ResourceGroup -n $AciName --follow"
}

# ─── 4. SUMMARY ───────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "Deployment updated." -ForegroundColor Cyan
Write-Host ""
Write-Host "   ACI direct  : http://$AciFqdn:8080/mcp"
Write-Host "   Front Door  : https://$FrontDoorHost/mcp   (preferred)" -ForegroundColor Green
Write-Host "   Health      : $FdHealth"
Write-Host ""
Write-Host "   Stream logs : az container logs -g $ResourceGroup -n $AciName --follow"
