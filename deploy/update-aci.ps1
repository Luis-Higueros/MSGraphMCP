# =============================================================================
# update-aci.ps1 - Rebuild image in ACR and restart existing ACI container
#
# Usage:
#   .\deploy\update-aci.ps1
#
# Prerequisites:
#   - Azure CLI installed and logged in (az login)
#   - Resources already provisioned
# =============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Live infrastructure
$TenantId      = "425a5546-5a6e-4f1b-ab62-23d91d07d893"
$ClientId      = "ba14f7ed-4216-450f-a2ff-7a93ae92fc74"
$Subscription  = "380d9153-35d4-45fd-a604-fe72aaf453ae"
$ResourceGroup = "euw-ea_teamsai-sandbox-rg"
$AcrName       = "msgraphmcp95932"
$AciName       = "msgraph-mcp-27992"
$ImageName     = "msgraphmcp"
$ImageTag      = "latest"
$FrontDoorHost = "ep-msgraphmcp-43613-c6dvbtfyfccmhzf8.a03.azurefd.net"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot  = Split-Path -Parent $ScriptDir
if (-not (Test-Path (Join-Path $RepoRoot "Dockerfile"))) {
    $RepoRoot = $ScriptDir
}

Write-Host ""
Write-Host "MSGraphMCP - Update deployment" -ForegroundColor Cyan
Write-Host "   Subscription  : $Subscription"
Write-Host "   Resource Group: $ResourceGroup"
Write-Host "   ACR           : $AcrName"
Write-Host "   ACI           : $AciName"
Write-Host ""

az account set --subscription $Subscription | Out-Null

Write-Host "Building image in ACR (cloud build)..." -ForegroundColor Yellow
az acr build `
    --registry $AcrName `
    --image "${ImageName}:${ImageTag}" `
    (Resolve-Path $RepoRoot).Path `
    --file (Join-Path $RepoRoot "Dockerfile")

Write-Host ""
Write-Host "Build complete. New image: $AcrName.azurecr.io/${ImageName}:${ImageTag}" -ForegroundColor Green

Write-Host ""
Write-Host "Restarting ACI container group: $AciName..." -ForegroundColor Yellow
az container stop  --resource-group $ResourceGroup --name $AciName --output none
az container start --resource-group $ResourceGroup --name $AciName --output none

Write-Host ""
Write-Host "Waiting for container to become healthy..." -ForegroundColor Yellow

$AciFqdn = (az container show `
    --resource-group $ResourceGroup `
    --name $AciName `
    --query ipAddress.fqdn -o tsv).Trim()

$AciHealth = "http://$AciFqdn:8080/health"
$FdHealth  = "https://$FrontDoorHost/health"

$healthy = $false
for ($i = 1; $i -le 30; $i++) {
    Start-Sleep -Seconds 6
    try {
        $resp = Invoke-WebRequest -Uri $AciHealth -UseBasicParsing -TimeoutSec 10
        if ($resp.Content -match "Healthy") {
            $healthy = $true
            break
        }
    }
    catch {
        # continue waiting
    }

    Write-Host "   Attempt $i/30 - waiting..." -ForegroundColor DarkGray
}

Write-Host ""
if ($healthy) {
    Write-Host "Container is healthy." -ForegroundColor Green
}
else {
    Write-Host "Container did not respond healthy within timeout." -ForegroundColor Yellow
    Write-Host "   Check logs: az container logs -g $ResourceGroup -n $AciName --follow"
}

Write-Host ""
Write-Host "Deployment updated." -ForegroundColor Cyan
Write-Host ""
Write-Host "   ACI direct : http://$AciFqdn:8080/mcp"
Write-Host "   Front Door : https://$FrontDoorHost/mcp (preferred)" -ForegroundColor Green
Write-Host "   Health     : $FdHealth"
Write-Host ""
Write-Host "   Stream logs: az container logs -g $ResourceGroup -n $AciName --follow"
