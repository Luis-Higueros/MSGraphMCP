param(
    [string]$Subscription = "380d9153-35d4-45fd-a604-fe72aaf453ae",
    [string]$ResourceGroup = "euw-ea_teamsai-sandbox-rg",
    [string]$AciName = "msgraph-mcp-27992",
    [string]$NotifyEmail = "luis.higueros@aveva.com"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$aciResourceId = "/subscriptions/$Subscription/resourceGroups/$ResourceGroup/providers/Microsoft.ContainerInstance/containerGroups/$AciName"
$actionGroupName = "aci-msgraph-ops-ag"
$alertName = "aci-msgraph-stop-alert"

Write-Host "Configuring alerting for ACI: $AciName" -ForegroundColor Cyan
az account set --subscription $Subscription | Out-Null

$existingAg = az monitor action-group list --resource-group $ResourceGroup --query "[?name=='$actionGroupName'].id" -o tsv
if (-not $existingAg) {
    Write-Host "Creating action group: $actionGroupName" -ForegroundColor Yellow
    az monitor action-group create --name $actionGroupName --short-name aciops --resource-group $ResourceGroup --action email ops-email $NotifyEmail --output none
}

$actionGroupId = az monitor action-group list --resource-group $ResourceGroup --query "[?name=='$actionGroupName'].id | [0]" -o tsv

$existingAlert = az monitor activity-log alert list --resource-group $ResourceGroup --query "[?name=='$alertName'].id" -o tsv
if ($existingAlert) {
    Write-Host "Deleting existing alert: $alertName" -ForegroundColor Yellow
    az monitor activity-log alert delete --name $alertName --resource-group $ResourceGroup
}

Write-Host "Creating alert: $alertName" -ForegroundColor Yellow
az monitor activity-log alert create --name $alertName --resource-group $ResourceGroup --scope "/subscriptions/$Subscription" --condition "category=Administrative and operationName=Microsoft.ContainerInstance/containerGroups/stop/action and resourceId=$aciResourceId and status=Succeeded" --action-group $actionGroupId --description "Email alert when MSGraphMCP ACI is explicitly stopped" --output none

Write-Host "Alert setup complete." -ForegroundColor Green
Write-Host "Action Group: $actionGroupName"
Write-Host "Alert Name  : $alertName"
Write-Host "Notify Email: $NotifyEmail"