param(
    [string]$CorruptedExtensionName = 'application-insights'
)

$ErrorActionPreference = 'Stop'

Write-Host 'Fixing Azure CLI extension cache...' -ForegroundColor Cyan

$extDir = Join-Path $env:USERPROFILE '.azure\cliextensions'
$target = Join-Path $extDir $CorruptedExtensionName

if (Test-Path $target) {
    try {
        Remove-Item -Recurse -Force $target
        Write-Host "Removed: $target" -ForegroundColor Yellow
    }
    catch {
        Write-Host "Could not remove locked extension folder: $target" -ForegroundColor Yellow
        Write-Host "Applying permanent workaround by moving Azure CLI extension directory." -ForegroundColor Yellow
    }
}
else {
    Write-Host "No local cache found for extension: $CorruptedExtensionName" -ForegroundColor DarkGray
}

$newExtensionDir = Join-Path $env:USERPROFILE '.azure\\cliextensions-clean'
New-Item -ItemType Directory -Force -Path $newExtensionDir | Out-Null
az config set extension.dir=$newExtensionDir | Out-Null

az config set extension.use_dynamic_install=yes_without_prompt | Out-Null
az config set extension.dynamic_install_allow_preview=true | Out-Null

Write-Host 'Azure CLI extension dynamic-install settings updated.' -ForegroundColor Green
Write-Host "Extension directory: $newExtensionDir" -ForegroundColor Green
az account show --query "{subscription:id,user:user.name}" -o table
