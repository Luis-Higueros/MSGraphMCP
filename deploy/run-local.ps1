param(
    [switch]$NoAzurite
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$azuritePath = Join-Path $repoRoot ".azurite"
$projectPath = Join-Path $repoRoot "src/MSGraphMCP"
$builtDllPath = Join-Path $projectPath "bin/Debug/net9.0/MSGraphMCP.dll"

# Uses dedicated queue/table ports to avoid conflicts with other local Azurite instances.
$azuriteConn = "DefaultEndpointsProtocol=http;AccountName=devstoreaccount1;AccountKey=Eby8vdM02xNOcqFlqUwJPLlmEtlCDXJ1OUzFT50uSRZ6IFsuFq2UVErCz4I6tq/K1SZFPTOtr/KBHBeksoGMGw==;BlobEndpoint=http://127.0.0.1:10000/devstoreaccount1;QueueEndpoint=http://127.0.0.1:11001/devstoreaccount1;TableEndpoint=http://127.0.0.1:11002/devstoreaccount1;"

function Test-TcpPort {
    param([string]$HostName, [int]$Port)

    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $client.Connect($HostName, $Port)
        return $client.Connected
    }
    catch {
        return $false
    }
    finally {
        $client.Dispose()
    }
}

if (-not $NoAzurite) {
    if (-not (Get-Command azurite -ErrorAction SilentlyContinue)) {
        throw "Azurite is not installed. Install with: npm i -g azurite"
    }

    if (-not (Test-TcpPort -HostName "127.0.0.1" -Port 10000)) {
        New-Item -ItemType Directory -Force -Path $azuritePath | Out-Null
        Start-Process -FilePath "azurite" -ArgumentList "--silent --location `"$azuritePath`" --blobHost 127.0.0.1 --blobPort 10000 --queueHost 127.0.0.1 --queuePort 11001 --tableHost 127.0.0.1 --tablePort 11002" | Out-Null
        Start-Sleep -Seconds 2
    }
}

$env:ASPNETCORE_ENVIRONMENT = "Development"
$env:TokenCache__StorageConnectionString = $azuriteConn

if (Test-TcpPort -HostName "127.0.0.1" -Port 8080) {
    Write-Host "MSGraphMCP is already running on http://127.0.0.1:8080"
    Write-Host "Stop the existing process first if you want to restart it."
    return
}

Write-Host "Starting MSGraphMCP with local Azurite token cache..."
if (Test-Path $builtDllPath) {
    dotnet run --project $projectPath --no-build
}
else {
    dotnet run --project $projectPath
}
