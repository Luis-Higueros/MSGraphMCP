param(
    [switch]$NoAzurite
)

$ErrorActionPreference = "Stop"

# Uses dedicated queue/table ports to avoid conflicts with other local Azurite instances.
$azuriteConn = "DefaultEndpointsProtocol=http;AccountName=devstoreaccount1;AccountKey=Eby8vdM02xNOcqFlqUwJPLlmEtlCDXJ1OUzFT50uSRZ6IFsuFq2UVErCz4I6tq/K1SZFPTOtr/KBHBeksoGMGw==;BlobEndpoint=http://127.0.0.1:10000/devstoreaccount1;QueueEndpoint=http://127.0.0.1:11001/devstoreaccount1;TableEndpoint=http://127.0.0.1:11002/devstoreaccount1;"

function Test-TcpPort {
    param([string]$Host, [int]$Port)

    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $client.Connect($Host, $Port)
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

    if (-not (Test-TcpPort -Host "127.0.0.1" -Port 10000)) {
        New-Item -ItemType Directory -Force -Path ".azurite" | Out-Null
        Start-Process -FilePath "azurite" -ArgumentList "--silent --location .azurite --blobHost 127.0.0.1 --blobPort 10000 --queueHost 127.0.0.1 --queuePort 11001 --tableHost 127.0.0.1 --tablePort 11002" | Out-Null
        Start-Sleep -Seconds 2
    }
}

$env:ASPNETCORE_ENVIRONMENT = "Development"
$env:TokenCache__StorageConnectionString = $azuriteConn

Write-Host "Starting MSGraphMCP with local Azurite token cache..."
dotnet run --project src/MSGraphMCP
