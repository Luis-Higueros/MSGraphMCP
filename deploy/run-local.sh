#!/usr/bin/env bash
set -euo pipefail

NO_AZURITE="${NO_AZURITE:-0}"

# Uses dedicated queue/table ports to avoid conflicts with other local Azurite instances.
AZURITE_CONN="DefaultEndpointsProtocol=http;AccountName=devstoreaccount1;AccountKey=Eby8vdM02xNOcqFlqUwJPLlmEtlCDXJ1OUzFT50uSRZ6IFsuFq2UVErCz4I6tq/K1SZFPTOtr/KBHBeksoGMGw==;BlobEndpoint=http://127.0.0.1:10000/devstoreaccount1;QueueEndpoint=http://127.0.0.1:11001/devstoreaccount1;TableEndpoint=http://127.0.0.1:11002/devstoreaccount1;"

if [[ "$NO_AZURITE" != "1" ]]; then
  if ! command -v azurite >/dev/null 2>&1; then
    echo "Azurite is not installed. Install with: npm i -g azurite"
    exit 1
  fi

  if ! (echo > /dev/tcp/127.0.0.1/10000) >/dev/null 2>&1; then
    mkdir -p .azurite
    azurite --silent --location .azurite --blobHost 127.0.0.1 --blobPort 10000 --queueHost 127.0.0.1 --queuePort 11001 --tableHost 127.0.0.1 --tablePort 11002 >/tmp/azurite.log 2>&1 &
    sleep 2
  fi
fi

export ASPNETCORE_ENVIRONMENT=Development
export TokenCache__StorageConnectionString="$AZURITE_CONN"

echo "Starting MSGraphMCP with local Azurite token cache..."
dotnet run --project src/MSGraphMCP
