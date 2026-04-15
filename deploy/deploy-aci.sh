#!/usr/bin/env bash
# =============================================================================
# deploy-aci.sh — Deploy MSGraphMCP to Azure Container Instances
#
# Usage:
#   chmod +x deploy-aci.sh
#   ./deploy-aci.sh
#
# Prerequisites:
#   - Azure CLI installed and logged in (az login)
#   - Docker installed and running
#   - Populate the variables in the CONFIGURATION section below
# =============================================================================

set -euo pipefail

# ─── CONFIGURATION ────────────────────────────────────────────────────────────
# Edit these values before running

RESOURCE_GROUP="rg-msgraph-mcp"
LOCATION="eastus"

# Azure Container Registry
ACR_NAME="msgraphmcpacr"           # Must be globally unique, lowercase, 5-50 chars

# Azure Storage Account (for MSAL token cache)
STORAGE_ACCOUNT="msgraphmcpstore"  # Must be globally unique, lowercase, 3-24 chars
STORAGE_CONTAINER="msal-token-cache"

# Azure Container Instance
ACI_NAME="msgraph-mcp"
ACI_DNS_LABEL="msgraph-mcp-${RANDOM}"   # Must be unique in the region
IMAGE_TAG="latest"

# Azure AD App Registration (from portal.azure.com)
AAD_TENANT_ID="${AAD_TENANT_ID:-}"      # Set via env or hardcode here
AAD_CLIENT_ID="${AAD_CLIENT_ID:-}"      # Set via env or hardcode here

# ─── VALIDATION ───────────────────────────────────────────────────────────────

if [[ -z "$AAD_TENANT_ID" || -z "$AAD_CLIENT_ID" ]]; then
  echo "❌  Set AAD_TENANT_ID and AAD_CLIENT_ID environment variables before running."
  exit 1
fi

echo "🚀 Deploying MSGraphMCP to Azure Container Instances"
echo "   Resource Group : $RESOURCE_GROUP ($LOCATION)"
echo "   ACR            : $ACR_NAME"
echo "   ACI            : $ACI_NAME"
echo ""

# ─── 1. RESOURCE GROUP ────────────────────────────────────────────────────────

echo "📦 Creating resource group..."
az group create \
  --name "$RESOURCE_GROUP" \
  --location "$LOCATION" \
  --output none

# ─── 2. AZURE CONTAINER REGISTRY ─────────────────────────────────────────────

echo "🐳 Creating Azure Container Registry..."
az acr create \
  --resource-group "$RESOURCE_GROUP" \
  --name "$ACR_NAME" \
  --sku Basic \
  --admin-enabled true \
  --output none

ACR_LOGIN_SERVER=$(az acr show --name "$ACR_NAME" --query loginServer -o tsv)
ACR_USERNAME=$(az acr credential show --name "$ACR_NAME" --query username -o tsv)
ACR_PASSWORD=$(az acr credential show --name "$ACR_NAME" --query "passwords[0].value" -o tsv)

# ─── 3. BUILD & PUSH IMAGE ───────────────────────────────────────────────────

echo "🔨 Building Docker image..."
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(dirname "$SCRIPT_DIR")"
FULL_IMAGE="${ACR_LOGIN_SERVER}/msgraphmcp:${IMAGE_TAG}"

docker build -t "$FULL_IMAGE" "$REPO_ROOT"

echo "⬆️  Pushing image to ACR..."
echo "$ACR_PASSWORD" | docker login "$ACR_LOGIN_SERVER" -u "$ACR_USERNAME" --password-stdin
docker push "$FULL_IMAGE"

# ─── 4. STORAGE ACCOUNT (MSAL TOKEN CACHE) ───────────────────────────────────

echo "💾 Creating Storage Account for MSAL token cache..."
az storage account create \
  --name "$STORAGE_ACCOUNT" \
  --resource-group "$RESOURCE_GROUP" \
  --location "$LOCATION" \
  --sku Standard_LRS \
  --kind StorageV2 \
  --min-tls-version TLS1_2 \
  --output none

STORAGE_CONN_STR=$(az storage account show-connection-string \
  --name "$STORAGE_ACCOUNT" \
  --resource-group "$RESOURCE_GROUP" \
  --query connectionString -o tsv)

echo "   Creating blob container: $STORAGE_CONTAINER"
az storage container create \
  --name "$STORAGE_CONTAINER" \
  --connection-string "$STORAGE_CONN_STR" \
  --output none

# ─── 5. DEPLOY CONTAINER INSTANCE ────────────────────────────────────────────

echo "🚢 Deploying Azure Container Instance..."
az container create \
  --resource-group "$RESOURCE_GROUP" \
  --name "$ACI_NAME" \
  --image "$FULL_IMAGE" \
  --registry-login-server "$ACR_LOGIN_SERVER" \
  --registry-username "$ACR_USERNAME" \
  --registry-password "$ACR_PASSWORD" \
  --dns-name-label "$ACI_DNS_LABEL" \
  --ports 8080 \
  --protocol TCP \
  --os-type Linux \
  --cpu 1 \
  --memory 1.5 \
  --restart-policy Always \
  --environment-variables \
    ASPNETCORE_ENVIRONMENT=Production \
    AzureAd__TenantId="$AAD_TENANT_ID" \
    AzureAd__ClientId="$AAD_CLIENT_ID" \
    TokenCache__ContainerName="$STORAGE_CONTAINER" \
  --secure-environment-variables \
    TokenCache__StorageConnectionString="$STORAGE_CONN_STR" \
  --output none

# ─── 6. GET PUBLIC URL ────────────────────────────────────────────────────────

FQDN=$(az container show \
  --resource-group "$RESOURCE_GROUP" \
  --name "$ACI_NAME" \
  --query ipAddress.fqdn -o tsv)

echo ""
echo "✅ Deployment complete!"
echo ""
echo "   MCP Server URL : http://${FQDN}:8080/mcp"
echo "   Health Check   : http://${FQDN}:8080/health"
echo ""
echo "   Use in Relevance AI:"
echo "   POST http://${FQDN}:8080/mcp"
echo "   Header: X-Session-Id: <sessionId>"
echo ""
echo "   Next steps:"
echo "   1. Add HTTPS via Azure Application Gateway or API Management"
echo "   2. Restrict ACI to VNet if Relevance AI supports private endpoints"
echo "   3. Set up Azure Monitor alerts on /health"
echo ""
echo "   To view logs:"
echo "   az container logs --resource-group $RESOURCE_GROUP --name $ACI_NAME --follow"
