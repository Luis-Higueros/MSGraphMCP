# Connection Reset Error - Analysis and Mitigation

## Problem
Relevance reports: `MCP network error: ('Connection aborted.', ConnectionResetError(104, 'Connection reset by peer'))`

This occurs **after a couple of successful calls**, indicating an idle connection timeout issue.

## Root Cause
**Azure Front Door Idle Connection Timeout**

1. Relevance establishes HTTP connection to Front Door endpoint
2. Makes several successful MCP requests using that connection
3. If gap between requests exceeds Front Door's idle timeout (~60-90 seconds), Front Door closes the connection
4. When Relevance tries to reuse the closed connection → **Connection Reset Error (errno 104)**

## Evidence
- Error occurs "after a couple of other calls" (not immediately)
- Error type: `ConnectionResetError(104)` - server actively closed connection  
- Pattern consistent with proxy/load balancer idle timeout behavior

## Solutions

### 1. Client-Side Fix (RECOMMENDED - Relevance Team)
Relevance should implement **automatic retry logic** for connection errors:

```python
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import requests

# Create session with retry logic
session = requests.Session()
retry_strategy = Retry(
    total=3,                          # Retry up to 3 times
    backoff_factor=1,                 # Wait 1, 2, 4 seconds between retries
    status_forcelist=[104, 502, 503, 504],  # Retry on these errors
    allowed_methods=["POST", "GET"]   # Methods to retry
)
adapter = HTTPAdapter(max_retries=retry_strategy)
session.mount("https://", adapter)
session.mount("http://", adapter)

# Use this session for all MCP requests
response = session.post(mcp_endpoint, json=payload)
```

**Why this is the best solution:**
- Handles connection resets transparently
- Works regardless of timeout configuration
- Industry standard pattern for resilient HTTP clients
- Minimal performance impact (only retries on actual failures)

### 2. Server-Side Mitigations (MSGraphMCP)

#### A. Extended Kestrel Keep-Alive (✅ APPLIED)
Added to `Program.cs`:
```csharp
builder.Services.Configure<KestrelServerOptions>(options =>
{
    options.Limits.KeepAliveTimeout = TimeSpan.FromMinutes(5);
    options.Limits.RequestHeadersTimeout = TimeSpan.FromMinutes(2);
});
```

**Impact:** Keeps application server connections alive longer, but **does not control Front Door timeout**.

#### B. Front Door Backend Timeout Configuration (TODO)
Requires Azure portal or CLI with appropriate permissions:
```bash
# Check current timeout (Standard/Premium tier)
az network front-door backend-pool show \
  --front-door-name <name> \
  --resource-group <rg> \
  --name DefaultBackendPool \
  --query 'backends[].sendRecvTimeoutSeconds'

# Increase timeout if needed (max varies by tier)
# Typically 60s default, can go up to 120s on Premium
```

**Note:** Front Door timeout increases might not fully solve the issue if Relevance has very long gaps between requests.

### 3. Workarounds (Not Recommended)

#### Connection Keep-Alive Pings
Could implement periodic "ping" requests from client to keep connection alive.
- **Downside:** Wasteful, adds unnecessary traffic
- **Better:** Just implement retry logic

## Diagnostic Data Needed

To confirm root cause, need to know:
1. **Time gap between Relevance requests** when error occurs (in seconds)
2. **Does immediate retry succeed?** (indicates stale connection, not app crash)
3. **Frequency:** Every session or intermittent?

## Testing
Test script created to reproduce: `deploy/test-connection-reset.ps1`
```powershell
# Test with 70-second delays to trigger timeout
.\deploy\test-connection-reset.ps1 -DelaySeconds 70 -RequestCount 3
```

## Monitoring

### KQL Query - Detect Connection Reset Patterns
```kql
traces
| where timestamp > ago(1h)
| where message contains "error" or message contains "connection" or message contains "reset"
| summarize count() by bin(timestamp, 5m), message
| order by timestamp desc
```

### Check for App Crashes vs Connection Issues
```kql
requests
| where timestamp > ago(1h)
| summarize 
    successCount = countif(success == true),
    failureCount = countif(success == false),
    avgDuration = avg(duration)
    by bin(timestamp, 5m)
| order by timestamp desc
```

If `failureCount` is low/zero but Relevance sees connection resets → confirms client-side connection reuse issue.

## Recommended Action Plan

1. **Immediate (Relevance Team):** Implement retry logic in MCP client
2. **Short-term (MSGraphMCP):** Deploy Kestrel keep-alive update (already prepared)
3. **Optional (MSGraphMCP):** Increase Front Door backend timeout if possible
4. **Long-term:** Monitor connection error rates via App Insights

## Status
- ✅ Kestrel keep-alive configuration added
- ⏳ Testing to confirm 60-70s timeout threshold
- 📋 Need Relevance team to implement retry logic
- 📋 Need to check Front Door timeout configuration (permissions required)
