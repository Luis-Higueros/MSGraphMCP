#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Test for connection reset issues by simulating Relevance's connection reuse pattern
.DESCRIPTION
    Makes multiple requests with delays to trigger idle timeout and connection reset scenarios
#>

param(
    [string]$Endpoint = "https://ep-msgraphmcp-43613-c6dvbtfyfccmhzf8.a03.azurefd.net",
    [int]$RequestCount = 5,
    [int]$DelaySeconds = 30
)

$ErrorActionPreference = "Continue"

Write-Host "🔍 Testing connection reset scenario" -ForegroundColor Cyan
Write-Host "Endpoint: $Endpoint"
Write-Host "Requests: $RequestCount"
Write-Host "Delay between requests: $DelaySeconds seconds"
Write-Host ""

# Test payload - simple health check
$healthUri = "$Endpoint/health"

# Test with keep-alive connection (similar to how HTTP clients work)
$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

for ($i = 1; $i -le $RequestCount; $i++) {
    Write-Host "[$i/$RequestCount] Making request at $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor Yellow
    
    try {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $response = Invoke-WebRequest -Uri $healthUri -WebSession $session -TimeoutSec 60 -UseBasicParsing
        $sw.Stop()
        
        Write-Host "  ✅ Success: $($response.StatusCode) in $($sw.ElapsedMilliseconds)ms" -ForegroundColor Green
        Write-Host "  Response: $($response.Content)" -ForegroundColor Gray
    }
    catch {
        Write-Host "  ❌ FAILED: $($_.Exception.Message)" -ForegroundColor Red
        
        # Check for connection reset
        if ($_.Exception.Message -match "reset|abort|connection.*closed") {
            Write-Host "  ⚠️  CONNECTION RESET DETECTED" -ForegroundColor Magenta
            Write-Host "  This indicates the server/proxy closed the connection" -ForegroundColor Magenta
        }
        
        # Print full exception for diagnosis
        Write-Host "  Exception type: $($_.Exception.GetType().FullName)" -ForegroundColor Gray
        if ($_.Exception.InnerException) {
            Write-Host "  Inner exception: $($_.Exception.InnerException.Message)" -ForegroundColor Gray
        }
    }
    
    # Wait before next request (simulates delay between Relevance calls)
    if ($i -lt $RequestCount) {
        Write-Host "  ⏳ Waiting $DelaySeconds seconds before next request..." -ForegroundColor Cyan
        Start-Sleep -Seconds $DelaySeconds
        Write-Host ""
    }
}

Write-Host ""
Write-Host "🏁 Test complete" -ForegroundColor Cyan

# Summary guidance
Write-Host ""
Write-Host "📋 If connection resets occurred:" -ForegroundColor Yellow
Write-Host "  1. Check Front Door timeout settings (default ~60s for Standard tier)"
Write-Host "  2. Check if ACI has resource constraints or is restarting"
Write-Host "  3. Consider enabling Front Door session affinity"
Write-Host "  4. Relevance may need to implement connection retry logic"
