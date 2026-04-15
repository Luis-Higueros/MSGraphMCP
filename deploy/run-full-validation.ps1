param(
    [string]$UserHint,
    [string]$GraphSessionId,
    [ValidateSet('dev','aci','afd')]
    [string]$Target = 'afd',
    [string]$ReportsDir = 'deploy/reports'
)

$ErrorActionPreference = 'Stop'

New-Item -ItemType Directory -Path $ReportsDir -Force | Out-Null
$ts = Get-Date -Format 'yyyyMMdd-HHmmss'
$logFile = Join-Path $ReportsDir ("full-validation-{0}.log" -f $ts)
$allToolsJson = Join-Path $ReportsDir ("all-tools-fullrun-{0}.json" -f $ts)
$mailSummJson = Join-Path $ReportsDir ("mail-summarize-context-{0}.json" -f $ts)

function Write-Log {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format 'o'), $Message
    Add-Content -Path $logFile -Value $line -Encoding utf8
}

function Run-ScriptStep {
    param(
        [string]$Name,
        [string]$ScriptPath,
        [string[]]$StepArgs
    )

    Write-Log "=== $Name ==="
    Write-Log ("Command: powershell -NoProfile -ExecutionPolicy Bypass -File {0} {1}" -f $ScriptPath, ($StepArgs -join ' '))

    $outFile = Join-Path $env:TEMP ("val_out_{0}.log" -f ([Guid]::NewGuid().ToString('N')))
    $errFile = Join-Path $env:TEMP ("val_err_{0}.log" -f ([Guid]::NewGuid().ToString('N')))

    try {
        $argList = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $ScriptPath) + $StepArgs
        $proc = Start-Process -FilePath "powershell" -ArgumentList $argList -RedirectStandardOutput $outFile -RedirectStandardError $errFile -NoNewWindow -PassThru -Wait

        if (Test-Path $outFile) {
            Get-Content $outFile | ForEach-Object { Write-Log ("[stdout] {0}" -f $_) }
        }
        if (Test-Path $errFile) {
            Get-Content $errFile | ForEach-Object { Write-Log ("[stderr] {0}" -f $_) }
        }

        Write-Log ("ExitCode: {0}" -f $proc.ExitCode)
        return [int]$proc.ExitCode
    }
    finally {
        Remove-Item $outFile -ErrorAction SilentlyContinue
        Remove-Item $errFile -ErrorAction SilentlyContinue
    }
}

Write-Log "Full validation started"
Write-Log ("Target={0}" -f $Target)
Write-Log ("UserHint={0}" -f $UserHint)
Write-Log ("Input GraphSessionId={0}" -f $GraphSessionId)

$step1Args = @("-Target", $Target, "-JsonOutputPath", $allToolsJson)
if (-not [string]::IsNullOrWhiteSpace($GraphSessionId)) {
    $step1Args += @("-GraphSessionId", $GraphSessionId)
}
elseif (-not [string]::IsNullOrWhiteSpace($UserHint)) {
    $step1Args += @("-UserHint", $UserHint)
}

$step1 = Run-ScriptStep -Name "Step 1: test-all-tools" -ScriptPath ".\deploy\test-all-tools.ps1" -StepArgs $step1Args

$resolvedGraphSessionId = $GraphSessionId
if (Test-Path $allToolsJson) {
    try {
        $all = Get-Content $allToolsJson -Raw | ConvertFrom-Json
        if (-not [string]::IsNullOrWhiteSpace([string]$all.graphSessionId)) {
            $resolvedGraphSessionId = [string]$all.graphSessionId
        }
        Write-Log ("Step1 totals: pass={0} fail={1} skipped={2} total={3}" -f $all.totals.pass, $all.totals.fail, $all.totals.skipped, $all.totals.total)
    }
    catch {
        Write-Log ("Could not parse all-tools JSON: {0}" -f $_.Exception.Message)
    }
}

if (-not [string]::IsNullOrWhiteSpace($resolvedGraphSessionId)) {
    $resolvedGraphSessionId = $resolvedGraphSessionId.Trim('"')
}
else {
    Write-Log "Step1 JSON output was not created."
}

if ([string]::IsNullOrWhiteSpace($resolvedGraphSessionId)) {
    Write-Log "No GraphSessionId available after Step 1; Step 2 may be skipped."
}
else {
    Write-Log ("Resolved GraphSessionId={0}" -f $resolvedGraphSessionId)
}

$step2 = 99
if (-not [string]::IsNullOrWhiteSpace($resolvedGraphSessionId)) {
    $step2Args = @(
        "-GraphSessionId", $resolvedGraphSessionId,
        "-BaseUrl", "https://ep-msgraphmcp-43613-c6dvbtfyfccmhzf8.a03.azurefd.net",
        "-Since", "2026-03-19",
        "-Until", "2026-04-09",
        "-Keywords", "Meeting Summarized",
        "-Matrix"
    )
    $step2 = Run-ScriptStep -Name "Step 2: test-mail-search matrix" -ScriptPath ".\deploy\test-mail-search.ps1" -StepArgs $step2Args
}
else {
    Write-Log "=== Step 2: test-mail-search matrix ==="
    Write-Log "Skipped (no GraphSessionId available)."
}

$step3Args = @("-Target", $Target, "-JsonOutputPath", $mailSummJson)
if (-not [string]::IsNullOrWhiteSpace($resolvedGraphSessionId)) {
    $step3Args += @("-GraphSessionId", $resolvedGraphSessionId)
}
elseif (-not [string]::IsNullOrWhiteSpace($UserHint)) {
    $step3Args += @("-UserHint", $UserHint)
}

$step3 = Run-ScriptStep -Name "Step 3: test-mail-summarize-context" -ScriptPath ".\deploy\test-mail-summarize-context.ps1" -StepArgs $step3Args

Write-Log "=== Final Summary ==="
Write-Log ("test-all-tools exit code: {0}" -f $step1)
Write-Log ("test-mail-search exit code: {0}" -f $step2)
Write-Log ("test-mail-summarize-context exit code: {0}" -f $step3)
Write-Log ("all-tools JSON: {0}" -f $allToolsJson)
Write-Log ("mail-summarize JSON: {0}" -f $mailSummJson)
Write-Log "Full validation completed"

Write-Output ("LOG_FILE={0}" -f $logFile)
Write-Output ("ALL_TOOLS_JSON={0}" -f $allToolsJson)
Write-Output ("MAIL_SUMMARIZE_JSON={0}" -f $mailSummJson)
Write-Output ("EXIT_CODES step1={0} step2={1} step3={2}" -f $step1, $step2, $step3)

if ($step1 -ne 0 -or $step2 -ne 0 -or $step3 -ne 0) {
    exit 1
}
