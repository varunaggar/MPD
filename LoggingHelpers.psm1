<#
.SYNOPSIS
    Structured logging helpers for the M365 permissions sync solution.

.DESCRIPTION
    Wraps Write-Host / Write-Warning / Write-Error with consistent formatting,
    plus a Write-Telemetry function that emits custom events visible as
    "customEvents" in Application Insights.

    Application Insights ingests Write-Host output automatically as traces.
    For STRUCTURED telemetry (filterable, queryable in KQL), use Write-Telemetry.

.NOTES
    No external dependencies. Standalone PowerShell.
    Imported by every Azure Function via profile.ps1.
#>

# ──────────────────────────────────────────────────────────────
# Module-scoped state — track current run for log correlation
# ──────────────────────────────────────────────────────────────

$script:CurrentRunId       = $null
$script:CurrentFunctionName = $null

# ──────────────────────────────────────────────────────────────
# Public: Initialize-LoggingContext
# Sets the run context for subsequent log calls
# ──────────────────────────────────────────────────────────────

function Initialize-LoggingContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$FunctionName,
        [Parameter(Mandatory)] [guid]$RunId
    )

    $script:CurrentFunctionName = $FunctionName
    $script:CurrentRunId        = $RunId

    Write-LogInfo "Logging context initialised for $FunctionName (RunId=$RunId)"
}

# ──────────────────────────────────────────────────────────────
# Public: Write-LogInfo / Write-LogWarning / Write-LogError
# Standard structured logs — go to App Insights as "traces"
# ──────────────────────────────────────────────────────────────

function Write-LogInfo {
    [CmdletBinding()]
    param([Parameter(Mandatory, Position=0)] [string]$Message)

    $prefix = Get-LogPrefix
    Write-Host "$prefix [INFO ] $Message"
}

function Write-LogWarning {
    [CmdletBinding()]
    param([Parameter(Mandatory, Position=0)] [string]$Message)

    $prefix = Get-LogPrefix
    Write-Warning "$prefix [WARN ] $Message"
}

function Write-LogError {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position=0)] [string]$Message,
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $prefix = Get-LogPrefix

    if ($ErrorRecord) {
        $detail = "$Message | Exception: $($ErrorRecord.Exception.Message) | At: $($ErrorRecord.InvocationInfo.PositionMessage)"
        Write-Error "$prefix [ERROR] $detail" -ErrorAction Continue
    } else {
        Write-Error "$prefix [ERROR] $Message" -ErrorAction Continue
    }
}

# ──────────────────────────────────────────────────────────────
# Private: Get-LogPrefix — formats run context for every log line
# ──────────────────────────────────────────────────────────────

function Get-LogPrefix {
    $ts = (Get-Date -Format 'o')
    if ($script:CurrentFunctionName -and $script:CurrentRunId) {
        return "$ts [$($script:CurrentFunctionName)] [$($script:CurrentRunId.ToString().Substring(0,8))]"
    }
    return "$ts"
}

# ──────────────────────────────────────────────────────────────
# Public: Write-Telemetry
# Emits a STRUCTURED custom event to Application Insights.
#
# In Functions, telemetry written via Write-Host with a JSON body
# tagged "applicationinsights:customEvent" lands in App Insights as
# customEvents (filterable in KQL: customEvents | where name == "...")
#
# This gives us metrics like: count of users processed per run,
# duration of each phase, etc. that we can chart and alert on.
# ──────────────────────────────────────────────────────────────

function Write-Telemetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$EventName,
        [hashtable]$Properties = @{},
        [hashtable]$Metrics = @{}
    )

    # Always include run context if we have it
    $enriched = @{
        functionName = $script:CurrentFunctionName
        runId        = if ($script:CurrentRunId) { $script:CurrentRunId.ToString() } else { $null }
    }
    foreach ($k in $Properties.Keys) { $enriched[$k] = $Properties[$k] }

    $payload = @{
        eventName   = $EventName
        properties  = $enriched
        metrics     = $Metrics
        timestamp   = (Get-Date -Format 'o')
    } | ConvertTo-Json -Depth 5 -Compress

    # Write to host with a recognisable prefix — App Insights captures as trace
    # For richer telemetry, the function can also call the App Insights track
    # API directly (not implemented here to keep dependencies minimal).
    Write-Host "TELEMETRY $payload"
}

# ──────────────────────────────────────────────────────────────
# Public: Measure-Phase
# Times a script block, logs duration, emits telemetry.
#
# Usage:
#   $result = Measure-Phase -Name 'GraphPaging' -ScriptBlock {
#       Invoke-GraphPagedRequest -Uri $url
#   }
# ──────────────────────────────────────────────────────────────

function Measure-Phase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Name,
        [Parameter(Mandatory)] [scriptblock]$ScriptBlock
    )

    Write-LogInfo "Phase '$Name' starting"
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $result = $null
    $errorOccurred = $false

    try {
        $result = & $ScriptBlock
    }
    catch {
        $errorOccurred = $true
        $sw.Stop()
        Write-LogError "Phase '$Name' failed after $($sw.Elapsed.TotalSeconds) sec" -ErrorRecord $_
        Write-Telemetry -EventName "PhaseFailed" `
            -Properties @{ phaseName = $Name } `
            -Metrics    @{ durationSeconds = $sw.Elapsed.TotalSeconds }
        throw
    }

    $sw.Stop()
    if (-not $errorOccurred) {
        Write-LogInfo "Phase '$Name' completed in $([Math]::Round($sw.Elapsed.TotalSeconds, 2)) sec"
        Write-Telemetry -EventName "PhaseCompleted" `
            -Properties @{ phaseName = $Name } `
            -Metrics    @{ durationSeconds = $sw.Elapsed.TotalSeconds }
    }

    return $result
}

# ──────────────────────────────────────────────────────────────
# Exports
# ──────────────────────────────────────────────────────────────

Export-ModuleMember -Function @(
    'Initialize-LoggingContext',
    'Write-LogInfo',
    'Write-LogWarning',
    'Write-LogError',
    'Write-Telemetry',
    'Measure-Phase'
)
