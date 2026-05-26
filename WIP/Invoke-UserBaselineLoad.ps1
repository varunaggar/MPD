<#
.SYNOPSIS
    Full baseline load — pages all users from Graph into the Users table.

.DESCRIPTION
    Run once on initial deployment, and again whenever the delta token
    expires and needs to be reset.

    Steps performed:
      1. Pages through /users fetching all user objects
      2. MERGEs every user into the Users table (idempotent)
      3. Initiates a /users/delta call to capture a starting delta token
      4. Stores the token in DeltaTokens so Invoke-UserDeltaSync can run

    Re-runnable: yes. Re-running resets the Users table to current Graph
    state and replaces the delta token.

.PARAMETER ConfigPath
    Path to config.xml. Defaults to config.xml in the same folder as
    this script.

.NOTES
    Typical duration: 10–30 minutes for a 40,000 user tenant.
    Schedule as a one-off task or as a weekly reconciliation run.

    Scheduled Task example:
      Program : pwsh.exe
      Arguments: -NonInteractive -File "C:\M365PermSync\Invoke-UserBaselineLoad.ps1"
      Run As  : The service account that has access to the certificate.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.xml")
)

$ErrorActionPreference = "Stop"
$scriptName            = "Invoke-UserBaselineLoad"
# Paths to shared or helper modules
$sharedPath            = Join-Path $PSScriptRoot "shared"
# Path to Powershell modules
$modulesRoot           = Join-Path $PSScriptRoot "Modules"

# ──────────────────────────────────────────────────────────────
# Bootstrap — import shared / helper modules
# ──────────────────────────────────────────────────────────────

Import-Module (Join-Path $sharedPath "ConfigHelpers.psm1")  -Force
Import-Module (Join-Path $sharedPath "LoggingHelpers.psm1") -Force
Import-Module (Join-Path $sharedPath "GraphHelpers.psm1")   -Force
Import-Module (Join-Path $sharedPath "SqlHelpers.psm1")     -Force
Import-Module (Join-Path $sharedPath "DependencyHelpers.psm1") -Force

# ──────────────────────────────────────────────────────────────
# Load configuration
# ──────────────────────────────────────────────────────────────

try {
    $Config = Import-SyncConfig -Path $ConfigPath

    # ──────────────────────────────────────────────────────────────
    # Initialise logging
    # ──────────────────────────────────────────────────────────────
    Initialize-Logging -Config $Config -ProcessName $scriptName
}
catch {
    Write-Host "FATAL BOOTSTRAP ERROR in $scriptName" -ForegroundColor Red
    Write-Host "Location: $($_.InvocationInfo.ScriptName) Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Yellow
    Write-Host "Message : $($_.Exception.Message)" -ForegroundColor White
    exit 1
}

Write-LogSection "Invoke-UserBaselineLoad"
Write-LogInfo "Script   : $scriptName"
Write-LogInfo "Config   : $ConfigPath"
Write-LogInfo "Server   : $($Config.Database.Server)"
Write-LogInfo "Database : $($Config.Database.Name)"
Write-LogInfo "AppId    : $($Config.Authentication.AppID)"
Write-LogInfo "Tenant   : $($Config.Authentication.TenantID)"

# ──────────────────────────────────────────────────────────────
# Load PS modules as per the config.xml file
# ──────────────────────────────────────────────────────────────
Write-LogSection "Dependency Validation"
Initialize-ModuleDependencies -Config $Config

# ──────────────────────────────────────────────────────────────
# Authenticate to Azure AD using App ID 
# ──────────────────────────────────────────────────────────────
Write-LogSection "Authentication"

try {
    Connect-SyncServicePrincipal -Config $Config
    Write-LogInfo "Service principal connected successfully"
}
catch {
    Write-LogError "Authentication failed — cannot continue" -ErrorRecord $_
    Close-Logging -Status "Failed"
    exit 1
}

# ──────────────────────────────────────────────────────────────
# Initialise Graph and SQL contexts
# ──────────────────────────────────────────────────────────────
Initialize-GraphContext -Config $Config
Initialize-SqlContext   -Config $Config

# ────────────────────────────────────────────────────────────────────────────────────────────────────────
# Start run tracking - Create a entry in SyncLog table to keep a track of the execution and result
# The function Start-SyncLogEntry acquires a AAD toekn  for SQL auth and creates a connection to SQL database
# ────────────────────────────────────────────────────────────────────────────────────────────────────────

$runId = Start-SyncLogEntry -FunctionName $scriptName
Write-LogInfo "Run ID: $runId"

$inserted      = 0
$processed     = 0
$errors        = 0
$capturedToken = $null
$overallStart  = [System.Diagnostics.Stopwatch]::StartNew()

# ──────────────────────────────────────────────────────────────
# Phase 1 — Page through AAD users using Graph /users endpoint
# The select statement has attributes which are queried from Graph
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 1 — Graph paged fetch"

try {
    $allUsers = Invoke-LoggedPhase -Name "GraphPagedFetch" -ScriptBlock {

        $pageSize = [int]($Config.Graph.PageSize ?? 999)
        $select   = "id,userPrincipalName,displayName,mail,accountEnabled," +
                    "userType,onPremisesSyncEnabled,department,jobTitle,createdDateTime"
        $url      = "$($Config.Graph.BaseUrl)/users?`$select=$select&`$top=$pageSize"

        Invoke-GraphPagedRequest -Uri $url
    }
    Write-LogInfo "Graph returned $($allUsers.Count) users"
}
catch {
    Write-LogError "Phase 1 (Graph fetch) failed — aborting" -ErrorRecord $_
    Complete-SyncLogEntry -RunId $runId -Status "Failed" `
        -ErrorMessage "Phase 1 failed: $($_.Exception.Message)"
    Close-Logging -Status "Failed"
    exit 1
}

# ──────────────────────────────────────────────────────────────
# Phase 2 — MERGE all users into the Users table
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 2 — SQL upsert"

$mergeSql = @"
MERGE Users AS target
USING (SELECT @UserId AS UserId) AS source
ON target.UserId = source.UserId
WHEN MATCHED THEN UPDATE SET
    UserPrincipalName     = @UPN,
    DisplayName           = @DisplayName,
    Mail                  = @Mail,
    AccountEnabled        = @AccountEnabled,
    UserType              = @UserType,
    OnPremisesSyncEnabled = @OnPremisesSyncEnabled,
    Department            = @Department,
    JobTitle              = @JobTitle,
    EntraCreatedDateTime  = @CreatedDateTime,
    LastSyncedAt          = SYSUTCDATETIME(),
    LastModifiedAt        = SYSUTCDATETIME(),
    SyncSource            = 'Baseline',
    LastSyncRunId         = @RunId,
    IsDeleted             = 0,
    DeletedAt             = NULL
WHEN NOT MATCHED THEN INSERT (
    UserId, UserPrincipalName, DisplayName, Mail, AccountEnabled,
    UserType, OnPremisesSyncEnabled, Department, JobTitle,
    EntraCreatedDateTime, SyncSource, LastSyncRunId
) VALUES (
    @UserId, @UPN, @DisplayName, @Mail, @AccountEnabled,
    @UserType, @OnPremisesSyncEnabled, @Department, @JobTitle,
    @CreatedDateTime, 'Baseline', @RunId
);
"@

$progressInterval = 5000
$phaseStart = [System.Diagnostics.Stopwatch]::StartNew()
$batchSize = 500
$batchStatements = @()

foreach ($u in $allUsers) {
    $batchStatements += @{
        Query      = $mergeSql
        Parameters = @{
            '@UserId'                = [guid]$u.id
            '@UPN'                   = $u.userPrincipalName
            '@DisplayName'           = $u.displayName
            '@Mail'                  = $u.mail
            '@AccountEnabled'        = if ($null -eq $u.accountEnabled)        { $null } else { [bool]$u.accountEnabled }
            '@UserType'              = $u.userType
            '@OnPremisesSyncEnabled' = if ($null -eq $u.onPremisesSyncEnabled) { $null } else { [bool]$u.onPremisesSyncEnabled }
            '@Department'            = $u.department
            '@JobTitle'              = $u.jobTitle
            '@CreatedDateTime'       = if ($u.createdDateTime) { [datetime]$u.createdDateTime } else { $null }
            '@RunId'                 = $runId
        }
    }

    if ($batchStatements.Count -ge $batchSize) {
        try {
            $inserted += Invoke-SqlBatch -Statements $batchStatements -UseTransaction
        }
        catch {
            $errors += $batchStatements.Count
            Write-LogWarning "Batch execution failed: $($_.Exception.Message)"
        }
        $processed += $batchStatements.Count
        $batchStatements = @()

        Write-LogInfo "Progress: $processed / $($allUsers.Count) users processed ($errors errors)"
    }
}

if ($batchStatements.Count -gt 0) {
    try {
        $inserted += Invoke-SqlBatch -Statements $batchStatements -UseTransaction
    }
    catch {
        $errors += $batchStatements.Count
        Write-LogWarning "Final batch execution failed: $($_.Exception.Message)"
    }
    $processed += $batchStatements.Count
}

$phaseStart.Stop()
Write-LogInfo "SQL upsert complete in $([Math]::Round($phaseStart.Elapsed.TotalSeconds,1))s — $inserted upserted, $errors errors"

# ──────────────────────────────────────────────────────────────
# Phase 3 — Capture starting delta token
# We use $select=id only — we don't need the payload, just the token.
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 3 — Capture delta token"

try {
    $deltaResult = Invoke-LoggedPhase -Name "CaptureDeltaToken" -ScriptBlock {
        $url = "$($Config.Graph.BaseUrl)/users/delta"
        Invoke-GraphDeltaQuery -Uri $url
    }

    if ($deltaResult.TokenExpired -or -not $deltaResult.DeltaToken) {
        throw "Delta initialisation did not return a valid token"
    }

    Save-DeltaToken -TokenName "users_delta" -TokenValue $deltaResult.DeltaToken
    $capturedToken = $deltaResult.DeltaToken
    Write-LogInfo "Starting delta token saved to DeltaTokens table"
}
catch {
    Write-LogError "Phase 3 (delta token capture) failed" -ErrorRecord $_
    # This is non-fatal for the data — users are loaded — but delta sync won't work.
    # Mark as PartialFailure rather than Failed.
    $errors++
}

# ──────────────────────────────────────────────────────────────
# Finalise
# ──────────────────────────────────────────────────────────────

$overallStart.Stop()
$finalStatus = if ($errors -gt 0 -and -not $capturedToken) 
                { "Failed" }
               elseif ($errors -gt 0)                      
               { "PartialFailure" }
               else                                         
               { "Success" }

Complete-SyncLogEntry `
    -RunId          $runId `
    -Status         $finalStatus `
    -UsersInserted  $inserted `
    -UsersProcessed $processed `
    -ErrorCount     $errors `
    -TokenAdvancedTo $capturedToken

Write-LogSection "Run Summary"
Write-LogSummary @{
    "Status"          = $finalStatus
    "Users processed" = $processed
    "Users upserted"  = $inserted
    "Errors"          = $errors
    "Token captured"  = ($null -ne $capturedToken)
    "Duration"        = "$([Math]::Round($overallStart.Elapsed.TotalMinutes, 2)) minutes"
    "Run ID"          = $runId
}

Remove-OldLogFiles
Close-Logging -Status $finalStatus

if ($finalStatus -eq "Failed") { exit 1 }
