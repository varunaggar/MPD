<#
.SYNOPSIS
    Incremental delta sync — keeps the Users table aligned with Entra ID.

.DESCRIPTION
    Scheduled to run every 15 minutes via Windows Task Scheduler.

    Steps performed:
      1. Reads the active users_delta token from DeltaTokens
      2. Calls /users/delta?$deltatoken=... to get only changes since last run
      3. Applies changes:
           @removed marker  → soft-delete (IsDeleted = 1)
           New object ID    → INSERT
           Existing ID      → UPDATE (COALESCE prevents overwriting with NULL
                              when Graph returns partial properties)
      4. Saves the new delta token ONLY after all processing succeeds
         (crash-safe: next run reprocesses rather than skips)
      5. On HTTP 410 Gone (expired token): deactivates token and exits cleanly.
         Operator must re-run Invoke-UserBaselineLoad to recover.

.PARAMETER ConfigPath
    Path to config.xml. Defaults to config.xml in the same folder as
    this script.

.NOTES
    Scheduled Task example:
      Trigger : Daily, repeat every 15 minutes, indefinitely
      Program : pwsh.exe
      Arguments: -NonInteractive -File "C:\M365PermSync\Invoke-UserDeltaSync.ps1"
      Run As  : The service account that has access to the certificate.
      Settings: Do not start a new instance if already running.
                Run task as soon as possible after a scheduled start is missed.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.xml")
)

$ErrorActionPreference = "Stop"
$scriptName            = "Invoke-UserDeltaSync"
$sharedPath            = Join-Path $PSScriptRoot "shared"

# ──────────────────────────────────────────────────────────────
# Bootstrap — import shared modules
# ──────────────────────────────────────────────────────────────

Import-Module (Join-Path $sharedPath "ConfigHelpers.psm1")  -Force
Import-Module (Join-Path $sharedPath "LoggingHelpers.psm1") -Force
Import-Module (Join-Path $sharedPath "GraphHelpers.psm1")   -Force
Import-Module (Join-Path $sharedPath "SqlHelpers.psm1")     -Force

# ──────────────────────────────────────────────────────────────
# Load configuration
# ──────────────────────────────────────────────────────────────

$Config = Import-SyncConfig -Path $ConfigPath

# ──────────────────────────────────────────────────────────────
# Initialise logging
# Log file: Logs\Invoke-UserDeltaSync_yyyy-MM-dd_HH-mm-ss.log
# ──────────────────────────────────────────────────────────────

Initialize-Logging -Config $Config -ProcessName $scriptName

Write-LogSection "Invoke-UserDeltaSync"
Write-LogInfo "Config   : $ConfigPath"
Write-LogInfo "Server   : $($Config.Database.Server)"
Write-LogInfo "Database : $($Config.Database.Name)"

# ──────────────────────────────────────────────────────────────
# Authenticate and initialise helpers
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

Initialize-GraphContext -Config $Config
Initialize-SqlContext   -Config $Config

# ──────────────────────────────────────────────────────────────
# Start run tracking
# ──────────────────────────────────────────────────────────────

$runId        = Start-SyncLogEntry -FunctionName $scriptName
$inserted     = 0
$updated      = 0
$softDeleted  = 0
$processed    = 0
$errors       = 0
$newToken     = $null
$overallStart = [System.Diagnostics.Stopwatch]::StartNew()

Write-LogInfo "Run ID: $runId"

# ──────────────────────────────────────────────────────────────
# Phase 1 — Fetch stored delta token
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 1 — Fetch delta token"

$storedToken = Get-DeltaToken -TokenName "users_delta"

if (-not $storedToken) {
    $msg = "No active users_delta token found. Run Invoke-UserBaselineLoad first."
    Write-LogError $msg
    Complete-SyncLogEntry -RunId $runId -Status "Failed" -ErrorMessage $msg
    Close-Logging -Status "Failed"
    exit 1
}

Write-LogInfo "Active token found (length: $($storedToken.Length))"

# ──────────────────────────────────────────────────────────────
# Phase 2 — Call /users/delta with the stored token
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 2 — Graph delta fetch"

$select = "id,userPrincipalName,displayName,mail,accountEnabled," +
          "userType,onPremisesSyncEnabled,department,jobTitle,createdDateTime"
$url    = "$($Config.Graph.BaseUrl)/users/delta?`$select=$select&`$deltatoken=$storedToken"

try {
    $deltaResult = Invoke-LoggedPhase -Name "GraphDeltaFetch" -ScriptBlock {
        Invoke-GraphDeltaQuery -Uri $url
    }
}
catch {
    Write-LogError "Graph delta fetch failed" -ErrorRecord $_
    Complete-SyncLogEntry -RunId $runId -Status "Failed" `
        -ErrorMessage "Graph delta fetch failed: $($_.Exception.Message)"
    Close-Logging -Status "Failed"
    exit 1
}

# Token expired — deactivate and exit cleanly
if ($deltaResult.TokenExpired) {
    Write-LogWarning "Delta token expired (HTTP 410 Gone)"
    Write-LogWarning "Action required: re-run Invoke-UserBaselineLoad to reset"

    Disable-DeltaToken -TokenName "users_delta" `
        -Reason "HTTP 410 Gone — expired. Re-run Invoke-UserBaselineLoad."

    Complete-SyncLogEntry -RunId $runId -Status "Failed" `
        -ErrorMessage "Delta token expired (410 Gone). Baseline reload required."
    Close-Logging -Status "Failed — token expired"
    exit 1
}

Write-LogInfo "Delta returned $($deltaResult.Objects.Count) changed objects"

if ($deltaResult.Objects.Count -eq 0) {
    Write-LogInfo "No changes since last run — nothing to process"
    # Still save the new token and mark success
    if ($deltaResult.DeltaToken) {
        Save-DeltaToken -TokenName "users_delta" -TokenValue $deltaResult.DeltaToken
        $newToken = $deltaResult.DeltaToken
    }
    Complete-SyncLogEntry -RunId $runId -Status "Success" -TokenAdvancedTo $newToken
    Write-LogSummary @{
        "Status"         = "Success (no changes)"
        "Objects changed"= 0
        "Duration"       = "$([Math]::Round($overallStart.Elapsed.TotalSeconds, 1))s"
    }
    Remove-OldLogFiles
    Close-Logging -Status "Success"
    exit 0
}

# ──────────────────────────────────────────────────────────────
# Phase 3 — Apply changes to Users table
# Process: deletions first, then inserts/updates
# COALESCE on UPDATE prevents overwriting with NULL when Graph
# returns only the changed properties in a delta response.
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 3 — Apply changes"

$softDeleteSql = @"
UPDATE Users SET
    IsDeleted      = 1,
    DeletedAt      = SYSUTCDATETIME(),
    LastSyncedAt   = SYSUTCDATETIME(),
    LastModifiedAt = SYSUTCDATETIME(),
    SyncSource     = 'Delta',
    LastSyncRunId  = @RunId
WHERE UserId = @UserId AND IsDeleted = 0;
"@

$upsertSql = @"
MERGE Users AS target
USING (SELECT @UserId AS UserId) AS source
ON target.UserId = source.UserId
WHEN MATCHED THEN UPDATE SET
    UserPrincipalName     = COALESCE(@UPN,                  UserPrincipalName),
    DisplayName           = COALESCE(@DisplayName,           DisplayName),
    Mail                  = COALESCE(@Mail,                  Mail),
    AccountEnabled        = COALESCE(@AccountEnabled,        AccountEnabled),
    UserType              = COALESCE(@UserType,              UserType),
    OnPremisesSyncEnabled = COALESCE(@OnPremisesSyncEnabled, OnPremisesSyncEnabled),
    Department            = COALESCE(@Department,            Department),
    JobTitle              = COALESCE(@JobTitle,              JobTitle),
    EntraCreatedDateTime  = COALESCE(@CreatedDateTime,       EntraCreatedDateTime),
    LastSyncedAt          = SYSUTCDATETIME(),
    LastModifiedAt        = SYSUTCDATETIME(),
    SyncSource            = 'Delta',
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
    @CreatedDateTime, 'Delta', @RunId
);
"@

$existsSql = "SELECT COUNT(1) FROM Users WHERE UserId = @UserId"

$phaseStart = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($u in $deltaResult.Objects) {
    try {
        # ── Deleted user ─────────────────────────────────────
        if ($u.'@removed') {
            $rows = Invoke-SqlNonQuery -Query $softDeleteSql -Parameters @{
                '@UserId' = [guid]$u.id
                '@RunId'  = $runId
            }
            if ($rows -gt 0) {
                $softDeleted++
                Write-LogInfo "Soft-deleted: $($u.id)"
            }
            $processed++
            continue
        }

        # ── New or changed user ───────────────────────────────
        $exists = Invoke-SqlScalar -Query $existsSql `
            -Parameters @{ '@UserId' = [guid]$u.id }

        Invoke-SqlNonQuery -Query $upsertSql -Parameters @{
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
        } | Out-Null

        if ([int]$exists -eq 0) { $inserted++ } else { $updated++ }
        $processed++
    }
    catch {
        $errors++
        Write-LogWarning "Failed to process user $($u.id): $($_.Exception.Message)"
    }
}

$phaseStart.Stop()
Write-LogInfo "Changes applied in $([Math]::Round($phaseStart.Elapsed.TotalSeconds,1))s"

# ──────────────────────────────────────────────────────────────
# Phase 4 — Save new delta token
# IMPORTANT: token is saved only AFTER all processing completes.
# If the script crashes before this point, the next run will
# reprocess the same window rather than skip any changes.
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 4 — Advance delta token"

if ($deltaResult.DeltaToken) {
    Save-DeltaToken -TokenName "users_delta" -TokenValue $deltaResult.DeltaToken
    $newToken = $deltaResult.DeltaToken
    Write-LogInfo "Delta token advanced successfully"
}
else {
    Write-LogWarning "No new delta token returned — token NOT advanced"
}

# ──────────────────────────────────────────────────────────────
# Finalise
# ──────────────────────────────────────────────────────────────

$overallStart.Stop()
$finalStatus = if ($errors -gt 0) { "PartialFailure" } else { "Success" }

Complete-SyncLogEntry `
    -RunId           $runId `
    -Status          $finalStatus `
    -UsersInserted   $inserted `
    -UsersUpdated    $updated `
    -UsersSoftDeleted $softDeleted `
    -UsersProcessed  $processed `
    -ErrorCount      $errors `
    -TokenAdvancedTo $newToken

Write-LogSummary @{
    "Status"          = $finalStatus
    "Inserted"        = $inserted
    "Updated"         = $updated
    "Soft-deleted"    = $softDeleted
    "Errors"          = $errors
    "Total processed" = $processed
    "Duration"        = "$([Math]::Round($overallStart.Elapsed.TotalSeconds, 1))s"
    "Run ID"          = $runId
}

Remove-OldLogFiles
Close-Logging -Status $finalStatus

if ($finalStatus -eq "Failed") { exit 1 }
