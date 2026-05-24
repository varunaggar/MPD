<#
.SYNOPSIS
    Incremental delta sync — keeps the Mailboxes table aligned
    with Exchange Online.

.DESCRIPTION
    Scheduled to run every 15 minutes via Windows Task Scheduler.

    Uses a WhenChangedUTC timestamp filter rather than Graph delta
    tokens (no mailbox delta API exists in Graph). The timestamp
    of the last successful run is stored in DeltaTokens and used
    as the filter window for the next run, with a configurable
    overlap (DeltaOverlapMinutes) to prevent missed changes.

    Steps performed:
      1. Reads the stored mailboxes_delta_timestamp from DeltaTokens
      2. Subtracts DeltaOverlapMinutes to create the since window
      3. Calls Get-EXOMailbox with WhenChangedUTC filter
      4. Also checks soft-deleted mailboxes in the same time window
      5. Applies changes: inserts new, updates changed, soft-deletes removed
      6. Saves the NEW timestamp ONLY after all processing succeeds
         (crash-safe: next run reprocesses rather than skips)

    Deletion handling:
      - Soft-deleted mailboxes appearing in the EXO SoftDeletedMailbox
        API within the time window are caught here and marked IsDeleted=1
      - Permanent deletions (after 30-day Exchange retention) are
        caught by the weekly Invoke-MailboxBaselineLoad reconciliation

.PARAMETER ConfigPath
    Path to config.xml. Defaults to config.xml in the same folder.

.NOTES
    Scheduled Task:
      Program  : pwsh.exe
      Arguments: -NonInteractive -File "C:\M365PermSync\Invoke-MailboxDeltaSync.ps1"
      Trigger  : Daily, repeat every 15 minutes, indefinitely
      Settings : ExecutionTimeLimit = 10 minutes
                 MultipleInstances = IgnoreNew
                 StartWhenAvailable = true
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.xml")
)

$ErrorActionPreference = "Stop"
$scriptName            = "Invoke-MailboxDeltaSync"
$sharedPath            = Join-Path $PSScriptRoot "shared"

# ──────────────────────────────────────────────────────────────
# Bootstrap
# ──────────────────────────────────────────────────────────────

Import-Module (Join-Path $sharedPath "ConfigHelpers.psm1")  -Force
Import-Module (Join-Path $sharedPath "LoggingHelpers.psm1") -Force
Import-Module (Join-Path $sharedPath "SqlHelpers.psm1")     -Force
Import-Module (Join-Path $sharedPath "ExoHelpers.psm1")     -Force

# ──────────────────────────────────────────────────────────────
# Configuration and logging
# ──────────────────────────────────────────────────────────────

$Config = Import-SyncConfig -Path $ConfigPath

Initialize-Logging -Config $Config -ProcessName $scriptName

Write-LogSection "Invoke-MailboxDeltaSync"
Write-LogInfo "Config   : $ConfigPath"
Write-LogInfo "EXO Org  : $($Config.ExchangeOnline.Organisation)"

# ──────────────────────────────────────────────────────────────
# Authentication
# ──────────────────────────────────────────────────────────────

Write-LogSection "Authentication"

try {
    Connect-SyncServicePrincipal -Config $Config
    Initialize-SqlContext -Config $Config
    Initialize-ExoContext -Config $Config
    Connect-ExoSession    -Config $Config
    Write-LogInfo "Azure and Exchange Online connected"
}
catch {
    Write-LogError "Authentication failed" -ErrorRecord $_
    Disconnect-ExoSession
    Close-Logging -Status "Failed"
    exit 1
}

# ──────────────────────────────────────────────────────────────
# Start run tracking
# ──────────────────────────────────────────────────────────────

$runId        = Start-SyncLogEntry -FunctionName $scriptName
$inserted     = 0
$updated      = 0
$softDeleted  = 0
$processed    = 0
$errors       = 0
$newTimestamp = $null
$overallSw    = [System.Diagnostics.Stopwatch]::StartNew()

Write-LogInfo "Run ID: $runId"

# ──────────────────────────────────────────────────────────────
# Phase 1 — Determine the sync window
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 1 — Determine sync window"

$storedTimestamp = Get-DeltaToken -TokenName "mailboxes_delta_timestamp"

if (-not $storedTimestamp) {
    $msg = "No active mailboxes_delta_timestamp found. Run Invoke-MailboxBaselineLoad first."
    Write-LogError $msg
    Complete-SyncLogEntry -RunId $runId -Status "Failed" -ErrorMessage $msg
    Disconnect-ExoSession
    Close-Logging -Status "Failed"
    exit 1
}

# Apply overlap window to avoid missing changes on boundary
$overlapMinutes = [int]($Config.Sync.DeltaOverlapMinutes ?? 30)
$sinceUtc       = [datetime]::Parse($storedTimestamp).ToUniversalTime().AddMinutes(-$overlapMinutes)

Write-LogInfo "Stored timestamp : $storedTimestamp"
Write-LogInfo "Overlap minutes  : $overlapMinutes"
Write-LogInfo "Query window from: $($sinceUtc.ToString('o')) UTC"

# ──────────────────────────────────────────────────────────────
# Phase 2 — Fetch changed mailboxes from EXO
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 2 — EXO delta fetch"

$changedMailboxes = $null
try {
    $changedMailboxes = Invoke-LoggedPhase -Name "ExoChangedFetch" -ScriptBlock {
        Get-ChangedExoMailboxes -Config $Config -SinceUtc $sinceUtc
    }
    Write-LogInfo "EXO returned $($changedMailboxes.Count) changed objects"
}
catch {
    Write-LogError "Phase 2 (EXO delta fetch) failed" -ErrorRecord $_
    Disconnect-ExoSession
    Complete-SyncLogEntry -RunId $runId -Status "Failed" `
        -ErrorMessage "EXO delta fetch failed: $($_.Exception.Message)"
    Close-Logging -Status "Failed"
    exit 1
}

if ($changedMailboxes.Count -eq 0) {
    Write-LogInfo "No changes in window — nothing to process"

    # Still advance the timestamp so next run's window stays current
    $newTimestamp = (Get-Date).ToUniversalTime().ToString('o')
    Save-DeltaToken -TokenName "mailboxes_delta_timestamp" -TokenValue $newTimestamp

    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET
    CompletedAt        = SYSUTCDATETIME(),
    Status             = 'Success',
    MailboxesProcessed = 0,
    TokenAdvancedTo    = @Token
WHERE RunId = @RunId
"@ -Parameters @{ '@Token' = $newTimestamp; '@RunId' = $runId } | Out-Null

    Write-LogSummary @{ "Status" = "Success (no changes)"; "Duration" = "$([Math]::Round($overallSw.Elapsed.TotalSeconds,1))s" }
    Disconnect-ExoSession
    Remove-OldLogFiles
    Close-Logging -Status "Success"
    exit 0
}

# ──────────────────────────────────────────────────────────────
# Phase 3 — Apply changes to Mailboxes table
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 3 — Apply changes"

$softDeleteSql = @"
UPDATE dbo.Mailboxes SET
    IsDeleted     = 1,
    DeletedAt     = SYSUTCDATETIME(),
    LastSyncedAt  = SYSUTCDATETIME(),
    SyncSource    = 'Delta',
    LastSyncRunId = @RunId
WHERE ExchangeGuid = @ExchangeGuid AND IsDeleted = 0;
"@

$upsertSql = @"
MERGE dbo.Mailboxes AS target
USING (SELECT @ExchangeGuid AS ExchangeGuid) AS source
ON target.ExchangeGuid = source.ExchangeGuid
WHEN MATCHED THEN UPDATE SET
    UserId                  = COALESCE(@UserId,                  UserId),
    PrimarySmtpAddress      = COALESCE(@PrimarySmtp,             PrimarySmtpAddress),
    UserPrincipalName       = COALESCE(@UPN,                     UserPrincipalName),
    DisplayName             = COALESCE(@DisplayName,             DisplayName),
    Alias                   = COALESCE(@Alias,                   Alias),
    RecipientTypeDetails    = COALESCE(@RecipientTypeDetails,    RecipientTypeDetails),
    MailboxType             = COALESCE(@MailboxType,             MailboxType),
    HiddenFromAddressLists  = COALESCE(@HiddenFromAddressLists,  HiddenFromAddressLists),
    LitigationHoldEnabled   = COALESCE(@LitigationHoldEnabled,   LitigationHoldEnabled),
    ArchiveStatus           = COALESCE(@ArchiveStatus,           ArchiveStatus),
    ForwardingAddress       = COALESCE(@ForwardingAddress,       ForwardingAddress),
    ForwardingSmtpAddress   = COALESCE(@ForwardingSmtpAddress,   ForwardingSmtpAddress),
    GrantSendOnBehalfTo     = COALESCE(@GrantSendOnBehalfTo,     GrantSendOnBehalfTo),
    IsDirSynced             = COALESCE(@IsDirSynced,             IsDirSynced),
    WhenMailboxCreated      = COALESCE(@WhenMailboxCreated,      WhenMailboxCreated),
    WhenChangedUTC          = COALESCE(@WhenChangedUTC,          WhenChangedUTC),
    LastSyncedAt            = SYSUTCDATETIME(),
    LastModifiedAt          = SYSUTCDATETIME(),
    IsDeleted               = 0,
    DeletedAt               = NULL,
    SyncSource              = 'Delta',
    LastSyncRunId           = @RunId
WHEN NOT MATCHED THEN INSERT (
    ExchangeGuid, UserId, PrimarySmtpAddress, UserPrincipalName,
    DisplayName, Alias, RecipientTypeDetails, MailboxType,
    HiddenFromAddressLists, LitigationHoldEnabled, ArchiveStatus,
    ForwardingAddress, ForwardingSmtpAddress, GrantSendOnBehalfTo,
    IsDirSynced, WhenMailboxCreated, WhenChangedUTC,
    SyncSource, LastSyncRunId
) VALUES (
    @ExchangeGuid, @UserId, @PrimarySmtp, @UPN,
    @DisplayName, @Alias, @RecipientTypeDetails, @MailboxType,
    @HiddenFromAddressLists, @LitigationHoldEnabled, @ArchiveStatus,
    @ForwardingAddress, @ForwardingSmtpAddress, @GrantSendOnBehalfTo,
    @IsDirSynced, @WhenMailboxCreated, @WhenChangedUTC,
    'Delta', @RunId
);
"@

$existsSql = "SELECT COUNT(1) FROM dbo.Mailboxes WHERE ExchangeGuid = @ExchangeGuid"
$phaseSw   = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($m in $changedMailboxes) {
    try {
        # ── Soft-deleted mailbox ──────────────────────────────
        if ($m.IsDeleted) {
            $rows = Invoke-SqlNonQuery -Query $softDeleteSql -Parameters @{
                '@ExchangeGuid' = $m.ExchangeGuid
                '@RunId'        = $runId
            }
            if ($rows -gt 0) {
                $softDeleted++
                Write-LogInfo "Soft-deleted: $($m.PrimarySmtpAddress) ($($m.ExchangeGuid))"
            }
            $processed++
            continue
        }

        # ── New or changed mailbox ────────────────────────────
        $exists = Invoke-SqlScalar -Query $existsSql `
            -Parameters @{ '@ExchangeGuid' = $m.ExchangeGuid }

        Invoke-SqlNonQuery -Query $upsertSql -Parameters @{
            '@ExchangeGuid'          = $m.ExchangeGuid
            '@UserId'                = $m.UserId
            '@PrimarySmtp'           = $m.PrimarySmtpAddress
            '@UPN'                   = $m.UserPrincipalName
            '@DisplayName'           = $m.DisplayName
            '@Alias'                 = $m.Alias
            '@RecipientTypeDetails'  = $m.RecipientTypeDetails
            '@MailboxType'           = $m.MailboxType
            '@HiddenFromAddressLists'= $m.HiddenFromAddressLists
            '@LitigationHoldEnabled' = $m.LitigationHoldEnabled
            '@ArchiveStatus'         = $m.ArchiveStatus
            '@ForwardingAddress'     = $m.ForwardingAddress
            '@ForwardingSmtpAddress' = $m.ForwardingSmtpAddress
            '@GrantSendOnBehalfTo'   = $m.GrantSendOnBehalfTo
            '@IsDirSynced'           = $m.IsDirSynced
            '@WhenMailboxCreated'    = $m.WhenMailboxCreated
            '@WhenChangedUTC'        = $m.WhenChangedUTC
            '@RunId'                 = $runId
        } | Out-Null

        if ([int]$exists -eq 0) { $inserted++ } else { $updated++ }
        $processed++
    }
    catch {
        $errors++
        Write-LogWarning "Failed to process $($m.PrimarySmtpAddress): $($_.Exception.Message)"
    }
}

$phaseSw.Stop()
Write-LogInfo "Changes applied in $([Math]::Round($phaseSw.Elapsed.TotalSeconds,1))s"

# ──────────────────────────────────────────────────────────────
# Phase 4 — Advance timestamp
# Saved ONLY after processing completes successfully.
# If the script crashes in Phase 3, the next run reprocesses
# the same window rather than skipping any changes.
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 4 — Advance timestamp"

$newTimestamp = (Get-Date).ToUniversalTime().ToString('o')
Save-DeltaToken -TokenName "mailboxes_delta_timestamp" -TokenValue $newTimestamp
Write-LogInfo "Timestamp advanced to: $newTimestamp"

# ──────────────────────────────────────────────────────────────
# Finalise
# ──────────────────────────────────────────────────────────────

$overallSw.Stop()
$finalStatus = if ($errors -gt 0) { "PartialFailure" } else { "Success" }

Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET
    CompletedAt          = SYSUTCDATETIME(),
    Status               = @Status,
    MailboxesInserted    = @Inserted,
    MailboxesUpdated     = @Updated,
    MailboxesSoftDeleted = @Deleted,
    MailboxesProcessed   = @Processed,
    ErrorCount           = @Errors,
    TokenAdvancedTo      = @Token
WHERE RunId = @RunId
"@ -Parameters @{
    '@RunId'     = $runId
    '@Status'    = $finalStatus
    '@Inserted'  = $inserted
    '@Updated'   = $updated
    '@Deleted'   = $softDeleted
    '@Processed' = $processed
    '@Errors'    = $errors
    '@Token'     = $newTimestamp
} | Out-Null

Write-LogSummary @{
    "Status"      = $finalStatus
    "Inserted"    = $inserted
    "Updated"     = $updated
    "Soft-deleted"= $softDeleted
    "Errors"      = $errors
    "Processed"   = $processed
    "Duration"    = "$([Math]::Round($overallSw.Elapsed.TotalSeconds,1))s"
    "Run ID"      = $runId
}

Disconnect-ExoSession
Remove-OldLogFiles
Close-Logging -Status $finalStatus

if ($finalStatus -eq "Failed") { exit 1 }
