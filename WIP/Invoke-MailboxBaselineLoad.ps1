<#
.SYNOPSIS
    Full baseline load — pages all mailboxes from Exchange Online
    into the Mailboxes table.

.DESCRIPTION
    Run once on initial deployment, and re-run weekly as a
    reconciliation pass that catches permanent deletions and
    any drift not caught by the delta sync.

    Steps performed:
      1. Connects to Exchange Online using certificate auth
      2. Fetches all mailboxes of the configured types via Get-EXOMailbox
      3. MERGEs every mailbox into the Mailboxes table (idempotent)
      4. Reconciles deletions: soft-deletes any DB row whose
         ExchangeGuid is no longer present in EXO
      5. Saves the current UTC timestamp to DeltaTokens
         so Invoke-MailboxDeltaSync can run from this point forward
      6. Disconnects from Exchange Online

    Mailbox types collected (configured in config.xml):
      UserMailbox | SharedMailbox | RoomMailbox | EquipmentMailbox

.PARAMETER ConfigPath
    Path to config.xml. Defaults to config.xml in the same folder.

.NOTES
    Typical duration: 15–45 minutes for a 40,000 mailbox tenant.
    Schedule as a weekly task (e.g. Sunday 03:00) after the
    user baseline load has completed.

    Scheduled Task:
      Program  : pwsh.exe
      Arguments: -NonInteractive -File "C:\M365PermSync\Invoke-MailboxBaselineLoad.ps1"
      Trigger  : Weekly, Sunday 03:00
      Settings : ExecutionTimeLimit = 2 hours
                 MultipleInstances = IgnoreNew
                 StartWhenAvailable = true
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.xml")
)

$ErrorActionPreference = "Stop"
$scriptName            = "Invoke-MailboxBaselineLoad"
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

Write-LogSection "Invoke-MailboxBaselineLoad"
Write-LogInfo "Config       : $ConfigPath"
Write-LogInfo "Server       : $($Config.Database.Server)"
Write-LogInfo "Database     : $($Config.Database.Name)"
Write-LogInfo "EXO Org      : $($Config.ExchangeOnline.Organisation)"
Write-LogInfo "Mailbox types: $($Config.ExchangeOnline.MailboxTypes)"

# ──────────────────────────────────────────────────────────────
# Authentication
# ──────────────────────────────────────────────────────────────

Write-LogSection "Authentication"

try {
    Connect-SyncServicePrincipal -Config $Config
    Write-LogInfo "Azure service principal connected"
    Initialize-SqlContext -Config $Config

    Initialize-ExoContext -Config $Config
    Connect-ExoSession    -Config $Config
    Write-LogInfo "Exchange Online connected"
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
$savedTimestamp = $null
$overallSw    = [System.Diagnostics.Stopwatch]::StartNew()

Write-LogInfo "Run ID: $runId"

# ──────────────────────────────────────────────────────────────
# Phase 1 — Fetch all mailboxes from EXO
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 1 — EXO mailbox fetch"

$allMailboxes = $null
try {
    $allMailboxes = Invoke-LoggedPhase -Name "ExoMailboxFetch" -ScriptBlock {
        Get-AllExoMailboxes -Config $Config
    }
    Write-LogInfo "EXO returned $($allMailboxes.Count) mailboxes"
}
catch {
    Write-LogError "Phase 1 (EXO fetch) failed — aborting" -ErrorRecord $_
    Disconnect-ExoSession
    Complete-SyncLogEntry -RunId $runId -Status "Failed" `
        -ErrorMessage "EXO fetch failed: $($_.Exception.Message)"
    Close-Logging -Status "Failed"
    exit 1
}

# Build a lookup set of ExchangeGuids from EXO for reconciliation in Phase 3
$exoGuids = [System.Collections.Generic.HashSet[string]]::new(
    ($allMailboxes | ForEach-Object { $_.ExchangeGuid.ToString() })
)

# ──────────────────────────────────────────────────────────────
# Phase 2 — MERGE each mailbox into the Mailboxes table
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 2 — SQL upsert"

$mergeSql = @"
MERGE dbo.Mailboxes AS target
USING (SELECT @ExchangeGuid AS ExchangeGuid) AS source
ON target.ExchangeGuid = source.ExchangeGuid
WHEN MATCHED THEN UPDATE SET
    UserId                  = @UserId,
    PrimarySmtpAddress      = @PrimarySmtp,
    UserPrincipalName       = @UPN,
    DisplayName             = @DisplayName,
    Alias                   = @Alias,
    RecipientTypeDetails    = @RecipientTypeDetails,
    MailboxType             = @MailboxType,
    HiddenFromAddressLists  = @HiddenFromAddressLists,
    LitigationHoldEnabled   = @LitigationHoldEnabled,
    ArchiveStatus           = @ArchiveStatus,
    ForwardingAddress       = @ForwardingAddress,
    ForwardingSmtpAddress   = @ForwardingSmtpAddress,
    GrantSendOnBehalfTo     = @GrantSendOnBehalfTo,
    IsDirSynced             = @IsDirSynced,
    WhenMailboxCreated      = @WhenMailboxCreated,
    WhenChangedUTC          = @WhenChangedUTC,
    LastSyncedAt            = SYSUTCDATETIME(),
    LastModifiedAt          = SYSUTCDATETIME(),
    IsDeleted               = 0,
    DeletedAt               = NULL,
    SyncSource              = 'Baseline',
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
    'Baseline', @RunId
);
"@

$existsSql        = "SELECT COUNT(1) FROM dbo.Mailboxes WHERE ExchangeGuid = @ExchangeGuid"
$progressInterval = 2000
$phaseSw          = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($m in $allMailboxes) {
    try {
        $exists = Invoke-SqlScalar -Query $existsSql `
            -Parameters @{ '@ExchangeGuid' = $m.ExchangeGuid }

        Invoke-SqlNonQuery -Query $mergeSql -Parameters @{
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
    }
    catch {
        $errors++
        Write-LogWarning "MERGE failed for $($m.PrimarySmtpAddress): $($_.Exception.Message)"
    }
    $processed++

    if ($processed % $progressInterval -eq 0) {
        Write-LogInfo "Progress: $processed / $($allMailboxes.Count) ($errors errors)"
    }
}

$phaseSw.Stop()
Write-LogInfo "SQL upsert complete in $([Math]::Round($phaseSw.Elapsed.TotalSeconds,1))s — $inserted inserted, $updated updated, $errors errors"

# ──────────────────────────────────────────────────────────────
# Phase 3 — Reconcile deletions
# Any mailbox in the DB (not deleted) that is no longer in EXO
# gets soft-deleted. This is the key advantage of the weekly
# baseline over the delta sync alone.
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 3 — Deletion reconciliation"

try {
    # Pull active ExchangeGuids from the database
    $dbTable = Invoke-SqlQuery -Query @"
SELECT CAST(ExchangeGuid AS NVARCHAR(36)) AS ExchangeGuid
FROM dbo.Mailboxes
WHERE IsDeleted = 0
"@

    $dbGuids = $dbTable | ForEach-Object { $_.ExchangeGuid }

    Write-LogInfo "DB has $($dbGuids.Count) active mailboxes; EXO returned $($exoGuids.Count)"

    $softDeleteSql = @"
UPDATE dbo.Mailboxes SET
    IsDeleted     = 1,
    DeletedAt     = SYSUTCDATETIME(),
    LastSyncedAt  = SYSUTCDATETIME(),
    SyncSource    = 'Baseline',
    LastSyncRunId = @RunId
WHERE ExchangeGuid = @ExchangeGuid AND IsDeleted = 0;
"@

    foreach ($guid in $dbGuids) {
        if (-not $exoGuids.Contains($guid)) {
            try {
                $rows = Invoke-SqlNonQuery -Query $softDeleteSql `
                    -Parameters @{ '@ExchangeGuid' = [guid]$guid; '@RunId' = $runId }
                if ($rows -gt 0) {
                    $softDeleted++
                    Write-LogInfo "Soft-deleted: $guid (no longer in EXO)"
                }
            }
            catch {
                $errors++
                Write-LogWarning "Failed to soft-delete $guid : $($_.Exception.Message)"
            }
        }
    }

    Write-LogInfo "Reconciliation complete — $softDeleted mailboxes soft-deleted"
}
catch {
    Write-LogError "Phase 3 (reconciliation) encountered an error" -ErrorRecord $_
    $errors++
}

# ──────────────────────────────────────────────────────────────
# Phase 4 — Save timestamp for delta sync
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 4 — Save delta timestamp"

try {
    $savedTimestamp = (Get-Date).ToUniversalTime().ToString('o')
    Save-DeltaToken -TokenName "mailboxes_delta_timestamp" -TokenValue $savedTimestamp
    Write-LogInfo "Delta timestamp saved: $savedTimestamp"
}
catch {
    Write-LogError "Failed to save delta timestamp" -ErrorRecord $_
    $errors++
}

# ──────────────────────────────────────────────────────────────
# Finalise
# ──────────────────────────────────────────────────────────────

$overallSw.Stop()
$finalStatus = if ($errors -gt 0 -and $inserted -eq 0 -and $updated -eq 0) { "Failed" }
               elseif ($errors -gt 0) { "PartialFailure" }
               else                    { "Success" }

Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET
    CompletedAt          = SYSUTCDATETIME(),
    Status               = @Status,
    MailboxesInserted    = @Inserted,
    MailboxesUpdated     = @Updated,
    MailboxesSoftDeleted = @Deleted,
    MailboxesProcessed   = @Processed,
    ErrorCount           = @Errors,
    ErrorMessage         = @ErrorMessage,
    TokenAdvancedTo      = @Token
WHERE RunId = @RunId
"@ -Parameters @{
    '@RunId'        = $runId
    '@Status'       = $finalStatus
    '@Inserted'     = $inserted
    '@Updated'      = $updated
    '@Deleted'      = $softDeleted
    '@Processed'    = $processed
    '@Errors'       = $errors
    '@ErrorMessage' = $null
    '@Token'        = $savedTimestamp
} | Out-Null

Write-LogSummary @{
    "Status"             = $finalStatus
    "Mailboxes inserted" = $inserted
    "Mailboxes updated"  = $updated
    "Soft-deleted"       = $softDeleted
    "Errors"             = $errors
    "Total processed"    = $processed
    "Duration"           = "$([Math]::Round($overallSw.Elapsed.TotalMinutes, 2)) minutes"
    "Run ID"             = $runId
}

Disconnect-ExoSession
Remove-OldLogFiles
Close-Logging -Status $finalStatus

if ($finalStatus -eq "Failed") { exit 1 }
