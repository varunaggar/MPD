<#
.SYNOPSIS
    Delta sync: keeps the Users table aligned with Entra ID.

.DESCRIPTION
    Timer-triggered function that runs every 15 minutes.

    1. Reads the stored users_delta token from DeltaTokens.
    2. Calls /users/delta?$deltatoken=... to get only changes since last run.
    3. For each returned object:
         - "@removed" marker → soft-delete in Users
         - new id            → INSERT
         - existing id       → UPDATE
    4. Persists the new delta token ONLY after all rows processed
       (so a mid-run crash means the next run reprocesses, never skips).
    5. On HTTP 410 Gone (token expired), deactivates the token and exits;
       the operator must re-run Invoke-UserBaselineLoad to recover.

.NOTES
    Trigger: Timer "0 */15 * * * *"
    Idempotent: yes — MERGE upsert plus token-save-after-success pattern.
#>

param($Timer)

$ErrorActionPreference = 'Stop'

# Helpers (Graph, SQL, Logging) imported via profile.ps1

# ──────────────────────────────────────────────────────────────
# Initialise run context
# ──────────────────────────────────────────────────────────────

$functionName = 'Invoke-UserDeltaSync'
$runId = Start-SyncLogEntry -FunctionName $functionName
Initialize-LoggingContext -FunctionName $functionName -RunId $runId

if ($Timer.IsPastDue) {
    Write-LogWarning "Timer is past due — execution may have been delayed"
}

Write-LogInfo "Delta sync started"

$inserted    = 0
$updated     = 0
$softDeleted = 0
$processed   = 0
$errors      = 0
$newToken    = $null

try {
    # ──────────────────────────────────────────────────────────
    # Phase 1 — Fetch stored delta token
    # ──────────────────────────────────────────────────────────

    $storedToken = Get-DeltaToken -TokenName 'users_delta'

    if (-not $storedToken) {
        $msg = "No active users_delta token. Run Invoke-UserBaselineLoad first."
        Write-LogError $msg
        Complete-SyncLogEntry -RunId $runId -Status 'Failed' -ErrorMessage $msg
        Write-Telemetry -EventName 'DeltaSyncSkipped' `
            -Properties @{ reason = 'NoActiveToken' }
        return
    }

    # ──────────────────────────────────────────────────────────
    # Phase 2 — Run the delta query
    # ──────────────────────────────────────────────────────────

    $deltaResult = Measure-Phase -Name 'GraphDeltaFetch' -ScriptBlock {
        $url = 'https://graph.microsoft.com/v1.0/users/delta' +
               '?$select=id,userPrincipalName,displayName,mail,accountEnabled,' +
                        'userType,onPremisesSyncEnabled,department,jobTitle,createdDateTime' +
               "&`$deltatoken=$storedToken"

        Invoke-GraphDeltaQuery -Uri $url
    }

    # Token expired — disable and bail; baseline load is required to recover
    if ($deltaResult.TokenExpired) {
        Write-LogWarning "Delta token expired (HTTP 410). Deactivating and exiting."
        Disable-DeltaToken -TokenName 'users_delta' `
            -Reason 'HTTP 410 Gone — run Invoke-UserBaselineLoad to reset'

        Complete-SyncLogEntry -RunId $runId -Status 'Failed' `
            -ErrorMessage 'Delta token expired; baseline reload required.'

        Write-Telemetry -EventName 'DeltaTokenExpired'
        return
    }

    $changes = $deltaResult.Objects
    Write-LogInfo "Delta returned $($changes.Count) changes"

    # ──────────────────────────────────────────────────────────
    # Phase 3 — Apply changes to Users table
    # ──────────────────────────────────────────────────────────

    Measure-Phase -Name 'ApplyChanges' -ScriptBlock {

        # Pre-built SQL templates we'll re-use
        $upsertSql = @"
MERGE Users AS target
USING (SELECT @UserId AS UserId) AS source
ON target.UserId = source.UserId
WHEN MATCHED THEN UPDATE SET
    UserPrincipalName     = COALESCE(@UPN, UserPrincipalName),
    DisplayName           = COALESCE(@DisplayName, DisplayName),
    Mail                  = COALESCE(@Mail, Mail),
    AccountEnabled        = COALESCE(@AccountEnabled, AccountEnabled),
    UserType              = COALESCE(@UserType, UserType),
    OnPremisesSyncEnabled = COALESCE(@OnPremisesSyncEnabled, OnPremisesSyncEnabled),
    Department            = COALESCE(@Department, Department),
    JobTitle              = COALESCE(@JobTitle, JobTitle),
    EntraCreatedDateTime  = COALESCE(@CreatedDateTime, EntraCreatedDateTime),
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

        $deleteSql = @"
UPDATE Users SET
    IsDeleted     = 1,
    DeletedAt     = SYSUTCDATETIME(),
    LastSyncedAt  = SYSUTCDATETIME(),
    LastModifiedAt= SYSUTCDATETIME(),
    SyncSource    = 'Delta',
    LastSyncRunId = @RunId
WHERE UserId = @UserId AND IsDeleted = 0;
"@

        $existsSql = "SELECT COUNT(*) FROM Users WHERE UserId = @UserId"

        foreach ($u in $changes) {
            try {
                # Deletion marker → soft-delete
                if ($u.'@removed') {
                    $rowsAffected = Invoke-SqlNonQuery -Query $deleteSql -Parameters @{
                        '@UserId' = [guid]$u.id
                        '@RunId'  = $runId
                    }
                    if ($rowsAffected -gt 0) { $script:softDeleted++ }
                    $script:processed++
                    continue
                }

                # Distinguish insert from update for accurate counters
                $exists = Invoke-SqlScalar -Query $existsSql `
                    -Parameters @{ '@UserId' = [guid]$u.id }

                Invoke-SqlNonQuery -Query $upsertSql -Parameters @{
                    '@UserId'                = [guid]$u.id
                    '@UPN'                   = $u.userPrincipalName
                    '@DisplayName'           = $u.displayName
                    '@Mail'                  = $u.mail
                    '@AccountEnabled'        = if ($null -eq $u.accountEnabled) { $null } else { [bool]$u.accountEnabled }
                    '@UserType'              = $u.userType
                    '@OnPremisesSyncEnabled' = if ($null -eq $u.onPremisesSyncEnabled) { $null } else { [bool]$u.onPremisesSyncEnabled }
                    '@Department'            = $u.department
                    '@JobTitle'              = $u.jobTitle
                    '@CreatedDateTime'       = if ($u.createdDateTime) { [datetime]$u.createdDateTime } else { $null }
                    '@RunId'                 = $runId
                } | Out-Null

                if ($exists -eq 0) { $script:inserted++ } else { $script:updated++ }
                $script:processed++
            }
            catch {
                $script:errors++
                Write-LogWarning "Failed to process $($u.id): $($_.Exception.Message)"
            }
        }
    }

    # ──────────────────────────────────────────────────────────
    # Phase 4 — Persist new token AFTER all processing succeeded
    # If we crashed before this, the next run will reprocess the same window
    # rather than skip changes.
    # ──────────────────────────────────────────────────────────

    if ($deltaResult.DeltaToken) {
        Save-DeltaToken -TokenName 'users_delta' -TokenValue $deltaResult.DeltaToken
        $newToken = $deltaResult.DeltaToken
    }

    # ──────────────────────────────────────────────────────────
    # Final SyncLog update
    # ──────────────────────────────────────────────────────────

    $finalStatus = if ($errors -gt 0) { 'PartialFailure' } else { 'Success' }

    Complete-SyncLogEntry `
        -RunId $runId `
        -Status $finalStatus `
        -UsersInserted $inserted `
        -UsersUpdated $updated `
        -UsersSoftDeleted $softDeleted `
        -UsersProcessed $processed `
        -ErrorCount $errors `
        -TokenAdvancedTo $newToken

    Write-Telemetry -EventName 'DeltaSyncCompleted' `
        -Properties @{ status = $finalStatus } `
        -Metrics    @{
            usersInserted    = $inserted
            usersUpdated     = $updated
            usersSoftDeleted = $softDeleted
            usersProcessed   = $processed
            errorCount       = $errors
        }

    Write-LogInfo "Delta sync complete: $inserted new, $updated updated, $softDeleted deleted, $errors errors"
}
catch {
    Write-LogError "Delta sync failed" -ErrorRecord $_

    Complete-SyncLogEntry `
        -RunId $runId `
        -Status 'Failed' `
        -UsersInserted $inserted `
        -UsersUpdated $updated `
        -UsersSoftDeleted $softDeleted `
        -UsersProcessed $processed `
        -ErrorCount ($errors + 1) `
        -ErrorMessage $_.Exception.Message

    Write-Telemetry -EventName 'DeltaSyncFailed' `
        -Properties @{ errorMessage = $_.Exception.Message }

    throw
}
