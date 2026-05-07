<#
.SYNOPSIS
    Baseline load: full snapshot of all users from Graph into the Users table.

.DESCRIPTION
    HTTP-triggered function that runs once on initial deployment.

    Performs three tasks:
      1. Pages through /users and upserts every record into Users (SyncSource='Baseline').
      2. Initiates /users/delta to capture a starting delta token.
      3. Stores the delta token in the DeltaTokens table.

    After this completes successfully, the Invoke-UserDeltaSync timer function
    can run from the captured token to keep the table in sync going forward.

.NOTES
    Trigger: HTTP POST  (function key required)
    Expected duration: 10–30 minutes for a 40K-user tenant.
    Re-runnable: yes — MERGE makes upserts idempotent.
                 Re-running will reset the Users table to the current Graph state
                 and replace the existing delta token.
#>

using namespace System.Net

param($Request, $TriggerMetadata)

$ErrorActionPreference = 'Stop'

# Helpers (Graph, SQL, Logging) are imported via profile.ps1 — available here.

# ──────────────────────────────────────────────────────────────
# Initialise run context
# ──────────────────────────────────────────────────────────────

$functionName = 'Invoke-UserBaselineLoad'
$runId = Start-SyncLogEntry -FunctionName $functionName
Initialize-LoggingContext -FunctionName $functionName -RunId $runId

Write-LogInfo "Baseline load started"
Write-Telemetry -EventName 'BaselineLoadStarted'

$inserted  = 0
$processed = 0
$errors    = 0
$capturedToken = $null

try {
    # ──────────────────────────────────────────────────────────
    # Phase 1 — Page through /users and upsert every row
    # ──────────────────────────────────────────────────────────

    Measure-Phase -Name 'GraphPagedFetch' -ScriptBlock {

        $url = 'https://graph.microsoft.com/v1.0/users' +
               '?$select=id,userPrincipalName,displayName,mail,accountEnabled,' +
                        'userType,onPremisesSyncEnabled,department,jobTitle,createdDateTime' +
               '&$top=999'

        $allUsers = Invoke-GraphPagedRequest -Uri $url
        Write-LogInfo "Graph returned $($allUsers.Count) users"
        $script:fetchedUsers = $allUsers
    }

    # ──────────────────────────────────────────────────────────
    # Phase 2 — Upsert every user. MERGE handles insert-or-update.
    # ──────────────────────────────────────────────────────────

    Measure-Phase -Name 'SqlUpsert' -ScriptBlock {

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

        foreach ($u in $script:fetchedUsers) {
            try {
                Invoke-SqlNonQuery -Query $mergeSql -Parameters @{
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

                $script:inserted++
            }
            catch {
                $script:errors++
                Write-LogWarning "Upsert failed for $($u.userPrincipalName): $($_.Exception.Message)"
            }
            $script:processed++

            if ($script:processed % 5000 -eq 0) {
                Write-LogInfo "Upserted $($script:processed) users so far"
            }
        }
    }

    # ──────────────────────────────────────────────────────────
    # Phase 3 — Capture starting delta token
    # We use $select=id only since we don't care about returned objects here,
    # only the deltaLink at the end of the response chain.
    # ──────────────────────────────────────────────────────────

    Measure-Phase -Name 'CaptureDeltaToken' -ScriptBlock {

        $deltaInitUrl = 'https://graph.microsoft.com/v1.0/users/delta?$select=id'

        $result = Invoke-GraphDeltaQuery -Uri $deltaInitUrl

        if ($result.TokenExpired -or -not $result.DeltaToken) {
            throw "Could not capture starting delta token"
        }

        Save-DeltaToken -TokenName 'users_delta' -TokenValue $result.DeltaToken
        $script:capturedToken = $result.DeltaToken

        Write-LogInfo "Starting delta token saved"
    }

    # ──────────────────────────────────────────────────────────
    # Final SyncLog update
    # ──────────────────────────────────────────────────────────

    $finalStatus = if ($errors -gt 0) { 'PartialFailure' } else { 'Success' }

    Complete-SyncLogEntry `
        -RunId $runId `
        -Status $finalStatus `
        -UsersInserted $inserted `
        -UsersProcessed $processed `
        -ErrorCount $errors `
        -TokenAdvancedTo $capturedToken

    Write-Telemetry -EventName 'BaselineLoadCompleted' `
        -Properties @{ status = $finalStatus } `
        -Metrics    @{
            usersInserted  = $inserted
            usersProcessed = $processed
            errorCount     = $errors
        }

    Write-LogInfo "Baseline load complete: $processed processed, $inserted upserted, $errors errors"

    # HTTP response
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body = (@{
            status         = $finalStatus
            runId          = $runId.ToString()
            usersProcessed = $processed
            usersInserted  = $inserted
            errorCount     = $errors
            tokenCaptured  = ($null -ne $capturedToken)
        } | ConvertTo-Json)
        Headers = @{ 'Content-Type' = 'application/json' }
    })
}
catch {
    Write-LogError "Baseline load failed" -ErrorRecord $_

    Complete-SyncLogEntry `
        -RunId $runId `
        -Status 'Failed' `
        -UsersInserted $inserted `
        -UsersProcessed $processed `
        -ErrorCount ($errors + 1) `
        -ErrorMessage $_.Exception.Message

    Write-Telemetry -EventName 'BaselineLoadFailed' `
        -Properties @{ errorMessage = $_.Exception.Message } `
        -Metrics    @{ usersProcessed = $processed }

    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::InternalServerError
        Body = (@{
            status       = 'Failed'
            runId        = $runId.ToString()
            errorMessage = $_.Exception.Message
        } | ConvertTo-Json)
        Headers = @{ 'Content-Type' = 'application/json' }
    })

    throw
}
