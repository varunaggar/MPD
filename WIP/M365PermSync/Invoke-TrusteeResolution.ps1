<#
.SYNOPSIS
    Resolves raw trustee identities to Users.UserId across all permission tables.

.DESCRIPTION
    Phase 4 of the M365 Permissions Sync solution. Permission tables store
    trustee identities in their raw form as returned by EXO / Admin API:
      - PrimarySmtpAddress (MFC group members)
      - UPN (EXO cmdlets)
      - DOMAIN\username (Get-EXOMailboxPermission)
      - Distinguished Name segments (GrantSendOnBehalfTo)
      - Display names (various)

    This script matches each TrusteeRawIdentity to a user in the Users table
    and sets the ResolvedUserId FK. It processes only rows where
    ResolvedUserId IS NULL, making it safe to run frequently.

    RESOLUTION ORDER (per trustee identity):
      1. Exact match on Users.UserPrincipalName
      2. Exact match on Users.Mail
      3. DOMAIN\user format → extract local part, match UPN prefix
      4. DN/path format → extract CN or last segment, match DisplayName
      5. No match → left as NULL (reported as unresolved)

    Ambiguous matches (e.g. two users with the same UPN prefix from
    different domains) are skipped and reported.

    TABLES UPDATED:
      FullAccessPermissions.ResolvedUserId
      SendAsPermissions.ResolvedUserId
      SendOnBehalfPermissions.ResolvedUserId
      FolderPermissions.ResolvedUserId
      MfcGroupMembers.ResolvedUserId

.PARAMETER ConfigPath
    Path to config.xml. Defaults to config.xml in the same folder.

.NOTES
    Schedule hourly or after each delta sync cycle.
    Typical duration: 1–5 minutes (mostly SQL reads + writes).

    To force a full re-resolution, run:
      UPDATE dbo.FullAccessPermissions SET ResolvedUserId = NULL WHERE IsDeleted = 0
      (repeat for each table)
    Then re-run this script.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.xml")
)

$ErrorActionPreference = "Stop"
$scriptName            = "Invoke-TrusteeResolution"
$sharedPath            = Join-Path $PSScriptRoot "shared"

# ══════════════════════════════════════════════════════════════
# Bootstrap
# ══════════════════════════════════════════════════════════════

Import-Module (Join-Path $sharedPath "ConfigHelpers.psm1")  -Force
Import-Module (Join-Path $sharedPath "LoggingHelpers.psm1") -Force
Import-Module (Join-Path $sharedPath "SqlHelpers.psm1")     -Force

try {
    $Config = Import-SyncConfig -Path $ConfigPath
    Initialize-Logging -Config $Config -ProcessName $scriptName
}
catch {
    Write-Host "FATAL BOOTSTRAP ERROR in $scriptName" -ForegroundColor Red
    Write-Host "Message : $($_.Exception.Message)"    -ForegroundColor White
    exit 1
}

Write-LogSection "Invoke-TrusteeResolution"
Write-LogInfo "Config   : $ConfigPath"
Write-LogInfo "Server   : $($Config.Database.Server)"
Write-LogInfo "Database : $($Config.Database.Name)"

# ══════════════════════════════════════════════════════════════
# Authentication
# ══════════════════════════════════════════════════════════════

Write-LogSection "Authentication"

try {
    Connect-SyncServicePrincipal -Config $Config
    Initialize-SqlContext -Config $Config
    Write-LogInfo "Service principal and SQL context connected"
}
catch {
    Write-LogError "Authentication failed" -ErrorRecord $_
    Close-Logging -Status "Failed"
    exit 1
}

# ══════════════════════════════════════════════════════════════
# Run tracking
# ══════════════════════════════════════════════════════════════

$runId     = [guid]::NewGuid()
$overallSw = [System.Diagnostics.Stopwatch]::StartNew()

Invoke-SqlNonQuery -Query @"
INSERT INTO dbo.SyncLog (RunId, FunctionName, StartedAt, Status)
VALUES (@RunId, @Function, SYSUTCDATETIME(), 'Running')
"@ -Parameters @{ '@RunId' = $runId; '@Function' = $scriptName } | Out-Null

$counters = @{
    Resolved   = 0   # distinct identities successfully matched
    Updated    = 0   # total DB rows updated across all tables
    Unresolved = 0   # distinct identities with no match
    Ambiguous  = 0   # distinct identities with multiple possible matches
    Skipped    = 0   # system trustees (Default, Anonymous)
    Errors     = 0
}

Write-LogInfo "Run ID: $runId"

# ══════════════════════════════════════════════════════════════
# PHASE 1 — Build lookup dictionaries from Users table
#
# Four lookup tiers, tried in order per trustee:
#   1. UPN (exact)        — most reliable
#   2. Mail/SMTP (exact)  — common for Admin API data
#   3. UPN local part     — for DOMAIN\user format
#   4. DisplayName        — last resort, for DN format
#
# Ambiguous entries (multiple users map to same key) are set to
# $null so they are skipped rather than resolved incorrectly.
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 1 — Build user lookup dictionaries"

$upnLookup         = @{}   # UPN (lower) → UserId
$mailLookup        = @{}   # Mail (lower) → UserId
$localPartLookup   = @{}   # UPN local part (lower) → UserId or $null if ambiguous
$displayNameLookup = @{}   # DisplayName (lower) → UserId or $null if ambiguous

try {
    $allUsers = Invoke-SqlQuery -Query @"
SELECT
    CAST(UserId AS NVARCHAR(36)) AS UserId,
    UserPrincipalName,
    Mail,
    DisplayName
FROM dbo.Users
WHERE IsDeleted = 0
"@

    foreach ($u in $allUsers) {
        $uid = $u.UserId

        # Tier 1: UPN
        if (-not [string]::IsNullOrWhiteSpace($u.UserPrincipalName)) {
            $upnLookup[$u.UserPrincipalName.ToLower()] = $uid
        }

        # Tier 2: Mail
        if (-not [string]::IsNullOrWhiteSpace($u.Mail)) {
            $key = $u.Mail.ToLower()
            if (-not $mailLookup.ContainsKey($key)) { $mailLookup[$key] = $uid }
        }

        # Tier 3: UPN local part (before @) for DOMAIN\user resolution
        if (-not [string]::IsNullOrWhiteSpace($u.UserPrincipalName)) {
            $local = ($u.UserPrincipalName -split '@')[0].ToLower()
            if ($localPartLookup.ContainsKey($local)) {
                $localPartLookup[$local] = $null   # ambiguous — two users with same local part
            }
            else {
                $localPartLookup[$local] = $uid
            }
        }

        # Tier 4: DisplayName (least reliable — collisions common)
        if (-not [string]::IsNullOrWhiteSpace($u.DisplayName)) {
            $dn = $u.DisplayName.ToLower()
            if ($displayNameLookup.ContainsKey($dn)) {
                $displayNameLookup[$dn] = $null   # ambiguous
            }
            else {
                $displayNameLookup[$dn] = $uid
            }
        }
    }

    Write-LogInfo "User lookups built — UPN: $($upnLookup.Count), Mail: $($mailLookup.Count), LocalPart: $($localPartLookup.Count), DisplayName: $($displayNameLookup.Count)"
}
catch {
    Write-LogError "Phase 1 failed — aborting" -ErrorRecord $_
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET CompletedAt=SYSUTCDATETIME(), Status='Failed',
    ErrorCount=1, ErrorMessage=@Msg WHERE RunId=@RunId
"@ -Parameters @{ '@RunId' = $runId; '@Msg' = "Phase 1 failed: $($_.Exception.Message)" } | Out-Null
    Close-Logging -Status "Failed"
    exit 1
}

# ══════════════════════════════════════════════════════════════
# PHASE 2 — Collect distinct unresolved trustee identities
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 2 — Collect unresolved trustees"

# The five tables and their IsDeleted filter
$permTables = @(
    @{ Table = 'FullAccessPermissions';    HasIsDeleted = $true  }
    @{ Table = 'SendAsPermissions';        HasIsDeleted = $true  }
    @{ Table = 'SendOnBehalfPermissions';  HasIsDeleted = $true  }
    @{ Table = 'FolderPermissions';        HasIsDeleted = $true  }
    @{ Table = 'MfcGroupMembers';          HasIsDeleted = $false }
)

$allUnresolved = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

foreach ($t in $permTables) {
    $whereClause = "WHERE ResolvedUserId IS NULL"
    if ($t.HasIsDeleted) { $whereClause += " AND IsDeleted = 0" }

    $rows = Invoke-SqlQuery -Query "SELECT DISTINCT TrusteeRawIdentity FROM dbo.$($t.Table) $whereClause"
    foreach ($r in $rows) {
        if (-not [string]::IsNullOrWhiteSpace($r.TrusteeRawIdentity)) {
            [void]$allUnresolved.Add($r.TrusteeRawIdentity)
        }
    }
}

Write-LogInfo "Total distinct unresolved trustee identities: $($allUnresolved.Count)"

if ($allUnresolved.Count -eq 0) {
    Write-LogInfo "Nothing to resolve — all trustees already resolved or no permission data"
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET CompletedAt=SYSUTCDATETIME(), Status='Success',
    UsersProcessed=0, ErrorCount=0 WHERE RunId=@RunId
"@ -Parameters @{ '@RunId' = $runId } | Out-Null
    Write-LogSummary @{ "Status" = "Success (nothing to resolve)" }
    Remove-OldLogFiles
    Close-Logging -Status "Success"
    exit 0
}

# ══════════════════════════════════════════════════════════════
# PHASE 3 — Resolve each identity
#
# Resolution order:
#   1. Skip system trustees (Default, Anonymous)
#   2. Exact UPN match
#   3. Exact Mail/SMTP match
#   4. DOMAIN\user → extract user, match UPN local part
#   5. DN or path format → extract CN / last segment, match DisplayName
#   6. Unresolved
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 3 — Resolve identities"

$resolvedMap = @{}   # TrusteeRawIdentity → UserId (only resolved ones)

foreach ($identity in $allUnresolved) {
    # System trustees — not real users
    if ($identity -in @('Default', 'Anonymous')) {
        $counters.Skipped++
        continue
    }

    $lower  = $identity.ToLower().Trim()
    $userId = $null

    # Tier 1: exact UPN
    $userId = $upnLookup[$lower]

    # Tier 2: exact Mail/SMTP
    if (-not $userId) {
        $userId = $mailLookup[$lower]
    }

    # Tier 3: DOMAIN\user format
    if (-not $userId -and $identity -match '^[^\\@]+\\(.+)$') {
        $localPart = $Matches[1].ToLower()
        if ($localPartLookup.ContainsKey($localPart)) {
            $resolved = $localPartLookup[$localPart]
            if ($null -eq $resolved) {
                $counters.Ambiguous++
                Write-LogWarning "Ambiguous DOMAIN\user: '$identity' — multiple users with local part '$localPart'"
                continue
            }
            $userId = $resolved
        }
    }

    # Tier 4: DN or path format → extract CN or last segment
    if (-not $userId) {
        $displayCandidate = $null

        # CN=John Smith,OU=Users,DC=contoso,DC=com
        if ($identity -match 'CN=([^,]+)') {
            $displayCandidate = $Matches[1].ToLower()
        }
        # contoso.com/Users/John Smith
        elseif ($identity -match '/([^/]+)$') {
            $displayCandidate = $Matches[1].ToLower()
        }

        if ($displayCandidate) {
            if ($displayNameLookup.ContainsKey($displayCandidate)) {
                $resolved = $displayNameLookup[$displayCandidate]
                if ($null -eq $resolved) {
                    $counters.Ambiguous++
                    Write-LogWarning "Ambiguous display name: '$identity' → '$displayCandidate' matches multiple users"
                    continue
                }
                $userId = $resolved
            }
        }
    }

    if ($userId) {
        $resolvedMap[$identity] = $userId
        $counters.Resolved++
    }
    else {
        $counters.Unresolved++
    }
}

$resolveRate = if ($allUnresolved.Count -gt 0) {
    [Math]::Round(($counters.Resolved / ($allUnresolved.Count - $counters.Skipped)) * 100, 1)
} else { 0 }

Write-LogInfo "Resolution complete — Resolved: $($counters.Resolved), Unresolved: $($counters.Unresolved), Ambiguous: $($counters.Ambiguous), Skipped: $($counters.Skipped) ($resolveRate% match rate)"

# ══════════════════════════════════════════════════════════════
# PHASE 4 — Bulk UPDATE each permission table
#
# For each resolved identity, UPDATE all matching rows across
# all five tables. One SQL connection held open for the phase.
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 4 — Bulk UPDATE"

if ($resolvedMap.Count -eq 0) {
    Write-LogInfo "No identities were resolved — skipping UPDATE"
}
else {
    $phaseSw = [System.Diagnostics.Stopwatch]::StartNew()

    $conn = Open-SqlConnection
    try {
        foreach ($t in $permTables) {
            $tableName   = $t.Table
            $whereFilter = if ($t.HasIsDeleted) { "AND IsDeleted = 0" } else { "" }

            $updateSql = @"
UPDATE dbo.$tableName
SET ResolvedUserId = @UserId
WHERE TrusteeRawIdentity = @Trustee AND ResolvedUserId IS NULL $whereFilter
"@
            $tableUpdated = 0

            foreach ($entry in $resolvedMap.GetEnumerator()) {
                try {
                    $rows = Invoke-SqlNonQuery -Connection $conn -Query $updateSql -Parameters @{
                        '@UserId'  = [guid]$entry.Value
                        '@Trustee' = $entry.Key
                    }
                    $tableUpdated    += $rows
                    $counters.Updated += $rows
                }
                catch {
                    $counters.Errors++
                    Write-LogWarning "UPDATE failed on $tableName for '$($entry.Key)': $($_.Exception.Message)"
                }
            }

            Write-LogInfo "  $tableName — $tableUpdated rows updated"
        }
    }
    finally {
        $conn.Dispose()
    }

    $phaseSw.Stop()
    Write-LogInfo "Phase 4 complete in $([Math]::Round($phaseSw.Elapsed.TotalSeconds,1))s — $($counters.Updated) total rows updated"
}

# ══════════════════════════════════════════════════════════════
# PHASE 5 — Report unresolved trustees
# Log the top unresolved identities for investigation.
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 5 — Unresolved report"

if ($counters.Unresolved -gt 0) {
    $unresolvedList = $allUnresolved |
        Where-Object { $_ -notin @('Default', 'Anonymous') -and -not $resolvedMap.ContainsKey($_) } |
        Select-Object -First 50

    Write-LogInfo "Top unresolved trustees (max 50):"
    foreach ($u in $unresolvedList) {
        Write-LogInfo "  UNRESOLVED: $u"
    }

    if ($counters.Unresolved -gt 50) {
        Write-LogInfo "  ... and $($counters.Unresolved - 50) more"
    }
}
else {
    Write-LogInfo "All trustees resolved successfully"
}

# ══════════════════════════════════════════════════════════════
# Finalise
# ══════════════════════════════════════════════════════════════

$overallSw.Stop()
$finalStatus = if ($counters.Errors -gt 0) { "PartialFailure" } else { "Success" }

Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET
    CompletedAt     = SYSUTCDATETIME(),
    Status          = @Status,
    UsersProcessed  = @Processed,
    UsersInserted   = @Resolved,
    UsersUpdated    = @Updated,
    MailboxesSkipped= @Unresolved,
    ErrorCount      = @Errors,
    ErrorMessage    = @Msg
WHERE RunId = @RunId
"@ -Parameters @{
    '@RunId'      = $runId
    '@Status'     = $finalStatus
    '@Processed'  = $allUnresolved.Count
    '@Resolved'   = $counters.Resolved
    '@Updated'    = $counters.Updated
    '@Unresolved' = $counters.Unresolved
    '@Errors'     = $counters.Errors
    '@Msg'        = if ($counters.Unresolved -gt 0) { "$($counters.Unresolved) unresolved, $($counters.Ambiguous) ambiguous — check log" } else { $null }
} | Out-Null

Write-LogSummary @{
    "Status"              = $finalStatus
    "Distinct identities" = $allUnresolved.Count
    "Resolved"            = $counters.Resolved
    "DB rows updated"     = $counters.Updated
    "Unresolved"          = $counters.Unresolved
    "Ambiguous"           = $counters.Ambiguous
    "System (skipped)"    = $counters.Skipped
    "Match rate"          = "${resolveRate}%"
    "Errors"              = $counters.Errors
    "Duration"            = "$([Math]::Round($overallSw.Elapsed.TotalSeconds, 1))s"
    "Run ID"              = $runId
}

Remove-OldLogFiles
Close-Logging -Status $finalStatus

if ($finalStatus -eq "Failed") { exit 1 }
