<#
.SYNOPSIS
    Reusable SQL helpers for the M365 permissions sync solution.

.DESCRIPTION
    Imported by every Azure Function via profile.ps1.
    Provides parameterised query helpers using Azure AD token authentication
    via the Managed Identity. Connection pooling is handled by the .NET
    SqlClient library automatically — we open and close on each call.

.NOTES
    Dependencies: Az.Accounts (loaded via requirements.psd1)
    Assumes Connect-AzAccount -Identity has run (in profile.ps1).

    Environment variables required:
      SQL_SERVER   — fully-qualified server name (e.g. sql-foo.database.windows.net)
      SQL_DATABASE — database name (e.g. db-m365permissions)
#>

# ──────────────────────────────────────────────────────────────
# Module-scoped state
# ──────────────────────────────────────────────────────────────

$script:SqlServer   = $env:SQL_SERVER
$script:SqlDatabase = $env:SQL_DATABASE

if (-not $script:SqlServer -or -not $script:SqlDatabase) {
    throw "SqlHelpers: SQL_SERVER and SQL_DATABASE environment variables must be set."
}

# Token cache — Azure SQL tokens are valid ~1 hour. Cache to avoid reacquiring
# on every query within a function execution.
$script:SqlTokenCache = @{
    Token     = $null
    ExpiresAt = [datetime]::MinValue
}

# ──────────────────────────────────────────────────────────────
# Private: Get-SqlAccessToken
# Acquires (or returns cached) Azure AD token for SQL
# ──────────────────────────────────────────────────────────────

function Get-SqlAccessToken {
    param([switch]$ForceRefresh)

    $now = [datetime]::UtcNow

    if (-not $ForceRefresh -and
        $script:SqlTokenCache.Token -and
        $script:SqlTokenCache.ExpiresAt -gt $now.AddMinutes(2)) {
        return $script:SqlTokenCache.Token
    }

    try {
        $tokenInfo = Get-AzAccessToken -ResourceUrl "https://database.windows.net/" -ErrorAction Stop

        $script:SqlTokenCache.Token     = $tokenInfo.Token
        # Get-AzAccessToken returns ExpiresOn as DateTimeOffset; convert to UTC datetime
        $script:SqlTokenCache.ExpiresAt = $tokenInfo.ExpiresOn.UtcDateTime.AddMinutes(-5)

        return $tokenInfo.Token
    }
    catch {
        throw "Failed to acquire SQL access token: $($_.Exception.Message)"
    }
}

# ──────────────────────────────────────────────────────────────
# Private: New-SqlConnection
# Builds and opens a SqlConnection using AAD token auth
# ──────────────────────────────────────────────────────────────

function New-SqlConnection {
    $conn = New-Object System.Data.SqlClient.SqlConnection
    $conn.ConnectionString = "Server=$($script:SqlServer);Database=$($script:SqlDatabase);Encrypt=True;TrustServerCertificate=False;Connection Timeout=30"
    $conn.AccessToken = Get-SqlAccessToken
    $conn.Open()
    return $conn
}

# ──────────────────────────────────────────────────────────────
# Private: Add-SqlParameters
# Adds parameters to a SqlCommand, handling $null → DBNull
# ──────────────────────────────────────────────────────────────

function Add-SqlParameters {
    param(
        [System.Data.SqlClient.SqlCommand]$Command,
        [hashtable]$Parameters
    )

    if (-not $Parameters) { return }

    foreach ($p in $Parameters.GetEnumerator()) {
        $value = if ($null -eq $p.Value) { [System.DBNull]::Value } else { $p.Value }
        [void]$Command.Parameters.AddWithValue($p.Key, $value)
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Invoke-SqlNonQuery
# Executes INSERT/UPDATE/DELETE/MERGE. Returns rows affected.
# ──────────────────────────────────────────────────────────────

function Invoke-SqlNonQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Query,
        [hashtable]$Parameters = @{},
        [int]$TimeoutSec = 60
    )

    $conn = $null
    try {
        $conn = New-SqlConnection
        $cmd = $conn.CreateCommand()
        $cmd.CommandText    = $Query
        $cmd.CommandTimeout = $TimeoutSec
        Add-SqlParameters -Command $cmd -Parameters $Parameters

        return $cmd.ExecuteNonQuery()
    }
    catch {
        throw "SQL non-query failed: $($_.Exception.Message)"
    }
    finally {
        if ($conn) { $conn.Close(); $conn.Dispose() }
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Invoke-SqlScalar
# Executes a query that returns a single scalar value.
# ──────────────────────────────────────────────────────────────

function Invoke-SqlScalar {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Query,
        [hashtable]$Parameters = @{},
        [int]$TimeoutSec = 30
    )

    $conn = $null
    try {
        $conn = New-SqlConnection
        $cmd = $conn.CreateCommand()
        $cmd.CommandText    = $Query
        $cmd.CommandTimeout = $TimeoutSec
        Add-SqlParameters -Command $cmd -Parameters $Parameters

        $result = $cmd.ExecuteScalar()

        if ($null -eq $result -or $result -is [System.DBNull]) { return $null }
        return $result
    }
    catch {
        throw "SQL scalar query failed: $($_.Exception.Message)"
    }
    finally {
        if ($conn) { $conn.Close(); $conn.Dispose() }
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Invoke-SqlQuery
# Executes a SELECT and returns rows as a DataTable
# ──────────────────────────────────────────────────────────────

function Invoke-SqlQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Query,
        [hashtable]$Parameters = @{},
        [int]$TimeoutSec = 60
    )

    $conn = $null
    try {
        $conn = New-SqlConnection
        $cmd = $conn.CreateCommand()
        $cmd.CommandText    = $Query
        $cmd.CommandTimeout = $TimeoutSec
        Add-SqlParameters -Command $cmd -Parameters $Parameters

        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        $dataset = New-Object System.Data.DataSet
        [void]$adapter.Fill($dataset)

        if ($dataset.Tables.Count -eq 0) { return @() }
        return $dataset.Tables[0]
    }
    catch {
        throw "SQL query failed: $($_.Exception.Message)"
    }
    finally {
        if ($conn) { $conn.Close(); $conn.Dispose() }
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Invoke-SqlBatch
# Executes many parameterised statements within a single connection.
# Use for bulk operations to avoid open/close per statement.
# Each item in -Statements is @{ Query = "..."; Parameters = @{...} }
# Returns total rows affected across all statements.
# ──────────────────────────────────────────────────────────────

function Invoke-SqlBatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [array]$Statements,
        [int]$TimeoutSec = 60,
        [switch]$UseTransaction
    )

    $conn = $null
    $tx = $null
    $totalRows = 0

    try {
        $conn = New-SqlConnection
        if ($UseTransaction) { $tx = $conn.BeginTransaction() }

        foreach ($stmt in $Statements) {
            $cmd = $conn.CreateCommand()
            $cmd.CommandText    = $stmt.Query
            $cmd.CommandTimeout = $TimeoutSec
            if ($tx) { $cmd.Transaction = $tx }
            Add-SqlParameters -Command $cmd -Parameters $stmt.Parameters

            $totalRows += $cmd.ExecuteNonQuery()
        }

        if ($tx) { $tx.Commit() }
        return $totalRows
    }
    catch {
        if ($tx) {
            try { $tx.Rollback() } catch {}
        }
        throw "SQL batch failed: $($_.Exception.Message)"
    }
    finally {
        if ($tx)   { $tx.Dispose() }
        if ($conn) { $conn.Close(); $conn.Dispose() }
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Start-SyncLogEntry
# Inserts a new SyncLog row in 'Running' state. Returns the RunId.
# ──────────────────────────────────────────────────────────────

function Start-SyncLogEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$FunctionName
    )

    $runId = [guid]::NewGuid()

    Invoke-SqlNonQuery -Query @"
INSERT INTO SyncLog (RunId, FunctionName, StartedAt, Status)
VALUES (@RunId, @FunctionName, SYSUTCDATETIME(), 'Running')
"@ -Parameters @{
        "@RunId"        = $runId
        "@FunctionName" = $FunctionName
    } | Out-Null

    return $runId
}

# ──────────────────────────────────────────────────────────────
# Public: Complete-SyncLogEntry
# Updates an existing SyncLog row with final counts and status.
# ──────────────────────────────────────────────────────────────

function Complete-SyncLogEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [guid]$RunId,
        [Parameter(Mandatory)] [ValidateSet('Success','PartialFailure','Failed')] [string]$Status,
        [int]$UsersInserted    = 0,
        [int]$UsersUpdated     = 0,
        [int]$UsersSoftDeleted = 0,
        [int]$UsersProcessed   = 0,
        [int]$ErrorCount       = 0,
        [string]$ErrorMessage  = $null,
        [string]$TokenAdvancedTo = $null
    )

    Invoke-SqlNonQuery -Query @"
UPDATE SyncLog SET
    CompletedAt      = SYSUTCDATETIME(),
    Status           = @Status,
    UsersInserted    = @Inserted,
    UsersUpdated     = @Updated,
    UsersSoftDeleted = @Deleted,
    UsersProcessed   = @Processed,
    ErrorCount       = @Errors,
    ErrorMessage     = @ErrorMessage,
    TokenAdvancedTo  = @Token
WHERE RunId = @RunId
"@ -Parameters @{
        "@RunId"        = $RunId
        "@Status"       = $Status
        "@Inserted"     = $UsersInserted
        "@Updated"      = $UsersUpdated
        "@Deleted"      = $UsersSoftDeleted
        "@Processed"    = $UsersProcessed
        "@Errors"       = $ErrorCount
        "@ErrorMessage" = $ErrorMessage
        "@Token"        = $TokenAdvancedTo
    } | Out-Null
}

# ──────────────────────────────────────────────────────────────
# Public: Get-DeltaToken / Save-DeltaToken / Disable-DeltaToken
# DeltaTokens table convenience helpers
# ──────────────────────────────────────────────────────────────

function Get-DeltaToken {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string]$TokenName)

    return Invoke-SqlScalar -Query @"
SELECT TokenValue FROM DeltaTokens
WHERE TokenName = @Name AND IsActive = 1
"@ -Parameters @{ "@Name" = $TokenName }
}

function Save-DeltaToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$TokenName,
        [Parameter(Mandatory)] [string]$TokenValue
    )

    Invoke-SqlNonQuery -Query @"
MERGE DeltaTokens AS target
USING (SELECT @Name AS TokenName, @Value AS TokenValue) AS source
ON target.TokenName = source.TokenName
WHEN MATCHED THEN UPDATE SET
    TokenValue        = source.TokenValue,
    UpdatedAt         = SYSUTCDATETIME(),
    IsActive          = 1,
    DeactivatedAt     = NULL,
    DeactivationReason = NULL
WHEN NOT MATCHED THEN INSERT
    (TokenName, TokenValue, CreatedAt, UpdatedAt, IsActive)
    VALUES (source.TokenName, source.TokenValue, SYSUTCDATETIME(), SYSUTCDATETIME(), 1);
"@ -Parameters @{
        "@Name"  = $TokenName
        "@Value" = $TokenValue
    } | Out-Null
}

function Disable-DeltaToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$TokenName,
        [string]$Reason = 'Deactivated by sync function'
    )

    Invoke-SqlNonQuery -Query @"
UPDATE DeltaTokens SET
    IsActive           = 0,
    DeactivatedAt      = SYSUTCDATETIME(),
    DeactivationReason = @Reason
WHERE TokenName = @Name
"@ -Parameters @{
        "@Name"   = $TokenName
        "@Reason" = $Reason
    } | Out-Null
}

# ──────────────────────────────────────────────────────────────
# Exports
# ──────────────────────────────────────────────────────────────

Export-ModuleMember -Function @(
    'Invoke-SqlNonQuery',
    'Invoke-SqlScalar',
    'Invoke-SqlQuery',
    'Invoke-SqlBatch',
    'Start-SyncLogEntry',
    'Complete-SyncLogEntry',
    'Get-DeltaToken',
    'Save-DeltaToken',
    'Disable-DeltaToken'
)
