<#
.SYNOPSIS
    Azure SQL Database helpers — certificate-based authentication.

.DESCRIPTION
    Provides parameterised query helpers authenticating to Azure SQL
    using an Azure AD token (obtained via the service principal
    context established by Connect-SyncServicePrincipal).

    No SQL username/password is used anywhere. The app registration
    must have an EXTERNAL PROVIDER user in the target database with
    the required roles.

.NOTES
    Dependencies: Az.Accounts (for Get-AzAccessToken)
    Config is passed in via Initialize-SqlContext.

    Required SQL setup (run once as SQL admin):
        CREATE USER [your-app-display-name] FROM EXTERNAL PROVIDER;
        ALTER ROLE db_datareader ADD MEMBER [your-app-display-name];
        ALTER ROLE db_datawriter ADD MEMBER [your-app-display-name];
        GRANT EXECUTE TO [your-app-display-name];
#>

# ──────────────────────────────────────────────────────────────
# Module-scoped state
# ──────────────────────────────────────────────────────────────

$script:SqlServer          = $null
$script:SqlDatabase        = $null
$script:ConnectionTimeout  = 30
$script:CommandTimeout     = 120

$script:SqlTokenCache = @{
    Token     = $null
    ExpiresAt = [datetime]::MinValue
}

# ──────────────────────────────────────────────────────────────
# Public: Initialize-SqlContext
# Called once per script after loading config.
# ──────────────────────────────────────────────────────────────

function Initialize-SqlContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Config
    )

    $script:SqlServer         = $Config.Database.Server
    $script:SqlDatabase       = $Config.Database.Name
    $script:ConnectionTimeout = [int]($Config.Database.ConnectionTimeoutSec ?? 30)
    $script:CommandTimeout    = [int]($Config.Database.CommandTimeoutSec ?? 120)

    # Reset token cache
    $script:SqlTokenCache.Token     = $null
    $script:SqlTokenCache.ExpiresAt = [datetime]::MinValue

    Write-Verbose "SQL context initialised (Server=$($script:SqlServer), DB=$($script:SqlDatabase))"
}

# ──────────────────────────────────────────────────────────────
# Private: Get-SqlAccessToken
# Acquires (or returns cached) Azure AD token for SQL.
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
        $tokenInfo = Get-AzAccessToken `
            -ResourceUrl "https://database.windows.net/" `
            -ErrorAction Stop

        $script:SqlTokenCache.Token     = $tokenInfo.Token
        $script:SqlTokenCache.ExpiresAt = $tokenInfo.ExpiresOn.UtcDateTime.AddMinutes(-5)

        Write-Verbose "SQL access token acquired. Valid until $($script:SqlTokenCache.ExpiresAt) UTC"
        return $tokenInfo.Token
    }
    catch {
        throw "Failed to acquire SQL access token: $($_.Exception.Message)"
    }
}

# ──────────────────────────────────────────────────────────────
# Private: New-SqlConnection
# Opens an authenticated connection to Azure SQL.
# ──────────────────────────────────────────────────────────────

function New-SqlConnection {
    $connStr = "Server=$($script:SqlServer);Database=$($script:SqlDatabase);" +
               "Encrypt=True;TrustServerCertificate=False;" +
               "Connection Timeout=$($script:ConnectionTimeout)"

    $conn             = New-Object System.Data.SqlClient.SqlConnection
    $conn.ConnectionString = $connStr
    $conn.AccessToken      = Get-SqlAccessToken
    $conn.Open()
    return $conn
}

# ──────────────────────────────────────────────────────────────
# Private: Set-SqlParameters
# Adds parameters to a SqlCommand, converting $null → DBNull.
# ──────────────────────────────────────────────────────────────

function Set-SqlParameters {
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
# Executes INSERT / UPDATE / DELETE / MERGE.
# Returns number of rows affected.
# ──────────────────────────────────────────────────────────────

function Invoke-SqlNonQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Query,
        [hashtable]$Parameters  = @{},
        [int]$TimeoutSec        = $null
    )

    if (-not $TimeoutSec) { $TimeoutSec = $script:CommandTimeout }

    $conn = $null
    try {
        $conn                   = New-SqlConnection
        $cmd                    = $conn.CreateCommand()
        $cmd.CommandText        = $Query
        $cmd.CommandTimeout     = $TimeoutSec
        Set-SqlParameters -Command $cmd -Parameters $Parameters
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
# Executes a query and returns a single scalar value.
# ──────────────────────────────────────────────────────────────

function Invoke-SqlScalar {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Query,
        [hashtable]$Parameters  = @{},
        [int]$TimeoutSec        = 30
    )

    $conn = $null
    try {
        $conn               = New-SqlConnection
        $cmd                = $conn.CreateCommand()
        $cmd.CommandText    = $Query
        $cmd.CommandTimeout = $TimeoutSec
        Set-SqlParameters -Command $cmd -Parameters $Parameters

        $result = $cmd.ExecuteScalar()
        return if ($null -eq $result -or $result -is [System.DBNull]) { $null } else { $result }
    }
    catch {
        throw "SQL scalar failed: $($_.Exception.Message)"
    }
    finally {
        if ($conn) { $conn.Close(); $conn.Dispose() }
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Invoke-SqlQuery
# Executes a SELECT and returns a DataTable.
# ──────────────────────────────────────────────────────────────

function Invoke-SqlQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Query,
        [hashtable]$Parameters  = @{},
        [int]$TimeoutSec        = $null
    )

    if (-not $TimeoutSec) { $TimeoutSec = $script:CommandTimeout }

    $conn = $null
    try {
        $conn               = New-SqlConnection
        $cmd                = $conn.CreateCommand()
        $cmd.CommandText    = $Query
        $cmd.CommandTimeout = $TimeoutSec
        Set-SqlParameters -Command $cmd -Parameters $Parameters

        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        $ds      = New-Object System.Data.DataSet
        [void]$adapter.Fill($ds)

        return if ($ds.Tables.Count -eq 0) { @() } else { $ds.Tables[0] }
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
# Executes multiple statements within a single connection.
# Each item: @{ Query = "..."; Parameters = @{} }
# Optional transaction wrapping.
# ──────────────────────────────────────────────────────────────

function Invoke-SqlBatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [array]$Statements,
        [int]$TimeoutSec    = $null,
        [switch]$UseTransaction
    )

    if (-not $TimeoutSec) { $TimeoutSec = $script:CommandTimeout }

    $conn = $null; $tx = $null; $total = 0
    try {
        $conn = New-SqlConnection
        if ($UseTransaction) { $tx = $conn.BeginTransaction() }

        foreach ($s in $Statements) {
            $cmd                = $conn.CreateCommand()
            $cmd.CommandText    = $s.Query
            $cmd.CommandTimeout = $TimeoutSec
            if ($tx) { $cmd.Transaction = $tx }
            Set-SqlParameters -Command $cmd -Parameters $s.Parameters
            $total += $cmd.ExecuteNonQuery()
        }

        if ($tx) { $tx.Commit() }
        return $total
    }
    catch {
        if ($tx) { try { $tx.Rollback() } catch {} }
        throw "SQL batch failed: $($_.Exception.Message)"
    }
    finally {
        if ($tx)   { $tx.Dispose() }
        if ($conn) { $conn.Close(); $conn.Dispose() }
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Start-SyncLogEntry
# Inserts a SyncLog row with Status = 'Running'.
# Returns the RunId GUID.
# ──────────────────────────────────────────────────────────────

function Start-SyncLogEntry {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string]$FunctionName)

    $runId = [guid]::NewGuid()

    Invoke-SqlNonQuery -Query @"
INSERT INTO SyncLog (RunId, FunctionName, StartedAt, Status)
VALUES (@RunId, @FunctionName, SYSUTCDATETIME(), 'Running')
"@ -Parameters @{
        '@RunId'        = $runId
        '@FunctionName' = $FunctionName
    } | Out-Null

    return $runId
}

# ──────────────────────────────────────────────────────────────
# Public: Complete-SyncLogEntry
# Updates the SyncLog row at the end of a run.
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
    ErrorCount       = @ErrorCount,
    ErrorMessage     = @ErrorMessage,
    TokenAdvancedTo  = @Token
WHERE RunId = @RunId
"@ -Parameters @{
        '@RunId'       = $RunId
        '@Status'      = $Status
        '@Inserted'    = $UsersInserted
        '@Updated'     = $UsersUpdated
        '@Deleted'     = $UsersSoftDeleted
        '@Processed'   = $UsersProcessed
        '@ErrorCount'  = $ErrorCount
        '@ErrorMessage'= $ErrorMessage
        '@Token'       = $TokenAdvancedTo
    } | Out-Null
}

# ──────────────────────────────────────────────────────────────
# Public: Get-DeltaToken
# Returns the active token value for a named token, or $null.
# ──────────────────────────────────────────────────────────────

function Get-DeltaToken {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string]$TokenName)

    return Invoke-SqlScalar -Query @"
SELECT TokenValue FROM DeltaTokens
WHERE TokenName = @Name AND IsActive = 1
"@ -Parameters @{ '@Name' = $TokenName }
}

# ──────────────────────────────────────────────────────────────
# Public: Save-DeltaToken
# Upserts a delta token (insert or update).
# ──────────────────────────────────────────────────────────────

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
    TokenValue         = source.TokenValue,
    UpdatedAt          = SYSUTCDATETIME(),
    IsActive           = 1,
    DeactivatedAt      = NULL,
    DeactivationReason = NULL
WHEN NOT MATCHED THEN INSERT
    (TokenName, TokenValue, CreatedAt, UpdatedAt, IsActive)
    VALUES (source.TokenName, source.TokenValue, SYSUTCDATETIME(), SYSUTCDATETIME(), 1);
"@ -Parameters @{
        '@Name'  = $TokenName
        '@Value' = $TokenValue
    } | Out-Null
}

# ──────────────────────────────────────────────────────────────
# Public: Disable-DeltaToken
# Marks an active token as inactive with a reason.
# ──────────────────────────────────────────────────────────────

function Disable-DeltaToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$TokenName,
        [string]$Reason = 'Deactivated'
    )

    Invoke-SqlNonQuery -Query @"
UPDATE DeltaTokens SET
    IsActive           = 0,
    DeactivatedAt      = SYSUTCDATETIME(),
    DeactivationReason = @Reason
WHERE TokenName = @Name
"@ -Parameters @{
        '@Name'   = $TokenName
        '@Reason' = $Reason
    } | Out-Null
}

# ──────────────────────────────────────────────────────────────
# Exports
# ──────────────────────────────────────────────────────────────

Export-ModuleMember -Function @(
    'Initialize-SqlContext',
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
