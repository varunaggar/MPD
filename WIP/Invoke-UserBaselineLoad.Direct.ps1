<#
.SYNOPSIS
    Direct baseline load test script without dependency helper functions.

.DESCRIPTION
    This script imports required local PowerShell modules directly from the
    workspace Modules folder, authenticates using the configured auth method,
    and validates Azure SQL access by inserting and reading back a test row
    in the DeltaTokens table.

    It avoids using the shared DependencyHelpers module or any helper module
    loader functions.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.xml")
)

$ErrorActionPreference = 'Stop'

$scriptName = 'Invoke-UserBaselineLoad.Direct'
$workspaceRoot = $PSScriptRoot
$modulesRoot   = Join-Path $workspaceRoot 'Modules'

function Get-LocalModuleManifestPath {
    param(
        [Parameter(Mandatory)] [string]$Root,
        [Parameter(Mandatory)] [string]$ModuleNamePattern
    )
    $moduleFolder = Get-ChildItem -Path $Root -Directory -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -like $ModuleNamePattern } |
                    Sort-Object Name -Descending | Select-Object -First 1
    if (-not $moduleFolder) { return $null }

    $manifest = Get-ChildItem -Path $moduleFolder.FullName -Filter '*.psd1' -Recurse -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '\.psd1$' } |
                Sort-Object FullName | Select-Object -First 1
    if ($manifest) { return $manifest.FullName }
    return $null
}

function Import-LocalModule {
    param(
        [Parameter(Mandatory)] [string]$ModuleNamePattern
    )
    $manifestPath = Get-LocalModuleManifestPath -Root $modulesRoot -ModuleNamePattern $ModuleNamePattern
    if (-not $manifestPath) {
        throw "Could not find local module matching '$ModuleNamePattern' in $modulesRoot"
    }
    Import-Module $manifestPath -Force -ErrorAction Stop
    Write-Host "Imported local module from $manifestPath"
}

function Read-Config {
    param([string]$Path)
    if (-not (Test-Path $Path)) { throw "Config file not found: $Path" }
    [xml]$xml = Get-Content -Path $Path -Raw
    return $xml.M365PermsSyncConfig
}

function Get-ConfigValue {
    param(
        [Parameter(Mandatory)] [object]$Config,
        [Parameter(Mandatory)] [string]$Section,
        [Parameter(Mandatory)] [string]$Key
    )
    try {
        $value = $Config.$Section.$Key
        if ([string]::IsNullOrWhiteSpace($value)) { return $null }
        return $value
    }
    catch {
        return $null
    }
}

function Convert-SecureTokenToPlainText {
    param([System.Security.SecureString]$SecureToken)
    if (-not $SecureToken) { return $null }
    $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureToken)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringUni($ptr)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    }
}

function Authenticate-Azure {
    param([object]$Config)

    $authType              = Get-ConfigValue -Config $Config -Section Authentication -Key AuthType
    $appId                 = Get-ConfigValue -Config $Config -Section Authentication -Key AppId
    $tenantId              = Get-ConfigValue -Config $Config -Section Authentication -Key TenantId
    $clientSecret          = Get-ConfigValue -Config $Config -Section Authentication -Key ClientSecret
    $thumbprint            = Get-ConfigValue -Config $Config -Section Authentication -Key CertificateThumbprint
    $managedIdentityClientId = Get-ConfigValue -Config $Config -Section Authentication -Key ManagedIdentityClientId
    $userPrincipalName     = Get-ConfigValue -Config $Config -Section Authentication -Key UserPrincipalName

    Write-Host "Authenticating to Azure (AuthType=$authType, Tenant=$tenantId, AppId=$appId)"

    switch ($authType) {
        'Certificate' {
            if (-not $thumbprint) { throw 'CertificateThumbprint is required for Certificate auth.' }
            $thumbprint = $thumbprint.Replace(' ', '').Trim()
            $cert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $thumbprint }
            if (-not $cert) { throw "Certificate with thumbprint $thumbprint not found in LocalMachine\\My." }
            Connect-AzAccount -ServicePrincipal -ApplicationId $appId -Tenant $tenantId -CertificateThumbprint $thumbprint -ErrorAction Stop | Out-Null
        }
        'Secret' {
            if (-not $clientSecret) { throw 'ClientSecret is required for Secret auth.' }
            $secureSecret = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($appId, $secureSecret)
            Connect-AzAccount -ServicePrincipal -Tenant $tenantId -Credential $credential -ErrorAction Stop | Out-Null
        }
        'ManagedIdentity' {
            if ([string]::IsNullOrWhiteSpace($managedIdentityClientId)) {
                Connect-AzAccount -Identity -ErrorAction Stop | Out-Null
            }
            else {
                Connect-AzAccount -Identity -AccountId $managedIdentityClientId -ErrorAction Stop | Out-Null
            }
        }
        'User' {
            if ([string]::IsNullOrWhiteSpace($userPrincipalName)) {
                Connect-AzAccount -Tenant $tenantId -ErrorAction Stop | Out-Null
            }
            else {
                Write-Host 'Using device/browser login because Connect-AzAccount in this Az.Accounts version does not accept -Username.'
                Connect-AzAccount -Tenant $tenantId -UseDeviceAuthentication -ErrorAction Stop | Out-Null
            }
        }
        default {
            throw "Unsupported AuthType: $authType"
        }
    }

    Write-Host 'Azure authentication completed.'
}

function Get-SqlAccessToken {
    $tokenResponse = Get-AzAccessToken -ResourceUrl 'https://database.windows.net/' -ErrorAction Stop
    return Convert-SecureTokenToPlainText -SecureToken $tokenResponse.Token
}

function New-SqlConnection {
    param([string]$Server, [string]$Database, [string]$AccessToken)
    $conn = New-Object System.Data.SqlClient.SqlConnection
    $conn.ConnectionString = "Server=tcp:$Server,1433;Database=$Database;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"
    $conn.AccessToken = $AccessToken
    return $conn
}

function Execute-SqlNonQuery {
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string]$Query
    )
    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = $Query
    $cmd.CommandTimeout = 120
    return $cmd.ExecuteNonQuery()
}

function Execute-SqlQuery {
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string]$Query
    )
    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = $Query
    $cmd.CommandTimeout = 120

    $reader = $cmd.ExecuteReader()
    $rows = @()
    while ($reader.Read()) {
        $row = @{}
        for ($i = 0; $i -lt $reader.FieldCount; $i++) {
            $row[$reader.GetName($i)] = $reader.GetValue($i)
        }
        $rows += [pscustomobject]$row
    }
    $reader.Close()
    return $rows
}

try {
    Write-Host "Loading local Az modules from $modulesRoot"
    Import-LocalModule -ModuleNamePattern 'az.accounts*'
    Import-LocalModule -ModuleNamePattern 'exchangeonlinemanagement*'

    $config = Read-Config -Path $ConfigPath
    if (-not $config) { throw "Configuration not loaded from $ConfigPath" }

    Authenticate-Azure -Config $config

    $sqlServer = Get-ConfigValue -Config $config -Section Database -Key Server
    $sqlDatabase = Get-ConfigValue -Config $config -Section Database -Key Name
    if (-not $sqlServer -or -not $sqlDatabase) { throw 'Database.Server and Database.Name must be set in config.xml' }

    $accessToken = Get-SqlAccessToken
    if (-not $accessToken) { throw 'Failed to acquire SQL access token.' }

    Write-Host "Acquired SQL token. Creating connection to $sqlServer/$sqlDatabase"
    $connection = New-SqlConnection -Server $sqlServer -Database $sqlDatabase -AccessToken $accessToken
    $connection.Open()

    $testTokenName = "InvokeUserBaselineLoadTest_$([DateTime]::UtcNow.ToString('yyyyMMddHHmmss'))"
    $insertSql = @"
INSERT INTO DeltaTokens (TokenName, TokenValue, CreatedAt, UpdatedAt, IsActive, DeactivatedAt, DeactivationReason)
VALUES ('$testTokenName', 'ok', SYSUTCDATETIME(), SYSUTCDATETIME(), 1, NULL, NULL);
"@
    Execute-SqlNonQuery -Connection $connection -Query $insertSql | Out-Null

    $selectSql = "SELECT TOP 1 TokenName, TokenValue, CreatedAt, UpdatedAt, IsActive, DeactivatedAt, DeactivationReason FROM DeltaTokens WHERE TokenName = '$testTokenName' ORDER BY CreatedAt DESC"
    $rows = Execute-SqlQuery -Connection $connection -Query $selectSql
    $connection.Close()

    if ($rows.Count -eq 0) {
        throw 'No test row returned from DeltaTokens.'
    }

    Write-Host "SQL insert/read validation succeeded. Retrieved row:"
    $row = $rows[0]
    foreach ($property in $row.PSObject.Properties) {
        Write-Host "  $($property.Name): $($property.Value)"
    }
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.Exception.InnerException) { Write-Host "INNER: $($_.Exception.InnerException.Message)" -ForegroundColor Red }
    exit 1
}
