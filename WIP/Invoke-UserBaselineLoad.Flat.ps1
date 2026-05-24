<#
.SYNOPSIS
    Flat baseline load validation script.

.DESCRIPTION
    This script uses only ConfigHelpers and LoggingHelpers modules.
    All other logic is written inline in simple sequential PowerShell.
    It authenticates with Azure, acquires a SQL access token, and
    executes a simple SELECT 1 against the configured Azure SQL DB.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.xml")
)

$ErrorActionPreference = 'Stop'
$scriptName = 'Invoke-UserBaselineLoad.Flat'
$sharedPath = Join-Path $PSScriptRoot 'shared'
$modulesRoot = Join-Path $PSScriptRoot 'Modules'

Import-Module (Join-Path $sharedPath 'ConfigHelpers.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $sharedPath 'LoggingHelpers.psm1') -Force -ErrorAction Stop

$config = Import-SyncConfig -Path $ConfigPath
Initialize-Logging -Config $config -ProcessName $scriptName

try {
    Write-LogSection 'Bootstrap'
    Write-LogInfo "Config loaded from: $ConfigPath"
    Write-LogInfo "Azure AppId: $($config.Authentication.AppId)"
    Write-LogInfo "SQL Server: $($config.Database.Server)"
    Write-LogInfo "SQL DB: $($config.Database.Name)"

    Write-LogSection 'Az Module Import'
    $azManifest = Get-ChildItem -Path $modulesRoot -Recurse -Filter 'Az.Accounts.psd1' -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $azManifest) {
        throw "Could not locate Az.Accounts.psd1 under $modulesRoot"
    }
    Import-Module $azManifest.FullName -Force -ErrorAction Stop
    Write-LogInfo "Imported Az.Accounts from $($azManifest.FullName)"

    if (-not (Get-Command Connect-AzAccount -ErrorAction SilentlyContinue)) {
        throw 'Connect-AzAccount is not available after importing Az.Accounts.'
    }

    Write-LogSection 'Authentication'
    $authType = $config.Authentication.AuthType
    Write-LogInfo "AuthType=$authType"

    switch ($authType) {
        'Certificate' {
            $thumbprint = $config.Authentication.CertificateThumbprint.Replace(' ', '').Trim()
            Connect-AzAccount -ServicePrincipal -ApplicationId $config.Authentication.AppId -Tenant $config.Authentication.TenantId -CertificateThumbprint $thumbprint -ErrorAction Stop | Out-Null
        }
        'Secret' {
            $secureSecret = ConvertTo-SecureString -String $config.Authentication.ClientSecret -AsPlainText -Force
            $creds = New-Object System.Management.Automation.PSCredential($config.Authentication.AppId, $secureSecret)
            Connect-AzAccount -ServicePrincipal -Tenant $config.Authentication.TenantId -Credential $creds -ErrorAction Stop | Out-Null
        }
        'ManagedIdentity' {
            if ([string]::IsNullOrWhiteSpace($config.Authentication.ManagedIdentityClientId)) {
                Connect-AzAccount -Identity -ErrorAction Stop | Out-Null
            } else {
                Connect-AzAccount -Identity -AccountId $config.Authentication.ManagedIdentityClientId -ErrorAction Stop | Out-Null
            }
        }
        'User' {
            if ([string]::IsNullOrWhiteSpace($config.Authentication.UserPrincipalName)) {
                Connect-AzAccount -Tenant $config.Authentication.TenantId -ErrorAction Stop | Out-Null
            } else {
                Connect-AzAccount -Tenant $config.Authentication.TenantId -UseDeviceAuthentication -ErrorAction Stop | Out-Null
            }
        }
        default {
            throw "Unsupported AuthType: $authType"
        }
    }

    Write-LogInfo 'Azure authentication completed.'

    Write-LogSection 'SQL Token'
    $tokenResponse = Get-AzAccessToken -ResourceUrl 'https://database.windows.net/' -ErrorAction Stop
    if ($null -eq $tokenResponse -or [string]::IsNullOrWhiteSpace($tokenResponse.Token)) {
        throw 'Failed to acquire SQL access token.'
    }

    $token = if ($tokenResponse.Token -is [System.Security.SecureString]) {
        $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenResponse.Token)
        try { [Runtime.InteropServices.Marshal]::PtrToStringUni($ptr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) }
    } else {
        $tokenResponse.Token
    }

    Write-LogInfo "SQL token acquired. Expires: $($tokenResponse.ExpiresOn)"

    Write-LogSection 'SQL Connection'
    $sqlServer = $config.Database.Server
    $sqlDb = $config.Database.Name
    $connStr = "Server=tcp:$sqlServer,1433;Database=$sqlDb;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"
    $connection = New-Object System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = $connStr
    $connection.AccessToken      = $token

    Write-LogInfo "Opening SQL connection to $sqlServer/$sqlDb"
    $connection.Open()

    $command = $connection.CreateCommand()
    $command.CommandText = 'SELECT 1'
    $command.CommandTimeout = 30
    $result = $command.ExecuteScalar()
    $connection.Close()

    Write-LogInfo "SQL connection succeeded. Query result: $result"
    Close-Logging -Status 'Success'
}
catch {
    Write-LogError "Flat script failed: $($_.Exception.Message)" -ErrorRecord $_
    Close-Logging -Status 'Failed'
    exit 1
}
