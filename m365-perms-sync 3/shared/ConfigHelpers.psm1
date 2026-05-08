<#
.SYNOPSIS
    Configuration loader and Azure authentication helpers.

.DESCRIPTION
    Reads config.xml and exposes it as a typed object.
    Handles one-time Connect-AzAccount using the certificate
    stored in the Windows certificate store.

.NOTES
    Every script imports this module first, loads config,
    then calls Connect-SyncServicePrincipal before doing
    any Graph or SQL work.

    Dependencies: Az.Accounts (must be installed on the server)
    Install: Install-Module Az.Accounts -Scope AllUsers -Force
#>

# ──────────────────────────────────────────────────────────────
# Public: Import-SyncConfig
# Reads config.xml and returns the configuration object.
# ──────────────────────────────────────────────────────────────

function Import-SyncConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string]$Path
    )

    try {
        [xml]$raw = Get-Content -Path $Path -Encoding UTF8 -ErrorAction Stop
        $cfg = $raw.M365PermsSyncConfig

        # Validate mandatory placeholders have been replaced
        $checks = @{
            'Authentication.TenantId'            = $cfg.Authentication.TenantId
            'Authentication.AppId'               = $cfg.Authentication.AppId
            'Authentication.CertificateThumbprint' = $cfg.Authentication.CertificateThumbprint
            'Database.Server'                    = $cfg.Database.Server
            'Database.Name'                      = $cfg.Database.Name
            'Logging.Directory'                  = $cfg.Logging.Directory
        }

        foreach ($field in $checks.GetEnumerator()) {
            if ([string]::IsNullOrWhiteSpace($field.Value) -or $field.Value -like 'REPLACE-*') {
                throw "config.xml: '$($field.Key)' has not been set. Replace the placeholder before running."
            }
        }

        return $cfg
    }
    catch {
        throw "Failed to load configuration from '$Path': $($_.Exception.Message)"
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Connect-SyncServicePrincipal
# Authenticates to Azure using the app registration certificate.
# Called once per script; Get-AzAccessToken then works for
# both Graph and SQL tokens within the same session.
# ──────────────────────────────────────────────────────────────

function Connect-SyncServicePrincipal {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Config
    )

    $thumbprint = $Config.Authentication.CertificateThumbprint
    $appId      = $Config.Authentication.AppId
    $tenantId   = $Config.Authentication.TenantId
    $storeLoc   = $Config.Authentication.CertificateStoreLocation
    $storeName  = $Config.Authentication.CertificateStoreName

    # Verify certificate is present in the store before attempting auth
    $certPath = "Cert:\$storeLoc\$storeName\$thumbprint"
    if (-not (Test-Path $certPath)) {
        throw "Certificate with thumbprint '$thumbprint' not found in $certPath. " +
              "Ensure the certificate is installed on this server."
    }

    try {
        Connect-AzAccount `
            -ServicePrincipal `
            -ApplicationId $appId `
            -Tenant $tenantId `
            -CertificateThumbprint $thumbprint `
            -CertificateStoreLocation $storeLoc `
            -ErrorAction Stop | Out-Null

        Write-Verbose "Connected to Azure as service principal (AppId=$appId)"
    }
    catch {
        throw "Failed to connect to Azure as service principal: $($_.Exception.Message)"
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Get-ConfigValue
# Safe helper to read a string value from the config with a default.
# ──────────────────────────────────────────────────────────────

function Get-ConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Config,
        [Parameter(Mandatory)] [string]$Section,
        [Parameter(Mandatory)] [string]$Key,
        [string]$Default = $null
    )

    try {
        $value = $Config.$Section.$Key
        if ([string]::IsNullOrWhiteSpace($value)) { return $Default }
        return $value
    }
    catch {
        return $Default
    }
}

# ──────────────────────────────────────────────────────────────
# Exports
# ──────────────────────────────────────────────────────────────

Export-ModuleMember -Function @(
    'Import-SyncConfig',
    'Connect-SyncServicePrincipal',
    'Get-ConfigValue'
)
