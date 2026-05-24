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

        if ($null -eq $cfg) {
            throw "Invalid configuration: Root element <M365PermsSyncConfig> not found in '$Path'."
        }

        # Validate mandatory placeholders have been replaced
        $requiredFields = @(
            @{ Path = "Authentication.TenantId";            Value = $cfg.Authentication.TenantId }
            @{ Path = "Authentication.AppId";               Value = $cfg.Authentication.AppId }
            @{ Path = "Authentication.CertificateThumbprint"; Value = $cfg.Authentication.CertificateThumbprint }
            @{ Path = "Database.Server";                    Value = $cfg.Database.Server }
            @{ Path = "Database.Name";                      Value = $cfg.Database.Name }
            #@{ Path = "Logging.Directory";                  Value = $cfg.Logging.Directory }
        )

        foreach ($field in $requiredFields) {
            if ([string]::IsNullOrWhiteSpace($field.Value) -or $field.Value -match "^REPLACE-") {
                throw "Configuration error: The field '$($field.Path)' is missing or contains a placeholder. Please update config.xml."
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

    # Sanitize thumbprint (remove spaces/newlines often introduced by copy-paste)
    $thumbprint = $Config.Authentication.CertificateThumbprint.Replace(" ", "").Trim()
    
    $appId      = $Config.Authentication.AppId
    $tenantId   = $Config.Authentication.TenantId
    $storeLoc   = $Config.Authentication.CertificateStoreLocation
    $storeName  = $Config.Authentication.CertificateStoreName

    Write-LogInfo "Establishing Azure connection (Tenant: $tenantId, AppId: $appId)..."

    # Verify certificate is present in the store before attempting auth
    $certPath = "Cert:\$storeLoc\$storeName\$thumbprint"
    if (-not (Test-Path $certPath)) {
        $errMsg = "Authentication Certificate not found in store: $certPath"
        Write-LogError $errMsg
        throw $errMsg
    }

    try {
        # Connect-AzAccount requires the Az.Accounts module
        Connect-AzAccount `
            -ServicePrincipal `
            -ApplicationId $appId `
            -Tenant $tenantId `
            -CertificateThumbprint $thumbprint `
            -CertificateStoreLocation $storeLoc `
            -ErrorAction Stop | Out-Null

        Write-LogInfo "Successfully authenticated to Azure as service principal."
    }
    catch {
        $errMsg = "Service Principal authentication failed: $($_.Exception.Message)"
        Write-LogError $errMsg -ErrorRecord $_
        throw $errMsg
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
