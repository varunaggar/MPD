<#
.SYNOPSIS
    Exchange Online PowerShell helpers — certificate-based authentication.

.DESCRIPTION
    Provides connection management and mailbox data retrieval using
    the ExchangeOnlineManagement V3 module with certificate-based
    service principal authentication.

    The Entra app registration used here is the SAME one used for
    Graph — the certificate is already in the Windows cert store.
    However, the app requires an additional permission for EXO:

      Microsoft APIs → Office 365 Exchange Online
        Application permission: Exchange.ManageAsApp

    AND the service principal must be assigned the Exchange Online
    management role. Run once in Exchange Online PowerShell as admin:

      New-ManagementRoleAssignment
          -App "<your-app-display-name>"
          -Role "Mail Recipients"      (or Exchange Recipient Administrator)

.NOTES
    Dependencies:
      Install-Module ExchangeOnlineManagement -Scope AllUsers -Force
      Minimum version: 3.0

    Config consumed from config.xml:
      Authentication.AppId
      Authentication.CertificateThumbprint
      Authentication.CertificateStoreLocation
      ExchangeOnline.Organisation
      ExchangeOnline.MailboxTypes
#>

# ──────────────────────────────────────────────────────────────
# Module-scoped state
# ──────────────────────────────────────────────────────────────

$script:ExoConfig     = $null
$script:IsConnected   = $false

# Properties fetched for every mailbox call.
# Keeping this list explicit avoids pulling unnecessary data.
$script:MailboxProperties = @(
    'ExchangeGuid',
    'ExternalDirectoryObjectId',    # Entra Object ID → FK to Users table
    'PrimarySmtpAddress',
    'UserPrincipalName',
    'DisplayName',
    'Alias',
    'RecipientTypeDetails',
    'HiddenFromAddressListsEnabled',
    'LitigationHoldEnabled',
    'ArchiveStatus',
    'ForwardingAddress',
    'ForwardingSmtpAddress',
    'GrantSendOnBehalfTo',
    'IsDirSynced',
    'WhenMailboxCreated',
    'WhenChangedUTC'
)

# ──────────────────────────────────────────────────────────────
# Public: Initialize-ExoContext
# Called once per script after loading config.
# ──────────────────────────────────────────────────────────────

function Initialize-ExoContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Config
    )

    $script:ExoConfig   = $Config
    $script:IsConnected = $false

    # Validate required EXO settings
    $org = $Config.ExchangeOnline.Organisation
    if ([string]::IsNullOrWhiteSpace($org) -or $org -like 'REPLACE-*') {
        throw "config.xml ExchangeOnline.Organisation has not been set."
    }

    # Check ExchangeOnlineManagement module is available
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "ExchangeOnlineManagement module is not installed. " +
              "Run: Install-Module ExchangeOnlineManagement -Scope AllUsers -Force"
    }

    Write-LogInfo "EXO context initialised (Organisation=$org)"
}

# ──────────────────────────────────────────────────────────────
# Public: Connect-ExoSession
# Connects to Exchange Online using certificate-based auth.
# Safe to call multiple times — checks IsConnected first.
# ──────────────────────────────────────────────────────────────

function Connect-ExoSession {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Config
    )

    if ($script:IsConnected) {
        return
    }

    Write-LogInfo "Connecting to Exchange Online..."

    $appId      = $Config.Authentication.AppId
    $thumbprint = $Config.Authentication.CertificateThumbprint.Replace(" ", "").Trim()
    $storeLoc   = $Config.Authentication.CertificateStoreLocation
    $org        = $Config.ExchangeOnline.Organisation

    # Verify certificate before attempting connection
    $certPath = "Cert:\$storeLoc\My\$thumbprint"
    if (-not (Test-Path $certPath)) {
        throw "Certificate '$thumbprint' not found in $certPath"
    }

    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop

        Connect-ExchangeOnline `
            -AppId                  $appId `
            -CertificateThumbprint  $thumbprint `
            -Organization           $org `
            -ShowBanner:            $false `
            -ErrorAction            Stop

        $script:IsConnected = $true
        Write-LogInfo "Successfully connected to Exchange Online session (Org: $org)"
    }
    catch {
        $script:IsConnected = $false
        $errMsg = "Failed to connect to Exchange Online: $($_.Exception.Message)"
        Write-LogError $errMsg -ErrorRecord $_
        throw $errMsg
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Disconnect-ExoSession
# Cleanly disconnects the EXO session.
# Always call this at the end of each script (in a finally block).
# ──────────────────────────────────────────────────────────────

function Disconnect-ExoSession {
    [CmdletBinding()]
    param()

    if (-not $script:IsConnected) { return }

    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-LogInfo "Disconnected from Exchange Online session."
    }
    catch {
        Write-LogWarning "EXO disconnect encountered an error: $($_.Exception.Message)"
    }
    finally {
        $script:IsConnected = $false
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Get-AllExoMailboxes
# Full retrieval — all mailboxes of the configured types.
# Used by the baseline load process.
# Returns an array of normalised mailbox objects.
# ──────────────────────────────────────────────────────────────

function Get-AllExoMailboxes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Config
    )

    $types  = $Config.ExchangeOnline.MailboxTypes -split ',' | ForEach-Object { $_.Trim() }
    Write-LogInfo "Fetching all mailboxes of types: $($types -join ', ')..."

    try {
        $raw = Get-EXOMailbox `
            -ResultSize Unlimited `
            -Properties $script:MailboxProperties `
            -ErrorAction Stop |
            Where-Object { $_.RecipientTypeDetails -in $types }

        Write-LogInfo "EXO returned $($raw.Count) mailboxes"
        return $raw | ForEach-Object { ConvertTo-MailboxObject $_ }
    }
    catch {
        throw "Get-AllExoMailboxes failed: $($_.Exception.Message)"
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Get-ChangedExoMailboxes
# Retrieval of mailboxes modified since a given UTC timestamp.
# Used by the delta sync process.
# Includes both live and soft-deleted mailboxes changed in the window.
# Returns an array of normalised mailbox objects.
# ──────────────────────────────────────────────────────────────

function Get-ChangedExoMailboxes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Config,
        [Parameter(Mandatory)] [datetime]$SinceUtc
    )

    $types       = $Config.ExchangeOnline.MailboxTypes -split ',' | ForEach-Object { $_.Trim() }
    $filterValue = $SinceUtc.ToString('MM/dd/yyyy HH:mm:ss')
    $filter      = "WhenChangedUTC -ge '$filterValue'"

    Write-LogInfo "Fetching changed mailboxes since $filterValue UTC..."

    try {
        # ── Active mailboxes that changed ─────────────────────
        $active = Get-EXOMailbox `
            -Filter      $filter `
            -ResultSize  Unlimited `
            -Properties  $script:MailboxProperties `
            -ErrorAction Stop |
            Where-Object { $_.RecipientTypeDetails -in $types } |
            ForEach-Object {
                $obj = ConvertTo-MailboxObject $_
                $obj.IsDeleted = $false
                $obj
            }

        # ── Soft-deleted mailboxes that changed ───────────────
        # Exchange soft-deletes mailboxes for 30 days before hard deletion.
        # Checking this window here means the delta sync catches removals
        # without waiting for the weekly baseline reconciliation.
        $softDeleted = Get-EXOMailbox `
            -SoftDeletedMailbox `
            -Filter      $filter `
            -ResultSize  Unlimited `
            -Properties  $script:MailboxProperties `
            -ErrorAction SilentlyContinue |
            Where-Object { $_ -ne $null -and $_.RecipientTypeDetails -in $types } |
            ForEach-Object {
                $obj = ConvertTo-MailboxObject $_
                $obj.IsDeleted = $true
                $obj
            }

        $all = @($active) + @($softDeleted)
        Write-LogInfo "Delta found $($active.Count) changed + $($softDeleted.Count) soft-deleted mailboxes"
        return $all
    }
    catch {
        throw "Get-ChangedExoMailboxes failed: $($_.Exception.Message)"
    }
}

# ──────────────────────────────────────────────────────────────
# Private: ConvertTo-MailboxObject
# Normalises a raw EXO mailbox PSObject into a clean hashtable
# with consistent types — makes the SQL MERGE parameters simpler.
# ──────────────────────────────────────────────────────────────

function ConvertTo-MailboxObject {
    param($Mailbox)

    # Flatten GrantSendOnBehalfTo multi-value to semicolon string
    $grantSoB = if ($Mailbox.GrantSendOnBehalfTo -and $Mailbox.GrantSendOnBehalfTo.Count -gt 0) {
        ($Mailbox.GrantSendOnBehalfTo | ForEach-Object { $_.ToString() }) -join ';'
    } else { $null }

    # Normalise RecipientTypeDetails to a simple MailboxType label
    $mailboxType = switch ($Mailbox.RecipientTypeDetails) {
        'UserMailbox'      { 'User' }
        'SharedMailbox'    { 'Shared' }
        'RoomMailbox'      { 'Room' }
        'EquipmentMailbox' { 'Equipment' }
        default            { $Mailbox.RecipientTypeDetails }
    }

    # ExternalDirectoryObjectId is the Entra Object ID → FK to Users.UserId
    $userId = if (-not [string]::IsNullOrWhiteSpace($Mailbox.ExternalDirectoryObjectId)) {
        try { [guid]$Mailbox.ExternalDirectoryObjectId } catch { $null }
    } else { $null }

    return [PSCustomObject]@{
        ExchangeGuid              = [guid]$Mailbox.ExchangeGuid
        UserId                    = $userId
        PrimarySmtpAddress        = $Mailbox.PrimarySmtpAddress
        UserPrincipalName         = $Mailbox.UserPrincipalName
        DisplayName               = $Mailbox.DisplayName
        Alias                     = $Mailbox.Alias
        RecipientTypeDetails      = $Mailbox.RecipientTypeDetails
        MailboxType               = $mailboxType
        HiddenFromAddressLists    = if ($null -ne $Mailbox.HiddenFromAddressListsEnabled) { [bool]$Mailbox.HiddenFromAddressListsEnabled } else { $null }
        LitigationHoldEnabled     = if ($null -ne $Mailbox.LitigationHoldEnabled)         { [bool]$Mailbox.LitigationHoldEnabled }         else { $null }
        ArchiveStatus             = if ($null -ne $Mailbox.ArchiveStatus) { $Mailbox.ArchiveStatus.ToString() } else { $null }
        ForwardingAddress         = $Mailbox.ForwardingAddress
        ForwardingSmtpAddress     = $Mailbox.ForwardingSmtpAddress
        GrantSendOnBehalfTo       = $grantSoB
        IsDirSynced               = if ($null -ne $Mailbox.IsDirSynced) { [bool]$Mailbox.IsDirSynced } else { $null }
        WhenMailboxCreated        = $Mailbox.WhenMailboxCreated
        WhenChangedUTC            = $Mailbox.WhenChangedUTC
        IsDeleted                 = $false   # overridden by caller if soft-deleted
    }
}

# ──────────────────────────────────────────────────────────────
# Exports
# ──────────────────────────────────────────────────────────────

Export-ModuleMember -Function @(
    'Initialize-ExoContext',
    'Connect-ExoSession',
    'Disconnect-ExoSession',
    'Get-AllExoMailboxes',
    'Get-ChangedExoMailboxes'
)
