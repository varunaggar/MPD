<#
.SYNOPSIS
    Reusable Graph API helpers for the M365 permissions sync solution.

.DESCRIPTION
    Imported by every Azure Function via profile.ps1.
    Provides token acquisition, paged requests, and delta query support
    on top of Invoke-RestMethod. Handles throttling, auth expiry,
    and delta token expiry automatically.

.NOTES
    Dependencies: Az.Accounts, Az.KeyVault (loaded via requirements.psd1)
    Assumes Managed Identity authentication has already occurred
    (Connect-AzAccount -Identity) in profile.ps1.

    Environment variables required:
      KEY_VAULT_NAME — name of the Key Vault holding the three secrets:
                       TenantId, GraphAppId, GraphAppSecret
#>

# ──────────────────────────────────────────────────────────────
# Module-scoped state
# ──────────────────────────────────────────────────────────────

# Token cache — avoid reacquiring Graph tokens on every call.
# Access tokens typically last ~1 hour. We refresh at 50 min.
$script:GraphTokenCache = @{
    Token     = $null
    ExpiresAt = [datetime]::MinValue
}

# Key Vault name — set once per function app via environment variable
$script:KeyVaultName = $env:KEY_VAULT_NAME

if (-not $script:KeyVaultName) {
    throw "GraphHelpers: KEY_VAULT_NAME environment variable must be set."
}

# ──────────────────────────────────────────────────────────────
# Private: Get-GraphCredentialsFromKeyVault
# Retrieves tenant ID, Graph app ID, and Graph app secret from KV
# ──────────────────────────────────────────────────────────────

function Get-GraphCredentialsFromKeyVault {
    try {
        $tenantId    = (Get-AzKeyVaultSecret -VaultName $script:KeyVaultName -Name "TenantId"       -AsPlainText)
        $graphAppId  = (Get-AzKeyVaultSecret -VaultName $script:KeyVaultName -Name "GraphAppId"     -AsPlainText)
        $graphSecret = (Get-AzKeyVaultSecret -VaultName $script:KeyVaultName -Name "GraphAppSecret" -AsPlainText)

        if (-not $tenantId -or -not $graphAppId -or -not $graphSecret) {
            throw "One or more required Key Vault secrets are missing."
        }

        return @{
            TenantId = $tenantId
            AppId    = $graphAppId
            Secret   = $graphSecret
        }
    }
    catch {
        throw "Failed to retrieve Graph credentials from Key Vault '$($script:KeyVaultName)': $($_.Exception.Message)"
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Get-GraphToken
# Acquires (or returns cached) Graph access token
# ──────────────────────────────────────────────────────────────

function Get-GraphToken {
    [CmdletBinding()]
    param(
        [switch]$ForceRefresh
    )

    $now = [datetime]::UtcNow

    # Return cached token if valid and refresh not forced
    if (-not $ForceRefresh -and
        $script:GraphTokenCache.Token -and
        $script:GraphTokenCache.ExpiresAt -gt $now.AddMinutes(2)) {
        Write-Verbose "Returning cached Graph token (expires at $($script:GraphTokenCache.ExpiresAt))"
        return $script:GraphTokenCache.Token
    }

    Write-Verbose "Acquiring new Graph token"

    $creds = Get-GraphCredentialsFromKeyVault

    $tokenBody = @{
        grant_type    = "client_credentials"
        client_id     = $creds.AppId
        client_secret = $creds.Secret
        scope         = "https://graph.microsoft.com/.default"
    }

    try {
        $response = Invoke-RestMethod `
            -Uri    "https://login.microsoftonline.com/$($creds.TenantId)/oauth2/v2.0/token" `
            -Method POST `
            -Body   $tokenBody `
            -ErrorAction Stop

        # Cache with safety margin — expire 10 min before actual expiry
        $script:GraphTokenCache.Token     = $response.access_token
        $script:GraphTokenCache.ExpiresAt = $now.AddSeconds($response.expires_in - 600)

        Write-Verbose "Graph token acquired. Valid until $($script:GraphTokenCache.ExpiresAt) UTC"
        return $response.access_token
    }
    catch {
        throw "Failed to acquire Graph token: $($_.Exception.Message)"
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Invoke-GraphRequest
# Single Graph REST call with 429 retry and 401 auto-refresh
# ──────────────────────────────────────────────────────────────

function Invoke-GraphRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Uri,

        [ValidateSet("GET","POST","PATCH","PUT","DELETE")]
        [string]$Method = "GET",

        [object]$Body = $null,

        [int]$MaxRetries = 5,

        [int]$TimeoutSec = 100
    )

    $attempt = 0
    $lastError = $null

    while ($attempt -lt $MaxRetries) {
        $attempt++

        try {
            $token = Get-GraphToken
            $headers = @{
                Authorization    = "Bearer $token"
                "Content-Type"   = "application/json"
                ConsistencyLevel = "eventual"
            }

            $params = @{
                Uri         = $Uri
                Method      = $Method
                Headers     = $headers
                TimeoutSec  = $TimeoutSec
                ErrorAction = "Stop"
            }

            if ($Body -and $Method -in @("POST","PATCH","PUT")) {
                if ($Body -isnot [string]) {
                    $Body = $Body | ConvertTo-Json -Depth 10
                }
                $params.Body = $Body
            }

            return Invoke-RestMethod @params
        }
        catch {
            $statusCode = $null
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}

            # 429 — throttled. Respect Retry-After header, exponential fallback
            if ($statusCode -eq 429) {
                $retryAfter = 30
                try {
                    $headerValue = $_.Exception.Response.Headers["Retry-After"]
                    if ($headerValue) { $retryAfter = [int]$headerValue }
                } catch {}

                # Cap backoff at 5 minutes to avoid Function timeout
                $retryAfter = [Math]::Min($retryAfter, 300)

                Write-Warning "Graph throttled (429) on attempt $attempt. Waiting $retryAfter seconds."
                Start-Sleep -Seconds $retryAfter
                $lastError = $_
                continue
            }

            # 401 — token expired or rejected. Force refresh and retry once
            if ($statusCode -eq 401 -and $attempt -eq 1) {
                Write-Warning "Graph returned 401 on attempt $attempt. Refreshing token."
                Get-GraphToken -ForceRefresh | Out-Null
                $lastError = $_
                continue
            }

            # 5xx — transient server error. Exponential backoff
            if ($statusCode -ge 500 -and $statusCode -le 599) {
                $wait = [Math]::Min([Math]::Pow(2, $attempt), 60)
                Write-Warning "Graph returned $statusCode on attempt $attempt. Waiting $wait seconds."
                Start-Sleep -Seconds $wait
                $lastError = $_
                continue
            }

            # 410 Gone — specifically meaningful for delta queries
            # Re-throw immediately so caller can handle re-initialisation
            if ($statusCode -eq 410) { throw }

            # All other errors — do not retry
            throw
        }
    }

    throw "Graph request failed after $MaxRetries attempts. Last error: $($lastError.Exception.Message)"
}

# ──────────────────────────────────────────────────────────────
# Public: Invoke-GraphPagedRequest
# Automatically follows @odata.nextLink, returns all values
# ──────────────────────────────────────────────────────────────

function Invoke-GraphPagedRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Uri
    )

    $allObjects = [System.Collections.Generic.List[object]]::new()
    $currentUri = $Uri
    $pageCount = 0

    do {
        $pageCount++
        Write-Verbose "Fetching page $pageCount"

        $response = Invoke-GraphRequest -Uri $currentUri -Method GET

        if ($response.value) {
            $allObjects.AddRange([object[]]$response.value)
        }

        $currentUri = $response.'@odata.nextLink'

    } while ($currentUri)

    Write-Verbose "Paged request complete: $pageCount pages, $($allObjects.Count) objects"
    return $allObjects.ToArray()
}

# ──────────────────────────────────────────────────────────────
# Public: Invoke-GraphDeltaQuery
# Handles delta pattern: paging + deltaLink capture + 410 Gone
# ──────────────────────────────────────────────────────────────

function Invoke-GraphDeltaQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Uri
    )

    $allObjects = [System.Collections.Generic.List[object]]::new()
    $currentUri = $Uri
    $deltaLink = $null
    $pageCount = 0

    try {
        do {
            $pageCount++
            Write-Verbose "Fetching delta page $pageCount"

            $response = Invoke-GraphRequest -Uri $currentUri -Method GET

            if ($response.value) {
                $allObjects.AddRange([object[]]$response.value)
            }

            # Delta query termination: deltaLink means "no more pages"
            if ($response.'@odata.deltaLink') {
                $deltaLink = $response.'@odata.deltaLink'
                $currentUri = $null
            }
            elseif ($response.'@odata.nextLink') {
                $currentUri = $response.'@odata.nextLink'
            }
            else {
                $currentUri = $null
            }

        } while ($currentUri)

        if (-not $deltaLink) {
            throw "Delta query completed without returning @odata.deltaLink"
        }

        Write-Verbose "Delta query complete: $pageCount pages, $($allObjects.Count) changed objects"

        return @{
            Objects      = $allObjects.ToArray()
            DeltaLink    = $deltaLink
            DeltaToken   = Get-DeltaTokenFromUrl -DeltaLink $deltaLink
            TokenExpired = $false
        }
    }
    catch {
        # 410 Gone = delta token has expired. Signal caller to re-initialise
        $statusCode = $null
        try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}

        if ($statusCode -eq 410) {
            Write-Warning "Delta token expired (410 Gone). Re-initialisation required."
            return @{
                Objects      = @()
                DeltaLink    = $null
                DeltaToken   = $null
                TokenExpired = $true
            }
        }

        throw
    }
}

# ──────────────────────────────────────────────────────────────
# Private: Get-DeltaTokenFromUrl
# Extracts the $deltatoken value from a deltaLink URL
# ──────────────────────────────────────────────────────────────

function Get-DeltaTokenFromUrl {
    param([string]$DeltaLink)

    if ([string]::IsNullOrEmpty($DeltaLink)) { return $null }

    if ($DeltaLink -match '\$deltatoken=([^&]+)') {
        return [uri]::UnescapeDataString($Matches[1])
    }

    # Some Graph endpoints return $skiptoken instead of $deltatoken
    if ($DeltaLink -match '\$skiptoken=([^&]+)') {
        return [uri]::UnescapeDataString($Matches[1])
    }

    return $null
}

# ──────────────────────────────────────────────────────────────
# Exports
# ──────────────────────────────────────────────────────────────

Export-ModuleMember -Function @(
    'Get-GraphToken',
    'Invoke-GraphRequest',
    'Invoke-GraphPagedRequest',
    'Invoke-GraphDeltaQuery'
)
