<#
.SYNOPSIS
    Microsoft Graph REST API helpers — certificate-based authentication.

.DESCRIPTION
    Provides token acquisition, paged requests, and delta query support.
    Assumes Connect-SyncServicePrincipal has already been called
    (which establishes an Az context using the certificate).

    Token is cached for its lifetime. All retry/throttle/paging
    logic is encapsulated here — calling scripts just call the
    top-level functions and get results.

.NOTES
    Dependencies: Az.Accounts (must be installed on the server)
    Config is passed in via Initialize-GraphContext.
#>

# ──────────────────────────────────────────────────────────────
# Module-scoped state
# ──────────────────────────────────────────────────────────────

$script:GraphConfig = $null

$script:TokenCache = @{
    Token     = $null
    ExpiresAt = [datetime]::MinValue
}

# ──────────────────────────────────────────────────────────────
# Public: Initialize-GraphContext
# Called once per script after loading config.
# ──────────────────────────────────────────────────────────────

function Initialize-GraphContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Config
    )

    $script:GraphConfig = $Config
    # Reset token cache so any previous session token is not reused
    $script:TokenCache.Token     = $null
    $script:TokenCache.ExpiresAt = [datetime]::MinValue
    
    Write-LogInfo "Graph context initialised (BaseUrl=$($Config.Graph.BaseUrl))"
}

# ──────────────────────────────────────────────────────────────
# Public: Get-GraphToken
# Returns a valid Graph access token, refreshing if needed.
# The Az context established by Connect-SyncServicePrincipal
# allows Get-AzAccessToken to work without re-authenticating.
# ──────────────────────────────────────────────────────────────

function Get-GraphToken {
    [CmdletBinding()]
    param([switch]$ForceRefresh)

    $now = [datetime]::UtcNow

    if (-not $ForceRefresh -and
        $script:TokenCache.Token -and
        $script:TokenCache.ExpiresAt -gt $now.AddMinutes(2)) {
        # Silence frequent cache hits to keep logs clean
        return $script:TokenCache.Token
    }

    Write-LogInfo "Acquiring Graph access token..."

    try {
        $tokenInfo = Get-AzAccessToken `
            -ResourceUrl "https://graph.microsoft.com/" `
            -ErrorAction Stop

        # Ensure the token is a plain string. Modern Az modules return SecureString.
        $plainToken = if ($tokenInfo.Token -is [System.Security.SecureString]) {
            $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenInfo.Token)
            try {
                [Runtime.InteropServices.Marshal]::PtrToStringUni($ptr)
            }
            finally {
                [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
            }
        } else {
            $tokenInfo.Token
        }

        $script:TokenCache.Token     = $plainToken
        # Expire 10 minutes early as a safety margin
        $script:TokenCache.ExpiresAt = $tokenInfo.ExpiresOn.UtcDateTime.AddMinutes(-10)

        Write-LogInfo "Graph token acquired. Valid until $($script:TokenCache.ExpiresAt) UTC"
        return $plainToken
    }
    catch {
        throw "Failed to acquire Graph access token: $($_.Exception.Message)"
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Invoke-GraphRequest
# Single REST call with retry on 429, 401 and 5xx.
# ──────────────────────────────────────────────────────────────

function Invoke-GraphRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Uri,
        [ValidateSet("GET","POST","PATCH","PUT","DELETE")] [string]$Method = "GET",
        [object]$Body = $null,
        [int]$MaxRetries = $null,
        [int]$TimeoutSec = 100
    )

    # Use config value if not explicitly passed
    if (-not $MaxRetries) {
        $MaxRetries = if ($script:GraphConfig -and $script:GraphConfig.Graph.MaxRetries) { [int]$script:GraphConfig.Graph.MaxRetries } else { 5 }
    }
    $throttleMax = if ($script:GraphConfig -and $script:GraphConfig.Graph.ThrottleBackoffMaxSec) { [int]$script:GraphConfig.Graph.ThrottleBackoffMaxSec } else { 300 }

    $attempt   = 0
    $lastError = $null

    while ($attempt -lt $MaxRetries) {
        $attempt++

        try {
            $token   = Get-GraphToken
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
                $params.Body = if ($Body -isnot [string]) { $Body | ConvertTo-Json -Depth 10 } else { $Body }
            }

            return Invoke-RestMethod @params
        }
        catch {
            $statusCode = $null
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
            
            # Attempt to extract the actual Graph error message from the response body
            $apiErrorMessage = ""
            try {
                $responseStream = $_.Exception.Response.GetResponseStream()
                if ($null -ne $responseStream) {
                    $reader = New-Object System.IO.StreamReader($responseStream)
                    $responseBody = $reader.ReadToEnd()
                    # Basic regex to pull the 'message' field from Graph's error JSON
                    if ($responseBody -match '"message":"([^"]+)"') {
                        $apiErrorMessage = " | API Message: $($Matches[1])"
                    }
                }
            } catch {}

            # 429 — throttled
            if ($statusCode -eq 429) {
                $wait = 30
                try {
                    $h = $_.Exception.Response.Headers["Retry-After"]
                    if ($h) { $wait = [int]$h }
                } catch {}
                $wait = [Math]::Min($wait, $throttleMax)
                
                $logMsg = "Graph throttled (429). Waiting ${wait}s before retry $attempt/$MaxRetries$apiErrorMessage"
                Write-LogWarning $logMsg
                
                Start-Sleep -Seconds $wait
                $lastError = $_
                continue
            }

            # 401 — force token refresh once
            if ($statusCode -eq 401 -and $attempt -eq 1) {
                $logMsg = "Graph returned 401. Refreshing token and retrying...$apiErrorMessage"
                Write-LogWarning $logMsg
                
                Get-GraphToken -ForceRefresh | Out-Null
                $lastError = $_
                continue
            }

            # 5xx — transient server error, exponential backoff
            if ($statusCode -ge 500 -and $statusCode -le 599) {
                $wait = [Math]::Min([Math]::Pow(2, $attempt), 60)
                
                $logMsg = "Graph returned $statusCode. Waiting ${wait}s before retry $attempt/$MaxRetries$apiErrorMessage"
                Write-LogWarning $logMsg
                
                Start-Sleep -Seconds $wait
                $lastError = $_
                continue
            }

            # 410 Gone — delta token expired; re-throw immediately for caller to handle
            if ($statusCode -eq 410) { throw }

            # All other errors — do not retry
            throw
        }
    }

    $finalError = "Invoke-GraphRequest ($Method) failed after $MaxRetries attempts. URI: $Uri | Error: $($lastError.Exception.Message)"
    # Include API message in the final throw if we found one
    if ($apiErrorMessage) { $finalError += $apiErrorMessage }
    
    throw $finalError
}

# ──────────────────────────────────────────────────────────────
# Public: Invoke-GraphPagedRequest
# Follows @odata.nextLink automatically.
# Returns all result objects as a single array.
# ──────────────────────────────────────────────────────────────

function Invoke-GraphPagedRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Uri
    )

    $allObjects = [System.Collections.Generic.List[object]]::new()
    $currentUri = $Uri
    $page       = 0

    do {
        $page++
        Write-LogInfo "Fetching page $page from Graph..."

        $response = Invoke-GraphRequest -Uri $currentUri

        if ($response.value) {
            $allObjects.AddRange([object[]]$response.value)
        }

        $currentUri = $response.'@odata.nextLink'
    } while ($currentUri)

    Write-LogInfo "Paged request complete: $page pages, $($allObjects.Count) total objects"
    return $allObjects.ToArray()
}

# ──────────────────────────────────────────────────────────────
# Public: Invoke-GraphDeltaQuery
# Handles the full delta pattern:
#   - pages through all results following @odata.nextLink
#   - captures @odata.deltaLink at the end of the page chain
#   - handles HTTP 410 Gone (expired token) gracefully
# Returns a hashtable with:
#   Objects      — changed objects
#   DeltaLink    — full delta link URL
#   DeltaToken   — extracted token value
#   TokenExpired — $true if HTTP 410 was received
# ──────────────────────────────────────────────────────────────

function Invoke-GraphDeltaQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Uri
    )

    $allObjects = [System.Collections.Generic.List[object]]::new()
    $currentUri = $Uri
    $deltaLink  = $null
    $page       = 0

    try {
        do {
            $page++
            Write-LogInfo "Fetching delta page $page..."

            $response = Invoke-GraphRequest -Uri $currentUri

            if ($response.value) {
                $allObjects.AddRange([object[]]$response.value)
            }

            if ($response.'@odata.deltaLink') {
                $deltaLink  = $response.'@odata.deltaLink'
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
            throw "Delta query completed $page pages but no @odata.deltaLink was returned"
        }

        Write-LogInfo "Delta query complete: $page pages, $($allObjects.Count) changed objects" 

        return @{
            Objects      = $allObjects.ToArray()
            DeltaLink    = $deltaLink
            DeltaToken   = Get-DeltaTokenFromUrl -Url $deltaLink
            TokenExpired = $false
        }
    }
    catch {
        $statusCode = $null
        try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}

        if ($statusCode -eq 410) {
            $msg = "Delta token expired (HTTP 410 Gone). Re-initialisation required."
            Write-LogWarning $msg
            
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
# Extracts the $deltatoken query parameter from a deltaLink URL.
# ──────────────────────────────────────────────────────────────

function Get-DeltaTokenFromUrl {
    param([string]$Url)

    if ([string]::IsNullOrEmpty($Url)) { return $null }

    if ($Url -match '\$deltatoken=([^&]+)') { return [uri]::UnescapeDataString($Matches[1]) }
    if ($Url -match '\$skiptoken=([^&]+)')  { return [uri]::UnescapeDataString($Matches[1]) }

    return $null
}

# ──────────────────────────────────────────────────────────────
# Exports
# ──────────────────────────────────────────────────────────────

Export-ModuleMember -Function @(
    'Initialize-GraphContext',
    'Get-GraphToken',
    'Invoke-GraphRequest',
    'Invoke-GraphPagedRequest',
    'Invoke-GraphDeltaQuery'
)
