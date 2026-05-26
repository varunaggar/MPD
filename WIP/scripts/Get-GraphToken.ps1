<#
.SYNOPSIS
    Acquire an access token for Microsoft Graph and display selected claims.

.DESCRIPTION
    Imports a local Az.Accounts module if present, attempts to acquire an
    access token for the Microsoft Graph resource, decodes the JWT payload
    (without printing the raw token) and shows key claims for debugging.

.PARAMETER Interactive
    If supplied, will call Connect-AzAccount interactively if no context exists.
#>

param(
    [switch]$Interactive
)

Set-StrictMode -Version Latest

function Import-LocalAzAccounts {
    $candidate = Join-Path $PSScriptRoot ".." | Resolve-Path -Relative
    $candidate = Join-Path (Split-Path -Parent $PSScriptRoot) "Modules/az.accounts.5.4.1-preview/Az.Accounts.psd1"
    if (Test-Path $candidate) {
        Import-Module $candidate -Force -ErrorAction SilentlyContinue
    }
}

function Get-PlainToken {
    param($tResp)
    if ($null -eq $tResp) { return $null }
    if ($tResp.Token -is [System.Security.SecureString]) {
        $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($tResp.Token)
        try {
            return [Runtime.InteropServices.Marshal]::PtrToStringUni($ptr)
        }
        finally {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
        }
    }
    return $tResp.Token
}

function Decode-JwtPayload {
    param([string]$token)
    if (-not $token) { return $null }
    $parts = $token.Split('.')
    if ($parts.Length -lt 2) { return $null }
    $payload = $parts[1]
    $padding = (4 - ($payload.Length % 4)) % 4
    $payload += '=' * $padding
    try { [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) } catch { return $null }
}

Import-LocalAzAccounts

if (-not (Get-Command Get-AzAccessToken -ErrorAction SilentlyContinue)) {
    Write-Host "Get-AzAccessToken not found in session."
    if ($Interactive) {
        Write-Host "Opening interactive sign-in..."
        Connect-AzAccount -UseDeviceAuthentication | Out-Null
    } else {
        Write-Host "Re-run with -Interactive to sign in interactively." -ForegroundColor Yellow
        exit 2
    }
}

try {
    $tResp = Get-AzAccessToken -ResourceUrl 'https://graph.microsoft.com/'
}
catch {
    Write-Error "Failed to acquire Graph token: $($_.Exception.Message)"
    exit 3
}

$token = Get-PlainToken $tResp
Write-Host "Type: $($tResp.Type)"; Write-Host "TenantId: $($tResp.TenantId)"; Write-Host "UserId: $($tResp.UserId)"; Write-Host "ExpiresOn: $($tResp.ExpiresOn)"

$json = Decode-JwtPayload -token $token
if ($json) {
    $obj = $json | ConvertFrom-Json
    Write-Host "--- Graph token claims ---"
    $obj | Format-List appid, oid, aud, upn, unique_name, tid, exp
} else {
    Write-Host "Token is not a JWT or decoding failed." -ForegroundColor Yellow
}

Write-Host "Done." 
