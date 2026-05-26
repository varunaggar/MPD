<#
.SYNOPSIS
    Acquire both Graph and SQL tokens and display their claims side-by-side.

.DESCRIPTION
    Convenience wrapper that calls the Graph and SQL token scripts (or
    performs the same logic inline) and prints a short comparison to help
    identify whether the session principal differs between resource tokens.

.PARAMETER Interactive
    If supplied, will call Connect-AzAccount interactively if no context exists.
#>

param(
    [switch]$Interactive
)

Set-StrictMode -Version Latest

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$graphScript = Join-Path $scriptDir 'Get-GraphToken.ps1'
$sqlScript   = Join-Path $scriptDir 'Get-SqlToken.ps1'

if ((Test-Path $graphScript) -and (Test-Path $sqlScript)) {
    Write-Host "Running Graph token acquisition..." -ForegroundColor Cyan
    & $graphScript -Interactive:$Interactive
    Write-Host "`nRunning SQL token acquisition..." -ForegroundColor Cyan
    & $sqlScript -Interactive:$Interactive

    Write-Host "`nInserting a second test row into DeltaTokens..." -ForegroundColor Cyan
    $localAz = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "Modules/az.accounts.5.4.1-preview/Az.Accounts.psd1"
    if (Test-Path $localAz) { Import-Module $localAz -Force -ErrorAction Stop }
    if (-not (Get-Command Get-AzAccessToken -ErrorAction SilentlyContinue)) {
        Write-Error "Get-AzAccessToken is not available; cannot acquire SQL token."
        exit 1
    }
    $tResp = Get-AzAccessToken -ResourceUrl 'https://database.windows.net/'
    $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($tResp.Token)
    try { $token = [Runtime.InteropServices.Marshal]::PtrToStringUni($ptr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) }

    $connStr = 'Server=tcp:mpd-sql-01.database.windows.net,1433;Database=mpd-db01;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;'
    $cn = New-Object System.Data.SqlClient.SqlConnection $connStr
    $cn.AccessToken = $token
    $cmd = $cn.CreateCommand()
    $cmd.CommandText = "INSERT INTO DeltaTokens (TokenName, TokenValue, CreatedAt, UpdatedAt, IsActive, DeactivatedAt, DeactivationReason) VALUES ('BothTokensTest', 'ok', SYSUTCDATETIME(), SYSUTCDATETIME(), 1, NULL, NULL);"
    try {
        $cn.Open()
        $cmd.ExecuteNonQuery() | Out-Null
        $cn.Close()
        Write-Host "Inserted test row into DeltaTokens." -ForegroundColor Green
    }
    catch {
        Write-Error "SQL insert failed: $($_.Exception.Message)"
        exit 1
    }
} else {
    Write-Error "Required test scripts not found in $scriptDir."
}

Write-Host "Done." 
