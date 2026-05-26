# 1. Setup Variables
$TenantId = "c2efc329-9485-4475-bdc6-267f4b9954ef"
$ClientId = "29aacb3a-c35b-4f76-a071-835936475a48"
$ClientSecret = "MI98Q~hPuhCuK46zr7aUlS0kVNJHPFliXnd8BbeL"

# 2. Get Access Token
$TokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
$Body = @{
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "https://graph.microsoft.com/.default"
    grant_type    = "client_credentials"
}

$TokenResponse = Invoke-RestMethod -Uri $TokenUrl -Method Post -Body $Body
$Headers = @{
    "Authorization" = "Bearer $($TokenResponse.access_token)"
    "Content-Type"  = "application/json"
}

# 3. Define the Delta Endpoint
# Use a local file to store the delta link for subsequent runs
$DeltaLinkFile = "./deltaLink.txt"
if (Test-Path $DeltaLinkFile) {
    $TargetUrl = Get-Content $DeltaLinkFile
    Write-Host "Resuming from saved Delta Link..." -ForegroundColor Cyan
} else {
    $TargetUrl = "https://graph.microsoft.com/v1.0/users/delta"
    Write-Host "Starting fresh sync..." -ForegroundColor Cyan
}

# 4. Fetch the Changes
$Results = @()

do {
    $Response = Invoke-RestMethod -Uri $TargetUrl -Method Get -Headers $Headers
    $Results += $Response.value

    # Check if there is more data (paging)
    if ($Response.'@odata.nextLink') {
        $TargetUrl = $Response.'@odata.nextLink'
    } else {
        # Save the final Delta Link for the NEXT time you run the script
        $Response.'@odata.deltaLink' | Out-File $DeltaLinkFile
        $TargetUrl = $null
    }
} while ($TargetUrl)

# 5. Output Results
$Results | Select-Object displayName, mail, id, @{Name="Action"; Expression={if($_.'@removed'){"Deleted"}else{"Changed/Created"}}}