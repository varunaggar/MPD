function Get-MailboxFolderPermissionsAllFolders {
    <#
    .SYNOPSIS
        Retrieves mailbox folder permissions for all visible folders using Graph-based APIs.
    .DESCRIPTION
        Enumerates folders via Microsoft Graph and retrieves permissions via the Exchange Online Admin API.
        Supports pipeline or CSV input and writes results to CSV for large environments.
    .PARAMETER Mailbox
        Mailbox UPN/alias. Accepts pipeline input.
    .PARAMETER CsvPath
        CSV input file containing a Mailbox column (or UserPrincipalName/UPN).
    .PARAMETER OutputCsvPath
        Destination CSV path for results.
    .PARAMETER ConfigPath
        Path to Custom-Config.xml used for tenant and app credentials.
    .PARAMETER BaseUrl
        Exchange Online Admin API base URL (default: https://outlook.office365.com).
    .PARAMETER IncludeHiddenFolders
        If specified, includes hidden folders in enumeration.
    .PARAMETER ThrottleLimit
        Parallel mailbox processing limit.
    .EXAMPLE
        "user@contoso.com" | Get-MailboxFolderPermissionsAllFolders -OutputCsvPath ./permissions.csv
    .EXAMPLE
        Get-MailboxFolderPermissionsAllFolders -CsvPath ./mailboxes.csv -OutputCsvPath ./permissions.csv
    #>
    [CmdletBinding(DefaultParameterSetName = 'Pipeline')]
    param (
        [Parameter(ParameterSetName = 'Pipeline', ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$Mailbox,

        [Parameter(ParameterSetName = 'Csv', Mandatory = $true)]
        [string]$CsvPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputCsvPath,

        [Parameter(Mandatory = $false)]
        [string]$ConfigPath = (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) "Template/Custom-Config.xml"),

        [Parameter(Mandatory = $false)]
        [string]$BaseUrl = "https://outlook.office365.com",

        [Parameter(Mandatory = $false)]
        [switch]$IncludeHiddenFolders,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 64)]
        [int]$ThrottleLimit = 6
    )

    begin {
        $mailboxes = New-Object System.Collections.Generic.List[string]

        if (-not (Test-Path $ConfigPath)) {
            throw "Config file not found: $ConfigPath"
        }

        [xml]$configXml = Get-Content -Path $ConfigPath -ErrorAction Stop
        $defaults = $configXml.Configuration.Services.Defaults
        $services = @($configXml.Configuration.Services.Service)

        $exSvc = $services | Where-Object { $_.Name -eq 'ExchangeOnline' } | Select-Object -First 1
        if (-not $exSvc) {
            throw "ExchangeOnline service not found in config: $ConfigPath"
        }

        $tenantId = if ($exSvc.TenantId) { $exSvc.TenantId } else { $defaults.TenantId }
        $clientId = if ($exSvc.ClientId) { $exSvc.ClientId } elseif ($exSvc.AppId) { $exSvc.AppId } elseif ($defaults.ClientId) { $defaults.ClientId } else { $defaults.AppId }
        $clientSecret = if ($exSvc.ClientSecret) { $exSvc.ClientSecret } else { $null }

        if (-not $tenantId -or -not $clientId -or -not $clientSecret) {
            throw "Missing TenantId/ClientId/ClientSecret in ExchangeOnline config. ClientSecret is required for app-only token acquisition."
        }

        $commonModulePath = Join-Path (Split-Path $PSScriptRoot -Parent) "Common.psd1"

        if (Test-Path $OutputCsvPath) {
            Remove-Item -Path $OutputCsvPath -Force -ErrorAction SilentlyContinue
        }

        $outputDir = Split-Path -Parent $OutputCsvPath
        if ($outputDir -and -not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }

        $tempFolder = Join-Path $outputDir ("mailbox-perms-" + (Get-Date -Format 'yyyyMMdd-HHmmss'))
        New-Item -ItemType Directory -Path $tempFolder -Force | Out-Null
    }

    process {
        if ($PSCmdlet.ParameterSetName -eq 'Pipeline') {
            if (-not [string]::IsNullOrWhiteSpace($Mailbox)) {
                $mailboxes.Add($Mailbox)
            }
        }
    }

    end {
        if ($PSCmdlet.ParameterSetName -eq 'Csv') {
            if (-not (Test-Path $CsvPath)) {
                throw "CSV file not found: $CsvPath"
            }

            Import-Csv -Path $CsvPath | ForEach-Object {
                $mbx = $_.Mailbox
                if (-not $mbx) { $mbx = $_.UserPrincipalName }
                if (-not $mbx) { $mbx = $_.UPN }

                if ($mbx) {
                    $mailboxes.Add($mbx)
                }
                else {
                    Write-FastLog -Message "CSV row missing Mailbox/UserPrincipalName/UPN: $($_ | ConvertTo-Json -Compress)" -Level 'WARN' -Context 'MailboxPerms'
                }
            }
        }

        if ($mailboxes.Count -eq 0) {
            Write-FastLog -Message "No mailboxes provided. Supply pipeline input or -CsvPath." -Level 'WARN' -Context 'MailboxPerms'
            return
        }

        $tokenEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

        $mailboxes | ForEach-Object -Parallel {
            param($mbx, $tenantId, $clientId, $clientSecret, $tokenEndpoint, $baseUrl, $includeHidden, $tempFolder, $commonModulePath)

            function Get-AccessToken {
                param(
                    [string]$Scope
                )
                $body = @{ client_id = $clientId; client_secret = $clientSecret; grant_type = 'client_credentials'; scope = $Scope }
                $resp = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body -ErrorAction Stop
                return $resp.access_token
            }

            function Get-GraphFolders {
                param(
                    [string]$Mailbox,
                    [string]$GraphToken,
                    [switch]$IncludeHidden
                )

                $headers = @{ Authorization = "Bearer $GraphToken" }
                $select = 'id,displayName,childFolderCount'
                $includeHiddenQuery = if ($IncludeHidden) { '&includeHiddenFolders=true' } else { '' }

                $folders = New-Object System.Collections.Generic.List[object]
                $queue = New-Object System.Collections.Generic.Queue[object]

                $rootUrl = "https://graph.microsoft.com/v1.0/users/$Mailbox/mailFolders?`$top=200&`$select=$select$includeHiddenQuery"
                $queue.Enqueue([PSCustomObject]@{ Url = $rootUrl; Path = '' })

                while ($queue.Count -gt 0) {
                    $item = $queue.Dequeue()
                    $nextUrl = $item.Url
                    $parentPath = $item.Path

                    while ($nextUrl) {
                        $resp = Invoke-RestMethod -Method Get -Uri $nextUrl -Headers $headers -ErrorAction Stop
                        foreach ($f in $resp.value) {
                            $path = if ($parentPath) { "$parentPath\$($f.displayName)" } else { "\$($f.displayName)" }
                            $folders.Add([PSCustomObject]@{ Id = $f.id; Path = $path; HasChildren = ($f.childFolderCount -gt 0) })

                            if ($f.childFolderCount -gt 0) {
                                $childUrl = "https://graph.microsoft.com/v1.0/users/$Mailbox/mailFolders/$($f.id)/childFolders?`$top=200&`$select=$select$includeHiddenQuery"
                                $queue.Enqueue([PSCustomObject]@{ Url = $childUrl; Path = $path })
                            }
                        }
                        $nextUrl = $resp.'@odata.nextLink'
                    }
                }

                return $folders
            }

            function Invoke-AdminApiPaged {
                param(
                    [string]$TenantId,
                    [string]$BaseUrl,
                    [string]$AccessToken,
                    [string]$AnchorMailbox,
                    [string]$Identity
                )

                $headers = @{ Authorization = "Bearer $AccessToken"; 'Content-Type' = 'application/json'; 'X-AnchorMailbox' = $AnchorMailbox }
                $body = @{ CmdletInput = @{ CmdletName = 'Get-MailboxFolderPermission'; Parameters = @{ Identity = $Identity; ResultSize = 'Unlimited' } } } | ConvertTo-Json -Depth 6
                $url = "$BaseUrl/adminapi/v2.0/$TenantId/MailboxFolderPermission"

                $items = @()
                $resp = Invoke-RestMethod -Method Post -Uri $url -Headers $headers -Body $body -ErrorAction Stop

                $currentItems = if ($resp.value) { $resp.value } else { $resp }
                if ($currentItems) { $items += $currentItems }

                $nextLink = $resp.'@odata.nextLink'
                while ($nextLink) {
                    $resp = Invoke-RestMethod -Method Post -Uri $nextLink -Headers $headers -Body $body -ErrorAction Stop
                    $currentItems = if ($resp.value) { $resp.value } else { $resp }
                    if ($currentItems) { $items += $currentItems }
                    $nextLink = $resp.'@odata.nextLink'
                }

                return $items
            }

            try {
                Import-Module $commonModulePath -Force -ErrorAction Stop
                Write-FastLog -Message "Processing mailbox: $mbx" -Context 'MailboxPerms'

                $graphToken = Get-AccessToken -Scope 'https://graph.microsoft.com/.default'
                $adminToken = Get-AccessToken -Scope 'https://outlook.office365.com/.default'

                $folders = Get-GraphFolders -Mailbox $mbx -GraphToken $graphToken -IncludeHidden:$includeHidden
                $anchor = "UPN:$mbx"

                $outFile = Join-Path $tempFolder ("$($mbx -replace '[^a-zA-Z0-9@._-]', '_').csv")

                foreach ($f in $folders) {
                    $identity = "$mbx $($f.Path)"
                    try {
                        $perms = Invoke-AdminApiPaged -TenantId $tenantId -BaseUrl $baseUrl -AccessToken $adminToken -AnchorMailbox $anchor -Identity $identity

                        $records = foreach ($perm in $perms) {
                            [PSCustomObject]@{
                                Mailbox                = $mbx
                                FolderPath             = $f.Path
                                User                   = $perm.User
                                AccessRights           = ($perm.AccessRights -join ',')
                                SharingPermissionFlags = ($perm.SharingPermissionFlags -join ',')
                                IsValid                = $perm.IsValid
                                ObjectState            = $perm.ObjectState
                            }
                        }

                        if ($records) {
                            $records | Export-Csv -Path $outFile -Append -NoTypeInformation -Encoding UTF8
                        }
                    }
                    catch {
                        Write-FastLog -Message "Failed permissions for $identity $($_.Exception.Message)" -Level 'WARN' -Context 'MailboxPerms'
                    }
                }
            }
            catch {
                Write-FastLog -Message "Failed mailbox $mbx $($_.Exception.Message)" -Level 'ERROR' -Context 'MailboxPerms'
            }
        } -ThrottleLimit $ThrottleLimit -ArgumentList $tenantId, $clientId, $clientSecret, $tokenEndpoint, $BaseUrl, $IncludeHiddenFolders, $tempFolder, $commonModulePath

        Merge-CsvFolder -FolderPath $tempFolder | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8
        Write-FastLog -Message "Completed. Output: $OutputCsvPath" -Level 'SUCCESS' -Context 'MailboxPerms'
    }
}
