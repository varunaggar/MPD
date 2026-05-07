function Get-SplunkMailboxPermissionEvents {
    <#
    .SYNOPSIS
        Queries Splunk for Add-MailboxPermission and Add-MailboxFolderPermission events.
    .DESCRIPTION
        Submits a Splunk search job using the REST API and returns matching events.
        Requires a Splunk HEC or REST token with search permissions.
    .PARAMETER SplunkUrl
        Base URL of Splunk (e.g., https://splunk.contoso.com:8089).
    .PARAMETER Token
        Splunk token used for REST authentication.
    .PARAMETER Index
        Splunk index to search. Default: main.
    .PARAMETER Earliest
        Splunk earliest time (e.g., -24h, -7d, 2026-02-01T00:00:00).
    .PARAMETER Latest
        Splunk latest time (e.g., now, 2026-02-01T23:59:59).
    .PARAMETER AdditionalSearch
        Additional raw search conditions appended to the query.
    .PARAMETER Limit
        Maximum results to return. Default: 200.
    .PARAMETER SkipCertificateCheck
        Skip TLS certificate validation (useful for lab environments).
    .EXAMPLE
        Get-SplunkMailboxPermissionEvents -SplunkUrl "https://splunk.contoso.com:8089" -Token $env:SPLUNK_TOKEN -Index "o365" -Earliest "-7d" -Latest "now"
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$SplunkUrl,

        [Parameter(Mandatory = $true)]
        [string]$Token,

        [Parameter(Mandatory = $false)]
        [string]$Index = 'main',

        [Parameter(Mandatory = $false)]
        [string]$Earliest = '-24h',

        [Parameter(Mandatory = $false)]
        [string]$Latest = 'now',

        [Parameter(Mandatory = $false)]
        [string]$AdditionalSearch,

        [Parameter(Mandatory = $false)]
        [int]$Limit = 200,

        [Parameter(Mandatory = $false)]
        [switch]$SkipCertificateCheck
    )

    process {
        try {
            $baseUri = $SplunkUrl.TrimEnd('/')
            $searchCmd = "search index=$Index (`"Add-MailboxPermission`" OR `"Add-MailboxFolderPermission`")"
            if ($AdditionalSearch) { $searchCmd += " $AdditionalSearch" }

            Write-FastLog -Message "Submitting Splunk search job..." -Context 'Splunk'

            $headers = @{ Authorization = "Splunk $Token" }
            $body = @{
                search        = $searchCmd
                earliest_time = $Earliest
                latest_time   = $Latest
                output_mode   = 'json'
                count         = $Limit
            }

            $invokeParams = @{
                Uri         = "$baseUri/services/search/jobs"
                Method      = 'Post'
                Headers     = $headers
                Body        = $body
                ErrorAction = 'Stop'
            }
            if ($SkipCertificateCheck) { $invokeParams['SkipCertificateCheck'] = $true }

            $job = Invoke-RestMethod @invokeParams
            if (-not $job.sid) { throw 'Splunk did not return a search job SID.' }

            $sid = $job.sid
            $resultsUri = "$baseUri/services/search/jobs/$sid/results?output_mode=json&count=$Limit"

            Write-FastLog -Message "Polling Splunk job results..." -Context 'Splunk'

            $resultsParams = @{
                Uri         = $resultsUri
                Method      = 'Get'
                Headers     = $headers
                ErrorAction = 'Stop'
            }
            if ($SkipCertificateCheck) { $resultsParams['SkipCertificateCheck'] = $true }

            $results = Invoke-RestMethod @resultsParams
            if ($results.results) { return $results.results }
            return $results
        }
        catch {
            Write-FastLog -Message "Splunk query failed: $($_.Exception.Message)" -Level 'ERROR' -Context 'Splunk'
            throw
        }
    }
}
