function Invoke-GraphHuntingQuery {
    <#
    .SYNOPSIS
        Runs a Microsoft 365 Defender Advanced Hunting KQL query.
    .DESCRIPTION
        Executes an Advanced Hunting query against Microsoft Graph and returns the results.
        Requires an active Microsoft Graph connection with ThreatHunting.Read.All permission.
    .PARAMETER Query
        The KQL query to run.
    .PARAMETER ApiVersion
        The Microsoft Graph API version to target (v1.0 or beta). Default is v1.0.
    .PARAMETER ReturnRaw
        If specified, returns the full response object including schema and results.
    .EXAMPLE
        Invoke-GraphHuntingQuery -Query "CloudAppEvents | limit 10"
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string]$Query,

        [Parameter(Mandatory = $false)]
        [ValidateSet('v1.0', 'beta')]
        [string]$ApiVersion = 'v1.0',

        [Parameter(Mandatory = $false)]
        [switch]$ReturnRaw
    )

    process {
        try {
            $mgContext = Get-MgContext -ErrorAction SilentlyContinue
            if (-not $mgContext) {
                throw "No Microsoft Graph context found. Run Connect-MgGraph or Connect-Service first."
            }

            Write-FastLog -Message "Running Advanced Hunting query..." -Context 'Hunting'

            $response = $null
            $cmd = Get-Command Invoke-MgSecurityHuntingQuery -ErrorAction SilentlyContinue
            if ($cmd) {
                if ($cmd.Parameters.ContainsKey('Query')) {
                    $response = Invoke-MgSecurityHuntingQuery -Query $Query -ErrorAction Stop
                }
                elseif ($cmd.Parameters.ContainsKey('BodyParameter')) {
                    $response = Invoke-MgSecurityHuntingQuery -BodyParameter @{ Query = $Query } -ErrorAction Stop
                }
                else {
                    throw "Invoke-MgSecurityHuntingQuery found, but no supported parameter set was detected."
                }
            }
            else {
                $uri = "https://graph.microsoft.com/$ApiVersion/security/runHuntingQuery"
                $bodyJson = @{ query = $Query } | ConvertTo-Json -Depth 10
                $response = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $bodyJson -ContentType "application/json" -ErrorAction Stop
            }

            if ($ReturnRaw) { return $response }
            if ($null -ne $response.results) { return $response.results }
            return $response
        }
        catch {
            Write-FastLog -Message "Advanced Hunting query failed: $($_.Exception.Message)" -Level 'ERROR' -Context 'Hunting'
            throw
        }
    }
}
