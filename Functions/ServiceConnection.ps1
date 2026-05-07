function Connect-Service {
    <#
    .SYNOPSIS
        Connects to services defined in the XML configuration.
    .DESCRIPTION
        Reads the <Services> section of the configuration file and establishes connections
        to enabled services. It supports Interactive, Certificate, ManagedIdentity, and ClientSecret flows.
        
        Specifically supports:
        - Exchange Online: Interactive, Certificate, Managed Identity, and Client Secret (via manual token acquisition).
        - Microsoft Graph: Interactive, Certificate, Managed Identity, Client Secret.
        - Azure (Az): Interactive, Service Principal (Cert/Secret), Managed Identity.
        
    .PARAMETER ConfigPath
        Path to the XML configuration file. Requires <Services> section.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )

    process {
        Write-FastLog -Message "Connecting to services defined in: $ConfigPath" -Context 'Connect'

        # 1. Read XML
        if (-not (Test-Path $ConfigPath)) {
            Write-FastLog -Message "Config file not found: $ConfigPath" -Level 'ERROR' -Context 'Connect'
            return
        }

        try {
            [xml]$configXml = Get-Content -Path $ConfigPath -ErrorAction Stop
        }
        catch {
            Write-FastLog -Message "Failed to read XML: $_" -Level 'ERROR' -Context 'Connect'
            return
        }

        # 2. Parse Services
        if (-not ($configXml.Configuration.Services -and $configXml.Configuration.Services.Service)) {
            Write-FastLog -Message "No services defined in configuration." -Level 'WARN' -Context 'Connect'
            return
        }

        $defaults = $configXml.Configuration.Services.Defaults
        $services = @($configXml.Configuration.Services.Service)

        foreach ($svc in $services) {
            # Skip if explicitly disabled
            if ($svc.Enabled -eq 'false') {
                Write-FastLog -Message "Skipping disabled service: $($svc.Name)" -Level 'INFO' -Context 'Connect'
                continue
            }

            # Merge Defaults & Resolve Variables
            $authMethod = if ($svc.AuthMethod) { $svc.AuthMethod } else { $defaults.AuthMethod }
            $tenantId   = if ($svc.TenantId)   { $svc.TenantId }   else { $defaults.TenantId }
            
            # Helper to check ClientId or AppId
            $clientId   = if ($svc.ClientId) { $svc.ClientId } elseif ($svc.AppId) { $svc.AppId } elseif ($defaults.ClientId) { $defaults.ClientId } else { $defaults.AppId }
            
            $thumbprint = if ($svc.CertificateThumbprint) { $svc.CertificateThumbprint } else { $defaults.CertificateThumbprint }
            $organization = if ($svc.Organization) { $svc.Organization } else { if ($tenantId) { $tenantId } else { $defaults.TenantId } }
            
            # Specifics
            $clientSecret = if ($svc.ClientSecret) { $svc.ClientSecret } else { $null }
            $identityClientId = if ($svc.IdentityClientId) { $svc.IdentityClientId } else { $null }
            $subscriptionId = if ($svc.SubscriptionId) { $svc.SubscriptionId } else { $defaults.SubscriptionId }

            Write-FastLog -Message "Connecting to $($svc.Name) using $authMethod..." -Context 'Connect'

            try {
                switch ($svc.Name) {
                    # =========================================================================================
                    # MICROSOFT GRAPH
                    # =========================================================================================
                    'Graph' {
                        $scopes = if ($svc.Scopes -and $svc.Scopes.Scope) { @($svc.Scopes.Scope) } else { @("User.Read") }
                        
                        # Graph SDK Robustness Parameters
                        $graphEnv = if ($svc.Environment) { $svc.Environment } else { "Global" }
                        $contextScope = if ($svc.ContextScope) { $svc.ContextScope } else { "Process" }

                        $commonGraphParams = @{
                            Environment  = $graphEnv
                            ContextScope = $contextScope
                            NoWelcome    = $true
                            ErrorAction  = "Stop"
                        }

                        switch ($authMethod) {
                            'Interactive' {
                                Connect-MgGraph @commonGraphParams -Scopes $scopes
                            }
                            'Certificate' {
                                if (-not $clientId -or -not $tenantId -or -not $thumbprint) { throw "Missing ClientId, TenantId, or CertificateThumbprint." }
                                Connect-MgGraph @commonGraphParams -ClientId $clientId -TenantId $tenantId -CertificateThumbprint $thumbprint
                            }
                            'ManagedIdentity' {
                                $miParams = @{ Identity = $true } + $commonGraphParams
                                if ($identityClientId) { $miParams['ClientId'] = $identityClientId } # User Assigned Identity
                                Connect-MgGraph @miParams
                            }
                            'ClientSecret' {
                                if (-not $clientId -or -not $tenantId -or -not $clientSecret) { throw "Missing ClientId, TenantId, or ClientSecret." }
                                $secStr = ConvertTo-SecureString $clientSecret -AsPlainText -Force
                                $cred = [System.Management.Automation.PSCredential]::new($clientId, $secStr)
                                Connect-MgGraph @commonGraphParams -TenantId $tenantId -ClientSecretCredential $cred
                            }
                            Default { throw "Unsupported AuthMethod for Graph: $authMethod" }
                        }
                        
                        # Validate Connection
                        if (-not (Get-MgContext)) { 
                            throw "Critical: 'Get-MgContext' returned null. Connection failed silently." 
                        }
                    }

                    # =========================================================================================
                    # EXCHANGE ONLINE
                    # =========================================================================================
                    'ExchangeOnline' {
                        switch ($authMethod) {
                            'Interactive' {
                                $userParam = if ($svc.UserPrincipalName) { @{ UserPrincipalName = $svc.UserPrincipalName } } else { @{} }
                                Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop @userParam
                            }
                            'Certificate' {
                                if (-not $clientId -or -not $organization -or -not $thumbprint) { throw "Missing ClientId, Organization, or CertificateThumbprint." }
                                Connect-ExchangeOnline -AppId $clientId -Organization $organization -CertificateThumbprint $thumbprint -ShowBanner:$false -ErrorAction Stop
                            }
                            'ManagedIdentity' {
                                $params = @{ ManagedIdentity = $true; Organization = $organization }
                                Connect-ExchangeOnline @params -ShowBanner:$false -ErrorAction Stop
                            }
                            'ClientSecret' {
                                # ---------------------------------------------------------------------------------
                                # SPECIAL HANDLING: Exchange Online Module does NOT support ClientSecret directly.
                                # Workaround: Acquire Access Token manually via MSAL/REST, then pass to Connect-ExchangeOnline.
                                # ---------------------------------------------------------------------------------
                                if (-not $clientId -or -not $tenantId -or -not $clientSecret) { throw "Missing ClientId, TenantId, or ClientSecret for Exchange ClientSecret flow." }
                                
                                Write-FastLog -Message "Acquiring access token for Exchange Online..." -Context 'Connect'
                                
                                $tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
                                $body = @{
                                    client_id     = $clientId
                                    scope         = "https://outlook.office365.com/.default"
                                    client_secret = $clientSecret
                                    grant_type    = "client_credentials"
                                }

                                $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ErrorAction Stop
                                $accessToken = $response.access_token

                                if (-not $accessToken) { throw "Failed to acquire access token for Exchange Online." }
                                
                                # Use the token to connect
                                Connect-ExchangeOnline -Organization $organization -AccessToken ($accessToken | ConvertTo-SecureString -AsPlainText -Force) -ShowBanner:$false -ErrorAction Stop
                            }
                            Default { throw "Unsupported AuthMethod for ExchangeOnline: $authMethod" }
                        }
                        
                        # Validate Connection
                        # NOTE: Using Get-ConnectionInformation implies EXO V3 module. For V2 compatibility, check Get-PSSession.
                        if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
                            throw "Critical: 'Get-ConnectionInformation' returned empty. Exchange connection failed silently."
                        }
                    }

                    # =========================================================================================
                    # AZURE (Az Module)
                    # =========================================================================================
                    'Az' {
                        switch ($authMethod) {
                            'Interactive' {
                                $params = @{}
                                if ($tenantId) { $params['Tenant'] = $tenantId }
                                if ($subscriptionId) { $params['Subscription'] = $subscriptionId }
                                Connect-AzAccount @params -ErrorAction Stop
                            }
                            'Certificate' {
                                if (-not $clientId -or -not $tenantId -or -not $thumbprint) { throw "Missing ClientId, TenantId, or CertificateThumbprint." }
                                $params = @{
                                    ServicePrincipal = $true
                                    Tenant           = $tenantId
                                    ApplicationId    = $clientId
                                    CertificateThumbprint = $thumbprint
                                }
                                if ($subscriptionId) { $params['Subscription'] = $subscriptionId }
                                Connect-AzAccount @params -ErrorAction Stop
                            }
                            'ClientSecret' {
                                if (-not $clientId -or -not $tenantId -or -not $clientSecret) { throw "Missing ClientId, TenantId, or ClientSecret." }
                                $secStr = ConvertTo-SecureString $clientSecret -AsPlainText -Force
                                $params = @{
                                    ServicePrincipal = $true
                                    Tenant           = $tenantId
                                    ApplicationId    = $clientId
                                    Credential       = [System.Management.Automation.PSCredential]::new($clientId, $secStr)
                                }
                                if ($subscriptionId) { $params['Subscription'] = $subscriptionId }
                                Connect-AzAccount @params -ErrorAction Stop
                            }
                            'ManagedIdentity' {
                                $params = @{ Identity = $true }
                                if ($identityClientId) { $params['AccountId'] = $identityClientId }
                                if ($subscriptionId) { $params['Subscription'] = $subscriptionId }
                                Connect-AzAccount @params -ErrorAction Stop
                            }
                            Default { throw "Unsupported AuthMethod for Az: $authMethod" }
                        }
                        
                        # Validate Connection
                        if (-not (Get-AzContext)) {
                            throw "Critical: 'Get-AzContext' returned null. Azure connection failed silently."
                        }
                    }

                    Default {
                        Write-FastLog -Message "Unknown service name: $($svc.Name)" -Level 'WARN' -Context 'Connect'
                    }
                }
                
                Write-FastLog -Message "Successfully connected to $($svc.Name)." -Context 'Connect'
            }
            catch {
                Write-FastLog -Message "Failed to connect to $($svc.Name): $($_.Exception.Message)" -Level 'ERROR' -Context 'Connect'
            }
        }
    }
}
