#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Demonstration script for the Common module.
.DESCRIPTION
    Shows how to import the module, use logging, connect to services, and export data.
#>

# 1. Import the Common Module
# Adjust path if script is moved
$modulePath = Join-Path $PSScriptRoot "../Modules/Common"
Import-Module $modulePath -Force

# Initialize Project Structure (Setup will handle logging init automatically if folders exist)
$configPath = Join-Path $PSScriptRoot "Custom-Config.xml"
Initialize-ModuleDependencies -Setup -ConfigFile $configPath

# Variable names containing hyphens need braces: ${Global:Var-Name}
Write-FastLog -Message "Starting Demo Script" -Level 'INFO' -Context 'Demo'
Write-FastLog -Message "Using Input Path: ${Global:Demo-Usage_Input}" -Level 'INFO' -Context 'Setup'

# 3. Connect to Services (XML-Driven)
# This will read authentication settings from Custom-Config.xml
try {
    Write-FastLog -Message "Attempting Service Connections (Check 'Custom-Config.xml' to enable specific services)..." -Context 'Demo'
    Connect-Service -ConfigPath $configPath
} catch {
    Write-FastLog -Message "Connection Logic Failed: $_" -Level 'ERROR' -Context 'Demo'
}

# 4. Data Processing Example
try {
    Write-FastLog -Message "Creating dummy data..." -Context 'Demo'
    
    $dummyData = 1..5 | ForEach-Object {
        [PSCustomObject]@{
            Id = $_
            Name = "User_$_"
            Email = "user$_@contoso.com"
            Date = Get-Date
        }
    }

    # Use the scaffolded output folder. Variable with hyphen needs {}
    $exportPath = Join-Path ${Global:Demo-Usage_Output} "Result.xlsx"
    
    # export to Excel (or CSV fallback)
    $dummyData | Export-SimpleExcel -Path $exportPath -AutoFit
    
    Write-FastLog -Message "Demo completed successfully. Output at: $exportPath" -Level 'SUCCESS' -Context 'Demo'
}
catch {
    Write-FastLog -Message "An error occurred: $_" -Level 'ERROR' -Context 'Demo'
}
