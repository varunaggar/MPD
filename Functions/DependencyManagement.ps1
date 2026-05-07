function Initialize-ModuleDependencies {
    <#
    .SYNOPSIS
        Validates and loads modules defined in an XML configuration file.
    .DESCRIPTION
        Reads a configuration file, checks if required modules are loaded, and attempts to import them if missing.
        Also provides project structure setup.
    .PARAMETER ConfigFile
        Optional path to the XML configuration file.
        If not provided, defaults to: [ScriptDir]\[ScriptName]\[ScriptName]-Dependencies\[ScriptName]-Config.xml
    .PARAMETER Setup
        If specified, creates a standard directory structure (Log, Input, Output, Config, Dependencies) and returns path variables.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$ConfigFile,

        [Parameter(Mandatory = $false)]
        [switch]$Setup
    )

    process {
        # 0. Resolve Caller Context
        $callStack = Get-PSCallStack
        $caller = if ($callStack.Count -gt 1) { $callStack[1] } else { $null }
        $scriptName = $null
        $scriptDir = $null
        
        if ($caller -and $caller.ScriptName) {
            $scriptName = [System.IO.Path]::GetFileNameWithoutExtension($caller.ScriptName)
            $scriptDir = [System.IO.Path]::GetDirectoryName($caller.ScriptName)
        } else {
             Write-FastLog -Message "Could not determine calling script context. Please ensure this function is called from a script file." -Level 'ERROR' -Context 'Init'
             return
        }

        # define standard Folder Paths
        $rootFolder = Join-Path $scriptDir $scriptName
        $logFolder = Join-Path $rootFolder "$scriptName-Log"
        $inputFolder = Join-Path $rootFolder "$scriptName-Input"
        $outputFolder = Join-Path $rootFolder "$scriptName-Output"
        $configFolder = Join-Path $rootFolder "$scriptName-Config"
        $depFolder = Join-Path $rootFolder "$scriptName-Dependencies"

        # 1. Handle Setup Switch
        if ($Setup) {
            $folders = @($rootFolder, $logFolder, $inputFolder, $outputFolder, $configFolder, $depFolder)
            
            foreach ($f in $folders) {
                if (-not (Test-Path $f)) {
                    try {
                        New-Item -Path $f -ItemType Directory -Force | Out-Null
                        Write-FastLog -Message "Created directory: $f" -Context 'Setup'
                    }
                    catch {
                        Write-FastLog -Message "Failed to create directory $f : $_" -Level 'ERROR' -Context 'Setup'
                    }
                }
            }

            # Set Global Variables for easy access
            Set-Variable -Name "${scriptName}_Root" -Value $rootFolder -Scope Global -Force
            Set-Variable -Name "${scriptName}_Log" -Value $logFolder -Scope Global -Force
            Set-Variable -Name "${scriptName}_Input" -Value $inputFolder -Scope Global -Force
            Set-Variable -Name "${scriptName}_Output" -Value $outputFolder -Scope Global -Force
            Set-Variable -Name "${scriptName}_Config" -Value $configFolder -Scope Global -Force
            Set-Variable -Name "${scriptName}_Dependencies" -Value $depFolder -Scope Global -Force
            
            # Also generic globals for current run (optional, but helpful)
            $global:ProjectPaths = @{
                Root         = $rootFolder
                Log          = $logFolder
                Input        = $inputFolder
                Output       = $outputFolder
                Config       = $configFolder
                Dependencies = $depFolder
            }

            # Initialize Logging to the new Log folder immediately?
            # User requirement: "log file should be create in ScriptName-Log folder"
            if (Test-Path $logFolder) {
                 Start-FastLog -LogPath (Join-Path $logFolder "${scriptName}_$(Get-Date -Format 'yyyyMMdd-HHmmss').log")
            }

            Write-FastLog -Message "Project structure setup completed." -Level 'SUCCESS' -Context 'Setup'
        }

        # 2. Determine Config File Path
        $targetConfig = $null
        if (-not [string]::IsNullOrWhiteSpace($ConfigFile)) {
            $targetConfig = $ConfigFile
        } else {
            # Default location: in Dependencies folder
            $targetConfig = Join-Path $depFolder "$scriptName-Config.xml"
        }

        # 3. Check if file exists (Only proceed with dependency check if config exists or explicit ConfigFile was passed)
        if (-not (Test-Path $targetConfig)) {
            if ($Setup) {
                # If we just ran setup, the config file probably doesn't exist yet, so we return gracefully.
                Write-FastLog -Message "Setup complete. No config file found at $targetConfig (this is expected for a fresh setup)." -Level 'INFO' -Context 'DepCheck'
                return
            } else {
                # If not setup, this is a critical failure
                Write-FastLog -Message "Configuration file not found at: $targetConfig" -Level 'ERROR' -Context 'DepCheck'
                exit 1
            }
        }

        Write-FastLog -Message "Initializing module dependencies from: $targetConfig" -Context 'DepCheck'

        # 4. Read XML
        try {
            [xml]$configXml = Get-Content -Path $targetConfig -ErrorAction Stop
        }
        catch {
            Write-FastLog -Message "Failed to parse XML config: $_" -Level 'ERROR' -Context 'DepCheck'
            exit 1
        }

        # 4. Process Modules
        # Supports structure: <Configuration><Modules><Module Name="..." Version="..." /></Modules></Configuration>
        # Check if <Modules> or <Module> exists
        if ($configXml.Configuration -and $configXml.Configuration.Modules) {
            $modules = $configXml.Configuration.Modules.Module
        } else {
            # Fallback or empty
            $modules = $null
        }

        if (-not $modules) {
             Write-FastLog -Message "No modules definition found in config XML." -Level 'WARN' -Context 'DepCheck'
             return
        }

        # Ensure $modules is invalid array even if single item
        $modules = @($modules)

        foreach ($mod in $modules) {
            $name = $mod.Name
            $version = $mod.Version
            $path = $mod.Path

            # Check if loaded (only check if Name available)
            $isLoaded = $false
            if (-not [string]::IsNullOrWhiteSpace($name)) {
                $isLoaded = Get-Module -Name $name -ErrorAction SilentlyContinue
                if ($isLoaded) {
                     Write-FastLog -Message "Module '$name' is already loaded." -Level 'INFO' -Context 'DepCheck'
                     continue
                }
            }

            Write-FastLog -Message "Loading Module: $(if($name){$name}else{$path})..." -Context 'DepCheck'

            # Try to load
            try {
                if (-not [string]::IsNullOrWhiteSpace($path)) {
                    # XML Path might be relative to config file location or absolute.
                    # Verify path resolution.
                    $resolvedPath = $path
                    if (-not (Test-Path $resolvedPath)) {
                        # Try relative to config file directory
                        $configDir = Split-Path $targetConfig -Parent
                        $tryPath = Join-Path $configDir $path
                        if (Test-Path $tryPath) {
                            $resolvedPath = $tryPath
                        }
                    }
                    
                    Write-FastLog -Message "Importing from path: $resolvedPath" -Context 'DepCheck'
                    Import-Module -Name $resolvedPath -ErrorAction Stop
                }
                elseif (-not [string]::IsNullOrWhiteSpace($name)) {
                    # Import by name (and version if present)
                    $params = @{ Name = $name; ErrorAction = 'Stop' }
                    if (-not [string]::IsNullOrWhiteSpace($version)) { 
                        $params['RequiredVersion'] = $version 
                    }
                    Import-Module @params
                }
                else {
                    throw "Invalid module definition in XML. 'Name' or 'Path' attribute is required."
                }

                Write-FastLog -Message "Successfully loaded module." -Level 'SUCCESS' -Context 'DepCheck'
            }
            catch {
                Write-FastLog -Message "CRITICAL: Failed to load required module. Error: $_" -Level 'ERROR' -Context 'DepCheck'
                exit 1
            }
        }
    }
}
