function Start-FastLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$LogPath,
        
        [Parameter(Mandatory = $false)]
        [string]$EventLogSource
    )

    process {
        # Initialize Event Log config if on Windows and Source provided
        $script:FastLogEventSource = $null
        if ($EventLogSource) {
            if ($IsWindows) {
                try {
                    if (-not ([System.Diagnostics.EventLog]::SourceExists($EventLogSource))) {
                        Write-Warning "Event Log source '$EventLogSource' does not exist. Attempting to create it (Requires Admin)..."
                        New-EventLog -LogName Application -Source $EventLogSource -ErrorAction Stop
                        Write-Host "Created Event Log source: $EventLogSource" -ForegroundColor Green
                    }
                    $script:FastLogEventSource = $EventLogSource
                }
                catch {
                    Write-Warning "Failed to configure Event Log source. Ensure you are running as Administrator to create new sources. Error: $_"
                }
            } else {
                Write-Warning "Event Log logging is only supported on Windows."
            }
        }

        if ([string]::IsNullOrWhiteSpace($LogPath)) {
            $callStack = Get-PSCallStack
            # Index 1 is the immediate caller. If called from script, ScriptName is populated.
            # If called interactively or from unsaved block, it might be empty.
            $caller = if ($callStack.Count -gt 1) { $callStack[1] } else { $null }
            
            $prefix = if ($caller -and $caller.ScriptName) { 
                [System.IO.Path]::GetFileNameWithoutExtension($caller.ScriptName) 
            } else { 
                "ScriptLog" 
            }
            
            $LogPath = Join-Path $PWD "${prefix}_$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
        }

        try {
            $script:FastLogPath = $LogPath
            $msg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [INFO] Logging started to: $LogPath"
            Write-Host $msg -ForegroundColor Green
            $msg | Out-File -FilePath $LogPath -Append -Encoding UTF8 -ErrorAction Stop
        }
        catch {
            Write-Warning "Failed to initialize log file: $_"
        }
    }
}

function Write-FastLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO',

        [Parameter(Mandatory = $false)]
        [string]$Context
    )

    process {
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $fullMessage = "[$timestamp] [$Level]"
        if ($Context) { $fullMessage += " [$Context]" }
        $fullMessage += " $Message"

        # Console Output
        switch ($Level) {
            'INFO'    { Write-Host $fullMessage -ForegroundColor Cyan }
            'WARN'    { Write-Host $fullMessage -ForegroundColor Yellow }
            'ERROR'   { Write-Host $fullMessage -ForegroundColor Red }
            'SUCCESS' { Write-Host $fullMessage -ForegroundColor Green }
        }

        # File Output
        if ($script:FastLogPath) {
            try {
                $fullMessage | Out-File -FilePath $script:FastLogPath -Append -Encoding UTF8 -ErrorAction Stop
            }
            catch {
                Write-Warning "Failed to write to log file: $_"
            }
        }

        # Event Log Output (Windows Only)
        if ($script:FastLogEventSource -and $IsWindows) {
            try {
                $evtType = switch ($Level) {
                    'INFO'    { 'Information' }
                    'SUCCESS' { 'Information' } # Map SUCCESS to Info
                    'WARN'    { 'Warning' }
                    'ERROR'   { 'Error' }
                    Default   { 'Information' }
                }
                # EventID 1 is generic
                Write-EventLog -LogName Application -Source $script:FastLogEventSource -EntryType $evtType -EventId 1 -Message "$fullMessage" -ErrorAction Stop
            }
            catch {
                # Suppress errors to avoid spamming console if EventLog fails intermittently
            }
        }
    }
}
