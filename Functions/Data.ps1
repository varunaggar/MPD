function Export-SimpleExcel {
    <#
    .SYNOPSIS
        Exports data to Excel using the ImportExcel module if available, otherwise warns.
    .DESCRIPTION
        Simplifies export to Excel (xlsx). Requires 'ImportExcel' module.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [PSObject]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [string]$WorksheetName = 'Data',

        [Parameter(Mandatory = $false)]
        [switch]$AutoFit,

        [Parameter(Mandatory = $false)]
        [switch]$Show
    )

    begin {
        $data = @()
        $moduleAvailable = Get-Module -ListAvailable -Name ImportExcel
    }

    process {
        $data += $InputObject
    }

    end {
        if (-not $moduleAvailable) {
            Write-FastLog -Message "Module 'ImportExcel' not found. Falling back to CSV export." -Level 'WARN' -Context 'Data'
            $csvPath = [System.IO.Path]::ChangeExtension($Path, ".csv")
            
            try {
                $data | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force -ErrorAction Stop
                Write-FastLog -Message "Exported to CSV: $csvPath" -Level 'SUCCESS' -Context 'Data'
            }
            catch {
                Write-FastLog -Message "Failed to export CSV: $_" -Level 'ERROR' -Context 'Data'
                throw
            }
        }
        else {
            try {
                $params = @{
                    Path = $Path
                    WorksheetName = $WorksheetName
                    AutoSize = $AutoFit
                    Show = $Show
                }
                $data | Export-Excel @params -ErrorAction Stop
                Write-FastLog -Message "Exported to Excel: $Path" -Level 'SUCCESS' -Context 'Data'
            }
            catch {
                Write-FastLog -Message "Failed to export Excel: $_" -Level 'ERROR' -Context 'Data'
                throw
            }
        }
    }
}

function Merge-CsvFolder {
    <#
    .SYNOPSIS
        Merges all CSV files in a folder into a single list of objects.
    .DESCRIPTION
        Reads all *.csv files in the target folder and outputs a combined array.
        Adds 'SourceFile' property to track origin.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$FolderPath
    )

    process {
        Write-FastLog -Message "Merging CSV files from: $FolderPath" -Context 'Data'
        
        if (-not (Test-Path $FolderPath)) {
            Write-FastLog -Message "Folder not found: $FolderPath" -Level 'ERROR' -Context 'Data'
            return
        }

        $files = Get-ChildItem -Path $FolderPath -Filter *.csv -File
        
        $totalRows = 0
        foreach ($file in $files) {
            Write-FastLog -Message "Processing $($file.Name)..." -Context 'Data'
            $content = Import-Csv -Path $file.FullName
            
            foreach ($row in $content) {
                $row | Add-Member -MemberType NoteProperty -Name 'SourceFile' -Value $file.Name -Force
                $row
            }
            $totalRows += $content.Count
        }
        
        Write-FastLog -Message "Merged $($files.Count) files. Total rows: $totalRows" -Level 'SUCCESS' -Context 'Data'
    }
}
