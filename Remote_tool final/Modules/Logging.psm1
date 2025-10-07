# Logging.psm1
# Enhanced logging system with daily log files and auto-cleanup

$script:LogPath = ""
$script:CurrentLogFile = ""
$script:LogRetentionDays = 30

function Initialize-Logging {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogPath,
        
        [Parameter()]
        [int]$RetentionDays = 30
    )
    
    try {
        # Set module-level variables
        $script:LogPath = $LogPath
        $script:LogRetentionDays = $RetentionDays
        
        # Create log directory if it doesn't exist
        if (!(Test-Path $LogPath)) {
            New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
        }
        
        # Set current log file
        $script:CurrentLogFile = Get-LogFileName
        
        # Clean up old logs
        Remove-OldLogFiles
        
        # Write initialization entry
        Write-Log "Logging system initialized" -Level INFO -SkipConsole
        
        return $true
    }
    catch {
        Write-Error "Failed to initialize logging: $_"
        return $false
    }
}

function Get-LogFileName {
    [CmdletBinding()]
    param()
    
    $dateString = (Get-Date).ToString("yyyy-MM-dd")
    $fileName = "RemoteHelpDesk_$dateString.log"
    return Join-Path $script:LogPath $fileName
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter()]
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS', 'DEBUG')]
        [string]$Level = 'INFO',
        
        [Parameter()]
        [string]$TargetComputer = 'localhost',
        
        [Parameter()]
        [string]$Action = '',
        
        [Parameter()]
        [switch]$SkipConsole
    )
    
    try {
        # Check if we need to rotate to a new log file (date changed)
        $currentFile = Get-LogFileName
        if ($currentFile -ne $script:CurrentLogFile) {
            $script:CurrentLogFile = $currentFile
            Remove-OldLogFiles
        }
        
        # Format log entry
        $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff")
        $username = "$env:USERDOMAIN\$env:USERNAME"
        
        # Build log entry
        $logEntry = "[$timestamp] [$Level] [$username] [$TargetComputer]"
        if (![string]::IsNullOrEmpty($Action)) {
            $logEntry += " [$Action]"
        }
        $logEntry += " $Message"
        
        # Write to file
        if (![string]::IsNullOrEmpty($script:CurrentLogFile)) {
            Add-Content -Path $script:CurrentLogFile -Value $logEntry -ErrorAction SilentlyContinue
        }
        
        # Write to console if requested
        if (!$SkipConsole) {
            $color = switch ($Level) {
                'ERROR' { 'Red' }
                'WARNING' { 'Yellow' }
                'SUCCESS' { 'Green' }
                'DEBUG' { 'Gray' }
                default { 'White' }
            }
            
            Write-Host $logEntry -ForegroundColor $color
        }
    }
    catch {
        # Failsafe - if logging fails, at least write to console
        if (!$SkipConsole) {
            Write-Host "LOGGING ERROR: $_" -ForegroundColor Red
            Write-Host "Original message: $Message" -ForegroundColor Yellow
        }
    }
}

function Remove-OldLogFiles {
    [CmdletBinding()]
    param()
    
    try {
        if ([string]::IsNullOrEmpty($script:LogPath) -or !(Test-Path $script:LogPath)) {
            return
        }
        
        $cutoffDate = (Get-Date).AddDays(-$script:LogRetentionDays)
        
        Get-ChildItem -Path $script:LogPath -Filter "RemoteHelpDesk_*.log" | 
            Where-Object { $_.LastWriteTime -lt $cutoffDate } |
            ForEach-Object {
                try {
                    Remove-Item $_.FullName -Force
                    Write-Verbose "Removed old log file: $($_.Name)"
                }
                catch {
                    Write-Warning "Failed to remove old log file: $($_.Name)"
                }
            }
    }
    catch {
        Write-Warning "Failed to clean up old log files: $_"
    }
}

function Get-LogEntries {
    [CmdletBinding()]
    param(
        [Parameter()]
        [DateTime]$StartDate = (Get-Date).Date,
        
        [Parameter()]
        [DateTime]$EndDate = (Get-Date).Date.AddDays(1),
        
        [Parameter()]
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS', 'DEBUG', 'ALL')]
        [string]$Level = 'ALL',
        
        [Parameter()]
        [string]$TargetComputer = '',
        
        [Parameter()]
        [string]$SearchText = ''
    )
    
    try {
        $entries = @()
        
        # Get log files in date range
        $logFiles = Get-ChildItem -Path $script:LogPath -Filter "RemoteHelpDesk_*.log" |
            Where-Object {
                # Extract date from filename
                if ($_.Name -match 'RemoteHelpDesk_(\d{4}-\d{2}-\d{2})\.log') {
                    $fileDate = [DateTime]::ParseExact($matches[1], 'yyyy-MM-dd', $null)
                    $fileDate -ge $StartDate.Date -and $fileDate -le $EndDate.Date
                }
            }
        
        foreach ($file in $logFiles) {
            $content = Get-Content -Path $file.FullName -ErrorAction SilentlyContinue
            
            foreach ($line in $content) {
                # Parse log entry
                if ($line -match '^\[([^\]]+)\]\s+\[([^\]]+)\]\s+\[([^\]]+)\]\s+\[([^\]]+)\](.*)') {
                    $entry = @{
                        Timestamp = $matches[1]
                        Level = $matches[2]
                        Username = $matches[3]
                        TargetComputer = $matches[4]
                        Message = $matches[5].Trim()
                    }
                    
                    # Apply filters
                    if ($Level -ne 'ALL' -and $entry.Level -ne $Level) { continue }
                    if (![string]::IsNullOrEmpty($TargetComputer) -and $entry.TargetComputer -notlike "*$TargetComputer*") { continue }
                    if (![string]::IsNullOrEmpty($SearchText) -and $line -notlike "*$SearchText*") { continue }
                    
                    $entries += New-Object PSObject -Property $entry
                }
            }
        }
        
        return $entries | Sort-Object Timestamp -Descending
    }
    catch {
        Write-Error "Failed to retrieve log entries: $_"
        return @()
    }
}

function Export-Logs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        
        [Parameter()]
        [ValidateSet('CSV', 'HTML', 'TXT')]
        [string]$Format = 'CSV',
        
        [Parameter()]
        [DateTime]$StartDate = (Get-Date).Date.AddDays(-7),
        
        [Parameter()]
        [DateTime]$EndDate = (Get-Date).Date.AddDays(1)
    )
    
    try {
        $entries = Get-LogEntries -StartDate $StartDate -EndDate $EndDate
        
        if ($entries.Count -eq 0) {
            Write-Warning "No log entries found for the specified date range"
            return $false
        }
        
        switch ($Format) {
            'CSV' {
                $entries | Export-Csv -Path $OutputPath -NoTypeInformation
            }
            'HTML' {
                $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Remote Help Desk Logs</title>
    <style>
        body { font-family: Arial, sans-serif; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #4CAF50; color: white; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .ERROR { color: red; font-weight: bold; }
        .WARNING { color: orange; }
        .SUCCESS { color: green; }
    </style>
</head>
<body>
    <h1>Remote Help Desk Logs</h1>
    <p>Period: $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))</p>
    <table>
        <tr>
            <th>Timestamp</th>
            <th>Level</th>
            <th>Username</th>
            <th>Target Computer</th>
            <th>Message</th>
        </tr>
"@
                foreach ($entry in $entries) {
                    $html += "<tr class='$($entry.Level)'>"
                    $html += "<td>$($entry.Timestamp)</td>"
                    $html += "<td>$($entry.Level)</td>"
                    $html += "<td>$($entry.Username)</td>"
                    $html += "<td>$($entry.TargetComputer)</td>"
                    $html += "<td>$($entry.Message)</td>"
                    $html += "</tr>`n"
                }
                
                $html += @"
    </table>
    <p>Generated: $(Get-Date)</p>
</body>
</html>
"@
                $html | Out-File -FilePath $OutputPath -Encoding UTF8
            }
            'TXT' {
                $entries | ForEach-Object {
                    "$($_.Timestamp) [$($_.Level)] [$($_.Username)] [$($_.TargetComputer)] $($_.Message)"
                } | Out-File -FilePath $OutputPath -Encoding UTF8
            }
        }
        
        Write-Log "Logs exported to $OutputPath" -Level SUCCESS
        return $true
    }
    catch {
        Write-Error "Failed to export logs: $_"
        Write-Log "Failed to export logs: $_" -Level ERROR
        return $false
    }
}

function Get-LogStatistics {
    [CmdletBinding()]
    param(
        [Parameter()]
        [DateTime]$StartDate = (Get-Date).Date,
        
        [Parameter()]
        [DateTime]$EndDate = (Get-Date).Date.AddDays(1)
    )
    
    try {
        $entries = Get-LogEntries -StartDate $StartDate -EndDate $EndDate
        
        $stats = @{
            TotalEntries = $entries.Count
            ByLevel = $entries | Group-Object Level | ForEach-Object { @{$_.Name = $_.Count} }
            ByUser = $entries | Group-Object Username | Select-Object -First 10 | ForEach-Object { @{$_.Name = $_.Count} }
            ByComputer = $entries | Group-Object TargetComputer | Select-Object -First 10 | ForEach-Object { @{$_.Name = $_.Count} }
            TimeRange = "$($StartDate.ToString('yyyy-MM-dd HH:mm')) to $($EndDate.ToString('yyyy-MM-dd HH:mm'))"
        }
        
        return New-Object PSObject -Property $stats
    }
    catch {
        Write-Error "Failed to get log statistics: $_"
        return $null
    }
}

function Clear-Logs {
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$Force
    )
    
    if (!$Force) {
        $confirm = Read-Host "Are you sure you want to clear all log files? (Y/N)"
        if ($confirm -ne 'Y') {
            Write-Host "Operation cancelled" -ForegroundColor Yellow
            return
        }
    }
    
    try {
        Get-ChildItem -Path $script:LogPath -Filter "RemoteHelpDesk_*.log" | Remove-Item -Force
        Write-Host "All log files have been cleared" -ForegroundColor Green
        
        # Reinitialize current log file
        $script:CurrentLogFile = Get-LogFileName
        Write-Log "Logs cleared by user" -Level WARNING
    }
    catch {
        Write-Error "Failed to clear logs: $_"
    }
}

# Export module functions
Export-ModuleMember -Function @(
    'Initialize-Logging',
    'Write-Log',
    'Get-LogEntries',
    'Export-Logs',
    'Get-LogStatistics',
    'Clear-Logs'
)