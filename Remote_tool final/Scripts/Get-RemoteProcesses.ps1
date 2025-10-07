# Get-RemoteProcesses.ps1
# Shows running processes sorted by CPU and Memory usage in grid view

param(
    [Parameter(Mandatory = $true)]
    [string]$ComputerName,
    
    [Parameter()]
    [System.Management.Automation.PSCredential]$Credential,
    
    [Parameter()]
    [int]$TopCount = 20
)

try {
    Write-Host "Retrieving running processes from $ComputerName..." -ForegroundColor Yellow
    Write-Host "Getting top $TopCount processes by CPU and Memory usage..." -ForegroundColor Cyan
    
    $scriptBlock = {
        param($Count)
        
        try {
            # Get process information with performance counters
            $processes = Get-Process | Where-Object { $_.ProcessName -ne "Idle" } | ForEach-Object {
                try {
                    # Calculate CPU percentage (approximate)
                    $cpuPercent = 0
                    if ($_.CPU) {
                        $cpuPercent = [math]::Round($_.CPU, 2)
                    }
                    
                    # Get memory usage in MB
                    $workingSetMB = [math]::Round($_.WorkingSet / 1MB, 2)
                    $privateMemoryMB = [math]::Round($_.PrivateMemorySize / 1MB, 2)
                    
                    # Get additional process details
                    $processInfo = [PSCustomObject]@{
                        ProcessName = $_.ProcessName
                        Id = $_.Id
                        CPU = $cpuPercent
                        WorkingSetMB = $workingSetMB
                        PrivateMemoryMB = $privateMemoryMB
                        VirtualMemoryMB = [math]::Round($_.VirtualMemorySize / 1MB, 2)
                        Threads = $_.Threads.Count
                        Handles = $_.HandleCount
                        StartTime = try { $_.StartTime } catch { "N/A" }
                        Company = try { $_.Company } catch { "N/A" }
                        Description = try { $_.Description } catch { "N/A" }
                        Path = try { $_.Path } catch { "N/A" }
                        WindowTitle = try { $_.MainWindowTitle } catch { "" }
                        Responding = $_.Responding
                    }
                    
                    return $processInfo
                }
                catch {
                    # Skip processes we can't access
                    return $null
                }
            } | Where-Object { $_ -ne $null }
            
            # Sort by CPU usage (descending) and get top processes
            $topCPUProcesses = $processes | Sort-Object CPU -Descending | Select-Object -First $Count
            
            # Sort by Memory usage (descending) and get top processes
            $topMemoryProcesses = $processes | Sort-Object WorkingSetMB -Descending | Select-Object -First $Count
            
            # Combine and deduplicate
            $topProcesses = ($topCPUProcesses + $topMemoryProcesses) | 
                Sort-Object @{Expression = "CPU"; Descending = $true}, @{Expression = "WorkingSetMB"; Descending = $true} |
                Group-Object Id | ForEach-Object { $_.Group | Select-Object -First 1 } |
                Select-Object -First $Count
            
            # Get system information for context
            $systemInfo = @{
                ComputerName = $env:COMPUTERNAME
                TotalProcesses = $processes.Count
                Timestamp = Get-Date
                OSInfo = (Get-WmiObject Win32_OperatingSystem | Select-Object Caption, TotalVisibleMemorySize, FreePhysicalMemory)
            }
            
            return @{
                Success = $true
                Processes = $topProcesses
                SystemInfo = $systemInfo
                AllProcesses = $processes
            }
        }
        catch {
            return @{
                Success = $false
                Error = $_.Exception.Message
            }
        }
    }
    
    # Execute on remote computer
    $session = New-PSSession -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop
    
    try {
        $result = Invoke-Command -Session $session -ScriptBlock $scriptBlock -ArgumentList $TopCount
        
        if ($result.Success) {
            Write-Host "Successfully retrieved process information!" -ForegroundColor Green
            Write-Host "Computer: $($result.SystemInfo.ComputerName)" -ForegroundColor Cyan
            Write-Host "Total Processes: $($result.SystemInfo.TotalProcesses)" -ForegroundColor Cyan
            Write-Host "Scan Time: $($result.SystemInfo.Timestamp)" -ForegroundColor Cyan
            
            # Display top processes in console
            Write-Host "`nTop $TopCount Processes by CPU and Memory Usage:" -ForegroundColor Yellow
            $result.Processes | Format-Table -Property ProcessName, Id, CPU, WorkingSetMB, PrivateMemoryMB, Threads, Handles, Responding -AutoSize
            
            # Show detailed view in grid
            $gridTitle = "Top Processes - $ComputerName (Top $TopCount by CPU/Memory)"
            $selectedProcesses = $result.Processes | Out-GridView -Title $gridTitle -PassThru
            
            # If user selected processes, show additional details
            if ($selectedProcesses) {
                foreach ($proc in $selectedProcesses) {
                    Write-Host "`n--- Process Details: $($proc.ProcessName) (PID: $($proc.Id)) ---" -ForegroundColor Yellow
                    Write-Host "Company: $($proc.Company)" -ForegroundColor White
                    Write-Host "Description: $($proc.Description)" -ForegroundColor White
                    Write-Host "Path: $($proc.Path)" -ForegroundColor White
                    Write-Host "Start Time: $($proc.StartTime)" -ForegroundColor White
                    Write-Host "Window Title: $($proc.WindowTitle)" -ForegroundColor White
                    Write-Host "Responding: $($proc.Responding)" -ForegroundColor White
                }
            }
            
            # Show system memory information
            if ($result.SystemInfo.OSInfo) {
                $osInfo = $result.SystemInfo.OSInfo
                $totalMemoryGB = [math]::Round($osInfo.TotalVisibleMemorySize / 1MB, 2)
                $freeMemoryGB = [math]::Round($osInfo.FreePhysicalMemory / 1MB, 2)
                $usedMemoryGB = $totalMemoryGB - $freeMemoryGB
                $memoryUsagePercent = [math]::Round(($usedMemoryGB / $totalMemoryGB) * 100, 1)
                
                Write-Host "`nSystem Memory Information:" -ForegroundColor Yellow
                Write-Host "OS: $($osInfo.Caption)" -ForegroundColor White
                Write-Host "Total Memory: $totalMemoryGB GB" -ForegroundColor White
                Write-Host "Used Memory: $usedMemoryGB GB ($memoryUsagePercent%)" -ForegroundColor White
                Write-Host "Free Memory: $freeMemoryGB GB" -ForegroundColor White
            }
            
            return @{
                Success = $true
                Message = "Process information retrieved successfully"
                ProcessCount = $result.Processes.Count
                TotalSystemProcesses = $result.SystemInfo.TotalProcesses
                SystemInfo = $result.SystemInfo
                TopProcesses = $result.Processes
            }
        }
        else {
            throw $result.Error
        }
    }
    finally {
        Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    }
}
catch {
    $errorMessage = "Failed to retrieve process information: $_"
    Write-Error $errorMessage
    
    [System.Windows.Forms.MessageBox]::Show(
        $errorMessage,
        "Process Retrieval Failed",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    
    return @{
        Success = $false
        Error = $_.Exception.Message
    }
}