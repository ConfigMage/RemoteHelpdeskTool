# Get-InstalledUpdates.ps1
# Shows recent Windows updates with KB numbers and installation dates

param(
    [Parameter(Mandatory = $true)]
    [string]$ComputerName,
    
    [Parameter()]
    [System.Management.Automation.PSCredential]$Credential,
    
    [Parameter()]
    [int]$DaysBack = 30
)

try {
    Write-Host "Retrieving Windows updates from $ComputerName (last $DaysBack days)..." -ForegroundColor Yellow
    
    $scriptBlock = {
        param($Days)
        
        try {
            # Get updates using Windows Update COM object
            $session = New-Object -ComObject Microsoft.Update.Session
            $searcher = $session.CreateUpdateSearcher()
            $historyCount = $searcher.GetTotalHistoryCount()
            
            if ($historyCount -eq 0) {
                return @{
                    Success = $true
                    Updates = @()
                    Message = "No update history found"
                }
            }
            
            # Get all history and filter by date
            $cutoffDate = (Get-Date).AddDays(-$Days)
            $history = $searcher.QueryHistory(0, $historyCount) | Where-Object {
                $_.Date -gt $cutoffDate
            }
            
            $updates = @()
            foreach ($update in $history) {
                # Extract KB number from title if present
                $kbNumber = "N/A"
                if ($update.Title -match 'KB(\d+)') {
                    $kbNumber = "KB$($matches[1])"
                }
                
                # Map result codes to friendly names
                $result = switch ($update.ResultCode) {
                    1 { "In Progress" }
                    2 { "Succeeded" }
                    3 { "Succeeded with Errors" }
                    4 { "Failed" }
                    5 { "Aborted" }
                    default { "Unknown ($($update.ResultCode))" }
                }
                
                $updates += [PSCustomObject]@{
                    Title = $update.Title
                    KBNumber = $kbNumber
                    Date = $update.Date
                    Result = $result
                    SupportUrl = $update.SupportUrl
                    Description = $update.Description
                }
            }
            
            return @{
                Success = $true
                Updates = ($updates | Sort-Object Date -Descending)
                TotalCount = $updates.Count
            }
        }
        catch {
            return @{
                Success = $false
                Error = $_.Exception.Message
                Updates = @()
            }
        }
    }
    
    # Execute on remote computer
    $session = New-PSSession -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop
    
    try {
        $result = Invoke-Command -Session $session -ScriptBlock $scriptBlock -ArgumentList $DaysBack
        
        if ($result.Success) {
            if ($result.Updates.Count -gt 0) {
                Write-Host "Found $($result.TotalCount) updates installed in the last $DaysBack days" -ForegroundColor Green
                
                # Display in console
                $result.Updates | Format-Table -Property Date, KBNumber, Result, Title -AutoSize
                
                # Also show in grid view for better interaction
                $result.Updates | Out-GridView -Title "Windows Updates - $ComputerName (Last $DaysBack days)" -PassThru
                
                return @{
                    Success = $true
                    Updates = $result.Updates
                    Count = $result.TotalCount
                    Message = "Retrieved $($result.TotalCount) updates"
                }
            }
            else {
                Write-Host "No updates found in the last $DaysBack days" -ForegroundColor Yellow
                return @{
                    Success = $true
                    Updates = @()
                    Message = "No updates found in specified timeframe"
                }
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
    Write-Error "Failed to retrieve Windows updates: $_"
    return @{
        Success = $false
        Error = $_.Exception.Message
        Updates = @()
    }
}