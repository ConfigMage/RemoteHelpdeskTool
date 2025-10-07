# Clear-AppCache.ps1
# Cleans up Teams and Outlook cache folders for the logged-in user

param(
    [Parameter(Mandatory = $true)]
    [string]$ComputerName,
    
    [Parameter()]
    [System.Management.Automation.PSCredential]$Credential,
    
    [Parameter()]
    [string]$UserInput
)

try {
    # Get target user if specified
    $targetUser = $UserInput
    if ([string]::IsNullOrEmpty($targetUser)) {
        $targetUser = "current"
    }
    
    Write-Host "Clearing application caches on $ComputerName..." -ForegroundColor Yellow
    
    if ($targetUser -ne "current") {
        Write-Host "Target user: $targetUser" -ForegroundColor Cyan
    }
    
    $scriptBlock = {
        param($User)
        
        $results = @{
            Success = $true
            TeamsCleared = $false
            OutlookCleared = $false
            SizeFreed = 0
            Errors = @()
            Details = @()
        }
        
        try {
            # Get the current user or specified user
            if ($User -eq "current" -or [string]::IsNullOrEmpty($User)) {
                $currentUser = $env:USERNAME
                $userProfile = $env:USERPROFILE
            } else {
                $currentUser = $User
                $userProfile = "C:\Users\$User"
                
                if (!(Test-Path $userProfile)) {
                    throw "User profile not found for: $User"
                }
            }
            
            $results.Details += "Processing caches for user: $currentUser"
            $results.Details += "User profile path: $userProfile"
            
            # Teams cache locations
            $teamsPaths = @(
                "$userProfile\AppData\Roaming\Microsoft\Teams\Cache",
                "$userProfile\AppData\Roaming\Microsoft\Teams\GPUCache",
                "$userProfile\AppData\Roaming\Microsoft\Teams\Logs",
                "$userProfile\AppData\Roaming\Microsoft\Teams\tmp",
                "$userProfile\AppData\Roaming\Microsoft\Teams\IndexedDB",
                "$userProfile\AppData\Roaming\Microsoft\Teams\Local Storage",
                "$userProfile\AppData\Roaming\Microsoft\Teams\Session Storage"
            )
            
            # Outlook cache locations
            $outlookPaths = @(
                "$userProfile\AppData\Local\Microsoft\Outlook\RoamCache",
                "$userProfile\AppData\Local\Microsoft\Office\16.0\OfficeFileCache",
                "$userProfile\AppData\Local\Microsoft\Office\Outlook\Offline Address Books",
                "$userProfile\AppData\Roaming\Microsoft\Outlook\Forms"
            )
            
            # Function to get folder size
            function Get-FolderSize {
                param($Path)
                if (Test-Path $Path) {
                    try {
                        $size = (Get-ChildItem $Path -Recurse -Force | Measure-Object -Property Length -Sum).Sum
                        return $size
                    } catch {
                        return 0
                    }
                }
                return 0
            }
            
            # Function to clear folder
            function Clear-Folder {
                param($Path, $AppName)
                
                if (Test-Path $Path) {
                    try {
                        $sizeBefore = Get-FolderSize -Path $Path
                        $results.Details += "Clearing $AppName cache: $Path"
                        
                        # Remove all contents but keep the folder
                        Get-ChildItem -Path $Path -Recurse -Force | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
                        
                        $sizeAfter = Get-FolderSize -Path $Path
                        $sizeFreed = $sizeBefore - $sizeAfter
                        
                        if ($sizeFreed -gt 0) {
                            $results.SizeFreed += $sizeFreed
                            $results.Details += "  Freed: $([math]::Round($sizeFreed / 1MB, 2)) MB"
                        }
                        
                        return $true
                    } catch {
                        $results.Errors += "Failed to clear $AppName cache at $Path : $($_.Exception.Message)"
                        return $false
                    }
                } else {
                    $results.Details += "$AppName cache not found: $Path"
                    return $false
                }
            }
            
            # Clear Teams caches
            $results.Details += "`n--- Clearing Microsoft Teams Caches ---"
            $teamsCleared = $false
            foreach ($path in $teamsPaths) {
                if (Clear-Folder -Path $path -AppName "Teams") {
                    $teamsCleared = $true
                }
            }
            $results.TeamsCleared = $teamsCleared
            
            # Clear Outlook caches
            $results.Details += "`n--- Clearing Microsoft Outlook Caches ---"
            $outlookCleared = $false
            foreach ($path in $outlookPaths) {
                if (Clear-Folder -Path $path -AppName "Outlook") {
                    $outlookCleared = $true
                }
            }
            $results.OutlookCleared = $outlookCleared
            
            # Additional cleanup - Windows temporary files related to Office
            $tempPaths = @(
                "$env:TEMP\*Outlook*",
                "$env:TEMP\*Teams*",
                "$userProfile\AppData\Local\Temp\*Outlook*",
                "$userProfile\AppData\Local\Temp\*Teams*"
            )
            
            $results.Details += "`n--- Clearing Temporary Files ---"
            foreach ($tempPath in $tempPaths) {
                try {
                    $items = Get-ChildItem -Path $tempPath -Force -ErrorAction SilentlyContinue
                    if ($items) {
                        $tempSize = Get-FolderSize -Path $tempPath
                        Remove-Item -Path $tempPath -Force -Recurse -ErrorAction SilentlyContinue
                        $results.SizeFreed += $tempSize
                        $results.Details += "Cleared temp files: $tempPath"
                    }
                } catch {
                    # Ignore temp file cleanup errors
                }
            }
            
            $results.Details += "`n--- Summary ---"
            $results.Details += "Teams cache cleared: $($results.TeamsCleared)"
            $results.Details += "Outlook cache cleared: $($results.OutlookCleared)"
            $results.Details += "Total space freed: $([math]::Round($results.SizeFreed / 1MB, 2)) MB"
            
            return $results
        }
        catch {
            $results.Success = $false
            $results.Errors += $_.Exception.Message
            return $results
        }
    }
    
    # Execute on remote computer
    $session = New-PSSession -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop
    
    try {
        $result = Invoke-Command -Session $session -ScriptBlock $scriptBlock -ArgumentList $targetUser
        
        if ($result.Success) {
            Write-Host "`nCache clearing completed!" -ForegroundColor Green
            Write-Host "Teams cache cleared: $($result.TeamsCleared)" -ForegroundColor Cyan
            Write-Host "Outlook cache cleared: $($result.OutlookCleared)" -ForegroundColor Cyan
            Write-Host "Total space freed: $([math]::Round($result.SizeFreed / 1MB, 2)) MB" -ForegroundColor Yellow
            
            # Show detailed results
            Write-Host "`nDetailed Results:" -ForegroundColor Yellow
            foreach ($detail in $result.Details) {
                Write-Host "  $detail" -ForegroundColor White
            }
            
            # Show errors if any
            if ($result.Errors.Count -gt 0) {
                Write-Host "`nWarnings/Errors:" -ForegroundColor Yellow
                foreach ($error in $result.Errors) {
                    Write-Host "  $error" -ForegroundColor Red
                }
            }
            
            # Show summary message
            $summary = "Cache clearing completed on $ComputerName`n`n"
            $summary += "Teams cache: $($result.TeamsCleared ? 'Cleared' : 'Not found/cleared')`n"
            $summary += "Outlook cache: $($result.OutlookCleared ? 'Cleared' : 'Not found/cleared')`n"
            $summary += "Space freed: $([math]::Round($result.SizeFreed / 1MB, 2)) MB`n`n"
            
            if ($result.SizeFreed -gt 0) {
                $summary += "Users may need to restart Teams and Outlook for optimal performance."
            }
            
            [System.Windows.Forms.MessageBox]::Show(
                $summary,
                "Cache Cleanup Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            return @{
                Success = $true
                Message = "Application caches cleared successfully"
                SizeFreedMB = [math]::Round($result.SizeFreed / 1MB, 2)
                Details = $result
            }
        }
        else {
            throw "Cache clearing failed: $($result.Errors -join '; ')"
        }
    }
    finally {
        Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    }
}
catch {
    $errorMessage = "Failed to clear application caches: $_"
    Write-Error $errorMessage
    
    [System.Windows.Forms.MessageBox]::Show(
        $errorMessage,
        "Cache Cleanup Failed",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    
    return @{
        Success = $false
        Error = $_.Exception.Message
    }
}