# Force-GPUpdate.ps1
# Runs gpupdate /force remotely with success confirmation

param(
    [Parameter(Mandatory = $true)]
    [string]$ComputerName,
    
    [Parameter()]
    [System.Management.Automation.PSCredential]$Credential,
    
    [Parameter()]
    [string]$UserInput
)

try {
    # Confirm action with user if not already confirmed
    if (-not $UserInput -or $UserInput -ne "confirmed") {
        $confirmation = [System.Windows.Forms.MessageBox]::Show(
            "This will force a Group Policy update on $ComputerName.`n`nThis action may temporarily affect user sessions and network performance.`n`nDo you want to continue?",
            "Confirm Group Policy Update",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        if ($confirmation -ne [System.Windows.Forms.DialogResult]::Yes) {
            return @{
                Success = $false
                Message = "Operation cancelled by user"
            }
        }
    }
    
    Write-Host "Forcing Group Policy update on $ComputerName..." -ForegroundColor Yellow
    Write-Host "This may take a few minutes to complete..." -ForegroundColor Cyan
    
    $scriptBlock = {
        try {
            # Run gpupdate /force
            $process = Start-Process -FilePath "gpupdate.exe" -ArgumentList "/force" -Wait -PassThru -NoNewWindow -RedirectStandardOutput "C:\temp\gpupdate_output.txt" -RedirectStandardError "C:\temp\gpupdate_error.txt"
            
            # Get output and error content
            $output = ""
            $error = ""
            
            if (Test-Path "C:\temp\gpupdate_output.txt") {
                $output = Get-Content "C:\temp\gpupdate_output.txt" -Raw
                Remove-Item "C:\temp\gpupdate_output.txt" -Force -ErrorAction SilentlyContinue
            }
            
            if (Test-Path "C:\temp\gpupdate_error.txt") {
                $error = Get-Content "C:\temp\gpupdate_error.txt" -Raw
                Remove-Item "C:\temp\gpupdate_error.txt" -Force -ErrorAction SilentlyContinue
            }
            
            # Check if gpupdate completed successfully
            $success = $process.ExitCode -eq 0
            
            # Get additional policy information
            $lastPolicyUpdate = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine" -Name "LastRegistryUpdate" -ErrorAction SilentlyContinue).LastRegistryUpdate
            
            return @{
                Success = $success
                ExitCode = $process.ExitCode
                Output = $output
                Error = $error
                ComputerName = $env:COMPUTERNAME
                LastPolicyUpdate = $lastPolicyUpdate
                Timestamp = Get-Date
            }
        }
        catch {
            return @{
                Success = $false
                Error = $_.Exception.Message
                ComputerName = $env:COMPUTERNAME
                Timestamp = Get-Date
            }
        }
    }
    
    # Execute on remote computer
    $session = New-PSSession -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop
    
    try {
        $result = Invoke-Command -Session $session -ScriptBlock $scriptBlock -ErrorAction Stop
        
        if ($result.Success) {
            Write-Host "Group Policy update completed successfully!" -ForegroundColor Green
            Write-Host "Computer: $($result.ComputerName)" -ForegroundColor Cyan
            Write-Host "Completion Time: $($result.Timestamp)" -ForegroundColor Cyan
            
            if ($result.LastPolicyUpdate) {
                Write-Host "Last Policy Update: $($result.LastPolicyUpdate)" -ForegroundColor Cyan
            }
            
            # Show output if available
            if ($result.Output) {
                Write-Host "`nGPUpdate Output:" -ForegroundColor Yellow
                Write-Host $result.Output -ForegroundColor White
            }
            
            # Show success message
            [System.Windows.Forms.MessageBox]::Show(
                "Group Policy update completed successfully on $ComputerName.`n`nThe policies have been refreshed and should take effect immediately.",
                "GP Update Successful",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            return @{
                Success = $true
                Message = "Group Policy update completed successfully"
                Details = $result
            }
        }
        else {
            $errorMsg = "Group Policy update failed"
            if ($result.Error) {
                $errorMsg += ": $($result.Error)"
            }
            if ($result.ExitCode) {
                $errorMsg += " (Exit Code: $($result.ExitCode))"
            }
            
            Write-Host $errorMsg -ForegroundColor Red
            
            if ($result.Error) {
                Write-Host "Error Details: $($result.Error)" -ForegroundColor Red
            }
            
            throw $errorMsg
        }
    }
    finally {
        Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    }
}
catch {
    $errorMessage = "Failed to force Group Policy update: $_"
    Write-Error $errorMessage
    
    [System.Windows.Forms.MessageBox]::Show(
        $errorMessage,
        "GP Update Failed",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    
    return @{
        Success = $false
        Error = $_.Exception.Message
    }
}