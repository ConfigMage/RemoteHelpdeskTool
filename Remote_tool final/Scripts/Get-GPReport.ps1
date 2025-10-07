# Get-GPReport.ps1
# Generates an HTML Group Policy report for troubleshooting

param(
    [Parameter(Mandatory = $true)]
    [string]$ComputerName,
    
    [Parameter()]
    [System.Management.Automation.PSCredential]$Credential,
    [Parameter()]
    [ValidateSet('Computer','User','Both')]
    [string]$Scope = 'Computer'
)

try {
    Write-Host "Generating Group Policy report for $ComputerName..." -ForegroundColor Yellow
    
    # Create output directory if it doesn't exist
    $outputDir = Join-Path $env:TEMP "GPReports"
    if (!(Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    
    # Generate unique filename
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $reportPath = Join-Path $outputDir "GPReport_${ComputerName}_${timestamp}.html"
    
    # Generate the Group Policy report using gpresult
    $scriptBlock = {
        param($ReportPath, $ScopeValue)
        
        # Run gpresult to generate HTML report
        $gpArgs = @('/h', $ReportPath, '/f')
        switch ($ScopeValue) {
            'Computer' { $gpArgs += '/scope'; $gpArgs += 'computer' }
            'User'     { $gpArgs += '/scope'; $gpArgs += 'user' }
            default    { }
        }
        $result = gpresult @gpArgs
        
        return @{
            Success = $LASTEXITCODE -eq 0
            ExitCode = $LASTEXITCODE
            Output = $result
            ReportPath = $ReportPath
        }
    }
    
    # Execute on remote computer
    $session = New-PSSession -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop
    
    try {
        # Create remote temp path
        $remoteTempPath = Invoke-Command -Session $session -ScriptBlock {
            $tempDir = Join-Path $env:TEMP "GPReports"
            if (!(Test-Path $tempDir)) {
                New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
            }
            return Join-Path $tempDir "GPReport_$(hostname)_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        }
        
        # Generate report on remote computer
        $result = Invoke-Command -Session $session -ScriptBlock $scriptBlock -ArgumentList $remoteTempPath, $Scope
        
        if ($result.Success) {
            # Copy report back to local machine
            Copy-Item -FromSession $session -Path $result.ReportPath -Destination $reportPath -ErrorAction Stop
            
            # Clean up remote file
            Invoke-Command -Session $session -ScriptBlock {
                param($Path)
                Remove-Item -Path $Path -Force -ErrorAction SilentlyContinue
            } -ArgumentList $result.ReportPath
            
            Write-Host "Group Policy report generated successfully!" -ForegroundColor Green
            Write-Host "Report saved to: $reportPath" -ForegroundColor Cyan
            
            # Open the report in default browser
            Start-Process $reportPath
            
            return @{
                Success = $true
                Message = "GP report generated successfully"
                ReportPath = $reportPath
                FileSize = (Get-Item $reportPath).Length
            }
        }
        else {
            throw "GPResult failed with exit code: $($result.ExitCode)"
        }
    }
    finally {
        Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    }
}
catch {
    Write-Error "Failed to generate Group Policy report: $_"
    return @{
        Success = $false
        Error = $_.Exception.Message
    }
}
