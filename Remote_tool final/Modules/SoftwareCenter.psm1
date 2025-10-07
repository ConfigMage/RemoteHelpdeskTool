# SoftwareCenter.psm1
# Handles remote software installation via MECM (Microsoft Endpoint Configuration Manager)

$script:SoftwareCenterTab = $null
$script:ComputerTextBox = $null
$script:ConnectButton = $null
$script:SoftwareListView = $null
$script:InstallButton = $null
$script:RefreshButton = $null
$script:StatusLabel = $null
$script:ConnectedComputer = ""
$script:ScriptRoot = ""

function Initialize-SoftwareCenterTab {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TabPage]$TabPage,
        
        [Parameter(Mandatory = $true)]
        [string]$ScriptRoot,
        
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.ToolStripStatusLabel]$StatusLabel
    )
    
    $script:SoftwareCenterTab = $TabPage
    $script:ScriptRoot = $ScriptRoot
    $script:StatusLabel = $StatusLabel
    
    # Create connection panel
    $connectionPanel = New-Object System.Windows.Forms.Panel
    $connectionPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $connectionPanel.Height = 60
    $connectionPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    
    # Computer name label
    $lblComputer = New-Object System.Windows.Forms.Label
    $lblComputer.Location = New-Object System.Drawing.Point(10, 20)
    $lblComputer.Size = New-Object System.Drawing.Size(140, 20)
    $lblComputer.Text = "Remote Computer:"
    $lblComputer.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $connectionPanel.Controls.Add($lblComputer)
    
    # Computer name textbox
    $script:ComputerTextBox = New-Object System.Windows.Forms.TextBox
    $script:ComputerTextBox.Location = New-Object System.Drawing.Point(160, 18)
    $script:ComputerTextBox.Size = New-Object System.Drawing.Size(220, 20)
    $script:ComputerTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:ComputerTextBox.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            Connect-ToMECM
        }
    })
    $connectionPanel.Controls.Add($script:ComputerTextBox)
    
    # Connect button
    $script:ConnectButton = New-Object System.Windows.Forms.Button
    $script:ConnectButton.Location = New-Object System.Drawing.Point(390, 16)
    $script:ConnectButton.Size = New-Object System.Drawing.Size(100, 25)
    $script:ConnectButton.Text = "Connect"
    $script:ConnectButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:ConnectButton.BackColor = [System.Drawing.Color]::LightBlue
    $script:ConnectButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $script:ConnectButton.Add_Click({
        Connect-ToMECM
    })
    $connectionPanel.Controls.Add($script:ConnectButton)
    
    # Refresh button
    $script:RefreshButton = New-Object System.Windows.Forms.Button
    $script:RefreshButton.Location = New-Object System.Drawing.Point(500, 16)
    $script:RefreshButton.Size = New-Object System.Drawing.Size(100, 25)
    $script:RefreshButton.Text = "Refresh"
    $script:RefreshButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:RefreshButton.BackColor = [System.Drawing.Color]::LightGreen
    $script:RefreshButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $script:RefreshButton.Enabled = $false
    $script:RefreshButton.Add_Click({
        Refresh-SoftwareList
    })
    $connectionPanel.Controls.Add($script:RefreshButton)
    
    $TabPage.Controls.Add($connectionPanel)
    
    # Create software list panel
    $listPanel = New-Object System.Windows.Forms.Panel
    $listPanel.Location = New-Object System.Drawing.Point(0, 60)
    $listPanel.Size = New-Object System.Drawing.Size($TabPage.Width, ($TabPage.Height - 110))
    $listPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                        [System.Windows.Forms.AnchorStyles]::Bottom -bor 
                        [System.Windows.Forms.AnchorStyles]::Left -bor 
                        [System.Windows.Forms.AnchorStyles]::Right
    
    # Create ListView for software display
    $script:SoftwareListView = New-Object System.Windows.Forms.ListView
    $script:SoftwareListView.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:SoftwareListView.View = [System.Windows.Forms.View]::Details
    $script:SoftwareListView.FullRowSelect = $true
    $script:SoftwareListView.GridLines = $true
    $script:SoftwareListView.MultiSelect = $false
    $script:SoftwareListView.Columns.Add("Application Name", 300) | Out-Null
    $script:SoftwareListView.Columns.Add("Publisher", 200) | Out-Null
    $script:SoftwareListView.Columns.Add("Version", 100) | Out-Null
    $script:SoftwareListView.Columns.Add("AppID", 150) | Out-Null
    $script:SoftwareListView.Columns.Add("Install Status", 120) | Out-Null
    
    # Double-click to install
    $script:SoftwareListView.Add_DoubleClick({
        if ($script:SoftwareListView.SelectedItems.Count -gt 0) {
            Install-SelectedSoftware
        }
    })
    
    $listPanel.Controls.Add($script:SoftwareListView)
    $TabPage.Controls.Add($listPanel)
    
    # Create action buttons panel
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $buttonPanel.Height = 50
    $buttonPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    
    # Install button
    $script:InstallButton = New-Object System.Windows.Forms.Button
    $script:InstallButton.Location = New-Object System.Drawing.Point(10, 10)
    $script:InstallButton.Size = New-Object System.Drawing.Size(140, 35)
    $script:InstallButton.Text = "Install Selected"
    $script:InstallButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:InstallButton.BackColor = [System.Drawing.Color]::RoyalBlue
    $script:InstallButton.ForeColor = [System.Drawing.Color]::White
    $script:InstallButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $script:InstallButton.Enabled = $false
    $script:InstallButton.Add_Click({
        Install-SelectedSoftware
    })
    $buttonPanel.Controls.Add($script:InstallButton)
    
    # Uninstall button
    $uninstallButton = New-Object System.Windows.Forms.Button
    $uninstallButton.Location = New-Object System.Drawing.Point(160, 10)
    $uninstallButton.Size = New-Object System.Drawing.Size(140, 35)
    $uninstallButton.Text = "Uninstall Selected"
    $uninstallButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $uninstallButton.BackColor = [System.Drawing.Color]::Crimson
    $uninstallButton.ForeColor = [System.Drawing.Color]::White
    $uninstallButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $uninstallButton.Enabled = $false
    $uninstallButton.Add_Click({
        Uninstall-SelectedSoftware
    })
    $buttonPanel.Controls.Add($uninstallButton)
    
    # Check Status button
    $statusButton = New-Object System.Windows.Forms.Button
    $statusButton.Location = New-Object System.Drawing.Point(310, 10)
    $statusButton.Size = New-Object System.Drawing.Size(130, 35)
    $statusButton.Text = "Check Status"
    $statusButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $statusButton.BackColor = [System.Drawing.Color]::Orange
    $statusButton.ForeColor = [System.Drawing.Color]::White
    $statusButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $statusButton.Enabled = $false
    $statusButton.Add_Click({
        Check-InstallationStatus
    })
    $buttonPanel.Controls.Add($statusButton)
    
    # Export List button
    $exportButton = New-Object System.Windows.Forms.Button
    $exportButton.Location = New-Object System.Drawing.Point(450, 10)
    $exportButton.Size = New-Object System.Drawing.Size(130, 35)
    $exportButton.Text = "Export List"
    $exportButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $exportButton.BackColor = [System.Drawing.Color]::Green
    $exportButton.ForeColor = [System.Drawing.Color]::White
    $exportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $exportButton.Enabled = $false
    $exportButton.Add_Click({
        Export-SoftwareList
    })
    $buttonPanel.Controls.Add($exportButton)
    
    $TabPage.Controls.Add($buttonPanel)
}

function Connect-ToMECM {
    [CmdletBinding()]
    param()
    
    $computerName = $script:ComputerTextBox.Text.Trim()
    
    if ([string]::IsNullOrEmpty($computerName)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a computer name.",
            "Computer Name Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    $script:StatusLabel.Text = "Connecting to $computerName..."
    $script:ConnectButton.Enabled = $false
    
    try {
        # Test connection
        $connectionResult = Test-RemoteConnection -ComputerName $computerName
        
        if (!$connectionResult.IsOnline) {
            throw "Computer $computerName is not reachable"
        }
        
        # Get MECM software list
        $software = Get-MECMSoftware -ComputerName $computerName
        
        if ($null -ne $software) {
            $script:ConnectedComputer = $computerName
            Update-SoftwareListView -Software $software
            
            # Enable buttons
            $script:RefreshButton.Enabled = $true
            $script:InstallButton.Enabled = $true
            $script:StatusLabel.Text = "Connected to $computerName - Found $($software.Count) applications"
            
            Write-Log "Connected to MECM on $computerName" -Level SUCCESS -TargetComputer $computerName
        }
        else {
            throw "Failed to retrieve software list"
        }
    }
    catch {
        $script:StatusLabel.Text = "Connection failed: $_"
        Write-Log "Failed to connect to MECM on $computerName : $_" -Level ERROR -TargetComputer $computerName
        
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to connect to $computerName : $_",
            "Connection Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        $script:ConnectButton.Enabled = $true
    }
}

function Get-MECMSoftware {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )
    
    try {
        $namespace = "root\ccm\clientSDK"
        $query = "SELECT * FROM CCM_Application"
        
        # Get all applications from MECM
        $software = Invoke-WmiObjectWithCredentials -ComputerName $ComputerName -Namespace $namespace -Query $query
        
        $result = @()
        foreach ($app in $software) {
            # Map install states to friendly names
            $installStatus = switch ($app.InstallState) {
                "NotInstalled" { "Not Installed" }
                "Unknown" { "Unknown" }
                "Error" { "Error" }
                "Installed" { "Installed" }
                "NotEvaluated" { "Not Evaluated" }
                "NotUpdated" { "Not Updated" }
                "NotConfigured" { "Not Configured" }
                default { $app.InstallState }
            }
            
            $result += [PSCustomObject]@{
                Name = if ($app.Name) { $app.Name } else { "N/A" }
                Publisher = if ($app.Publisher) { $app.Publisher } else { "N/A" }
                Version = if ($app.SoftwareVersion) { $app.SoftwareVersion } else { "N/A" }
                AppID = if ($app.Id) { $app.Id } else { "N/A" }
                InstallStatus = $installStatus
                Revision = if ($app.Revision) { $app.Revision } else { "N/A" }
                IsMachineTarget = $app.IsMachineTarget
            }
        }
        
        return $result | Sort-Object Name
    }
    catch {
        Write-Error "Failed to get MECM software from $ComputerName : $_"
        return $null
    }
}

function Update-SoftwareListView {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Software
    )
    
    $script:SoftwareListView.Items.Clear()
    
    foreach ($app in $Software) {
        $item = New-Object System.Windows.Forms.ListViewItem($app.Name)
        $item.SubItems.Add($app.Publisher) | Out-Null
        $item.SubItems.Add($app.Version) | Out-Null
        $item.SubItems.Add($app.AppID) | Out-Null
        $item.SubItems.Add($app.InstallStatus) | Out-Null
        $item.Tag = $app  # Store the full object for later use
        
        # Color code based on status
        switch ($app.InstallStatus) {
            "Installed" { $item.ForeColor = [System.Drawing.Color]::Green }
            "Not Installed" { $item.ForeColor = [System.Drawing.Color]::Black }
            "Error" { $item.ForeColor = [System.Drawing.Color]::Red }
            default { $item.ForeColor = [System.Drawing.Color]::Gray }
        }
        
        $script:SoftwareListView.Items.Add($item) | Out-Null
    }
}

function Install-SelectedSoftware {
    [CmdletBinding()]
    param()
    
    if ($script:SoftwareListView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a software to install.",
            "No Selection",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    $selectedApp = $script:SoftwareListView.SelectedItems[0].Tag
    $appName = $selectedApp.Name
    $appID = $selectedApp.AppID
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to install '$appName' on $($script:ConnectedComputer)?",
        "Confirm Installation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $script:StatusLabel.Text = "Installing $appName..."
        Write-Log "Initiating installation of $appName ($appID) on $($script:ConnectedComputer)" -Level INFO -TargetComputer $script:ConnectedComputer
        
        try {
            $installResult = Install-MECMSoftware -ComputerName $script:ConnectedComputer -AppID $appID -AppDetails $selectedApp
            
            if ($installResult) {
                $script:StatusLabel.Text = "Installation initiated successfully"
                [System.Windows.Forms.MessageBox]::Show(
                    "Installation of '$appName' has been initiated on $($script:ConnectedComputer).`n`nThe installation will proceed in the background.",
                    "Installation Started",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
                Write-Log "Installation of $appName initiated successfully" -Level SUCCESS -TargetComputer $script:ConnectedComputer
                
                # Refresh the list after a delay
                Start-Sleep -Seconds 2
                Refresh-SoftwareList
            }
            else {
                throw "Installation failed"
            }
        }
        catch {
            $script:StatusLabel.Text = "Installation failed"
            Write-Log "Failed to install $appName : $_" -Level ERROR -TargetComputer $script:ConnectedComputer
            
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to initiate installation of '$appName': $_",
                "Installation Failed",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
}

function Install-MECMSoftware {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $true)]
        [string]$AppID,
        
        [Parameter(Mandatory = $true)]
        [object]$AppDetails
    )
    
    try {
        $credential = Get-RemoteCredential
        $namespace = "root\ccm\clientSDK"
        
        # Get the specific application
        $application = Get-WmiObject -ComputerName $ComputerName -Namespace $namespace -Class CCM_Application -Filter "Id='$AppID'" -Credential $credential
        
        if ($null -eq $application) {
            throw "Application not found in MECM"
        }
        
        # Convert boolean values to integers for WMI method
        $isMachineTarget = if ($application.IsMachineTarget) { 1 } else { 0 }
        $isRebootIfNeeded = 0  # False - don't force reboot
        
        # Invoke the Install method
        $result = Invoke-WmiMethod -ComputerName $ComputerName -Namespace $namespace -Class CCM_Application -Name Install -ArgumentList @(
            0,  # EnforcePreference (0 = Foreground)
            $AppID,
            $isMachineTarget,
            $isRebootIfNeeded,
            "Normal",  # Priority
            $application.Revision
        ) -Credential $credential -ErrorAction Stop
        
        if ($result.ReturnValue -eq 0) {
            return $true
        }
        else {
            Write-Error "Installation returned code: $($result.ReturnValue)"
            return $false
        }
    }
    catch {
        Write-Error "Failed to install software: $_"
        return $false
    }
}

function Uninstall-SelectedSoftware {
    [CmdletBinding()]
    param()
    
    if ($script:SoftwareListView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a software to uninstall.",
            "No Selection",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    $selectedApp = $script:SoftwareListView.SelectedItems[0].Tag
    $appName = $selectedApp.Name
    $appID = $selectedApp.AppID
    
    if ($selectedApp.InstallStatus -ne "Installed") {
        [System.Windows.Forms.MessageBox]::Show(
            "'$appName' is not currently installed.",
            "Not Installed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to uninstall '$appName' from $($script:ConnectedComputer)?",
        "Confirm Uninstall",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $script:StatusLabel.Text = "Uninstalling $appName..."
        Write-Log "Initiating uninstall of $appName ($appID) on $($script:ConnectedComputer)" -Level INFO -TargetComputer $script:ConnectedComputer
        
        try {
            $credential = Get-RemoteCredential
            $namespace = "root\ccm\clientSDK"
            
            # Get the specific application
            $application = Get-WmiObject -ComputerName $script:ConnectedComputer -Namespace $namespace -Class CCM_Application -Filter "Id='$AppID'" -Credential $credential
            
            if ($null -eq $application) {
                throw "Application not found in MECM"
            }
            
            # Convert boolean values to integers for WMI method
            $isMachineTarget = if ($application.IsMachineTarget) { 1 } else { 0 }
            $isRebootIfNeeded = 0  # False
            
            # Invoke the Uninstall method
            $result = Invoke-WmiMethod -ComputerName $script:ConnectedComputer -Namespace $namespace -Class CCM_Application -Name Uninstall -ArgumentList @(
                0,  # EnforcePreference
                $AppID,
                $isMachineTarget,
                $isRebootIfNeeded,
                "Normal",  # Priority
                $application.Revision
            ) -Credential $credential -ErrorAction Stop
            
            if ($result.ReturnValue -eq 0) {
                $script:StatusLabel.Text = "Uninstall initiated successfully"
                [System.Windows.Forms.MessageBox]::Show(
                    "Uninstall of '$appName' has been initiated on $($script:ConnectedComputer).",
                    "Uninstall Started",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
                Write-Log "Uninstall of $appName initiated successfully" -Level SUCCESS -TargetComputer $script:ConnectedComputer
                
                # Refresh the list after a delay
                Start-Sleep -Seconds 2
                Refresh-SoftwareList
            }
            else {
                throw "Uninstall returned code: $($result.ReturnValue)"
            }
        }
        catch {
            $script:StatusLabel.Text = "Uninstall failed"
            Write-Log "Failed to uninstall $appName : $_" -Level ERROR -TargetComputer $script:ConnectedComputer
            
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to uninstall '$appName': $_",
                "Uninstall Failed",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
}

function Check-InstallationStatus {
    [CmdletBinding()]
    param()
    
    if ($script:SoftwareListView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a software to check status.",
            "No Selection",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    $selectedApp = $script:SoftwareListView.SelectedItems[0].Tag
    $appName = $selectedApp.Name
    $appID = $selectedApp.AppID
    
    $script:StatusLabel.Text = "Checking status of $appName..."
    
    try {
        $credential = Get-RemoteCredential
        $namespace = "root\ccm\clientSDK"
        
        # Get current application status
        $application = Get-WmiObject -ComputerName $script:ConnectedComputer -Namespace $namespace -Class CCM_Application -Filter "Id='$AppID'" -Credential $credential
        
        if ($null -eq $application) {
            throw "Application not found"
        }
        
        $statusInfo = @"
Application: $($application.Name)
Publisher: $($application.Publisher)
Version: $($application.SoftwareVersion)
Install State: $($application.InstallState)
Evaluation State: $($application.EvaluationState)
Percent Complete: $($application.PercentComplete)%
Error Code: $($application.ErrorCode)
Last Install Time: $($application.LastInstallTime)
Last Evaluation Time: $($application.LastEvalTime)
"@
        
        [System.Windows.Forms.MessageBox]::Show(
            $statusInfo,
            "Application Status",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        $script:StatusLabel.Text = "Status check completed"
    }
    catch {
        $script:StatusLabel.Text = "Failed to check status"
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to check status: $_",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

function Refresh-SoftwareList {
    [CmdletBinding()]
    param()
    
    if ([string]::IsNullOrEmpty($script:ConnectedComputer)) {
        return
    }
    
    $script:StatusLabel.Text = "Refreshing software list..."
    $script:RefreshButton.Enabled = $false
    
    try {
        $software = Get-MECMSoftware -ComputerName $script:ConnectedComputer
        
        if ($null -ne $software) {
            Update-SoftwareListView -Software $software
            $script:StatusLabel.Text = "Software list refreshed - Found $($software.Count) applications"
            Write-Log "Refreshed software list for $($script:ConnectedComputer)" -Level INFO -TargetComputer $script:ConnectedComputer
        }
        else {
            throw "Failed to retrieve software list"
        }
    }
    catch {
        $script:StatusLabel.Text = "Refresh failed"
        Write-Log "Failed to refresh software list: $_" -Level ERROR -TargetComputer $script:ConnectedComputer
    }
    finally {
        $script:RefreshButton.Enabled = $true
    }
}

function Export-SoftwareList {
    [CmdletBinding()]
    param()
    
    if ($script:SoftwareListView.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No software to export.",
            "Nothing to Export",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }
    
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $saveDialog.DefaultExt = "csv"
    $saveDialog.FileName = "SoftwareList_$($script:ConnectedComputer)_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $exportData = @()
            foreach ($item in $script:SoftwareListView.Items) {
                $app = $item.Tag
                $exportData += [PSCustomObject]@{
                    Name = $app.Name
                    Publisher = $app.Publisher
                    Version = $app.Version
                    AppID = $app.AppID
                    InstallStatus = $app.InstallStatus
                    Computer = $script:ConnectedComputer
                    ExportDate = Get-Date
                }
            }
            
            $exportData | Export-Csv -Path $saveDialog.FileName -NoTypeInformation
            
            $script:StatusLabel.Text = "Software list exported successfully"
            Write-Log "Exported software list to $($saveDialog.FileName)" -Level SUCCESS -TargetComputer $script:ConnectedComputer
            
            [System.Windows.Forms.MessageBox]::Show(
                "Software list exported successfully to:`n$($saveDialog.FileName)",
                "Export Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        catch {
            $script:StatusLabel.Text = "Export failed"
            Write-Log "Failed to export software list: $_" -Level ERROR
            
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to export software list: $_",
                "Export Failed",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
}

# Export only the functions called from main script
Export-ModuleMember -Function @(
    'Initialize-SoftwareCenterTab'
)