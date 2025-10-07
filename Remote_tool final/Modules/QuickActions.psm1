# QuickActions.psm1
# Handles Quick Actions panel and functionality

$script:QuickActionsPanel = $null
$script:ComputerNameTextBox = $null
$script:ConnectButton = $null
$script:StatusLabel = $null
$script:IPLabel = $null
$script:ActionButtons = @()
$script:CurrentComputer = ""
$script:ScriptRoot = ""
$script:MainStatusLabel = $null
$script:ConnectionLabel = $null

function Get-ConnectionAuthDescriptor {
    [CmdletBinding()]
    param()

    $summaryCommand = Get-Command Get-RemoteAuthenticationSummary -ErrorAction SilentlyContinue
    if ($summaryCommand) {
        try {
            $summary = Get-RemoteAuthenticationSummary
            if ($summary) { return $summary }
        }
        catch {
            Write-Verbose "Get-RemoteAuthenticationSummary failed: $_"
        }
    }

    $fallbackCommand = Get-Command Get-AuthenticationSummaryText -ErrorAction SilentlyContinue
    if ($fallbackCommand) {
        try {
            return Get-AuthenticationSummaryText
        }
        catch {
            Write-Verbose "Get-AuthenticationSummaryText failed: $_"
        }
    }

    return 'Auth: Unknown'
}


function Initialize-QuickActionsPanel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Panel]$Panel,
        
        [Parameter(Mandatory = $true)]
        [string]$ScriptRoot,
        
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.ToolStripStatusLabel]$StatusLabel,
        
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.ToolStripStatusLabel]$ConnectionLabel
    )
    
    $script:QuickActionsPanel = $Panel
    $script:ScriptRoot = $ScriptRoot
    $script:MainStatusLabel = $StatusLabel
    $script:ConnectionLabel = $ConnectionLabel
    
    # Top panel for computer connection
    $topPanel = New-Object System.Windows.Forms.Panel
    $topPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $topPanel.Height = 110
    $topPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    
    # Computer name label
    $lblComputer = New-Object System.Windows.Forms.Label
    $lblComputer.Location = New-Object System.Drawing.Point(10, 15)
    $lblComputer.Size = New-Object System.Drawing.Size(120, 23)
    $lblComputer.Text = "Computer Name:"
    $lblComputer.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $topPanel.Controls.Add($lblComputer)
    
    # Computer name textbox
    $script:ComputerNameTextBox = New-Object System.Windows.Forms.TextBox
    $script:ComputerNameTextBox.Location = New-Object System.Drawing.Point(135, 12)
    $script:ComputerNameTextBox.Size = New-Object System.Drawing.Size(250, 23)
    $script:ComputerNameTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $topPanel.Controls.Add($script:ComputerNameTextBox)
    
    # Connect button
    $script:ConnectButton = New-Object System.Windows.Forms.Button
    $script:ConnectButton.Location = New-Object System.Drawing.Point(395, 10)
    $script:ConnectButton.Size = New-Object System.Drawing.Size(100, 25)
    $script:ConnectButton.Text = "Connect"
    $script:ConnectButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:ConnectButton.BackColor = [System.Drawing.Color]::FromArgb(52, 152, 219)
    $script:ConnectButton.ForeColor = [System.Drawing.Color]::White
    $script:ConnectButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $script:ConnectButton.Add_Click({
        Connect-ToRemoteComputer
    })
    $topPanel.Controls.Add($script:ConnectButton)
    
    # Status panel
    $statusPanel = New-Object System.Windows.Forms.GroupBox
    $statusPanel.Location = New-Object System.Drawing.Point(10, 45)
    $statusPanel.Size = New-Object System.Drawing.Size(465, 55)
    $statusPanel.Text = "Connection Status"
    $topPanel.Controls.Add($statusPanel)
    
    # Status indicator (online/offline)
    $script:StatusLabel = New-Object System.Windows.Forms.Label
    $script:StatusLabel.Location = New-Object System.Drawing.Point(10, 20)
    $script:StatusLabel.Size = New-Object System.Drawing.Size(100, 20)
    $script:StatusLabel.Text = "$([char]0x25CF) Offline"
    $script:StatusLabel.ForeColor = [System.Drawing.Color]::Red
    $script:StatusLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $statusPanel.Controls.Add($script:StatusLabel)
    
    # IP address label
    $script:IPLabel = New-Object System.Windows.Forms.Label
    $script:IPLabel.Location = New-Object System.Drawing.Point(120, 20)
    $script:IPLabel.Size = New-Object System.Drawing.Size(280, 20)
    $script:IPLabel.Text = "IP: Not connected"
    $statusPanel.Controls.Add($script:IPLabel)
    
    # Keyboard shortcuts
    $script:ComputerNameTextBox.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            Connect-ToRemoteComputer
        }
    })
    
    $Panel.Controls.Add($topPanel)
    
    # Actions panel
    $actionsPanel = New-Object System.Windows.Forms.Panel
    $actionsPanel.Location = New-Object System.Drawing.Point(0, 115)
    $actionsPanel.Size = New-Object System.Drawing.Size($Panel.Width, ($Panel.Height - 115))
    $actionsPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                           [System.Windows.Forms.AnchorStyles]::Bottom -bor 
                           [System.Windows.Forms.AnchorStyles]::Left -bor 
                           [System.Windows.Forms.AnchorStyles]::Right
    $actionsPanel.AutoScroll = $true
    
    # Create action buttons grid (5 buttons) - larger for full tab
    $buttonWidth = 220
    $buttonHeight = 110
    $buttonSpacing = 15
    $startX = 15
    $startY = 15
    
    for ($i = 0; $i -lt 5; $i++) {
        $button = New-Object System.Windows.Forms.Button
        
        # Calculate position (2 columns layout)
        $col = $i % 2
        $row = [Math]::Floor($i / 2)
        $x = $startX + ($col * ($buttonWidth + $buttonSpacing))
        $y = $startY + ($row * ($buttonHeight + $buttonSpacing))
        
        $button.Location = New-Object System.Drawing.Point($x, $y)
        $button.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
        $button.Text = "Action $(($i + 1))"
        $button.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $button.FlatAppearance.BorderSize = 2
        $button.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $button.UseCompatibleTextRendering = $false
        $button.AutoEllipsis = $true
        $button.Enabled = $false
        $button.Name = "ActionButton_$i"  # Store index in Name property

        # Add click handler that resolves the action using the sender button
        $button.Add_Click({
            param($sender, $eventArgs)
            Invoke-QuickAction -ActionButton $sender
        })
        
        $actionsPanel.Controls.Add($button)
        $script:ActionButtons += $button
    }
    
    $Panel.Controls.Add($actionsPanel)
}

function Connect-ToRemoteComputer {
    [CmdletBinding()]
    param()
    
    $computerName = $script:ComputerNameTextBox.Text.Trim()
    
    if ([string]::IsNullOrEmpty($computerName)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a computer name.",
            "Computer Name Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    $authDescriptor = Get-ConnectionAuthDescriptor
    $authDescriptor = Get-ConnectionAuthDescriptor
    $script:MainStatusLabel.Text = "Connecting to $computerName..."
    $script:ConnectButton.Enabled = $false
    
    try {
        # Test connection
        $connectionResult = Test-RemoteConnection -ComputerName $computerName
        
        if ($connectionResult.IsOnline) {
            $script:CurrentComputer = $computerName
            $script:StatusLabel.Text = "$([char]0x25CF) Online"
            $script:StatusLabel.ForeColor = [System.Drawing.Color]::Green
            $script:IPLabel.Text = "IP: $($connectionResult.IPAddress)"
            
            # Enable action buttons
            foreach ($button in $script:ActionButtons) {
                $button.Enabled = $true
            }
            
            $script:ConnectionLabel.Text = "Connection: Connected to $computerName ($authDescriptor)"
            $script:MainStatusLabel.Text = "Connected successfully"
            
            Write-Log "Connected to $computerName" -Level SUCCESS -TargetComputer $computerName
        }
        else {
            $script:StatusLabel.Text = "$([char]0x25CF) Offline"
            $script:StatusLabel.ForeColor = [System.Drawing.Color]::Red
            $script:IPLabel.Text = "IP: Unable to connect"
            
            # Disable action buttons
            foreach ($button in $script:ActionButtons) {
                $button.Enabled = $false
            }
            
            $script:ConnectionLabel.Text = "Connection: Not Connected ($authDescriptor)"
            $script:MainStatusLabel.Text = "Connection failed"
            
            [System.Windows.Forms.MessageBox]::Show(
                "Unable to connect to $computerName. Please verify the computer name and network connectivity.",
                "Connection Failed",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            
            Write-Log "Failed to connect to $computerName" -Level ERROR -TargetComputer $computerName
        }
    }
    catch {
        $script:StatusLabel.Text = "$([char]0x25CF) Error"
        $script:StatusLabel.ForeColor = [System.Drawing.Color]::Red
        $script:IPLabel.Text = "IP: Error occurred"
        $script:MainStatusLabel.Text = "Connection error: $_"
        
        Write-Log "Connection error: $_" -Level ERROR -TargetComputer $computerName
        $script:ConnectionLabel.Text = "Connection: Not Connected ($authDescriptor)"
        
        [System.Windows.Forms.MessageBox]::Show(
            "An error occurred while connecting: $_",
            "Connection Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        $script:ConnectButton.Enabled = $true
    }
}

function Update-QuickActionButtons {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Actions
    )
    
    try {
        for ($i = 0; $i -lt [Math]::Min($Actions.Count, $script:ActionButtons.Count); $i++) {
            $button = $script:ActionButtons[$i]
            $action = $Actions[$i]
            
            # Update button properties with text wrapping (optimized for 220px buttons)
            $displayText = $action.DisplayName
            if ($displayText.Length -gt 20) {
                # Split long text into multiple lines for larger buttons  
                $words = $displayText -split ' '
                $lines = @()
                $currentLine = ""
                
                foreach ($word in $words) {
                    if (($currentLine + " " + $word).Length -le 18) {
                        if ($currentLine -eq "") {
                            $currentLine = $word
                        } else {
                            $currentLine += " " + $word
                        }
                    } else {
                        if ($currentLine -ne "") {
                            $lines += $currentLine
                        }
                        $currentLine = $word
                    }
                }
                if ($currentLine -ne "") {
                    $lines += $currentLine
                }
                $button.Text = $lines -join "`n"
            } else {
                $button.Text = $displayText
            }

            # Store the action object in Tag for reference
            $button.Tag = $action
            
            # Set button color
            try {
                $button.BackColor = [System.Drawing.ColorTranslator]::FromHtml($action.ButtonColor)
                $button.ForeColor = [System.Drawing.Color]::White
            }
            catch {
                # Use default color if parsing fails
                $button.BackColor = [System.Drawing.Color]::Gray
                $button.ForeColor = [System.Drawing.Color]::White
            }
            
            # Add tooltip with description
            $tooltip = New-Object System.Windows.Forms.ToolTip
            $tooltip.SetToolTip($button, $action.Description)
        }
        
        Write-Verbose "Updated $($Actions.Count) quick action buttons"
    }
    catch {
        Write-Error "Failed to update quick action buttons: $_"
    }
}

function Invoke-QuickAction {
    [CmdletBinding(DefaultParameterSetName = "ByButton")]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = "ByButton")]
        [System.Windows.Forms.Button]$ActionButton,

        [Parameter(Mandatory = $true, ParameterSetName = "ByIndex")]
        [object]$ActionIndex
    )

    if ([string]::IsNullOrEmpty($script:CurrentComputer)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please connect to a computer first.",
            "Not Connected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $button = $null
    $action = $null
    $resolvedIndex = $null

    switch ($PSCmdlet.ParameterSetName) {
        "ByButton" {
            $button = $ActionButton
            $resolvedIndex = [Array]::IndexOf($script:ActionButtons, $button)

            if ($resolvedIndex -lt 0) {
                if ($button.Tag) {
                    $matchedButton = $script:ActionButtons | Where-Object { $_.Tag -eq $button.Tag } | Select-Object -First 1
                    if ($matchedButton) {
                        $button = $matchedButton
                        $resolvedIndex = [Array]::IndexOf($script:ActionButtons, $matchedButton)
                    }
                }

                if ($resolvedIndex -lt 0) {
                    Write-Warning "Clicked control is not mapped to a quick action"
                    return
                }
            }

            $action = $button.Tag
        }
        "ByIndex" {
            $selector = $ActionIndex

            if ($selector -is [System.Windows.Forms.Button]) {
                $button = $selector
                $resolvedIndex = [Array]::IndexOf($script:ActionButtons, $button)
            }
            elseif ($selector -is [PSCustomObject]) {
                $action = $selector
                $button = $script:ActionButtons | Where-Object { $_.Tag -eq $action } | Select-Object -First 1
                if ($button) {
                    $resolvedIndex = [Array]::IndexOf($script:ActionButtons, $button)
                }
            }
            else {
                $parsedIndex = $null

                if ($selector -is [int]) {
                    $parsedIndex = [int]$selector
                }
                else {
                    $tempIndex = 0
                    if ([int]::TryParse([string]$selector, [ref]$tempIndex)) {
                        $parsedIndex = $tempIndex
                    }
                }

                if ($null -eq $parsedIndex) {
                    Write-Warning "Invalid quick action selector: $selector"
                    return
                }

                if ($parsedIndex -lt 0 -or $parsedIndex -ge $script:ActionButtons.Count) {
                    Write-Warning "Invalid quick action index: $parsedIndex"
                    return
                }

                $resolvedIndex = $parsedIndex
                $button = $script:ActionButtons[$resolvedIndex]
            }

            if ($null -eq $button) {
                Write-Warning "Unable to resolve quick action from selector: $selector"
                return
            }

            if ($null -eq $action) {
                $action = $button.Tag
            }
        }
    }

    if ($null -eq $button -or $button -isnot [System.Windows.Forms.Button]) {
        Write-Warning "Unable to resolve quick action button"
        return
    }

    if ($null -eq $action -or $action -isnot [PSCustomObject]) {
        Write-Warning "No action configured for this button"
        return
    }

    $script:MainStatusLabel.Text = "Executing: $($action.DisplayName)..."
    $button.Enabled = $false

    try {
        Write-Log "Starting action: $($action.DisplayName)" -Level INFO -TargetComputer $script:CurrentComputer -Action $action.Name

        # Check if action requires input
        $userInput = $null
        if ($action.RequiresInput) {
            $inputForm = New-Object System.Windows.Forms.Form
            $inputForm.Text = $action.DisplayName
            $inputForm.Size = New-Object System.Drawing.Size(400, 200)
            $inputForm.StartPosition = "CenterParent"
            $inputForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

            $lblPrompt = New-Object System.Windows.Forms.Label
            $lblPrompt.Location = New-Object System.Drawing.Point(10, 20)
            $lblPrompt.Size = New-Object System.Drawing.Size(360, 40)
            $lblPrompt.Text = "Enter required input for this action:"
            $inputForm.Controls.Add($lblPrompt)

            $txtInput = New-Object System.Windows.Forms.TextBox
            $txtInput.Location = New-Object System.Drawing.Point(10, 70)
            $txtInput.Size = New-Object System.Drawing.Size(360, 23)
            $inputForm.Controls.Add($txtInput)

            $btnOK = New-Object System.Windows.Forms.Button
            $btnOK.Location = New-Object System.Drawing.Point(210, 110)
            $btnOK.Size = New-Object System.Drawing.Size(75, 23)
            $btnOK.Text = "OK"
            $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $inputForm.Controls.Add($btnOK)

            $btnCancel = New-Object System.Windows.Forms.Button
            $btnCancel.Location = New-Object System.Drawing.Point(295, 110)
            $btnCancel.Size = New-Object System.Drawing.Size(75, 23)
            $btnCancel.Text = "Cancel"
            $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $inputForm.Controls.Add($btnCancel)

            $result = $inputForm.ShowDialog()

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $userInput = $txtInput.Text
            }
            else {
                Write-Log "Action cancelled by user" -Level INFO -TargetComputer $script:CurrentComputer
                return
            }
        }

        # Build script path
        $scriptPath = Join-Path $script:ScriptRoot $action.ScriptPath

        if (!(Test-Path $scriptPath)) {
            throw "Script not found: $scriptPath"
        }

        # Prepare parameters
        $parameters = @{
            ComputerName = $script:CurrentComputer
        }

        if ($userInput) {
            $parameters['UserInput'] = $userInput
        }

        # Execute script
        $result = Invoke-RemoteScript -ComputerName $script:CurrentComputer -ScriptPath $scriptPath -Parameters $parameters

        if ($result.Success) {
            $script:MainStatusLabel.Text = "$($action.DisplayName) completed successfully"
            Write-Log "Action completed: $($action.DisplayName)" -Level SUCCESS -TargetComputer $script:CurrentComputer -Action $action.Name

            # Display results if available
            if ($result.Result) {
                Show-ActionResults -Title $action.DisplayName -Results $result.Result
            }
        }
        else {
            throw $result.Error
        }
    }
    catch {
        $script:MainStatusLabel.Text = "Action failed: $_"
        Write-Log "Action failed: $_" -Level ERROR -TargetComputer $script:CurrentComputer -Action $action.Name

        [System.Windows.Forms.MessageBox]::Show(
            "Failed to execute action: $_",
            "Action Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        if ($button) {
            $button.Enabled = $true
        }
    }
}

function Show-ActionResults {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,
        
        [Parameter(Mandatory = $true)]
        [object]$Results
    )
    
    # Create results form
    $resultsForm = New-Object System.Windows.Forms.Form
    $resultsForm.Text = "$Title - Results"
    $resultsForm.Size = New-Object System.Drawing.Size(800, 600)
    $resultsForm.StartPosition = "CenterParent"
    
    # Results text box
    $txtResults = New-Object System.Windows.Forms.TextBox
    $txtResults.Multiline = $true
    $txtResults.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $txtResults.ReadOnly = $true
    $txtResults.Font = New-Object System.Drawing.Font("Consolas", 10)
    $txtResults.Dock = [System.Windows.Forms.DockStyle]::Fill
    
    # Format results
    if ($Results -is [Array] -or $Results -is [System.Collections.IEnumerable]) {
        $txtResults.Text = $Results | Format-Table -AutoSize | Out-String
    }
    else {
        $txtResults.Text = $Results | Format-List | Out-String
    }
    
    $resultsForm.Controls.Add($txtResults)
    
    # Close button
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $btnClose.Height = 30
    $btnClose.Add_Click({
        $resultsForm.Close()
    })
    $resultsForm.Controls.Add($btnClose)
    
    $resultsForm.ShowDialog()
}

# Export only the functions called from main script
Export-ModuleMember -Function @(
    'Initialize-QuickActionsPanel',
    'Update-QuickActionButtons'
)






