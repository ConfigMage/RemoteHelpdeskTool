# ScriptLibrary.psm1
# Handles Script Library functionality

$script:ScriptLibraryPanel = $null
$script:ScriptListView = $null
$script:ScriptPreviewBox = $null
$script:SearchBox = $null
$script:Scripts = @()
$script:ScriptRoot = ""
$script:MainStatusLabel = $null

function Initialize-ScriptLibraryPanel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Panel]$Panel,
        
        [Parameter(Mandatory = $true)]
        [string]$ScriptRoot,
        
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.ToolStripStatusLabel]$StatusLabel
    )
    
    $script:ScriptLibraryPanel = $Panel
    $script:ScriptRoot = $ScriptRoot
    $script:MainStatusLabel = $StatusLabel
    
    # Create toolbar
    $toolbar = New-Object System.Windows.Forms.Panel
    $toolbar.Dock = [System.Windows.Forms.DockStyle]::Top
    $toolbar.Height = 40
    $toolbar.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    
    # Search label
    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Location = New-Object System.Drawing.Point(10, 12)
    $lblSearch.Size = New-Object System.Drawing.Size(50, 20)
    $lblSearch.Text = "Search:"
    $lblSearch.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $toolbar.Controls.Add($lblSearch)
    
    # Search textbox
    $script:SearchBox = New-Object System.Windows.Forms.TextBox
    $script:SearchBox.Location = New-Object System.Drawing.Point(65, 10)
    $script:SearchBox.Size = New-Object System.Drawing.Size(250, 23)
    $script:SearchBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:SearchBox.Add_TextChanged({
        Filter-Scripts
    })
    $toolbar.Controls.Add($script:SearchBox)
    
    # Add Script button
    $btnAddScript = New-Object System.Windows.Forms.Button
    $btnAddScript.Location = New-Object System.Drawing.Point(330, 8)
    $btnAddScript.Size = New-Object System.Drawing.Size(100, 25)
    $btnAddScript.Text = "Add Script"
    $btnAddScript.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnAddScript.BackColor = [System.Drawing.Color]::FromArgb(46, 204, 113)
    $btnAddScript.ForeColor = [System.Drawing.Color]::White
    $btnAddScript.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnAddScript.Add_Click({
        Show-AddScriptDialog
    })
    $toolbar.Controls.Add($btnAddScript)
    
    # Edit Script button
    $btnEditScript = New-Object System.Windows.Forms.Button
    $btnEditScript.Location = New-Object System.Drawing.Point(440, 8)
    $btnEditScript.Size = New-Object System.Drawing.Size(100, 25)
    $btnEditScript.Text = "Edit Script"
    $btnEditScript.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnEditScript.BackColor = [System.Drawing.Color]::FromArgb(52, 152, 219)
    $btnEditScript.ForeColor = [System.Drawing.Color]::White
    $btnEditScript.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnEditScript.Add_Click({
        Show-EditScriptDialog
    })
    $toolbar.Controls.Add($btnEditScript)
    
    # Delete Script button
    $btnDeleteScript = New-Object System.Windows.Forms.Button
    $btnDeleteScript.Location = New-Object System.Drawing.Point(550, 8)
    $btnDeleteScript.Size = New-Object System.Drawing.Size(100, 25)
    $btnDeleteScript.Text = "Delete Script"
    $btnDeleteScript.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnDeleteScript.BackColor = [System.Drawing.Color]::FromArgb(231, 76, 60)
    $btnDeleteScript.ForeColor = [System.Drawing.Color]::White
    $btnDeleteScript.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnDeleteScript.Add_Click({
        Remove-SelectedScript
    })
    $toolbar.Controls.Add($btnDeleteScript)
    
    $Panel.Controls.Add($toolbar)
    
    # Create split container
    $splitContainer = New-Object System.Windows.Forms.SplitContainer
    $splitContainer.Location = New-Object System.Drawing.Point(0, 40)
    $splitContainer.Size = New-Object System.Drawing.Size($Panel.Width, ($Panel.Height - 40))
    $splitContainer.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                            [System.Windows.Forms.AnchorStyles]::Bottom -bor 
                            [System.Windows.Forms.AnchorStyles]::Left -bor 
                            [System.Windows.Forms.AnchorStyles]::Right
    $splitContainer.SplitterDistance = 300
    
    # Left panel - Script list
    $script:ScriptListView = New-Object System.Windows.Forms.ListView
    $script:ScriptListView.Dock = [System.Windows.Forms.DockStyle]::Fill
    $script:ScriptListView.View = [System.Windows.Forms.View]::Details
    $script:ScriptListView.FullRowSelect = $true
    $script:ScriptListView.GridLines = $true
    $script:ScriptListView.MultiSelect = $false
    $script:ScriptListView.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # Add columns (adjusted for 300px width) 
    $script:ScriptListView.Columns.Add("Name", 120) | Out-Null
    $script:ScriptListView.Columns.Add("Category", 70) | Out-Null
    $script:ScriptListView.Columns.Add("Description", 130) | Out-Null
    $script:ScriptListView.Columns.Add("Author", 60) | Out-Null
    $script:ScriptListView.Columns.Add("Date Added", 80) | Out-Null
    
    # Selection changed event
    $script:ScriptListView.Add_SelectedIndexChanged({
        if ($script:ScriptListView.SelectedItems.Count -gt 0) {
            $selectedItem = $script:ScriptListView.SelectedItems[0]
            $scriptId = $selectedItem.Tag
            $script = $script:Scripts | Where-Object { $_.id -eq $scriptId }
            if ($script) {
                Display-ScriptPreview -Script $script
            }
        }
    })
    
    # Double-click to copy
    $script:ScriptListView.Add_DoubleClick({
        Copy-ScriptToClipboard
    })
    
    $splitContainer.Panel1.Controls.Add($script:ScriptListView)
    
    # Right panel - Script preview
    $previewPanel = New-Object System.Windows.Forms.Panel
    $previewPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    
    # Preview label
    $lblPreview = New-Object System.Windows.Forms.Label
    $lblPreview.Location = New-Object System.Drawing.Point(5, 5)
    $lblPreview.Size = New-Object System.Drawing.Size(100, 20)
    $lblPreview.Text = "Script Preview:"
    $lblPreview.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $previewPanel.Controls.Add($lblPreview)
    
    # Copy button
    $btnCopy = New-Object System.Windows.Forms.Button
    $btnCopy.Location = New-Object System.Drawing.Point(($previewPanel.Width - 110), 2)
    $btnCopy.Size = New-Object System.Drawing.Size(100, 25)
    $btnCopy.Text = "Copy to Clipboard"
    $btnCopy.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnCopy.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnCopy.BackColor = [System.Drawing.Color]::FromArgb(155, 89, 182)
    $btnCopy.ForeColor = [System.Drawing.Color]::White
    $btnCopy.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCopy.Add_Click({
        Copy-ScriptToClipboard
    })
    $previewPanel.Controls.Add($btnCopy)
    
    # Preview textbox
    $script:ScriptPreviewBox = New-Object System.Windows.Forms.TextBox
    $script:ScriptPreviewBox.Location = New-Object System.Drawing.Point(5, 30)
    $script:ScriptPreviewBox.Size = New-Object System.Drawing.Size(($previewPanel.Width - 10), ($previewPanel.Height - 35))
    $script:ScriptPreviewBox.Multiline = $true
    $script:ScriptPreviewBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $script:ScriptPreviewBox.ReadOnly = $true
    $script:ScriptPreviewBox.Font = New-Object System.Drawing.Font("Consolas", 10)
    $script:ScriptPreviewBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                                      [System.Windows.Forms.AnchorStyles]::Bottom -bor 
                                      [System.Windows.Forms.AnchorStyles]::Left -bor 
                                      [System.Windows.Forms.AnchorStyles]::Right
    $previewPanel.Controls.Add($script:ScriptPreviewBox)
    
    $splitContainer.Panel2.Controls.Add($previewPanel)
    
    $Panel.Controls.Add($splitContainer)
}

function Update-ScriptLibrary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Scripts
    )
    
    try {
        $script:Scripts = $Scripts
        $script:ScriptListView.Items.Clear()
        
        foreach ($scriptItem in $Scripts) {
            $item = New-Object System.Windows.Forms.ListViewItem($scriptItem.name)
            $item.SubItems.Add($scriptItem.category) | Out-Null
            $item.SubItems.Add($scriptItem.description) | Out-Null
            $item.SubItems.Add($scriptItem.author) | Out-Null
            $item.SubItems.Add($scriptItem.dateAdded) | Out-Null
            $item.Tag = $scriptItem.id
            
            $script:ScriptListView.Items.Add($item) | Out-Null
        }
        
        Write-Verbose "Loaded $($Scripts.Count) scripts into library"
    }
    catch {
        Write-Error "Failed to update script library: $_"
    }
}

function Filter-Scripts {
    [CmdletBinding()]
    param()
    
    $searchText = $script:SearchBox.Text.ToLower()
    
    if ([string]::IsNullOrEmpty($searchText)) {
        # Show all scripts
        foreach ($item in $script:ScriptListView.Items) {
            $item.Remove()
        }
        Update-ScriptLibrary -Scripts $script:Scripts
        return
    }
    
    # Filter scripts
    $script:ScriptListView.Items.Clear()
    
    foreach ($scriptItem in $script:Scripts) {
        if ($scriptItem.name.ToLower().Contains($searchText) -or 
            $scriptItem.description.ToLower().Contains($searchText) -or
            $scriptItem.category.ToLower().Contains($searchText)) {
            
            $item = New-Object System.Windows.Forms.ListViewItem($scriptItem.name)
            $item.SubItems.Add($scriptItem.category) | Out-Null
            $item.SubItems.Add($scriptItem.description) | Out-Null
            $item.SubItems.Add($scriptItem.author) | Out-Null
            $item.SubItems.Add($scriptItem.dateAdded) | Out-Null
            $item.Tag = $scriptItem.id
            
            $script:ScriptListView.Items.Add($item) | Out-Null
        }
    }
}

function Display-ScriptPreview {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Script
    )
    
    $preview = @"
# Script: $($Script.name)
# Category: $($Script.category)
# Description: $($Script.description)
# Author: $($Script.author)
# Date Added: $($Script.dateAdded)
# ----------------------------------------

$($Script.content)
"@
    
    $script:ScriptPreviewBox.Text = $preview
}

function Copy-ScriptToClipboard {
    [CmdletBinding()]
    param()
    
    if ($script:ScriptListView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a script to copy.",
            "No Selection",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }
    
    $selectedItem = $script:ScriptListView.SelectedItems[0]
    $scriptId = $selectedItem.Tag
    $script = $script:Scripts | Where-Object { $_.id -eq $scriptId }
    
    if ($script) {
        [System.Windows.Forms.Clipboard]::SetText($script.content)
        $script:MainStatusLabel.Text = "Script copied to clipboard: $($script.name)"
        Write-Log "Script copied to clipboard: $($script.name)" -Level INFO
    }
}

function Show-AddScriptDialog {
    [CmdletBinding()]
    param()
    
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Add New Script"
    $dialog.Size = New-Object System.Drawing.Size(600, 500)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    
    # Name
    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Location = New-Object System.Drawing.Point(10, 15)
    $lblName.Size = New-Object System.Drawing.Size(100, 20)
    $lblName.Text = "Script Name:"
    $dialog.Controls.Add($lblName)
    
    $txtName = New-Object System.Windows.Forms.TextBox
    $txtName.Location = New-Object System.Drawing.Point(120, 12)
    $txtName.Size = New-Object System.Drawing.Size(450, 23)
    $dialog.Controls.Add($txtName)
    
    # Category
    $lblCategory = New-Object System.Windows.Forms.Label
    $lblCategory.Location = New-Object System.Drawing.Point(10, 45)
    $lblCategory.Size = New-Object System.Drawing.Size(100, 20)
    $lblCategory.Text = "Category:"
    $dialog.Controls.Add($lblCategory)
    
    $cmbCategory = New-Object System.Windows.Forms.ComboBox
    $cmbCategory.Location = New-Object System.Drawing.Point(120, 42)
    $cmbCategory.Size = New-Object System.Drawing.Size(200, 23)
    $cmbCategory.Items.AddRange(@("Diagnostics", "Services", "Network", "Storage", "Security", "Performance", "Other"))
    $dialog.Controls.Add($cmbCategory)
    
    # Description
    $lblDescription = New-Object System.Windows.Forms.Label
    $lblDescription.Location = New-Object System.Drawing.Point(10, 75)
    $lblDescription.Size = New-Object System.Drawing.Size(100, 20)
    $lblDescription.Text = "Description:"
    $dialog.Controls.Add($lblDescription)
    
    $txtDescription = New-Object System.Windows.Forms.TextBox
    $txtDescription.Location = New-Object System.Drawing.Point(120, 72)
    $txtDescription.Size = New-Object System.Drawing.Size(450, 50)
    $txtDescription.Multiline = $true
    $dialog.Controls.Add($txtDescription)
    
    # Author
    $lblAuthor = New-Object System.Windows.Forms.Label
    $lblAuthor.Location = New-Object System.Drawing.Point(10, 130)
    $lblAuthor.Size = New-Object System.Drawing.Size(100, 20)
    $lblAuthor.Text = "Author:"
    $dialog.Controls.Add($lblAuthor)
    
    $txtAuthor = New-Object System.Windows.Forms.TextBox
    $txtAuthor.Location = New-Object System.Drawing.Point(120, 127)
    $txtAuthor.Size = New-Object System.Drawing.Size(200, 23)
    $txtAuthor.Text = $env:USERNAME
    $dialog.Controls.Add($txtAuthor)
    
    # Script Content
    $lblContent = New-Object System.Windows.Forms.Label
    $lblContent.Location = New-Object System.Drawing.Point(10, 160)
    $lblContent.Size = New-Object System.Drawing.Size(100, 20)
    $lblContent.Text = "Script Code:"
    $dialog.Controls.Add($lblContent)
    
    $txtContent = New-Object System.Windows.Forms.TextBox
    $txtContent.Location = New-Object System.Drawing.Point(10, 185)
    $txtContent.Size = New-Object System.Drawing.Size(560, 220)
    $txtContent.Multiline = $true
    $txtContent.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $txtContent.Font = New-Object System.Drawing.Font("Consolas", 9)
    $dialog.Controls.Add($txtContent)
    
    # Buttons
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Location = New-Object System.Drawing.Point(400, 420)
    $btnSave.Size = New-Object System.Drawing.Size(80, 25)
    $btnSave.Text = "Save"
    $btnSave.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dialog.Controls.Add($btnSave)
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(490, 420)
    $btnCancel.Size = New-Object System.Drawing.Size(80, 25)
    $btnCancel.Text = "Cancel"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dialog.Controls.Add($btnCancel)
    
    $result = $dialog.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        if ([string]::IsNullOrEmpty($txtName.Text) -or [string]::IsNullOrEmpty($txtContent.Text)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Script name and content are required.",
                "Validation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }
        
        $newScript = @{
            name = $txtName.Text
            category = if ($cmbCategory.SelectedItem) { $cmbCategory.SelectedItem } else { "Other" }
            description = $txtDescription.Text
            author = $txtAuthor.Text
            content = $txtContent.Text
        }
        
        $configPath = Join-Path $script:ScriptRoot "Config\ScriptLibrary.json"
        if (Add-ScriptToLibrary -ConfigPath $configPath -ScriptData $newScript) {
            $script:MainStatusLabel.Text = "Script added successfully"
            
            # Reload library
            $scripts = Get-ScriptLibraryConfig -ConfigPath $configPath
            if ($scripts) {
                Update-ScriptLibrary -Scripts $scripts
            }
        }
    }
}

function Show-EditScriptDialog {
    [CmdletBinding()]
    param()
    
    if ($script:ScriptListView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a script to edit.",
            "No Selection",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }
    
    $selectedItem = $script:ScriptListView.SelectedItems[0]
    $scriptId = $selectedItem.Tag
    $scriptToEdit = $script:Scripts | Where-Object { $_.id -eq $scriptId }
    
    if (!$scriptToEdit) {
        return
    }
    
    # Create dialog (similar to Add dialog but pre-filled)
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Edit Script"
    $dialog.Size = New-Object System.Drawing.Size(600, 500)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    
    # Name
    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Location = New-Object System.Drawing.Point(10, 15)
    $lblName.Size = New-Object System.Drawing.Size(100, 20)
    $lblName.Text = "Script Name:"
    $dialog.Controls.Add($lblName)
    
    $txtName = New-Object System.Windows.Forms.TextBox
    $txtName.Location = New-Object System.Drawing.Point(120, 12)
    $txtName.Size = New-Object System.Drawing.Size(450, 23)
    $txtName.Text = $scriptToEdit.name
    $dialog.Controls.Add($txtName)
    
    # Category
    $lblCategory = New-Object System.Windows.Forms.Label
    $lblCategory.Location = New-Object System.Drawing.Point(10, 45)
    $lblCategory.Size = New-Object System.Drawing.Size(100, 20)
    $lblCategory.Text = "Category:"
    $dialog.Controls.Add($lblCategory)
    
    $cmbCategory = New-Object System.Windows.Forms.ComboBox
    $cmbCategory.Location = New-Object System.Drawing.Point(120, 42)
    $cmbCategory.Size = New-Object System.Drawing.Size(200, 23)
    $cmbCategory.Items.AddRange(@("Diagnostics", "Services", "Network", "Storage", "Security", "Performance", "Other"))
    $cmbCategory.Text = $scriptToEdit.category
    $dialog.Controls.Add($cmbCategory)
    
    # Description
    $lblDescription = New-Object System.Windows.Forms.Label
    $lblDescription.Location = New-Object System.Drawing.Point(10, 75)
    $lblDescription.Size = New-Object System.Drawing.Size(100, 20)
    $lblDescription.Text = "Description:"
    $dialog.Controls.Add($lblDescription)
    
    $txtDescription = New-Object System.Windows.Forms.TextBox
    $txtDescription.Location = New-Object System.Drawing.Point(120, 72)
    $txtDescription.Size = New-Object System.Drawing.Size(450, 50)
    $txtDescription.Multiline = $true
    $txtDescription.Text = $scriptToEdit.description
    $dialog.Controls.Add($txtDescription)
    
    # Author
    $lblAuthor = New-Object System.Windows.Forms.Label
    $lblAuthor.Location = New-Object System.Drawing.Point(10, 130)
    $lblAuthor.Size = New-Object System.Drawing.Size(100, 20)
    $lblAuthor.Text = "Author:"
    $dialog.Controls.Add($lblAuthor)
    
    $txtAuthor = New-Object System.Windows.Forms.TextBox
    $txtAuthor.Location = New-Object System.Drawing.Point(120, 127)
    $txtAuthor.Size = New-Object System.Drawing.Size(200, 23)
    $txtAuthor.Text = $scriptToEdit.author
    $dialog.Controls.Add($txtAuthor)
    
    # Script Content
    $lblContent = New-Object System.Windows.Forms.Label
    $lblContent.Location = New-Object System.Drawing.Point(10, 160)
    $lblContent.Size = New-Object System.Drawing.Size(100, 20)
    $lblContent.Text = "Script Code:"
    $dialog.Controls.Add($lblContent)
    
    $txtContent = New-Object System.Windows.Forms.TextBox
    $txtContent.Location = New-Object System.Drawing.Point(10, 185)
    $txtContent.Size = New-Object System.Drawing.Size(560, 220)
    $txtContent.Multiline = $true
    $txtContent.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $txtContent.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtContent.Text = $scriptToEdit.content
    $dialog.Controls.Add($txtContent)
    
    # Buttons
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Location = New-Object System.Drawing.Point(400, 420)
    $btnSave.Size = New-Object System.Drawing.Size(80, 25)
    $btnSave.Text = "Save"
    $btnSave.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dialog.Controls.Add($btnSave)
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(490, 420)
    $btnCancel.Size = New-Object System.Drawing.Size(80, 25)
    $btnCancel.Text = "Cancel"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dialog.Controls.Add($btnCancel)
    
    $result = $dialog.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $updatedScript = @{
            name = $txtName.Text
            category = if ($cmbCategory.SelectedItem) { $cmbCategory.SelectedItem } else { "Other" }
            description = $txtDescription.Text
            author = $txtAuthor.Text
            content = $txtContent.Text
        }
        
        $configPath = Join-Path $script:ScriptRoot "Config\ScriptLibrary.json"
        if (Update-ScriptInLibrary -ConfigPath $configPath -ScriptId $scriptId -UpdatedData $updatedScript) {
            $script:MainStatusLabel.Text = "Script updated successfully"
            
            # Reload library
            $scripts = Get-ScriptLibraryConfig -ConfigPath $configPath
            if ($scripts) {
                Update-ScriptLibrary -Scripts $scripts
            }
        }
    }
}

function Remove-SelectedScript {
    [CmdletBinding()]
    param()
    
    if ($script:ScriptListView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a script to delete.",
            "No Selection",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }
    
    $selectedItem = $script:ScriptListView.SelectedItems[0]
    $scriptId = $selectedItem.Tag
    $scriptName = $selectedItem.Text
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to delete the script '$scriptName'?",
        "Confirm Delete",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $configPath = Join-Path $script:ScriptRoot "Config\ScriptLibrary.json"
        if (Remove-ScriptFromLibrary -ConfigPath $configPath -ScriptId $scriptId) {
            $script:MainStatusLabel.Text = "Script deleted successfully"
            
            # Reload library
            $scripts = Get-ScriptLibraryConfig -ConfigPath $configPath
            if ($scripts) {
                Update-ScriptLibrary -Scripts $scripts
            }
            
            # Clear preview
            $script:ScriptPreviewBox.Clear()
        }
    }
}

# Export only the functions called from main script
Export-ModuleMember -Function @(
    'Initialize-ScriptLibraryPanel',
    'Update-ScriptLibrary'
)