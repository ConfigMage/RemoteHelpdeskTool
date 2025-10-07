#Requires -Version 5.1
# NOTE: Administrator privileges recommended for remote operations

<#
.SYNOPSIS
    Enhanced Remote Help Desk Tool - Simple Version
.DESCRIPTION
    A comprehensive help desk toolkit with remote software installation, quick actions, and script library
.AUTHOR
    Help Desk Team
.VERSION
    1.0.0
#>

[CmdletBinding()]
param()

# Set strict mode for better error handling
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Get script directory
if ($PSScriptRoot) {
    $script:ScriptRoot = $PSScriptRoot
}
elseif ($PSCommandPath) {
    $script:ScriptRoot = Split-Path -Parent $PSCommandPath
}
elseif ($MyInvocation.ScriptName) {
    $script:ScriptRoot = Split-Path -Parent $MyInvocation.ScriptName
}
elseif ($MyInvocation.MyCommand -and $MyInvocation.MyCommand.PSObject.Properties['Path']) {
    $script:ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}
else {
    $script:ScriptRoot = (Get-Location).ProviderPath
    Write-Warning "Unable to determine script root from invocation metadata; using current location: $script:ScriptRoot"
}

# Import required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


# Ensure previous module versions are unloaded so we always import the latest functions
foreach ($moduleName in @("ConfigManager","Logging","RemoteConnection","QuickActions","ScriptLibrary","SoftwareCenter")) {
    Remove-Module -Name $moduleName -ErrorAction SilentlyContinue
}

# Import modules
$moduleFiles = @(
    "ConfigManager.psm1",
    "Logging.psm1",
    "RemoteConnection.psm1",
    "QuickActions.psm1",
    "ScriptLibrary.psm1",
    "SoftwareCenter.psm1"
)

foreach ($module in $moduleFiles) {
    $modulePath = Join-Path -Path "$ScriptRoot\Modules" -ChildPath $module
    if (Test-Path $modulePath) {
        try {
            Import-Module $modulePath -Force -ErrorAction Stop
            Write-Verbose "Successfully imported module: $module"
        }
        catch {
            Write-Error "Failed to import module $module : $_"
            exit 1
        }
    }
    else {
        Write-Warning "Module not found: $modulePath"
    }
}

# Initialize logging
Initialize-Logging -LogPath "$ScriptRoot\Logs"
Write-Log "Starting Remote Help Desk Tool" -Level INFO

# Global variables to store form data
$script:FormData = @{
    ScriptRoot = $ScriptRoot
    Credential = $null
    AuthenticationMethod = 'Unknown'
    AuthenticationSummary = 'Auth: Unknown'
    AuthenticationWarningShown = $false
    StatusLabel = $null
    ConnectionLabel = $null
    SoftwareCenterTab = $null
    QuickActionsTab = $null
    ScriptLibraryTab = $null
}

# Function to prompt for credentials
function Get-AuthenticationSummaryText {
    [CmdletBinding()]
    param()

    switch ($script:FormData.AuthenticationMethod) {
        'SmartCard' { 'Auth: Smart Card' }
        'CurrentUser' { 'Auth: Current User' }
        'UsernamePassword' { 'Auth: Username/Password' }
        default { 'Auth: Unknown' }
    }
}

function Get-UserCredential {
    [CmdletBinding()]
    param()

    $message = 'Authenticate to use the Remote Help Desk Tool. Smart card users can insert their card and select it from the drop-down.'

    try {
        $credential = Get-Credential -Message $message -ErrorAction Stop
    }
    catch {
        Write-Log "Credential prompt failed: $_" -Level ERROR
        return $null
    }

    if ($null -eq $credential) {
        $script:FormData.AuthenticationMethod = 'Unknown'
        $script:FormData.AuthenticationSummary = Get-AuthenticationSummaryText
        return $null
    }

    if ($credential -eq [System.Management.Automation.PSCredential]::Empty) {
        $script:FormData.AuthenticationMethod = 'CurrentUser'
    }
    else {
        $password = $credential.GetNetworkCredential().Password
        if ([string]::IsNullOrWhiteSpace($password)) {
            $script:FormData.AuthenticationMethod = 'SmartCard'
        }
        else {
            $script:FormData.AuthenticationMethod = 'UsernamePassword'
        }
    }

    $script:FormData.AuthenticationSummary = Get-AuthenticationSummaryText
    return $credential
}

# Function to load configurations
function Invoke-LoadConfigurations {
    try {
        # Load Quick Actions configuration
        $quickActionsConfig = Get-QuickActionsConfig -ConfigPath (Join-Path $script:FormData.ScriptRoot "Config\QuickActions.csv")
        if ($quickActionsConfig) {
            Update-QuickActionButtons -Actions $quickActionsConfig
            Write-Log "Loaded Quick Actions configuration" -Level SUCCESS
        }
        
        # Load Script Library configuration
        $scriptLibraryConfig = Get-ScriptLibraryConfig -ConfigPath (Join-Path $script:FormData.ScriptRoot "Config\ScriptLibrary.json")
        if ($scriptLibraryConfig) {
            Update-ScriptLibrary -Scripts $scriptLibraryConfig
            Write-Log "Loaded Script Library configuration" -Level SUCCESS
        }
        
        $script:FormData.StatusLabel.Text = "Configuration loaded successfully"
    }
    catch {
        Write-Log "Failed to load configuration: $_" -Level ERROR
        $script:FormData.StatusLabel.Text = "Configuration load failed"
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to load configuration files: $_",
            "Configuration Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
}

# Function to reload configuration
function Invoke-ReloadConfiguration {
    $script:FormData.StatusLabel.Text = "Reloading configuration..."
    Write-Log "Reloading configuration files" -Level INFO
    
    Invoke-LoadConfigurations
    
    [System.Windows.Forms.MessageBox]::Show(
        "Configuration files have been reloaded successfully.",
        "Configuration Reload",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}

# Create the main form
function New-MainForm {
    # Create main form
    $mainForm = New-Object System.Windows.Forms.Form
    $mainForm.Text = "Remote Help Desk Tool v1.0"
    $mainForm.Size = New-Object System.Drawing.Size(1000, 700)
    $mainForm.StartPosition = "CenterScreen"
    $mainForm.Icon = [System.Drawing.SystemIcons]::Information
    $mainForm.MinimumSize = New-Object System.Drawing.Size(800, 600)
    
    # Create menu bar
    $menuStrip = New-Object System.Windows.Forms.MenuStrip
    
    # File menu
    $fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $fileMenu.Text = "&File"
    
    $reloadConfigItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $reloadConfigItem.Text = "&Reload Configuration"
    $reloadConfigItem.ShortcutKeys = [System.Windows.Forms.Keys]::F5
    $reloadConfigItem.Add_Click({
        Invoke-ReloadConfiguration
    })
    
    $openConfigItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $openConfigItem.Text = "&Open Config Folder"
    $openConfigItem.Add_Click({
        Start-Process explorer.exe (Join-Path $script:FormData.ScriptRoot "Config")
    })
    
    $exitItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $exitItem.Text = "E&xit"
    $exitItem.ShortcutKeys = [System.Windows.Forms.Keys]::Alt, [System.Windows.Forms.Keys]::F4
    $exitItem.Add_Click({
        $mainForm.Close()
    })
    
    $separator = New-Object System.Windows.Forms.ToolStripSeparator
    $fileMenu.DropDownItems.AddRange(@($reloadConfigItem, $openConfigItem, $separator, $exitItem)) | Out-Null
    
    # Help menu
    $helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $helpMenu.Text = "&Help"
    
    $aboutItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $aboutItem.Text = "&About"
    $aboutItem.Add_Click({
        [System.Windows.Forms.MessageBox]::Show(
            "Remote Help Desk Tool v1.0`n`nA comprehensive toolkit for remote system administration",
            "About",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    })
    
    $helpMenu.DropDownItems.Add($aboutItem) | Out-Null
    
    $menuStrip.Items.AddRange(@($fileMenu, $helpMenu)) | Out-Null
    $mainForm.Controls.Add($menuStrip)
    
    # Force layout calculation to get actual menu height
    $mainForm.PerformLayout()
    
    # Create TabControl with proper positioning
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControlY = $menuStrip.Height + 5
    $tabControlHeight = $mainForm.ClientSize.Height - $menuStrip.Height - 35  # Account for status bar
    $tabControl.Location = New-Object System.Drawing.Point(10, $tabControlY)
    $tabControl.Size = New-Object System.Drawing.Size(($mainForm.ClientSize.Width - 20), $tabControlHeight)
    $tabControl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                          [System.Windows.Forms.AnchorStyles]::Bottom -bor 
                          [System.Windows.Forms.AnchorStyles]::Left -bor 
                          [System.Windows.Forms.AnchorStyles]::Right
    
    # Create Software Center Tab
    $softwareCenterTab = New-Object System.Windows.Forms.TabPage
    $softwareCenterTab.Text = "Software Center"
    $softwareCenterTab.UseVisualStyleBackColor = $true
    $script:FormData.SoftwareCenterTab = $softwareCenterTab
    
    # Create Quick Actions Tab
    $quickActionsTab = New-Object System.Windows.Forms.TabPage
    $quickActionsTab.Text = "Quick Actions"
    $quickActionsTab.UseVisualStyleBackColor = $true
    $script:FormData.QuickActionsTab = $quickActionsTab
    
    # Create Script Library Tab
    $scriptLibraryTab = New-Object System.Windows.Forms.TabPage
    $scriptLibraryTab.Text = "Script Library"
    $scriptLibraryTab.UseVisualStyleBackColor = $true
    $script:FormData.ScriptLibraryTab = $scriptLibraryTab
    
    # Add tabs to control
    $tabControl.TabPages.AddRange(@($softwareCenterTab, $quickActionsTab, $scriptLibraryTab)) | Out-Null
    $mainForm.Controls.Add($tabControl)
    
    # Create Status Strip
    $statusStrip = New-Object System.Windows.Forms.StatusStrip
    
    $connectionLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $connectionLabel.Text = "Connection: Not Connected"
    $connectionLabel.BorderSides = [System.Windows.Forms.ToolStripStatusLabelBorderSides]::Right
    $connectionLabel.Width = 200
    $script:FormData.ConnectionLabel = $connectionLabel
    
    $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $statusLabel.Text = "Ready"
    $statusLabel.Spring = $true
    $statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $script:FormData.StatusLabel = $statusLabel
    
    $statusStrip.Items.AddRange(@($connectionLabel, $statusLabel)) | Out-Null
    $mainForm.Controls.Add($statusStrip)
    
    # Form Load event
    $mainForm.Add_Load({
        $credential = Get-UserCredential

        if ($null -eq $credential) {
            [System.Windows.Forms.MessageBox]::Show(
                "Credentials are required to use this tool.",
                "Authentication Required",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            $mainForm.Close()
            return
        }

        $script:FormData.Credential = $credential
        $setRemoteCommand = Get-Command Set-RemoteCredential -ErrorAction SilentlyContinue
        if ($setRemoteCommand -and $setRemoteCommand.Parameters.ContainsKey("AuthenticationMethod")) {
            Set-RemoteCredential -Credential $credential -AuthenticationMethod $script:FormData.AuthenticationMethod
        }
        elseif ($setRemoteCommand) {
            if (-not $script:FormData.AuthenticationWarningShown) {
                Write-Warning "Set-RemoteCredential does not support AuthenticationMethod; caching credential without method metadata."
                $script:FormData.AuthenticationWarningShown = $true
            }
            Set-RemoteCredential -Credential $credential
        }
        else {
            if (-not $script:FormData.AuthenticationWarningShown) {
                Write-Warning "Set-RemoteCredential command not found. Remote operations will prompt for credentials each time."
                $script:FormData.AuthenticationWarningShown = $true
            }
        }

        $authSummary = Get-AuthenticationSummaryText
        $script:FormData.AuthenticationSummary = $authSummary
        if ($script:FormData.ConnectionLabel) {
            $script:FormData.ConnectionLabel.Text = "Connection: Not Connected ($authSummary)"
        }

        # Initialize tabs
        Initialize-SoftwareCenterTab -TabPage $script:FormData.SoftwareCenterTab -ScriptRoot $script:FormData.ScriptRoot -StatusLabel $script:FormData.StatusLabel

        # Initialize Quick Actions tab (now standalone)
        Initialize-QuickActionsPanel -Panel $script:FormData.QuickActionsTab -ScriptRoot $script:FormData.ScriptRoot -StatusLabel $script:FormData.StatusLabel -ConnectionLabel $script:FormData.ConnectionLabel

        # Initialize Script Library tab (now standalone)
        Initialize-ScriptLibraryPanel -Panel $script:FormData.ScriptLibraryTab -ScriptRoot $script:FormData.ScriptRoot -StatusLabel $script:FormData.StatusLabel

        # Load configurations
        Invoke-LoadConfigurations

        $script:FormData.StatusLabel.Text = "Ready"
        Write-Log "Form loaded successfully" -Level SUCCESS
    })
    
    # Form Closing event
    $mainForm.Add_FormClosing({
        Write-Log "Closing Remote Help Desk Tool" -Level INFO
    })
    
    # Ensure we return the form object explicitly
    return $mainForm
}

# Main execution
try {
    Write-Host "Starting Remote Help Desk Tool..." -ForegroundColor Green
    
    # Ensure required directories exist
    $requiredDirs = @("Logs", "Config", "Scripts", "Modules", "Resources")
    foreach ($dir in $requiredDirs) {
        $dirPath = Join-Path $ScriptRoot $dir
        if (!(Test-Path $dirPath)) {
            New-Item -ItemType Directory -Path $dirPath -Force | Out-Null
            Write-Verbose "Created directory: $dirPath"
        }
    }
    
    # Create and show the main form
    $mainForm = New-MainForm
    [System.Windows.Forms.Application]::Run($mainForm)
}
catch {
    Write-Error "Fatal error: $_"
    Write-Log "Fatal error: $_" -Level ERROR
    [System.Windows.Forms.MessageBox]::Show(
        "A fatal error occurred:`n`n$_",
        "Fatal Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}
finally {
    # Cleanup
    Write-Log "Application closed" -Level INFO
}

