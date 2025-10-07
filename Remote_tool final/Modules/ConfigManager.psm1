# ConfigManager.psm1
# Handles all configuration file operations

function Get-QuickActionsConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )
    
    try {
        # Check if config file exists
        if (!(Test-Path $ConfigPath)) {
            Write-Warning "Quick Actions config not found. Creating default configuration..."
            New-DefaultQuickActionsConfig -Path $ConfigPath
        }
        
        # Import CSV
        $actions = Import-Csv -Path $ConfigPath -ErrorAction Stop
        
        # Validate required columns
        $requiredColumns = @('Name', 'DisplayName', 'Description', 'ScriptPath', 'ButtonColor', 'RequiresInput')
        foreach ($column in $requiredColumns) {
            if ($actions[0].PSObject.Properties.Name -notcontains $column) {
                throw "Missing required column: $column"
            }
        }
        
        # Validate each action
        $validatedActions = @()
        foreach ($action in $actions) {
            # Convert RequiresInput to boolean
            $action.RequiresInput = [System.Convert]::ToBoolean($action.RequiresInput)
            
            # Validate button color
            if ($action.ButtonColor -notmatch '^#[0-9A-Fa-f]{6}$') {
                Write-Warning "Invalid color format for action '$($action.Name)'. Using default color."
                $action.ButtonColor = "#808080"
            }
            
            $validatedActions += $action
        }
        
        return $validatedActions
    }
    catch {
        Write-Error "Failed to load Quick Actions configuration: $_"
        return $null
    }
}

function New-DefaultQuickActionsConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    
    $defaultConfig = @"
# Quick Actions Configuration File
# Columns: Name,DisplayName,Description,ScriptPath,ButtonColor,RequiresInput
# ButtonColor: Hex color code (e.g., #4A90E2)
# RequiresInput: true/false - whether the action requires user input
Name,DisplayName,Description,ScriptPath,ButtonColor,RequiresInput
"GPReport","Generate GPO Report","Exports Group Policy results to HTML","Scripts\Get-GPReport.ps1","#4A90E2","false"
"InstalledUpdates","Get Updates","Shows recent Windows updates","Scripts\Get-InstalledUpdates.ps1","#2ECC71","false"
"ForceGPUpdate","Force GP Update","Runs gpupdate /force remotely","Scripts\Force-GPUpdate.ps1","#E74C3C","true"
"ClearCache","Clear Teams/Outlook","Cleans application caches","Scripts\Clear-AppCache.ps1","#F39C12","true"
"GetProcesses","Remote Processes","Shows processes by CPU/Memory","Scripts\Get-RemoteProcesses.ps1","#9B59B6","false"
"@
    
    try {
        $defaultConfig | Out-File -FilePath $Path -Encoding UTF8
        Write-Host "Created default Quick Actions configuration at: $Path" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to create default Quick Actions configuration: $_"
    }
}

function Get-ScriptLibraryConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )
    
    try {
        # Check if config file exists
        if (!(Test-Path $ConfigPath)) {
            Write-Warning "Script Library config not found. Creating default configuration..."
            New-DefaultScriptLibraryConfig -Path $ConfigPath
        }
        
        # Read JSON file
        $jsonContent = Get-Content -Path $ConfigPath -Raw -ErrorAction Stop
        $scripts = $jsonContent | ConvertFrom-Json
        
        # Validate each script entry
        $validatedScripts = @()
        foreach ($script in $scripts) {
            # Ensure all required properties exist
            if ($script.PSObject.Properties.Name -contains 'id' -and
                $script.PSObject.Properties.Name -contains 'name' -and
                $script.PSObject.Properties.Name -contains 'description' -and
                $script.PSObject.Properties.Name -contains 'content') {
                $validatedScripts += $script
            }
            else {
                Write-Warning "Skipping invalid script entry: $($script.name)"
            }
        }
        
        return $validatedScripts
    }
    catch {
        Write-Error "Failed to load Script Library configuration: $_"
        return $null
    }
}

function New-DefaultScriptLibraryConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    
    $defaultScripts = @(
        @{
            id = "1"
            name = "Check Service Status"
            description = "Gets status of any Windows service"
            category = "Services"
            content = @'
$serviceName = Read-Host 'Enter service name'
Get-Service -Name $serviceName -ComputerName $computerName | Select-Object Name, Status, StartType, DisplayName
'@
            author = "Admin"
            dateAdded = (Get-Date).ToString("yyyy-MM-dd")
        },
        @{
            id = "2"
            name = "Get Event Log Errors"
            description = "Retrieves last 50 error events from System log"
            category = "Diagnostics"
            content = @'
Get-EventLog -LogName System -EntryType Error -Newest 50 -ComputerName $computerName | 
    Select-Object TimeGenerated, Source, EventID, Message | 
    Out-GridView -Title "System Error Events"
'@
            author = "Admin"
            dateAdded = (Get-Date).ToString("yyyy-MM-dd")
        },
        @{
            id = "3"
            name = "Get Disk Space"
            description = "Shows disk space usage for all drives"
            category = "Storage"
            content = @'
Get-WmiObject Win32_LogicalDisk -ComputerName $computerName -Filter "DriveType=3" | 
    Select-Object DeviceID, 
        @{Name="Size(GB)";Expression={[math]::Round($_.Size/1GB,2)}},
        @{Name="FreeSpace(GB)";Expression={[math]::Round($_.FreeSpace/1GB,2)}},
        @{Name="PercentFree";Expression={[math]::Round(($_.FreeSpace/$_.Size)*100,2)}} |
    Format-Table -AutoSize
'@
            author = "Admin"
            dateAdded = (Get-Date).ToString("yyyy-MM-dd")
        },
        @{
            id = "4"
            name = "Get Network Configuration"
            description = "Displays network adapter configuration"
            category = "Network"
            content = @'
Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $computerName -Filter "IPEnabled=True" | 
    Select-Object Description, IPAddress, IPSubnet, DefaultIPGateway, DNSServerSearchOrder, MACAddress |
    Format-List
'@
            author = "Admin"
            dateAdded = (Get-Date).ToString("yyyy-MM-dd")
        },
        @{
            id = "5"
            name = "Restart Service"
            description = "Restarts a specified Windows service"
            category = "Services"
            content = @'
$serviceName = Read-Host 'Enter service name to restart'
$service = Get-Service -Name $serviceName -ComputerName $computerName -ErrorAction SilentlyContinue
if ($service) {
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to restart the service '$serviceName'?",
        "Confirm Service Restart",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($confirmation -eq 'Yes') {
        Restart-Service -Name $serviceName -Force
        Write-Host "Service '$serviceName' has been restarted successfully." -ForegroundColor Green
    }
} else {
    Write-Warning "Service '$serviceName' not found on $computerName"
}
'@
            author = "Admin"
            dateAdded = (Get-Date).ToString("yyyy-MM-dd")
        }
    )
    
    try {
        $defaultScripts | ConvertTo-Json -Depth 10 | Out-File -FilePath $Path -Encoding UTF8
        Write-Host "Created default Script Library configuration at: $Path" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to create default Script Library configuration: $_"
    }
}

function Add-ScriptToLibrary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$ScriptData
    )
    
    try {
        # Load existing scripts
        $scripts = Get-ScriptLibraryConfig -ConfigPath $ConfigPath
        if ($null -eq $scripts) {
            $scripts = @()
        }
        
        # Generate new ID
        $newId = ([int]($scripts | ForEach-Object { $_.id } | Measure-Object -Maximum).Maximum + 1).ToString()
        $ScriptData.id = $newId
        $ScriptData.dateAdded = (Get-Date).ToString("yyyy-MM-dd")
        
        # Add to collection
        $scripts += $ScriptData
        
        # Save back to file
        $scripts | ConvertTo-Json -Depth 10 | Out-File -FilePath $ConfigPath -Encoding UTF8
        
        Write-Host "Script added successfully with ID: $newId" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to add script to library: $_"
        return $false
    }
}

function Update-ScriptInLibrary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath,
        
        [Parameter(Mandatory = $true)]
        [string]$ScriptId,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$UpdatedData
    )
    
    try {
        # Load existing scripts
        $scripts = Get-ScriptLibraryConfig -ConfigPath $ConfigPath
        if ($null -eq $scripts) {
            throw "No scripts found in library"
        }
        
        # Find and update the script
        $scriptFound = $false
        for ($i = 0; $i -lt $scripts.Count; $i++) {
            if ($scripts[$i].id -eq $ScriptId) {
                # Update properties
                foreach ($key in $UpdatedData.Keys) {
                    if ($key -ne 'id') {  # Don't allow ID changes
                        $scripts[$i].$key = $UpdatedData[$key]
                    }
                }
                $scriptFound = $true
                break
            }
        }
        
        if (!$scriptFound) {
            throw "Script with ID $ScriptId not found"
        }
        
        # Save back to file
        $scripts | ConvertTo-Json -Depth 10 | Out-File -FilePath $ConfigPath -Encoding UTF8
        
        Write-Host "Script updated successfully" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to update script in library: $_"
        return $false
    }
}

function Remove-ScriptFromLibrary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath,
        
        [Parameter(Mandatory = $true)]
        [string]$ScriptId
    )
    
    try {
        # Load existing scripts
        $scripts = Get-ScriptLibraryConfig -ConfigPath $ConfigPath
        if ($null -eq $scripts) {
            throw "No scripts found in library"
        }
        
        # Filter out the script to remove
        $filteredScripts = $scripts | Where-Object { $_.id -ne $ScriptId }
        
        if ($filteredScripts.Count -eq $scripts.Count) {
            throw "Script with ID $ScriptId not found"
        }
        
        # Save back to file
        $filteredScripts | ConvertTo-Json -Depth 10 | Out-File -FilePath $ConfigPath -Encoding UTF8
        
        Write-Host "Script removed successfully" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to remove script from library: $_"
        return $false
    }
}

function Test-ConfigurationFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )
    
    $results = @{
        QuickActions = $false
        ScriptLibrary = $false
        Errors = @()
    }
    
    # Test Quick Actions config
    $quickActionsPath = Join-Path $ConfigPath "QuickActions.csv"
    try {
        $actions = Get-QuickActionsConfig -ConfigPath $quickActionsPath
        if ($actions) {
            $results.QuickActions = $true
        }
    }
    catch {
        $results.Errors += "Quick Actions: $_"
    }
    
    # Test Script Library config
    $scriptLibraryPath = Join-Path $ConfigPath "ScriptLibrary.json"
    try {
        $scripts = Get-ScriptLibraryConfig -ConfigPath $scriptLibraryPath
        if ($scripts) {
            $results.ScriptLibrary = $true
        }
    }
    catch {
        $results.Errors += "Script Library: $_"
    }
    
    return $results
}

# Export module functions
Export-ModuleMember -Function @(
    'Get-QuickActionsConfig',
    'New-DefaultQuickActionsConfig',
    'Get-ScriptLibraryConfig',
    'New-DefaultScriptLibraryConfig',
    'Add-ScriptToLibrary',
    'Update-ScriptInLibrary',
    'Remove-ScriptFromLibrary',
    'Test-ConfigurationFiles'
)