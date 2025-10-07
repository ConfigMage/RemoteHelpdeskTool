# RemoteConnection.psm1
# Handles credential management and remote connections


# Script-level variable to store credentials
$script:CredentialContext = [pscustomobject]@{
    Credential = $null
    AuthenticationMethod = 'Unknown'
}
$script:ConnectedComputers = @{}

function Set-RemoteCredential {
    [CmdletBinding()]
    param(
        [Parameter()]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter()]
        [ValidateSet('UsernamePassword','SmartCard','CurrentUser','Unknown')]
        [string]$AuthenticationMethod = 'Unknown'
    )

    if ($AuthenticationMethod -ne 'CurrentUser' -and $null -eq $Credential) {
        throw "Credential cannot be null when AuthenticationMethod is $AuthenticationMethod."
    }

    if ($AuthenticationMethod -eq 'CurrentUser') {
        $Credential = [System.Management.Automation.PSCredential]::Empty
    }

    $script:CredentialContext = [pscustomobject]@{
        Credential = $Credential
        AuthenticationMethod = $AuthenticationMethod
    }

    Write-Verbose "Credentials cached for session using method: $AuthenticationMethod"
}

function Get-RemoteCredential {
    [CmdletBinding()]
    param()

    if ($null -eq $script:CredentialContext) {
        throw "No credentials cached. Please authenticate first."
    }

    $credential = $script:CredentialContext.Credential

    if ($credential -eq [System.Management.Automation.PSCredential]::Empty) {
        return $null  # This will use current user context
    }

    return $credential
}

function Get-RemoteAuthenticationMethod {
    [CmdletBinding()]
    param()

    if ($null -eq $script:CredentialContext) {
        return 'Unknown'
    }

    return $script:CredentialContext.AuthenticationMethod
}
function Get-RemoteAuthenticationSummary {
    [CmdletBinding()]
    param()

    $method = Get-RemoteAuthenticationMethod

    switch ($method) {
        'SmartCard' { 'Auth: Smart Card' }
        'CurrentUser' { 'Auth: Current User' }
        'UsernamePassword' { 'Auth: Username/Password' }
        default { 'Auth: Unknown' }
    }
}

function Invoke-WmiObjectWithCredentials {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $true)]
        [string]$Namespace,
        
        [Parameter()]
        [string]$Query,
        
        [Parameter()]
        [string]$Class,
        
        [Parameter()]
        [string]$Filter,
        
        [Parameter()]
        [hashtable]$Property
    )
    
    try {
        $credential = Get-RemoteCredential
        $params = @{
            ComputerName = $ComputerName
            Namespace = $Namespace
            ErrorAction = 'Stop'
        }
        
        # Add credential only if not using current user
        if ($null -ne $credential) {
            $params['Credential'] = $credential
        }
        
        # Add appropriate parameters
        if ($Query) {
            $params['Query'] = $Query
        }
        if ($Class) {
            $params['Class'] = $Class
        }
        if ($Filter) {
            $params['Filter'] = $Filter
        }
        if ($Property) {
            $params['Property'] = $Property
        }
        
        return Get-WmiObject @params
    }
    catch {
        Write-Error "WMI operation failed: $_"
        return $null
    }
}

function Invoke-WmiMethodWithCredentials {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $true)]
        [string]$Namespace,
        
        [Parameter(Mandatory = $true)]
        [string]$Class,
        
        [Parameter(Mandatory = $true)]
        [string]$Name,
        
        [Parameter()]
        [array]$ArgumentList
    )
    
    try {
        $credential = Get-RemoteCredential
        $params = @{
            ComputerName = $ComputerName
            Namespace = $Namespace
            Class = $Class
            Name = $Name
            ErrorAction = 'Stop'
        }
        
        # Add credential only if not using current user
        if ($null -ne $credential) {
            $params['Credential'] = $credential
        }
        
        if ($ArgumentList) {
            $params['ArgumentList'] = $ArgumentList
        }
        
        return Invoke-WmiMethod @params
    }
    catch {
        Write-Error "WMI method invocation failed: $_"
        return $null
    }
}

function Test-RemoteConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter()]
        [int]$TimeoutSeconds = 5
    )
    
    try {
        # Quick ping test
        $pingResult = Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction SilentlyContinue
        
        if ($pingResult) {
            # Try to resolve IP address
            try {
                $ipAddress = [System.Net.Dns]::GetHostAddresses($ComputerName) | 
                    Where-Object { $_.AddressFamily -eq 'InterNetwork' } | 
                    Select-Object -First 1 -ExpandProperty IPAddressToString
            }
            catch {
                $ipAddress = "Unable to resolve"
            }
            
            # Cache connection info
            $script:ConnectedComputers[$ComputerName] = @{
                LastChecked = Get-Date
                IsOnline = $true
                IPAddress = $ipAddress
            }
            
            return @{
                IsOnline = $true
                IPAddress = $ipAddress
                ResponseTime = (Test-Connection -ComputerName $ComputerName -Count 1).ResponseTime
            }
        }
        else {
            $script:ConnectedComputers[$ComputerName] = @{
                LastChecked = Get-Date
                IsOnline = $false
                IPAddress = $null
            }
            
            return @{
                IsOnline = $false
                IPAddress = $null
                ResponseTime = $null
            }
        }
    }
    catch {
        Write-Error "Failed to test connection to $ComputerName : $_"
        return @{
            IsOnline = $false
            IPAddress = $null
            ResponseTime = $null
            Error = $_.Exception.Message
        }
    }
}

function Invoke-RemoteCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,
        
        [Parameter()]
        [hashtable]$ArgumentList = @{},
        
        [Parameter()]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        # Use cached credentials if not provided
        if ($null -eq $Credential) {
            $Credential = Get-RemoteCredential
        }
        
        # Test connection first
        $connectionTest = Test-RemoteConnection -ComputerName $ComputerName
        if (!$connectionTest.IsOnline) {
            throw "Computer $ComputerName is not reachable"
        }
        
        # Create session
        $sessionParams = @{
            ComputerName = $ComputerName
            Credential = $Credential
            ErrorAction = 'Stop'
        }
        
        $session = New-PSSession @sessionParams
        
        try {
            # Invoke command
            $result = Invoke-Command -Session $session -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList
            return @{
                Success = $true
                Result = $result
                Error = $null
            }
        }
        finally {
            Remove-PSSession -Session $session -ErrorAction SilentlyContinue
        }
    }
    catch {
        return @{
            Success = $false
            Result = $null
            Error = $_.Exception.Message
        }
    }
}

function Invoke-RemoteScript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,
        
        [Parameter()]
        [hashtable]$Parameters = @{},
        
        [Parameter()]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        # Validate script exists
        if (!(Test-Path $ScriptPath)) {
            throw "Script not found: $ScriptPath"
        }
        
        # Read script content
        $scriptContent = Get-Content -Path $ScriptPath -Raw
        $scriptBlock = [scriptblock]::Create($scriptContent)
        
        # Use cached credentials if not provided
        if ($null -eq $Credential) {
            $Credential = Get-RemoteCredential
        }
        
        # Add computer name to parameters
        $Parameters['ComputerName'] = $ComputerName
        $Parameters['Credential'] = $Credential
        
        # Execute script
        $result = Invoke-RemoteCommand -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $Parameters -Credential $Credential
        
        return $result
    }
    catch {
        return @{
            Success = $false
            Result = $null
            Error = $_.Exception.Message
        }
    }
}

function Get-RemoteSystemInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter()]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    $scriptBlock = {
        $os = Get-WmiObject Win32_OperatingSystem
        $cs = Get-WmiObject Win32_ComputerSystem
        $cpu = Get-WmiObject Win32_Processor | Select-Object -First 1
        $mem = Get-WmiObject Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum
        
        @{
            ComputerName = $env:COMPUTERNAME
            Domain = $cs.Domain
            Manufacturer = $cs.Manufacturer
            Model = $cs.Model
            OS = $os.Caption
            OSVersion = $os.Version
            ServicePack = $os.ServicePackMajorVersion
            Architecture = $os.OSArchitecture
            LastBoot = $os.ConvertToDateTime($os.LastBootUpTime)
            LoggedOnUser = (Get-WmiObject Win32_ComputerSystem).UserName
            CPU = $cpu.Name
            CPUCores = $cpu.NumberOfCores
            TotalMemoryGB = [Math]::Round($mem.Sum / 1GB, 2)
            FreeMemoryGB = [Math]::Round($os.FreePhysicalMemory / 1MB, 2)
        }
    }
    
    try {
        # Use cached credentials if not provided
        if ($null -eq $Credential) {
            $Credential = Get-RemoteCredential
        }
        
        $result = Invoke-RemoteCommand -ComputerName $ComputerName -ScriptBlock $scriptBlock -Credential $Credential
        
        if ($result.Success) {
            return $result.Result
        }
        else {
            throw $result.Error
        }
    }
    catch {
        Write-Error "Failed to get system info from $ComputerName : $_"
        return $null
    }
}

function Start-RemoteProcess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $true)]
        [string]$ProcessPath,
        
        [Parameter()]
        [string]$Arguments = "",
        
        [Parameter()]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        # Use cached credentials if not provided
        if ($null -eq $Credential) {
            $Credential = Get-RemoteCredential
        }
        
        $scriptBlock = {
            param($Path, $Args)
            
            if ($Args) {
                Start-Process -FilePath $Path -ArgumentList $Args -PassThru
            }
            else {
                Start-Process -FilePath $Path -PassThru
            }
        }
        
        $result = Invoke-RemoteCommand -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList @{Path = $ProcessPath; Args = $Arguments} -Credential $Credential
        
        if ($result.Success) {
            return $result.Result
        }
        else {
            throw $result.Error
        }
    }
    catch {
        Write-Error "Failed to start process on $ComputerName : $_"
        return $null
    }
}

function Get-RemoteService {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter()]
        [string]$ServiceName = "*",
        
        [Parameter()]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        # Use cached credentials if not provided
        if ($null -eq $Credential) {
            $Credential = Get-RemoteCredential
        }
        
        Get-Service -ComputerName $ComputerName -Name $ServiceName -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to get services from $ComputerName : $_"
        return $null
    }
}

function Restart-RemoteService {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $true)]
        [string]$ServiceName,
        
        [Parameter()]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        # Use cached credentials if not provided
        if ($null -eq $Credential) {
            $Credential = Get-RemoteCredential
        }
        
        $scriptBlock = {
            param($Name)
            Restart-Service -Name $Name -Force -PassThru
        }
        
        $result = Invoke-RemoteCommand -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList @{Name = $ServiceName} -Credential $Credential
        
        if ($result.Success) {
            return $result.Result
        }
        else {
            throw $result.Error
        }
    }
    catch {
        Write-Error "Failed to restart service on $ComputerName : $_"
        return $null
    }
}

function Copy-ToRemoteComputer {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,
        
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,
        
        [Parameter()]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        # Use cached credentials if not provided
        if ($null -eq $Credential) {
            $Credential = Get-RemoteCredential
        }
        
        # Create UNC path
        $remotePath = "\\$ComputerName\$($DestinationPath.Replace(':', '$'))"
        
        # Create PSDrive with credentials
        $driveName = "RemoteCopy" + (Get-Random -Maximum 9999)
        New-PSDrive -Name $driveName -PSProvider FileSystem -Root "\\$ComputerName\C$" -Credential $Credential -ErrorAction Stop | Out-Null
        
        try {
            # Copy file
            Copy-Item -Path $SourcePath -Destination "${driveName}:\$($DestinationPath.Substring(3))" -Force -ErrorAction Stop
            return $true
        }
        finally {
            Remove-PSDrive -Name $driveName -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Error "Failed to copy to $ComputerName : $_"
        return $false
    }
}

function Get-ConnectedComputers {
    [CmdletBinding()]
    param()
    
    return $script:ConnectedComputers
}

function Clear-ConnectionCache {
    [CmdletBinding()]
    param()
    
    $script:ConnectedComputers = @{}
    Write-Verbose "Connection cache cleared"
}

# Export module functions
Export-ModuleMember -Function @(
    'Set-RemoteCredential',
    'Get-RemoteCredential',
    'Get-RemoteAuthenticationMethod',
    'Get-RemoteAuthenticationSummary',
    'Invoke-WmiObjectWithCredentials',
    'Invoke-WmiMethodWithCredentials',
    'Test-RemoteConnection',
    'Invoke-RemoteCommand',
    'Invoke-RemoteScript',
    'Get-RemoteSystemInfo',
    'Start-RemoteProcess',
    'Get-RemoteService',
    'Restart-RemoteService',
    'Copy-ToRemoteComputer',
    'Get-ConnectedComputers',
    'Clear-ConnectionCache'
)
