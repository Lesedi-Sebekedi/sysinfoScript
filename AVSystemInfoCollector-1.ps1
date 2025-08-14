<#
.SYNOPSIS
    Windows 7 System Information Collector with Enhanced Network Details
.DESCRIPTION
    Collects essential system information including DHCP/static IP configuration
    Output matches the structure of SystemInfoCollector.ps1
    Automatically handles PowerShell version compatibility
.NOTES
    Author: Lesedi Sebekedi
    Version: 1.3
#>

#region Initialization and Configuration
$Script:Config = @{
    NetFx4Url = "https://download.microsoft.com/download/9/5/A/95A9616B-7A37-4AF6-BC36-D6EA96C8DAAE/dotNetFx40_Full_x86_x64.exe"
    PS3InstallerUrl = "https://download.microsoft.com/download/E/7/6/E76850B8-DA6E-4FF5-8CCE-A24FC513FD16/Windows6.1-KB2506143-x64.msu"
    AssetNumberPattern = '^0\d{10}$'
    OutputDir = "$env:USERPROFILE\Desktop\SystemReports"
}
#endregion

#region Core Functions

function Test-NetFramework4Installed {
    try {
        $netReg = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -ErrorAction Stop
        return ($netReg.Release -ge 378389) # .NET 4.5+ also qualifies
    }
    catch { return $false }
}

function Install-NetFramework4 {
    try {
        $installerPath = "$env:TEMP\dotNet40Setup.exe"
        
        # Download if needed
        if (-not (Test-Path $installerPath)) {
            Write-Host "Downloading .NET 4.0 (~48MB)..." -ForegroundColor Cyan
            $ProgressPreference = 'SilentlyContinue' # Faster download
            Invoke-WebRequest -Uri $Script:Config.NetFx4Url -OutFile $installerPath
        }

        # Install silently
        Write-Host "Installing .NET 4.0 (may take 5-10 minutes)..." -ForegroundColor Cyan
        $proc = Start-Process -FilePath $installerPath -ArgumentList "/quiet /norestart" -PassThru -Wait
        
        if ($proc.ExitCode -eq 0) {
            Write-Host ".NET 4.0 installed successfully!" -ForegroundColor Green
            return $true
        }
        else {
            Write-Host "Installation failed (Exit code: $($proc.ExitCode))" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host ".NET 4.0 installation error: $_" -ForegroundColor Red
        return $false
    }
}

function Install-PowerShell3 {
    try {
        $installerPath = "$env:TEMP\PS3Setup.msu"
        
        # Download if needed
        if (-not (Test-Path $installerPath)) {
            Write-Host "Downloading PowerShell 3.0 (~8MB)..." -ForegroundColor Cyan
            Invoke-WebRequest -Uri $Script:Config.PS3InstallerUrl -OutFile $installerPath
        }

        # Install update
        Write-Host "Installing PowerShell 3.0..." -ForegroundColor Cyan
        $proc = Start-Process "wusa.exe" -ArgumentList "$installerPath /quiet /norestart" -PassThru -Wait
        
        if ($proc.ExitCode -eq 0) {
            Write-Host "PowerShell 3.0 installed! Please restart your computer." -ForegroundColor Green
            return $true
        }
        else {
            Write-Host "Installation failed (Exit code: $($proc.ExitCode))" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "PowerShell 3.0 installation error: $_" -ForegroundColor Red
        return $false
    }
}

function Initialize-Environment {
    # Check Windows version
    if ((Get-WmiObject Win32_OperatingSystem).Version -notlike "6.1*") {
        Write-Host "This script is designed for Windows 7 only" -ForegroundColor Red
        exit 1
    }

    # Skip if already on PS 3.0+
    if ($PSVersionTable.PSVersion.Major -ge 3) { return }

    Write-Host "`nWARNING: Running on PowerShell $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    Write-Host "Recommended: Upgrade to PowerShell 3.0 for better performance`n" -ForegroundColor Yellow

    $choice = Read-Host "Automatically install PowerShell 3.0? [Y/N]"
    if ($choice -notmatch '^[Yy]') { return }

    # .NET 4.0 check
    if (-not (Test-NetFramework4Installed)) {
        Write-Host "`n.NET 4.0 is required for PowerShell 3.0" -ForegroundColor Yellow
        $netChoice = Read-Host "Automatically download and install .NET 4.0? (~48MB) [Y/N]"
        
        if ($netChoice -match '^[Yy]') {
            if (-not (Install-NetFramework4)) {
                Write-Host "Continuing with PowerShell 2.0..." -ForegroundColor Yellow
                return
            }
        }
        else {
            Write-Host "PowerShell 3.0 requires .NET 4.0. Continuing with PS 2.0..." -ForegroundColor Yellow
            return
        }
    }

    # Install PS 3.0 if we got this far
    if (Install-PowerShell3) {
        # Exit to let user restart
        exit 0
    }
}

function ConvertTo-JsonCompatible {
    <#
    .SYNOPSIS
        Version-aware JSON conversion with fallback for PS 2.0
    #>
    param($InputObject, [int]$Depth = 5)
    
    if ($PSVersionTable.PSVersion.Major -ge 3) {
        return $InputObject | ConvertTo-Json -Depth $Depth
    }
    else {
        # PS 2.0 fallback formatter
        function Format-JsonInner($obj, $indent = 0) {
            $space = " " * $indent
            switch ($obj) {
                { $_ -is [System.Collections.IDictionary] } {
                    $entries = foreach ($key in $obj.Keys) {
                        "$space`"$key`": $(Format-JsonInner $obj[$key] ($indent + 2))"
                    }
                    return "{`n" + ($entries -join ",`n") + "`n$space}"
                }
                { $_ -is [Array] } {
                    $entries = $obj | ForEach-Object { Format-JsonInner $_ ($indent + 2) }
                    return "[`n" + ($entries -join ",`n") + "`n$space]"
                }
                { $_ -is [string] } { return '"' + $obj.Replace('"', '\"') + '"' }
                { $null -eq $_ } { return 'null' }
                default { return $obj.ToString() }
            }
        }
        return Format-JsonInner $InputObject
    }
}

function Convert-WmiDateTime {
    <#
    .SYNOPSIS
        Converts WMI date format to standard datetime string
    #>
    param([string]$wmiDate)
    
    if ([string]::IsNullOrEmpty($wmiDate)) { return $null }
    
    try {
        return [Management.ManagementDateTimeConverter]::ToDateTime($wmiDate).ToString('yyyy-MM-dd HH:mm:ss')
    } 
    catch {
        return $null
    }
}
#endregion

#region Data Collection Functions

function Get-NetworkConfiguration {
    <#
    .SYNOPSIS
        Collects detailed network adapter configuration
    .OUTPUTS
        PSCustomObject[] containing network adapter information
    #>
    try {
        $adapters = @()
        $nics = Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled }

        foreach ($nic in $nics) {
            $adapterInfo = @{
                Name                = $nic.Description
                InterfaceDescription = $nic.Description
                MacAddress          = $nic.MACAddress
                Speed               = if ($nic.Speed) { "$($nic.Speed / 1MB) Mbps" } else { "N/A" }
                IPAddress           = if ($nic.IPAddress) { $nic.IPAddress[0] } else { $null }
                IPConfiguration     = if ($nic.DHCPEnabled) { "DHCP" } else { "Static" }
                DHCPServer         = if ($nic.DHCPEnabled) { $nic.DHCPServer } else { $null }
                LeaseObtained      = if ($nic.DHCPEnabled) { Convert-WmiDateTime $nic.DHCPLeaseObtained } else { $null }
                LeaseExpires       = if ($nic.DHCPEnabled) { Convert-WmiDateTime $nic.DHCPLeaseExpires } else { $null }
                DefaultGateway     = if ($nic.DefaultIPGateway) { $nic.DefaultIPGateway[0] } else { $null }
                DNSServers         = if ($nic.DNSServerSearchOrder) { $nic.DNSServerSearchOrder } else { $null }
            }
            $adapters += [PSCustomObject]$adapterInfo
        }

        return $adapters
    } 
    catch {
        Write-Warning "Error collecting network information: $_"
        return @()
    }
}

function Get-InstalledSoftware {
    <#
    .SYNOPSIS
        Retrieves list of installed applications from registry
    #>
    $installedApps = @()
    $regPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    foreach ($path in $regPaths) {
        if (Test-Path $path) {
            $apps = Get-ItemProperty $path -ErrorAction SilentlyContinue | 
                   Where-Object { $_.DisplayName -and -not $_.SystemComponent }
            $installedApps += $apps | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
        }
    }
    
    return $installedApps
}

function Get-SystemHotfixes {
    <#
    .SYNOPSIS
        Retrieves installed hotfixes if available
    #>
    if (Get-Command Get-HotFix -ErrorAction SilentlyContinue) {
        return Get-HotFix -ErrorAction SilentlyContinue | Select-Object HotFixID, Description, InstalledOn
    }
    return @()
}

function Get-Win7SystemInfo {
    <#
    .SYNOPSIS
        Collects comprehensive system information
    .OUTPUTS
        PSCustomObject containing system, hardware, network, and software information
    #>
    try {
        # Basic System Information
        $os = Get-WmiObject Win32_OperatingSystem
        $computer = Get-WmiObject Win32_ComputerSystem
        $bios = Get-WmiObject Win32_BIOS
        $product = Get-WmiObject Win32_ComputerSystemProduct

        # Hardware Information
        $cpu = Get-WmiObject Win32_Processor | Select-Object -First 1
        $memory = Get-WmiObject Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum
        $pageFile = Get-WmiObject Win32_PageFileUsage | Select-Object @{Name="PageFileGB"; Expression={[math]::Round($_.AllocatedBaseSize / 1KB, 2)}}
        
        $disks = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } | Select-Object DeviceID, VolumeName,
            @{Name="SizeGB"; Expression={[math]::Round($_.Size / 1GB, 2)}},
            @{Name="FreeGB"; Expression={[math]::Round($_.FreeSpace / 1GB, 2)}},
            @{Name="Type"; Expression={"Fixed"}}
        
        $gpu = Get-WmiObject Win32_VideoController | Select-Object Name,
            @{Name="AdapterRAMGB"; Expression={[math]::Round($_.AdapterRAM / 1GB, 2)}},
            DriverVersion

        # Build output object
        return [PSCustomObject]@{
            System = @{
                HostName     = $env:COMPUTERNAME
                OS           = $os.Caption
                Version      = $os.Version
                Build        = $os.BuildNumber
                Architecture = if ([IntPtr]::Size -eq 8) { "64-bit" } else { "32-bit" }
                BootTime     = Convert-WmiDateTime $os.LastBootUpTime
                Manufacturer = $computer.Manufacturer
                Model        = $computer.Model
                BIOS         = @{
                    Serial = $bios.SerialNumber
                    Version = $bios.SMBIOSBIOSVersion
                }
            }
            Hardware = @{
                CPU    = [PSCustomObject]@{
                    Name = $cpu.Name
                    Cores = $cpu.NumberOfCores
                    Threads = $cpu.NumberOfLogicalProcessors
                    ClockSpeed = "$($cpu.MaxClockSpeed) MHz"
                }
                Memory = @{
                    TotalGB    = [math]::Round($memory.Sum / 1GB, 2)
                    Sticks     = $memory.Count
                    PageFileGB = if ($pageFile) { $pageFile.PageFileGB } else { 0 }
                }
                Disks  = $disks
                GPU    = $gpu
            }
            Network = if (($netAdapters = Get-NetworkConfiguration).Count -eq 1) { $netAdapters[0] } else { $netAdapters }
            Software = @{
                InstalledApps = Get-InstalledSoftware
                Hotfixes     = Get-SystemHotfixes
            }
            UUID    = $product.UUID
            PSVersion = $PSVersionTable.PSVersion.ToString()
        }
    } 
    catch {
        Write-Warning "Error collecting system information: $_"
        return $null
    }
}
#endregion

#region Main Execution
function Main {
    # Initialize environment checks
    Initialize-Environment

    # Asset number input with validation
    do {
        $assetNumber = Read-Host "Enter Asset Number (11 digits starting with 0)"
    } while (-not ($assetNumber -match $Script:Configuration.AssetNumberPattern))

    # Collect system information
    if (-not ($systemInfo = Get-Win7SystemInfo)) {
        Write-Host "Failed to collect system information" -ForegroundColor Red
        exit 1
    }

    # Add asset number to collected data
    $systemInfo | Add-Member -NotePropertyName "AssetNumber" -NotePropertyValue $assetNumber -Force

    # Prepare output directory
    if (-not (Test-Path $Script:Configuration.OutputDirectory)) { 
        New-Item -ItemType Directory -Path $Script:Configuration.OutputDirectory -Force | Out-Null 
    }

    # Generate output filename
    $computerName = $env:COMPUTERNAME -replace '[\\/:*?"<>|]', '_'
    $reportPath = "$($Script:Configuration.OutputDirectory)\$computerName-$(Get-Date -Format 'yyyyMMdd-HHmmss').json"

    # Export data to JSON file
    ConvertTo-JsonCompatible -InputObject $systemInfo -Depth 5 | Out-File $reportPath -Force

    Write-Host "System information successfully saved to:`n$reportPath" -ForegroundColor Green
}

# Execute main function
Main
#endregion