<#
.SYNOPSIS
    Windows 7 System Information Collector with Enhanced Network Details
.DESCRIPTION
    Collects essential system information including DHCP/static IP configuration
    Output matches the structure of SystemInfoCollector.ps1
.NOTES
    Author: Lesedi Sebekedi
    Version: 1.1
    Last Updated: $(Get-Date -Format "yyyy-MM-dd")
#>

#region Initialization and Validation
# Check if running on Windows 7 (Version 6.1)
if ((Get-WmiObject Win32_OperatingSystem).Version -notlike "6.1*") {
    Write-Host "This script is designed for Windows 7 only" -ForegroundColor Red
    exit 1
}
#endregion

#region Helper Functions
<#
.SYNOPSIS
    Converts WMI date format to human-readable format
.PARAMETER wmiDate
    The WMI date string to convert
#>
function Convert-WmiDateTime {
    param([string]$wmiDate)
    
    if ([string]::IsNullOrEmpty($wmiDate)) { 
        return $null
    }
    
    try {
        return [Management.ManagementDateTimeConverter]::ToDateTime($wmiDate).ToString('yyyy-MM-dd HH:mm:ss')
    } 
    catch {
        return $null
    }
}
#endregion

#region Data Collection Functions
<#
.SYNOPSIS
    Collects detailed network adapter configuration
.OUTPUTS
    PSCustomObject[] containing network adapter information
#>
function Get-NetworkConfiguration {
    try {
        $adapters = @()
        $nics = Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled }

        foreach ($nic in $nics) {
            $adapters += [PSCustomObject]@{
                Name = $nic.Description
                InterfaceDescription = $nic.Description
                MacAddress = $nic.MACAddress
                Speed = if ($nic.Speed) { "$($nic.Speed / 1MB) Mbps" } else { "N/A" }
                IPAddress = if ($nic.IPAddress) { $nic.IPAddress[0] } else { $null }
                IPConfiguration = if ($nic.DHCPEnabled) { "DHCP" } else { "Static" }
                DHCPServer = if ($nic.DHCPEnabled) { $nic.DHCPServer } else { $null }
                LeaseObtained = if ($nic.DHCPEnabled) { Convert-WmiDateTime $nic.DHCPLeaseObtained } else { $null }
                LeaseExpires = if ($nic.DHCPEnabled) { Convert-WmiDateTime $nic.DHCPLeaseExpires } else { $null }
                DefaultGateway = if ($nic.DefaultIPGateway) { $nic.DefaultIPGateway[0] } else { $null }
                DNSServers = if ($nic.DNSServerSearchOrder) { $nic.DNSServerSearchOrder } else { $null }
            }
        }

        return $adapters
    } 
    catch {
        Write-Warning "Error collecting network information: $_"
        return @()
    }
}

<#
.SYNOPSIS
    Collects comprehensive system information
.OUTPUTS
    PSCustomObject containing system, hardware, network, and software information
#>
function Get-Win7SystemInfo {
    try {
        #region Basic System Information
        $os = Get-WmiObject Win32_OperatingSystem
        $computer = Get-WmiObject Win32_ComputerSystem
        $bios = Get-WmiObject Win32_BIOS
        $product = Get-WmiObject Win32_ComputerSystemProduct
        #endregion

        #region Hardware Information
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
        #endregion

        #region Network Information
        $networkAdapters = Get-NetworkConfiguration
        #endregion

        #region Installed Software
        $installedApps = @()
        $regPaths = @(
            "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        
        foreach ($path in $regPaths) {
            if (Test-Path $path) {
                $apps = Get-ItemProperty $path | Where-Object { $_.DisplayName -and -not $_.SystemComponent }
                $installedApps += $apps | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
            }
        }
        #endregion

        #region Hotfixes
        $hotfixes = @()
        if (Get-Command Get-HotFix -ErrorAction SilentlyContinue) {
            $hotfixes = Get-HotFix -ErrorAction SilentlyContinue | Select-Object HotFixID, Description, InstalledOn
        }
        #endregion

        #region Build Output Object
        $result = [PSCustomObject]@{
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
            Network = if ($networkAdapters.Count -eq 1) { $networkAdapters[0] } else { $networkAdapters }
            Software = @{
                InstalledApps = $installedApps
                Hotfixes     = $hotfixes
            }
            UUID    = $product.UUID
            PSVersion = $PSVersionTable.PSVersion.ToString()
        }
        #endregion

        return $result
    } 
    catch {
        Write-Warning "Error collecting system information: $_"
        return $null
    }
}
#endregion

#region Main Execution
# Asset number input with validation (11 digits starting with 0)
do {
    $assetNumber = Read-Host "Enter Asset Number (11 digits starting with 0)"
} while (-not ($assetNumber -match '^0\d{10}$'))

# Collect system information
$systemInfo = Get-Win7SystemInfo

if ($systemInfo) {
    # Add asset number to collected data
    $systemInfo | Add-Member -NotePropertyName "AssetNumber" -NotePropertyValue $assetNumber -Force

    # Create output directory if it doesn't exist
    $outputDir = "$env:USERPROFILE\Desktop\SystemReports"
    if (-not (Test-Path $outputDir)) { 
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null 
    }

    # Generate output filename with sanitized computer name and timestamp
    $computerName = $env:COMPUTERNAME -replace '[\\/:*?"<>|]', '_'
    $reportPath = "$outputDir\$computerName-$(Get-Date -Format 'yyyyMMdd-HHmmss').json"

    # Export data to JSON file
    #region JSON Conversion (PS 2.0 Compatible)
    function ConvertTo-JsonFallback {
        param($InputObject, [int]$Depth = 5)
        
        if ($PSVersionTable.PSVersion.Major -ge 3) {
            # Use native ConvertTo-Json if available (PS 3.0+)
            return $InputObject | ConvertTo-Json -Depth $Depth
        }
        else {
            # Fallback to custom formatter (simplified version of your original)
            function Format-JsonInner($obj, $indent = 0) {
                $space = " " * $indent
                if ($obj -is [System.Collections.IDictionary]) {
                    $entries = @()
                    foreach ($key in $obj.Keys) {
                        $entries += "$space`"$key`": $(Format-JsonInner $obj[$key] ($indent + 2))"
                    }
                    return "{`n" + ($entries -join ",`n") + "`n$space}"
                }
                elseif ($obj -is [Array]) {
                    $entries = $obj | ForEach-Object { Format-JsonInner $_ ($indent + 2) }
                    return "[`n" + ($entries -join ",`n") + "`n$space]"
                }
                elseif ($obj -is [string]) {
                    return '"' + $obj.Replace('"', '\"') + '"'
                }
                elseif ($null -eq $obj) {
                    return 'null'
                }
                else {
                    return $obj.ToString()
                }
            }
            return Format-JsonInner $InputObject
        }
    }
    #endregion

    # Replace this line in Main Execution:
    # $systemInfo | ConvertTo-Json -Depth 5 | Out-File ...
    # With:
    ConvertTo-JsonFallback -InputObject $systemInfo -Depth 5 | Out-File $reportPath -Force

    # Provide user feedback
    Write-Host "System information successfully saved to:`n$reportPath" -ForegroundColor Green
} 
else {
    Write-Host "Failed to collect system information" -ForegroundColor Red
}
#endregion