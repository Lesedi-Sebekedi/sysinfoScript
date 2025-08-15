<#
.SYNOPSIS
    Windows 7 System Information Collector with Enhanced Network Details
.DESCRIPTION
    Collects essential system information including DHCP/static IP configuration
    Outputs to both TXT and JSON formats
#>

# Check if running on Windows 7
if ((Get-WmiObject Win32_OperatingSystem).Version -notlike "6.1*") {
    Write-Host "This script is designed for Windows 7 only" -ForegroundColor Red
    exit 1
}

function Convert-WmiDateTime {
    param([string]$wmiDate)
    if ([string]::IsNullOrEmpty($wmiDate)) { return "N/A" }
    try {
        return [Management.ManagementDateTimeConverter]::ToDateTime($wmiDate).ToString('yyyy-MM-dd HH:mm:ss')
    } catch {
        return "InvalidDate"
    }
}

function Get-NetworkConfiguration {
    try {
        $networkAdapters = @()
        $nics = Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled }

        foreach ($nic in $nics) {
            $ipConfigType = if ($nic.DHCPEnabled) { "DHCP" } else { "Static" }

            $adapter = [PSCustomObject]@{
                Name = $nic.Description
                InterfaceDescription = $nic.Description
                MacAddress = $nic.MACAddress
                Speed = if ($nic.Speed) { "$([math]::Round($nic.Speed / 1MB, 2)) Mbps" } else { "N/A" }
                IPAddress = if ($nic.IPAddress) { $nic.IPAddress[0] } else { "N/A" }
                IPConfigType = $ipConfigType
                SubnetMask = if ($nic.IPSubnet) { $nic.IPSubnet[0] } else { "N/A" }
                DefaultGateway = if ($nic.DefaultIPGateway) { $nic.DefaultIPGateway[0] } else { "N/A" }
                DNSServers = if ($nic.DNSServerSearchOrder) { $nic.DNSServerSearchOrder -join ", " } else { "N/A" }
                DHCPEnabled = $nic.DHCPEnabled
                DHCPServer = if ($nic.DHCPServer) { $nic.DHCPServer } else { "N/A" }
                LeaseObtained = if ($nic.DHCPEnabled) { Convert-WmiDateTime $nic.DHCPLeaseObtained } else { "N/A" }
                LeaseExpires = if ($nic.DHCPEnabled) { Convert-WmiDateTime $nic.DHCPLeaseExpires } else { "N/A" }
            }
            $networkAdapters += $adapter
        }

        if ($networkAdapters.Count -eq 1) { return $networkAdapters[0] }
        return $networkAdapters
    }
    catch {
        Write-Warning "Error collecting network information: $_"
        return $null
    }
}

function Get-Win7SystemInfo {
    try {
        # Basic System Information
        $os = Get-WmiObject Win32_OperatingSystem
        $computerSystem = Get-WmiObject Win32_ComputerSystem
        $bios = Get-WmiObject Win32_BIOS
        $systemProduct = Get-WmiObject Win32_ComputerSystemProduct

        # Hardware Information
        $cpu = Get-WmiObject Win32_Processor | Select-Object -First 1 | Select-Object Name,
            @{Name="Cores"; Expression={$_.NumberOfCores}},
            @{Name="Threads"; Expression={$_.NumberOfLogicalProcessors}},
            @{Name="ClockSpeed"; Expression={"$($_.MaxClockSpeed) MHz"}}

        $physicalMemory = Get-WmiObject Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum
        $memorySticks = Get-WmiObject Win32_PhysicalMemory
        $pageFile = Get-WmiObject Win32_PageFileUsage | Select-Object @{Name="PageFileGB"; Expression={[math]::Round($_.AllocatedBaseSize / 1KB, 2)}}
        
        $disks = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } | Select-Object DeviceID, VolumeName,
            @{Name="SizeGB"; Expression={[math]::Round($_.Size / 1GB, 2)}},
            @{Name="FreeGB"; Expression={[math]::Round($_.FreeSpace / 1GB, 2)}},
            @{Name="Type"; Expression={"Fixed"}}

        $gpu = Get-WmiObject Win32_VideoController | Select-Object Name,
            @{Name="AdapterRAMGB"; Expression={[math]::Round($_.AdapterRAM / 1GB, 2)}},
            DriverVersion

        # Network Information
        $networkAdapters = Get-NetworkConfiguration

        # Installed Software
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

        # Hotfixes
        $hotfixes = @()
        if (Get-Command Get-HotFix -ErrorAction SilentlyContinue) {
            $hotfixes = Get-HotFix -ErrorAction SilentlyContinue | Select-Object HotFixID, Description, InstalledOn
        }

        # Build output object matching SystemInfoCollector's structure
        return [PSCustomObject]@{
            System = @{
                HostName     = $env:COMPUTERNAME
                OS           = $os.Caption
                Version      = $os.Version
                Build        = $os.BuildNumber
                Architecture = if ([IntPtr]::Size -eq 8) { "64-bit" } else { "32-bit" }
                BootTime     = Convert-WmiDateTime $os.LastBootUpTime
                Manufacturer = $computerSystem.Manufacturer
                Model        = $computerSystem.Model
                BIOS         = @{
                    Version = $bios.SMBIOSBIOSVersion
                    Serial  = $bios.SerialNumber
                }
            }
            Hardware = @{
                CPU    = $cpu
                Memory = @{
                    TotalGB    = [math]::Round($physicalMemory.Sum / 1GB, 2)
                    Sticks     = $memorySticks.Count
                    PageFileGB = if ($pageFile) { $pageFile.PageFileGB } else { 0 }
                }
                Disks  = $disks
                GPU    = $gpu
            }
            Network = $networkAdapters
            Software = @{
                InstalledApps = $installedApps
                Hotfixes     = $hotfixes
            }
            UUID       = $systemProduct.UUID
            PSVersion  = $PSVersionTable.PSVersion.ToString()
            Timestamp  = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        }
    } catch {
        Write-Warning "Error collecting system information: $_"
        return $null
    }
}

function ConvertTo-TextReport {
    param(
        [Parameter(Mandatory=$true)]
        $SystemInfo
    )
    
    $report = @"
=== SYSTEM INFORMATION REPORT ===
Generated: $($SystemInfo.Timestamp)

ASSET NUMBER: $($SystemInfo.AssetNumber)
COMPUTER NAME: $($SystemInfo.System.HostName)

--- SYSTEM ---
OS: $($SystemInfo.System.OS)
Version: $($SystemInfo.System.Version)
Architecture: $($SystemInfo.System.Architecture)
Manufacturer: $($SystemInfo.System.Manufacturer)
Model: $($SystemInfo.System.Model)
Last Boot: $($SystemInfo.System.BootTime)

--- HARDWARE ---
CPU: $($SystemInfo.Hardware.CPU.Name)
Cores: $($SystemInfo.Hardware.CPU.Cores)
Threads: $($SystemInfo.Hardware.CPU.Threads)
Clock Speed: $($SystemInfo.Hardware.CPU.ClockSpeed)
Memory: $($SystemInfo.Hardware.Memory.TotalGB) GB ($($SystemInfo.Hardware.Memory.Sticks) sticks)
Page File: $($SystemInfo.Hardware.Memory.PageFileGB) GB

Disks:
$($SystemInfo.Hardware.Disks | ForEach-Object {
    "  Drive $($_.DeviceID): $($_.SizeGB)GB total, $($_.FreeGB)GB free ($($_.VolumeName))"
})

GPU: $($SystemInfo.Hardware.GPU.Name)
VRAM: $($SystemInfo.Hardware.GPU.AdapterRAMGB) GB
Driver: $($SystemInfo.Hardware.GPU.DriverVersion)

--- NETWORK ---
$($SystemInfo.Network | ForEach-Object {
    if ($_.Name) {
        "Adapter: $($_.Name)`r`n"
        "IPAddress: $($_.IPAddress)`r`n"
        "SubnetMask: $($_.SubnetMask)`r`n"
        "DefaultGateway: $($_.DefaultGateway)`r`n"
        "MacAddress: $($_.MacAddress)`r`n"
        "DNSServer: $($_.DNSServers)`r`n"
    }
})

--- SOFTWARE ---
Installed Applications:
$($SystemInfo.Software.InstalledApps | ForEach-Object {
    "  $($_.DisplayName) ($($_.DisplayVersion)) by $($_.Publisher)`r`n"
})

Hotfixes:
$($SystemInfo.Software.Hotfixes | ForEach-Object {
    "  $($_.HotFixID) - $($_.Description) (Installed: $($_.InstalledOn))`r`n"
})

--- SYSTEM IDENTIFIERS ---
UUID: $($SystemInfo.UUID)
BIOS Version: $($SystemInfo.System.BIOS.Version)
BIOS Serial: $($SystemInfo.System.BIOS.Serial)

Report generated with PowerShell $($SystemInfo.PSVersion)
"@

    return $report
}

# Main Execution
Write-Host "Windows 7 System Information Collector" -ForegroundColor Cyan
Write-Host "------------------------------------" -ForegroundColor Cyan

# Asset number input with validation
do {
    $assetNumber = Read-Host "Enter Asset Number (11 digits starting with 0)"
} while (-not ($assetNumber -match '^0\d{10}$'))

# Collect system info
$systemInfo = Get-Win7SystemInfo

if ($systemInfo -ne $null) {
    # Add asset number property
    $systemInfo | Add-Member -NotePropertyName "AssetNumber" -NotePropertyValue $assetNumber -Force

    # Create output directory
    $outputDir = "$env:USERPROFILE\Desktop\SystemReports"
    if (-not (Test-Path $outputDir)) {
        try {
            New-Item -ItemType Directory -Path $outputDir -ErrorAction Stop | Out-Null
        } catch {
            Write-Warning "Failed to create directory: $outputDir. $_"
            exit 1
        }
    }

    $timestampSafe = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $computerName = $env:COMPUTERNAME -replace '[\\/:*?"<>|]', '_'

    # Generate and save TXT report
    $txtReport = ConvertTo-TextReport -SystemInfo $systemInfo
    $txtFile = Join-Path $outputDir "${computerName}_SystemInfo_${timestampSafe}.txt"
    $txtReport | Out-File -FilePath $txtFile -Encoding UTF8
    Write-Host "`nText report saved to: $txtFile" -ForegroundColor Green

    # Display report in console
    Write-Host "`n=== SYSTEM INFORMATION SUMMARY ===`n"
    Write-Host $txtReport
}
else {
    Write-Host "Failed to collect system information" -ForegroundColor Red
}