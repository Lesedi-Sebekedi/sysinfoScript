<#
.SYNOPSIS
    Windows 7 System Information Collector with Enhanced Network Details
.DESCRIPTION
    Collects system and network information (DHCP/static IP details) 
    and outputs the data in TXT and JSON-friendly structures.
#>

# --- Validate OS Version ---
if ((Get-WmiObject Win32_OperatingSystem).Version -notlike "6.1*") {
    Write-Host "This script is designed for Windows 7 only" -ForegroundColor Red
    exit 1
}

# --- Helpers ---
function Convert-WmiDateTime {
    param([string]$wmiDate)
    if ([string]::IsNullOrEmpty($wmiDate)) { return "N/A" }
    try   { [Management.ManagementDateTimeConverter]::ToDateTime($wmiDate).ToString('yyyy-MM-dd HH:mm:ss') }
    catch { "InvalidDate" }
}

# --- Network Info ---
function Get-NetworkConfiguration {
    try {
        $adapters = @()
        $nics = Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled }

        foreach ($nic in $nics) {
            $adapters += [PSCustomObject]@{
                Name              = $nic.Description
                InterfaceDescription = $nic.Description
                MacAddress        = $nic.MACAddress
                Speed             = if ($nic.Speed) { "$([math]::Round($nic.Speed / 1MB, 2)) Mbps" } else { "N/A" }
                IPAddress         = $nic.IPAddress[0]    ?? "N/A"
                IPConfigType      = if ($nic.DHCPEnabled) { "DHCP" } else { "Static" }
                SubnetMask        = $nic.IPSubnet[0]     ?? "N/A"
                DefaultGateway    = $nic.DefaultIPGateway[0] ?? "N/A"
                DNSServers        = ($nic.DNSServerSearchOrder -join ", ") ?? "N/A"
                DHCPEnabled       = $nic.DHCPEnabled
                DHCPServer        = $nic.DHCPServer      ?? "N/A"
                LeaseObtained     = if ($nic.DHCPEnabled) { Convert-WmiDateTime $nic.DHCPLeaseObtained } else { "N/A" }
                LeaseExpires      = if ($nic.DHCPEnabled) { Convert-WmiDateTime $nic.DHCPLeaseExpires  } else { "N/A" }
            }
        }

        return ($adapters.Count -eq 1) ? $adapters[0] : $adapters
    }
    catch {
        Write-Warning "Error collecting network information: $_"
        return $null
    }
}

# --- System Info ---
function Get-Win7SystemInfo {
    try {
        # Core details
        $os         = Get-WmiObject Win32_OperatingSystem
        $system     = Get-WmiObject Win32_ComputerSystem
        $bios       = Get-WmiObject Win32_BIOS
        $sysProduct = Get-WmiObject Win32_ComputerSystemProduct

        # CPU
        $cpu = Get-WmiObject Win32_Processor | 
               Select-Object -First 1 Name,
                   @{Name="Cores";   Expression={$_.NumberOfCores}},
                   @{Name="Threads"; Expression={$_.NumberOfLogicalProcessors}},
                   @{Name="ClockSpeed"; Expression={"$($_.MaxClockSpeed) MHz"}}

        # Memory
        $physMem    = Get-WmiObject Win32_PhysicalMemory
        $memTotal   = [math]::Round(($physMem | Measure-Object Capacity -Sum).Sum / 1GB, 2)
        $pageFile   = Get-WmiObject Win32_PageFileUsage |
                      Select-Object @{Name="PageFileGB"; Expression={[math]::Round($_.AllocatedBaseSize / 1KB, 2)}}

        # Storage
        $disks = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } |
                 Select-Object DeviceID, VolumeName,
                     @{Name="SizeGB"; Expression={[math]::Round($_.Size / 1GB, 2)}},
                     @{Name="FreeGB"; Expression={[math]::Round($_.FreeSpace / 1GB, 2)}},
                     @{Name="Type"; Expression={"Fixed"}}

        # GPU
        $gpu = Get-WmiObject Win32_VideoController |
               Select-Object Name,
                   @{Name="AdapterRAMGB"; Expression={[math]::Round($_.AdapterRAM / 1GB, 2)}},
                   DriverVersion

        # Network
        $netConfig = Get-NetworkConfiguration

        # Software
        $apps = @()
        $regPaths = @(
            "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )

        foreach ($path in $regPaths) {
            if (Test-Path $path) {
                $apps += Get-ItemProperty $path | 
                         Where-Object { $_.DisplayName -and -not $_.SystemComponent } |
                         Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
            }
        }

        $hotfixes = if (Get-Command Get-HotFix -ErrorAction SilentlyContinue) {
            Get-HotFix | Select-Object HotFixID, Description, InstalledOn
        }

        # Output object
        [PSCustomObject]@{
            System = @{
                HostName     = $env:COMPUTERNAME
                OS           = $os.Caption
                Version      = $os.Version
                Build        = $os.BuildNumber
                Architecture = if ([IntPtr]::Size -eq 8) { "64-bit" } else { "32-bit" }
                BootTime     = Convert-WmiDateTime $os.LastBootUpTime
                Manufacturer = $system.Manufacturer
                Model        = $system.Model
                BIOS         = @{
                    Version = $bios.SMBIOSBIOSVersion
                    Serial  = $bios.SerialNumber
                }
            }
            Hardware = @{
                CPU    = $cpu
                Memory = @{
                    TotalGB    = $memTotal
                    Sticks     = $physMem.Count
                    PageFileGB = $pageFile.PageFileGB ?? 0
                }
                Disks  = $disks
                GPU    = $gpu
            }
            Network   = $netConfig
            Software  = @{
                InstalledApps = $apps
                Hotfixes      = $hotfixes
            }
            UUID      = $sysProduct.UUID
            PSVersion = $PSVersionTable.PSVersion.ToString()
            Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        }
    }
    catch {
        Write-Warning "Error collecting system information: $_"
        return $null
    }
}

# --- Text Report Formatter ---
function ConvertTo-TextReport {
    param([Parameter(Mandatory=$true)] $SystemInfo)

    @"
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
    @"
Adapter: $($_.Name)
IPAddress: $($_.IPAddress)
SubnetMask: $($_.SubnetMask)
DefaultGateway: $($_.DefaultGateway)
MacAddress: $($_.MacAddress)
DNSServer: $($_.DNSServers)
IP Configuration: $($_.IPConfigType)
DHCP Server: $($_.DHCPServer)
"@
})

--- SOFTWARE ---
Installed Applications:
$($SystemInfo.Software.InstalledApps | ForEach-Object {
    "  $($_.DisplayName) ($($_.DisplayVersion)) by $($_.Publisher)"
})

Hotfixes:
$($SystemInfo.Software.Hotfixes | ForEach-Object {
    "  $($_.HotFixID) - $($_.Description) (Installed: $($_.InstalledOn))"
})

--- SYSTEM IDENTIFIERS ---
UUID: $($SystemInfo.UUID)
BIOS Version: $($SystemInfo.System.BIOS.Version)
BIOS Serial: $($SystemInfo.System.BIOS.Serial)

Report generated with PowerShell $($SystemInfo.PSVersion)
"@
}

# --- Main Execution ---
Write-Host "Windows 7 System Information Collector" -ForegroundColor Cyan
Write-Host "------------------------------------" -ForegroundColor Cyan

# Asset Number
do {
    $assetNumber = Read-Host "Enter Asset Number (11 digits starting with 0)"
} while ($assetNumber -notmatch '^0\d{10}$')

# Collect Info
$sysInfo = Get-Win7SystemInfo
if (-not $sysInfo) {
    Write-Host "Failed to collect system information" -ForegroundColor Red
    exit 1
}

# Attach Asset Number
$sysInfo | Add-Member -NotePropertyName "AssetNumber" -NotePropertyValue $assetNumber -Force

# Prepare Output Dir
$outputDir = "$env:USERPROFILE\Desktop\SystemReports"
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }

# Save Report
$timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
$pcName    = $env:COMPUTERNAME -replace '[\\/:*?"<>|]', '_'

$txtFile   = Join-Path $outputDir "${pcName}_SystemInfo_${timestamp}.txt"
ConvertTo-TextReport $sysInfo | Out-File $txtFile -Encoding UTF8

Write-Host "`nText report saved to: $txtFile" -ForegroundColor Green
Write-Host "`n=== SYSTEM INFORMATION SUMMARY ===`n"
Write-Host (ConvertTo-TextReport $sysInfo)
