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
        $adapters = @()
        $nics = Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled }

        foreach ($nic in $nics) {
            $ipConfigType = if ($nic.DHCPEnabled) { "DHCP" } else { "Static" }

            $adapter = New-Object PSObject -Property @{
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
            $adapters += $adapter
        }

        if ($adapters.Count -eq 1) { return $adapters[0] }
        return $adapters
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
        $computer = Get-WmiObject Win32_ComputerSystem
        $bios = Get-WmiObject Win32_BIOS
        $systemProduct = Get-WmiObject Win32_ComputerSystemProduct

        # Hardware Information
        $cpu = Get-WmiObject Win32_Processor | Select-Object -First 1
        $memory = Get-WmiObject Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum
        $memorySticks = Get-WmiObject Win32_PhysicalMemory
        $disks = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
        $gpu = Get-WmiObject Win32_VideoController | Select-Object Name, 
            @{Name="AdapterRAMGB"; Expression={[math]::Round($_.AdapterRAM / 1GB, 2)}}, 
            DriverVersion

        # Network Information
        $network = Get-NetworkConfiguration

        # Installed Software
        $installedApps = @()
        $regPaths = @(
            "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        
        foreach ($path in $regPaths) {
            if (Test-Path $path) {
                $apps = Get-ItemProperty $path | Where-Object { $_.DisplayName -and -not $_.SystemComponent }
                foreach ($app in $apps) {
                    $installedApps += [PSCustomObject]@{
                        DisplayName = $app.DisplayName
                        DisplayVersion = if ($app.DisplayVersion) { $app.DisplayVersion } else { "N/A" }
                        Publisher = if ($app.Publisher) { $app.Publisher } else { "N/A" }
                        InstallDate = if ($app.InstallDate) { $app.InstallDate } else { $null }
                    }
                }
            }
        }

        # Hotfixes
        $hotfixes = @()
        if (Get-Command Get-HotFix -ErrorAction SilentlyContinue) {
            $hotfixes = Get-HotFix -ErrorAction SilentlyContinue | ForEach-Object {
                [PSCustomObject]@{
                    HotFixID = $_.HotFixID
                    Description = $_.Description
                    InstalledOn = @{
                        value = $_.InstalledOn
                        DateTime = Convert-WmiDateTime $_.InstalledOn
                    }
                }
            }
        }

        # Build output object
        $result = New-Object PSObject -Property @{
            System = New-Object PSObject -Property @{
                HostName = $env:COMPUTERNAME
                OS = $os.Caption
                Version = $os.Version
                Build = $os.BuildNumber
                Architecture = if ([IntPtr]::Size -eq 8) { "64-bit" } else { "32-bit" }
                BootTime = Convert-WmiDateTime $os.LastBootUpTime
                Manufacturer = $computer.Manufacturer
                Model = $computer.Model
                BIOS = New-Object PSObject -Property @{
                    Version = $bios.SMBIOSBIOSVersion
                    Serial = $bios.SerialNumber
                }
            }
            Hardware = New-Object PSObject -Property @{
                CPU = New-Object PSObject -Property @{
                    Name = $cpu.Name
                    Cores = $cpu.NumberOfCores
                    Threads = $cpu.NumberOfLogicalProcessors
                    ClockSpeed = "$($cpu.MaxClockSpeed) MHz"
                }
                Memory = New-Object PSObject -Property @{
                    TotalGB = [math]::Round($memory.Sum / 1GB, 2)
                    Sticks = $memorySticks.Count
                    PageFileGB = [math]::Round((Get-WmiObject Win32_PageFileUsage).AllocatedBaseSize / 1KB, 2)
                }
                Disks = $disks | ForEach-Object {
                    New-Object PSObject -Property @{
                        DeviceID = $_.DeviceID
                        VolumeName = if ($_.VolumeName) { $_.VolumeName } else { "N/A" }
                        SizeGB = if ($_.Size) { [math]::Round($_.Size / 1GB, 2) } else { 0 }
                        FreeGB = if ($_.FreeSpace) { [math]::Round($_.FreeSpace / 1GB, 2) } else { 0 }
                        Type = "Fixed"
                    }
                }
                GPU = $gpu
            }
            Network = $network
            Software = New-Object PSObject -Property @{
                InstalledApps = $installedApps
                Hotfixes = $hotfixes
            }
            UUID = $systemProduct.UUID
            PSVersion = $PSVersionTable.PSVersion.ToString()
            Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        }

        return $result
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
        "Adapter: $($_.Name)"
        "  IP Address: $($_.IPAddress)"
        "  Subnet Mask: $($_.SubnetMask)"
        "  Gateway: $($_.DefaultGateway)"
        "  MAC: $($_.MacAddress)"
        "  DNS: $($_.DNSServers)`n"
    }
})

--- SOFTWARE ---
Installed Applications:
$($SystemInfo.Software.InstalledApps | ForEach-Object {
    "  $($_.DisplayName) ($($_.DisplayVersion)) by $($_.Publisher)"
})

Hotfixes:
$($SystemInfo.Software.Hotfixes | ForEach-Object {
    "  $($_.HotFixID) - $($_.Description) (Installed: $($_.InstalledOn.DateTime))"
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
    $systemInfo | Add-Member -MemberType NoteProperty -Name "AssetNumber" -Value $assetNumber -Force

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
    $computerName = $env:COMPUTERNAME

    # Generate and save TXT report
    $txtReport = ConvertTo-TextReport -SystemInfo $systemInfo
    $txtFile = Join-Path $outputDir "${computerName}_SystemInfo_${timestampSafe}.txt"
    $txtReport | Out-File -FilePath $txtFile -Encoding UTF8
    Write-Host "`nText report saved to: $txtFile" -ForegroundColor Green

    # Generate and save JSON report (optional)
    $saveJson = Read-Host "`nAlso save as JSON? (Y/N)"
    if ($saveJson -match '^[Yy]$') {
        $jsonOutput = $systemInfo | ConvertTo-Json -Depth 5
        $jsonFile = Join-Path $outputDir "${computerName}_SystemInfo_${timestampSafe}.json"
        $jsonOutput | Out-File -FilePath $jsonFile -Encoding UTF8
        Write-Host "JSON report saved to: $jsonFile" -ForegroundColor Green
    }

    # Display report in console
    Write-Host "`n=== SYSTEM INFORMATION SUMMARY ===`n"
    Write-Host $txtReport
}
else {
    Write-Host "Failed to collect system information" -ForegroundColor Red
}