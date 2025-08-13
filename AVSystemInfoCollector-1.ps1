<#
.SYNOPSIS
    Windows 7 System Information Collector with Enhanced Network Details
.DESCRIPTION
    Collects essential system information including DHCP/static IP configuration
    Displays clean JSON output without internal object properties
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

            $leaseInfo = @{}
            if ($nic.DHCPEnabled) {
                $leaseInfo = @{
                    DHCPServer    = $nic.DHCPServer
                    LeaseObtained = Convert-WmiDateTime $nic.DHCPLeaseObtained
                    LeaseExpires  = Convert-WmiDateTime $nic.DHCPLeaseExpires
                }
            }

            $adapters += [PSCustomObject]@{
                AdapterName    = $nic.Description
                IPConfiguration= $ipConfigType
                IPAddress      = if ($nic.IPAddress) { $nic.IPAddress } else { @() }
                SubnetMask     = if ($nic.IPSubnet) { $nic.IPSubnet } else { @() }
                DefaultGateway = if ($nic.DefaultIPGateway) { $nic.DefaultIPGateway } else { @() }
                DNSServers     = if ($nic.DNSServerSearchOrder) { $nic.DNSServerSearchOrder } else { @() }
                MACAddress     = $nic.MACAddress
                DHCPInfo       = $leaseInfo
            }
        }

        return $adapters

    } catch {
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

        # Hardware Information
        $cpu = Get-WmiObject Win32_Processor | Select-Object -First 1
        $memory = Get-WmiObject Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum
        $disks = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }

        # Network Information
        $network = Get-NetworkConfiguration

        # Installed Software
        $software = @()
        $regPaths = @(
            "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        
        foreach ($path in $regPaths) {
            if (Test-Path $path) {
                $apps = Get-ItemProperty $path | Where-Object { $_.DisplayName }
                foreach ($app in $apps) {
                    $software += [PSCustomObject]@{
                        Name = $app.DisplayName
                        Version = if ($app.DisplayVersion) { $app.DisplayVersion } else { "N/A" }
                        Publisher = if ($app.Publisher) { $app.Publisher } else { "N/A" }
                    }
                }
            }
        }

        # Build output object
        $result = [PSCustomObject]@{
            System = [PSCustomObject]@{
                HostName = $env:COMPUTERNAME
                OS = $os.Caption
                Version = $os.Version
                Architecture = if ([IntPtr]::Size -eq 8) { "64-bit" } else { "32-bit" }
                Manufacturer = $computer.Manufacturer
                Model = $computer.Model
                BIOS = [PSCustomObject]@{
                    Version = $bios.SMBIOSBIOSVersion
                    Serial = $bios.SerialNumber
                }
            }
            Hardware = [PSCustomObject]@{
                CPU = [PSCustomObject]@{
                    Name = $cpu.Name
                    Cores = $cpu.NumberOfCores
                    Threads = $cpu.NumberOfLogicalProcessors
                    ClockSpeed = "$($cpu.MaxClockSpeed) MHz"
                }
                Memory = [PSCustomObject]@{
                    TotalGB = [math]::Round($memory.Sum / 1GB, 2)
                }
                Disks = $disks | ForEach-Object {
                    [PSCustomObject]@{
                        Drive = $_.DeviceID
                        SizeGB = if ($_.Size) { [math]::Round($_.Size / 1GB, 2) } else { 0 }
                        FreeGB = if ($_.FreeSpace) { [math]::Round($_.FreeSpace / 1GB, 2) } else { 0 }
                    }
                }
            }
            NetworkAdapters = $network
            Software = $software
            Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        }

        return $result

    } catch {
        Write-Warning "Error collecting system information: $_"
        return $null
    }
}

function Format-JsonOutput {
    param(
        $InputObject,
        $Indent = 0
    )
    $indentSpace = " " * $Indent
    $nextIndent = $Indent + 2

    if ($InputObject -is [System.Collections.IDictionary]) {
        $items = @()
        foreach ($key in $InputObject.Keys) {
            $val = Format-JsonOutput -InputObject $InputObject[$key] -Indent $nextIndent
            $items += "$([string]::Format('"{0}" : {1}', $key, $val))"
        }
        return "{`n$($indentSpace + '  ' + ($items -join ",`n$indentSpace  "))`n$indentSpace}"
    }
    elseif ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string])) {
        $items = $InputObject | ForEach-Object { Format-JsonOutput -InputObject $_ -Indent $nextIndent }
        return "[`n$($indentSpace + '  ' + ($items -join ",`n$indentSpace  "))`n$indentSpace]"
    }
    elseif ($InputObject -is [string]) {
        $escaped = $InputObject.Replace('"', '\"')
        return '"' + $escaped + '"'
    }
    elseif ($InputObject -is [bool]) {
        return $InputObject.ToString().ToLower()
    }
    elseif ($null -eq $InputObject) {
        return "null"
    }
    else {
        return $InputObject.ToString()
    }
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

if ($systemInfo) {
    $systemInfo | Add-Member -NotePropertyName AssetNumber -NotePropertyValue $assetNumber

    $jsonOutput = Format-JsonOutput $systemInfo

    Write-Host "`nSystem Information (JSON Format):`n" -ForegroundColor Green
    Write-Host $jsonOutput

    # Save option
    $saveToFile = Read-Host "`nSave to file? (Y/N)"
    if ($saveToFile -match '^[Yy]$') {
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
        $outputFile = Join-Path $outputDir "$($env:COMPUTERNAME)_SystemInfo_$timestampSafe.json"
        try {
            $jsonOutput | Out-File -FilePath $outputFile -Encoding UTF8 -Force
            Write-Host "`nSaved to: $outputFile" -ForegroundColor Green
        } catch {
            Write-Warning "Failed to save file: $_"
        }
    }
} else {
    Write-Host "Failed to collect system information" -ForegroundColor Red
}
