<#
.SYNOPSIS
    Windows 7 System Information Collector with Enhanced Network Details
.DESCRIPTION
    Collects essential system information including DHCP/static IP configuration
    Displays clean JSON output without internal object properties
    Features:
    - Comprehensive hardware/software inventory
    - Detailed network configuration
    - Custom JSON formatting
    - Asset number validation
.NOTES
    Author: Lesedi Sebekedi
    Version: 1.0
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
        return "N/A" 
    }
    
    try {
        return [Management.ManagementDateTimeConverter]::ToDateTime($wmiDate).ToString('yyyy-MM-dd HH:mm:ss')
    } 
    catch {
        return "InvalidDate"
    }
}

<#
.SYNOPSIS
    Formats output as pretty JSON with custom indentation
.PARAMETER InputObject
    The object to format as JSON
.PARAMETER Indent
    Current indentation level (used for recursion)
#>
function Format-JsonOutput {
    param(
        $InputObject,
        [int]$Indent = 0
    )
    
    $indentSpace = " " * $Indent
    $nextIndent = $Indent + 2

    switch ($InputObject) {
        { $_ -is [System.Collections.IDictionary] } {
            $items = @()
            foreach ($key in $InputObject.Keys) {
                $val = Format-JsonOutput -InputObject $InputObject[$key] -Indent $nextIndent
                $items += "$([string]::Format('"{0}" : {1}', $key, $val))"
            }
            return "{`n$($indentSpace + '  ' + ($items -join ",`n$indentSpace  "))`n$indentSpace}"
        }
        
        { $_ -is [System.Collections.IEnumerable] -and -not ($_ -is [string]) } {
            $items = $_ | ForEach-Object { Format-JsonOutput -InputObject $_ -Indent $nextIndent }
            return "[`n$($indentSpace + '  ' + ($items -join ",`n$indentSpace  "))`n$indentSpace]"
        }
        
        { $_ -is [string] } {
            $escaped = $_.Replace('"', '\"')
            return '"' + $escaped + '"'
        }
        
        { $_ -is [bool] } {
            return $_.ToString().ToLower()
        }
        
        { $null -eq $_ } {
            return "null"
        }
        
        default {
            return $_.ToString()
        }
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
            # Determine IP configuration type
            $ipConfigType = if ($nic.DHCPEnabled) { "DHCP" } else { "Static" }

            # Collect DHCP lease information if applicable
            $leaseInfo = @{}
            if ($nic.DHCPEnabled) {
                $leaseInfo = @{
                    DHCPServer    = $nic.DHCPServer
                    LeaseObtained = Convert-WmiDateTime $nic.DHCPLeaseObtained
                    LeaseExpires  = Convert-WmiDateTime $nic.DHCPLeaseExpires
                }
            }

            $adapters += [PSCustomObject]@{
                AdapterName     = $nic.Description
                IPConfiguration = $ipConfigType
                IPAddress       = if ($nic.IPAddress) { $nic.IPAddress } else { @() }
                SubnetMask      = if ($nic.IPSubnet) { $nic.IPSubnet } else { @() }
                DefaultGateway  = if ($nic.DefaultIPGateway) { $nic.DefaultIPGateway } else { @() }
                DNSServers      = if ($nic.DNSServerSearchOrder) { $nic.DNSServerSearchOrder } else { @() }
                MACAddress      = $nic.MACAddress
                DHCPInfo       = $leaseInfo
            }
        }

        return $adapters
    } 
    catch {
        Write-Warning "Error collecting network information: $_"
        return $null
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
        #endregion

        #region Hardware Information
        $cpu = Get-WmiObject Win32_Processor | Select-Object -First 1
        $memory = Get-WmiObject Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum
        $disks = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
        #endregion

        #region Network Information
        $network = Get-NetworkConfiguration
        #endregion

        #region Installed Software
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
                        Name      = $app.DisplayName
                        Version   = if ($app.DisplayVersion) { $app.DisplayVersion } else { "N/A" }
                        Publisher = if ($app.Publisher) { $app.Publisher } else { "N/A" }
                    }
                }
            }
        }
        #endregion

        #region Build Output Object
        $result = [PSCustomObject]@{
            System = [PSCustomObject]@{
                HostName      = $env:COMPUTERNAME
                OS           = $os.Caption
                Version       = $os.Version
                Architecture  = if ([IntPtr]::Size -eq 8) { "64-bit" } else { "32-bit" }
                Manufacturer  = $computer.Manufacturer
                Model         = $computer.Model
                BIOS          = [PSCustomObject]@{
                    Version = $bios.SMBIOSBIOSVersion
                    Serial  = $bios.SerialNumber
                }
            }
            Hardware = [PSCustomObject]@{
                CPU = [PSCustomObject]@{
                    Name       = $cpu.Name
                    Cores      = $cpu.NumberOfCores
                    Threads    = $cpu.NumberOfLogicalProcessors
                    ClockSpeed = "$($cpu.MaxClockSpeed) MHz"
                }
                Memory = [PSCustomObject]@{
                    TotalGB = [math]::Round($memory.Sum / 1GB, 2)
                }
                Disks = $disks | ForEach-Object {
                    [PSCustomObject]@{
                        Drive  = $_.DeviceID
                        SizeGB = if ($_.Size) { [math]::Round($_.Size / 1GB, 2) } else { 0 }
                        FreeGB = if ($_.FreeSpace) { [math]::Round($_.FreeSpace / 1GB, 2) } else { 0 }
                    }
                }
            }
            NetworkAdapters = $network
            Software        = $software
            Timestamp       = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
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
# Display header
Write-Host "Windows 7 System Information Collector" -ForegroundColor Cyan
Write-Host "------------------------------------" -ForegroundColor Cyan

# Asset number input with validation (11 digits starting with 0)
do {
    $assetNumber = Read-Host "Enter Asset Number (11 digits starting with 0)"
} while (-not ($assetNumber -match '^0\d{10}$'))

# Collect system information
$systemInfo = Get-Win7SystemInfo

if ($systemInfo) {
    # Add asset number to collected data
    $systemInfo | Add-Member -NotePropertyName AssetNumber -NotePropertyValue $assetNumber

    # Format and display JSON output
    $jsonOutput = Format-JsonOutput $systemInfo
    Write-Host "`nSystem Information (JSON Format):`n" -ForegroundColor Green
    Write-Host $jsonOutput

    # Save to file option
    $saveToFile = Read-Host "`nSave to file? (Y/N)"
    if ($saveToFile -match '^[Yy]$') {
        $outputDir = "$env:USERPROFILE\Desktop\SystemReports"
        
        try {
            # Create output directory if it doesn't exist
            if (-not (Test-Path $outputDir)) {
                New-Item -ItemType Directory -Path $outputDir -ErrorAction Stop | Out-Null
            }
            
            # Generate filename with computer name and timestamp
            $timestampSafe = (Get-Date).ToString('yyyyMMdd_HHmmss')
            $outputFile = Join-Path $outputDir "$($env:COMPUTERNAME)_SystemInfo_$timestampSafe.json"
            
            # Save JSON to file
            $jsonOutput | Out-File -FilePath $outputFile -Encoding UTF8 -Force
            Write-Host "`nSaved to: $outputFile" -ForegroundColor Green
        } 
        catch {
            Write-Warning "Failed to save file: $_"
        }
    }
} 
else {
    Write-Host "Failed to collect system information" -ForegroundColor Red
}
#endregion