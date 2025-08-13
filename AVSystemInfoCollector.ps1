# SystemInfoCollector.ps1
# Collects detailed system information including DHCP/static network config
# Fully compatible with PowerShell 1.0 through latest versions
# Fixed progress tracking and DHCP lease time conversion issues

# Check hostname before proceeding

<#
.SYNOPSIS
    Imports system inventory data to SQL Server
.COMPANY
    Your Company Name
.AUTHOR
    Your Name
.VERSION
    1.0
.SECURITY
    Requires SQL write permissions
#>
$hostname = $env:COMPUTERNAME
if (-not $hostname.StartsWith("PT")) {
    Write-Host "This script can only run on systems with hostnames starting with 'PT'" -ForegroundColor Red
    Write-Host "Current hostname: $hostname" -ForegroundColor Yellow
    Write-Host "Script execution aborted." -ForegroundColor Red
    exit
}
function Write-ProgressHelper {
    param(
        [int]$Step,
        [int]$TotalSteps,
        [string]$Message
    )
    if ($PSVersionTable.PSVersion.Major -ge 2) {
        Write-Progress -Activity "Collecting System Information" -Status $Message -PercentComplete (($Step / $TotalSteps) * 100)
    } else {
        Write-Host "[$Step/$TotalSteps] $Message" -ForegroundColor Cyan
    }
}

function Get-SystemInfo {
    [CmdletBinding()]
    param()

    try {
        $totalSteps = 7
        $currentStep = 1

        # 1. Basic System Info
        Write-ProgressHelper -Step $currentStep -TotalSteps $totalSteps -Message "Collecting basic system information"
        $os = Get-WmiObject -Class Win32_OperatingSystem -ErrorAction Stop
        $computerSystem = Get-WmiObject -Class Win32_ComputerSystem -ErrorAction Stop
        $bios = Get-WmiObject -Class Win32_BIOS -ErrorAction Stop
        $currentStep++

        # 2. CPU Info
        Write-ProgressHelper -Step $currentStep -TotalSteps $totalSteps -Message "Collecting CPU information"
        $cpu = Get-WmiObject -Class Win32_Processor -ErrorAction Stop | Select-Object Name, 
            @{Name="Cores"; Expression={$_.NumberOfCores}},
            @{Name="Threads"; Expression={$_.NumberOfLogicalProcessors}},
            @{Name="ClockSpeed"; Expression={"$($_.MaxClockSpeed) MHz"}}
        $currentStep++

        # 3. Memory Information
        Write-ProgressHelper -Step $currentStep -TotalSteps $totalSteps -Message "Collecting memory information"
        $physicalMemory = Get-WmiObject -Class Win32_PhysicalMemory -ErrorAction SilentlyContinue | 
                          Measure-Object -Property Capacity -Sum
        $pageFile = Get-WmiObject -Class Win32_PageFileUsage -ErrorAction SilentlyContinue | 
                    Select-Object @{Name="PageFileGB"; Expression={[math]::Round($_.AllocatedBaseSize / 1KB, 2)}}
        $totalMemoryGB = [math]::Round(($physicalMemory.Sum / 1GB), 2)
        $currentStep++

        # 4. Disk Information
        Write-ProgressHelper -Step $currentStep -TotalSteps $totalSteps -Message "Collecting disk information"
        $disks = Get-WmiObject -Class Win32_LogicalDisk -ErrorAction SilentlyContinue | 
                 Where-Object { $_.Size -gt 0 } | 
                 Select-Object DeviceID, VolumeName,
                    @{Name="SizeGB"; Expression={[math]::Round($_.Size / 1GB, 2)}},
                    @{Name="FreeGB"; Expression={[math]::Round($_.FreeSpace / 1GB, 2)}},
                    @{Name="Type"; Expression={switch($_.DriveType){2{"Removable"}3{"Fixed"}4{"Network"}5{"Optical"}}}}
        $currentStep++

        # 5. Network Configuration
        Write-ProgressHelper -Step $currentStep -TotalSteps $totalSteps -Message "Collecting network information"
        $networkAdapters = @()
        $nics = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ErrorAction SilentlyContinue | 
                Where-Object { $_.IPEnabled -eq $true }
        
        foreach ($nic in $nics) {
            # Determine DHCP vs Static configuration
            $ipConfigType = if ($nic.DHCPEnabled -ne $null) {
                if ($nic.DHCPEnabled) { "DHCP" } else { "Static" }
            } else { "Unknown" }
            
            # Get network configuration
            $dnsServers = if ($nic.DNSServerSearchOrder) { $nic.DNSServerSearchOrder } else { @() }
            $gateways = if ($nic.DefaultIPGateway) { $nic.DefaultIPGateway } else { @() }
            $subnets = if ($nic.IPSubnet) { $nic.IPSubnet } else { @() }
            
            # Handle DHCP lease times safely
            $leaseObtained = "N/A"
            $leaseExpires = "N/A"
            if ($nic.DHCPLeaseObtained -and $nic.DHCPLeaseObtained -match '^\d+$') {
                try {
                    $leaseObtained = (Get-Date "01/01/1970").AddSeconds($nic.DHCPLeaseObtained).ToString('yyyy-MM-dd HH:mm:ss')
                } catch {
                    $leaseObtained = "N/A (Error: $($_.Exception.Message))"
                }
            }
            
            if ($nic.DHCPLeaseExpires -and $nic.DHCPLeaseExpires -match '^\d+$') {
                try {
                    $leaseExpires = (Get-Date "01/01/1970").AddSeconds($nic.DHCPLeaseExpires).ToString('yyyy-MM-dd HH:mm:ss')
                } catch {
                    $leaseExpires = "N/A (Error: $($_.Exception.Message))"
                }
            }
            
            $networkAdapters += [PSCustomObject]@{
                Name = $nic.Description
                InterfaceDescription = $nic.Description
                MacAddress = $nic.MACAddress
                Speed = if ($nic.Speed) { "$([math]::Round($nic.Speed/1MB,2)) Mbps" } else { "N/A" }
                IPAddress = if ($nic.IPAddress) { $nic.IPAddress[0] } else { "N/A" }
                IPConfigType = $ipConfigType
                SubnetMask = if ($subnets.Count -gt 0) { $subnets[0] } else { "N/A" }
                DefaultGateway = if ($gateways.Count -gt 0) { $gateways[0] } else { "N/A" }
                DNSServers = $dnsServers -join ", "
                DHCPEnabled = $nic.DHCPEnabled
                DHCPServer = if ($nic.DHCPServer) { $nic.DHCPServer } else { "N/A" }
                LeaseObtained = $leaseObtained
                LeaseExpires = $leaseExpires
            }
        }
        $currentStep++

        # 6. Installed Applications
        Write-ProgressHelper -Step $currentStep -TotalSteps $totalSteps -Message "Collecting installed applications"
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
        $currentStep++

        # 7. Hardware Identifiers
        Write-ProgressHelper -Step $currentStep -TotalSteps $totalSteps -Message "Collecting hardware identifiers"
        $uuid = (Get-WmiObject -Class Win32_ComputerSystemProduct -ErrorAction SilentlyContinue).UUID
        $gpu = Get-WmiObject -Class Win32_VideoController -ErrorAction SilentlyContinue | 
               Select-Object Name, 
               @{Name="AdapterRAMGB"; Expression={[math]::Round($_.AdapterRAM / 1GB, 2)}}, 
               DriverVersion
        $currentStep++

        # Return structured object
        return [PSCustomObject]@{
            System = @{
                HostName     = $env:COMPUTERNAME
                OS           = $os.Caption
                Version      = $os.Version
                Build        = $os.BuildNumber
                Architecture = if ([IntPtr]::Size -eq 8) { "64-bit" } else { "32-bit" }
                BootTime     = [Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime).ToString('yyyy-MM-dd HH:mm:ss')
                Manufacturer = $computerSystem.Manufacturer
                Model        = $computerSystem.Model
                BIOS         = @{
                    Serial = $bios.SerialNumber
                    Version = $bios.SMBIOSBIOSVersion
                }
            }
            Hardware = @{
                CPU    = $cpu
                Memory = @{
                    TotalGB    = $totalMemoryGB
                    Sticks     = $physicalMemory.Count
                    PageFileGB = $pageFile.PageFileGB
                }
                Disks  = $disks
                GPU    = $gpu
            }
            Network = if ($networkAdapters.Count -eq 1) { $networkAdapters[0] } else { $networkAdapters }
            Software = @{
                InstalledApps = $installedApps
                Hotfixes     = if (Get-Command Get-HotFix -ErrorAction SilentlyContinue) {
                    (Get-HotFix -ErrorAction SilentlyContinue | 
                     Select-Object HotFixID, Description, InstalledOn)
                } else { @() }
            }
            UUID    = $uuid
            PSVersion = $PSVersionTable.PSVersion.ToString()
        }
    }
    catch {
        Write-Warning "[ERROR] Failed to collect system info: $_"
        return $null
    }
    finally {
        if ($PSVersionTable.PSVersion.Major -ge 2) {
            Write-Progress -Activity "Collecting System Information" -Completed
        }
    }
}

# Main execution
Write-Host "`nSystem Information Collector" -ForegroundColor Cyan
Write-Host "Compatible with PowerShell 1.0 through latest versions" -ForegroundColor Yellow
Write-Host "Collecting detailed system configuration including network settings`n" -ForegroundColor Gray

# Prompt for AssetNumber
do {
    $assetNumber = Read-Host "Please enter the Asset Number for this system (required)"
    
    # Validation checks
    $isValid = $assetNumber -match '^0\d{10}$' -and 
               $assetNumber[0] -eq '0' -and 
               ($assetNumber -as [long] -ge 7000000000 -and $assetNumber -as [long] -le 7999999999)
    
    if (-not $isValid) {
        Write-Host "Invalid Asset Number format!" -ForegroundColor Red
        Write-Host "Required format: 11 digits starting with 0 (e.g., 07001010000)" -ForegroundColor Yellow
        Write-Host "Valid range: 07000000000 to 07999999999" -ForegroundColor Yellow
    }
} while (-not $isValid)

# Collect system information
$systemInfo = Get-SystemInfo
if ($systemInfo) {
    $systemInfo | Add-Member -NotePropertyName "AssetNumber" -NotePropertyValue $assetNumber -Force

    $outputDir = "$env:USERPROFILE\Desktop\SystemReports"
    if (-not (Test-Path $outputDir)) { 
        New-Item -ItemType Directory -Path $outputDir | Out-Null 
    }

    $computerName = $env:COMPUTERNAME -replace '[\\/:*?"<>|]', '_'
    $reportPath = "$outputDir\$computerName-$(Get-Date -Format 'yyyyMMdd-HHmmss').json"

    # Convert to JSON with fallback for PS 2.0 and earlier
    if ($PSVersionTable.PSVersion.Major -ge 3) {
        $systemInfo | ConvertTo-Json -Depth 10 | Out-File -FilePath $reportPath -Force
    } else {
        # Manual JSON conversion for PS 1.0-2.0
        function ConvertTo-BasicJson {
            param(
                [object]$InputObject,
                [string]$Indent = ""
            )
            
            if ($InputObject -is [System.Collections.IDictionary] -or $InputObject -is [PSObject]) {
                $json = "{`r`n"
                $newIndent = $Indent + "  "
                foreach ($prop in $InputObject.PSObject.Properties) {
                    $json += "$newIndent`"$($prop.Name)`": $(ConvertTo-BasicJson -InputObject $prop.Value -Indent $newIndent),`r`n"
                }
                $json = $json.TrimEnd(",`r`n") + "`r`n$Indent}"
                return $json
            } elseif ($InputObject -is [System.Array]) {
                $json = "[`r`n"
                $newIndent = $Indent + "  "
                foreach ($item in $InputObject) {
                    $json += "$newIndent$(ConvertTo-BasicJson -InputObject $item -Indent $newIndent),`r`n"
                }
                $json = $json.TrimEnd(",`r`n") + "`r`n$Indent]"
                return $json
            } else {
                return "`"$($InputObject)`""
            }
        }

        ConvertTo-BasicJson -InputObject $systemInfo | Out-File -FilePath $reportPath -Force
    }

    Write-Host "`nSystem information successfully collected and saved to:`n$reportPath`n" -ForegroundColor Green
} else {
    Write-Host "`nFailed to collect system information`n" -ForegroundColor Red
}

# End of script