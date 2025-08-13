<#
    .SYNOPSIS
        Collects comprehensive system information and exports to JSON
    .DESCRIPTION
        Gathers hardware, software, and network configuration details
        Compatible with PowerShell 2.0 through latest versions
        Outputs structured JSON file to user's desktop

    .COMPANY
        North West Provincial Treasury 
    .AUTHOR
        Lesedi Sebekedi
#>

function Get-SystemInfo {
    [CmdletBinding()]
    param()

    try {
        # Determine if we can use modern CIM cmdlets (PowerShell 3.0+) or fallback to WMI
        $useCim = $PSVersionTable.PSVersion.Major -ge 3
        
        # REGION: BASIC SYSTEM INFORMATION
        # -------------------------------
        # Collect core system details (OS, BIOS, computer model)
        if ($useCim) {
            $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
            $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
            $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop
        } 
        else {
            $computerSystem = Get-WmiObject -Class Win32_ComputerSystem -ErrorAction Stop
            $os = Get-WmiObject -Class Win32_OperatingSystem -ErrorAction Stop
            $bios = Get-WmiObject -Class Win32_BIOS -ErrorAction Stop
        }

        # REGION: PROCESSOR INFORMATION
        # ----------------------------
        # Get CPU details including cores, threads and clock speed
        $cpuParams = @{
            Property = @(
                'Name'
                @{Name="Cores"; Expression={$_.NumberOfCores}}
                @{Name="Threads"; Expression={$_.NumberOfLogicalProcessors}}
                @{Name="ClockSpeed"; Expression={"$($_.MaxClockSpeed) MHz"}}
            )
        }

        if ($useCim) {
            $cpu = Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop | Select-Object @cpuParams
        } 
        else {
            $cpu = Get-WmiObject -Class Win32_Processor -ErrorAction Stop | Select-Object @cpuParams
        }

        # REGION: MEMORY INFORMATION
        # --------------------------
        # Calculate total physical memory and page file size
        if ($useCim) {
            $physicalMemory = Get-CimInstance -ClassName Win32_PhysicalMemory -ErrorAction SilentlyContinue | 
                            Measure-Object -Property Capacity -Sum
            $pageFile = Get-CimInstance -ClassName Win32_PageFileUsage -ErrorAction SilentlyContinue | 
                       Select-Object @{Name="PageFileGB"; Expression={[math]::Round($_.AllocatedBaseSize / 1KB, 2)}}
        } 
        else {
            $physicalMemory = Get-WmiObject -Class Win32_PhysicalMemory -ErrorAction SilentlyContinue | 
                            Measure-Object -Property Capacity -Sum
            $pageFile = Get-WmiObject -Class Win32_PageFileUsage -ErrorAction SilentlyContinue | 
                       Select-Object @{Name="PageFileGB"; Expression={[math]::Round($_.AllocatedBaseSize / 1KB, 2)}}
        }
        
        $totalMemoryGB = [math]::Round(($physicalMemory.Sum / 1GB), 2)

        # REGION: STORAGE INFORMATION
        # ---------------------------
        # Gather disk information including size and free space
        $diskParams = @{
            Property = @(
                'DeviceID'
                'VolumeName'
                @{Name="SizeGB"; Expression={[math]::Round($_.Size / 1GB, 2)}}
                @{Name="FreeGB"; Expression={[math]::Round($_.FreeSpace / 1GB, 2)}}
                @{Name="Type"; Expression={
                    switch($_.DriveType) {
                        2 {"Removable"}
                        3 {"Fixed"} 
                        4 {"Network"}
                        5 {"Optical"}
                    }
                }}
            )
        }

        if ($useCim) {
            $disks = Get-CimInstance -ClassName Win32_LogicalDisk -ErrorAction SilentlyContinue | 
                    Where-Object { $_.Size -gt 0 } | 
                    Select-Object @diskParams
        } 
        else {
            $disks = Get-WmiObject -Class Win32_LogicalDisk -ErrorAction SilentlyContinue | 
                    Where-Object { $_.Size -gt 0 } | 
                    Select-Object @diskParams
        }

        # REGION: NETWORK CONFIGURATION
        # -----------------------------
        # Collect network adapter details with fallback for older PowerShell
        $networkAdapters = @()
        
        # Modern approach using NetAdapter cmdlets if available
        if (Get-Command Get-NetAdapter -ErrorAction SilentlyContinue) {
            $networkAdapters = Get-NetAdapter -Physical -ErrorAction SilentlyContinue | 
                             Where-Object { $_.Status -eq 'Up' } | 
                             Select-Object Name, InterfaceDescription, MacAddress, 
                                 @{Name="Speed"; Expression={"$($_.LinkSpeed)"}},
                                 @{Name="IPAddress"; Expression={
                                     (Get-NetIPAddress -InterfaceIndex $_.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).IPAddress
                                 }}
        }
        # Fallback to WMI/CIM for older systems
        else {
            if ($useCim) {
                $nics = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -ErrorAction SilentlyContinue | 
                        Where-Object { $_.IPEnabled -eq $true }
            } 
            else {
                $nics = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ErrorAction SilentlyContinue | 
                        Where-Object { $_.IPEnabled -eq $true }
            }
            
            foreach ($nic in $nics) {
                $networkAdapters += [PSCustomObject]@{
                    Name = $nic.Description
                    InterfaceDescription = $nic.Description
                    MacAddress = $nic.MACAddress
                    Speed = if ($nic.Speed) { "$($nic.Speed / 1MB) Mbps" } else { "N/A" }
                    IPAddress = $nic.IPAddress[0]
                }
            }
        }

        # REGION: INSTALLED SOFTWARE
        # --------------------------
        # Query registry for installed applications
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

        # REGION: HARDWARE IDENTIFIERS
        # ---------------------------
        # Get system UUID and GPU information
        if ($useCim) {
            $uuid = (Get-CimInstance -ClassName Win32_ComputerSystemProduct -ErrorAction SilentlyContinue).UUID
            $gpu = Get-CimInstance -ClassName Win32_VideoController -ErrorAction SilentlyContinue | 
                   Select-Object Name, 
                       @{Name="AdapterRAMGB"; Expression={[math]::Round($_.AdapterRAM / 1GB, 2)}}, 
                       DriverVersion
        } 
        else {
            $uuid = (Get-WmiObject -Class Win32_ComputerSystemProduct -ErrorAction SilentlyContinue).UUID
            $gpu = Get-WmiObject -Class Win32_VideoController -ErrorAction SilentlyContinue | 
                   Select-Object Name, 
                       @{Name="AdapterRAMGB"; Expression={[math]::Round($_.AdapterRAM / 1GB, 2)}}, 
                       DriverVersion
        }

        # Build and return structured system information object
        return [PSCustomObject]@{
            System = @{
                HostName     = $env:COMPUTERNAME
                OS           = $os.Caption
                Version      = $os.Version
                Build        = $os.BuildNumber
                Architecture = if ([Environment]::Is64BitOperatingSystem) { "64-bit" } else { "32-bit" }
                BootTime     = $os.LastBootUpTime.ToString('yyyy-MM-dd HH:mm:ss')
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
                    PageFileGB = if ($pageFile) { $pageFile.PageFileGB } else { 0 }
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
}

# MAIN EXECUTION
# --------------

# Prompt for and validate asset number
do {
    $assetNumber = Read-Host "Please enter the Asset Number for this system (required)"
    if ([string]::IsNullOrWhiteSpace($assetNumber)) {
        Write-Host "Asset Number cannot be empty!" -ForegroundColor Red
    }
} while ([string]::IsNullOrWhiteSpace($assetNumber))

# Collect system information
$systemInfo = Get-SystemInfo

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
    $systemInfo | ConvertTo-Json -Depth 5 | Out-File -FilePath $reportPath -Force

    # Provide user feedback
    Write-Host "System information successfully saved to:`n$reportPath" -ForegroundColor Green
}
else {
    Write-Host "Failed to collect system information" -ForegroundColor Red
}