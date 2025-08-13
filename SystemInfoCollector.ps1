# SystemInfoCollector.ps1
# Collects detailed system information and saves it to a JSON file on the desktop.
function Get-SystemInfo {
    [CmdletBinding()]
    param()

    try {
        # Detect PowerShell version and set appropriate cmdlets
        $useCim = $PSVersionTable.PSVersion.Major -ge 3
        
        # 1. Basic System Info (Compatible with all Windows versions)
        if ($useCim) {
            $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
            $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
            $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop
        } else {
            $computerSystem = Get-WmiObject -Class Win32_ComputerSystem -ErrorAction Stop
            $os = Get-WmiObject -Class Win32_OperatingSystem -ErrorAction Stop
            $bios = Get-WmiObject -Class Win32_BIOS -ErrorAction Stop
        }

        # 2. CPU Info (Works with modern Intel/AMD)
        if ($useCim) {
            $cpu = Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop | Select-Object Name, 
                @{Name="Cores"; Expression={$_.NumberOfCores}},
                @{Name="Threads"; Expression={$_.NumberOfLogicalProcessors}},
                @{Name="ClockSpeed"; Expression={"$($_.MaxClockSpeed) MHz"}}
        } else {
            $cpu = Get-WmiObject -Class Win32_Processor -ErrorAction Stop | Select-Object Name, 
                @{Name="Cores"; Expression={$_.NumberOfCores}},
                @{Name="Threads"; Expression={$_.NumberOfLogicalProcessors}},
                @{Name="ClockSpeed"; Expression={"$($_.MaxClockSpeed) MHz"}}
        }

        # 3. Memory (Physical + Virtual)
        if ($useCim) {
            $physicalMemory = Get-CimInstance -ClassName Win32_PhysicalMemory -ErrorAction SilentlyContinue | 
                              Measure-Object -Property Capacity -Sum
            $pageFile = Get-CimInstance -ClassName Win32_PageFileUsage -ErrorAction SilentlyContinue | 
                        Select-Object @{Name="PageFileGB"; Expression={[math]::Round($_.AllocatedBaseSize / 1KB, 2)}}
        } else {
            $physicalMemory = Get-WmiObject -Class Win32_PhysicalMemory -ErrorAction SilentlyContinue | 
                              Measure-Object -Property Capacity -Sum
            $pageFile = Get-WmiObject -Class Win32_PageFileUsage -ErrorAction SilentlyContinue | 
                        Select-Object @{Name="PageFileGB"; Expression={[math]::Round($_.AllocatedBaseSize / 1KB, 2)}}
        }
        $totalMemoryGB = [math]::Round(($physicalMemory.Sum / 1GB), 2)

        # 4. Disks (Including SSDs/HDDs)
        if ($useCim) {
            $disks = Get-CimInstance -ClassName Win32_LogicalDisk -ErrorAction SilentlyContinue | 
                     Where-Object { $_.Size -gt 0 } | 
                     Select-Object DeviceID, VolumeName,
                        @{Name="SizeGB"; Expression={[math]::Round($_.Size / 1GB, 2)}},
                        @{Name="FreeGB"; Expression={[math]::Round($_.FreeSpace / 1GB, 2)}},
                        @{Name="Type"; Expression={switch($_.DriveType){2{"Removable"}3{"Fixed"}4{"Network"}5{"Optical"}}}}
        } else {
            $disks = Get-WmiObject -Class Win32_LogicalDisk -ErrorAction SilentlyContinue | 
                     Where-Object { $_.Size -gt 0 } | 
                     Select-Object DeviceID, VolumeName,
                        @{Name="SizeGB"; Expression={[math]::Round($_.Size / 1GB, 2)}},
                        @{Name="FreeGB"; Expression={[math]::Round($_.FreeSpace / 1GB, 2)}},
                        @{Name="Type"; Expression={switch($_.DriveType){2{"Removable"}3{"Fixed"}4{"Network"}5{"Optical"}}}}
        }

        # 5. Network (IPv4/IPv6)
        $networkAdapters = @()
        if (Get-Command Get-NetAdapter -ErrorAction SilentlyContinue) {
            $networkAdapters = Get-NetAdapter -Physical -ErrorAction SilentlyContinue | 
                               Where-Object { $_.Status -eq 'Up' } | 
                               Select-Object Name, InterfaceDescription, MacAddress, 
                                   @{Name="Speed"; Expression={"$($_.LinkSpeed)"}},
                                   @{Name="IPAddress"; Expression={
                                       (Get-NetIPAddress -InterfaceIndex $_.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).IPAddress
                                   }}
        } else {
            # Fallback for PowerShell 2.0 without NetAdapter cmdlets
            if ($useCim) {
                $nics = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -ErrorAction SilentlyContinue | 
                        Where-Object { $_.IPEnabled -eq $true }
            } else {
                $nics = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ErrorAction SilentlyContinue | 
                        Where-Object { $_.IPEnabled -eq $true }
            }
            
            foreach ($nic in $nics) {
                $networkAdapters += [PSCustomObject]@{
                    Name = $nic.Description
                    InterfaceDescription = $nic.Description
                    MacAddress = $nic.MACAddress
                    Speed = "N/A"
                    IPAddress = $nic.IPAddress[0]
                }
            }
        }

        # 6. Installed Apps (Registry)
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

        # 7. System UUID and BIOS
        if ($useCim) {
            $uuid = (Get-CimInstance -ClassName Win32_ComputerSystemProduct -ErrorAction SilentlyContinue).UUID
            $gpu = Get-CimInstance -ClassName Win32_VideoController -ErrorAction SilentlyContinue | 
                   Select-Object Name, AdapterRAM, DriverVersion
        } else {
            $uuid = (Get-WmiObject -Class Win32_ComputerSystemProduct -ErrorAction SilentlyContinue).UUID
            $gpu = Get-WmiObject -Class Win32_VideoController -ErrorAction SilentlyContinue | 
                   Select-Object Name, AdapterRAM, DriverVersion
        }

        # Return structured object
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
                    PageFileGB = $pageFile.PageFileGB
                }
                Disks  = $disks
                GPU    = $gpu
            }
            Network = $networkAdapters = if ($networkAdapters.Count -eq 1) { $networkAdapters[0] } else { $networkAdapters }
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

# Prompt for AssetNumber
$assetNumber = Read-Host "Please enter the Asset Number for this system (required)"
while ([string]::IsNullOrWhiteSpace($assetNumber)) {
    Write-Host "Asset Number cannot be empty!" -ForegroundColor Red
    $assetNumber = Read-Host "Please enter the Asset Number for this system (required)"
}

$systemInfo = Get-SystemInfo
if ($systemInfo) {
    $systemInfo | Add-Member -NotePropertyName "AssetNumber" -NotePropertyValue $assetNumber -Force

    $outputDir = "$env:USERPROFILE\Desktop\SystemReports"
    if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

    $computerName = $env:COMPUTERNAME -replace '[\\/:*?"<>|]', '_'
    $reportPath = "$outputDir\$computerName-$(Get-Date -Format 'yyyyMMdd-HHmmss').json"

    $systemInfo | ConvertTo-Json -Depth 10 | Out-File -FilePath $reportPath -Force

    Write-Host "System info saved to $reportPath" -ForegroundColor Green
} else {
    Write-Host "Failed to collect system info" -ForegroundColor Red
}

# End of SystemInfoCollector.ps1
