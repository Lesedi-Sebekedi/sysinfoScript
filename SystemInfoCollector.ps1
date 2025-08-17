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
    .VERSION
        2.0
#>

function Get-SystemInfo {
    [CmdletBinding()]
    param()

    try {
        Write-Host "  🔍 Determining PowerShell capabilities..." -ForegroundColor DarkYellow
        # Determine if we can use modern CIM cmdlets (PowerShell 3.0+) or fallback to WMI
        $useCim = $PSVersionTable.PSVersion.Major -ge 3
        Write-Host "  ✅ Using $($useCim ? 'CIM' : 'WMI') commands" -ForegroundColor DarkGreen
        
        # REGION: BASIC SYSTEM INFORMATION
        # -------------------------------
        Write-Host "  📋 Collecting basic system information..." -ForegroundColor DarkYellow
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
        Write-Host "  ✅ Basic system information collected" -ForegroundColor DarkGreen

        # REGION: PROCESSOR INFORMATION
        # ----------------------------
        Write-Host "  🖥️  Collecting processor information..." -ForegroundColor DarkYellow
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
        Write-Host "  ✅ Processor information collected" -ForegroundColor DarkGreen

        # REGION: MEMORY INFORMATION
        # --------------------------
        Write-Host "  💾 Collecting memory information..." -ForegroundColor DarkYellow
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
        
        # Enhanced memory calculation with null checking
        $totalMemoryGB = if ($physicalMemory.Sum) { 
            [math]::Round(($physicalMemory.Sum / 1GB), 2) 
        } else { 
            0 
        }
        Write-Host "  ✅ Memory information collected ($totalMemoryGB GB)" -ForegroundColor DarkGreen

        # REGION: STORAGE INFORMATION
        # ---------------------------
        Write-Host "  💿 Collecting storage information..." -ForegroundColor DarkYellow
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
                        default {"Unknown"}
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
        Write-Host "  ✅ Storage information collected ($($disks.Count) drives)" -ForegroundColor DarkGreen

        # REGION: NETWORK CONFIGURATION
        # -----------------------------
        Write-Host "  🌐 Collecting network configuration..." -ForegroundColor DarkYellow
        $networkAdapters = @()

        if (Get-Command Get-NetAdapter -ErrorAction SilentlyContinue) {
            $adapters = Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq 'Up' }
            foreach ($adapter in $adapters) {
                $ipConfig = Get-NetIPAddress -InterfaceIndex $adapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue
                $ipInterface = Get-NetIPInterface -InterfaceIndex $adapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue
                $route = Get-NetRoute -InterfaceIndex $adapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | 
                        Where-Object { $_.DestinationPrefix -eq '0.0.0.0/0' } | Select-Object -First 1
                $dns = Get-DnsClientServerAddress -InterfaceIndex $adapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue
                
                $networkAdapters += [PSCustomObject]@{
                    Name                = $adapter.Name
                    InterfaceDescription = $adapter.InterfaceDescription
                    MacAddress          = $adapter.MacAddress
                    Speed               = "$($adapter.LinkSpeed)"
                    IPAddress           = $ipConfig.IPAddress
                    SubnetMask          = $ipConfig.PrefixLength
                    DefaultGateway      = $route.NextHop
                    DNSServers          = $dns.ServerAddresses -join ','
                    DHCPEnabled         = ($ipInterface.Dhcp -eq "Enabled")
                    IPConfigType        = if ($ipInterface.Dhcp -eq "Enabled") { "DHCP" } else { "Static" }
                }
            }
        }
        Write-Host "  ✅ Network configuration collected ($($networkAdapters.Count) adapters)" -ForegroundColor DarkGreen

        # REGION: INSTALLED SOFTWARE
        # --------------------------
        Write-Host "  📦 Collecting software inventory..." -ForegroundColor DarkYellow
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
        Write-Host "  ✅ Software inventory collected ($($installedApps.Count) applications)" -ForegroundColor DarkGreen

        # REGION: HARDWARE IDENTIFIERS
        # ---------------------------
        Write-Host "  🔧 Collecting hardware identifiers..." -ForegroundColor DarkYellow
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

        # Enhanced UUID handling with fallback
        if (-not $uuid) { 
            $uuid = "Unknown" 
            Write-Warning "  ⚠️  UUID not collected, using 'Unknown'"
        }
        Write-Host "  ✅ Hardware identifiers collected" -ForegroundColor DarkGreen

        # Build and return structured system information object
        Write-Host "  🏗️  Building system information object..." -ForegroundColor DarkYellow
        $systemInfo = [PSCustomObject]@{
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
            Network = $networkAdapters  # Always return array for consistency
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

        Write-Host "  ✅ System information object built successfully" -ForegroundColor DarkGreen
        return $systemInfo
    }
    catch {
        Write-Error "❌ Failed to collect system info: $_"
        return $null
    }
}

function Test-DataIntegrity {
    param([Parameter(Mandatory)][PSObject]$SystemInfo)
    
    $warnings = @()
    
    # Check critical fields
    if (-not $SystemInfo.System.HostName) {
        $warnings += "Hostname not collected - this may cause import issues"
    }
    if (-not $SystemInfo.System.BIOS.Serial) {
        $warnings += "BIOS Serial not collected - this may cause import issues"
    }
    if (-not $SystemInfo.UUID -or $SystemInfo.UUID -eq "Unknown") {
        $warnings += "System UUID not collected - this may cause import issues"
    }
    if ($SystemInfo.Network.Count -eq 0) {
        $warnings += "No network adapters found - network import will be skipped"
    }
    if ($SystemInfo.Hardware.Disks.Count -eq 0) {
        $warnings += "No disk information found - disk import will be skipped"
    }
    
    return $warnings
}

# MAIN EXECUTION
# --------------

Write-Host "🚀 System Information Collector v2.0" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "North West Provincial Treasury" -ForegroundColor DarkGray
Write-Host "Author: Lesedi Sebekedi" -ForegroundColor DarkGray
Write-Host ""

# Enhanced asset number input with validation
do {
    $assetNumber = Read-Host "Please enter the Asset Number for this system (required)"
    if ([string]::IsNullOrWhiteSpace($assetNumber)) {
        Write-Host "❌ Asset Number cannot be empty!" -ForegroundColor Red
    } elseif ($assetNumber -notmatch '^[A-Z0-9]+$') {
        Write-Host "❌ Asset Number should contain only letters and numbers!" -ForegroundColor Red
        Write-Host "   Example format: ABC123456" -ForegroundColor DarkYellow
    } else {
        Write-Host "✅ Asset Number format is valid: $assetNumber" -ForegroundColor Green
        break
    }
} while ($true)

Write-Host ""
Write-Host "🔍 Collecting system information..." -ForegroundColor Yellow

# Collect system information
$systemInfo = Get-SystemInfo

if ($systemInfo) {
    Write-Host ""
    Write-Host "📊 Data integrity check..." -ForegroundColor Yellow
    
    # Add asset number to collected data
    $systemInfo | Add-Member -NotePropertyName "AssetNumber" -NotePropertyValue $assetNumber -Force

    # Validate data integrity
    $warnings = Test-DataIntegrity -SystemInfo $systemInfo
    if ($warnings.Count -gt 0) {
        Write-Host "⚠️  Data collection warnings:" -ForegroundColor Yellow
        foreach ($warning in $warnings) {
            Write-Host "   - $warning" -ForegroundColor DarkYellow
        }
    } else {
        Write-Host "✅ All critical data collected successfully" -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "💾 Preparing to export data..." -ForegroundColor Yellow
    
    # Create output directory if it doesn't exist
    $outputDir = "$env:USERPROFILE\Desktop\SystemReports"
    if (-not (Test-Path $outputDir)) { 
        Write-Host "  📁 Creating output directory..." -ForegroundColor DarkYellow
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null 
        Write-Host "  ✅ Output directory created: $outputDir" -ForegroundColor DarkGreen
    } else {
        Write-Host "  ✅ Output directory exists: $outputDir" -ForegroundColor DarkGreen
    }

    # Generate output filename with sanitized computer name and timestamp
    $computerName = $env:COMPUTERNAME -replace '[\\/:*?"<>|]', '_'
    $reportPath = "$outputDir\$computerName-$(Get-Date -Format 'yyyyMMdd-HHmmss').json"

    Write-Host "  📝 Exporting to JSON file..." -ForegroundColor DarkYellow
    
    # Export data to JSON file
    $systemInfo | ConvertTo-Json -Depth 5 | Out-File -FilePath $reportPath -Force

    # Enhanced file validation
    if (Test-Path $reportPath) {
        $fileSize = (Get-Item $reportPath).Length
        if ($fileSize -gt 0) {
            Write-Host ""
            Write-Host "🎉 SUCCESS!" -ForegroundColor Green
            Write-Host "===========" -ForegroundColor Green
            Write-Host "✅ System information successfully saved to:" -ForegroundColor Green
            Write-Host "   📁 File: $reportPath" -ForegroundColor Cyan
            Write-Host "   📊 Size: $([math]::Round($fileSize/1KB, 2)) KB" -ForegroundColor Cyan
            Write-Host "   🕒 Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
            Write-Host ""
            Write-Host "📋 Summary:" -ForegroundColor Yellow
            Write-Host "   - System: $($systemInfo.System.HostName)" -ForegroundColor DarkGray
            Write-Host "   - Asset: $assetNumber" -ForegroundColor DarkGray
            Write-Host "   - OS: $($systemInfo.System.OS)" -ForegroundColor DarkGray
            Write-Host "   - Memory: $($systemInfo.Hardware.Memory.TotalGB) GB" -ForegroundColor DarkGray
            Write-Host "   - Network Adapters: $($systemInfo.Network.Count)" -ForegroundColor DarkGray
            Write-Host "   - Applications: $($systemInfo.Software.InstalledApps.Count)" -ForegroundColor DarkGray
        } else {
            Write-Error "❌ File created but appears to be empty"
        }
    } else {
        Write-Error "❌ Failed to create output file"
    }
}
else {
    Write-Host ""
    Write-Host "❌ FAILED TO COLLECT SYSTEM INFORMATION" -ForegroundColor Red
    Write-Host "=========================================" -ForegroundColor Red
    Write-Host "The script encountered an error during data collection." -ForegroundColor Red
    Write-Host "Please check the error messages above and try again." -ForegroundColor Red
    Write-Host ""
    Write-Host "Common troubleshooting steps:" -ForegroundColor Yellow
    Write-Host "1. Run PowerShell as Administrator" -ForegroundColor DarkYellow
    Write-Host "2. Check Windows Management Instrumentation (WMI) service" -ForegroundColor DarkYellow
    Write-Host "3. Verify system has proper permissions" -ForegroundColor DarkYellow
}

Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor DarkGray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")