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
        Write-Host "   Determining PowerShell capabilities..." -ForegroundColor DarkYellow
        
        # PSVersion check compatible with PS2.0
        $useCim = $false
        if ($PSVersionTable -and $PSVersionTable.PSVersion) {
            $useCim = $PSVersionTable.PSVersion.Major -ge 3
        }
        Write-Host "   Using $(if ($useCim) {'CIM'} else {'WMI'}) commands" -ForegroundColor DarkGreen        

        # REGION: BASIC SYSTEM INFORMATION
        Write-Host "   Collecting basic system information..." -ForegroundColor DarkYellow
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
        Write-Host "  Basic system information collected" -ForegroundColor DarkGreen

        # REGION: PROCESSOR INFORMATION
        Write-Host "    Collecting processor information..." -ForegroundColor DarkYellow
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
        Write-Host "   Processor information collected" -ForegroundColor DarkGreen

        # REGION: MEMORY INFORMATION
        Write-Host "   Collecting memory information..." -ForegroundColor DarkYellow
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
        
        $totalMemoryGB = if ($physicalMemory.Sum) { 
            [math]::Round(($physicalMemory.Sum / 1GB), 2) 
        } else { 
            0 
        }
        Write-Host "   Memory information collected ($totalMemoryGB GB)" -ForegroundColor DarkGreen

        # REGION: STORAGE INFORMATION
        Write-Host "   Collecting storage information..." -ForegroundColor DarkYellow
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
        Write-Host "   Storage information collected ($(($disks | Measure-Object).Count) drives)" -ForegroundColor DarkGreen

        # REGION: NETWORK CONFIGURATION
        Write-Host "   Collecting network configuration..." -ForegroundColor DarkYellow
        $networkAdapters = @()
        
        # PS2.0 compatible network adapter collection
        $adapters = Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }
        foreach ($adapter in $adapters) {
            $networkAdapters += New-Object PSObject -Property @{
                Name           = $adapter.Description
                IPAddress      = if ($adapter.IPAddress) { $adapter.IPAddress[0] } else { $null }
                SubnetMask     = if ($adapter.IPSubnet) { $adapter.IPSubnet[0] } else { $null }
                DefaultGateway = if ($adapter.DefaultIPGateway) { $adapter.DefaultIPGateway[0] } else { $null }
                DNSServers     = if ($adapter.DNSServerSearchOrder) { $adapter.DNSServerSearchOrder -join ',' } else { $null }
                DHCPEnabled    = $adapter.DHCPEnabled
                MacAddress     = $adapter.MACAddress
            }
        }
        Write-Host "   Network configuration collected ($(($networkAdapters | Measure-Object).Count) adapters)" -ForegroundColor DarkGreen

        # REGION: INSTALLED SOFTWARE
        Write-Host " Collecting software inventory..." -ForegroundColor DarkYellow
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
        Write-Host "   Software inventory collected ($(($installedApps | Measure-Object).Count) applications)" -ForegroundColor DarkGreen

        # REGION: HARDWARE IDENTIFIERS
        Write-Host "Collecting hardware identifiers..." -ForegroundColor DarkYellow
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

        if (-not $uuid) { 
            $uuid = "Unknown" 
            Write-Warning "    UUID not collected, using 'Unknown'"
        }
        Write-Host "   Hardware identifiers collected" -ForegroundColor DarkGreen

        # Build system information object (PS2.0 compatible)
        Write-Host "    Building system information object..." -ForegroundColor DarkYellow
        
        # Create nested objects using New-Object
        $biosObj = New-Object PSObject -Property @{
            Serial = $bios.SerialNumber
            Version = $bios.SMBIOSBIOSVersion
        }
        
        $memoryObj = New-Object PSObject -Property @{
            TotalGB = $totalMemoryGB
            Sticks = $physicalMemory.Count
            PageFileGB = if ($pageFile) { $pageFile.PageFileGB } else { 0 }
        }
        
        $hardwareObj = New-Object PSObject -Property @{
            CPU = $cpu
            Memory = $memoryObj
            Disks = $disks
            GPU = $gpu
        }
        
        $softwareObj = New-Object PSObject -Property @{
            InstalledApps = $installedApps
            Hotfixes = if (Get-Command Get-HotFix -ErrorAction SilentlyContinue) {
                (Get-HotFix -ErrorAction SilentlyContinue | Select-Object HotFixID, Description, InstalledOn)
            } else { @() }
        }
        
        $systemObj = New-Object PSObject -Property @{
            HostName = $env:COMPUTERNAME
            OS = $os.Caption
            Version = $os.Version
            Build = $os.BuildNumber
            Architecture = if ([Environment]::Is64BitOperatingSystem) { "64-bit" } else { "32-bit" }
            BootTime = $os.LastBootUpTime.ToString('yyyy-MM-dd HH:mm:ss')
            Manufacturer = $computerSystem.Manufacturer
            Model = $computerSystem.Model
            BIOS = $biosObj
        }
        
        $systemInfo = New-Object PSObject -Property @{
            System = $systemObj
            Hardware = $hardwareObj
            Network = $networkAdapters
            Software = $softwareObj
            UUID = $uuid
            PSVersion = if ($PSVersionTable.PSVersion) { $PSVersionTable.PSVersion.ToString() } else { "2.0" }
        }

        Write-Host "   System information object built successfully" -ForegroundColor DarkGreen
        return $systemInfo
    }
    catch {
        Write-Error " Failed to collect system info: $_"
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
    if (($SystemInfo.Network | Measure-Object).Count -eq 0) {
        $warnings += "No network adapters found - network import will be skipped"
    }
    if (($SystemInfo.Hardware.Disks | Measure-Object).Count -eq 0) {
        $warnings += "No disk information found - disk import will be skipped"
    }
    
    return $warnings
}

# JSON conversion function for PS2.0
function ConvertTo-Json20 {
    param(
        $obj,
        [int]$depth = 0,
        [int]$maxDepth = 5
    )
    
    if ($depth -gt $maxDepth) {
        return '"max depth reached"'
    }
    
    $indent = '    ' * $depth
    $nextIndent = '    ' * ($depth + 1)
    $newLine = "`r`n"
    
    $js = New-Object System.Text.StringBuilder
    
    if ($obj -eq $null) {
        $js.Append("null") | Out-Null
    }
    elseif ($obj -is [string]) {
        $js.AppendFormat('"{0}"', $obj) | Out-Null
    }
    elseif ($obj -is [bool]) {
        $js.AppendFormat('{0}', $obj.ToString().ToLower()) | Out-Null
    }
    elseif ($obj -is [int] -or $obj -is [float] -or $obj -is [double] -or $obj -is [decimal]) {
        $js.AppendFormat('{0}', $obj) | Out-Null
    }
    elseif ($obj -is [array]) {
        $js.Append("[$newLine") | Out-Null
        $first = $true
        foreach ($item in $obj) {
            if (-not $first) {
                $js.Append(",$newLine") | Out-Null
            }
            $first = $false
            $js.Append($nextIndent) | Out-Null
            $js.Append((ConvertTo-Json20 $item ($depth + 1) $maxDepth)) | Out-Null
        }
        $js.Append("$newLine$indent]") | Out-Null
    }
    elseif ($obj -is [System.Collections.IDictionary] -or $obj -is [PSObject]) {
        $js.Append("{$newLine") | Out-Null
        $first = $true
        
        $properties = if ($obj -is [System.Collections.IDictionary]) {
            $obj.Keys
        } else {
            $obj.PSObject.Properties | Where-Object { $_.Name -notlike '_*' } | Select-Object -ExpandProperty Name
        }
        
        foreach ($prop in $properties) {
            if (-not $first) {
                $js.Append(",$newLine") | Out-Null
            }
            $first = $false
            
            $value = if ($obj -is [System.Collections.IDictionary]) {
                $obj[$prop]
            } else {
                $obj.$prop
            }
            
            $js.Append($nextIndent) | Out-Null
            $js.AppendFormat('"{0}": ', $prop) | Out-Null
            $js.Append((ConvertTo-Json20 $value ($depth + 1) $maxDepth)) | Out-Null
        }
        $js.Append("$newLine$indent}") | Out-Null
    }
    else {
        $js.AppendFormat('"{0}"', $obj.ToString()) | Out-Null
    }
    
    return $js.ToString()
}

# MAIN EXECUTION
Write-Host " System Information Collector v2.0 (PS2.0 Compatible)" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "North West Provincial Treasury" -ForegroundColor DarkGray
Write-Host "Author: Lesedi Sebekedi" -ForegroundColor DarkGray
Write-Host ""

# Asset number validation
do {
    $assetNumber = Read-Host "Please enter the Asset Number for this system (required)"
    if ([string]::IsNullOrWhiteSpace($assetNumber)) {
        Write-Host " Asset Number cannot be empty!" -ForegroundColor Red
    } elseif ($assetNumber -notmatch '^[A-Z0-9]+$') {
        Write-Host " Asset Number should contain only letters and numbers!" -ForegroundColor Red
        Write-Host "   Example format: ABC123456" -ForegroundColor DarkYellow
    } else {
        Write-Host " Asset Number format is valid: $assetNumber" -ForegroundColor Green
        break
    }
} while ($true)

Write-Host ""
Write-Host " Collecting system information..." -ForegroundColor Yellow

$systemInfo = Get-SystemInfo

if ($systemInfo) {
    Write-Host ""
    Write-Host " Data integrity check..." -ForegroundColor Yellow
    
    # Add asset number
    $systemInfo | Add-Member -NotePropertyName "AssetNumber" -NotePropertyValue $assetNumber -Force

    # Check system name starts with PT
    if ($systemInfo.System.HostName -notlike "PT*") {
        Write-Host " Error: System name must start with 'PT' (current name: '$($systemInfo.System.HostName)')" -ForegroundColor Red
        Write-Host "Press any key to exit..." -ForegroundColor DarkGray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit
    }

    # Validate data
    $warnings = Test-DataIntegrity -SystemInfo $systemInfo
    if ($warnings.Count -gt 0) {
        Write-Host " Data collection warnings:" -ForegroundColor Yellow
        foreach ($warning in $warnings) {
            Write-Host "   - $warning" -ForegroundColor DarkYellow
        }
    } else {
        Write-Host " All critical data collected successfully" -ForegroundColor Green
    }

    Write-Host ""
    Write-Host " Preparing to export data..." -ForegroundColor Yellow
    
    # Create output directory
    $outputDir = "$env:USERPROFILE\Desktop\SystemReports"
    if (-not (Test-Path $outputDir)) { 
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null 
        Write-Host "   Output directory created: $outputDir" -ForegroundColor DarkGreen
    }

    # Generate filename
    $computerName = $env:COMPUTERNAME -replace '[\\/:*?"<>|]', '_'
    $reportPath = "$outputDir\$computerName-$(Get-Date -Format 'yyyyMMdd-HHmmss').json"

    Write-Host "   Exporting to JSON file..." -ForegroundColor DarkYellow
    
    # Export to JSON (using our PS2.0 compatible function)
    ConvertTo-Json20 $systemInfo | Out-File -FilePath $reportPath -Force

    # Verify export
    if (Test-Path $reportPath -PathType Leaf) {
        $fileSize = (Get-Item $reportPath).Length
        if ($fileSize -gt 0) {
            Write-Host ""
            Write-Host " SUCCESS!" -ForegroundColor Green
            Write-Host "===========" -ForegroundColor Green
            Write-Host " System information successfully saved to:" -ForegroundColor Green
            Write-Host "   File: $reportPath" -ForegroundColor Cyan
            Write-Host "   Size: $([math]::Round($fileSize/1KB, 2)) KB" -ForegroundColor Cyan
            Write-Host "   Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
            Write-Host ""
            Write-Host " Summary:" -ForegroundColor Yellow
            Write-Host "   - System: $($systemInfo.System.HostName)" -ForegroundColor DarkGray
            Write-Host "   - Asset: $assetNumber" -ForegroundColor DarkGray
            Write-Host "   - OS: $($systemInfo.System.OS)" -ForegroundColor DarkGray
            Write-Host "   - Memory: $($systemInfo.Hardware.Memory.TotalGB) GB" -ForegroundColor DarkGray
            Write-Host "   - Network Adapters: $(($systemInfo.Network | Measure-Object).Count)" -ForegroundColor DarkGray
            Write-Host "   - Applications: $(($systemInfo.Software.InstalledApps | Measure-Object).Count)" -ForegroundColor DarkGray
        }
    }
}
else {
    Write-Host ""
    Write-Host " FAILED TO COLLECT SYSTEM INFORMATION" -ForegroundColor Red
    Write-Host "=========================================" -ForegroundColor Red
    Write-Host "The script encountered an error during data collection." -ForegroundColor Red
}

Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor DarkGray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")