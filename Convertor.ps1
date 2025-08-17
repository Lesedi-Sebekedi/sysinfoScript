<#
.SYNOPSIS
    Batch converts system info .txt files to PTJSEBOLA-style JSON (master-map version).
.DESCRIPTION
    Reads all .txt files in a folder, parses them into JSON based on a single master field map,
    preserving extra PTTESTING fields. Output goes into a "Converted" subfolder.
#>

$FolderPath = Join-Path $env:USERPROFILE "Desktop\SystemReports"
if (-not (Test-Path $FolderPath)) { throw "Folder not found: $FolderPath" }

$ConvertedFolder = Join-Path $FolderPath "Converted"
if (-not (Test-Path $ConvertedFolder)) { New-Item $ConvertedFolder -ItemType Directory | Out-Null }

function Extract($text, $pattern, [switch]$Number) {
    if ($text -match $pattern) {
        return $Number ? [decimal]$matches[1].Trim() : $matches[1].Trim()
    }
    $null
}

# Master map: nested sections → fields → regex + type
$map = @{
    System = @{
        BIOS = @{
            Version      = @{ pattern = "BIOS Version: (.+?)\r?\n" }
            Serial       = @{ pattern = "BIOS Serial: (.+?)\r?\n" }
        }
        Version      = @{ pattern = "Version: (.+?)\r?\n" }
        Architecture = @{ pattern = "Architecture: (.+?)\r?\n" }
        HostName     = @{ pattern = "COMPUTER NAME: (.+?)\r?\n" }
        OS           = @{ pattern = "OS: (.+?)\r?\n" }
        Model        = @{ pattern = "Model: (.+?)\r?\n" }
        Manufacturer = @{ pattern = "Manufacturer: (.+?)\r?\n" }
        BootTime     = @{ pattern = "Last Boot: (.+?)\r?\n" }
        Build        = @{ pattern = "Build: (.+?)\r?\n" }
    }
    Hardware = @{
        GPU = @{
            Name          = @{ pattern = "GPU: (.+?)\r?\n" }
            AdapterRAMGB  = @{ pattern = "VRAM: (.+?) GB"; number = $true }
            DriverVersion = @{ pattern = "Driver: (.+?)\r?\n" }
        }
        CPU = @{
            Name       = @{ pattern = "CPU: (.+?)\r?\n" }
            Cores      = @{ pattern = "Cores: (\d+)\r?\n"; number = $true }
            Threads    = @{ pattern = "Threads: (\d+)\r?\n"; number = $true }
            ClockSpeed = @{ pattern = "Clock Speed: (.+?)\r?\n" }
        }
        Memory = @{
            TotalGB    = @{ pattern = "Memory: (.+?) GB"; number = $true }
            PageFileGB = @{ pattern = "Page File: (.+?) GB"; number = $true }
            Sticks     = @{ pattern = "Memory: \d+ GB \((\d+) sticks\)"; number = $true }
        }
    }
    UUID        = @{ pattern = "UUID: (.+?)\r?\n" }
    PSVersion   = @{ pattern = "Report generated with PowerShell (.+?)\r?\n" }
    AssetNumber = @{ pattern = "ASSET NUMBER: (.+?)\r?\n" }
}

function Build-FromMap($mapNode, $text) {
    $obj = @{}
    foreach ($key in $mapNode.Keys) {
        $item = $mapNode[$key]
        if ($item -is [hashtable] -and $item.ContainsKey('pattern')) {
            $obj[$key] = Extract $text $item.pattern ($item.number)
        }
        elseif ($item -is [hashtable]) {
            $obj[$key] = Build-FromMap $item $text
        }
    }
    return $obj
}

Get-ChildItem -Path $FolderPath -Filter *.txt | ForEach-Object {
    Write-Host "Processing $($_.Name)..."
    $text = Get-Content $_.FullName -Raw

    $output = Build-FromMap $map $text

    # Add arrays for sections that aren't in the map
    $output.Hardware.Disks = @()
    $output.Network        = @()
    $output.Software       = @{ Hotfixes = @(); InstalledApps = @() }

    # Parse disks
    foreach ($d in [regex]::Matches($text, "Drive (.+?): (.+?)GB total, (.+?)GB free \((.+?)\)")) {
        $output.Hardware.Disks += @{
            DeviceID   = $d.Groups[1].Value.Trim()
            SizeGB     = [decimal]$d.Groups[2].Value
            FreeGB     = [decimal]$d.Groups[3].Value
            VolumeName = $d.Groups[4].Value.Trim()
            Type       = "Fixed"
        }
    }

    # Parse network
    $netSec = [regex]::Match($text, "--- NETWORK ---(.+?)--- SOFTWARE ---", 'Singleline')
    if ($netSec.Success) {
        foreach ($b in ($netSec.Groups[1].Value -split "\r?\n\r?\n" | Where-Object { $_ -match "Adapter:" })) {
            $dns = if ($b -match "DNSServer: (.+?)(\r?\n|$)") { $matches[1] -split '\s*,\s*' }
            $output.Network += @{
                Name           = Extract $b "Adapter: (.+?)\r?\n"
                IPAddress      = Extract $b "IPAddress: (.+?)\r?\n"
                SubnetMask     = Extract $b "SubnetMask: (.+?)\r?\n"
                DefaultGateway = Extract $b "DefaultGateway: (.+?)\r?\n"
                MacAddress     = Extract $b "MacAddress: (.+?)\r?\n"
                DNSServers     = $dns
                Speed          = "N/A"
            }
        }
    }

    # Installed Apps
    foreach ($a in [regex]::Matches($text, "  (.+?) \((.+?)\) by (.+?)\r?\n")) {
        $output.Software.InstalledApps += @{
            DisplayName    = $a.Groups[1].Value.Trim()
            DisplayVersion = $a.Groups[2].Value.Trim()
            Publisher      = $a.Groups[3].Value.Trim()
            InstallDate    = $null
        }
    }

    # Hotfixes
    foreach ($h in [regex]::Matches($text, "  (.+?) - (.+?) \(Installed: (.+?)\)\r?\n")) {
        try {
            $date = [DateTime]::Parse($h.Groups[3].Value)
            $ms   = [int][double]::Parse((Get-Date $date -UFormat %s)) * 1000
        } catch { $date = $null; $ms = $null }
        $output.Software.Hotfixes += @{
            HotFixID    = $h.Groups[1].Value.Trim()
            Description = $h.Groups[2].Value.Trim()
            InstalledOn = @{
                value    = if ($ms) { "/Date($ms)/" } else { $null }
                DateTime = if ($date) { $date.ToString("yyyy-MM-dd HH:mm:ss") } else { $h.Groups[3].Value.Trim() }
            }
        }
    }

    # Save JSON
    $outFile = Join-Path $ConvertedFolder ($_.BaseName + "_converted.json")
    $output | ConvertTo-Json -Depth 6 | Out-File $outFile -Encoding UTF8
    Write-Host "Saved: $outFile"
}

Write-Host "All files processed. Output in: $ConvertedFolder"
