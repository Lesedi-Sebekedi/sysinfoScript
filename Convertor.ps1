<#
.SYNOPSIS
    Batch converts system info .txt files to PTJSEBOLA-style JSON.
.DESCRIPTION
    Reads all .txt system info files in a specified folder, parses them, 
    and outputs JSON matching PTJSEBOLA's structure while preserving extra PTTESTING fields.
    Output is saved in a "Converted" subfolder.
#>

$FolderPath = Join-Path $env:USERPROFILE "Desktop\SystemReports"

if (-not (Test-Path $FolderPath)) {
    Write-Error "Folder not found: $FolderPath"
    exit
}

# Create Converted folder
$ConvertedFolder = Join-Path $FolderPath "Converted"
if (-not (Test-Path $ConvertedFolder)) {
    New-Item -Path $ConvertedFolder -ItemType Directory | Out-Null
}

# Process each .txt file
Get-ChildItem -Path $FolderPath -Filter *.txt | ForEach-Object {
    Write-Host "Processing $($_.Name)..."

    $textContent = Get-Content -Path $_.FullName -Raw

    # Extract memory sticks count from the text
    $memorySticks = if ($textContent -match "Memory: \d+ GB \((\d+) sticks\)") {
        [int]$matches[1]
    } else {
        0
    }

    # Build PTJSEBOLA-style object
    $output = @{
        System = @{
            BIOS = @{
                Version = ($textContent -match "BIOS Version: (.+?)\r?\n") ? $matches[1].Trim() : $null
                Serial  = ($textContent -match "BIOS Serial: (.+?)\r?\n") ? $matches[1].Trim() : $null
            }
            Version      = ($textContent -match "Version: (.+?)\r?\n") ? $matches[1].Trim() : $null
            Architecture = ($textContent -match "Architecture: (.+?)\r?\n") ? $matches[1].Trim() : $null
            HostName     = ($textContent -match "COMPUTER NAME: (.+?)\r?\n") ? $matches[1].Trim() : $null
            OS           = ($textContent -match "OS: (.+?)\r?\n") ? $matches[1].Trim() : $null
            Model        = ($textContent -match "Model: (.+?)\r?\n") ? $matches[1].Trim() : $null
            Manufacturer = ($textContent -match "Manufacturer: (.+?)\r?\n") ? $matches[1].Trim() : $null
            BootTime     = ($textContent -match "Last Boot: (.+?)\r?\n") ? $matches[1].Trim() : $null
            Build        = ($textContent -match "Build: (.+?)\r?\n") ? $matches[1].Trim() : $null
        }
        Hardware = @{
            GPU = @{
                Name         = ($textContent -match "GPU: (.+?)\r?\n") ? $matches[1].Trim() : $null
                AdapterRAMGB = ($textContent -match "VRAM: (.+?) GB") ? [decimal]$matches[1].Trim() : $null
                DriverVersion= ($textContent -match "Driver: (.+?)\r?\n") ? $matches[1].Trim() : $null
            }
            CPU = @{
                Name       = ($textContent -match "CPU: (.+?)\r?\n") ? $matches[1].Trim() : $null
                Cores      = ($textContent -match "Cores: (\d+)\r?\n") ? [int]$matches[1].Trim() : $null
                Threads    = ($textContent -match "Threads: (\d+)\r?\n") ? [int]$matches[1].Trim() : $null
                ClockSpeed = ($textContent -match "Clock Speed: (.+?)\r?\n") ? $matches[1].Trim() : $null
            }
            Memory = @{
                TotalGB   = ($textContent -match "Memory: (.+?) GB") ? [decimal]$matches[1].Trim() : $null
                PageFileGB= ($textContent -match "Page File: (.+?) GB") ? [decimal]$matches[1].Trim() : $null
                Sticks    = $memorySticks
            }
            Disks = @()
        }
        Network = @()
        Software = @{
            Hotfixes = @()
            InstalledApps = @()
        }
        UUID      = ($textContent -match "UUID: (.+?)\r?\n") ? $matches[1].Trim() : $null
        PSVersion = ($textContent -match "Report generated with PowerShell (.+?)\r?\n") ? $matches[1].Trim() : $null
        AssetNumber = ($textContent -match "ASSET NUMBER: (.+?)\r?\n") ? $matches[1].Trim() : $null
    }

    # Parse disks
    $diskMatches = [regex]::Matches($textContent, "Drive (.+?): (.+?)GB total, (.+?)GB free \((.+?)\)")
    foreach ($disk in $diskMatches) {
        $output.Hardware.Disks += @{
            DeviceID = $disk.Groups[1].Value.Trim()
            SizeGB = [decimal]$disk.Groups[2].Value.Trim()
            FreeGB = [decimal]$disk.Groups[3].Value.Trim()
            VolumeName = $disk.Groups[4].Value.Trim()
            Type = "Fixed"
        }
    }

    # Parse network adapters
    $networkSection = [regex]::Match($textContent, "--- NETWORK ---(.+?)--- SOFTWARE ---", [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if ($networkSection.Success) {
        $networkContent = $networkSection.Groups[1].Value
        $adapterBlocks = $networkContent -split "\r?\n\r?\n" | Where-Object { $_ -match "Adapter:" }

        foreach ($block in $adapterBlocks) {
              # Extract DNS servers - handle multi-line case
            $dnsServers = $null
            if ($block -match "DNSServer: (.+?)(\r?\n|$)") {
                $dnsServers = $matches[1].Trim() -split '\s*,\s*'
            }
            
            $adapter = @{
                Name = ($block -match "Adapter: (.+?)\r?\n") ? $matches[1].Trim() : $null
                IPAddress = ($block -match "IPAddress: (.+?)\r?\n") ? $matches[1].Trim() : $null
                SubnetMask = ($block -match "SubnetMask: (.+?)\r?\n") ? $matches[1].Trim() : $null
                DefaultGateway = ($block -match "DefaultGateway: (.+?)\r?\n") ? $matches[1].Trim() : $null
                MacAddress = ($block -match "MacAddress: (.+?)\r?\n") ? $matches[1].Trim() : $null
                DNSServers = $dnsServers
                Speed = "N/A"
            }
            $output.Network += $adapter
        }
    }

    # Parse installed applications
    $appMatches = [regex]::Matches($textContent, "  (.+?) \((.+?)\) by (.+?)\r?\n")
    foreach ($app in $appMatches) {
        $output.Software.InstalledApps += @{
            DisplayName = $app.Groups[1].Value.Trim()
            DisplayVersion = $app.Groups[2].Value.Trim()
            Publisher = $app.Groups[3].Value.Trim()
            InstallDate = $null
        }
    }

    # Parse hotfixes (PowerShell 2.0 compatible version)
    $hotfixMatches = [regex]::Matches($textContent, "  (.+?) - (.+?) \(Installed: (.+?)\)\r?\n")
    foreach ($hotfix in $hotfixMatches) {
        $dateStr = $hotfix.Groups[3].Value.Trim()
        $dateObj = $null
        $dateMs = $null
        
        # PowerShell 2.0 compatible date parsing
        try {
            $dateObj = [DateTime]::Parse($dateStr)
            $dateMs = [int][double]::Parse((Get-Date $dateObj -UFormat %s)) * 1000
        }
        catch {
            $dateObj = $null
        }
        
        $output.Software.Hotfixes += @{
            HotFixID = $hotfix.Groups[1].Value.Trim()
            Description = $hotfix.Groups[2].Value.Trim()
            InstalledOn = @{
                value = if ($dateMs) { "/Date($dateMs)/" } else { $null }
                DateTime = if ($dateObj) { $dateObj.ToString("yyyy-MM-dd HH:mm:ss") } else { $dateStr }
            }
        }
    }

    # Save output
    $outFile = Join-Path $ConvertedFolder ($_.BaseName + "_converted.json")
    $output | ConvertTo-Json -Depth 6 | Out-File -FilePath $outFile -Encoding UTF8

    Write-Host "Saved: $outFile"
}

Write-Host "All files processed. Output in: $ConvertedFolder"