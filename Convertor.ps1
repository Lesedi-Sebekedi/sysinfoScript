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
                AdapterRAMGB = ($textContent -match "GPU RAM: (.+?) GB") ? [decimal]$matches[1].Trim() : $null
                DriverVersion= ($textContent -match "GPU Driver: (.+?)\r?\n") ? $matches[1].Trim() : $null
            }
            CPU = @{
                Name       = ($textContent -match "CPU: (.+?)\r?\n") ? $matches[1].Trim() : $null
                Cores      = ($textContent -match "Cores: (.+?)\r?\n") ? [int]$matches[1].Trim() : $null
                Threads    = ($textContent -match "Threads: (.+?)\r?\n") ? [int]$matches[1].Trim() : $null
                ClockSpeed = ($textContent -match "Clock Speed: (.+?)\r?\n") ? $matches[1].Trim() : $null
            }
            Memory = @{
                TotalGB   = ($textContent -match "Memory: (.+?) GB") ? [int]$matches[1].Trim() : $null
                PageFileGB= ($textContent -match "Page File: (.+?) GB") ? [decimal]$matches[1].Trim() : $null
                Sticks    = ($textContent -match "Memory Sticks: (.+?)\r?\n") ? [int]$matches[1].Trim() : $null
            }
            Disks = @{
                DeviceID   = "C:"
                VolumeName = ($textContent -match "Disk C: Label: (.+?)\r?\n") ? $matches[1].Trim() : $null
                SizeGB     = ($textContent -match "Disk C: Size: (.+?) GB") ? [decimal]$matches[1].Trim() : $null
                FreeGB     = ($textContent -match "Disk C: Free: (.+?) GB") ? [decimal]$matches[1].Trim() : $null
                Type       = "Fixed"
                ExtraDisks = @() # Will store any other disks
            }
        }

        Network = @()
        Software = @{
            Hotfixes = @()
            InstalledApps = @()
        }

        UUID      = ($textContent -match "UUID: (.+?)\r?\n") ? $matches[1].Trim() : $null
        PSVersion = ($textContent -match "PSVersion: (.+?)\r?\n") ? $matches[1].Trim() : $null
        AssetNumber = ($textContent -match "Asset Number: (.+?)\r?\n") ? $matches[1].Trim() : $null
    }

    # Extra disks
    $diskMatches = [regex]::Matches($textContent, "Disk ([A-Z]): Size: (.+?) GB, Free: (.+?) GB")
    foreach ($m in $diskMatches) {
        if ($m.Groups[1].Value -ne "C") {
            $output.Hardware.Disks.ExtraDisks += @{
                DeviceID = "$($m.Groups[1].Value):"
                SizeGB   = [decimal]$m.Groups[2].Value
                FreeGB   = [decimal]$m.Groups[3].Value
                Type     = "Fixed"
            }
        }
    }

    # Network parsing (always array)
    $nicMatches = [regex]::Matches($textContent, "NIC: (.+?)\r?\nMAC: (.+?)\r?\nIP: (.+?)\r?\n")
    foreach ($n in $nicMatches) {
        $output.Network += @{
            Name = $n.Groups[1].Value.Trim()
            MacAddress = $n.Groups[2].Value.Trim()
            IPAddress = $n.Groups[3].Value.Trim()
        }
    }

    # Installed apps
    $appMatches = [regex]::Matches($textContent, "App: (.+?) \| Version: (.*?) \| Publisher: (.*?) \| Date: (.*?)\r?\n")
    foreach ($a in $appMatches) {
        $output.Software.InstalledApps += @{
            DisplayName    = $a.Groups[1].Value.Trim()
            DisplayVersion = $a.Groups[2].Value.Trim()
            Publisher      = $a.Groups[3].Value.Trim()
            InstallDate    = $a.Groups[4].Value.Trim()
        }
    }

    # Hotfixes
    $hotfixMatches = [regex]::Matches($textContent, "Hotfix: (.+?) \| Desc: (.+?) \| Date: (.*?)\r?\n")
    foreach ($h in $hotfixMatches) {
        $dateStr = $h.Groups[3].Value.Trim()
        $dateObj = $null
        $dateMs  = $null
        if ([DateTime]::TryParse($dateStr, [ref]$dateObj)) {
            $dateMs = [int][double]::Parse((Get-Date $dateObj -UFormat %s)) * 1000
        }
        $output.Software.Hotfixes += @{
            HotFixID = $h.Groups[1].Value.Trim()
            Description = $h.Groups[2].Value.Trim()
            InstalledOn = @{
                value = if ($dateMs) { "/Date($dateMs)/" } else { $null }
                DateTime = if ($dateObj) { $dateObj.ToString("dddd, dd MMMM yyyy HH:mm:ss") } else { "N/A" }
            }
        }
    }

    # Save output
    $outFile = Join-Path $ConvertedFolder ($_.BaseName + "_converted.json")
    $output | ConvertTo-Json -Depth 6 | Out-File -FilePath $outFile -Encoding UTF8

    Write-Host "Saved: $outFile"
}

Write-Host "All files processed. Output in: $ConvertedFolder"
