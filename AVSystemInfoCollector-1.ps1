# AVSystemInfoCollector-1.ps1
# Collect and flatten full system info for PowerApps/CSV import

Write-Host "Collecting system information..." -ForegroundColor Cyan

# --- UUID ---
$uuid = (Get-WmiObject Win32_ComputerSystemProduct).UUID

# --- Hardware Info ---
$hardware = @{
    Manufacturer = (Get-WmiObject Win32_ComputerSystem).Manufacturer
    Model        = (Get-WmiObject Win32_ComputerSystem).Model
    TotalRAMGB   = [math]::Round((Get-WmiObject Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
    CPU          = (Get-WmiObject Win32_Processor).Name
}

# --- Software Info ---
$softwareList = @()
$uninstallKeys = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

foreach ($path in $uninstallKeys) {
    Get-ItemProperty $path -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -and $_.DisplayName.Trim() -ne "" } |
    ForEach-Object {
        $softwareList += [PSCustomObject]@{
            Name        = $_.DisplayName
            Version     = $_.DisplayVersion
            Vendor      = $_.Publisher
            InstallDate = $_.InstallDate
        }
    }
}

# --- Network Info ---
$network = @{
    HostName = $env:COMPUTERNAME
    IPs      = (Get-WmiObject Win32_NetworkAdapterConfiguration |
                Where-Object { $_.IPEnabled -eq $true } |
                ForEach-Object { $_.IPAddress } | Where-Object { $_ -notmatch ":" }) -join ", "
    MACs     = (Get-WmiObject Win32_NetworkAdapterConfiguration |
                Where-Object { $_.IPEnabled -eq $true } |
                ForEach-Object { $_.MACAddress }) -join ", "
}

# --- System Info ---
$os = Get-WmiObject Win32_OperatingSystem
$system = @{
    OSName      = $os.Caption
    OSVersion   = $os.Version
    LastBootUp  = $os.ConvertToDateTime($os.LastBootUpTime)
}

# --- PSVersion ---
$psver = $PSVersionTable.PSVersion.ToString()

# --- Flatten into rows ---
$flatRows = @()
foreach ($soft in $softwareList) {
    $flatRows += [PSCustomObject]@{
        AssetNumber = $uuid
        PSVersion   = $psver
        HW_Manufacturer = $hardware.Manufacturer
        HW_Model        = $hardware.Model
        HW_TotalRAMGB   = $hardware.TotalRAMGB
        HW_CPU          = $hardware.CPU
        NW_HostName     = $network.HostName
        NW_IPs          = $network.IPs
        NW_MACs         = $network.MACs
        SYS_OSName      = $system.OSName
        SYS_OSVersion   = $system.OSVersion
        SYS_LastBootUp  = $system.LastBootUp
        SW_Name         = $soft.Name
        SW_Version      = $soft.Version
        SW_Vendor       = $soft.Vendor
        SW_InstallDate  = $soft.InstallDate
    }
}

# --- Output ---
$csvPath = ".\FullSystemInfo.csv"
$flatRows | Export-Csv $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "Full system info collected and saved to $csvPath" -ForegroundColor Green
