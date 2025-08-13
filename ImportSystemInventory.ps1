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

param(
    [string]$ReportsFolder = "$env:USERPROFILE\Desktop\SystemReports",
    [string]$ConnectionString = "Server=PTLSEBEKEDI;Database=AssetDB;Integrated Security=True"
)

Add-Type -AssemblyName "System.Data"
$jsonFiles = Get-ChildItem -Path $ReportsFolder -Filter *.json -File

foreach ($file in $jsonFiles) {
    Write-Host "Importing $($file.Name)..." -ForegroundColor Cyan

    try {
        $json = Get-Content $file.FullName -Raw | ConvertFrom-Json
        $conn = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
        $conn.Open()
        $tran = $conn.BeginTransaction()

        # --- Check if system exists ---
        $checkCmd = $conn.CreateCommand()
        $checkCmd.Transaction = $tran
        $checkCmd.CommandText = "SELECT SystemID FROM Systems WHERE AssetNumber=@AssetNumber OR UUID=@UUID"
        $checkCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@AssetNumber",[System.Data.SqlDbType]::VarChar,50))).Value = $json.AssetNumber
        $checkCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@UUID",[System.Data.SqlDbType]::VarChar,36))).Value = $json.UUID
        $sysId = $checkCmd.ExecuteScalar()

        if ($sysId) {
            # Update existing
            $updateCmd = $conn.CreateCommand()
            $updateCmd.Transaction = $tran
            $updateCmd.CommandText = "UPDATE Systems SET HostName=@HostName, SerialNumber=@SerialNumber, ScanDate=@ScanDate WHERE SystemID=@SystemID"
            $updateCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@HostName",[System.Data.SqlDbType]::VarChar,100))).Value = $json.System.HostName
            $updateCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@SerialNumber",[System.Data.SqlDbType]::VarChar,100))).Value = $json.System.BIOS.Serial
            $updateCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@ScanDate",[System.Data.SqlDbType]::DateTime))).Value = [DateTime]::Now
            $updateCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@SystemID",[System.Data.SqlDbType]::Int))).Value = $sysId
            $updateCmd.ExecuteNonQuery() | Out-Null
        }
        else {
            # Insert new
            $insertCmd = $conn.CreateCommand()
            $insertCmd.Transaction = $tran
            $insertCmd.CommandText = @"
INSERT INTO Systems (AssetNumber, HostName, UUID, SerialNumber, ScanDate)
VALUES (@AssetNumber, @HostName, @UUID, @SerialNumber, @ScanDate);
SELECT SCOPE_IDENTITY();
"@
            $insertCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@AssetNumber",[System.Data.SqlDbType]::VarChar,50))).Value = $json.AssetNumber
            $insertCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@HostName",[System.Data.SqlDbType]::VarChar,100))).Value = $json.System.HostName
            $insertCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@UUID",[System.Data.SqlDbType]::VarChar,36))).Value = $json.UUID
            $insertCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@SerialNumber",[System.Data.SqlDbType]::VarChar,100))).Value = $json.System.BIOS.Serial
            $insertCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@ScanDate",[System.Data.SqlDbType]::DateTime))).Value = [DateTime]::Now
            $sysId = [int]$insertCmd.ExecuteScalar()
        }

        # --- Clear old details ---
        foreach ($table in "SystemSpecs","SystemDisks","InstalledApps") {
            $del = $conn.CreateCommand()
            $del.Transaction = $tran
            $del.CommandText = "DELETE FROM $table WHERE SystemID=@ID"
            $del.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@ID",[System.Data.SqlDbType]::Int))).Value = $sysId
            $del.ExecuteNonQuery() | Out-Null
        }

        # --- Insert SystemSpecs ---
# --- Insert SystemSpecs ---
$specCmd = $conn.CreateCommand()
$specCmd.Transaction = $tran
$specCmd.CommandText = @"
INSERT INTO SystemSpecs (SystemID,OSName,OSVersion,OSBuild,Architecture,LastBoot,Manufacturer,Model,BIOSVersion,CPUCores,CPUThreads,CPUClockSpeed,TotalRAMGB)
VALUES (@SystemID,@OSName,@OSVersion,@OSBuild,@Architecture,@LastBoot,@Manufacturer,@Model,@BIOSVersion,@CPUCores,@CPUThreads,@CPUClockSpeed,@TotalRAMGB)
"@

$specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@SystemID",[System.Data.SqlDbType]::Int))).Value = $sysId
$specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@OSName",[System.Data.SqlDbType]::VarChar,100))).Value = $json.System.OS
$specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@OSVersion",[System.Data.SqlDbType]::VarChar,50))).Value = $json.System.Version
$specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@OSBuild",[System.Data.SqlDbType]::VarChar,20))).Value = $json.System.Build
$specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@Architecture",[System.Data.SqlDbType]::VarChar,10))).Value = $json.System.Architecture
$specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@LastBoot",[System.Data.SqlDbType]::DateTime))).Value = [DateTime]::ParseExact($json.System.BootTime,'yyyy-MM-dd HH:mm:ss',$null)
$specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@Manufacturer",[System.Data.SqlDbType]::VarChar,100))).Value = $json.System.Manufacturer
$specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@Model",[System.Data.SqlDbType]::VarChar,100))).Value = $json.System.Model
$specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@BIOSVersion",[System.Data.SqlDbType]::VarChar,100))).Value = $json.System.BIOS.Version
$specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@CPUCores",[System.Data.SqlDbType]::Int))).Value = $json.Hardware.CPU.Cores
$specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@CPUThreads",[System.Data.SqlDbType]::Int))).Value = $json.Hardware.CPU.Threads
$specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@CPUClockSpeed",[System.Data.SqlDbType]::VarChar,20))).Value = $json.Hardware.CPU.ClockSpeed

$ramParam = New-Object System.Data.SqlClient.SqlParameter("@TotalRAMGB",[System.Data.SqlDbType]::Decimal)
$ramParam.Precision = 5
$ramParam.Scale = 2
$ramParam.Value = $json.Hardware.Memory.TotalGB
$specCmd.Parameters.Add($ramParam)

$specCmd.ExecuteNonQuery() | Out-Null


        # --- Insert Disks ---
        foreach ($disk in $json.Hardware.Disks) {
            $dCmd = $conn.CreateCommand()
            $dCmd.Transaction = $tran
            $dCmd.CommandText = "INSERT INTO SystemDisks (SystemID,DriveLetter,SizeGB,FreeGB,DiskType) VALUES (@SystemID,@DL,@Size,@Free,@Type)"
            $dCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@SystemID",[System.Data.SqlDbType]::Int))).Value = $sysId
            $dCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@DL",[System.Data.SqlDbType]::VarChar,5))).Value = $disk.DeviceID
            $dCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@Size",[System.Data.SqlDbType]::Decimal))).Value = $disk.SizeGB
            $dCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@Free",[System.Data.SqlDbType]::Decimal))).Value = $disk.FreeGB
            $dCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@Type",[System.Data.SqlDbType]::VarChar,20))).Value = $disk.Type
            $dCmd.ExecuteNonQuery() | Out-Null
        }

        # --- Insert Apps ---
        foreach ($app in $json.Software.InstalledApps) {
            if (-not $app.DisplayName) { continue }
            $aCmd = $conn.CreateCommand()
            $aCmd.Transaction = $tran
            $aCmd.CommandText = "INSERT INTO InstalledApps (SystemID,AppName,AppVersion,Publisher,InstallDate) VALUES (@SystemID,@Name,@Ver,@Pub,@Date)"
            $aCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@SystemID",[System.Data.SqlDbType]::Int))).Value = $sysId
            $aCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@Name",[System.Data.SqlDbType]::VarChar,255))).Value = $app.DisplayName
            $aCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@Ver",[System.Data.SqlDbType]::VarChar,100))).Value = if ($app.DisplayVersion) { $app.DisplayVersion } else { [DBNull]::Value }
            $aCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@Pub",[System.Data.SqlDbType]::VarChar,255))).Value = if ($app.Publisher) { $app.Publisher } else { [DBNull]::Value }
            $aCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@Date",[System.Data.SqlDbType]::Date))).Value = if ($app.InstallDate -match "^\d{8}$") { [DateTime]::ParseExact($app.InstallDate,"yyyyMMdd",$null) } elseif ($app.InstallDate) { [DateTime]$app.InstallDate } else { [DBNull]::Value }
            $aCmd.ExecuteNonQuery() | Out-Null
        }

        $tran.Commit()
        Write-Host "✅ Imported Asset $($json.AssetNumber)" -ForegroundColor Green

    } catch {
        $tran.Rollback()
        Write-Host "❌ Error importing $($file.Name): $($_.Exception.Message)" -ForegroundColor Red
    } finally {
        $conn.Close()
    }
}

Write-Host "All imports complete." -ForegroundColor Green
    