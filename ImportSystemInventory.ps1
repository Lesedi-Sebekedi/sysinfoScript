<#
.SYNOPSIS
    Imports system inventory data from JSON files to SQL Server using the new schema
.DESCRIPTION
    Processes JSON inventory files and imports them into a SQL database.
    Handles both new systems and updates to existing records.
.COMPANY
    North West Provincial Treasury
.AUTHOR
    Lesedi Sebekedi
.VERSION
    4.0
.SECURITY
    Requires SQL write permissions to the AssetDB database
.PARAMETER ReportsFolder
    Path to folder containing JSON files to import
.PARAMETER ConnectionString
    SQL Server connection string
#>

param(
    [string]$ReportsFolder = "$env:USERPROFILE\Desktop\SystemReports",
    [string]$ConnectionString = "Server=PTLSEBEKEDI;Database=AssetDB;Integrated Security=True"
)

#region Helper Functions

function Safe-AddSqlParameter {
    <#
    .SYNOPSIS
        Safely adds a parameter to a SQL command with proper null handling and optional size/precision.
    #>
    param(
        [Parameter(Mandatory)][System.Data.SqlClient.SqlCommand]$Command,
        [Parameter(Mandatory)][string]$Name,
        [Parameter()][object]$Value,
        [Parameter(Mandatory)][System.Data.SqlDbType]$Type,
        [int]$Size = 0,
        [int]$Precision = 0,
        [int]$Scale = 0
    )

    # Create parameter with or without size
    $param = if ($Size -gt 0) {
        $Command.Parameters.Add($Name, $Type, $Size)
    } else {
        $Command.Parameters.Add($Name, $Type)
    }

    # Apply numeric precision/scale for decimals
    if ($Type -eq [System.Data.SqlDbType]::Decimal -and $Precision -gt 0) {
        $param.Precision = [byte]$Precision
        $param.Scale     = [byte]$Scale
    }

    # Normalize value
    if ($null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value))) {
        $param.Value = [DBNull]::Value
    }
    elseif ($Value -is [array]) {
        $param.Value = ($Value -join ', ')
    }
    else {
        # Do NOT blindly ToString() – keep native type for correct SqlDbType
        $param.Value = $Value
    }
}

function Get-FormattedAssetNumber {
    <#
    .SYNOPSIS
        Extracts and formats the asset number from system data.
    #>
    param([Parameter(Mandatory)][PSObject]$SystemData)

    $raw = $SystemData.AssetNumber
    if ($null -ne $raw) {
        $text = $raw.ToString().Trim()
        if ($text) { return $text }
    }
    throw "Invalid or missing AssetNumber in system data."
}

function Get-SafeDateTime {
    <#
    .SYNOPSIS
        Best-effort DateTime parser returning [DBNull]::Value on failure.
    #>
    param([string]$DateString)

    if ([string]::IsNullOrWhiteSpace($DateString)) { return [DBNull]::Value }

    try {
        # Try parsing as DateTimeOffset
        $dto = [DateTimeOffset]::Parse($DateString)
        return $dto.DateTime
    } catch {
        # Fallback to DateTime parse
        try {
            $dt = [DateTime]::Parse($DateString)
            return $dt
        } catch {
            return [DBNull]::Value
        }
    }
}

function Get-SafeSqlDateTime {
    <#
    .SYNOPSIS
        Wrapper used by Software import to keep original call sites.
    #>
    param([string]$DateString)
    return (Get-SafeDateTime -DateString $DateString)
}

#endregion Helper Functions

#region Import Functions

function Import-SystemRecord {
    <#
    .SYNOPSIS
        Inserts or updates a system record in the database.
    #>
    param(
        [Parameter(Mandatory)][PSObject]$SystemData,
        [Parameter(Mandatory)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory)][System.Data.SqlClient.SqlTransaction]$Transaction
    )
    try {
        $assetNumber = Get-FormattedAssetNumber -SystemData $SystemData

        # Existence check
        $check = $Connection.CreateCommand()
        $check.Transaction = $Transaction
        $check.CommandText = "SELECT 1 FROM Systems WHERE AssetNumber = @AssetNumber"
        Safe-AddSqlParameter -Command $check -Name "@AssetNumber" -Value $assetNumber -Type ([System.Data.SqlDbType]::VarChar) -Size 50
        $exists = $null -ne $check.ExecuteScalar()

        $cmd = $Connection.CreateCommand()
        $cmd.Transaction = $Transaction
        $cmd.CommandText = if ($exists) {
@"
UPDATE Systems
SET HostName=@HostName,
    UUID=@UUID,
    SerialNumber=@SerialNumber,
    OS=@OS,
    Version=@Version,
    Architecture=@Architecture,
    Build=@Build,
    Manufacturer=@Manufacturer,
    Model=@Model,
    BootTime=@BootTime,
    BIOSVersion=@BIOSVersion,
    ScanDate=@ScanDate,
    PSVersion=@PSVersion
WHERE AssetNumber=@AssetNumber
"@
        } else {
@"
INSERT INTO Systems (
    AssetNumber, HostName, UUID, SerialNumber, OS, Version,
    Architecture, Build, Manufacturer, Model, BootTime,
    BIOSVersion, ScanDate, PSVersion
)
VALUES (
    @AssetNumber, @HostName, @UUID, @SerialNumber, @OS, @Version,
    @Architecture, @Build, @Manufacturer, @Model, @BootTime,
    @BIOSVersion, @ScanDate, @PSVersion
)
"@
        }

        # Parameters
        Safe-AddSqlParameter -Command $cmd -Name "@AssetNumber"  -Value $assetNumber                         -Type ([System.Data.SqlDbType]::VarChar) -Size 50
        Safe-AddSqlParameter -Command $cmd -Name "@HostName"     -Value $SystemData.System.HostName          -Type ([System.Data.SqlDbType]::VarChar) -Size 100
        Safe-AddSqlParameter -Command $cmd -Name "@UUID"         -Value $SystemData.UUID                     -Type ([System.Data.SqlDbType]::VarChar) -Size 36
        Safe-AddSqlParameter -Command $cmd -Name "@SerialNumber" -Value $SystemData.System.BIOS.Serial       -Type ([System.Data.SqlDbType]::VarChar) -Size 100
        Safe-AddSqlParameter -Command $cmd -Name "@OS"           -Value $SystemData.System.OS                -Type ([System.Data.SqlDbType]::VarChar) -Size 100
        Safe-AddSqlParameter -Command $cmd -Name "@Version"      -Value $SystemData.System.Version           -Type ([System.Data.SqlDbType]::VarChar) -Size 100
        Safe-AddSqlParameter -Command $cmd -Name "@Architecture" -Value $SystemData.System.Architecture      -Type ([System.Data.SqlDbType]::VarChar) -Size 50
        Safe-AddSqlParameter -Command $cmd -Name "@Build"        -Value $SystemData.System.Build             -Type ([System.Data.SqlDbType]::VarChar) -Size 50
        Safe-AddSqlParameter -Command $cmd -Name "@Manufacturer" -Value $SystemData.System.Manufacturer      -Type ([System.Data.SqlDbType]::VarChar) -Size 100
        Safe-AddSqlParameter -Command $cmd -Name "@Model"        -Value $SystemData.System.Model             -Type ([System.Data.SqlDbType]::VarChar) -Size 100
        Safe-AddSqlParameter -Command $cmd -Name "@BootTime"     -Value (Get-SafeDateTime $SystemData.System.BootTime) -Type ([System.Data.SqlDbType]::DateTime)
        Safe-AddSqlParameter -Command $cmd -Name "@BIOSVersion"  -Value $SystemData.System.BIOS.Version      -Type ([System.Data.SqlDbType]::VarChar) -Size 100
        Safe-AddSqlParameter -Command $cmd -Name "@ScanDate"     -Value ([DateTime]::UtcNow)                 -Type ([System.Data.SqlDbType]::DateTime)
        Safe-AddSqlParameter -Command $cmd -Name "@PSVersion"    -Value $SystemData.PSVersion                -Type ([System.Data.SqlDbType]::VarChar) -Size 50

        $cmd.ExecuteNonQuery() | Out-Null
        return $assetNumber
    }
    catch {
        Write-Error "Failed to import system record: $_"
        throw
    }
}

function Import-Hardware {
    param(
        [Parameter(Mandatory)][string]$AssetNumber,
        [Parameter(Mandatory)][PSObject]$SystemData,
        [Parameter(Mandatory)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory)][System.Data.SqlClient.SqlTransaction]$Transaction
    )
    try {
        # Check existing hardware
        $check = $Connection.CreateCommand()
        $check.Transaction = $Transaction
        $check.CommandText = "SELECT HardwareID FROM Hardware WHERE AssetNumber = @AssetNumber"
        Safe-AddSqlParameter -Command $check -Name "@AssetNumber" -Value $AssetNumber -Type ([System.Data.SqlDbType]::VarChar) -Size 50
        $hardwareId = $check.ExecuteScalar()

        $cmd = $Connection.CreateCommand()
        $cmd.Transaction = $Transaction
        
        # Handle GPU data (single object or array)
        $gpuData = $SystemData.Hardware.GPU
        $gpuName = $null
        $gpuRam = 0
        $gpuDriver = $null
        
        if ($gpuData -is [array]) {
            # Take first GPU if multiple exist
            $firstGpu = $gpuData[0]
            $gpuName = $firstGpu.Name
            $gpuRam = [decimal]$firstGpu.AdapterRAMGB
            $gpuDriver = $firstGpu.DriverVersion
        }
        else {
            # Single GPU object
            $gpuName = $gpuData.Name
            $gpuRam = [decimal]$gpuData.AdapterRAMGB
            $gpuDriver = $gpuData.DriverVersion
        }

        $cmd.CommandText = if ($hardwareId) {
@"
UPDATE Hardware SET
    CPUName=@CPUName,
    CPUCores=@CPUCores,
    CPUThreads=@CPUThreads,
    CPUClockSpeed=@CPUClockSpeed,
    TotalRAMGB=@TotalRAMGB,
    PageFileGB=@PageFileGB,
    MemorySticks=@MemorySticks,
    GPUName=@GPUName,
    GPUAdapterRAMGB=@GPUAdapterRAMGB,
    GPUDriverVersion=@GPUDriverVersion
WHERE AssetNumber=@AssetNumber
"@
        } else {
@" 
INSERT INTO Hardware (
    AssetNumber, CPUName, CPUCores, CPUThreads, CPUClockSpeed,
    TotalRAMGB, PageFileGB, MemorySticks, GPUName, GPUAdapterRAMGB, GPUDriverVersion
) VALUES (
    @AssetNumber, @CPUName, @CPUCores, @CPUThreads, @CPUClockSpeed,
    @TotalRAMGB, @PageFileGB, @MemorySticks, @GPUName, @GPUAdapterRAMGB, @GPUDriverVersion
)
"@
        }

        # Parameters
        Safe-AddSqlParameter -Command $cmd -Name "@AssetNumber"       -Value $AssetNumber                              -Type ([System.Data.SqlDbType]::VarChar) -Size 50
        Safe-AddSqlParameter -Command $cmd -Name "@CPUName"           -Value $SystemData.Hardware.CPU.Name            -Type ([System.Data.SqlDbType]::VarChar) -Size 100
        Safe-AddSqlParameter -Command $cmd -Name "@CPUCores"          -Value $SystemData.Hardware.CPU.Cores           -Type ([System.Data.SqlDbType]::Int)
        Safe-AddSqlParameter -Command $cmd -Name "@CPUThreads"        -Value $SystemData.Hardware.CPU.Threads         -Type ([System.Data.SqlDbType]::Int)
        Safe-AddSqlParameter -Command $cmd -Name "@CPUClockSpeed"     -Value $SystemData.Hardware.CPU.ClockSpeed      -Type ([System.Data.SqlDbType]::VarChar) -Size 20
        Safe-AddSqlParameter -Command $cmd -Name "@TotalRAMGB"        -Value ([decimal]$SystemData.Hardware.Memory.TotalGB)      -Type ([System.Data.SqlDbType]::Decimal) -Precision 5 -Scale 2
        Safe-AddSqlParameter -Command $cmd -Name "@PageFileGB"        -Value ([decimal]$SystemData.Hardware.Memory.PageFileGB)   -Type ([System.Data.SqlDbType]::Decimal) -Precision 5 -Scale 2
        Safe-AddSqlParameter -Command $cmd -Name "@MemorySticks"      -Value $SystemData.Hardware.Memory.Sticks       -Type ([System.Data.SqlDbType]::Int)
        Safe-AddSqlParameter -Command $cmd -Name "@GPUName"           -Value $gpuName            -Type ([System.Data.SqlDbType]::VarChar) -Size 100
        Safe-AddSqlParameter -Command $cmd -Name "@GPUAdapterRAMGB"   -Value $gpuRam    -Type ([System.Data.SqlDbType]::Decimal) -Precision 5 -Scale 2
        Safe-AddSqlParameter -Command $cmd -Name "@GPUDriverVersion"  -Value $gpuDriver   -Type ([System.Data.SqlDbType]::VarChar) -Size 50

        $cmd.ExecuteNonQuery() | Out-Null

        # Rest of the function remains the same...
        # Ensure HardwareID for disks
        if (-not $hardwareId) {
            $getId = $Connection.CreateCommand()
            $getId.Transaction = $Transaction
            $getId.CommandText = "SELECT HardwareID FROM Hardware WHERE AssetNumber = @AssetNumber"
            Safe-AddSqlParameter -Command $getId -Name "@AssetNumber" -Value $AssetNumber -Type ([System.Data.SqlDbType]::VarChar) -Size 50
            $hardwareId = [int]$getId.ExecuteScalar()
        }

        if ($SystemData.Hardware.Disks) {
            Import-Disks -HardwareId $hardwareId -Disks $SystemData.Hardware.Disks -Connection $Connection -Transaction $Transaction
        }

        return $hardwareId
    }
    catch {
        Write-Error "Failed to import hardware specs: $_"
        throw
    }
}

function Import-Disks {
    <#
    .SYNOPSIS
        Imports disk information into the database.
    #>
    param(
        [Parameter(Mandatory)][int]$HardwareId,
        [Parameter()][PSObject[]]$Disks,
        [Parameter(Mandatory)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory)][System.Data.SqlClient.SqlTransaction]$Transaction
    )
    try {
        foreach ($disk in $Disks) {
            if (-not $disk.DeviceID) { continue }

            # Check existence
            $check = $Connection.CreateCommand()
            $check.Transaction = $Transaction
            $check.CommandText = "SELECT COUNT(*) FROM Disks WHERE HardwareID = @HardwareID AND DeviceID = @DeviceID"
            Safe-AddSqlParameter -Command $check -Name "@HardwareID" -Value $HardwareId -Type ([System.Data.SqlDbType]::Int)
            Safe-AddSqlParameter -Command $check -Name "@DeviceID"   -Value $disk.DeviceID -Type ([System.Data.SqlDbType]::VarChar) -Size 5
            $exists = [int]$check.ExecuteScalar() -gt 0

            $cmd = $Connection.CreateCommand()
            $cmd.Transaction = $Transaction
            $cmd.CommandText = if ($exists) {
@"
UPDATE Disks SET
    VolumeName=@VolumeName,
    SizeGB=@SizeGB,
    FreeGB=@FreeGB,
    Type=@Type
WHERE HardwareID=@HardwareID AND DeviceID=@DeviceID
"@
            } else {
@"
INSERT INTO Disks (
    HardwareID, DeviceID, VolumeName, SizeGB, FreeGB, Type
) VALUES (
    @HardwareID, @DeviceID, @VolumeName, @SizeGB, @FreeGB, @Type
)
"@
            }

            Safe-AddSqlParameter -Command $cmd -Name "@HardwareID" -Value $HardwareId          -Type ([System.Data.SqlDbType]::Int)
            Safe-AddSqlParameter -Command $cmd -Name "@DeviceID"   -Value $disk.DeviceID        -Type ([System.Data.SqlDbType]::VarChar) -Size 5
            Safe-AddSqlParameter -Command $cmd -Name "@VolumeName" -Value $disk.VolumeName      -Type ([System.Data.SqlDbType]::VarChar) -Size 100
            Safe-AddSqlParameter -Command $cmd -Name "@SizeGB"     -Value ([decimal]$disk.SizeGB) -Type ([System.Data.SqlDbType]::Decimal) -Precision 10 -Scale 2
            Safe-AddSqlParameter -Command $cmd -Name "@FreeGB"     -Value ([decimal]$disk.FreeGB) -Type ([System.Data.SqlDbType]::Decimal) -Precision 10 -Scale 2
            Safe-AddSqlParameter -Command $cmd -Name "@Type"       -Value $disk.Type            -Type ([System.Data.SqlDbType]::VarChar) -Size 20

            $cmd.ExecuteNonQuery() | Out-Null
        }
    }
    catch {
        Write-Error "Failed to import disk information: $_"
        throw
    }
}

function Convert-ToBool {
    param([string]$value)
    if ([string]::IsNullOrWhiteSpace($value)) { return $false }
    switch ($value.ToLower()) {
        'true'  { return $true }
        'false' { return $false }
        default { return $false }  # fallback for invalid values
    }
}

function Import-Network {
    param(
        [Parameter(Mandatory)][string]$AssetNumber,
        [Parameter(Mandatory)][pscustomobject]$NetworkData,
        [Parameter(Mandatory)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory)][System.Data.SqlClient.SqlTransaction]$Transaction
    )
    
    try {
        $cmd = $Connection.CreateCommand()
        $cmd.Transaction = $Transaction
        $cmd.CommandText = @"
INSERT INTO Network (
    AssetNumber, AdapterName, InterfaceDescription, MacAddress, 
    Speed, IPAddress, DHCPEnabled
) VALUES (
    @AssetNumber, @AdapterName, @InterfaceDescription, @MacAddress,
    @Speed, @IPAddress, @DHCPEnabled
)
"@

        Safe-AddSqlParameter -Command $cmd -Name "@AssetNumber" -Value $AssetNumber -Type ([System.Data.SqlDbType]::VarChar) -Size 50
        Safe-AddSqlParameter -Command $cmd -Name "@AdapterName" -Value $NetworkData.Name -Type ([System.Data.SqlDbType]::VarChar) -Size 100
        Safe-AddSqlParameter -Command $cmd -Name "@InterfaceDescription" -Value $NetworkData.InterfaceDescription -Type ([System.Data.SqlDbType]::VarChar) -Size 255
        Safe-AddSqlParameter -Command $cmd -Name "@MacAddress" -Value $NetworkData.MacAddress -Type ([System.Data.SqlDbType]::VarChar) -Size 20
        Safe-AddSqlParameter -Command $cmd -Name "@Speed" -Value $NetworkData.Speed -Type ([System.Data.SqlDbType]::VarChar) -Size 20
        Safe-AddSqlParameter -Command $cmd -Name "@IPAddress" -Value $NetworkData.IPAddress -Type ([System.Data.SqlDbType]::VarChar) -Size 50
        Safe-AddSqlParameter -Command $cmd -Name "@DHCPEnabled" -Value $false -Type ([System.Data.SqlDbType]::Bit) # Default value

        $cmd.ExecuteNonQuery() | Out-Null
    }
    catch {
        Write-Error "Failed to import network information for asset ${AssetNumber}: $_"
        throw
    }
}


function Import-Software {
    <#
    .SYNOPSIS
        Imports installed applications and hotfixes into the database.
    #>
    param(
        [Parameter(Mandatory)][string]$AssetNumber,
        [Parameter()][PSObject]$SoftwareData,
        [Parameter(Mandatory)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory)][System.Data.SqlClient.SqlTransaction]$Transaction
    )
    try {
        # Applications
        if ($SoftwareData.InstalledApps) {
            foreach ($app in $SoftwareData.InstalledApps) {
                if (-not $app.DisplayName) { continue }

                $check = $Connection.CreateCommand()
                $check.Transaction = $Transaction
                $check.CommandText = "SELECT COUNT(*) FROM Software WHERE AssetNumber=@AssetNumber AND AppName=@AppName AND IsApplication=1"
                $check.Parameters.AddWithValue("@AssetNumber", $AssetNumber) | Out-Null
                $check.Parameters.AddWithValue("@AppName", $app.DisplayName) | Out-Null
                $exists = [int]$check.ExecuteScalar() -gt 0

                $installDate = Get-SafeSqlDateTime -DateString $app.InstallDate

                $cmd = $Connection.CreateCommand()
                $cmd.Transaction = $Transaction
                $cmd.CommandText = if ($exists) {
                    "UPDATE Software SET AppVersion=@AppVersion, Publisher=@Publisher, InstallDate=@InstallDate WHERE AssetNumber=@AssetNumber AND AppName=@AppName AND IsApplication=1"
                } else {
                    "INSERT INTO Software (AssetNumber, AppName, AppVersion, Publisher, InstallDate, IsApplication) VALUES (@AssetNumber, @AppName, @AppVersion, @Publisher, @InstallDate, 1)"
                }

                $cmd.Parameters.AddWithValue("@AssetNumber", $AssetNumber) | Out-Null
                $cmd.Parameters.AddWithValue("@AppName", $app.DisplayName) | Out-Null
                $cmd.Parameters.AddWithValue("@AppVersion", ($app.DisplayVersion ?? [DBNull]::Value)) | Out-Null
                $cmd.Parameters.AddWithValue("@Publisher", ($app.Publisher ?? [DBNull]::Value)) | Out-Null
                $cmd.Parameters.AddWithValue("@InstallDate", $installDate) | Out-Null

                $cmd.ExecuteNonQuery() | Out-Null
            }
        }

        # Hotfixes
        if ($SoftwareData.Hotfixes) {
            foreach ($hotfix in $SoftwareData.Hotfixes) {
                if (-not $hotfix.HotFixID) { continue }

                $check = $Connection.CreateCommand()
                $check.Transaction = $Transaction
                $check.CommandText = "SELECT COUNT(*) FROM Software WHERE AssetNumber=@AssetNumber AND HotFixID=@HotFixID AND IsApplication=0"
                $check.Parameters.AddWithValue("@AssetNumber", $AssetNumber) | Out-Null
                $check.Parameters.AddWithValue("@HotFixID", $hotfix.HotFixID) | Out-Null
                $exists = [int]$check.ExecuteScalar() -gt 0

                $installedDate =
                    if ($hotfix.InstalledOn -and $hotfix.InstalledOn.DateTime) {
                        Get-SafeSqlDateTime -DateString $hotfix.InstalledOn.DateTime
                    } else { [DBNull]::Value }

                $cmd = $Connection.CreateCommand()
                $cmd.Transaction = $Transaction
                $cmd.CommandText = if ($exists) {
                    "UPDATE Software SET HotFixDescription=@HotFixDescription, HotFixInstalledDate=@HotFixInstalledDate WHERE AssetNumber=@AssetNumber AND HotFixID=@HotFixID AND IsApplication=0"
                } else {
                    "INSERT INTO Software (AssetNumber, HotFixID, HotFixDescription, HotFixInstalledDate, IsApplication) VALUES (@AssetNumber, @HotFixID, @HotFixDescription, @HotFixInstalledDate, 0)"
                }

                $cmd.Parameters.AddWithValue("@AssetNumber", $AssetNumber) | Out-Null
                $cmd.Parameters.AddWithValue("@HotFixID", $hotfix.HotFixID) | Out-Null
                $cmd.Parameters.AddWithValue("@HotFixDescription", ($hotfix.Description ?? [DBNull]::Value)) | Out-Null
                $cmd.Parameters.AddWithValue("@HotFixInstalledDate", $installedDate) | Out-Null

                $cmd.ExecuteNonQuery() | Out-Null
            }
        }
    }
    catch {
        Write-Error "Failed to import software information: $_"
        throw
    }
}

#endregion Import Functions

#region Summary / Reporting

function Show-SummaryReport {
    <#
    .SYNOPSIS
        Displays a summary report of the import process.
    #>
    param([Parameter(Mandatory)][hashtable]$SummaryData)

    Write-Host "`nIMPORT SUMMARY REPORT" -ForegroundColor Magenta
    Write-Host "====================" -ForegroundColor Magenta
    Write-Host "Start Time: $($SummaryData.StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan
    Write-Host "End Time:   $($SummaryData.EndTime.ToString('yyyy-MM-dd HH:mm:ss'))"   -ForegroundColor Cyan
    Write-Host "Duration:   $($SummaryData.Duration.ToString('hh\:mm\:ss'))"           -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor DarkGray
    Write-Host "Total files processed: $($SummaryData.TotalFiles)" -ForegroundColor Cyan
    Write-Host "Successfully imported: $($SummaryData.SuccessCount)" -ForegroundColor Green
    Write-Host "Failed imports:       $($SummaryData.FailedCount)"   -ForegroundColor Red

    if ($SummaryData.SuccessCount -gt 0) {
        Write-Host "`nSUCCESSFUL IMPORTS ($($SummaryData.SuccessAssets.Count)):" -ForegroundColor Green
        $SummaryData.SuccessAssets | ForEach-Object { Write-Host "  - $_" -ForegroundColor Green }
    }

    if ($SummaryData.FailedCount -gt 0) {
        Write-Host "`nFAILED IMPORTS ($($SummaryData.FailedAssets.Count)):" -ForegroundColor Red
        $SummaryData.FailedAssets | ForEach-Object {
            Write-Host "  - File:  $($_.FileName)"    -ForegroundColor Red
            Write-Host "    Asset: $($_.AssetNumber)" -ForegroundColor Yellow
            Write-Host "    Error: $($_.Error)"       -ForegroundColor DarkYellow
            Write-Host "    ----------------------------" -ForegroundColor DarkGray
        }
    }

    Write-Host "`nImport process completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Magenta
}

#endregion Summary / Reporting

#region Main Execution

# Initialize summary tracking
$summary = @{
    TotalFiles    = 0
    SuccessCount  = 0
    FailedCount   = 0
    SuccessAssets = [System.Collections.Generic.List[string]]::new()
    FailedAssets  = [System.Collections.Generic.List[hashtable]]::new()
    StartTime     = Get-Date
}

try {
    # Load required assembly for SQL operations
    Add-Type -AssemblyName "System.Data" -ErrorAction Stop

    # Get all JSON files in the reports folder
    $jsonFiles = Get-ChildItem -Path $ReportsFolder -Filter *.json -File -ErrorAction Stop
    $summary.TotalFiles = $jsonFiles.Count

    if ($jsonFiles.Count -eq 0) {
        Write-Host "No JSON files found in $ReportsFolder" -ForegroundColor Yellow
        exit 0
    }

    foreach ($file in $jsonFiles) {
        Write-Host "`nProcessing $($file.Name)..." -ForegroundColor Cyan

        $conn = $null
        $tran = $null
        $assetNumber = $null
        $json = $null

        try {
            # Load JSON data from file
            $json = Get-Content $file.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop

            # Set up database connection with transaction
            $conn = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
            $conn.Open()
            $tran = $conn.BeginTransaction()

            # Import all data
            $assetNumber = Import-SystemRecord -SystemData $json -Connection $conn -Transaction $tran
            Import-Hardware -AssetNumber $assetNumber -SystemData $json -Connection $conn -Transaction $tran
            Import-Network  -AssetNumber $assetNumber -NetworkData $json.Network   -Connection $conn -Transaction $tran
            Import-Software -AssetNumber $assetNumber -SoftwareData $json.Software -Connection $conn -Transaction $tran

            # Commit transaction if all operations succeeded
            $tran.Commit()
            Write-Host "✅ Successfully imported $assetNumber" -ForegroundColor Green

            # Update success summary
            $summary.SuccessCount++
            $summary.SuccessAssets.Add($assetNumber)
        }
        catch {
            # Rollback transaction on error
            if ($tran -and $tran.Connection -eq $conn) {
                try { $tran.Rollback(); Write-Host "❌ Transaction rolled back for $($file.Name)" -ForegroundColor Red }
                catch { Write-Host "❌ Failed to rollback transaction: $_" -ForegroundColor DarkRed }
            }

            $errorMsg = "Error importing $($file.Name): $($_.Exception.Message)"
            Write-Host "❌ $errorMsg" -ForegroundColor Red
            Write-Host "Error details: $($_.ScriptStackTrace)" -ForegroundColor DarkYellow

            # Update failure summary
            $summary.FailedCount++
            $summary.FailedAssets.Add(@{
                FileName    = $file.Name
                AssetNumber = if ($assetNumber) { $assetNumber } else { try { $json.AssetNumber } catch { "N/A" } }
                Error       = $errorMsg
            })
        }
        finally {
            if ($conn) {
                if ($conn.State -eq [System.Data.ConnectionState]::Open) { $conn.Close() }
                $conn.Dispose()
            }
        }
    }
}
catch {
    Write-Host "❌ Fatal error during import process: $($_.Exception.Message)" -ForegroundColor Red
    $summary.FailedCount = $summary.TotalFiles
}
finally {
    $summary.EndTime = Get-Date
    $summary.Duration = $summary.EndTime - $summary.StartTime
    Show-SummaryReport -SummaryData $summary
}

#endregion Main Execution

# Return summary object if needed by calling script
$summary
