<#
.SYNOPSIS
    Imports system inventory data from JSON files to SQL Server
.DESCRIPTION
    Processes JSON inventory files and imports them into a SQL database
    Handles both new systems and updates to existing records
.COMPANY
    North West Provincial Treasury
.AUTHOR
    Lesedi Sebekedi
.VERSION
    2.0
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

# Load required assembly for SQL operations
Add-Type -AssemblyName "System.Data"

function Import-SystemRecord {
    <#
    .SYNOPSIS
        Imports a single system record into the database
    .DESCRIPTION
        Handles both new system creation and updates to existing systems
        with full transaction support
    #>
    param(
        [PSObject]$SystemData,
        [System.Data.SqlClient.SqlConnection]$Connection,
        [System.Data.SqlClient.SqlTransaction]$Transaction
    )
    
    try {
        # Check if system already exists
        $checkCmd = $Connection.CreateCommand()
        $checkCmd.Transaction = $Transaction
        $checkCmd.CommandText = "SELECT SystemID FROM Systems WHERE AssetNumber=@AssetNumber OR UUID=@UUID"
        
        $checkCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@AssetNumber", [System.Data.SqlDbType]::VarChar, 50))).Value = $SystemData.AssetNumber
        $checkCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@UUID", [System.Data.SqlDbType]::VarChar, 36))).Value = $SystemData.UUID
        
        $systemId = $checkCmd.ExecuteScalar()

        if ($systemId) {
            # Update existing system record
            $updateCmd = $Connection.CreateCommand()
            $updateCmd.Transaction = $Transaction
            $updateCmd.CommandText = @"
                UPDATE Systems 
                SET HostName = @HostName, 
                    SerialNumber = @SerialNumber, 
                    ScanDate = @ScanDate 
                WHERE SystemID = @SystemID
"@
            
            $updateCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@HostName", [System.Data.SqlDbType]::VarChar, 100))).Value = $SystemData.System.HostName
            $updateCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@SerialNumber", [System.Data.SqlDbType]::VarChar, 100))).Value = $SystemData.System.BIOS.Serial
            $updateCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@ScanDate", [System.Data.SqlDbType]::DateTime))).Value = [DateTime]::Now
            $updateCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@SystemID", [System.Data.SqlDbType]::Int))).Value = $systemId
            
            $updateCmd.ExecuteNonQuery() | Out-Null
        }
        else {
            # Insert new system record
            $insertCmd = $Connection.CreateCommand()
            $insertCmd.Transaction = $Transaction
            $insertCmd.CommandText = @"
                INSERT INTO Systems (AssetNumber, HostName, UUID, SerialNumber, ScanDate)
                VALUES (@AssetNumber, @HostName, @UUID, @SerialNumber, @ScanDate);
                SELECT SCOPE_IDENTITY();
"@
            
            $insertCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@AssetNumber", [System.Data.SqlDbType]::VarChar, 50))).Value = $SystemData.AssetNumber
            $insertCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@HostName", [System.Data.SqlDbType]::VarChar, 100))).Value = $SystemData.System.HostName
            $insertCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@UUID", [System.Data.SqlDbType]::VarChar, 36))).Value = $SystemData.UUID
            $insertCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@SerialNumber", [System.Data.SqlDbType]::VarChar, 100))).Value = $SystemData.System.BIOS.Serial
            $insertCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@ScanDate", [System.Data.SqlDbType]::DateTime))).Value = [DateTime]::Now
            
            $systemId = [int]$insertCmd.ExecuteScalar()
        }

        return $systemId
    }
    catch {
        Write-Error "Failed to import system record: $_"
        throw
    }
}

function Clear-SystemDetails {
    <#
    .SYNOPSIS
        Clears existing system details before importing new ones
    #>
    param(
        [int]$SystemId,
        [System.Data.SqlClient.SqlConnection]$Connection,
        [System.Data.SqlClient.SqlTransaction]$Transaction,
        [string[]]$Tables = @("SystemSpecs", "SystemDisks", "InstalledApps")
    )
    
    foreach ($table in $Tables) {
        $deleteCmd = $Connection.CreateCommand()
        $deleteCmd.Transaction = $Transaction
        $deleteCmd.CommandText = "DELETE FROM $table WHERE SystemID = @SystemID"
        
        $deleteCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@SystemID", [System.Data.SqlDbType]::Int))).Value = $SystemId
        
        $deleteCmd.ExecuteNonQuery() | Out-Null
    }
}

function Import-SystemSpecs {
    <#
    .SYNOPSIS
        Imports system specifications into the database
    #>
    param(
        [int]$SystemId,
        [PSObject]$SystemData,
        [System.Data.SqlClient.SqlConnection]$Connection,
        [System.Data.SqlClient.SqlTransaction]$Transaction
    )
    
    try {
        $specCmd = $Connection.CreateCommand()
        $specCmd.Transaction = $Transaction
        $specCmd.CommandText = @"
            INSERT INTO SystemSpecs (
                SystemID, OSName, OSVersion, OSBuild, Architecture,
                LastBoot, Manufacturer, Model, BIOSVersion,
                CPUCores, CPUThreads, CPUClockSpeed, TotalRAMGB
            )
            VALUES (
                @SystemID, @OSName, @OSVersion, @OSBuild, @Architecture,
                @LastBoot, @Manufacturer, @Model, @BIOSVersion,
                @CPUCores, @CPUThreads, @CPUClockSpeed, @TotalRAMGB
            )
"@

        # Add all parameters with proper data types
        $specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@SystemID", [System.Data.SqlDbType]::Int))).Value = $SystemId
        $specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@OSName", [System.Data.SqlDbType]::VarChar, 100))).Value = $SystemData.System.OS
        $specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@OSVersion", [System.Data.SqlDbType]::VarChar, 50))).Value = $SystemData.System.Version
        $specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@OSBuild", [System.Data.SqlDbType]::VarChar, 20))).Value = $SystemData.System.Build
        $specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@Architecture", [System.Data.SqlDbType]::VarChar, 10))).Value = $SystemData.System.Architecture
        $specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@LastBoot", [System.Data.SqlDbType]::DateTime))).Value = [DateTime]::ParseExact($SystemData.System.BootTime, 'yyyy-MM-dd HH:mm:ss', $null)
        $specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@Manufacturer", [System.Data.SqlDbType]::VarChar, 100))).Value = $SystemData.System.Manufacturer
        $specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@Model", [System.Data.SqlDbType]::VarChar, 100))).Value = $SystemData.System.Model
        $specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@BIOSVersion", [System.Data.SqlDbType]::VarChar, 100))).Value = $SystemData.System.BIOS.Version
        $specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@CPUCores", [System.Data.SqlDbType]::Int))).Value = $SystemData.Hardware.CPU.Cores
        $specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@CPUThreads", [System.Data.SqlDbType]::Int))).Value = $SystemData.Hardware.CPU.Threads
        $specCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@CPUClockSpeed", [System.Data.SqlDbType]::VarChar, 20))).Value = $SystemData.Hardware.CPU.ClockSpeed
        
        # Special handling for decimal parameter
        $ramParam = New-Object System.Data.SqlClient.SqlParameter("@TotalRAMGB", [System.Data.SqlDbType]::Decimal)
        $ramParam.Precision = 5
        $ramParam.Scale = 2
        $ramParam.Value = $SystemData.Hardware.Memory.TotalGB
        $specCmd.Parameters.Add($ramParam)

        $specCmd.ExecuteNonQuery() | Out-Null
    }
    catch {
        Write-Error "Failed to import system specs: $_"
        throw
    }
}

function Import-Disks {
    <#
    .SYNOPSIS
        Imports disk information into the database
    #>
    param(
        [int]$SystemId,
        [PSObject[]]$Disks,
        [System.Data.SqlClient.SqlConnection]$Connection,
        [System.Data.SqlClient.SqlTransaction]$Transaction
    )
    
    try {
        $diskCmd = $Connection.CreateCommand()
        $diskCmd.Transaction = $Transaction
        $diskCmd.CommandText = @"
            INSERT INTO SystemDisks (
                SystemID, DriveLetter, SizeGB, FreeGB, DiskType
            )
            VALUES (
                @SystemID, @DriveLetter, @SizeGB, @FreeGB, @DiskType
            )
"@

        # Add parameters that will be reused for each disk
        $diskCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@SystemID", [System.Data.SqlDbType]::Int))).Value = $SystemId
        $diskCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@DriveLetter", [System.Data.SqlDbType]::VarChar, 5))).Value = $null
        $diskCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@SizeGB", [System.Data.SqlDbType]::Decimal))).Value = $null
        $diskCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@FreeGB", [System.Data.SqlDbType]::Decimal))).Value = $null
        $diskCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@DiskType", [System.Data.SqlDbType]::VarChar, 20))).Value = $null

        foreach ($disk in $Disks) {
            $diskCmd.Parameters["@DriveLetter"].Value = $disk.DeviceID
            $diskCmd.Parameters["@SizeGB"].Value = $disk.SizeGB
            $diskCmd.Parameters["@FreeGB"].Value = $disk.FreeGB
            $diskCmd.Parameters["@DiskType"].Value = $disk.Type
            
            $diskCmd.ExecuteNonQuery() | Out-Null
        }
    }
    catch {
        Write-Error "Failed to import disk information: $_"
        throw
    }
}

function Import-Applications {
    <#
    .SYNOPSIS
        Imports installed applications into the database
    #>
    param(
        [int]$SystemId,
        [PSObject[]]$Applications,
        [System.Data.SqlClient.SqlConnection]$Connection,
        [System.Data.SqlClient.SqlTransaction]$Transaction
    )
    
    try {
        $appCmd = $Connection.CreateCommand()
        $appCmd.Transaction = $Transaction
        $appCmd.CommandText = @"
            INSERT INTO InstalledApps (
                SystemID, AppName, AppVersion, Publisher, InstallDate
            )
            VALUES (
                @SystemID, @AppName, @AppVersion, @Publisher, @InstallDate
            )
"@

        # Add parameters with proper null handling
        $appCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@SystemID", [System.Data.SqlDbType]::Int))).Value = $SystemId
        $appCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@AppName", [System.Data.SqlDbType]::VarChar, 255))).Value = $null
        $appCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@AppVersion", [System.Data.SqlDbType]::VarChar, 100))).Value = $null
        $appCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@Publisher", [System.Data.SqlDbType]::VarChar, 255))).Value = $null
        $appCmd.Parameters.Add((New-Object System.Data.SqlClient.SqlParameter("@InstallDate", [System.Data.SqlDbType]::Date))).Value = $null

        foreach ($app in $Applications) {
            if (-not $app.DisplayName) { continue }

            $appCmd.Parameters["@AppName"].Value = $app.DisplayName
            
            # Handle null/empty values properly
            if ($app.DisplayVersion) {
                $appCmd.Parameters["@AppVersion"].Value = $app.DisplayVersion
            } else {
                $appCmd.Parameters["@AppVersion"].Value = [DBNull]::Value
            }
            
            if ($app.Publisher) {
                $appCmd.Parameters["@Publisher"].Value = $app.Publisher
            } else {
                $appCmd.Parameters["@Publisher"].Value = [DBNull]::Value
            }
            
            # Parse install date with proper format handling
            if ($app.InstallDate -match "^\d{8}$") {
                $appCmd.Parameters["@InstallDate"].Value = [DateTime]::ParseExact($app.InstallDate, "yyyyMMdd", $null)
            }
            elseif ($app.InstallDate) {
                try {
                    $appCmd.Parameters["@InstallDate"].Value = [DateTime]$app.InstallDate
                }
                catch {
                    $appCmd.Parameters["@InstallDate"].Value = [DBNull]::Value
                }
            }
            else {
                $appCmd.Parameters["@InstallDate"].Value = [DBNull]::Value
            }
            
            $appCmd.ExecuteNonQuery() | Out-Null
        }
    }
    catch {
        Write-Error "Failed to import applications: $_"
        throw
    }
}

# MAIN SCRIPT EXECUTION
# ---------------------

# Get all JSON files in the reports folder
$jsonFiles = Get-ChildItem -Path $ReportsFolder -Filter *.json -File

if ($jsonFiles.Count -eq 0) {
    Write-Host "No JSON files found in $ReportsFolder" -ForegroundColor Yellow
    exit 0
}

foreach ($file in $jsonFiles) {
    Write-Host "Importing $($file.Name)..." -ForegroundColor Cyan

    try {
        # Load JSON data from file
        $json = Get-Content $file.FullName -Raw | ConvertFrom-Json
        
        # Set up database connection with transaction
        $conn = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
        $conn.Open()
        $tran = $conn.BeginTransaction()

        # Import or update system record
        $systemId = Import-SystemRecord -SystemData $json -Connection $conn -Transaction $tran

        # Clear existing details before importing new ones
        Clear-SystemDetails -SystemId $systemId -Connection $conn -Transaction $tran

        # Import system specifications
        Import-SystemSpecs -SystemId $systemId -SystemData $json -Connection $conn -Transaction $tran

        # Import disk information
        Import-Disks -SystemId $systemId -Disks $json.Hardware.Disks -Connection $conn -Transaction $tran

        # Import installed applications
        Import-Applications -SystemId $systemId -Applications $json.Software.InstalledApps -Connection $conn -Transaction $tran

        # Commit transaction if all operations succeeded
        $tran.Commit()
        Write-Host "✅ Successfully imported $($json.AssetNumber)" -ForegroundColor Green
    }
    catch {
        # Rollback transaction on error
        if ($tran -and $tran.Connection -eq $conn) {
            $tran.Rollback()
            Write-Host "❌ Transaction rolled back for $($file.Name)" -ForegroundColor Red
        }
        
        Write-Host "❌ Error importing $($file.Name): $($_.Exception.Message)" -ForegroundColor Red
    }
    finally {
        # Ensure connection is closed
        if ($conn -and $conn.State -eq [System.Data.ConnectionState]::Open) {
            $conn.Close()
        }
    }
}

Write-Host "`nImport process completed. Processed $($jsonFiles.Count) files." -ForegroundColor Green