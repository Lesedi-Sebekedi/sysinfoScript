<#
.SYNOPSIS
    Enhanced system inventory import script with improved network data handling
.DESCRIPTION
    Processes JSON inventory files and imports them into SQL database with:
    - Better network data validation
    - Duplicate data prevention
    - Comprehensive error handling
.COMPANY
    North West Provincial Treasury
.AUTHOR
    Lesedi Sebekedi
.VERSION
    4.1
#>

param(
    [ValidateNotNullOrEmpty()]
    [string]$ReportsFolder = "$env:USERPROFILE\Desktop\SystemReports",
    
    [ValidateNotNullOrEmpty()]
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

function Convert-CIDRToSubnetMask {
    param([int]$CIDR)
    
    if ($CIDR -lt 0 -or $CIDR -gt 32) {
        throw "CIDR must be between 0 and 32, got: $CIDR"
    }
    
    try {
        $mask = [System.Net.IPAddress]::Parse(([math]::Pow(2, 32) - [math]::Pow(2, (32 - $CIDR))).ToString())
        return $mask
    }
    catch {
        throw "Failed to convert CIDR $CIDR to subnet mask: $_"
    }
}

function Test-ValidIPAddress {
    param([string]$IPAddress)
    
    if ([string]::IsNullOrWhiteSpace($IPAddress)) { return $false }
    return $IPAddress -match '^(\d{1,3}\.){3}\d{1,3}$'
}

function Test-JsonStructure {
    param([Parameter(Mandatory)][PSObject]$JsonData)
    
    $errors = @()
    
    # Check required top-level sections
    if (-not $JsonData.System) { $errors += "Missing 'System' section" }
    if (-not $JsonData.Hardware) { $errors += "Missing 'Hardware' section" }
    if (-not $JsonData.Network) { $errors += "Missing 'Network' section" }
    if (-not $JsonData.Software) { $errors += "Missing 'Software' section" }
    if (-not $JsonData.AssetNumber) { $errors += "Missing 'AssetNumber' field" }
    
    # Check System section
    if ($JsonData.System) {
        if (-not $JsonData.System.HostName) { $errors += "Missing 'System.HostName'" }
        if (-not $JsonData.System.OS) { $errors += "Missing 'System.OS'" }
        if (-not $JsonData.System.BIOS) { $errors += "Missing 'System.BIOS' section" }
        if ($JsonData.System.BIOS -and -not $JsonData.System.BIOS.Serial) { $errors += "Missing 'System.BIOS.Serial'" }
    }
    
    # Check Hardware section
    if ($JsonData.Hardware) {
        if (-not $JsonData.Hardware.CPU) { $errors += "Missing 'Hardware.CPU' section" }
        if (-not $JsonData.Hardware.Memory) { $errors += "Missing 'Hardware.Memory' section" }
        if ($JsonData.Hardware.CPU -and -not $JsonData.Hardware.CPU.Name) { $errors += "Missing 'Hardware.CPU.Name'" }
        if ($JsonData.Hardware.Memory -and -not $JsonData.Hardware.Memory.TotalGB) { $errors += "Missing 'Hardware.Memory.TotalGB'" }
    }
    
    # Check Network section
    if ($JsonData.Network) {
        if (-not $JsonData.Network.MacAddress) { $errors += "Missing 'Network.MacAddress'" }
        if (-not $JsonData.Network.Name) { $errors += "Missing 'Network.Name'" }
    }
    
    if ($errors.Count -gt 0) {
        throw "JSON validation failed:`n" + ($errors -join "`n")
    }
    
    return $true
}

function Test-DatabaseSchema {
    param([Parameter(Mandatory)][System.Data.SqlClient.SqlConnection]$Connection)
    
    $errors = @()
    
    try {
        # Test Systems table
        $cmd = $Connection.CreateCommand()
        $cmd.CommandText = @"
SELECT COLUMN_NAME, DATA_TYPE 
FROM INFORMATION_SCHEMA.COLUMNS 
WHERE TABLE_NAME = 'Systems' 
ORDER BY ORDINAL_POSITION
"@
        $result = $cmd.ExecuteReader()
        $systemsColumns = @()
        while ($result.Read()) {
            $systemsColumns += $result["COLUMN_NAME"]
        }
        $result.Close()
        
        # Test Network table
        $cmd.CommandText = @"
SELECT COLUMN_NAME, DATA_TYPE 
FROM INFORMATION_SCHEMA.COLUMNS 
WHERE TABLE_NAME = 'Network' 
ORDER BY ORDINAL_POSITION
"@
        $result = $cmd.ExecuteReader()
        $networkColumns = @()
        while ($result.Read()) {
            $networkColumns += $result["COLUMN_NAME"]
        }
        $result.Close()
        
        # Test Hardware table
        $cmd.CommandText = @"
SELECT COLUMN_NAME, DATA_TYPE 
FROM INFORMATION_SCHEMA.COLUMNS 
WHERE TABLE_NAME = 'Hardware' 
ORDER BY ORDINAL_POSITION
"@
        $result = $cmd.ExecuteReader()
        $hardwareColumns = @()
        while ($result.Read()) {
            $hardwareColumns += $result["COLUMN_NAME"]
        }
        $result.Close()
        
        # Test Software table
        $cmd.CommandText = @"
SELECT COLUMN_NAME, DATA_TYPE 
FROM INFORMATION_SCHEMA.COLUMNS 
WHERE TABLE_NAME = 'Software' 
ORDER BY ORDINAL_POSITION
"@
        $result = $cmd.ExecuteReader()
        $softwareColumns = @()
        while ($result.Read()) {
            $softwareColumns += $result["COLUMN_NAME"]
        }
        $result.Close()
        
        Write-Host "Database Schema Information:" -ForegroundColor Cyan
        Write-Host "Systems table columns: $($systemsColumns -join ', ')" -ForegroundColor DarkGray
        Write-Host "Network table columns: $($networkColumns -join ', ')" -ForegroundColor DarkGray
        Write-Host "Hardware table columns: $($hardwareColumns -join ', ')" -ForegroundColor DarkGray
        Write-Host "Software table columns: $($softwareColumns -join ', ')" -ForegroundColor DarkGray
        
        # Store schema info in global variables for use in other functions
        $script:DatabaseSchema = @{
            Systems = $systemsColumns
            Network = $networkColumns
            Hardware = $hardwareColumns
            Software = $softwareColumns
        }
        
        return $true
    }
    catch {
        Write-Warning "Could not retrieve database schema information: $_"
        return $false
    }
}

function Get-ColumnList {
    param(
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)][string[]]$RequiredColumns
    )
    
    if (-not $script:DatabaseSchema -or -not $script:DatabaseSchema[$TableName]) {
        return $RequiredColumns
    }
    
    $availableColumns = $script:DatabaseSchema[$TableName]
    $validColumns = @()
    
    foreach ($col in $RequiredColumns) {
        if ($availableColumns -contains $col) {
            $validColumns += $col
        } else {
            Write-Warning "Column '$col' not found in $TableName table, skipping..."
        }
    }
    
    return $validColumns
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
        Safe-AddSqlParameter -Command $cmd -Name "@PSVersion"    -Value ($SystemData.PSVersion ?? "Unknown") -Type ([System.Data.SqlDbType]::VarChar) -Size 50

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
            # Handle both AdapterRAM and AdapterRAMGB fields
            if ($firstGpu.AdapterRAMGB) {
                $gpuRam = [decimal]$firstGpu.AdapterRAMGB
            } elseif ($firstGpu.AdapterRAM) {
                $gpuRam = [decimal]($firstGpu.AdapterRAM / 1GB)
            } else {
                $gpuRam = 0
            }
            $gpuDriver = $firstGpu.DriverVersion
        }
        else {
            # Single GPU object
            $gpuName = $gpuData.Name
            # Handle both AdapterRAM and AdapterRAMGB fields
            if ($gpuData.AdapterRAMGB) {
                $gpuRam = [decimal]$gpuData.AdapterRAMGB
            } elseif ($gpuData.AdapterRAM) {
                $gpuRam = [decimal]($gpuData.AdapterRAM / 1GB)
            } else {
                $gpuRam = 0
            }
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

        # Ensure HardwareID for disks
        if (-not $hardwareId) {
            $getId = $Connection.CreateCommand()
            $getId.Transaction = $Transaction
            $getId.CommandText = "SELECT HardwareID FROM Hardware WHERE AssetNumber = @AssetNumber"
            Safe-AddSqlParameter -Command $getId -Name "@AssetNumber" -Value $AssetNumber -Type ([System.Data.SqlDbType]::VarChar) -Size 50
            $hardwareId = [int]$getId.ExecuteScalar()
        }

        # Handle Disks - convert single object to array if needed
        $disks = $SystemData.Hardware.Disks
        if ($disks) {
            if ($disks -isnot [array]) {
                $disks = @($disks)
            }
            Import-Disks -HardwareId $hardwareId -Disks $disks -Connection $Connection -Transaction $Transaction
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
        # Validate required network fields
        if (-not $NetworkData.MacAddress) {
            throw "Missing required MAC address in network data"
        }

        # Convert CIDR to subnet mask if needed
        $subnetMask = if ($NetworkData.SubnetMask -match '^\d+$') {
            try {
                Convert-CIDRToSubnetMask -CIDR $NetworkData.SubnetMask
            } catch {
                Write-Warning "Invalid CIDR value: $($NetworkData.SubnetMask), using as-is"
                $NetworkData.SubnetMask
            }
        } else {
            $NetworkData.SubnetMask
        }

        # Validate IP configuration
        $ipConfigType = if ($NetworkData.IPConfigType) {
            $NetworkData.IPConfigType
        } else {
            if ($NetworkData.DHCPEnabled) { "DHCP" } else { "Static" }
        }

        # Check for existing network adapter by MAC address
        $check = $Connection.CreateCommand()
        $check.Transaction = $Transaction
        $check.CommandText = "SELECT COUNT(*) FROM Network WHERE AssetNumber = @AssetNumber AND MacAddress = @MacAddress"
        Safe-AddSqlParameter -Command $check -Name "@AssetNumber" -Value $AssetNumber -Type ([System.Data.SqlDbType]::VarChar) -Size 50
        Safe-AddSqlParameter -Command $check -Name "@MacAddress" -Value $NetworkData.MacAddress -Type ([System.Data.SqlDbType]::VarChar) -Size 20
        $exists = [int]$check.ExecuteScalar() -gt 0

        $cmd = $Connection.CreateCommand()
        $cmd.Transaction = $Transaction
        $cmd.CommandText = if ($exists) {
@"
UPDATE Network SET
    AdapterName = @AdapterName,
    InterfaceDescription = @InterfaceDescription,
    Speed = @Speed,
    IPAddress = @IPAddress,
    IPConfigType = @IPConfigType,
    SubnetMask = @SubnetMask,
    DefaultGateway = @DefaultGateway,
    DNSServers = @DNSServers,
    DHCPEnabled = @DHCPEnabled
WHERE AssetNumber = @AssetNumber AND MacAddress = @MacAddress
"@
        } else {
@"
INSERT INTO Network (
    AssetNumber, AdapterName, InterfaceDescription, MacAddress,
    Speed, IPAddress, IPConfigType, SubnetMask, DefaultGateway, DNSServers,
    DHCPEnabled
) VALUES (
    @AssetNumber, @AdapterName, @InterfaceDescription, @MacAddress,
    @Speed, @IPAddress, @IPConfigType, @SubnetMask, @DefaultGateway, @DNSServers,
    @DHCPEnabled
)
"@
        }

        # Parameters with validation
        Safe-AddSqlParameter -Command $cmd -Name "@AssetNumber" -Value $AssetNumber -Type ([System.Data.SqlDbType]::VarChar) -Size 50
        Safe-AddSqlParameter -Command $cmd -Name "@AdapterName" -Value $NetworkData.Name -Type ([System.Data.SqlDbType]::VarChar) -Size 100
        Safe-AddSqlParameter -Command $cmd -Name "@InterfaceDescription" -Value $NetworkData.InterfaceDescription -Type ([System.Data.SqlDbType]::VarChar) -Size 255
        Safe-AddSqlParameter -Command $cmd -Name "@MacAddress" -Value $NetworkData.MacAddress -Type ([System.Data.SqlDbType]::VarChar) -Size 20
        Safe-AddSqlParameter -Command $cmd -Name "@Speed" -Value $NetworkData.Speed -Type ([System.Data.SqlDbType]::VarChar) -Size 20
        Safe-AddSqlParameter -Command $cmd -Name "@IPAddress" -Value $NetworkData.IPAddress -Type ([System.Data.SqlDbType]::VarChar) -Size 50
        Safe-AddSqlParameter -Command $cmd -Name "@IPConfigType" -Value $ipConfigType -Type ([System.Data.SqlDbType]::VarChar) -Size 20
        Safe-AddSqlParameter -Command $cmd -Name "@SubnetMask" -Value $subnetMask -Type ([System.Data.SqlDbType]::VarChar) -Size 20
        Safe-AddSqlParameter -Command $cmd -Name "@DefaultGateway" -Value $NetworkData.DefaultGateway -Type ([System.Data.SqlDbType]::VarChar) -Size 50
        Safe-AddSqlParameter -Command $cmd -Name "@DNSServers" -Value $NetworkData.DNSServers -Type ([System.Data.SqlDbType]::VarChar) -Size 255
        
        # Handle DHCPEnabled field - could be boolean or string
        $dhcpEnabled = if ($NetworkData.DHCPEnabled -is [bool]) {
            [int]$NetworkData.DHCPEnabled
        } elseif ($NetworkData.DHCPEnabled -is [string]) {
            [int](Convert-ToBool $NetworkData.DHCPEnabled)
        } else {
            0
        }
        Safe-AddSqlParameter -Command $cmd -Name "@DHCPEnabled" -Value $dhcpEnabled -Type ([System.Data.SqlDbType]::Bit)

        $rowsAffected = $cmd.ExecuteNonQuery()
        
        if ($rowsAffected -eq 0) {
            Write-Warning "No rows were affected in Network table for asset $AssetNumber, MAC: $($NetworkData.MacAddress)"
            # Don't throw an error, just log a warning
        } else {
            Write-Verbose "Successfully processed network adapter $($NetworkData.MacAddress) - $rowsAffected rows affected"
        }

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
                Safe-AddSqlParameter -Command $check -Name "@AssetNumber" -Value $AssetNumber -Type ([System.Data.SqlDbType]::VarChar) -Size 50
                Safe-AddSqlParameter -Command $check -Name "@AppName" -Value $app.DisplayName -Type ([System.Data.SqlDbType]::VarChar) -Size 255
                $exists = [int]$check.ExecuteScalar() -gt 0

                $installDate = Get-SafeSqlDateTime -DateString $app.InstallDate

                $cmd = $Connection.CreateCommand()
                $cmd.Transaction = $Transaction
                $cmd.CommandText = if ($exists) {
                    "UPDATE Software SET AppVersion=@AppVersion, Publisher=@Publisher, InstallDate=@InstallDate WHERE AssetNumber=@AssetNumber AND AppName=@AppName AND IsApplication=1"
                } else {
                    "INSERT INTO Software (AssetNumber, AppName, AppVersion, Publisher, InstallDate, IsApplication) VALUES (@AssetNumber, @AppName, @AppVersion, @Publisher, @InstallDate, 1)"
                }

                Safe-AddSqlParameter -Command $cmd -Name "@AssetNumber" -Value $AssetNumber -Type ([System.Data.SqlDbType]::VarChar) -Size 50
                Safe-AddSqlParameter -Command $cmd -Name "@AppName" -Value $app.DisplayName -Type ([System.Data.SqlDbType]::VarChar) -Size 255
                Safe-AddSqlParameter -Command $cmd -Name "@AppVersion" -Value ($app.DisplayVersion ?? [DBNull]::Value) -Type ([System.Data.SqlDbType]::VarChar) -Size 100
                Safe-AddSqlParameter -Command $cmd -Name "@Publisher" -Value ($app.Publisher ?? [DBNull]::Value) -Type ([System.Data.SqlDbType]::VarChar) -Size 255
                Safe-AddSqlParameter -Command $cmd -Name "@InstallDate" -Value $installDate -Type ([System.Data.SqlDbType]::DateTime)

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
                Safe-AddSqlParameter -Command $check -Name "@AssetNumber" -Value $AssetNumber -Type ([System.Data.SqlDbType]::VarChar) -Size 50
                Safe-AddSqlParameter -Command $check -Name "@HotFixID" -Value $hotfix.HotFixID -Type ([System.Data.SqlDbType]::VarChar) -Size 50
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

                Safe-AddSqlParameter -Command $cmd -Name "@AssetNumber" -Value $AssetNumber -Type ([System.Data.SqlDbType]::VarChar) -Size 50
                Safe-AddSqlParameter -Command $cmd -Name "@HotFixID" -Value $hotfix.HotFixID -Type ([System.Data.SqlDbType]::VarChar) -Size 50
                Safe-AddSqlParameter -Command $cmd -Name "@HotFixDescription" -Value ($hotfix.Description ?? [DBNull]::Value) -Type ([System.Data.SqlDbType]::VarChar) -Size 255
                Safe-AddSqlParameter -Command $cmd -Name "@HotFixInstalledDate" -Value $installedDate -Type ([System.Data.SqlDbType]::DateTime)

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
        Write-Host "File size: $([math]::Round($file.Length / 1KB, 2)) KB" -ForegroundColor DarkGray

        $conn = $null
        $tran = $null
        $assetNumber = $null
        $json = $null

        try {
            Write-Host "  📖 Loading JSON data..." -ForegroundColor DarkYellow
            $json = Get-Content $file.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop

            Write-Host "  ✅ JSON loaded successfully" -ForegroundColor DarkGreen
            Write-Host "  🔍 Validating JSON structure..." -ForegroundColor DarkYellow
            
            # Validate JSON structure before import
            Test-JsonStructure -JsonData $json
            Write-Host "  ✅ JSON validation passed" -ForegroundColor DarkGreen

            Write-Host "  🗄️  Connecting to database..." -ForegroundColor DarkYellow
            $conn = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
            $conn.Open()
            $tran = $conn.BeginTransaction()
            Write-Host "  ✅ Database connected" -ForegroundColor DarkGreen

            # Validate database schema
            Write-Host "  🔍 Validating database schema..." -ForegroundColor DarkYellow
            Test-DatabaseSchema -Connection $conn
            Write-Host "  ✅ Database schema validated" -ForegroundColor DarkGreen

            Write-Host "  📝 Importing system record..." -ForegroundColor DarkYellow
            $assetNumber = Import-SystemRecord -SystemData $json -Connection $conn -Transaction $tran
            Write-Host "  ✅ System record imported: $assetNumber" -ForegroundColor DarkGreen

            Write-Host "  🔧 Importing hardware information..." -ForegroundColor DarkYellow
            Import-Hardware -AssetNumber $assetNumber -SystemData $json -Connection $conn -Transaction $tran
            Write-Host "  ✅ Hardware information imported" -ForegroundColor DarkGreen
            
            Write-Host "  🌐 Importing network information..." -ForegroundColor DarkYellow
            # Handle Network data - could be single object or array
            $networkData = $json.Network
            if ($networkData -is [array]) {
                Write-Host "    📡 Processing $($networkData.Count) network adapters..." -ForegroundColor DarkCyan
                foreach ($network in $networkData) {
                    Import-Network -AssetNumber $assetNumber -NetworkData $network -Connection $conn -Transaction $tran
                }
            } else {
                Write-Host "    📡 Processing single network adapter..." -ForegroundColor DarkCyan
                Import-Network -AssetNumber $assetNumber -NetworkData $networkData -Connection $conn -Transaction $tran
            }
            Write-Host "  ✅ Network information imported" -ForegroundColor DarkGreen
            
            Write-Host "  💾 Importing software information..." -ForegroundColor DarkYellow
            Import-Software -AssetNumber $assetNumber -SoftwareData $json.Software -Connection $conn -Transaction $tran
            Write-Host "  ✅ Software information imported" -ForegroundColor DarkGreen

            Write-Host "  💾 Committing transaction..." -ForegroundColor DarkYellow
            $tran.Commit()
            Write-Host "✅ Successfully imported $assetNumber" -ForegroundColor Green

            # Update success summary
            $summary.SuccessCount++
            $summary.SuccessAssets.Add($assetNumber)
        }
        catch {
            if ($tran -and $tran.Connection -eq $conn) {
                try { $tran.Rollback() } catch { }
                Write-Host "❌ Transaction rolled back for $($file.Name)" -ForegroundColor Red
            }
            Write-Host "❌ Error importing $($file.Name): $_" -ForegroundColor Red

            # Update failure summary
            $summary.FailedCount++
            $summary.FailedAssets.Add(@{
                FileName    = $file.Name
                AssetNumber = if ($assetNumber) { $assetNumber } else { try { $json.AssetNumber } catch { "N/A" } }
                Error       = $_.Exception.Message
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
    Write-Host "❌ Fatal error during import process: $_" -ForegroundColor Red
}
finally {
    $summary.EndTime = Get-Date
    $summary.Duration = $summary.EndTime - $summary.StartTime
    Show-SummaryReport -SummaryData $summary
}

#endregion Main Execution

# Return summary object if needed by calling script
$summary
