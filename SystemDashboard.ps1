<#
.SYNOPSIS
    System Inventory Dashboard Web Server
.DESCRIPTION
    Creates a responsive local web dashboard to view system inventory from SQL database
    Features:
    - Paginated system listings with sorting
    - Detailed system views with tabs
    - Advanced search functionality
    - Automatic refresh
    - Secure HTML rendering
    - Performance monitoring
    - Export capabilities
.NOTES
    Author: Lesedi Sebekedi
    Version: 2.0
    Last Updated: $(Get-Date -Format "yyyy-MM-dd")
#>

#region CONFIGURATION
# Dashboard settings
$port = 8080  # Web server listening port
$dashboardTitle = "System Inventory Dashboard"
$companyName = "North West Provincial Treasury"
$defaultRowCount = 20  # Number of items per page
$connectionString = "Server=PTLSEBEKEDI;Database=AssetDB;Integrated Security=True;TrustServerCertificate=True"
$enablePerformanceLogging = $true
$performanceLogPath = "$env:TEMP\DashboardPerformance.log"
#endregion

# Load required SQL module
try {
    Import-Module SqlServer -ErrorAction Stop
}
catch {
    Write-Host "SQL Server module not found. Installing..."
    Install-Module -Name SqlServer -Force -AllowClobber -Scope CurrentUser
    Import-Module SqlServer
}

#region HTML TEMPLATES
# Main HTML structure with embedded CSS and JavaScript
$htmlHeader = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$dashboardTitle</title>
    <!-- External CSS libraries -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css">
    <style>
        /* Custom styling for dashboard elements */
        .system-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 10px 20px rgba(0,0,0,0.1);
            transition: all 0.3s ease;
        }
        .search-box {
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .logo-header {
            background: linear-gradient(135deg, #6e8efb 0%, #4a6cf7 100%);
        }
        .pagination .page-item.active .page-link {
            background-color: #4a6cf7;
            border-color: #4a6cf7;
        }
        .results-count {
            font-size: 0.9rem;
            color: #6c757d;
        }
        .app-icon {
            width: 24px;
            height: 24px;
            margin-right: 10px;
        }
        .back-button {
            margin-bottom: 20px;
        }
        .nav-tabs .nav-link.active {
            font-weight: bold;
            border-bottom: 3px solid #4a6cf7;
        }
        .badge-custom {
            font-size: 0.8em;
            font-weight: normal;
        }
        .disk-usage {
            height: 10px;
            border-radius: 5px;
        }
        .disk-usage-bar {
            height: 100%;
            border-radius: 5px;
        }
        .tab-content {
            padding: 15px 0;
        }
        .export-btn {
            margin-left: 10px;
        }
    </style>
</head>
<body>
    <!-- Navigation header with live clock -->
    <nav class="navbar navbar-expand-lg navbar-dark logo-header mb-4">
        <div class="container-fluid">
            <a class="navbar-brand" href="#">
                <i class="fas fa-server me-2"></i>$dashboardTitle
            </a>
            <div class="navbar-text ms-auto">
                <i class="fas fa-clock me-1"></i>
                <span id="live-clock"></span>
            </div>
        </div>
    </nav>
    <div class="container-fluid">
"@

$htmlFooter = @"
        <!-- Footer with copyright and timestamp -->
        <footer class="mt-5 py-3 text-center text-muted">
            <p>© $(Get-Date -Format yyyy) $companyName. All rights reserved.</p>
            <p class="small">Last updated: <span id="last-updated">$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</span></p>
        </footer>
    </div>
    
    <!-- JavaScript libraries and functions -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/dataTables.bootstrap5.min.js"></script>
    <script>
        // Live clock update function
        function updateClock() {
            const now = new Date();
            document.getElementById('live-clock').textContent = now.toLocaleTimeString();
            setTimeout(updateClock, 1000);
        }
        updateClock();
        
        // Auto-refresh every 5 minutes (300,000 ms)
        setTimeout(function() {
            window.location.reload();
        }, 300000);
        
        // Initialize DataTables for tables with class 'datatable'
        `$(document).ready(function() {
            `$('.datatable').DataTable({
                responsive: true,
                pageLength: 10,
                lengthMenu: [5, 10, 25, 50, 100]
            });
        });
        
        // Export button functionality
        function exportData(format) {
            const searchParams = new URLSearchParams(window.location.search);
            const assetNumber = searchParams.get('AssetNumber') || '';
            const searchTerm = searchParams.get('search') || '';
            const page = searchParams.get('page') || 1;
            
            if (assetNumber) {
                window.location.href = `/export?AssetNumber=${assetNumber}&format=${format}`;
            } else {
                window.location.href = `/export?search=${searchTerm}&page=${page}&format=${format}`;
            }
        }
    </script>
</body>
</html>
"@
#endregion

#region DASHBOARD FUNCTIONS

<#
.SYNOPSIS
    Retrieves paginated system records from the database with sorting
#>
function Get-SystemRecords {
    param(
        [string]$searchTerm = "",
        [int]$page = 1,
        [int]$pageSize = $defaultRowCount,
        [string]$sortColumn = "ScanDate",
        [string]$sortDirection = "DESC"
    )
    
    try {
        $offset = ($page - 1) * $pageSize
        
        # Create connection and command objects
        $conn = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $conn.Open()
        
        # Validate sort column to prevent SQL injection
        $validColumns = @("AssetNumber", "HostName", "OS", "TotalRAMGB", "ScanDate")
        if ($validColumns -notcontains $sortColumn) {
            $sortColumn = "ScanDate"
        }
        
        # Validate sort direction
        $sortDirection = if ($sortDirection -eq "ASC") { "ASC" } else { "DESC" }
        
        # Build base query with explicit table references
        $query = @"
        SELECT 
            s.AssetNumber, 
            s.HostName, 
            s.OS, 
            h.TotalRAMGB,
            s.ScanDate,
            COUNT(*) OVER() AS TotalCount
        FROM Systems s
        LEFT JOIN Hardware h ON s.AssetNumber = h.AssetNumber
"@

        # Add WHERE clause if search term provided
        $whereClause = ""
        $params = @{}
        
        if (-not [string]::IsNullOrWhiteSpace($searchTerm)) {
            $whereClause = "WHERE s.AssetNumber LIKE @searchTerm OR s.HostName LIKE @searchTerm OR s.OS LIKE @searchTerm"
            $params["@searchTerm"] = "%$searchTerm%"
        }

        # Add ORDER BY and paging
        $query += @"
        $whereClause
        ORDER BY $sortColumn $sortDirection
        OFFSET $offset ROWS
        FETCH NEXT $pageSize ROWS ONLY
"@

        # Execute query with parameters
        $cmd = New-Object System.Data.SqlClient.SqlCommand($query, $conn)
        foreach ($key in $params.Keys) {
            $cmd.Parameters.AddWithValue($key, $params[$key]) | Out-Null
        }

        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dataset = New-Object System.Data.DataSet
        $adapter.Fill($dataset) | Out-Null
        
        $results = $dataset.Tables[0]
        $conn.Close()

        return @{
            Records = $results | Select-Object AssetNumber, HostName, OS, TotalRAMGB, ScanDate
            TotalCount = if ($results.Rows.Count -gt 0) { $results.Rows[0].TotalCount } else { 0 }
        }
    }
    catch {
        Write-Error "Database query failed: $_"
        return @{ Records = @(); TotalCount = 0 }
    }
}

<#
.SYNOPSIS
    Retrieves detailed information about a specific system
#>
function Show-SystemDetails {
    param($AssetNumber)
    
    try {
        # Create connection
        $conn = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $conn.Open()

        # Query for basic system details
        $systemQuery = @"
        SELECT 
            s.AssetNumber,
            s.HostName,
            s.SerialNumber,
            s.OS,
            s.Version,
            s.Architecture,
            s.Manufacturer,
            s.Model,
            s.BootTime,
            s.BIOSVersion,
            s.ScanDate,
            s.PSVersion,
            h.CPUName,
            h.CPUCores,
            h.CPUThreads,
            h.CPUClockSpeed,
            h.TotalRAMGB,
            h.PageFileGB,
            h.MemorySticks,
            h.GPUName,
            h.GPUAdapterRAMGB,
            h.GPUDriverVersion,
            n.IPAddress,
            n.MacAddress,
            n.SubnetMask,
            n.DefaultGateway,
            n.DNSServers
        FROM Systems s
        LEFT JOIN Hardware h ON s.AssetNumber = h.AssetNumber
        LEFT JOIN Network n ON s.AssetNumber = n.AssetNumber
        WHERE s.AssetNumber = @AssetNumber
"@
        $cmd = New-Object System.Data.SqlClient.SqlCommand($systemQuery, $conn)
        $cmd.Parameters.AddWithValue("@AssetNumber", $AssetNumber) | Out-Null
        
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dataset = New-Object System.Data.DataSet
        $adapter.Fill($dataset) | Out-Null
        $details = $dataset.Tables[0].Rows[0]

        # Query for installed applications
        $appsQuery = @"
        SELECT AppName, AppVersion, Publisher, InstallDate
        FROM Software
        WHERE AssetNumber = @AssetNumber AND IsApplication = 1
        ORDER BY AppName
"@
        $cmd = New-Object System.Data.SqlClient.SqlCommand($appsQuery, $conn)
        $cmd.Parameters.AddWithValue("@AssetNumber", $AssetNumber) | Out-Null
        
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dataset = New-Object System.Data.DataSet
        $adapter.Fill($dataset) | Out-Null
        $installedApps = $dataset.Tables[0]

        # Query for hotfixes
        $hotfixQuery = @"
        SELECT HotFixID, HotFixDescription, HotFixInstalledDate
        FROM Software
        WHERE AssetNumber = @AssetNumber AND IsApplication = 0
        ORDER BY HotFixInstalledDate DESC
"@
        $cmd = New-Object System.Data.SqlClient.SqlCommand($hotfixQuery, $conn)
        $cmd.Parameters.AddWithValue("@AssetNumber", $AssetNumber) | Out-Null
        
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dataset = New-Object System.Data.DataSet
        $adapter.Fill($dataset) | Out-Null
        $hotfixes = $dataset.Tables[0]

        # Query for disks
        $disksQuery = @"
        SELECT d.DeviceID, d.VolumeName, d.SizeGB, d.FreeGB, d.Type
        FROM Disks d
        JOIN Hardware h ON d.HardwareID = h.HardwareID
        WHERE h.AssetNumber = @AssetNumber
        ORDER BY d.DeviceID
"@
        $cmd = New-Object System.Data.SqlClient.SqlCommand($disksQuery, $conn)
        $cmd.Parameters.AddWithValue("@AssetNumber", $AssetNumber) | Out-Null
        
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dataset = New-Object System.Data.DataSet
        $adapter.Fill($dataset) | Out-Null
        $disks = $dataset.Tables[0]

        $conn.Close()

        # Generate HTML for installed applications list
        $appsHtml = ""
        if ($installedApps.Rows.Count -gt 0) {
            $appsHtml += "<table class='table table-sm datatable'>"
            $appsHtml += "<thead><tr><th>Application</th><th>Version</th><th>Publisher</th><th>Installed</th></tr></thead>"
            $appsHtml += "<tbody>"
            foreach ($row in $installedApps.Rows) {
                $displayName = Encode-HTML $row["AppName"]
                $version = if ([string]::IsNullOrWhiteSpace($row["AppVersion"])) { 
                    "Unknown" 
                } else { 
                    Encode-HTML $row["AppVersion"]
                }
                $publisher = Encode-HTML $row["Publisher"]
                $installDate = if ([DBNull]::Value.Equals($row["InstallDate"])) { 
                    "Unknown" 
                } else { 
                    ([DateTime]$row["InstallDate"]).ToString("yyyy-MM-dd")
                }
                
                $appsHtml += @"
<tr>
    <td>$displayName</td>
    <td>$version</td>
    <td>$publisher</td>
    <td>$installDate</td>
</tr>
"@
            }
            $appsHtml += "</tbody></table>"
        } else {
            $appsHtml = "<div class='alert alert-info'>No application data available</div>"
        }

        # Generate HTML for hotfixes
        $hotfixHtml = ""
        if ($hotfixes.Rows.Count -gt 0) {
            $hotfixHtml += "<table class='table table-sm datatable'>"
            $hotfixHtml += "<thead><tr><th>Hotfix ID</th><th>Description</th><th>Installed On</th></tr></thead>"
            $hotfixHtml += "<tbody>"
            foreach ($row in $hotfixes.Rows) {
                $hotfixId = Encode-HTML $row["HotFixID"]
                $description = Encode-HTML $row["HotFixDescription"]
                $installedDate = if ([DBNull]::Value.Equals($row["HotFixInstalledDate"])) { 
                    "Unknown" 
                } else { 
                    ([DateTime]$row["HotFixInstalledDate"]).ToString("yyyy-MM-dd")
                }
                
                $hotfixHtml += @"
<tr>
    <td>$hotfixId</td>
    <td>$description</td>
    <td>$installedDate</td>
</tr>
"@
            }
            $hotfixHtml += "</tbody></table>"
        } else {
            $hotfixHtml = "<div class='alert alert-info'>No hotfix data available</div>"
        }

        # Generate HTML for disks
        $disksHtml = ""
        if ($disks.Rows.Count -gt 0) {
            $disksHtml += "<table class='table table-sm datatable'>"
            $disksHtml += "<thead><tr><th>Device</th><th>Volume</th><th>Type</th><th>Size</th><th>Free</th><th>Usage</th></tr></thead>"
            $disksHtml += "<tbody>"
            foreach ($row in $disks.Rows) {
                $deviceId = Encode-HTML $row["DeviceID"]
                $volumeName = Encode-HTML $row["VolumeName"]
                $type = Encode-HTML $row["Type"]
                $sizeGB = [math]::Round($row["SizeGB"], 2)
                $freeGB = [math]::Round($row["FreeGB"], 2)
                $usedGB = $sizeGB - $freeGB
                $usedPercent = if ($sizeGB -gt 0) { [math]::Round(($usedGB / $sizeGB) * 100) } else { 0 }
                
                $usageColor = if ($usedPercent -gt 90) { "bg-danger" } elseif ($usedPercent -gt 70) { "bg-warning" } else { "bg-success" }
                
                $disksHtml += @"
<tr>
    <td>$deviceId</td>
    <td>$volumeName</td>
    <td>$type</td>
    <td>$sizeGB GB</td>
    <td>$freeGB GB</td>
    <td>
        <div class="disk-usage bg-light">
            <div class="disk-usage-bar $usageColor" style="width: $usedPercent%"></div>
        </div>
        <small>$usedPercent% used</small>
    </td>
</tr>
"@
            }
            $disksHtml += "</tbody></table>"
        } else {
            $disksHtml = "<div class='alert alert-info'>No disk data available</div>"
        }

        # Build the complete details HTML with tabs
        $html = @"
        <div class="card mb-4">
            <div class="card-header bg-primary text-white">
                <div class="d-flex justify-content-between align-items-center">
                    <h5 class="mb-0">System Details: $(Encode-HTML $details["HostName"])</h5>
                    <div>
                        <button class="btn btn-sm btn-light export-btn" onclick="exportData('csv')">
                            <i class="fas fa-file-csv me-1"></i>Export CSV
                        </button>
                        <button class="btn btn-sm btn-light export-btn" onclick="exportData('json')">
                            <i class="fas fa-file-code me-1"></i>Export JSON
                        </button>
                    </div>
                </div>
            </div>
            <div class="card-body">
                <!-- System summary badges -->
                <div class="mb-4">
                    <span class="badge bg-secondary me-2 badge-custom">
                        <i class="fas fa-hashtag me-1"></i>$(Encode-HTML $details["AssetNumber"])
                    </span>
                    <span class="badge bg-info me-2 badge-custom">
                        <i class="fas fa-microchip me-1"></i>$(Encode-HTML $details["CPUName"])
                    </span>
                    <span class="badge bg-info me-2 badge-custom">
                        <i class="fas fa-memory me-1"></i>$([math]::Round($details["TotalRAMGB"], 2)) GB RAM
                    </span>
                    <span class="badge bg-info me-2 badge-custom">
                        <i class="fas fa-network-wired me-1"></i>$(Encode-HTML $details["IPAddress"])
                    </span>
                    <span class="badge bg-info me-2 badge-custom">
                        <i class="fas fa-clock me-1"></i>Last scanned: $(Encode-HTML $details["ScanDate"])
                    </span>
                </div>
                
                <!-- Tab navigation -->
                <ul class="nav nav-tabs" id="systemDetailsTabs" role="tablist">
                    <li class="nav-item" role="presentation">
                        <button class="nav-link active" id="overview-tab" data-bs-toggle="tab" data-bs-target="#overview" type="button" role="tab">
                            <i class="fas fa-info-circle me-1"></i>Overview
                        </button>
                    </li>
                    <li class="nav-item" role="presentation">
                        <button class="nav-link" id="hardware-tab" data-bs-toggle="tab" data-bs-target="#hardware" type="button" role="tab">
                            <i class="fas fa-microchip me-1"></i>Hardware
                        </button>
                    </li>
                    <li class="nav-item" role="presentation">
                        <button class="nav-link" id="software-tab" data-bs-toggle="tab" data-bs-target="#software" type="button" role="tab">
                            <i class="fas fa-windows me-1"></i>Software
                        </button>
                    </li>
                    <li class="nav-item" role="presentation">
                        <button class="nav-link" id="storage-tab" data-bs-toggle="tab" data-bs-target="#storage" type="button" role="tab">
                            <i class="fas fa-hdd me-1"></i>Storage
                        </button>
                    </li>
                    <li class="nav-item" role="presentation">
                        <button class="nav-link" id="network-tab" data-bs-toggle="tab" data-bs-target="#network" type="button" role="tab">
                            <i class="fas fa-network-wired me-1"></i>Network
                        </button>
                    </li>
                </ul>
                
                <!-- Tab content -->
                <div class="tab-content" id="systemDetailsTabsContent">
                    <!-- Overview tab -->
                    <div class="tab-pane fade show active" id="overview" role="tabpanel">
                        <div class="row mt-3">
                            <div class="col-md-6">
                                <h6><i class="fas fa-info-circle me-2"></i>System Information</h6>
                                <table class="table table-sm">
                                    $(Render-DetailRow "Host Name" $details["HostName"])
                                    $(Render-DetailRow "Asset Number" $details["AssetNumber"])
                                    $(Render-DetailRow "Serial Number" $details["SerialNumber"])
                                    $(Render-DetailRow "Manufacturer" $details["Manufacturer"])
                                    $(Render-DetailRow "Model" $details["Model"])
                                    $(Render-DetailRow "Last Scan" $details["ScanDate"])
                                </table>
                            </div>
                            <div class="col-md-6">
                                <h6><i class="fas fa-windows me-2"></i>Operating System</h6>
                                <table class="table table-sm">
                                    $(Render-DetailRow "OS" $details["OS"])
                                    $(Render-DetailRow "Version" $details["Version"])
                                    $(Render-DetailRow "Architecture" $details["Architecture"])
                                    $(Render-DetailRow "Build" $details["Build"])
                                    $(Render-DetailRow "PowerShell" $details["PSVersion"])
                                    $(Render-DetailRow "Boot Time" $details["BootTime"])
                                </table>
                            </div>
                        </div>
                    </div>
                    
                    <!-- Hardware tab -->
                    <div class="tab-pane fade" id="hardware" role="tabpanel">
                        <div class="row mt-3">
                            <div class="col-md-6">
                                <h6><i class="fas fa-microchip me-2"></i>Processor</h6>
                                <table class="table table-sm">
                                    $(Render-DetailRow "CPU Name" $details["CPUName"])
                                    $(Render-DetailRow "Cores" $details["CPUCores"])
                                    $(Render-DetailRow "Threads" $details["CPUThreads"])
                                    $(Render-DetailRow "Clock Speed" $details["CPUClockSpeed"])
                                </table>
                                
                                <h6><i class="fas fa-memory me-2"></i>Memory</h6>
                                <table class="table table-sm">
                                    $(Render-DetailRow "Total RAM" "$([math]::Round($details["TotalRAMGB"], 2)) GB")
                                    $(Render-DetailRow "Page File" "$([math]::Round($details["PageFileGB"], 2)) GB")
                                    $(Render-DetailRow "Memory Sticks" $details["MemorySticks"])
                                </table>
                            </div>
                            <div class="col-md-6">
                                <h6><i class="fas fa-desktop me-2"></i>Graphics</h6>
                                <table class="table table-sm">
                                    $(Render-DetailRow "GPU Name" $details["GPUName"])
                                    $(Render-DetailRow "GPU RAM" "$([math]::Round($details["GPUAdapterRAMGB"], 2)) GB")
                                    $(Render-DetailRow "Driver Version" $details["GPUDriverVersion"])
                                </table>
                                
                                <h6><i class="fas fa-barcode me-2"></i>BIOS</h6>
                                <table class="table table-sm">
                                    $(Render-DetailRow "Version" $details["BIOSVersion"])
                                    $(Render-DetailRow "Serial Number" $details["SerialNumber"])
                                </table>
                            </div>
                        </div>
                    </div>
                    
                    <!-- Software tab -->
                    <div class="tab-pane fade" id="software" role="tabpanel">
                        <ul class="nav nav-pills mb-3" id="software-tabs" role="tablist">
                            <li class="nav-item" role="presentation">
                                <button class="nav-link active" id="applications-tab" data-bs-toggle="pill" data-bs-target="#applications" type="button">
                                    <i class="fas fa-box me-1"></i>Applications ($($installedApps.Rows.Count))
                                </button>
                            </li>
                            <li class="nav-item" role="presentation">
                                <button class="nav-link" id="hotfixes-tab" data-bs-toggle="pill" data-bs-target="#hotfixes" type="button">
                                    <i class="fas fa-patch me-1"></i>Hotfixes ($($hotfixes.Rows.Count))
                                </button>
                            </li>
                        </ul>
                        
                        <div class="tab-content">
                            <div class="tab-pane fade show active" id="applications" role="tabpanel">
                                $appsHtml
                            </div>
                            <div class="tab-pane fade" id="hotfixes" role="tabpanel">
                                $hotfixHtml
                            </div>
                        </div>
                    </div>
                    
                    <!-- Storage tab -->
                    <div class="tab-pane fade" id="storage" role="tabpanel">
                        $disksHtml
                    </div>
                    
                    <!-- Network tab -->
                    <div class="tab-pane fade" id="network" role="tabpanel">
                        <div class="row mt-3">
                            <div class="col-md-6">
                                <h6><i class="fas fa-network-wired me-2"></i>Network Configuration</h6>
                                <table class="table table-sm">
                                    $(Render-DetailRow "IP Address" $details["IPAddress"])
                                    $(Render-DetailRow "MAC Address" $details["MacAddress"])
                                    $(Render-DetailRow "Subnet Mask" $details["SubnetMask"])
                                    $(Render-DetailRow "Default Gateway" $details["DefaultGateway"])
                                    $(Render-DetailRow "DNS Servers" $details["DNSServers"])
                                </table>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
"@
        return $html
    }
    catch {
        return "<div class='alert alert-danger'>Error loading system details: $(Encode-HTML $_.Exception.Message)</div>"
    }
}

<#
.SYNOPSIS
    Helper function to render a table row for system details
#>
function Render-DetailRow {
    param($label, $value)
    
    if ([DBNull]::Value.Equals($value) -or $value -eq $null) {
        $value = "N/A"
    }
    elseif ($value -is [DateTime]) {
        $value = $value.ToString("yyyy-MM-dd HH:mm:ss")
    }
    return @"
<tr>
    <th>$(Encode-HTML $label)</th>
    <td>$(Encode-HTML $value)</td>
</tr>
"@
}

<#
.SYNOPSIS
    Generates HTML for pagination controls
#>
function Get-PaginationHtml {
    param(
        [int]$currentPage,
        [int]$totalItems,
        [int]$pageSize = $defaultRowCount,
        [string]$searchTerm = ""
    )
    
    $totalPages = [math]::Ceiling($totalItems / $pageSize)
    if ($totalPages -le 1) { return "" }

    $html = "<nav aria-label='Page navigation'><ul class='pagination justify-content-center'>"
    
    # Previous button
    $prevDisabled = if ($currentPage -le 1) { "disabled" } else { "" }
    $prevPage = [math]::Max(1, $currentPage - 1)
    $html += "<li class='page-item $prevDisabled'>" +
             "<a class='page-link' href='/?page=$prevPage&search=$(Encode-HTML $searchTerm)'>" +
             "<i class='fas fa-chevron-left'></i></a></li>"
    
    # Calculate page range to show (5 pages around current)
    $startPage = [math]::Max(1, $currentPage - 2)
    $endPage = [math]::Min($totalPages, $currentPage + 2)
    
    # Show first page + ellipsis if needed
    if ($startPage -gt 1) {
        $html += "<li class='page-item'>" +
                 "<a class='page-link' href='/?page=1&search=$(Encode-HTML $searchTerm)'>1</a></li>"
        if ($startPage -gt 2) {
            $html += "<li class='page-item disabled'><span class='page-link'>...</span></li>"
        }
    }
    
    # Show calculated page range
    for ($i = $startPage; $i -le $endPage; $i++) {
        $active = if ($i -eq $currentPage) { "active" } else { "" }
        $html += "<li class='page-item $active'>" +
                 "<a class='page-link' href='/?page=$i&search=$(Encode-HTML $searchTerm)'>$i</a></li>"
    }
    
    # Show ellipsis + last page if needed
    if ($endPage -lt $totalPages) {
        if ($endPage -lt $totalPages - 1) {
            $html += "<li class='page-item disabled'><span class='page-link'>...</span></li>"
        }
        $html += "<li class='page-item'>" +
                 "<a class='page-link' href='/?page=$totalPages&search=$(Encode-HTML $searchTerm)'>$totalPages</a></li>"
    }
    
    # Next button
    $nextDisabled = if ($currentPage -ge $totalPages) { "disabled" } else { "" }
    $nextPage = [math]::Min($totalPages, $currentPage + 1)
    $html += "<li class='page-item $nextDisabled'>" +
             "<a class='page-link' href='/?page=$nextPage&search=$(Encode-HTML $searchTerm)'>" +
             "<i class='fas fa-chevron-right'></i></a></li>"
    
    $html += "</ul></nav>"
    return $html
}

<#
.SYNOPSIS
    Exports system data in specified format
#>
function Export-SystemData {
    param(
        [string]$AssetNumber = "",
        [string]$searchTerm = "",
        [int]$page = 1,
        [string]$format = "csv"
    )
    
    try {
        $conn = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $conn.Open()

        if (-not [string]::IsNullOrWhiteSpace($AssetNumber)) {
            # Export single system details
            $query = @"
            SELECT 
                s.*, 
                h.*,
                n.*
            FROM Systems s
            LEFT JOIN Hardware h ON s.AssetNumber = h.AssetNumber
            LEFT JOIN Network n ON s.AssetNumber = n.AssetNumber
            WHERE s.AssetNumber = @AssetNumber
"@
            $cmd = New-Object System.Data.SqlClient.SqlCommand($query, $conn)
            $cmd.Parameters.AddWithValue("@AssetNumber", $AssetNumber) | Out-Null
            
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
            $dataset = New-Object System.Data.DataSet
            $adapter.Fill($dataset) | Out-Null
            $data = $dataset.Tables[0]
        }
        else {
            # Export multiple systems
            $offset = ($page - 1) * $defaultRowCount
            
            $query = @"
            SELECT 
                s.AssetNumber, 
                s.HostName, 
                s.OS, 
                s.Version,
                s.Architecture,
                s.Manufacturer,
                s.Model,
                h.CPUName,
                h.CPUCores,
                h.TotalRAMGB,
                s.ScanDate
            FROM Systems s
            LEFT JOIN Hardware h ON s.AssetNumber = h.AssetNumber
"@
            $whereClause = ""
            $params = @{}
            
            if (-not [string]::IsNullOrWhiteSpace($searchTerm)) {
                $whereClause = "WHERE s.AssetNumber LIKE @searchTerm OR s.HostName LIKE @searchTerm OR s.OS LIKE @searchTerm"
                $params["@searchTerm"] = "%$searchTerm%"
            }

            $query += @"
            $whereClause
            ORDER BY s.ScanDate DESC
            OFFSET $offset ROWS
            FETCH NEXT $defaultRowCount ROWS ONLY
"@

            $cmd = New-Object System.Data.SqlClient.SqlCommand($query, $conn)
            foreach ($key in $params.Keys) {
                $cmd.Parameters.AddWithValue($key, $params[$key]) | Out-Null
            }

            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
            $dataset = New-Object System.Data.DataSet
            $adapter.Fill($dataset) | Out-Null
            $data = $dataset.Tables[0]
        }

        $conn.Close()

        # Convert to requested format
        switch ($format.ToLower()) {
            "csv" {
                $output = New-Object System.Text.StringBuilder
                
                # Add header row
                $headerRow = $data.Columns | ForEach-Object { $_.ColumnName }
                $output.AppendLine(($headerRow -join ",")) | Out-Null
                
                # Add data rows
                foreach ($row in $data.Rows) {
                    $dataRow = $data.Columns | ForEach-Object { 
                        $value = $row[$_]
                        if ($value -is [DateTime]) {
                            $value.ToString("yyyy-MM-dd HH:mm:ss")
                        }
                        elseif ($value -eq [DBNull]::Value) {
                            ""
                        }
                        else {
                            '"' + $value.ToString().Replace('"', '""') + '"'
                        }
                    }
                    $output.AppendLine(($dataRow -join ",")) | Out-Null
                }
                
                return $output.ToString()
            }
            "json" {
                $result = @()
                foreach ($row in $data.Rows) {
                    $item = @{}
                    foreach ($column in $data.Columns) {
                        $value = $row[$column]
                        if ($value -is [DateTime]) {
                            $item[$column.ColumnName] = $value.ToString("yyyy-MM-ddTHH:mm:ss")
                        }
                        elseif ($value -eq [DBNull]::Value) {
                            $item[$column.ColumnName] = $null
                        }
                        else {
                            $item[$column.ColumnName] = $value
                        }
                    }
                    $result += $item
                }
                
                if ($result.Count -eq 1) {
                    return $result[0] | ConvertTo-Json -Depth 5
                }
                else {
                    return $result | ConvertTo-Json -Depth 5
                }
            }
            default {
                throw "Unsupported export format: $format"
            }
        }
    }
    catch {
        Write-Error "Export failed: $_"
        throw
    }
}

<#
.SYNOPSIS
    Logs performance metrics for dashboard operations
#>
function Log-Performance {
    param(
        [string]$operation,
        [double]$durationMs,
        [string]$details = ""
    )
    
    if (-not $enablePerformanceLogging) { return }
    
    $logEntry = @{
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Operation = $operation
        DurationMs = $durationMs
        Details = $details
    } | ConvertTo-Json -Compress
    
    try {
        Add-Content -Path $performanceLogPath -Value $logEntry -ErrorAction SilentlyContinue
    }
    catch {
        Write-Warning "Failed to write to performance log: $_"
    }
}

<#
.SYNOPSIS
    Renders the complete dashboard HTML
#>
function Render-Dashboard {
    param(
        [string]$searchTerm = "",
        [int]$page = 1,
        [string]$AssetNumber = ""
    )
    
    $startTime = Get-Date
    
    # Build search form
    $searchForm = @"
        <div class="card search-box mb-4">
            <div class="card-body">
                <form method="GET" action="/" class="row g-3">
                    <div class="col-md-8">
                        <div class="input-group">
                            <span class="input-group-text"><i class="fas fa-search"></i></span>
                            <input type="text" 
                                   class="form-control" 
                                   name="search" 
                                   placeholder="Search by asset number, host name or OS..." 
                                   value="$(Encode-HTML $searchTerm)">
                        </div>
                    </div>
                    <div class="col-md-4">
                        <button type="submit" class="btn btn-primary me-2">
                            <i class="fas fa-search me-1"></i>Search
                        </button>
                        <a href="/" class="btn btn-outline-secondary">
                            <i class="fas fa-times me-1"></i>Clear
                        </a>
                    </div>
                </form>
            </div>
        </div>
"@

    # Build system details section if an asset number was requested
    if (-not [string]::IsNullOrWhiteSpace($AssetNumber)) {
        $detailsSection = Show-SystemDetails -AssetNumber $AssetNumber
        $backButton = @"
        <div class="mb-3">
            <a href="/?page=$page&search=$(Encode-HTML $searchTerm)" class="btn btn-outline-secondary">
                <i class="fas fa-arrow-left me-1"></i> Back to System List
            </a>
        </div>
"@
        # Combine all components for details view
        $output = $htmlHeader + $searchForm + $backButton + $detailsSection + $htmlFooter
        
        $duration = ((Get-Date) - $startTime).TotalMilliseconds
        Log-Performance -operation "RenderDetails" -durationMs $duration -details "AssetNumber=$AssetNumber"
        
        return $output
    }

    # If no asset number, show the system list view
    $systemData = Get-SystemRecords -searchTerm $searchTerm -page $page
    
    # Results count indicator
    $resultsCount = @"
    <div class="row mb-3">
        <div class="col">
            <p class="results-count">
                Showing $([math]::Min(($page - 1) * $defaultRowCount + 1, $systemData.TotalCount)) - 
                $([math]::Min($page * $defaultRowCount, $systemData.TotalCount)) of 
                $($systemData.TotalCount) systems
            </p>
        </div>
        <div class="col-auto">
            <div class="btn-group">
                <button class="btn btn-sm btn-outline-secondary" onclick="exportData('csv')">
                    <i class="fas fa-file-csv me-1"></i>Export CSV
                </button>
                <button class="btn btn-sm btn-outline-secondary" onclick="exportData('json')">
                    <i class="fas fa-file-code me-1"></i>Export JSON
                </button>
            </div>
        </div>
    </div>
"@

    # Build system cards for the main listing
    $systemCards = ""
    foreach ($system in $systemData.Records) {
        $systemCards += @"
        <div class="col-md-6 col-lg-4 mb-4">
            <div class="card system-card h-100">
                <div class="card-header">
                    <h5 class="card-title mb-0">
                        <i class="fas fa-desktop me-2"></i>$(Encode-HTML $system.HostName)
                    </h5>
                </div>
                <div class="card-body">
                    <div class="d-flex justify-content-between mb-2">
                        <span class="badge bg-primary">$(Encode-HTML $system.AssetNumber)</span>
                        <span class="badge bg-info">$([math]::Round($system.TotalRAMGB, 2)) GB RAM</span>
                    </div>
                    <p class="card-text">
                        <i class="fas fa-windows me-2"></i>$(Encode-HTML $system.OS)<br>
                        <small class="text-muted">Last scanned: $(Encode-HTML $system.ScanDate)</small>
                    </p>
                </div>
                <div class="card-footer bg-transparent">
                    <a href="/?AssetNumber=$(Encode-HTML $system.AssetNumber)&search=$(Encode-HTML $searchTerm)&page=$page" 
                       class="btn btn-sm btn-outline-primary">
                        <i class="fas fa-info-circle me-1"></i>Details
                    </a>
                </div>
            </div>
        </div>
"@
    }

    # Generate pagination controls (top and bottom)
    $paginationTop = Get-PaginationHtml -currentPage $page -totalItems $systemData.TotalCount -searchTerm $searchTerm
    $paginationBottom = $paginationTop

    # Combine all components for list view
    $output = $htmlHeader + $searchForm + $resultsCount + $paginationTop + @"
        <div class="row">
            $systemCards
        </div>
        $paginationBottom
"@ + $htmlFooter

    $duration = ((Get-Date) - $startTime).TotalMilliseconds
    Log-Performance -operation "RenderList" -durationMs $duration -details "Page=$page, SearchTerm=$searchTerm"
    
    return $output
}

<#
.SYNOPSIS
    Encodes text for safe HTML output
#>
function Encode-HTML {
    param([string]$text)
    if ([string]::IsNullOrEmpty($text)) { return "" }
    return [System.Net.WebUtility]::HtmlEncode($text)
}
#endregion

#region WEB SERVER
try {
    # Initialize HTTP listener
    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://localhost:$port/")
    $listener.Start()
    Write-Host "Dashboard running at http://localhost:$port" -ForegroundColor Green

    # Main request handling loop
    while ($listener.IsListening) {
        try {
            $context = $listener.GetContext()
            $request = $context.Request
            $response = $context.Response

            try {
                # Parse query parameters from URL
                $queryParams = @{}
                if ($request.Url.Query) {
                    $request.Url.Query.TrimStart('?').Split('&') | ForEach-Object {
                        $key, $value = $_.Split('=', 2)
                        $queryParams[$key] = if ($value) { [System.Uri]::UnescapeDataString($value) } else { "" }
                    }
                }

                # Handle export requests
                if ($request.Url.LocalPath -eq "/export") {
                    $startTime = Get-Date
                    
                    # Get parameters with defaults
                    $searchTerm = $queryParams["search"] ?? ""
                    $page = [int]($queryParams["page"] ?? 1)
                    $AssetNumber = $queryParams["AssetNumber"] ?? ""
                    $format = $queryParams["format"] ?? "csv"
                    
                    $exportData = Export-SystemData -searchTerm $searchTerm -page $page -AssetNumber $AssetNumber -format $format
                    
                    # Set appropriate content type and headers
                    if ($format -eq "csv") {
                        $response.ContentType = "text/csv"
                        $filename = if ($AssetNumber) { "SystemDetails_$AssetNumber.csv" } else { "Systems_$page.csv" }
                        $response.AddHeader("Content-Disposition", "attachment; filename=$filename")
                    }
                    else {
                        $response.ContentType = "application/json"
                        $filename = if ($AssetNumber) { "SystemDetails_$AssetNumber.json" } else { "Systems_$page.json" }
                        $response.AddHeader("Content-Disposition", "attachment; filename=$filename")
                    }
                    
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($exportData)
                    $response.ContentLength64 = $buffer.Length
                    $response.OutputStream.Write($buffer, 0, $buffer.Length)
                    $response.OutputStream.Close()
                    
                    $duration = ((Get-Date) - $startTime).TotalMilliseconds
                    Log-Performance -operation "ExportData" -durationMs $duration -details "Format=$format, AssetNumber=$AssetNumber"
                    
                    continue
                }

                # Get parameters with defaults for regular requests
                $searchTerm = $queryParams["search"] ?? ""
                $page = [int]($queryParams["page"] ?? 1)
                $AssetNumber = $queryParams["AssetNumber"] ?? ""

                # Generate and send HTML response
                $html = Render-Dashboard -searchTerm $searchTerm -page $page -AssetNumber $AssetNumber
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)

                $response.ContentLength64 = $buffer.Length
                $response.ContentType = "text/html; charset=utf-8"
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }
            catch {
                # Error page for request processing failures
                $errorHtml = @"
<!DOCTYPE html>
<html>
<head>
    <title>Error</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body>
    <div class="container mt-5">
        <div class="alert alert-danger">
            <h4 class="alert-heading">Error Processing Request</h4>
            <p>An unexpected error occurred while processing your request.</p>
            <hr>
            <p class="mb-0">$(Encode-HTML $_.Exception.Message)</p>
        </div>
        <a href="/" class="btn btn-primary">Return to Dashboard</a>
    </div>
</body>
</html>
"@
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($errorHtml)
                $response.ContentLength64 = $buffer.Length
                $response.ContentType = "text/html; charset=utf-8"
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
                
                Write-Host "Error processing request: $_" -ForegroundColor Red
            }
        }
        catch {
            Write-Host "Error in request loop: $_" -ForegroundColor Red
        }
    }
}
catch {
    Write-Host "Failed to start dashboard: $_" -ForegroundColor Red
}
finally {
    # Clean up listener when done
    if ($listener.IsListening) {
        $listener.Stop()
    }
    
    # Log shutdown
    Write-Host "Dashboard server stopped" -ForegroundColor Yellow
}
#endregion