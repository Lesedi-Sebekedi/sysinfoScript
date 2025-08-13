<#
.SYNOPSIS
    System Inventory Dashboard Web Server
.DESCRIPTION
    Creates a responsive local web dashboard to view system inventory from SQL database
    Features:
    - Paginated system listings
    - Detailed system views
    - Search functionality
    - Automatic refresh
    - Secure HTML rendering
.NOTES
    Author: Your Name
    Version: 1.2
    Last Updated: $(Get-Date -Format "yyyy-MM-dd")
#>

#region CONFIGURATION
# Dashboard settings
$port = 8080  # Web server listening port
$dashboardTitle = "System Inventory Dashboard"
$companyName = "Nowth West Provincial Treasury"
$defaultRowCount = 20  # Number of items per page
$connectionString = "Server=PTLSEBEKEDI;Database=AssetDB;Integrated Security=True;TrustServerCertificate=True"
#endregion

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
    </script>
</body>
</html>
"@
#endregion

#region DASHBOARD FUNCTIONS

<#
.SYNOPSIS
    Retrieves paginated system records from the database
.DESCRIPTION
    Executes SQL query to fetch system records with optional search filtering
    Returns both the records and total count for pagination
.PARAMETER searchTerm
    Optional term to filter systems by asset number or host name
.PARAMETER page
    Page number to retrieve (1-based)
.PARAMETER pageSize
    Number of records per page
#>
function Get-SystemRecords {
    param(
        [string]$searchTerm = "",
        [int]$page = 1,
        [int]$pageSize = $defaultRowCount
    )
    
    try {
        $offset = ($page - 1) * $pageSize
        
        # Build SQL query based on whether we're filtering
        if ([string]::IsNullOrWhiteSpace($searchTerm)) {
            $query = @"
            SELECT 
                s.SystemID,
                s.AssetNumber, 
                s.HostName, 
                sp.OSName, 
                sp.TotalRAMGB,
                s.ScanDate,
                COUNT(*) OVER() AS TotalCount
            FROM Systems s
            JOIN SystemSpecs sp ON s.SystemID = sp.SystemID
            ORDER BY s.ScanDate DESC
            OFFSET $offset ROWS
            FETCH NEXT $pageSize ROWS ONLY
"@
        } else {
            $escapedTerm = $searchTerm -replace "'", "''"  # Escape single quotes for SQL
            $query = @"
            SELECT 
                s.SystemID,
                s.AssetNumber, 
                s.HostName, 
                sp.OSName, 
                sp.TotalRAMGB,
                s.ScanDate,
                COUNT(*) OVER() AS TotalCount
            FROM Systems s
            JOIN SystemSpecs sp ON s.SystemID = sp.SystemID
            WHERE s.AssetNumber LIKE '%$escapedTerm%' 
               OR s.HostName LIKE '%$escapedTerm%'
            ORDER BY s.ScanDate DESC
            OFFSET $offset ROWS
            FETCH NEXT $pageSize ROWS ONLY
"@
        }

        $results = Invoke-Sqlcmd -ConnectionString $connectionString -Query $query
        
        return @{
            Records = $results | Select-Object SystemID, AssetNumber, HostName, OSName, TotalRAMGB, ScanDate
            TotalCount = if ($results) { $results[0].TotalCount } else { 0 }
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
.DESCRIPTION
    Fetches system details and installed applications for display
.PARAMETER systemId
    The ID of the system to retrieve details for
#>
function Show-SystemDetails {
    param($systemId)
    
    try {
        # Query for basic system details
        $query = @"
        SELECT 
            s.AssetNumber,
            s.HostName,
            s.SerialNumber,
            s.ScanDate,
            sp.*
        FROM Systems s
        JOIN SystemSpecs sp ON s.SystemID = sp.SystemID
        WHERE s.SystemID = $systemId
"@
        $details = Invoke-Sqlcmd -ConnectionString $connectionString -Query $query
        
        # Query for installed applications
        $installedAppsQuery = @"
        SELECT AppName, AppVersion
        FROM InstalledApps
        WHERE SystemID = $systemId
        ORDER BY AppName
"@
        $installedApps = Invoke-Sqlcmd -ConnectionString $connectionString -Query $installedAppsQuery

        # Handle case where query returns an array
        if ($details -is [System.Array]) { 
            $details = $details[0] 
        }

        if ($details) {
            # Generate HTML for installed applications list
            $appsHtml = ""
            if ($installedApps -and $installedApps.Count -gt 0) {
                $appsHtml += "<ul class='list-group'>"
                foreach ($app in $installedApps) {
                    $displayName = Encode-HTML $app.AppName
                    $version = if ([string]::IsNullOrWhiteSpace($app.AppVersion)) { 
                        "Unknown" 
                    } else { 
                        Encode-HTML $app.AppVersion 
                    }
                    $appsHtml += @"
<li class='list-group-item d-flex align-items-center'>
    <img src='https://cdn-icons-png.flaticon.com/512/888/888879.png' 
         alt='App' 
         class='app-icon'>
    <div>
        <strong>$displayName</strong><br>
        <small class='text-muted'>Version: $version</small>
    </div>
</li>
"@
                }
                $appsHtml += "</ul>"
            } else {
                $appsHtml = "<em>No application data available</em>"
            }

            # Build the complete details HTML
            $html = @"
            <div class="card mb-4">
                <div class="card-header bg-primary text-white">
                    <h5 class="mb-0">System Details: $(Encode-HTML $details.HostName)</h5>
                </div>
                <div class="card-body">
                    <div class="row">
                        <!-- Basic Information Column -->
                        <div class="col-md-6">
                            <h6>Basic Information</h6>
                            <table class="table table-sm">
                                $(Render-DetailRow "Asset Number" $details.AssetNumber)
                                $(Render-DetailRow "Host Name" $details.HostName)
                                $(Render-DetailRow "Serial Number" $details.SerialNumber)
                                $(Render-DetailRow "Last Scan" $details.ScanDate)
                            </table>
                        </div>
                        
                        <!-- System Specifications Column -->
                        <div class="col-md-6">
                            <h6>System Specifications</h6>
                            <table class="table table-sm">
                                $(Render-DetailRow "OS" $details.OSName)
                                $(Render-DetailRow "Architecture" $details.Architecture)
                                $(Render-DetailRow "CPU Cores" $details.CPUCores)
                                $(Render-DetailRow "Total RAM" "$([math]::Round($details.TotalRAMGB, 2)) GB")
                            </table>
                        </div>
                    </div>
                    
                    <!-- Installed Applications Section -->
                    <div class="mt-3">
                        <h6>Installed Applications</h6>
                        <div class="bg-light p-3 rounded">
                            $appsHtml
                        </div>
                    </div>
                </div>
            </div>
"@
            return $html
        }
        return "<div class='alert alert-warning'>System details not found</div>"
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
.DESCRIPTION
    Creates pagination links with proper active/disables states
    Shows limited page range around current page with ellipsis
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
    
    # Previous button (disabled if on first page)
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
    
    # Next button (disabled if on last page)
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
    Renders the complete dashboard HTML
.DESCRIPTION
    Combines all components (search, details, listings, pagination)
    into the final HTML page
#>
function Render-Dashboard {
    param(
        [string]$searchTerm = "",
        [int]$page = 1,
        [int]$systemId = 0
    )
    
    # Get paginated system records
    $systemData = Get-SystemRecords -searchTerm $searchTerm -page $page
    
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
                                   placeholder="Search by asset number or host name..." 
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

    # Build system details section if a system ID was requested
    $detailsSection = ""
    if ($systemId -gt 0) {
        $detailsSection = Show-SystemDetails -systemId $systemId
    }

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
                        <i class="fas fa-windows me-2"></i>$(Encode-HTML $system.OSName)<br>
                        <small class="text-muted">Last scanned: $(Encode-HTML $system.ScanDate)</small>
                    </p>
                </div>
                <div class="card-footer bg-transparent">
                    <a href="/?id=$($system.SystemID)&search=$(Encode-HTML $searchTerm)&page=$page" 
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
    </div>
"@

    # Combine all components into final HTML
    $output = $htmlHeader + $searchForm + $detailsSection + $resultsCount + $paginationTop + @"
        <div class="row">
            $systemCards
        </div>
        $paginationBottom
"@ + $htmlFooter

    return $output
}

<#
.SYNOPSIS
    Encodes text for safe HTML output
.DESCRIPTION
    Prevents XSS by encoding special characters
.PARAMETER text
    The text to encode
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
                        $queryParams[$key] = [System.Uri]::UnescapeDataString($value)
                    }
                }

                # Get parameters with defaults
                $searchTerm = $queryParams["search"] ?? ""
                $page = [int]($queryParams["page"] ?? 1)
                $systemId = [int]($queryParams["id"] ?? 0)

                # Generate and send HTML response
                $html = Render-Dashboard -searchTerm $searchTerm -page $page -systemId $systemId
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
}
#endregion