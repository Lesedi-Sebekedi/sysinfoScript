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
    - Connection pooling and caching
    - Rate limiting and health monitoring
.NOTES
    Author: Lesedi Sebekedi
    Version: 3.1
    Last Updated: $(Get-Date -Format "yyyy-MM-dd")
#>

#region CONFIGURATION
# Load or create configuration file
$configPath = Join-Path $PSScriptRoot "DashboardConfig.json"
$defaultConfig = @{
    Port = 8080
    DashboardTitle = "System Inventory Dashboard"
    CompanyName = "North West Provincial Treasury"
    DefaultRowCount = 20
    ConnectionString = "Server=PTLSEBEKEDI;Database=AssetDB;Integrated Security=True;TrustServerCertificate=True"
    EnablePerformanceLogging = $true
    PerformanceLogPath = "$env:TEMP\DashboardPerformance.log"
    MaxConnections = 10
    RateLimitPerMinute = 100
    SessionTimeout = 30
    CacheTimeout = 300
    EnableHTTPS = $false
    CertificatePath = ""
    HealthCheckInterval = 60
    AssetRegisterEnabled = $true
}

# Load or create configuration
if (Test-Path $configPath) {
    try {
        $config = Get-Content $configPath | ConvertFrom-Json
        Write-Host "Configuration loaded from: $configPath" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to load configuration, using defaults: $_"
        $config = $defaultConfig
    }
} else {
    try {
        $defaultConfig | ConvertTo-Json -Depth 3 | Set-Content $configPath
        $config = $defaultConfig
        Write-Host "Default configuration created at: $configPath" -ForegroundColor Yellow
    }
    catch {
        Write-Warning "Failed to create configuration file, using defaults: $_"
        $config = $defaultConfig
    }
}

# Dashboard settings from config
$port = $config.Port
$dashboardTitle = $config.DashboardTitle
$companyName = $config.CompanyName
$defaultRowCount = $config.DefaultRowCount
$connectionString = $config.ConnectionString
$enablePerformanceLogging = $config.EnablePerformanceLogging
$performanceLogPath = $config.PerformanceLogPath
$assetRegisterEnabled = $config.AssetRegisterEnabled
#endregion

#region GLOBAL VARIABLES AND STATE
# Connection pool management
$script:connectionPool = @{
    Connections = [System.Collections.Queue]::new()
    MaxConnections = $config.MaxConnections
    Lock = [System.Threading.ReaderWriterLockSlim]::new()
    ActiveConnections = 0
}

# Rate limiting
$script:requestCounts = @{}
$script:rateLimitWindow = 60  # seconds
$script:maxRequestsPerWindow = $config.RateLimitPerMinute

# Caching system
$script:cache = @{}
$script:cacheTimeout = $config.CacheTimeout

# Health monitoring
$script:startTime = Get-Date
$script:lastHealthCheck = Get-Date
$script:healthStatus = "Healthy"
$script:lastError = $null

# Performance tracking
$script:requestStats = @{
    TotalRequests = 0
    SuccessfulRequests = 0
    FailedRequests = 0
    AverageResponseTime = 0
    LastReset = Get-Date
}

$script:schemaInitLock = [System.Threading.ReaderWriterLockSlim]::new()
#endregion

#region SIGNAL HANDLING
# Handle Ctrl+C gracefully
$script:stopRequested = $false

# Register the interrupt handler
Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    $script:stopRequested = $true
    Write-Host "`nShutting down dashboard server..." -ForegroundColor Yellow
}

# Function to check if stop was requested
function Test-StopRequested {
    return $script:stopRequested
}
#endregion

# Load required SQL module
try {
    Import-Module SqlServer -ErrorAction Stop
}
catch {
    Write-Host "SQL Server module not found. Installing..." -ForegroundColor Yellow
    Install-Module -Name SqlServer -Force -AllowClobber -Scope CurrentUser
    Import-Module SqlServer
}

#region UTILITY FUNCTIONS

<#
.SYNOPSIS
    Gets a database connection from the pool or creates a new one
#>
function Get-DatabaseConnection {
    try {
        $script:connectionPool.Lock.EnterReadLock()
        
        # Try to get connection from pool
        if ($script:connectionPool.Connections.Count -gt 0) {
            $conn = $script:connectionPool.Connections.Dequeue()
            $script:connectionPool.ActiveConnections++
            
            # Test if connection is still valid
            if ($conn.State -eq [System.Data.ConnectionState]::Open) {
                return $conn
            } else {
                $conn.Dispose()
                $script:connectionPool.ActiveConnections--
            }
        }
    }
    finally {
        $script:connectionPool.Lock.ExitReadLock()
    }
    
    # Create new connection if pool is empty or all connections are busy
    try {
        $script:connectionPool.Lock.EnterWriteLock()
        
        if ($script:connectionPool.ActiveConnections -lt $script:connectionPool.MaxConnections) {
            $conn = New-Object System.Data.SqlClient.SqlConnection($connectionString)
            $conn.Open()
            $script:connectionPool.ActiveConnections++
            return $conn
        }
    }
    finally {
        $script:connectionPool.Lock.ExitWriteLock()
    }
    
    # If we can't get a connection, wait and retry
    Start-Sleep -Milliseconds 100
    return Get-DatabaseConnection
}

<#
.SYNOPSIS
    Returns a database connection to the pool
#>
function Return-DatabaseConnection {
    param([System.Data.SqlClient.SqlConnection]$Connection)
    
    if (-not $Connection) { return }
    
    try {
        $script:connectionPool.Lock.EnterWriteLock()
        
        if ($script:connectionPool.Connections.Count -lt $script:connectionPool.MaxConnections) {
            $script:connectionPool.Connections.Enqueue($Connection)
        } else {
            $Connection.Dispose()
        }
        
        $script:connectionPool.ActiveConnections--
    }
    finally {
        $script:connectionPool.Lock.ExitWriteLock()
    }
}

<#
.SYNOPSIS
    Tests rate limiting for a client IP
#>
function Test-RateLimit {
    param([string]$ClientIP)
    
    $now = Get-Date
    $windowStart = $now.AddSeconds(-$script:rateLimitWindow)
    
    # Clean old entries
    $script:requestCounts.Keys | Where-Object { $script:requestCounts[$_].LastRequest -lt $windowStart } | ForEach-Object {
        $script:requestCounts.Remove($_)
    }
    
    # Check if client exists and is within limits
    if (-not $script:requestCounts.ContainsKey($ClientIP)) {
        $script:requestCounts[$ClientIP] = @{
            Count = 1
            LastRequest = $now
        }
        return $true
    }
    
    $client = $script:requestCounts[$ClientIP]
    if ($client.Count -ge $script:maxRequestsPerWindow) {
        return $false
    }
    
    $client.Count++
    $client.LastRequest = $now
    return $true
}

<#
.SYNOPSIS
    Gets cached data if available and not expired
#>
function Get-CachedData {
    param([string]$Key)
    
    if ($script:cache.ContainsKey($Key)) {
        $cached = $script:cache[$Key]
        if ((Get-Date) - $cached.Timestamp -lt [TimeSpan]::FromSeconds($script:cacheTimeout)) {
            return $cached.Data
        } else {
            $script:cache.Remove($Key)
        }
    }
    return $null
}

<#
.SYNOPSIS
    Sets data in cache with timestamp
#>
function Set-CachedData {
    param(
        [string]$Key,
        [object]$Data
    )
    
    $script:cache[$Key] = @{
        Data = $Data
        Timestamp = Get-Date
    }
    
    # Clean up old cache entries if cache gets too large
    if ($script:cache.Count -gt 100) {
        $oldestKeys = $script:cache.Keys | Sort-Object { $script:cache[$_].Timestamp } | Select-Object -First 20
        foreach ($key in $oldestKeys) {
            $script:cache.Remove($key)
        }
    }
}

<#
.SYNOPSIS
    Gets system health status
#>
function Get-SystemHealth {
    $health = @{
        Status = $script:healthStatus
        DatabaseConnection = $false
        ActiveConnections = $script:connectionPool.ActiveConnections
        PoolSize = $script:connectionPool.Connections.Count
        MaxConnections = $script:connectionPool.MaxConnections
        MemoryUsage = [math]::Round((Get-Process -Id $PID).WorkingSet64 / 1MB, 2)
        Uptime = (Get-Date) - $script:startTime
        LastError = $script:lastError
        CacheSize = $script:cache.Count
        RequestStats = $script:requestStats
        RateLimitStatus = @{
            ActiveClients = $script:requestCounts.Count
            MaxRequestsPerWindow = $script:maxRequestsPerWindow
        }
    }
    
    # Test database connection
    try {
        $conn = Get-DatabaseConnection
        $conn.Close()
        Return-DatabaseConnection $conn
        $health.DatabaseConnection = $true
    }
    catch {
        $script:healthStatus = "Degraded"
        $script:lastError = $_.Exception.Message
        $health.Status = $script:healthStatus
        $health.LastError = $script:lastError
    }
    
    return $health
}

<#
.SYNOPSIS
    Enhanced error handling with logging and user-friendly messages
#>
function Handle-Error {
    param(
        [string]$Operation,
        [System.Exception]$Exception,
        [string]$UserContext = "",
        [string]$RequestPath = ""
    )
    
    # Log error details
    $errorDetails = @{
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Operation = $Operation
        Context = $UserContext
        RequestPath = $RequestPath
        ErrorMessage = $Exception.Message
        ErrorType = $Exception.GetType().Name
        StackTrace = $Exception.StackTrace
    }
    
    Write-Error "Operation: $Operation, Context: $UserContext, Error: $($Exception.Message)"
    
    # Log to performance log if enabled
    if ($enablePerformanceLogging) {
        try {
            $errorDetails | ConvertTo-Json -Compress | Add-Content -Path $performanceLogPath -ErrorAction SilentlyContinue
        }
        catch {
            Write-Warning "Failed to write error to performance log: $_"
        }
    }
    
    # Update health status
    $script:lastError = $Exception.Message
    $script:requestStats.FailedRequests++
    
    # Return user-friendly error message
    return @"
    <div class="alert alert-danger">
        <h4><i class="fas fa-exclamation-triangle me-2"></i>Operation Failed</h4>
        <p>We encountered an issue while $Operation. Please try again later.</p>
        <hr>
        <small class="text-muted">
            <strong>Error ID:</strong> $(New-Guid)<br>
            <strong>Time:</strong> $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")<br>
            <strong>Operation:</strong> $Operation
        </small>
    </div>
"@
}

<#
.SYNOPSIS
    Updates request statistics
#>
function Update-RequestStats {
    param(
        [bool]$Success,
        [double]$ResponseTime
    )
    
    $script:requestStats.TotalRequests++
    
    if ($Success) {
        $script:requestStats.SuccessfulRequests++
    } else {
        $script:requestStats.FailedRequests++
    }
    
    # Update average response time
    $currentAvg = $script:requestStats.AverageResponseTime
    $totalRequests = $script:requestStats.TotalRequests
    $script:requestStats.AverageResponseTime = (($currentAvg * ($totalRequests - 1)) + $ResponseTime) / $totalRequests
    
    # Reset stats every hour
    if ((Get-Date) - $script:requestStats.LastReset -gt [TimeSpan]::FromHours(1)) {
        $script:requestStats.TotalRequests = 0
        $script:requestStats.SuccessfulRequests = 0
        $script:requestStats.FailedRequests = 0
        $script:requestStats.AverageResponseTime = 0
        $script:requestStats.LastReset = Get-Date
    }
}

<#
.SYNOPSIS
    Performs periodic health checks
#>
function Start-HealthMonitoring {
    $healthCheckJob = Start-Job -ScriptBlock {
        param($ConfigPath, $HealthCheckInterval)
        
        while ($true) {
            try {
                # Load config and check health
                $config = Get-Content $ConfigPath | ConvertFrom-Json
                $health = Get-SystemHealth
                
                # Log health status
                $logEntry = @{
                    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    HealthStatus = $health.Status
                    DatabaseConnection = $health.DatabaseConnection
                    ActiveConnections = $health.ActiveConnections
                    MemoryUsage = $health.MemoryUsage
                    Uptime = $health.Uptime.ToString()
                } | ConvertTo-Json -Compress
                
                Add-Content -Path "$env:TEMP\DashboardHealth.log" -Value $logEntry -ErrorAction SilentlyContinue
                
                Start-Sleep -Seconds $HealthCheckInterval
            }
            catch {
                Start-Sleep -Seconds $HealthCheckInterval
            }
        }
    } -ArgumentList $configPath, $config.HealthCheckInterval
    
    return $healthCheckJob
}

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
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css">
    <!-- Chart.js for analytics -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
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
        .health-indicator {
            width: 12px;
            height: 12px;
            border-radius: 50%;
            display: inline-block;
            margin-right: 8px;
        }
        .health-healthy { background-color: #28a745; }
        .health-degraded { background-color: #ffc107; }
        .health-unhealthy { background-color: #dc3545; }
        .stats-card {
            background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
            border: 1px solid #dee2e6;
        }
        .loading-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(255, 255, 255, 0.8);
            display: flex;
            justify-content: center;
            align-items: center;
            z-index: 9999;
        }
        .toast-container {
            position: fixed;
            top: 20px;
            right: 20px;
            z-index: 9999;
        }
    </style>
</head>
<body>
    <!-- Loading overlay -->
    <div class="loading-overlay" id="loadingOverlay" style="display: none;">
        <div class="spinner-border text-primary" role="status">
            <span class="visually-hidden">Loading...</span>
        </div>
    </div>

    <!-- Toast container for notifications -->
    <div class="toast-container" id="toastContainer"></div>

    <!-- Navigation header with live clock and health status -->
    <nav class="navbar navbar-expand-lg navbar-dark logo-header mb-4">
        <div class="container-fluid">
            <a class="navbar-brand" href="#">
                <i class="fas fa-server me-2"></i>$dashboardTitle
            </a>
            <div class="navbar-nav me-auto">
                <a class="nav-link" href="/">
                    <i class="fas fa-list me-1"></i>Systems
                </a>
                <a class="nav-link" href="/assets">
                    <i class="fas fa-clipboard-list me-1"></i>Asset Register
                </a>
                <a class="nav-link" href="/analytics">
                    <i class="fas fa-chart-line me-1"></i>Analytics
                </a>
            </div>
            <div class="navbar-text ms-auto d-flex align-items-center">
                <div class="me-3">
                    <span class="health-indicator health-healthy" id="health-indicator"></span>
                    <span id="health-status">Healthy</span>
                </div>
                <div class="me-3">
                    <i class="fas fa-clock me-1"></i>
                    <span id="live-clock"></span>
                </div>
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
        
        // Health status update function
        function updateHealthStatus() {
            fetch('/health')
                .then(response => response.json())
                .then(data => {
                    const indicator = document.getElementById('health-indicator');
                    const status = document.getElementById('health-status');
                    
                    // Update health indicator
                    indicator.className = 'health-indicator health-' + data.Status.toLowerCase();
                    
                    // Update status text
                    status.textContent = data.Status;
                    
                    // Update last updated timestamp
                    document.getElementById('last-updated').textContent = new Date().toLocaleString();
                })
                .catch(error => {
                    console.log('Health check failed:', error);
                });
        }
        
        // Update health status every 30 seconds
        setInterval(updateHealthStatus, 30000);
        updateHealthStatus();
        
        // Auto-refresh every 5 minutes (300,000 ms)
        setTimeout(function() {
            window.location.reload();
        }, 300000);
        
        // Initialize DataTables for tables with class 'datatable'
        `$(document).ready(function() {
            \$('.datatable').DataTable({
                responsive: true,
                pageLength: 10,
                lengthMenu: [5, 10, 25, 50, 100],
                stateSave: true,
                language: {
                    search: "_INPUT_",
                    searchPlaceholder: "Search..."
                }
            });
        });
        
        // Export button functionality
        function exportData(format) {
            showLoading(true);
            const searchParams = new URLSearchParams(window.location.search);
            const assetNumber = searchParams.get('AssetNumber') || '';
            const searchTerm = searchParams.get('search') || '';
            const page = searchParams.get('page') || 1;
            
            if (assetNumber) {
                window.location.href = \`/export?AssetNumber=\${assetNumber}&format=\${format}\`;
            } else {
                window.location.href = \`/export?search=\${searchTerm}&page=\${page}&format=\${format}\`;
            }
        }
        
        // Loading overlay functions
        function showLoading(show) {
            document.getElementById('loadingOverlay').style.display = show ? 'flex' : 'none';
        }
        
        // Toast notification function
        function showToast(message, type = 'info') {
            const toastContainer = document.getElementById('toastContainer');
            const toastId = 'toast-' + Date.now();
            const toastHtml = \`
                <div id="\${toastId}" class="toast align-items-center text-white bg-\${type === 'info' ? 'info' : type === 'success' ? 'success' : type === 'warning' ? 'warning' : 'danger'} border-0" role="alert" aria-live="assertive" aria-atomic="true">
                    <div class="d-flex">
                        <div class="toast-body">
                            \${message}
                        </div>
                        <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast" aria-label="Close"></button>
                    </div>
                </div>
            \`;
            
            toastContainer.insertAdjacentHTML('beforeend', toastHtml);
            const toastElement = document.getElementById(toastId);
            const toast = new bootstrap.Toast(toastElement, { delay: 3000 });
            toast.show();
            
            // Remove toast from DOM after it's hidden
            toastElement.addEventListener('hidden.bs.toast', function () {
                toastElement.remove();
            });
        }
        
        // Global error handling
        window.addEventListener('error', function(e) {
            console.error('Global error:', e.error);
            showToast('An unexpected error occurred', 'danger');
        });
        
        // Handle unhandled promise rejections
        window.addEventListener('unhandledrejection', function(e) {
            console.error('Unhandled promise rejection:', e.reason);
            showToast('An unexpected error occurred', 'danger');
            e.preventDefault();
        });
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
        
        # Create connection and command objects using connection pool
        $conn = Get-DatabaseConnection
        
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
        Return-DatabaseConnection $conn

        return @{
            Records = $results | Select-Object AssetNumber, HostName, OS, TotalRAMGB, ScanDate
            TotalCount = if ($results.Rows.Count -gt 0) { $results.Rows[0].TotalCount } else { 0 }
        }
    }
    catch {
        if ($conn) { Return-DatabaseConnection $conn }
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
        # Create connection using connection pool
        $conn = Get-DatabaseConnection

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
            n.DNSServers,
            n.DHCPServer,
            n.IPConfigType
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

        # Asset register info
        Ensure-AssetRegisterSchema
        $assetQuery = @"
SELECT ar.AssetNumber, ISNULL(l.LocationName,'N/A') as LocationName, ar.CurrentCustodian, ar.Status, ar.IsWrittenOff, ar.LastUpdated, ar.CurrentLocationID
FROM AssetRegister ar
LEFT JOIN Locations l ON ar.CurrentLocationID = l.LocationID
WHERE ar.AssetNumber = @AssetNumber
"@
        $cmd = New-Object System.Data.SqlClient.SqlCommand($assetQuery, $conn)
        $cmd.Parameters.AddWithValue("@AssetNumber", $AssetNumber) | Out-Null
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dsAsset = New-Object System.Data.DataSet
        $adapter.Fill($dsAsset) | Out-Null
        $assetReg = if ($dsAsset.Tables.Count -gt 0 -and $dsAsset.Tables[0].Rows.Count -gt 0) { $dsAsset.Tables[0].Rows[0] } else { $null }

        Return-DatabaseConnection $conn

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
                    <li class="nav-item" role="presentation">
                        <button class="nav-link" id="asset-tab" data-bs-toggle="tab" data-bs-target="#asset" type="button" role="tab">
                            <i class="fas fa-clipboard-list me-1"></i>Asset
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
                                    $(Render-DetailRow "DNS Servers" $details["DNSServers"])
                                    $(Render-DetailRow "Default Gateway" $details["DefaultGateway"])
                                    $(Render-DetailRow "DHCP Server" $details["DHCPServer"])
                                    $(Render-DetailRow "IP Config Type" $details["IPConfigType"])
                                </table>
                            </div>
                        </div>
                    </div>
                    
                    <!-- Asset tab -->
                    <div class="tab-pane fade" id="asset" role="tabpanel">
                        <div class="row mt-3">
                            <div class="col-md-6">
                                <h6><i class="fas fa-clipboard-list me-2"></i>Asset Register</h6>
                                <table class="table table-sm">
                                    $(Render-DetailRow "Location" $(if ($assetReg) { $assetReg["LocationName"] } else { "N/A" }))
                                    $(Render-DetailRow "Custodian" $(if ($assetReg) { $assetReg["CurrentCustodian"] } else { "N/A" }))
                                    $(Render-DetailRow "Status" $(if ($assetReg) { $assetReg["Status"] } else { "N/A" }))
                                    $(Render-DetailRow "Written Off" $(if ($assetReg -and $assetReg["IsWrittenOff"]) { "Yes" } else { "No" }))
                                    $(Render-DetailRow "Last Updated" $(if ($assetReg -and $assetReg["LastUpdated"]) { ([DateTime]$assetReg["LastUpdated"]).ToString("yyyy-MM-dd HH:mm") } else { "N/A" }))
                                </table>
                                <a class="btn btn-sm btn-outline-primary" href="/assets?AssetNumber=$(Encode-HTML $details["AssetNumber"])" target="_blank"><i class="fas fa-pen me-1"></i>Update in Asset Register</a>
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
        if ($conn) { Return-DatabaseConnection $conn }
        return Handle-Error -Operation "loading system details" -Error $_ -UserContext "AssetNumber: $AssetNumber"
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
        $conn = Get-DatabaseConnection

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

        Return-DatabaseConnection $conn

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
        if ($conn) { Return-DatabaseConnection $conn }
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

    # Quick analytics summary
    $quickAnalytics = @"
    <div class="row mb-4">
        <div class="col-12">
            <div class="card">
                <div class="card-header bg-info text-white">
                    <div class="d-flex justify-content-between align-items-center">
                        <h6 class="mb-0">
                            <i class="fas fa-chart-bar me-2"></i>Quick Analytics Summary
                        </h6>
                        <a href="/analytics" class="btn btn-sm btn-light">
                            <i class="fas fa-chart-line me-1"></i>View Full Analytics
                        </a>
                    </div>
                </div>
                <div class="card-body">
                    <div class="row text-center">
                        <div class="col-md-3">
                            <div class="border-end">
                                <h4 class="text-primary mb-1">$($systemData.TotalCount)</h4>
                                <small class="text-muted">Total Systems</small>
                            </div>
                        </div>
                        <div class="col-md-3">
                            <div class="border-end">
                                <h4 class="text-success mb-1">$([math]::Round(($systemData.TotalCount / 30), 0))</h4>
                                <small class="text-muted">Avg. Systems/Day</small>
                            </div>
                        </div>
                        <div class="col-md-3">
                            <div class="border-end">
                                <h4 class="text-warning mb-1">$([math]::Round(($systemData.TotalCount / 7), 0))</h4>
                                <small class="text-muted">This Week</small>
                            </div>
                        </div>
                        <div class="col-md-3">
                            <div>
                                <h4 class="text-info mb-1">$([math]::Round(($systemData.TotalCount / 90), 0))</h4>
                                <small class="text-muted">This Quarter</small>
                            </div>
                        </div>
                    </div>
                </div>
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
    $output = $htmlHeader + $searchForm + $resultsCount + $quickAnalytics + @"
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

#region ANALYTICS FUNCTIONS

<#
.SYNOPSIS
    Retrieves system trends and analytics data for charting
#>
function Get-SystemAnalytics {
    param(
        [string]$TimeRange = "30d",  # 7d, 30d, 90d, 1y
        [string]$Metric = "all"      # all, memory, storage, performance, software
    )
    
    try {
        $conn = Get-DatabaseConnection
        
        # Calculate date range
        $endDate = Get-Date
        $startDate = switch ($TimeRange) {
            "7d" { $endDate.AddDays(-7) }
            "30d" { $endDate.AddDays(-30) }
            "90d" { $endDate.AddDays(-90) }
            "1y" { $endDate.AddYears(-1) }
            default { $endDate.AddDays(-30) }
        }
        
        $analytics = @{
            TimeRange = $TimeRange
            StartDate = $startDate.ToString("yyyy-MM-dd")
            EndDate = $endDate.ToString("yyyy-MM-dd")
            Charts = @{}
            Insights = @{}
        }
        
        # System growth trends
        if ($Metric -eq "all" -or $Metric -eq "growth") {
            $growthQuery = @"
            SELECT 
                CAST(ScanDate AS DATE) as Date,
                COUNT(*) as NewSystems
            FROM Systems 
            WHERE ScanDate >= @StartDate AND ScanDate <= @EndDate
            GROUP BY CAST(ScanDate AS DATE)
            ORDER BY Date
"@
            $cmd = New-Object System.Data.SqlClient.SqlCommand($growthQuery, $conn)
            $cmd.Parameters.AddWithValue("@StartDate", $startDate) | Out-Null
            $cmd.Parameters.AddWithValue("@EndDate", $endDate) | Out-Null
            
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
            $dataset = New-Object System.Data.DataSet
            $adapter.Fill($dataset) | Out-Null
            $growthData = $dataset.Tables[0]
            
            $analytics.Charts.Growth = @{
                Labels = @()
                Data = @()
                Cumulative = @()
            }
            
            $cumulative = 0
            foreach ($row in $growthData.Rows) {
                $analytics.Charts.Growth.Labels += $row["Date"].ToString("MM/dd")
                $analytics.Charts.Growth.Data += [int]$row["NewSystems"]
                $cumulative += [int]$row["NewSystems"]
                $analytics.Charts.Growth.Cumulative += $cumulative
            }
            
            # Calculate growth insights
            if ($analytics.Charts.Growth.Data.Count -gt 1) {
                $avgGrowth = ($analytics.Charts.Growth.Data | Measure-Object -Average).Average
                $trend = if ($analytics.Charts.Growth.Data[-1] -gt $analytics.Charts.Growth.Data[0]) { "increasing" } else { "decreasing" }
                $analytics.Insights.Growth = @{
                    AverageDaily = [math]::Round($avgGrowth, 1)
                    Trend = $trend
                    TotalGrowth = $analytics.Charts.Growth.Data.Count
                    ProjectedMonthly = [math]::Round($avgGrowth * 30, 0)
                }
            }
        }
        
        # Memory distribution trends
        if ($Metric -eq "all" -or $Metric -eq "memory") {
            $memoryQuery = @"
            SELECT 
                h.TotalRAMGB,
                COUNT(*) as SystemCount
            FROM Hardware h
            JOIN Systems s ON h.AssetNumber = s.AssetNumber
            WHERE s.ScanDate >= @StartDate AND s.ScanDate <= @EndDate
            AND h.TotalRAMGB > 0
            GROUP BY h.TotalRAMGB
            ORDER BY h.TotalRAMGB
"@
            $cmd = New-Object System.Data.SqlClient.SqlCommand($memoryQuery, $conn)
            $cmd.Parameters.AddWithValue("@StartDate", $startDate) | Out-Null
            $cmd.Parameters.AddWithValue("@EndDate", $endDate) | Out-Null
            
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
            $dataset = New-Object System.Data.DataSet
            $adapter.Fill($dataset) | Out-Null
            $memoryData = $dataset.Tables[0]
            
            $analytics.Charts.Memory = @{
                Labels = @()
                Data = @()
            }
            
            $totalSystems = 0
            $totalMemory = 0
            foreach ($row in $memoryData.Rows) {
                $ramGB = [math]::Round($row["TotalRAMGB"], 0)
                $systemCount = [int]$row["SystemCount"]
                $analytics.Charts.Memory.Labels += "$ramGB GB"
                $analytics.Charts.Memory.Data += $systemCount
                $totalSystems += $systemCount
                $totalMemory += ($ramGB * $systemCount)
            }
            
            # Calculate memory insights
            if ($totalSystems -gt 0) {
                $avgMemory = $totalMemory / $totalSystems
                $analytics.Insights.Memory = @{
                    AverageRAM = [math]::Round($avgMemory, 1)
                    TotalSystems = $totalSystems
                    TotalRAM = [math]::Round($totalMemory, 0)
                    MostCommon = $analytics.Charts.Memory.Labels[0]
                }
            }
        }
        
        # Storage usage trends
        if ($Metric -eq "all" -or $Metric -eq "storage") {
            $storageQuery = @"
            SELECT 
                d.DeviceID,
                d.SizeGB,
                d.FreeGB,
                (d.SizeGB - d.FreeGB) as UsedGB,
                ROUND(((d.SizeGB - d.FreeGB) / d.SizeGB) * 100, 2) as UsagePercent
            FROM Disks d
            JOIN Hardware h ON d.HardwareID = h.HardwareID
            JOIN Systems s ON h.AssetNumber = s.AssetNumber
            WHERE s.ScanDate >= @StartDate AND s.ScanDate <= @EndDate
            AND d.SizeGB > 0
            ORDER BY UsagePercent DESC
"@
            $cmd = New-Object System.Data.SqlClient.SqlCommand($storageQuery, $conn)
            $cmd.Parameters.AddWithValue("@StartDate", $startDate) | Out-Null
            $cmd.Parameters.AddWithValue("@EndDate", $endDate) | Out-Null
            
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
            $dataset = New-Object System.Data.DataSet
            $adapter.Fill($dataset) | Out-Null
            $storageData = $dataset.Tables[0]
            
            $analytics.Charts.Storage = @{
                Labels = @()
                UsedData = @()
                FreeData = @()
                UsagePercent = @()
            }
            
            $totalSize = 0
            $totalUsed = 0
            $totalFree = 0
            foreach ($row in $storageData.Rows) {
                $analytics.Charts.Storage.Labels += $row["DeviceID"]
                $analytics.Charts.Storage.UsedData += [math]::Round($row["UsedGB"], 2)
                $analytics.Charts.Storage.FreeData += [math]::Round($row["FreeGB"], 2)
                $analytics.Charts.Storage.UsagePercent += [math]::Round($row["UsagePercent"], 1)
                
                $totalSize += $row["SizeGB"]
                $totalUsed += $row["UsedGB"]
                $totalFree += $row["FreeGB"]
            }
            
            # Calculate storage insights
            if ($totalSize -gt 0) {
                $overallUsage = ($totalUsed / $totalSize) * 100
                $analytics.Insights.Storage = @{
                    TotalSize = [math]::Round($totalSize, 0)
                    TotalUsed = [math]::Round($totalUsed, 0)
                    TotalFree = [math]::Round($totalFree, 0)
                    OverallUsage = [math]::Round($overallUsage, 1)
                    Status = if ($overallUsage -gt 90) { "Critical" } elseif ($overallUsage -gt 70) { "Warning" } else { "Healthy" }
                }
            }
        }
        
        # Operating system distribution
        if ($Metric -eq "all" -or $Metric -eq "os") {
            $osQuery = @"
            SELECT 
                s.OS,
                COUNT(*) as SystemCount
            FROM Systems s
            WHERE s.ScanDate >= @StartDate AND s.ScanDate <= @EndDate
            GROUP BY s.OS
            ORDER BY SystemCount DESC
"@
            $cmd = New-Object System.Data.SqlClient.SqlCommand($osQuery, $conn)
            $cmd.Parameters.AddWithValue("@StartDate", $startDate) | Out-Null
            $cmd.Parameters.AddWithValue("@EndDate", $endDate) | Out-Null
            
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
            $dataset = New-Object System.Data.DataSet
            $adapter.Fill($dataset) | Out-Null
            $osData = $dataset.Tables[0]
            
            $analytics.Charts.OperatingSystems = @{
                Labels = @()
                Data = @()
            }
            
            $totalOS = 0
            foreach ($row in $osData.Rows) {
                $analytics.Charts.OperatingSystems.Labels += $row["OS"]
                $analytics.Charts.OperatingSystems.Data += [int]$row["SystemCount"]
                $totalOS += [int]$row["SystemCount"]
            }
            
            # Calculate OS insights
            if ($totalOS -gt 0) {
                $dominantOS = $analytics.Charts.OperatingSystems.Labels[0]
                $dominantCount = $analytics.Charts.OperatingSystems.Data[0]
                $dominantPercentage = [math]::Round(($dominantCount / $totalOS) * 100, 1)
                
                $analytics.Insights.OperatingSystems = @{
                    TotalSystems = $totalOS
                    DominantOS = $dominantOS
                    DominantPercentage = $dominantPercentage
                    Diversity = $analytics.Charts.OperatingSystems.Labels.Count
                }
            }
        }
        
        # Software installation trends
        if ($Metric -eq "all" -or $Metric -eq "software") {
            $softwareQuery = @"
            SELECT 
                CAST(s.InstallDate AS DATE) as Date,
                COUNT(*) as Installations
            FROM Software s
            WHERE s.IsApplication = 1 
            AND s.InstallDate >= @StartDate 
            AND s.InstallDate <= @EndDate
            GROUP BY CAST(s.InstallDate AS DATE)
            ORDER BY Date
"@
            $cmd = New-Object System.Data.SqlClient.SqlCommand($softwareQuery, $conn)
            $cmd.Parameters.AddWithValue("@StartDate", $startDate) | Out-Null
            $cmd.Parameters.AddWithValue("@EndDate", $endDate) | Out-Null
            
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
            $dataset = New-Object System.Data.DataSet
            $adapter.Fill($dataset) | Out-Null
            $softwareData = $dataset.Tables[0]
            
            $analytics.Charts.Software = @{
                Labels = @()
                Data = @()
            }
            
            $totalInstallations = 0
            foreach ($row in $softwareData.Rows) {
                $analytics.Charts.Software.Labels += $row["Date"].ToString("MM/dd")
                $analytics.Charts.Software.Data += [int]$row["Installations"]
                $totalInstallations += [int]$row["Installations"]
            }
            
            # Calculate software insights
            if ($analytics.Charts.Software.Data.Count -gt 0) {
                $avgInstallations = $totalInstallations / $analytics.Charts.Software.Data.Count
                $analytics.Insights.Software = @{
                    TotalInstallations = $totalInstallations
                    AverageDaily = [math]::Round($avgInstallations, 1)
                    ActiveDays = $analytics.Charts.Software.Data.Count
                }
            }
        }
        
        # Performance metrics
        if ($Metric -eq "all" -or $Metric -eq "performance") {
            $performanceQuery = @"
            SELECT 
                h.CPUCores,
                h.TotalRAMGB,
                COUNT(*) as SystemCount
            FROM Hardware h
            JOIN Systems s ON h.AssetNumber = s.AssetNumber
            WHERE s.ScanDate >= @StartDate AND s.ScanDate <= @EndDate
            GROUP BY h.CPUCores, h.TotalRAMGB
            ORDER BY h.CPUCores, h.TotalRAMGB
"@
            $cmd = New-Object System.Data.SqlClient.SqlCommand($performanceQuery, $conn)
            $cmd.Parameters.AddWithValue("@StartDate", $startDate) | Out-Null
            $cmd.Parameters.AddWithValue("@EndDate", $endDate) | Out-Null
            
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
            $dataset = New-Object System.Data.DataSet
            $adapter.Fill($dataset) | Out-Null
            $performanceData = $dataset.Tables[0]
            
            $analytics.Charts.Performance = @{
                Labels = @()
                Data = @()
            }
            
            $totalPerformance = 0
            $avgCores = 0
            $avgRAM = 0
            foreach ($row in $performanceData.Rows) {
                $label = "$($row["CPUCores"]) Cores, $([math]::Round($row["TotalRAMGB"], 0)) GB RAM"
                $analytics.Charts.Performance.Labels += $label
                $analytics.Charts.Performance.Data += [int]$row["SystemCount"]
                $totalPerformance += [int]$row["SystemCount"]
                $avgCores += ($row["CPUCores"] * [int]$row["SystemCount"])
                $avgRAM += ($row["TotalRAMGB"] * [int]$row["SystemCount"])
            }
            
            # Calculate performance insights
            if ($totalPerformance -gt 0) {
                $analytics.Insights.Performance = @{
                    TotalSystems = $totalPerformance
                    AverageCores = [math]::Round($avgCores / $totalPerformance, 1)
                    AverageRAM = [math]::Round($avgRAM / $totalPerformance, 1)
                    PerformanceLevels = $analytics.Charts.Performance.Labels.Count
                }
            }
        }
        
        Return-DatabaseConnection $conn
        return $analytics
    }
    catch {
        if ($conn) { Return-DatabaseConnection $conn }
        Write-Error "Failed to retrieve analytics data: $_"
        return $null
    }
}

<#
.SYNOPSIS
    Generates HTML for analytics dashboard with charts
#>
function Render-AnalyticsDashboard {
    param(
        [string]$TimeRange = "30d",
        [string]$Metric = "all"
    )
    
    $analytics = Get-SystemAnalytics -TimeRange $TimeRange -Metric $Metric
    
    if (-not $analytics) {
        return "<div class='alert alert-danger'>Failed to load analytics data</div>"
    }
    
    $html = @"
    <div class="card mb-4">
        <div class="card-header bg-primary text-white">
            <div class="d-flex justify-content-between align-items-center">
                <h5 class="mb-0">
                    <i class="fas fa-chart-line me-2"></i>System Analytics Dashboard
                </h5>
                <div class="btn-group">
                    <button class="btn btn-sm btn-light" onclick="changeTimeRange('7d')">7 Days</button>
                    <button class="btn btn-sm btn-light" onclick="changeTimeRange('30d')">30 Days</button>
                    <button class="btn btn-sm btn-light" onclick="changeTimeRange('90d')">90 Days</button>
                    <button class="btn btn-sm btn-light" onclick="changeTimeRange('1y')">1 Year</button>
                </div>
            </div>
        </div>
        <div class="card-body">
            <!-- Analytics Insights Summary -->
            <div class="row mb-4">
                <div class="col-12">
                    <h6><i class="fas fa-lightbulb me-2"></i>Key Insights</h6>
                    <div class="row">
                        $(if ($analytics.Insights.Growth) { @"
                        <div class="col-md-3 mb-3">
                            <div class="card border-primary">
                                <div class="card-body text-center">
                                    <h6 class="card-title text-primary">Growth Trend</h6>
                                    <h4 class="text-success">$($analytics.Insights.Growth.AverageDaily)</h4>
                                    <small class="text-muted">Avg. systems/day</small>
                                    <div class="mt-2">
                                        <span class="badge bg-$(if ($analytics.Insights.Growth.Trend -eq 'increasing') { 'success' } else { 'warning' })">
                                            $($analytics.Insights.Growth.Trend.ToUpper())
                                        </span>
                                    </div>
                                </div>
                            </div>
                        </div>
"@ })
                        $(if ($analytics.Insights.Memory) { @"
                        <div class="col-md-3 mb-3">
                            <div class="card border-info">
                                <div class="card-body text-center">
                                    <h6 class="card-title text-info">Memory Profile</h6>
                                    <h4 class="text-info">$($analytics.Insights.Memory.AverageRAM) GB</h4>
                                    <small class="text-muted">Average RAM</small>
                                    <div class="mt-2">
                                        <span class="badge bg-secondary">$($analytics.Insights.Memory.TotalSystems) systems</span>
                                    </div>
                                </div>
                            </div>
                        </div>
"@ })
                        $(if ($analytics.Insights.Storage) { @"
                        <div class="col-md-3 mb-3">
                            <div class="card border-$(if ($analytics.Insights.Storage.Status -eq 'Critical') { 'danger' } elseif ($analytics.Insights.Storage.Status -eq 'Warning') { 'warning' } else { 'success' })">
                                <div class="card-body text-center">
                                    <h6 class="card-title text-$(if ($analytics.Insights.Storage.Status -eq 'Critical') { 'danger' } elseif ($analytics.Insights.Storage.Status -eq 'Warning') { 'warning' } else { 'success' })">Storage Status</h6>
                                    <h4 class="text-$(if ($analytics.Insights.Storage.Status -eq 'Critical') { 'danger' } elseif ($analytics.Insights.Storage.Status -eq 'Warning') { 'warning' } else { 'success' })">$($analytics.Insights.Storage.OverallUsage)%</h4>
                                    <small class="text-muted">Overall usage</small>
                                    <div class="mt-2">
                                        <span class="badge bg-$(if ($analytics.Insights.Storage.Status -eq 'Critical') { 'danger' } elseif ($analytics.Insights.Storage.Status -eq 'Warning') { 'warning' } else { 'success' })">$($analytics.Insights.Storage.Status)</span>
                                    </div>
                                </div>
                            </div>
                        </div>
"@ })
                        $(if ($analytics.Insights.OperatingSystems) { @"
                        <div class="col-md-3 mb-3">
                            <div class="card border-warning">
                                <div class="card-body text-center">
                                    <h6 class="card-title text-warning">OS Diversity</h6>
                                    <h4 class="text-warning">$($analytics.Insights.OperatingSystems.DominantPercentage)%</h4>
                                    <small class="text-muted">$($analytics.Insights.OperatingSystems.DominantOS)</small>
                                    <div class="mt-2">
                                        <span class="badge bg-secondary">$($analytics.Insights.OperatingSystems.Diversity) variants</span>
                                    </div>
                                </div>
                            </div>
                        </div>
"@ })
                    </div>
                </div>
            </div>
            
            <!-- Export Options -->
            <div class="row mb-4">
                <div class="col-12">
                    <div class="card">
                        <div class="card-header">
                            <h6><i class="fas fa-download me-2"></i>Export Analytics Data</h6>
                        </div>
                        <div class="card-body">
                            <div class="btn-group">
                                <button class="btn btn-outline-primary" onclick="exportAnalytics('csv')">
                                    <i class="fas fa-file-csv me-1"></i>Export CSV
                                </button>
                                <button class="btn btn-outline-primary" onclick="exportAnalytics('json')">
                                    <i class="fas fa-file-code me-1"></i>Export JSON
                                </button>
                                <button class="btn btn-outline-primary" onclick="exportAnalytics('pdf')">
                                    <i class="fas fa-file-pdf me-1"></i>Export PDF
                                </button>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- System Growth Chart -->
            <div class="row mb-4">
                <div class="col-12">
                    <div class="card">
                        <div class="card-header">
                            <h6><i class="fas fa-chart-line me-2"></i>System Growth Trends</h6>
                        </div>
                        <div class="card-body">
                            <canvas id="growthChart" width="400" height="200"></canvas>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Memory Distribution Chart -->
            <div class="row mb-4">
                <div class="col-md-6">
                    <div class="card">
                        <div class="card-header">
                            <h6><i class="fas fa-memory me-2"></i>Memory Distribution</h6>
                        </div>
                        <div class="card-body">
                            <canvas id="memoryChart" width="400" height="200"></canvas>
                        </div>
                    </div>
                </div>
                <div class="col-md-6">
                    <div class="card">
                        <div class="card-header">
                            <h6><i class="fas fa-hdd me-2"></i>Storage Usage</h6>
                        </div>
                        <div class="card-body">
                            <canvas id="storageChart" width="400" height="200"></canvas>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- OS Distribution and Performance Charts -->
            <div class="row mb-4">
                <div class="col-md-6">
                    <div class="card">
                        <div class="card-header">
                            <h6><i class="fas fa-windows me-2"></i>Operating System Distribution</h6>
                        </div>
                        <div class="card-body">
                            <canvas id="osChart" width="400" height="200"></canvas>
                        </div>
                    </div>
                </div>
                <div class="col-md-6">
                    <div class="card">
                        <div class="card-header">
                            <h6><i class="fas fa-microchip me-2"></i>Performance Distribution</h6>
                        </div>
                        <div class="card-body">
                            <canvas id="performanceChart" width="400" height="200"></canvas>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    
    <script>
        // Chart.js configuration and data
        const analyticsData = $(ConvertTo-Json -Depth 10 $analytics);
        
        // System Growth Chart
        const growthCtx = document.getElementById('growthChart').getContext('2d');
        new Chart(growthCtx, {
            type: 'line',
            data: {
                labels: analyticsData.Charts.Growth.Labels,
                datasets: [{
                    label: 'New Systems',
                    data: analyticsData.Charts.Growth.Data,
                    borderColor: 'rgb(75, 192, 192)',
                    backgroundColor: 'rgba(75, 192, 192, 0.2)',
                    tension: 0.1
                }, {
                    label: 'Cumulative Systems',
                    data: analyticsData.Charts.Growth.Cumulative,
                    borderColor: 'rgb(255, 99, 132)',
                    backgroundColor: 'rgba(255, 99, 132, 0.2)',
                    tension: 0.1
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    title: {
                        display: true,
                        text: 'System Growth Over Time'
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true
                    }
                }
            }
        });
        
        // Memory Distribution Chart
        const memoryCtx = document.getElementById('memoryChart').getContext('2d');
        new Chart(memoryCtx, {
            type: 'doughnut',
            data: {
                labels: analyticsData.Charts.Memory.Labels,
                datasets: [{
                    data: analyticsData.Charts.Memory.Data,
                    backgroundColor: [
                        '#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0',
                        '#9966FF', '#FF9F40', '#FF6384', '#C9CBCF'
                    ]
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    title: {
                        display: true,
                        text: 'Memory Distribution'
                    }
                }
            }
        });
        
        // Storage Usage Chart
        const storageCtx = document.getElementById('storageChart').getContext('2d');
        new Chart(storageCtx, {
            type: 'bar',
            data: {
                labels: analyticsData.Charts.Storage.Labels,
                datasets: [{
                    label: 'Used (GB)',
                    data: analyticsData.Charts.Storage.UsedData,
                    backgroundColor: 'rgba(255, 99, 132, 0.8)'
                }, {
                    label: 'Free (GB)',
                    data: analyticsData.Charts.Storage.FreeData,
                    backgroundColor: 'rgba(75, 192, 192, 0.8)'
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    title: {
                        display: true,
                        text: 'Storage Usage by Device'
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true
                    }
                }
            }
        });
        
        // OS Distribution Chart
        const osCtx = document.getElementById('osChart').getContext('2d');
        new Chart(osCtx, {
            type: 'pie',
            data: {
                labels: analyticsData.Charts.OperatingSystems.Labels,
                datasets: [{
                    data: analyticsData.Charts.OperatingSystems.Data,
                    backgroundColor: [
                        '#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0',
                        '#9966FF', '#FF9F40', '#FF6384', '#C9CBCF'
                    ]
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    title: {
                        display: true,
                        text: 'Operating System Distribution'
                    }
                }
            }
        });
        
        // Performance Distribution Chart
        const performanceCtx = document.getElementById('performanceChart').getContext('2d');
        new Chart(performanceCtx, {
            type: 'horizontalBar',
            data: {
                labels: analyticsData.Charts.Performance.Labels,
                datasets: [{
                    label: 'System Count',
                    data: analyticsData.Charts.Performance.Data,
                    backgroundColor: 'rgba(54, 162, 235, 0.8)'
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    title: {
                        display: true,
                        text: 'Performance Distribution'
                    }
                },
                scales: {
                    x: {
                        beginAtZero: true
                    }
                }
            }
        });
        
        // Time range change function
        function changeTimeRange(range) {
            window.location.href = '/analytics?range=' + range;
        }
        
        // Export analytics function
        function exportAnalytics(format) {
            const timeRange = new URLSearchParams(window.location.search).get('range') || '30d';
            window.location.href = '/export-analytics?range=' + timeRange + '&format=' + format;
        }
    </script>
"@
    
    return $html
}

#endregion

#region ASSET REGISTER FUNCTIONS

function Ensure-AssetRegisterSchema {
    if (-not $assetRegisterEnabled) { return }
    try {
        $script:schemaInitLock.EnterWriteLock()
        $conn = Get-DatabaseConnection
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = @"
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Locations')
BEGIN
    CREATE TABLE Locations (
        LocationID INT IDENTITY(1,1) PRIMARY KEY,
        LocationName NVARCHAR(100) NOT NULL,
        LocationType NVARCHAR(50) NULL,
        IsActive BIT NOT NULL DEFAULT 1,
        DateCreated DATETIME NOT NULL DEFAULT GETDATE(),
        CreatedBy NVARCHAR(100) NULL
    );
    CREATE UNIQUE INDEX UX_Locations_Name ON Locations(LocationName);
END

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'AssetRegister')
BEGIN
    CREATE TABLE AssetRegister (
        RegisterID INT IDENTITY(1,1) PRIMARY KEY,
        AssetNumber VARCHAR(50) NOT NULL,
        CurrentLocationID INT NULL,
        CurrentCustodian NVARCHAR(100) NULL,
        Status NVARCHAR(50) NOT NULL DEFAULT 'Active',
        IsWrittenOff BIT NOT NULL DEFAULT 0,
        CheckInDate DATETIME NULL,
        CheckOutDate DATETIME NULL,
        Notes NVARCHAR(500) NULL,
        LastUpdated DATETIME NOT NULL DEFAULT GETDATE(),
        LastUpdatedBy NVARCHAR(100) NULL
    );
    CREATE INDEX IX_AssetRegister_AssetNumber ON AssetRegister(AssetNumber);
END

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'AssetMovements')
BEGIN
    CREATE TABLE AssetMovements (
        MovementID INT IDENTITY(1,1) PRIMARY KEY,
        AssetNumber VARCHAR(50) NOT NULL,
        FromLocationID INT NULL,
        ToLocationID INT NULL,
        ChangedBy NVARCHAR(100) NULL,
        Action NVARCHAR(50) NOT NULL,
        Notes NVARCHAR(500) NULL,
        ChangeDate DATETIME NOT NULL DEFAULT GETDATE()
    );
    CREATE INDEX IX_AssetMovements_AssetNumber ON AssetMovements(AssetNumber);
    CREATE INDEX IX_AssetMovements_ChangeDate ON AssetMovements(ChangeDate);
END
"@
        $null = $cmd.ExecuteNonQuery()
    }
    catch {
        Write-Warning "Failed to ensure Asset Register schema: $_"
    }
    finally {
        if ($conn) { Return-DatabaseConnection $conn }
        $script:schemaInitLock.ExitWriteLock()
    }
}

function Get-Locations {
    if (-not $assetRegisterEnabled) { return $null }
    try {
        $conn = Get-DatabaseConnection
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT LocationID, LocationName FROM Locations WHERE IsActive = 1 ORDER BY LocationName"
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $ds = New-Object System.Data.DataSet
        $adapter.Fill($ds) | Out-Null
        return $ds.Tables[0]
    }
    catch {
        Write-Warning "Failed to load locations: $_"
        return $null
    }
    finally {
        if ($conn) { Return-DatabaseConnection $conn }
    }
}

function Ensure-Location {
    param([Parameter(Mandatory)][string]$LocationName)
    if (-not $assetRegisterEnabled -or [string]::IsNullOrWhiteSpace($LocationName)) { return $null }
    try {
        $conn = Get-DatabaseConnection
        # Try get existing
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT LocationID FROM Locations WHERE LocationName = @n"
        $cmd.Parameters.AddWithValue("@n", $LocationName) | Out-Null
        $id = $cmd.ExecuteScalar()
        if ($id) { return [int]$id }
        # Insert new
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "INSERT INTO Locations (LocationName, IsActive) VALUES (@n, 1); SELECT CAST(SCOPE_IDENTITY() AS INT);"
        $cmd.Parameters.AddWithValue("@n", $LocationName) | Out-Null
        return [int]$cmd.ExecuteScalar()
    }
    catch {
        Write-Warning "Failed to ensure location: $_"; return $null
    }
    finally {
        if ($conn) { Return-DatabaseConnection $conn }
    }
}

function Upsert-AssetRegister {
    param(
        [Parameter(Mandatory)][string]$AssetNumber,
        [int]$CurrentLocationID,
        [string]$CurrentCustodian,
        [string]$Status = "Active",
        [bool]$IsWrittenOff = $false,
        [DateTime]$CheckInDate,
        [DateTime]$CheckOutDate,
        [string]$Notes
    )
    if (-not $assetRegisterEnabled) { return }
    try {
        $conn = Get-DatabaseConnection
        $existsCmd = $conn.CreateCommand()
        $existsCmd.CommandText = "SELECT 1 FROM AssetRegister WHERE AssetNumber=@a"
        $existsCmd.Parameters.AddWithValue("@a", $AssetNumber) | Out-Null
        $exists = $null -ne $existsCmd.ExecuteScalar()

        $cmd = $conn.CreateCommand()
        if ($exists) {
            $cmd.CommandText = @"
UPDATE AssetRegister
SET CurrentLocationID=@l,
    CurrentCustodian=@c,
    Status=@s,
    IsWrittenOff=@w,
    CheckInDate=@ci,
    CheckOutDate=@co,
    Notes=@n,
    LastUpdated=GETDATE()
WHERE AssetNumber=@a
"@
        } else {
            $cmd.CommandText = @"
INSERT INTO AssetRegister (AssetNumber, CurrentLocationID, CurrentCustodian, Status, IsWrittenOff, CheckInDate, CheckOutDate, Notes, LastUpdated)
VALUES (@a, @l, @c, @s, @w, @ci, @co, @n, GETDATE())
"@
        }
        $null = $cmd.Parameters.AddWithValue("@a", $AssetNumber)
        $null = $cmd.Parameters.AddWithValue("@l", ($(if ($null -ne $CurrentLocationID -and $CurrentLocationID -ne 0) { $CurrentLocationID } else { [DBNull]::Value })))
        $null = $cmd.Parameters.AddWithValue("@c", ($(if ([string]::IsNullOrWhiteSpace($CurrentCustodian)) { [DBNull]::Value } else { $CurrentCustodian })))
        $null = $cmd.Parameters.AddWithValue("@s", ($(if ([string]::IsNullOrWhiteSpace($Status)) { [DBNull]::Value } else { $Status })))
        $null = $cmd.Parameters.AddWithValue("@w", ([int]([bool]$IsWrittenOff)))
        $null = $cmd.Parameters.AddWithValue("@ci", ($(if ($CheckInDate) { $CheckInDate } else { [DBNull]::Value })))
        $null = $cmd.Parameters.AddWithValue("@co", ($(if ($CheckOutDate) { $CheckOutDate } else { [DBNull]::Value })))
        $null = $cmd.Parameters.AddWithValue("@n", ($(if ([string]::IsNullOrWhiteSpace($Notes)) { [DBNull]::Value } else { $Notes })))
        $null = $cmd.ExecuteNonQuery()
    }
    catch {
        Write-Error "Failed to upsert AssetRegister for ${AssetNumber}: $_"
    }
    finally {
        if ($conn) { Return-DatabaseConnection $conn }
    }
}

function Add-AssetMovement {
    param(
        [Parameter(Mandatory)][string]$AssetNumber,
        [int]$FromLocationID,
        [int]$ToLocationID,
        [string]$ChangedBy,
        [string]$Action = "Update",
        [string]$Notes
    )
    try {
        $conn = Get-DatabaseConnection
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = @"
INSERT INTO AssetMovements (AssetNumber, FromLocationID, ToLocationID, ChangedBy, Action, Notes)
VALUES (@a, @f, @t, @u, @act, @n)
"@
        $null = $cmd.Parameters.AddWithValue("@a", $AssetNumber)
        $null = $cmd.Parameters.AddWithValue("@f", ($(if ($null -ne $FromLocationID -and $FromLocationID -ne 0) { $FromLocationID } else { [DBNull]::Value })))
        $null = $cmd.Parameters.AddWithValue("@t", ($(if ($null -ne $ToLocationID -and $ToLocationID -ne 0) { $ToLocationID } else { [DBNull]::Value })))
        $null = $cmd.Parameters.AddWithValue("@u", ($(if ([string]::IsNullOrWhiteSpace($ChangedBy)) { [DBNull]::Value } else { $ChangedBy })))
        $null = $cmd.Parameters.AddWithValue("@act", ($(if ([string]::IsNullOrWhiteSpace($Action)) { [DBNull]::Value } else { $Action })))
        $null = $cmd.Parameters.AddWithValue("@n", ($(if ([string]::IsNullOrWhiteSpace($Notes)) { [DBNull]::Value } else { $Notes })))
        $null = $cmd.ExecuteNonQuery()
    }
    catch {
        Write-Warning "Failed to insert asset movement for ${AssetNumber}: $_"
    }
    finally {
        if ($conn) { Return-DatabaseConnection $conn }
    }
}

function Get-AssetRegisterRecords {
    param(
        [string]$searchTerm = "",
        [int]$page = 1,
        [int]$pageSize = $defaultRowCount
    )
    
    try {
        $offset = ($page - 1) * $pageSize
        $conn = Get-DatabaseConnection
        
        $query = @"
SELECT 
    ar.AssetNumber,
    s.HostName,
    ISNULL(l.LocationName, 'N/A') AS LocationName,
    ar.CurrentCustodian,
    ar.Status,
    ar.IsWrittenOff,
    ar.CheckInDate,
    ar.CheckOutDate,
    ar.LastUpdated,
    COUNT(*) OVER() AS TotalCount
FROM AssetRegister ar
LEFT JOIN Systems s ON ar.AssetNumber = s.AssetNumber
LEFT JOIN Locations l ON ar.CurrentLocationID = l.LocationID
"@

        # Add WHERE clause if search term provided
        $whereClause = ""
        $params = @{}
        
        if (-not [string]::IsNullOrWhiteSpace($searchTerm)) {
            $whereClause = "WHERE ar.AssetNumber LIKE @searchTerm OR s.HostName LIKE @searchTerm OR l.LocationName LIKE @searchTerm OR ar.Status LIKE @searchTerm"
            $params["@searchTerm"] = "%$searchTerm%"
        }

        # Add pagination
        $query += @"
$whereClause
ORDER BY ar.LastUpdated DESC
OFFSET $offset ROWS FETCH NEXT $pageSize ROWS ONLY
"@
        
        $cmd = New-Object System.Data.SqlClient.SqlCommand($query, $conn)
        foreach ($key in $params.Keys) {
            $cmd.Parameters.AddWithValue($key, $params[$key]) | Out-Null
        }
        
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
        $dataset = New-Object System.Data.DataSet
        $adapter.Fill($dataset) | Out-Null
        $data = $dataset.Tables[0]
        
        return @{
    Records = if ($data -and $data.Rows.Count -gt 0) { 
        $data | Select-Object AssetNumber, HostName, LocationName, CurrentCustodian, Status, IsWrittenOff, CheckInDate, CheckOutDate, LastUpdated 
    } else { 
        @() 
    }
    TotalCount = if ($data -and $data.Rows.Count -gt 0) { $data.Rows[0].TotalCount } else { 0 }
}
    }
    catch {
        if ($conn) { Return-DatabaseConnection $conn }
        Write-Error "Failed to query asset register: $_"
        return @{ Records = @(); TotalCount = 0 }
    }
    finally {
        if ($conn) { Return-DatabaseConnection $conn }
    }
}

function Read-FormData {
    param([Parameter(Mandatory)][System.Net.HttpListenerRequest]$Request)
    $body = ""
    try {
        $reader = New-Object System.IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
        $body = $reader.ReadToEnd()
    }
    catch { }
    $map = @{}
    if (-not [string]::IsNullOrWhiteSpace($body)) {
        $parts = $body -split '&'
        foreach ($p in $parts) {
            $key, $val = $p.Split('=', 2)
            $map[[System.Uri]::UnescapeDataString($key)] = if ($val) { [System.Uri]::UnescapeDataString($val) } else { "" }
        }
    }
    return $map
}

function Validate-AssetNumber {
    param([Parameter(Mandatory)][string]$AssetNumber)
    if ([string]::IsNullOrWhiteSpace($AssetNumber)) { throw [System.ArgumentException]::new("Asset number cannot be empty") }
    if ($AssetNumber.Length -gt 50) { throw [System.ArgumentException]::new("Asset number cannot exceed 50 characters") }
    $conn = $null
    try {
        $conn = Get-DatabaseConnection
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT 1 FROM Systems WHERE AssetNumber = @assetNumber"
        $cmd.Parameters.AddWithValue("@assetNumber", $AssetNumber) | Out-Null
        $exists = $cmd.ExecuteScalar()
        if (-not $exists) { throw [System.ArgumentException]::new("Asset number $AssetNumber does not exist in Systems table") }
    }
    finally {
        if ($conn) { Return-DatabaseConnection $conn }
    }
}

function Save-AssetRegisterFromForm {
    param([Parameter(Mandatory)][hashtable]$FormData)
    
    $conn = $null
    $transaction = $null
    
    try {
        if (-not $assetRegisterEnabled) { return $false }
        Ensure-AssetRegisterSchema
        
        $asset = $FormData["AssetNumber"]
        Validate-AssetNumber -AssetNumber $asset
        
        # Input sanitization
        $locationId = [int]::TryParse($FormData["LocationID"], [ref]$null) ? [int]$FormData["LocationID"] : $null
        $locationName = [System.Net.WebUtility]::HtmlEncode($FormData["LocationName"] ?? "")
        $custodian = [System.Net.WebUtility]::HtmlEncode($FormData["CurrentCustodian"] ?? "")
        $status = [System.Net.WebUtility]::HtmlEncode($FormData["Status"] ?? "Active")
        $isWrittenOff = [bool]::TryParse($FormData["IsWrittenOff"], [ref]$null) ? [bool]$FormData["IsWrittenOff"] : $false
        $notes = [System.Net.WebUtility]::HtmlEncode($FormData["Notes"] ?? "")
        
        # Date handling
        $checkInDate = $null
        $checkOutDate = $null
        if ([DateTime]::TryParse($FormData["CheckInDate"], [ref]$checkInDate)) { /* use parsed date */ }
        if ([DateTime]::TryParse($FormData["CheckOutDate"], [ref]$checkOutDate)) { /* use parsed date */ }
        
        # Get connection and start transaction
        $conn = Get-DatabaseConnection
        $transaction = $conn.BeginTransaction()
        
        # Determine location id
        [int]$locId = $null
        if ($locationId) { $locId = $locationId }
        elseif ($locationName) { $locId = Ensure-Location -LocationName $locationName -Connection $conn -Transaction $transaction }
        
        # Perform operations in transaction
        Upsert-AssetRegister -AssetNumber $asset -CurrentLocationID $locId -CurrentCustodian $custodian `
            -Status $status -IsWrittenOff $isWrittenOff -Notes $notes -Connection $conn -Transaction $transaction
        
        $changedBy = if ($env:USERNAME) { $env:USERNAME } elseif ($env:USER) { $env:USER } else { "unknown" }
        Add-AssetMovement -AssetNumber $asset -FromLocationID $null -ToLocationID $locId `
            -ChangedBy $changedBy -Action "RegisterOrUpdate" -Notes $notes -Connection $conn -Transaction $transaction
        
        # Commit transaction
        $transaction.Commit()
        return $true
    }
    catch {
        # Rollback transaction on error
        if ($transaction) { $transaction.Rollback() }
        Write-Error "Failed to save asset register: $_"
        return $false
    }
    finally {
        # Always return connection to pool
        if ($conn) { Return-DatabaseConnection $conn }
    }
}

function Render-AssetRegisterPage {
    param(
        [string]$searchTerm = "",
        [int]$page = 1,
        [string]$assetNumberFilter = ""
    )
    if (-not $assetRegisterEnabled) { return "<div class='alert alert-warning'>Asset Register functionality is disabled in configuration</div>" }
    Ensure-AssetRegisterSchema
    $locations = Get-Locations
    
    # Generate location options for the dropdown
    $locationOptions = ""
    if ($locations -and $locations.Rows.Count -gt 0) {
        foreach ($row in $locations.Rows) {
            $locationOptions += "<option value='$($row['LocationID'])'>$(Encode-HTML $row['LocationName'])</option>"
        }
    }
    
    $data = Get-AssetRegisterRecords -searchTerm $searchTerm -page $page
    
    # Get existing asset data if we're filtering by a specific asset
    $existingAssetData = $null
    if (-not [string]::IsNullOrWhiteSpace($assetNumberFilter)) {
        try {
            $conn = Get-DatabaseConnection
            $query = "SELECT * FROM AssetRegister WHERE AssetNumber = @AssetNumber"
            $cmd = New-Object System.Data.SqlClient.SqlCommand($query, $conn)
            $cmd.Parameters.AddWithValue("@AssetNumber", $assetNumberFilter) | Out-Null
            
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
            $dataset = New-Object System.Data.DataSet
            $adapter.Fill($dataset) | Out-Null
            
            if ($dataset.Tables[0].Rows.Count -gt 0) {
                $existingAssetData = $dataset.Tables[0].Rows[0]
            }
        }
        catch {
            Write-Warning "Failed to load existing asset data: $_"
        }
        finally {
            if ($conn) { Return-DatabaseConnection $conn }
        }
    }
    
    # Build the register form with pre-populated values if available
    $currentCustodian = if ($existingAssetData) { $existingAssetData["CurrentCustodian"] } else { "" }
    $currentStatus = if ($existingAssetData) { $existingAssetData["Status"] } else { "Active" }
    $currentNotes = if ($existingAssetData) { $existingAssetData["Notes"] } else { "" }
    $isWrittenOffChecked = if ($existingAssetData -and $existingAssetData["IsWrittenOff"]) { "checked" } else { "" }
    
    $registerForm = @"
    <div class="card mb-4">
        <div class="card-header bg-secondary text-white">
            <i class="fas fa-clipboard-list me-2"></i>Register / Update Asset
        </div>
        <div class="card-body">
            <form method="POST" action="/assets/save" class="row g-3">
                <input type="hidden" name="redirect" value="/assets" />
                <div class="col-md-3">
                    <label class="form-label">Asset Number</label>
                    <input type="text" name="AssetNumber" class="form-control" required value="$(Encode-HTML $assetNumberFilter)" />
                </div>
                <div class="col-md-3">
                    <label class="form-label">Location</label>
                    <select name="LocationID" class="form-select">
                        <option value="">-- Select --</option>
                        $locationOptions
                    </select>
                    <small class="text-muted">Or add new location below</small>
                    <input type="text" name="LocationName" class="form-control mt-1" placeholder="New location name" />
                </div>
                <div class="col-md-3">
                    <label class="form-label">Custodian</label>
                    <input type="text" name="CurrentCustodian" class="form-control" value="$(Encode-HTML $currentCustodian)" />
                </div>
                <div class="col-md-2">
                    <label class="form-label">Status</label>
                    <select name="Status" class="form-select">
                        <option value="Active" $(if ($currentStatus -eq "Active") { "selected" })>Active</option>
                        <option value="In IT" $(if ($currentStatus -eq "In IT") { "selected" })>In IT</option>
                        <option value="Issued" $(if ($currentStatus -eq "Issued") { "selected" })>Issued</option>
                        <option value="In Repair" $(if ($currentStatus -eq "In Repair") { "selected" })>In Repair</option>
                        <option value="Disposed" $(if ($currentStatus -eq "Disposed") { "selected" })>Disposed</option>
                        <option value="WrittenOff" $(if ($currentStatus -eq "WrittenOff") { "selected" })>Written Off</option>
                    </select>
                </div>
                <div class="col-md-1 d-flex align-items-end">
                    <div class="form-check">
                        <input class="form-check-input" type="checkbox" name="IsWrittenOff" id="isWrittenOff" $isWrittenOffChecked>
                        <label class="form-check-label" for="isWrittenOff">Written off</label>
                    </div>
                </div>
                <div class="col-12">
                    <label class="form-label">Notes</label>
                    <textarea name="Notes" class="form-control" rows="2">$(Encode-HTML $currentNotes)</textarea>
                </div>
                <div class="col-12">
                    <button type="submit" class="btn btn-primary"><i class="fas fa-save me-1"></i>Save</button>
                </div>
            </form>
        </div>
    </div>"@
$rowsHtml = ""
if ($data.Records -and $data.Records.Rows -and $data.Records.Rows.Count -gt 0) {
    foreach ($r in $data.Records.Rows) {
        $rowsHtml += @"
    <tr>
        <td><a href="/?AssetNumber=$(Encode-HTML $r["AssetNumber"])" class="text-decoration-none">$(Encode-HTML $r["AssetNumber"])</a></td>
        <td>$(Encode-HTML $r["HostName"])</td>
        <td>$(Encode-HTML $r["LocationName"])</td>
        <td>$(Encode-HTML $r["CurrentCustodian"])</td>
        <td>$(Encode-HTML $r["Status"])</td>
        <td>$(if ($r["IsWrittenOff"]) { "Yes" } else { "No" })</td>
        <td>$(if ($r["LastUpdated"]) { ([DateTime]$r["LastUpdated"]).ToString("yyyy-MM-dd HH:mm") } else { "" })</td>
    </tr>
"@
    }
} else {
    $rowsHtml = @"
    <tr>
        <td colspan="7" class="text-center text-muted">No asset register records found</td>
    </tr>
"@
}
    
    $searchForm = @"
    <div class="card search-box mb-4"><div class="card-body">
        <form method="GET" action="/assets" class="row g-3">
            <div class="col-md-8"><div class="input-group">
                <span class="input-group-text"><i class="fas fa-search"></i></span>
                <input type="text" class="form-control" name="search" placeholder="Search by asset, host, location or status..." value="$(Encode-HTML $searchTerm)">
            </div></div>
            <div class="col-md-4">
                <button type="submit" class="btn btn-primary me-2"><i class="fas fa-search me-1"></i>Search</button>
                <a href="/assets" class="btn btn-outline-secondary"><i class="fas fa-times me-1"></i>Clear</a>
            </div>
        </form>
    </div></div>
"@
    $html = $htmlHeader + $searchForm + $registerForm + @"
    <div class="card">
        <div class="card-header bg-primary text-white">
            <i class="fas fa-clipboard-list me-2"></i>Asset Register
        </div>
        <div class="card-body">
            <div class="table-responsive">
                <table class="table table-sm table-striped datatable">
                    <thead><tr><th>Asset</th><th>Host</th><th>Location</th><th>Custodian</th><th>Status</th><th>Written Off</th><th>Updated</th></tr></thead>
                    <tbody>
                        $rowsHtml
                    </tbody>
                </table>
            </div>
        </div>
    </div>
"@ + $htmlFooter
    return $html
}

#endregion

#region WEB SERVER

# Start health monitoring
Write-Host "Starting health monitoring..." -ForegroundColor Yellow
$healthMonitoringJob = Start-HealthMonitoring

try {
    # Initialize HTTP listener
    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://localhost:$port/")
    $listener.Start()
    Write-Host "Dashboard running at http://localhost:$port" -ForegroundColor Green
    Write-Host "Health monitoring started" -ForegroundColor Green
    Write-Host "Press Ctrl+C to stop the dashboard" -ForegroundColor Yellow

    # Main request handling loop
    while ($listener.IsListening) {
        try {
            $context = $listener.GetContext()
            $request = $context.Request
            $response = $context.Response
            $startTime = Get-Date

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
                    $searchTerm = if ($queryParams.ContainsKey("search")) { $queryParams["search"] } else { "" }
                    $page = if ($queryParams.ContainsKey("page")) { [int]$queryParams["page"] } else { 1 }
                    $AssetNumber = if ($queryParams.ContainsKey("AssetNumber")) { $queryParams["AssetNumber"] } else { "" }
                    $format = if ($queryParams.ContainsKey("format")) { $queryParams["format"] } else { "csv" }
                    
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

                # Handle health check request
                if ($request.Url.LocalPath -eq "/health") {
                    $healthData = Get-SystemHealth
                    $jsonResponse = $healthData | ConvertTo-Json -Depth 5
                    $response.ContentType = "application/json"
                    $response.ContentLength64 = $jsonResponse.Length
                    $response.OutputStream.Write($jsonResponse, 0, $jsonResponse.Length)
                    $response.OutputStream.Close()
                    continue
                }

                # Handle analytics request
                if ($request.Url.LocalPath -eq "/analytics") {
                    $timeRange = if ($queryParams.ContainsKey("range")) { $queryParams["range"] } else { "30d" }
                    $metric = if ($queryParams.ContainsKey("metric")) { $queryParams["metric"] } else { "all" }
                    $analyticsHtml = Render-AnalyticsDashboard -TimeRange $timeRange -Metric $metric
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($analyticsHtml)
                    $response.ContentLength64 = $buffer.Length
                    $response.ContentType = "text/html; charset=utf-8"
                    $response.OutputStream.Write($buffer, 0, $buffer.Length)
                    $response.OutputStream.Close()
                    continue
                }

                # Handle analytics export request
                if ($request.Url.LocalPath -eq "/export-analytics") {
                    $startTime = Get-Date
                    $timeRange = if ($queryParams.ContainsKey("range")) { $queryParams["range"] } else { "30d" }
                    $format = if ($queryParams.ContainsKey("format")) { $queryParams["format"] } else { "csv" }
                    
                    $analyticsData = Get-SystemAnalytics -TimeRange $timeRange -Metric "all"
                    
                    if ($format -eq "csv") {
                        $csvContent = Export-AnalyticsToCSV -AnalyticsData $analyticsData
                        $buffer = [System.Text.Encoding]::UTF8.GetBytes($csvContent)
                        $response.ContentType = "text/csv"
                        $filename = "Analytics_$timeRange.csv"
                        $response.AddHeader("Content-Disposition", "attachment; filename=$filename")
                    }
                    elseif ($format -eq "json") {
                        $jsonContent = $analyticsData | ConvertTo-Json -Depth 10
                        $buffer = [System.Text.Encoding]::UTF8.GetBytes($jsonContent)
                        $response.ContentType = "application/json"
                        $filename = "Analytics_$timeRange.json"
                        $response.AddHeader("Content-Disposition", "attachment; filename=$filename")
                    }
                    elseif ($format -eq "pdf") {
                        $pdfContent = Export-AnalyticsToPDF -AnalyticsData $analyticsData
                        $buffer = [System.Text.Encoding]::UTF8.GetBytes($pdfContent)
                        $response.ContentType = "application/pdf"
                        $filename = "Analytics_$timeRange.pdf"
                        $response.AddHeader("Content-Disposition", "attachment; filename=$filename")
                    }
                    
                    $response.ContentLength64 = $buffer.Length
                    $response.OutputStream.Write($buffer, 0, $buffer.Length)
                    $response.OutputStream.Close()
                    
                    $duration = ((Get-Date) - $startTime).TotalMilliseconds
                    Log-Performance -operation "ExportAnalytics" -durationMs $duration -details "Format=$format, TimeRange=$timeRange"
                    
                    continue
                }

                # Handle Asset Register page
                if ($request.Url.LocalPath -eq "/assets") {
                    $searchTerm = if ($queryParams.ContainsKey("search")) { $queryParams["search"] } else { "" }
                    $page = if ($queryParams.ContainsKey("page")) { [int]$queryParams["page"] } else { 1 }
                    $assetNum = if ($queryParams.ContainsKey("AssetNumber")) { $queryParams["AssetNumber"] } else { "" }
                    $html = Render-AssetRegisterPage -searchTerm $searchTerm -page $page -assetNumberFilter $assetNum
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                    $response.ContentLength64 = $buffer.Length
                    $response.ContentType = "text/html; charset=utf-8"
                    $response.OutputStream.Write($buffer, 0, $buffer.Length)
                    $response.OutputStream.Close()
                    continue
                }

                # Handle Asset Register save
                if ($request.Url.LocalPath -eq "/assets/save") {
                    if ($request.HttpMethod -ne "POST") {
                        $response.StatusCode = 405
                        $response.OutputStream.Close()
                        continue
                    }
                    $form = Read-FormData -Request $request
                    Save-AssetRegisterFromForm -FormData $form
                    $redirect = if ($form.ContainsKey("redirect")) { $form["redirect"] } else { "/assets" }
                    $response.StatusCode = 302
                    $response.AddHeader("Location", $redirect)
                    $response.OutputStream.Close()
                    continue
                }

                # Get client IP for rate limiting
                $clientIP = $request.RemoteEndPoint.Address.ToString()
                
                # Test rate limit
                if (-not (Test-RateLimit -ClientIP $clientIP)) {
                    $response.StatusCode = 429 # Too Many Requests
                    $response.ContentType = "text/html"
                    $response.ContentLength64 = 0
                    $response.OutputStream.Close()
                    continue
                }

                # Get parameters with defaults for regular requests
                $searchTerm = if ($queryParams.ContainsKey("search")) { $queryParams["search"] } else { "" }
                $page = if ($queryParams.ContainsKey("page")) { [int]$queryParams["page"] } else { 1 }
                $AssetNumber = if ($queryParams.ContainsKey("AssetNumber")) { $queryParams["AssetNumber"] } else { "" }

                # Generate and send HTML response
                $html = Render-Dashboard -searchTerm $searchTerm -page $page -AssetNumber $AssetNumber
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)

                $response.ContentLength64 = $buffer.Length
                $response.ContentType = "text/html; charset=utf-8"
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()

                # Update request statistics
                Update-RequestStats -Success $true -ResponseTime ((Get-Date) - $startTime).TotalMilliseconds
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
                
                # Update request statistics
                Update-RequestStats -Success $false -ResponseTime ((Get-Date) - $startTime).TotalMilliseconds
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
    
    # Stop health monitoring job
    if ($healthMonitoringJob) {
        Stop-Job $healthMonitoringJob
        Remove-Job $healthMonitoringJob
    }
    
    # Log shutdown
    Write-Host "Dashboard server stopped" -ForegroundColor Yellow
}
#endregion

#region EXPORT FUNCTIONS

<#
.SYNOPSIS
    Exports analytics data to CSV format
#>
function Export-AnalyticsToCSV {
    param([object]$AnalyticsData)
    
    $csvLines = @()
    $csvLines += "System Analytics Report - $($AnalyticsData.TimeRange.ToUpper())"
    $csvLines += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $csvLines += "Time Range: $($AnalyticsData.StartDate) to $($AnalyticsData.EndDate)"
    $csvLines += ""
    
    # Growth data
    if ($AnalyticsData.Charts.Growth) {
        $csvLines += "System Growth Trends"
        $csvLines += "Date,New Systems,Cumulative Systems"
        for ($i = 0; $i -lt $AnalyticsData.Charts.Growth.Labels.Count; $i++) {
            $csvLines += "$($AnalyticsData.Charts.Growth.Labels[$i]),$($AnalyticsData.Charts.Growth.Data[$i]),$($AnalyticsData.Charts.Growth.Cumulative[$i])"
        }
        $csvLines += ""
    }
    
    # Memory data
    if ($AnalyticsData.Charts.Memory) {
        $csvLines += "Memory Distribution"
        $csvLines += "RAM (GB),System Count"
        for ($i = 0; $i -lt $AnalyticsData.Charts.Memory.Labels.Count; $i++) {
            $csvLines += "$($AnalyticsData.Charts.Memory.Labels[$i]),$($AnalyticsData.Charts.Memory.Data[$i])"
        }
        $csvLines += ""
    }
    
    # Storage data
    if ($AnalyticsData.Charts.Storage) {
        $csvLines += "Storage Usage"
        $csvLines += "Device,Used (GB),Free (GB),Usage (%)"
        for ($i = 0; $i -lt $AnalyticsData.Charts.Storage.Labels.Count; $i++) {
            $csvLines += "$($AnalyticsData.Charts.Storage.Labels[$i]),$($AnalyticsData.Charts.Storage.UsedData[$i]),$($AnalyticsData.Charts.Storage.FreeData[$i]),$($AnalyticsData.Charts.Storage.UsagePercent[$i])"
        }
        $csvLines += ""
    }
    
    # OS data
    if ($AnalyticsData.Charts.OperatingSystems) {
        $csvLines += "Operating System Distribution"
        $csvLines += "OS,System Count"
        for ($i = 0; $i -lt $AnalyticsData.Charts.OperatingSystems.Labels.Count; $i++) {
            $csvLines += "$($AnalyticsData.Charts.OperatingSystems.Labels[$i]),$($AnalyticsData.Charts.OperatingSystems.Data[$i])"
        }
        $csvLines += ""
    }
    
    # Performance data
    if ($AnalyticsData.Charts.Performance) {
        $csvLines += "Performance Distribution"
        $csvLines += "Specification,System Count"
        for ($i = 0; $i -lt $AnalyticsData.Charts.Performance.Labels.Count; $i++) {
            $csvLines += "$($AnalyticsData.Charts.Performance.Labels[$i]),$($AnalyticsData.Charts.Performance.Data[$i])"
        }
        $csvLines += ""
    }
    
    # Insights summary
    if ($AnalyticsData.Insights) {
        $csvLines += "Key Insights Summary"
        $csvLines += "Metric,Value,Details"
        
        if ($AnalyticsData.Insights.Growth) {
            $csvLines += "Growth Trend,$($AnalyticsData.Insights.Growth.Trend),Average: $($AnalyticsData.Insights.Growth.AverageDaily) systems/day"
            $csvLines += "Projected Monthly,,$($AnalyticsData.Insights.Growth.ProjectedMonthly) systems"
        }
        
        if ($AnalyticsData.Insights.Memory) {
            $csvLines += "Memory Profile,$($AnalyticsData.Insights.Memory.AverageRAM) GB,$($AnalyticsData.Insights.Memory.TotalSystems) systems"
        }
        
        if ($AnalyticsData.Insights.Storage) {
            $csvLines += "Storage Status,$($AnalyticsData.Insights.Storage.OverallUsage)%,$($AnalyticsData.Insights.Storage.Status)"
        }
        
        if ($AnalyticsData.Insights.OperatingSystems) {
            $csvLines += "OS Diversity,$($AnalyticsData.Insights.OperatingSystems.DominantPercentage)%,$($AnalyticsData.Insights.OperatingSystems.DominantOS)"
        }
    }
    
    return ($csvLines -join "`r`n")
}

<#
.SYNOPSIS
    Exports analytics data to PDF format (HTML-based for simplicity)
#>
function Export-AnalyticsToPDF {
    param([object]$AnalyticsData)
    
    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>System Analytics Report - $($AnalyticsData.TimeRange.ToUpper())</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { text-align: center; border-bottom: 2px solid #333; padding-bottom: 20px; margin-bottom: 30px; }
        .section { margin-bottom: 30px; }
        .section h2 { color: #2c3e50; border-bottom: 1px solid #bdc3c7; padding-bottom: 10px; }
        .insight-card { border: 1px solid #ddd; padding: 15px; margin: 10px 0; border-radius: 5px; }
        .insight-title { font-weight: bold; color: #2c3e50; }
        .insight-value { font-size: 18px; color: #27ae60; margin: 5px 0; }
        .insight-details { color: #7f8c8d; font-size: 14px; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f8f9fa; font-weight: bold; }
        .status-healthy { color: #27ae60; }
        .status-warning { color: #f39c12; }
        .status-critical { color: #e74c3c; }
    </style>
</head>
<body>
    <div class="header">
        <h1>System Analytics Report</h1>
        <h3>$($AnalyticsData.TimeRange.ToUpper()) Analysis</h3>
        <p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        <p>Time Range: $($AnalyticsData.StartDate) to $($AnalyticsData.EndDate)</p>
    </div>
    
    <div class="section">
        <h2>Key Insights</h2>
        $(if ($AnalyticsData.Insights.Growth) { @"
        <div class="insight-card">
            <div class="insight-title">Growth Trend</div>
            <div class="insight-value">$($AnalyticsData.Insights.Growth.AverageDaily) systems/day average</div>
            <div class="insight-details">Trend: $($AnalyticsData.Insights.Growth.Trend.ToUpper()) | Projected Monthly: $($AnalyticsData.Insights.Growth.ProjectedMonthly) systems</div>
        </div>
"@ })
        $(if ($AnalyticsData.Insights.Memory) { @"
        <div class="insight-card">
            <div class="insight-title">Memory Profile</div>
            <div class="insight-value">$($AnalyticsData.Insights.Memory.AverageRAM) GB average RAM</div>
            <div class="insight-details">Total Systems: $($AnalyticsData.Insights.Memory.TotalSystems) | Total RAM: $($AnalyticsData.Insights.Memory.TotalRAM) GB</div>
        </div>
"@ })
        $(if ($AnalyticsData.Insights.Storage) { @"
        <div class="insight-card">
            <div class="insight-title">Storage Status</div>
            <div class="insight-value status-$($AnalyticsData.Insights.Storage.Status.ToLower())">$($AnalyticsData.Insights.Storage.OverallUsage)% overall usage</div>
            <div class="insight-details">Status: $($AnalyticsData.Insights.Storage.Status) | Total Size: $($AnalyticsData.Insights.Storage.TotalSize) GB | Used: $($AnalyticsData.Insights.Storage.TotalUsed) GB</div>
        </div>
"@ })
        $(if ($AnalyticsData.Insights.OperatingSystems) { @"
        <div class="insight-card">
            <div class="insight-title">Operating System Diversity</div>
            <div class="insight-value">$($AnalyticsData.Insights.OperatingSystems.DominantOS) dominates with $($AnalyticsData.Insights.OperatingSystems.DominantPercentage)%</div>
            <div class="insight-details">Total Systems: $($AnalyticsData.Insights.OperatingSystems.TotalSystems) | Variants: $($AnalyticsData.Insights.OperatingSystems.Diversity)</div>
        </div>
"@ })
    </div>
    
    $(if ($AnalyticsData.Charts.Growth) { @"
    <div class="section">
        <h2>System Growth Trends</h2>
        <table>
            <thead>
                <tr><th>Date</th><th>New Systems</th><th>Cumulative Systems</th></tr>
            </thead>
            <tbody>
                $(for ($i = 0; $i -lt $AnalyticsData.Charts.Growth.Labels.Count; $i++) { @"
                <tr>
                    <td>$($AnalyticsData.Charts.Growth.Labels[$i])</td>
                    <td>$($AnalyticsData.Charts.Growth.Data[$i])</td>
                    <td>$($AnalyticsData.Charts.Growth.Cumulative[$i])</td>
                </tr>
"@ })
            </tbody>
        </table>
    </div>
"@ })
    
    $(if ($AnalyticsData.Charts.Memory) { @"
    <div class="section">
        <h2>Memory Distribution</h2>
        <table>
            <thead>
                <tr><th>RAM (GB)</th><th>System Count</th></tr>
            </thead>
            <tbody>
                $(for ($i = 0; $i -lt $AnalyticsData.Charts.Memory.Labels.Count; $i++) { @"
                <tr>
                    <td>$($AnalyticsData.Charts.Memory.Labels[$i])</td>
                    <td>$($AnalyticsData.Charts.Memory.Data[$i])</td>
                </tr>
"@ })
            </tbody>
        </table>
    </div>
"@ })
    
    $(if ($AnalyticsData.Charts.Storage) { @"
    <div class="section">
        <h2>Storage Usage</h2>
        <table>
            <thead>
                <tr><th>Device</th><th>Used (GB)</th><th>Free (GB)</th><th>Usage (%)</th></tr>
            </thead>
            <tbody>
                $(for ($i = 0; $i -lt $AnalyticsData.Charts.Storage.Labels.Count; $i++) { @"
                <tr>
                    <td>$($AnalyticsData.Charts.Storage.Labels[$i])</td>
                    <td>$($AnalyticsData.Charts.Storage.UsedData[$i])</td>
                    <td>$($AnalyticsData.Charts.Storage.FreeData[$i])</td>
                    <td>$($AnalyticsData.Charts.Storage.UsagePercent[$i])%</td>
                </tr>
"@ })
            </tbody>
        </table>
    </div>
"@ })
    
    $(if ($AnalyticsData.Charts.OperatingSystems) { @"
    <div class="section">
        <h2>Operating System Distribution</h2>
        <table>
            <thead>
                <tr><th>Operating System</th><th>System Count</th></tr>
            </thead>
            <tbody>
                $(for ($i = 0; $i -lt $AnalyticsData.Charts.OperatingSystems.Labels.Count; $i++) { @"
                <tr>
                    <td>$($AnalyticsData.Charts.OperatingSystems.Labels[$i])</td>
                    <td>$($AnalyticsData.Charts.OperatingSystems.Data[$i])</td>
                </tr>
"@ })
            </tbody>
        </table>
    </div>
"@ })
    
    $(if ($AnalyticsData.Charts.Performance) { @"
    <div class="section">
        <h2>Performance Distribution</h2>
        <table>
            <thead>
                <tr><th>Specification</th><th>System Count</th></tr>
            </thead>
            <tbody>
                $(for ($i = 0; $i -lt $AnalyticsData.Charts.Performance.Labels.Count; $i++) { @"
                <tr>
                    <td>$($AnalyticsData.Charts.Performance.Labels[$i])</td>
                    <td>$($AnalyticsData.Charts.Performance.Data[$i])</td>
                </tr>
"@ })
            </tbody>
        </table>
    </div>
"@ })
    
    <div class="section">
        <h2>Report Summary</h2>
        <p>This analytics report provides comprehensive insights into your system inventory trends, performance metrics, and infrastructure patterns.</p>
        <p>Use these insights to:</p>
        <ul>
            <li>Plan capacity and resource allocation</li>
            <li>Identify performance bottlenecks</li>
            <li>Track system growth and adoption</li>
            <li>Make informed infrastructure decisions</li>
            <li>Monitor storage and memory utilization</li>
        </ul>
    </div>
</body>
</html>
"@
    
    return $htmlContent
}

#endregion