<#
.SYNOPSIS
    System Inventory Dashboard Web Server
.DESCRIPTION
    Creates a local web dashboard to view system inventory from SQL database
    Runs on port 8080 with search functionality
.COMPANY
    North West Provincial Treasury 
.AUTHOR
    Lesedi Sebekedi
.VERSION
    2.0
.SECURITY
    Requires SQL read permissions and local admin for port binding
#>

# Import required modules
Import-Module SqlServer

# REGION: WEB SERVER CONFIGURATION
# -------------------------------
$port = 8080
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")

try {
    $listener.Start()
    Write-Host "Dashboard running at http://localhost:$port" -ForegroundColor Green
}
catch {
    Write-Host "Failed to start HTTP listener: $_" -ForegroundColor Red
    exit 1
}

# REGION: HELPER FUNCTIONS
# ------------------------

function Encode-HTML {
    <#
    .SYNOPSIS
        Encodes text for safe HTML output
    #>
    param([string]$text)
    
    if ([string]::IsNullOrEmpty($text)) { 
        return "" 
    }
    return [System.Net.WebUtility]::HtmlEncode($text)
}

function Parse-QueryString {
    <#
    .SYNOPSIS
        Parses URL query string into key-value pairs
    #>
    param([string]$query)
    
    $query = $query.TrimStart('?')
    $dict = @{}

    if (-not [string]::IsNullOrEmpty($query)) {
        $query.Split('&') | ForEach-Object {
            $key, $value = $_.Split('=', 2)
            $dict[$key] = [System.Uri]::UnescapeDataString($value)
        }
    }

    return $dict
}

function Escape-SqlLike {
    <#
    .SYNOPSIS
        Escapes special characters for SQL LIKE queries
    #>
    param([string]$input)
    
    if ([string]::IsNullOrEmpty($input)) { 
        return "" 
    }
    # Escape SQL wildcards by surrounding with []
    return $input -replace '([%_\[\]])', '[$1]'
}

function Get-SystemRecords {
    <#
    .SYNOPSIS
        Retrieves system records from database with optional search
    #>
    param(
        [string]$searchTerm,
        [string]$connectionString
    )
    
    # Configure secure connection
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    
    if ([string]::IsNullOrWhiteSpace($searchTerm)) {
        $sqlQuery = @"
        SELECT TOP 20 
            s.AssetNumber, 
            s.HostName, 
            sp.OSName, 
            sp.TotalRAMGB,
            s.ScanDate
        FROM Systems s
        JOIN SystemSpecs sp ON s.SystemID = sp.SystemID
        ORDER BY s.ScanDate DESC
"@
    }
    else {
        $escapedSearchTerm = Escape-SqlLike $searchTerm
        $searchPattern = "%$escapedSearchTerm%"

        $sqlQuery = @"
        SELECT TOP 20 
            s.AssetNumber, 
            s.HostName, 
            sp.OSName, 
            sp.TotalRAMGB,
            s.ScanDate
        FROM Systems s
        JOIN SystemSpecs sp ON s.SystemID = sp.SystemID
        WHERE s.AssetNumber LIKE '$searchPattern' ESCAPE '[' 
           OR s.HostName LIKE '$searchPattern' ESCAPE '['
        ORDER BY s.ScanDate DESC
"@
    }

    return Invoke-Sqlcmd -ConnectionString $connectionString -Query $sqlQuery
}

function Show-ErrorPage {
    <#
    .SYNOPSIS
        Displays error page when exceptions occur
    #>
    param($response, $errorMessage)
    
    $errorHtml = @"
<!DOCTYPE html>
<html>
<head>
    <title>Error</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-gray-100">
    <div class="container mx-auto px-4 py-8">
        <div class="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded relative" role="alert">
            <strong class="font-bold">Application Error!</strong>
            <span class="block sm:inline">$($errorMessage)</span>
        </div>
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

# REGION: MAIN WEB SERVER LOOP
# ----------------------------
while ($listener.IsListening) {
    try {
        # Get HTTP context
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response

        # Parse search query if present
        $searchTerm = ""
        if (-not [string]::IsNullOrEmpty($request.Url.Query)) {
            $queryParams = Parse-QueryString $request.Url.Query
            $searchTerm = $queryParams["search"]
        }

        # Database configuration
        $connectionString = "Server=PTLSEBEKEDI;Database=AssetDB;Integrated Security=True;TrustServerCertificate=True"
        
        # Get systems from database
        $systems = Get-SystemRecords -searchTerm $searchTerm -connectionString $connectionString

        # Prepare HTML response
        $escapedSearch = Encode-HTML $searchTerm
        $currentDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        # Build HTML header
        $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>System Inventory Dashboard</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" />
    <style>
        .hover-scale:hover {
            transform: scale(1.02);
            transition: transform 0.2s ease-in-out;
        }
    </style>
</head>
<body class="bg-gray-50">
    <div class="container mx-auto px-4 py-8">
        <header class="mb-8">
            <h1 class="text-3xl font-bold text-gray-800 mb-2">System Inventory Dashboard</h1>
            <p class="text-gray-600">View and manage your IT assets</p>
        </header>

        <div class="bg-white rounded-lg shadow-md p-6 mb-8">
            <form method="GET" action="/" class="flex flex-col md:flex-row gap-4">
                <div class="flex-grow relative">
                    <div class="absolute inset-y-0 left-0 flex items-center pl-3 pointer-events-none">
                        <i class="fas fa-search text-gray-400"></i>
                    </div>
                    <input type="text" name="search" placeholder="Search by Asset # or Host Name..." 
                           value="$escapedSearch" 
                           class="w-full pl-10 pr-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500" />
                </div>
                <button type="submit" 
                        class="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2">
                    <i class="fas fa-search mr-2"></i>Search
                </button>
                <a href="/" 
                   class="px-4 py-2 bg-gray-200 text-gray-800 rounded-lg hover:bg-gray-300 focus:outline-none focus:ring-2 focus:ring-gray-500 focus:ring-offset-2 text-center">
                    <i class="fas fa-times mr-2"></i>Clear
                </a>
            </form>
        </div>

        <div class="bg-white rounded-lg shadow-md overflow-hidden">
            <div class="overflow-x-auto">
                <table class="min-w-full divide-y divide-gray-200">
                    <thead class="bg-gray-50">
                        <tr>
                            <th scope="col" class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Asset #</th>
                            <th scope="col" class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Host Name</th>
                            <th scope="col" class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">OS</th>
                            <th scope="col" class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">RAM (GB)</th>
                            <th scope="col" class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Last Scanned</th>
                        </tr>
                    </thead>
                    <tbody class="bg-white divide-y divide-gray-200">
"@

        # Add table rows for each system
        foreach ($sys in $systems) {
            $assetNum = Encode-HTML $sys.AssetNumber
            $hostName = Encode-HTML $sys.HostName
            $osName = Encode-HTML $sys.OSName
            $ram = [math]::Round($sys.TotalRAMGB, 2)
            $scanDate = $sys.ScanDate.ToString("yyyy-MM-dd HH:mm")

            $html += @"
                        <tr class="hover:bg-gray-50 hover-scale">
                            <td class="px-6 py-4 whitespace-nowrap text-sm font-medium text-blue-600">$assetNum</td>
                            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900">$hostName</td>
                            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">$osName</td>
                            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">$ram</td>
                            <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">$scanDate</td>
                        </tr>
"@
        }

        # Close HTML document
        $html += @"
                    </tbody>
                </table>
            </div>
        </div>

        <footer class="mt-8 text-center text-sm text-gray-500">
            <p>Generated on $currentDate</p>
            <p class="mt-1">© $(Get-Date -Format "yyyy") IT Inventory System. All rights reserved.</p>
        </footer>
    </div>
</body>
</html>
"@

        # Send response
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
        $response.ContentLength64 = $buffer.Length
        $response.ContentType = "text/html; charset=utf-8"
        $response.OutputStream.Write($buffer, 0, $buffer.Length)
        $response.OutputStream.Close()
    }
    catch {
        Write-Warning "Error processing request: $_"
        Show-ErrorPage -response $response -errorMessage "An unexpected error occurred while processing your request."
    }
}

# Clean up
$listener.Stop()
Write-Host "Dashboard server stopped" -ForegroundColor Yellow