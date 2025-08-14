# **System Inventory Management Suite - Walkthrough**

This guide explains how to use the three PowerShell scripts that make up the **North West Provincial Treasury's System Inventory Management Suite**:

1. **SystemInfoCollector.ps1** - Collects system information
2. **AVSystemInfoCollector-1.ps1** - Windows 7-compatible version with auto-upgrade
3. **ImportSystemInventory.ps1** - Imports data to SQL Server
4. **SystemDashboard.ps1** - Web-based dashboard for viewing inventory

---

## **1. System Information Collection**

### **Option A: For Modern Systems (Windows 8/10/11)**
📜 **File:** `SystemInfoCollector.ps1`

**What it does:**
- Collects comprehensive system data (hardware, software, network)
- Saves as JSON to `Desktop\SystemReports`
- Works on PowerShell 3.0+ (default on Windows 8+)

**How to run:**
```powershell
.\SystemInfoCollector.ps1
```
1. When prompted, enter the **11-digit Asset Number** (starting with 0)
2. Script saves a JSON file like:  
   `C:\Users\[You]\Desktop\SystemReports\COMPUTERNAME-20230814-143022.json`

---

### **Option B: For Windows 7 Systems**
📜 **File:** `AVSystemInfoCollector-1.ps1`

**What it does:**
- **Auto-upgrades PowerShell 2.0 → 3.0** if needed
- Installs .NET 4.0 automatically (required for PS 3.0)
- Falls back to PS 2.0 mode if upgrade fails
- Same output format as `SystemInfoCollector.ps1`

**How to run:**
```powershell
.\AVSystemInfoCollector-1.ps1
```
1. If on PS 2.0, it will prompt to upgrade (recommended)
2. Enter the **11-digit Asset Number** when asked
3. JSON file saved to `Desktop\SystemReports`

---

## **2. Importing Data to SQL Server**
📜 **File:** `ImportSystemInventory.ps1`

**What it does:**
- Imports JSON reports into SQL Server (`AssetDB`)
- Updates existing records or creates new ones
- Handles:
  - System info (hostname, BIOS, etc.)
  - Hardware specs (CPU, RAM, disks)
  - Installed software

**Prerequisites:**
- SQL Server with `AssetDB` database (schema must exist)
- Permissions to write to the database

**How to run:**
```powershell
# Default (uses Desktop\SystemReports and local SQL Server)
.\ImportSystemInventory.ps1

# Custom folder/SQL Server:
.\ImportSystemInventory.ps1 -ReportsFolder "C:\Inventory" -ConnectionString "Server=SQL01;Database=AssetDB;Integrated Security=True"
```

**Process:**
1. Scans `SystemReports` folder for JSON files
2. For each file:
   - Checks if system exists (by AssetNumber/UUID)
   - Updates or inserts records in these tables:
     - `Systems` (basic info)
     - `SystemSpecs` (hardware details)
     - `SystemDisks` (storage)
     - `InstalledApps` (software)

---

## **3. Viewing the Dashboard**
📜 **File:** `SystemDashboard.ps1`

**What it does:**
- Starts a **local web server** on port `8080`
- Shows paginated system inventory
- Search by hostname/asset number
- Detailed system views

**How to run:**
```powershell
.\SystemDashboard.ps1
```
1. Open browser to:  
   `http://localhost:8080`
2. Features:
   - 🔍 **Search** systems
   - 📊 **Filterable tables**
   - ℹ️ **Detailed views** (click "Details")
   - ⏱️ **Auto-refreshes** every 5 minutes

---

## **Workflow Summary**
1. **Collect Data**  
   - Run `SystemInfoCollector.ps1` (modern OS) or `AVSystemInfoCollector-1.ps1` (Win7)
2. **Import to SQL**  
   - Run `ImportSystemInventory.ps1` periodically
3. **View Dashboard**  
   - Run `SystemDashboard.ps1` to monitor inventory

---

## **Troubleshooting**
- **PowerShell errors**: Run as Administrator
- **SQL connection issues**: Verify `ConnectionString` in `ImportSystemInventory.ps1`
- **Win7 upgrade fails**: Manually install .NET 4.0 first
- **Dashboard not loading**: Check firewall allows port `8080`

---

## **Security Notes**
- 🔒 All scripts use **least-privilege principles**
- 🔐 SQL connection uses **Windows Integrated Auth**
- 🛡️ HTML output is **XSS-protected** (encoded)
- 📁 JSON files contain **asset numbers** - store securely


