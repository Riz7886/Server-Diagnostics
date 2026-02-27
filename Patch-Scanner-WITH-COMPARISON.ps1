# PATCH COMPLIANCE SCANNER v4.0 - WITH LIST COMPARISON
# Compares YOUR server list against reports
# Author: Syed Ahmad
# Date: February 27, 2026

[CmdletBinding()]
param(
    [string]$ScanPath = "C:\Reports-alerts",
    [string]$ServerListFile = "",
    [string]$OutputPath = "",
    [switch]$ExportCSV = $true
)

# Setup
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path (Split-Path $ScanPath -Parent) "Reports"
}

if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  PATCH SCANNER v4.0 - WITH COMPARISON" -ForegroundColor Yellow
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Scan Path: $ScanPath" -ForegroundColor White
Write-Host "Output Path: $OutputPath" -ForegroundColor White
Write-Host ""

# Function to read Excel
function Read-ExcelFile {
    param([string]$FilePath)
    
    Write-Host "Reading server list from Excel..." -ForegroundColor Cyan
    Write-Host "  File: $([System.IO.Path]::GetFileName($FilePath))" -ForegroundColor Gray
    
    try {
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false
        
        $Workbook = $Excel.Workbooks.Open($FilePath)
        $Worksheet = $Workbook.Sheets.Item(1)
        
        $Range = $Worksheet.UsedRange
        $RowCount = $Range.Rows.Count
        $ColCount = $Range.Columns.Count
        
        $Headers = @()
        for ($col = 1; $col -le $ColCount; $col++) {
            $Headers += $Worksheet.Cells.Item(1, $col).Text
        }
        
        Write-Host "  Columns found: $($Headers -join ', ')" -ForegroundColor Gray
        
        $Data = @()
        for ($row = 2; $row -le $RowCount; $row++) {
            $RowData = @{}
            $HasData = $false
            
            for ($col = 1; $col -le $Headers.Count; $col++) {
                $Value = $Worksheet.Cells.Item($row, $col).Text
                $RowData[$Headers[$col - 1]] = $Value
                if (-not [string]::IsNullOrWhiteSpace($Value)) {
                    $HasData = $true
                }
            }
            
            if ($HasData) {
                $Data += [PSCustomObject]$RowData
            }
        }
        
        $Workbook.Close($false)
        $Excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
        
        Write-Host "  Loaded $($Data.Count) servers from list" -ForegroundColor Green
        return $Data
        
    } catch {
        Write-Host "  ERROR: $_" -ForegroundColor Red
        return @()
    }
}

# Function to find IPs
function Find-IPAddresses {
    param([string]$Text)
    
    $IPPattern = '\b(?:\d{1,3}\.){3}\d{1,3}\b'
    $Matches = [regex]::Matches($Text, $IPPattern)
    return $Matches | ForEach-Object { $_.Value } | Select-Object -Unique
}

# Function to read MSG files
function Read-MSGFile {
    param([string]$FilePath)
    
    try {
        $Outlook = New-Object -ComObject Outlook.Application
        $Msg = $Outlook.Session.OpenSharedItem($FilePath)
        $Content = $Msg.HTMLBody
        if ([string]::IsNullOrWhiteSpace($Content)) {
            $Content = $Msg.Body
        }
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
        return $Content
    } catch {
        try {
            $Bytes = [System.IO.File]::ReadAllBytes($FilePath)
            $Content = [System.Text.Encoding]::Unicode.GetString($Bytes)
            return $Content
        } catch {
            return ""
        }
    }
}

# Function to scan reports
function Scan-Reports {
    param([string]$Path)
    
    Write-Host "Step 1: Scanning compliance reports..." -ForegroundColor Cyan
    
    $AllFiles = @()
    $Extensions = @("*.msg", "*.html", "*.htm", "*.txt", "*.csv")
    
    foreach ($Ext in $Extensions) {
        $Files = Get-ChildItem -Path $Path -Filter $Ext -Recurse -ErrorAction SilentlyContinue
        if ($Files) {
            $AllFiles += $Files
            Write-Host "  Found $($Files.Count) $Ext files" -ForegroundColor Green
        }
    }
    
    if ($AllFiles.Count -eq 0) {
        Write-Host "  ERROR: No report files found!" -ForegroundColor Red
        return @()
    }
    
    Write-Host "  Total report files: $($AllFiles.Count)" -ForegroundColor Yellow
    Write-Host ""
    
    $ReportData = @{}
    
    foreach ($File in $AllFiles) {
        Write-Host "  Processing: $($File.Name)" -ForegroundColor Gray
        
        try {
            $Content = ""
            
            if ($File.Extension -eq ".msg") {
                $Content = Read-MSGFile -FilePath $File.FullName
            } else {
                $Content = Get-Content $File.FullName -Raw -ErrorAction Stop
            }
            
            if ([string]::IsNullOrWhiteSpace($Content)) {
                continue
            }
            
            $IPs = Find-IPAddresses -Text $Content
            
            foreach ($IP in $IPs) {
                $IPIndex = $Content.IndexOf($IP)
                if ($IPIndex -ge 0) {
                    $ContextStart = [Math]::Max(0, $IPIndex - 1000)
                    $ContextEnd = [Math]::Min($Content.Length, $IPIndex + 1000)
                    $Context = $Content.Substring($ContextStart, $ContextEnd - $ContextStart)
                    
                    $Status = "UNKNOWN"
                    $Details = ""
                    
                    # Check for Compliant
                    if ($Context -match "Trend\s*Micro[^\w]*Compliant" -or 
                        $Context -match "Trellix[^\w]*Compliant" -or
                        $Context -match "CrowdStrike[^\w]*Compliant") {
                        $Status = "COMPLIANT"
                    }
                    
                    # Check for Non-Compliant
                    if ($Context -match "NonCompliant|Non-Compliant|Missing|Failed") {
                        $Status = "NON-COMPLIANT"
                        
                        # Extract details
                        if ($Context -match "Trend\s*Micro[^\w]*(NonCompliant|Non-Compliant)") {
                            $Details += "Trend Micro: Non-Compliant; "
                        }
                        if ($Context -match "Trellix[^\w]*(NonCompliant|Non-Compliant)") {
                            $Details += "Trellix: Non-Compliant; "
                        }
                        if ($Context -match "CrowdStrike[^\w]*(NonCompliant|Non-Compliant)") {
                            $Details += "CrowdStrike: Non-Compliant; "
                        }
                        if ($Context -match "CloudWatch[^\w]*(NonCompliant|Non-Compliant)") {
                            $Details += "CloudWatch: Non-Compliant; "
                        }
                    }
                    
                    # Extract server name
                    $ServerName = "Unknown"
                    if ($Context -match "(EC2AMAZ-[A-Z0-9]+|[a-z]\d{3}app\d{2}[a-z]{3}[^\s<>,]*)") {
                        $ServerName = $Matches[1]
                    }
                    
                    if (-not $ReportData.ContainsKey($IP)) {
                        $ReportData[$IP] = @{
                            IP = $IP
                            ServerName = $ServerName
                            Status = $Status
                            Details = $Details
                            SourceReport = $File.Name
                            LastChecked = $File.LastWriteTime
                        }
                    }
                }
            }
            
        } catch {
            Write-Host "    ERROR: $_" -ForegroundColor Red
        }
    }
    
    Write-Host ""
    Write-Host "  Found data for $($ReportData.Count) unique servers in reports" -ForegroundColor Green
    Write-Host ""
    
    return $ReportData
}

# Main execution
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "STEP 1: LOAD YOUR SERVER LIST" -ForegroundColor Yellow
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# Find server list
if ([string]::IsNullOrWhiteSpace($ServerListFile)) {
    $ParentPath = Split-Path $ScanPath -Parent
    $PossibleFiles = @(
        (Join-Path $ParentPath "Rick-patch-List.xlsx"),
        (Join-Path $ScanPath "Rick-patch-List.xlsx"),
        (Get-ChildItem -Path $ParentPath -Filter "*patch*list*.xlsx" -ErrorAction SilentlyContinue | Select-Object -First 1).FullName,
        (Get-ChildItem -Path $ParentPath -Filter "*server*list*.xlsx" -ErrorAction SilentlyContinue | Select-Object -First 1).FullName
    )
    
    foreach ($File in $PossibleFiles) {
        if ($File -and (Test-Path $File)) {
            $ServerListFile = $File
            break
        }
    }
}

if (-not $ServerListFile -or -not (Test-Path $ServerListFile)) {
    Write-Host "ERROR: Server list not found!" -ForegroundColor Red
    Write-Host "Please specify: -ServerListFile 'C:\path\to\Rick-patch-List.xlsx'" -ForegroundColor Yellow
    exit 1
}

$ServerList = Read-ExcelFile -FilePath $ServerListFile

if ($ServerList.Count -eq 0) {
    Write-Host "ERROR: No servers in list!" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "STEP 2: SCAN COMPLIANCE REPORTS" -ForegroundColor Yellow
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

$ReportData = Scan-Reports -Path $ScanPath

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "STEP 3: COMPARE YOUR LIST VS REPORTS" -ForegroundColor Yellow
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

$ComparisonResults = @()

foreach ($Server in $ServerList) {
    # Get IP from list (try different column names)
    $IP = $null
    foreach ($Prop in $Server.PSObject.Properties.Name) {
        $Value = $Server.$Prop
        if ($Value -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
            $IP = $Value
            break
        }
    }
    
    if ([string]::IsNullOrWhiteSpace($IP)) {
        continue
    }
    
    # Check if IP is in reports
    if ($ReportData.ContainsKey($IP)) {
        $Found = $ReportData[$IP]
        
        $ComparisonResults += [PSCustomObject]@{
            IP = $IP
            ServerName = $Found.ServerName
            Status = $Found.Status
            Details = $Found.Details
            SourceReport = $Found.SourceReport
            LastChecked = $Found.LastChecked
            InYourList = "YES"
        }
    } else {
        $ComparisonResults += [PSCustomObject]@{
            IP = $IP
            ServerName = "Not Found"
            Status = "NOT IN REPORTS"
            Details = "This server from your list was not found in any compliance report"
            SourceReport = "N/A"
            LastChecked = "N/A"
            InYourList = "YES"
        }
    }
}

# Calculate statistics
$TotalServers = $ComparisonResults.Count
$Compliant = ($ComparisonResults | Where-Object { $_.Status -eq "COMPLIANT" }).Count
$NonCompliant = ($ComparisonResults | Where-Object { $_.Status -eq "NON-COMPLIANT" }).Count
$NotInReports = ($ComparisonResults | Where-Object { $_.Status -eq "NOT IN REPORTS" }).Count
$Unknown = $TotalServers - $Compliant - $NonCompliant - $NotInReports

Write-Host "Comparison complete!" -ForegroundColor Green
Write-Host "  From your list: $TotalServers servers" -ForegroundColor White
Write-Host "  Found Compliant: $Compliant" -ForegroundColor Green
Write-Host "  Found Non-Compliant: $NonCompliant" -ForegroundColor Red
Write-Host "  Not in Reports: $NotInReports" -ForegroundColor Yellow
Write-Host "  Unknown Status: $Unknown" -ForegroundColor Gray
Write-Host ""

# Generate HTML
Write-Host "Generating comparison report..." -ForegroundColor Cyan

$CompliancePercent = if ($TotalServers -gt 0) { [math]::Round(($Compliant / $TotalServers) * 100, 1) } else { 0 }

$HTML = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Patch Compliance Comparison Report</title>
    <style>
        body { font-family: Arial; background: #f0f0f0; padding: 20px; margin: 0; }
        .container { max-width: 1600px; margin: 0 auto; background: white; border-radius: 10px; box-shadow: 0 0 30px rgba(0,0,0,0.2); }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 40px; text-align: center; border-radius: 10px 10px 0 0; }
        .header h1 { margin: 0 0 10px 0; font-size: 36px; }
        .header p { margin: 5px 0; opacity: 0.95; }
        .alert { background: #fef3c7; border-left: 4px solid #f59e0b; padding: 15px; margin: 20px; border-radius: 5px; }
        .alert-success { background: #d1fae5; border-left-color: #10b981; }
        .alert-danger { background: #fee2e2; border-left-color: #ef4444; }
        .stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; padding: 30px; background: #f8f9fa; }
        .stat-card { background: white; padding: 25px; border-radius: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); text-align: center; }
        .stat-label { font-size: 14px; color: #666; margin-bottom: 10px; text-transform: uppercase; }
        .stat-value { font-size: 42px; font-weight: bold; color: #667eea; }
        .stat-card.success .stat-value { color: #10b981; }
        .stat-card.danger .stat-value { color: #ef4444; }
        .stat-card.warning .stat-value { color: #f59e0b; }
        .content { padding: 30px; }
        .section-title { font-size: 24px; color: #1f2937; margin: 20px 0; padding-bottom: 10px; border-bottom: 2px solid #667eea; }
        .search-box { width: 100%; max-width: 400px; padding: 12px; border: 2px solid #d1d5db; border-radius: 8px; margin-bottom: 15px; }
        .filter-btns { margin-bottom: 20px; }
        .filter-btn { padding: 10px 20px; margin-right: 10px; margin-bottom: 10px; border: 2px solid #667eea; background: white; color: #667eea; border-radius: 8px; cursor: pointer; font-weight: bold; }
        .filter-btn:hover { background: #667eea; color: white; }
        .filter-btn.active { background: #667eea; color: white; }
        table { width: 100%; border-collapse: collapse; margin: 20px 0; }
        th { background: #667eea; color: white; padding: 15px; text-align: left; font-weight: bold; position: sticky; top: 0; }
        td { padding: 12px 15px; border-bottom: 1px solid #e5e7eb; }
        tr:hover { background: #f9fafb; }
        .badge { padding: 6px 14px; border-radius: 20px; font-weight: bold; font-size: 12px; display: inline-block; }
        .badge.success { background: #d1fae5; color: #065f46; }
        .badge.danger { background: #fee2e2; color: #991b1b; }
        .badge.warning { background: #fef3c7; color: #92400e; }
        .details { font-size: 12px; color: #666; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Patch Compliance Comparison Report</h1>
            <p>YOUR SERVER LIST vs COMPLIANCE REPORTS</p>
            <p>Generated: $(Get-Date -Format 'MMMM dd, yyyy - hh:mm:ss tt')</p>
        </div>
        
        <div class="stats">
            <div class="stat-card">
                <div class="stat-label">Your Server List</div>
                <div class="stat-value">$TotalServers</div>
            </div>
            <div class="stat-card success">
                <div class="stat-label">Compliant</div>
                <div class="stat-value">$Compliant</div>
            </div>
            <div class="stat-card danger">
                <div class="stat-label">Non-Compliant</div>
                <div class="stat-value">$NonCompliant</div>
            </div>
            <div class="stat-card warning">
                <div class="stat-label">Not in Reports</div>
                <div class="stat-value">$NotInReports</div>
            </div>
        </div>
        
        <div class="content">
            <div class="section-title">Comparison Results - Servers from YOUR List</div>
            
            <input type="text" class="search-box" id="searchBox" placeholder="Search by IP..." onkeyup="filterTable()">
            
            <div class="filter-btns">
                <button class="filter-btn active" onclick="filterStatus('all')">All ($TotalServers)</button>
                <button class="filter-btn" onclick="filterStatus('COMPLIANT')">Compliant ($Compliant)</button>
                <button class="filter-btn" onclick="filterStatus('NON-COMPLIANT')">Non-Compliant ($NonCompliant)</button>
                <button class="filter-btn" onclick="filterStatus('NOT IN REPORTS')">Not in Reports ($NotInReports)</button>
            </div>
            
            <table id="serverTable">
                <thead>
                    <tr>
                        <th>IP Address</th>
                        <th>Server Name</th>
                        <th>Status</th>
                        <th>Details / Issues</th>
                        <th>Source Report</th>
                        <th>Last Checked</th>
                    </tr>
                </thead>
                <tbody>
"@

foreach ($Server in $ComparisonResults) {
    $BadgeClass = switch ($Server.Status) {
        "COMPLIANT" { "success" }
        "NON-COMPLIANT" { "danger" }
        default { "warning" }
    }
    
    $HTML += @"
                    <tr class="data-row" data-status="$($Server.Status)">
                        <td><strong>$($Server.IP)</strong></td>
                        <td>$($Server.ServerName)</td>
                        <td><span class="badge $BadgeClass">$($Server.Status)</span></td>
                        <td class="details">$($Server.Details)</td>
                        <td>$($Server.SourceReport)</td>
                        <td>$($Server.LastChecked)</td>
                    </tr>
"@
}

$HTML += @"
                </tbody>
            </table>
        </div>
    </div>
    
    <script>
        function filterStatus(status) {
            const rows = document.querySelectorAll('.data-row');
            const buttons = document.querySelectorAll('.filter-btn');
            
            buttons.forEach(btn => btn.classList.remove('active'));
            event.target.classList.add('active');
            
            rows.forEach(row => {
                row.style.display = (status === 'all' || row.getAttribute('data-status') === status) ? '' : 'none';
            });
        }
        
        function filterTable() {
            const filter = document.getElementById('searchBox').value.toUpperCase();
            const rows = document.querySelectorAll('.data-row');
            
            rows.forEach(row => {
                const text = row.textContent || row.innerText;
                row.style.display = text.toUpperCase().indexOf(filter) > -1 ? '' : 'none';
            });
        }
    </script>
</body>
</html>
"@

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$HTMLFile = Join-Path $OutputPath "PatchComparison_$Timestamp.html"
$HTML | Out-File -FilePath $HTMLFile -Encoding UTF8 -Force

if ($ExportCSV) {
    $CSVFile = Join-Path $OutputPath "PatchComparison_$Timestamp.csv"
    $ComparisonResults | Export-Csv -Path $CSVFile -NoTypeInformation -Encoding UTF8
    Write-Host "  CSV: $CSVFile" -ForegroundColor Green
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  COMPARISON COMPLETE!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Your Server List: $TotalServers servers" -ForegroundColor White
Write-Host "  Compliant: $Compliant ($CompliancePercent%)" -ForegroundColor Green
Write-Host "  Non-Compliant: $NonCompliant" -ForegroundColor Red  
Write-Host "  Not in Reports: $NotInReports" -ForegroundColor Yellow
Write-Host ""
Write-Host "Report: $HTMLFile" -ForegroundColor Cyan
Write-Host ""

Start-Process $HTMLFile

Write-Host "DONE! Your comparison report is ready!" -ForegroundColor Green
