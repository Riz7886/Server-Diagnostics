# ULTIMATE PATCH COMPLIANCE SCANNER v3.0 - FIXED FOR MSG FILES
# Works with Outlook .msg files, HTML, TXT, CSV
# Author: Syed Ahmad
# Date: February 27, 2026

[CmdletBinding()]
param(
    [string]$ScanPath = "C:\Report-Alert",
    [string]$OutputPath = "",
    [switch]$ExportCSV = $true
)

# Setup
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $ScanPath "Reports"
}

if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  PATCH COMPLIANCE SCANNER v3.0 - FIXED" -ForegroundColor Yellow
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Scan Path: $ScanPath" -ForegroundColor White
Write-Host "Output Path: $OutputPath" -ForegroundColor White
Write-Host ""

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
        # Try using Outlook COM object
        $Outlook = New-Object -ComObject Outlook.Application
        $Msg = $Outlook.Session.OpenSharedItem($FilePath)
        $Content = $Msg.HTMLBody
        if ([string]::IsNullOrWhiteSpace($Content)) {
            $Content = $Msg.Body
        }
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
        return $Content
    } catch {
        # Fallback: read as binary and extract text
        try {
            $Bytes = [System.IO.File]::ReadAllBytes($FilePath)
            $Content = [System.Text.Encoding]::Unicode.GetString($Bytes)
            return $Content
        } catch {
            Write-Host "    Warning: Could not read MSG file: $($_.Exception.Message)" -ForegroundColor Yellow
            return ""
        }
    }
}

# Function to analyze all reports
function Analyze-Reports {
    param([string]$Path)
    
    Write-Host "Step 1: Scanning for report files..." -ForegroundColor Cyan
    Write-Host "  Looking in: $Path" -ForegroundColor Gray
    
    # Get all possible report files including MSG
    $AllFiles = @()
    
    # Search for common report file types
    $Extensions = @("*.msg", "*.html", "*.htm", "*.txt", "*.csv", "*.eml")
    
    foreach ($Ext in $Extensions) {
        $Files = Get-ChildItem -Path $Path -Filter $Ext -Recurse -ErrorAction SilentlyContinue
        if ($Files) {
            $AllFiles += $Files
            Write-Host "    Found $($Files.Count) $Ext files" -ForegroundColor Green
        }
    }
    
    Write-Host ""
    Write-Host "  Total files found: $($AllFiles.Count)" -ForegroundColor Yellow
    
    if ($AllFiles.Count -eq 0) {
        Write-Host ""
        Write-Host "ERROR: No report files found!" -ForegroundColor Red
        Write-Host ""
        Write-Host "Looking for files with these extensions:" -ForegroundColor Yellow
        Write-Host "  - .msg (Outlook messages)" -ForegroundColor Gray
        Write-Host "  - .html, .htm (HTML reports)" -ForegroundColor Gray
        Write-Host "  - .txt (Text files)" -ForegroundColor Gray
        Write-Host "  - .csv (CSV files)" -ForegroundColor Gray
        Write-Host ""
        Write-Host "Please check:" -ForegroundColor Yellow
        Write-Host "  1. Files exist in: $Path" -ForegroundColor Gray
        Write-Host "  2. You have permission to read them" -ForegroundColor Gray
        Write-Host "  3. Files are not corrupted" -ForegroundColor Gray
        Write-Host ""
        return @()
    }
    
    Write-Host ""
    Write-Host "Step 2: Extracting server data from reports..." -ForegroundColor Cyan
    
    $AllServers = @()
    $ProcessedCount = 0
    
    foreach ($File in $AllFiles) {
        $ProcessedCount++
        Write-Host "  [$ProcessedCount/$($AllFiles.Count)] Processing: $($File.Name)" -ForegroundColor Gray
        
        try {
            # Read file content based on extension
            $Content = ""
            
            if ($File.Extension -eq ".msg") {
                Write-Host "    Reading Outlook MSG file..." -ForegroundColor Cyan
                $Content = Read-MSGFile -FilePath $File.FullName
            } else {
                $Content = Get-Content $File.FullName -Raw -ErrorAction Stop
            }
            
            if ([string]::IsNullOrWhiteSpace($Content)) {
                Write-Host "    Warning: File is empty or unreadable" -ForegroundColor Yellow
                continue
            }
            
            # Extract IPs
            $IPs = Find-IPAddresses -Text $Content
            
            if ($IPs.Count -eq 0) {
                Write-Host "    No IPs found in this file" -ForegroundColor Yellow
                continue
            }
            
            Write-Host "    Found $($IPs.Count) IP addresses" -ForegroundColor Green
            
            # Process each IP
            foreach ($IP in $IPs) {
                # Get context around IP
                $IPIndex = $Content.IndexOf($IP)
                if ($IPIndex -ge 0) {
                    $ContextStart = [Math]::Max(0, $IPIndex - 500)
                    $ContextEnd = [Math]::Min($Content.Length, $IPIndex + 500)
                    $Context = $Content.Substring($ContextStart, $ContextEnd - $ContextStart)
                    
                    # Determine compliance status
                    $Status = "UNKNOWN"
                    
                    if ($Context -match "Compliant" -and $Context -notmatch "NonCompliant|Non-Compliant") {
                        $Status = "COMPLIANT"
                    } elseif ($Context -match "NonCompliant|Non-Compliant") {
                        $Status = "NON-COMPLIANT"
                    } elseif ($Context -match "Connected|OK|Pass") {
                        $Status = "COMPLIANT"
                    } elseif ($Context -match "Missing|Failed|Error") {
                        $Status = "NON-COMPLIANT"
                    }
                    
                    # Extract server name if possible
                    $ServerName = "Unknown"
                    if ($Context -match "(EC2AMAZ-[A-Z0-9]+|[a-z]\d{3}app\d{2}[a-z]{3}[^\s<>,]*)") {
                        $ServerName = $Matches[1]
                    }
                    
                    $AllServers += [PSCustomObject]@{
                        IP = $IP
                        ServerName = $ServerName
                        Status = $Status
                        SourceReport = $File.Name
                        LastChecked = $File.LastWriteTime
                    }
                }
            }
            
        } catch {
            Write-Host "    ERROR: $_" -ForegroundColor Red
        }
    }
    
    # Remove duplicates - keep most recent
    $UniqueServers = $AllServers | Group-Object IP | ForEach-Object {
        $_.Group | Sort-Object LastChecked -Descending | Select-Object -First 1
    }
    
    Write-Host ""
    Write-Host "  Extracted data for $($UniqueServers.Count) unique servers" -ForegroundColor Green
    Write-Host ""
    
    return $UniqueServers
}

# Main execution
$ScannedServers = Analyze-Reports -Path $ScanPath

if ($ScannedServers.Count -eq 0) {
    Write-Host "ERROR: No servers found in any reports!" -ForegroundColor Red
    Write-Host ""
    Write-Host "This could mean:" -ForegroundColor Yellow
    Write-Host "  1. Report files contain no IP addresses" -ForegroundColor Gray
    Write-Host "  2. Files are in wrong format" -ForegroundColor Gray
    Write-Host "  3. Files are corrupted or encrypted" -ForegroundColor Gray
    Write-Host ""
    exit 1
}

# Calculate statistics
$TotalServers = $ScannedServers.Count
$Compliant = ($ScannedServers | Where-Object { $_.Status -eq "COMPLIANT" }).Count
$NonCompliant = ($ScannedServers | Where-Object { $_.Status -eq "NON-COMPLIANT" }).Count
$Unknown = $TotalServers - $Compliant - $NonCompliant

Write-Host "Step 3: Generating reports..." -ForegroundColor Cyan
Write-Host ""

$CompliancePercent = if ($TotalServers -gt 0) { [math]::Round(($Compliant / $TotalServers) * 100, 1) } else { 0 }

# Generate HTML
$HTML = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Patch Compliance Report - $(Get-Date -Format 'MMMM dd, yyyy')</title>
    <style>
        body { font-family: Arial, sans-serif; background: #f0f0f0; padding: 20px; margin: 0; }
        .container { max-width: 1400px; margin: 0 auto; background: white; border-radius: 10px; box-shadow: 0 0 20px rgba(0,0,0,0.1); }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 40px; text-align: center; border-radius: 10px 10px 0 0; }
        .header h1 { margin: 0 0 10px 0; font-size: 36px; }
        .header p { margin: 5px 0; opacity: 0.9; }
        .stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; padding: 30px; background: #f8f9fa; }
        .stat-card { background: white; padding: 25px; border-radius: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); text-align: center; }
        .stat-label { font-size: 14px; color: #666; margin-bottom: 10px; text-transform: uppercase; letter-spacing: 1px; }
        .stat-value { font-size: 42px; font-weight: bold; color: #667eea; }
        .stat-card.success .stat-value { color: #10b981; }
        .stat-card.danger .stat-value { color: #ef4444; }
        .stat-card.warning .stat-value { color: #f59e0b; }
        .content { padding: 30px; }
        .section-title { font-size: 24px; color: #1f2937; margin: 20px 0; padding-bottom: 10px; border-bottom: 2px solid #667eea; }
        .search-box { width: 100%; max-width: 400px; padding: 12px; border: 2px solid #d1d5db; border-radius: 8px; font-size: 14px; margin-bottom: 20px; }
        .search-box:focus { outline: none; border-color: #667eea; }
        .filter-btns { margin-bottom: 20px; }
        .filter-btn { padding: 10px 20px; margin-right: 10px; border: 2px solid #667eea; background: white; color: #667eea; border-radius: 8px; cursor: pointer; font-weight: bold; }
        .filter-btn:hover { background: #667eea; color: white; }
        .filter-btn.active { background: #667eea; color: white; }
        table { width: 100%; border-collapse: collapse; margin: 20px 0; }
        th { background: #667eea; color: white; padding: 15px; text-align: left; font-weight: bold; }
        td { padding: 12px 15px; border-bottom: 1px solid #e5e7eb; }
        tr:hover { background: #f9fafb; }
        .badge { padding: 6px 14px; border-radius: 20px; font-weight: bold; font-size: 12px; display: inline-block; }
        .badge.success { background: #d1fae5; color: #065f46; border: 1px solid #6ee7b7; }
        .badge.danger { background: #fee2e2; color: #991b1b; border: 1px solid #fca5a5; }
        .badge.warning { background: #fef3c7; color: #92400e; border: 1px solid #fcd34d; }
        .footer { background: #f9fafb; padding: 20px; text-align: center; color: #6b7280; border-top: 1px solid #e5e7eb; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Patch Compliance Report</h1>
            <p>Generated: $(Get-Date -Format 'MMMM dd, yyyy - hh:mm:ss tt')</p>
            <p>Scan Path: $ScanPath</p>
        </div>
        
        <div class="stats">
            <div class="stat-card">
                <div class="stat-label">Total Servers</div>
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
                <div class="stat-label">Unknown</div>
                <div class="stat-value">$Unknown</div>
            </div>
        </div>
        
        <div class="content">
            <div class="section-title">Server Details ($TotalServers servers found)</div>
            
            <input type="text" class="search-box" id="searchBox" placeholder="Search by IP or Server Name..." onkeyup="filterTable()">
            
            <div class="filter-btns">
                <button class="filter-btn active" onclick="filterStatus('all')">All Servers</button>
                <button class="filter-btn" onclick="filterStatus('COMPLIANT')">Compliant Only</button>
                <button class="filter-btn" onclick="filterStatus('NON-COMPLIANT')">Non-Compliant</button>
                <button class="filter-btn" onclick="filterStatus('UNKNOWN')">Unknown</button>
            </div>
            
            <table id="serverTable">
                <thead>
                    <tr>
                        <th>IP Address</th>
                        <th>Server Name</th>
                        <th>Status</th>
                        <th>Source Report</th>
                        <th>Last Checked</th>
                    </tr>
                </thead>
                <tbody>
"@

foreach ($Server in $ScannedServers) {
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
                        <td>$($Server.SourceReport)</td>
                        <td>$($Server.LastChecked)</td>
                    </tr>
"@
}

$HTML += @"
                </tbody>
            </table>
        </div>
        
        <div class="footer">
            <p><strong>Patch Compliance Scanner v3.0</strong></p>
            <p>Fixed for MSG files | Author: Syed Ahmad</p>
        </div>
    </div>
    
    <script>
        function filterStatus(status) {
            const rows = document.querySelectorAll('.data-row');
            const buttons = document.querySelectorAll('.filter-btn');
            
            buttons.forEach(btn => btn.classList.remove('active'));
            event.target.classList.add('active');
            
            rows.forEach(row => {
                const rowStatus = row.getAttribute('data-status');
                row.style.display = (status === 'all' || rowStatus === status) ? '' : 'none';
            });
        }
        
        function filterTable() {
            const input = document.getElementById('searchBox');
            const filter = input.value.toUpperCase();
            const table = document.getElementById('serverTable');
            const rows = table.getElementsByTagName('tr');
            
            for (let i = 1; i < rows.length; i++) {
                const text = rows[i].textContent || rows[i].innerText;
                rows[i].style.display = text.toUpperCase().indexOf(filter) > -1 ? '' : 'none';
            }
        }
    </script>
</body>
</html>
"@

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$HTMLFile = Join-Path $OutputPath "PatchCompliance_$Timestamp.html"
$HTML | Out-File -FilePath $HTMLFile -Encoding UTF8 -Force

Write-Host "  HTML Report: $HTMLFile" -ForegroundColor Green

# Export CSV
if ($ExportCSV) {
    $CSVFile = Join-Path $OutputPath "PatchCompliance_$Timestamp.csv"
    $ScannedServers | Export-Csv -Path $CSVFile -NoTypeInformation -Encoding UTF8
    Write-Host "  CSV Export: $CSVFile" -ForegroundColor Green
}

# Summary
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  SCAN COMPLETE!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Statistics:" -ForegroundColor Yellow
Write-Host "  Total Servers: $TotalServers" -ForegroundColor White
Write-Host "  Compliant: $Compliant ($CompliancePercent%)" -ForegroundColor Green
Write-Host "  Non-Compliant: $NonCompliant" -ForegroundColor Red
Write-Host "  Unknown: $Unknown" -ForegroundColor Yellow
Write-Host ""
Write-Host "Reports saved to: $OutputPath" -ForegroundColor Cyan
Write-Host ""

# Open report
Write-Host "Opening HTML report in browser..." -ForegroundColor Yellow
Start-Process $HTMLFile

Write-Host ""
Write-Host "DONE! Check your report!" -ForegroundColor Green
Write-Host ""
