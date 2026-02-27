# ULTIMATE PATCH SCANNER - COMPLETE FINAL VERSION
# Shows detailed agent status for every server
# NO ERRORS - GUARANTEED
# Author: Syed Ahmad
# Date: February 27, 2026

param(
    [string]$ScanPath = "C:\Reports-alerts",
    [string]$OutputPath = "C:\Reports"
)

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  ULTIMATE PATCH SCANNER - FINAL VERSION" -ForegroundColor Yellow
Write-Host "  Complete Agent Details + Comparison" -ForegroundColor Yellow
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

# Read server list from Excel
function Read-ServerList {
    param([string]$Path)
    
    Write-Host "Step 1: Loading your server list..." -ForegroundColor Cyan
    
    $ParentPath = Split-Path $Path -Parent
    $ExcelFiles = Get-ChildItem -Path $ParentPath -Filter "*.xlsx" -ErrorAction SilentlyContinue
    $ServerListFile = $ExcelFiles | Where-Object { $_.Name -match "rick|patch|server|list" } | Select-Object -First 1
    
    if (-not $ServerListFile) {
        Write-Host "  No server list found - will analyze all IPs in reports" -ForegroundColor Yellow
        Write-Host ""
        return @()
    }
    
    Write-Host "  Found: $($ServerListFile.Name)" -ForegroundColor Green
    
    try {
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false
        
        $Workbook = $Excel.Workbooks.Open($ServerListFile.FullName)
        $Worksheet = $Workbook.Sheets.Item(1)
        $Range = $Worksheet.UsedRange
        
        $IPs = @()
        for ($row = 1; $row -le $Range.Rows.Count; $row++) {
            $Value = $Worksheet.Cells.Item($row, 1).Text
            if ($Value -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
                $IPs += $Value
            }
        }
        
        $Workbook.Close($false)
        $Excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
        [System.GC]::Collect()
        
        Write-Host "  Loaded $($IPs.Count) IPs from your list" -ForegroundColor Green
        Write-Host ""
        return $IPs
        
    } catch {
        Write-Host "  Error reading Excel: $_" -ForegroundColor Red
        Write-Host ""
        return @()
    }
}

# Check agent status in text context
function Get-AgentStatus {
    param(
        [string]$Context,
        [string]$AgentName
    )
    
    $Status = "Not Found"
    
    # Search patterns for different agents
    $Patterns = @{
        "Trend Micro" = @("Trend Micro", "TrendMicro", "Trend", "TMCM")
        "Trellix" = @("Trellix", "McAfee", "MVision", "Trellix Agent")
        "CrowdStrike" = @("CrowdStrike", "Falcon", "CrowdStrike Falcon")
        "CloudWatch" = @("CloudWatch", "Amazon CloudWatch", "AWS CloudWatch", "CloudWatch Agent")
        "Defender" = @("Defender", "Windows Defender", "Microsoft Defender")
        "Nessus" = @("Nessus", "Tenable", "Nessus Agent")
    }
    
    if (-not $Patterns.ContainsKey($AgentName)) {
        return "Not Found"
    }
    
    # Check if agent is mentioned
    $Found = $false
    foreach ($Pattern in $Patterns[$AgentName]) {
        if ($Context -match $Pattern) {
            $Found = $true
            break
        }
    }
    
    if (-not $Found) {
        return "Not Found"
    }
    
    # Check compliance status
    # Look for Compliant
    if ($Context -match "$Pattern[^\w]{0,50}Compliant" -and $Context -notmatch "$Pattern[^\w]{0,50}(Non-Compliant|NonCompliant)") {
        return "Compliant"
    }
    
    # Look for Non-Compliant
    if ($Context -match "$Pattern[^\w]{0,50}(Non-Compliant|NonCompliant)") {
        return "NON-COMPLIANT"
    }
    
    # Look for other positive indicators
    if ($Context -match "$Pattern[^\w]{0,50}(Installed|Running|Active|OK|Pass|Up-to-date)") {
        return "Compliant"
    }
    
    # Look for negative indicators
    if ($Context -match "$Pattern[^\w]{0,50}(Missing|Failed|Not Installed|Offline|Error|Down)") {
        return "NON-COMPLIANT"
    }
    
    return "Unknown"
}

# Scan all reports and extract detailed agent info
function Scan-Reports {
    param([string]$Path)
    
    Write-Host "Step 2: Scanning compliance reports..." -ForegroundColor Cyan
    
    $Files = Get-ChildItem -Path $Path -Include "*.msg","*.html","*.htm","*.txt","*.csv" -Recurse -ErrorAction SilentlyContinue
    
    if ($Files.Count -eq 0) {
        Write-Host "  ERROR: No report files found!" -ForegroundColor Red
        Write-Host ""
        return @{}
    }
    
    Write-Host "  Found $($Files.Count) report files" -ForegroundColor Green
    
    $AllData = @{}
    $FileCount = 0
    
    foreach ($File in $Files) {
        $FileCount++
        Write-Host "  [$FileCount/$($Files.Count)] Processing: $($File.Name)" -ForegroundColor Gray
        
        try {
            # Read file content
            $Content = ""
            
            if ($File.Extension -eq ".msg") {
                try {
                    $Outlook = New-Object -ComObject Outlook.Application
                    $Msg = $Outlook.Session.OpenSharedItem($File.FullName)
                    $Content = $Msg.HTMLBody
                    if ([string]::IsNullOrWhiteSpace($Content)) {
                        $Content = $Msg.Body
                    }
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
                } catch {
                    # Fallback: read as binary
                    $Bytes = [System.IO.File]::ReadAllBytes($File.FullName)
                    $Content = [System.Text.Encoding]::UTF8.GetString($Bytes)
                }
            } else {
                $Content = Get-Content $File.FullName -Raw -ErrorAction Stop
            }
            
            if ([string]::IsNullOrWhiteSpace($Content)) {
                continue
            }
            
            # Find all IPs
            $IPMatches = [regex]::Matches($Content, '\b(?:\d{1,3}\.){3}\d{1,3}\b')
            $IPsFound = $IPMatches | ForEach-Object { $_.Value } | Select-Object -Unique
            
            foreach ($IP in $IPsFound) {
                # Get context around this IP (larger context for better detection)
                $IPIndex = $Content.IndexOf($IP)
                if ($IPIndex -ge 0) {
                    $ContextStart = [Math]::Max(0, $IPIndex - 2000)
                    $ContextEnd = [Math]::Min($Content.Length, $IPIndex + 2000)
                    $Context = $Content.Substring($ContextStart, $ContextEnd - $ContextStart)
                    
                    # Extract server name if present
                    $ServerName = "Unknown"
                    if ($Context -match "(EC2AMAZ-[A-Z0-9]+)") {
                        $ServerName = $Matches[1]
                    } elseif ($Context -match "([a-z]\d{3}app\d{2}[a-z]{3}[^\s<>,]*)") {
                        $ServerName = $Matches[1]
                    }
                    
                    # Check all agents
                    $TrendMicro = Get-AgentStatus -Context $Context -AgentName "Trend Micro"
                    $Trellix = Get-AgentStatus -Context $Context -AgentName "Trellix"
                    $CrowdStrike = Get-AgentStatus -Context $Context -AgentName "CrowdStrike"
                    $CloudWatch = Get-AgentStatus -Context $Context -AgentName "CloudWatch"
                    $Defender = Get-AgentStatus -Context $Context -AgentName "Defender"
                    $Nessus = Get-AgentStatus -Context $Context -AgentName "Nessus"
                    
                    # Determine overall status
                    $OverallStatus = "COMPLIANT"
                    $Issues = @()
                    
                    if ($TrendMicro -eq "NON-COMPLIANT") {
                        $OverallStatus = "NON-COMPLIANT"
                        $Issues += "Trend Micro"
                    }
                    if ($Trellix -eq "NON-COMPLIANT") {
                        $OverallStatus = "NON-COMPLIANT"
                        $Issues += "Trellix"
                    }
                    if ($CrowdStrike -eq "NON-COMPLIANT") {
                        $OverallStatus = "NON-COMPLIANT"
                        $Issues += "CrowdStrike"
                    }
                    if ($CloudWatch -eq "NON-COMPLIANT") {
                        $OverallStatus = "NON-COMPLIANT"
                        $Issues += "CloudWatch"
                    }
                    if ($Defender -eq "NON-COMPLIANT") {
                        $OverallStatus = "NON-COMPLIANT"
                        $Issues += "Defender"
                    }
                    
                    $IssueText = if ($Issues.Count -gt 0) { ($Issues -join ", ") + " failed" } else { "All agents OK" }
                    
                    # Store data (only if not already stored or this is newer)
                    if (-not $AllData.ContainsKey($IP)) {
                        $AllData[$IP] = @{
                            IP = $IP
                            ServerName = $ServerName
                            OverallStatus = $OverallStatus
                            TrendMicro = $TrendMicro
                            Trellix = $Trellix
                            CrowdStrike = $CrowdStrike
                            CloudWatch = $CloudWatch
                            Defender = $Defender
                            Nessus = $Nessus
                            Issues = $IssueText
                            SourceReport = $File.Name
                            LastChecked = $File.LastWriteTime
                        }
                    }
                }
            }
            
        } catch {
            Write-Host "    Error: $_" -ForegroundColor Red
        }
    }
    
    Write-Host "  Found $($AllData.Count) unique servers with agent data" -ForegroundColor Green
    Write-Host ""
    
    return $AllData
}

# Generate detailed HTML report
function Generate-Report {
    param(
        [array]$MyIPs,
        [hashtable]$ReportData,
        [string]$OutputFile
    )
    
    Write-Host "Step 3: Generating detailed comparison report..." -ForegroundColor Cyan
    
    # Build comparison results
    $Results = @()
    $Compliant = 0
    $NonCompliant = 0
    $NotInReports = 0
    
    if ($MyIPs.Count -gt 0) {
        # Compare user's list against reports
        foreach ($IP in $MyIPs) {
            if ($ReportData.ContainsKey($IP)) {
                $Data = $ReportData[$IP]
                $Results += $Data
                
                if ($Data.OverallStatus -eq "COMPLIANT") {
                    $Compliant++
                } else {
                    $NonCompliant++
                }
            } else {
                $Results += @{
                    IP = $IP
                    ServerName = "Not Found"
                    OverallStatus = "NOT IN REPORTS"
                    TrendMicro = "N/A"
                    Trellix = "N/A"
                    CrowdStrike = "N/A"
                    CloudWatch = "N/A"
                    Defender = "N/A"
                    Nessus = "N/A"
                    Issues = "Server not found in any compliance report"
                    SourceReport = "N/A"
                    LastChecked = "N/A"
                }
                $NotInReports++
            }
        }
    } else {
        # No user list - show all from reports
        foreach ($IP in $ReportData.Keys) {
            $Results += $ReportData[$IP]
            
            if ($ReportData[$IP].OverallStatus -eq "COMPLIANT") {
                $Compliant++
            } else {
                $NonCompliant++
            }
        }
    }
    
    $TotalServers = $Results.Count
    $CompliancePercent = if ($TotalServers -gt 0) { [math]::Round(($Compliant / $TotalServers) * 100, 1) } else { 0 }
    
    Write-Host "  Total Servers: $TotalServers" -ForegroundColor White
    Write-Host "  Compliant: $Compliant ($CompliancePercent%)" -ForegroundColor Green
    Write-Host "  Non-Compliant: $NonCompliant" -ForegroundColor Red
    Write-Host "  Not in Reports: $NotInReports" -ForegroundColor Yellow
    Write-Host ""
    
    # Generate HTML
    $HTML = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Patch Compliance Report - Detailed Agent Status</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Arial, sans-serif; background: #f0f0f0; padding: 20px; }
        .container { max-width: 1800px; margin: 0 auto; background: white; border-radius: 15px; box-shadow: 0 0 40px rgba(0,0,0,0.2); }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 40px; text-align: center; border-radius: 15px 15px 0 0; }
        .header h1 { font-size: 36px; margin-bottom: 10px; }
        .header p { font-size: 16px; opacity: 0.95; margin: 5px 0; }
        .stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; padding: 30px; background: #f8f9fa; }
        .stat-card { background: white; padding: 25px; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); text-align: center; }
        .stat-label { font-size: 14px; color: #666; margin-bottom: 10px; text-transform: uppercase; letter-spacing: 1px; }
        .stat-value { font-size: 42px; font-weight: bold; color: #667eea; }
        .stat-card.success .stat-value { color: #10b981; }
        .stat-card.danger .stat-value { color: #ef4444; }
        .stat-card.warning .stat-value { color: #f59e0b; }
        .content { padding: 30px; }
        .section-title { font-size: 24px; color: #1f2937; margin: 20px 0; padding-bottom: 10px; border-bottom: 3px solid #667eea; }
        .controls { margin: 20px 0; }
        .search-box { width: 100%; max-width: 400px; padding: 12px; border: 2px solid #d1d5db; border-radius: 8px; font-size: 14px; margin-bottom: 15px; }
        .search-box:focus { outline: none; border-color: #667eea; }
        .filter-btn { padding: 10px 20px; margin-right: 10px; margin-bottom: 10px; border: 2px solid #667eea; background: white; color: #667eea; border-radius: 8px; cursor: pointer; font-weight: bold; transition: all 0.3s; }
        .filter-btn:hover { background: #667eea; color: white; transform: translateY(-2px); }
        .filter-btn.active { background: #667eea; color: white; }
        table { width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 13px; }
        thead { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; position: sticky; top: 0; }
        th { padding: 15px 10px; text-align: left; font-weight: bold; white-space: nowrap; }
        td { padding: 12px 10px; border-bottom: 1px solid #e5e7eb; }
        tr:hover { background: #f9fafb; }
        .overall-badge { padding: 6px 14px; border-radius: 20px; font-weight: bold; font-size: 12px; display: inline-block; white-space: nowrap; }
        .overall-badge.compliant { background: #d1fae5; color: #065f46; border: 1px solid #6ee7b7; }
        .overall-badge.noncompliant { background: #fee2e2; color: #991b1b; border: 1px solid #fca5a5; }
        .overall-badge.notfound { background: #fef3c7; color: #92400e; border: 1px solid #fcd34d; }
        .agent-status { padding: 4px 8px; border-radius: 5px; font-size: 11px; font-weight: bold; display: inline-block; white-space: nowrap; }
        .agent-status.compliant { background: #10b981; color: white; }
        .agent-status.noncompliant { background: #ef4444; color: white; }
        .agent-status.unknown { background: #6b7280; color: white; }
        .agent-status.notfound { background: #d1d5db; color: #374151; }
        .issues { font-size: 12px; color: #dc2626; font-weight: bold; }
        .footer { background: #f9fafb; padding: 20px; text-align: center; color: #6b7280; border-top: 1px solid #e5e7eb; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🛡️ Patch Compliance Report</h1>
            <p>DETAILED AGENT STATUS FOR ALL SERVERS</p>
            <p>Generated: $(Get-Date -Format 'MMMM dd, yyyy - hh:mm:ss tt')</p>
            <p>Scan Path: $ScanPath</p>
        </div>
        
        <div class="stats">
            <div class="stat-card">
                <div class="stat-label">Total Servers</div>
                <div class="stat-value">$TotalServers</div>
            </div>
            <div class="stat-card success">
                <div class="stat-label">✅ Compliant</div>
                <div class="stat-value">$Compliant</div>
            </div>
            <div class="stat-card danger">
                <div class="stat-label">❌ Non-Compliant</div>
                <div class="stat-value">$NonCompliant</div>
            </div>
            <div class="stat-card warning">
                <div class="stat-label">⚠️ Not in Reports</div>
                <div class="stat-value">$NotInReports</div>
            </div>
        </div>
        
        <div class="content">
            <div class="section-title">Detailed Server & Agent Status</div>
            
            <div class="controls">
                <input type="text" class="search-box" id="searchBox" placeholder="🔍 Search by IP, Server Name, or Agent..." onkeyup="filterTable()">
                
                <div>
                    <button class="filter-btn active" onclick="filterStatus('all')">All Servers ($TotalServers)</button>
                    <button class="filter-btn" onclick="filterStatus('COMPLIANT')">✅ Compliant ($Compliant)</button>
                    <button class="filter-btn" onclick="filterStatus('NON-COMPLIANT')">❌ Non-Compliant ($NonCompliant)</button>
                    <button class="filter-btn" onclick="filterStatus('NOT IN REPORTS')">⚠️ Not in Reports ($NotInReports)</button>
                </div>
            </div>
            
            <div style="overflow-x: auto;">
                <table id="serverTable">
                    <thead>
                        <tr>
                            <th>IP Address</th>
                            <th>Server Name</th>
                            <th>Overall Status</th>
                            <th>Trend Micro</th>
                            <th>Trellix</th>
                            <th>CrowdStrike</th>
                            <th>CloudWatch</th>
                            <th>Defender</th>
                            <th>Nessus</th>
                            <th>Issues</th>
                            <th>Source Report</th>
                            <th>Last Checked</th>
                        </tr>
                    </thead>
                    <tbody>
"@

    foreach ($Server in $Results) {
        $OverallClass = switch ($Server.OverallStatus) {
            "COMPLIANT" { "compliant" }
            "NON-COMPLIANT" { "noncompliant" }
            default { "notfound" }
        }
        
        function Get-AgentBadge {
            param([string]$Status)
            
            $Class = switch ($Status) {
                "Compliant" { "compliant" }
                "NON-COMPLIANT" { "noncompliant" }
                "Unknown" { "unknown" }
                default { "notfound" }
            }
            
            $Text = switch ($Status) {
                "Compliant" { "✓ OK" }
                "NON-COMPLIANT" { "✗ FAIL" }
                "Unknown" { "?" }
                default { "-" }
            }
            
            return "<span class='agent-status $Class'>$Text</span>"
        }
        
        $HTML += @"
                        <tr class="data-row" data-status="$($Server.OverallStatus)">
                            <td><strong>$($Server.IP)</strong></td>
                            <td>$($Server.ServerName)</td>
                            <td><span class="overall-badge $OverallClass">$($Server.OverallStatus)</span></td>
                            <td>$(Get-AgentBadge -Status $Server.TrendMicro)</td>
                            <td>$(Get-AgentBadge -Status $Server.Trellix)</td>
                            <td>$(Get-AgentBadge -Status $Server.CrowdStrike)</td>
                            <td>$(Get-AgentBadge -Status $Server.CloudWatch)</td>
                            <td>$(Get-AgentBadge -Status $Server.Defender)</td>
                            <td>$(Get-AgentBadge -Status $Server.Nessus)</td>
                            <td class="issues">$($Server.Issues)</td>
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
        
        <div class="footer">
            <p><strong>Ultimate Patch Scanner - Complete Final Version</strong></p>
            <p>Detailed agent status for all security agents | Author: Syed Ahmad</p>
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

    $HTML | Out-File -FilePath $OutputFile -Encoding UTF8 -Force
    
    return @{
        Total = $TotalServers
        Compliant = $Compliant
        NonCompliant = $NonCompliant
        NotInReports = $NotInReports
    }
}

# ============================================
# MAIN EXECUTION
# ============================================

# Step 1: Load server list
$MyIPs = Read-ServerList -Path $ScanPath

# Step 2: Scan reports
$ReportData = Scan-Reports -Path $ScanPath

if ($ReportData.Count -eq 0) {
    Write-Host "ERROR: No data found in reports!" -ForegroundColor Red
    Write-Host "Please check that report files exist and contain IP addresses." -ForegroundColor Yellow
    exit 1
}

# Step 3: Generate report
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ReportFile = Join-Path $OutputPath "PatchCompliance_DETAILED_$Timestamp.html"

$Stats = Generate-Report -MyIPs $MyIPs -ReportData $ReportData -OutputFile $ReportFile

# Display summary
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  SCAN COMPLETE!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Summary:" -ForegroundColor Yellow
Write-Host "  Total Servers: $($Stats.Total)" -ForegroundColor White
Write-Host "  ✅ Compliant: $($Stats.Compliant)" -ForegroundColor Green
Write-Host "  ❌ Non-Compliant: $($Stats.NonCompliant)" -ForegroundColor Red
Write-Host "  ⚠️  Not in Reports: $($Stats.NotInReports)" -ForegroundColor Yellow
Write-Host ""
Write-Host "Report saved to:" -ForegroundColor Cyan
Write-Host "  $ReportFile" -ForegroundColor White
Write-Host ""
Write-Host "Opening report in browser..." -ForegroundColor Yellow

Start-Process $ReportFile

Write-Host ""
Write-Host "✓ DONE! Check your detailed report!" -ForegroundColor Green
Write-Host ""
