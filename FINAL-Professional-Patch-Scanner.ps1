<#
.SYNOPSIS
    Professional Patch Compliance Scanner with Server List Comparison
    
.DESCRIPTION
    Compares YOUR server list (Rick-patch-List.xlsx) against compliance reports.
    Shows detailed agent status ONLY for servers in YOUR list.
    
.PARAMETER ScanPath
    Path containing compliance reports (default: C:\Reports-alerts)
    
.PARAMETER OutputPath
    Where to save reports (default: C:\Reports)
    
.EXAMPLE
    .\FINAL-Professional-Patch-Scanner.ps1 -ScanPath "C:\Reports-alerts"
#>

param(
    [string]$ScanPath = "C:\Reports-alerts",
    [string]$OutputPath = "C:\Reports"
)

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  PROFESSIONAL PATCH COMPLIANCE SCANNER" -ForegroundColor White
Write-Host "  Server List Comparison with Agent Details" -ForegroundColor White
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

# ================================================================
# STEP 1: FIND AND READ YOUR SERVER LIST
# ================================================================

Write-Host "[STEP 1] Loading YOUR server list..." -ForegroundColor Yellow
Write-Host ""

$ParentPath = Split-Path $ScanPath -Parent
$ServerListFile = $null

# Search for Excel file
$ExcelFiles = @()
$ExcelFiles += Get-ChildItem -Path $ParentPath -Filter "*.xlsx" -ErrorAction SilentlyContinue
$ExcelFiles += Get-ChildItem -Path $ScanPath -Filter "*.xlsx" -ErrorAction SilentlyContinue

if ($ExcelFiles.Count -eq 0) {
    Write-Host "ERROR: No Excel file found for server list!" -ForegroundColor Red
    Write-Host "Searched in: $ParentPath and $ScanPath" -ForegroundColor Yellow
    exit 1
}

# Try to find Rick-patch-List or similar
$ServerListFile = $ExcelFiles | Where-Object { $_.Name -match "rick|patch|server|list" } | Select-Object -First 1

if (-not $ServerListFile) {
    $ServerListFile = $ExcelFiles[0]
}

Write-Host "  Using file: $($ServerListFile.Name)" -ForegroundColor Green
Write-Host "  Location: $($ServerListFile.DirectoryName)" -ForegroundColor Gray

# Read Excel file
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false

$Workbook = $Excel.Workbooks.Open($ServerListFile.FullName)
$Worksheet = $Workbook.Sheets.Item(1)
$Range = $Worksheet.UsedRange

# Extract IPs from Column A
$MyServerIPs = @()
for ($row = 1; $row -le $Range.Rows.Count; $row++) {
    $CellValue = $Worksheet.Cells.Item($row, 1).Text
    if ($CellValue -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
        $MyServerIPs += $CellValue
    }
}

$Workbook.Close($false)
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
[System.GC]::Collect()

if ($MyServerIPs.Count -eq 0) {
    Write-Host "ERROR: No IP addresses found in Excel file!" -ForegroundColor Red
    exit 1
}

Write-Host "  Loaded: $($MyServerIPs.Count) servers from YOUR list" -ForegroundColor Green
Write-Host ""

# ================================================================
# STEP 2: SCAN COMPLIANCE REPORTS
# ================================================================

Write-Host "[STEP 2] Scanning compliance reports..." -ForegroundColor Yellow
Write-Host ""

$ReportFiles = Get-ChildItem -Path $ScanPath -Include "*.msg","*.html","*.txt" -Recurse -ErrorAction SilentlyContinue

if ($ReportFiles.Count -eq 0) {
    Write-Host "ERROR: No report files found!" -ForegroundColor Red
    exit 1
}

Write-Host "  Found: $($ReportFiles.Count) report files" -ForegroundColor Green

# Store all server data from reports
$AllReportServers = @{}

foreach ($File in $ReportFiles) {
    Write-Host "  Processing: $($File.Name)" -ForegroundColor Gray
    
    # Read file
    $Content = ""
    if ($File.Extension -eq ".msg") {
        try {
            $Outlook = New-Object -ComObject Outlook.Application
            $Msg = $Outlook.Session.OpenSharedItem($File.FullName)
            $Content = $Msg.Body
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
        } catch {
            $Bytes = [System.IO.File]::ReadAllBytes($File.FullName)
            $Content = [System.Text.Encoding]::UTF8.GetString($Bytes)
        }
    } else {
        $Content = Get-Content $File.FullName -Raw -ErrorAction SilentlyContinue
    }
    
    if ([string]::IsNullOrWhiteSpace($Content)) { continue }
    
    # Find IPs
    $IPMatches = [regex]::Matches($Content, '\b(?:\d{1,3}\.){3}\d{1,3}\b')
    
    foreach ($IPMatch in $IPMatches) {
        $IP = $IPMatch.Value
        
        # Get context around IP
        $IPIndex = $Content.IndexOf($IP)
        $Start = [Math]::Max(0, $IPIndex - 1500)
        $End = [Math]::Min($Content.Length, $IPIndex + 1500)
        $Context = $Content.Substring($Start, $End - $Start)
        
        # Extract server name
        $ServerName = "Unknown"
        if ($Context -match "(EC2AMAZ-[A-Z0-9]+)") {
            $ServerName = $Matches[1]
        }
        
        # Check each agent
        $Agents = @{
            TrendMicro = "-"
            Trellix = "-"
            CrowdStrike = "-"
            CloudWatch = "-"
            Defender = "-"
            Nessus = "-"
        }
        
        # TrendMicro
        if ($Context -match "Trend") {
            if ($Context -match "Trend.{0,100}Compliant" -and $Context -notmatch "Trend.{0,100}NonCompliant") {
                $Agents.TrendMicro = "OK"
            } elseif ($Context -match "Trend.{0,100}(NonCompliant|Non-Compliant)") {
                $Agents.TrendMicro = "FAIL"
            } else {
                $Agents.TrendMicro = "?"
            }
        }
        
        # Trellix
        if ($Context -match "Trellix") {
            if ($Context -match "Trellix.{0,100}Compliant" -and $Context -notmatch "Trellix.{0,100}NonCompliant") {
                $Agents.Trellix = "OK"
            } elseif ($Context -match "Trellix.{0,100}(NonCompliant|Non-Compliant)") {
                $Agents.Trellix = "FAIL"
            } else {
                $Agents.Trellix = "?"
            }
        }
        
        # CrowdStrike
        if ($Context -match "CrowdStrike") {
            if ($Context -match "CrowdStrike.{0,100}Compliant" -and $Context -notmatch "CrowdStrike.{0,100}NonCompliant") {
                $Agents.CrowdStrike = "OK"
            } elseif ($Context -match "CrowdStrike.{0,100}(NonCompliant|Non-Compliant)") {
                $Agents.CrowdStrike = "FAIL"
            } else {
                $Agents.CrowdStrike = "?"
            }
        }
        
        # CloudWatch
        if ($Context -match "CloudWatch") {
            if ($Context -match "CloudWatch.{0,100}Compliant" -and $Context -notmatch "CloudWatch.{0,100}NonCompliant") {
                $Agents.CloudWatch = "OK"
            } elseif ($Context -match "CloudWatch.{0,100}(NonCompliant|Non-Compliant)") {
                $Agents.CloudWatch = "FAIL"
            } else {
                $Agents.CloudWatch = "?"
            }
        }
        
        # Defender
        if ($Context -match "Defender") {
            if ($Context -match "Defender.{0,100}Compliant" -and $Context -notmatch "Defender.{0,100}NonCompliant") {
                $Agents.Defender = "OK"
            } elseif ($Context -match "Defender.{0,100}(NonCompliant|Non-Compliant)") {
                $Agents.Defender = "FAIL"
            } else {
                $Agents.Defender = "?"
            }
        }
        
        # Nessus
        if ($Context -match "Nessus") {
            if ($Context -match "Nessus.{0,100}Compliant" -and $Context -notmatch "Nessus.{0,100}NonCompliant") {
                $Agents.Nessus = "OK"
            } elseif ($Context -match "Nessus.{0,100}(NonCompliant|Non-Compliant)") {
                $Agents.Nessus = "FAIL"
            } else {
                $Agents.Nessus = "?"
            }
        }
        
        # Determine overall status
        $FailedAgents = @()
        foreach ($Agent in $Agents.Keys) {
            if ($Agents[$Agent] -eq "FAIL") {
                $FailedAgents += $Agent
            }
        }
        
        $OverallStatus = if ($FailedAgents.Count -gt 0) { "NON-COMPLIANT" } else { "COMPLIANT" }
        $Issues = if ($FailedAgents.Count -gt 0) { ($FailedAgents -join ", ") + " failed" } else { "All OK" }
        
        # Store (only if not already stored)
        if (-not $AllReportServers.ContainsKey($IP)) {
            $AllReportServers[$IP] = @{
                IP = $IP
                ServerName = $ServerName
                OverallStatus = $OverallStatus
                TrendMicro = $Agents.TrendMicro
                Trellix = $Agents.Trellix
                CrowdStrike = $Agents.CrowdStrike
                CloudWatch = $Agents.CloudWatch
                Defender = $Agents.Defender
                Nessus = $Agents.Nessus
                Issues = $Issues
                SourceReport = $File.Name
            }
        }
    }
}

Write-Host "  Found: $($AllReportServers.Count) total servers in all reports" -ForegroundColor Green
Write-Host ""

# ================================================================
# STEP 3: COMPARE YOUR LIST vs REPORTS
# ================================================================

Write-Host "[STEP 3] Comparing YOUR list against reports..." -ForegroundColor Yellow
Write-Host ""

$FinalResults = @()
$Stats = @{
    Total = $MyServerIPs.Count
    Compliant = 0
    NonCompliant = 0
    NotInReports = 0
}

foreach ($MyIP in $MyServerIPs) {
    if ($AllReportServers.ContainsKey($MyIP)) {
        # Server found in reports
        $ServerData = $AllReportServers[$MyIP]
        $FinalResults += $ServerData
        
        if ($ServerData.OverallStatus -eq "COMPLIANT") {
            $Stats.Compliant++
        } else {
            $Stats.NonCompliant++
        }
    } else {
        # Server NOT found in reports
        $FinalResults += @{
            IP = $MyIP
            ServerName = "Not Found"
            OverallStatus = "NOT IN REPORTS"
            TrendMicro = "N/A"
            Trellix = "N/A"
            CrowdStrike = "N/A"
            CloudWatch = "N/A"
            Defender = "N/A"
            Nessus = "N/A"
            Issues = "This server was not found in any compliance report"
            SourceReport = "N/A"
        }
        $Stats.NotInReports++
    }
}

Write-Host "  YOUR List: $($Stats.Total) servers" -ForegroundColor White
Write-Host "  Found Compliant: $($Stats.Compliant)" -ForegroundColor Green
Write-Host "  Found Non-Compliant: $($Stats.NonCompliant)" -ForegroundColor Red
Write-Host "  Not in Reports: $($Stats.NotInReports)" -ForegroundColor Yellow
Write-Host ""

# ================================================================
# STEP 4: GENERATE HTML REPORT
# ================================================================

Write-Host "[STEP 4] Generating professional HTML report..." -ForegroundColor Yellow
Write-Host ""

$HTML = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Patch Compliance Report - YOUR Server List</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: Arial, sans-serif; background: #f5f5f5; padding: 20px; }
.container { max-width: 1900px; margin: 0 auto; background: white; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.15); }
.header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 40px; text-align: center; border-radius: 12px 12px 0 0; }
.header h1 { font-size: 32px; margin-bottom: 8px; }
.header p { font-size: 15px; opacity: 0.95; }
.alert { background: #fff3cd; border-left: 4px solid #ffc107; padding: 15px 20px; margin: 20px 30px; border-radius: 4px; color: #856404; }
.stats { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; padding: 30px; background: #fafafa; }
.stat { background: white; padding: 25px; border-radius: 8px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
.stat-label { font-size: 13px; color: #666; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px; }
.stat-value { font-size: 36px; font-weight: bold; color: #333; }
.stat.success .stat-value { color: #10b981; }
.stat.danger .stat-value { color: #ef4444; }
.stat.warning .stat-value { color: #f59e0b; }
.content { padding: 30px; }
h2 { color: #333; margin-bottom: 20px; font-size: 22px; }
table { width: 100%; border-collapse: collapse; font-size: 13px; background: white; }
thead { background: linear-gradient(135deg, #667eea, #764ba2); color: white; }
th { padding: 14px 10px; text-align: left; font-weight: 600; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; }
td { padding: 12px 10px; border-bottom: 1px solid #eee; }
tr:hover { background: #f9fafb; }
.status-ok { background: #d1fae5; color: #065f46; padding: 6px 12px; border-radius: 16px; font-weight: 600; font-size: 11px; display: inline-block; }
.status-fail { background: #fee2e2; color: #991b1b; padding: 6px 12px; border-radius: 16px; font-weight: 600; font-size: 11px; display: inline-block; }
.status-notfound { background: #fef3c7; color: #92400e; padding: 6px 12px; border-radius: 16px; font-weight: 600; font-size: 11px; display: inline-block; }
.agent-ok { background: #10b981; color: white; padding: 4px 8px; border-radius: 4px; font-size: 10px; font-weight: 600; display: inline-block; }
.agent-fail { background: #ef4444; color: white; padding: 4px 8px; border-radius: 4px; font-size: 10px; font-weight: 600; display: inline-block; }
.agent-unknown { background: #6b7280; color: white; padding: 4px 8px; border-radius: 4px; font-size: 10px; font-weight: 600; display: inline-block; }
.agent-na { background: #d1d5db; color: #4b5563; padding: 4px 8px; border-radius: 4px; font-size: 10px; font-weight: 600; display: inline-block; }
.footer { background: #fafafa; padding: 20px; text-align: center; color: #666; font-size: 13px; border-top: 1px solid #eee; }
</style>
</head>
<body>
<div class="container">
<div class="header">
<h1>Patch Compliance Report</h1>
<p>YOUR Server List Comparison with Detailed Agent Status</p>
<p>Generated: $(Get-Date -Format 'MMMM dd, yyyy - hh:mm:ss tt')</p>
</div>

<div class="alert">
<strong>Note:</strong> This report shows ONLY the $($Stats.Total) servers from YOUR Rick-patch-List.xlsx file. It does NOT include all servers from the compliance reports.
</div>

<div class="stats">
<div class="stat">
<div class="stat-label">YOUR List Total</div>
<div class="stat-value">$($Stats.Total)</div>
</div>
<div class="stat success">
<div class="stat-label">Compliant</div>
<div class="stat-value">$($Stats.Compliant)</div>
</div>
<div class="stat danger">
<div class="stat-label">Non-Compliant</div>
<div class="stat-value">$($Stats.NonCompliant)</div>
</div>
<div class="stat warning">
<div class="stat-label">Not Found</div>
<div class="stat-value">$($Stats.NotInReports)</div>
</div>
</div>

<div class="content">
<h2>Server Details - ONLY Servers from YOUR List</h2>
<table>
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
</tr>
</thead>
<tbody>
"@

foreach ($Server in $FinalResults) {
    $StatusClass = if ($Server.OverallStatus -eq "COMPLIANT") { "status-ok" } elseif ($Server.OverallStatus -eq "NON-COMPLIANT") { "status-fail" } else { "status-notfound" }
    
    function Get-AgentBadge($Status) {
        if ($Status -eq "OK") { return "<span class='agent-ok'>OK</span>" }
        elseif ($Status -eq "FAIL") { return "<span class='agent-fail'>FAIL</span>" }
        elseif ($Status -eq "?") { return "<span class='agent-unknown'>?</span>" }
        elseif ($Status -eq "-") { return "<span class='agent-na'>-</span>" }
        else { return "<span class='agent-na'>N/A</span>" }
    }
    
    $HTML += "<tr>"
    $HTML += "<td><strong>$($Server.IP)</strong></td>"
    $HTML += "<td>$($Server.ServerName)</td>"
    $HTML += "<td><span class='$StatusClass'>$($Server.OverallStatus)</span></td>"
    $HTML += "<td>$(Get-AgentBadge $Server.TrendMicro)</td>"
    $HTML += "<td>$(Get-AgentBadge $Server.Trellix)</td>"
    $HTML += "<td>$(Get-AgentBadge $Server.CrowdStrike)</td>"
    $HTML += "<td>$(Get-AgentBadge $Server.CloudWatch)</td>"
    $HTML += "<td>$(Get-AgentBadge $Server.Defender)</td>"
    $HTML += "<td>$(Get-AgentBadge $Server.Nessus)</td>"
    $HTML += "<td>$($Server.Issues)</td>"
    $HTML += "<td>$($Server.SourceReport)</td>"
    $HTML += "</tr>"
}

$HTML += @"
</tbody>
</table>
</div>

<div class="footer">
<strong>Professional Patch Compliance Scanner</strong><br>
Server List File: $($ServerListFile.Name) | Scan Path: $ScanPath
</div>
</div>
</body>
</html>
"@

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ReportFile = Join-Path $OutputPath "Professional_Comparison_$Timestamp.html"
$HTML | Out-File -FilePath $ReportFile -Encoding UTF8 -Force

Write-Host "  Report saved: $ReportFile" -ForegroundColor Green
Write-Host ""

# ================================================================
# SUMMARY
# ================================================================

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  COMPARISON COMPLETE!" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "YOUR Server List: $($Stats.Total) servers" -ForegroundColor White
Write-Host "  Compliant:      $($Stats.Compliant)" -ForegroundColor Green
Write-Host "  Non-Compliant:  $($Stats.NonCompliant)" -ForegroundColor Red
Write-Host "  Not in Reports: $($Stats.NotInReports)" -ForegroundColor Yellow
Write-Host ""
Write-Host "Report File: $ReportFile" -ForegroundColor Cyan
Write-Host ""
Write-Host "Opening report in browser..." -ForegroundColor Yellow

Start-Process $ReportFile

Write-Host ""
Write-Host "DONE! Report shows ONLY YOUR $($Stats.Total) servers from Rick-patch-List!" -ForegroundColor Green
Write-Host ""
