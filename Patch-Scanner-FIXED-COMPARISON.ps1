# PATCH SCANNER - FIXED COMPARISON VERSION
# Actually compares YOUR list vs reports!
# Author: Syed Ahmad

param(
    [string]$ScanPath = "C:\Reports-alerts",
    [string]$OutputPath = "C:\Reports"
)

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  PATCH SCANNER - FIXED COMPARISON" -ForegroundColor Yellow
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

function Read-ServerList {
    param([string]$Path)
    
    Write-Host "Step 1: Reading YOUR server list..." -ForegroundColor Cyan
    
    $ParentPath = Split-Path $Path -Parent
    $ExcelFiles = Get-ChildItem -Path $ParentPath -Filter "*.xlsx" -ErrorAction SilentlyContinue
    $ServerListFile = $ExcelFiles | Where-Object { $_.Name -match "rick|patch|server|list" } | Select-Object -First 1
    
    if (-not $ServerListFile) {
        Write-Host "  ERROR: No Rick-patch-List found!" -ForegroundColor Red
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
        
        $MyServers = @()
        for ($row = 1; $row -le $Range.Rows.Count; $row++) {
            $IP = $Worksheet.Cells.Item($row, 1).Text
            if ($IP -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
                $MyServers += $IP
                Write-Host "    Added from YOUR list: $IP" -ForegroundColor Gray
            }
        }
        
        $Workbook.Close($false)
        $Excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
        [System.GC]::Collect()
        
        Write-Host "  Loaded $($MyServers.Count) IPs from YOUR list" -ForegroundColor Green
        Write-Host ""
        return $MyServers
        
    } catch {
        Write-Host "  Error: $_" -ForegroundColor Red
        return @()
    }
}

function Scan-Reports {
    param([string]$Path)
    
    Write-Host "Step 2: Scanning compliance reports..." -ForegroundColor Cyan
    
    $Files = Get-ChildItem -Path $Path -Include "*.msg","*.html","*.txt" -Recurse -ErrorAction SilentlyContinue
    
    if ($Files.Count -eq 0) {
        Write-Host "  No files found!" -ForegroundColor Red
        return @{}
    }
    
    Write-Host "  Found $($Files.Count) files" -ForegroundColor Green
    
    $AllData = @{}
    
    foreach ($File in $Files) {
        Write-Host "  Processing: $($File.Name)" -ForegroundColor Gray
        
        try {
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
                $Content = Get-Content $File.FullName -Raw
            }
            
            $IPMatches = [regex]::Matches($Content, '\b(?:\d{1,3}\.){3}\d{1,3}\b')
            
            foreach ($Match in $IPMatches) {
                $IP = $Match.Value
                
                $Index = $Content.IndexOf($IP)
                $Start = [Math]::Max(0, $Index - 1000)
                $End = [Math]::Min($Content.Length, $Index + 1000)
                $Context = $Content.Substring($Start, $End - $Start)
                
                $ServerName = "Unknown"
                if ($Context -match "(EC2AMAZ-[A-Z0-9]+)") {
                    $ServerName = $Matches[1]
                }
                
                $TrendMicro = "-"
                $Trellix = "-"
                $CrowdStrike = "-"
                $CloudWatch = "-"
                $Defender = "-"
                $Nessus = "-"
                
                if ($Context -match "Trend") {
                    if ($Context -match "Trend.{0,50}Compliant" -and $Context -notmatch "Trend.{0,50}NonCompliant") {
                        $TrendMicro = "OK"
                    } elseif ($Context -match "Trend.{0,50}NonCompliant") {
                        $TrendMicro = "FAIL"
                    } else {
                        $TrendMicro = "?"
                    }
                }
                
                if ($Context -match "Trellix") {
                    if ($Context -match "Trellix.{0,50}Compliant" -and $Context -notmatch "Trellix.{0,50}NonCompliant") {
                        $Trellix = "OK"
                    } elseif ($Context -match "Trellix.{0,50}NonCompliant") {
                        $Trellix = "FAIL"
                    } else {
                        $Trellix = "?"
                    }
                }
                
                if ($Context -match "CrowdStrike") {
                    if ($Context -match "CrowdStrike.{0,50}Compliant" -and $Context -notmatch "CrowdStrike.{0,50}NonCompliant") {
                        $CrowdStrike = "OK"
                    } elseif ($Context -match "CrowdStrike.{0,50}NonCompliant") {
                        $CrowdStrike = "FAIL"
                    } else {
                        $CrowdStrike = "?"
                    }
                }
                
                if ($Context -match "CloudWatch") {
                    if ($Context -match "CloudWatch.{0,50}Compliant" -and $Context -notmatch "CloudWatch.{0,50}NonCompliant") {
                        $CloudWatch = "OK"
                    } elseif ($Context -match "CloudWatch.{0,50}NonCompliant") {
                        $CloudWatch = "FAIL"
                    } else {
                        $CloudWatch = "?"
                    }
                }
                
                if ($Context -match "Defender") {
                    if ($Context -match "Defender.{0,50}Compliant" -and $Context -notmatch "Defender.{0,50}NonCompliant") {
                        $Defender = "OK"
                    } elseif ($Context -match "Defender.{0,50}NonCompliant") {
                        $Defender = "FAIL"
                    } else {
                        $Defender = "?"
                    }
                }
                
                if ($Context -match "Nessus") {
                    if ($Context -match "Nessus.{0,50}Compliant" -and $Context -notmatch "Nessus.{0,50}NonCompliant") {
                        $Nessus = "OK"
                    } elseif ($Context -match "Nessus.{0,50}NonCompliant") {
                        $Nessus = "FAIL"
                    } else {
                        $Nessus = "?"
                    }
                }
                
                $OverallStatus = "COMPLIANT"
                $Issues = @()
                
                if ($TrendMicro -eq "FAIL") { $OverallStatus = "NON-COMPLIANT"; $Issues += "TrendMicro" }
                if ($Trellix -eq "FAIL") { $OverallStatus = "NON-COMPLIANT"; $Issues += "Trellix" }
                if ($CrowdStrike -eq "FAIL") { $OverallStatus = "NON-COMPLIANT"; $Issues += "CrowdStrike" }
                if ($CloudWatch -eq "FAIL") { $OverallStatus = "NON-COMPLIANT"; $Issues += "CloudWatch" }
                if ($Defender -eq "FAIL") { $OverallStatus = "NON-COMPLIANT"; $Issues += "Defender" }
                
                $IssueText = if ($Issues.Count -gt 0) { ($Issues -join ", ") + " failed" } else { "All OK" }
                
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
                        Source = $File.Name
                    }
                }
            }
        } catch {
            Write-Host "    Error: $_" -ForegroundColor Red
        }
    }
    
    Write-Host "  Found data for $($AllData.Count) total servers in reports" -ForegroundColor Green
    Write-Host ""
    
    return $AllData
}

function Compare-MyListVsReports {
    param(
        [array]$MyIPs,
        [hashtable]$AllReportData
    )
    
    Write-Host "Step 3: COMPARING YOUR LIST vs REPORTS..." -ForegroundColor Cyan
    Write-Host ""
    
    $Results = @()
    $Compliant = 0
    $NonCompliant = 0
    $NotFound = 0
    
    foreach ($IP in $MyIPs) {
        Write-Host "  Checking YOUR IP: $IP ..." -ForegroundColor Gray -NoNewline
        
        if ($AllReportData.ContainsKey($IP)) {
            $Data = $AllReportData[$IP]
            Write-Host " FOUND in reports!" -ForegroundColor Green
            
            $Results += $Data
            
            if ($Data.OverallStatus -eq "COMPLIANT") {
                $Compliant++
            } else {
                $NonCompliant++
            }
        } else {
            Write-Host " NOT FOUND in any report!" -ForegroundColor Red
            
            $Results += @{
                IP = $IP
                ServerName = "NotFound"
                OverallStatus = "NOT IN REPORTS"
                TrendMicro = "N/A"
                Trellix = "N/A"
                CrowdStrike = "N/A"
                CloudWatch = "N/A"
                Defender = "N/A"
                Nessus = "N/A"
                Issues = "This IP from YOUR list was not found in any compliance report"
                Source = "N/A"
            }
            $NotFound++
        }
    }
    
    Write-Host ""
    Write-Host "Comparison Results:" -ForegroundColor Yellow
    Write-Host "  From YOUR list: $($MyIPs.Count) IPs" -ForegroundColor White
    Write-Host "  Found Compliant: $Compliant" -ForegroundColor Green
    Write-Host "  Found Non-Compliant: $NonCompliant" -ForegroundColor Red
    Write-Host "  Not in Reports: $NotFound" -ForegroundColor Yellow
    Write-Host ""
    
    return @{
        Results = $Results
        Total = $MyIPs.Count
        Compliant = $Compliant
        NonCompliant = $NonCompliant
        NotFound = $NotFound
    }
}

function Generate-Report {
    param(
        [array]$Results,
        [int]$Total,
        [int]$Compliant,
        [int]$NonCompliant,
        [int]$NotFound,
        [string]$OutputFile
    )
    
    Write-Host "Step 4: Generating HTML report..." -ForegroundColor Cyan
    
    $HTML = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Patch Compliance Report - YOUR Server List</title>
<style>
body { font-family: Arial; background: #f0f0f0; padding: 20px; }
.container { max-width: 1800px; margin: 0 auto; background: white; border-radius: 10px; box-shadow: 0 0 20px rgba(0,0,0,0.2); }
.header { background: linear-gradient(135deg, #667eea, #764ba2); color: white; padding: 40px; text-align: center; }
.header h1 { margin: 0; font-size: 36px; }
.header p { margin: 10px 0; font-size: 16px; }
.alert { background: #fef3c7; border-left: 4px solid #f59e0b; padding: 15px; margin: 20px 30px; border-radius: 5px; }
.stats { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; padding: 30px; background: #f8f9fa; }
.stat-card { background: white; padding: 25px; border-radius: 10px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
.stat-label { font-size: 14px; color: #666; margin-bottom: 10px; text-transform: uppercase; }
.stat-value { font-size: 42px; font-weight: bold; }
.content { padding: 30px; }
.search-box { width: 400px; padding: 12px; border: 2px solid #ddd; border-radius: 8px; margin-bottom: 15px; }
.filter-btn { padding: 10px 20px; margin-right: 10px; margin-bottom: 10px; border: 2px solid #667eea; background: white; color: #667eea; border-radius: 8px; cursor: pointer; font-weight: bold; }
.filter-btn:hover { background: #667eea; color: white; }
.filter-btn.active { background: #667eea; color: white; }
table { width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 13px; }
thead { background: #667eea; color: white; }
th { padding: 15px 10px; text-align: left; }
td { padding: 12px 10px; border-bottom: 1px solid #e5e7eb; }
tr:hover { background: #f9fafb; }
.badge-ok { background: #10b981; color: white; padding: 4px 8px; border-radius: 5px; font-size: 11px; font-weight: bold; }
.badge-fail { background: #ef4444; color: white; padding: 4px 8px; border-radius: 5px; font-size: 11px; font-weight: bold; }
.badge-unknown { background: #6b7280; color: white; padding: 4px 8px; border-radius: 5px; font-size: 11px; font-weight: bold; }
.badge-na { background: #d1d5db; color: #374151; padding: 4px 8px; border-radius: 5px; font-size: 11px; font-weight: bold; }
.overall-ok { background: #d1fae5; color: #065f46; padding: 6px 14px; border-radius: 20px; font-weight: bold; font-size: 12px; }
.overall-fail { background: #fee2e2; color: #991b1b; padding: 6px 14px; border-radius: 20px; font-weight: bold; font-size: 12px; }
.overall-notfound { background: #fef3c7; color: #92400e; padding: 6px 14px; border-radius: 20px; font-weight: bold; font-size: 12px; }
</style>
</head>
<body>
<div class="container">
<div class="header">
<h1>Patch Compliance Report</h1>
<p>COMPARISON: YOUR Server List vs Compliance Reports</p>
<p>Generated: $(Get-Date -Format 'MMMM dd, yyyy - hh:mm tt')</p>
</div>

<div class="alert">
<strong>Note:</strong> This report shows ONLY the $Total servers from YOUR Rick-patch-List.xlsx file, NOT all servers from the compliance reports.
</div>

<div class="stats">
<div class="stat-card"><div class="stat-label">YOUR List Total</div><div class="stat-value">$Total</div></div>
<div class="stat-card"><div class="stat-label">Compliant</div><div class="stat-value" style="color:#10b981">$Compliant</div></div>
<div class="stat-card"><div class="stat-label">Non-Compliant</div><div class="stat-value" style="color:#ef4444">$NonCompliant</div></div>
<div class="stat-card"><div class="stat-label">Not Found</div><div class="stat-value" style="color:#f59e0b">$NotFound</div></div>
</div>

<div class="content">
<h2>YOUR Server List - Compliance Status</h2>

<input type="text" class="search-box" id="searchBox" placeholder="Search..." onkeyup="filterTable()">

<div>
<button class="filter-btn active" onclick="filterStatus('all')">All ($Total)</button>
<button class="filter-btn" onclick="filterStatus('COMPLIANT')">Compliant ($Compliant)</button>
<button class="filter-btn" onclick="filterStatus('NON-COMPLIANT')">Non-Compliant ($NonCompliant)</button>
<button class="filter-btn" onclick="filterStatus('NOT IN REPORTS')">Not Found ($NotFound)</button>
</div>

<table id="serverTable">
<thead>
<tr>
<th>IP (from YOUR list)</th>
<th>Server</th>
<th>Overall</th>
<th>TrendMicro</th>
<th>Trellix</th>
<th>CrowdStrike</th>
<th>CloudWatch</th>
<th>Defender</th>
<th>Nessus</th>
<th>Issues</th>
<th>Source</th>
</tr>
</thead>
<tbody>
"@

    foreach ($Server in $Results) {
        $OverallClass = if ($Server.OverallStatus -eq "COMPLIANT") { "overall-ok" } elseif ($Server.OverallStatus -eq "NON-COMPLIANT") { "overall-fail" } else { "overall-notfound" }
        
        function Get-Badge {
            param([string]$Status)
            
            if ($Status -eq "OK") {
                return "<span class='badge-ok'>OK</span>"
            } elseif ($Status -eq "FAIL") {
                return "<span class='badge-fail'>FAIL</span>"
            } elseif ($Status -eq "?") {
                return "<span class='badge-unknown'>?</span>"
            } elseif ($Status -eq "-") {
                return "<span class='badge-na'>-</span>"
            } else {
                return "<span class='badge-na'>N/A</span>"
            }
        }
        
        $HTML += "<tr class='data-row' data-status='$($Server.OverallStatus)'>"
        $HTML += "<td><strong>$($Server.IP)</strong></td>"
        $HTML += "<td>$($Server.ServerName)</td>"
        $HTML += "<td><span class='$OverallClass'>$($Server.OverallStatus)</span></td>"
        $HTML += "<td>$(Get-Badge -Status $Server.TrendMicro)</td>"
        $HTML += "<td>$(Get-Badge -Status $Server.Trellix)</td>"
        $HTML += "<td>$(Get-Badge -Status $Server.CrowdStrike)</td>"
        $HTML += "<td>$(Get-Badge -Status $Server.CloudWatch)</td>"
        $HTML += "<td>$(Get-Badge -Status $Server.Defender)</td>"
        $HTML += "<td>$(Get-Badge -Status $Server.Nessus)</td>"
        $HTML += "<td>$($Server.Issues)</td>"
        $HTML += "<td>$($Server.Source)</td>"
        $HTML += "</tr>"
    }

    $HTML += @"
</tbody>
</table>
</div>
</div>

<script>
function filterStatus(status) {
var rows = document.querySelectorAll('.data-row');
var buttons = document.querySelectorAll('.filter-btn');
buttons.forEach(function(btn) { btn.classList.remove('active'); });
event.target.classList.add('active');
rows.forEach(function(row) {
var rowStatus = row.getAttribute('data-status');
row.style.display = (status === 'all' || rowStatus === status) ? '' : 'none';
});
}

function filterTable() {
var filter = document.getElementById('searchBox').value.toUpperCase();
var rows = document.querySelectorAll('.data-row');
rows.forEach(function(row) {
var text = row.textContent || row.innerText;
row.style.display = text.toUpperCase().indexOf(filter) > -1 ? '' : 'none';
});
}
</script>
</body>
</html>
"@

    $HTML | Out-File -FilePath $OutputFile -Encoding UTF8 -Force
    Write-Host "  Report saved!" -ForegroundColor Green
    Write-Host ""
}

# MAIN EXECUTION
$MyIPs = Read-ServerList -Path $ScanPath

if ($MyIPs.Count -eq 0) {
    Write-Host "ERROR: No IPs found in YOUR Rick-patch-List!" -ForegroundColor Red
    exit 1
}

$AllReportData = Scan-Reports -Path $ScanPath

if ($AllReportData.Count -eq 0) {
    Write-Host "ERROR: No data in compliance reports!" -ForegroundColor Red
    exit 1
}

$Comparison = Compare-MyListVsReports -MyIPs $MyIPs -AllReportData $AllReportData

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ReportFile = Join-Path $OutputPath "YourList_Comparison_$Timestamp.html"

Generate-Report -Results $Comparison.Results -Total $Comparison.Total -Compliant $Comparison.Compliant -NonCompliant $Comparison.NonCompliant -NotFound $Comparison.NotFound -OutputFile $ReportFile

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  COMPARISON COMPLETE!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "YOUR Server List: $($Comparison.Total) IPs" -ForegroundColor White
Write-Host "  Compliant: $($Comparison.Compliant)" -ForegroundColor Green
Write-Host "  Non-Compliant: $($Comparison.NonCompliant)" -ForegroundColor Red
Write-Host "  Not Found: $($Comparison.NotFound)" -ForegroundColor Yellow
Write-Host ""
Write-Host "Report: $ReportFile" -ForegroundColor Cyan
Write-Host ""

Start-Process $ReportFile

Write-Host "DONE! This shows ONLY YOUR servers from Rick-patch-List!" -ForegroundColor Green
Write-Host ""
