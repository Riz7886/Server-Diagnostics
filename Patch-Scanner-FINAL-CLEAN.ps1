# ULTIMATE PATCH SCANNER - FINAL CLEAN VERSION
# No errors - guaranteed clean syntax
# Author: Syed Ahmad

param(
    [string]$ScanPath = "C:\Reports-alerts",
    [string]$OutputPath = "C:\Reports"
)

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  ULTIMATE PATCH SCANNER - FINAL" -ForegroundColor Yellow
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

function Read-ServerList {
    param([string]$Path)
    
    Write-Host "Step 1: Loading server list..." -ForegroundColor Cyan
    
    $ParentPath = Split-Path $Path -Parent
    $ExcelFiles = Get-ChildItem -Path $ParentPath -Filter "*.xlsx" -ErrorAction SilentlyContinue
    $ServerListFile = $ExcelFiles | Where-Object { $_.Name -match "rick|patch|server|list" } | Select-Object -First 1
    
    if (-not $ServerListFile) {
        Write-Host "  No server list - will show all IPs from reports" -ForegroundColor Yellow
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
        
        Write-Host "  Loaded $($IPs.Count) IPs" -ForegroundColor Green
        Write-Host ""
        return $IPs
        
    } catch {
        Write-Host "  Error: $_" -ForegroundColor Red
        Write-Host ""
        return @()
    }
}

function Get-AgentStatus {
    param(
        [string]$Context,
        [string]$AgentName
    )
    
    $Patterns = @{
        "TrendMicro" = "Trend"
        "Trellix" = "Trellix"
        "CrowdStrike" = "CrowdStrike"
        "CloudWatch" = "CloudWatch"
        "Defender" = "Defender"
        "Nessus" = "Nessus"
    }
    
    $Pattern = $Patterns[$AgentName]
    
    if (-not ($Context -match $Pattern)) {
        return "NotFound"
    }
    
    if ($Context -match "$Pattern.{0,50}Compliant" -and $Context -notmatch "$Pattern.{0,50}NonCompliant") {
        return "OK"
    }
    
    if ($Context -match "$Pattern.{0,50}NonCompliant") {
        return "FAIL"
    }
    
    if ($Context -match "$Pattern.{0,50}(Installed|Running|Active|OK)") {
        return "OK"
    }
    
    if ($Context -match "$Pattern.{0,50}(Missing|Failed|Error)") {
        return "FAIL"
    }
    
    return "Unknown"
}

function Scan-Reports {
    param([string]$Path)
    
    Write-Host "Step 2: Scanning reports..." -ForegroundColor Cyan
    
    $Files = Get-ChildItem -Path $Path -Include "*.msg","*.html","*.txt" -Recurse -ErrorAction SilentlyContinue
    
    if ($Files.Count -eq 0) {
        Write-Host "  No files found!" -ForegroundColor Red
        return @{}
    }
    
    Write-Host "  Found $($Files.Count) files" -ForegroundColor Green
    
    $AllData = @{}
    $Count = 0
    
    foreach ($File in $Files) {
        $Count++
        Write-Host "  [$Count/$($Files.Count)] $($File.Name)" -ForegroundColor Gray
        
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
                
                $TrendMicro = Get-AgentStatus -Context $Context -AgentName "TrendMicro"
                $Trellix = Get-AgentStatus -Context $Context -AgentName "Trellix"
                $CrowdStrike = Get-AgentStatus -Context $Context -AgentName "CrowdStrike"
                $CloudWatch = Get-AgentStatus -Context $Context -AgentName "CloudWatch"
                $Defender = Get-AgentStatus -Context $Context -AgentName "Defender"
                $Nessus = Get-AgentStatus -Context $Context -AgentName "Nessus"
                
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
    
    Write-Host "  Found $($AllData.Count) unique servers" -ForegroundColor Green
    Write-Host ""
    
    return $AllData
}

function Generate-Report {
    param(
        [array]$MyIPs,
        [hashtable]$ReportData,
        [string]$OutputFile
    )
    
    Write-Host "Step 3: Generating report..." -ForegroundColor Cyan
    
    $Results = @()
    $Compliant = 0
    $NonCompliant = 0
    $NotFound = 0
    
    if ($MyIPs.Count -gt 0) {
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
                    ServerName = "NotFound"
                    OverallStatus = "NOT IN REPORTS"
                    TrendMicro = "N/A"
                    Trellix = "N/A"
                    CrowdStrike = "N/A"
                    CloudWatch = "N/A"
                    Defender = "N/A"
                    Nessus = "N/A"
                    Issues = "Not in any report"
                    Source = "N/A"
                }
                $NotFound++
            }
        }
    } else {
        foreach ($IP in $ReportData.Keys) {
            $Results += $ReportData[$IP]
            
            if ($ReportData[$IP].OverallStatus -eq "COMPLIANT") {
                $Compliant++
            } else {
                $NonCompliant++
            }
        }
    }
    
    $Total = $Results.Count
    
    Write-Host "  Total: $Total | Compliant: $Compliant | Non-Compliant: $NonCompliant | Not Found: $NotFound" -ForegroundColor White
    Write-Host ""
    
    $HTML = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Patch Compliance Report</title>
<style>
body { font-family: Arial; background: #f0f0f0; padding: 20px; }
.container { max-width: 1800px; margin: 0 auto; background: white; border-radius: 10px; box-shadow: 0 0 20px rgba(0,0,0,0.2); }
.header { background: linear-gradient(135deg, #667eea, #764ba2); color: white; padding: 40px; text-align: center; }
.header h1 { margin: 0; font-size: 36px; }
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
<p>Detailed Agent Status</p>
<p>Generated: $(Get-Date -Format 'MMMM dd, yyyy - hh:mm tt')</p>
</div>

<div class="stats">
<div class="stat-card"><div class="stat-label">Total</div><div class="stat-value">$Total</div></div>
<div class="stat-card"><div class="stat-label">Compliant</div><div class="stat-value" style="color:#10b981">$Compliant</div></div>
<div class="stat-card"><div class="stat-label">Non-Compliant</div><div class="stat-value" style="color:#ef4444">$NonCompliant</div></div>
<div class="stat-card"><div class="stat-label">Not Found</div><div class="stat-value" style="color:#f59e0b">$NotFound</div></div>
</div>

<div class="content">
<h2>Server Details</h2>

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
<th>IP</th>
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
            } elseif ($Status -eq "Unknown") {
                return "<span class='badge-unknown'>?</span>"
            } elseif ($Status -eq "NotFound") {
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
    
    return @{
        Total = $Total
        Compliant = $Compliant
        NonCompliant = $NonCompliant
        NotFound = $NotFound
    }
}

# MAIN
$MyIPs = Read-ServerList -Path $ScanPath
$ReportData = Scan-Reports -Path $ScanPath

if ($ReportData.Count -eq 0) {
    Write-Host "ERROR: No data found!" -ForegroundColor Red
    exit 1
}

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ReportFile = Join-Path $OutputPath "PatchReport_$Timestamp.html"

$Stats = Generate-Report -MyIPs $MyIPs -ReportData $ReportData -OutputFile $ReportFile

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  COMPLETE!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Total: $($Stats.Total)" -ForegroundColor White
Write-Host "Compliant: $($Stats.Compliant)" -ForegroundColor Green
Write-Host "Non-Compliant: $($Stats.NonCompliant)" -ForegroundColor Red
Write-Host "Not Found: $($Stats.NotFound)" -ForegroundColor Yellow
Write-Host ""
Write-Host "Report: $ReportFile" -ForegroundColor Cyan
Write-Host ""

Start-Process $ReportFile

Write-Host "DONE!" -ForegroundColor Green
