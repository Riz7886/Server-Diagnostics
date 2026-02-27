param(
    [ValidateSet("Compare","FullReport")]
    [string]$Mode = "Compare",
    [string]$ScanPath = "C:\Reports-alerts",
    [string]$ServerListFile = "",
    [string]$OutputPath = "C:\Reports"
)

$ErrorActionPreference = "SilentlyContinue"

if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Compliance Scanner - Mode: $Mode" -ForegroundColor White
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

function Get-ServerList {
    param([string]$FilePath)
    
    if ([string]::IsNullOrWhiteSpace($FilePath)) {
        $SearchPaths = @((Split-Path $ScanPath -Parent), $ScanPath)
        $ExcelFiles = @()
        
        foreach ($Path in $SearchPaths) {
            if (Test-Path $Path) {
                $ExcelFiles += Get-ChildItem -Path $Path -Filter "*.xlsx" -ErrorAction SilentlyContinue
            }
        }
        
        if ($ExcelFiles.Count -eq 0) { return @() }
        
        $Match = $ExcelFiles | Where-Object { $_.Name -match "rick|patch|server|list" } | Select-Object -First 1
        $FilePath = if ($Match) { $Match.FullName } else { $ExcelFiles[0].FullName }
    }
    
    if (-not (Test-Path $FilePath)) { return @() }
    
    Write-Host "Loading server list: $([System.IO.Path]::GetFileName($FilePath))" -ForegroundColor Yellow
    
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    
    $Workbook = $Excel.Workbooks.Open($FilePath)
    $Worksheet = $Workbook.Sheets.Item(1)
    $Range = $Worksheet.UsedRange
    
    $IPs = @()
    for ($i = 1; $i -le $Range.Rows.Count; $i++) {
        $Value = $Worksheet.Cells.Item($i, 1).Text
        if ($Value -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
            $IPs += $Value
        }
    }
    
    $Workbook.Close($false)
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Host "Loaded: $($IPs.Count) servers" -ForegroundColor Green
    Write-Host ""
    
    return $IPs
}

function Scan-Reports {
    param([string]$Path)
    
    Write-Host "Scanning compliance reports..." -ForegroundColor Yellow
    
    $Files = Get-ChildItem -Path $Path -Include "*.msg","*.html","*.txt","*.htm" -Recurse -ErrorAction SilentlyContinue
    
    if ($Files.Count -eq 0) {
        Write-Host "ERROR: No report files found in $Path" -ForegroundColor Red
        exit 1
    }
    
    Write-Host "Found: $($Files.Count) report files" -ForegroundColor Green
    Write-Host ""
    
    $AllServers = @{}
    
    foreach ($File in $Files) {
        $Content = ""
        
        if ($File.Extension -eq ".msg") {
            try {
                $Outlook = New-Object -ComObject Outlook.Application
                $Msg = $Outlook.Session.OpenSharedItem($File.FullName)
                $Content = $Msg.Body
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
            } catch {
                try {
                    $Bytes = [System.IO.File]::ReadAllBytes($File.FullName)
                    $Content = [System.Text.Encoding]::UTF8.GetString($Bytes)
                } catch { }
            }
        } else {
            $Content = Get-Content $File.FullName -Raw -ErrorAction SilentlyContinue
        }
        
        if ([string]::IsNullOrWhiteSpace($Content)) { continue }
        
        $IPMatches = [regex]::Matches($Content, '\b(?:\d{1,3}\.){3}\d{1,3}\b')
        
        foreach ($IPMatch in $IPMatches) {
            $IP = $IPMatch.Value
            
            if ($AllServers.ContainsKey($IP)) { continue }
            
            $IPIndex = $Content.IndexOf($IP)
            $Start = [Math]::Max(0, $IPIndex - 2000)
            $End = [Math]::Min($Content.Length, $IPIndex + 2000)
            $Context = $Content.Substring($Start, $End - $Start)
            
            $ServerName = "Unknown"
            if ($Context -match "(EC2AMAZ-[A-Z0-9]+)") {
                $ServerName = $Matches[1]
            }
            
            $Agents = @{
                TrendMicro = "-"
                Trellix = "-"
                CrowdStrike = "-"
                CloudWatch = "-"
                Defender = "-"
                Nessus = "-"
            }
            
            $AgentPatterns = @{
                "TrendMicro" = "Trend"
                "Trellix" = "Trellix"
                "CrowdStrike" = "CrowdStrike"
                "CloudWatch" = "CloudWatch"
                "Defender" = "Defender"
                "Nessus" = "Nessus"
            }
            
            foreach ($AgentKey in $AgentPatterns.Keys) {
                $Pattern = $AgentPatterns[$AgentKey]
                
                if ($Context -match $Pattern) {
                    if ($Context -match "$Pattern.{0,150}Compliant" -and $Context -notmatch "$Pattern.{0,150}(NonCompliant|Non-Compliant)") {
                        $Agents[$AgentKey] = "OK"
                    } elseif ($Context -match "$Pattern.{0,150}(NonCompliant|Non-Compliant)") {
                        $Agents[$AgentKey] = "FAIL"
                    } else {
                        $Agents[$AgentKey] = "?"
                    }
                }
            }
            
            $Failed = @()
            foreach ($Key in $Agents.Keys) {
                if ($Agents[$Key] -eq "FAIL") {
                    $Failed += $Key
                }
            }
            
            $Status = if ($Failed.Count -gt 0) { "NON-COMPLIANT" } else { "COMPLIANT" }
            $Issues = if ($Failed.Count -gt 0) { ($Failed -join ", ") + " failed" } else { "All OK" }
            
            $AllServers[$IP] = @{
                IP = $IP
                ServerName = $ServerName
                OverallStatus = $Status
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
    
    Write-Host "Processed: $($AllServers.Count) unique servers" -ForegroundColor Green
    Write-Host ""
    
    return $AllServers
}

function Generate-HTMLReport {
    param($Results, $Stats, $ReportMode, $OutputFile)
    
    Write-Host "Generating HTML report..." -ForegroundColor Yellow
    
    $HTML = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Compliance Report</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}body{font-family:Arial,sans-serif;background:#f5f5f5;padding:20px}.container{max-width:1900px;margin:0 auto;background:#fff;border-radius:12px;box-shadow:0 4px 20px rgba(0,0,0,.15)}.header{background:linear-gradient(135deg,#667eea,#764ba2);color:#fff;padding:40px;text-align:center;border-radius:12px 12px 0 0}.header h1{font-size:32px;margin-bottom:8px}.header p{font-size:14px;opacity:.95}.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:20px;padding:30px;background:#fafafa}.stat{background:#fff;padding:25px;border-radius:8px;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,.08)}.stat-label{font-size:13px;color:#666;text-transform:uppercase;margin-bottom:8px;letter-spacing:.5px}.stat-value{font-size:36px;font-weight:700;color:#333}.stat.success .stat-value{color:#10b981}.stat.danger .stat-value{color:#ef4444}.stat.warning .stat-value{color:#f59e0b}.content{padding:30px}h2{color:#333;margin-bottom:20px;font-size:22px}table{width:100%;border-collapse:collapse;font-size:13px;background:#fff}thead{background:linear-gradient(135deg,#667eea,#764ba2);color:#fff}th{padding:14px 10px;text-align:left;font-weight:600;font-size:12px;text-transform:uppercase;letter-spacing:.5px}td{padding:12px 10px;border-bottom:1px solid #eee}tr:hover{background:#f9fafb}.status-ok{background:#d1fae5;color:#065f46;padding:6px 12px;border-radius:16px;font-weight:600;font-size:11px;display:inline-block}.status-fail{background:#fee2e2;color:#991b1b;padding:6px 12px;border-radius:16px;font-weight:600;font-size:11px;display:inline-block}.status-notfound{background:#fef3c7;color:#92400e;padding:6px 12px;border-radius:16px;font-weight:600;font-size:11px;display:inline-block}.agent-ok{background:#10b981;color:#fff;padding:4px 8px;border-radius:4px;font-size:10px;font-weight:600;display:inline-block}.agent-fail{background:#ef4444;color:#fff;padding:4px 8px;border-radius:4px;font-size:10px;font-weight:600;display:inline-block}.agent-unknown{background:#6b7280;color:#fff;padding:4px 8px;border-radius:4px;font-size:10px;font-weight:600;display:inline-block}.agent-na{background:#d1d5db;color:#4b5563;padding:4px 8px;border-radius:4px;font-size:10px;font-weight:600;display:inline-block}.footer{background:#fafafa;padding:20px;text-align:center;color:#666;font-size:13px;border-top:1px solid #eee}
</style>
</head>
<body>
<div class="container">
<div class="header">
<h1>Compliance Report</h1>
<p>Mode: $ReportMode | Generated: $(Get-Date -Format 'MMMM dd, yyyy - hh:mm:ss tt')</p>
</div>
<div class="stats">
<div class="stat"><div class="stat-label">Total Servers</div><div class="stat-value">$($Stats.Total)</div></div>
<div class="stat success"><div class="stat-label">Compliant</div><div class="stat-value">$($Stats.Compliant)</div></div>
<div class="stat danger"><div class="stat-label">Non-Compliant</div><div class="stat-value">$($Stats.NonCompliant)</div></div>
<div class="stat warning"><div class="stat-label">Not Found</div><div class="stat-value">$($Stats.NotFound)</div></div>
</div>
<div class="content">
<h2>Server Details</h2>
<table>
<thead>
<tr><th>IP Address</th><th>Server Name</th><th>Overall Status</th><th>Trend Micro</th><th>Trellix</th><th>CrowdStrike</th><th>CloudWatch</th><th>Defender</th><th>Nessus</th><th>Issues</th><th>Source Report</th></tr>
</thead>
<tbody>
"@

    foreach ($Server in $Results) {
        $StatusClass = switch ($Server.OverallStatus) {
            "COMPLIANT" { "status-ok" }
            "NON-COMPLIANT" { "status-fail" }
            default { "status-notfound" }
        }
        
        function Get-Badge($Value) {
            switch ($Value) {
                "OK" { return "<span class='agent-ok'>OK</span>" }
                "FAIL" { return "<span class='agent-fail'>FAIL</span>" }
                "?" { return "<span class='agent-unknown'>?</span>" }
                "-" { return "<span class='agent-na'>-</span>" }
                default { return "<span class='agent-na'>N/A</span>" }
            }
        }
        
        $HTML += "<tr>"
        $HTML += "<td><strong>$($Server.IP)</strong></td>"
        $HTML += "<td>$($Server.ServerName)</td>"
        $HTML += "<td><span class='$StatusClass'>$($Server.OverallStatus)</span></td>"
        $HTML += "<td>$(Get-Badge $Server.TrendMicro)</td>"
        $HTML += "<td>$(Get-Badge $Server.Trellix)</td>"
        $HTML += "<td>$(Get-Badge $Server.CrowdStrike)</td>"
        $HTML += "<td>$(Get-Badge $Server.CloudWatch)</td>"
        $HTML += "<td>$(Get-Badge $Server.Defender)</td>"
        $HTML += "<td>$(Get-Badge $Server.Nessus)</td>"
        $HTML += "<td>$($Server.Issues)</td>"
        $HTML += "<td>$($Server.SourceReport)</td>"
        $HTML += "</tr>"
    }
    
    $HTML += @"
</tbody>
</table>
</div>
<div class="footer">
Enterprise Compliance Scanner | Scan Path: $ScanPath
</div>
</div>
</body>
</html>
"@

    $HTML | Out-File -FilePath $OutputFile -Encoding UTF8 -Force
    
    Write-Host "Report saved: $OutputFile" -ForegroundColor Green
    Write-Host ""
}

$MyServerList = @()
$AllReportData = Scan-Reports -Path $ScanPath

if ($AllReportData.Count -eq 0) {
    Write-Host "ERROR: No data found in reports" -ForegroundColor Red
    exit 1
}

if ($Mode -eq "Compare") {
    $MyServerList = Get-ServerList -FilePath $ServerListFile
    
    if ($MyServerList.Count -eq 0) {
        Write-Host "WARNING: No server list found - switching to FullReport mode" -ForegroundColor Yellow
        Write-Host ""
        $Mode = "FullReport"
    }
}

if ($Mode -eq "Compare") {
    Write-Host "Comparing server list against reports..." -ForegroundColor Yellow
    Write-Host ""
    
    $Results = @()
    $Stats = @{
        Total = $MyServerList.Count
        Compliant = 0
        NonCompliant = 0
        NotFound = 0
    }
    
    foreach ($IP in $MyServerList) {
        if ($AllReportData.ContainsKey($IP)) {
            $ServerData = $AllReportData[$IP]
            $Results += $ServerData
            
            if ($ServerData.OverallStatus -eq "COMPLIANT") {
                $Stats.Compliant++
            } else {
                $Stats.NonCompliant++
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
                Issues = "Not found in any compliance report"
                SourceReport = "N/A"
            }
            $Stats.NotFound++
        }
    }
} else {
    $Results = @()
    foreach ($IP in $AllReportData.Keys) {
        $Results += $AllReportData[$IP]
    }
    
    $Stats = @{
        Total = $Results.Count
        Compliant = ($Results | Where-Object { $_.OverallStatus -eq "COMPLIANT" }).Count
        NonCompliant = ($Results | Where-Object { $_.OverallStatus -eq "NON-COMPLIANT" }).Count
        NotFound = 0
    }
}

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ReportFile = Join-Path $OutputPath "ComplianceReport_${Mode}_$Timestamp.html"

Generate-HTMLReport -Results $Results -Stats $Stats -ReportMode $Mode -OutputFile $ReportFile

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Scan Complete" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Mode: $Mode" -ForegroundColor White
Write-Host "Total: $($Stats.Total) | Compliant: $($Stats.Compliant) | Non-Compliant: $($Stats.NonCompliant) | Not Found: $($Stats.NotFound)" -ForegroundColor White
Write-Host ""

Start-Process $ReportFile
