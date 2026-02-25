param(
    [switch]$AutoFix,
    [int]$DTUThreshold = 80
)

$ErrorActionPreference = "SilentlyContinue"
$ReportPath = "C:\Temp\SQL_DTU_Reports"
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogFile = Join-Path $ReportPath "DTU_Fix_$timestamp.log"

if (!(Test-Path $ReportPath)) { New-Item -Path $ReportPath -ItemType Directory -Force | Out-Null }

function Write-Log {
    param([string]$Message, [string]$Color = "White")
    $msg = "$(Get-Date -Format 'HH:mm:ss') - $Message"
    Write-Host $msg -ForegroundColor $Color
    $msg | Out-File -FilePath $LogFile -Append
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  MYCARELOOP SQL DTU AUTO-FIX" -ForegroundColor Cyan
Write-Host "  Checking Tony's decreased DTUs" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

Write-Log "Installing required modules..." "Yellow"
$modules = @('Az.Accounts', 'Az.Sql', 'Az.Monitor')
foreach ($mod in $modules) {
    if (!(Get-Module -ListAvailable -Name $mod)) {
        Install-Module -Name $mod -Force -AllowClobber -Scope CurrentUser -ErrorAction SilentlyContinue
    }
    Import-Module $mod -ErrorAction SilentlyContinue
}

Write-Log "Connecting to Azure..." "Yellow"
try {
    $ctx = Get-AzContext
    if (!$ctx) {
        Connect-AzAccount | Out-Null
    }
    Write-Log "  Connected: $($ctx.Account.Id)" "Green"
} catch {
    Write-Log "ERROR: Run Connect-AzAccount first" "Red"
    exit 1
}

Write-Host ""
Write-Log "Scanning for MyCareLoop SQL Databases..." "Yellow"

$servers = Get-AzSqlServer
$issues = @()
$fixed = @()
$totalDBs = 0

foreach ($server in $servers) {
    $dbs = Get-AzSqlDatabase -ServerName $server.ServerName -ResourceGroupName $server.ResourceGroupName | 
           Where-Object { $_.DatabaseName -ne 'master' }
    
    foreach ($db in $dbs) {
        $totalDBs++
        
        if ($db.DatabaseName -notlike "*mycareloop*" -and $db.DatabaseName -notlike "*careloop*") {
            continue
        }
        
        Write-Log "  Checking: $($db.DatabaseName)" "Gray"
        
        $currentDTU = 0
        $currentTier = $db.SkuName
        
        if ($db.CurrentServiceObjectiveName -match '(\d+)') {
            $currentDTU = [int]$matches[1]
        }
        
        $endTime = Get-Date
        $startTime = $endTime.AddHours(-24)
        
        $dtuMetric = Get-AzMetric -ResourceId $db.ResourceId `
                                  -MetricName "dtu_consumption_percent" `
                                  -StartTime $startTime `
                                  -EndTime $endTime `
                                  -TimeGrain 01:00:00 `
                                  -AggregationType Average `
                                  -ErrorAction SilentlyContinue
        
        if ($dtuMetric -and $dtuMetric.Data) {
            $avgDTU = ($dtuMetric.Data | Where-Object { $_.Average } | Measure-Object -Property Average -Average).Average
            $maxDTU = ($dtuMetric.Data | Where-Object { $_.Average } | Measure-Object -Property Average -Maximum).Maximum
            
            if ($avgDTU -gt $DTUThreshold -or $maxDTU -gt 90) {
                $recommendedDTU = $currentDTU
                $action = "NEEDS INCREASE"
                
                if ($maxDTU -gt 95) {
                    $recommendedDTU = $currentDTU * 2
                    $action = "CRITICAL - Double DTUs"
                } elseif ($avgDTU -gt 85) {
                    $recommendedDTU = [math]::Ceiling($currentDTU * 1.5)
                    $action = "HIGH - Increase 50%"
                } elseif ($avgDTU -gt $DTUThreshold) {
                    $recommendedDTU = [math]::Ceiling($currentDTU * 1.25)
                    $action = "MEDIUM - Increase 25%"
                }
                
                $issue = [PSCustomObject]@{
                    Database = $db.DatabaseName
                    Server = $server.ServerName
                    ResourceGroup = $server.ResourceGroupName
                    CurrentDTU = $currentDTU
                    CurrentTier = $currentTier
                    AvgDTUPercent = [math]::Round($avgDTU, 2)
                    MaxDTUPercent = [math]::Round($maxDTU, 2)
                    RecommendedDTU = $recommendedDTU
                    Action = $action
                    ResourceId = $db.ResourceId
                }
                
                $issues += $issue
                
                Write-Log "    ISSUE FOUND!" "Red"
                Write-Log "      Current DTU: $currentDTU" "Yellow"
                Write-Log "      Avg Usage: $([math]::Round($avgDTU, 2))%" "Yellow"
                Write-Log "      Max Usage: $([math]::Round($maxDTU, 2))%" "Red"
                Write-Log "      Recommended: $recommendedDTU DTUs" "Green"
                
                if ($AutoFix) {
                    Write-Log "    AUTO-FIXING..." "Cyan"
                    
                    $newSku = switch ($recommendedDTU) {
                        {$_ -le 10} { "Basic" }
                        {$_ -le 50} { "S2" }
                        {$_ -le 100} { "S3" }
                        {$_ -le 200} { "S6" }
                        {$_ -le 400} { "S9" }
                        default { "S12" }
                    }
                    
                    try {
                        Set-AzSqlDatabase -ResourceGroupName $server.ResourceGroupName `
                                        -ServerName $server.ServerName `
                                        -DatabaseName $db.DatabaseName `
                                        -RequestedServiceObjectiveName $newSku `
                                        -ErrorAction Stop | Out-Null
                        
                        Write-Log "      FIXED! Changed to $newSku" "Green"
                        
                        $fixed += [PSCustomObject]@{
                            Database = $db.DatabaseName
                            OldDTU = $currentDTU
                            NewSku = $newSku
                            NewDTU = $recommendedDTU
                        }
                    } catch {
                        Write-Log "      ERROR: $($_.Exception.Message)" "Red"
                    }
                }
            } else {
                Write-Log "    OK - DTU usage: $([math]::Round($avgDTU, 2))%" "Green"
            }
        }
    }
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Green
Write-Host "  ANALYSIS COMPLETE" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green
Write-Host ""

Write-Host "SUMMARY:" -ForegroundColor Cyan
Write-Host "  Total databases scanned: $totalDBs" -ForegroundColor White
Write-Host "  MyCareLoop databases found: $($issues.Count)" -ForegroundColor White
Write-Host "  Databases needing DTU increase: $($issues.Count)" -ForegroundColor $(if ($issues.Count -gt 0) { "Red" } else { "Green" })

if ($AutoFix) {
    Write-Host "  Databases fixed: $($fixed.Count)" -ForegroundColor Green
}

Write-Host ""

if ($issues.Count -gt 0) {
    Write-Host "DATABASES NEEDING ATTENTION:" -ForegroundColor Yellow
    Write-Host ""
    
    $num = 1
    foreach ($issue in $issues | Sort-Object MaxDTUPercent -Descending) {
        $color = if ($issue.MaxDTUPercent -gt 95) { "Red" } elseif ($issue.MaxDTUPercent -gt 85) { "Yellow" } else { "White" }
        
        Write-Host "$num. $($issue.Database)" -ForegroundColor $color
        Write-Host "   Current: $($issue.CurrentDTU) DTUs ($($issue.CurrentTier))" -ForegroundColor Gray
        Write-Host "   Usage: Avg $($issue.AvgDTUPercent)% | Max $($issue.MaxDTUPercent)%" -ForegroundColor $color
        Write-Host "   Recommended: $($issue.RecommendedDTU) DTUs" -ForegroundColor Green
        Write-Host "   Action: $($issue.Action)" -ForegroundColor Cyan
        Write-Host ""
        $num++
    }
}

if ($AutoFix -and $fixed.Count -gt 0) {
    Write-Host ""
    Write-Host "DATABASES FIXED:" -ForegroundColor Green
    foreach ($fix in $fixed) {
        Write-Host "  $($fix.Database): $($fix.OldDTU) DTUs -> $($fix.NewSku) ($($fix.NewDTU) DTUs)" -ForegroundColor Green
    }
}

$issues | Export-Csv -Path (Join-Path $ReportPath "DTU_Issues_$timestamp.csv") -NoTypeInformation

if ($AutoFix) {
    $fixed | Export-Csv -Path (Join-Path $ReportPath "DTU_Fixed_$timestamp.csv") -NoTypeInformation
}

$htmlReport = @"
<!DOCTYPE html>
<html>
<head>
<style>
body{font-family:Arial;margin:20px;background:#f5f5f5}
.container{max-width:1000px;margin:0 auto;background:white;padding:30px;box-shadow:0 0 10px rgba(0,0,0,0.1)}
h1{color:#dc2626;border-bottom:3px solid #dc2626;padding-bottom:10px}
table{width:100%;border-collapse:collapse;margin:20px 0}
th{background:#dc2626;color:white;padding:12px;text-align:left}
td{padding:10px;border:1px solid #ddd}
tr:nth-child(even){background:#f9f9f9}
.critical{background:#fee2e2;font-weight:bold}
.high{background:#fef3c7}
.summary{background:#fef3c7;border-left:5px solid #f59e0b;padding:20px;margin:20px 0}
</style>
</head>
<body>
<div class="container">
<h1>MyCareLoop SQL DTU Analysis Report</h1>
<p><strong>Report Date:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
<p><strong>Requested By:</strong> Tony Schlak</p>

<div class="summary">
<h2>Summary</h2>
<p>Total Databases Scanned: <strong>$totalDBs</strong></p>
<p>MyCareLoop Databases Found: <strong>$($issues.Count)</strong></p>
<p>Databases Needing DTU Increase: <strong style="color:#dc2626">$($issues.Count)</strong></p>
$(if ($AutoFix) { "<p>Databases Fixed: <strong style='color:#059669'>$($fixed.Count)</strong></p>" } else { "" })
</div>

<h2>Databases Requiring Attention</h2>
<table>
<tr>
<th>Database</th>
<th>Current DTU</th>
<th>Avg Usage %</th>
<th>Max Usage %</th>
<th>Recommended DTU</th>
<th>Action</th>
</tr>
"@

foreach ($issue in ($issues | Sort-Object MaxDTUPercent -Descending)) {
    $rowClass = if ($issue.MaxDTUPercent -gt 95) { "critical" } elseif ($issue.MaxDTUPercent -gt 85) { "high" } else { "" }
    $htmlReport += @"
<tr class="$rowClass">
<td>$($issue.Database)</td>
<td>$($issue.CurrentDTU) ($($issue.CurrentTier))</td>
<td>$($issue.AvgDTUPercent)%</td>
<td>$($issue.MaxDTUPercent)%</td>
<td>$($issue.RecommendedDTU)</td>
<td>$($issue.Action)</td>
</tr>
"@
}

$htmlReport += "</table>"

if ($AutoFix -and $fixed.Count -gt 0) {
    $htmlReport += "<h2>Databases Fixed</h2><table><tr><th>Database</th><th>Old DTU</th><th>New SKU</th><th>New DTU</th></tr>"
    foreach ($fix in $fixed) {
        $htmlReport += "<tr><td>$($fix.Database)</td><td>$($fix.OldDTU)</td><td>$($fix.NewSku)</td><td>$($fix.NewDTU)</td></tr>"
    }
    $htmlReport += "</table>"
}

$htmlReport += @"
<h2>Recommendation</h2>
<p>Tony decreased DTUs to save costs. Based on 24-hour performance metrics, the databases above are experiencing high DTU utilization and may need to be increased to maintain performance.</p>
<p><strong>Command to fix:</strong><br>
<code style="background:#f3f4f6;padding:10px;display:block;margin:10px 0">
.\MyCareLoop-DTU-Fix.ps1 -AutoFix
</code>
</p>
</div>
</body>
</html>
"@

$htmlPath = Join-Path $ReportPath "DTU_Report_$timestamp.html"
$htmlReport | Out-File -FilePath $htmlPath -Encoding UTF8
Start-Process $htmlPath

Write-Host ""
Write-Host "Report saved: $htmlPath" -ForegroundColor Cyan
Write-Host "Log saved: $LogFile" -ForegroundColor Cyan
Write-Host ""

if (!$AutoFix -and $issues.Count -gt 0) {
    Write-Host "TO FIX THESE ISSUES, RUN:" -ForegroundColor Yellow
    Write-Host "  .\MyCareLoop-DTU-Fix.ps1 -AutoFix" -ForegroundColor White
    Write-Host ""
}

Write-Log "Analysis complete. Found $($issues.Count) databases needing attention." "Green"
