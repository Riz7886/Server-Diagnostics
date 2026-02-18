$ErrorActionPreference = "Continue"

Clear-Host
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  SINGLE SERVER AGENT REPORT - 100% SAFE" -ForegroundColor Cyan
Write-Host "  NO REMOTE CONNECTIONS - NO CHANGES MADE" -ForegroundColor Cyan
Write-Host "  Created by: Syed Rizvi" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "This script ONLY reads information on THIS server" -ForegroundColor Green
Write-Host "It will NOT install, update, or modify anything" -ForegroundColor Green
Write-Host ""

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$serverName = $env:COMPUTERNAME
$outputFolder = "C:\Temp"
$outputFile = Join-Path $outputFolder "AgentReport_${serverName}_${timestamp}.html"

# Create output folder
try {
    New-Item -ItemType Directory -Path $outputFolder -Force -ErrorAction SilentlyContinue | Out-Null
    Write-Host "Output folder: $outputFolder" -ForegroundColor Cyan
} catch {
    Write-Host "Cannot create output folder, using desktop..." -ForegroundColor Yellow
    $outputFolder = [Environment]::GetFolderPath("Desktop")
    $outputFile = Join-Path $outputFolder "AgentReport_${serverName}_${timestamp}.html"
}

Write-Host ""
Write-Host "Collecting information (READ-ONLY)..." -ForegroundColor Cyan
Write-Host ""

$agentStatus = @()

# Helper function to safely get service
function Get-SafeService {
    param($ServiceName)
    try {
        $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($svc) {
            return @{
                Exists = $true
                Status = $svc.Status.ToString()
                StartType = $svc.StartType.ToString()
            }
        }
    } catch {}
    return @{ Exists = $false }
}

# Helper function to safely get installed software
function Get-SafeSoftware {
    param($Pattern)
    try {
        # Check both 64-bit and 32-bit registry
        $paths = @(
            "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        
        foreach ($path in $paths) {
            $software = Get-ItemProperty $path -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName -like "*$Pattern*" } |
                Select-Object -First 1
            
            if ($software) {
                return @{
                    Exists = $true
                    Name = $software.DisplayName
                    Version = $software.DisplayVersion
                    InstallDate = $software.InstallDate
                    Publisher = $software.Publisher
                }
            }
        }
    } catch {}
    return @{ Exists = $false }
}

# 1. Microsoft Defender
Write-Host "Checking: Microsoft Defender..." -ForegroundColor Yellow
try {
    $defender = Get-MpComputerStatus -ErrorAction SilentlyContinue
    if ($defender) {
        $agentStatus += [PSCustomObject]@{
            AgentName = "Microsoft Defender"
            Status = if ($defender.AntivirusEnabled) { "Running" } else { "Stopped" }
            Version = if ($defender.AMProductVersion) { $defender.AMProductVersion } else { "Unknown" }
            LastUpdate = if ($defender.AntivirusSignatureLastUpdated) { $defender.AntivirusSignatureLastUpdated.ToString("yyyy-MM-dd HH:mm") } else { "Unknown" }
            SignatureAge = if ($defender.AntivirusSignatureAge -ne $null) { "$($defender.AntivirusSignatureAge) days" } else { "Unknown" }
            Compliant = if ($defender.AntivirusSignatureAge -le 7) { "Compliant" } else { "NonCompliant" }
        }
        Write-Host "  Found - Version: $($defender.AMProductVersion)" -ForegroundColor Green
    } else {
        throw "Not available"
    }
} catch {
    $agentStatus += [PSCustomObject]@{
        AgentName = "Microsoft Defender"
        Status = "Not Available"
        Version = "N/A"
        LastUpdate = "N/A"
        SignatureAge = "N/A"
        Compliant = "Not Available"
    }
    Write-Host "  Not available or access denied" -ForegroundColor Red
}

# 2. Nessus Agent
Write-Host "Checking: Nessus Agent..." -ForegroundColor Yellow
$nessusService = Get-SafeService "Nessus Agent"
if (!$nessusService.Exists) { $nessusService = Get-SafeService "NessusAgent" }
$nessusSoftware = Get-SafeSoftware "Nessus Agent"
if (!$nessusSoftware.Exists) { $nessusSoftware = Get-SafeSoftware "Tenable Nessus Agent" }

if ($nessusService.Exists -or $nessusSoftware.Exists) {
    $agentStatus += [PSCustomObject]@{
        AgentName = "Nessus Agent"
        Status = if ($nessusService.Exists) { $nessusService.Status } else { "Unknown" }
        Version = if ($nessusSoftware.Exists) { $nessusSoftware.Version } else { "Unknown" }
        LastUpdate = if ($nessusSoftware.Exists -and $nessusSoftware.InstallDate) { $nessusSoftware.InstallDate } else { "Unknown" }
        SignatureAge = "N/A"
        Compliant = if ($nessusService.Status -eq "Running") { "Compliant" } else { "NonCompliant" }
    }
    Write-Host "  Found - Status: $($nessusService.Status)" -ForegroundColor Green
} else {
    $agentStatus += [PSCustomObject]@{
        AgentName = "Nessus Agent"
        Status = "Not Installed"
        Version = "N/A"
        LastUpdate = "N/A"
        SignatureAge = "N/A"
        Compliant = "Not Installed"
    }
    Write-Host "  Not installed" -ForegroundColor Gray
}

# 3. Trend Micro
Write-Host "Checking: Trend Micro..." -ForegroundColor Yellow
$trendService = Get-SafeService "TMBMServer"
if (!$trendService.Exists) { $trendService = Get-SafeService "TmListen" }
$trendSoftware = Get-SafeSoftware "Trend Micro"

if ($trendService.Exists -or $trendSoftware.Exists) {
    $agentStatus += [PSCustomObject]@{
        AgentName = "Trend Micro"
        Status = if ($trendService.Exists) { $trendService.Status } else { "Unknown" }
        Version = if ($trendSoftware.Exists) { $trendSoftware.Version } else { "Unknown" }
        LastUpdate = if ($trendSoftware.Exists -and $trendSoftware.InstallDate) { $trendSoftware.InstallDate } else { "Unknown" }
        SignatureAge = "N/A"
        Compliant = if ($trendService.Status -eq "Running") { "Compliant" } else { "NonCompliant" }
    }
    Write-Host "  Found - Status: $($trendService.Status)" -ForegroundColor Green
} else {
    $agentStatus += [PSCustomObject]@{
        AgentName = "Trend Micro"
        Status = "Not Installed"
        Version = "N/A"
        LastUpdate = "N/A"
        SignatureAge = "N/A"
        Compliant = "Not Installed"
    }
    Write-Host "  Not installed" -ForegroundColor Gray
}

# 4. CrowdStrike
Write-Host "Checking: CrowdStrike Falcon..." -ForegroundColor Yellow
$crowdService = Get-SafeService "CSFalconService"
if (!$crowdService.Exists) { $crowdService = Get-SafeService "CSAgent" }
$crowdSoftware = Get-SafeSoftware "CrowdStrike"

if ($crowdService.Exists -or $crowdSoftware.Exists) {
    $agentStatus += [PSCustomObject]@{
        AgentName = "CrowdStrike Falcon"
        Status = if ($crowdService.Exists) { $crowdService.Status } else { "Unknown" }
        Version = if ($crowdSoftware.Exists) { $crowdSoftware.Version } else { "Unknown" }
        LastUpdate = if ($crowdSoftware.Exists -and $crowdSoftware.InstallDate) { $crowdSoftware.InstallDate } else { "Unknown" }
        SignatureAge = "N/A"
        Compliant = if ($crowdService.Status -eq "Running") { "Compliant" } else { "NonCompliant" }
    }
    Write-Host "  Found - Status: $($crowdService.Status)" -ForegroundColor Green
} else {
    $agentStatus += [PSCustomObject]@{
        AgentName = "CrowdStrike Falcon"
        Status = "Not Installed"
        Version = "N/A"
        LastUpdate = "N/A"
        SignatureAge = "N/A"
        Compliant = "Not Installed"
    }
    Write-Host "  Not installed" -ForegroundColor Gray
}

# 5. Trellix/McAfee
Write-Host "Checking: Trellix Agent..." -ForegroundColor Yellow
$trellixService = Get-SafeService "masvc"
if (!$trellixService.Exists) { $trellixService = Get-SafeService "McAfeeFramework" }
$trellixSoftware = Get-SafeSoftware "Trellix"
if (!$trellixSoftware.Exists) { $trellixSoftware = Get-SafeSoftware "McAfee" }

if ($trellixService.Exists -or $trellixSoftware.Exists) {
    $agentStatus += [PSCustomObject]@{
        AgentName = "Trellix Agent"
        Status = if ($trellixService.Exists) { $trellixService.Status } else { "Unknown" }
        Version = if ($trellixSoftware.Exists) { $trellixSoftware.Version } else { "Unknown" }
        LastUpdate = if ($trellixSoftware.Exists -and $trellixSoftware.InstallDate) { $trellixSoftware.InstallDate } else { "Unknown" }
        SignatureAge = "N/A"
        Compliant = if ($trellixService.Status -eq "Running") { "Compliant" } else { "NonCompliant" }
    }
    Write-Host "  Found - Status: $($trellixService.Status)" -ForegroundColor Green
} else {
    $agentStatus += [PSCustomObject]@{
        AgentName = "Trellix Agent"
        Status = "Not Installed"
        Version = "N/A"
        LastUpdate = "N/A"
        SignatureAge = "N/A"
        Compliant = "Not Installed"
    }
    Write-Host "  Not installed" -ForegroundColor Gray
}

# 6. Amazon CloudWatch
Write-Host "Checking: Amazon CloudWatch Agent..." -ForegroundColor Yellow
$cloudwatchService = Get-SafeService "AmazonCloudWatchAgent"
$cloudwatchSoftware = Get-SafeSoftware "Amazon CloudWatch Agent"

if ($cloudwatchService.Exists -or $cloudwatchSoftware.Exists) {
    $agentStatus += [PSCustomObject]@{
        AgentName = "Amazon CloudWatch"
        Status = if ($cloudwatchService.Exists) { $cloudwatchService.Status } else { "Unknown" }
        Version = if ($cloudwatchSoftware.Exists) { $cloudwatchSoftware.Version } else { "Unknown" }
        LastUpdate = if ($cloudwatchSoftware.Exists -and $cloudwatchSoftware.InstallDate) { $cloudwatchSoftware.InstallDate } else { "Unknown" }
        SignatureAge = "N/A"
        Compliant = if ($cloudwatchService.Status -eq "Running") { "Compliant" } else { "NonCompliant" }
    }
    Write-Host "  Found - Status: $($cloudwatchService.Status)" -ForegroundColor Green
} else {
    $agentStatus += [PSCustomObject]@{
        AgentName = "Amazon CloudWatch"
        Status = "Not Installed"
        Version = "N/A"
        LastUpdate = "N/A"
        SignatureAge = "N/A"
        Compliant = "Not Installed"
    }
    Write-Host "  Not installed" -ForegroundColor Gray
}

# 7. Centrify
Write-Host "Checking: Centrify Agent..." -ForegroundColor Yellow
$centrifyService = Get-SafeService "CentrifyDC"
$centrifySoftware = Get-SafeSoftware "Centrify"

if ($centrifyService.Exists -or $centrifySoftware.Exists) {
    $agentStatus += [PSCustomObject]@{
        AgentName = "Centrify"
        Status = if ($centrifyService.Exists) { $centrifyService.Status } else { "Unknown" }
        Version = if ($centrifySoftware.Exists) { $centrifySoftware.Version } else { "Unknown" }
        LastUpdate = if ($centrifySoftware.Exists -and $centrifySoftware.InstallDate) { $centrifySoftware.InstallDate } else { "Unknown" }
        SignatureAge = "N/A"
        Compliant = if ($centrifyService.Status -eq "Running") { "Compliant" } else { "NonCompliant" }
    }
    Write-Host "  Found - Status: $($centrifyService.Status)" -ForegroundColor Green
} else {
    $agentStatus += [PSCustomObject]@{
        AgentName = "Centrify"
        Status = "Not Installed"
        Version = "N/A"
        LastUpdate = "N/A"
        SignatureAge = "N/A"
        Compliant = "Not Installed"
    }
    Write-Host "  Not installed" -ForegroundColor Gray
}

Write-Host ""
Write-Host "Generating HTML report..." -ForegroundColor Cyan

$compliantCount = ($agentStatus | Where-Object { $_.Compliant -eq "Compliant" }).Count
$nonCompliantCount = ($agentStatus | Where-Object { $_.Compliant -eq "NonCompliant" }).Count
$notInstalledCount = ($agentStatus | Where-Object { $_.Compliant -eq "Not Installed" -or $_.Compliant -eq "Not Available" }).Count

$html = @"
<!DOCTYPE html>
<html>
<head>
<title>Agent Status Report - $serverName</title>
<style>
body{font-family:Arial;margin:20px;background:#f5f5f5}
.container{max-width:1200px;margin:0 auto;background:white;padding:30px;box-shadow:0 0 10px rgba(0,0,0,0.1)}
.header{background:#0078d4;color:white;padding:30px;margin:-30px -30px 30px -30px}
h1{margin:0;font-size:32px}
.warning{background:#fff3cd;padding:15px;margin:20px 0;border-left:4px solid #ffc107;font-weight:bold}
.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:20px;margin:30px 0}
.stat-box{background:#f8f9fa;padding:20px;text-align:center;border-radius:5px}
.stat-box h2{font-size:36px;margin:10px 0}
.compliant h2{color:#28a745}
.noncompliant h2{color:#dc3545}
.notinstalled h2{color:#6c757d}
table{width:100%;border-collapse:collapse;margin:20px 0}
th{background:#343a40;color:white;padding:12px;text-align:left}
td{padding:12px;border-bottom:1px solid #dee2e6}
tr:nth-child(even){background:#f8f9fa}
.status-running{color:#28a745;font-weight:bold}
.status-stopped{color:#dc3545;font-weight:bold}
.comp{background:#d4edda;color:#155724;padding:5px 10px;border-radius:3px;font-weight:bold}
.noncomp{background:#f8d7da;color:#721c24;padding:5px 10px;border-radius:3px;font-weight:bold}
.notinst{background:#e2e3e5;color:#383d41;padding:5px 10px;border-radius:3px}
</style>
</head>
<body>
<div class="container">
<div class="header">
<h1>Agent Status Report</h1>
<p>Server: $serverName</p>
<p>Report Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
<p>Created by: Syed Rizvi</p>
</div>
<div class="warning">READ-ONLY REPORT - No changes made to this server</div>
<div class="stats">
<div class="stat-box compliant"><h2>$compliantCount</h2><p>Compliant</p></div>
<div class="stat-box noncompliant"><h2>$nonCompliantCount</h2><p>Non-Compliant</p></div>
<div class="stat-box notinstalled"><h2>$notInstalledCount</h2><p>Not Installed</p></div>
</div>
<h2>Agent Details</h2>
<table>
<tr><th>Agent</th><th>Status</th><th>Version</th><th>Last Update</th><th>Signature Age</th><th>Compliance</th></tr>
"@

foreach ($agent in $agentStatus) {
    $statusClass = if ($agent.Status -eq "Running") { "status-running" } else { "" }
    $compClass = switch ($agent.Compliant) {
        "Compliant" { "comp" }
        "NonCompliant" { "noncomp" }
        default { "notinst" }
    }
    
    $html += "<tr><td><strong>$($agent.AgentName)</strong></td><td class='$statusClass'>$($agent.Status)</td><td>$($agent.Version)</td><td>$($agent.LastUpdate)</td><td>$($agent.SignatureAge)</td><td><span class='$compClass'>$($agent.Compliant)</span></td></tr>"
}

$html += @"
</table>
<div style="text-align:center;padding:20px;color:#666;font-size:14px;margin-top:30px;border-top:1px solid #dee2e6">
<p>Agent Status Report - Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
<p>Server: $serverName - Created by: Syed Rizvi</p>
<p><strong>READ-ONLY REPORT - No changes made</strong></p>
</div>
</div>
</body>
</html>
"@

try {
    $html | Out-File -FilePath $outputFile -Encoding UTF8 -Force
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host "  SUCCESS" -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Report saved: $outputFile" -ForegroundColor White
    Write-Host ""
    Write-Host "Summary:" -ForegroundColor Yellow
    Write-Host "  Compliant:     $compliantCount" -ForegroundColor Green
    Write-Host "  Non-Compliant: $nonCompliantCount" -ForegroundColor Red
    Write-Host "  Not Installed: $notInstalledCount" -ForegroundColor Gray
    Write-Host ""
    Write-Host "NO CHANGES WERE MADE TO THIS SERVER" -ForegroundColor Green
    Write-Host ""
    
    Start-Process $outputFile
    
    Write-Host "Report opened in browser" -ForegroundColor Green
    Write-Host ""
} catch {
    Write-Host "ERROR saving report: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Trying alternative location..." -ForegroundColor Yellow
    $altFile = Join-Path ([Environment]::GetFolderPath("Desktop")) "AgentReport_${serverName}_${timestamp}.html"
    $html | Out-File -FilePath $altFile -Encoding UTF8 -Force
    Write-Host "Saved to: $altFile" -ForegroundColor Green
    Start-Process $altFile
}

Write-Host "Script completed successfully" -ForegroundColor Green
Write-Host ""
