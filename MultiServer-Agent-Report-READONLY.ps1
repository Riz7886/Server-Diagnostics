$ErrorActionPreference = "Continue"

Clear-Host
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  MULTI-SERVER AGENT STATUS REPORT - READ-ONLY" -ForegroundColor Cyan
Write-Host "  REMOTE SCANNING - NO CHANGES MADE TO ANY SERVER" -ForegroundColor Cyan
Write-Host "  Created by: Syed Rizvi" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "This script is READ-ONLY and will NOT:" -ForegroundColor Green
Write-Host "  - Install any software" -ForegroundColor Green
Write-Host "  - Update any agents" -ForegroundColor Green
Write-Host "  - Modify any configurations" -ForegroundColor Green
Write-Host "  - Change any settings" -ForegroundColor Green
Write-Host ""
Write-Host "It will remotely query servers and report their status" -ForegroundColor Green
Write-Host ""

$serverList = @(
    "10.133.39.41",
    "10.116.20.98",
    "10.116.52.137",
    "10.174.8.24",
    "10.174.16.13",
    "10.133.7.16",
    "10.133.39.23",
    "10.116.33.62",
    "10.116.21.83",
    "10.116.52.11"
)

Write-Host "Servers to scan:" -ForegroundColor Yellow
foreach ($server in $serverList) {
    Write-Host "  - $server" -ForegroundColor Gray
}
Write-Host ""

$proceed = Read-Host "Proceed with scan? (yes/no)"
if ($proceed -ne "yes") {
    Write-Host "Cancelled" -ForegroundColor Yellow
    exit 0
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFile = "C:\Temp\MultiServer_AgentReport_$timestamp.html"
New-Item -ItemType Directory -Path "C:\Temp" -Force -ErrorAction SilentlyContinue | Out-Null

$allResults = @()

$scriptBlock = {
    $agentStatus = @()
    
    try {
        $defender = Get-MpComputerStatus -ErrorAction SilentlyContinue
        if ($defender) {
            $agentStatus += [PSCustomObject]@{
                AgentName = "Microsoft Defender"
                Status = if ($defender.AntivirusEnabled) { "Running" } else { "Stopped" }
                Version = $defender.AMProductVersion
                LastUpdate = $defender.AntivirusSignatureLastUpdated
                SignatureVersion = $defender.AntivirusSignatureVersion
                Compliant = if ($defender.AntivirusSignatureAge -le 7) { "Compliant" } else { "NonCompliant" }
            }
        }
    } catch {
        $agentStatus += [PSCustomObject]@{
            AgentName = "Microsoft Defender"
            Status = "Not Found"
            Version = "N/A"
            LastUpdate = "N/A"
            SignatureVersion = "N/A"
            Compliant = "Not Installed"
        }
    }
    
    $nessusService = Get-Service -Name "Nessus Agent" -ErrorAction SilentlyContinue
    $nessusSoftware = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like "*Nessus Agent*" } | Select-Object -First 1
    
    if ($nessusService -or $nessusSoftware) {
        $agentStatus += [PSCustomObject]@{
            AgentName = "Nessus Agent"
            Status = if ($nessusService) { $nessusService.Status } else { "Unknown" }
            Version = if ($nessusSoftware) { $nessusSoftware.DisplayVersion } else { "Unknown" }
            LastUpdate = if ($nessusSoftware) { $nessusSoftware.InstallDate } else { "Unknown" }
            SignatureVersion = "N/A"
            Compliant = if ($nessusService -and $nessusService.Status -eq "Running") { "Compliant" } else { "NonCompliant" }
        }
    } else {
        $agentStatus += [PSCustomObject]@{
            AgentName = "Nessus Agent"
            Status = "Not Installed"
            Version = "N/A"
            LastUpdate = "N/A"
            SignatureVersion = "N/A"
            Compliant = "Not Installed"
        }
    }
    
    $trendService = Get-Service -Name "TMBMServer" -ErrorAction SilentlyContinue
    $trendSoftware = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like "*Trend Micro*" } | Select-Object -First 1
    
    if ($trendService -or $trendSoftware) {
        $agentStatus += [PSCustomObject]@{
            AgentName = "Trend Micro"
            Status = if ($trendService) { $trendService.Status } else { "Unknown" }
            Version = if ($trendSoftware) { $trendSoftware.DisplayVersion } else { "Unknown" }
            LastUpdate = if ($trendSoftware) { $trendSoftware.InstallDate } else { "Unknown" }
            SignatureVersion = "N/A"
            Compliant = if ($trendService -and $trendService.Status -eq "Running") { "Compliant" } else { "NonCompliant" }
        }
    } else {
        $agentStatus += [PSCustomObject]@{
            AgentName = "Trend Micro"
            Status = "Not Installed"
            Version = "N/A"
            LastUpdate = "N/A"
            SignatureVersion = "N/A"
            Compliant = "Not Installed"
        }
    }
    
    $crowdService = Get-Service -Name "CSFalconService" -ErrorAction SilentlyContinue
    $crowdSoftware = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like "*CrowdStrike*" } | Select-Object -First 1
    
    if ($crowdService -or $crowdSoftware) {
        $agentStatus += [PSCustomObject]@{
            AgentName = "CrowdStrike Falcon"
            Status = if ($crowdService) { $crowdService.Status } else { "Unknown" }
            Version = if ($crowdSoftware) { $crowdSoftware.DisplayVersion } else { "Unknown" }
            LastUpdate = if ($crowdSoftware) { $crowdSoftware.InstallDate } else { "Unknown" }
            SignatureVersion = "N/A"
            Compliant = if ($crowdService -and $crowdService.Status -eq "Running") { "Compliant" } else { "NonCompliant" }
        }
    } else {
        $agentStatus += [PSCustomObject]@{
            AgentName = "CrowdStrike Falcon"
            Status = "Not Installed"
            Version = "N/A"
            LastUpdate = "N/A"
            SignatureVersion = "N/A"
            Compliant = "Not Installed"
        }
    }
    
    $trellixService = Get-Service -Name "masvc" -ErrorAction SilentlyContinue
    $trellixSoftware = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like "*Trellix*" -or $_.DisplayName -like "*McAfee*" } | Select-Object -First 1
    
    if ($trellixService -or $trellixSoftware) {
        $agentStatus += [PSCustomObject]@{
            AgentName = "Trellix Agent"
            Status = if ($trellixService) { $trellixService.Status } else { "Unknown" }
            Version = if ($trellixSoftware) { $trellixSoftware.DisplayVersion } else { "Unknown" }
            LastUpdate = if ($trellixSoftware) { $trellixSoftware.InstallDate } else { "Unknown" }
            SignatureVersion = "N/A"
            Compliant = if ($trellixService -and $trellixService.Status -eq "Running") { "Compliant" } else { "NonCompliant" }
        }
    } else {
        $agentStatus += [PSCustomObject]@{
            AgentName = "Trellix Agent"
            Status = "Not Installed"
            Version = "N/A"
            LastUpdate = "N/A"
            SignatureVersion = "N/A"
            Compliant = "Not Installed"
        }
    }
    
    $cloudwatchService = Get-Service -Name "AmazonCloudWatchAgent" -ErrorAction SilentlyContinue
    $cloudwatchSoftware = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like "*CloudWatch*" } | Select-Object -First 1
    
    if ($cloudwatchService -or $cloudwatchSoftware) {
        $agentStatus += [PSCustomObject]@{
            AgentName = "Amazon CloudWatch Agent"
            Status = if ($cloudwatchService) { $cloudwatchService.Status } else { "Unknown" }
            Version = if ($cloudwatchSoftware) { $cloudwatchSoftware.DisplayVersion } else { "Unknown" }
            LastUpdate = if ($cloudwatchSoftware) { $cloudwatchSoftware.InstallDate } else { "Unknown" }
            SignatureVersion = "N/A"
            Compliant = if ($cloudwatchService -and $cloudwatchService.Status -eq "Running") { "Compliant" } else { "NonCompliant" }
        }
    } else {
        $agentStatus += [PSCustomObject]@{
            AgentName = "Amazon CloudWatch Agent"
            Status = "Not Installed"
            Version = "N/A"
            LastUpdate = "N/A"
            SignatureVersion = "N/A"
            Compliant = "Not Installed"
        }
    }
    
    $centrifyService = Get-Service -Name "CentrifyDC" -ErrorAction SilentlyContinue
    $centrifySoftware = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like "*Centrify*" } | Select-Object -First 1
    
    if ($centrifyService -or $centrifySoftware) {
        $agentStatus += [PSCustomObject]@{
            AgentName = "Centrify Agent"
            Status = if ($centrifyService) { $centrifyService.Status } else { "Unknown" }
            Version = if ($centrifySoftware) { $centrifySoftware.DisplayVersion } else { "Unknown" }
            LastUpdate = if ($centrifySoftware) { $centrifySoftware.InstallDate } else { "Unknown" }
            SignatureVersion = "N/A"
            Compliant = if ($centrifyService -and $centrifyService.Status -eq "Running") { "Compliant" } else { "NonCompliant" }
        }
    } else {
        $agentStatus += [PSCustomObject]@{
            AgentName = "Centrify Agent"
            Status = "Not Installed"
            Version = "N/A"
            LastUpdate = "N/A"
            SignatureVersion = "N/A"
            Compliant = "Not Installed"
        }
    }
    
    return $agentStatus
}

Write-Host ""
Write-Host "Scanning servers remotely (READ-ONLY)..." -ForegroundColor Cyan
Write-Host ""

foreach ($server in $serverList) {
    Write-Host "Checking: $server..." -ForegroundColor Yellow
    
    try {
        $result = Invoke-Command -ComputerName $server -ScriptBlock $scriptBlock -ErrorAction Stop
        
        foreach ($agent in $result) {
            $allResults += [PSCustomObject]@{
                ServerName = $server
                AgentName = $agent.AgentName
                Status = $agent.Status
                Version = $agent.Version
                LastUpdate = $agent.LastUpdate
                SignatureVersion = $agent.SignatureVersion
                Compliant = $agent.Compliant
            }
        }
        
        Write-Host "  SUCCESS" -ForegroundColor Green
    }
    catch {
        Write-Host "  FAILED: $($_.Exception.Message)" -ForegroundColor Red
        
        $allResults += [PSCustomObject]@{
            ServerName = $server
            AgentName = "Connection Failed"
            Status = "Error"
            Version = "N/A"
            LastUpdate = "N/A"
            SignatureVersion = "N/A"
            Compliant = "Cannot Connect"
        }
    }
}

Write-Host ""
Write-Host "Generating HTML report..." -ForegroundColor Cyan

$html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Multi-Server Agent Status Report</title>
    <style>
        body { font-family: Arial; margin: 20px; background: #f5f5f5; }
        .container { max-width: 1600px; margin: 0 auto; background: white; padding: 30px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
        .header { background: #0078d4; color: white; padding: 30px; margin: -30px -30px 30px -30px; }
        h1 { margin: 0; font-size: 32px; }
        .warning { background: #fff3cd; padding: 15px; margin: 20px 0; border-left: 4px solid #ffc107; font-weight: bold; }
        .stats { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin: 30px 0; }
        .stat-box { background: #f8f9fa; padding: 20px; text-align: center; border-radius: 5px; }
        .stat-box h2 { font-size: 36px; margin: 10px 0; }
        table { width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 13px; }
        th { background: #343a40; color: white; padding: 10px; text-align: left; position: sticky; top: 0; }
        td { padding: 10px; border-bottom: 1px solid #dee2e6; }
        tr:nth-child(even) { background: #f8f9fa; }
        .status-running { color: #28a745; font-weight: bold; }
        .status-stopped { color: #dc3545; font-weight: bold; }
        .compliant { background: #d4edda; color: #155724; padding: 3px 8px; border-radius: 3px; font-weight: bold; font-size: 11px; }
        .noncompliant { background: #f8d7da; color: #721c24; padding: 3px 8px; border-radius: 3px; font-weight: bold; font-size: 11px; }
        .notinstalled { background: #e2e3e5; color: #383d41; padding: 3px 8px; border-radius: 3px; font-size: 11px; }
        .server-section { margin: 30px 0; }
        .server-name { background: #e7f3ff; padding: 15px; margin: 20px 0; border-left: 4px solid #0078d4; font-weight: bold; font-size: 18px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Multi-Server Agent Status Report</h1>
            <p>Report Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
            <p>Servers Scanned: $($serverList.Count)</p>
            <p>Created by: Syed Rizvi</p>
        </div>
        
        <div class="warning">
            READ-ONLY REPORT - No changes were made to any server
        </div>
        
        <div class="stats">
            <div class="stat-box">
                <h2>$($serverList.Count)</h2>
                <p>Servers Scanned</p>
            </div>
            <div class="stat-box">
                <h2>$(($allResults | Where-Object { $_.Compliant -eq "Compliant" }).Count)</h2>
                <p>Compliant Agents</p>
            </div>
            <div class="stat-box">
                <h2>$(($allResults | Where-Object { $_.Compliant -eq "NonCompliant" }).Count)</h2>
                <p>Non-Compliant</p>
            </div>
            <div class="stat-box">
                <h2>$(($allResults | Where-Object { $_.Compliant -eq "Not Installed" }).Count)</h2>
                <p>Not Installed</p>
            </div>
        </div>
        
        <h2>All Servers - Agent Status</h2>
        <table>
            <tr>
                <th>Server</th>
                <th>Agent Name</th>
                <th>Status</th>
                <th>Version</th>
                <th>Last Update</th>
                <th>Signature Version</th>
                <th>Compliance</th>
            </tr>
"@

foreach ($result in $allResults) {
    $statusClass = if ($result.Status -eq "Running") { "status-running" } else { "status-stopped" }
    $complianceClass = switch ($result.Compliant) {
        "Compliant" { "compliant" }
        "NonCompliant" { "noncompliant" }
        default { "notinstalled" }
    }
    
    $html += @"
            <tr>
                <td><strong>$($result.ServerName)</strong></td>
                <td>$($result.AgentName)</td>
                <td class="$statusClass">$($result.Status)</td>
                <td>$($result.Version)</td>
                <td>$($result.LastUpdate)</td>
                <td>$($result.SignatureVersion)</td>
                <td><span class="$complianceClass">$($result.Compliant)</span></td>
            </tr>
"@
}

$html += @"
        </table>
        
        <div style="text-align: center; padding: 20px; color: #666; font-size: 14px; margin-top: 30px; border-top: 1px solid #dee2e6;">
            <p>Multi-Server Agent Status Report</p>
            <p>Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
            <p>Created by: Syed Rizvi - SAP Cloud Infrastructure Team</p>
            <p><strong>This is a READ-ONLY report - No changes were made to any server</strong></p>
        </div>
    </div>
</body>
</html>
"@

$html | Out-File -FilePath $outputFile -Encoding UTF8

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  REPORT GENERATED SUCCESSFULLY" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Report saved to: $outputFile" -ForegroundColor White
Write-Host ""
Write-Host "Servers scanned: $($serverList.Count)" -ForegroundColor Yellow
Write-Host "Agents checked per server: 7" -ForegroundColor Yellow
Write-Host ""
Write-Host "NO CHANGES WERE MADE TO ANY SERVER" -ForegroundColor Green
Write-Host ""
Write-Host "Opening report..." -ForegroundColor Cyan
Start-Process $outputFile

Write-Host ""
Write-Host "DONE - 100% READ-ONLY - ZERO MODIFICATIONS" -ForegroundColor Green
Write-Host ""
