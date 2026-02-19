<#
.SYNOPSIS
    Nessus Plugin 65057 - SAP Insecure Windows Service Permissions Assessment
    
.DESCRIPTION
    Scans SAP sidecar servers for Plugin 65057 vulnerabilities, determines if patch 
    or permission fix is needed, and generates HTML report with remediation.
    
.NOTES
    Author: Security Remediation Team
    Date: 2026-02-19
    Environment: DOD FedRAMP - AWS GovCloud
    Ticket: INC0135584
    Purpose: SAP Sidecar Service Permission Analysis
    
.REQUIREMENTS
    - Run as Administrator
    - Network access to target servers
    - Domain credentials with admin rights
#>

#Requires -Version 5.1
#Requires -RunAsAdministrator

# ============================================================================
# CONFIGURATION
# ============================================================================

$Config = @{
    # Target SAP sidecar servers from Plugin 65057 findings
    TargetServers = @(
        "10.134.4.171",
        "10.134.4.153",
        "10.134.4.254",
        "10.134.4.109",
        "10.134.4.247",
        "10.134.4.6"
    )
    
    OutputPath = "C:\SecurityRemediation\Plugin65057_SAP"
    ReportName = "SAP_Plugin65057_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    LogName = "SAP_Plugin65057_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    
    # Vulnerable identities
    VulnerableIdentities = @(
        "BUILTIN\Users",
        "Everyone", 
        "NT AUTHORITY\Authenticated Users",
        "BUILTIN\Power Users",
        "NT AUTHORITY\INTERACTIVE"
    )
    
    # Dangerous permissions
    DangerousRights = @(
        "FullControl",
        "Modify",
        "Write",
        "WriteData",
        "AppendData",
        "WriteExtendedAttributes",
        "Delete",
        "ChangePermissions",
        "TakeOwnership"
    )
    
    # SAP-specific service patterns
    SAPServicePatterns = @(
        "*SAP*",
        "*sap*",
        "*sidecar*",
        "*SIDECAR*"
    )
}

# ============================================================================
# LOGGING FUNCTIONS
# ============================================================================

if (-not (Test-Path $Config.OutputPath)) {
    New-Item -ItemType Directory -Path $Config.OutputPath -Force | Out-Null
}

$Script:LogFile = Join-Path $Config.OutputPath $Config.LogName
$Script:ReportFile = Join-Path $Config.OutputPath $Config.ReportName

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARNING','ERROR','SUCCESS')]
        [string]$Level = 'INFO'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Level] $Message"
    
    Add-Content -Path $Script:LogFile -Value $logMessage
    
    switch ($Level) {
        'ERROR'   { Write-Host $logMessage -ForegroundColor Red }
        'WARNING' { Write-Host $logMessage -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $logMessage -ForegroundColor Green }
        default   { Write-Host $logMessage -ForegroundColor White }
    }
}

# ============================================================================
# CREDENTIAL FUNCTION
# ============================================================================

function Get-RemoteCredential {
    Write-Host "`n╔══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║          CREDENTIAL AUTHENTICATION                  ║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan
    
    Write-Host "Enter credentials with administrative access to SAP servers" -ForegroundColor Yellow
    Write-Host "Format: DOMAIN\Username`n" -ForegroundColor Gray
    
    $credential = Get-Credential -Message "Enter Domain\Username and Password"
    
    if (-not $credential) {
        Write-Log "No credentials provided. Exiting." -Level ERROR
        exit 1
    }
    
    Write-Log "Credentials obtained for: $($credential.UserName)" -Level SUCCESS
    return $credential
}

# ============================================================================
# CONNECTIVITY TEST
# ============================================================================

function Test-ServerConnectivity {
    param(
        [string]$ServerName,
        [PSCredential]$Credential
    )
    
    Write-Log "Testing connectivity to $ServerName"
    
    # Ping test
    if (-not (Test-Connection -ComputerName $ServerName -Count 2 -Quiet)) {
        Write-Log "Ping failed to $ServerName" -Level ERROR
        return $false
    }
    
    # WinRM test
    try {
        $testResult = Invoke-Command -ComputerName $ServerName -Credential $Credential -ScriptBlock {
            $env:COMPUTERNAME
        } -ErrorAction Stop
        
        Write-Log "WinRM connection successful to $ServerName" -Level SUCCESS
        return $true
    } catch {
        Write-Log "WinRM connection failed to $ServerName : $_" -Level ERROR
        return $false
    }
}

# ============================================================================
# PATCH STATUS CHECK
# ============================================================================

function Get-PatchStatus {
    param(
        [string]$ServerName,
        [PSCredential]$Credential
    )
    
    Write-Log "Checking patch status on $ServerName"
    
    try {
        $patchInfo = Invoke-Command -ComputerName $ServerName -Credential $Credential -ScriptBlock {
            $os = Get-CimInstance Win32_OperatingSystem
            $lastUpdate = Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 1
            
            $recentUpdates = Get-HotFix | Where-Object {
                $_.Description -match "Security|Update" -and
                $_.InstalledOn -gt (Get-Date).AddDays(-90)
            }
            
            $pendingReboot = $false
            if (Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") {
                $pendingReboot = $true
            }
            if (Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") {
                $pendingReboot = $true
            }
            
            [PSCustomObject]@{
                ComputerName = $env:COMPUTERNAME
                OSVersion = $os.Version
                OSBuild = $os.BuildNumber
                OSCaption = $os.Caption
                LastUpdate = $lastUpdate.InstalledOn
                LastUpdateKB = $lastUpdate.HotFixID
                RecentUpdateCount = $recentUpdates.Count
                PendingReboot = $pendingReboot
                LastBootTime = $os.LastBootUpTime
            }
        } -ErrorAction Stop
        
        Write-Log "Patch status retrieved for $ServerName" -Level SUCCESS
        return $patchInfo
        
    } catch {
        Write-Log "Failed to get patch status on $ServerName : $_" -Level ERROR
        return $null
    }
}

# ============================================================================
# SERVICE PERMISSION ANALYSIS
# ============================================================================

function Get-ServicePermissionVulnerabilities {
    param(
        [string]$ServerName,
        [PSCredential]$Credential
    )
    
    Write-Log "Analyzing service permissions on $ServerName"
    
    try {
        $results = Invoke-Command -ComputerName $ServerName -Credential $Credential -ScriptBlock {
            param($VulnIdentities, $DangerousRights, $SAPPatterns)
            
            $findings = @()
            $services = Get-WmiObject Win32_Service
            
            foreach ($service in $services) {
                try {
                    $svcName = $service.Name
                    $svcKey = "HKLM:\SYSTEM\CurrentControlSet\Services\$svcName"
                    
                    if (-not (Test-Path $svcKey)) { continue }
                    
                    $acl = Get-Acl $svcKey -ErrorAction Stop
                    
                    # Check if SAP-related service
                    $isSAPService = $false
                    foreach ($pattern in $SAPPatterns) {
                        if ($svcName -like $pattern -or $service.DisplayName -like $pattern) {
                            $isSAPService = $true
                            break
                        }
                    }
                    
                    foreach ($ace in $acl.Access) {
                        # Check vulnerable identity
                        $isVulnIdentity = $false
                        foreach ($vulnId in $VulnIdentities) {
                            if ($ace.IdentityReference.Value -like "*$vulnId*" -or 
                                $ace.IdentityReference.Value -eq $vulnId) {
                                $isVulnIdentity = $true
                                break
                            }
                        }
                        
                        if ($isVulnIdentity) {
                            # Check dangerous rights
                            $rights = $ace.RegistryRights.ToString()
                            $hasDangerousRights = $false
                            
                            foreach ($dangerousRight in $DangerousRights) {
                                if ($rights -match $dangerousRight) {
                                    $hasDangerousRights = $true
                                    break
                                }
                            }
                            
                            if ($hasDangerousRights) {
                                $findings += [PSCustomObject]@{
                                    ServiceName = $svcName
                                    DisplayName = $service.DisplayName
                                    StartMode = $service.StartMode
                                    State = $service.State
                                    ServiceAccount = $service.StartName
                                    PathName = $service.PathName
                                    VulnerableIdentity = $ace.IdentityReference.Value
                                    DangerousRights = $rights
                                    AccessType = $ace.AccessControlType
                                    IsInherited = $ace.IsInherited
                                    IsSAPService = $isSAPService
                                    Severity = if($isSAPService){"CRITICAL"}else{"HIGH"}
                                }
                            }
                        }
                    }
                } catch {
                    continue
                }
            }
            
            return $findings
            
        } -ArgumentList $Config.VulnerableIdentities, $Config.DangerousRights, $Config.SAPServicePatterns -ErrorAction Stop
        
        $sapCount = ($results | Where-Object IsSAPService -eq $true).Count
        $totalCount = $results.Count
        
        Write-Log "Found $totalCount vulnerable services on $ServerName ($sapCount are SAP-related)" -Level $(if($totalCount -gt 0){'WARNING'}else{'SUCCESS'})
        
        return $results
        
    } catch {
        Write-Log "Failed to analyze permissions on $ServerName : $_" -Level ERROR
        return @()
    }
}

# ============================================================================
# DETERMINE REMEDIATION TYPE
# ============================================================================

function Get-RemediationStrategy {
    param(
        [object]$PatchStatus,
        [array]$Vulnerabilities
    )
    
    $needsPatch = $false
    $needsPermFix = $false
    $reasons = @()
    
    # Check patch status
    if ($PatchStatus) {
        if ($PatchStatus.LastUpdate) {
            $daysSinceUpdate = ((Get-Date) - $PatchStatus.LastUpdate).Days
            
            if ($daysSinceUpdate -gt 60) {
                $needsPatch = $true
                $reasons += "Last Windows update: $daysSinceUpdate days ago (KB: $($PatchStatus.LastUpdateKB))"
            }
        }
        
        if ($PatchStatus.PendingReboot) {
            $needsPatch = $true
            $reasons += "Pending reboot detected - installed updates not applied"
        }
        
        if ($PatchStatus.RecentUpdateCount -lt 3) {
            $needsPatch = $true
            $reasons += "Only $($PatchStatus.RecentUpdateCount) updates in last 90 days"
        }
    }
    
    # Check vulnerabilities
    if ($Vulnerabilities.Count -gt 0) {
        $needsPermFix = $true
        $sapCount = ($Vulnerabilities | Where-Object IsSAPService -eq $true).Count
        
        if ($sapCount -gt 0) {
            $reasons += "Found $sapCount SAP service(s) with insecure permissions [CRITICAL]"
        }
        
        $nonSapCount = $Vulnerabilities.Count - $sapCount
        if ($nonSapCount -gt 0) {
            $reasons += "Found $nonSapCount non-SAP service(s) with insecure permissions"
        }
        
        $inheritedCount = ($Vulnerabilities | Where-Object IsInherited -eq $true).Count
        if ($inheritedCount -gt 0) {
            $reasons += "$inheritedCount inherited permissions (may need GPO review)"
        }
    }
    
    # Determine action
    $action = if ($needsPatch -and $needsPermFix) {
        "PATCH_AND_PERMISSIONS"
    } elseif ($needsPatch) {
        "PATCH_ONLY"
    } elseif ($needsPermFix) {
        "PERMISSIONS_ONLY"
    } else {
        "NO_ACTION"
    }
    
    return [PSCustomObject]@{
        NeedsPatch = $needsPatch
        NeedsPermissionFix = $needsPermFix
        Reasons = $reasons
        Action = $action
        Priority = if(($Vulnerabilities | Where-Object IsSAPService -eq $true).Count -gt 0){"CRITICAL"}else{"HIGH"}
    }
}

# ============================================================================
# HTML REPORT GENERATION
# ============================================================================

function New-HTMLReport {
    param([array]$Results)
    
    Write-Log "Generating HTML report"
    
    $totalServers = $Results.Count
    $vulnServers = ($Results | Where-Object {$_.Vulnerabilities.Count -gt 0}).Count
    $totalVulns = ($Results | ForEach-Object {$_.Vulnerabilities.Count} | Measure-Object -Sum).Sum
    $sapVulns = ($Results | ForEach-Object {$_.Vulnerabilities | Where-Object IsSAPService -eq $true}).Count
    $patchNeeded = ($Results | Where-Object {$_.Remediation.NeedsPatch}).Count
    $permNeeded = ($Results | Where-Object {$_.Remediation.NeedsPermissionFix}).Count
    
$html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>SAP Plugin 65057 Assessment Report</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            background: #f0f2f5;
            padding: 20px;
            line-height: 1.6;
        }
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            border-radius: 8px;
            margin-bottom: 25px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .header h1 {
            font-size: 28px;
            margin-bottom: 10px;
        }
        .header p {
            opacity: 0.9;
            font-size: 14px;
        }
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 25px;
        }
        .summary-card {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.08);
            border-left: 4px solid #667eea;
        }
        .summary-card h3 {
            color: #555;
            font-size: 14px;
            margin-bottom: 10px;
            text-transform: uppercase;
        }
        .summary-card .value {
            font-size: 32px;
            font-weight: bold;
            color: #333;
        }
        .summary-card.critical { border-left-color: #dc3545; }
        .summary-card.critical .value { color: #dc3545; }
        .summary-card.warning { border-left-color: #ffc107; }
        .summary-card.warning .value { color: #ff8c00; }
        .summary-card.success { border-left-color: #28a745; }
        .summary-card.success .value { color: #28a745; }
        
        .server-section {
            background: white;
            padding: 25px;
            border-radius: 8px;
            margin-bottom: 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.08);
        }
        .server-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            padding-bottom: 15px;
            border-bottom: 2px solid #f0f2f5;
        }
        .server-header h2 {
            color: #333;
            font-size: 22px;
        }
        .badge {
            display: inline-block;
            padding: 6px 12px;
            border-radius: 4px;
            font-size: 12px;
            font-weight: bold;
            text-transform: uppercase;
        }
        .badge-critical { background: #dc3545; color: white; }
        .badge-warning { background: #ffc107; color: #333; }
        .badge-success { background: #28a745; color: white; }
        .badge-info { background: #17a2b8; color: white; }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 15px 0;
            font-size: 14px;
        }
        th {
            background: #667eea;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: 600;
        }
        td {
            padding: 10px 12px;
            border-bottom: 1px solid #e9ecef;
        }
        tr:hover {
            background: #f8f9fa;
        }
        .highlight-sap {
            background: #fff3cd !important;
            font-weight: bold;
        }
        
        .info-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin: 15px 0;
        }
        .info-item {
            padding: 10px;
            background: #f8f9fa;
            border-radius: 4px;
        }
        .info-item label {
            display: block;
            font-size: 12px;
            color: #666;
            margin-bottom: 5px;
        }
        .info-item .value {
            font-size: 14px;
            font-weight: 600;
            color: #333;
        }
        
        .remediation-box {
            background: #fff3cd;
            border-left: 4px solid #ffc107;
            padding: 20px;
            margin: 20px 0;
            border-radius: 4px;
        }
        .remediation-box h3 {
            color: #856404;
            margin-bottom: 15px;
        }
        .remediation-box ul {
            margin-left: 20px;
        }
        .remediation-box li {
            margin: 8px 0;
            color: #856404;
        }
        
        .code-block {
            background: #1e1e1e;
            color: #d4d4d4;
            padding: 20px;
            border-radius: 6px;
            overflow-x: auto;
            margin: 15px 0;
            font-family: 'Courier New', monospace;
            font-size: 13px;
            line-height: 1.5;
        }
        .code-block pre {
            margin: 0;
        }
        
        .alert {
            padding: 15px;
            border-radius: 4px;
            margin: 15px 0;
        }
        .alert-danger {
            background: #f8d7da;
            border-left: 4px solid #dc3545;
            color: #721c24;
        }
        .alert-warning {
            background: #fff3cd;
            border-left: 4px solid #ffc107;
            color: #856404;
        }
        .alert-success {
            background: #d4edda;
            border-left: 4px solid #28a745;
            color: #155724;
        }
        
        .next-steps {
            background: white;
            padding: 25px;
            border-radius: 8px;
            margin: 20px 0;
            box-shadow: 0 2px 4px rgba(0,0,0,0.08);
        }
        .next-steps h2 {
            color: #333;
            margin-bottom: 20px;
        }
        .next-steps ol {
            margin-left: 25px;
        }
        .next-steps li {
            margin: 12px 0;
            line-height: 1.8;
        }
        
        .footer {
            text-align: center;
            padding: 20px;
            color: #666;
            font-size: 12px;
        }
        
        @media print {
            body { background: white; }
            .server-section { page-break-inside: avoid; }
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>🔒 SAP Security Assessment - Nessus Plugin 65057</h1>
        <p><strong>Insecure Windows Service Permissions Analysis</strong></p>
        <p>Environment: DOD FedRAMP AWS GovCloud | Ticket: INC0135584 | Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    </div>
    
    <div class="summary-grid">
        <div class="summary-card">
            <h3>Total Servers Scanned</h3>
            <div class="value">$totalServers</div>
        </div>
        <div class="summary-card critical">
            <h3>Servers with Vulnerabilities</h3>
            <div class="value">$vulnServers</div>
        </div>
        <div class="summary-card warning">
            <h3>Total Vulnerable Services</h3>
            <div class="value">$totalVulns</div>
        </div>
        <div class="summary-card critical">
            <h3>SAP Services Affected</h3>
            <div class="value">$sapVulns</div>
        </div>
        <div class="summary-card warning">
            <h3>Servers Need Patching</h3>
            <div class="value">$patchNeeded</div>
        </div>
        <div class="summary-card warning">
            <h3>Servers Need Perm Fix</h3>
            <div class="value">$permNeeded</div>
        </div>
    </div>
"@

    foreach ($result in $Results) {
        $statusBadge = switch ($result.Remediation.Action) {
            "PATCH_AND_PERMISSIONS" { "<span class='badge badge-critical'>⚠️ PATCH + PERMISSIONS</span>" }
            "PATCH_ONLY" { "<span class='badge badge-warning'>🔄 PATCH REQUIRED</span>" }
            "PERMISSIONS_ONLY" { "<span class='badge badge-warning'>🔧 PERMISSIONS FIX</span>" }
            "NO_ACTION" { "<span class='badge badge-success'>✓ NO ACTION NEEDED</span>" }
            default { "<span class='badge badge-info'>UNKNOWN</span>" }
        }
        
        if ($result.Remediation.Priority -eq "CRITICAL") {
            $statusBadge += " <span class='badge badge-critical'>🔴 CRITICAL - SAP AFFECTED</span>"
        }
        
        $html += @"
    <div class="server-section">
        <div class="server-header">
            <h2>🖥️ $($result.ServerName)</h2>
            <div>$statusBadge</div>
        </div>
"@
        
        if ($result.Status -eq "UNREACHABLE") {
            $html += @"
        <div class="alert alert-danger">
            <strong>⚠️ Server Unreachable</strong><br>
            Unable to connect to this server. Please verify network connectivity and credentials.
        </div>
"@
        } else {
            # System Info
            if ($result.PatchStatus) {
                $html += @"
        <h3>📋 System Information</h3>
        <div class="info-grid">
            <div class="info-item">
                <label>Operating System</label>
                <div class="value">$($result.PatchStatus.OSCaption)</div>
            </div>
            <div class="info-item">
                <label>OS Build</label>
                <div class="value">$($result.PatchStatus.OSBuild)</div>
            </div>
            <div class="info-item">
                <label>Last Update</label>
                <div class="value">$($result.PatchStatus.LastUpdate)</div>
            </div>
            <div class="info-item">
                <label>Last Update KB</label>
                <div class="value">$($result.PatchStatus.LastUpdateKB)</div>
            </div>
            <div class="info-item">
                <label>Recent Updates (90d)</label>
                <div class="value">$($result.PatchStatus.RecentUpdateCount)</div>
            </div>
            <div class="info-item">
                <label>Pending Reboot</label>
                <div class="value" style="color: $(if($result.PatchStatus.PendingReboot){'#dc3545'}else{'#28a745'});">
                    $(if($result.PatchStatus.PendingReboot){'YES ⚠️'}else{'NO ✓'})
                </div>
            </div>
        </div>
"@
            }
            
            # Vulnerabilities
            $html += @"
        <h3>🔍 Vulnerability Analysis</h3>
        <p><strong>Total Vulnerable Services:</strong> 
           <span style="color: $(if($result.Vulnerabilities.Count -gt 0){'#dc3545'}else{'#28a745'}); font-weight: bold;">
               $($result.Vulnerabilities.Count)
           </span>
        </p>
"@
            
            if ($result.Vulnerabilities.Count -gt 0) {
                $sapVulns = $result.Vulnerabilities | Where-Object IsSAPService -eq $true
                if ($sapVulns.Count -gt 0) {
                    $html += "<p style='color: #dc3545; font-weight: bold;'>⚠️ $($sapVulns.Count) SAP service(s) affected - CRITICAL PRIORITY</p>"
                }
                
                $html += @"
        <table>
            <thead>
                <tr>
                    <th>Service Name</th>
                    <th>Display Name</th>
                    <th>State</th>
                    <th>Vulnerable Identity</th>
                    <th>Dangerous Rights</th>
                    <th>Type</th>
                </tr>
            </thead>
            <tbody>
"@
                foreach ($vuln in $result.Vulnerabilities) {
                    $rowClass = if ($vuln.IsSAPService) { "highlight-sap" } else { "" }
                    $typeLabel = if ($vuln.IsSAPService) { "🔴 SAP" } else { "Standard" }
                    
                    $html += @"
                <tr class="$rowClass">
                    <td><code>$($vuln.ServiceName)</code></td>
                    <td>$($vuln.DisplayName)</td>
                    <td>$($vuln.State)</td>
                    <td style="color: #dc3545; font-weight: bold;">$($vuln.VulnerableIdentity)</td>
                    <td style="color: #dc3545;">$($vuln.DangerousRights)</td>
                    <td>$typeLabel</td>
                </tr>
"@
                }
                $html += @"
            </tbody>
        </table>
"@
            }
            
            # Remediation
            $html += @"
        <div class="remediation-box">
            <h3>🔧 Remediation Required</h3>
            <p><strong>Action Plan:</strong> $($result.Remediation.Action)</p>
            <p><strong>Priority Level:</strong> $($result.Remediation.Priority)</p>
            <ul>
"@
            foreach ($reason in $result.Remediation.Reasons) {
                $html += "<li>$reason</li>"
            }
            $html += @"
            </ul>
        </div>
"@
            
            # Permission Fix Script
            if ($result.Remediation.NeedsPermissionFix -and $result.Vulnerabilities.Count -gt 0) {
                $serviceList = ($result.Vulnerabilities | Select-Object -ExpandProperty ServiceName -Unique | ForEach-Object { "        `"$_`"" }) -join ",`n"
                
                $html += @"
        <h3>📝 Permission Remediation Script for $($result.ServerName)</h3>
        <div class="alert alert-warning">
            <strong>⚠️ Important:</strong> Test this script in non-production first. Create a change request before running in production.
        </div>
        <div class="code-block">
<pre># ============================================================================
# SAP Service Permission Remediation Script
# Server: $($result.ServerName)
# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
# ============================================================================

#Requires -RunAsAdministrator

# Services to remediate
`$services = @(
$serviceList
)

# Create backup directory
`$backupPath = "C:\SecurityBackups\SAP_ServicePerms_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -ItemType Directory -Path `$backupPath -Force | Out-Null
Write-Host "Backup location: `$backupPath`n" -ForegroundColor Cyan

# Remediation log
`$logFile = Join-Path `$backupPath "remediation.log"

function Write-RemediationLog {
    param([string]`$Message, [string]`$Level = "INFO")
    `$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    `$logMsg = "[`$timestamp] [`$Level] `$Message"
    Add-Content -Path `$logFile -Value `$logMsg
    
    switch (`$Level) {
        "ERROR"   { Write-Host `$logMsg -ForegroundColor Red }
        "WARNING" { Write-Host `$logMsg -ForegroundColor Yellow }
        "SUCCESS" { Write-Host `$logMsg -ForegroundColor Green }
        default   { Write-Host `$logMsg -ForegroundColor White }
    }
}

Write-Host "=== SAP Service Permission Remediation ===" -ForegroundColor Cyan
Write-Host "Server: $($result.ServerName)" -ForegroundColor White
Write-Host "Services to fix: `$(`$services.Count)`n" -ForegroundColor White

`$successCount = 0
`$failCount = 0

foreach (`$serviceName in `$services) {
    try {
        `$serviceKey = "HKLM:\SYSTEM\CurrentControlSet\Services\`$serviceName"
        
        if (-not (Test-Path `$serviceKey)) {
            Write-RemediationLog "Service key not found: `$serviceName" -Level WARNING
            continue
        }
        
        Write-Host "Processing: `$serviceName" -ForegroundColor Cyan
        
        # Backup current ACL
        `$acl = Get-Acl `$serviceKey
        `$acl | Export-Clixml (Join-Path `$backupPath "`$serviceName-acl-backup.xml")
        Write-RemediationLog "Backed up ACL for `$serviceName"
        
        # Remove vulnerable permissions
        `$removed = 0
        `$acl.Access | Where-Object {
            (`$_.IdentityReference -match "Users|Everyone|Authenticated Users|Power Users|INTERACTIVE") -and
            (`$_.RegistryRights -match "FullControl|Modify|Write|Delete|ChangePermissions|TakeOwnership")
        } | ForEach-Object {
            `$acl.RemoveAccessRule(`$_) | Out-Null
            Write-RemediationLog "  Removed: `$(`$_.IdentityReference) - `$(`$_.RegistryRights)" -Level WARNING
            `$removed++
        }
        
        if (`$removed -gt 0) {
            # Apply corrected ACL
            Set-Acl -Path `$serviceKey -AclObject `$acl
            Write-Host "  ✓ Fixed `$serviceName - Removed `$removed vulnerable permission(s)" -ForegroundColor Green
            Write-RemediationLog "Successfully remediated `$serviceName" -Level SUCCESS
            `$successCount++
        } else {
            Write-Host "  ℹ `$serviceName - No vulnerable permissions found" -ForegroundColor Gray
            Write-RemediationLog "No action needed for `$serviceName"
        }
        
    } catch {
        Write-Host "  ✗ Failed: `$serviceName - `$_" -ForegroundColor Red
        Write-RemediationLog "Failed to remediate `$serviceName : `$_" -Level ERROR
        `$failCount++
    }
}

Write-Host "`n=== Remediation Complete ===" -ForegroundColor Cyan
Write-Host "Successfully remediated: `$successCount" -ForegroundColor Green
Write-Host "Failed: `$failCount" -ForegroundColor $(if(`$failCount -gt 0){'Red'}else{'Green'})
Write-Host "Backup location: `$backupPath" -ForegroundColor Cyan
Write-Host "Log file: `$logFile`n" -ForegroundColor Cyan

# Verify services still work
Write-Host "Verifying service status..." -ForegroundColor Cyan
foreach (`$serviceName in `$services) {
    try {
        `$svc = Get-Service `$serviceName -ErrorAction Stop
        if (`$svc.StartType -eq 'Automatic' -and `$svc.Status -ne 'Running') {
            Write-Host "  ⚠️  `$serviceName is set to Automatic but not running" -ForegroundColor Yellow
        } else {
            Write-Host "  ✓ `$serviceName - Status: `$(`$svc.Status)" -ForegroundColor Green
        }
    } catch {
        Write-Host "  ⚠️  Could not verify `$serviceName" -ForegroundColor Yellow
    }
}

Write-Host "`n✓ Script execution completed" -ForegroundColor Green
Write-Host "Please test SAP functionality before closing this ticket.`n" -ForegroundColor Yellow</pre>
        </div>
"@
            }
            
            # Patching Instructions
            if ($result.Remediation.NeedsPatch) {
                $html += @"
        <h3>🔄 Windows Update Required</h3>
        <div class="alert alert-warning">
            <strong>Action Required:</strong> Install pending Windows updates
        </div>
        <div class="code-block">
<pre># Check for available updates
Install-Module PSWindowsUpdate -Force
Get-WindowsUpdate

# Install all updates (requires approval)
Install-WindowsUpdate -AcceptAll -AutoReboot

# Or manual via WSUS/SCCM based on your org policy</pre>
        </div>
"@
            }
        }
        
        $html += "</div>" # Close server section
    }
    
    # Next Steps
    $html += @"
    <div class="next-steps">
        <h2>📋 Implementation Plan</h2>
        <ol>
            <li><strong>Review Report:</strong> Prioritize servers with SAP service vulnerabilities (marked in yellow)</li>
            <li><strong>Test Scripts:</strong> Run remediation scripts in non-production environment first</li>
            <li><strong>Create Change Request:</strong> Document changes for ticket INC0135584</li>
            <li><strong>Coordinate with SAP Team:</strong> Notify SAP administrators before making changes</li>
            <li><strong>Schedule Maintenance:</strong> Plan maintenance window for production servers</li>
            <li><strong>Execute Remediation:</strong> Run scripts during approved maintenance window</li>
            <li><strong>Verify Functionality:</strong> Test all SAP services after remediation</li>
            <li><strong>Rescan with Nessus:</strong> Verify Plugin 65057 is resolved (wait 24-48 hours)</li>
            <li><strong>Update Ticket:</strong> Close INC0135584 with documentation</li>
        </ol>
    </div>
    
    <div class="next-steps">
        <h2>⚠️ Critical Notes for DOD FedRAMP Environment</h2>
        <ul style="margin-left: 25px; line-height: 2;">
            <li>All changes must follow your organization's change management process</li>
            <li>Backup ACLs are created automatically by remediation scripts</li>
            <li>SAP services (marked in yellow) are CRITICAL - coordinate with SAP team</li>
            <li>Test service functionality immediately after permission changes</li>
            <li>Document all changes in your compliance tracking system</li>
            <li>Keep backups for audit purposes (retain for 7+ days)</li>
            <li>If a service fails to start, restore from backup and investigate</li>
        </ul>
    </div>
    
    <div class="footer">
        <p><strong>Report Generated:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | <strong>Environment:</strong> DOD FedRAMP AWS GovCloud</p>
        <p><strong>Ticket:</strong> INC0135584 | <strong>Plugin:</strong> Nessus 65057 - Insecure Windows Service Permissions</p>
        <p><strong>Assessment Script Version:</strong> 1.0 | For questions, contact Security Remediation Team</p>
    </div>
    
</body>
</html>
"@
    
    $html | Out-File -FilePath $Script:ReportFile -Encoding UTF8
    Write-Log "HTML report saved: $Script:ReportFile" -Level SUCCESS
    
    return $Script:ReportFile
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

function Start-SAPPlugin65057Assessment {
    
    Write-Host @"

╔══════════════════════════════════════════════════════════════════╗
║                                                                  ║
║     SAP SECURITY ASSESSMENT - NESSUS PLUGIN 65057               ║
║     Insecure Windows Service Permissions Analysis               ║
║                                                                  ║
║     Environment: DOD FedRAMP AWS GovCloud                       ║
║     Ticket: INC0135584                                          ║
║     Servers: 6                                                  ║
║                                                                  ║
╚══════════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Cyan
    
    Write-Log "=== SAP Plugin 65057 Assessment Started ===" -Level INFO
    Write-Log "Target servers: $($Config.TargetServers -join ', ')"
    
    # Get credentials once for all servers
    $credential = Get-RemoteCredential
    
    $allResults = @()
    $serverNum = 0
    $totalServers = $Config.TargetServers.Count
    
    foreach ($server in $Config.TargetServers) {
        $serverNum++
        
        Write-Host "`n" ("=" * 70) -ForegroundColor Cyan
        Write-Host "[$serverNum/$totalServers] Processing: $server" -ForegroundColor Cyan
        Write-Host ("=" * 70) -ForegroundColor Cyan
        
        # Test connectivity
        $isReachable = Test-ServerConnectivity -ServerName $server -Credential $credential
        
        if (-not $isReachable) {
            Write-Log "Skipping $server - not reachable" -Level ERROR
            
            $allResults += [PSCustomObject]@{
                ServerName = $server
                Status = "UNREACHABLE"
                PatchStatus = $null
                Vulnerabilities = @()
                Remediation = [PSCustomObject]@{
                    NeedsPatch = $false
                    NeedsPermissionFix = $false
                    Reasons = @("Server unreachable - check network/credentials")
                    Action = "CHECK_CONNECTIVITY"
                    Priority = "HIGH"
                }
            }
            continue
        }
        
        # Get patch status
        Write-Host "`n[1/2] Checking patch status..." -ForegroundColor Yellow
        $patchStatus = Get-PatchStatus -ServerName $server -Credential $credential
        
        # Get vulnerabilities
        Write-Host "[2/2] Analyzing service permissions..." -ForegroundColor Yellow
        $vulns = Get-ServicePermissionVulnerabilities -ServerName $server -Credential $credential
        
        # Determine remediation
        $remediation = Get-RemediationStrategy -PatchStatus $patchStatus -Vulnerabilities $vulns
        
        # Store results
        $allResults += [PSCustomObject]@{
            ServerName = $server
            Status = "SUCCESS"
            PatchStatus = $patchStatus
            Vulnerabilities = $vulns
            Remediation = $remediation
        }
        
        # Summary for this server
        $sapVulns = ($vulns | Where-Object IsSAPService -eq $true).Count
        Write-Host "`n✓ Assessment complete for $server" -ForegroundColor Green
        Write-Host "  Total vulnerabilities: $($vulns.Count)" -ForegroundColor $(if($vulns.Count -gt 0){'Yellow'}else{'Green'})
        if ($sapVulns -gt 0) {
            Write-Host "  SAP services affected: $sapVulns [CRITICAL]" -ForegroundColor Red
        }
        Write-Host "  Recommended action: $($remediation.Action)" -ForegroundColor Cyan
    }
    
    # Generate report
    Write-Host "`n" ("=" * 70) -ForegroundColor Cyan
    Write-Host "Generating HTML Report..." -ForegroundColor Cyan
    Write-Host ("=" * 70) -ForegroundColor Cyan
    
    $reportPath = New-HTMLReport -Results $allResults
    
    # Export CSV
    $csvPath = Join-Path $Config.OutputPath "SAP_Plugin65057_Data_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $allResults | ForEach-Object {
        $server = $_.ServerName
        $remediation = $_.Remediation
        foreach ($vuln in $_.Vulnerabilities) {
            [PSCustomObject]@{
                Server = $server
                ServiceName = $vuln.ServiceName
                DisplayName = $vuln.DisplayName
                IsSAPService = $vuln.IsSAPService
                State = $vuln.State
                VulnerableIdentity = $vuln.VulnerableIdentity
                DangerousRights = $vuln.DangerousRights
                IsInherited = $vuln.IsInherited
                Severity = $vuln.Severity
                Action = $remediation.Action
                Priority = $remediation.Priority
            }
        }
    } | Export-Csv -Path $csvPath -NoTypeInformation
    
    Write-Log "CSV data exported: $csvPath" -Level SUCCESS
    
    # Final summary
    $totalVulns = ($allResults | ForEach-Object {$_.Vulnerabilities.Count} | Measure-Object -Sum).Sum
    $sapVulns = ($allResults | ForEach-Object {$_.Vulnerabilities | Where-Object IsSAPService}).Count
    
    Write-Host "`n" ("=" * 70) -ForegroundColor Green
    Write-Host "ASSESSMENT COMPLETE" -ForegroundColor Green
    Write-Host ("=" * 70) -ForegroundColor Green
    Write-Host "`n📊 Results Summary:" -ForegroundColor Cyan
    Write-Host "  Servers scanned: $totalServers" -ForegroundColor White
    Write-Host "  Total vulnerabilities: $totalVulns" -ForegroundColor $(if($totalVulns -gt 0){'Yellow'}else{'Green'})
    Write-Host "  SAP services affected: $sapVulns" -ForegroundColor $(if($sapVulns -gt 0){'Red'}else{'Green'})
    
    Write-Host "`n📁 Output Files:" -ForegroundColor Cyan
    Write-Host "  HTML Report: $reportPath" -ForegroundColor White
    Write-Host "  CSV Data: $csvPath" -ForegroundColor White
    Write-Host "  Log File: $Script:LogFile" -ForegroundColor White
    
    Write-Host "`n✓ Opening HTML report..." -ForegroundColor Green
    Start-Process $reportPath
    
    Write-Log "=== Assessment Complete ===" -Level SUCCESS
    
    return $allResults
}

# ============================================================================
# RUN ASSESSMENT
# ============================================================================

try {
    $results = Start-SAPPlugin65057Assessment
    
    Write-Host "`n" ("=" * 70) -ForegroundColor Cyan
    Write-Host "NEXT STEPS FOR INC0135584:" -ForegroundColor Cyan
    Write-Host ("=" * 70) -ForegroundColor Cyan
    Write-Host "1. Review HTML report (opened automatically)" -ForegroundColor White
    Write-Host "2. Prioritize servers with SAP service issues" -ForegroundColor White  
    Write-Host "3. Test remediation scripts in non-prod first" -ForegroundColor White
    Write-Host "4. Create change request for production" -ForegroundColor White
    Write-Host "5. Coordinate with SAP team before changes" -ForegroundColor White
    Write-Host "6. Execute during maintenance window" -ForegroundColor White
    Write-Host "7. Verify with Nessus rescan" -ForegroundColor White
    Write-Host "8. Update and close ticket INC0135584" -ForegroundColor White
    
} catch {
    Write-Log "Critical error: $_" -Level ERROR
    Write-Host "`n✗ Assessment failed. Check log: $Script:LogFile" -ForegroundColor Red
    exit 1
}
