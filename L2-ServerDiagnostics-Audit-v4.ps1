<#
.SYNOPSIS
    L2 Server Diagnostics - AUDIT REPORT (READ-ONLY)
    Author: Syed Rizvi
    Version: 4.0

.DESCRIPTION
    Safe READ-ONLY diagnostic script for L2 team.
    - NO dangerous commands (no kill, purge, reset, repair)
    - Checks GPO, Replication, Security Tools (Trellix, Trend, Nessus)
    - SSL Certificate Expiration
    - CPU Usage & Top Processes
    - NIC Teaming Status
    - Failed Scheduled Tasks
    - Windows Backup Status
    - Failed Login Attempts
    - Windows Activation
    - Page File Usage
    - Last 20 Event Errors
    - Creates CSV and HTML reports automatically
    - Saves to Desktop AND C:\L2_Reports

    *** THIS SCRIPT DOES NOT MAKE ANY CHANGES ***

.EXAMPLE
    .\L2-ServerDiagnostics-Audit-v4.ps1
    .\L2-ServerDiagnostics-Audit-v4.ps1 -OutputPath "C:\Reports"
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "$env:USERPROFILE\Desktop",
    
    [Parameter(Mandatory=$false)]
    [string]$SecondaryPath = "C:\L2_Reports"
)

$ErrorActionPreference = "SilentlyContinue"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$computerName = $env:COMPUTERNAME

# Create output directories if they don't exist
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}
if (-not (Test-Path $SecondaryPath)) {
    New-Item -ItemType Directory -Path $SecondaryPath -Force | Out-Null
}

# Initialize results collection
$auditResults = @()

function Add-AuditResult {
    param($Category, $Check, $Status, $Details, $Severity)
    $script:auditResults += [PSCustomObject]@{
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Computer = $computerName
        Category = $Category
        Check = $Check
        Status = $Status
        Details = $Details
        Severity = $Severity
    }
}

Write-Host ""
Write-Host "=============================================================" -ForegroundColor Cyan
Write-Host "    L2 SERVER DIAGNOSTICS - AUDIT REPORT v4.0" -ForegroundColor Cyan
Write-Host "    Author: Syed Rizvi" -ForegroundColor Cyan
Write-Host "=============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "*** THIS SCRIPT IS READ-ONLY - NO CHANGES WILL BE MADE ***" -ForegroundColor Green
Write-Host ""
Write-Host "Computer: $computerName" -ForegroundColor White
Write-Host "Started:  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor White
Write-Host ""
Write-Host "Reports will be saved to:" -ForegroundColor Yellow
Write-Host "  1. $OutputPath" -ForegroundColor Cyan
Write-Host "  2. $SecondaryPath" -ForegroundColor Cyan
Write-Host ""

$totalChecks = 25

# ============================================================
# [1/25] SYSTEM INFORMATION
# ============================================================
Write-Host "[1/$totalChecks] Gathering System Information..." -ForegroundColor Yellow

$os = Get-WmiObject Win32_OperatingSystem
$cs = Get-WmiObject Win32_ComputerSystem
$lastBoot = $os.ConvertToDateTime($os.LastBootUpTime)
$uptime = (Get-Date) - $lastBoot

$systemInfo = @{
    ComputerName = $computerName
    OS = $os.Caption
    Version = $os.Version
    LastBoot = $lastBoot.ToString("yyyy-MM-dd HH:mm:ss")
    Uptime = "$($uptime.Days)d $($uptime.Hours)h $($uptime.Minutes)m"
    TotalMemoryGB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
    Domain = $cs.Domain
    DomainJoined = $cs.PartOfDomain
}

Add-AuditResult "System" "Computer Name" "INFO" $computerName "Info"
Add-AuditResult "System" "Operating System" "INFO" $os.Caption "Info"
Add-AuditResult "System" "Last Boot" "INFO" $systemInfo.LastBoot "Info"
Add-AuditResult "System" "Uptime" $(if ($uptime.Days -gt 30) { "WARNING" } else { "OK" }) $systemInfo.Uptime $(if ($uptime.Days -gt 30) { "Medium" } else { "Info" })
Add-AuditResult "System" "Domain" "INFO" $cs.Domain "Info"

Write-Host "      Computer: $computerName" -ForegroundColor Green
Write-Host "      Uptime: $($systemInfo.Uptime)" -ForegroundColor $(if ($uptime.Days -gt 30) { "Yellow" } else { "Green" })

# ============================================================
# [2/25] WINDOWS ACTIVATION STATUS
# ============================================================
Write-Host "[2/$totalChecks] Checking Windows Activation..." -ForegroundColor Yellow

try {
    $licenseStatus = Get-WmiObject SoftwareLicensingProduct | Where-Object { $_.PartialProductKey -and $_.Name -like "*Windows*" } | Select-Object -First 1
    
    if ($licenseStatus) {
        $statusText = switch ($licenseStatus.LicenseStatus) {
            0 { "Unlicensed" }
            1 { "Licensed" }
            2 { "Out-of-Box Grace" }
            3 { "Out-of-Tolerance Grace" }
            4 { "Non-Genuine Grace" }
            5 { "Notification" }
            6 { "Extended Grace" }
            default { "Unknown" }
        }
        
        if ($licenseStatus.LicenseStatus -eq 1) {
            Add-AuditResult "Activation" "Windows License" "OK" "Windows is activated - $statusText" "Info"
            Write-Host "      Windows Activation: ACTIVATED" -ForegroundColor Green
        } else {
            Add-AuditResult "Activation" "Windows License" "CRITICAL" "Windows is NOT activated - $statusText" "Critical"
            Write-Host "      Windows Activation: NOT ACTIVATED - $statusText" -ForegroundColor Red
        }
    } else {
        Add-AuditResult "Activation" "Windows License" "WARNING" "Could not determine license status" "Medium"
        Write-Host "      Windows Activation: Unknown" -ForegroundColor Yellow
    }
} catch {
    Add-AuditResult "Activation" "Windows License" "WARNING" "Could not check activation status" "Medium"
}

# ============================================================
# [3/25] DOMAIN & TRUST STATUS
# ============================================================
Write-Host "[3/$totalChecks] Checking Domain & Trust Status..." -ForegroundColor Yellow

$domainStatus = "Unknown"
$trustStatus = "Unknown"
$dcName = "N/A"

if ($cs.PartOfDomain) {
    try {
        $dcInfo = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        $dcName = $dcInfo.FindDomainController().Name
        $domainStatus = "Connected"
        Add-AuditResult "Domain" "Domain Controller" "OK" $dcName "Info"
        
        $secureChannel = Test-ComputerSecureChannel -ErrorAction SilentlyContinue
        if ($secureChannel) {
            $trustStatus = "Healthy"
            Add-AuditResult "Domain" "Trust Relationship" "OK" "Secure channel is valid" "Info"
            Write-Host "      Trust Relationship: HEALTHY" -ForegroundColor Green
        } else {
            $trustStatus = "BROKEN"
            Add-AuditResult "Domain" "Trust Relationship" "CRITICAL" "Secure channel is BROKEN - needs repair" "Critical"
            Write-Host "      Trust Relationship: BROKEN" -ForegroundColor Red
        }
    } catch {
        $domainStatus = "Error"
        Add-AuditResult "Domain" "Domain Status" "WARNING" "Cannot contact domain controller" "High"
        Write-Host "      Domain Status: Cannot contact DC" -ForegroundColor Red
    }
} else {
    $domainStatus = "Workgroup"
    Add-AuditResult "Domain" "Domain Status" "INFO" "Computer is in a Workgroup (not domain joined)" "Info"
    Write-Host "      Domain Status: Workgroup (not domain joined)" -ForegroundColor Yellow
}

# ============================================================
# [4/25] GPO (GROUP POLICY) STATUS
# ============================================================
Write-Host "[4/$totalChecks] Checking Group Policy Status..." -ForegroundColor Yellow

$gpoStatus = "Unknown"
$lastGPORefresh = "Unknown"

try {
    $gpoEvents = Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-GroupPolicy/Operational'; Level=2,3; StartTime=(Get-Date).AddDays(-7)} -MaxEvents 10 -ErrorAction SilentlyContinue
    
    if ($gpoEvents) {
        $gpoErrorCount = $gpoEvents.Count
        $gpoStatus = "ERRORS"
        Add-AuditResult "GPO" "Group Policy Status" "WARNING" "$gpoErrorCount GPO errors in last 7 days" "High"
        Write-Host "      GPO Status: $gpoErrorCount errors found" -ForegroundColor Yellow
        
        foreach ($evt in $gpoEvents | Select-Object -First 3) {
            $msgPreview = if ($evt.Message.Length -gt 150) { $evt.Message.Substring(0, 150) + "..." } else { $evt.Message }
            Add-AuditResult "GPO" "GPO Error $($evt.Id)" "WARNING" $msgPreview "Medium"
        }
    } else {
        $gpoStatus = "Healthy"
        Add-AuditResult "GPO" "Group Policy Status" "OK" "No GPO errors in last 7 days" "Info"
        Write-Host "      GPO Status: Healthy (no errors)" -ForegroundColor Green
    }
    
    $gpResult = gpresult /r 2>$null | Select-String "Last time Group Policy was applied"
    if ($gpResult) {
        $lastGPORefresh = ($gpResult -split ":")[-1].Trim()
        Add-AuditResult "GPO" "Last GPO Refresh" "INFO" $lastGPORefresh "Info"
    }
} catch {
    Add-AuditResult "GPO" "Group Policy Status" "WARNING" "Could not check GPO status" "Medium"
}

# ============================================================
# [5/25] AD REPLICATION STATUS (if DC)
# ============================================================
Write-Host "[5/$totalChecks] Checking AD Replication Status..." -ForegroundColor Yellow

$isDC = (Get-WmiObject Win32_ComputerSystem).DomainRole -ge 4
$replicationStatus = "N/A"

if ($isDC) {
    try {
        $replStatus = repadmin /showrepl /csv 2>$null | ConvertFrom-Csv
        $replErrors = $replStatus | Where-Object { $_.'Number of Failures' -gt 0 }
        
        if ($replErrors) {
            $replicationStatus = "ERRORS"
            Add-AuditResult "Replication" "AD Replication" "CRITICAL" "$($replErrors.Count) replication failures detected" "Critical"
            Write-Host "      AD Replication: ERRORS DETECTED" -ForegroundColor Red
        } else {
            $replicationStatus = "Healthy"
            Add-AuditResult "Replication" "AD Replication" "OK" "All replication partners healthy" "Info"
            Write-Host "      AD Replication: Healthy" -ForegroundColor Green
        }
    } catch {
        Add-AuditResult "Replication" "AD Replication" "INFO" "Could not check replication" "Info"
    }
} else {
    Add-AuditResult "Replication" "AD Replication" "INFO" "Not a Domain Controller - skipped" "Info"
    Write-Host "      AD Replication: N/A (not a DC)" -ForegroundColor Gray
}

# ============================================================
# [6/25] CPU USAGE
# ============================================================
Write-Host "[6/$totalChecks] Checking CPU Usage..." -ForegroundColor Yellow

try {
    $cpuLoad = (Get-WmiObject Win32_Processor | Measure-Object -Property LoadPercentage -Average).Average
    $cpuLoad = [math]::Round($cpuLoad, 1)
    
    if ($cpuLoad -gt 90) {
        Add-AuditResult "CPU" "CPU Usage" "CRITICAL" "$cpuLoad% - Very High!" "Critical"
        Write-Host "      CPU Usage: $cpuLoad% - CRITICAL" -ForegroundColor Red
    } elseif ($cpuLoad -gt 80) {
        Add-AuditResult "CPU" "CPU Usage" "WARNING" "$cpuLoad% - High" "High"
        Write-Host "      CPU Usage: $cpuLoad% - WARNING" -ForegroundColor Yellow
    } else {
        Add-AuditResult "CPU" "CPU Usage" "OK" "$cpuLoad% - Normal" "Info"
        Write-Host "      CPU Usage: $cpuLoad% - OK" -ForegroundColor Green
    }
} catch {
    Add-AuditResult "CPU" "CPU Usage" "WARNING" "Could not check CPU usage" "Medium"
}

# ============================================================
# [7/25] TOP PROCESSES (CPU & MEMORY)
# ============================================================
Write-Host "[7/$totalChecks] Checking Top Processes..." -ForegroundColor Yellow

try {
    # Top 5 CPU consuming processes
    $topCPU = Get-Process | Sort-Object CPU -Descending | Select-Object -First 5
    foreach ($proc in $topCPU) {
        $cpuTime = [math]::Round($proc.CPU, 2)
        Add-AuditResult "Processes" "Top CPU: $($proc.ProcessName)" "INFO" "CPU Time: $cpuTime sec, Memory: $([math]::Round($proc.WorkingSet64/1MB, 1)) MB" "Info"
    }
    
    # Top 5 Memory consuming processes
    $topMem = Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 5
    foreach ($proc in $topMem) {
        $memMB = [math]::Round($proc.WorkingSet64 / 1MB, 1)
        Add-AuditResult "Processes" "Top Memory: $($proc.ProcessName)" "INFO" "Memory: $memMB MB" "Info"
    }
    
    Write-Host "      Top Processes: Captured" -ForegroundColor Green
} catch {
    Add-AuditResult "Processes" "Process List" "WARNING" "Could not get process list" "Medium"
}

# ============================================================
# [8/25] MEMORY CHECK
# ============================================================
Write-Host "[8/$totalChecks] Checking Memory Usage..." -ForegroundColor Yellow

$totalMem = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
$freeMem = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
$usedMem = $totalMem - $freeMem
$memPercent = [math]::Round(($usedMem / $totalMem) * 100, 1)

if ($memPercent -gt 90) {
    Add-AuditResult "Memory" "RAM Usage" "CRITICAL" "$memPercent% used ($usedMem GB of $totalMem GB)" "Critical"
    Write-Host "      Memory: $memPercent% used - CRITICAL" -ForegroundColor Red
} elseif ($memPercent -gt 80) {
    Add-AuditResult "Memory" "RAM Usage" "WARNING" "$memPercent% used ($usedMem GB of $totalMem GB)" "High"
    Write-Host "      Memory: $memPercent% used - WARNING" -ForegroundColor Yellow
} else {
    Add-AuditResult "Memory" "RAM Usage" "OK" "$memPercent% used ($usedMem GB of $totalMem GB)" "Info"
    Write-Host "      Memory: $memPercent% used - OK" -ForegroundColor Green
}

# ============================================================
# [9/25] PAGE FILE USAGE
# ============================================================
Write-Host "[9/$totalChecks] Checking Page File Usage..." -ForegroundColor Yellow

try {
    $pageFile = Get-WmiObject Win32_PageFileUsage
    if ($pageFile) {
        $pfUsedMB = $pageFile.CurrentUsage
        $pfMaxMB = $pageFile.AllocatedBaseSize
        $pfPercent = if ($pfMaxMB -gt 0) { [math]::Round(($pfUsedMB / $pfMaxMB) * 100, 1) } else { 0 }
        
        if ($pfPercent -gt 80) {
            Add-AuditResult "PageFile" "Page File Usage" "WARNING" "$pfPercent% used ($pfUsedMB MB of $pfMaxMB MB)" "High"
            Write-Host "      Page File: $pfPercent% used - WARNING" -ForegroundColor Yellow
        } else {
            Add-AuditResult "PageFile" "Page File Usage" "OK" "$pfPercent% used ($pfUsedMB MB of $pfMaxMB MB)" "Info"
            Write-Host "      Page File: $pfPercent% used - OK" -ForegroundColor Green
        }
    } else {
        Add-AuditResult "PageFile" "Page File" "WARNING" "No page file configured" "High"
        Write-Host "      Page File: Not configured" -ForegroundColor Yellow
    }
} catch {
    Add-AuditResult "PageFile" "Page File" "WARNING" "Could not check page file" "Medium"
}

# ============================================================
# [10/25] DISK SPACE CHECK
# ============================================================
Write-Host "[10/$totalChecks] Checking Disk Space..." -ForegroundColor Yellow

$disks = Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3"
foreach ($disk in $disks) {
    $freePercent = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 1)
    $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
    $totalGB = [math]::Round($disk.Size / 1GB, 2)
    
    if ($freePercent -lt 10) {
        Add-AuditResult "Disk" "$($disk.DeviceID) Space" "CRITICAL" "$freeGB GB free ($freePercent%) of $totalGB GB" "Critical"
        Write-Host "      $($disk.DeviceID) $freePercent% free - CRITICAL" -ForegroundColor Red
    } elseif ($freePercent -lt 20) {
        Add-AuditResult "Disk" "$($disk.DeviceID) Space" "WARNING" "$freeGB GB free ($freePercent%) of $totalGB GB" "High"
        Write-Host "      $($disk.DeviceID) $freePercent% free - WARNING" -ForegroundColor Yellow
    } else {
        Add-AuditResult "Disk" "$($disk.DeviceID) Space" "OK" "$freeGB GB free ($freePercent%) of $totalGB GB" "Info"
        Write-Host "      $($disk.DeviceID) $freePercent% free - OK" -ForegroundColor Green
    }
}

# ============================================================
# [11/25] CRITICAL SERVICES CHECK
# ============================================================
Write-Host "[11/$totalChecks] Checking Critical Services..." -ForegroundColor Yellow

$criticalServices = @(
    @{Name="Netlogon"; Display="Netlogon (Domain Auth)"; Critical=$true},
    @{Name="W32Time"; Display="Windows Time"; Critical=$true},
    @{Name="gpsvc"; Display="Group Policy Client"; Critical=$true},
    @{Name="DNS"; Display="DNS Server"; Critical=$false},
    @{Name="NTDS"; Display="AD Domain Services"; Critical=$false},
    @{Name="Dnscache"; Display="DNS Client"; Critical=$true},
    @{Name="LanmanWorkstation"; Display="Workstation"; Critical=$true},
    @{Name="LanmanServer"; Display="Server"; Critical=$true},
    @{Name="EventLog"; Display="Windows Event Log"; Critical=$true},
    @{Name="Schedule"; Display="Task Scheduler"; Critical=$true},
    @{Name="TermService"; Display="Remote Desktop Services"; Critical=$true},
    @{Name="CryptSvc"; Display="Cryptographic Services"; Critical=$true},
    @{Name="BITS"; Display="BITS (Background Transfer)"; Critical=$true},
    @{Name="wuauserv"; Display="Windows Update"; Critical=$false},
    @{Name="AmazonSSMAgent"; Display="AWS SSM Agent"; Critical=$false}
)

$stoppedCritical = 0
foreach ($svc in $criticalServices) {
    $service = Get-Service -Name $svc.Name -ErrorAction SilentlyContinue
    if ($service) {
        if ($service.Status -eq "Running") {
            Add-AuditResult "Services" $svc.Display "OK" "Running" "Info"
        } else {
            $severity = if ($svc.Critical) { "Critical" } else { "Medium" }
            Add-AuditResult "Services" $svc.Display "STOPPED" "Service is $($service.Status)" $severity
            if ($svc.Critical) { $stoppedCritical++ }
        }
    }
}

if ($stoppedCritical -gt 0) {
    Write-Host "      Critical Services: $stoppedCritical STOPPED" -ForegroundColor Red
} else {
    Write-Host "      Critical Services: All Running" -ForegroundColor Green
}

# ============================================================
# [12/25] NIC TEAMING STATUS
# ============================================================
Write-Host "[12/$totalChecks] Checking NIC Teaming Status..." -ForegroundColor Yellow

try {
    $nicTeams = Get-NetLbfoTeam -ErrorAction SilentlyContinue
    
    if ($nicTeams) {
        foreach ($team in $nicTeams) {
            if ($team.Status -eq "Up") {
                Add-AuditResult "NIC Team" $team.Name "OK" "Status: $($team.Status), Mode: $($team.TeamingMode)" "Info"
                Write-Host "      NIC Team '$($team.Name)': UP" -ForegroundColor Green
            } else {
                Add-AuditResult "NIC Team" $team.Name "CRITICAL" "Status: $($team.Status) - Team is DOWN!" "Critical"
                Write-Host "      NIC Team '$($team.Name)': DOWN" -ForegroundColor Red
            }
            
            # Check team members
            $members = Get-NetLbfoTeamMember -Team $team.Name -ErrorAction SilentlyContinue
            foreach ($member in $members) {
                if ($member.AdministrativeMode -eq "Active" -and $member.ReceiveLinkSpeed -gt 0) {
                    Add-AuditResult "NIC Team" "$($team.Name) - $($member.Name)" "OK" "Active, Speed: $($member.ReceiveLinkSpeed)" "Info"
                } else {
                    Add-AuditResult "NIC Team" "$($team.Name) - $($member.Name)" "WARNING" "Status: $($member.AdministrativeMode)" "High"
                }
            }
        }
    } else {
        Add-AuditResult "NIC Team" "NIC Teaming" "INFO" "No NIC teams configured" "Info"
        Write-Host "      NIC Teaming: Not configured" -ForegroundColor Gray
    }
} catch {
    Add-AuditResult "NIC Team" "NIC Teaming" "INFO" "NIC Teaming not available or not configured" "Info"
    Write-Host "      NIC Teaming: Not available" -ForegroundColor Gray
}

# ============================================================
# [13/25] SSL CERTIFICATE EXPIRATION
# ============================================================
Write-Host "[13/$totalChecks] Checking SSL Certificate Expiration..." -ForegroundColor Yellow

try {
    $certs = Get-ChildItem Cert:\LocalMachine\My -ErrorAction SilentlyContinue
    $expiringCerts = 0
    $expiredCerts = 0
    
    foreach ($cert in $certs) {
        $daysUntilExpiry = ($cert.NotAfter - (Get-Date)).Days
        $certName = if ($cert.FriendlyName) { $cert.FriendlyName } else { $cert.Subject.Substring(0, [Math]::Min(50, $cert.Subject.Length)) }
        
        if ($daysUntilExpiry -lt 0) {
            Add-AuditResult "SSL Certs" $certName "CRITICAL" "EXPIRED $([Math]::Abs($daysUntilExpiry)) days ago!" "Critical"
            $expiredCerts++
        } elseif ($daysUntilExpiry -lt 30) {
            Add-AuditResult "SSL Certs" $certName "WARNING" "Expires in $daysUntilExpiry days ($($cert.NotAfter.ToString('yyyy-MM-dd')))" "High"
            $expiringCerts++
        } elseif ($daysUntilExpiry -lt 90) {
            Add-AuditResult "SSL Certs" $certName "INFO" "Expires in $daysUntilExpiry days ($($cert.NotAfter.ToString('yyyy-MM-dd')))" "Medium"
        }
    }
    
    if ($expiredCerts -gt 0) {
        Write-Host "      SSL Certificates: $expiredCerts EXPIRED!" -ForegroundColor Red
    } elseif ($expiringCerts -gt 0) {
        Write-Host "      SSL Certificates: $expiringCerts expiring soon" -ForegroundColor Yellow
    } else {
        Add-AuditResult "SSL Certs" "Certificate Status" "OK" "All certificates valid" "Info"
        Write-Host "      SSL Certificates: All valid" -ForegroundColor Green
    }
} catch {
    Add-AuditResult "SSL Certs" "Certificate Check" "WARNING" "Could not check certificates" "Medium"
}

# ============================================================
# [14/25] SECURITY TOOLS CHECK (Trellix, Trend, Nessus)
# ============================================================
Write-Host "[14/$totalChecks] Checking Security Tools..." -ForegroundColor Yellow

$securityToolsFound = @()

# --- TRELLIX/McAfee ---
Write-Host "      --- TRELLIX/McAfee ---" -ForegroundColor Cyan

$trellixServices = @(
    @{Name="macmnsvc"; Display="Trellix Agent (MA)"; Critical=$true},
    @{Name="masvc"; Display="Trellix Agent Service"; Critical=$true},
    @{Name="mfefire"; Display="Trellix Firewall"; Critical=$false},
    @{Name="mfemms"; Display="Trellix Management Service"; Critical=$true},
    @{Name="mfevtp"; Display="Trellix Validation Trust"; Critical=$false},
    @{Name="McShield"; Display="Trellix On-Access Scanner"; Critical=$true},
    @{Name="McAfeeFramework"; Display="Trellix Framework"; Critical=$true}
)

$trellixInstalled = Test-Path "C:\Program Files\McAfee" -or Test-Path "C:\Program Files (x86)\McAfee" -or Test-Path "C:\Program Files\Trellix"
$trellixRunning = 0
$trellixStopped = 0

if ($trellixInstalled) {
    foreach ($svc in $trellixServices) {
        $service = Get-Service -Name $svc.Name -ErrorAction SilentlyContinue
        if ($service) {
            if ($service.Status -eq "Running") {
                Add-AuditResult "Trellix" $svc.Display "OK" "Running" "Info"
                $trellixRunning++
            } else {
                $severity = if ($svc.Critical) { "High" } else { "Medium" }
                Add-AuditResult "Trellix" $svc.Display "STOPPED" "Service is $($service.Status)" $severity
                $trellixStopped++
            }
        }
    }
    
    # Check ePO connection
    $cmdAgentPath = "C:\Program Files\McAfee\Agent\cmdagent.exe"
    if (Test-Path $cmdAgentPath) {
        try {
            $agentInfo = & $cmdAgentPath -i 2>&1 | Out-String
            if ($agentInfo -match "ePO Server") {
                Add-AuditResult "Trellix" "ePO Connection" "OK" "Connected to ePO Server" "Info"
                $securityToolsFound += "Trellix: Connected to ePO"
                Write-Host "      Trellix ePO: CONNECTED" -ForegroundColor Green
            }
        } catch { }
    }
    
    if ($trellixRunning -gt 0) {
        $securityToolsFound += "Trellix: $trellixRunning services running"
        Write-Host "      Trellix: $trellixRunning services running" -ForegroundColor $(if ($trellixStopped -gt 0) { "Yellow" } else { "Green" })
    }
} else {
    Add-AuditResult "Trellix" "Installation" "INFO" "Trellix/McAfee not installed" "Info"
    Write-Host "      Trellix: Not Installed" -ForegroundColor Gray
}

# --- TREND MICRO ---
Write-Host "      --- TREND MICRO ---" -ForegroundColor Cyan

$trendServices = @(
    @{Name="ds_agent"; Display="Deep Security Agent"; Critical=$true},
    @{Name="Trend Micro Deep Security Agent"; Display="DSA (Alt)"; Critical=$true},
    @{Name="TmListen"; Display="Trend Micro Listen"; Critical=$true},
    @{Name="ntrtscan"; Display="Trend Real-Time Scan"; Critical=$true},
    @{Name="TmCCSF"; Display="Trend Common Client"; Critical=$false},
    @{Name="Apex One NT RealTime Scan"; Display="Apex One RealTime"; Critical=$true}
)

$trendInstalled = Test-Path "C:\Program Files\Trend Micro" -or Test-Path "C:\Program Files (x86)\Trend Micro"
$trendRunning = 0
$trendStopped = 0

if ($trendInstalled) {
    foreach ($svc in $trendServices) {
        $service = Get-Service -Name $svc.Name -ErrorAction SilentlyContinue
        if ($service) {
            if ($service.Status -eq "Running") {
                Add-AuditResult "Trend Micro" $svc.Display "OK" "Running" "Info"
                $trendRunning++
            } else {
                $severity = if ($svc.Critical) { "High" } else { "Medium" }
                Add-AuditResult "Trend Micro" $svc.Display "STOPPED" "Service is $($service.Status)" $severity
                $trendStopped++
            }
        }
    }
    
    if ($trendRunning -gt 0) {
        $securityToolsFound += "Trend Micro: $trendRunning services running"
        Write-Host "      Trend Micro: $trendRunning services running" -ForegroundColor $(if ($trendStopped -gt 0) { "Yellow" } else { "Green" })
    }
} else {
    Add-AuditResult "Trend Micro" "Installation" "INFO" "Trend Micro not installed" "Info"
    Write-Host "      Trend Micro: Not Installed" -ForegroundColor Gray
}

# --- TENABLE NESSUS ---
Write-Host "      --- TENABLE NESSUS ---" -ForegroundColor Cyan

$nessusService = Get-Service -Name "Tenable Nessus Agent" -ErrorAction SilentlyContinue
$nessusPath = "C:\Program Files\Tenable\Nessus Agent"

if ($nessusService) {
    if ($nessusService.Status -eq "Running") {
        Add-AuditResult "Nessus" "Agent Service" "OK" "Running" "Info"
        
        $nessuscli = Join-Path $nessusPath "nessuscli.exe"
        if (Test-Path $nessuscli) {
            try {
                $nessusStatus = & $nessuscli agent status 2>&1 | Out-String
                
                if ($nessusStatus -match "Linked to") {
                    Add-AuditResult "Nessus" "Link Status" "OK" "Agent is linked to Tenable server" "Info"
                    $securityToolsFound += "Nessus: Linked"
                    Write-Host "      Nessus Agent: LINKED" -ForegroundColor Green
                } elseif ($nessusStatus -match "Running: Yes") {
                    Add-AuditResult "Nessus" "Link Status" "OK" "Agent running" "Info"
                    $securityToolsFound += "Nessus: Running"
                    Write-Host "      Nessus Agent: Running" -ForegroundColor Green
                } else {
                    Add-AuditResult "Nessus" "Link Status" "WARNING" "Agent may not be linked" "High"
                    $securityToolsFound += "Nessus: NOT LINKED"
                    Write-Host "      Nessus Agent: NOT LINKED" -ForegroundColor Yellow
                }
            } catch { }
        }
    } else {
        Add-AuditResult "Nessus" "Agent Service" "STOPPED" "Service is $($nessusService.Status)" "High"
        $securityToolsFound += "Nessus: STOPPED"
        Write-Host "      Nessus Service: STOPPED" -ForegroundColor Red
    }
} elseif (Test-Path $nessusPath) {
    Add-AuditResult "Nessus" "Agent Service" "WARNING" "Installed but service not found" "High"
    Write-Host "      Nessus: Installed but service missing" -ForegroundColor Yellow
} else {
    Add-AuditResult "Nessus" "Installation" "INFO" "Nessus Agent not installed" "Info"
    Write-Host "      Nessus: Not Installed" -ForegroundColor Gray
}

# --- OTHER SECURITY TOOLS ---
$otherTools = @(
    @{Name="CrowdStrike Falcon"; Service="CSFalconService"},
    @{Name="Carbon Black"; Service="CbDefense"},
    @{Name="SentinelOne"; Service="SentinelAgent"},
    @{Name="Windows Defender"; Service="WinDefend"}
)

foreach ($tool in $otherTools) {
    $service = Get-Service -Name $tool.Service -ErrorAction SilentlyContinue
    if ($service) {
        if ($service.Status -eq "Running") {
            Add-AuditResult "Security Tools" $tool.Name "OK" "Running" "Info"
            $securityToolsFound += "$($tool.Name): Running"
            Write-Host "      $($tool.Name): RUNNING" -ForegroundColor Green
        } else {
            Add-AuditResult "Security Tools" $tool.Name "STOPPED" "Service is $($service.Status)" "High"
            $securityToolsFound += "$($tool.Name): STOPPED"
        }
    }
}

if ($securityToolsFound.Count -eq 0) {
    Add-AuditResult "Security Tools" "Overall Security" "CRITICAL" "No security tools detected!" "Critical"
    Write-Host "      SECURITY: NO TOOLS DETECTED!" -ForegroundColor Red
}

# ============================================================
# [15/25] TIME SYNC CHECK
# ============================================================
Write-Host "[15/$totalChecks] Checking Time Synchronization..." -ForegroundColor Yellow

try {
    $w32tm = w32tm /query /status 2>&1 | Out-String
    $timeSource = if ($w32tm -match "Source:\s*(.+)") { $Matches[1].Trim() } else { "Unknown" }
    
    if ($timeSource -and $timeSource -ne "Local CMOS Clock" -and $timeSource -ne "Free-running System Clock") {
        Add-AuditResult "Time" "Time Source" "OK" $timeSource "Info"
        Write-Host "      Time Source: $timeSource - OK" -ForegroundColor Green
    } else {
        Add-AuditResult "Time" "Time Source" "WARNING" "Using local clock - not synced to domain" "High"
        Write-Host "      Time Source: Local Clock - WARNING" -ForegroundColor Yellow
    }
} catch {
    Add-AuditResult "Time" "Time Sync" "WARNING" "Could not check time sync" "Medium"
}

# ============================================================
# [16/25] PENDING REBOOT CHECK
# ============================================================
Write-Host "[16/$totalChecks] Checking Pending Reboot..." -ForegroundColor Yellow

$pendingReboot = $false
$rebootReasons = @()

if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") {
    $pendingReboot = $true
    $rebootReasons += "CBS"
}
if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") {
    $pendingReboot = $true
    $rebootReasons += "Windows Update"
}
if (Test-Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations") {
    $pendingReboot = $true
    $rebootReasons += "File Rename"
}

if ($pendingReboot) {
    Add-AuditResult "Reboot" "Pending Reboot" "WARNING" "Reboot required: $($rebootReasons -join ', ')" "High"
    Write-Host "      Pending Reboot: YES - $($rebootReasons -join ', ')" -ForegroundColor Yellow
} else {
    Add-AuditResult "Reboot" "Pending Reboot" "OK" "No pending reboot" "Info"
    Write-Host "      Pending Reboot: None" -ForegroundColor Green
}

# ============================================================
# [17/25] RDP CONFIGURATION CHECK
# ============================================================
Write-Host "[17/$totalChecks] Checking RDP Configuration..." -ForegroundColor Yellow

$rdpEnabled = (Get-ItemProperty "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -ErrorAction SilentlyContinue).fDenyTSConnections -eq 0
$nlaEnabled = (Get-ItemProperty "HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" -Name "UserAuthentication" -ErrorAction SilentlyContinue).UserAuthentication -eq 1

if ($rdpEnabled) {
    Add-AuditResult "RDP" "RDP Status" "OK" "RDP is enabled" "Info"
    Write-Host "      RDP: Enabled" -ForegroundColor Green
} else {
    Add-AuditResult "RDP" "RDP Status" "WARNING" "RDP is disabled" "Medium"
    Write-Host "      RDP: Disabled" -ForegroundColor Yellow
}

Add-AuditResult "RDP" "NLA Status" "INFO" $(if ($nlaEnabled) { "NLA Enabled (Secure)" } else { "NLA Disabled" }) "Info"

# ============================================================
# [18/25] FIREWALL STATUS
# ============================================================
Write-Host "[18/$totalChecks] Checking Firewall Status..." -ForegroundColor Yellow

$fwProfiles = Get-NetFirewallProfile -ErrorAction SilentlyContinue
foreach ($profile in $fwProfiles) {
    $status = if ($profile.Enabled) { "Enabled" } else { "Disabled" }
    Add-AuditResult "Firewall" "$($profile.Name) Profile" $(if ($profile.Enabled) { "OK" } else { "WARNING" }) $status $(if ($profile.Enabled) { "Info" } else { "Medium" })
}
Write-Host "      Firewall Profiles Checked" -ForegroundColor Green

# ============================================================
# [19/25] FAILED SCHEDULED TASKS
# ============================================================
Write-Host "[19/$totalChecks] Checking Failed Scheduled Tasks..." -ForegroundColor Yellow

try {
    $failedTasks = Get-ScheduledTask | Get-ScheduledTaskInfo -ErrorAction SilentlyContinue | 
                   Where-Object { $_.LastTaskResult -ne 0 -and $_.LastTaskResult -ne 267009 -and $_.LastRunTime -gt (Get-Date).AddDays(-7) }
    
    $failedCount = 0
    foreach ($task in $failedTasks | Select-Object -First 10) {
        $taskName = $task.TaskName
        $lastResult = $task.LastTaskResult
        Add-AuditResult "Scheduled Tasks" $taskName "WARNING" "Last result: $lastResult (Failed)" "Medium"
        $failedCount++
    }
    
    if ($failedCount -gt 0) {
        Write-Host "      Scheduled Tasks: $failedCount failed in last 7 days" -ForegroundColor Yellow
    } else {
        Add-AuditResult "Scheduled Tasks" "Task Status" "OK" "No failed tasks in last 7 days" "Info"
        Write-Host "      Scheduled Tasks: All OK" -ForegroundColor Green
    }
} catch {
    Add-AuditResult "Scheduled Tasks" "Task Check" "WARNING" "Could not check scheduled tasks" "Medium"
}

# ============================================================
# [20/25] WINDOWS BACKUP STATUS
# ============================================================
Write-Host "[20/$totalChecks] Checking Windows Backup Status..." -ForegroundColor Yellow

try {
    # Check Windows Server Backup
    $wsbJob = Get-WBJob -Previous 1 -ErrorAction SilentlyContinue
    
    if ($wsbJob) {
        if ($wsbJob.JobState -eq "Completed" -or $wsbJob.HResult -eq 0) {
            Add-AuditResult "Backup" "Windows Server Backup" "OK" "Last backup successful: $($wsbJob.EndTime)" "Info"
            Write-Host "      Windows Backup: Last backup successful" -ForegroundColor Green
        } else {
            Add-AuditResult "Backup" "Windows Server Backup" "WARNING" "Last backup failed: $($wsbJob.ErrorDescription)" "High"
            Write-Host "      Windows Backup: FAILED" -ForegroundColor Red
        }
    } else {
        # Check for VSS/Backup events
        $backupEvents = Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Backup'; Level=2,3} -MaxEvents 5 -ErrorAction SilentlyContinue
        
        if ($backupEvents) {
            Add-AuditResult "Backup" "Backup Status" "WARNING" "Backup errors found in event log" "High"
            Write-Host "      Backup: Errors in event log" -ForegroundColor Yellow
        } else {
            Add-AuditResult "Backup" "Backup Status" "INFO" "Windows Server Backup not configured or no recent jobs" "Info"
            Write-Host "      Backup: No recent backup jobs" -ForegroundColor Gray
        }
    }
} catch {
    Add-AuditResult "Backup" "Backup Status" "INFO" "Could not check backup status (may not be configured)" "Info"
    Write-Host "      Backup: Not configured or not available" -ForegroundColor Gray
}

# ============================================================
# [21/25] FAILED LOGIN ATTEMPTS (Security)
# ============================================================
Write-Host "[21/$totalChecks] Checking Failed Login Attempts..." -ForegroundColor Yellow

try {
    # Event ID 4625 = Failed login
    $failedLogins = Get-WinEvent -FilterHashtable @{LogName='Security'; Id=4625; StartTime=(Get-Date).AddHours(-24)} -MaxEvents 100 -ErrorAction SilentlyContinue
    
    $failedCount = if ($failedLogins) { $failedLogins.Count } else { 0 }
    
    if ($failedCount -gt 50) {
        Add-AuditResult "Security" "Failed Logins (24h)" "CRITICAL" "$failedCount failed login attempts - Possible brute force!" "Critical"
        Write-Host "      Failed Logins: $failedCount in 24h - CRITICAL" -ForegroundColor Red
    } elseif ($failedCount -gt 20) {
        Add-AuditResult "Security" "Failed Logins (24h)" "WARNING" "$failedCount failed login attempts" "High"
        Write-Host "      Failed Logins: $failedCount in 24h - WARNING" -ForegroundColor Yellow
    } else {
        Add-AuditResult "Security" "Failed Logins (24h)" "OK" "$failedCount failed login attempts" "Info"
        Write-Host "      Failed Logins: $failedCount in 24h - OK" -ForegroundColor Green
    }
    
    # Get top offending IPs/Users
    if ($failedCount -gt 0) {
        $topUsers = $failedLogins | ForEach-Object {
            $xml = [xml]$_.ToXml()
            $xml.Event.EventData.Data | Where-Object { $_.Name -eq "TargetUserName" } | Select-Object -ExpandProperty '#text'
        } | Group-Object | Sort-Object Count -Descending | Select-Object -First 3
        
        foreach ($user in $topUsers) {
            Add-AuditResult "Security" "Failed Login User" "INFO" "$($user.Name): $($user.Count) attempts" "Info"
        }
    }
} catch {
    Add-AuditResult "Security" "Failed Logins" "WARNING" "Could not check failed logins" "Medium"
}

# ============================================================
# [22/25] LAST 20 EVENT ERRORS
# ============================================================
Write-Host "[22/$totalChecks] Gathering Last 20 Event Errors..." -ForegroundColor Yellow

$eventErrors = @()

# System Errors
$sysErrors = Get-WinEvent -FilterHashtable @{LogName='System'; Level=2} -MaxEvents 10 -ErrorAction SilentlyContinue
foreach ($evt in $sysErrors) {
    $msgPreview = if ($evt.Message.Length -gt 150) { $evt.Message.Substring(0, 150) + "..." } else { $evt.Message }
    $eventErrors += [PSCustomObject]@{
        Log = "System"
        Time = $evt.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
        ID = $evt.Id
        Source = $evt.ProviderName
        Message = $msgPreview
    }
    Add-AuditResult "Events" "System Error $($evt.Id)" "ERROR" "$($evt.ProviderName): $msgPreview" "Medium"
}

# Application Errors
$appErrors = Get-WinEvent -FilterHashtable @{LogName='Application'; Level=2} -MaxEvents 10 -ErrorAction SilentlyContinue
foreach ($evt in $appErrors) {
    $msgPreview = if ($evt.Message.Length -gt 150) { $evt.Message.Substring(0, 150) + "..." } else { $evt.Message }
    $eventErrors += [PSCustomObject]@{
        Log = "Application"
        Time = $evt.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
        ID = $evt.Id
        Source = $evt.ProviderName
        Message = $msgPreview
    }
    Add-AuditResult "Events" "Application Error $($evt.Id)" "ERROR" "$($evt.ProviderName): $msgPreview" "Medium"
}

Write-Host "      Found $($eventErrors.Count) recent errors" -ForegroundColor $(if ($eventErrors.Count -gt 10) { "Yellow" } else { "Green" })

# ============================================================
# [23/25] NETWORK CONNECTIVITY
# ============================================================
Write-Host "[23/$totalChecks] Checking Network Connectivity..." -ForegroundColor Yellow

$networkTests = @(
    @{Target=$dcName; Port=389; Name="LDAP to DC"},
    @{Target=$dcName; Port=88; Name="Kerberos to DC"},
    @{Target="8.8.8.8"; Port=53; Name="External DNS"}
)

foreach ($test in $networkTests) {
    if ($test.Target -and $test.Target -ne "N/A") {
        $result = Test-NetConnection -ComputerName $test.Target -Port $test.Port -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
        if ($result.TcpTestSucceeded) {
            Add-AuditResult "Network" $test.Name "OK" "Port $($test.Port) reachable" "Info"
        } else {
            Add-AuditResult "Network" $test.Name "WARNING" "Port $($test.Port) NOT reachable" "High"
        }
    }
}
Write-Host "      Network Tests Complete" -ForegroundColor Green

# ============================================================
# [24/25] INSTALLED PATCHES (Last 10)
# ============================================================
Write-Host "[24/$totalChecks] Checking Recent Patches..." -ForegroundColor Yellow

$hotfixes = Get-HotFix | Sort-Object InstalledOn -Descending -ErrorAction SilentlyContinue | Select-Object -First 10
foreach ($hf in $hotfixes) {
    if ($hf.InstalledOn) {
        Add-AuditResult "Patches" $hf.HotFixID "INFO" "Installed: $($hf.InstalledOn.ToString('yyyy-MM-dd'))" "Info"
    }
}

$lastPatch = $hotfixes | Where-Object { $_.InstalledOn } | Select-Object -First 1
if ($lastPatch -and $lastPatch.InstalledOn) {
    $daysSincePatch = ((Get-Date) - $lastPatch.InstalledOn).Days
    if ($daysSincePatch -gt 60) {
        Add-AuditResult "Patches" "Patch Status" "WARNING" "Last patch was $daysSincePatch days ago" "High"
        Write-Host "      Last Patch: $daysSincePatch days ago - WARNING" -ForegroundColor Yellow
    } else {
        Add-AuditResult "Patches" "Patch Status" "OK" "Last patch was $daysSincePatch days ago" "Info"
        Write-Host "      Last Patch: $daysSincePatch days ago - OK" -ForegroundColor Green
    }
}

# ============================================================
# [25/25] SUMMARY COUNTS
# ============================================================
Write-Host "[25/$totalChecks] Generating Summary..." -ForegroundColor Yellow

$criticalCount = ($auditResults | Where-Object { $_.Severity -eq "Critical" }).Count
$highCount = ($auditResults | Where-Object { $_.Severity -eq "High" }).Count
$warningCount = ($auditResults | Where-Object { $_.Status -eq "WARNING" -or $_.Status -eq "STOPPED" -or $_.Status -eq "ERROR" }).Count

# ============================================================
# GENERATE REPORTS - SAVE TO BOTH LOCATIONS
# ============================================================
Write-Host ""
Write-Host "Generating Reports..." -ForegroundColor Yellow

$csvFileName = "L2_Audit_${computerName}_$timestamp.csv"
$htmlFileName = "L2_Audit_${computerName}_$timestamp.html"

$csvPath = Join-Path $OutputPath $csvFileName
$htmlPath = Join-Path $OutputPath $htmlFileName
$csvPath2 = Join-Path $SecondaryPath $csvFileName
$htmlPath2 = Join-Path $SecondaryPath $htmlFileName

# Export CSV to BOTH locations
$auditResults | Export-Csv -Path $csvPath -NoTypeInformation
$auditResults | Export-Csv -Path $csvPath2 -NoTypeInformation
Write-Host "      CSV saved to: $csvPath" -ForegroundColor Green
Write-Host "      CSV saved to: $csvPath2" -ForegroundColor Green

# HTML Report
$html = @"
<!DOCTYPE html>
<html>
<head>
    <title>L2 Server Audit - $computerName</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 1400px; margin: 0 auto; }
        .header { background: linear-gradient(135deg, #5c2d91 0%, #8e44ad 100%); color: white; padding: 30px; border-radius: 8px; margin-bottom: 20px; }
        .header h1 { margin: 0; font-size: 28px; }
        .badge { display: inline-block; padding: 5px 15px; border-radius: 20px; font-size: 12px; margin-top: 10px; }
        .badge-safe { background: #27ae60; }
        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 15px; margin-bottom: 30px; }
        .card { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); text-align: center; }
        .card h3 { margin: 0 0 10px 0; color: #666; font-size: 11px; }
        .card .value { font-size: 24px; font-weight: bold; }
        .card .value.critical { color: #e74c3c; }
        .card .value.warning { color: #f39c12; }
        .card .value.ok { color: #27ae60; }
        table { width: 100%; border-collapse: collapse; background: white; box-shadow: 0 2px 10px rgba(0,0,0,0.1); margin-bottom: 20px; }
        th { background: #5c2d91; color: white; padding: 12px 8px; text-align: left; font-size: 12px; }
        td { padding: 10px 8px; border-bottom: 1px solid #eee; font-size: 11px; }
        tr:hover { background: #f8f4fc; }
        .status-ok { background: #d4edda; color: #155724; padding: 3px 8px; border-radius: 4px; }
        .status-warning { background: #fff3cd; color: #856404; padding: 3px 8px; border-radius: 4px; }
        .status-critical { background: #f8d7da; color: #721c24; padding: 3px 8px; border-radius: 4px; }
        .status-error { background: #f8d7da; color: #721c24; padding: 3px 8px; border-radius: 4px; }
        .status-info { background: #e7e7e7; color: #666; padding: 3px 8px; border-radius: 4px; }
        .status-stopped { background: #f8d7da; color: #721c24; padding: 3px 8px; border-radius: 4px; }
        .section { margin-bottom: 30px; }
        .section h2 { color: #5c2d91; border-bottom: 2px solid #5c2d91; padding-bottom: 10px; }
        .timestamp { color: rgba(255,255,255,0.8); margin-top: 10px; }
    </style>
</head>
<body>
<div class="container">
    <div class="header">
        <h1>🔍 L2 Server Diagnostic Audit Report v4.0</h1>
        <p>Computer: <strong>$computerName</strong> | Domain: <strong>$($cs.Domain)</strong></p>
        <p class="timestamp">Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") | Author: Syed Rizvi</p>
        <span class="badge badge-safe">✓ READ-ONLY AUDIT - NO CHANGES MADE</span>
    </div>
    
    <div class="summary">
        <div class="card">
            <h3>Total Checks</h3>
            <div class="value">$($auditResults.Count)</div>
        </div>
        <div class="card">
            <h3>Critical</h3>
            <div class="value critical">$criticalCount</div>
        </div>
        <div class="card">
            <h3>High Priority</h3>
            <div class="value warning">$highCount</div>
        </div>
        <div class="card">
            <h3>Warnings</h3>
            <div class="value warning">$warningCount</div>
        </div>
        <div class="card">
            <h3>Trust</h3>
            <div class="value $(if ($trustStatus -eq 'Healthy') { 'ok' } else { 'critical' })">$trustStatus</div>
        </div>
        <div class="card">
            <h3>GPO</h3>
            <div class="value $(if ($gpoStatus -eq 'Healthy') { 'ok' } else { 'warning' })">$gpoStatus</div>
        </div>
        <div class="card">
            <h3>CPU</h3>
            <div class="value $(if ($cpuLoad -lt 80) { 'ok' } elseif ($cpuLoad -lt 90) { 'warning' } else { 'critical' })">$cpuLoad%</div>
        </div>
        <div class="card">
            <h3>Memory</h3>
            <div class="value $(if ($memPercent -lt 80) { 'ok' } elseif ($memPercent -lt 90) { 'warning' } else { 'critical' })">$memPercent%</div>
        </div>
    </div>
    
    <div class="section">
        <h2>📋 All Audit Results</h2>
        <table>
            <tr>
                <th>Category</th>
                <th>Check</th>
                <th>Status</th>
                <th>Details</th>
                <th>Severity</th>
            </tr>
"@

foreach ($result in $auditResults) {
    $statusClass = switch ($result.Status) {
        "OK" { "status-ok" }
        "WARNING" { "status-warning" }
        "CRITICAL" { "status-critical" }
        "ERROR" { "status-error" }
        "STOPPED" { "status-stopped" }
        default { "status-info" }
    }
    
    $html += @"
            <tr>
                <td><strong>$($result.Category)</strong></td>
                <td>$($result.Check)</td>
                <td><span class="$statusClass">$($result.Status)</span></td>
                <td>$($result.Details)</td>
                <td>$($result.Severity)</td>
            </tr>
"@
}

$html += @"
        </table>
    </div>
    
    <div class="section">
        <h2>🛡️ Security Tools Summary</h2>
        <table>
            <tr><th>Tool</th><th>Status</th></tr>
"@

foreach ($tool in $securityToolsFound) {
    $parts = $tool -split ":"
    $toolStatus = if ($parts[1] -match "Running|Connected|Linked|OK") { "status-ok" } else { "status-warning" }
    $html += "<tr><td>$($parts[0])</td><td><span class='$toolStatus'>$($parts[1])</span></td></tr>"
}

if ($securityToolsFound.Count -eq 0) {
    $html += "<tr><td>No Security Tools</td><td><span class='status-critical'>NONE DETECTED</span></td></tr>"
}

$html += @"
        </table>
    </div>
    
    <div class="section">
        <h2>📊 Recent Event Errors (Last 20)</h2>
        <table>
            <tr><th>Log</th><th>Time</th><th>Event ID</th><th>Source</th><th>Message</th></tr>
"@

foreach ($evt in $eventErrors | Select-Object -First 20) {
    $html += "<tr><td>$($evt.Log)</td><td>$($evt.Time)</td><td>$($evt.ID)</td><td>$($evt.Source)</td><td>$($evt.Message)</td></tr>"
}

$html += @"
        </table>
    </div>
    
    <p style="margin-top:30px; color:#27ae60; font-weight:bold; text-align:center;">
        ✓ This is a READ-ONLY audit report. No changes were made to this server.
    </p>
</div>
</body>
</html>
"@

$html | Out-File -FilePath $htmlPath -Encoding UTF8
$html | Out-File -FilePath $htmlPath2 -Encoding UTF8
Write-Host "      HTML saved to: $htmlPath" -ForegroundColor Green
Write-Host "      HTML saved to: $htmlPath2" -ForegroundColor Green

# ============================================================
# FINAL SUMMARY
# ============================================================
Write-Host ""
Write-Host "=============================================================" -ForegroundColor Green
Write-Host "                    AUDIT COMPLETE!" -ForegroundColor Green
Write-Host "            *** NO CHANGES WERE MADE ***" -ForegroundColor Green
Write-Host "=============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "SUMMARY:" -ForegroundColor Yellow
Write-Host "  Total Checks:      $($auditResults.Count)" -ForegroundColor White
Write-Host "  Critical Issues:   $criticalCount" -ForegroundColor $(if ($criticalCount -gt 0) { "Red" } else { "Green" })
Write-Host "  High Priority:     $highCount" -ForegroundColor $(if ($highCount -gt 0) { "Yellow" } else { "Green" })
Write-Host "  Warnings:          $warningCount" -ForegroundColor $(if ($warningCount -gt 0) { "Yellow" } else { "Green" })
Write-Host ""
Write-Host "KEY STATUS:" -ForegroundColor Yellow
Write-Host "  Trust Status:      $trustStatus" -ForegroundColor $(if ($trustStatus -eq "Healthy") { "Green" } else { "Red" })
Write-Host "  GPO Status:        $gpoStatus" -ForegroundColor $(if ($gpoStatus -eq "Healthy") { "Green" } else { "Yellow" })
Write-Host "  Replication:       $replicationStatus" -ForegroundColor $(if ($replicationStatus -eq "Healthy" -or $replicationStatus -eq "N/A") { "Green" } else { "Red" })
Write-Host "  CPU Usage:         $cpuLoad%" -ForegroundColor $(if ($cpuLoad -lt 80) { "Green" } else { "Yellow" })
Write-Host "  Memory Usage:      $memPercent%" -ForegroundColor $(if ($memPercent -lt 80) { "Green" } else { "Yellow" })
Write-Host ""
Write-Host "OUTPUT FILES (Saved to BOTH locations):" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Location 1 (Desktop):" -ForegroundColor Cyan
Write-Host "    CSV:  $csvPath" -ForegroundColor White
Write-Host "    HTML: $htmlPath" -ForegroundColor White
Write-Host ""
Write-Host "  Location 2 (C:\L2_Reports):" -ForegroundColor Cyan
Write-Host "    CSV:  $csvPath2" -ForegroundColor White
Write-Host "    HTML: $htmlPath2" -ForegroundColor White
Write-Host ""

# Open HTML report
try {
    Start-Process $htmlPath
} catch {
    Write-Host "Open manually: $htmlPath" -ForegroundColor Yellow
}
