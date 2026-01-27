<#
.SYNOPSIS
    L2 Server Diagnostics - Audit Report
    Author: Syed Rizvi
    Version: 4.0
.DESCRIPTION
    Read-only diagnostic script for L2 team.
.EXAMPLE
    .\L2-ServerDiagnostics-Audit-v4.ps1
#>

param(
    [string]$OutputPath = "$env:USERPROFILE\Desktop",
    [string]$SecondaryPath = "C:\L2_Reports"
)

$ErrorActionPreference = "SilentlyContinue"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$computerName = $env:COMPUTERNAME

if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
if (-not (Test-Path $SecondaryPath)) { New-Item -ItemType Directory -Path $SecondaryPath -Force | Out-Null }

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
Write-Host "L2 SERVER DIAGNOSTICS - AUDIT REPORT v4.0" -ForegroundColor Cyan
Write-Host "Author: Syed Rizvi"
Write-Host "Computer: $computerName"
Write-Host "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host ""

$totalChecks = 25

Write-Host "[1/$totalChecks] Gathering System Information..." -ForegroundColor Yellow
$os = Get-WmiObject Win32_OperatingSystem
$cs = Get-WmiObject Win32_ComputerSystem
$lastBoot = $os.ConvertToDateTime($os.LastBootUpTime)
$uptime = (Get-Date) - $lastBoot

Add-AuditResult "System" "Computer Name" "INFO" $computerName "Info"
Add-AuditResult "System" "Operating System" "INFO" $os.Caption "Info"
Add-AuditResult "System" "Last Boot" "INFO" $lastBoot.ToString("yyyy-MM-dd HH:mm:ss") "Info"
Add-AuditResult "System" "Uptime" $(if ($uptime.Days -gt 30) { "WARNING" } else { "OK" }) "$($uptime.Days)d $($uptime.Hours)h $($uptime.Minutes)m" $(if ($uptime.Days -gt 30) { "Medium" } else { "Info" })
Add-AuditResult "System" "Domain" "INFO" $cs.Domain "Info"
Write-Host "      Uptime: $($uptime.Days)d $($uptime.Hours)h" -ForegroundColor Green

Write-Host "[2/$totalChecks] Checking Windows Activation..." -ForegroundColor Yellow
try {
    $licenseStatus = Get-WmiObject SoftwareLicensingProduct | Where-Object { $_.PartialProductKey -and $_.Name -like "*Windows*" } | Select-Object -First 1
    if ($licenseStatus) {
        if ($licenseStatus.LicenseStatus -eq 1) {
            Add-AuditResult "Activation" "Windows License" "OK" "Activated" "Info"
            Write-Host "      Windows: Activated" -ForegroundColor Green
        } else {
            Add-AuditResult "Activation" "Windows License" "CRITICAL" "NOT Activated" "Critical"
            Write-Host "      Windows: NOT ACTIVATED" -ForegroundColor Red
        }
    }
} catch { }

Write-Host "[3/$totalChecks] Checking Domain and Trust..." -ForegroundColor Yellow
$domainStatus = "Unknown"
$trustStatus = "Unknown"
$dcName = "N/A"

if ($cs.PartOfDomain) {
    try {
        $dcInfo = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        $dcName = $dcInfo.FindDomainController().Name
        Add-AuditResult "Domain" "Domain Controller" "OK" $dcName "Info"
        $secureChannel = Test-ComputerSecureChannel -ErrorAction SilentlyContinue
        if ($secureChannel) {
            $trustStatus = "Healthy"
            Add-AuditResult "Domain" "Trust Relationship" "OK" "Secure channel valid" "Info"
            Write-Host "      Trust: Healthy" -ForegroundColor Green
        } else {
            $trustStatus = "BROKEN"
            Add-AuditResult "Domain" "Trust Relationship" "CRITICAL" "Secure channel BROKEN" "Critical"
            Write-Host "      Trust: BROKEN" -ForegroundColor Red
        }
    } catch {
        Add-AuditResult "Domain" "Domain Status" "WARNING" "Cannot contact DC" "High"
        Write-Host "      Domain: Cannot contact DC" -ForegroundColor Red
    }
} else {
    Add-AuditResult "Domain" "Domain Status" "INFO" "Workgroup" "Info"
    Write-Host "      Domain: Workgroup" -ForegroundColor Yellow
}

Write-Host "[4/$totalChecks] Checking Group Policy..." -ForegroundColor Yellow
$gpoStatus = "Unknown"
try {
    $gpoEvents = Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-GroupPolicy/Operational'; Level=2,3; StartTime=(Get-Date).AddDays(-7)} -MaxEvents 10 -ErrorAction SilentlyContinue
    if ($gpoEvents) {
        $gpoStatus = "ERRORS"
        Add-AuditResult "GPO" "Group Policy Status" "WARNING" "$($gpoEvents.Count) errors in last 7 days" "High"
        Write-Host "      GPO: $($gpoEvents.Count) errors" -ForegroundColor Yellow
    } else {
        $gpoStatus = "Healthy"
        Add-AuditResult "GPO" "Group Policy Status" "OK" "No errors" "Info"
        Write-Host "      GPO: Healthy" -ForegroundColor Green
    }
} catch { }

Write-Host "[5/$totalChecks] Checking AD Replication..." -ForegroundColor Yellow
$isDC = (Get-WmiObject Win32_ComputerSystem).DomainRole -ge 4
$replicationStatus = "N/A"
if ($isDC) {
    try {
        $replStatus = repadmin /showrepl /csv 2>$null | ConvertFrom-Csv
        $replErrors = $replStatus | Where-Object { $_.'Number of Failures' -gt 0 }
        if ($replErrors) {
            $replicationStatus = "ERRORS"
            Add-AuditResult "Replication" "AD Replication" "CRITICAL" "$($replErrors.Count) failures" "Critical"
            Write-Host "      Replication: ERRORS" -ForegroundColor Red
        } else {
            $replicationStatus = "Healthy"
            Add-AuditResult "Replication" "AD Replication" "OK" "Healthy" "Info"
            Write-Host "      Replication: Healthy" -ForegroundColor Green
        }
    } catch { }
} else {
    Add-AuditResult "Replication" "AD Replication" "INFO" "Not a DC" "Info"
    Write-Host "      Replication: N/A (not DC)" -ForegroundColor Gray
}

Write-Host "[6/$totalChecks] Checking CPU Usage..." -ForegroundColor Yellow
$cpuLoad = 0
try {
    $cpuLoad = [math]::Round((Get-WmiObject Win32_Processor | Measure-Object -Property LoadPercentage -Average).Average, 1)
    if ($cpuLoad -gt 90) {
        Add-AuditResult "CPU" "CPU Usage" "CRITICAL" "$cpuLoad%" "Critical"
        Write-Host "      CPU: $cpuLoad% - CRITICAL" -ForegroundColor Red
    } elseif ($cpuLoad -gt 80) {
        Add-AuditResult "CPU" "CPU Usage" "WARNING" "$cpuLoad%" "High"
        Write-Host "      CPU: $cpuLoad% - WARNING" -ForegroundColor Yellow
    } else {
        Add-AuditResult "CPU" "CPU Usage" "OK" "$cpuLoad%" "Info"
        Write-Host "      CPU: $cpuLoad%" -ForegroundColor Green
    }
} catch { }

Write-Host "[7/$totalChecks] Checking Top Processes..." -ForegroundColor Yellow
try {
    $topCPU = Get-Process | Sort-Object CPU -Descending | Select-Object -First 5
    foreach ($proc in $topCPU) {
        Add-AuditResult "Processes" "Top CPU: $($proc.ProcessName)" "INFO" "CPU: $([math]::Round($proc.CPU, 2))s, Mem: $([math]::Round($proc.WorkingSet64/1MB, 1))MB" "Info"
    }
    Write-Host "      Top Processes: Captured" -ForegroundColor Green
} catch { }

Write-Host "[8/$totalChecks] Checking Memory..." -ForegroundColor Yellow
$totalMem = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
$freeMem = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
$usedMem = $totalMem - $freeMem
$memPercent = [math]::Round(($usedMem / $totalMem) * 100, 1)
if ($memPercent -gt 90) {
    Add-AuditResult "Memory" "RAM Usage" "CRITICAL" "$memPercent% used" "Critical"
    Write-Host "      Memory: $memPercent% - CRITICAL" -ForegroundColor Red
} elseif ($memPercent -gt 80) {
    Add-AuditResult "Memory" "RAM Usage" "WARNING" "$memPercent% used" "High"
    Write-Host "      Memory: $memPercent% - WARNING" -ForegroundColor Yellow
} else {
    Add-AuditResult "Memory" "RAM Usage" "OK" "$memPercent% used" "Info"
    Write-Host "      Memory: $memPercent%" -ForegroundColor Green
}

Write-Host "[9/$totalChecks] Checking Page File..." -ForegroundColor Yellow
try {
    $pageFile = Get-WmiObject Win32_PageFileUsage
    if ($pageFile) {
        $pfPercent = if ($pageFile.AllocatedBaseSize -gt 0) { [math]::Round(($pageFile.CurrentUsage / $pageFile.AllocatedBaseSize) * 100, 1) } else { 0 }
        if ($pfPercent -gt 80) {
            Add-AuditResult "PageFile" "Page File" "WARNING" "$pfPercent% used" "High"
        } else {
            Add-AuditResult "PageFile" "Page File" "OK" "$pfPercent% used" "Info"
        }
        Write-Host "      Page File: $pfPercent%" -ForegroundColor Green
    }
} catch { }

Write-Host "[10/$totalChecks] Checking Disk Space..." -ForegroundColor Yellow
$disks = Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3"
foreach ($disk in $disks) {
    $freePercent = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 1)
    $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
    if ($freePercent -lt 10) {
        Add-AuditResult "Disk" "$($disk.DeviceID) Space" "CRITICAL" "$freeGB GB free ($freePercent%)" "Critical"
        Write-Host "      $($disk.DeviceID) $freePercent% free - CRITICAL" -ForegroundColor Red
    } elseif ($freePercent -lt 20) {
        Add-AuditResult "Disk" "$($disk.DeviceID) Space" "WARNING" "$freeGB GB free ($freePercent%)" "High"
        Write-Host "      $($disk.DeviceID) $freePercent% free - WARNING" -ForegroundColor Yellow
    } else {
        Add-AuditResult "Disk" "$($disk.DeviceID) Space" "OK" "$freeGB GB free ($freePercent%)" "Info"
        Write-Host "      $($disk.DeviceID) $freePercent% free" -ForegroundColor Green
    }
}

Write-Host "[11/$totalChecks] Checking Critical Services..." -ForegroundColor Yellow
$criticalServices = @("Netlogon", "W32Time", "gpsvc", "Dnscache", "LanmanWorkstation", "LanmanServer", "EventLog", "Schedule", "TermService", "CryptSvc", "BITS")
$stoppedCritical = 0
foreach ($svcName in $criticalServices) {
    $service = Get-Service -Name $svcName -ErrorAction SilentlyContinue
    if ($service) {
        if ($service.Status -eq "Running") {
            Add-AuditResult "Services" $svcName "OK" "Running" "Info"
        } else {
            Add-AuditResult "Services" $svcName "STOPPED" "$($service.Status)" "Critical"
            $stoppedCritical++
        }
    }
}
if ($stoppedCritical -gt 0) {
    Write-Host "      Services: $stoppedCritical STOPPED" -ForegroundColor Red
} else {
    Write-Host "      Services: All Running" -ForegroundColor Green
}

Write-Host "[12/$totalChecks] Checking NIC Teaming..." -ForegroundColor Yellow
try {
    $nicTeams = Get-NetLbfoTeam -ErrorAction SilentlyContinue
    if ($nicTeams) {
        foreach ($team in $nicTeams) {
            if ($team.Status -eq "Up") {
                Add-AuditResult "NIC Team" $team.Name "OK" "Status: Up" "Info"
                Write-Host "      NIC Team $($team.Name): UP" -ForegroundColor Green
            } else {
                Add-AuditResult "NIC Team" $team.Name "CRITICAL" "Status: $($team.Status)" "Critical"
                Write-Host "      NIC Team $($team.Name): DOWN" -ForegroundColor Red
            }
        }
    } else {
        Add-AuditResult "NIC Team" "NIC Teaming" "INFO" "Not configured" "Info"
        Write-Host "      NIC Teaming: Not configured" -ForegroundColor Gray
    }
} catch { }

Write-Host "[13/$totalChecks] Checking SSL Certificates..." -ForegroundColor Yellow
try {
    $certs = Get-ChildItem Cert:\LocalMachine\My -ErrorAction SilentlyContinue
    $expiredCerts = 0
    $expiringCerts = 0
    foreach ($cert in $certs) {
        $daysUntilExpiry = ($cert.NotAfter - (Get-Date)).Days
        $certName = if ($cert.FriendlyName) { $cert.FriendlyName } else { $cert.Subject.Substring(0, [Math]::Min(40, $cert.Subject.Length)) }
        if ($daysUntilExpiry -lt 0) {
            Add-AuditResult "SSL Certs" $certName "CRITICAL" "EXPIRED" "Critical"
            $expiredCerts++
        } elseif ($daysUntilExpiry -lt 30) {
            Add-AuditResult "SSL Certs" $certName "WARNING" "Expires in $daysUntilExpiry days" "High"
            $expiringCerts++
        }
    }
    if ($expiredCerts -gt 0) {
        Write-Host "      SSL Certs: $expiredCerts EXPIRED" -ForegroundColor Red
    } elseif ($expiringCerts -gt 0) {
        Write-Host "      SSL Certs: $expiringCerts expiring soon" -ForegroundColor Yellow
    } else {
        Add-AuditResult "SSL Certs" "Certificate Status" "OK" "All valid" "Info"
        Write-Host "      SSL Certs: All valid" -ForegroundColor Green
    }
} catch { }

Write-Host "[14/$totalChecks] Checking Security Tools..." -ForegroundColor Yellow
$securityToolsFound = @()

$trellixServices = @("macmnsvc", "masvc", "mfefire", "mfemms", "McShield", "McAfeeFramework")
$trellixRunning = 0
foreach ($svcName in $trellixServices) {
    $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
    if ($svc -and $svc.Status -eq "Running") {
        Add-AuditResult "Trellix" $svcName "OK" "Running" "Info"
        $trellixRunning++
    } elseif ($svc) {
        Add-AuditResult "Trellix" $svcName "STOPPED" "$($svc.Status)" "High"
    }
}
if ($trellixRunning -gt 0) {
    $securityToolsFound += "Trellix: $trellixRunning services"
    Write-Host "      Trellix: $trellixRunning services running" -ForegroundColor Green
}

$trendServices = @("ds_agent", "TmListen", "ntrtscan")
$trendRunning = 0
foreach ($svcName in $trendServices) {
    $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
    if ($svc -and $svc.Status -eq "Running") {
        Add-AuditResult "Trend Micro" $svcName "OK" "Running" "Info"
        $trendRunning++
    } elseif ($svc) {
        Add-AuditResult "Trend Micro" $svcName "STOPPED" "$($svc.Status)" "High"
    }
}
if ($trendRunning -gt 0) {
    $securityToolsFound += "Trend Micro: $trendRunning services"
    Write-Host "      Trend Micro: $trendRunning services running" -ForegroundColor Green
}

$nessusService = Get-Service -Name "Tenable Nessus Agent" -ErrorAction SilentlyContinue
if ($nessusService) {
    if ($nessusService.Status -eq "Running") {
        Add-AuditResult "Nessus" "Agent Service" "OK" "Running" "Info"
        $securityToolsFound += "Nessus: Running"
        Write-Host "      Nessus: Running" -ForegroundColor Green
        $nessuscli = "C:\Program Files\Tenable\Nessus Agent\nessuscli.exe"
        if (Test-Path $nessuscli) {
            $nessusStatus = & $nessuscli agent status 2>&1 | Out-String
            if ($nessusStatus -match "Linked to") {
                Add-AuditResult "Nessus" "Link Status" "OK" "Linked" "Info"
            } else {
                Add-AuditResult "Nessus" "Link Status" "WARNING" "Not linked" "High"
            }
        }
    } else {
        Add-AuditResult "Nessus" "Agent Service" "STOPPED" "$($nessusService.Status)" "High"
        Write-Host "      Nessus: STOPPED" -ForegroundColor Red
    }
}

$otherTools = @(@{Name="CrowdStrike"; Service="CSFalconService"}, @{Name="Windows Defender"; Service="WinDefend"})
foreach ($tool in $otherTools) {
    $svc = Get-Service -Name $tool.Service -ErrorAction SilentlyContinue
    if ($svc -and $svc.Status -eq "Running") {
        Add-AuditResult "Security Tools" $tool.Name "OK" "Running" "Info"
        $securityToolsFound += "$($tool.Name): Running"
        Write-Host "      $($tool.Name): Running" -ForegroundColor Green
    } elseif ($svc) {
        Add-AuditResult "Security Tools" $tool.Name "STOPPED" "$($svc.Status)" "High"
    }
}
if ($securityToolsFound.Count -eq 0) {
    Add-AuditResult "Security Tools" "Overall" "CRITICAL" "No security tools detected" "Critical"
    Write-Host "      Security Tools: NONE DETECTED" -ForegroundColor Red
}

Write-Host "[15/$totalChecks] Checking Time Sync..." -ForegroundColor Yellow
try {
    $w32tm = w32tm /query /status 2>&1 | Out-String
    $timeSource = if ($w32tm -match "Source:\s*(.+)") { $Matches[1].Trim() } else { "Unknown" }
    if ($timeSource -and $timeSource -notmatch "Local CMOS|Free-running") {
        Add-AuditResult "Time" "Time Source" "OK" $timeSource "Info"
        Write-Host "      Time: Synced" -ForegroundColor Green
    } else {
        Add-AuditResult "Time" "Time Source" "WARNING" "Not synced to domain" "High"
        Write-Host "      Time: Not synced" -ForegroundColor Yellow
    }
} catch { }

Write-Host "[16/$totalChecks] Checking Pending Reboot..." -ForegroundColor Yellow
$pendingReboot = $false
if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") { $pendingReboot = $true }
if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") { $pendingReboot = $true }
if ($pendingReboot) {
    Add-AuditResult "Reboot" "Pending Reboot" "WARNING" "Reboot required" "High"
    Write-Host "      Pending Reboot: YES" -ForegroundColor Yellow
} else {
    Add-AuditResult "Reboot" "Pending Reboot" "OK" "None" "Info"
    Write-Host "      Pending Reboot: None" -ForegroundColor Green
}

Write-Host "[17/$totalChecks] Checking RDP..." -ForegroundColor Yellow
$rdpEnabled = (Get-ItemProperty "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -ErrorAction SilentlyContinue).fDenyTSConnections -eq 0
if ($rdpEnabled) {
    Add-AuditResult "RDP" "RDP Status" "OK" "Enabled" "Info"
    Write-Host "      RDP: Enabled" -ForegroundColor Green
} else {
    Add-AuditResult "RDP" "RDP Status" "WARNING" "Disabled" "Medium"
    Write-Host "      RDP: Disabled" -ForegroundColor Yellow
}

Write-Host "[18/$totalChecks] Checking Firewall..." -ForegroundColor Yellow
$fwProfiles = Get-NetFirewallProfile -ErrorAction SilentlyContinue
foreach ($profile in $fwProfiles) {
    Add-AuditResult "Firewall" "$($profile.Name) Profile" $(if ($profile.Enabled) { "OK" } else { "WARNING" }) $(if ($profile.Enabled) { "Enabled" } else { "Disabled" }) "Info"
}
Write-Host "      Firewall: Checked" -ForegroundColor Green

Write-Host "[19/$totalChecks] Checking Scheduled Tasks..." -ForegroundColor Yellow
try {
    $failedTasks = Get-ScheduledTask | Get-ScheduledTaskInfo -ErrorAction SilentlyContinue | Where-Object { $_.LastTaskResult -ne 0 -and $_.LastTaskResult -ne 267009 -and $_.LastRunTime -gt (Get-Date).AddDays(-7) }
    $failedCount = ($failedTasks | Measure-Object).Count
    if ($failedCount -gt 0) {
        foreach ($task in $failedTasks | Select-Object -First 5) {
            Add-AuditResult "Scheduled Tasks" $task.TaskName "WARNING" "Failed: $($task.LastTaskResult)" "Medium"
        }
        Write-Host "      Scheduled Tasks: $failedCount failed" -ForegroundColor Yellow
    } else {
        Add-AuditResult "Scheduled Tasks" "Task Status" "OK" "No failures" "Info"
        Write-Host "      Scheduled Tasks: OK" -ForegroundColor Green
    }
} catch { }

Write-Host "[20/$totalChecks] Checking Backup Status..." -ForegroundColor Yellow
try {
    $wsbJob = Get-WBJob -Previous 1 -ErrorAction SilentlyContinue
    if ($wsbJob) {
        if ($wsbJob.HResult -eq 0) {
            Add-AuditResult "Backup" "Windows Backup" "OK" "Last backup successful" "Info"
            Write-Host "      Backup: Successful" -ForegroundColor Green
        } else {
            Add-AuditResult "Backup" "Windows Backup" "WARNING" "Last backup failed" "High"
            Write-Host "      Backup: FAILED" -ForegroundColor Red
        }
    } else {
        Add-AuditResult "Backup" "Backup Status" "INFO" "Not configured" "Info"
        Write-Host "      Backup: Not configured" -ForegroundColor Gray
    }
} catch { Add-AuditResult "Backup" "Backup Status" "INFO" "Not available" "Info" }

Write-Host "[21/$totalChecks] Checking Failed Logins..." -ForegroundColor Yellow
try {
    $failedLogins = Get-WinEvent -FilterHashtable @{LogName='Security'; Id=4625; StartTime=(Get-Date).AddHours(-24)} -MaxEvents 100 -ErrorAction SilentlyContinue
    $failedCount = if ($failedLogins) { $failedLogins.Count } else { 0 }
    if ($failedCount -gt 50) {
        Add-AuditResult "Security" "Failed Logins (24h)" "CRITICAL" "$failedCount attempts" "Critical"
        Write-Host "      Failed Logins: $failedCount - CRITICAL" -ForegroundColor Red
    } elseif ($failedCount -gt 20) {
        Add-AuditResult "Security" "Failed Logins (24h)" "WARNING" "$failedCount attempts" "High"
        Write-Host "      Failed Logins: $failedCount - WARNING" -ForegroundColor Yellow
    } else {
        Add-AuditResult "Security" "Failed Logins (24h)" "OK" "$failedCount attempts" "Info"
        Write-Host "      Failed Logins: $failedCount" -ForegroundColor Green
    }
} catch { }

Write-Host "[22/$totalChecks] Gathering Event Errors..." -ForegroundColor Yellow
$eventErrors = @()
$sysErrors = Get-WinEvent -FilterHashtable @{LogName='System'; Level=2} -MaxEvents 10 -ErrorAction SilentlyContinue
foreach ($evt in $sysErrors) {
    $msg = if ($evt.Message.Length -gt 150) { $evt.Message.Substring(0, 150) } else { $evt.Message }
    $eventErrors += [PSCustomObject]@{ Log="System"; Time=$evt.TimeCreated.ToString("yyyy-MM-dd HH:mm"); ID=$evt.Id; Source=$evt.ProviderName; Message=$msg }
    Add-AuditResult "Events" "System Error $($evt.Id)" "ERROR" "$($evt.ProviderName): $msg" "Medium"
}
$appErrors = Get-WinEvent -FilterHashtable @{LogName='Application'; Level=2} -MaxEvents 10 -ErrorAction SilentlyContinue
foreach ($evt in $appErrors) {
    $msg = if ($evt.Message.Length -gt 150) { $evt.Message.Substring(0, 150) } else { $evt.Message }
    $eventErrors += [PSCustomObject]@{ Log="Application"; Time=$evt.TimeCreated.ToString("yyyy-MM-dd HH:mm"); ID=$evt.Id; Source=$evt.ProviderName; Message=$msg }
    Add-AuditResult "Events" "App Error $($evt.Id)" "ERROR" "$($evt.ProviderName): $msg" "Medium"
}
Write-Host "      Event Errors: $($eventErrors.Count) found" -ForegroundColor $(if ($eventErrors.Count -gt 10) { "Yellow" } else { "Green" })

Write-Host "[23/$totalChecks] Checking Network..." -ForegroundColor Yellow
if ($dcName -and $dcName -ne "N/A") {
    $ldapTest = Test-NetConnection -ComputerName $dcName -Port 389 -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    if ($ldapTest.TcpTestSucceeded) {
        Add-AuditResult "Network" "LDAP to DC" "OK" "Reachable" "Info"
    } else {
        Add-AuditResult "Network" "LDAP to DC" "WARNING" "Not reachable" "High"
    }
}
Write-Host "      Network: Checked" -ForegroundColor Green

Write-Host "[24/$totalChecks] Checking Patches..." -ForegroundColor Yellow
$hotfixes = Get-HotFix | Sort-Object InstalledOn -Descending -ErrorAction SilentlyContinue | Select-Object -First 5
foreach ($hf in $hotfixes) {
    if ($hf.InstalledOn) { Add-AuditResult "Patches" $hf.HotFixID "INFO" "Installed: $($hf.InstalledOn.ToString('yyyy-MM-dd'))" "Info" }
}
$lastPatch = $hotfixes | Where-Object { $_.InstalledOn } | Select-Object -First 1
if ($lastPatch) {
    $daysSincePatch = ((Get-Date) - $lastPatch.InstalledOn).Days
    if ($daysSincePatch -gt 60) {
        Add-AuditResult "Patches" "Patch Status" "WARNING" "Last patch $daysSincePatch days ago" "High"
        Write-Host "      Patches: $daysSincePatch days ago - WARNING" -ForegroundColor Yellow
    } else {
        Add-AuditResult "Patches" "Patch Status" "OK" "Last patch $daysSincePatch days ago" "Info"
        Write-Host "      Patches: $daysSincePatch days ago" -ForegroundColor Green
    }
}

Write-Host "[25/$totalChecks] Generating Summary..." -ForegroundColor Yellow
$criticalCount = ($auditResults | Where-Object { $_.Severity -eq "Critical" }).Count
$highCount = ($auditResults | Where-Object { $_.Severity -eq "High" }).Count
$warningCount = ($auditResults | Where-Object { $_.Status -eq "WARNING" -or $_.Status -eq "STOPPED" -or $_.Status -eq "ERROR" }).Count

Write-Host ""
Write-Host "Generating Reports..." -ForegroundColor Yellow

$csvFileName = "L2_Audit_${computerName}_$timestamp.csv"
$htmlFileName = "L2_Audit_${computerName}_$timestamp.html"
$csvPath = Join-Path $OutputPath $csvFileName
$htmlPath = Join-Path $OutputPath $htmlFileName
$csvPath2 = Join-Path $SecondaryPath $csvFileName
$htmlPath2 = Join-Path $SecondaryPath $htmlFileName

$auditResults | Export-Csv -Path $csvPath -NoTypeInformation
$auditResults | Export-Csv -Path $csvPath2 -NoTypeInformation

$html = "<!DOCTYPE html><html><head><title>L2 Server Audit - $computerName</title><style>body{font-family:Segoe UI,Arial,sans-serif;margin:20px;background:#f5f5f5}.container{max-width:1400px;margin:0 auto}.header{background:linear-gradient(135deg,#2c3e50 0%,#3498db 100%);color:white;padding:30px;margin-bottom:20px}.header h1{margin:0;font-size:24px}.summary{display:flex;flex-wrap:wrap;gap:15px;margin-bottom:30px}.card{background:white;padding:20px;min-width:120px;text-align:center}.card h3{margin:0 0 10px 0;color:#666;font-size:11px}.card .value{font-size:24px;font-weight:bold}.critical{color:#e74c3c}.warning{color:#f39c12}.ok{color:#27ae60}table{width:100%;border-collapse:collapse;background:white;margin-bottom:20px}th{background:#2c3e50;color:white;padding:10px;text-align:left;font-size:12px}td{padding:8px;border-bottom:1px solid #eee;font-size:11px}tr:hover{background:#f8f9fa}.status-ok{background:#d4edda;color:#155724;padding:2px 8px}.status-warning{background:#fff3cd;color:#856404;padding:2px 8px}.status-critical{background:#f8d7da;color:#721c24;padding:2px 8px}.status-error{background:#f8d7da;color:#721c24;padding:2px 8px}.status-stopped{background:#f8d7da;color:#721c24;padding:2px 8px}.status-info{background:#e9ecef;color:#666;padding:2px 8px}</style></head><body><div class='container'><div class='header'><h1>L2 Server Diagnostic Audit Report</h1><p>Computer: $computerName | Domain: $($cs.Domain)</p><p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Author: Syed Rizvi</p></div><div class='summary'><div class='card'><h3>Total Checks</h3><div class='value'>$($auditResults.Count)</div></div><div class='card'><h3>Critical</h3><div class='value critical'>$criticalCount</div></div><div class='card'><h3>High</h3><div class='value warning'>$highCount</div></div><div class='card'><h3>Warnings</h3><div class='value warning'>$warningCount</div></div><div class='card'><h3>Trust</h3><div class='value $(if ($trustStatus -eq 'Healthy') { 'ok' } else { 'critical' })'>$trustStatus</div></div><div class='card'><h3>GPO</h3><div class='value $(if ($gpoStatus -eq 'Healthy') { 'ok' } else { 'warning' })'>$gpoStatus</div></div><div class='card'><h3>CPU</h3><div class='value'>$cpuLoad%</div></div><div class='card'><h3>Memory</h3><div class='value'>$memPercent%</div></div></div><h2>Audit Results</h2><table><tr><th>Category</th><th>Check</th><th>Status</th><th>Details</th><th>Severity</th></tr>"

foreach ($result in $auditResults) {
    $statusClass = switch ($result.Status) { "OK" { "status-ok" } "WARNING" { "status-warning" } "CRITICAL" { "status-critical" } "ERROR" { "status-error" } "STOPPED" { "status-stopped" } default { "status-info" } }
    $html += "<tr><td>$($result.Category)</td><td>$($result.Check)</td><td><span class='$statusClass'>$($result.Status)</span></td><td>$($result.Details)</td><td>$($result.Severity)</td></tr>"
}

$html += "</table><h2>Security Tools</h2><table><tr><th>Tool</th><th>Status</th></tr>"
foreach ($tool in $securityToolsFound) { $html += "<tr><td>$tool</td><td><span class='status-ok'>Running</span></td></tr>" }
$html += "</table><h2>Recent Event Errors</h2><table><tr><th>Log</th><th>Time</th><th>ID</th><th>Source</th><th>Message</th></tr>"
foreach ($evt in $eventErrors | Select-Object -First 20) { $html += "<tr><td>$($evt.Log)</td><td>$($evt.Time)</td><td>$($evt.ID)</td><td>$($evt.Source)</td><td>$($evt.Message)</td></tr>" }
$html += "</table></div></body></html>"

$html | Out-File -FilePath $htmlPath -Encoding UTF8
$html | Out-File -FilePath $htmlPath2 -Encoding UTF8

Write-Host ""
Write-Host "AUDIT COMPLETE" -ForegroundColor Green
Write-Host ""
Write-Host "Summary:"
Write-Host "  Total Checks:    $($auditResults.Count)"
Write-Host "  Critical:        $criticalCount" -ForegroundColor $(if ($criticalCount -gt 0) { "Red" } else { "Green" })
Write-Host "  High Priority:   $highCount" -ForegroundColor $(if ($highCount -gt 0) { "Yellow" } else { "Green" })
Write-Host ""
Write-Host "Reports saved to:"
Write-Host "  $csvPath"
Write-Host "  $htmlPath"
Write-Host "  $csvPath2"
Write-Host "  $htmlPath2"
Write-Host ""

try { Start-Process $htmlPath } catch { }
