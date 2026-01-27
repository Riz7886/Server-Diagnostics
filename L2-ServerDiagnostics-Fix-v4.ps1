<#
.SYNOPSIS
    L2 Server Diagnostics - FIX SCRIPT (WITH PROMPTS)
    Author: Syed Rizvi
    Version: 4.0

.DESCRIPTION
    Safe FIX script for L2 team - PROMPTS before each action.
    - Prompts before every fix
    - Shows what will be done
    - Safe operations only
    - Logs all actions taken

    Run AUDIT script first to identify issues!

.EXAMPLE
    .\L2-ServerDiagnostics-Fix-v4.ps1
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "$env:USERPROFILE\Desktop",
    
    [Parameter(Mandatory=$false)]
    [string]$SecondaryPath = "C:\L2_Reports"
)

$ErrorActionPreference = "Stop"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$computerName = $env:COMPUTERNAME

# Create directories
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
if (-not (Test-Path $SecondaryPath)) { New-Item -ItemType Directory -Path $SecondaryPath -Force | Out-Null }

$logPath = Join-Path $OutputPath "L2_Fix_Log_${computerName}_$timestamp.txt"
$logPath2 = Join-Path $SecondaryPath "L2_Fix_Log_${computerName}_$timestamp.txt"

function Write-Log {
    param($Message, $Color = "White")
    $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
    Add-Content -Path $logPath -Value $logEntry
    Add-Content -Path $logPath2 -Value $logEntry
    Write-Host $Message -ForegroundColor $Color
}

function Prompt-Action {
    param($ActionName, $Description)
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  ACTION: $ActionName" -ForegroundColor Yellow
    Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Description: $Description" -ForegroundColor White
    Write-Host ""
    $response = Read-Host "  Do you want to proceed? (Y/N)"
    return $response -eq "Y" -or $response -eq "y"
}

Write-Host ""
Write-Host "=============================================================" -ForegroundColor Magenta
Write-Host "    L2 SERVER DIAGNOSTICS - FIX SCRIPT v4.0" -ForegroundColor Magenta
Write-Host "    Author: Syed Rizvi" -ForegroundColor Magenta
Write-Host "=============================================================" -ForegroundColor Magenta
Write-Host ""
Write-Host "*** THIS SCRIPT WILL PROMPT BEFORE EVERY ACTION ***" -ForegroundColor Yellow
Write-Host "*** RUN THE AUDIT SCRIPT FIRST TO IDENTIFY ISSUES ***" -ForegroundColor Yellow
Write-Host ""
Write-Host "Computer: $computerName" -ForegroundColor White
Write-Host "Log Files:" -ForegroundColor Cyan
Write-Host "  1. $logPath" -ForegroundColor White
Write-Host "  2. $logPath2" -ForegroundColor White
Write-Host ""

Write-Log "L2 Fix Script v4.0 Started on $computerName"

$fixCount = 0
$skipCount = 0

# ============================================================
# FIX 1: RESTART STOPPED CRITICAL SERVICES
# ============================================================
Write-Host ""
Write-Host "[FIX 1/12] Checking Stopped Services..." -ForegroundColor Yellow

$criticalServices = @(
    @{Name="Netlogon"; Display="Netlogon"},
    @{Name="W32Time"; Display="Windows Time"},
    @{Name="gpsvc"; Display="Group Policy Client"},
    @{Name="Dnscache"; Display="DNS Client"},
    @{Name="TermService"; Display="Remote Desktop Services"},
    @{Name="EventLog"; Display="Windows Event Log"},
    @{Name="CryptSvc"; Display="Cryptographic Services"},
    @{Name="LanmanWorkstation"; Display="Workstation"},
    @{Name="LanmanServer"; Display="Server"},
    @{Name="BITS"; Display="BITS (Background Transfer)"},
    @{Name="Schedule"; Display="Task Scheduler"}
)

foreach ($svc in $criticalServices) {
    $service = Get-Service -Name $svc.Name -ErrorAction SilentlyContinue
    if ($service -and $service.Status -ne "Running") {
        if (Prompt-Action "Start $($svc.Display) Service" "This service is currently $($service.Status). Starting it may resolve connectivity/authentication issues.") {
            try {
                Start-Service -Name $svc.Name -ErrorAction Stop
                Write-Log "SUCCESS: Started $($svc.Display) service" "Green"
                $fixCount++
            } catch {
                Write-Log "FAILED: Could not start $($svc.Display) - $($_.Exception.Message)" "Red"
            }
        } else {
            Write-Log "SKIPPED: User chose not to start $($svc.Display)" "Yellow"
            $skipCount++
        }
    }
}

# ============================================================
# FIX 2: RESTART SECURITY TOOL SERVICES
# ============================================================
Write-Host ""
Write-Host "[FIX 2/12] Checking Security Tool Services..." -ForegroundColor Yellow

$securityServices = @(
    @{Name="Tenable Nessus Agent"; Display="Nessus Agent"},
    @{Name="mfefire"; Display="Trellix Firewall"},
    @{Name="mfemms"; Display="Trellix Management"},
    @{Name="macmnsvc"; Display="Trellix Agent"},
    @{Name="McShield"; Display="Trellix Scanner"},
    @{Name="TmListen"; Display="Trend Micro"},
    @{Name="ntrtscan"; Display="Trend Real-Time Scan"},
    @{Name="ds_agent"; Display="Trend Deep Security"},
    @{Name="CSFalconService"; Display="CrowdStrike"},
    @{Name="WinDefend"; Display="Windows Defender"}
)

foreach ($svc in $securityServices) {
    $service = Get-Service -Name $svc.Name -ErrorAction SilentlyContinue
    if ($service -and $service.Status -ne "Running") {
        if (Prompt-Action "Start $($svc.Display)" "This security service is stopped. Starting it will restore security monitoring.") {
            try {
                Start-Service -Name $svc.Name -ErrorAction Stop
                Write-Log "SUCCESS: Started $($svc.Display)" "Green"
                $fixCount++
            } catch {
                Write-Log "FAILED: Could not start $($svc.Display) - $($_.Exception.Message)" "Red"
            }
        } else {
            Write-Log "SKIPPED: User chose not to start $($svc.Display)" "Yellow"
            $skipCount++
        }
    } elseif ($service) {
        Write-Host "  $($svc.Display): Running - No fix needed" -ForegroundColor Green
    }
}

# ============================================================
# FIX 3: TRUST RELATIONSHIP REPAIR
# ============================================================
Write-Host ""
Write-Host "[FIX 3/12] Checking Trust Relationship..." -ForegroundColor Yellow

$secureChannel = Test-ComputerSecureChannel -ErrorAction SilentlyContinue
if (-not $secureChannel) {
    Write-Host ""
    Write-Host "  ⚠️  TRUST RELATIONSHIP IS BROKEN!" -ForegroundColor Red
    Write-Host ""
    
    if (Prompt-Action "Repair Trust Relationship" "This will reset the computer account password with the domain. This is SAFE and commonly done to fix trust issues.") {
        try {
            Test-ComputerSecureChannel -Repair -Credential (Get-Credential -Message "Enter Domain Admin credentials to repair trust")
            Write-Log "SUCCESS: Trust relationship repaired" "Green"
            $fixCount++
        } catch {
            Write-Log "FAILED: Could not repair trust - $($_.Exception.Message)" "Red"
            Write-Host "  Alternative: You may need to rejoin the domain or escalate to L3" -ForegroundColor Yellow
        }
    } else {
        Write-Log "SKIPPED: User chose not to repair trust relationship" "Yellow"
        $skipCount++
    }
} else {
    Write-Host "  Trust Relationship: Healthy - No fix needed" -ForegroundColor Green
}

# ============================================================
# FIX 4: FORCE GROUP POLICY REFRESH
# ============================================================
Write-Host ""
Write-Host "[FIX 4/12] Group Policy Refresh..." -ForegroundColor Yellow

if (Prompt-Action "Force Group Policy Refresh" "This will refresh all Group Policy settings from the domain. This is SAFE and commonly done.") {
    try {
        Write-Host "  Running gpupdate /force..." -ForegroundColor Cyan
        $gpResult = gpupdate /force 2>&1
        Write-Log "SUCCESS: Group Policy refreshed" "Green"
        $fixCount++
    } catch {
        Write-Log "FAILED: GPUpdate failed - $($_.Exception.Message)" "Red"
    }
} else {
    Write-Log "SKIPPED: User chose not to refresh Group Policy" "Yellow"
    $skipCount++
}

# ============================================================
# FIX 5: TIME SYNCHRONIZATION
# ============================================================
Write-Host ""
Write-Host "[FIX 5/12] Time Synchronization..." -ForegroundColor Yellow

$w32tm = w32tm /query /status 2>&1 | Out-String
$timeSource = if ($w32tm -match "Source:\s*(.+)") { $Matches[1].Trim() } else { "" }

if ($timeSource -match "Local CMOS" -or $timeSource -match "Free-running" -or [string]::IsNullOrEmpty($timeSource)) {
    Write-Host "  ⚠️  Time is NOT synced to domain!" -ForegroundColor Yellow
    
    if (Prompt-Action "Force Time Sync" "This will resync time with the domain controller. Critical for Kerberos authentication.") {
        try {
            w32tm /resync /force | Out-Null
            Restart-Service W32Time -Force -ErrorAction SilentlyContinue
            Write-Log "SUCCESS: Time synchronized with domain" "Green"
            $fixCount++
        } catch {
            Write-Log "FAILED: Time sync failed - $($_.Exception.Message)" "Red"
        }
    } else {
        Write-Log "SKIPPED: User chose not to sync time" "Yellow"
        $skipCount++
    }
} else {
    Write-Host "  Time Sync: OK - Synced to $timeSource" -ForegroundColor Green
}

# ============================================================
# FIX 6: DNS CACHE FLUSH
# ============================================================
Write-Host ""
Write-Host "[FIX 6/12] DNS Cache..." -ForegroundColor Yellow

if (Prompt-Action "Flush DNS Cache" "This clears the local DNS cache. SAFE and commonly done to resolve name resolution issues.") {
    try {
        Clear-DnsClientCache -ErrorAction SilentlyContinue
        ipconfig /flushdns | Out-Null
        Write-Log "SUCCESS: DNS cache flushed" "Green"
        $fixCount++
    } catch {
        Write-Log "FAILED: DNS flush failed - $($_.Exception.Message)" "Red"
    }
} else {
    Write-Log "SKIPPED: User chose not to flush DNS" "Yellow"
    $skipCount++
}

# ============================================================
# FIX 7: REGISTER DNS
# ============================================================
Write-Host ""
Write-Host "[FIX 7/12] DNS Registration..." -ForegroundColor Yellow

if (Prompt-Action "Register DNS" "This will re-register this computer's DNS records. SAFE and helps with name resolution issues.") {
    try {
        ipconfig /registerdns | Out-Null
        Write-Log "SUCCESS: DNS registration initiated" "Green"
        $fixCount++
    } catch {
        Write-Log "FAILED: DNS registration failed - $($_.Exception.Message)" "Red"
    }
} else {
    Write-Log "SKIPPED: User chose not to register DNS" "Yellow"
    $skipCount++
}

# ============================================================
# FIX 8: CLEAR KERBEROS TICKETS
# ============================================================
Write-Host ""
Write-Host "[FIX 8/12] Kerberos Tickets..." -ForegroundColor Yellow

if (Prompt-Action "Purge Kerberos Tickets" "This clears cached Kerberos tickets. SAFE and commonly done when experiencing authentication issues.") {
    try {
        klist purge | Out-Null
        Write-Log "SUCCESS: Kerberos tickets purged" "Green"
        $fixCount++
    } catch {
        Write-Log "FAILED: Klist purge failed - $($_.Exception.Message)" "Red"
    }
} else {
    Write-Log "SKIPPED: User chose not to purge Kerberos tickets" "Yellow"
    $skipCount++
}

# ============================================================
# FIX 9: CLEAR TEMP FILES (If disk low)
# ============================================================
Write-Host ""
Write-Host "[FIX 9/12] Disk Space Cleanup..." -ForegroundColor Yellow

$cDrive = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'"
$freePercent = [math]::Round(($cDrive.FreeSpace / $cDrive.Size) * 100, 1)

if ($freePercent -lt 15) {
    Write-Host "  ⚠️  C: drive is at $freePercent% free - LOW!" -ForegroundColor Yellow
    
    if (Prompt-Action "Clean Temp Files" "This will remove temporary files from Windows Temp and User Temp folders. SAFE operation.") {
        try {
            Remove-Item "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
            Remove-Item "C:\Windows\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue
            Get-ChildItem "C:\Windows\Logs" -Recurse -File -ErrorAction SilentlyContinue | 
                Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } | 
                Remove-Item -Force -ErrorAction SilentlyContinue
            
            Write-Log "SUCCESS: Temp files cleaned" "Green"
            $fixCount++
            
            $cDriveNew = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'"
            $newFreePercent = [math]::Round(($cDriveNew.FreeSpace / $cDriveNew.Size) * 100, 1)
            Write-Host "  New free space: $newFreePercent%" -ForegroundColor Cyan
        } catch {
            Write-Log "FAILED: Cleanup failed - $($_.Exception.Message)" "Red"
        }
    } else {
        Write-Log "SKIPPED: User chose not to clean temp files" "Yellow"
        $skipCount++
    }
} else {
    Write-Host "  Disk Space: $freePercent% free - OK" -ForegroundColor Green
}

# ============================================================
# FIX 10: WINDOWS UPDATE SERVICE
# ============================================================
Write-Host ""
Write-Host "[FIX 10/12] Windows Update Service..." -ForegroundColor Yellow

$wuService = Get-Service -Name wuauserv -ErrorAction SilentlyContinue
if ($wuService -and $wuService.Status -ne "Running") {
    if (Prompt-Action "Start Windows Update Service" "This service is needed for patching. Starting it is SAFE.") {
        try {
            Start-Service -Name wuauserv -ErrorAction Stop
            Write-Log "SUCCESS: Windows Update service started" "Green"
            $fixCount++
        } catch {
            Write-Log "FAILED: Could not start Windows Update - $($_.Exception.Message)" "Red"
        }
    } else {
        Write-Log "SKIPPED: User chose not to start Windows Update" "Yellow"
        $skipCount++
    }
} else {
    Write-Host "  Windows Update: Running - OK" -ForegroundColor Green
}

# ============================================================
# FIX 11: CLEAR WINDOWS UPDATE CACHE (If WU issues)
# ============================================================
Write-Host ""
Write-Host "[FIX 11/12] Windows Update Cache..." -ForegroundColor Yellow

if (Prompt-Action "Clear Windows Update Cache" "This clears the Windows Update download cache. Helps fix stuck updates. Service will be restarted.") {
    try {
        Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
        Stop-Service -Name bits -Force -ErrorAction SilentlyContinue
        
        Remove-Item "C:\Windows\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
        
        Start-Service -Name bits -ErrorAction SilentlyContinue
        Start-Service -Name wuauserv -ErrorAction SilentlyContinue
        
        Write-Log "SUCCESS: Windows Update cache cleared" "Green"
        $fixCount++
    } catch {
        Write-Log "FAILED: Could not clear WU cache - $($_.Exception.Message)" "Red"
    }
} else {
    Write-Log "SKIPPED: User chose not to clear Windows Update cache" "Yellow"
    $skipCount++
}

# ============================================================
# FIX 12: RESTART PRINT SPOOLER (Common Issue)
# ============================================================
Write-Host ""
Write-Host "[FIX 12/12] Print Spooler Service..." -ForegroundColor Yellow

$spooler = Get-Service -Name Spooler -ErrorAction SilentlyContinue
if ($spooler) {
    if (Prompt-Action "Restart Print Spooler" "This restarts the Print Spooler service. Fixes common printing issues.") {
        try {
            Restart-Service -Name Spooler -Force -ErrorAction Stop
            Write-Log "SUCCESS: Print Spooler restarted" "Green"
            $fixCount++
        } catch {
            Write-Log "FAILED: Could not restart Print Spooler - $($_.Exception.Message)" "Red"
        }
    } else {
        Write-Log "SKIPPED: User chose not to restart Print Spooler" "Yellow"
        $skipCount++
    }
}

# ============================================================
# FINAL SUMMARY
# ============================================================
Write-Host ""
Write-Host "=============================================================" -ForegroundColor Green
Write-Host "                    FIX SCRIPT COMPLETE!" -ForegroundColor Green
Write-Host "=============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "SUMMARY:" -ForegroundColor Yellow
Write-Host "  Fixes Applied:  $fixCount" -ForegroundColor Green
Write-Host "  Fixes Skipped:  $skipCount" -ForegroundColor Yellow
Write-Host ""
Write-Host "LOG FILES:" -ForegroundColor Cyan
Write-Host "  1. $logPath" -ForegroundColor White
Write-Host "  2. $logPath2" -ForegroundColor White
Write-Host ""

Write-Log "Fix Script Completed - $fixCount fixes applied, $skipCount skipped"

Write-Host "RECOMMENDATION:" -ForegroundColor Yellow
Write-Host "  Run the AUDIT script again to verify fixes were successful." -ForegroundColor White
Write-Host ""
