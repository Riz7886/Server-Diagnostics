<#
.SYNOPSIS
    L2 Server Diagnostics - Fix Script
    Author: Syed Rizvi
    Version: 4.0
.DESCRIPTION
    Fix script for L2 team with prompts before each action.
.EXAMPLE
    .\L2-ServerDiagnostics-Fix-v4.ps1
#>

param(
    [string]$OutputPath = "$env:USERPROFILE\Desktop",
    [string]$SecondaryPath = "C:\L2_Reports"
)

$ErrorActionPreference = "Stop"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$computerName = $env:COMPUTERNAME

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
    Write-Host "ACTION: $ActionName" -ForegroundColor Yellow
    Write-Host "Description: $Description"
    Write-Host ""
    $response = Read-Host "Do you want to proceed? (Y/N)"
    return $response -eq "Y" -or $response -eq "y"
}

Write-Host ""
Write-Host "L2 SERVER DIAGNOSTICS - FIX SCRIPT v4.0" -ForegroundColor Cyan
Write-Host "Author: Syed Rizvi"
Write-Host "Computer: $computerName"
Write-Host ""
Write-Host "This script will prompt before each action."
Write-Host ""

Write-Log "Fix Script Started on $computerName"

$fixCount = 0
$skipCount = 0

Write-Host "[FIX 1/12] Checking Stopped Services..." -ForegroundColor Yellow
$criticalServices = @("Netlogon", "W32Time", "gpsvc", "Dnscache", "TermService", "EventLog", "CryptSvc", "LanmanWorkstation", "LanmanServer", "BITS", "Schedule")
foreach ($svcName in $criticalServices) {
    $service = Get-Service -Name $svcName -ErrorAction SilentlyContinue
    if ($service -and $service.Status -ne "Running") {
        if (Prompt-Action "Start $svcName Service" "This service is $($service.Status). Starting it may resolve issues.") {
            try {
                Start-Service -Name $svcName -ErrorAction Stop
                Write-Log "SUCCESS: Started $svcName" "Green"
                $fixCount++
            } catch {
                Write-Log "FAILED: Could not start $svcName - $($_.Exception.Message)" "Red"
            }
        } else {
            Write-Log "SKIPPED: $svcName" "Yellow"
            $skipCount++
        }
    }
}

Write-Host ""
Write-Host "[FIX 2/12] Checking Security Tool Services..." -ForegroundColor Yellow
$securityServices = @(
    @{Name="Tenable Nessus Agent"; Display="Nessus Agent"},
    @{Name="mfefire"; Display="Trellix Firewall"},
    @{Name="macmnsvc"; Display="Trellix Agent"},
    @{Name="TmListen"; Display="Trend Micro"},
    @{Name="CSFalconService"; Display="CrowdStrike"},
    @{Name="WinDefend"; Display="Windows Defender"}
)
foreach ($svc in $securityServices) {
    $service = Get-Service -Name $svc.Name -ErrorAction SilentlyContinue
    if ($service -and $service.Status -ne "Running") {
        if (Prompt-Action "Start $($svc.Display)" "Security service is stopped. Starting it restores protection.") {
            try {
                Start-Service -Name $svc.Name -ErrorAction Stop
                Write-Log "SUCCESS: Started $($svc.Display)" "Green"
                $fixCount++
            } catch {
                Write-Log "FAILED: Could not start $($svc.Display)" "Red"
            }
        } else {
            Write-Log "SKIPPED: $($svc.Display)" "Yellow"
            $skipCount++
        }
    }
}

Write-Host ""
Write-Host "[FIX 3/12] Checking Trust Relationship..." -ForegroundColor Yellow
$secureChannel = Test-ComputerSecureChannel -ErrorAction SilentlyContinue
if (-not $secureChannel) {
    Write-Host "Trust relationship is BROKEN" -ForegroundColor Red
    if (Prompt-Action "Repair Trust Relationship" "This repairs the secure channel between this server and the domain.") {
        try {
            Test-ComputerSecureChannel -Repair -Credential (Get-Credential -Message "Enter Domain Admin credentials")
            Write-Log "SUCCESS: Trust relationship repaired" "Green"
            $fixCount++
        } catch {
            Write-Log "FAILED: Could not repair trust - $($_.Exception.Message)" "Red"
        }
    } else {
        Write-Log "SKIPPED: Trust repair" "Yellow"
        $skipCount++
    }
} else {
    Write-Host "Trust relationship: Healthy" -ForegroundColor Green
}

Write-Host ""
Write-Host "[FIX 4/12] Group Policy Refresh..." -ForegroundColor Yellow
if (Prompt-Action "Force Group Policy Refresh" "Refreshes all Group Policy settings from the domain.") {
    try {
        gpupdate /force 2>&1 | Out-Null
        Write-Log "SUCCESS: Group Policy refreshed" "Green"
        $fixCount++
    } catch {
        Write-Log "FAILED: GPUpdate failed" "Red"
    }
} else {
    Write-Log "SKIPPED: Group Policy refresh" "Yellow"
    $skipCount++
}

Write-Host ""
Write-Host "[FIX 5/12] Time Synchronization..." -ForegroundColor Yellow
$w32tm = w32tm /query /status 2>&1 | Out-String
if ($w32tm -match "Local CMOS|Free-running") {
    Write-Host "Time is NOT synced to domain" -ForegroundColor Yellow
    if (Prompt-Action "Force Time Sync" "Resyncs time with the domain controller.") {
        try {
            w32tm /resync /force | Out-Null
            Restart-Service W32Time -Force -ErrorAction SilentlyContinue
            Write-Log "SUCCESS: Time synchronized" "Green"
            $fixCount++
        } catch {
            Write-Log "FAILED: Time sync failed" "Red"
        }
    } else {
        Write-Log "SKIPPED: Time sync" "Yellow"
        $skipCount++
    }
} else {
    Write-Host "Time: Already synced" -ForegroundColor Green
}

Write-Host ""
Write-Host "[FIX 6/12] DNS Cache..." -ForegroundColor Yellow
if (Prompt-Action "Flush DNS Cache" "Clears the local DNS cache to resolve name resolution issues.") {
    try {
        Clear-DnsClientCache -ErrorAction SilentlyContinue
        ipconfig /flushdns | Out-Null
        Write-Log "SUCCESS: DNS cache flushed" "Green"
        $fixCount++
    } catch {
        Write-Log "FAILED: DNS flush failed" "Red"
    }
} else {
    Write-Log "SKIPPED: DNS flush" "Yellow"
    $skipCount++
}

Write-Host ""
Write-Host "[FIX 7/12] DNS Registration..." -ForegroundColor Yellow
if (Prompt-Action "Register DNS" "Re-registers this computer DNS records.") {
    try {
        ipconfig /registerdns | Out-Null
        Write-Log "SUCCESS: DNS registration initiated" "Green"
        $fixCount++
    } catch {
        Write-Log "FAILED: DNS registration failed" "Red"
    }
} else {
    Write-Log "SKIPPED: DNS registration" "Yellow"
    $skipCount++
}

Write-Host ""
Write-Host "[FIX 8/12] Kerberos Tickets..." -ForegroundColor Yellow
if (Prompt-Action "Clear Kerberos Tickets" "Clears cached Kerberos tickets to fix authentication issues.") {
    try {
        klist purge | Out-Null
        Write-Log "SUCCESS: Kerberos tickets cleared" "Green"
        $fixCount++
    } catch {
        Write-Log "FAILED: Could not clear tickets" "Red"
    }
} else {
    Write-Log "SKIPPED: Kerberos clear" "Yellow"
    $skipCount++
}

Write-Host ""
Write-Host "[FIX 9/12] Disk Cleanup..." -ForegroundColor Yellow
$cDrive = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'"
$freePercent = [math]::Round(($cDrive.FreeSpace / $cDrive.Size) * 100, 1)
if ($freePercent -lt 15) {
    Write-Host "C: drive is at $freePercent% free - LOW" -ForegroundColor Yellow
    if (Prompt-Action "Clean Temp Files" "Removes temporary files from Windows Temp folders.") {
        try {
            Remove-Item "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
            Remove-Item "C:\Windows\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "SUCCESS: Temp files cleaned" "Green"
            $fixCount++
        } catch {
            Write-Log "FAILED: Cleanup failed" "Red"
        }
    } else {
        Write-Log "SKIPPED: Temp cleanup" "Yellow"
        $skipCount++
    }
} else {
    Write-Host "Disk space: $freePercent% free - OK" -ForegroundColor Green
}

Write-Host ""
Write-Host "[FIX 10/12] Windows Update Service..." -ForegroundColor Yellow
$wuService = Get-Service -Name wuauserv -ErrorAction SilentlyContinue
if ($wuService -and $wuService.Status -ne "Running") {
    if (Prompt-Action "Start Windows Update Service" "Windows Update service is needed for patching.") {
        try {
            Start-Service -Name wuauserv -ErrorAction Stop
            Write-Log "SUCCESS: Windows Update started" "Green"
            $fixCount++
        } catch {
            Write-Log "FAILED: Could not start Windows Update" "Red"
        }
    } else {
        Write-Log "SKIPPED: Windows Update" "Yellow"
        $skipCount++
    }
} else {
    Write-Host "Windows Update: Running" -ForegroundColor Green
}

Write-Host ""
Write-Host "[FIX 11/12] Windows Update Cache..." -ForegroundColor Yellow
if (Prompt-Action "Clear Windows Update Cache" "Clears the Windows Update download cache to fix stuck updates.") {
    try {
        Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
        Stop-Service -Name bits -Force -ErrorAction SilentlyContinue
        Remove-Item "C:\Windows\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
        Start-Service -Name bits -ErrorAction SilentlyContinue
        Start-Service -Name wuauserv -ErrorAction SilentlyContinue
        Write-Log "SUCCESS: Windows Update cache cleared" "Green"
        $fixCount++
    } catch {
        Write-Log "FAILED: Could not clear WU cache" "Red"
    }
} else {
    Write-Log "SKIPPED: WU cache clear" "Yellow"
    $skipCount++
}

Write-Host ""
Write-Host "[FIX 12/12] Print Spooler..." -ForegroundColor Yellow
$spooler = Get-Service -Name Spooler -ErrorAction SilentlyContinue
if ($spooler) {
    if (Prompt-Action "Restart Print Spooler" "Restarts the Print Spooler service to fix printing issues.") {
        try {
            Restart-Service -Name Spooler -Force -ErrorAction Stop
            Write-Log "SUCCESS: Print Spooler restarted" "Green"
            $fixCount++
        } catch {
            Write-Log "FAILED: Could not restart Print Spooler" "Red"
        }
    } else {
        Write-Log "SKIPPED: Print Spooler" "Yellow"
        $skipCount++
    }
}

Write-Host ""
Write-Host "FIX SCRIPT COMPLETE" -ForegroundColor Green
Write-Host ""
Write-Host "Summary:"
Write-Host "  Fixes Applied: $fixCount" -ForegroundColor Green
Write-Host "  Fixes Skipped: $skipCount" -ForegroundColor Yellow
Write-Host ""
Write-Host "Log Files:"
Write-Host "  $logPath"
Write-Host "  $logPath2"
Write-Host ""

Write-Log "Fix Script Completed - $fixCount applied, $skipCount skipped"
