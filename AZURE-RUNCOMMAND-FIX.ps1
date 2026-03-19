# ============================================================
#  AZURE VM RUN COMMAND - RDP EMERGENCY FIX
#  Target: c006app01otpd3ab462fa6fb707d  (10.168.0.32)
#  OS: Windows Server 2019 Datacenter
#  Paste this directly into Azure Portal > VM > Run Command
# ============================================================

Write-Host "=== STEP 1: Checking current Winlogon registry values ===" -ForegroundColor Cyan
$wlPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'
$wl = Get-ItemProperty -Path $wlPath
Write-Host "  Current Shell    : $($wl.Shell)"
Write-Host "  Current Userinit : $($wl.Userinit)"

# ---- Fix Winlogon registry (root cause of RDP black screen / immediate logoff) ----
Write-Host ""
Write-Host "=== STEP 2: Restoring Winlogon Shell and Userinit ===" -ForegroundColor Cyan
Set-ItemProperty -Path $wlPath -Name 'Shell'    -Value 'explorer.exe'
Set-ItemProperty -Path $wlPath -Name 'Userinit' -Value 'C:\Windows\system32\userinit.exe,'
Write-Host "  Shell set to    : explorer.exe"         -ForegroundColor Green
Write-Host "  Userinit set to : userinit.exe,"        -ForegroundColor Green

# ---- Also confirm RDP is enabled (fDenyTSConnections must be 0) ----
Write-Host ""
Write-Host "=== STEP 3: Verifying RDP is enabled ===" -ForegroundColor Cyan
$rdpPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server'
$rdp = Get-ItemProperty -Path $rdpPath
if ($rdp.fDenyTSConnections -ne 0) {
    Set-ItemProperty -Path $rdpPath -Name 'fDenyTSConnections' -Value 0
    Write-Host "  RDP was DISABLED - now ENABLED" -ForegroundColor Yellow
} else {
    Write-Host "  RDP already enabled (fDenyTSConnections = 0)" -ForegroundColor Green
}

# ---- Check NLA setting ----
$nlaPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp'
$nla = Get-ItemProperty -Path $nlaPath -ErrorAction SilentlyContinue
Write-Host "  NLA (UserAuthentication): $($nla.UserAuthentication)"

# ---- Restart Remote Desktop Services ----
Write-Host ""
Write-Host "=== STEP 4: Restarting Remote Desktop Services ===" -ForegroundColor Cyan
try {
    Stop-Service -Name TermService -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 4
    Start-Service -Name TermService -ErrorAction Stop
    Write-Host "  TermService restarted successfully" -ForegroundColor Green
} catch {
    Write-Host "  TermService restart warning: $($_.Message)" -ForegroundColor Yellow
}

# ---- Check for bad KB on this server ----
Write-Host ""
Write-Host "=== STEP 5: Checking for bad January 2026 patches ===" -ForegroundColor Cyan
$badKBs = @('KB5074109','KB5073457','KB5073450','KB5073455')
$installed = Get-HotFix | Where-Object { $badKBs -contains $_.HotFixID }
if ($installed) {
    Write-Host "  BAD PATCHES FOUND - scheduling removal:" -ForegroundColor Red
    foreach ($kb in $installed) {
        Write-Host "    $($kb.HotFixID) - installed $($kb.InstalledOn)" -ForegroundColor Red
        $kbNum = $kb.HotFixID -replace '[Kk][Bb]',''
        $proc = Start-Process -FilePath 'wusa.exe' `
            -ArgumentList "/uninstall /kb:$kbNum /quiet /norestart" `
            -Wait -PassThru
        Write-Host "    Uninstall exit code: $($proc.ExitCode)" -ForegroundColor Yellow
    }
} else {
    Write-Host "  No bad January 2026 patches found on this server." -ForegroundColor Green
}

# ---- Check if March 2026 fix is already installed ----
Write-Host ""
Write-Host "=== STEP 6: Checking for March 2026 permanent fix (KB5078766) ===" -ForegroundColor Cyan
$fixKB = Get-HotFix -Id 'KB5078766' -ErrorAction SilentlyContinue
if ($fixKB) {
    Write-Host "  KB5078766 is already installed. Server is fully patched." -ForegroundColor Green
} else {
    Write-Host "  KB5078766 NOT installed - attempting install via Windows Update..." -ForegroundColor Yellow
    try {
        $sess   = New-Object -ComObject Microsoft.Update.Session
        $srch   = $sess.CreateUpdateSearcher()
        $res    = $srch.Search('IsInstalled=0')
        $target = $null
        foreach ($u in $res.Updates) {
            if ($u.KBArticleIDs -contains '5078766') { $target = $u; break }
        }
        if ($target) {
            $coll = New-Object -ComObject Microsoft.Update.UpdateColl
            $coll.Add($target) | Out-Null
            $dl         = $sess.CreateUpdateDownloader()
            $dl.Updates = $coll
            $dl.Download() | Out-Null
            $ins         = $sess.CreateUpdateInstaller()
            $ins.Updates = $coll
            $r           = $ins.Install()
            Write-Host "  KB5078766 installed. Result: $($r.ResultCode). Reboot required: $($r.RebootRequired)" -ForegroundColor Green
        } else {
            Write-Host "  KB5078766 not found in Windows Update feed (no internet or WSUS not synced)." -ForegroundColor Yellow
            Write-Host "  Download manually from:" -ForegroundColor Yellow
            Write-Host "  https://catalog.update.microsoft.com/Search.aspx?q=KB5078766" -ForegroundColor Cyan
            Write-Host "  Copy the .msu to C:\Windows\Temp\ then run:" -ForegroundColor Yellow
            Write-Host "  wusa.exe C:\Windows\Temp\KB5078766.msu /quiet /norestart" -ForegroundColor Cyan
        }
    } catch {
        Write-Host "  Windows Update COM error: $($_.Message)" -ForegroundColor Yellow
    }
}

# ---- Final status ----
Write-Host ""
Write-Host "=== FINAL STATUS ===" -ForegroundColor Cyan
$wlFinal = Get-ItemProperty -Path $wlPath
$rdpFinal = Get-ItemProperty -Path $rdpPath
Write-Host "  Winlogon Shell    : $($wlFinal.Shell)"
Write-Host "  Winlogon Userinit : $($wlFinal.Userinit)"
Write-Host "  RDP Enabled       : $(if ($rdpFinal.fDenyTSConnections -eq 0) {'YES'} else {'NO - PROBLEM'})"
Write-Host ""
Write-Host "=== EMERGENCY FIX COMPLETE ===" -ForegroundColor Green
Write-Host "  Test RDP connection to 10.168.0.32 now." -ForegroundColor Green
Write-Host "  If it works, schedule a reboot in your SAP maintenance window." -ForegroundColor Yellow
Write-Host "  The reboot is required to fully apply the patch removal." -ForegroundColor Yellow
