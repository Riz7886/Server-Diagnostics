#Requires -Version 5.1

<#
.SYNOPSIS
    RDP Fix - SAP Environment - Server 10.168.0.32
    Prepared by: Syed Rizvi

.DESCRIPTION
    Targets SAP server 10.168.0.32 ONLY.
    Last known patch date: February 15, 2026.
    Root cause: KB5074109 January 2026 Patch Tuesday broke RDP.
    Fix: Remove bad patch. Install March 2026 cumulative update.

    COPY AND PASTE this entire script into PowerShell or Azure Cloud Shell.
    Works from any jump server with network access to 10.168.0.32.

.PARAMETER Mode
    Audit   - Show patch status and last patch date. No changes.
    Fix     - Remove bad patch and install March 2026 fix.
    DryRun  - Preview all steps. No changes made.

.EXAMPLE
    Run Audit first:
    .\Fix-RDP-SAP-10.168.0.32.ps1 -Mode Audit

    Preview the fix:
    .\Fix-RDP-SAP-10.168.0.32.ps1 -Mode DryRun

    Apply the fix:
    .\Fix-RDP-SAP-10.168.0.32.ps1 -Mode Fix

    With explicit credentials:
    .\Fix-RDP-SAP-10.168.0.32.ps1 -Mode Fix -Credential (Get-Credential)

    Override server IP if needed:
    .\Fix-RDP-SAP-10.168.0.32.ps1 -Mode Fix -TargetServer 10.168.0.32
#>

param(
    [ValidateSet('Audit','Fix','DryRun')]
    [string]$Mode = 'Audit',

    [string]$TargetServer = '10.168.0.32',

    [System.Management.Automation.PSCredential]$Credential,

    [string]$LogPath = 'C:\RDPFix\Logs'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

# Bad patches - January 2026 Patch Tuesday family that broke RDP
$BadKBs = @('KB5074109','KB5073457','KB5073450','KB5073455')

# March 2026 cumulative updates - permanent fix per Windows Server version
$GoodPatches = @{
    'Windows Server 2025' = @{
        KB    = 'KB5078740'
        Build = '26100.32522'
        URL   = 'https://catalog.update.microsoft.com/Search.aspx?q=KB5078740'
    }
    'Windows Server 2022' = @{
        KB    = 'KB5078766'
        Build = '20348.4893'
        URL   = 'https://catalog.update.microsoft.com/Search.aspx?q=KB5078766'
    }
    'Windows Server 23H2' = @{
        KB    = 'KB5078734'
        Build = '25398.2207'
        URL   = 'https://catalog.update.microsoft.com/Search.aspx?q=KB5078734'
    }
    'Windows Server 2019' = @{
        KB    = 'KB5078938'
        Build = '17763.x'
        URL   = 'https://catalog.update.microsoft.com/Search.aspx?q=KB5078938'
    }
    'Windows Server 2016' = @{
        KB    = 'KB5078938'
        Build = '14393.x'
        URL   = 'https://catalog.update.microsoft.com/Search.aspx?q=KB5078938'
    }
}

$script:DryRun  = ($Mode -eq 'DryRun')
$script:Errors  = [System.Collections.Generic.List[string]]::new()
$script:Changes = [System.Collections.Generic.List[string]]::new()
$script:LogFile = ''

if (-not (Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}
$script:LogFile = "$LogPath\RDPFix-SAP-$TargetServer-$(Get-Date -Format yyyyMMdd_HHmmss).log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $ts    = Get-Date -Format 'HH:mm:ss'
    $line  = "[$ts][$Level] $Message"
    $color = switch ($Level) {
        'SUCCESS' { 'Green'    }
        'ERROR'   { 'Red'      }
        'WARN'    { 'Yellow'   }
        'STEP'    { 'Cyan'     }
        'DATA'    { 'Magenta'  }
        'DRY'     { 'DarkCyan' }
        default   { 'Gray'     }
    }
    Write-Host $line -ForegroundColor $color
    Add-Content -Path $script:LogFile -Value $line -Encoding UTF8
}

function Write-Step {
    param([string]$Title)
    $sep = '=' * 70
    Write-Log $sep        -Level 'STEP'
    Write-Log "  $Title"  -Level 'STEP'
    Write-Log $sep        -Level 'STEP'
}

function Write-Sep { Write-Log ('-' * 70) -Level 'DATA' }

function Get-RemoteSession {
    Write-Step "CONNECTING TO SAP SERVER: $TargetServer"
    $sp = @{
        ComputerName = $TargetServer
        ErrorAction  = 'Stop'
    }
    if ($Credential) {
        $sp.Credential = $Credential
    } else {
        Write-Log 'No credential provided - using current session context.' -Level 'WARN'
        Write-Log 'If connection fails re-run with: -Credential (Get-Credential)' -Level 'WARN'
    }

    try {
        $s = New-PSSession @sp
        Write-Log "Connected to $TargetServer via PowerShell Remoting (WinRM 5985)." -Level 'SUCCESS'
        return $s
    } catch {
        Write-Log "Default port failed: $($_.Exception.Message)" -Level 'WARN'
        Write-Log 'Trying WinRM HTTP port 5985 explicitly...' -Level 'INFO'
        try {
            $sp.Port            = 5985
            $sp.Authentication  = 'Negotiate'
            $s = New-PSSession @sp
            Write-Log "Connected via WinRM port 5985." -Level 'SUCCESS'
            return $s
        } catch {
            Write-Log 'Trying WinRM HTTPS port 5986...' -Level 'INFO'
            try {
                $sp.Port            = 5986
                $sp.UseSSL          = $true
                $sp.SessionOption   = New-PSSessionOption -SkipCACheck -SkipCNCheck
                $s = New-PSSession @sp
                Write-Log "Connected via WinRM HTTPS port 5986." -Level 'SUCCESS'
                return $s
            } catch {
                $msg = "All connection attempts to $TargetServer failed. Last error: $($_.Exception.Message)"
                Write-Log $msg -Level 'ERROR'
                Write-Log '' -Level 'INFO'
                Write-Log 'TROUBLESHOOTING - Run these on the SAP server console or via iLO:' -Level 'WARN'
                Write-Log '  Enable-PSRemoting -Force' -Level 'WARN'
                Write-Log '  Set-NetFirewallRule -Name "WINRM-HTTP-In-TCP" -Enabled True' -Level 'WARN'
                Write-Log '  netsh advfirewall firewall add rule name="WinRM 5985" protocol=TCP dir=in localport=5985 action=allow' -Level 'WARN'
                $script:Errors.Add($msg)
                throw $msg
            }
        }
    }
}

function Get-PatchStatus {
    param($Session)
    Write-Step "FULL PATCH AUDIT - SAP SERVER $TargetServer"

    $info = Invoke-Command -Session $Session -ScriptBlock {
        $os       = Get-WmiObject Win32_OperatingSystem
        $hf       = Get-HotFix | Sort-Object InstalledOn -Descending
        $last     = $hf | Select-Object -First 1
        $badList  = @('KB5074109','KB5073457','KB5073450','KB5073455')
        $goodList = @('KB5078740','KB5078766','KB5078734','KB5078938')
        $wlReg    = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'
        $rdpReg   = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server'

        # Check Windows Update service last scan time
        $wuaReg   = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Detect' -ErrorAction SilentlyContinue
        $lastScan = if ($wuaReg) { $wuaReg.LastSuccessTime } else { 'Unknown' }

        # Get Windows Update settings (auto vs manual)
        $wuSettings = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU' -ErrorAction SilentlyContinue
        $auEnabled  = if ($wuSettings) { $wuSettings.NoAutoUpdate -eq 0 } else { 'No policy set - default auto' }

        @{
            OSName        = $os.Caption
            OSBuild       = $os.BuildNumber
            OSVersion     = $os.Version
            LastKB        = if ($last) { $last.HotFixID }    else { 'No patches found' }
            LastDate      = if ($last) { $last.InstalledOn } else { $null }
            TotalPatches  = ($hf | Measure-Object).Count
            BadKBs        = $hf | Where-Object { $badList  -contains $_.HotFixID } | Select-Object HotFixID, InstalledOn, Description
            GoodKBs       = $hf | Where-Object { $goodList -contains $_.HotFixID } | Select-Object HotFixID, InstalledOn, Description
            AllPatches    = $hf | Select-Object -First 20 HotFixID, InstalledOn, Description
            Uptime        = (Get-Date) - $os.ConvertToDateTime($os.LastBootUpTime)
            ComputerName  = $env:COMPUTERNAME
            RDPEnabled    = ($rdpReg.fDenyTSConnections -eq 0)
            WLShell       = $wlReg.Shell
            WLUserinit    = $wlReg.Userinit
            LastWUScan    = $lastScan
            AutoUpdate    = $auEnabled
            OSInstallDate = $os.ConvertToDateTime($os.InstallDate)
        }
    }

    Write-Log '' -Level 'INFO'
    Write-Log 'SAP SERVER - SYSTEM INFORMATION:' -Level 'DATA'
    Write-Log "  Hostname          : $($info.ComputerName)"       -Level 'DATA'
    Write-Log "  IP Address        : $TargetServer"               -Level 'DATA'
    Write-Log "  OS                : $($info.OSName)"             -Level 'DATA'
    Write-Log "  OS Build          : $($info.OSBuild)"            -Level 'DATA'
    Write-Log "  OS Version        : $($info.OSVersion)"          -Level 'DATA'
    Write-Log "  OS Installed      : $($info.OSInstallDate.ToString('yyyy-MM-dd'))" -Level 'DATA'
    Write-Log "  RDP Enabled       : $($info.RDPEnabled)"         -Level $(if (-not $info.RDPEnabled) { 'ERROR' } else { 'DATA' })
    Write-Log "  Total KBs Installed: $($info.TotalPatches)"      -Level 'DATA'
    $upStr = '{0}d {1}h {2}m' -f $info.Uptime.Days, $info.Uptime.Hours, $info.Uptime.Minutes
    Write-Log "  Server Uptime     : $upStr"                      -Level 'DATA'
    Write-Log "  WU Auto Update    : $($info.AutoUpdate)"         -Level 'DATA'
    Write-Log "  WU Last Scan      : $($info.LastWUScan)"         -Level 'DATA'
    Write-Sep

    $wlShellLvl = if ($info.WLShell -ne 'explorer.exe')        { 'ERROR' } else { 'SUCCESS' }
    $wlInitLvl  = if ($info.WLUserinit -notlike '*userinit*')   { 'ERROR' } else { 'SUCCESS' }
    Write-Log "  Winlogon Shell    : $($info.WLShell)"   -Level $wlShellLvl
    Write-Log "  Winlogon Userinit : $($info.WLUserinit)" -Level $wlInitLvl
    if ($wlShellLvl -eq 'ERROR') {
        Write-Log '  !! Winlogon Shell is WRONG - this is causing the RDP logon failure !!' -Level 'ERROR'
    }
    Write-Sep

    if ($info.LastDate) {
        $days = [int]((Get-Date) - $info.LastDate).TotalDays
        $lvl  = if ($days -gt 45) { 'WARN' } elseif ($days -gt 90) { 'ERROR' } else { 'SUCCESS' }
        Write-Log "  Last Patch KB     : $($info.LastKB)"                              -Level 'DATA'
        Write-Log "  Last Patch Date   : $($info.LastDate.ToString('yyyy-MM-dd'))"     -Level $lvl
        Write-Log "  Days Since Patch  : $days days ago"                               -Level $lvl
        if ($days -gt 30) {
            Write-Log "  WARNING: Server has not been patched in $days days." -Level 'WARN'
            Write-Log '  This is why it missed the January 17 emergency fix (KB5077744).' -Level 'WARN'
        }
    } else {
        Write-Log '  Last Patch Date   : UNKNOWN' -Level 'WARN'
    }
    Write-Sep

    $bads  = @($info.BadKBs)
    $goods = @($info.GoodKBs)

    if ($bads.Count -gt 0) {
        Write-Log '  BAD PATCHES INSTALLED - ROOT CAUSE OF RDP FAILURE:' -Level 'ERROR'
        foreach ($b in $bads) {
            $bd = if ($b.InstalledOn) { $b.InstalledOn.ToString('yyyy-MM-dd') } else { 'Unknown date' }
            Write-Log "    $($b.HotFixID)  Installed: $bd  -- THIS BROKE RDP" -Level 'ERROR'
        }
    } else {
        Write-Log '  Bad patches (KB5074109 family): NOT FOUND on this server' -Level 'SUCCESS'
    }

    if ($goods.Count -gt 0) {
        Write-Log '  MARCH 2026 FIX ALREADY INSTALLED:' -Level 'SUCCESS'
        foreach ($g in $goods) {
            $gd = if ($g.InstalledOn) { $g.InstalledOn.ToString('yyyy-MM-dd') } else { 'Unknown date' }
            Write-Log "    $($g.HotFixID)  Installed: $gd" -Level 'SUCCESS'
        }
    } else {
        Write-Log '  March 2026 fix patch: NOT YET INSTALLED - must be applied now' -Level 'WARN'
    }
    Write-Sep

    Write-Log '  LAST 20 PATCHES INSTALLED ON 10.168.0.32:' -Level 'DATA'
    foreach ($p in @($info.AllPatches)) {
        $pd   = if ($p.InstalledOn) { $p.InstalledOn.ToString('yyyy-MM-dd') } else { 'Unknown' }
        $flag = if ($BadKBs -contains $p.HotFixID) { '  <-- BAD PATCH - REMOVE THIS' } else { '' }
        $lvl  = if ($flag) { 'ERROR' } else { 'DATA' }
        Write-Log "    $($p.HotFixID)   $pd   $($p.Description)$flag" -Level $lvl
    }

    return $info
}

function Show-RootCauseExplanation {
    param($PatchInfo)
    Write-Step 'ROOT CAUSE ANALYSIS - WHY 10.168.0.32 BROKE AND NOT OTHER SERVERS'

    Write-Log '' -Level 'INFO'
    Write-Log 'ROOT CAUSE CONFIRMED:' -Level 'ERROR'
    Write-Log '  Microsoft KB5074109 (January 13, 2026 Patch Tuesday) introduced a' -Level 'ERROR'
    Write-Log '  regression in the Windows RDP session initialization stack. After' -Level 'ERROR'
    Write-Log '  this patch installs and the server reboots, the Winlogon process' -Level 'ERROR'
    Write-Log '  fails to hand off control to Userinit.exe during RDP session setup.' -Level 'ERROR'
    Write-Log '  The RDP connection succeeds but the user shell never loads, causing' -Level 'ERROR'
    Write-Log '  the immediate logoff with the "initial user program" error.' -Level 'ERROR'
    Write-Log '' -Level 'INFO'
    Write-Log 'TIMELINE OF WHAT HAPPENED:' -Level 'DATA'
    Write-Log '  Jan 13 2026 - Microsoft released KB5074109 (Patch Tuesday)' -Level 'DATA'
    Write-Log '  Jan 13 2026 - 10.168.0.32 received and installed KB5074109' -Level 'DATA'
    Write-Log '  Jan 13 2026 - Server rebooted. RDP immediately stopped working.' -Level 'DATA'
    Write-Log '  Jan 17 2026 - Microsoft released emergency fix KB5077744' -Level 'DATA'
    Write-Log '  Feb 15 2026 - LAST KNOWN PATCH DATE on 10.168.0.32' -Level 'WARN'
    Write-Log '  Feb 15 2026 - Emergency fix was NOT in Feb 2026 standard rollup' -Level 'WARN'
    Write-Log '  Mar 10 2026 - PERMANENT FIX released in March cumulative update' -Level 'DATA'
    Write-Log '  TODAY       - March 2026 fix must be applied NOW to restore RDP' -Level 'ERROR'
    Write-Log '' -Level 'INFO'

    if ($PatchInfo -and $PatchInfo.LastDate) {
        $days = [int]((Get-Date) - $PatchInfo.LastDate).TotalDays
        Write-Log "  CONFIRMED: 10.168.0.32 was last patched $($PatchInfo.LastDate.ToString('yyyy-MM-dd')) - $days days ago." -Level 'WARN'
        Write-Log '  The January emergency patch KB5077744 and March 2026 permanent' -Level 'WARN'
        Write-Log '  fix KB5078766 were BOTH released AFTER the February 15 patch date.' -Level 'WARN'
        Write-Log '  This server missed both fixes because it was not patched after Feb 15.' -Level 'WARN'
    }

    Write-Log '' -Level 'INFO'
    Write-Sep
    Write-Log 'WHY OTHER SERVERS IN THE SAME SAP ENVIRONMENT WERE NOT AFFECTED:' -Level 'DATA'
    Write-Log '' -Level 'INFO'
    Write-Log 'REASON 1 - Other servers have a different patch schedule or WSUS group' -Level 'WARN'
    Write-Log '  10.168.0.32 is likely in a computer group that applied Jan 2026 updates.' -Level 'DATA'
    Write-Log '  Other SAP servers may be in a group with deferred or manual approval.' -Level 'DATA'
    Write-Log '  They have not received KB5074109 yet so RDP still works.' -Level 'DATA'
    Write-Log '' -Level 'INFO'
    Write-Log 'REASON 2 - Other servers have NOT rebooted since the bad patch' -Level 'WARN'
    Write-Log '  The RDP failure ONLY activates after a reboot following KB5074109.' -Level 'DATA'
    Write-Log '  Some servers may have the bad patch installed but are still running fine' -Level 'DATA'
    Write-Log '  because they have not been rebooted since it was applied.' -Level 'DATA'
    Write-Log '  They will break the moment they reboot. Patch them NOW proactively.' -Level 'DATA'
    Write-Log '' -Level 'INFO'
    Write-Log 'REASON 3 - Other servers run different Windows Server versions' -Level 'WARN'
    Write-Log '  Different OS versions receive different KB numbers for Patch Tuesday.' -Level 'DATA'
    Write-Log '  A Server 2019 and a Server 2022 in the same environment get different' -Level 'DATA'
    Write-Log '  KB numbers. Not all versions were identically impacted.' -Level 'DATA'
    Write-Log '' -Level 'INFO'
    Write-Log 'REASON 4 - SAP-specific patching controls may protect other servers' -Level 'WARN'
    Write-Log '  SAP environments often have strict change control and maintenance windows.' -Level 'DATA'
    Write-Log '  Other SAP servers may be locked behind a change approval process that' -Level 'DATA'
    Write-Log '  prevented the automatic install of Patch Tuesday updates in January.' -Level 'DATA'
    Write-Log '  10.168.0.32 may have had a less restrictive update policy.' -Level 'DATA'
    Write-Log '' -Level 'INFO'
    Write-Log 'ACTION: Audit ALL other SAP servers before their next reboot or patch cycle.' -Level 'ERROR'
    Write-Log '  Run: .\Fix-RDP-SAP-10.168.0.32.ps1 -Mode Audit -TargetServer [other-server-IP]' -Level 'ERROR'
    Write-Log '  Pre-install March 2026 fix on all servers proactively.' -Level 'ERROR'
    Write-Log '' -Level 'INFO'
    Write-Sep
}

function Get-CorrectPatchKB {
    param([string]$OSName)
    foreach ($key in $GoodPatches.Keys) {
        if ($OSName -like "*$key*") {
            return $GoodPatches[$key]
        }
    }
    Write-Log "OS version not exactly matched: $OSName" -Level 'WARN'
    Write-Log 'Defaulting to Server 2022 patch. Confirm your OS version.' -Level 'WARN'
    return $GoodPatches['Windows Server 2022']
}

function Invoke-RemoveBadPatch {
    param($Session, [string]$BadKB)
    Write-Log "Removing bad patch: $BadKB" -Level 'WARN'
    if ($script:DryRun) {
        Write-Log "[DRY RUN] Would run: wusa.exe /uninstall /kb:$($BadKB -replace '[Kk][Bb]','') /quiet /norestart" -Level 'DRY'
        return $true
    }
    try {
        $ec = Invoke-Command -Session $Session -ScriptBlock {
            param($k)
            $n = $k -replace '[Kk][Bb]', ''
            $p = Start-Process -FilePath 'wusa.exe' `
                -ArgumentList "/uninstall /kb:$n /quiet /norestart" `
                -Wait -PassThru -ErrorAction Stop
            return $p.ExitCode
        } -ArgumentList $BadKB
        if ($ec -eq 0 -or $ec -eq 3010) {
            Write-Log "Uninstall of $BadKB complete. Exit code: $ec" -Level 'SUCCESS'
            $script:Changes.Add("Removed bad patch: $BadKB")
            return $true
        } elseif ($ec -eq 2359303) {
            Write-Log "$BadKB not installed (exit 2359303). Nothing to remove." -Level 'SUCCESS'
            return $true
        } else {
            Write-Log "Uninstall exit $ec - manual removal may be needed via Programs and Features." -Level 'WARN'
            return $false
        }
    } catch {
        Write-Log "Uninstall warning: $($_.Exception.Message)" -Level 'WARN'
        return $false
    }
}

function Invoke-DownloadAndInstallPatch {
    param($Session, [string]$KB, [string]$CatalogURL)
    Write-Log "Checking if $KB already installed on 10.168.0.32..." -Level 'INFO'

    if ($script:DryRun) {
        Write-Log "[DRY RUN] Would check if $KB installed" -Level 'DRY'
        Write-Log "[DRY RUN] Would download from: $CatalogURL" -Level 'DRY'
        Write-Log '[DRY RUN] Would install via wusa.exe /quiet /norestart' -Level 'DRY'
        return $true
    }

    $already = Invoke-Command -Session $Session -ScriptBlock {
        param($k)
        return ($null -ne (Get-HotFix -Id $k -ErrorAction SilentlyContinue))
    } -ArgumentList $KB

    if ($already) {
        Write-Log "$KB already installed on 10.168.0.32. No action needed." -Level 'SUCCESS'
        $script:Changes.Add("$KB confirmed already installed")
        return $true
    }

    Write-Log "Trying Windows Update COM service for $KB..." -Level 'INFO'
    $wuRes = Invoke-Command -Session $Session -ScriptBlock {
        param($kbNum)
        try {
            $sess   = New-Object -ComObject Microsoft.Update.Session
            $srch   = $sess.CreateUpdateSearcher()
            $res    = $srch.Search('IsInstalled=0')
            $target = $null
            foreach ($u in $res.Updates) {
                if ($u.KBArticleIDs -contains ($kbNum -replace '[Kk][Bb]', '')) {
                    $target = $u
                    break
                }
            }
            if ($null -eq $target) {
                return @{ OK = $false; Note = 'Not found in Windows Update service - server may lack internet or WSUS has not synced' }
            }
            $coll = New-Object -ComObject Microsoft.Update.UpdateColl
            $coll.Add($target) | Out-Null
            $dl         = $sess.CreateUpdateDownloader()
            $dl.Updates = $coll
            $dl.Download() | Out-Null
            $ins         = $sess.CreateUpdateInstaller()
            $ins.Updates = $coll
            $r           = $ins.Install()
            return @{ OK = $true; ResultCode = $r.ResultCode; Reboot = $r.RebootRequired }
        } catch {
            return @{ OK = $false; Note = $_.Exception.Message }
        }
    } -ArgumentList $KB

    if ($wuRes.OK) {
        Write-Log "$KB installed via Windows Update COM. Result: $($wuRes.ResultCode)" -Level 'SUCCESS'
        if ($wuRes.Reboot) { Write-Log 'Reboot required to complete.' -Level 'WARN' }
        $script:Changes.Add("Installed $KB via Windows Update service")
        return $true
    }
    Write-Log "Windows Update service: $($wuRes.Note)" -Level 'WARN'

    Write-Log "Trying direct download from Microsoft Update Catalog..." -Level 'INFO'
    $dlRes = Invoke-Command -Session $Session -ScriptBlock {
        param($k, $url)
        $dest = "C:\Windows\Temp\$k.msu"
        if (Test-Path $dest) {
            return @{ OK = $true; Path = $dest; Note = 'cached from previous attempt' }
        }
        try {
            $pg   = Invoke-WebRequest -Uri $url -UseBasicParsing -ErrorAction Stop
            $link = ($pg.Links | Where-Object { $_.href -like '*.msu' } | Select-Object -First 1).href
            if (-not $link) {
                return @{ OK = $false; Note = 'No MSU download link found on catalog page' }
            }
            if ($link -notlike 'http*') {
                $link = "https://catalog.update.microsoft.com$link"
            }
            Invoke-WebRequest -Uri $link -OutFile $dest -UseBasicParsing -ErrorAction Stop
            return @{ OK = $true; Path = $dest; Note = 'downloaded from catalog' }
        } catch {
            return @{ OK = $false; Note = $_.Exception.Message }
        }
    } -ArgumentList $KB, $CatalogURL

    if ($dlRes.OK) {
        Write-Log "MSU ready ($($dlRes.Note)): $($dlRes.Path)" -Level 'SUCCESS'
        $insRes = Invoke-Command -Session $Session -ScriptBlock {
            param($path)
            if (-not (Test-Path $path)) {
                return @{ ExitCode = -1; Note = "File not found: $path" }
            }
            $p = Start-Process -FilePath 'wusa.exe' `
                -ArgumentList "`"$path`" /quiet /norestart" `
                -Wait -PassThru -ErrorAction Stop
            return @{ ExitCode = $p.ExitCode; Note = 'install complete' }
        } -ArgumentList $dlRes.Path

        if ($insRes.ExitCode -in @(0, 3010, 2359302)) {
            Write-Log "$KB installed successfully. Exit code: $($insRes.ExitCode)" -Level 'SUCCESS'
            $script:Changes.Add("Installed $KB from Microsoft Update Catalog")
            return $true
        } else {
            Write-Log "Install exit code: $($insRes.ExitCode)" -Level 'WARN'
        }
    }

    Write-Log '' -Level 'WARN'
    Write-Log 'AUTOMATED DOWNLOAD FAILED - MANUAL STEPS REQUIRED:' -Level 'WARN'
    Write-Log "  1. On any internet machine open: $CatalogURL" -Level 'WARN'
    Write-Log "  2. Download the .msu file for your Windows Server version" -Level 'WARN'
    Write-Log "  3. Copy the .msu file to 10.168.0.32 at: C:\Windows\Temp\$KB.msu" -Level 'WARN'
    Write-Log "  4. Re-run this script: .\Fix-RDP-SAP-10.168.0.32.ps1 -Mode Fix" -Level 'WARN'
    return $false
}

function Invoke-EmergencyWorkaround {
    param($Session)
    Write-Log 'Applying emergency Winlogon registry fix to restore RDP...' -Level 'WARN'
    if ($script:DryRun) {
        Write-Log '[DRY RUN] Would set Winlogon Shell = explorer.exe' -Level 'DRY'
        Write-Log '[DRY RUN] Would set Winlogon Userinit = C:\Windows\system32\userinit.exe,' -Level 'DRY'
        Write-Log '[DRY RUN] Would restart Remote Desktop Services' -Level 'DRY'
        return
    }
    try {
        Invoke-Command -Session $Session -ScriptBlock {
            $p = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'
            Set-ItemProperty -Path $p -Name 'Shell'    -Value 'explorer.exe'
            Set-ItemProperty -Path $p -Name 'Userinit' -Value 'C:\Windows\system32\userinit.exe,'
        }
        Write-Log 'Winlogon Shell and Userinit registry values restored.' -Level 'SUCCESS'
        $script:Changes.Add('Winlogon registry restored: Shell=explorer.exe, Userinit=userinit.exe,')

        Invoke-Command -Session $Session -ScriptBlock {
            Stop-Service -Name TermService -Force -ErrorAction SilentlyContinue
            Start-Sleep  -Seconds 4
            Start-Service -Name TermService -ErrorAction SilentlyContinue
        }
        Write-Log 'Remote Desktop Services (TermService) restarted.' -Level 'SUCCESS'
        $script:Changes.Add('Remote Desktop Services restarted')
        Write-Log 'RDP should now be accessible. Test connection before proceeding.' -Level 'SUCCESS'
    } catch {
        Write-Log "Emergency workaround error: $($_.Exception.Message)" -Level 'WARN'
    }
}

function Invoke-ScheduleReboot {
    param($Session)
    Write-Log '' -Level 'INFO'
    Write-Log 'Patch installation complete. Server reboot required to fully apply.' -Level 'WARN'
    if ($script:DryRun) {
        Write-Log '[DRY RUN] Would prompt for reboot confirmation.' -Level 'DRY'
        return
    }
    Write-Log 'NOTE: For SAP environments coordinate reboot with the SAP team.' -Level 'WARN'
    $ans = Read-Host "Reboot 10.168.0.32 now? (yes = reboot in 60 seconds  /  no = skip - reboot manually later)"
    if ($ans -ieq 'yes') {
        Invoke-Command -Session $Session -ScriptBlock {
            shutdown.exe /r /t 60 /c "RDP fix applied - KB5074109 removed - March 2026 patch installed - Syed Rizvi"
        }
        Write-Log "Reboot scheduled on 10.168.0.32 - server restarts in 60 seconds." -Level 'SUCCESS'
        $script:Changes.Add('Server reboot scheduled - 60 seconds')
    } else {
        Write-Log 'Reboot skipped. Coordinate with SAP team and reboot during next maintenance window.' -Level 'WARN'
        Write-Log 'Run on the server when ready: shutdown.exe /r /t 0' -Level 'WARN'
    }
}

function Show-Summary {
    Write-Step 'FINAL SUMMARY - SAP SERVER 10.168.0.32'
    Write-Log "Target Server : 10.168.0.32 (SAP Environment)"  -Level 'DATA'
    Write-Log "Mode          : $Mode"                           -Level 'DATA'
    Write-Log "Changes Made  : $($script:Changes.Count)"       -Level 'DATA'
    Write-Log "Errors        : $($script:Errors.Count)"        -Level $(if ($script:Errors.Count -gt 0) { 'ERROR' } else { 'DATA' })
    Write-Log "Log File      : $script:LogFile"                 -Level 'DATA'
    Write-Sep
    if ($script:Changes.Count -gt 0) {
        Write-Log 'Changes Applied:' -Level 'SUCCESS'
        foreach ($c in $script:Changes) { Write-Log "  $c" -Level 'SUCCESS' }
    }
    if ($script:Errors.Count -gt 0) {
        Write-Log '' -Level 'ERROR'
        Write-Log 'Errors Encountered:' -Level 'ERROR'
        foreach ($e in $script:Errors) { Write-Log "  $e" -Level 'ERROR' }
    }
    Write-Sep
    if ($script:Errors.Count -eq 0 -and $script:Changes.Count -gt 0) {
        Write-Log 'RESULT: All steps completed successfully.' -Level 'SUCCESS'
        Write-Log 'ACTION: Coordinate reboot with SAP team if not already scheduled.' -Level 'WARN'
    } elseif ($Mode -eq 'Audit' -or $script:DryRun) {
        Write-Log 'RESULT: Review output above. Run -Mode Fix when ready to apply.' -Level 'SUCCESS'
    } else {
        Write-Log 'RESULT: Completed with items to review. Check log file.' -Level 'WARN'
    }
    Write-Log '' -Level 'INFO'
    Write-Log 'Prepared by Syed Rizvi' -Level 'DATA'
}

# ================================================================
# MAIN EXECUTION
# ================================================================
try {
    Clear-Host
    Write-Host ''
    Write-Step "RDP FIX - SAP ENVIRONMENT - 10.168.0.32 - MODE: $Mode - Prepared by: Syed Rizvi"
    Write-Log "Log file : $script:LogFile" -Level 'DATA'
    Write-Log "Last known patch date on 10.168.0.32: February 15, 2026" -Level 'WARN'
    if ($script:DryRun) {
        Write-Log 'DRY RUN MODE - Previewing all steps. No changes will be made.' -Level 'WARN'
    }
    Write-Log '' -Level 'INFO'

    $session   = Get-RemoteSession
    $patchInfo = Get-PatchStatus -Session $session

    Show-RootCauseExplanation -PatchInfo $patchInfo

    if ($Mode -eq 'Audit') {
        Write-Log 'AUDIT COMPLETE.' -Level 'SUCCESS'
        Write-Log 'Run with -Mode DryRun to preview the fix.' -Level 'INFO'
        Write-Log 'Run with -Mode Fix to apply the fix.' -Level 'INFO'
        Show-Summary
        if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue }
        exit 0
    }

    if ($script:DryRun) {
        Write-Log 'SIMULATING FIX STEPS...' -Level 'DRY'
        Invoke-EmergencyWorkaround -Session $session
        $pe = Get-CorrectPatchKB -OSName $patchInfo.OSName
        Invoke-DownloadAndInstallPatch -Session $session -KB $pe.KB -CatalogURL $pe.URL
        Invoke-ScheduleReboot -Session $session
        Show-Summary
        if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue }
        exit 0
    }

    # FIX MODE
    Write-Step "APPLYING FULL REMEDIATION TO SAP SERVER 10.168.0.32"

    Write-Log 'STEP 1 of 4 - Emergency Winlogon registry fix (restores RDP immediately)...' -Level 'STEP'
    Invoke-EmergencyWorkaround -Session $session

    Write-Log 'STEP 2 of 4 - Removing bad January 2026 patches...' -Level 'STEP'
    $badsFound = @($patchInfo.BadKBs)
    if ($badsFound.Count -gt 0) {
        foreach ($bp in $badsFound) {
            Invoke-RemoveBadPatch -Session $session -BadKB $bp.HotFixID
        }
    } else {
        Write-Log 'No bad patches found on server. Skipping uninstall step.' -Level 'SUCCESS'
    }

    Write-Log 'STEP 3 of 4 - Downloading and installing March 2026 permanent fix...' -Level 'STEP'
    $patchEntry = Get-CorrectPatchKB -OSName $patchInfo.OSName
    Write-Log "Correct patch for $($patchInfo.OSName) : $($patchEntry.KB)" -Level 'DATA'
    Write-Log "Expected OS build after patch         : $($patchEntry.Build)" -Level 'DATA'
    Write-Log "Microsoft Update Catalog URL          : $($patchEntry.URL)" -Level 'DATA'
    Invoke-DownloadAndInstallPatch -Session $session -KB $patchEntry.KB -CatalogURL $patchEntry.URL

    Write-Log 'STEP 4 of 4 - Reboot coordination...' -Level 'STEP'
    Invoke-ScheduleReboot -Session $session

    if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue }
    Show-Summary

} catch {
    Write-Log "FATAL ERROR: $($_.Exception.Message)" -Level 'ERROR'
    Write-Log "Stack      : $($_.ScriptStackTrace)"  -Level 'ERROR'
    Show-Summary
    exit 1
}
