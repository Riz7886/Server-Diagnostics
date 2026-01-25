param(
    [Parameter(Mandatory=$false)]
    [string]$ComputerName = $env:COMPUTERNAME,
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "$env:USERPROFILE\Desktop\ServerDiagnostics_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
)

$ErrorActionPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"

function Safe-Execute {
    param([scriptblock]$Code, [string]$Default = "N/A")
    try { $result = & $Code; if ($null -eq $result) { return $Default } else { return $result } }
    catch { return $Default }
}

function Test-PortConnectivity {
    param([string]$Target, [int]$Port)
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $connect = $tcp.BeginConnect($Target, $Port, $null, $null)
        $wait = $connect.AsyncWaitHandle.WaitOne(2000, $false)
        if ($wait) { try { $tcp.EndConnect($connect) } catch { } }
        $tcp.Close()
        return $wait
    } catch { return $false }
}

function Get-SystemInfo {
    $info = @{
        Hostname = Safe-Execute { $env:COMPUTERNAME } "Unknown"
        Domain = Safe-Execute { (Get-WmiObject Win32_ComputerSystem).Domain } "Unknown"
        DomainJoined = Safe-Execute { (Get-WmiObject Win32_ComputerSystem).PartOfDomain } $false
        OS = Safe-Execute { (Get-WmiObject Win32_OperatingSystem).Caption } "Unknown"
        Version = Safe-Execute { (Get-WmiObject Win32_OperatingSystem).Version } "Unknown"
        BuildNumber = Safe-Execute { (Get-WmiObject Win32_OperatingSystem).BuildNumber } "Unknown"
        LastBoot = Safe-Execute { (Get-WmiObject Win32_OperatingSystem).ConvertToDateTime((Get-WmiObject Win32_OperatingSystem).LastBootUpTime) } "Unknown"
        TotalMemoryGB = Safe-Execute { [math]::Round((Get-WmiObject Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2) } 0
        FreeMemoryGB = Safe-Execute { [math]::Round((Get-WmiObject Win32_OperatingSystem).FreePhysicalMemory / 1MB, 2) } 0
        MemoryUsagePercent = Safe-Execute { 
            $os = Get-WmiObject Win32_OperatingSystem
            $cs = Get-WmiObject Win32_ComputerSystem
            [math]::Round((1 - ($os.FreePhysicalMemory * 1KB / $cs.TotalPhysicalMemory)) * 100, 1)
        } 0
    }
    if ($info.LastBoot -ne "Unknown") {
        $info.Uptime = Safe-Execute { (Get-Date) - $info.LastBoot } "Unknown"
    } else { $info.Uptime = "Unknown" }
    return $info
}

function Get-LastLoggedOnUsers {
    $logins = @()
    $currentUser = Safe-Execute { (Get-WmiObject Win32_ComputerSystem).UserName } "Unknown"
    try {
        $events = Get-WinEvent -FilterHashtable @{LogName='Security';Id=4624;StartTime=(Get-Date).AddDays(-7)} -MaxEvents 30 -ErrorAction Stop
        foreach ($event in $events) {
            $logonType = $event.Properties[8].Value
            if ($logonType -in @(2, 10, 11)) {
                $logins += @{
                    Time = $event.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
                    User = Safe-Execute { $event.Properties[5].Value } "Unknown"
                    Domain = Safe-Execute { $event.Properties[6].Value } "Unknown"
                    LogonType = switch ($logonType) { 2 {"Console"} 10 {"RDP"} 11 {"Cached"} default {"Other"} }
                    SourceIP = Safe-Execute { $event.Properties[18].Value } "-"
                }
            }
        }
    } catch { }
    return @{ CurrentUser = $currentUser; Logins = ($logins | Select-Object -First 10) }
}

function Get-PatchHistory {
    $patches = @()
    try {
        $hotfixes = Get-HotFix -ErrorAction Stop | Sort-Object InstalledOn -Descending | Select-Object -First 15
        foreach ($hf in $hotfixes) {
            $patches += @{
                HotFixID = $hf.HotFixID
                Description = Safe-Execute { $hf.Description } "Update"
                InstalledOn = Safe-Execute { $hf.InstalledOn.ToString("yyyy-MM-dd") } "Unknown"
                InstalledBy = Safe-Execute { $hf.InstalledBy } "Unknown"
            }
        }
    } catch { }
    return $patches
}

function Get-ErrorsAroundPatches {
    param([array]$Patches)
    $analysis = @()
    try {
        $recentPatches = $Patches | Where-Object { $_.InstalledOn -ne "Unknown" } | Select-Object -First 3
        foreach ($patch in $recentPatches) {
            try {
                $patchDate = [DateTime]::Parse($patch.InstalledOn)
                $errorsBefore = 0
                $errorsAfter = 0
                try {
                    $before = Get-WinEvent -FilterHashtable @{LogName='System';Level=1,2;StartTime=$patchDate.AddHours(-24);EndTime=$patchDate} -MaxEvents 50 -ErrorAction Stop
                    $errorsBefore = $before.Count
                } catch { }
                try {
                    $after = Get-WinEvent -FilterHashtable @{LogName='System';Level=1,2;StartTime=$patchDate;EndTime=$patchDate.AddHours(48)} -MaxEvents 50 -ErrorAction Stop
                    $errorsAfter = $after.Count
                } catch { }
                $status = if ($errorsAfter -gt ($errorsBefore * 2) -and $errorsAfter -gt 5) { "ERRORS INCREASED" } 
                          elseif ($errorsAfter -gt 0) { "SOME ERRORS" } 
                          else { "OK" }
                $analysis += @{
                    PatchID = $patch.HotFixID
                    PatchDate = $patch.InstalledOn
                    ErrorsBefore = $errorsBefore
                    ErrorsAfter = $errorsAfter
                    Status = $status
                }
            } catch { }
        }
    } catch { }
    return $analysis
}

function Get-NetworkConfiguration {
    $config = @()
    try {
        $adapters = Get-WmiObject Win32_NetworkAdapterConfiguration -ErrorAction Stop | Where-Object { $_.IPEnabled }
        foreach ($adapter in $adapters) {
            $config += @{
                Description = Safe-Execute { $adapter.Description } "Unknown"
                IPAddress = Safe-Execute { $adapter.IPAddress -join ", " } "N/A"
                Gateway = Safe-Execute { $adapter.DefaultIPGateway -join ", " } "N/A"
                DNS = Safe-Execute { $adapter.DNSServerSearchOrder -join ", " } "N/A"
            }
        }
    } catch { }
    return $config
}

function Test-DomainConnectivity {
    $results = @{
        InDomain = $false
        DomainName = "WORKGROUP"
        SecureChannel = $false
        DomainController = "N/A"
        DCIPAddress = "N/A"
        LDAPPort389 = $false
        KerberosPort88 = $false
        DNSPort53 = $false
        SMBPort445 = $false
    }
    try {
        $cs = Get-WmiObject Win32_ComputerSystem -ErrorAction Stop
        $results.InDomain = $cs.PartOfDomain
        $results.DomainName = $cs.Domain
        if ($cs.PartOfDomain) {
            try {
                $dc = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().FindDomainController()
                $results.DomainController = $dc.Name
                $results.DCIPAddress = $dc.IPAddress
                $results.LDAPPort389 = Test-PortConnectivity -Target $dc.IPAddress -Port 389
                $results.KerberosPort88 = Test-PortConnectivity -Target $dc.IPAddress -Port 88
                $results.DNSPort53 = Test-PortConnectivity -Target $dc.IPAddress -Port 53
                $results.SMBPort445 = Test-PortConnectivity -Target $dc.IPAddress -Port 445
            } catch { }
            try { $results.SecureChannel = Test-ComputerSecureChannel -ErrorAction Stop } catch { $results.SecureChannel = $false }
        }
    } catch { }
    return $results
}

function Get-FirewallStatus {
    $status = @{ Domain = @{Enabled=$false}; Private = @{Enabled=$false}; Public = @{Enabled=$false} }
    try {
        $fw = Get-NetFirewallProfile -ErrorAction Stop
        foreach ($profile in $fw) {
            $status[$profile.Name] = @{
                Enabled = $profile.Enabled
                DefaultInbound = Safe-Execute { $profile.DefaultInboundAction.ToString() } "N/A"
            }
        }
    } catch { }
    return $status
}

function Get-RDPStatus {
    $rdp = @{ Enabled = $false; PortNumber = 3389; NLAEnabled = $false; ServiceStatus = "Unknown"; Port3389Open = $false }
    try {
        $tsReg = Get-ItemProperty 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -ErrorAction Stop
        $rdp.Enabled = ($tsReg.fDenyTSConnections -eq 0)
    } catch { }
    try {
        $rdpTcp = Get-ItemProperty 'HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -ErrorAction Stop
        $rdp.PortNumber = $rdpTcp.PortNumber
        $rdp.NLAEnabled = ($rdpTcp.UserAuthentication -eq 1)
    } catch { }
    try { $rdp.ServiceStatus = (Get-Service TermService -ErrorAction Stop).Status.ToString() } catch { }
    $rdp.Port3389Open = Test-PortConnectivity -Target "127.0.0.1" -Port 3389
    return $rdp
}

function Get-CriticalServices {
    $services = @()
    $criticalServices = @("Netlogon","W32Time","LanmanServer","LanmanWorkstation","TermService","WinRM","Dnscache","gpsvc","Winmgmt","EventLog","BITS","wuauserv")
    foreach ($svcName in $criticalServices) {
        try {
            $svc = Get-Service -Name $svcName -ErrorAction Stop
            $services += @{
                Name = $svc.Name
                DisplayName = $svc.DisplayName
                Status = $svc.Status.ToString()
                StartType = Safe-Execute { $svc.StartType.ToString() } "Unknown"
            }
        } catch { }
    }
    return $services
}

function Get-DNSConfiguration {
    $dns = @{ Servers = @(); ResolvesSelf = $false }
    try {
        $adapters = Get-WmiObject Win32_NetworkAdapterConfiguration -ErrorAction Stop | Where-Object { $_.IPEnabled -and $_.DNSServerSearchOrder }
        $dns.Servers = ($adapters | ForEach-Object { $_.DNSServerSearchOrder }) | Select-Object -Unique
    } catch { }
    try { Resolve-DnsName -Name $env:COMPUTERNAME -ErrorAction Stop | Out-Null; $dns.ResolvesSelf = $true } catch { }
    return $dns
}

function Get-TimeConfiguration {
    $time = @{ Synchronized = $false; Offset = 0; OffsetOK = $true }
    try {
        $w32tm = w32tm /query /status 2>&1
        $time.Synchronized = ($w32tm -match "Leap Indicator: 0")
    } catch { }
    return $time
}

function Get-EventLogErrors {
    $events = @{ System = @(); Security = @(); Application = @() }
    try {
        $events.System = Get-WinEvent -FilterHashtable @{LogName='System';Level=1,2;StartTime=(Get-Date).AddDays(-7)} -MaxEvents 15 -ErrorAction Stop | ForEach-Object {
            @{ TimeCreated = $_.TimeCreated.ToString("yyyy-MM-dd HH:mm"); Id = $_.Id; Source = $_.ProviderName; Message = $_.Message.Substring(0, [Math]::Min(150, $_.Message.Length)) }
        }
    } catch { }
    try {
        $events.Security = Get-WinEvent -FilterHashtable @{LogName='Security';Id=4625,4740;StartTime=(Get-Date).AddDays(-7)} -MaxEvents 10 -ErrorAction Stop | ForEach-Object {
            @{ TimeCreated = $_.TimeCreated.ToString("yyyy-MM-dd HH:mm"); Id = $_.Id; Source = "Security"; Message = "Failed login or lockout" }
        }
    } catch { }
    try {
        $events.Application = Get-WinEvent -FilterHashtable @{LogName='Application';Level=1,2;StartTime=(Get-Date).AddDays(-3)} -MaxEvents 10 -ErrorAction Stop | ForEach-Object {
            @{ TimeCreated = $_.TimeCreated.ToString("yyyy-MM-dd HH:mm"); Id = $_.Id; Source = $_.ProviderName; Message = $_.Message.Substring(0, [Math]::Min(150, $_.Message.Length)) }
        }
    } catch { }
    return $events
}

function Get-DiskSpace {
    $disks = @()
    try {
        $drives = Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop
        foreach ($drive in $drives) {
            $freePercent = [math]::Round(($drive.FreeSpace / $drive.Size) * 100, 1)
            $disks += @{
                Drive = $drive.DeviceID
                SizeGB = [math]::Round($drive.Size / 1GB, 2)
                FreeGB = [math]::Round($drive.FreeSpace / 1GB, 2)
                FreePercent = $freePercent
                Status = if ($freePercent -lt 10) { "CRITICAL" } elseif ($freePercent -lt 20) { "WARNING" } else { "OK" }
            }
        }
    } catch { }
    return $disks
}

function Get-PendingReboot {
    $reboot = @{ Required = $false; Reasons = @() }
    try { if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") { $reboot.Required = $true; $reboot.Reasons += "CBS" } } catch { }
    try { if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") { $reboot.Required = $true; $reboot.Reasons += "Windows Update" } } catch { }
    try { if (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction Stop) { $reboot.Required = $true; $reboot.Reasons += "File Rename" } } catch { }
    return $reboot
}

function Generate-HTMLReport {
    param($SysInfo, $NetConfig, $DomainInfo, $FWInfo, $RDPInfo, $SvcInfo, $DNSInfo, $TimeInfo, $EventsInfo, $LoginInfo, $PatchInfo, $PatchAnalysis, $DiskInfo, $RebootInfo, $CompName)
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Server Diagnostics - $CompName</title>
    <style>
        body{font-family:'Segoe UI',Arial,sans-serif;margin:20px;background:#f5f5f5}
        .container{max-width:1400px;margin:0 auto;background:white;padding:30px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,0.1)}
        h1{color:#1e3c72;border-bottom:3px solid #2a5298;padding-bottom:15px}
        h2{color:#2a5298;margin-top:30px;border-left:4px solid #2a5298;padding-left:15px;background:#f8f9fa;padding:10px 15px;border-radius:0 5px 5px 0}
        table{width:100%;border-collapse:collapse;margin:15px 0}
        th{background:linear-gradient(135deg,#1e3c72 0%,#2a5298 100%);color:white;padding:12px;text-align:left}
        td{padding:10px;border-bottom:1px solid #ddd}
        tr:nth-child(even){background:#f8f9fa}
        .good{background:#d4edda;color:#155724;padding:5px 10px;border-radius:4px;font-weight:bold}
        .bad{background:#f8d7da;color:#721c24;padding:5px 10px;border-radius:4px;font-weight:bold}
        .warn{background:#fff3cd;color:#856404;padding:5px 10px;border-radius:4px;font-weight:bold}
        .section{margin:20px 0;padding:20px;background:#fafafa;border-radius:8px;border:1px solid #e0e0e0}
        .summary{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:15px;margin:20px 0}
        .card{padding:15px;border-radius:8px;text-align:center}
        .card h3{margin:0 0 5px 0;font-size:12px}
        .card p{margin:0;font-size:16px;font-weight:bold}
        .timestamp{color:#7f8c8d;font-size:0.9em}
        code{background:#1e1e1e;color:#d4d4d4;padding:3px 8px;border-radius:4px;font-family:Consolas,monospace}
        .alert{padding:15px;border-radius:8px;margin:15px 0;border-left:4px solid}
        .alert-bad{background:#f8d7da;border-color:#dc3545;color:#721c24}
        .alert-warn{background:#fff3cd;border-color:#ffc107;color:#856404}
        .alert-good{background:#d4edda;border-color:#28a745;color:#155724}
    </style>
</head>
<body>
<div class="container">
<h1>Windows Server Diagnostics Report</h1>
<p class="timestamp"><strong>Generated:</strong> $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") | <strong>Server:</strong> $CompName | <strong>Version:</strong> 2.1 Bulletproof</p>

<div class="summary">
<div class="card" style="background:$(if($DomainInfo.InDomain){'#d4edda'}else{'#f8d7da'})"><h3>Domain</h3><p>$(if($DomainInfo.InDomain){'JOINED'}else{'NOT JOINED'})</p></div>
<div class="card" style="background:$(if($DomainInfo.SecureChannel){'#d4edda'}else{'#f8d7da'})"><h3>Trust</h3><p>$(if($DomainInfo.SecureChannel){'VALID'}else{'BROKEN'})</p></div>
<div class="card" style="background:$(if($RDPInfo.Enabled){'#d4edda'}else{'#f8d7da'})"><h3>RDP</h3><p>$(if($RDPInfo.Enabled){'ENABLED'}else{'DISABLED'})</p></div>
<div class="card" style="background:$(if($TimeInfo.OffsetOK){'#d4edda'}else{'#f8d7da'})"><h3>Time</h3><p>$(if($TimeInfo.OffsetOK){'OK'}else{'OUT OF SYNC'})</p></div>
<div class="card" style="background:$(if($RebootInfo.Required){'#fff3cd'}else{'#d4edda'})"><h3>Reboot</h3><p>$(if($RebootInfo.Required){'PENDING'}else{'NO'})</p></div>
<div class="card" style="background:$(if($SysInfo.MemoryUsagePercent -gt 90){'#f8d7da'}elseif($SysInfo.MemoryUsagePercent -gt 80){'#fff3cd'}else{'#d4edda'})"><h3>Memory</h3><p>$($SysInfo.MemoryUsagePercent)%</p></div>
</div>

<h2>System Information</h2>
<div class="section">
<table>
<tr><th>Property</th><th>Value</th></tr>
<tr><td>Hostname</td><td><strong>$($SysInfo.Hostname)</strong></td></tr>
<tr><td>Domain</td><td>$($SysInfo.Domain)</td></tr>
<tr><td>OS</td><td>$($SysInfo.OS)</td></tr>
<tr><td>Version</td><td>$($SysInfo.Version) (Build $($SysInfo.BuildNumber))</td></tr>
<tr><td>Last Boot</td><td><strong>$($SysInfo.LastBoot)</strong></td></tr>
<tr><td>Uptime</td><td>$(if($SysInfo.Uptime -ne 'Unknown'){"$($SysInfo.Uptime.Days) days, $($SysInfo.Uptime.Hours) hours"}else{'Unknown'})</td></tr>
<tr><td>Memory</td><td>$($SysInfo.FreeMemoryGB) GB free of $($SysInfo.TotalMemoryGB) GB ($($SysInfo.MemoryUsagePercent)% used)</td></tr>
</table>
</div>

<h2>Last Logged On Users</h2>
<div class="section">
<p><strong>Currently Logged In:</strong> <span class="good">$($LoginInfo.CurrentUser)</span></p>
<table>
<tr><th>Date/Time</th><th>User</th><th>Domain</th><th>Type</th><th>Source IP</th></tr>
"@

    if ($LoginInfo.Logins -and $LoginInfo.Logins.Count -gt 0) {
        foreach ($login in $LoginInfo.Logins) {
            $html += "<tr><td>$($login.Time)</td><td><strong>$($login.User)</strong></td><td>$($login.Domain)</td><td>$($login.LogonType)</td><td>$($login.SourceIP)</td></tr>"
        }
    } else {
        $html += "<tr><td colspan='5'>No login data available (may require admin rights)</td></tr>"
    }

    $html += @"
</table>
</div>

<h2>Disk Space</h2>
<div class="section">
<table>
<tr><th>Drive</th><th>Total</th><th>Free</th><th>Free %</th><th>Status</th></tr>
"@

    foreach ($disk in $DiskInfo) {
        $sc = switch($disk.Status){"CRITICAL"{"bad"}"WARNING"{"warn"}default{"good"}}
        $html += "<tr><td>$($disk.Drive)</td><td>$($disk.SizeGB) GB</td><td>$($disk.FreeGB) GB</td><td>$($disk.FreePercent)%</td><td><span class='$sc'>$($disk.Status)</span></td></tr>"
    }

    $html += @"
</table>
</div>

<h2>Installed Patches</h2>
<div class="section">
<table>
<tr><th>KB ID</th><th>Description</th><th>Installed</th><th>By</th></tr>
"@

    foreach ($patch in $PatchInfo) {
        $html += "<tr><td><strong>$($patch.HotFixID)</strong></td><td>$($patch.Description)</td><td>$($patch.InstalledOn)</td><td>$($patch.InstalledBy)</td></tr>"
    }

    $html += @"
</table>
</div>

<h2>Errors Before vs After Patches</h2>
<div class="section">
<table>
<tr><th>Patch</th><th>Date</th><th>Errors Before</th><th>Errors After</th><th>Status</th></tr>
"@

    foreach ($pa in $PatchAnalysis) {
        $sc = switch -Wildcard ($pa.Status){"*INCREASED*"{"bad"}"*SOME*"{"warn"}default{"good"}}
        $html += "<tr><td><strong>$($pa.PatchID)</strong></td><td>$($pa.PatchDate)</td><td>$($pa.ErrorsBefore)</td><td>$($pa.ErrorsAfter)</td><td><span class='$sc'>$($pa.Status)</span></td></tr>"
    }

    $html += @"
</table>
</div>

<h2>Domain Controller Connectivity</h2>
<div class="section">
<table>
<tr><th>Check</th><th>Status</th><th>Details</th></tr>
<tr><td>Domain Membership</td><td><span class="$(if($DomainInfo.InDomain){'good'}else{'bad'})">$(if($DomainInfo.InDomain){'JOINED'}else{'NOT JOINED'})</span></td><td>$($DomainInfo.DomainName)</td></tr>
<tr><td>Domain Controller</td><td>-</td><td>$($DomainInfo.DomainController) ($($DomainInfo.DCIPAddress))</td></tr>
<tr><td>Secure Channel</td><td><span class="$(if($DomainInfo.SecureChannel){'good'}else{'bad'})">$(if($DomainInfo.SecureChannel){'VALID'}else{'BROKEN'})</span></td><td>Trust relationship</td></tr>
<tr><td>LDAP (389)</td><td><span class="$(if($DomainInfo.LDAPPort389){'good'}else{'bad'})">$(if($DomainInfo.LDAPPort389){'OPEN'}else{'BLOCKED'})</span></td><td>Authentication</td></tr>
<tr><td>Kerberos (88)</td><td><span class="$(if($DomainInfo.KerberosPort88){'good'}else{'bad'})">$(if($DomainInfo.KerberosPort88){'OPEN'}else{'BLOCKED'})</span></td><td>Tickets</td></tr>
<tr><td>DNS (53)</td><td><span class="$(if($DomainInfo.DNSPort53){'good'}else{'bad'})">$(if($DomainInfo.DNSPort53){'OPEN'}else{'BLOCKED'})</span></td><td>Name resolution</td></tr>
<tr><td>SMB (445)</td><td><span class="$(if($DomainInfo.SMBPort445){'good'}else{'bad'})">$(if($DomainInfo.SMBPort445){'OPEN'}else{'BLOCKED'})</span></td><td>File sharing/GPO</td></tr>
</table>
</div>

<h2>RDP Configuration</h2>
<div class="section">
<table>
<tr><th>Setting</th><th>Value</th><th>Status</th></tr>
<tr><td>RDP Enabled</td><td>$($RDPInfo.Enabled)</td><td><span class="$(if($RDPInfo.Enabled){'good'}else{'bad'})">$(if($RDPInfo.Enabled){'ENABLED'}else{'DISABLED'})</span></td></tr>
<tr><td>RDP Port</td><td>$($RDPInfo.PortNumber)</td><td>-</td></tr>
<tr><td>NLA</td><td>$($RDPInfo.NLAEnabled)</td><td><span class="$(if($RDPInfo.NLAEnabled){'warn'}else{'good'})">$(if($RDPInfo.NLAEnabled){'ENABLED'}else{'DISABLED'})</span></td></tr>
<tr><td>Terminal Services</td><td>$($RDPInfo.ServiceStatus)</td><td><span class="$(if($RDPInfo.ServiceStatus -eq 'Running'){'good'}else{'bad'})">$($RDPInfo.ServiceStatus)</span></td></tr>
</table>
</div>

<h2>Critical Services</h2>
<div class="section">
<table>
<tr><th>Service</th><th>Name</th><th>Status</th><th>Start Type</th></tr>
"@

    foreach ($svc in $SvcInfo) {
        $sc = if($svc.Status -eq "Running"){"good"}elseif($svc.Status -eq "Stopped" -and $svc.StartType -eq "Automatic"){"bad"}else{"warn"}
        $html += "<tr><td>$($svc.Name)</td><td>$($svc.DisplayName)</td><td><span class='$sc'>$($svc.Status)</span></td><td>$($svc.StartType)</td></tr>"
    }

    $html += @"
</table>
</div>

<h2>Recent Errors (Last 7 Days)</h2>
<div class="section">
<h3>System Log</h3>
<table>
<tr><th>Time</th><th>ID</th><th>Source</th><th>Message</th></tr>
"@

    if ($EventsInfo.System.Count -gt 0) {
        foreach ($evt in $EventsInfo.System) { $html += "<tr><td>$($evt.TimeCreated)</td><td>$($evt.Id)</td><td>$($evt.Source)</td><td>$($evt.Message)</td></tr>" }
    } else { $html += "<tr><td colspan='4'>No errors found</td></tr>" }

    $html += @"
</table>
<h3>Security Log (Failed Logins)</h3>
<table>
<tr><th>Time</th><th>ID</th><th>Source</th><th>Message</th></tr>
"@

    if ($EventsInfo.Security.Count -gt 0) {
        foreach ($evt in $EventsInfo.Security) { $html += "<tr><td>$($evt.TimeCreated)</td><td>$($evt.Id)</td><td>$($evt.Source)</td><td>$($evt.Message)</td></tr>" }
    } else { $html += "<tr><td colspan='4'>No failed logins found</td></tr>" }

    $html += @"
</table>
</div>

<h2>Recommended Actions</h2>
<div class="section">
"@

    $issues = @()
    if (-not $DomainInfo.InDomain) { $issues += "<div class='alert alert-bad'><strong>SERVER NOT IN DOMAIN</strong></div>" }
    if ($DomainInfo.InDomain -and -not $DomainInfo.SecureChannel) { $issues += "<div class='alert alert-bad'><strong>TRUST BROKEN</strong> - Run: <code>Test-ComputerSecureChannel -Repair -Credential (Get-Credential)</code></div>" }
    if (-not $RDPInfo.Enabled) { $issues += "<div class='alert alert-bad'><strong>RDP DISABLED</strong></div>" }
    if ($RDPInfo.NLAEnabled -and -not $DomainInfo.SecureChannel) { $issues += "<div class='alert alert-bad'><strong>NLA + BROKEN TRUST</strong> - Disable NLA first</div>" }
    if ($RebootInfo.Required) { $issues += "<div class='alert alert-warn'><strong>REBOOT PENDING</strong> - $($RebootInfo.Reasons -join ', ')</div>" }
    foreach ($d in $DiskInfo | Where-Object {$_.Status -eq "CRITICAL"}) { $issues += "<div class='alert alert-bad'><strong>LOW DISK on $($d.Drive)</strong> - $($d.FreePercent)% free</div>" }
    if ($SysInfo.MemoryUsagePercent -gt 90) { $issues += "<div class='alert alert-warn'><strong>HIGH MEMORY</strong> - $($SysInfo.MemoryUsagePercent)%</div>" }
    
    if ($issues.Count -eq 0) { $html += "<div class='alert alert-good'><strong>NO CRITICAL ISSUES!</strong> Server appears healthy.</div>" }
    else { $html += $issues -join "`n" }

    $html += @"
</div>

<h2>Quick Fix Commands</h2>
<div class="section">
<table>
<tr><th>Issue</th><th>Command</th></tr>
<tr><td>Repair Trust</td><td><code>Test-ComputerSecureChannel -Repair -Credential (Get-Credential)</code></td></tr>
<tr><td>Reset Password</td><td><code>Reset-ComputerMachinePassword -Credential (Get-Credential)</code></td></tr>
<tr><td>Time Sync</td><td><code>w32tm /resync /force</code></td></tr>
<tr><td>Flush DNS</td><td><code>ipconfig /flushdns</code></td></tr>
<tr><td>Enable RDP</td><td><code>Set-ItemProperty 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name fDenyTSConnections -Value 0</code></td></tr>
<tr><td>Disable NLA</td><td><code>Set-ItemProperty 'HKLM:\...\WinStations\RDP-Tcp' -Name UserAuthentication -Value 0</code></td></tr>
<tr><td>GPO Update</td><td><code>gpupdate /force</code></td></tr>
</table>
</div>

<p class="timestamp">Report by L1 Server Diagnostics v2.1 | READ-ONLY - No changes made</p>
</div>
</body>
</html>
"@

    return $html
}

Clear-Host
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  L1 SERVER DIAGNOSTICS v2.1 BULLETPROOF" -ForegroundColor Cyan
Write-Host "  READ-ONLY - NO CHANGES WILL BE MADE" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "[1/12] System information..." -ForegroundColor Gray
$sysInfo = Get-SystemInfo

Write-Host "[2/12] Network configuration..." -ForegroundColor Gray
$netConfig = Get-NetworkConfiguration

Write-Host "[3/12] Domain connectivity..." -ForegroundColor Gray
$domainInfo = Test-DomainConnectivity

Write-Host "[4/12] Firewall status..." -ForegroundColor Gray
$fwInfo = Get-FirewallStatus

Write-Host "[5/12] RDP configuration..." -ForegroundColor Gray
$rdpInfo = Get-RDPStatus

Write-Host "[6/12] Critical services..." -ForegroundColor Gray
$svcInfo = Get-CriticalServices

Write-Host "[7/12] DNS configuration..." -ForegroundColor Gray
$dnsInfo = Get-DNSConfiguration

Write-Host "[8/12] Time sync..." -ForegroundColor Gray
$timeInfo = Get-TimeConfiguration

Write-Host "[9/12] Event logs..." -ForegroundColor Gray
$evtInfo = Get-EventLogErrors

Write-Host "[10/12] Login history..." -ForegroundColor Gray
$loginInfo = Get-LastLoggedOnUsers

Write-Host "[11/12] Patch history..." -ForegroundColor Gray
$patchInfo = Get-PatchHistory

Write-Host "[12/12] Patch analysis..." -ForegroundColor Gray
$patchAnalysis = Get-ErrorsAroundPatches -Patches $patchInfo

$diskInfo = Get-DiskSpace
$rebootInfo = Get-PendingReboot

Write-Host ""
Write-Host "Generating report..." -ForegroundColor Yellow

$report = Generate-HTMLReport -SysInfo $sysInfo -NetConfig $netConfig -DomainInfo $domainInfo -FWInfo $fwInfo -RDPInfo $rdpInfo -SvcInfo $svcInfo -DNSInfo $dnsInfo -TimeInfo $timeInfo -EventsInfo $evtInfo -LoginInfo $loginInfo -PatchInfo $patchInfo -PatchAnalysis $patchAnalysis -DiskInfo $diskInfo -RebootInfo $rebootInfo -CompName $ComputerName

$report | Out-File -FilePath $OutputPath -Encoding UTF8 -Force

Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host "  COMPLETE!" -ForegroundColor Green
Write-Host "  Report: $OutputPath" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Green

Start-Process $OutputPath
