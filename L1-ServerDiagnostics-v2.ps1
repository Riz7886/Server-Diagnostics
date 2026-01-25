param(
    [Parameter(Mandatory=$false)]
    [string]$ComputerName = $env:COMPUTERNAME,
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "$env:USERPROFILE\Desktop\ServerDiagnostics_$(Get-Date -Format 'yyyyMMdd_HHmmss').html",
    [Parameter(Mandatory=$false)]
    [switch]$ConnectAzure,
    [Parameter(Mandatory=$false)]
    [switch]$ConnectAWS,
    [Parameter(Mandatory=$false)]
    [switch]$AutoConnect
)

$ErrorActionPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"

function Install-RequiredModules {
    $modules = @("Az.Accounts", "Az.Compute", "AWS.Tools.Common", "AWS.Tools.EC2", "AWS.Tools.SSM")
    foreach ($module in $modules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            try {
                Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser -SkipPublisherCheck
            } catch { }
        }
    }
}

function Connect-CloudProviders {
    param([bool]$Azure, [bool]$AWS, [bool]$Auto)
    $results = @{Azure = $false; AWS = $false}
    if ($Azure) {
        try {
            Import-Module Az.Accounts -ErrorAction SilentlyContinue
            if ($Auto) {
                $context = Get-AzContext
                if ($context) { $results.Azure = $true }
            } else {
                Connect-AzAccount -ErrorAction Stop | Out-Null
                $results.Azure = $true
            }
        } catch { }
    }
    if ($AWS) {
        try {
            Import-Module AWS.Tools.Common -ErrorAction SilentlyContinue
            if ($Auto) {
                $creds = Get-AWSCredential
                if ($creds) { $results.AWS = $true }
            } else {
                Set-AWSCredential -ProfileName default -ErrorAction SilentlyContinue
                $results.AWS = $true
            }
        } catch { }
    }
    return $results
}

function Test-PortConnectivity {
    param([string]$Target, [int]$Port, [int]$Timeout = 2000)
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $connect = $tcp.BeginConnect($Target, $Port, $null, $null)
        $wait = $connect.AsyncWaitHandle.WaitOne($Timeout, $false)
        if ($wait) {
            $tcp.EndConnect($connect)
            $tcp.Close()
            return $true
        }
        $tcp.Close()
        return $false
    } catch { return $false }
}

function Get-SystemInfo {
    param([string]$Computer)
    $info = @{}
    try {
        $os = Get-WmiObject Win32_OperatingSystem -ComputerName $Computer
        $cs = Get-WmiObject Win32_ComputerSystem -ComputerName $Computer
        $info.Hostname = $cs.Name
        $info.Domain = $cs.Domain
        $info.DomainJoined = $cs.PartOfDomain
        $info.OS = $os.Caption
        $info.Version = $os.Version
        $info.BuildNumber = $os.BuildNumber
        $info.LastBoot = $os.ConvertToDateTime($os.LastBootUpTime)
        $info.Uptime = (Get-Date) - $info.LastBoot
        $info.InstallDate = $os.ConvertToDateTime($os.InstallDate)
        $info.TotalMemoryGB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
        $info.FreeMemoryGB = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
        $info.MemoryUsagePercent = [math]::Round((1 - ($os.FreePhysicalMemory * 1KB / $cs.TotalPhysicalMemory)) * 100, 1)
    } catch {
        $info.Error = "Unable to retrieve system information"
    }
    return $info
}

function Get-LastLoggedOnUsers {
    param([string]$Computer)
    $logins = @()
    try {
        $events = Get-WinEvent -ComputerName $Computer -FilterHashtable @{
            LogName = 'Security'
            Id = 4624
            StartTime = (Get-Date).AddDays(-7)
        } -MaxEvents 50 | Where-Object {
            $_.Properties[8].Value -in @(2, 10, 11)
        }
        foreach ($event in $events) {
            $logins += @{
                Time = $event.TimeCreated
                User = $event.Properties[5].Value
                Domain = $event.Properties[6].Value
                LogonType = switch ($event.Properties[8].Value) {
                    2 { "Interactive (Console)" }
                    10 { "RemoteInteractive (RDP)" }
                    11 { "CachedInteractive" }
                    default { $event.Properties[8].Value }
                }
                SourceIP = $event.Properties[18].Value
            }
        }
        $lastUser = Get-WmiObject Win32_ComputerSystem -ComputerName $Computer
        $logins = @(@{
            CurrentUser = $lastUser.UserName
            Logins = $logins | Select-Object -First 10
        })
    } catch {
        $logins = @(@{ Error = "Unable to retrieve login history" })
    }
    return $logins
}

function Get-PatchHistory {
    param([string]$Computer)
    $patches = @()
    try {
        $hotfixes = Get-HotFix -ComputerName $Computer | Sort-Object InstalledOn -Descending | Select-Object -First 20
        foreach ($hf in $hotfixes) {
            $patches += @{
                HotFixID = $hf.HotFixID
                Description = $hf.Description
                InstalledOn = $hf.InstalledOn
                InstalledBy = $hf.InstalledBy
            }
        }
    } catch { }
    try {
        $session = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $historyCount = $searcher.GetTotalHistoryCount()
        if ($historyCount -gt 0) {
            $history = $searcher.QueryHistory(0, [Math]::Min(20, $historyCount))
            foreach ($update in $history) {
                if ($update.Title) {
                    $patches += @{
                        HotFixID = if ($update.Title -match 'KB\d+') { $matches[0] } else { "N/A" }
                        Description = $update.Title
                        InstalledOn = $update.Date
                        InstalledBy = "Windows Update"
                        ResultCode = switch ($update.ResultCode) {
                            0 { "Not Started" }
                            1 { "In Progress" }
                            2 { "Succeeded" }
                            3 { "Succeeded With Errors" }
                            4 { "Failed" }
                            5 { "Aborted" }
                            default { $update.ResultCode }
                        }
                    }
                }
            }
        }
    } catch { }
    return $patches | Sort-Object { $_.InstalledOn } -Descending | Select-Object -First 25
}

function Get-ErrorsAroundPatches {
    param([string]$Computer, [array]$Patches)
    $analysis = @()
    try {
        $patchDates = $Patches | Where-Object { $_.InstalledOn } | ForEach-Object { 
            if ($_.InstalledOn -is [DateTime]) { $_.InstalledOn } 
            else { [DateTime]::Parse($_.InstalledOn) }
        } | Select-Object -Unique | Sort-Object -Descending | Select-Object -First 5
        foreach ($patchDate in $patchDates) {
            $beforeStart = $patchDate.AddHours(-24)
            $afterEnd = $patchDate.AddHours(48)
            $errorsBefore = @()
            $errorsAfter = @()
            try {
                $errorsBefore = Get-WinEvent -ComputerName $Computer -FilterHashtable @{
                    LogName = 'System'
                    Level = 1,2
                    StartTime = $beforeStart
                    EndTime = $patchDate
                } -MaxEvents 10 | ForEach-Object {
                    @{ Time = $_.TimeCreated; Id = $_.Id; Message = $_.Message.Substring(0, [Math]::Min(150, $_.Message.Length)) }
                }
            } catch { }
            try {
                $errorsAfter = Get-WinEvent -ComputerName $Computer -FilterHashtable @{
                    LogName = 'System'
                    Level = 1,2
                    StartTime = $patchDate
                    EndTime = $afterEnd
                } -MaxEvents 10 | ForEach-Object {
                    @{ Time = $_.TimeCreated; Id = $_.Id; Message = $_.Message.Substring(0, [Math]::Min(150, $_.Message.Length)) }
                }
            } catch { }
            $relatedPatch = $Patches | Where-Object { 
                $d = if ($_.InstalledOn -is [DateTime]) { $_.InstalledOn } else { try { [DateTime]::Parse($_.InstalledOn) } catch { $null } }
                $d -and [Math]::Abs(($d - $patchDate).TotalHours) -lt 2
            } | Select-Object -First 1
            $analysis += @{
                PatchDate = $patchDate
                PatchID = if ($relatedPatch) { $relatedPatch.HotFixID } else { "Unknown" }
                PatchDescription = if ($relatedPatch) { $relatedPatch.Description } else { "N/A" }
                ErrorsBefore = $errorsBefore
                ErrorsBeforeCount = $errorsBefore.Count
                ErrorsAfter = $errorsAfter
                ErrorsAfterCount = $errorsAfter.Count
                Status = if ($errorsAfter.Count -gt $errorsBefore.Count) { "ERRORS INCREASED AFTER PATCH" } 
                         elseif ($errorsAfter.Count -gt 0) { "SOME ERRORS AFTER PATCH" } 
                         else { "NO ISSUES DETECTED" }
            }
        }
    } catch { }
    return $analysis
}

function Get-NetworkConfiguration {
    param([string]$Computer)
    $config = @()
    try {
        $adapters = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $Computer | Where-Object { $_.IPEnabled }
        foreach ($adapter in $adapters) {
            $config += @{
                Description = $adapter.Description
                IPAddress = $adapter.IPAddress -join ", "
                Subnet = $adapter.IPSubnet -join ", "
                Gateway = $adapter.DefaultIPGateway -join ", "
                DNS = $adapter.DNSServerSearchOrder -join ", "
                DHCPEnabled = $adapter.DHCPEnabled
                MACAddress = $adapter.MACAddress
            }
        }
    } catch { }
    return $config
}

function Test-DomainConnectivity {
    param([string]$Computer)
    $results = @{}
    try {
        $cs = Get-WmiObject Win32_ComputerSystem -ComputerName $Computer
        if ($cs.PartOfDomain) {
            $results.InDomain = $true
            $results.DomainName = $cs.Domain
            try {
                $dc = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().FindDomainController()
                $results.DomainController = $dc.Name
                $results.DCIPAddress = $dc.IPAddress
                $results.LDAPPort389 = Test-PortConnectivity -Target $dc.IPAddress -Port 389
                $results.LDAPPort636 = Test-PortConnectivity -Target $dc.IPAddress -Port 636
                $results.KerberosPort88 = Test-PortConnectivity -Target $dc.IPAddress -Port 88
                $results.DNSPort53 = Test-PortConnectivity -Target $dc.IPAddress -Port 53
                $results.SMBPort445 = Test-PortConnectivity -Target $dc.IPAddress -Port 445
                $results.RPCPort135 = Test-PortConnectivity -Target $dc.IPAddress -Port 135
                $results.GCPort3268 = Test-PortConnectivity -Target $dc.IPAddress -Port 3268
            } catch {
                $results.DomainController = "Unable to find DC"
                $results.DCIPAddress = "N/A"
            }
            try {
                $nltest = nltest /dsgetdc:$($cs.Domain) 2>&1
                $results.NLTestSuccess = $LASTEXITCODE -eq 0
            } catch { $results.NLTestSuccess = $false }
            try {
                $results.SecureChannel = Test-ComputerSecureChannel -ErrorAction Stop
            } catch { $results.SecureChannel = $false }
        } else {
            $results.InDomain = $false
            $results.DomainName = "WORKGROUP"
        }
    } catch {
        $results.Error = $_.Exception.Message
    }
    return $results
}

function Get-FirewallStatus {
    param([string]$Computer)
    $status = @{}
    try {
        $fw = Get-NetFirewallProfile -CimSession $Computer
        foreach ($profile in $fw) {
            $status[$profile.Name] = @{
                Enabled = $profile.Enabled
                DefaultInbound = $profile.DefaultInboundAction
                DefaultOutbound = $profile.DefaultOutboundAction
            }
        }
        $rules = Get-NetFirewallRule -CimSession $Computer | Where-Object { 
            $_.Enabled -eq $true -and $_.Direction -eq "Inbound" -and $_.Action -eq "Block"
        } | Select-Object -First 20
        $status.BlockingRules = $rules | ForEach-Object { $_.DisplayName }
    } catch {
        $status.Error = "Unable to retrieve firewall status"
    }
    return $status
}

function Get-RDPStatus {
    param([string]$Computer)
    $rdp = @{}
    try {
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer)
        $key = $reg.OpenSubKey('SYSTEM\CurrentControlSet\Control\Terminal Server')
        $rdp.Enabled = ($key.GetValue('fDenyTSConnections') -eq 0)
        $key2 = $reg.OpenSubKey('SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp')
        $rdp.PortNumber = $key2.GetValue('PortNumber')
        $rdp.UserAuthentication = $key2.GetValue('UserAuthentication')
        $rdp.NLAEnabled = ($rdp.UserAuthentication -eq 1)
        $svc = Get-Service -ComputerName $Computer -Name TermService
        $rdp.ServiceStatus = $svc.Status.ToString()
        $rdp.Port3389Open = Test-PortConnectivity -Target $Computer -Port 3389
    } catch {
        $rdp.Error = "Unable to retrieve RDP status"
    }
    return $rdp
}

function Get-CriticalServices {
    param([string]$Computer)
    $services = @()
    $criticalServices = @(
        "Netlogon", "NTDS", "DNS", "W32Time", "LanmanServer", "LanmanWorkstation",
        "RemoteRegistry", "TermService", "WinRM", "DFSR", "Dnscache", "Dhcp",
        "NlaSvc", "gpsvc", "CryptSvc", "Winmgmt", "Schedule", "EventLog",
        "BITS", "wuauserv", "TrustedInstaller", "AppIDSvc"
    )
    foreach ($svcName in $criticalServices) {
        try {
            $svc = Get-Service -ComputerName $Computer -Name $svcName -ErrorAction SilentlyContinue
            if ($svc) {
                $services += @{
                    Name = $svc.Name
                    DisplayName = $svc.DisplayName
                    Status = $svc.Status.ToString()
                    StartType = $svc.StartType.ToString()
                }
            }
        } catch { }
    }
    return $services
}

function Get-DNSConfiguration {
    param([string]$Computer)
    $dns = @{}
    try {
        $adapters = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $Computer | Where-Object { $_.IPEnabled }
        $dns.Servers = ($adapters | ForEach-Object { $_.DNSServerSearchOrder }) | Where-Object { $_ } | Select-Object -Unique
        foreach ($server in $dns.Servers) {
            if ($server) {
                $dns["DNS_$server"] = Test-PortConnectivity -Target $server -Port 53
            }
        }
        try {
            $resolve = Resolve-DnsName -Name $Computer -ErrorAction Stop
            $dns.ResolvesSelf = $true
        } catch { $dns.ResolvesSelf = $false }
    } catch {
        $dns.Error = "Unable to retrieve DNS configuration"
    }
    return $dns
}

function Get-TimeConfiguration {
    param([string]$Computer)
    $time = @{}
    try {
        $w32tm = w32tm /query /status /computer:$Computer 2>&1
        $time.Output = $w32tm -join "`n"
        $time.Synchronized = ($w32tm -match "Leap Indicator: 0")
        $localTime = Get-Date
        $remoteTime = Get-WmiObject Win32_LocalTime -ComputerName $Computer
        $remoteDateTime = Get-Date -Year $remoteTime.Year -Month $remoteTime.Month -Day $remoteTime.Day -Hour $remoteTime.Hour -Minute $remoteTime.Minute -Second $remoteTime.Second
        $time.Offset = [math]::Abs(($localTime - $remoteDateTime).TotalSeconds)
        $time.OffsetOK = ($time.Offset -lt 300)
    } catch {
        $time.Error = "Unable to retrieve time configuration"
    }
    return $time
}

function Get-EventLogErrors {
    param([string]$Computer)
    $events = @{}
    try {
        $events.System = Get-WinEvent -ComputerName $Computer -FilterHashtable @{
            LogName = 'System'
            Level = 1,2
            StartTime = (Get-Date).AddDays(-7)
        } -MaxEvents 25 | ForEach-Object {
            @{
                TimeCreated = $_.TimeCreated
                Id = $_.Id
                Source = $_.ProviderName
                Message = $_.Message.Substring(0, [Math]::Min(200, $_.Message.Length))
            }
        }
    } catch { $events.System = @() }
    try {
        $events.Security = Get-WinEvent -ComputerName $Computer -FilterHashtable @{
            LogName = 'Security'
            Id = 4625,4771,4776,4740
            StartTime = (Get-Date).AddDays(-7)
        } -MaxEvents 15 | ForEach-Object {
            @{
                TimeCreated = $_.TimeCreated
                Id = $_.Id
                Source = $_.ProviderName
                Message = $_.Message.Substring(0, [Math]::Min(200, $_.Message.Length))
            }
        }
    } catch { $events.Security = @() }
    try {
        $events.Application = Get-WinEvent -ComputerName $Computer -FilterHashtable @{
            LogName = 'Application'
            Level = 1,2
            StartTime = (Get-Date).AddDays(-3)
        } -MaxEvents 15 | ForEach-Object {
            @{
                TimeCreated = $_.TimeCreated
                Id = $_.Id
                Source = $_.ProviderName
                Message = $_.Message.Substring(0, [Math]::Min(200, $_.Message.Length))
            }
        }
    } catch { $events.Application = @() }
    return $events
}

function Get-DiskSpace {
    param([string]$Computer)
    $disks = @()
    try {
        $drives = Get-WmiObject Win32_LogicalDisk -ComputerName $Computer -Filter "DriveType=3"
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
    param([string]$Computer)
    $reboot = @{ Required = $false; Reasons = @() }
    try {
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Computer)
        if ($reg.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending')) {
            $reboot.Required = $true
            $reboot.Reasons += "Component Based Servicing"
        }
        if ($reg.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')) {
            $reboot.Required = $true
            $reboot.Reasons += "Windows Update"
        }
        $pfro = $reg.OpenSubKey('SYSTEM\CurrentControlSet\Control\Session Manager')
        if ($pfro) {
            $val = $pfro.GetValue('PendingFileRenameOperations')
            if ($val) {
                $reboot.Required = $true
                $reboot.Reasons += "Pending File Rename Operations"
            }
        }
    } catch { }
    return $reboot
}

function Generate-HTMLReport {
    param(
        [hashtable]$SystemInfo,
        [array]$NetworkConfig,
        [hashtable]$DomainInfo,
        [hashtable]$FirewallInfo,
        [hashtable]$RDPInfo,
        [array]$ServicesInfo,
        [hashtable]$DNSInfo,
        [hashtable]$TimeInfo,
        [hashtable]$EventsInfo,
        [array]$LoginInfo,
        [array]$PatchInfo,
        [array]$PatchAnalysis,
        [array]$DiskInfo,
        [hashtable]$RebootInfo,
        [string]$ComputerName
    )
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Server Diagnostics Report - $ComputerName</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 1400px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #1e3c72; border-bottom: 3px solid #2a5298; padding-bottom: 15px; }
        h2 { color: #2a5298; margin-top: 30px; border-left: 4px solid #2a5298; padding-left: 15px; background: #f8f9fa; padding: 10px 15px; border-radius: 0 5px 5px 0; }
        h3 { color: #34495e; margin-top: 20px; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th { background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); color: white; padding: 12px; text-align: left; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr:nth-child(even) { background: #f8f9fa; }
        tr:hover { background: #e8f4fc; }
        .status-good { background: #d4edda; color: #155724; padding: 5px 10px; border-radius: 4px; font-weight: bold; }
        .status-bad { background: #f8d7da; color: #721c24; padding: 5px 10px; border-radius: 4px; font-weight: bold; }
        .status-warn { background: #fff3cd; color: #856404; padding: 5px 10px; border-radius: 4px; font-weight: bold; }
        .section { margin: 20px 0; padding: 20px; background: #fafafa; border-radius: 8px; border: 1px solid #e0e0e0; }
        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 15px; margin: 20px 0; }
        .summary-card { padding: 20px; border-radius: 8px; text-align: center; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
        .summary-card h3 { margin: 0 0 10px 0; font-size: 14px; }
        .summary-card p { margin: 0; font-size: 18px; font-weight: bold; }
        .timestamp { color: #7f8c8d; font-size: 0.9em; }
        code { background: #1e1e1e; color: #d4d4d4; padding: 3px 8px; border-radius: 4px; font-family: Consolas, monospace; }
        .alert { padding: 15px; border-radius: 8px; margin: 15px 0; border-left: 4px solid; }
        .alert-danger { background: #f8d7da; border-color: #dc3545; color: #721c24; }
        .alert-warning { background: #fff3cd; border-color: #ffc107; color: #856404; }
        .alert-success { background: #d4edda; border-color: #28a745; color: #155724; }
        .patch-timeline { border-left: 3px solid #2a5298; padding-left: 20px; margin: 20px 0; }
        .patch-item { margin: 15px 0; padding: 10px; background: #f8f9fa; border-radius: 5px; }
        .patch-date { font-weight: bold; color: #1e3c72; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Windows Server Diagnostics Report</h1>
        <p class="timestamp"><strong>Generated:</strong> $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") | <strong>Target:</strong> $ComputerName | <strong>Script Version:</strong> 2.0</p>
        
        <div class="summary">
            <div class="summary-card" style="background: $(if($DomainInfo.InDomain){'#d4edda'}else{'#f8d7da'});">
                <h3>Domain Status</h3>
                <p>$(if($DomainInfo.InDomain){'JOINED'}else{'NOT JOINED'})</p>
            </div>
            <div class="summary-card" style="background: $(if($DomainInfo.SecureChannel){'#d4edda'}else{'#f8d7da'});">
                <h3>Trust Relationship</h3>
                <p>$(if($DomainInfo.SecureChannel){'VALID'}else{'BROKEN'})</p>
            </div>
            <div class="summary-card" style="background: $(if($RDPInfo.Enabled -and $RDPInfo.Port3389Open){'#d4edda'}else{'#f8d7da'});">
                <h3>RDP Status</h3>
                <p>$(if($RDPInfo.Enabled -and $RDPInfo.Port3389Open){'ACCESSIBLE'}else{'BLOCKED'})</p>
            </div>
            <div class="summary-card" style="background: $(if($TimeInfo.OffsetOK){'#d4edda'}else{'#f8d7da'});">
                <h3>Time Sync</h3>
                <p>$(if($TimeInfo.OffsetOK){'SYNCHRONIZED'}else{'OUT OF SYNC'})</p>
            </div>
            <div class="summary-card" style="background: $(if($RebootInfo.Required){'#fff3cd'}else{'#d4edda'});">
                <h3>Pending Reboot</h3>
                <p>$(if($RebootInfo.Required){'YES'}else{'NO'})</p>
            </div>
            <div class="summary-card" style="background: $(if($SystemInfo.MemoryUsagePercent -gt 90){'#f8d7da'}elseif($SystemInfo.MemoryUsagePercent -gt 80){'#fff3cd'}else{'#d4edda'});">
                <h3>Memory Usage</h3>
                <p>$($SystemInfo.MemoryUsagePercent)%</p>
            </div>
        </div>

        <h2>System Information</h2>
        <div class="section">
            <table>
                <tr><th>Property</th><th>Value</th></tr>
                <tr><td>Hostname</td><td><strong>$($SystemInfo.Hostname)</strong></td></tr>
                <tr><td>Domain</td><td>$($SystemInfo.Domain)</td></tr>
                <tr><td>Domain Joined</td><td><span class="$(if($SystemInfo.DomainJoined){'status-good'}else{'status-bad'})">$($SystemInfo.DomainJoined)</span></td></tr>
                <tr><td>Operating System</td><td>$($SystemInfo.OS)</td></tr>
                <tr><td>Version / Build</td><td>$($SystemInfo.Version) (Build $($SystemInfo.BuildNumber))</td></tr>
                <tr><td>Last Boot Time</td><td><strong>$($SystemInfo.LastBoot)</strong></td></tr>
                <tr><td>Uptime</td><td>$($SystemInfo.Uptime.Days) days, $($SystemInfo.Uptime.Hours) hours, $($SystemInfo.Uptime.Minutes) minutes</td></tr>
                <tr><td>Total Memory</td><td>$($SystemInfo.TotalMemoryGB) GB</td></tr>
                <tr><td>Free Memory</td><td>$($SystemInfo.FreeMemoryGB) GB</td></tr>
                <tr><td>Memory Usage</td><td><span class="$(if($SystemInfo.MemoryUsagePercent -gt 90){'status-bad'}elseif($SystemInfo.MemoryUsagePercent -gt 80){'status-warn'}else{'status-good'})">$($SystemInfo.MemoryUsagePercent)%</span></td></tr>
            </table>
        </div>

        <h2>Last Logged On Users</h2>
        <div class="section">
"@

    if ($LoginInfo -and $LoginInfo[0].CurrentUser) {
        $html += "<p><strong>Currently Logged In:</strong> <span class='status-good'>$($LoginInfo[0].CurrentUser)</span></p>"
    }
    
    $html += @"
            <table>
                <tr><th>Date/Time</th><th>User</th><th>Domain</th><th>Logon Type</th><th>Source IP</th></tr>
"@

    if ($LoginInfo -and $LoginInfo[0].Logins) {
        foreach ($login in $LoginInfo[0].Logins) {
            $html += "<tr><td>$($login.Time)</td><td><strong>$($login.User)</strong></td><td>$($login.Domain)</td><td>$($login.LogonType)</td><td>$($login.SourceIP)</td></tr>"
        }
    } else {
        $html += "<tr><td colspan='5'>Unable to retrieve login history (may require elevated permissions)</td></tr>"
    }
    
    $html += @"
            </table>
        </div>

        <h2>Disk Space</h2>
        <div class="section">
            <table>
                <tr><th>Drive</th><th>Total Size</th><th>Free Space</th><th>Free %</th><th>Status</th></tr>
"@

    foreach ($disk in $DiskInfo) {
        $statusClass = switch ($disk.Status) { "CRITICAL" { "status-bad" } "WARNING" { "status-warn" } default { "status-good" } }
        $html += "<tr><td>$($disk.Drive)</td><td>$($disk.SizeGB) GB</td><td>$($disk.FreeGB) GB</td><td>$($disk.FreePercent)%</td><td><span class='$statusClass'>$($disk.Status)</span></td></tr>"
    }

    $html += @"
            </table>
        </div>

        <h2>Pending Reboot Status</h2>
        <div class="section">
"@

    if ($RebootInfo.Required) {
        $html += "<div class='alert alert-warning'><strong>REBOOT REQUIRED!</strong> Reasons: $($RebootInfo.Reasons -join ', ')</div>"
    } else {
        $html += "<div class='alert alert-success'><strong>No pending reboot detected.</strong></div>"
    }

    $html += @"
        </div>

        <h2>Installed Patches (Recent)</h2>
        <div class="section">
            <table>
                <tr><th>KB/HotFix ID</th><th>Description</th><th>Installed Date</th><th>Installed By</th></tr>
"@

    foreach ($patch in $PatchInfo | Select-Object -First 15) {
        $html += "<tr><td><strong>$($patch.HotFixID)</strong></td><td>$($patch.Description)</td><td>$($patch.InstalledOn)</td><td>$($patch.InstalledBy)</td></tr>"
    }

    $html += @"
            </table>
        </div>

        <h2>Errors Before vs After Patches</h2>
        <div class="section">
"@

    if ($PatchAnalysis -and $PatchAnalysis.Count -gt 0) {
        foreach ($analysis in $PatchAnalysis) {
            $statusClass = switch -Wildcard ($analysis.Status) { 
                "*INCREASED*" { "alert-danger" } 
                "*SOME*" { "alert-warning" } 
                default { "alert-success" } 
            }
            $html += @"
            <div class="patch-item">
                <p class="patch-date">Patch: $($analysis.PatchID) - Installed: $($analysis.PatchDate)</p>
                <p>$($analysis.PatchDescription)</p>
                <table>
                    <tr><th>Period</th><th>Error Count</th><th>Details</th></tr>
                    <tr><td>24 Hours BEFORE Patch</td><td><span class="$(if($analysis.ErrorsBeforeCount -gt 0){'status-warn'}else{'status-good'})">$($analysis.ErrorsBeforeCount) errors</span></td><td>$(if($analysis.ErrorsBefore){($analysis.ErrorsBefore | ForEach-Object { "[$($_.Time)] Event $($_.Id)" }) -join '<br>'}else{'None'})</td></tr>
                    <tr><td>48 Hours AFTER Patch</td><td><span class="$(if($analysis.ErrorsAfterCount -gt 5){'status-bad'}elseif($analysis.ErrorsAfterCount -gt 0){'status-warn'}else{'status-good'})">$($analysis.ErrorsAfterCount) errors</span></td><td>$(if($analysis.ErrorsAfter){($analysis.ErrorsAfter | ForEach-Object { "[$($_.Time)] Event $($_.Id)" }) -join '<br>'}else{'None'})</td></tr>
                </table>
                <div class="$statusClass" style="margin-top:10px; padding:10px; border-radius:5px;"><strong>Analysis:</strong> $($analysis.Status)</div>
            </div>
"@
        }
    } else {
        $html += "<p>No patch analysis data available.</p>"
    }

    $html += @"
        </div>

        <h2>Domain Controller Connectivity</h2>
        <div class="section">
            <table>
                <tr><th>Check</th><th>Status</th><th>Details</th></tr>
                <tr><td>Domain Membership</td><td><span class="$(if($DomainInfo.InDomain){'status-good'}else{'status-bad'})">$(if($DomainInfo.InDomain){'JOINED'}else{'NOT JOINED'})</span></td><td>$($DomainInfo.DomainName)</td></tr>
                <tr><td>Domain Controller</td><td>-</td><td>$($DomainInfo.DomainController) ($($DomainInfo.DCIPAddress))</td></tr>
                <tr><td>Secure Channel (Trust)</td><td><span class="$(if($DomainInfo.SecureChannel){'status-good'}else{'status-bad'})">$(if($DomainInfo.SecureChannel){'VALID'}else{'BROKEN'})</span></td><td>Trust relationship with domain</td></tr>
                <tr><td>LDAP (389)</td><td><span class="$(if($DomainInfo.LDAPPort389){'status-good'}else{'status-bad'})">$(if($DomainInfo.LDAPPort389){'OPEN'}else{'BLOCKED'})</span></td><td>LDAP authentication</td></tr>
                <tr><td>LDAPS (636)</td><td><span class="$(if($DomainInfo.LDAPPort636){'status-good'}else{'status-bad'})">$(if($DomainInfo.LDAPPort636){'OPEN'}else{'BLOCKED'})</span></td><td>Secure LDAP</td></tr>
                <tr><td>Kerberos (88)</td><td><span class="$(if($DomainInfo.KerberosPort88){'status-good'}else{'status-bad'})">$(if($DomainInfo.KerberosPort88){'OPEN'}else{'BLOCKED'})</span></td><td>Kerberos authentication</td></tr>
                <tr><td>DNS (53)</td><td><span class="$(if($DomainInfo.DNSPort53){'status-good'}else{'status-bad'})">$(if($DomainInfo.DNSPort53){'OPEN'}else{'BLOCKED'})</span></td><td>DNS resolution</td></tr>
                <tr><td>SMB (445)</td><td><span class="$(if($DomainInfo.SMBPort445){'status-good'}else{'status-bad'})">$(if($DomainInfo.SMBPort445){'OPEN'}else{'BLOCKED'})</span></td><td>File sharing and GPO</td></tr>
                <tr><td>RPC (135)</td><td><span class="$(if($DomainInfo.RPCPort135){'status-good'}else{'status-bad'})">$(if($DomainInfo.RPCPort135){'OPEN'}else{'BLOCKED'})</span></td><td>RPC endpoint mapper</td></tr>
                <tr><td>Global Catalog (3268)</td><td><span class="$(if($DomainInfo.GCPort3268){'status-good'}else{'status-bad'})">$(if($DomainInfo.GCPort3268){'OPEN'}else{'BLOCKED'})</span></td><td>Global catalog queries</td></tr>
                <tr><td>NLTEST</td><td><span class="$(if($DomainInfo.NLTestSuccess){'status-good'}else{'status-bad'})">$(if($DomainInfo.NLTestSuccess){'PASSED'}else{'FAILED'})</span></td><td>Domain controller discovery</td></tr>
            </table>
        </div>

        <h2>RDP Configuration</h2>
        <div class="section">
            <table>
                <tr><th>Setting</th><th>Value</th><th>Status</th></tr>
                <tr><td>RDP Enabled</td><td>$($RDPInfo.Enabled)</td><td><span class="$(if($RDPInfo.Enabled){'status-good'}else{'status-bad'})">$(if($RDPInfo.Enabled){'ENABLED'}else{'DISABLED'})</span></td></tr>
                <tr><td>RDP Port</td><td>$($RDPInfo.PortNumber)</td><td>-</td></tr>
                <tr><td>NLA (Network Level Auth)</td><td>$($RDPInfo.NLAEnabled)</td><td><span class="$(if($RDPInfo.NLAEnabled){'status-warn'}else{'status-good'})">$(if($RDPInfo.NLAEnabled){'ENABLED - May block access'}else{'DISABLED'})</span></td></tr>
                <tr><td>Port 3389 Accessible</td><td>$($RDPInfo.Port3389Open)</td><td><span class="$(if($RDPInfo.Port3389Open){'status-good'}else{'status-bad'})">$(if($RDPInfo.Port3389Open){'OPEN'}else{'BLOCKED'})</span></td></tr>
                <tr><td>Terminal Services</td><td>$($RDPInfo.ServiceStatus)</td><td><span class="$(if($RDPInfo.ServiceStatus -eq 'Running'){'status-good'}else{'status-bad'})">$($RDPInfo.ServiceStatus)</span></td></tr>
            </table>
        </div>

        <h2>Firewall Status</h2>
        <div class="section">
            <table>
                <tr><th>Profile</th><th>Enabled</th><th>Default Inbound</th><th>Default Outbound</th></tr>
"@

    foreach ($profile in @("Domain", "Private", "Public")) {
        if ($FirewallInfo.$profile) {
            $html += "<tr><td>$profile</td><td><span class='$(if($FirewallInfo.$profile.Enabled){"status-warn"}else{"status-good"})'>$($FirewallInfo.$profile.Enabled)</span></td><td>$($FirewallInfo.$profile.DefaultInbound)</td><td>$($FirewallInfo.$profile.DefaultOutbound)</td></tr>"
        }
    }

    $html += @"
            </table>
        </div>

        <h2>Critical Services</h2>
        <div class="section">
            <table>
                <tr><th>Service</th><th>Display Name</th><th>Status</th><th>Start Type</th></tr>
"@

    foreach ($svc in $ServicesInfo) {
        $statusClass = if($svc.Status -eq "Running"){"status-good"}elseif($svc.Status -eq "Stopped" -and $svc.StartType -eq "Automatic"){"status-bad"}else{"status-warn"}
        $html += "<tr><td>$($svc.Name)</td><td>$($svc.DisplayName)</td><td><span class='$statusClass'>$($svc.Status)</span></td><td>$($svc.StartType)</td></tr>"
    }

    $html += @"
            </table>
        </div>

        <h2>DNS Configuration</h2>
        <div class="section">
            <table>
                <tr><th>Check</th><th>Status</th></tr>
                <tr><td>DNS Servers</td><td>$($DNSInfo.Servers -join ', ')</td></tr>
                <tr><td>Self Resolution</td><td><span class="$(if($DNSInfo.ResolvesSelf){'status-good'}else{'status-bad'})">$(if($DNSInfo.ResolvesSelf){'SUCCESS'}else{'FAILED'})</span></td></tr>
"@

    foreach ($key in $DNSInfo.Keys | Where-Object { $_ -like "DNS_*" }) {
        $server = $key.Replace("DNS_", "")
        $html += "<tr><td>DNS Server $server (Port 53)</td><td><span class='$(if($DNSInfo.$key){"status-good"}else{"status-bad"})'>$(if($DNSInfo.$key){'REACHABLE'}else{'UNREACHABLE'})</span></td></tr>"
    }

    $html += @"
            </table>
        </div>

        <h2>Time Synchronization</h2>
        <div class="section">
            <table>
                <tr><th>Check</th><th>Status</th></tr>
                <tr><td>Time Synchronized</td><td><span class="$(if($TimeInfo.Synchronized){'status-good'}else{'status-bad'})">$(if($TimeInfo.Synchronized){'YES'}else{'NO'})</span></td></tr>
                <tr><td>Time Offset</td><td><span class="$(if($TimeInfo.OffsetOK){'status-good'}else{'status-bad'})">$($TimeInfo.Offset) seconds $(if(-not $TimeInfo.OffsetOK){'(EXCEEDS 5 MIN THRESHOLD)'})</span></td></tr>
            </table>
        </div>

        <h2>Recent Error Events (Last 7 Days)</h2>
        <div class="section">
            <h3>System Log (Critical & Error)</h3>
            <table>
                <tr><th>Time</th><th>Event ID</th><th>Source</th><th>Message</th></tr>
"@

    if ($EventsInfo.System) {
        foreach ($evt in $EventsInfo.System) {
            $html += "<tr><td>$($evt.TimeCreated)</td><td>$($evt.Id)</td><td>$($evt.Source)</td><td>$($evt.Message)</td></tr>"
        }
    } else {
        $html += "<tr><td colspan='4'>No critical errors in last 7 days</td></tr>"
    }

    $html += @"
            </table>
            
            <h3>Security Log (Failed Logins & Account Lockouts)</h3>
            <table>
                <tr><th>Time</th><th>Event ID</th><th>Source</th><th>Message</th></tr>
"@

    if ($EventsInfo.Security) {
        foreach ($evt in $EventsInfo.Security) {
            $html += "<tr><td>$($evt.TimeCreated)</td><td>$($evt.Id)</td><td>$($evt.Source)</td><td>$($evt.Message)</td></tr>"
        }
    } else {
        $html += "<tr><td colspan='4'>No failed login attempts in last 7 days</td></tr>"
    }

    $html += @"
            </table>

            <h3>Application Log (Critical & Error)</h3>
            <table>
                <tr><th>Time</th><th>Event ID</th><th>Source</th><th>Message</th></tr>
"@

    if ($EventsInfo.Application) {
        foreach ($evt in $EventsInfo.Application) {
            $html += "<tr><td>$($evt.TimeCreated)</td><td>$($evt.Id)</td><td>$($evt.Source)</td><td>$($evt.Message)</td></tr>"
        }
    } else {
        $html += "<tr><td colspan='4'>No application errors in last 3 days</td></tr>"
    }

    $html += @"
            </table>
        </div>

        <h2>Recommended Actions</h2>
        <div class="section">
"@

    $issues = @()
    if (-not $DomainInfo.InDomain) { $issues += "<div class='alert alert-danger'><strong>SERVER NOT IN DOMAIN</strong> - Rejoin required</div>" }
    if (-not $DomainInfo.SecureChannel) { $issues += "<div class='alert alert-danger'><strong>TRUST RELATIONSHIP BROKEN</strong> - Run: <code>Test-ComputerSecureChannel -Repair -Credential (Get-Credential)</code></div>" }
    if (-not $RDPInfo.Enabled) { $issues += "<div class='alert alert-danger'><strong>RDP DISABLED</strong> - Enable via registry or GPO</div>" }
    if (-not $RDPInfo.Port3389Open) { $issues += "<div class='alert alert-danger'><strong>RDP PORT BLOCKED</strong> - Check firewall rules and NSG (AWS/Azure)</div>" }
    if ($RDPInfo.NLAEnabled -and -not $DomainInfo.SecureChannel) { $issues += "<div class='alert alert-danger'><strong>NLA ENABLED WITH BROKEN TRUST</strong> - Disable NLA temporarily or fix trust first</div>" }
    if (-not $DomainInfo.LDAPPort389 -or -not $DomainInfo.KerberosPort88) { $issues += "<div class='alert alert-danger'><strong>DC PORTS BLOCKED</strong> - Check firewall for LDAP(389) and Kerberos(88)</div>" }
    if (-not $TimeInfo.OffsetOK) { $issues += "<div class='alert alert-danger'><strong>TIME OUT OF SYNC</strong> - Run: <code>w32tm /resync /force</code></div>" }
    if (-not $DNSInfo.ResolvesSelf) { $issues += "<div class='alert alert-danger'><strong>DNS RESOLUTION FAILED</strong> - Check DNS server configuration</div>" }
    if ($RebootInfo.Required) { $issues += "<div class='alert alert-warning'><strong>REBOOT PENDING</strong> - Schedule a maintenance window for reboot</div>" }
    foreach ($disk in $DiskInfo | Where-Object { $_.Status -eq "CRITICAL" }) { $issues += "<div class='alert alert-danger'><strong>LOW DISK SPACE on $($disk.Drive)</strong> - Only $($disk.FreeGB) GB free ($($disk.FreePercent)%)</div>" }
    if ($SystemInfo.MemoryUsagePercent -gt 90) { $issues += "<div class='alert alert-warning'><strong>HIGH MEMORY USAGE</strong> - $($SystemInfo.MemoryUsagePercent)% used</div>" }

    if ($issues.Count -eq 0) {
        $html += "<div class='alert alert-success'><strong>NO CRITICAL ISSUES DETECTED!</strong> Server appears to be healthy.</div>"
    } else {
        $html += $issues -join "`n"
    }

    $html += @"
        </div>

        <h2>Remediation Commands Reference</h2>
        <div class="section">
            <table>
                <tr><th>Issue</th><th>Command</th></tr>
                <tr><td>Repair Trust Relationship</td><td><code>Test-ComputerSecureChannel -Repair -Credential (Get-Credential)</code></td></tr>
                <tr><td>Reset Computer Password</td><td><code>Reset-ComputerMachinePassword -Credential (Get-Credential)</code></td></tr>
                <tr><td>Force Time Sync</td><td><code>w32tm /resync /force</code></td></tr>
                <tr><td>Restart Netlogon</td><td><code>Restart-Service Netlogon -Force</code></td></tr>
                <tr><td>Flush DNS Cache</td><td><code>ipconfig /flushdns</code></td></tr>
                <tr><td>Register DNS</td><td><code>ipconfig /registerdns</code></td></tr>
                <tr><td>Enable RDP</td><td><code>Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name fDenyTSConnections -Value 0</code></td></tr>
                <tr><td>Disable NLA</td><td><code>Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -Name UserAuthentication -Value 0</code></td></tr>
                <tr><td>Open RDP Firewall</td><td><code>Enable-NetFirewallRule -DisplayGroup "Remote Desktop"</code></td></tr>
                <tr><td>Update Group Policy</td><td><code>gpupdate /force</code></td></tr>
                <tr><td>Check Secure Channel</td><td><code>Test-ComputerSecureChannel -Verbose</code></td></tr>
                <tr><td>Find Domain Controller</td><td><code>nltest /dsgetdc:DOMAIN</code></td></tr>
            </table>
        </div>

        <p class="timestamp">Report generated by L1 Server Diagnostics Script v2.0 | No changes were made to the server</p>
    </div>
</body>
</html>
"@

    return $html
}

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  L1 SERVER DIAGNOSTICS SCRIPT v2.0" -ForegroundColor Cyan
Write-Host "  READ-ONLY - NO CHANGES WILL BE MADE" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

Install-RequiredModules

if ($ConnectAzure -or $ConnectAWS) {
    Write-Host "Connecting to cloud providers..." -ForegroundColor Yellow
    $cloudStatus = Connect-CloudProviders -Azure $ConnectAzure -AWS $ConnectAWS -Auto $AutoConnect
}

Write-Host "Starting diagnostics for $ComputerName..." -ForegroundColor Cyan
Write-Host ""

Write-Host "[1/12] Gathering system information..." -ForegroundColor Gray
$sysInfo = Get-SystemInfo -Computer $ComputerName

Write-Host "[2/12] Checking network configuration..." -ForegroundColor Gray
$netConfig = Get-NetworkConfiguration -Computer $ComputerName

Write-Host "[3/12] Testing domain connectivity..." -ForegroundColor Gray
$domainInfo = Test-DomainConnectivity -Computer $ComputerName

Write-Host "[4/12] Checking firewall status..." -ForegroundColor Gray
$fwInfo = Get-FirewallStatus -Computer $ComputerName

Write-Host "[5/12] Checking RDP configuration..." -ForegroundColor Gray
$rdpInfo = Get-RDPStatus -Computer $ComputerName

Write-Host "[6/12] Checking critical services..." -ForegroundColor Gray
$svcInfo = Get-CriticalServices -Computer $ComputerName

Write-Host "[7/12] Checking DNS configuration..." -ForegroundColor Gray
$dnsInfo = Get-DNSConfiguration -Computer $ComputerName

Write-Host "[8/12] Checking time synchronization..." -ForegroundColor Gray
$timeInfo = Get-TimeConfiguration -Computer $ComputerName

Write-Host "[9/12] Gathering event log errors..." -ForegroundColor Gray
$evtInfo = Get-EventLogErrors -Computer $ComputerName

Write-Host "[10/12] Getting last logged on users..." -ForegroundColor Gray
$loginInfo = Get-LastLoggedOnUsers -Computer $ComputerName

Write-Host "[11/12] Getting patch history..." -ForegroundColor Gray
$patchInfo = Get-PatchHistory -Computer $ComputerName

Write-Host "[12/12] Analyzing errors around patches..." -ForegroundColor Gray
$patchAnalysis = Get-ErrorsAroundPatches -Computer $ComputerName -Patches $patchInfo

$diskInfo = Get-DiskSpace -Computer $ComputerName
$rebootInfo = Get-PendingReboot -Computer $ComputerName

Write-Host ""
Write-Host "Generating HTML report..." -ForegroundColor Yellow

$report = Generate-HTMLReport -SystemInfo $sysInfo -NetworkConfig $netConfig -DomainInfo $domainInfo -FirewallInfo $fwInfo -RDPInfo $rdpInfo -ServicesInfo $svcInfo -DNSInfo $dnsInfo -TimeInfo $timeInfo -EventsInfo $evtInfo -LoginInfo $loginInfo -PatchInfo $patchInfo -PatchAnalysis $patchAnalysis -DiskInfo $diskInfo -RebootInfo $rebootInfo -ComputerName $ComputerName

$report | Out-File -FilePath $OutputPath -Encoding UTF8

Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host "  REPORT COMPLETE!" -ForegroundColor Green
Write-Host "  Saved to: $OutputPath" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Green

Start-Process $OutputPath
