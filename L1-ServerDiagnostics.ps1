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
        $info.LastBoot = $os.ConvertToDateTime($os.LastBootUpTime)
        $info.Uptime = (Get-Date) - $info.LastBoot
    } catch {
        $info.Error = "Unable to retrieve system information"
    }
    return $info
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
            
            try {
                $nltest = nltest /dsgetdc:$($cs.Domain) 2>&1
                $results.NLTestSuccess = $LASTEXITCODE -eq 0
                $results.NLTestOutput = $nltest -join "`n"
            } catch { $results.NLTestSuccess = $false }
            
            try {
                $secureChannel = Test-ComputerSecureChannel -ErrorAction Stop
                $results.SecureChannel = $secureChannel
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
                LogAllowed = $profile.LogAllowed
                LogBlocked = $profile.LogBlocked
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
        $rdp.SecurityLayer = $key2.GetValue('SecurityLayer')
        $rdp.UserAuthentication = $key2.GetValue('UserAuthentication')
        
        $rdp.NLAEnabled = ($rdp.UserAuthentication -eq 1)
        
        $svc = Get-Service -ComputerName $Computer -Name TermService
        $rdp.ServiceStatus = $svc.Status
        $rdp.ServiceStartType = $svc.StartType
        
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
        "Netlogon",
        "NTDS",
        "DNS",
        "W32Time",
        "LanmanServer",
        "LanmanWorkstation",
        "RemoteRegistry",
        "TermService",
        "WinRM",
        "DFSR",
        "Dnscache",
        "Dhcp",
        "NlaSvc",
        "gpsvc"
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

function Get-GPOStatus {
    param([string]$Computer)
    $gpo = @{}
    try {
        $rsop = Get-WmiObject -Namespace "root\rsop\computer" -Class RSOP_GPO -ComputerName $Computer
        $gpo.AppliedGPOs = $rsop | ForEach-Object { $_.Name }
        $gpo.GPOCount = ($rsop | Measure-Object).Count
        
        try {
            $gpresult = gpresult /r /scope:computer 2>&1
            $gpo.LastGPUpdate = ($gpresult | Select-String "Last time Group Policy was applied").ToString()
        } catch { }
        
    } catch {
        $gpo.Error = "Unable to retrieve GPO status"
    }
    return $gpo
}

function Get-DNSConfiguration {
    param([string]$Computer)
    $dns = @{}
    try {
        $adapters = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $Computer | Where-Object { $_.IPEnabled }
        $dns.Servers = ($adapters | ForEach-Object { $_.DNSServerSearchOrder }) | Select-Object -Unique
        $dns.Suffix = ($adapters | ForEach-Object { $_.DNSDomainSuffixSearchOrder }) | Select-Object -Unique
        
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
            StartTime = (Get-Date).AddHours(-24)
        } -MaxEvents 10 | ForEach-Object {
            @{
                TimeCreated = $_.TimeCreated
                Id = $_.Id
                Message = $_.Message.Substring(0, [Math]::Min(200, $_.Message.Length))
            }
        }
    } catch { $events.System = @() }
    
    try {
        $events.Security = Get-WinEvent -ComputerName $Computer -FilterHashtable @{
            LogName = 'Security'
            Id = 4625,4771,4776
            StartTime = (Get-Date).AddHours(-24)
        } -MaxEvents 10 | ForEach-Object {
            @{
                TimeCreated = $_.TimeCreated
                Id = $_.Id
                Message = $_.Message.Substring(0, [Math]::Min(200, $_.Message.Length))
            }
        }
    } catch { $events.Security = @() }
    
    return $events
}

function Get-TrustRelationship {
    param([string]$Computer)
    $trust = @{}
    try {
        $result = Test-ComputerSecureChannel -Verbose 4>&1
        $trust.SecureChannel = Test-ComputerSecureChannel
        $trust.Details = $result -join "`n"
    } catch {
        $trust.SecureChannel = $false
        $trust.Error = $_.Exception.Message
    }
    return $trust
}

function Generate-HTMLReport {
    param(
        [hashtable]$SystemInfo,
        [array]$NetworkConfig,
        [hashtable]$DomainInfo,
        [hashtable]$FirewallInfo,
        [hashtable]$RDPInfo,
        [array]$ServicesInfo,
        [hashtable]$GPOInfo,
        [hashtable]$DNSInfo,
        [hashtable]$TimeInfo,
        [hashtable]$EventsInfo,
        [hashtable]$TrustInfo,
        [string]$ComputerName
    )
    
    $statusGood = "background-color: #d4edda; color: #155724;"
    $statusBad = "background-color: #f8d7da; color: #721c24;"
    $statusWarn = "background-color: #fff3cd; color: #856404;"
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Server Diagnostics Report - $ComputerName</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 15px; }
        h2 { color: #34495e; margin-top: 30px; border-left: 4px solid #3498db; padding-left: 15px; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th { background: #3498db; color: white; padding: 12px; text-align: left; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr:hover { background: #f8f9fa; }
        .status-good { $statusGood padding: 5px 10px; border-radius: 4px; }
        .status-bad { $statusBad padding: 5px 10px; border-radius: 4px; }
        .status-warn { $statusWarn padding: 5px 10px; border-radius: 4px; }
        .section { margin: 20px 0; padding: 20px; background: #fafafa; border-radius: 8px; }
        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin: 20px 0; }
        .summary-card { padding: 20px; border-radius: 8px; text-align: center; }
        .timestamp { color: #7f8c8d; font-size: 0.9em; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Windows Server Diagnostics Report</h1>
        <p class="timestamp">Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") | Target: $ComputerName</p>
        
        <div class="summary">
            <div class="summary-card" style="$(if($DomainInfo.InDomain){$statusGood}else{$statusBad})">
                <h3>Domain Status</h3>
                <p>$(if($DomainInfo.InDomain){"JOINED"}else{"NOT JOINED"})</p>
            </div>
            <div class="summary-card" style="$(if($DomainInfo.SecureChannel){$statusGood}else{$statusBad})">
                <h3>Trust Relationship</h3>
                <p>$(if($DomainInfo.SecureChannel){"VALID"}else{"BROKEN"})</p>
            </div>
            <div class="summary-card" style="$(if($RDPInfo.Enabled -and $RDPInfo.Port3389Open){$statusGood}else{$statusBad})">
                <h3>RDP Status</h3>
                <p>$(if($RDPInfo.Enabled -and $RDPInfo.Port3389Open){"ACCESSIBLE"}else{"BLOCKED"})</p>
            </div>
            <div class="summary-card" style="$(if($TimeInfo.OffsetOK){$statusGood}else{$statusBad})">
                <h3>Time Sync</h3>
                <p>$(if($TimeInfo.OffsetOK){"SYNCHRONIZED"}else{"OUT OF SYNC"})</p>
            </div>
        </div>
        
        <h2>System Information</h2>
        <div class="section">
            <table>
                <tr><th>Property</th><th>Value</th></tr>
                <tr><td>Hostname</td><td>$($SystemInfo.Hostname)</td></tr>
                <tr><td>Domain</td><td>$($SystemInfo.Domain)</td></tr>
                <tr><td>Domain Joined</td><td><span class="$(if($SystemInfo.DomainJoined){'status-good'}else{'status-bad'})">$($SystemInfo.DomainJoined)</span></td></tr>
                <tr><td>Operating System</td><td>$($SystemInfo.OS)</td></tr>
                <tr><td>Version</td><td>$($SystemInfo.Version)</td></tr>
                <tr><td>Last Boot</td><td>$($SystemInfo.LastBoot)</td></tr>
                <tr><td>Uptime</td><td>$($SystemInfo.Uptime.Days) days, $($SystemInfo.Uptime.Hours) hours</td></tr>
            </table>
        </div>
        
        <h2>Network Configuration</h2>
        <div class="section">
            <table>
                <tr><th>Adapter</th><th>IP Address</th><th>Gateway</th><th>DNS Servers</th><th>DHCP</th></tr>
"@

    foreach ($adapter in $NetworkConfig) {
        $html += "<tr><td>$($adapter.Description)</td><td>$($adapter.IPAddress)</td><td>$($adapter.Gateway)</td><td>$($adapter.DNS)</td><td>$($adapter.DHCPEnabled)</td></tr>"
    }
    
    $html += @"
            </table>
        </div>
        
        <h2>Domain Controller Connectivity</h2>
        <div class="section">
            <table>
                <tr><th>Check</th><th>Status</th><th>Details</th></tr>
                <tr><td>Domain Membership</td><td><span class="$(if($DomainInfo.InDomain){'status-good'}else{'status-bad'})">$(if($DomainInfo.InDomain){'JOINED'}else{'NOT JOINED'})</span></td><td>$($DomainInfo.DomainName)</td></tr>
                <tr><td>Domain Controller</td><td>-</td><td>$($DomainInfo.DomainController) ($($DomainInfo.DCIPAddress))</td></tr>
                <tr><td>Secure Channel</td><td><span class="$(if($DomainInfo.SecureChannel){'status-good'}else{'status-bad'})">$(if($DomainInfo.SecureChannel){'VALID'}else{'BROKEN'})</span></td><td>Trust relationship with domain</td></tr>
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
        
        <h2>Recent Error Events (Last 24 Hours)</h2>
        <div class="section">
            <h3>System Log</h3>
            <table>
                <tr><th>Time</th><th>Event ID</th><th>Message</th></tr>
"@

    if ($EventsInfo.System) {
        foreach ($evt in $EventsInfo.System) {
            $html += "<tr><td>$($evt.TimeCreated)</td><td>$($evt.Id)</td><td>$($evt.Message)</td></tr>"
        }
    } else {
        $html += "<tr><td colspan='3'>No critical errors in last 24 hours</td></tr>"
    }
    
    $html += @"
            </table>
            <h3>Security Log (Failed Logins)</h3>
            <table>
                <tr><th>Time</th><th>Event ID</th><th>Message</th></tr>
"@

    if ($EventsInfo.Security) {
        foreach ($evt in $EventsInfo.Security) {
            $html += "<tr><td>$($evt.TimeCreated)</td><td>$($evt.Id)</td><td>$($evt.Message)</td></tr>"
        }
    } else {
        $html += "<tr><td colspan='3'>No failed login attempts in last 24 hours</td></tr>"
    }
    
    $html += @"
            </table>
        </div>
        
        <h2>Recommended Actions</h2>
        <div class="section">
            <ul>
"@

    if (-not $DomainInfo.InDomain) {
        $html += "<li class='status-bad'>SERVER NOT IN DOMAIN - Rejoin required</li>"
    }
    if (-not $DomainInfo.SecureChannel) {
        $html += "<li class='status-bad'>TRUST RELATIONSHIP BROKEN - Run: Test-ComputerSecureChannel -Repair -Credential (Get-Credential)</li>"
    }
    if (-not $RDPInfo.Enabled) {
        $html += "<li class='status-bad'>RDP DISABLED - Enable via registry or GPO</li>"
    }
    if (-not $RDPInfo.Port3389Open) {
        $html += "<li class='status-bad'>RDP PORT BLOCKED - Check firewall rules and NSG (AWS/Azure)</li>"
    }
    if ($RDPInfo.NLAEnabled -and -not $DomainInfo.SecureChannel) {
        $html += "<li class='status-bad'>NLA ENABLED WITH BROKEN TRUST - Disable NLA temporarily or fix trust first</li>"
    }
    if (-not $DomainInfo.LDAPPort389 -or -not $DomainInfo.KerberosPort88) {
        $html += "<li class='status-bad'>DC PORTS BLOCKED - Check firewall for LDAP(389) and Kerberos(88)</li>"
    }
    if (-not $TimeInfo.OffsetOK) {
        $html += "<li class='status-bad'>TIME OUT OF SYNC - Run: w32tm /resync /force</li>"
    }
    if (-not $DNSInfo.ResolvesSelf) {
        $html += "<li class='status-bad'>DNS RESOLUTION FAILED - Check DNS server configuration</li>"
    }
    
    $html += @"
            </ul>
        </div>
        
        <h2>Remediation Commands Reference</h2>
        <div class="section">
            <table>
                <tr><th>Issue</th><th>Command</th></tr>
                <tr><td>Repair Trust Relationship</td><td><code>Test-ComputerSecureChannel -Repair -Credential (Get-Credential)</code></td></tr>
                <tr><td>Rejoin Domain</td><td><code>Reset-ComputerMachinePassword -Credential (Get-Credential)</code></td></tr>
                <tr><td>Force Time Sync</td><td><code>w32tm /resync /force</code></td></tr>
                <tr><td>Restart Netlogon</td><td><code>Restart-Service Netlogon -Force</code></td></tr>
                <tr><td>Flush DNS Cache</td><td><code>ipconfig /flushdns</code></td></tr>
                <tr><td>Register DNS</td><td><code>ipconfig /registerdns</code></td></tr>
                <tr><td>Enable RDP</td><td><code>Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name fDenyTSConnections -Value 0</code></td></tr>
                <tr><td>Disable NLA</td><td><code>Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -Name UserAuthentication -Value 0</code></td></tr>
                <tr><td>Open RDP Firewall</td><td><code>Enable-NetFirewallRule -DisplayGroup "Remote Desktop"</code></td></tr>
                <tr><td>Update Group Policy</td><td><code>gpupdate /force</code></td></tr>
            </table>
        </div>
        
        <p class="timestamp">Report generated by L1 Server Diagnostics Script v1.0</p>
    </div>
</body>
</html>
"@

    return $html
}

Install-RequiredModules

if ($ConnectAzure -or $ConnectAWS) {
    $cloudStatus = Connect-CloudProviders -Azure $ConnectAzure -AWS $ConnectAWS -Auto $AutoConnect
}

Write-Host "Starting diagnostics for $ComputerName..." -ForegroundColor Cyan

$sysInfo = Get-SystemInfo -Computer $ComputerName
$netConfig = Get-NetworkConfiguration -Computer $ComputerName
$domainInfo = Test-DomainConnectivity -Computer $ComputerName
$fwInfo = Get-FirewallStatus -Computer $ComputerName
$rdpInfo = Get-RDPStatus -Computer $ComputerName
$svcInfo = Get-CriticalServices -Computer $ComputerName
$gpoInfo = Get-GPOStatus -Computer $ComputerName
$dnsInfo = Get-DNSConfiguration -Computer $ComputerName
$timeInfo = Get-TimeConfiguration -Computer $ComputerName
$evtInfo = Get-EventLogErrors -Computer $ComputerName
$trustInfo = Get-TrustRelationship -Computer $ComputerName

$report = Generate-HTMLReport -SystemInfo $sysInfo -NetworkConfig $netConfig -DomainInfo $domainInfo -FirewallInfo $fwInfo -RDPInfo $rdpInfo -ServicesInfo $svcInfo -GPOInfo $gpoInfo -DNSInfo $dnsInfo -TimeInfo $timeInfo -EventsInfo $evtInfo -TrustInfo $trustInfo -ComputerName $ComputerName

$report | Out-File -FilePath $OutputPath -Encoding UTF8

Write-Host "Report saved to: $OutputPath" -ForegroundColor Green
Start-Process $OutputPath
