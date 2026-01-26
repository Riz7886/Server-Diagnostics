param(
    [Parameter(Mandatory=$false)]
    [string]$ComputerName = $env:COMPUTERNAME,
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "$env:USERPROFILE\Desktop\L2_ServerDiagnostics_$(Get-Date -Format 'yyyyMMdd_HHmmss').html",
    [Parameter(Mandatory=$false)]
    [switch]$ConnectAzure,
    [Parameter(Mandatory=$false)]
    [switch]$ConnectAWS,
    [Parameter(Mandatory=$false)]
    [switch]$AutoConnect,
    [Parameter(Mandatory=$false)]
    [string]$AWSRegion = "us-east-1",
    [Parameter(Mandatory=$false)]
    [string]$AzureSubscriptionId
)

$ErrorActionPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"

# ============================================
# L2 SERVER DIAGNOSTICS SCRIPT v2.0
# Author: Syed Rizvi
# READ-ONLY - NO CHANGES WILL BE MADE
# ============================================

function Write-StatusResult {
    param([string]$Test, [bool]$Passed, [string]$Details = "")
    $status = if ($Passed) { "PASS" } else { "FAIL" }
    $color = if ($Passed) { "Green" } else { "Red" }
    Write-Host "  [$status] " -ForegroundColor $color -NoNewline
    Write-Host "$Test" -NoNewline
    if ($Details) { Write-Host " - $Details" -ForegroundColor Gray } else { Write-Host "" }
    return @{ Test = $Test; Passed = $Passed; Details = $Details }
}

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
    $results = @{Azure = $false; AWS = $false; AzureError = ""; AWSError = ""}
    if ($Azure) {
        try {
            Import-Module Az.Accounts -ErrorAction SilentlyContinue
            if ($Auto) {
                $context = Get-AzContext
                if ($context) { $results.Azure = $true }
                else { $results.AzureError = "No existing Azure context found" }
            } else {
                Connect-AzAccount -ErrorAction Stop | Out-Null
                $results.Azure = $true
            }
        } catch { $results.AzureError = $_.Exception.Message }
    }
    if ($AWS) {
        try {
            Import-Module AWS.Tools.Common -ErrorAction SilentlyContinue
            Import-Module AWS.Tools.SSM -ErrorAction SilentlyContinue
            Import-Module AWS.Tools.EC2 -ErrorAction SilentlyContinue
            if ($Auto) {
                $creds = Get-AWSCredential -ErrorAction SilentlyContinue
                if ($creds) { $results.AWS = $true }
                else { $results.AWSError = "No AWS credentials configured" }
            } else {
                Set-AWSCredential -ProfileName default -ErrorAction SilentlyContinue
                $results.AWS = $true
            }
        } catch { $results.AWSError = $_.Exception.Message }
    }
    return $results
}

function Test-PortConnectivity {
    param([string]$Target, [int]$Port, [int]$Timeout = 3000)
    $result = @{
        Target = $Target
        Port = $Port
        Open = $false
        Status = "BLOCKED"
        ResponseTime = 0
        Error = ""
    }
    try {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $tcp = New-Object System.Net.Sockets.TcpClient
        $connect = $tcp.BeginConnect($Target, $Port, $null, $null)
        $wait = $connect.AsyncWaitHandle.WaitOne($Timeout, $false)
        $stopwatch.Stop()
        if ($wait) {
            try {
                $tcp.EndConnect($connect)
                $result.Open = $true
                $result.Status = "OPEN"
                $result.ResponseTime = $stopwatch.ElapsedMilliseconds
            } catch {
                $result.Status = "REFUSED"
                $result.Error = "Connection refused"
            }
        } else {
            $result.Status = "TIMEOUT"
            $result.Error = "Connection timed out after ${Timeout}ms"
        }
        $tcp.Close()
    } catch {
        $result.Status = "ERROR"
        $result.Error = $_.Exception.Message
    }
    return $result
}

function Test-AllCriticalPorts {
    param([string]$Target, [string]$DCTarget)
    $results = @()
    
    # RDP Port
    $rdp = Test-PortConnectivity -Target $Target -Port 3389
    $results += @{ Name = "RDP (3389)"; Target = $Target; Port = 3389; Open = $rdp.Open; Status = $rdp.Status; ResponseTime = $rdp.ResponseTime }
    
    # WinRM Ports
    $winrmHttp = Test-PortConnectivity -Target $Target -Port 5985
    $results += @{ Name = "WinRM HTTP (5985)"; Target = $Target; Port = 5985; Open = $winrmHttp.Open; Status = $winrmHttp.Status; ResponseTime = $winrmHttp.ResponseTime }
    
    $winrmHttps = Test-PortConnectivity -Target $Target -Port 5986
    $results += @{ Name = "WinRM HTTPS (5986)"; Target = $Target; Port = 5986; Open = $winrmHttps.Open; Status = $winrmHttps.Status; ResponseTime = $winrmHttps.ResponseTime }
    
    # SMB
    $smb = Test-PortConnectivity -Target $Target -Port 445
    $results += @{ Name = "SMB (445)"; Target = $Target; Port = 445; Open = $smb.Open; Status = $smb.Status; ResponseTime = $smb.ResponseTime }
    
    if ($DCTarget) {
        # Domain Controller Ports
        $ldap = Test-PortConnectivity -Target $DCTarget -Port 389
        $results += @{ Name = "LDAP (389)"; Target = $DCTarget; Port = 389; Open = $ldap.Open; Status = $ldap.Status; ResponseTime = $ldap.ResponseTime }
        
        $ldaps = Test-PortConnectivity -Target $DCTarget -Port 636
        $results += @{ Name = "LDAPS (636)"; Target = $DCTarget; Port = 636; Open = $ldaps.Open; Status = $ldaps.Status; ResponseTime = $ldaps.ResponseTime }
        
        $kerberos = Test-PortConnectivity -Target $DCTarget -Port 88
        $results += @{ Name = "Kerberos (88)"; Target = $DCTarget; Port = 88; Open = $kerberos.Open; Status = $kerberos.Status; ResponseTime = $kerberos.ResponseTime }
        
        $dns = Test-PortConnectivity -Target $DCTarget -Port 53
        $results += @{ Name = "DNS (53)"; Target = $DCTarget; Port = 53; Open = $dns.Open; Status = $dns.Status; ResponseTime = $dns.ResponseTime }
        
        $rpc = Test-PortConnectivity -Target $DCTarget -Port 135
        $results += @{ Name = "RPC (135)"; Target = $DCTarget; Port = 135; Open = $rpc.Open; Status = $rpc.Status; ResponseTime = $rpc.ResponseTime }
        
        $gc = Test-PortConnectivity -Target $DCTarget -Port 3268
        $results += @{ Name = "Global Catalog (3268)"; Target = $DCTarget; Port = 3268; Open = $gc.Open; Status = $gc.Status; ResponseTime = $gc.ResponseTime }
    }
    
    return $results
}

function Get-FirewallRulesAnalysis {
    param([string]$Computer)
    $analysis = @{
        Profiles = @()
        RDPRules = @()
        WinRMRules = @()
        BlockingRules = @()
        Issues = @()
    }
    
    try {
        # Get Firewall Profiles
        $profiles = Get-NetFirewallProfile -CimSession $Computer -ErrorAction Stop
        foreach ($profile in $profiles) {
            $analysis.Profiles += @{
                Name = $profile.Name
                Enabled = $profile.Enabled
                DefaultInbound = $profile.DefaultInboundAction.ToString()
                DefaultOutbound = $profile.DefaultOutboundAction.ToString()
            }
            if ($profile.Enabled -and $profile.DefaultInboundAction -eq "Block") {
                $analysis.Issues += "Firewall profile '$($profile.Name)' is enabled with default BLOCK inbound"
            }
        }
        
        # Check RDP Rules
        $rdpRules = Get-NetFirewallRule -CimSession $Computer -DisplayGroup "Remote Desktop" -ErrorAction SilentlyContinue
        foreach ($rule in $rdpRules) {
            $analysis.RDPRules += @{
                Name = $rule.DisplayName
                Enabled = $rule.Enabled
                Direction = $rule.Direction.ToString()
                Action = $rule.Action.ToString()
                Profile = $rule.Profile.ToString()
            }
        }
        $rdpEnabled = $rdpRules | Where-Object { $_.Enabled -eq $true -and $_.Direction -eq "Inbound" -and $_.Action -eq "Allow" }
        if (-not $rdpEnabled) {
            $analysis.Issues += "No enabled inbound RDP firewall rules found"
        }
        
        # Check WinRM Rules
        $winrmRules = Get-NetFirewallRule -CimSession $Computer -DisplayName "*WinRM*", "*Windows Remote Management*" -ErrorAction SilentlyContinue
        foreach ($rule in $winrmRules) {
            $analysis.WinRMRules += @{
                Name = $rule.DisplayName
                Enabled = $rule.Enabled
                Direction = $rule.Direction.ToString()
                Action = $rule.Action.ToString()
            }
        }
        
        # Check for Blocking Rules
        $blockRules = Get-NetFirewallRule -CimSession $Computer -Enabled True -Direction Inbound -Action Block -ErrorAction SilentlyContinue | Select-Object -First 30
        foreach ($rule in $blockRules) {
            $portFilter = Get-NetFirewallPortFilter -AssociatedNetFirewallRule $rule -CimSession $Computer -ErrorAction SilentlyContinue
            $analysis.BlockingRules += @{
                Name = $rule.DisplayName
                LocalPort = $portFilter.LocalPort
                Protocol = $portFilter.Protocol
                Profile = $rule.Profile.ToString()
            }
        }
        
    } catch {
        $analysis.Issues += "Unable to retrieve firewall rules: $($_.Exception.Message)"
    }
    
    return $analysis
}

function Get-AWSConnectivityInfo {
    param([string]$InstanceId, [string]$Region)
    $awsInfo = @{
        Available = $false
        SSMAgentStatus = "Unknown"
        SSMConnectivity = $false
        InstanceState = "Unknown"
        SecurityGroups = @()
        RDPSecurityGroupCheck = $false
        SSMEndpointReachable = $false
        Issues = @()
        Recommendations = @()
    }
    
    try {
        Import-Module AWS.Tools.EC2 -ErrorAction Stop
        Import-Module AWS.Tools.SSM -ErrorAction Stop
        
        # Check if instance exists and get details
        if ($InstanceId) {
            $instance = Get-EC2Instance -InstanceId $InstanceId -Region $Region -ErrorAction Stop
            if ($instance) {
                $awsInfo.Available = $true
                $awsInfo.InstanceState = $instance.Instances[0].State.Name
                
                # Get Security Groups
                foreach ($sg in $instance.Instances[0].SecurityGroups) {
                    $sgDetails = Get-EC2SecurityGroup -GroupId $sg.GroupId -Region $Region
                    $rdpRule = $sgDetails.IpPermissions | Where-Object { $_.FromPort -le 3389 -and $_.ToPort -ge 3389 }
                    $awsInfo.SecurityGroups += @{
                        GroupId = $sg.GroupId
                        GroupName = $sg.GroupName
                        RDPAllowed = ($null -ne $rdpRule)
                    }
                    if ($rdpRule) { $awsInfo.RDPSecurityGroupCheck = $true }
                }
                
                if (-not $awsInfo.RDPSecurityGroupCheck) {
                    $awsInfo.Issues += "No Security Group allows RDP (port 3389) inbound"
                    $awsInfo.Recommendations += "Add inbound rule for TCP port 3389 to Security Group"
                }
                
                # Check SSM Agent Status
                $ssmInfo = Get-SSMInstanceInformation -InstanceInformationFilterList @{Key="InstanceIds";ValueSet=$InstanceId} -Region $Region -ErrorAction SilentlyContinue
                if ($ssmInfo) {
                    $awsInfo.SSMAgentStatus = $ssmInfo.PingStatus
                    $awsInfo.SSMConnectivity = ($ssmInfo.PingStatus -eq "Online")
                    if (-not $awsInfo.SSMConnectivity) {
                        $awsInfo.Issues += "SSM Agent is not online (Status: $($ssmInfo.PingStatus))"
                        $awsInfo.Recommendations += "Check SSM Agent service on the instance"
                        $awsInfo.Recommendations += "Verify IAM instance profile has SSM permissions"
                        $awsInfo.Recommendations += "Check VPC endpoints or internet connectivity for SSM"
                    }
                } else {
                    $awsInfo.Issues += "Instance not registered with SSM"
                    $awsInfo.Recommendations += "Install/restart SSM Agent on the instance"
                    $awsInfo.Recommendations += "Attach IAM role with AmazonSSMManagedInstanceCore policy"
                }
            }
        }
        
        # Test SSM endpoint reachability
        $ssmEndpoint = "ssm.$Region.amazonaws.com"
        $ssmTest = Test-PortConnectivity -Target $ssmEndpoint -Port 443
        $awsInfo.SSMEndpointReachable = $ssmTest.Open
        if (-not $ssmTest.Open) {
            $awsInfo.Issues += "Cannot reach SSM endpoint ($ssmEndpoint)"
            $awsInfo.Recommendations += "Check VPC endpoint configuration or NAT Gateway"
        }
        
    } catch {
        $awsInfo.Issues += "AWS API Error: $($_.Exception.Message)"
    }
    
    return $awsInfo
}

function Get-AzureConnectivityInfo {
    param([string]$VMName, [string]$ResourceGroup)
    $azureInfo = @{
        Available = $false
        VMState = "Unknown"
        SerialConsoleEnabled = $false
        BootDiagnosticsEnabled = $false
        NSGRules = @()
        RDPNSGCheck = $false
        Issues = @()
        Recommendations = @()
    }
    
    try {
        Import-Module Az.Compute -ErrorAction Stop
        
        if ($VMName -and $ResourceGroup) {
            $vm = Get-AzVM -Name $VMName -ResourceGroupName $ResourceGroup -Status -ErrorAction Stop
            if ($vm) {
                $azureInfo.Available = $true
                $azureInfo.VMState = ($vm.Statuses | Where-Object { $_.Code -like "PowerState/*" }).DisplayStatus
                
                # Check Boot Diagnostics (required for Serial Console)
                $vmConfig = Get-AzVM -Name $VMName -ResourceGroupName $ResourceGroup
                if ($vmConfig.DiagnosticsProfile.BootDiagnostics.Enabled) {
                    $azureInfo.BootDiagnosticsEnabled = $true
                    $azureInfo.SerialConsoleEnabled = $true
                } else {
                    $azureInfo.Issues += "Boot Diagnostics not enabled (required for Serial Console)"
                    $azureInfo.Recommendations += "Enable Boot Diagnostics in VM settings"
                }
                
                # Get NIC and NSG
                foreach ($nic in $vmConfig.NetworkProfile.NetworkInterfaces) {
                    $nicResource = Get-AzNetworkInterface -ResourceId $nic.Id
                    if ($nicResource.NetworkSecurityGroup) {
                        $nsg = Get-AzNetworkSecurityGroup -ResourceGroupName $ResourceGroup -Name $nicResource.NetworkSecurityGroup.Id.Split('/')[-1] -ErrorAction SilentlyContinue
                        if ($nsg) {
                            foreach ($rule in $nsg.SecurityRules) {
                                if ($rule.DestinationPortRange -contains "3389" -or $rule.DestinationPortRange -contains "*") {
                                    if ($rule.Access -eq "Allow" -and $rule.Direction -eq "Inbound") {
                                        $azureInfo.RDPNSGCheck = $true
                                    }
                                }
                                $azureInfo.NSGRules += @{
                                    Name = $rule.Name
                                    Priority = $rule.Priority
                                    Direction = $rule.Direction
                                    Access = $rule.Access
                                    DestinationPort = $rule.DestinationPortRange
                                }
                            }
                        }
                    }
                }
                
                if (-not $azureInfo.RDPNSGCheck) {
                    $azureInfo.Issues += "No NSG rule allows RDP (port 3389) inbound"
                    $azureInfo.Recommendations += "Add inbound NSG rule for TCP port 3389"
                }
            }
        }
        
    } catch {
        $azureInfo.Issues += "Azure API Error: $($_.Exception.Message)"
    }
    
    return $azureInfo
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
        $info.Error = "Unable to retrieve system information: $($_.Exception.Message)"
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
    return $patches | Sort-Object { $_.InstalledOn } -Descending | Select-Object -First 25
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
        $rdp.ServiceStartType = $svc.StartType.ToString()
        $portTest = Test-PortConnectivity -Target $Computer -Port 3389
        $rdp.Port3389Open = $portTest.Open
        $rdp.Port3389Status = $portTest.Status
        $rdp.Port3389ResponseTime = $portTest.ResponseTime
    } catch {
        $rdp.Error = "Unable to retrieve RDP status: $($_.Exception.Message)"
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
        "BITS", "wuauserv", "TrustedInstaller", "AppIDSvc", "AmazonSSMAgent"
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
                    IsRunning = ($svc.Status -eq "Running")
                }
            }
        } catch { }
    }
    return $services
}

function Get-DNSConfiguration {
    param([string]$Computer)
    $dns = @{ Servers = @(); Tests = @() }
    try {
        $adapters = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $Computer | Where-Object { $_.IPEnabled }
        $dns.Servers = ($adapters | ForEach-Object { $_.DNSServerSearchOrder }) | Where-Object { $_ } | Select-Object -Unique
        foreach ($server in $dns.Servers) {
            if ($server) {
                $test = Test-PortConnectivity -Target $server -Port 53
                $dns.Tests += @{
                    Server = $server
                    Port53Open = $test.Open
                    Status = $test.Status
                }
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

function Get-GPOStatus {
    param([string]$Computer)
    $gpo = @{ Applied = @(); NotApplied = @(); LastRefresh = $null; Issues = @() }
    try {
        $gpresult = gpresult /r /scope:computer 2>&1
        $gpo.RawOutput = $gpresult -join "`n"
        
        # Get last GP refresh time
        $gpEvents = Get-WinEvent -ComputerName $Computer -FilterHashtable @{
            LogName = 'Microsoft-Windows-GroupPolicy/Operational'
            Id = 8001
        } -MaxEvents 1 -ErrorAction SilentlyContinue
        if ($gpEvents) {
            $gpo.LastRefresh = $gpEvents[0].TimeCreated
        }
        
        # Check for GP errors
        $gpErrors = Get-WinEvent -ComputerName $Computer -FilterHashtable @{
            LogName = 'Microsoft-Windows-GroupPolicy/Operational'
            Level = 2,3
            StartTime = (Get-Date).AddDays(-1)
        } -MaxEvents 10 -ErrorAction SilentlyContinue
        foreach ($err in $gpErrors) {
            $gpo.Issues += @{
                Time = $err.TimeCreated
                Id = $err.Id
                Message = $err.Message.Substring(0, [Math]::Min(200, $err.Message.Length))
            }
        }
    } catch {
        $gpo.Error = $_.Exception.Message
    }
    return $gpo
}

function Get-KerberosInfo {
    param([string]$Computer)
    $kerberos = @{ Tickets = @(); Issues = @() }
    try {
        $klist = klist 2>&1
        $kerberos.RawOutput = $klist -join "`n"
        $kerberos.HasTickets = ($klist -match "krbtgt")
        
        # Check for Kerberos errors in Security log
        $kerbErrors = Get-WinEvent -ComputerName $Computer -FilterHashtable @{
            LogName = 'Security'
            Id = 4771
            StartTime = (Get-Date).AddDays(-1)
        } -MaxEvents 10 -ErrorAction SilentlyContinue
        foreach ($err in $kerbErrors) {
            $kerberos.Issues += @{
                Time = $err.TimeCreated
                Message = "Kerberos pre-auth failed"
            }
        }
    } catch {
        $kerberos.Error = $_.Exception.Message
    }
    return $kerberos
}

function Get-ADReplicationStatus {
    $repl = @{ Summary = ""; Issues = @() }
    try {
        $replSummary = repadmin /replsummary 2>&1
        $repl.Summary = $replSummary -join "`n"
        $repl.HasFailures = ($replSummary -match "fail")
    } catch {
        $repl.Error = $_.Exception.Message
    }
    return $repl
}

function Generate-HTMLReport {
    param(
        [hashtable]$SystemInfo,
        [array]$NetworkConfig,
        [hashtable]$DomainInfo,
        [hashtable]$FirewallAnalysis,
        [hashtable]$RDPInfo,
        [array]$ServicesInfo,
        [hashtable]$DNSInfo,
        [hashtable]$TimeInfo,
        [hashtable]$EventsInfo,
        [array]$LoginInfo,
        [array]$PatchInfo,
        [array]$DiskInfo,
        [hashtable]$RebootInfo,
        [array]$PortTests,
        [hashtable]$AWSInfo,
        [hashtable]$AzureInfo,
        [hashtable]$GPOInfo,
        [hashtable]$KerberosInfo,
        [string]$ComputerName
    )
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>L2 Server Diagnostics Report - $ComputerName</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 1400px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #5c2d91; border-bottom: 3px solid #5c2d91; padding-bottom: 15px; }
        h2 { color: #5c2d91; margin-top: 30px; border-left: 4px solid #5c2d91; padding-left: 15px; background: #f8f9fa; padding: 10px 15px; border-radius: 0 5px 5px 0; }
        h3 { color: #34495e; margin-top: 20px; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th { background: linear-gradient(135deg, #5c2d91 0%, #7b4bab 100%); color: white; padding: 12px; text-align: left; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr:nth-child(even) { background: #f8f9fa; }
        tr:hover { background: #e8f4fc; }
        .status-pass { background: #d4edda; color: #155724; padding: 5px 10px; border-radius: 4px; font-weight: bold; }
        .status-fail { background: #f8d7da; color: #721c24; padding: 5px 10px; border-radius: 4px; font-weight: bold; }
        .status-warn { background: #fff3cd; color: #856404; padding: 5px 10px; border-radius: 4px; font-weight: bold; }
        .status-open { background: #d4edda; color: #155724; padding: 3px 8px; border-radius: 4px; }
        .status-blocked { background: #f8d7da; color: #721c24; padding: 3px 8px; border-radius: 4px; }
        .section { margin: 20px 0; padding: 20px; background: #fafafa; border-radius: 8px; border: 1px solid #e0e0e0; }
        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 15px; margin: 20px 0; }
        .summary-card { padding: 20px; border-radius: 8px; text-align: center; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
        .summary-card h3 { margin: 0 0 10px 0; font-size: 14px; }
        .summary-card p { margin: 0; font-size: 18px; font-weight: bold; }
        .timestamp { color: #7f8c8d; font-size: 0.9em; }
        .l2-badge { background: #5c2d91; color: white; padding: 3px 8px; border-radius: 4px; font-size: 0.8em; margin-left: 10px; }
        code { background: #1e1e1e; color: #d4d4d4; padding: 3px 8px; border-radius: 4px; font-family: Consolas, monospace; display: inline-block; margin: 2px 0; }
        .alert { padding: 15px; border-radius: 8px; margin: 15px 0; border-left: 4px solid; }
        .alert-danger { background: #f8d7da; border-color: #dc3545; color: #721c24; }
        .alert-warning { background: #fff3cd; border-color: #ffc107; color: #856404; }
        .alert-success { background: #d4edda; border-color: #28a745; color: #155724; }
        .alert-info { background: #e8daef; border-color: #5c2d91; color: #5c2d91; }
        .cloud-section { background: #e8f4fc; border: 1px solid #b8daff; border-radius: 8px; padding: 15px; margin: 10px 0; }
        .port-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px; }
        .port-item { padding: 10px; border-radius: 5px; border: 1px solid #ddd; }
    </style>
</head>
<body>
    <div class="container">
        <h1>🖥️ L2 Windows Server Diagnostics Report <span class="l2-badge">LEVEL 2</span></h1>
        <p class="timestamp"><strong>Generated:</strong> $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") | <strong>Target:</strong> $ComputerName | <strong>Script Version:</strong> 2.0 | <strong>Author:</strong> Syed Rizvi</p>
        
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
                <p>$(if($TimeInfo.OffsetOK){'OK'}else{'OUT OF SYNC'})</p>
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

        <h2>🔌 Port Connectivity Tests</h2>
        <div class="section">
            <div class="port-grid">
"@
    
    foreach ($port in $PortTests) {
        $statusClass = if ($port.Open) { "status-open" } else { "status-blocked" }
        $statusText = if ($port.Open) { "OPEN" } else { $port.Status }
        $html += @"
                <div class="port-item">
                    <strong>$($port.Name)</strong><br/>
                    Target: $($port.Target)<br/>
                    Status: <span class="$statusClass">$statusText</span>
                    $(if($port.Open){"<br/>Response: $($port.ResponseTime)ms"})
                </div>
"@
    }
    
    $html += @"
            </div>
        </div>

        <h2>🔥 Firewall Analysis</h2>
        <div class="section">
            <h3>Firewall Profiles</h3>
            <table>
                <tr><th>Profile</th><th>Enabled</th><th>Default Inbound</th><th>Default Outbound</th></tr>
"@
    
    foreach ($profile in $FirewallAnalysis.Profiles) {
        $html += "<tr><td>$($profile.Name)</td><td><span class='$(if($profile.Enabled){"status-pass"}else{"status-warn"})'>$($profile.Enabled)</span></td><td>$($profile.DefaultInbound)</td><td>$($profile.DefaultOutbound)</td></tr>"
    }
    
    $html += @"
            </table>
            
            <h3>RDP Firewall Rules</h3>
            <table>
                <tr><th>Rule Name</th><th>Enabled</th><th>Direction</th><th>Action</th></tr>
"@
    
    foreach ($rule in $FirewallAnalysis.RDPRules) {
        $html += "<tr><td>$($rule.Name)</td><td><span class='$(if($rule.Enabled){"status-pass"}else{"status-fail"})'>$($rule.Enabled)</span></td><td>$($rule.Direction)</td><td>$($rule.Action)</td></tr>"
    }
    
    if ($FirewallAnalysis.Issues.Count -gt 0) {
        $html += "</table><h3>⚠️ Firewall Issues Detected</h3><ul>"
        foreach ($issue in $FirewallAnalysis.Issues) {
            $html += "<li class='alert alert-warning'>$issue</li>"
        }
        $html += "</ul>"
    } else {
        $html += "</table>"
    }
    
    $html += @"
        </div>

        <h2>☁️ Cloud Connectivity (AWS)</h2>
        <div class="cloud-section">
"@
    
    if ($AWSInfo.Available) {
        $html += @"
            <table>
                <tr><td><strong>Instance State</strong></td><td>$($AWSInfo.InstanceState)</td></tr>
                <tr><td><strong>SSM Agent Status</strong></td><td><span class='$(if($AWSInfo.SSMConnectivity){"status-pass"}else{"status-fail"})'>$($AWSInfo.SSMAgentStatus)</span></td></tr>
                <tr><td><strong>SSM Endpoint Reachable</strong></td><td><span class='$(if($AWSInfo.SSMEndpointReachable){"status-pass"}else{"status-fail"})'>$($AWSInfo.SSMEndpointReachable)</span></td></tr>
                <tr><td><strong>RDP in Security Group</strong></td><td><span class='$(if($AWSInfo.RDPSecurityGroupCheck){"status-pass"}else{"status-fail"})'>$($AWSInfo.RDPSecurityGroupCheck)</span></td></tr>
            </table>
"@
        if ($AWSInfo.Issues.Count -gt 0) {
            $html += "<h4>Issues:</h4><ul>"
            foreach ($issue in $AWSInfo.Issues) { $html += "<li>$issue</li>" }
            $html += "</ul>"
        }
        if ($AWSInfo.Recommendations.Count -gt 0) {
            $html += "<h4>Recommendations:</h4><ul>"
            foreach ($rec in $AWSInfo.Recommendations) { $html += "<li>$rec</li>" }
            $html += "</ul>"
        }
    } else {
        $html += "<p>AWS connectivity check not performed or credentials not available.</p>"
    }
    
    $html += @"
            <h4>AWS SSM Session Manager - How to Connect:</h4>
            <ol>
                <li>Go to <strong>AWS Console → Systems Manager → Session Manager</strong></li>
                <li>Click <strong>Start Session</strong></li>
                <li>Select your instance and click <strong>Start Session</strong></li>
            </ol>
            <h4>AWS SSM Troubleshooting:</h4>
            <ul>
                <li>Verify SSM Agent is running: <code>Get-Service AmazonSSMAgent</code></li>
                <li>Check IAM role has <code>AmazonSSMManagedInstanceCore</code> policy</li>
                <li>Verify VPC has SSM endpoints or NAT Gateway for internet access</li>
                <li>Check Security Group allows HTTPS (443) outbound</li>
            </ul>
        </div>

        <h2>☁️ Cloud Connectivity (Azure)</h2>
        <div class="cloud-section">
"@
    
    if ($AzureInfo.Available) {
        $html += @"
            <table>
                <tr><td><strong>VM State</strong></td><td>$($AzureInfo.VMState)</td></tr>
                <tr><td><strong>Boot Diagnostics</strong></td><td><span class='$(if($AzureInfo.BootDiagnosticsEnabled){"status-pass"}else{"status-fail"})'>$($AzureInfo.BootDiagnosticsEnabled)</span></td></tr>
                <tr><td><strong>Serial Console Available</strong></td><td><span class='$(if($AzureInfo.SerialConsoleEnabled){"status-pass"}else{"status-fail"})'>$($AzureInfo.SerialConsoleEnabled)</span></td></tr>
                <tr><td><strong>RDP in NSG</strong></td><td><span class='$(if($AzureInfo.RDPNSGCheck){"status-pass"}else{"status-fail"})'>$($AzureInfo.RDPNSGCheck)</span></td></tr>
            </table>
"@
        if ($AzureInfo.Issues.Count -gt 0) {
            $html += "<h4>Issues:</h4><ul>"
            foreach ($issue in $AzureInfo.Issues) { $html += "<li>$issue</li>" }
            $html += "</ul>"
        }
    } else {
        $html += "<p>Azure connectivity check not performed or credentials not available.</p>"
    }
    
    $html += @"
            <h4>Azure Serial Console - How to Connect:</h4>
            <ol>
                <li>Go to <strong>Azure Portal → Virtual Machines</strong></li>
                <li>Select your VM</li>
                <li>Under <strong>Help → Serial Console</strong></li>
                <li>Press <strong>Enter</strong> to activate SAC prompt</li>
                <li>Type: <code>cmd</code> to create command channel</li>
                <li>Type: <code>ch -sn Cmd0001</code> to connect</li>
                <li>Login with local admin credentials</li>
            </ol>
            <h4>Azure Serial Console Troubleshooting:</h4>
            <ul>
                <li><strong>Prerequisites:</strong> Boot Diagnostics must be enabled</li>
                <li><strong>Error "Serial console not available":</strong> Enable Boot Diagnostics in VM settings</li>
                <li><strong>Cannot login:</strong> Use local admin account (domain accounts may not work if DC unreachable)</li>
                <li>NSG must allow port 3389 for RDP, but Serial Console works independently</li>
            </ul>
        </div>

        <h2>System Information</h2>
        <div class="section">
            <table>
                <tr><th>Property</th><th>Value</th></tr>
                <tr><td>Hostname</td><td><strong>$($SystemInfo.Hostname)</strong></td></tr>
                <tr><td>Domain</td><td>$($SystemInfo.Domain)</td></tr>
                <tr><td>Domain Joined</td><td><span class="$(if($SystemInfo.DomainJoined){'status-pass'}else{'status-fail'})">$($SystemInfo.DomainJoined)</span></td></tr>
                <tr><td>Operating System</td><td>$($SystemInfo.OS)</td></tr>
                <tr><td>Version / Build</td><td>$($SystemInfo.Version) (Build $($SystemInfo.BuildNumber))</td></tr>
                <tr><td>Last Boot Time</td><td><strong>$($SystemInfo.LastBoot)</strong></td></tr>
                <tr><td>Uptime</td><td>$($SystemInfo.Uptime.Days) days, $($SystemInfo.Uptime.Hours) hours</td></tr>
                <tr><td>Memory Usage</td><td><span class="$(if($SystemInfo.MemoryUsagePercent -gt 90){'status-fail'}elseif($SystemInfo.MemoryUsagePercent -gt 80){'status-warn'}else{'status-pass'})">$($SystemInfo.MemoryUsagePercent)%</span> ($($SystemInfo.FreeMemoryGB) GB free of $($SystemInfo.TotalMemoryGB) GB)</td></tr>
            </table>
        </div>

        <h2>Domain & Trust Status</h2>
        <div class="section">
            <table>
                <tr><td>In Domain</td><td><span class="$(if($DomainInfo.InDomain){'status-pass'}else{'status-fail'})">$($DomainInfo.InDomain)</span></td></tr>
                <tr><td>Domain Name</td><td>$($DomainInfo.DomainName)</td></tr>
                <tr><td>Domain Controller</td><td>$($DomainInfo.DomainController)</td></tr>
                <tr><td>DC IP Address</td><td>$($DomainInfo.DCIPAddress)</td></tr>
                <tr><td>Secure Channel (Trust)</td><td><span class="$(if($DomainInfo.SecureChannel){'status-pass'}else{'status-fail'})">$(if($DomainInfo.SecureChannel){'VALID'}else{'BROKEN'})</span></td></tr>
                <tr><td>NLTest Success</td><td><span class="$(if($DomainInfo.NLTestSuccess){'status-pass'}else{'status-fail'})">$($DomainInfo.NLTestSuccess)</span></td></tr>
            </table>
        </div>

        <h2>RDP Configuration</h2>
        <div class="section">
            <table>
                <tr><td>RDP Enabled</td><td><span class="$(if($RDPInfo.Enabled){'status-pass'}else{'status-fail'})">$($RDPInfo.Enabled)</span></td></tr>
                <tr><td>RDP Port</td><td>$($RDPInfo.PortNumber)</td></tr>
                <tr><td>Port 3389 Status</td><td><span class="$(if($RDPInfo.Port3389Open){'status-pass'}else{'status-fail'})">$($RDPInfo.Port3389Status)</span> $(if($RDPInfo.Port3389Open){"($($RDPInfo.Port3389ResponseTime)ms)"})</td></tr>
                <tr><td>NLA Enabled</td><td><span class="$(if($RDPInfo.NLAEnabled){'status-warn'}else{'status-pass'})">$($RDPInfo.NLAEnabled)</span> $(if($RDPInfo.NLAEnabled -and -not $DomainInfo.SecureChannel){"<strong style='color:red;'>⚠️ NLA + Broken Trust = Cannot Connect!</strong>"})</td></tr>
                <tr><td>TermService Status</td><td><span class="$(if($RDPInfo.ServiceStatus -eq 'Running'){'status-pass'}else{'status-fail'})">$($RDPInfo.ServiceStatus)</span></td></tr>
            </table>
        </div>

        <h2>Disk Space</h2>
        <div class="section">
            <table>
                <tr><th>Drive</th><th>Size (GB)</th><th>Free (GB)</th><th>Free %</th><th>Status</th></tr>
"@
    
    foreach ($disk in $DiskInfo) {
        $statusClass = switch ($disk.Status) { "CRITICAL" { "status-fail" } "WARNING" { "status-warn" } default { "status-pass" } }
        $html += "<tr><td>$($disk.Drive)</td><td>$($disk.SizeGB)</td><td>$($disk.FreeGB)</td><td>$($disk.FreePercent)%</td><td><span class='$statusClass'>$($disk.Status)</span></td></tr>"
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
        $statusClass = if ($svc.IsRunning) { "status-pass" } else { "status-fail" }
        $html += "<tr><td>$($svc.Name)</td><td>$($svc.DisplayName)</td><td><span class='$statusClass'>$($svc.Status)</span></td><td>$($svc.StartType)</td></tr>"
    }
    
    $html += @"
            </table>
        </div>

        <h2>Time Synchronization</h2>
        <div class="section">
            <table>
                <tr><td>Synchronized</td><td><span class="$(if($TimeInfo.Synchronized){'status-pass'}else{'status-fail'})">$($TimeInfo.Synchronized)</span></td></tr>
                <tr><td>Time Offset</td><td><span class="$(if($TimeInfo.OffsetOK){'status-pass'}else{'status-fail'})">$($TimeInfo.Offset) seconds $(if(-not $TimeInfo.OffsetOK){'⚠️ EXCEEDS 5 MIN - KERBEROS WILL FAIL'})</span></td></tr>
            </table>
        </div>

        <h2>Recommended Actions</h2>
        <div class="section">
"@
    
    $issues = @()
    if (-not $DomainInfo.InDomain) { $issues += "<div class='alert alert-danger'><strong>SERVER NOT IN DOMAIN</strong> - Rejoin required</div>" }
    if (-not $DomainInfo.SecureChannel) { $issues += "<div class='alert alert-danger'><strong>TRUST RELATIONSHIP BROKEN</strong> - Run: <code>Test-ComputerSecureChannel -Repair -Credential (Get-Credential)</code></div>" }
    if (-not $RDPInfo.Enabled) { $issues += "<div class='alert alert-danger'><strong>RDP DISABLED</strong> - Run: <code>Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name fDenyTSConnections -Value 0</code></div>" }
    if (-not $RDPInfo.Port3389Open) { $issues += "<div class='alert alert-danger'><strong>RDP PORT BLOCKED</strong> - Check firewall: <code>Enable-NetFirewallRule -DisplayGroup 'Remote Desktop'</code> and check AWS Security Groups / Azure NSG</div>" }
    if ($RDPInfo.NLAEnabled -and -not $DomainInfo.SecureChannel) { $issues += "<div class='alert alert-danger'><strong>NLA ENABLED WITH BROKEN TRUST</strong> - Disable NLA: <code>Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -Name UserAuthentication -Value 0</code></div>" }
    if (-not $TimeInfo.OffsetOK) { $issues += "<div class='alert alert-danger'><strong>TIME OUT OF SYNC</strong> - Run: <code>w32tm /resync /force</code></div>" }
    if ($RebootInfo.Required) { $issues += "<div class='alert alert-warning'><strong>REBOOT PENDING</strong> - Reasons: $($RebootInfo.Reasons -join ', ')</div>" }
    foreach ($disk in $DiskInfo | Where-Object { $_.Status -eq "CRITICAL" }) { $issues += "<div class='alert alert-danger'><strong>LOW DISK SPACE on $($disk.Drive)</strong> - Only $($disk.FreeGB) GB free ($($disk.FreePercent)%)</div>" }

    if ($issues.Count -eq 0) {
        $html += "<div class='alert alert-success'><strong>NO CRITICAL ISSUES DETECTED!</strong> Server appears to be healthy.</div>"
    } else {
        $html += $issues -join "`n"
    }

    $html += @"
        </div>

        <h2>L2 Quick Reference Commands</h2>
        <div class="section">
            <table>
                <tr><th>Task</th><th>Command</th></tr>
                <tr><td>Test Trust</td><td><code>Test-ComputerSecureChannel -Verbose</code></td></tr>
                <tr><td>Repair Trust</td><td><code>Test-ComputerSecureChannel -Repair -Credential (Get-Credential)</code></td></tr>
                <tr><td>Reset Computer Password</td><td><code>Reset-ComputerMachinePassword -Credential (Get-Credential)</code></td></tr>
                <tr><td>Force Time Sync</td><td><code>w32tm /resync /force</code></td></tr>
                <tr><td>Check Time Status</td><td><code>w32tm /query /status</code></td></tr>
                <tr><td>Enable RDP</td><td><code>Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name fDenyTSConnections -Value 0</code></td></tr>
                <tr><td>Disable NLA</td><td><code>Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -Name UserAuthentication -Value 0</code></td></tr>
                <tr><td>Enable RDP Firewall</td><td><code>Enable-NetFirewallRule -DisplayGroup "Remote Desktop"</code></td></tr>
                <tr><td>Restart RDP Service</td><td><code>Restart-Service TermService -Force</code></td></tr>
                <tr><td>Find DC</td><td><code>nltest /dsgetdc:DOMAIN</code></td></tr>
                <tr><td>Test Port</td><td><code>Test-NetConnection -ComputerName SERVER -Port 3389</code></td></tr>
                <tr><td>GPO Report</td><td><code>gpresult /h C:\temp\gpo.html /f</code></td></tr>
                <tr><td>Force GPO Update</td><td><code>gpupdate /force /boot</code></td></tr>
                <tr><td>AD Replication</td><td><code>repadmin /replsummary</code></td></tr>
                <tr><td>DC Diagnostics</td><td><code>dcdiag /v /c /e</code></td></tr>
                <tr><td>Kerberos Tickets</td><td><code>klist</code></td></tr>
                <tr><td>Purge Tickets</td><td><code>klist purge</code></td></tr>
                <tr><td>Check SPNs</td><td><code>setspn -L computername</code></td></tr>
                <tr><td>Flush DNS</td><td><code>ipconfig /flushdns</code></td></tr>
                <tr><td>Register DNS</td><td><code>ipconfig /registerdns</code></td></tr>
            </table>
        </div>

        <p class="timestamp">Report generated by L2 Server Diagnostics Script v2.0 | Author: Syed Rizvi | No changes were made to the server</p>
    </div>
</body>
</html>
"@

    return $html
}

# ============================================
# MAIN EXECUTION
# ============================================

Write-Host ""
Write-Host "=============================================" -ForegroundColor Magenta
Write-Host "  L2 SERVER DIAGNOSTICS SCRIPT v2.0" -ForegroundColor Magenta
Write-Host "  Author: Syed Rizvi" -ForegroundColor Magenta
Write-Host "  READ-ONLY - NO CHANGES WILL BE MADE" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Magenta
Write-Host ""

# Install required modules if needed
Write-Host "Checking required modules..." -ForegroundColor Yellow
Install-RequiredModules

# Connect to cloud providers if requested
$cloudStatus = @{Azure = $false; AWS = $false}
if ($ConnectAzure -or $ConnectAWS) {
    Write-Host "Connecting to cloud providers..." -ForegroundColor Yellow
    $cloudStatus = Connect-CloudProviders -Azure $ConnectAzure -AWS $ConnectAWS -Auto $AutoConnect
    if ($cloudStatus.Azure) { Write-Host "  [CONNECTED] Azure" -ForegroundColor Green }
    if ($cloudStatus.AWS) { Write-Host "  [CONNECTED] AWS" -ForegroundColor Green }
}

Write-Host ""
Write-Host "Starting diagnostics for $ComputerName..." -ForegroundColor Cyan
Write-Host ""

# Run all diagnostic checks
Write-Host "[1/15] Gathering system information..." -ForegroundColor Gray
$sysInfo = Get-SystemInfo -Computer $ComputerName

Write-Host "[2/15] Checking network configuration..." -ForegroundColor Gray
$netConfig = Get-NetworkConfiguration -Computer $ComputerName

Write-Host "[3/15] Testing domain connectivity..." -ForegroundColor Gray
$domainInfo = Test-DomainConnectivity -Computer $ComputerName

Write-Host "[4/15] Testing port connectivity..." -ForegroundColor Gray
$dcIP = if ($domainInfo.DCIPAddress -and $domainInfo.DCIPAddress -ne "N/A") { $domainInfo.DCIPAddress } else { $null }
$portTests = Test-AllCriticalPorts -Target $ComputerName -DCTarget $dcIP
foreach ($test in $portTests) {
    Write-StatusResult -Test "$($test.Name) to $($test.Target)" -Passed $test.Open -Details $test.Status | Out-Null
}

Write-Host "[5/15] Analyzing firewall rules..." -ForegroundColor Gray
$fwAnalysis = Get-FirewallRulesAnalysis -Computer $ComputerName

Write-Host "[6/15] Checking RDP configuration..." -ForegroundColor Gray
$rdpInfo = Get-RDPStatus -Computer $ComputerName
Write-StatusResult -Test "RDP Enabled" -Passed $rdpInfo.Enabled | Out-Null
Write-StatusResult -Test "RDP Port Open" -Passed $rdpInfo.Port3389Open -Details $rdpInfo.Port3389Status | Out-Null

Write-Host "[7/15] Checking critical services..." -ForegroundColor Gray
$svcInfo = Get-CriticalServices -Computer $ComputerName
$stoppedServices = $svcInfo | Where-Object { -not $_.IsRunning -and $_.StartType -eq "Automatic" }
if ($stoppedServices) {
    foreach ($svc in $stoppedServices) {
        Write-StatusResult -Test "Service: $($svc.Name)" -Passed $false -Details "Not Running" | Out-Null
    }
}

Write-Host "[8/15] Checking DNS configuration..." -ForegroundColor Gray
$dnsInfo = Get-DNSConfiguration -Computer $ComputerName

Write-Host "[9/15] Checking time synchronization..." -ForegroundColor Gray
$timeInfo = Get-TimeConfiguration -Computer $ComputerName
Write-StatusResult -Test "Time Sync" -Passed $timeInfo.OffsetOK -Details "$($timeInfo.Offset) seconds offset" | Out-Null

Write-Host "[10/15] Gathering event log errors..." -ForegroundColor Gray
$evtInfo = Get-EventLogErrors -Computer $ComputerName

Write-Host "[11/15] Getting last logged on users..." -ForegroundColor Gray
$loginInfo = Get-LastLoggedOnUsers -Computer $ComputerName

Write-Host "[12/15] Getting patch history..." -ForegroundColor Gray
$patchInfo = Get-PatchHistory -Computer $ComputerName

Write-Host "[13/15] Checking disk space..." -ForegroundColor Gray
$diskInfo = Get-DiskSpace -Computer $ComputerName
foreach ($disk in $diskInfo) {
    Write-StatusResult -Test "Disk $($disk.Drive)" -Passed ($disk.Status -eq "OK") -Details "$($disk.FreePercent)% free" | Out-Null
}

Write-Host "[14/15] Checking pending reboot..." -ForegroundColor Gray
$rebootInfo = Get-PendingReboot -Computer $ComputerName

Write-Host "[15/15] Checking GPO and Kerberos status..." -ForegroundColor Gray
$gpoInfo = Get-GPOStatus -Computer $ComputerName
$kerbInfo = Get-KerberosInfo -Computer $ComputerName

# AWS/Azure checks
$awsInfo = @{ Available = $false }
$azureInfo = @{ Available = $false }
if ($cloudStatus.AWS) {
    Write-Host "Checking AWS connectivity..." -ForegroundColor Yellow
    # Note: Would need instance ID to be provided or discovered
    $awsInfo = Get-AWSConnectivityInfo -Region $AWSRegion
}
if ($cloudStatus.Azure) {
    Write-Host "Checking Azure connectivity..." -ForegroundColor Yellow
    $azureInfo = Get-AzureConnectivityInfo
}

Write-Host ""
Write-Host "Generating HTML report..." -ForegroundColor Yellow

$report = Generate-HTMLReport -SystemInfo $sysInfo -NetworkConfig $netConfig -DomainInfo $domainInfo -FirewallAnalysis $fwAnalysis -RDPInfo $rdpInfo -ServicesInfo $svcInfo -DNSInfo $dnsInfo -TimeInfo $timeInfo -EventsInfo $evtInfo -LoginInfo $loginInfo -PatchInfo $patchInfo -DiskInfo $diskInfo -RebootInfo $rebootInfo -PortTests $portTests -AWSInfo $awsInfo -AzureInfo $azureInfo -GPOInfo $gpoInfo -KerberosInfo $kerbInfo -ComputerName $ComputerName

$report | Out-File -FilePath $OutputPath -Encoding UTF8

Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host "  REPORT COMPLETE!" -ForegroundColor Green
Write-Host "  Saved to: $OutputPath" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Green
Write-Host ""

# Summary output
Write-Host "SUMMARY:" -ForegroundColor White
Write-StatusResult -Test "Domain Joined" -Passed $domainInfo.InDomain
Write-StatusResult -Test "Trust Relationship" -Passed $domainInfo.SecureChannel
Write-StatusResult -Test "RDP Accessible" -Passed ($rdpInfo.Enabled -and $rdpInfo.Port3389Open)
Write-StatusResult -Test "Time Synchronized" -Passed $timeInfo.OffsetOK
Write-StatusResult -Test "No Pending Reboot" -Passed (-not $rebootInfo.Required)

Write-Host ""
Start-Process $OutputPath
