# L1 Windows Server Troubleshooting Guide
## Domain Connectivity & RDP Issues
### Windows Server 2012 R2 - 2022

---

## 1. Initial Assessment

Before troubleshooting, gather the following information from the ticket:

| Information | Details to Collect |
|-------------|-------------------|
| Server Name/IP | Hostname or IP address of affected server |
| Domain Name | Active Directory domain (e.g., corp.company.com) |
| Error Message | Exact error message user is seeing |
| Last Working | When did it last work correctly? |
| Recent Changes | Any recent patches, updates, or changes? |
| Cloud Platform | AWS, Azure, On-Premises, or Hybrid |

---

## 2. Common Error Messages & Quick Fixes

| Error Message | Likely Cause |
|--------------|--------------|
| The trust relationship between this workstation and the primary domain failed | Computer account password out of sync with AD |
| There are currently no logon servers available | Cannot reach Domain Controller - DNS or network issue |
| The remote computer requires Network Level Authentication | NLA enabled but credentials cannot be verified |
| Remote Desktop can't connect to the remote computer | RDP disabled, firewall blocking, or service stopped |
| The user name or password is incorrect | Wrong credentials, locked account, or trust issue |
| An authentication error has occurred | Kerberos/time sync issue or NLA problem |

---

## 3. Step-by-Step Troubleshooting

### Step 3.1: Verify Basic Connectivity

1. Ping the server by IP address
```powershell
ping <server-ip>
```

2. Ping the server by hostname
```powershell
ping <server-hostname>
```

3. **If IP works but hostname fails = DNS issue**

---

### Step 3.2: Check Critical Ports

Test connectivity to required ports using PowerShell:

| Port | Service | Test Command |
|------|---------|--------------|
| 389 | LDAP | `Test-NetConnection <DC-IP> -Port 389` |
| 636 | LDAPS | `Test-NetConnection <DC-IP> -Port 636` |
| 88 | Kerberos | `Test-NetConnection <DC-IP> -Port 88` |
| 53 | DNS | `Test-NetConnection <DC-IP> -Port 53` |
| 445 | SMB | `Test-NetConnection <DC-IP> -Port 445` |
| 135 | RPC | `Test-NetConnection <DC-IP> -Port 135` |
| 3389 | RDP | `Test-NetConnection <Server-IP> -Port 3389` |
| 3268 | Global Catalog | `Test-NetConnection <DC-IP> -Port 3268` |

---

### Step 3.3: Verify Domain Controller Communication

1. Find available Domain Controllers
```powershell
nltest /dsgetdc:<domain-name>
```

2. Verify secure channel status
```powershell
Test-ComputerSecureChannel -Verbose
```

3. Check Netlogon service
```powershell
Get-Service Netlogon
```

---

### Step 3.4: Check DNS Configuration

1. View current DNS servers
```powershell
Get-DnsClientServerAddress
```

2. Test DNS resolution for domain
```powershell
Resolve-DnsName <domain-name>
```

3. Flush DNS cache
```powershell
ipconfig /flushdns
```

4. Register with DNS
```powershell
ipconfig /registerdns
```

---

### Step 3.5: Check Time Synchronization

**Kerberos requires time to be within 5 minutes of the DC.**

1. Check time status
```powershell
w32tm /query /status
```

2. Force time resync
```powershell
w32tm /resync /force
```

---

### Step 3.6: Check Firewall Status

1. View firewall profiles
```powershell
Get-NetFirewallProfile | Select Name, Enabled
```

2. Check RDP firewall rule
```powershell
Get-NetFirewallRule -DisplayGroup "Remote Desktop" | Select DisplayName, Enabled
```

---

### Step 3.7: Check RDP Configuration

1. Verify RDP is enabled
```powershell
(Get-ItemProperty 'HKLM:\System\CurrentControlSet\Control\Terminal Server').fDenyTSConnections
```
**Result: 0 = RDP Enabled, 1 = RDP Disabled**

2. Check NLA (Network Level Authentication) status
```powershell
(Get-ItemProperty 'HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp').UserAuthentication
```
**Result: 1 = NLA Enabled, 0 = NLA Disabled**

3. Check Terminal Services status
```powershell
Get-Service TermService
```

---

## 4. Access Methods When RDP Fails

### 4.1 AWS EC2 Serial Console

1. Go to AWS Console > EC2 > Instances
2. Select the instance
3. Click Connect > EC2 Serial Console tab
4. Click Connect (requires SAC enabled on instance)
5. Type: `cmd` to start command prompt
6. Type: `ch -sn cmd0001` to connect to channel

> **Note:** If SSM Agent is not online, Serial Console is the only option

---

### 4.2 Azure Serial Console

1. Go to Azure Portal > Virtual Machines
2. Select the VM
3. Under Support + troubleshooting, click Serial Console
4. Press Enter to activate SAC prompt
5. Type: `cmd` then `ch -sn cmd0001`

---

### 4.3 PowerShell Remoting

```powershell
Enter-PSSession -ComputerName <server> -Credential (Get-Credential)
```

---

## 5. Resolution Procedures

### 5.1 Repair Trust Relationship

**REQUIRES: Domain Admin credentials**

```powershell
Test-ComputerSecureChannel -Repair -Credential (Get-Credential)
```

Or reset computer password:
```powershell
Reset-ComputerMachinePassword -Credential (Get-Credential)
```

---

### 5.2 Rejoin Domain (If Trust Cannot Be Repaired)

**REQUIRES: Local Admin + Domain Admin credentials**

1. Remove from domain (use local admin)
```powershell
Remove-Computer -UnjoinDomainCredential (Get-Credential) -Force -Restart
```

2. After restart, rejoin domain
```powershell
Add-Computer -DomainName <domain> -Credential (Get-Credential) -Restart
```

---

### 5.3 Enable RDP

```powershell
Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name fDenyTSConnections -Value 0
Enable-NetFirewallRule -DisplayGroup "Remote Desktop"
Restart-Service TermService -Force
```

---

### 5.4 Disable NLA (Temporary)

*Use when trust is broken and NLA prevents login*

```powershell
Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -Name UserAuthentication -Value 0
```

---

### 5.5 Change DNS Server

```powershell
Set-DnsClientServerAddress -InterfaceAlias "Ethernet" -ServerAddresses ("10.0.0.1","10.0.0.2")
```

---

## 6. Escalation Criteria

**Escalate to L2/L3 if:**

- Trust repair fails after 2 attempts
- Domain Controller is unreachable from multiple servers
- AWS/Azure Serial Console is not working
- GPO is blocking changes
- Multiple servers affected simultaneously
- Issue persists after domain rejoin
- Customer production is down

---

## 7. Quick Reference Commands

| Task | Command |
|------|---------|
| Find DC | `nltest /dsgetdc:<domain>` |
| Test Trust | `Test-ComputerSecureChannel -Verbose` |
| Repair Trust | `Test-ComputerSecureChannel -Repair -Credential (Get-Credential)` |
| Reset Password | `Reset-ComputerMachinePassword -Credential (Get-Credential)` |
| Check Time | `w32tm /query /status` |
| Sync Time | `w32tm /resync /force` |
| Flush DNS | `ipconfig /flushdns` |
| Register DNS | `ipconfig /registerdns` |
| Test Port | `Test-NetConnection <IP> -Port <port>` |
| Get DNS Servers | `Get-DnsClientServerAddress` |
| Set DNS | `Set-DnsClientServerAddress -InterfaceAlias "Ethernet" -ServerAddresses "IP"` |
| Check RDP | `(Get-ItemProperty 'HKLM:\...\Terminal Server').fDenyTSConnections` |
| Enable RDP FW | `Enable-NetFirewallRule -DisplayGroup "Remote Desktop"` |
| GPO Update | `gpupdate /force` |
| Remote PS | `Enter-PSSession -ComputerName <server> -Credential (Get-Credential)` |

---

**Document Version:** 1.0 | **Last Updated:** January 2026
