# 🖥️ L1 Windows Server Troubleshooting Guide

**Domain Connectivity & RDP Issues | Windows Server 2012 R2 - 2022**

**Version 2.0 - January 2026**

---

## 📋 Table of Contents

1. [Initial Assessment](#1-initial-assessment)
2. [Common Error Messages](#2-common-error-messages--quick-fixes)
3. [Step-by-Step Troubleshooting](#3-step-by-step-troubleshooting)
4. [Check Last Logged On Users](#4-check-last-logged-on-users) 🆕
5. [Check Patch History & Errors](#5-check-patch-history--errors) 🆕
6. [Check Disk Space & Memory](#6-check-disk-space--memory) 🆕
7. [Check Pending Reboot](#7-check-pending-reboot) 🆕
8. [Access Methods When RDP Fails](#8-access-methods-when-rdp-fails)
9. [Resolution Procedures](#9-resolution-procedures)
10. [Rejoin Server to Domain](#10-rejoin-server-to-domain) 🆕
11. [Escalation Criteria](#11-escalation-criteria)
12. [Quick Reference Commands](#12-quick-reference-commands)

---

## 1. Initial Assessment

Before troubleshooting, gather the following information from the ticket:

| Information | Details to Collect |
|-------------|-------------------|
| **Server Name/IP** | Hostname or IP address of affected server |
| **Domain Name** | Active Directory domain (e.g., corp.company.com) |
| **Error Message** | Exact error message user is seeing |
| **Last Working** | When did it last work correctly? |
| **Recent Changes** | Any recent patches, updates, or changes? |
| **Cloud Platform** | AWS, Azure, On-Premises, or Hybrid |
| **Last Patch Date** | When was the last Windows Update installed? |

---

## 2. Common Error Messages & Quick Fixes

| Error Message | Likely Cause |
|---------------|--------------|
| ❌ **The trust relationship between this workstation and the primary domain failed** | Computer account password out of sync with AD |
| ❌ **There are currently no logon servers available** | Cannot reach Domain Controller - DNS or network issue |
| ⚠️ **The remote computer requires Network Level Authentication** | NLA enabled but credentials cannot be verified |
| ❌ **Remote Desktop can't connect to the remote computer** | RDP disabled, firewall blocking, or service stopped |
| ⚠️ **The user name or password is incorrect** | Wrong credentials, locked account, or trust issue |
| ❌ **An authentication error has occurred** | Kerberos/time sync issue or NLA problem |
| ❌ **Logon Failed** | January 2026 Patch KB5073723 bug - Install KB5077795 |

---

## 3. Step-by-Step Troubleshooting

### Step 1: Verify Network Connectivity

```powershell
Test-NetConnection -ComputerName <server-name-or-ip> -Port 3389
```

**Expected:** ✅ `TcpTestSucceeded: True`

**If False:** Check firewall rules, network connectivity, or if server is online

---

### Step 2: Check Domain Trust

```powershell
Test-ComputerSecureChannel -Verbose
```

**Expected:** ✅ `True`

**If False:** Trust relationship is broken - proceed to repair (Section 9.1)

---

### Step 3: Verify DNS Resolution

```powershell
nslookup <domain.com>
nltest /dsgetdc:<domain.com>
```

**Expected:** Should return Domain Controller IP address

**If fails:** DNS misconfiguration - check DNS settings

---

### Step 4: Check Time Synchronization

```powershell
w32tm /query /status
```

**Important:** Time difference >5 minutes causes Kerberos authentication failures

⚠️ **If time is wrong:** `w32tm /resync /force`

---

### Step 5: Verify RDP Service Status

```powershell
Get-Service TermService | Select-Object Status, StartType
```

**Expected:** ✅ `Status: Running, StartType: Automatic`

---

### Step 6: Check RDP Registry Setting

```powershell
Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name fDenyTSConnections
```

**Expected:** ✅ `fDenyTSConnections: 0` (RDP enabled)

**If 1:** ❌ RDP is disabled - See Section 9.2 to enable

---

### Step 7: Check Firewall Rules

```powershell
Get-NetFirewallRule -DisplayGroup "Remote Desktop" | Select-Object DisplayName, Enabled
```

**Expected:** All rules should show ✅ `Enabled: True`

---

### Step 8: Test with Local Admin

Try logging in with local administrator account (not domain account)

- **If local admin works:** Problem is with domain authentication
- **If local admin fails:** Problem is with RDP service/configuration

---

## 4. Check Last Logged On Users 🆕

This helps identify if the server was recently accessed and by whom:

```powershell
query user
```

**Shows:** Currently logged-in users and session information

### Get Last Login History

```powershell
Get-EventLog -LogName Security -InstanceId 4624 -Newest 10 | 
    Select-Object TimeGenerated, 
        @{N='User';E={$_.ReplacementStrings[5]}}, 
        @{N='LogonType';E={$_.ReplacementStrings[8]}}
```

### Logon Types

| Logon Type | Meaning |
|------------|---------|
| 2 | Interactive (local keyboard/screen) |
| 3 | Network (shared folder access) |
| 10 | Remote Desktop (RDP) |

---

## 5. Check Patch History & Errors 🆕

### 5.1 View Recently Installed Patches

```powershell
Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 10 | 
    Format-Table HotFixID, Description, InstalledOn -AutoSize
```

### 5.2 Check for Failed Windows Updates

```powershell
Get-WinEvent -FilterHashtable @{LogName='System'; ProviderName='Microsoft-Windows-WindowsUpdateClient'; Level=2} -MaxEvents 20 | 
    Select-Object TimeCreated, Message
```

### 🔥 Known Issue - January 2026

**KB5073723** causes "Logon Failed" errors

**Solution:** Install out-of-band patch **KB5077795** (See Section 9.5)

### 5.3 Check Last Boot Time

```powershell
systeminfo | find "System Boot Time"
# OR
(Get-CimInstance Win32_OperatingSystem).LastBootUpTime
```

---

## 6. Check Disk Space & Memory 🆕

### 6.1 Check Disk Space

```powershell
Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3" | 
    Select-Object DeviceID, 
        @{N='Size(GB)';E={[math]::Round($_.Size/1GB,2)}}, 
        @{N='FreeSpace(GB)';E={[math]::Round($_.FreeSpace/1GB,2)}}, 
        @{N='PercentFree';E={[math]::Round(($_.FreeSpace/$_.Size)*100,2)}}
```

### ⚠️ Warning Thresholds

- **<10% free:** Can cause application issues
- **<5% free:** Can prevent Windows Updates and cause system instability
- **<2GB free:** Critical - server may become unresponsive

### 6.2 Check Memory Usage

```powershell
$os = Get-WmiObject Win32_OperatingSystem
$totalMemory = [math]::Round($os.TotalVisibleMemorySize/1MB, 2)
$freeMemory = [math]::Round($os.FreePhysicalMemory/1MB, 2)
$usedMemory = $totalMemory - $freeMemory
$percentUsed = [math]::Round(($usedMemory/$totalMemory)*100, 2)

Write-Output "Total Memory: $totalMemory GB"
Write-Output "Used Memory: $usedMemory GB"
Write-Output "Free Memory: $freeMemory GB"
Write-Output "Percent Used: $percentUsed%"
```

### 6.3 Find Large Files (if disk space low)

```powershell
# Find files larger than 1GB
Get-ChildItem C:\ -Recurse -File -ErrorAction SilentlyContinue | 
    Where-Object {$_.Length -gt 1GB} | 
    Sort-Object Length -Descending | 
    Select-Object FullName, @{N='Size(GB)';E={[math]::Round($_.Length/1GB,2)}} | 
    Format-Table -AutoSize
```

---

## 7. Check Pending Reboot 🆕

Some issues require a reboot to resolve. Check if Windows is waiting for a restart:

```powershell
# Check multiple reboot indicators
$rebootPending = $false

# Check Windows Update
if (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue) {
    Write-Host "Reboot Required: Windows Update" -ForegroundColor Yellow
    $rebootPending = $true
}

# Check Component-Based Servicing
if (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -ErrorAction SilentlyContinue) {
    Write-Host "Reboot Required: Component-Based Servicing" -ForegroundColor Yellow
    $rebootPending = $true
}

# Check Pending File Rename Operations
if (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction SilentlyContinue) {
    Write-Host "Reboot Required: Pending File Operations" -ForegroundColor Yellow
    $rebootPending = $true
}

if ($rebootPending) {
    Write-Host "`n[ACTION REQUIRED] Server needs to be restarted" -ForegroundColor Red
} else {
    Write-Host "`nNo reboot pending" -ForegroundColor Green
}
```

💡 **Best Practice:** Always check with users before rebooting production servers. Schedule maintenance window if needed.

---

## 8. Access Methods When RDP Fails

### 8.1 AWS Systems Manager Session Manager

1. Go to **AWS Console → Systems Manager → Session Manager**
2. Click **Start Session**
3. Select the instance
4. You now have PowerShell access without RDP

### 8.2 Azure Serial Console

1. Go to **Azure Portal → Virtual Machines**
2. Select the VM
3. Under **Support + troubleshooting**, click **Serial Console**
4. Press **Enter** to activate SAC prompt
5. Type: `cmd` then `ch -sn cmd0001`

### 8.3 PowerShell Remoting

```powershell
Enter-PSSession -ComputerName <server> -Credential (Get-Credential)
```

### 8.4 Access via Network Share

```powershell
# Copy files to remote server
Copy-Item "C:\localfile.msu" "\\<server-ip>\c$\temp\"
```

---

## 9. Resolution Procedures

### 9.1 Repair Trust Relationship

🔐 **REQUIRES:** Domain Admin credentials

```powershell
Test-ComputerSecureChannel -Repair -Credential (Get-Credential)
```

Or reset computer password:

```powershell
Reset-ComputerMachinePassword -Credential (Get-Credential)
```

### 9.2 Enable RDP

```powershell
Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name fDenyTSConnections -Value 0
Enable-NetFirewallRule -DisplayGroup "Remote Desktop"
Restart-Service TermService -Force
```

### 9.3 Disable NLA (Temporary)

*Use when trust is broken and NLA prevents login*

```powershell
Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -Name UserAuthentication -Value 0
```

### 9.4 Change DNS Server

```powershell
Set-DnsClientServerAddress -InterfaceAlias "Ethernet" -ServerAddresses ("10.0.0.1","10.0.0.2")
```

### 9.5 Install OOB Patch (KB5077795)

```powershell
# Copy patch to server first, then:
Start-Process -FilePath "wusa.exe" -ArgumentList "C:\temp\windows10.0-kb5077795-x64.msu","/quiet","/norestart" -Wait
Restart-Computer -Force
```

---

## 10. Rejoin Server to Domain 🆕

🔐 **REQUIRES:** Local Admin + Domain Admin credentials

### Step 1: Remove from Domain

```powershell
# Use local admin account
Remove-Computer -UnjoinDomainCredential (Get-Credential) -Force -Restart
```

### Step 2: After Restart, Rejoin Domain

```powershell
# Log in with local admin, then run:
Add-Computer -DomainName "yourdomain.com" -Credential (Get-Credential) -Restart
```

### Alternative: Reset via netdom

```powershell
netdom resetpwd /server:<DC-name> /userd:DOMAIN\admin /passwordd:*
```

⚠️ **Important:** After rejoining, verify GPO applies correctly: `gpupdate /force`

---

## 11. Escalation Criteria

### ⬆️ Escalate to L2/L3 if:

- Trust repair fails after **2 attempts**
- Domain Controller is unreachable from **multiple servers**
- AWS/Azure Serial Console is **not working**
- GPO is **blocking changes**
- **Multiple servers** affected simultaneously
- Issue persists after **domain rejoin**
- **Customer production is down**
- Disk space is **critically low** and cannot free up space
- Errors **increased significantly after patch**

---

## 12. Quick Reference Commands

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
| Last Login | `query user` |
| Installed Patches | `Get-HotFix \| Sort InstalledOn -Desc \| Select -First 10` |
| Disk Space | `Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3"` |
| Memory Usage | `(Get-WmiObject Win32_OperatingSystem).FreePhysicalMemory` |
| Last Boot | `systeminfo \| find "Boot Time"` |
| Pending Reboot | `Get-Item "HKLM:\SOFTWARE\...\RebootPending"` |

---

## 📄 Document Information

**Document Version:** 2.0  
**Last Updated:** January 2026  
**Created by:** Syed Rizvi

---
