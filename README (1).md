# SAP Plugin 65057 - Insecure Windows Service Permissions
## Complete Assessment and Remediation Solution

### 📋 Overview
This solution addresses **Nessus Plugin 65057** vulnerabilities on SAP sidecar servers in your DOD FedRAMP AWS GovCloud environment.

**Ticket:** INC0135584  
**Environment:** DOD FedRAMP AWS GovCloud  
**Servers:** 6 Windows Server 2022 instances  
**Issue:** Insecure Windows Service Permissions  

---

## 🎯 What This Does

This script will:
- ✅ Connect to all 6 SAP servers automatically
- ✅ Check Windows patch status
- ✅ Analyze service permissions for vulnerabilities
- ✅ Identify SAP-specific services (marked as CRITICAL)
- ✅ Determine if patching or permission fix is needed
- ✅ Generate comprehensive HTML report with remediation scripts
- ✅ Export raw data to CSV for tracking
- ✅ Provide ready-to-run PowerShell scripts for each server

---

## 🚀 Quick Start Guide

### **Step 1: Download Files from Git**

1. Download `Plugin65057-SAP-Assessment.ps1` from your Git repository
2. Save it to your workstation (e.g., `C:\Scripts\`)

### **Step 2: Run the Assessment**

Open PowerShell **as Administrator**:

```powershell
# Navigate to script location
cd C:\Scripts

# Run the assessment
.\Plugin65057-SAP-Assessment.ps1
```

### **Step 3: Enter Credentials**

When prompted, enter your domain credentials:
- **Format:** `DOMAIN\username`
- **Requirements:** Admin rights on all target servers

### **Step 4: Wait for Completion**

The script will:
- Connect to each of the 6 servers
- Perform security analysis (takes 2-5 minutes per server)
- Generate HTML report automatically

### **Step 5: Review Report**

The HTML report will open automatically showing:
- ✅ Which servers need patching
- ✅ Which services have bad permissions
- ✅ SAP services affected (marked in yellow - CRITICAL)
- ✅ Ready-to-use remediation scripts

---

## 📊 Target Servers

The script is pre-configured for these 6 servers:
```
10.134.4.171
10.134.4.153
10.134.4.254
10.134.4.109
10.134.4.247
10.134.4.6
```

---

## 🔧 What You'll Get

### **Output Files**

All files are saved to: `C:\SecurityRemediation\Plugin65057_SAP\`

| File | Description |
|------|-------------|
| **HTML Report** | Complete analysis with color-coded status and remediation scripts |
| **CSV File** | Raw vulnerability data for import into tracking systems |
| **Log File** | Detailed execution log for audit trail |

### **HTML Report Features**

- 📊 Executive summary dashboard
- 🖥️ Per-server detailed analysis
- 🔴 SAP services highlighted (CRITICAL priority)
- 📝 Copy-paste ready remediation scripts
- 🔄 Patch status and recommendations
- ⚠️ Risk prioritization

---

## 🛠️ Remediation Process

### **For Permission Issues** (Most Common)

1. **Review** the HTML report
2. **Find** your server section
3. **Copy** the remediation script from the report
4. **Test** in non-production first
5. **Run** on production during maintenance window

**Example Remediation Script** (generated in report):
```powershell
# The report will contain a script like this for each affected server
# Just copy and paste it to fix permissions

#Requires -RunAsAdministrator

$services = @(
    "ServiceName1",
    "ServiceName2"
)

# Backup and fix permissions automatically
foreach ($serviceName in $services) {
    # Script handles backup, permission removal, and verification
}
```

### **For Patch Issues**

If the report shows "PATCH REQUIRED":

1. Review pending Windows updates
2. Schedule maintenance window
3. Install updates via WSUS/SCCM
4. Reboot if required
5. Re-run assessment script

---

## ⚠️ Important Notes for DOD Environment

### **Before Running**

- ✅ You have admin rights on target servers
- ✅ Network connectivity to all 6 servers
- ✅ Script can run from your workstation
- ✅ Change management process followed

### **Testing Requirements**

- ⚠️ **ALWAYS** test remediation scripts in non-production first
- ⚠️ Coordinate with SAP team before changes
- ⚠️ Create change request (CR) for production

### **SAP Service Changes**

- 🔴 SAP services are marked as **CRITICAL** in the report
- 🔴 Notify SAP administrators before fixing
- 🔴 Test SAP functionality after remediation
- 🔴 Have rollback plan ready

---

## 📝 Step-by-Step Remediation Workflow

### **Phase 1: Assessment**
1. Run `Plugin65057-SAP-Assessment.ps1`
2. Review HTML report
3. Identify affected servers and services
4. Prioritize SAP services (marked in yellow)

### **Phase 2: Planning**
1. Create change request for ticket INC0135584
2. Coordinate with SAP team
3. Schedule maintenance windows
4. Plan rollback strategy

### **Phase 3: Testing**
1. Copy remediation script from report
2. Test in non-production environment
3. Verify services start correctly
4. Test SAP functionality

### **Phase 4: Production Execution**
1. Execute remediation during maintenance window
2. Monitor service status
3. Verify SAP functionality
4. Check application logs

### **Phase 5: Verification**
1. Wait 24-48 hours
2. Re-scan with Nessus
3. Verify Plugin 65057 is resolved
4. Update ticket INC0135584
5. Document completion

---

## 🔍 Understanding the Report

### **Server Status Badges**

| Badge | Meaning | Action Required |
|-------|---------|-----------------|
| ⚠️ PATCH + PERMISSIONS | Both issues found | Install updates AND fix permissions |
| 🔄 PATCH REQUIRED | Outdated patches | Install Windows updates |
| 🔧 PERMISSIONS FIX | Permission issues only | Run remediation script |
| ✓ NO ACTION NEEDED | Clean | No action required |
| 🔴 CRITICAL - SAP AFFECTED | SAP service has issue | **PRIORITY - Coordinate with SAP team** |

### **Service Table Columns**

- **Service Name:** Windows service identifier
- **Display Name:** Human-readable name
- **State:** Current status (Running/Stopped)
- **Vulnerable Identity:** Who has excessive permissions (e.g., "Everyone", "Users")
- **Dangerous Rights:** What they can do (e.g., "FullControl", "Modify")
- **Type:** 🔴 SAP (critical) or Standard

### **Yellow Highlighted Rows**

Any row with **yellow background** = SAP service = **CRITICAL PRIORITY**

---

## 🆘 Troubleshooting

### **Problem: "Access Denied" Error**
**Solution:** 
- Run PowerShell as Administrator
- Verify credentials have admin rights on target servers
- Check domain membership

### **Problem: "Server Unreachable"**
**Solution:**
- Verify network connectivity: `Test-Connection 10.134.4.171`
- Check firewall rules for WinRM (port 5985/5986)
- Verify server is online

### **Problem: "WinRM Connection Failed"**
**Solution:**
```powershell
# Enable WinRM on target server (if allowed)
Enable-PSRemoting -Force

# Or check if WinRM service is running
Get-Service WinRM
```

### **Problem: Script Won't Run - Execution Policy**
**Solution:**
```powershell
# Check current policy
Get-ExecutionPolicy

# If needed, temporarily allow (as Admin)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
```

### **Problem: Service Won't Start After Fix**
**Solution:**
```powershell
# Restore from backup (script creates automatically)
$backupPath = "C:\SecurityBackups\SAP_ServicePerms_<TIMESTAMP>"
$serviceName = "YourServiceName"

$backupACL = Import-Clixml "$backupPath\$serviceName-acl-backup.xml"
Set-Acl -Path "HKLM:\SYSTEM\CurrentControlSet\Services\$serviceName" -AclObject $backupACL
```

---

## 📞 Support & Escalation

### **For Script Issues**
- Check log file: `C:\SecurityRemediation\Plugin65057_SAP\SAP_Plugin65057_Log_*.log`
- Review error messages in PowerShell console
- Verify prerequisites are met

### **For SAP Service Issues**
- Contact SAP team before making changes
- Have SAP service account information ready
- Coordinate testing schedule

### **For Change Management**
- Document all changes in ticket INC0135584
- Include before/after screenshots
- Attach HTML report and CSV data

---

## ✅ Success Criteria

You know remediation is successful when:

1. ✅ Remediation script completes without errors
2. ✅ All affected services show "Status: Running" (if set to Automatic)
3. ✅ SAP functionality tests pass
4. ✅ Nessus rescan shows Plugin 65057 resolved
5. ✅ No application errors in Event Viewer
6. ✅ Ticket INC0135584 can be closed

---

## 📋 Checklist for Ticket Closure

- [ ] Assessment report reviewed
- [ ] Change request created and approved
- [ ] SAP team notified and coordinated
- [ ] Remediation scripts tested in non-prod
- [ ] Production remediation executed successfully
- [ ] All affected services verified working
- [ ] SAP functionality confirmed
- [ ] Nessus rescan completed (Plugin 65057 resolved)
- [ ] Backups retained for 7+ days
- [ ] Documentation updated in compliance system
- [ ] Ticket INC0135584 updated with completion notes

---

## 🔐 Security & Compliance Notes

### **Audit Trail**
- All changes are logged automatically
- Backups created before any modifications
- Timestamps recorded for compliance

### **Backup Retention**
- Keep ACL backups for minimum 7 days
- Store in: `C:\SecurityBackups\SAP_ServicePerms_*`
- Document backup location in ticket

### **FedRAMP Requirements**
- Follow your organization's change management process
- Document all security remediation activities
- Maintain audit trail for compliance review
- Test changes before production deployment

---

## 📖 Additional Resources

### **Understanding Plugin 65057**
Nessus Plugin 65057 detects Windows services with permissions that allow non-administrative users to modify service configuration or binaries. This can lead to privilege escalation attacks.

### **Why This Matters**
In FedRAMP environments, service permissions must follow principle of least privilege:
- Only SYSTEM and Administrators should have Full Control
- Standard users should have Read access only
- "Everyone" and "Users" should never have Modify/Write access

### **Valid Service Permissions**
```
✅ GOOD:
- SYSTEM: Full Control
- Administrators: Full Control  
- Users: Read

❌ BAD:
- Everyone: Full Control
- Users: Modify
- Authenticated Users: Write
```

---

## 🎓 Training Resources

### **PowerShell Basics**
- How to run scripts as Administrator
- Understanding execution policies
- Reading error messages

### **Windows Services**
- Service account management
- ACL (Access Control List) basics
- Registry permissions

### **SAP on Windows**
- SAP service architecture
- Service dependencies
- SAP-specific considerations

---

## 💡 Pro Tips

1. **Run assessment during business hours** - You may need to contact server admins
2. **Save the HTML report** - You'll need it for documentation
3. **Test one server first** - Verify process before scaling
4. **Screenshot everything** - Good for change records
5. **Keep backups** - Script creates them automatically, don't delete
6. **Coordinate with SAP team** - They may have testing requirements

---

## 📅 Maintenance Schedule Recommendations

**Best Times to Run Remediation:**
- Outside business hours for production
- During scheduled maintenance windows
- When SAP team is available for verification
- After normal backups have completed

**Avoid:**
- Month-end processing periods
- During system backups
- Peak business hours
- When SAP team is unavailable

---

## 🏁 Quick Reference Commands

```powershell
# Run assessment
.\Plugin65057-SAP-Assessment.ps1

# Check if server is reachable
Test-Connection 10.134.4.171

# Check service status
Get-Service -ComputerName 10.134.4.171 | Where-Object {$_.Name -like "*SAP*"}

# View recent changes
Get-EventLog -LogName System -Source "Service Control Manager" -Newest 20

# Check Windows updates
Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 10
```

---

## 📧 Report Delivery

### **Who Gets the Report**
- Security team (for review)
- SAP administrators (for coordination)
- Change management team (for approval)
- Compliance team (for audit trail)

### **What to Include**
- Full HTML report
- CSV data file
- Execution log
- Ticket number (INC0135584)
- Planned remediation schedule

---

**VERSION:** 1.0  
**LAST UPDATED:** 2026-02-19  
**ENVIRONMENT:** DOD FedRAMP AWS GovCloud  
**TICKET:** INC0135584  

---

## 📞 Need Help?

If you encounter issues:
1. Review the Troubleshooting section above
2. Check the log file for detailed errors
3. Verify all prerequisites are met
4. Contact your security team for guidance

**Remember:** This is a production environment. Always test first, coordinate with teams, and follow change management processes.

---

*End of README*
