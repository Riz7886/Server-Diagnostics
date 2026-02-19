# 🚀 QUICK START GUIDE
## Plugin 65057 - SAP Service Permission Fix

### ⚡ ONE-PAGE QUICK REFERENCE

---

## 📋 TICKET INFO
- **Ticket:** INC0135584
- **Plugin:** Nessus 65057 - Insecure Windows Service Permissions
- **Servers:** 6 SAP sidecar servers (AWS GovCloud)
- **Environment:** DOD FedRAMP

---

## 🎯 3-STEP PROCESS

### STEP 1: RUN ASSESSMENT (5-10 minutes)
```powershell
# Open PowerShell as Administrator
cd C:\Scripts
.\Plugin65057-SAP-Assessment.ps1

# Enter domain credentials when prompted
# Wait for HTML report to open automatically
```

### STEP 2: REVIEW REPORT
- ✅ Check which servers have issues
- 🔴 **PRIORITY:** SAP services (yellow highlighted rows)
- 📝 Copy remediation script from report

### STEP 3: FIX PERMISSIONS
```powershell
# Copy script from HTML report
# Test in non-prod first
# Run during maintenance window
# Verify services afterward
```

---

## 🎨 REPORT COLOR CODES

| Color/Badge | Meaning | Action |
|-------------|---------|--------|
| 🔴 CRITICAL - SAP | SAP service affected | **DO FIRST** - Coordinate with SAP team |
| ⚠️ PATCH + PERMISSIONS | Both issues | Install updates AND fix permissions |
| 🔄 PATCH REQUIRED | Needs updates | Run Windows Update |
| 🔧 PERMISSIONS FIX | Permission issue | Run remediation script |
| ✓ NO ACTION | Clean | Nothing needed |
| 🟡 Yellow rows in table | SAP service | Handle with SAP team coordination |

---

## ⚠️ CRITICAL WARNINGS

### BEFORE YOU START
- ✅ Run as Administrator
- ✅ Have admin credentials for all 6 servers
- ✅ Coordinate with SAP team for SAP services (yellow rows)
- ✅ Create change request for production
- ✅ Test in non-prod first

### NEVER DO THIS
- ❌ Run in production without testing
- ❌ Fix SAP services without SAP team coordination
- ❌ Skip creating backups (script does it automatically)
- ❌ Ignore yellow highlighted services
- ❌ Run during business hours without approval

---

## 🔧 TYPICAL REMEDIATION SCRIPT

```powershell
#Requires -RunAsAdministrator

# Services from report (example)
$services = @(
    "ServiceName1",
    "ServiceName2"
)

# Create backup
$backupPath = "C:\SecurityBackups\SAP_ServicePerms_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -ItemType Directory -Path $backupPath -Force | Out-Null

# Fix each service
foreach ($serviceName in $services) {
    $serviceKey = "HKLM:\SYSTEM\CurrentControlSet\Services\$serviceName"
    
    # Backup current ACL
    $acl = Get-Acl $serviceKey
    $acl | Export-Clixml "$backupPath\$serviceName-acl.xml"
    
    # Remove bad permissions (Users, Everyone, etc.)
    $acl.Access | Where-Object {
        ($_.IdentityReference -match "Users|Everyone|Authenticated Users") -and
        ($_.RegistryRights -match "FullControl|Modify|Write")
    } | ForEach-Object {
        $acl.RemoveAccessRule($_) | Out-Null
    }
    
    # Apply fixed permissions
    Set-Acl -Path $serviceKey -AclObject $acl
    Write-Host "✓ Fixed: $serviceName" -ForegroundColor Green
}

# Verify services still work
foreach ($serviceName in $services) {
    $svc = Get-Service $serviceName
    Write-Host "$serviceName - Status: $($svc.Status)" -ForegroundColor $(if($svc.Status -eq 'Running'){'Green'}else{'Yellow'})
}
```

---

## 🆘 QUICK TROUBLESHOOTING

| Problem | Solution |
|---------|----------|
| Access Denied | Run PowerShell as Administrator |
| Server Unreachable | Check network: `Test-Connection 10.134.4.171` |
| WinRM Failed | Verify WinRM service running on target |
| Script Won't Run | `Set-ExecutionPolicy RemoteSigned -Scope Process` |
| Service Won't Start | Restore from backup in `C:\SecurityBackups\` |

---

## 📊 TARGET SERVERS (Pre-configured)
```
✅ 10.134.4.171
✅ 10.134.4.153
✅ 10.134.4.254
✅ 10.134.4.109
✅ 10.134.4.247
✅ 10.134.4.6
```

---

## 📁 OUTPUT LOCATIONS

All files saved to: `C:\SecurityRemediation\Plugin65057_SAP\`

| File | Purpose |
|------|---------|
| `SAP_Plugin65057_Report_*.html` | **Main report** - Open this first |
| `SAP_Plugin65057_Data_*.csv` | Raw data for tracking |
| `SAP_Plugin65057_Log_*.log` | Execution details |
| `C:\SecurityBackups\*` | ACL backups (auto-created) |

---

## ✅ SUCCESS CHECKLIST

- [ ] Assessment script ran successfully
- [ ] HTML report reviewed
- [ ] SAP services identified (yellow rows)
- [ ] Change request created (INC0135584)
- [ ] SAP team coordinated
- [ ] Tested in non-prod
- [ ] Executed in production during maintenance
- [ ] All services verified running
- [ ] SAP functionality tested
- [ ] Nessus rescan shows Plugin 65057 resolved
- [ ] Ticket INC0135584 closed

---

## 🎯 WORKFLOW AT A GLANCE

```
1. RUN SCRIPT (10 min)
   ↓
2. REVIEW HTML REPORT (15 min)
   ↓
3. CREATE CHANGE REQUEST (30 min)
   ↓
4. COORDINATE WITH SAP TEAM (1 hour)
   ↓
5. TEST IN NON-PROD (1-2 hours)
   ↓
6. SCHEDULE MAINTENANCE (plan ahead)
   ↓
7. EXECUTE REMEDIATION (30 min per server)
   ↓
8. VERIFY FUNCTIONALITY (1 hour)
   ↓
9. NESSUS RESCAN (wait 24-48 hours)
   ↓
10. CLOSE TICKET INC0135584 ✅
```

---

## 💡 PRO TIPS

1. **SAP Services = Priority** - Yellow rows in report are critical
2. **Test First** - Always run in non-prod before production
3. **Backup Automatically Created** - Script handles it, keep for 7+ days
4. **Coordinate** - SAP team must be involved for SAP services
5. **Timing** - Run during maintenance window, off-hours preferred

---

## 📞 EMERGENCY ROLLBACK

If service fails after remediation:

```powershell
# Find your backup
cd C:\SecurityBackups
dir | Sort LastWriteTime -Descending

# Restore service ACL
$serviceName = "YourServiceName"
$backupPath = "C:\SecurityBackups\SAP_ServicePerms_YYYYMMDD_HHMMSS"
$backupACL = Import-Clixml "$backupPath\$serviceName-acl-backup.xml"
Set-Acl -Path "HKLM:\SYSTEM\CurrentControlSet\Services\$serviceName" -AclObject $backupACL

# Restart service
Restart-Service $serviceName
```

---

## 📈 EXPECTED RESULTS

**Before Fix:**
- Nessus shows Plugin 65057 vulnerability
- Services have "Users" or "Everyone" with Modify/Write access
- Security risk for privilege escalation

**After Fix:**
- Only SYSTEM and Administrators have Full Control
- Users have Read-only access
- Plugin 65057 resolved in Nessus rescan
- Services function normally

---

## 🔐 COMPLIANCE NOTES

**For DOD FedRAMP:**
- Document all changes in ticket
- Maintain audit trail
- Keep backups for compliance review
- Follow change management process
- Test before production deployment

**Backup Retention:**
- Minimum 7 days
- Store in `C:\SecurityBackups\`
- Include in change documentation

---

## 🎓 KEY CONCEPTS

**What's the Issue?**
Services have permissions allowing non-admins to modify them. This is a security risk.

**What's the Fix?**
Remove excessive permissions so only admins can modify services.

**Why SAP Services Matter?**
SAP is critical for operations. Any change requires coordination and testing.

**What About Patches?**
If report says "PATCH REQUIRED", install Windows updates via WSUS/SCCM.

---

## 📧 CONTACT FOR ISSUES

| Issue Type | Contact |
|------------|---------|
| Script errors | Security Remediation Team |
| SAP services | SAP Administrators |
| Change requests | Change Management Team |
| Compliance questions | Compliance Team |

---

## 🏁 FINAL REMINDER

**This is PRODUCTION - Be Careful!**

1. ✅ Always test first
2. ✅ Coordinate with teams
3. ✅ Follow change management
4. ✅ Have rollback plan
5. ✅ Document everything

**The script is ready to use. No modifications needed. Just run it!**

---

**QUICK HELP:** If stuck, see full README.md for detailed troubleshooting and examples.

**VERSION:** 1.0  
**DATE:** 2026-02-19  
**TICKET:** INC0135584  
