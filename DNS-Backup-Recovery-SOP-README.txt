DNS CONFIGURATION BACKUP AND RECOVERY
Standard Operating Procedure - VM Upsize and Rebuild Operations
SAP NS2 PCE Environment - IaaS / OS Team

Prepared by: Syed Rizvi, Cloud Infrastructure Engineer
Date: March 2026
Classification: Controlled Unclassified Information - DoD IL4


================================================================================
BACKGROUND AND PURPOSE
================================================================================

During a VM Upsize with Rebuild operation on SAP NS2 PCE production application
servers, the DNS configuration was completely wiped when the new VM was launched.
The servers lost all DNS settings including nameserver IPs, search domains, and
hostname registration. The result was complete connectivity failure with no RDP
access, no SAP application access, and no NLB routing available.

The existing VM Upsize runbook did not include DNS backup or restore steps. This
document adds those missing steps and provides full guidance to prevent this issue
from occurring again.


================================================================================
SECTION 1 - VM RESIZE RISK ASSESSMENT
================================================================================

Before performing any VM resize operation the team must understand the two types
of resize and the risks each carries for SAP ERP production application servers.


TYPE 1 - UPSIZE WITHOUT REBUILD
--------------------------------
What happens:
The VM is stopped, the instance type is changed, and the same VM is started again.
The same EBS volumes, same network interface, and same private IP address are
retained. Nothing in the OS configuration is touched.

DNS impact: DNS is NOT affected. IP address stays the same.
Downtime: Approximately 7 to 15 minutes.
Risk level: LOW

TYPE 2 - UPSIZE WITH REBUILD
------------------------------
What happens:
The VM is stopped. A new instance is launched from an AMI or snapshot with a
larger size. Volumes are detached from the old instance and reattached to the
new instance. The new instance has a different instance ID.

DNS impact: DNS IS WIPED. The new instance loses NIC configuration, hostname
registration, and Active Directory domain membership.
Downtime: Approximately 60 to 180 minutes.
Risk level: HIGH


RECOMMENDATION
--------------
Always attempt Upsize Without Rebuild first on ERP+ PAAS application servers.

For AWS EC2 instances, changing the instance type such as moving from m5.4xlarge
to m5.8xlarge does not require a rebuild. The same EBS root volume and data
volumes stay attached. Only the compute layer changes.

Only use the rebuild approach when:
- The root EBS volume itself needs to be replaced
- Migrating from Xen hypervisor to Nitro hypervisor
- The instance store type needs to change


WHEN EACH APPROACH APPLIES
----------------------------
Increase CPU or RAM (instance type change)
- Use Without Rebuild. Same volumes reattach. No OS or DNS change needed.

Increase EBS root disk size
- Use Without Rebuild. EBS volumes can be extended online without stopping the VM.

Increase EBS data disk size
- Use Without Rebuild. AWS allows EBS resize without detaching the volume.

Add new data disk volumes
- Use Without Rebuild. New EBS volumes can be hot-attached to a running EC2 instance.

Migrate to Nitro hypervisor from Xen
- Rebuild required. Full DNS backup mandatory before proceeding.

Replace corrupted root volume
- Rebuild required. Full DNS backup mandatory before proceeding.


SAP-SPECIFIC CONSIDERATIONS FOR ERP+ PAAS SERVERS
---------------------------------------------------
PAS - Primary Application Server
The PAS hosts the message server and enqueue service. Stopping the PAS disconnects
all active SAP users. Schedule during a maintenance window and notify all users.
Coordinate with the BASIS team before any work begins.

AAS - Additional Application Server
AAS servers can be taken offline while the PAS handles load. Lower risk if done
one server at a time. Verify load shifts to other AAS nodes before taking down
the next server.

NLB Target Group
After a rebuild the instance ID changes. The old instance ID remains in the NLB
target group as unhealthy. Deregister the old instance ID from the NLB before
the rebuild. Register the new instance ID after the rebuild is complete.

SAP Profiles
SAP profiles reference the hostname directly. If the hostname changes after a
rebuild SAP will not start. The hostname must be identical after rebuild. Verify
with the BASIS team before starting SAP.

SMLG Logon Groups
SAP logon groups reference the server hostname and IP. Run transaction SMLG after
rebuild to verify all logon group entries are correct.


================================================================================
SECTION 2 - WHAT DNS DATA TO CAPTURE BEFORE ANY REBUILD
================================================================================

Before the change window opens the following information must be captured from
each server and saved to the ServiceNow CHG work notes.

Hostname (short and FQDN)
- Why: SAP profiles use the exact hostname. Must match after rebuild.

Primary DNS Server IP
- Why: Wiped on new instance NIC. Causes all name resolution to fail.

Secondary DNS Server IP
- Why: Redundant DNS fails. NLB health checks fail.

DNS Suffix Search List
- Why: SAP connections use short names. Search domain is required.

Private IP Address
- Why: NLB target group must be updated if IP changes.

Active Directory Domain Name
- Why: Rebuilt instance loses AD join. Must rejoin exact same domain and OU.

NLB Target Group Name
- Why: New instance ID must be registered. Old one must be deregistered.

Hosts File Contents (C:\Windows\System32\drivers\etc\hosts)
- Why: Custom host entries for SAP connections are wiped on rebuild.

Network Adapter Name
- Why: Needed to reapply DNS settings to the correct adapter after rebuild.


================================================================================
SECTION 3 - PRE-REBUILD DNS BACKUP COMMANDS (WINDOWS SERVER 2022)
================================================================================

Run all commands below on each server before the change window opens.
Paste the complete output into the ServiceNow CHG work notes.
Save the output file to the staging share.
Responsible team: IaaS / OS Team
Timing: Planning and Preparation phase of the runbook


STEP 1 - Run master DNS backup script on each server

Open PowerShell as Administrator and run the following. Replace ns2sw1app1
with the actual server hostname.

$server = $env:COMPUTERNAME
$outfile = "C:\Temp\$server-DNS-Backup-$(Get-Date -Format yyyyMMdd-HHmm).txt"

"===== DNS BACKUP: $server =====" | Out-File $outfile
"Date: $(Get-Date)" | Out-File $outfile -Append
"" | Out-File $outfile -Append

"--- HOSTNAME ---" | Out-File $outfile -Append
$env:COMPUTERNAME | Out-File $outfile -Append

"--- FULL IPCONFIG ---" | Out-File $outfile -Append
ipconfig /all | Out-File $outfile -Append

"--- DNS CLIENT SERVER ADDRESSES ---" | Out-File $outfile -Append
Get-DnsClientServerAddress | Out-File $outfile -Append

"--- DNS SUFFIX SEARCH LIST ---" | Out-File $outfile -Append
Get-DnsClientGlobalSetting | Select SuffixSearchList | Out-File $outfile -Append

"--- AD DOMAIN ---" | Out-File $outfile -Append
(Get-WmiObject Win32_ComputerSystem).Domain | Out-File $outfile -Append

"--- NET IP CONFIG ---" | Out-File $outfile -Append
Get-NetIPConfiguration | Out-File $outfile -Append

"--- HOSTS FILE ---" | Out-File $outfile -Append
Get-Content C:\Windows\System32\drivers\etc\hosts | Out-File $outfile -Append

"--- DNS RESOLUTION TEST ---" | Out-File $outfile -Append
Resolve-DnsName $env:COMPUTERNAME | Out-File $outfile -Append

"--- SERVER CONNECTIVITY TEST (all 6 app servers) ---" | Out-File $outfile -Append
Test-NetConnection 10.134.4.171 -Port 80 | Out-File $outfile -Append    # ns2sw1app1
Test-NetConnection 10.134.4.153 -Port 80 | Out-File $outfile -Append    # ns2sw2app1
Test-NetConnection 10.134.4.254 -Port 80 | Out-File $outfile -Append    # ns2sw3app1
Test-NetConnection 10.134.4.247 -Port 80 | Out-File $outfile -Append    # ns2sw4app1
Test-NetConnection 10.134.4.109 -Port 80 | Out-File $outfile -Append    # ns2sw5app1
Test-NetConnection 10.134.4.6   -Port 80 | Out-File $outfile -Append    # ns2sw6app1

Write-Host "Backup saved to $outfile" -ForegroundColor Green

Copy-Item $outfile "\\<STAGING_SHARE>\DNS-Backups\"

Note: Run this on all servers before proceeding. Do not begin the rebuild until
the backup file is confirmed saved and pasted into the CHG work notes.


================================================================================
SECTION 4 - POST-REBUILD DNS RESTORE STEPS (WINDOWS SERVER 2022)
================================================================================

After the new VM is launched, connect to it using AWS Systems Manager Session
Manager first. Session Manager works without DNS or network connectivity. Use it
to restore DNS before attempting anything else.

To access Session Manager:
AWS Console - EC2 - Select the new instance - Connect - Session Manager - Connect


STEP 2 - Connect via AWS Session Manager and verify hostname

hostname
$env:COMPUTERNAME


STEP 3 - Set hostname if it was reset during rebuild

Check current hostname first. If it is showing an EC2 default name such as
ip-10-134-4-171 instead of the original hostname, rename it.

Rename-Computer -NewName "ns2sw1app1" -Force

A restart is required after renaming. Reconnect via Session Manager after restart.

Note: The hostname must match the original exactly. SAP profiles and DNS records
depend on the exact hostname.


STEP 4 - Restore DNS server addresses from backup

$nic = Get-NetAdapter | Where-Object {$_.Status -eq "Up"} | Select-Object -First 1
Write-Host "Adapter: $($nic.Name)  Index: $($nic.InterfaceIndex)"

Set-DnsClientServerAddress -InterfaceIndex $nic.InterfaceIndex `
  -ServerAddresses ("10.134.4.171","10.134.4.153")

Set-DnsClientGlobalSetting `
  -SuffixSearchList @("ns2.sap.mil","ns2corp.local","amazonaws.com")

Get-DnsClientServerAddress
Get-DnsClientGlobalSetting | Select SuffixSearchList

Note: Use the DNS IPs from the pre-rebuild backup file. Do not guess the IPs.


STEP 5 - Flush DNS and force re-registration

ipconfig /flushdns
ipconfig /registerdns

Start-Sleep -Seconds 30

nslookup $env:COMPUTERNAME
nslookup ns2sw1app1
ping -n 3 ns2sw2app1

# Test connectivity to the other 5 app servers from this rebuilt server
Test-NetConnection 10.134.4.171 -Port 80    # ns2sw1app1
Test-NetConnection 10.134.4.153 -Port 80    # ns2sw2app1
Test-NetConnection 10.134.4.254 -Port 80    # ns2sw3app1
Test-NetConnection 10.134.4.247 -Port 80    # ns2sw4app1
Test-NetConnection 10.134.4.109 -Port 80    # ns2sw5app1
Test-NetConnection 10.134.4.6   -Port 80    # ns2sw6app1

All tests must pass before proceeding to Active Directory re-join.


STEP 6 - Re-join Active Directory domain if lost

Check current domain status first.

(Get-WmiObject Win32_ComputerSystem).Domain

If the output shows WORKGROUP instead of the domain name the server must be
rejoined. Use the exact domain and OU path from the pre-rebuild backup.

Add-Computer `
  -DomainName "ns2corp.local" `
  -OUPath "OU=PCE-Servers,OU=NS2,DC=ns2corp,DC=local" `
  -Credential (Get-Credential) `
  -Restart

After restart verify domain join.

(Get-WmiObject Win32_ComputerSystem).Domain
whoami /fqdn

Note: Contact the BASIS team before restarting. SAP must be fully stopped before
the domain rejoin restart.


STEP 7 - Restore custom hosts file entries

Compare current hosts file with the backup and add back any missing entries.

notepad C:\Windows\System32\drivers\etc\hosts

Or append entries via PowerShell:

Add-Content C:\Windows\System32\drivers\etc\hosts `
  "10.x.x.x  sapdb.ns2corp.local  sapdb"

Replace the IP and hostname with actual values from the backup file.


================================================================================
SECTION 5 - MISSING RUNBOOK STEPS - WHERE TO INSERT
================================================================================

The following steps are missing from the existing VM Upsize With Rebuild runbook
and must be added to the Excel tracking sheet.


INSERT AFTER SN 5 - Before Downtime Activity section

SN 5A - DNS Backup
Run the DNS backup script on all affected servers. Save output to the CHG work
notes and staging share. Verify the backup file contains hostname, DNS IPs,
search domains, AD domain, hosts file contents, and NLB connectivity test results.
Responsible team: IaaS / OS Team
Approximate time: 20 minutes per server

SN 5B - NLB Pre-Check
Record the current NLB target group name, current instance IDs, and health status
for all affected servers. Screenshot and paste into CHG work notes.
Responsible team: IaaS / OS Team
Approximate time: 10 minutes

SN 5C - SAP Pre-Check
Confirm with the BASIS team that SAP is fully stopped and no users are connected.
Run transaction SM50 to confirm no active work processes. Get BASIS sign-off
before proceeding to downtime activity.
Responsible team: BASIS Team
Approximate time: 10 minutes


INSERT AFTER SN 15 - Create new data disks, Before attaching data disks to VM

SN 15A - NLB Deregister
Deregister the old instance ID from the NLB target group before launching the
new VM. AWS Console - EC2 - Target Groups - Deregister old instance. Verify
health check shows Deregistered status.
Responsible team: IaaS / OS Team
Approximate time: 10 minutes

SN 15B - DNS Post-Rebuild Restore
After the new VM launches connect via Session Manager. Restore hostname, DNS
server IPs, DNS suffix search list, AD domain join, and hosts file using backed-up
values from Step 5A. Verify nslookup resolves correctly before proceeding.
Responsible team: IaaS / OS Team
Approximate time: 30 minutes

SN 15C - NLB Register
Register the new instance ID in the NLB target group. AWS Console - EC2 - Target
Groups - Register Targets - Select new instance. Wait for health check to show
Healthy status before proceeding.
Responsible team: IaaS / OS Team
Approximate time: 15 minutes

SN 15D - DNS Verification
Run the full DNS verification checklist on the rebuilt server. All checks must
pass before notifying the BASIS team to start SAP. No exceptions.
Responsible team: IaaS / OS Team
Approximate time: 15 minutes


INSERT AT END - After SAP Start, Before Closing the CHG

FINAL A - SAP Logon Groups
After SAP start, BASIS team verifies transaction SMLG logon groups show correct
server entries. Confirm all application servers appear in the correct logon groups.
Responsible team: BASIS Team
Approximate time: 10 minutes

FINAL B - AD DNS Cleanup
Check Active Directory DNS Manager or Route 53 for stale A and PTR records from
the old instance. Delete old records and verify new records resolve correctly from
at least two different clients.
Responsible team: IaaS / OS Team
Approximate time: 10 minutes


================================================================================
SECTION 6 - NLB RE-REGISTRATION AFTER REBUILD
================================================================================

The six ERP+ PAAS application servers and their IP addresses are listed below.
After a rebuild the DNS and network config must be restored on the specific server
that was rebuilt. The other servers remain unaffected.

ns2sw1app1 - 10.134.4.171
ns2sw2app1 - 10.134.4.153
ns2sw3app1 - 10.134.4.254
ns2sw4app1 - 10.134.4.247
ns2sw5app1 - 10.134.4.109
ns2sw6app1 - 10.134.4.6


STEP 8 - Deregister old instance from NLB before rebuild

Get the current instance ID before stopping the server.

aws ec2 describe-instances \
  --filters "Name=private-ip-address,Values=10.134.4.171" \
  --query "Reservations[0].Instances[0].InstanceId" --output text

Deregister from the target group.

aws elbv2 deregister-targets \
  --target-group-arn <TARGET_GROUP_ARN> \
  --targets Id=<OLD_INSTANCE_ID>

Or via AWS Console:
EC2 - Load Balancers - Target Groups - Select group - Targets tab
Select old instance - Deregister

Perform this step for both NLB target groups.


STEP 9 - Register new instance in NLB after rebuild and DNS restore

Get the new instance ID.

aws ec2 describe-instances \
  --filters "Name=private-ip-address,Values=<SERVER_IP>" \
  --query "Reservations[0].Instances[0].InstanceId" --output text

Register the new instance in the target group.

aws elbv2 register-targets \
  --target-group-arn <TARGET_GROUP_ARN> \
  --targets Id=<NEW_INSTANCE_ID>

Check for healthy status.

aws elbv2 describe-target-health \
  --target-group-arn <TARGET_GROUP_ARN>

Wait until the State field shows healthy before proceeding.

Note: Do not start SAP until the NLB shows Healthy for the new instance.


================================================================================
SECTION 7 - FINAL VERIFICATION CHECKLIST BEFORE NOTIFYING BASIS TO START SAP
================================================================================

All items below must be verified and confirmed before the BASIS team is notified
to start SAP. Document results in ServiceNow CHG work notes.

Verification 1 - Hostname correct
Command: hostname or $env:COMPUTERNAME
Expected result: Exact original hostname such as ns2sw1app1

Verification 2 - DNS servers set correctly
Command: Get-DnsClientServerAddress
Expected result: Primary and secondary IPs match the pre-rebuild backup

Verification 3 - DNS resolves hostname
Command: nslookup ns2sw1app1
Expected result: Returns correct 10.134.4.x IP address

Verification 4 - Reverse DNS works
Command: nslookup 10.134.4.171
Expected result: Returns ns2sw1app1 with domain suffix

Verification 5 - Active Directory domain joined
Command: (Get-WmiObject Win32_ComputerSystem).Domain
Expected result: Shows domain name and not WORKGROUP

Verification 6 - Can ping other application servers
Command: ping -n 3 ns2sw2app1
Expected result: All 3 pings succeed

Verification 7 - NLB health check green
Method: AWS Console - Target Groups
Expected result: Status shows Healthy

Verification 8 - NLB port test passes
Command: Test-NetConnection 10.134.4.171 -Port 80    # ns2sw1app1
         Test-NetConnection 10.134.4.153 -Port 80    # ns2sw2app1
         Test-NetConnection 10.134.4.254 -Port 80    # ns2sw3app1
         Test-NetConnection 10.134.4.247 -Port 80    # ns2sw4app1
         Test-NetConnection 10.134.4.109 -Port 80    # ns2sw5app1
         Test-NetConnection 10.134.4.6   -Port 80    # ns2sw6app1
Expected result: TcpTestSucceeded = True for all servers in the 10.134.4.x subnet

Verification 9 - SAP HANA DB reachable
Command: Test-NetConnection <HANA_IP> -Port 30015
Expected result: TcpTestSucceeded = True

Verification 10 - Hosts file correct
Command: Get-Content C:\Windows\System32\drivers\etc\hosts
Expected result: All custom entries from the backup are present

All 10 verifications must pass before notifying BASIS to start SAP.
If any verification fails stop and resolve the issue before proceeding.


================================================================================
Prepared by: Syed Rizvi
Cloud Infrastructure Engineer
SAP National Security Services
March 2026
Classification: Controlled Unclassified Information - DoD IL4
================================================================================
