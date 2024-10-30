# Useful PowerShell Scripts

This repo is a collection of PowerShell scripts for:

- Common tasks
- Automations
- Audit tools
- Housekeeping

## Use

Clone/download the entire repo:

```cmd
git clone https://github.com/cl-louis/Useful-PowerShell-Scripts.git
```

Open the directory of the required script.

Right-Click the script -> 'Run with PowerShell'.

## Contents

### Computers-LastLogonAudit

This PowerShell script creates a report of LastLogon for all Computer Accounts in the current logged on users domain. The script queries all DC's in the domain to find the most recent logon as LastLogon is stored independantly on each DC in a Forest.

See README.MD in Computers-LastLogonAudit for more information.

### Audit-VigoMapping

This script audits the mapping of Z drive to ensure it uses the correct UNC path, then searches the users desktop for the VigoGLogon shortcut to ensure it has been created from a valid drive mapping.

See README.MD in Audit-VigoMapping for more information.

### Groups-EmptySecGroupsAudit

This PowerShell script creates a report of any non-builtin Security Groups that have no members.

See README.MD in Groups-EmptySecGroupsAudit for more information.

### OU-GPOApplicationAuditTool

This script finds all OUs in a domain and their applied/inherited GPOs.

See README.MD in OU-GPOApplicationAuditTool for more information.

### Users-LastLogonAudit

This PowerShell script creates a report of LastLogon for all User Accounts in the current logged on users domain. The script queries all DC's in the domain to find the most recent logon as LastLogon is stored independantly on each DC in a Forest.

See README.MD in Users-LastLogonAudit for more information.

### Install-RSATTools

This script installs RSAT.ServerManager.Tools, RSAT.ActiveDirectory.DS-LDS.Tools, RSAT.CertificateServices.Tools, RSAT.DHCP.Tools, RSAT.Dns.Tools, and RSAT.GroupPolicy.Management.Tools onto a users PC.

See README.MD in Install-RSATTools for more information.

### Set-UserPassword

This script allows a user to search for a domain account and then generates a secure password for the account.
The password is set on the PDC of the local PCs domain.

See README.MD in Set-UserPassword for more information.

### Get-SoftwareInventory

This script finds installed software packages on the local computer and generates a
report.

See README.MD in Get-SoftwareInventory for more information.

### Get-OpenPortsWithProcess

This script finds open/listening TCP and UDP ports on the local computer with the associated process.

See README.MD in Get-OpenPortsWithProcess for more information.

### Groups-AuditGroups

This script creates a report of all Security Groups in the domain.

See README.MD in Get-AuditGroups for more information.

### Get-EOLEmailVolume

This script creates a report of EOL email volume (send and receive).

See README.MD in Get-EOLEmailVolume for more information.

### Get-LoggedInUser

This function gets the current user sessions on remote or local computers.

See README.MD in Get-LoggedInUser for more information.

### Get-SMBShareAudit

This function gets the current user sessions on remote or local computers.

See README.MD in Get-SMBShareAudit for more information.

### Get-EOLMailboxUsage

This script creates a report of EOL mailbox usage.

See README.MD in Get-EOLMailboxUsage for more information.
