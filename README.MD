# Useful PowerShell Scripts

This repo is a collection of PowerShell scripts for:

- Common tasks
- Automations
- Audit tools
- Housekeeping

## Use

Clone/download the entire repo:

```cmd
git clone https://github.com/cl-louis/PowerShell-Scripts.git
```

Open the directory of the required script.

Right-Click the script -> 'Run with PowerShell'.


## Contents

### Computers-LastLogonAudit

This PowerShell script creates a report of LastLogon for all Computer Accounts in the current logged on users domain. The script queries all DC's in the domain to find the most recent logon as LastLogon is stored independantly on each DC in a Forest.

See README.MD in Computers-LastLogonAudit for more information.