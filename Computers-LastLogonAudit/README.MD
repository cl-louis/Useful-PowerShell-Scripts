# Computers-LastLogonAudit Tool

This PowerShell script creates a report of LastLogon for all Computer Accounts in the current logged on users domain. The script queries all DC's in the domain to find the most recent logon as LastLogon is stored independantly on each DC in a Forest.

> [The LastLogon] attribute is not replicated and is maintained separately on each domain controller in the domain. To get an accurate value for the user's last logon in the domain, the Last-Logon attribute for the user must be retrieved from every domain controller in the domain. The largest value that is retrieved is the true last logon time for that user.

from [Last-Logon attribute](https://docs.microsoft.com/en-us/windows/desktop/adschema/a-lastlogon#remarks)

## Use case

Use this script to find stale Computer Accounts. They can then be deactivted or deleted depending on policy.

## Run

Right-Click -> 'Run with PowerShell'

## Output

The script will automatically create a directory called 'Computers-LastLogonAudit' in the users 'Documents' directory. In here, the script will create the report and log file(s).
