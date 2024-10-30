# Get-EOLMailboxUsage

This script generates a report of the mailbox size/usage for: a single mailbox, mailbox(es) from a file,
or all mailboxes in a 365 tenant.

To use the 'from a file' option create a 'users.txt' file in this directory. The file will need to contain a CRLF delimited list of email addresses. e.g.

```txt
user1@company.com
user2@company.com
manager3@company.com
```

## Use case

Use this script to audit mailbox usage for a Microsft 365 Tenant.

## Run

Right-Click -> 'Run with PowerShell'

## Output

The script will automatically create a directory called 'Get-EOLMailboxUsage' in the users 'Documents' directory.
In here, the script will create the report, log and transcript file(s).
