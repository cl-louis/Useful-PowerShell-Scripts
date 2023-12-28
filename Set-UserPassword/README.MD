# Set-UserPassword

This script allows a user to search for a domain account and then generates a secure password for the account.
The password is set on the PDC of the local PCs domain.

Passwords are constructed using 3 files: Words.txt, Numbers.txt, and SpecialCharacters.txt.
By default, the password uses 3 items from Words.txt, 3 from Numbers.txt, and 1 from SpecialCharacters.txt

An Example password could be: ForeverMagicCarpets821!

## Use case

Use this script to set secure passwords on AD User Accounts.

## Run

Right-Click -> 'Run with PowerShell'

## Output

The script will automatically create a directory called 'Set-UserPassword' in the users 'Documents' directory.
In here, the script will create the report, log and transcript file(s).
