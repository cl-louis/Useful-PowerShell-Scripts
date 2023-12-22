# Audit-VigoMapping

This script audits the mapping of Z drive to ensure it uses the correct UNC path, then searches the users desktop for the VigoGLogon shortcut to ensure it has been created from a valid drive mapping.

## Use case

Use this script to ensure that Z drive mappings and VigoGLogon shortcuts are correct and consistent across users.

## Run

Right-Click -> 'Run with PowerShell'

## Output

The script will automatically create a directory called 'Audit-VigoMapping' in the users 'Documents' directory. In here, the script will create the transcript and log file(s).
