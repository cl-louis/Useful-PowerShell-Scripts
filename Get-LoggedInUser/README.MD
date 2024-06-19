# Get-LoggedInUser

This function gets the current user sessions on a remote or local computer.

Borrowed from [MrPig91/SysAdminTools](https://github.com/MrPig91/SysAdminTools/wiki/Get%E2%80%90LoggedInUser)

## Use case

Use this script to find stale User Accounts. They can then be deactivted or deleted depending on policy.

## Run

The function can be used in several ways:

For current computer:

```ps
PS C:\> . .\Get-LoggedInUser.ps1
PS C:\> Get-LoggedInUser
```

For current and one remote computer:

```ps
PS C:\> . .\Get-LoggedInUser.ps1
PS C:\> Get-LoggedInUser -ComputerName $env:COMPUTERNAME,dc01v
```

Pipe from Get-ADComputer:

```ps
PS C:\> . .\Get-LoggedInUser.ps1
PS C:\> Get-ADComputer -Filter { OperatingSystem -like "*server*" } | Select-Object -ExpandProperty Name | Get-LoggedInUser
```

## Output

N/A
