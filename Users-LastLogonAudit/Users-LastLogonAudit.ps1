#requires -version 5
<#
.SYNOPSIS
  AD User Accounts Last Logon Audit Tool.

.DESCRIPTION
  This script queries all domain controllers for last logon timestamps of User Accounts
  and creates a report (Users-LastLogonAudit.csv) outlining the results.

.PARAMETER N/A
  N/A

.OUTPUTS Log File
  Log file stored in "Documents\$PSCommandPath\Users-LastLogonAudit.log"

.OUTPUTS Report File
  Report file stored in "Documents\$PSCommandPath\Users-LastLogonAudit.csv"

.NOTES
  Version:        1.2
  Author:         Louis Lawson
  Creation Date:  29/10/2023

.EXAMPLE
  # N/A
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
    # N/A
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

# Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

# Import Modules & Snap-ins
# PSLogging module is required for creating log files
if (-not (Get-Module PSLogging -ListAvailable)) {
    Install-Module PSLogging -Scope CurrentUser -Force
}
Import-Module PSLogging

# ActiveDirectory module is required for getting info on
# domain controller, accounts, etc.
if (-not (Get-Module ActiveDirectory -ListAvailable)) {
    Install-Module ActiveDirectory -Scope CurrentUser -Force
}
Import-Module ActiveDirectory

#----------------------------------------------------------[Declarations]----------------------------------------------------------

# Script Version
$sScriptVersion = "1.0"
# Get the basename of the script
$sScriptName = (Get-Item $PSCommandPath ).Basename

# Script output directory
$sOutputDir = "$(Join-Path -Path $env:userprofile -ChildPath "Documents" | Join-Path -ChildPath $sScriptName)"
# Creates output directory
New-Item -ItemType Directory -Force -Path $sOutputDir -ErrorAction $ErrorActionPreference | Out-Null

# Report file decs
$sReportName = "$sScriptName.csv"
$sReportFile = Join-Path -Path $sOutputDir -ChildPath $sReportName

# Log file decs
$sLogName = "$sScriptName.log"
$sLogFile = Join-Path -Path $sOutputDir -ChildPath $sLogName

# Transcript file decs
$sTranscriptName = "$sScriptName.transcript"
$sTranscriptFile = Join-Path -Path $sOutputDir -ChildPath $sTranscriptName

#-----------------------------------------------------------[Functions]------------------------------------------------------------

Function Get-AllDCNames {
    Param ()

    Begin {
        Write-LogInfo -LogPath $sLogFile -Message "Get-AllDCNames: Getting all domain controllers"
    }

    Process {
        Try {
            $dcNames = Get-ADDomainController -Filter * | Select-Object -ExpandProperty Name | Sort-Object
            return $dcNames
        }

        Catch {
            Write-LogError -LogPath $sLogFile -Message $_.Exception -ExitGracefully
            Break
        }
    }

    End {
        If ($?) {
            Write-LogInfo -LogPath $sLogFile -Message "Get-AllDCNames: Completed Successfully."
            Write-LogInfo -LogPath $sLogFile -Message " "
        }
    }
}

Function Get-AllUsersInDomain {
    Param ()

    Begin {
        Write-LogInfo -LogPath $sLogFile -Message "Get-AllUsersInDomain: Getting all users in domain"
    }

    Process {
        Try {
            $searchBase = Get-ADDomain -Current LoggedOnUser | Select-Object -ExpandProperty DistinguishedName
            $users = Get-ADUser -Filter * -SearchBase $searchBase
            return $users
        }

        Catch {
            Write-LogError -LogPath $sLogFile -Message $_.Exception -ExitGracefully
            Break
        }
    }

    End {
        If ($?) {
            Write-LogInfo -LogPath $sLogFile -Message "Get-AllUsersInDomain: Completed Successfully."
            Write-LogInfo -LogPath $sLogFile -Message " "
        }
    }
}

#-----------------------------------------------------------[Markdown]------------------------------------------------------------

if ($PSVersionTable.PSVersion.Major -ge 6) {
    Show-Markdown -Path "./README.MD"
}
else {
    Write-Host $sScriptName -ForegroundColor Black -BackgroundColor White
    Write-Host "Version: $sScriptVersion" -ForegroundColor Black -BackgroundColor White
    Write-Host "This PowerShell script creates a report of LastLogon for all User Accounts in the current logged on users domain." -ForegroundColor Black -BackgroundColor White
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Start-Transcript -Path $sTranscriptFile | Out-Null
Start-Log -LogPath $sOutputDir -LogName $sLogName -ScriptVersion $sScriptVersion | Out-Null
# SCRIPT START

# Get all domain controllers in domain
$dcNames = Get-AllDCNames

# Get a collection of users in specified Domain base DNctv
$users = Get-AllUsersInDomain

# Hashtable used for splatting for Get-ADUser in loop
$params = @{
    "Properties" = "lastLogon"
}

$Output = foreach ( $user in $users ) {
    Write-Host "Processing"$user.Name"..." -foregroundcolor yellow
    Write-LogInfo -LogPath $sLogFile -Message "Users-LastLogonAudit: Processing $($user.Name)..."
    # Set LDAPFilter to find specific user
    $params.LDAPFilter = "(sAMAccountName=$($user.SamAccountName))"
    # Clear variables
    $latestLogonFT = $latestLogonServer = $latestLogon = $null
    # Iterate every DC name
    foreach ( $dcName in $dcNames ) {
        Write-LogInfo -LogPath $sLogFile -Message "Users-LastLogonAudit: Querying $dcName for $($user.Name)"
        # Query specific DC
        $params.Server = $dcName
        Write-LogInfo -LogPath $sLogFile -Message "Users-LastLogonAudit: Getting lastlogon for $($user.Name)"
        # Get lastLogon attribute (a file time)
        $lastLogonFT = Get-ADUser @params |
        Select-Object -ExpandProperty lastLogon
        Write-LogInfo -LogPath $sLogFile -Message "Users-LastLogonAudit: Comparing lastlogon for $($user.Name)"
        Write-LogInfo -LogPath $sLogFile -Message " "
        # Remember most recent file time and DC name
        if ( $lastLogonFT -and ($lastLogonFT -gt $latestLogonFT) ) {
            $latestLogonFT = $lastLogonFT
            $latestLogonServer = $dcName
        }
    }
    if ( $latestLogonFT -and ($latestLogonFT -gt 0) ) {
        # If user ever logged on, get DateTime from file time
        $latestLogon = [DateTime]::FromFileTime($latestLogonFT)
    }
    else {
        # User never logged on
        $latestLogon = $latestLogonServer = $null
    }
    # Output user
    $user | Select-Object `
        name,
    ObjectClass,
    Enabled,
    @{Name = "LatestLogon"; Expression = { $latestLogon } },
    @{Name = "LatestLogonServer"; Expression = { $latestLogonServer } }
}

Write-LogInfo -LogPath $sLogFile -Message "Users-LastLogonAudit: Writing report file to $sReportFile"
$Output | Export-Csv $sReportFile -NoTypeInformation

Write-Host "Report created in $sReportFile" -ForegroundColor Green

Read-Host -Prompt "Press enter to exit..." | Out-Null

Write-LogInfo -LogPath $sLogFile -Message "Users-LastLogonAudit: Script finished"
# SCRIPT END
Stop-Log -LogPath $sLogFile
Stop-Transcript