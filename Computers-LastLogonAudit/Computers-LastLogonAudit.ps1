#requires -version 5
<#
.SYNOPSIS
  AD Computer Accounts Last Logon Audit Tool.

.DESCRIPTION
  This script queries all domain controllers for last logon timestamps of Computer Accounts
  and creates a report (Computers-LastLogonAudit.csv) outlining the results.

.PARAMETER N/A
  N/A

.OUTPUTS Log File
  Log file stored in "Documents\$PSCommandPath\Computers-LastLogonAudit.log"

.OUTPUTS Report File
  Report file stored in "Documents\$PSCommandPath\Computers-LastLogonAudit.csv"

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
$sScriptVersion = "1.2"
# Get the basename of the script
$sScriptName = (Get-Item $PSCommandPath ).Basename

# Script output directory
$sOutputDir = Join-Path -Path $env:userprofile -ChildPath "Documents" -AdditionalChildPath $sScriptName
# Creates output directory
New-Item -ItemType Directory -Force -Path $sOutputDir -ErrorAction $ErrorActionPreference | Out-Null

# Report file decs
$sReportName = "$sScriptName.csv"
$sReportFile = Join-Path -Path $sOutputDir -ChildPath $sReportName

# Log file decs
$sLogName = "$sScriptName.log"
$sLogFile = Join-Path -Path $sOutputDir -ChildPath $sLogName

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

Function Get-AllComputersInDomain {
  Param ()

  Begin {
    Write-LogInfo -LogPath $sLogFile -Message "Get-AllComputersInDomain: Getting all computers in domain"
  }

  Process {
    Try {
      $searchBase = Get-ADDomain -Current LoggedOnUser | Select-Object -ExpandProperty DistinguishedName
      $computers = Get-ADComputer -Filter * -SearchBase $searchBase
      return $computers
    }

    Catch {
      Write-LogError -LogPath $sLogFile -Message $_.Exception -ExitGracefully
      Break
    }
  }

  End {
    If ($?) {
      Write-LogInfo -LogPath $sLogFile -Message "Get-AllComputersInDomain: Completed Successfully."
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
  Write-Host "This PowerShell script creates a report of LastLogon for all Computer Accounts in the current logged on users domain." -ForegroundColor Black -BackgroundColor White
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Start-Log -LogPath $sOutputDir -LogName $sLogName -ScriptVersion $sScriptVersion
# SCRIPT START

# Get all domain controllers in domain
$dcNames = Get-AllDCNames

# Get a collection of computers in specified Domain base DNctv
$computers = Get-AllComputersInDomain

# Hashtable used for splatting for Get-ADComputer in loop
$params = @{
  "Properties" = "lastLogon"
}

$Output = foreach ( $computer in $computers ) {
  Write-Host "Processing"$computer.Name"..." -foregroundcolor yellow
  Write-LogInfo -LogPath $sLogFile -Message "Computers-LastLogonAudit: Processing $($computer.Name)..."
  # Set LDAPFilter to find specific user
  $params.LDAPFilter = "(sAMAccountName=$($computer.SamAccountName))"
  # Clear variables
  $latestLogonFT = $latestLogonServer = $latestLogon = $null
  # Iterate every DC name
  foreach ( $dcName in $dcNames ) {
    Write-LogInfo -LogPath $sLogFile -Message "Computers-LastLogonAudit: Querying $dcName for $($computer.Name)"
    # Query specific DC
    $params.Server = $dcName
    Write-LogInfo -LogPath $sLogFile -Message "Computers-LastLogonAudit: Getting lastlogon for $($computer.Name)"
    # Get lastLogon attribute (a file time)
    $lastLogonFT = Get-ADComputer @params |
    Select-Object -ExpandProperty lastLogon
    Write-LogInfo -LogPath $sLogFile -Message "Computers-LastLogonAudit: Comparing lastlogon for $($computer.Name)"
    Write-LogInfo -LogPath $sLogFile -Message " "
    # Remember most recent file time and DC name
    if ( $lastLogonFT -and ($lastLogonFT -gt $latestLogonFT) ) {
      $latestLogonFT = $lastLogonFT
      $latestLogonServer = $dcName
    }
  }
  if ( $latestLogonFT -and ($latestLogonFT -gt 0) ) {
    # If computer ever logged on, get DateTime from file time
    $latestLogon = [DateTime]::FromFileTime($latestLogonFT)
  }
  else {
    # Computer never logged on
    $latestLogon = $latestLogonServer = $null
  }
  # Output computer
  $computer | Select-Object `
    name,
  ObjectClass,
  Enabled,
  @{Name = "LatestLogon"; Expression = { $latestLogon } },
  @{Name = "LatestLogonServer"; Expression = { $latestLogonServer } }
}

Write-LogInfo -LogPath $sLogFile -Message "Computers-LastLogonAudit: Writing report file to $sReportFile"
$Output | Export-Csv $sReportFile -Encoding UTF8 -NoTypeInformation

Write-LogInfo -LogPath $sLogFile -Message "Computers-LastLogonAudit: Script finished"
# SCRIPT END
Stop-Log -LogPath $sLogFile
Pause
Exit