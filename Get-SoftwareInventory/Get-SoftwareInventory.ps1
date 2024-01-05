#requires -version 5
<#
.SYNOPSIS
  Gather installed software packages on the local computer

.DESCRIPTION
  This script gathers a list of installed software packages on the local PC and outputs them to a CSV for auditing

.PARAMETER N/A
  N/A

.OUTPUTS Log File
  Log file stored in "Documents\Get-SoftwareInventory\Get-SoftwareInventory.log"

.OUTPUTS Transcript File
  Transcript file stored in "Documents\Get-SoftwareInventory\Get-SoftwareInventory.transcript"

.OUTPUTS Report File
  Report file stored in "Documents\Get-SoftwareInventory\Get-SoftwareInventory.csv"

.NOTES
  Version:        1.0
  Author:         Louis Lawson
  Creation Date:  02/01/2024

.EXAMPLE
  N/A
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
  # Script parameters go here
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

# Set Error Action to Silently Continue
$ErrorActionPreference = "Stop"

# Import Modules & Snap-ins
if (-not (Get-Module PSLogging -ListAvailable)) {
  Install-Module PSLogging -Scope CurrentUser -Force
}
Import-Module PSLogging

<# 
if (-not (Get-Module <MODULE_NAME> -ListAvailable)) {
  Install-Module <MODULE_NAME> -Scope CurrentUser -Force
}
Import-Module <MODULE_NAME>
#>

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

Function Get-SoftwareFromRegistry {
  Param (
    #List of ComputerNames to process
    [Parameter(ValueFromPipeline = $true,
      ValueFromPipelineByPropertyName = $true)]
    [string[]]
    $ComputerName
  )

  Begin {
    Write-LogInfo -LogPath $sLogFile -Message "Get-SoftwareFromRegistry: Param -ComputerName $ComputerName"
    $SoftwareArray = @()
  }

  Process {
    Try {
      #Variable to hold the location of Currently Installed Programs
      $SoftwareRegKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"

      Write-LogInfo -LogPath $sLogFile -Message "Creating an instance of Registry Object for LocalMachine"
      #Create an instance of the Registry Object and open the HKLM base key
      $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)

      Write-LogInfo -LogPath $sLogFile -Message "Opening the Uninstall subkey"
      #Open the Uninstall subkey using the OpenSubKey Method
      $RegKey = $Reg.OpenSubKey($SoftwareRegKey)

      Write-LogInfo -LogPath $sLogFile -Message "Creating array of subkey names"
      #Create a string array containing all the subkey names
      [String[]]$SubKeys = $RegKey.GetSubKeyNames()

      Write-LogInfo -LogPath $sLogFile -Message "Iterating over subkeys"
      #Open each Subkey and use its GetValue method to return the required values
      foreach ($key in $SubKeys) {
        $UninstallKey = $SoftwareRegKey + "\\" + $key
        $UninstallSubKey = $reg.OpenSubKey($UninstallKey)
        $obj = [PSCustomObject]@{
          DisplayName     = $($UninstallSubKey.GetValue("DisplayName"))
          DisplayVersion  = $($UninstallSubKey.GetValue("DisplayVersion"))
          InstallLocation = $($UninstallSubKey.GetValue("InstallLocation"))
          Publisher       = $($UninstallSubKey.GetValue("Publisher"))
        }
        $SoftwareArray += $obj
      } 
    }

    Catch {
      Write-LogError -LogPath $sLogFile -Message $_.Exception -ExitGracefully
      Break
    }
  }

  End {
    return $SoftwareArray | Where-Object { $_.DisplayName } | Sort-Object -Property DisplayName | Select-Object DisplayName, DisplayVersion, Publisher
    If ($?) {
      Write-LogInfo -LogPath $sLogFile -Message "Get-SoftwareFromRegistry: Completed Successfully."
      Write-LogInfo -LogPath $sLogFile -Message " "
    }
  }
}

Function Show-OptionMenu {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'i', 
    Justification = 'False positive for $i. See https://github.com/PowerShell/PSScriptAnalyzer/issues/1354')]
  Param (
    [string]$Title = 'Option Menu',
    [Parameter(Mandatory = $true)][string[]]$Choices
  )

  Begin {
    Write-LogInfo -LogPath $sLogFile -Message "Show-OptionMenu: Param -Title $Title"
    Write-LogInfo -LogPath $sLogFile -Message "Get-PasswordPart: Param -Choices $($Choices.count)"
  }

  Process {
    Try {
      Write-Host "================ $Title ================"
      $Menu = @{}

      $Choices | ForEach-Object -Begin { $i = 1 } { 
        Write-Host "$_`: Press '$i' for this option." 
        $Menu.add("$i", $_)
        $i++
      }

      Write-Host "Q: Press 'Q' to quit."

      $Selection = Read-Host "Please make a selection"

      if ($Selection -eq 'Q') { Break } Else { Return $Menu.$Selection }
    }

    Catch {
      Write-LogError -LogPath $sLogFile -Message $_.Exception -ExitGracefully
      Break
    }
  }

  End {
    If ($?) {
      Write-LogInfo -LogPath $sLogFile -Message "Show-OptionMenu: Completed Successfully."
      Write-LogInfo -LogPath $sLogFile -Message " "
    }
  }
}

<#
Function <FunctionName> {
  Param ()

  Begin {
    Write-LogInfo -LogPath $sLogFile -Message "<description of what is going on>..."
  }

  Process {
    Try {
      <code goes here>
    }

    Catch {
      Write-LogError -LogPath $sLogFile -Message $_.Exception -ExitGracefully
      Break
    }
  }

  End {
    If ($?) {
      Write-LogInfo -LogPath $sLogFile -Message "Completed Successfully."
      Write-LogInfo -LogPath $sLogFile -Message " "
    }
  }
}
#>

#-----------------------------------------------------------[Markdown]------------------------------------------------------------

if ($PSVersionTable.PSVersion.Major -ge 6) {
  Show-Markdown -Path "./README.MD"
}
else {
  Write-Host $sScriptName -ForegroundColor Black -BackgroundColor White
  Write-Host "Version: $sScriptVersion" -ForegroundColor Black -BackgroundColor White
  Write-Host "Gather installed software packages on the local computer" -ForegroundColor Black -BackgroundColor White
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Start-Transcript -Path $sTranscriptFile | Out-Null
Start-Log -LogPath $sOutputDir -LogName $sLogName -ScriptVersion $sScriptVersion | Out-Null
# SCRIPT START

# $UserSelection = Show-OptionMenu -Title 'Computer to inventory' -Choices @('Local Computer', 'Remote Computer')

# if ($UserSelection -eq "Local Computer") {
#   Write-Host "Finding installed software on local computer: $env:COMPUTERNAME"
#   $UserSelection = $env:COMPUTERNAME
# }
# elseif ($UserSelection -eq "Remote Computer") {
#   $UserSelection = Read-Host -Prompt "Enter remote PC name"
#   Write-Host "Finding installed software on local computer: $UserSelection"
# }

Write-Host "Getting software from $($env:COMPUTERNAME)" -ForegroundColor Yellow

$SoftwareFromReg = Get-SoftwareFromRegistry -ComputerName $env:COMPUTERNAME

Write-Host "Found $($SoftwareFromReg.count) installed software packages"

Write-LogInfo -LogPath $sLogFile -Message "Users-LastLogonAudit: Writing report file to $sReportFile"
$SoftwareFromReg | Export-Csv $sReportFile -NoTypeInformation

Write-Host "Report created in $sReportFile" -ForegroundColor Green

Read-Host -Prompt "Press enter to exit..." | Out-Null

Write-LogInfo -LogPath $sLogFile -Message "Users-LastLogonAudit: Script finished"

# SCRIPT END
Stop-Log -LogPath $sLogFile
Stop-Transcript