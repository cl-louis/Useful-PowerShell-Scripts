#requires -version 5
<#
.SYNOPSIS
  Set a domain User Account to a secure password.

.DESCRIPTION
  This script allows a user to search for a domain account and then generates a secure password for the account. 
  The password is set on the PDC of the local PCs domain.

  Passwords are constructed using 3 files: Words.txt, Numbers.txt, and SpecialCharacters.txt.
  By default, the password uses 3 items from Words.txt, 3 from Numbers.txt, and 1 from SpecialCharacters.txt

  An Example password could be: ForeverMagicCarpets821!

.PARAMETER Words
  A path, absolute or relative, to CRLF delimited list of words used to generate passwords.

.PARAMETER Numbers
  A path, absolute or relative, to CRLF delimited list of numbers used to generate passwords.

.PARAMETER SpecialCharacters
  A path, absolute or relative, to CRLF delimited list of special characters used to generate passwords.

.OUTPUTS Log File
  Log file stored in "Documents\Set-UserPassword\Set-UserPassword.log"

.OUTPUTS Transcript File
  Transcript file stored in "Documents\Set-UserPassword\Set-UserPassword.transcript"

.NOTES
  Version:        1.0
  Author:         Louis Lawson
  Creation Date:  27/12/2023

.EXAMPLE
  N/A
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
  [Parameter(Mandatory = $false, HelpMessage = "Path to Words.txt, used to generate passwords. Default 'Words.txt' in current directory")][string]$Words = "Words.txt",
  [Parameter(Mandatory = $false, HelpMessage = "Path to Numbers.txt, used to generate passwords. Default 'Numbers.txt' in current directory")][string]$Numbers = "Numbers.txt",
  [Parameter(Mandatory = $false, HelpMessage = "Path to SpecialCharacters.txt, used to generate passwords. Default 'SpecialCharacters.txt' in current directory")][string]$SpecialCharacters = "SpecialCharacters.txt"
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

# Set Error Action to Silently Continue
$ErrorActionPreference = "Stop"

# Import Modules & Snap-ins
if (-not (Get-Module PSLogging -ListAvailable)) {
  Install-Module PSLogging -Scope CurrentUser -Force
}
Import-Module PSLogging

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

# Log file decs
$sLogName = "$sScriptName.log"
$sLogFile = Join-Path -Path $sOutputDir -ChildPath $sLogName

# Transcript file decs
$sTranscriptName = "$sScriptName.transcript"
$sTranscriptFile = Join-Path -Path $sOutputDir -ChildPath $sTranscriptName

#-----------------------------------------------------------[Functions]------------------------------------------------------------

Function Import-TxtFile {
  Param (
    [Parameter(Mandatory = $true, HelpMessage = "Path, relative or absolute, to txt file.")][string]$TxtPath
  )

  Begin {
    Write-LogInfo -LogPath $sLogFile -Message "Import-TxtFile: Param -TxtPath $TxtPath"
  }

  Process {
    Try {
      Write-LogInfo -LogPath $sLogFile -Message "Import-TxtFile: Getting contents of $TxtPath"
      $TxtContents = Get-Content -Path "$TxtPath"
      if ($TxtContents.Length -eq 0) {
        Write-LogError -LogPath $sLogFile -Message "Import-TxtFile: Missing, invalid or empty TxtFile"
        Write-Host "Import-TxtFile: Missing, invalid or empty TxtFile" -ForegroundColor Red
        Throw "Import-TxtFile: Missing, invalid or empty TxtFile"
      }
      return $TxtContents
    }

    Catch {
      Write-LogError -LogPath $sLogFile -Message $_.Exception -ExitGracefully
      Break
    }
  }

  End {
    If ($?) {
      Write-LogInfo -LogPath $sLogFile -Message "Import-TxtFile: Completed Successfully."
      Write-LogInfo -LogPath $sLogFile -Message " "
    }
  }
}

Function Get-PasswordPart {
  Param (
    [Parameter(Mandatory = $true)][string[]]$AllChoices,
    [Parameter(Mandatory = $true)][int]$PartQuantity
  )

  Begin {
    Write-LogInfo -LogPath $sLogFile -Message "Get-PasswordPart: Param -AllChoices $($AllChoices.count)"
    Write-LogInfo -LogPath $sLogFile -Message "Get-PasswordPart: Param -PartQuantity $PartQuantity"
  }

  Process {
    Try {
      Write-LogInfo -LogPath $sLogFile -Message "Get-PasswordPart: Init empty hashtable"
      $Hash = @{}

      Write-LogInfo -LogPath $sLogFile -Message "Get-PasswordPart: Enumerating choices"
      $AllChoices | ForEach-Object {
        $Hash.add($_, (Get-Random -Maximum $AllChoices.count))
      }

      Write-LogInfo -LogPath $sLogFile -Message "Get-PasswordPart: Getting $PartQuantity part(s) from hashtable"
      $PasswordParts = $Hash.GetEnumerator() | Sort-Object -Property value | Select-Object -First $PartQuantity | Select-Object -ExpandProperty Name

      Return $PasswordParts
    }

    Catch {
      Write-LogError -LogPath $sLogFile -Message $_.Exception -ExitGracefully
      Break
    }
  }

  End {
    If ($?) {
      Write-LogInfo -LogPath $sLogFile -Message "Get-PasswordPart: Completed Successfully."
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

      if ($Selection -eq 'Q') { Return } Else { $Menu.$Selection }
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
  Write-Host "Search for a user account, then create and set a secure password" -ForegroundColor Black -BackgroundColor White
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Start-Transcript -Path $sTranscriptFile | Out-Null
Start-Log -LogPath $sOutputDir -LogName $sLogName -ScriptVersion $sScriptVersion | Out-Null
# SCRIPT START


# if (-not($PSBoundParameters.ContainsKey('Words')) -and $Words) {
#   Write-Host "Using default value for Words param"
# }
# else {
#   Write-Host "User supplied value for Words param"
# }

$PossibleWords = Import-TxtFile -TxtPath "$Words"
$PossibleNumbers = Import-TxtFile -TxtPath "$Numbers"
$PossibleSpecialCharacters = Import-TxtFile -TxtPath "$SpecialCharacters"
# Write-Host "$($env:LOGONSERVER.replace('\\',''))"

# $PDCEmulator = Get-ADDomainController -Filter "$($env:LOGONSERVER.replace('\\',''))" | Select-Object 

# $UserSelection = Show-OptionMenu -Title 'Domain Controllers' -Choices (Get-ADDomainController -Filter * | Select-Object HostName).HostName

# Write-Host "$UserSelection"

$UserName = Read-Host -Prompt 'Enter user name'
Write-LogInfo -LogPath $sLogFile -Message "Searching Domain for $UserName"
$ADUserResult = @(Get-ADUser -Filter "Name -like '*$UserName*'" | Select-Object Name, SamAccountName)
if ($ADUserResult.count -eq 0) {
  Write-LogInfo -LogPath $sLogFile -Message "No users found for query $UserName"
  Write-Host "No users found for query $UserName" -ForegroundColor Red
}
else {
  Write-LogInfo -LogPath $sLogFile -Message "Found $($ADUserResult.Count) users"
  Write-LogInfo -LogPath $sLogFile -Message "Prompting for user choice"
  Write-Host "Found $($ADUserResult.Count) users" -ForegroundColor Yellow
  $User = $ADUserResult | Out-GridView -OutputMode Single -Title "Select a user."
}

Write-Host " "

Write-LogInfo -LogPath $sLogFile -Message "Starting password generation"
$pswdWords = Get-PasswordPart -AllChoices $PossibleWords -PartQuantity 3
$pswdNumbers = Get-PasswordPart -AllChoices $PossibleNumbers -PartQuantity 3
$pswdSpecialCharacters = Get-PasswordPart -AllChoices $PossibleSpecialCharacters -PartQuantity 1

$TextInfo = (Get-Culture).TextInfo
$pswd = ($TextInfo.ToTitleCase("$pswdWords")) + "$pswdNumbers" + "$pswdSpecialCharacters" -replace '\s', ''
Write-LogInfo -LogPath $sLogFile -Message "Secure password created"

Try {
  Write-LogInfo -LogPath $sLogFile -Message "Attempting to set user password on PDC"
  Set-ADAccountPassword -Identity ($User.SamAccountName) -Server "$($(Get-ADDomain | Select-Object -Property PDCEmulator).PDCEmulator)" -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $pswd -Force)

  Write-Host "Password set successfully" -ForegroundColor Green -BackgroundColor White
  Write-LogInfo -LogPath $sLogFile -Message "Password set successfully"
  Write-Host "User: $($User.SamAccountName)"
  Write-Host "Password: $pswd"
}
Catch {
  Write-Host "Error Changing Password" -ForegroundColor White -BackgroundColor Red
  Write-LogError -LogPath $sLogFile -Message $_.Exception -ExitGracefully
}

Read-Host -Prompt "Press enter to exit..." | Out-Null

# SCRIPT END
Stop-Log -LogPath $sLogFile
Stop-Transcript