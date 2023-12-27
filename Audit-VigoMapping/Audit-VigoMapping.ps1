#Requires -Version 5
<#
.SYNOPSIS
  VigoGLogon shortcut and Z drive mapping audit tool.

.DESCRIPTION
  This script audits the mapping of Z drive to ensure it uses the correct UNC path, then
  searches the users desktop for the VigoGLogon shortcut to ensure it has been created from
  a valid drive mapping.

  This script can be ran to ensure that issues that arrise from incorrect drive mappings
  and VigoGLogon shortcuts can be avoided before they create issues for other Vigo users.

.PARAMETER N/A
  N/A

.OUTPUTS Log File
  Log file stored in "Documents\Audit-VigoMapping\Audit-VigoMapping.log"

.OUTPUTS Transcript File
  Transcript file stored in "Documents\Audit-VigoMapping\Audit-VigoMapping.transcript"

.NOTES
  Version:        1.0
  Author:         Louis Lawson
  Creation Date:  21/12/2023

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
if (-not (Get-Module PSLogging -ListAvailable)) {
  Install-Module PSLogging -Scope CurrentUser -Force
}
Import-Module PSLogging

if (-not (Get-Module SmbShare -ListAvailable)) {
  Install-Module SmbShare -Scope CurrentUser -Force
}
Import-Module SmbShare

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

Function Test-DriveMappingRemotePath {
  Param (
    [Parameter(Mandatory = $true, HelpMessage = "Drive letter to test. Requires suffixing with a colon.")][string]$Drive, 
    [Parameter(Mandatory = $true, HelpMessage="Remote path to test, usually a UNC path.")][string]$RemotePath
  )

  Begin {
    Write-LogInfo -LogPath $sLogFile -Message "Test-DriveMapping: Param -Drive $Drive"
    Write-LogInfo -LogPath $sLogFile -Message "Test-DriveMapping: Param -RemotePath $RemotePath"
    
  }

  Process {
    Try {
      Write-LogInfo -LogPath $sLogFile -Message "Test-DriveMapping: Testing mapping of $Drive drive"
      $DriveRemotePath = Get-SmbMapping -LocalPath "$Drive" | Select-Object -ExpandProperty RemotePath
      
      if ("$DriveRemotePath" -eq "$RemotePath") {
        Write-LogInfo -LogPath $sLogFile -Message "Test-DriveMapping: Valid drive mapping found"
        return $true
      }
      else {
        Write-LogInfo -LogPath $sLogFile -Message "Test-DriveMapping: Invalid drive mapping found"
        return $false
      }
    }

    Catch {
      Write-LogError -LogPath $sLogFile -Message $_.Exception -ExitGracefully
      Break
    }
  }

  End {
    If ($?) {
      Write-LogInfo -LogPath $sLogFile -Message "Test-DriveMapping: Completed Successfully."
      Write-LogInfo -LogPath $sLogFile -Message " "
    }
  }
}

Function Get-ItemFromDesktop {
  Param ([Parameter(Mandatory = $true)][string]$Search)

  Begin {
    Write-LogInfo -LogPath $sLogFile -Message "Get-ItemFromDesktop: Param Search $Search"
  }

  Process {
    Try {
      Write-LogInfo -LogPath $sLogFile -Message "Get-ItemFromDesktop: Searching users 'Desktop' for $Search"
      $results = @(Get-ChildItem -Path $env:userprofile\Desktop\* -Filter "$Search" | Sort-Object -Property Name)
      Write-LogInfo -LogPath $sLogFile -Message "Found $($results.Length) items"
      return $results
    }

    Catch {
      Write-LogError -LogPath $sLogFile -Message $_.Exception -ExitGracefully
      Break
    }
  }

  End {
    If ($?) {
      Write-LogInfo -LogPath $sLogFile -Message "Get-ItemFromDesktop: Completed Successfully."
      Write-LogInfo -LogPath $sLogFile -Message " "
    }
  }
}

#-----------------------------------------------------------[Markdown]------------------------------------------------------------

if ($PSVersionTable.PSVersion.Major -ge 6) {
  Show-Markdown -Path "README.MD"
}
else {
  Write-Host $sScriptName -ForegroundColor Black -BackgroundColor White
  Write-Host "Version: $sScriptVersion" -ForegroundColor Black -BackgroundColor White
  Write-Host "<SCRIPT_DESCRIPTION>" -ForegroundColor Black -BackgroundColor White
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------
Start-Transcript -Path $sTranscriptFile | Out-Null
Start-Log -LogPath $sOutputDir -LogName $sLogName -ScriptVersion $sScriptVersion | Out-Null
# SCRIPT START

# Check drive mapping
Write-Host " "
Write-Host "Checking Z drive mapping" -ForegroundColor Yellow
Write-Host " "

if (Test-DriveMappingRemotePath -Drive "Z:" -RemotePath "\\CL-FSS01.jupiter.apc.local\VigoETASQL") {
  Write-Host "Drive is mapped correctly" -ForegroundColor Green
}
else {
  Write-Host "Incorrect Z Drive mapping" -ForegroundColor Red
  Write-Host "Please ensure drive is mapped as '\\CL-FSS01.jupiter.apc.local\VigoETASQL'"
  Write-Host "Disconnect Z drive and run 'gpupdate /force' in a new cmd window"
  Write-Host "See https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/gpupdate"
}

# A scope modifier is used to surpress warning on the $idx variable
# See https://github.com/PowerShell/PSScriptAnalyzer/issues/1641
# Docs on scopes: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_scopes?view=powershell-7.4#scope-modifiers
Write-Host " "
Write-Host "Searching for VigoGLogon shortcut on desktop" -ForegroundColor Yellow
Write-Host " "
$results = @(Get-ChildItem -Path $env:userprofile\Desktop\* -Filter "Vigo*.lnk" | Sort-Object -Property Name)
Write-Host "Found $($results.Length) items" -ForegroundColor Yellow
Write-LogInfo -LogPath $sLogFile "Found $($results.Length) items"
$results | ForEach-Object -Begin { $script:idx = 0 } -Process { 
  ++$idx
  $shell = New-Object -ComObject WScript.Shell
  Write-LogInfo -LogPath $sLogFile -Message "Processing $($_.Name)"
  Write-Host "Processing $($_.Name)"
  $targetPath = $shell.CreateShortcut($_).TargetPath
  Write-Host "Shortcut targets '$targetPath'"
  Write-LogInfo -LogPath $sLogFile "Shortcut targets '$targetPath'"
  if ($targetPath -ne "Z:\Haulage\VigoGLogon.exe") {
    Write-LogInfo -LogPath $sLogFile "Invalid shortcut found."
    Write-LogInfo -LogPath $sLogFile "Remove the shortcut $($_.Name) to prevent Vigo issues."
    Write-Host "Invalid shortcut found." -ForegroundColor Red
    Write-Host "Remove this shortcut to prevent Vigo issues."
    Write-Host "Ensure shortcuts are created from the correctly mapped Z Drive."
  }
  else {
    Write-Host "Valid shortcut found" -ForegroundColor Green
    Write-LogInfo -LogPath $sLogFile "Valid shortcut found."
  }
  Write-Host " "
}

Read-Host -Prompt "Press enter to exit..." | Out-Null

# SCRIPT END
Stop-Log -LogPath $sLogFile
Stop-Transcript