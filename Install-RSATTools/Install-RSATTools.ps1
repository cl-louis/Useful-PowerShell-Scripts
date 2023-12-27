#requires -version 5
<#
.SYNOPSIS
  Install essential RSAT tools onto a users PC.

.DESCRIPTION
  This script installs: RSAT.ServerManager.Tools, RSAT.ActiveDirectory.DS-LDS.Tools, 
    RSAT.CertificateServices.Tools, RSAT.DHCP.Tools, RSAT.Dns.Tools, and 
    RSAT.GroupPolicy.Management.Tools onto a users PC.
  On run the script will check for admin permissions and automatically 
  elevate if they aren't found. 

.PARAMETER N/A
  N/A

.OUTPUTS Log File
  Log file stored in "Documents\Install-RSATTools\Install-RSATTools.log"

.OUTPUTS Report File
  Report file stored in "Documents\Install-RSATTools\Install-RSATTools.csv"

.NOTES
  Version:        1.0
  Author:         Louis
  Creation Date:  27/12/2023

.EXAMPLE
  N/A
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
  [switch]$Elevated
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

# Tests for elevated powershell window
Function Test-IsElevated {
  Param ()

  Process {
    Try {
      $(New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    }

    Catch {
      Write-LogError -LogPath $sLogFile -Message $_.Exception -ExitGracefully
      Break
    }
  }
}

# Restarts script as admin if it isn't already
if ((Test-IsElevated) -eq $false) {
  if ($elevated) {
    # tried to elevate, did not work, aborting
  }
  else {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-NoProfile -ExecutionPolicy ByPass -NoExit  -File "{0}" -Elevated' -f ($myinvocation.MyCommand.Definition))
  }
  exit
}

# Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

# Import Modules & Snap-ins
if (-not (Get-Module PSLogging -ListAvailable)) {
  Install-Module PSLogging -Scope CurrentUser -Force
}
Import-Module PSLogging

if (-not (Get-Module DISM -ListAvailable)) {
  Install-Module DISM -Scope CurrentUser -Force
}
Import-Module DISM

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

#-----------------------------------------------------------[Markdown]------------------------------------------------------------

if ($PSVersionTable.PSVersion.Major -ge 6) {
  Show-Markdown -Path "./README.MD"
}
else {
  Write-Host $sScriptName -ForegroundColor Black -BackgroundColor White
  Write-Host "Version: $sScriptVersion" -ForegroundColor Black -BackgroundColor White
  Write-Host "This script installs: RSAT.ServerManager.Tools, RSAT.ActiveDirectory.DS-LDS.Tools, RSAT.CertificateServices.Tools, RSAT.DHCP.Tools, RSAT.Dns.Tools, and RSAT.GroupPolicy.Management.Tools onto a users PC." -ForegroundColor Black -BackgroundColor White
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Start-Transcript -Path $sTranscriptFile | Out-Null
Start-Log -LogPath $sOutputDir -LogName $sLogName -ScriptVersion $sScriptVersion | Out-Null
# SCRIPT START

@(
  'RSAT.ServerManager.Tools*' 
  'RSAT.ActiveDirectory.DS-LDS.Tools*' 
  'RSAT.CertificateServices.Tools*' 
  'RSAT.DHCP.Tools*'
  'RSAT.Dns.Tools*' 
  'RSAT.GroupPolicy.Management.Tools*' 
) | ForEach-Object { Get-WindowsCapability -Name $_ -Online } | ForEach-Object { Add-WindowsCapability -Name $_.Name -Online }

Get-WindowsCapability -Name 'RSAT.*' -Online | Where-Object State -eq 'Installed' | Sort-Object State, Name | Format-Table Name, State, DisplayName -AutoSize

Read-Host -Prompt "Press enter to exit..." | Out-Null

# SCRIPT END
Stop-Log -LogPath $sLogFile
Stop-Transcript