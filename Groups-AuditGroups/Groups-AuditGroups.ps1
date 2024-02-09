#requires -version 5
<#
.SYNOPSIS
  This PowerShell script creates a report of all Security Groups

.DESCRIPTION
  This PowerShell script creates a report of all Security Groups.
  Use this script to audit Security Groups members, names, and uses.

.PARAMETER N/A
  N/A

.OUTPUTS Log File
  Log file stored in "Documents\Groups-AuditGroups\Groups-AuditGroups.log"

.OUTPUTS Transcript File
  Log file stored in "Documents\Groups-AuditGroups\Groups-AuditGroups.transcript"

.OUTPUTS Report File
  Report file stored in "Documents\Groups-AuditGroups\Groups-AuditGroups.csv"

.NOTES
  Version:        1.0
  Author:         Louis Lawson
  Creation Date:  09/02/2023

.EXAMPLE
  N/A
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

# ActiveDirectory module is required for group info
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

#-----------------------------------------------------------[Markdown]------------------------------------------------------------

if ($PSVersionTable.PSVersion.Major -ge 6) {
  Show-Markdown -Path "./README.MD"
}
else {
  Write-Host $sScriptName -ForegroundColor Black -BackgroundColor White
  Write-Host "Version: $sScriptVersion" -ForegroundColor Black -BackgroundColor White
  Write-Host "This PowerShell script creates a report of all Security Groups." -ForegroundColor Black -BackgroundColor White
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Start-Transcript -Path $sTranscriptFile | Out-Null
Start-Log -LogPath $sOutputDir -LogName $sLogName -ScriptVersion $sScriptVersion | Out-Null
# SCRIPT START

Write-LogInfo -LogPath $sLogFile -Message "Querying DC for groups."
Write-Host "Querying DC for groups." -foregroundcolor yellow
Write-Host " "
# Use Get-ADGroup with a Filter to only return security groups
$Output = Get-ADGroup -Filter "*" -Properties Name, DistinguishedName, GroupScope, Members, isCriticalSystemObject, Description | 
Select-Object Name, GroupScope, @{Name = "Critical Object"; Expression = { $_.isCriticalSystemObject } }, @{Name = "Member Count"; Expression = { $_.Members.count } }, Description, @{Name = "Members"; Expression = { (Get-ADGroupMember -Identity "$($_.Name)" | Select-Object -ExpandProperty Name) -Join ", " } }

Write-Host "Found $($Output.Length) groups"
Write-Host " "
Write-LogInfo -LogPath $sLogFile -Message "Found $($Output.Length) groups"
Write-LogInfo -LogPath $sLogFile -Message "Groups-AuditGroups: Writing report file to $sReportFile"
$Output | Export-Csv $sReportFile -NoTypeInformation

Write-Host "Report created in $sReportFile" -ForegroundColor Green

Read-Host -Prompt "Press enter to exit..." | Out-Null

# SCRIPT END
Stop-Log -LogPath $sLogFile
Stop-Transcript