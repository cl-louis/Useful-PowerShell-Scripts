#requires -version 5
<#
.SYNOPSIS
  Identify all OUs and their applied or linked GPOs

.DESCRIPTION
  This script gets all OUs in a domain and finds the applied/inherited GPOs.

.PARAMETER N/A
  N/A

.OUTPUTS Log File
  Log file stored in "Documents\OU-GPOApplicationAuditTool\OU-GPOApplicationAuditTool.log"

.OUTPUTS Report File
  Report file stored in "Documents\OU-GPOApplicationAuditTool\OU-GPOApplicationAuditTool.csv"

.NOTES
  Version:        1.0
  Author:         Louis Lawson
  Creation Date:  29/10/2023

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

if (-not (Get-Module ActiveDirectory -ListAvailable)) {
  Install-Module ActiveDirectory -Scope CurrentUser -Force
}
Import-Module ActiveDirectory

if (-not (Get-Module GroupPolicy -ListAvailable)) {
  Install-Module GroupPolicy -Scope CurrentUser -Force
}
Import-Module GroupPolicy

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
  Write-Host "<SCRIPT_DESCRIPTION>" -ForegroundColor Black -BackgroundColor White
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Start-Transcript -Path $sTranscriptFile | Out-Null
Start-Log -LogPath $sOutputDir -LogName $sLogName -ScriptVersion $sScriptVersion | Out-Null
# SCRIPT START

# Get all OUs for the current domain
Write-LogInfo -LogPath $sLogFile -Message "OU-GPOApplicationAuditTool: Finding all OUs for domain"
[array]$AllOrgUnits = Get-ADOrganizationalUnit -Filter * -Properties *

# Setup a hash table and an array that'll be sent to an out file
$OutputCollection = @{} # Hashtable for each OU' data
$OutputData = @() # Array used for return value

ForEach ($OrgUnit in $AllOrgUnits) {
  # A notice to the console that the script is still functioning
  Write-Host "Processing $($OrgUnit.Name)..." -foregroundcolor yellow
  Write-LogInfo -LogPath $sLogFile -Message "OU-GPOApplicationAuditTool: Processing $($OrgUnit.Name)"
  
  # Get-GPInheritance for each OU in the $AllOrgUnits array
  Write-LogInfo -LogPath $sLogFile -Message "OU-GPOApplicationAuditTool: Finding GPOs for $($OrgUnit.DistinguishedName)"
  $GPInheritance = Get-GPInheritance -Target $OrgUnit.DistinguishedName

  # Values from the $OrgUnit
  $OutputCollection.Name = $OrgUnit.Name
  $OutputCollection.CN = $OrgUnit.CanonicalName
  $OutputCollection.DN = $OrgUnit.DistinguishedName
  $OutputCollection.GUID = $OrgUnit.ObjectGUID
  $OutputCollection.Created = $OrgUnit.Created
  $OutputCollection.Modified = $OrgUnit.Modified
  $OutputCollection.DeleteProtected = $OrgUnit.ProtectedFromAccidentalDeletion
  $OutputCollection.isDeleted = $OrgUnit.isDeleted
  $OutputCollection.Deleted = $OrgUnit.Deleted

  # Values from the $GPInheritance object
  $OutputCollection.GpoInheritanceBlocked = $GPInheritance.GpoInheritanceBlocked
  $OutputCollection.GpoLinks = $GPInheritance.GpoLinks
  $OutputCollection.InheritedGpoLinks = $GPInheritance.InheritedGpoLinks

  Write-LogInfo -LogPath $sLogFile -Message "OU-GPOApplicationAuditTool: Adding $($OrgUnit.Name) collection to output array"
  # Create custom PS object to be used to create the csv file
  $OutputData += New-Object PSObject -Property $OutputCollection
}

# Selects 'columns' from $OutputData array.
# GpoLinks and InheritedGpoLinks are processed further to become a comma seperated list of GPO names
Write-LogInfo -LogPath $sLogFile -Message "OU-GPOApplicationAuditTool: Processing collection..."
$Output = $OutputData | Select-Object Name, Cn, DN, GUID, Created, Modified, DeleteProtected, isDeleted, Deleted, GpoInheritanceBlocked, @{name = "GpoLinks" ; Expression = { ($_.GpoLinks | Select-Object DisplayName).DisplayName -join "," } }, @{name = "InheritedGpoLinks" ; Expression = { ($_.InheritedGpoLinks | Select-Object DisplayName).DisplayName -join "," } } | Sort-Object Name

$Output | Export-Csv $sReportFile -NoTypeInformation
Write-LogInfo -LogPath $sLogFile -Message "OU-GPOApplicationAuditTool: Report created in $sReportFile"
Write-Host "OU-GPOApplicationAuditTool: Report created in $sReportFile" -ForegroundColor Green

Read-Host -Prompt "Press enter to exit..." | Out-Null

# SCRIPT END
Stop-Log -LogPath $sLogFile
Stop-Transcript