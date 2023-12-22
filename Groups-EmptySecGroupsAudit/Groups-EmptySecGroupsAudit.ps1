#requires -version 5
<#
.SYNOPSIS
  <Overview of script>

.DESCRIPTION
  <Brief description of script>

.PARAMETER <Parameter_Name>
  <Brief description of parameter. Repeat this attribute if required>

.INPUTS <Input_Name>
  <Brief description of input. Repeat this attribute if required>

.OUTPUTS <Output_Name>
  <Brief description of output. Repeat this attribute if required>

.NOTES
  Version:        1.0
  Author:         <Name>
  Creation Date:  <Date>

.EXAMPLE
  <Example explanation goes here>
  
  <Example goes here. Repeat this attribute for more than one example>
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
    # Script parameters go here
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
    Write-Host "Groups-EmptySecGroupsAudit" -ForegroundColor Black -BackgroundColor White
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Start-Transcript -Path $sTranscriptFile | Out-Null
Start-Log -LogPath $sOutputDir -LogName $sLogName -ScriptVersion $sScriptVersion | Out-Null
# SCRIPT START

Write-LogInfo -LogPath $sLogFile -Message "Querying DC for groups."
Write-Host "Querying DC for groups." -foregroundcolor yellow
Write-Host " "
# Use Get-ADGroup with a Filter to only return security groups
$Output = Get-ADGroup -Filter { GroupCategory -eq "Security" } -Properties Name, DistinguishedName, GroupScope, Members, isCriticalSystemObject, Description | 
Where-Object { !($_.IsCriticalSystemObject) -and $_.Members.count -eq 0 } | # Filter the returned Security Groups for only Non System Critical Groups (i.e. Builtin) with a member count of 0
Select-Object Name, DistinguishedName, GroupScope, Description # Select only the Name, DistinguishedName, GroupScope, Description

Write-Host "Found $($Output.Length) potentially empty groups"
Write-Host " "
Write-LogInfo -LogPath $sLogFile -Message "Found $($Output.Length) potentially empty groups"
Write-LogInfo -LogPath $sLogFile -Message "Groups-EmptySecGroupsAudit: Writing report file to $sReportFile"
$Output | Export-Csv $sReportFile -NoTypeInformation

Write-Host "Report created in $sReportFile" -ForegroundColor Green

Read-Host -Prompt "Press enter to exit..." | Out-Null

# SCRIPT END
Stop-Log -LogPath $sLogFile
Stop-Transcript