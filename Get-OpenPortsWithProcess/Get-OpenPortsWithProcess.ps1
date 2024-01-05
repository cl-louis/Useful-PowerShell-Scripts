#requires -version 5
<#
.SYNOPSIS
  Find open/listening TCP and UDP ports on the local computer.

.DESCRIPTION
  This script finds open/listening TCP and UDP ports on the local computer with the associated process.

.PARAMETER N/A
  N/A

.OUTPUTS Log File
  Log file stored in "Documents\Get-OpenPortsWithProcess\Get-OpenPortsWithProcess.log"

.OUTPUTS Transcript File
  Transcript file stored in "Documents\Get-OpenPortsWithProcess\Get-OpenPortsWithProcess.transcript"

.OUTPUTS Report File
  Report file stored in "Documents\Get-OpenPortsWithProcess\Get-OpenPortsWithProcess.csv"

.NOTES
  Version:        1.0
  Author:         Louis Lawson
  Creation Date:  05/01/2024

.EXAMPLE
  N/A
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

#-----------------------------------------------------------[Markdown]------------------------------------------------------------

if ($PSVersionTable.PSVersion.Major -ge 6) {
  Show-Markdown -Path "./README.MD"
}
else {
  Write-Host $sScriptName -ForegroundColor Black -BackgroundColor White
  Write-Host "Version: $sScriptVersion" -ForegroundColor Black -BackgroundColor White
  Write-Host "his script finds open/listening TCP and UDP ports on the local computer with the associated process" -ForegroundColor Black -BackgroundColor White
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Start-Transcript -Path $sTranscriptFile | Out-Null
Start-Log -LogPath $sOutputDir -LogName $sLogName -ScriptVersion $sScriptVersion | Out-Null
# SCRIPT START

$Output = @()

Write-LogInfo -LogPath $sLogFile -Message "Get-OpenPortsWithProcess: Querying local computer for TCP ports"
Write-Host "Querying local computer for TCP ports" -ForegroundColor Yellow
$TCPConnections = Get-NetTCPConnection | Select-Object -Property @{n = "Proto"; e = { "TCP" } }, LocalPort, LocalAddress, OwningProcess, @{n = "ProcessName"; e = { (Get-Process -PID $_.OwningProcess).ProcessName } }

Write-LogInfo -LogPath $sLogFile -Message "Get-OpenPortsWithProcess: Querying local computer for UDP ports"
Write-Host "Querying local computer for UDP ports" -ForegroundColor Yellow
$UDPEndpoints = Get-NetUDPEndpoint | Select-Object -Property @{n = "Proto"; e = { "UDP" } }, LocalPort, LocalAddress, OwningProcess, @{n = "ProcessName"; e = { (Get-Process -PID $_.OwningProcess).ProcessName } }

Write-LogInfo -LogPath $sLogFile -Message "Get-OpenPortsWithProcess: Preparing output"
$Output += $TCPConnections
$Output += $UDPEndpoints

Write-LogInfo -LogPath $sLogFile -Message "Get-OpenPortsWithProcess: Writing report file to $sReportFile"
$Output | Sort-Object -Property LocalPort | Export-Csv $sReportFile -NoTypeInformation

Write-Host "Report created in $sReportFile" -ForegroundColor Green

Read-Host -Prompt "Press enter to exit..." | Out-Null

Write-LogInfo -LogPath $sLogFile -Message "Get-OpenPortsWithProcess: Script finished"

# SCRIPT END
Stop-Log -LogPath $sLogFile
Stop-Transcript