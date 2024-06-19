#requires -version 5
<#
.SYNOPSIS
  SMB share audit tool.

.DESCRIPTION
  This script generates an audit of SMB shared for all 'Server' computers on an AD domain.

.PARAMETER N/A
  N/A

.OUTPUTS Log File
  Log file stored in "Documents\Get-SMBShareAudit\Get-SMBShareAudit.log"

.OUTPUTS Transcript File
  Transcript file stored in "Documents\Get-SMBShareAudit\Get-SMBShareAudit.transcript"

.OUTPUTS Report File
  Report file stored in "Documents\Get-SMBShareAudit\Get-SMBShareAudit.csv"

.NOTES
  Version:        1.0
  Author:         Louis Lawson
  Creation Date:  11/06/2024

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

if (-not (Get-Module SmbShare -ListAvailable)) {
  Install-Module SmbShare -Scope CurrentUser -Force
}
Import-Module SmbShare

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
  Write-Host "This PowerShell script audits the SMB shares on any 'Server' of a given AD Domain" -ForegroundColor Black -BackgroundColor White
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Start-Transcript -Path $sTranscriptFile | Out-Null
Start-Log -LogPath $sOutputDir -LogName $sLogName -ScriptVersion $sScriptVersion | Out-Null
# SCRIPT START

# SCRIPT LOGIC GOES HERE

# Get all computer accounts from domain with the word 'server' in the OS field
$servers = Get-ADComputer -Filter { OperatingSystem -like "*server*" } # -and Name -like "CL-FSS01"

# Loop over computer accounts
$Output = foreach ( $server in $servers ) {
  Write-Host "Processing"$server.Name"..." -foregroundcolor yellow
  Write-LogInfo -LogPath $sLogFile -Message "Get-SMBShareAudit: Processing $($server.Name)..."
  # Use Invoke-Command to run the SMB audit code on the target computer
  # as Get-Acl doesn't have a -Computer param
  Invoke-Command -ComputerName "$($server.Name)" -ScriptBlock {

    # array used as a lookup for permission numbers to readable strings
    $permissions = @{
      '268435456'   = 'FullControl'
      '-536805376'  = 'Modify, Synchronize'
      '-1610612736' = 'ReadAndExecute, Synchronize'
    }
        
    # Get all of the shares on the computer
    $shares = Get-SmbShare -IncludeHidden

    # Loop over shares
    foreach ( $share in $shares ) {
      # Run Get-Acl on the local path of the share to get security permissions
      Get-Acl -Path "$($share.Path)" | ForEach-Object {
        # 'Access' is each permission on the directory
        foreach ($access in $_.Access) {
          # If the FileSystemRights is a number then lookup
          # a human friendly string in the $permissions array
          if ($access.FileSystemRights -match '\d') {
            $fsr = $access.FileSystemRights.ToString()
            $fileSystemRights = $($permissions[$fsr])
          }
          else {
            $fileSystemRights = $access.FileSystemRights
          }
          # Construct a custom object to return
          # Example
          #   ServerName: CL-SQL01
          #   ShareName: POShare
          #   ShareState: Online
          #   ShareType: FileSystemDirectory
          #   ShareCurrentUsers: 0
          #   ShareLocalPath: C:\Program Files (x86)\Draycir\Spindle Requisitions\POShare
          #   ShareDescription: 
          #   FileSystemRights: FullControl
          #   AccessControlType: Allow
          #   IdentityReference: CL-SQL01\SDCUpdaterUser
          #   IsInherited: TRUE
          #   InheritanceFlags: ContainerInherit, ObjectInherit
          #   PropagationFlags: InheritOnly
          [PSCustomObject]@{
            ServerName        = $env:ComputerName
            ShareName         = $($share.Name)
            ShareState        = $($share.ShareState)
            ShareType         = $($share.ShareType)
            ShareCurrentUsers = $($share.CurrentUsers)
            ShareLocalPath    = (($_.Path) -split ("::"))[-1]
            ShareDescription  = $($share.Description)
            FileSystemRights  = $fileSystemRights
            AccessControlType = $access.AccessControlType
            IdentityReference = $access.IdentityReference
            IsInherited       = $access.IsInherited
            InheritanceFlags  = $access.InheritanceFlags
            PropagationFlags  = $access.PropagationFlags
          }
        }
      }
    }
  }
}

Write-LogInfo -LogPath $sLogFile -Message "Get-SMBShareAudit: Writing report file to $sReportFile"
$Output | Export-Csv $sReportFile -NoTypeInformation

Write-Host "Report created in $sReportFile" -ForegroundColor Green

Read-Host -Prompt "Press enter to exit..." | Out-Null

# SCRIPT END
Stop-Log -LogPath $sLogFile
Stop-Transcript