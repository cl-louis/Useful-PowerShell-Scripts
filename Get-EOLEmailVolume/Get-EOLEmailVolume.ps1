#requires -version 5
<#
.SYNOPSIS
  Generate a report of the email volume for: a single mailbox, mailbox(es) from a file, or all mailboxes.

.DESCRIPTION
  This script generates a report of the email volume (send and receive) for: a single mailbox, mailbox(es) from a file,
or all mailboxes in a 365 tenant.

Use this script to audit email send/receive volume for a Microsft 365 Tenant.

.PARAMETER N/A
  N/A

.OUTPUTS Log File
  Log file stored in "Documents\Get-EOLEmailVolume\Get-EOLEmailVolume.log"

.OUTPUTS Transcript File
  Transcript file stored in "Documents\Get-EOLEmailVolume\Get-EOLEmailVolume.transcript"

.OUTPUTS Report File
  Report file stored in "Documents\Get-EOLEmailVolume\Get-EOLEmailVolume.csv"

.NOTES
  Version:        1.0
  Author:         Louis Lawson
  Creation Date:  18/06/2024

.EXAMPLE
  N/A
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
    [Parameter(Mandatory = $false, HelpMessage = "Path to users.txt, must contain CLRF delimited list of users. Default 'users.txt' in current directory")][string]$UsersFile = "users.txt"
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

# Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

# Import Modules & Snap-ins
if (-not (Get-Module PSLogging -ListAvailable)) {
    Install-Module PSLogging -Scope CurrentUser -Force
}
Import-Module PSLogging

if (-not (Get-Module ExchangeOnlineManagement -ListAvailable)) {
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
}
Import-Module ExchangeOnlineManagement

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

Function Invoke-ConnectToExchange {
    Param ()
    
    Begin {
        Write-LogInfo -LogPath $sLogFile -Message "Invoke-ConnectToExchange: Inititiating connection with ExchangeOnline"
    }
    
    Process {
        Try {
            $Exchange = (get-module ExchangeOnlineManagement -ListAvailable).Name
            if ($null -eq $Exchange) {
                Write-host "Important: ExchangeOnline PowerShell module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
                $confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No  
                if ($confirm -match "[yY]") { 
                    Write-host "Installing ExchangeOnlineManagement"
                    Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
                    Write-host "ExchangeOnline PowerShell module is installed in the machine successfully."`n
                }
                elseif ($confirm -cnotmatch "[yY]" ) { 
                    Write-host "Exiting. `nNote: ExchangeOnline PowerShell module must be available in your system to run the script." 
                    Exit 
                }
            }
        }
    
        Catch {
            Write-LogError -LogPath $sLogFile -Message $_.Exception -ExitGracefully
            Break
        }
    }
    
    End {
        Connect-ExchangeOnline -ShowBanner:$false | Out-Null
        Write-Host "Successfully connected to ExchangeOnline" -ForegroundColor Green
        If ($?) {
            Write-LogInfo -LogPath $sLogFile -Message "Invoke-ConnectToExchange: Completed Successfully."
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
    Write-Host "This script generates a report of the email volume (send and receive) of a single mailbox, mailbox(es) from a file,
or all mailboxes in a 365 tenant." -ForegroundColor Black -BackgroundColor White
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Start-Transcript -Path $sTranscriptFile | Out-Null
Start-Log -LogPath $sOutputDir -LogName $sLogName -ScriptVersion $sScriptVersion | Out-Null
# SCRIPT START

# SCRIPT LOGIC GOES HERE

Invoke-ConnectToExchange

$UserSelection = Show-OptionMenu -Title 'Usage method' -Choices @('From file', 'Single user', 'All users')

if ($UserSelection -eq "From file") {
    Write-Host "Running in from file audit mode" -ForegroundColor Yellow
    $UserSelection = Import-TxtFile -TxtPath "$UsersFile"
}
elseif ($UserSelection -eq "Single user") {
    Write-Host "Running in single user audit mode" -ForegroundColor Yellow
    $UserSelection = Read-Host -Prompt "Enter user email"
}
elseif ($UserSelection -eq "All users") {
    Write-Host "Running in all user audit mode" -ForegroundColor Yellow
    $UserSelection = Get-Mailbox -ResultSize unlimited | Select-Object -ExpandProperty PrimarySmtpAddress
}

Write-Host $UserSelection

$endDate = (Get-Date).Date
$startDate = ($endDate).AddDays(-7)
$pageSize = 5000
$maxPage = 1000

$logEveryXPages = 1

$gmtParams = @{
    StartDate = $startDate
    EndDate   = $endDate
    PageSize  = $pageSize
    Page      = 1
}

$Output = foreach ($user in $UserSelection) {
    do {
        $gmtParams.Page = 1
        $totalSent = 0
  
        Write-Host "Querying sent in range $($gmtParams.StartDate) - $($gmtParams.EndDate)" -ForegroundColor Yellow
  
        do {
            # Logging
            if ($gmtParams.Page % $logEveryXPages -eq 0) {
                Write-Host "Processing page $($gmtParams.Page)"
            }
      
            $currentMessages = Get-MessageTrace -SenderAddress $user @gmtParams
  
            $gmtParams.Page++
            # $currentMessages | Export-Csv -Path $outFilePath -Append
  
            $totalSent += $($currentMessages).Count
  
            # We need to grab the timestamp for last occurence
            if ($gmtParams.Page -gt $maxPage -and $null -ne $currentMessages) {
                $gmtParams.EndDate = $currentMessages.Received | Sort-Object | Select-Object -First 1
            }
  
            # Let's add a short break so we're not throttled
            Start-Sleep -Seconds 1
  
            # The loop should end when there's no more messages
            # or when we reach last page
        } until ( $null -eq $currentMessages -or $gmtParams.Page -gt $maxPage )
  
        $gmtParams.Page = 1
        $totalReceived = 0
  
        Write-Host "Querying received in range $($gmtParams.StartDate) - $($gmtParams.EndDate)" -ForegroundColor Yellow
  
        do {
            # Logging
            if ($gmtParams.Page % $logEveryXPages -eq 0) {
                Write-Host "Processing page $($gmtParams.Page)"
            }
      
            $currentMessages = Get-MessageTrace -RecipientAddress $user @gmtParams
  
            $gmtParams.Page++
            # $currentMessages | Export-Csv -Path $outFilePath -Append
  
            $totalReceived += $($currentMessages).Count
  
            # We need to grab the timestamp for last occurence
            if ($gmtParams.Page -gt $maxPage -and $null -ne $currentMessages) {
                $gmtParams.EndDate = $currentMessages.Received | Sort-Object | Select-Object -First 1
            }
  
            # Let's add a short break so we're not throttled
            Start-Sleep -Seconds 1
  
            # The loop should end when there's no more messages
            # or when we reach last page
        } until ( $null -eq $currentMessages -or $gmtParams.Page -gt $maxPage )
  
        # The outer loop should end when there's no more messages
    } until ( $null -eq $currentMessages )
  
    $userMailbox = Get-Mailbox -Identity $user
  
    [pscustomobject]@{
        DisplayName        = $($userMailbox.DisplayName)
        PrimarySmtpAddress = $($userMailbox.PrimarySmtpAddress)
        RecipientType      = $($userMailbox.RecipientTypeDetails)
        Sent               = $totalSent
        Received           = $totalReceived
    }
}

Write-LogInfo -LogPath $sLogFile -Message "Get-EOLEmailVolume: Writing report file to $sReportFile"
$Output | Export-Csv $sReportFile -NoTypeInformation

Write-Host "Report created in $sReportFile" -ForegroundColor Green

Write-LogInfo -LogPath $sLogFile -Message "Get-EOLEmailVolume: Script finished"

Read-Host -Prompt "Press enter to exit..." | Out-Null

# SCRIPT END
Stop-Log -LogPath $sLogFile
Stop-Transcript