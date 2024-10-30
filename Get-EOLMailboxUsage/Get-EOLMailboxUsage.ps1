#requires -version 5
<#
.SYNOPSIS
  Generate a report of mailbox size/usage for: a single mailbox, mailbox(es) from a file, or all mailboxes.

.DESCRIPTION
  This script generates a report of mailbox size/usage for: a single mailbox, mailbox(es) from a file,
or all mailboxes in a 365 tenant.

Use this script to audit mailbox size/usage for a Microsft 365 Tenant.

.PARAMETER N/A
  N/A

.OUTPUTS Log File
  Log file stored in "Documents\Get-EOLMailboxUsage\Get-EOLMailboxUsage.log"

.OUTPUTS Transcript File
  Transcript file stored in "Documents\Get-EOLMailboxUsage\Get-EOLMailboxUsage.transcript"

.OUTPUTS Report File
  Report file stored in "Documents\Get-EOLMailboxUsage\Get-EOLMailboxUsage.csv"

.NOTES
  Version:        1.0
  Author:         Louis Lawson
  Creation Date:  29/10/2024

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
        Write-LogInfo -LogPath $sLogFile -Message "Show-OptionMenu: Param -Choices $($Choices.count)"
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
            $Exchange = (Get-Module ExchangeOnlineManagement -ListAvailable).Name
            if ($null -eq $Exchange) {
                Write-host "Important: ExchangeOnline PowerShell module is unavailable. It is mandatory to have this module installed in the system to run the script successfully." 
                $confirm = Read-Host Are you sure you want to install module? [Y] Yes [N] No  
                if ($confirm -match "[yY]") { 
                    Write-host "Installing ExchangeOnlineManagement"
                    Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
                    Write-host "ExchangeOnline PowerShell module is installed in the machine successfully."`n
                }
                elseif ($confirm -cnotmatch "[yY]" ) { 
                    Write-host "Exiting. Note: ExchangeOnline PowerShell module must be available in your system to run the script." 
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
    Write-Host "This script generates a report of mailbox size/usage of a single mailbox, mailbox(es) from a file,
or all mailboxes in a 365 tenant." -ForegroundColor Black -BackgroundColor White
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Start-Transcript -Path $sTranscriptFile | Out-Null
Start-Log -LogPath $sOutputDir -LogName $sLogName -ScriptVersion $sScriptVersion | Out-Null
# SCRIPT START

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
    $UserSelection = Get-EXOMailbox -ResultSize unlimited | Select-Object -ExpandProperty PrimarySmtpAddress
}

$Output = foreach ($User in $UserSelection) {
    Write-Host "Processing user $User" -ForegroundColor Yellow
    $Mailbox = Get-EXOMailbox -PropertySets All -Identity "$User"
    $MailboxStats = Get-EXOMailboxStatistics -PropertySets All -Identity "$User"

    if ($Mailbox.ArchiveStatus -eq "Active") {
        $MailboxArchive = Get-EXOMailboxStatistics -PropertySets All -Archive -Identity "$User"
    }

    [pscustomobject]@{
        DisplayName                 = $($Mailbox.DisplayName)
        PrimaryEmail                = $($Mailbox.PrimarySmtpAddress)
        MailboxType                 = $($Mailbox.RecipientTypeDetails)
        ItemCount                   = $($MailboxStats.ItemCount)
        ItemSizeMB                  = $($MailboxStats.TotalItemSize.value.ToMB())
        DeletedItemCount            = $($MailboxStats.DeletedItemCount)
        DeletedItemSizeMB           = $($MailboxStats.TotalDeletedItemSize.value.ToMB())
        MessageTableSizeMB          = $($MailboxStats.MessageTableTotalSize.value.ToMB())
        AttachmentTableSizeMB       = $($MailboxStats.AttachmentTableTotalSize.value.ToMB())
        ProhibitSendQuotaMB         = $([math]::Round(($Mailbox.ProhibitSendQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0))
        ProhibitSendReceiveQuotaMB  = $([math]::Round(($Mailbox.ProhibitSendReceiveQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0))
        IssueWarningQuotaMB         = $([math]::Round(($Mailbox.IssueWarningQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0))
        ArchiveStatus               = if ($Mailbox.ArchiveStatus -eq "Active") { $($Mailbox.ArchiveStatus) } Else { "Inactive" }
        ArchiveName                 = if ($Mailbox.ArchiveStatus -eq "Active") { $($MailboxArchive.DisplayName) } Else { $null }
        ArchiveItemCount            = if ($Mailbox.ArchiveStatus -eq "Active") { $($MailboxArchive.ItemCount) } Else { $null }
        ArchiveItemSizeMB           = if ($Mailbox.ArchiveStatus -eq "Active") { $($MailboxArchive.TotalItemSize.value.ToMB()) } Else { $null }
        ArchiveQuotaMB              = if ($Mailbox.ArchiveStatus -eq "Active") { $([math]::Round(($Mailbox.ArchiveQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0)) } Else { $null } 
        ArchiveWarningQuotaMB       = if ($Mailbox.ArchiveStatus -eq "Active") { $([math]::Round(($Mailbox.ArchiveWarningQuota.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0)) } Else { $null }
        AutoExpandingArchiveEnabled = $($Mailbox.AutoExpandingArchiveEnabled)
        RetentionPolicy             = $($Mailbox.RetentionPolicy)
    }
}

Write-LogInfo -LogPath $sLogFile -Message "Get-EOLEmailVolume: Writing report file to $sReportFile"
$Output | Sort-Object "ItemSizeMB" | Export-Csv $sReportFile -NoTypeInformation

Write-Host "Report created in $sReportFile" -ForegroundColor Green

Write-LogInfo -LogPath $sLogFile -Message "Get-EOLEmailVolume: Script finished"

Read-Host -Prompt "Press enter to exit..." | Out-Null

# SCRIPT END
Stop-Log -LogPath $sLogFile
Stop-Transcript