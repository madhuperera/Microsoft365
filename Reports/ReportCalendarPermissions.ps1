#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Reports calendar permissions for all members of a specified distribution group.

.DESCRIPTION
    Connects to Exchange Online, retrieves all members of a specified distribution group, and outputs
    the calendar permissions configured for each member's mailbox calendar. Excludes
    the Default and Anonymous permission entries.

.PARAMETER DistributionGroupName
    The name or email address of the Exchange Online distribution group to process.

.EXAMPLE
    .\ReportCalendarPermissions.ps1 -DistributionGroupName "IT-Team"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]
    [string]$DistributionGroupName
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable))
{
    throw "ExchangeOnlineManagement module is not installed. Install it using: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}
Connect-ExchangeOnline -ShowBanner:$false

$S_AllMembers = Get-DistributionGroupMember -Identity $DistributionGroupName `
    | Select-Object DisplayName, PrimarySmtpAddress `
    | Sort-Object DisplayName

foreach ($Member in $S_AllMembers)
{
    Write-Host "$($Member.DisplayName)" -ForegroundColor Green
    Get-MailboxFolderPermission -Identity $($Member.PrimarySmtpAddress + ":\Calendar") `
        | Where-Object {($_.User.DisplayName -ne "Default") -and ($_.User.DisplayName -ne "Anonymous")}
    Write-Host "`n"
}