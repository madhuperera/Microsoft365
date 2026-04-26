#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Reports calendar permissions for all members of a specified distribution group.

.DESCRIPTION
    Retrieves all members of a specified Exchange Online distribution group and outputs
    the calendar permissions configured for each member's mailbox calendar. Excludes
    the Default and Anonymous permission entries.

.PARAMETER DistributionGroupName
    The name or email address of the Exchange Online distribution group to process.

.EXAMPLE
    .\ReportCalendarPermissions.ps1 -DistributionGroupName "IT-Team"
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory=$True)]
    [String] $DistributionGroupName
)

$ErrorActionPreference = 'Stop'

$AllMembers = Get-DistributionGroupMember -Identity $DistributionGroupName `
    | Select-Object DisplayName, PrimarySmtpAddress `
    | Sort-Object DisplayName

foreach ($Member in $AllMembers)
{
    Write-Host "$($Member.DisplayName)" -ForegroundColor Green
    Get-MailboxFolderPermission -Identity $($Member.PrimarySmtpAddress + ":\Calendar") `
        | Where-Object {($_.User.DisplayName -ne "Default") -and ($_.User.DisplayName -ne "Anonymous")}
    Write-Host "`n"
}