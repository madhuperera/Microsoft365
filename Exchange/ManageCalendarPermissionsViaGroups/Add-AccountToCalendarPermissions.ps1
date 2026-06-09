#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Grants a specified account calendar permissions on all members of a distribution group.

.DESCRIPTION
    Retrieves the members of the specified distribution group and adds the given account
    to the Calendar folder of each member's mailbox with the specified permission level.
    Requires an active Exchange Online session.

.PARAMETER DistributionGroupName
    The name or email address of the distribution group whose members will be processed.

.PARAMETER NewAccount
    The UPN or email address of the account to grant calendar access to.

.PARAMETER Permissions
    The calendar permission level to assign, e.g. Reviewer, Editor, or LimitedDetails.

.EXAMPLE
    .\Add-AccountToCalendarPermissions.ps1 -DistributionGroupName "Sales Team" -NewAccount "admin@contoso.com" -Permissions "Reviewer"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$DistributionGroupName,

    [Parameter(Mandatory = $true)]
    [string]$NewAccount,

    [Parameter(Mandatory = $true)]
    [string]$Permissions
)

$ErrorActionPreference = 'Stop'

$S_AllMembers = Get-DistributionGroupMember -Identity $DistributionGroupName `
    | Select-Object DisplayName, PrimarySmtpAddress `
    | Sort-Object DisplayName

foreach ($S_Member in $S_AllMembers)
{
    Write-Host "$($S_Member.DisplayName)" -ForegroundColor Green
    Add-MailboxFolderPermission -Identity $($S_Member.PrimarySmtpAddress + ":\Calendar") -User $NewAccount -AccessRights $Permissions
    Write-Host "`n"
}