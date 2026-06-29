#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Removes calendar permissions for a specified account from all members of a distribution group.

.DESCRIPTION
    Retrieves the members of the specified distribution group and removes the given account
    from the Calendar folder permissions of each member's mailbox. Requires an active
    Exchange Online session.

.PARAMETER DistributionGroupName
    The name or email address of the distribution group whose members will be processed.

.PARAMETER OldAccount
    The UPN or email address of the account whose calendar permissions will be removed.

.EXAMPLE
    .\Remove-AccountPermissions.ps1 -DistributionGroupName "Sales Team" -OldAccount "former@contoso.com"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$DistributionGroupName,

    [Parameter(Mandatory = $true)]
    [string]$OldAccount
)

$ErrorActionPreference = 'Stop'

$S_AllMembers = Get-DistributionGroupMember -Identity $DistributionGroupName `
    | Select-Object DisplayName, PrimarySmtpAddress `
    | Sort-Object DisplayName

foreach ($S_Member in $S_AllMembers)
{
    Write-Host "$($S_Member.DisplayName)" -ForegroundColor Green
    Remove-MailboxFolderPermission -Identity $($S_Member.PrimarySmtpAddress + ":\Calendar") -User $OldAccount -Confirm:$false
    Write-Host "`n"
}
