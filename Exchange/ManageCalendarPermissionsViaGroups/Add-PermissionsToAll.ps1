#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Grants a staff member calendar access to all user mailboxes in the tenant.

.DESCRIPTION
    Retrieves all user mailboxes and adds the specified staff member's account to the
    Calendar folder of each mailbox with the specified permission level. The staff
    member's own mailbox is excluded. Requires an active Exchange Online session.

.PARAMETER StaffMemberPrimarySmtpAddress
    The primary SMTP address of the staff member to grant calendar access to.

.PARAMETER Permissions
    The calendar permission level to assign, e.g. Reviewer, Editor, or LimitedDetails.

.EXAMPLE
    .\Add-PermissionsToAll.ps1 -StaffMemberPrimarySmtpAddress "manager@contoso.com" -Permissions "Reviewer"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$StaffMemberPrimarySmtpAddress,

    [Parameter(Mandatory = $true)]
    [string]$Permissions
)

$ErrorActionPreference = 'Stop'

$AllMembers = Get-Mailbox | Where-Object {$_.RecipientTypeDetails -eq "UserMailbox" -and $_.PrimarySmtpAddress -ne $StaffMemberPrimarySmtpAddress}`
    | Select-Object DisplayName, PrimarySmtpAddress `
    | Sort-Object DisplayName

foreach ($Member in $AllMembers)
{
    Write-Host "$($Member.DisplayName)" -ForegroundColor Green
    Add-MailboxFolderPermission -Identity $($Member.PrimarySmtpAddress + ":\Calendar") -User $StaffMemberPrimarySmtpAddress -AccessRights $Permissions
    Write-Host "`n"
}