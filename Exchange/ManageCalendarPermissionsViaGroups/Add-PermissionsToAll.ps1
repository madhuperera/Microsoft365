param
(
    [Parameter(Mandatory=$True)]
    [String] $StaffMemberPrimarySmtpAddress,
    [Parameter(Mandatory=$True)]
    [String] $Permissions
)

$AllMembers = Get-Mailbox | Where-Object {$_.RecipientTypeDetails -eq "UserMailbox" -and $_.PrimarySmtpAddress -ne $StaffMemberPrimarySmtpAddress}`
    | Select-Object DisplayName, PrimarySmtpAddress `
    | Sort-Object DisplayName

foreach ($Member in $AllMembers)
{
    Write-Host "$($Member.DisplayName)" -ForegroundColor Green
    Add-MailboxFolderPermission -Identity $($Member.PrimarySmtpAddress + ":\Calendar") -User $StaffMemberPrimarySmtpAddress -AccessRights $Permissions
    Write-Host "`n"
}