param
(
    [Parameter(Mandatory=$True)]
    [String] $DistributionGroupName,
    [Parameter(Mandatory=$True)]
    [String] $NewAccount,
    [Parameter(Mandatory=$True)]
    [String] $Permissions
)

$AllMembers = Get-DistributionGroupMember -Identity $DistributionGroupName `
    | Select-Object DisplayName, PrimarySmtpAddress `
    | Sort-Object DisplayName

foreach ($Member in $AllMembers)
{
    Write-Host "$($Member.DisplayName)" -ForegroundColor Green
    Add-MailboxFolderPermission -Identity $($Member.PrimarySmtpAddress + ":\Calendar") -User $NewAccount -AccessRights $Permissions
    Write-Host "`n"
}