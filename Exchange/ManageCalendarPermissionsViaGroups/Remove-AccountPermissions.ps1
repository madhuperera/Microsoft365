param
(
    [Parameter(Mandatory=$True)]
    [String] $DistributionGroupName,
    [Parameter(Mandatory=$True)]
    [String] $OldAccount
)

$AllMembers = Get-DistributionGroupMember -Identity $DistributionGroupName `
    | Select-Object DisplayName, PrimarySmtpAddress `
    | Sort-Object DisplayName

foreach ($Member in $AllMembers)
{
    Write-Host "$($Member.DisplayName)" -ForegroundColor Green
    Remove-MailboxFolderPermission -Identity $($Member.PrimarySmtpAddress + ":\Calendar") -User $OldAccount -Confirm:$false
    Write-Host "`n"
}