param
(
    [Parameter(Mandatory=$True)]
    [String] $DistributionGroupName
)

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