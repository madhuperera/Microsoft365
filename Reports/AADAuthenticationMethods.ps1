# Connect to Microsoft Graph with the UserAuthenticationMethod.Read.All permission scope
Connect-Graph -Scopes "UserAuthenticationMethod.Read.All"

# Get all users in Azure AD
 $AllUsers = Get-MgUser -Filter "accountEnabled eq true"

$AllData = @()

# Loop through each user and retrieve their authentication methods
foreach ($Member in $AllUsers)
{
    $MemberId = $Member.Id
    $MemberName = $Member.DisplayName
    Write-Host "Windows Hello for Business for $($MemberName):"
    $WHfBAuthMethods = Get-MgUserAuthenticationMethod -UserId $MemberId | Where-Object {$_.additionalProperties.'@odata.type' -like "*windowsHelloForBusinessAuthenticationMethod*"}
    foreach ($Method in $WHfBAuthMethods)
    {
        $DeviceName = $Method.additionalProperties.displayName
        $RegisteredDate = $Method.additionalProperties.createdDateTime

        $OBJ = New-Object PSObject
        $OBJ | Add-Member -MemberType NoteProperty -Name "STaffDisplayName" -Value $MemberName
        $OBJ | Add-Member -MemberType NoteProperty -Name "DeviceName" -Value $DeviceName
        $OBJ | Add-Member -MemberType NoteProperty -Name "RegisteredDate" -Value $RegisteredDate

        $AllData += $OBJ
    }
}
