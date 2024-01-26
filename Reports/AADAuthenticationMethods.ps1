# Connect to Microsoft Graph with the UserAuthenticationMethod.Read.All permission scope
Connect-Graph -Scopes "UserAuthenticationMethod.Read.All"

# Get all users in Azure AD
 $AllUsers = Get-MgUser -Filter "accountEnabled eq true"


# Loop through each user and retrieve their authentication methods
foreach ($Member in $AllUsers)
{
    $MemberId = $Member.Id
    $MemberName = $Member.DisplayName
    Write-Host "Windows Hello for Business for $($MemberName):"
    $WHfBAuthMethods = Get-MgUserAuthenticationMethod -UserId $MemberId | Where-Object {$_.additionalProperties.'@odata.type' -like "*WindowsHelloForBusiness*"}
    foreach ($Method in $WHfBAuthMethods)
    {
        $Method
    }
}
