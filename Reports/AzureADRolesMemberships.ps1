$AllUsers = Get-MsolUser -All | Sort-Object -Property DisplayName
$RolesAssignment = foreach ($User in $AllUsers)
{
    $RolesAssigned = Get-MsolUserRole -ObjectId $User.ObjectId
    if ($RolesAssigned)
    {
        $UserProps = @{
            'DisplayName' = $User.DisplayName;
            'UserPrincipalName' = $User.UserPrincipalName;
            'IsLicensed' = $User.IsLicensed;
            'UserType' = $User.UserType;
            'BlockedSignIn' = $User.BlockCredential;
            'Roles' = $RolesAssigned.Name
        }
        New-Object -TypeName PSObject -Property $UserProps
    }
}

$RolesAssignment | Select-Object DisplayName, UserPrincipalName, BlockedSignIn, IsLicensed, UserType, Roles