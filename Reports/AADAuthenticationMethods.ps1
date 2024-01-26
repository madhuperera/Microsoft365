param
(
    [String] $S_ReportFilePath = "C:\Temp\AllAuthMethods.csv"
)

# Connect to Microsoft Graph with the UserAuthenticationMethod.Read.All permission scope
Connect-Graph -Scopes "UserAuthenticationMethod.Read.All"

# Get all users in Azure AD
$AllUsers = Get-MgUser -Filter "accountEnabled eq true"
$AllUsers = $AllUsers | Sort-Object -Property DisplayName

$AllData = @()

# Loop through each user and retrieve their authentication methods
foreach ($Member in $AllUsers)
{
    $MemberId = $Member.Id
    $MemberName = $Member.DisplayName
    Write-Host "`nWindows Hello for Business Check for $($MemberName):"
    $WHfBAuthMethods = Get-MgUserAuthenticationMethod -UserId $MemberId | Where-Object {$_.additionalProperties.'@odata.type' -like "*windowsHelloForBusinessAuthenticationMethod*"}
    if ($WHfBAuthMethods)
    {
		foreach ($Method in $WHfBAuthMethods)
		{
			$DeviceName = $Method.additionalProperties.displayName
			$RegisteredDate = $Method.additionalProperties.createdDateTime
			Write-Host "Found on $DeviceName"
	
			$OBJ = New-Object PSObject
			$OBJ | Add-Member -MemberType NoteProperty -Name "StaffDisplayName" -Value $MemberName
			$OBJ | Add-Member -MemberType NoteProperty -Name "DeviceName" -Value $DeviceName
			$OBJ | Add-Member -MemberType NoteProperty -Name "RegisteredDate" -Value $RegisteredDate
	
			$AllData += $OBJ
		}
	}
	else
	{
		Write-Host "No Windows Hello for Business Registrations found!`n"
	}
}

$AllData | Export-Csv -Path $S_ReportFilePath -NoTypeInformation
