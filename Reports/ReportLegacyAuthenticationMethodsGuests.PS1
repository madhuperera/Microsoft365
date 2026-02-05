<#
.SYNOPSIS
  Generates a report of authentication methods used by Azure AD Guest accounts.

.DESCRIPTION
  This script connects to Microsoft Graph using the required scopes and retrieves all Azure AD Guest user accounts.
  For each guest user, it enumerates their registered authentication methods, classifies them as Modern or Legacy, and compiles a report.
  The report is displayed in an Out-GridView window and exported as a CSV file.

.PARAMETER None
  The script does not accept parameters. It operates on all Azure AD Guest users.

.NOTES
  Author: Adapted from Tony Redmond's Office365itpros scripts.
  Source: https://github.com/12Knocksinna/Office365itpros/blob/master/ReportAuthenticationMethods.PS1
  More Info: https://office365itpros.com

.REQUIREMENTS
  - Microsoft Graph PowerShell SDK (Connect-MgGraph, Get-MgUser, Get-MgUserAuthenticationMethod)
  - Appropriate permissions: UserAuthenticationMethod.Read.All, Directory.Read.All, User.Read.All

.OUTPUTS
  - Displays a summary of authentication methods in the console.
  - Shows detailed results in an Out-GridView window.
  - Exports results to 'ReportAuthenticationMethods.csv' in the current directory.

.EXAMPLE
  .\ReportLegacyAuthenticationMethodsGuests.PS1

  Connects to Microsoft Graph, processes guest users, and exports their authentication methods.

.DISCLAIMER
  Do not use this script in production until you have validated it in a non-production environment.
  Review and adapt the code to meet your organization's requirements.
#>
# ReportLegacyAuthenticationMethodsGuests.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportAuthenticationMethods.PS1
# A report of the authentication methods for Azure AD Guest accounts

Connect-MgGraph -Scopes UserAuthenticationMethod.Read.All, Directory.Read.All, User.Read.All
# Select-MgProfile Beta

Write-Host "Finding Azure AD Guest accounts"
[array]$Users = Get-MgUser -Filter "userType eq 'Guest'" -ConsistencyLevel eventual -CountVariable Records -All
If (!($Users)) { Write-Host "No guest users found in Azure AD... exiting!"; break }

$i = 0
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $Users)
{
 $i++
 Write-Host ("Processing user {0} {1}/{2}." -f $User.DisplayName, $i, $Users.Count)
$AuthMethods = Get-MgUserAuthenticationMethod -UserId $User.Id

$ModernTypes = @()
$LegacyTypes = @()
$ModernOdataTypes = @(
  "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod",
  "#microsoft.graph.fido2AuthenticationMethod"
)

foreach ($Method in $AuthMethods) {
  $Type = $Method.AdditionalProperties['@odata.type']
  switch ($Type) {
    "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" { $ModernTypes += "Microsoft Authenticator" }
    "#microsoft.graph.fido2AuthenticationMethod" { $ModernTypes += "Passkey" }
    "#microsoft.graph.passwordAuthenticationMethod" { $LegacyTypes += "Password" }
    "#microsoft.graph.phoneAuthenticationMethod" { $LegacyTypes += "Phone" }
    "#microsoft.graph.emailAuthenticationMethod" { $LegacyTypes += "Email" }
    "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" { $LegacyTypes += "Passwordless" }
    default { $LegacyTypes += $Type }
  }
}

if ($ModernTypes.Count -gt 0) 
{
  $DisplayMethod = "Modern Authentication"
  $P1 = $ModernTypes -join ", "
} 
elseif ($LegacyTypes.Count -gt 0) 
{
  $DisplayMethod = "Legacy Authentication"
  $P1 = $LegacyTypes -join ", "
} 
else 
{
  $DisplayMethod = "No Methods"
  $P1 = ""
}

$ReportLine = [PSCustomObject]@{
  User   = $User.DisplayName
  UPN    = $User.UserPrincipalName
  Type   = $DisplayMethod
  Id     = $User.Id
  Methods= $P1
  UserId = $User.Id
}
$Report.Add($ReportLine)

} #End ForEach User
 
   
$Report = $Report | Sort-Object User 
Write-Host ""
Write-Host "Authentication Methods found"
Write-Host "----------------------------"
Write-Host ""
$Report | Group-Object Method | Sort-Object Count -Descending | Select Name, Count

$Report | Out-GridView

$Report | Export-Csv -Path ".\ReportAuthenticationMethodsGuests.csv" -NoTypeInformation -Encoding UTF8

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
