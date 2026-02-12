<#
.SYNOPSIS
  Generates a report of authentication methods used by Azure AD (Microsoft 365) accounts.

.DESCRIPTION
  This script connects to Microsoft Graph using the required scopes and retrieves all Azure AD user accounts.
  For each user, it enumerates their registered authentication methods, classifies them as Modern or Legacy, and compiles a report.
  The report includes license status, on-premises sync status, and account enabled status.
  The report is displayed in an Out-GridView window and exported as a CSV file.

.PARAMETER None
  The script does not accept parameters. It operates on all Azure AD users.

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
  .\ReportLegacyAuthenticationMethods.PS1

  Connects to Microsoft Graph, processes all users, and exports their authentication methods with license and sync status.

.DISCLAIMER
  Do not use this script in production until you have validated it in a non-production environment.
  Review and adapt the code to meet your organization's requirements.
#>
# ReportAuthenticationMethods.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportAuthenticationMethods.PS1
# A report of the authentication methods for Azure AD licensed accounts

Connect-MgGraph -Scopes UserAuthenticationMethod.Read.All, Directory.Read.All, User.Read.All
# Select-MgProfile Beta

Write-Host "Finding Azure AD accounts"
[array]$Users = Get-MgUser -Filter "userType eq 'Member'" -ConsistencyLevel eventual -CountVariable Records -All -Property Id, DisplayName, UserPrincipalName, AccountEnabled, AssignedLicenses, OnPremisesSyncEnabled
If (!($Users)) { Write-Host "No users found in Azure AD... exiting!"; break }

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
  $P1 = $LegacyTypes -join ", "
  # Check if only Password method exists (No MFA)
  if ($LegacyTypes.Count -eq 1 -and $LegacyTypes[0] -eq "Password") 
  {
    $DisplayMethod = "No MFA"
  }
  else
  {
    $DisplayMethod = "Legacy Authentication"
  }
} 
else 
{
  $DisplayMethod = "No Methods"
  $P1 = ""
}

# Determine license status
$IsLicensed = if ($User.AssignedLicenses.Count -gt 0) { "Yes" } else { "No" }

# Determine on-premises sync status
$IsSynced = if ($User.OnPremisesSyncEnabled -eq $true) { "Yes" } elseif ($User.OnPremisesSyncEnabled -eq $false) { "No" } else { "Cloud-Only" }

# Determine account enabled status
$IsEnabled = if ($User.AccountEnabled -eq $true) { "Yes" } else { "No" }

$ReportLine = [PSCustomObject]@{
  User   = $User.DisplayName
  UPN    = $User.UserPrincipalName
  Type   = $DisplayMethod
  Methods= $P1
  Licensed = $IsLicensed
  OnPremisesSynced = $IsSynced
  AccountEnabled = $IsEnabled
  Id     = $User.Id
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

$Report | Export-Csv -Path ".\ReportAuthenticationMethods.csv" -NoTypeInformation -Encoding UTF8

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
