#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users

<#
.SYNOPSIS
    Reports authentication methods for Entra ID guest accounts, classifying each as Modern or Legacy.

.DESCRIPTION
    Connects to Microsoft Graph and retrieves all guest user accounts.
    For each guest, enumerates registered authentication methods, classifies them as
    Modern Authentication, Legacy Authentication, or No Methods, and exports results to CSV.
    Adapted from Tony Redmond's Office365itpros scripts.

.PARAMETER OutputPath
    Path for the output CSV file. Defaults to a timestamped file in the current directory.

.EXAMPLE
    .\ReportLegacyAuthenticationMethodsGuests.ps1

.EXAMPLE
    .\ReportLegacyAuthenticationMethodsGuests.ps1 -OutputPath "C:\Reports\GuestAuthMethods.csv"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

if (-not $OutputPath)
{
    $S_Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path -Path (Get-Location).Path -ChildPath "ReportLegacyAuthenticationMethodsGuests_$S_Timestamp.csv"
}

$S_RequiredGraphScopes = @(
    'UserAuthenticationMethod.Read.All'
    'Directory.Read.All'
    'User.Read.All'
)

$S_GraphRequestDelayMilliseconds = 5

$S_ExistingContext = Get-MgContext
if ($S_ExistingContext)
{
    Write-Host "Existing Graph session detected:" -ForegroundColor Yellow
    Write-Host "  Account : $($S_ExistingContext.Account)" -ForegroundColor Yellow
    Write-Host "  TenantId: $($S_ExistingContext.TenantId)" -ForegroundColor Yellow
    Write-Host "  Scopes  : $($S_ExistingContext.Scopes -join ', ')" -ForegroundColor Yellow
    Write-Host ""

    $S_Choice = Read-Host "Use existing session? [Y] Yes  [N] Disconnect and reconnect  (Default: Y)"
    if ($S_Choice -eq 'N')
    {
        Disconnect-MgGraph | Out-Null
        Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome
    }
}
else
{
    Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome
}

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
  User    = $User.DisplayName
  UPN     = $User.UserPrincipalName
  Type    = $DisplayMethod
  Methods = $P1
  Id      = $User.Id
}
$Report.Add($ReportLine)

} #End ForEach User
 
   
$Report = $Report | Sort-Object User
Write-Host ""
Write-Host "Authentication Methods found"
Write-Host "----------------------------"
Write-Host ""
$Report | Group-Object Type | Sort-Object Count -Descending | Select-Object Name, Count | Format-Table -AutoSize

$Report | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Report exported to: $OutputPath" -ForegroundColor Green

$S_DisconnectChoice = Read-Host "`nDisconnect from Microsoft Graph? [Y] Yes  [N] Keep session  (Default: N)"
if ($S_DisconnectChoice -eq 'Y')
{
    Disconnect-MgGraph | Out-Null
    Write-Host "Disconnected." -ForegroundColor Green
}
else
{
    Write-Host "Graph session kept alive." -ForegroundColor Green
}
