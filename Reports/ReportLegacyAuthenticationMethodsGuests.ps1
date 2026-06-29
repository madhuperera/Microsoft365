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
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

$S_OutputPath = $OutputPath
if (-not $S_OutputPath)
{
    $S_Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $S_OutputPath = Join-Path -Path (Get-Location).Path -ChildPath "ReportLegacyAuthenticationMethodsGuests_$S_Timestamp.csv"
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

$S_ActiveContext = Get-MgContext
Write-Host ""
Write-Host "Active Graph context:" -ForegroundColor Cyan
Write-Host "  Account    : $($S_ActiveContext.Account)" -ForegroundColor Cyan
Write-Host "  TenantId   : $($S_ActiveContext.TenantId)" -ForegroundColor Cyan
Write-Host "  Environment: $($S_ActiveContext.Environment)" -ForegroundColor Cyan
Write-Host "  Scopes     : $($S_ActiveContext.Scopes -join ', ')" -ForegroundColor Cyan
Write-Host ""

$S_ContextConfirmation = Read-Host "Proceed with this Graph context? [Y] Yes  [N] No  (Default: N)"
if ([string]::IsNullOrWhiteSpace($S_ContextConfirmation))
{
    $S_ContextConfirmation = 'N'
}
else
{
    $S_ContextConfirmation = $S_ContextConfirmation.ToUpperInvariant()
}
if ($S_ContextConfirmation -ne 'Y')
{
    throw "Operation cancelled. Please reconnect to the correct tenant and account, then run again."
}

Write-Host "Finding Azure AD Guest accounts"
[array]$S_Users = Get-MgUser -Filter "userType eq 'Guest'" -ConsistencyLevel eventual -CountVariable S_Records -All
if (-not $S_Users)
{
    Write-Host "No guest users found in Azure AD... exiting!"
    break
}

$S_i = 0
$S_Report = [System.Collections.Generic.List[Object]]::new()
foreach ($S_User in $S_Users)
{
    $S_i++
    Write-Host ("Processing user {0} {1}/{2}." -f $S_User.DisplayName, $S_i, $S_Users.Count)
    $S_AuthMethods = Get-MgUserAuthenticationMethod -UserId $S_User.Id
    Start-Sleep -Milliseconds $S_GraphRequestDelayMilliseconds

    $S_ModernTypes = @()
    $S_LegacyTypes = @()
    $S_ModernOdataTypes = @(
        "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod",
        "#microsoft.graph.fido2AuthenticationMethod"
    )

    foreach ($S_Method in $S_AuthMethods)
    {
        $S_Type = $S_Method.AdditionalProperties['@odata.type']
        switch ($S_Type)
        {
            "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" { $S_ModernTypes += "Microsoft Authenticator" }
            "#microsoft.graph.fido2AuthenticationMethod" { $S_ModernTypes += "Passkey" }
            "#microsoft.graph.passwordAuthenticationMethod" { $S_LegacyTypes += "Password" }
            "#microsoft.graph.phoneAuthenticationMethod" { $S_LegacyTypes += "Phone" }
            "#microsoft.graph.emailAuthenticationMethod" { $S_LegacyTypes += "Email" }
            "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" { $S_LegacyTypes += "Passwordless" }
            default { $S_LegacyTypes += $S_Type }
        }
    }

    if ($S_ModernTypes.Count -gt 0)
    {
        $S_DisplayMethod = "Modern Authentication"
        $S_P1 = $S_ModernTypes -join ", "
    }
    elseif ($S_LegacyTypes.Count -gt 0)
    {
        $S_DisplayMethod = "Legacy Authentication"
        $S_P1 = $S_LegacyTypes -join ", "
    }
    else
    {
        $S_DisplayMethod = "No Methods"
        $S_P1 = ""
    }

    $S_ReportLine = [PSCustomObject]@{
        User    = $S_User.DisplayName
        UPN     = $S_User.UserPrincipalName
        Type    = $S_DisplayMethod
        Methods = $S_P1
        Id      = $S_User.Id
    }
    $S_Report.Add($S_ReportLine)

} #End ForEach User
 
   
$S_Report = $S_Report | Sort-Object User
Write-Host ""
Write-Host "Authentication Methods found"
Write-Host "----------------------------"
Write-Host ""
$S_Report | Group-Object Type | Sort-Object Count -Descending | Select-Object Name, Count | Format-Table -AutoSize

$S_Report | Export-Csv -Path $S_OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Report exported to: $S_OutputPath" -ForegroundColor Green

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
