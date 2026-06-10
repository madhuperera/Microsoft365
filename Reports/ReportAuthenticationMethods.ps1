#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users

<#
.SYNOPSIS
    Reports authentication methods registered for licensed Entra ID member accounts.

.DESCRIPTION
    Connects to Microsoft Graph and retrieves all licensed member user accounts.
    For each user, enumerates their registered authentication methods and exports
    the results to a CSV file. Also displays an interactive grid view of the results.
    Adapted from Tony Redmond's Office365itpros scripts.

.PARAMETER OutputPath
    Path for the output CSV file. Defaults to a timestamped file in the current directory.

.EXAMPLE
    .\ReportAuthenticationMethods.ps1

.EXAMPLE
    .\ReportAuthenticationMethods.ps1 -OutputPath "C:\Reports\AuthMethods.csv"
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
    $OutputPath = Join-Path -Path (Get-Location).Path -ChildPath "ReportAuthenticationMethods_$S_Timestamp.csv"
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

Write-Host "Finding licensed Azure AD accounts"
[array]$S_Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel eventual -CountVariable S_Records -All
if (!($S_Users))
{
    Write-Host "No licensed users found in Azure AD... exiting!"
    break
}

$S_Counter = 0
$S_Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($S_User in $S_Users)
{
    $S_Counter++
    Write-Host ("Processing user {0} {1}/{2}." -f $S_User.DisplayName, $S_Counter, $S_Users.Count)
    $S_AuthMethods = Get-MgUserAuthenticationMethod -UserId $S_User.Id
    Start-Sleep -Milliseconds $S_GraphRequestDelayMilliseconds
    ForEach ($S_AuthMethod in $S_AuthMethods)
    {
        $S_P1 = $Null
        $S_P2 = $Null
        $S_Method = $S_AuthMethod.AdditionalProperties['@odata.type']
        Switch ($S_Method)
        {
            "#microsoft.graph.passwordAuthenticationMethod" {
                $S_DisplayMethod = "Password"
                $S_P1 = "Traditional password"
            }
            "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
                $S_DisplayMethod = "Microsoft Authenticator"
                $S_P1 = $S_AuthMethod.AdditionalProperties['displayName']
                $S_P2 = $S_AuthMethod.AdditionalProperties['deviceTag'] + ": " + $S_AuthMethod.AdditionalProperties['clientAppName']
            }
            "#microsoft.graph.fido2AuthenticationMethod" {
                $S_DisplayMethod = "Passkey"
                $S_P1 = $S_AuthMethod.AdditionalProperties['displayName']
                $S_P2 = If ($S_AuthMethod.AdditionalProperties['creationDateTime']) { Get-Date($S_AuthMethod.AdditionalProperties['creationDateTime']) -format g } Else { $Null }
            }
            "#microsoft.graph.phoneAuthenticationMethod" {
                $S_DisplayMethod = "Phone"
                $S_P1 = "Number: " + $S_AuthMethod.AdditionalProperties['phoneNumber']
                $S_P2 = "Type: " + $S_AuthMethod.AdditionalProperties['phoneType']
            }
            "#microsoft.graph.emailAuthenticationMethod" {
                $S_DisplayMethod = "Email"
                $S_P1 = "Address: " + $S_AuthMethod.AdditionalProperties['emailAddress']
            }
            "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" {
                $S_DisplayMethod = "Microsoft Authentication Passwordless"
                $S_P1 = $S_AuthMethod.AdditionalProperties['displayName']
                $S_P2 = If ($S_AuthMethod.AdditionalProperties['creationDateTime']) { Get-Date($S_AuthMethod.AdditionalProperties['creationDateTime']) -format g } Else { $Null }
            }
            "Default" {
                $S_DisplayMethod = $S_Method
            }
        }

        $S_ReportLine = [PSCustomObject] @{
            User   = $S_User.DisplayName
            UPN    = $S_User.UserPrincipalName
            Method = $S_DisplayMethod
            Id     = $S_AuthMethod.Id
            P1     = $S_P1
            P2     = $S_P2
            UserId = $S_User.Id
        }
        $S_Report.Add($S_ReportLine)
    } #End ForEach Authentication Method
} #End ForEach User
   
$S_Report = $S_Report | Sort-Object User 
Write-Host ""
Write-Host "Authentication Methods found"
Write-Host "----------------------------"
Write-Host ""
$S_Report | Group-Object Method | Sort-Object Count -Descending | Select Name, Count

$S_Report | Out-GridView

$S_Report | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
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
