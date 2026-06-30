#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users

<#
.SYNOPSIS
    Exports Entra ID users who have a manager assigned.

.DESCRIPTION
    Connects to Microsoft Graph and exports a list of licensed, enabled users who have
    a manager assigned, including their email address and manager information.

.PARAMETER OutputPath
    Path for the output CSV file. Defaults to a timestamped file in the current directory.

.NOTES
    Author: Madhu Perera
    Requires: Microsoft.Graph.Authentication, Microsoft.Graph.Users modules

.EXAMPLE
    .\ReportUsersWithManagers.ps1

.EXAMPLE
    .\ReportUsersWithManagers.ps1 -OutputPath "C:\Reports\UsersWithManagers.csv"
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
    $S_OutputPath = Join-Path -Path (Get-Location).Path -ChildPath "ReportUsersWithManagers_$S_Timestamp.csv"
}

$S_RequiredGraphScopes = @(
    'User.Read.All'
    'Directory.Read.All'
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

# Get all users with their manager information
Write-Host "Retrieving licensed and enabled users from Entra ID..." -ForegroundColor Cyan

$S_Users = Get-MgUser -All -Filter "assignedLicenses/`$count ne 0 and accountEnabled eq true" -ConsistencyLevel eventual -CountVariable Records -Property Id, DisplayName, UserPrincipalName, Mail, JobTitle, Department

$S_UsersWithManagers = @()

foreach ($user in $S_Users)
{
    Write-Progress -Activity "Processing users" -Status "Checking $($user.DisplayName)" -PercentComplete (($S_UsersWithManagers.Count / $S_Users.Count) * 100)
    
    # Get manager for each user
    try
    {
        Start-Sleep -Milliseconds $S_GraphRequestDelayMilliseconds
        $manager = Get-MgUserManager -UserId $user.Id -ErrorAction SilentlyContinue
        
        if ($manager)
        {
            $S_UsersWithManagers += [PSCustomObject]@{
                DisplayName        = $user.DisplayName
                UserPrincipalName  = $user.UserPrincipalName
                Email              = if ($user.Mail) { $user.Mail } else { $user.UserPrincipalName }
                JobTitle           = $user.JobTitle
                Department         = $user.Department
                ManagerName        = $manager.AdditionalProperties.displayName
                ManagerEmail       = $manager.AdditionalProperties.userPrincipalName
            }
        }
    }
    catch
    {
        # User has no manager assigned, skip
        continue
    }
}

Write-Progress -Activity "Processing users" -Completed

$S_UsersWithManagers | Export-Csv -Path $S_OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "`nExport complete!" -ForegroundColor Green
Write-Host "Total users with managers: $($S_UsersWithManagers.Count)" -ForegroundColor Yellow
Write-Host "File saved to: $S_OutputPath" -ForegroundColor Yellow

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
