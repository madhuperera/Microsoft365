#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users

<#
.SYNOPSIS
    Reports Windows Hello for Business authentication method registrations for enabled Entra ID users.

.DESCRIPTION
    Connects to Microsoft Graph and retrieves all enabled users. For each user, checks whether
    they have a Windows Hello for Business authentication method registered and exports the results
    to a CSV file.

.PARAMETER OutputPath
    Path for the output CSV file. Defaults to a timestamped file in the current directory.

.EXAMPLE
    .\ReportAADAuthenticationMethods.ps1

.EXAMPLE
    .\ReportAADAuthenticationMethods.ps1 -OutputPath "C:\Reports\WHfBReport.csv"
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
    $OutputPath = Join-Path -Path (Get-Location).Path -ChildPath "ReportAADAuthenticationMethods_$S_Timestamp.csv"
}

$S_RequiredGraphScopes = @(
    'UserAuthenticationMethod.Read.All'
)

$S_GraphRequestDelayMilliseconds = 5

# Connect to Microsoft Graph
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

# Get all enabled users in Entra ID
$S_ActiveContext = Get-MgContext
Write-Host ""
Write-Host "Active Graph context:" -ForegroundColor Cyan
Write-Host "  Account    : $($S_ActiveContext.Account)" -ForegroundColor Cyan
Write-Host "  TenantId   : $($S_ActiveContext.TenantId)" -ForegroundColor Cyan
Write-Host "  Environment: $($S_ActiveContext.Environment)" -ForegroundColor Cyan
Write-Host "  Scopes     : $($S_ActiveContext.Scopes -join ', ')" -ForegroundColor Cyan
Write-Host ""

$S_ContextConfirmation = Read-Host "Proceed with this Graph context? [Y] Yes  [N] No"
if ($S_ContextConfirmation -ne 'Y')
{
    throw "Operation cancelled. Please reconnect to the correct tenant and account, then run again."
}

$S_AllUsers = Get-MgUser -Filter "accountEnabled eq true" -All
$S_AllUsers = $S_AllUsers | Sort-Object -Property DisplayName

$S_AllData = @()

# Loop through each user and retrieve their authentication methods
foreach ($S_Member in $S_AllUsers)
{
    $S_MemberId = $S_Member.Id
    $S_MemberName = $S_Member.DisplayName
    Write-Host "`nWindows Hello for Business Check for $($S_MemberName):"
    Start-Sleep -Milliseconds $S_GraphRequestDelayMilliseconds
    $S_WHfBAuthMethods = Get-MgUserAuthenticationMethod -UserId $S_MemberId | Where-Object { $_.additionalProperties.'@odata.type' -like "*windowsHelloForBusinessAuthenticationMethod*" }
    if ($S_WHfBAuthMethods)
    {
        foreach ($S_Method in $S_WHfBAuthMethods)
        {
            $S_DeviceName = $S_Method.additionalProperties.displayName
            $S_RegisteredDate = $S_Method.additionalProperties.createdDateTime
            Write-Host "Found on $S_DeviceName"

            $S_OBJ = New-Object PSObject
            $S_OBJ | Add-Member -MemberType NoteProperty -Name "StaffDisplayName" -Value $S_MemberName
            $S_OBJ | Add-Member -MemberType NoteProperty -Name "DeviceName" -Value $S_DeviceName
            $S_OBJ | Add-Member -MemberType NoteProperty -Name "RegisteredDate" -Value $S_RegisteredDate

            $S_AllData += $S_OBJ
        }
    }
    else
    {
        Write-Host "No Windows Hello for Business Registrations found!`n"
    }
}

$S_AllData | Export-Csv -Path $OutputPath -NoTypeInformation
Write-Host "Report exported to: $OutputPath" -ForegroundColor Green

$S_DisconnectChoice = Read-Host "Disconnect from Microsoft Graph now? [Y] Yes  [N] No  (Default: N)"
if ($S_DisconnectChoice -eq 'Y')
{
    Disconnect-MgGraph | Out-Null
    Write-Host "Disconnected from Microsoft Graph." -ForegroundColor Yellow
}
