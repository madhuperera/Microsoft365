#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Establishes a single read-only Microsoft Graph connection for use by the
    reporting scripts in the Microsoft365\Reports folder.

.DESCRIPTION
    This helper script is intended to be dot-sourced or run once at the start
    of a reporting session, before any of the other Report*.ps1 scripts in
    this folder are executed.

    It connects to Microsoft Graph using a consolidated set of read-only
    delegated scopes derived from the Report*.ps1 scripts in this folder.
    Only *.Read.All and *.ReadBasic.All scopes are requested. No write,
    update, delete, manage, assign, or modify scopes are included.

    If a Microsoft Graph context already exists, the script displays the
    current Account, Tenant ID, Environment, and Scopes and asks the
    operator to confirm whether to keep using it or reconnect with the
    read-only scope set.

.PARAMETER Force
    Disconnects any existing Microsoft Graph context and connects again
    with the read-only scope set without prompting.

.EXAMPLE
    PS> . .\_ReadOnlyConnectionScript.ps1

    Dot-sources the script so the established Microsoft Graph context is
    available to subsequent Report*.ps1 scripts in the same session.

.EXAMPLE
    PS> .\_ReadOnlyConnectionScript.ps1 -Force

    Forces a fresh Microsoft Graph connection using the read-only scope
    set, regardless of any existing context.

.NOTES
    Run this script once before running the other Report*.ps1 scripts in
    the Microsoft365\Reports folder so they can reuse the existing
    Microsoft Graph context instead of prompting for consent each time.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Read-only Microsoft Graph scopes
#
# This array is the consolidated, de-duplicated, read-only scope set
# required by the Report*.ps1 scripts in this folder. Only *.Read.All and
# *.ReadBasic.All scopes are listed. Write, update, delete, manage, assign,
# and modify scopes are intentionally excluded.
# ---------------------------------------------------------------------------
$S_ReadOnlyGraphScopes = @(
    'Application.Read.All'
    'AuditLog.Read.All'
    'Device.Read.All'
    'DeviceManagementApps.Read.All'
    'DeviceManagementConfiguration.Read.All'
    'DeviceManagementManagedDevices.Read.All'
    'DeviceManagementServiceConfig.Read.All'
    'Directory.Read.All'
    'Domain.Read.All'
    'Group.Read.All'
    'GroupMember.Read.All'
    'Organization.Read.All'
    'Policy.Read.All'
    'RoleManagement.Read.Directory'
    'Team.ReadBasic.All'
    'User.Read.All'
    'UserAuthenticationMethod.Read.All'
)

# ---------------------------------------------------------------------------
# Inspect any existing Microsoft Graph context
# ---------------------------------------------------------------------------
$S_ExistingContext = $null
try
{
    $S_ExistingContext = Get-MgContext
}
catch
{
    $S_ExistingContext = $null
}

if ($S_ExistingContext -and -not $Force)
{
    Write-Host ""
    Write-Host "An existing Microsoft Graph connection was found:" -ForegroundColor Cyan
    Write-Host ("  Account     : {0}" -f $S_ExistingContext.Account)
    Write-Host ("  Tenant ID   : {0}" -f $S_ExistingContext.TenantId)
    Write-Host ("  Environment : {0}" -f $S_ExistingContext.Environment)
    Write-Host ("  Scopes      : {0}" -f (($S_ExistingContext.Scopes | Sort-Object) -join ', '))
    Write-Host ""

    $S_Answer = Read-Host "Continue using this existing Microsoft Graph connection? [Y/N]"
    if ($S_Answer -match '^(Y|y)')
    {
        Write-Host "Continuing with the existing Microsoft Graph connection." -ForegroundColor Green
        return
    }

    Write-Host "Disconnecting the existing Microsoft Graph session before reconnecting..." -ForegroundColor Yellow
    try
    {
        Disconnect-MgGraph | Out-Null
    }
    catch
    {
        Write-Warning "Failed to cleanly disconnect the existing Microsoft Graph session: $($_.Exception.Message)"
    }
}
elseif ($S_ExistingContext -and $Force)
{
    Write-Host "Force specified. Disconnecting the existing Microsoft Graph session..." -ForegroundColor Yellow
    try
    {
        Disconnect-MgGraph | Out-Null
    }
    catch
    {
        Write-Warning "Failed to cleanly disconnect the existing Microsoft Graph session: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# Connect to Microsoft Graph using the read-only scope set
# ---------------------------------------------------------------------------
Write-Host "Connecting to Microsoft Graph with the read-only scope set..." -ForegroundColor Cyan
try
{
    Connect-MgGraph -Scopes $S_ReadOnlyGraphScopes -NoWelcome
}
catch
{
    throw "Connect-MgGraph failed: $($_.Exception.Message)"
}

# ---------------------------------------------------------------------------
# Confirm the active Microsoft Graph context
# ---------------------------------------------------------------------------
$S_NewContext = Get-MgContext
if (-not $S_NewContext)
{
    throw "Microsoft Graph context could not be retrieved after Connect-MgGraph."
}

Write-Host ""
Write-Host "Connected to Microsoft Graph:" -ForegroundColor Green
Write-Host ("  Account     : {0}" -f $S_NewContext.Account)
Write-Host ("  Tenant ID   : {0}" -f $S_NewContext.TenantId)
Write-Host ("  Environment : {0}" -f $S_NewContext.Environment)
Write-Host ("  Scopes      : {0}" -f (($S_NewContext.Scopes | Sort-Object) -join ', '))
Write-Host ""
Write-Host "You can now run the Report*.ps1 scripts in this folder; they will reuse this connection." -ForegroundColor Green
