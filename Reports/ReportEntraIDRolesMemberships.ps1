#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Identity.DirectoryManagement

<#
.SYNOPSIS
    Reports on users assigned to Administrator or Global Reader directory roles,
    including their last sign-in date. Exports CSV and HTML.

.DESCRIPTION
    Connects to Microsoft Graph and retrieves all active Entra ID directory roles.
    For each role, reports the assigned users including their display name, UPN,
    account state, and last sign-in date. Exports results to CSV and HTML.

.PARAMETER OutputPath
    Path for the output CSV file. Defaults to a timestamped file in the current directory.

.PARAMETER Test
    When specified, limits processing for quick testing.

.EXAMPLE
    .\ReportEntraIDRolesMemberships.ps1

.EXAMPLE
    .\ReportEntraIDRolesMemberships.ps1 -OutputPath "C:\Reports\RoleMembers.csv"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]$Test
)

# ── Setup ──────────────────────────────────────────────────────────────────────
$ErrorActionPreference = 'Stop'

$S_OutputPath = $OutputPath

if (-not $S_OutputPath)
{
    $S_Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $S_OutputPath = Join-Path (Get-Location).Path "ReportEntraIDRolesMemberships_$S_Timestamp.csv"
}

$S_HtmlPath = [System.IO.Path]::ChangeExtension($S_OutputPath, '.html')

# ── Connect to Microsoft Graph ─────────────────────────────────────────────────
$S_RequiredGraphScopes = @(
    'User.Read.All'
    'Group.Read.All'
    'RoleManagement.Read.Directory'
    'AuditLog.Read.All'
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
        Write-Host "Disconnecting existing session..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Write-Host "Reconnecting with required scopes..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome
        Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
    }
    else
    {
        Write-Host "Using existing Graph session." -ForegroundColor Green
    }
}
else
{
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome
    Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
}

try
{
    # ── Get Privileged Directory Roles ─────────────────────────────────────────
    Write-Host "Retrieving directory roles..." -ForegroundColor Cyan
    $S_DirectoryRoles = Get-MgDirectoryRole -All

    $S_PrivilegedRoles = $S_DirectoryRoles | Where-Object {
        $_.DisplayName -like "*Administrator*" -or $_.DisplayName -eq "Global Reader"
    }

    $S_RoleCount = ($S_PrivilegedRoles | Measure-Object).Count
    Write-Host "Found $S_RoleCount privileged roles." -ForegroundColor Cyan

    # ── Get Members of Each Role (handles both Users and Groups) ────────────────
    Write-Host "Retrieving role members..." -ForegroundColor Cyan
    # Key = UserId, Value = list of @{ Role; ViaGroup }
    $S_RoleMemberMap = @{}
    $S_GroupViaMap   = @{}  # Key = UserId, Value = list of group display names

    foreach ($S_Role in $S_PrivilegedRoles)
    {
        $S_Members = Get-MgDirectoryRoleMember -DirectoryRoleId $S_Role.Id -All
        foreach ($S_Member in $S_Members)
        {
            $S_OdataType = $S_Member.AdditionalProperties.'@odata.type'

            if ($S_OdataType -eq '#microsoft.graph.group')
            {
                # This is a group assigned to the role — expand its members
                $S_GroupName = $S_Member.AdditionalProperties.displayName
                Write-Host "  Expanding group '$S_GroupName' in role '$($S_Role.DisplayName)'..." -ForegroundColor Yellow
                try
                {
                    $S_GroupMembers = Get-MgGroupMember -GroupId $S_Member.Id -All
                }
                catch
                {
                    Write-Warning "Could not expand group $($S_Member.Id) ($S_GroupName): $_"
                    $S_GroupMembers = @()
                }
                foreach ($S_Gm in $S_GroupMembers)
                {
                    $S_GmType = $S_Gm.AdditionalProperties.'@odata.type'
                    if ($S_GmType -eq '#microsoft.graph.user')
                    {
                        if (-not $S_RoleMemberMap.ContainsKey($S_Gm.Id))
                        {
                            $S_RoleMemberMap[$S_Gm.Id] = [System.Collections.Generic.List[string]]::new()
                            $S_GroupViaMap[$S_Gm.Id]   = [System.Collections.Generic.List[string]]::new()
                        }
                        $S_RoleMemberMap[$S_Gm.Id].Add($S_Role.DisplayName)
                        $S_GroupViaMap[$S_Gm.Id].Add($S_GroupName)
                    }
                }
            }
            else
            {
                # Direct user assignment
                if (-not $S_RoleMemberMap.ContainsKey($S_Member.Id))
                {
                    $S_RoleMemberMap[$S_Member.Id] = [System.Collections.Generic.List[string]]::new()
                    $S_GroupViaMap[$S_Member.Id]   = [System.Collections.Generic.List[string]]::new()
                }
                $S_RoleMemberMap[$S_Member.Id].Add($S_Role.DisplayName)
            }
        }
    }

    $S_UniqueUserIds = $S_RoleMemberMap.Keys
    $S_UserCount = ($S_UniqueUserIds | Measure-Object).Count
    Write-Host "Found $S_UserCount unique privileged users. Retrieving details..." -ForegroundColor Cyan

    # ── Retrieve User Details ──────────────────────────────────────────────────
    $S_UserProperties = @(
        'Id'
        'DisplayName'
        'UserPrincipalName'
        'Mail'
        'AccountEnabled'
        'AssignedLicenses'
        'OnPremisesSyncEnabled'
        'UserType'
        'SignInActivity'
    ) -join ','

    $S_Results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $S_CurrentUser = 0

    foreach ($S_UserId in $S_UniqueUserIds)
    {
        $S_CurrentUser++
        Write-Progress -Activity "Processing Privileged Users" -Status "$S_CurrentUser of $S_UserCount" -PercentComplete (($S_CurrentUser / $S_UserCount) * 100)

        try
        {
            $S_User = Get-MgUser -UserId $S_UserId -Property $S_UserProperties
        }
        catch
        {
            Write-Warning "Could not retrieve user $S_UserId : $_"
            continue
        }

        # ── Last Sign-In ──────────────────────────────────────────────────────
        $S_LastInteractive    = $S_User.SignInActivity.LastSignInDateTime
        $S_LastNonInteractive = $S_User.SignInActivity.LastNonInteractiveSignInDateTime

        $S_Dates = @($S_LastInteractive, $S_LastNonInteractive) | Where-Object { $_ -ne $null }

        if ($S_Dates.Count -eq 0)
        {
            $S_LastSignInDate = 'No Sign-In Recorded'
            $S_DaysSinceSignIn = 'N/A'
        }
        else
        {
            $S_LastSignIn = ($S_Dates | Sort-Object -Descending | Select-Object -First 1)
            $S_LastSignInDate = $S_LastSignIn.ToString('yyyy-MM-dd HH:mm')
            $S_DaysSinceSignIn = [math]::Floor(((Get-Date) - $S_LastSignIn).TotalDays)
        }

        # ── Account Status ─────────────────────────────────────────────────────
        if ($S_User.AccountEnabled)
        {
            $S_AccountStatus = 'Enabled'
        }
        else
        {
            $S_AccountStatus = 'Disabled'
        }

        # ── Licensing & Sync ───────────────────────────────────────────────────
        $S_IsLicensed = ($S_User.AssignedLicenses | Measure-Object).Count -gt 0
        $S_IsOnPremSynced = $S_User.OnPremisesSyncEnabled -eq $true

        # ── Roles & Group Assignment ─────────────────────────────────────────
        $S_AssignedRoles = ($S_RoleMemberMap[$S_UserId] | Select-Object -Unique) -join '; '
        $S_ViaGroups = ($S_GroupViaMap[$S_UserId] | Select-Object -Unique) -join '; '
        if ($S_ViaGroups)
        {
            $S_AssignedViaGroup = $S_ViaGroups
        }
        else
        {
            $S_AssignedViaGroup = 'Direct'
        }

        # ── Domain ─────────────────────────────────────────────────────────────
        $S_Domain = ($S_User.UserPrincipalName -split '@')[1]

        # ── Build result row ───────────────────────────────────────────────────
        $S_Results.Add([PSCustomObject]@{
            DisplayName       = $S_User.DisplayName
            UserPrincipalName = $S_User.UserPrincipalName
            Domain            = $S_Domain
            AccountStatus     = $S_AccountStatus
            IsLicensed        = $S_IsLicensed
            IsOnPremSynced    = $S_IsOnPremSynced
            UserType          = $S_User.UserType
            LastSignIn        = $S_LastSignInDate
            DaysSinceSignIn   = $S_DaysSinceSignIn
            AssignedViaGroup  = $S_AssignedViaGroup
            Roles             = $S_AssignedRoles
        })
    }

    Write-Progress -Activity "Processing Privileged Users" -Completed

    # Sort results by DisplayName
    $S_Results = $S_Results | Sort-Object DisplayName

    # ── Export CSV ─────────────────────────────────────────────────────────────
    $S_Results | Export-Csv -Path $S_OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nCSV exported to: $S_OutputPath" -ForegroundColor Green
    Write-Host "Total rows: $($S_Results.Count)" -ForegroundColor Green

    # ── Calculate Statistics ───────────────────────────────────────────────────
    $S_TotalPrivileged = $S_Results.Count
    $S_EnabledCount    = ($S_Results | Where-Object { $_.AccountStatus -eq 'Enabled' }).Count
    $S_DisabledCount   = ($S_Results | Where-Object { $_.AccountStatus -eq 'Disabled' }).Count
    $S_OnPremCount     = ($S_Results | Where-Object { $_.IsOnPremSynced -eq $true }).Count
    $S_CloudOnlyCount  = $S_TotalPrivileged - $S_OnPremCount
    $S_NoSignInCount   = ($S_Results | Where-Object { $_.LastSignIn -eq 'No Sign-In Recorded' }).Count
    $S_ViaGroupCount   = ($S_Results | Where-Object { $_.AssignedViaGroup -ne 'Direct' }).Count
    $S_DirectCount     = $S_TotalPrivileged - $S_ViaGroupCount

    # Role breakdown
    $S_RoleCounts = @{}
    foreach ($S_R in $S_Results)
    {
        foreach ($S_RoleName in ($S_R.Roles -split '; '))
        {
            $S_RoleName = $S_RoleName.Trim()
            if ($S_RoleName)
            {
                if (-not $S_RoleCounts.ContainsKey($S_RoleName))
                {
                    $S_RoleCounts[$S_RoleName] = 0
                }
                $S_RoleCounts[$S_RoleName]++
            }
        }
    }

    # Build role breakdown cards HTML
    $S_RoleCardsHtml = ($S_RoleCounts.GetEnumerator() | Sort-Object Name | ForEach-Object {
        "            <div class=`"card purple`"><div class=`"label`">$([System.Web.HttpUtility]::HtmlEncode($_.Key))</div><div class=`"value`">$($_.Value)</div></div>"
    }) -join "`n"

    # Build full table rows
    $S_TableRows = ($S_Results | ForEach-Object {
        if ($_.LastSignIn -eq 'No Sign-In Recorded')
        {
            $S_SignInClass = ' class="warn"'
        }
        elseif ($_.DaysSinceSignIn -ne 'N/A' -and [int]$_.DaysSinceSignIn -gt 90)
        {
            $S_SignInClass = ' class="warn"'
        }
        else
        {
            $S_SignInClass = ''
        }

        if ($_.AccountStatus -eq 'Disabled')
        {
            $S_StatusClass = ' class="warn"'
        }
        else
        {
            $S_StatusClass = ''
        }

        if ($_.AssignedViaGroup -ne 'Direct')
        {
            $S_GroupClass = ' class="group"'
        }
        else
        {
            $S_GroupClass = ''
        }

        "        <tr><td>$([System.Web.HttpUtility]::HtmlEncode($_.DisplayName))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.UserPrincipalName))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.Domain))</td><td$S_StatusClass>$($_.AccountStatus)</td><td>$($_.IsLicensed)</td><td>$($_.IsOnPremSynced)</td><td>$($_.UserType)</td><td$S_SignInClass>$($_.LastSignIn)</td><td$S_SignInClass>$($_.DaysSinceSignIn)</td><td$S_GroupClass>$([System.Web.HttpUtility]::HtmlEncode($_.AssignedViaGroup))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.Roles))</td></tr>"
    }) -join "`n"

    # ── Generate HTML Report ───────────────────────────────────────────────────
    $S_ReportDate = Get-Date -Format 'dd MMM yyyy HH:mm'
    $S_TenantId   = (Get-MgContext).TenantId

    $S_Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Entra ID Privileged Role Memberships</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; padding: 24px; }
        .header { text-align: center; margin-bottom: 32px; }
        .header h1 { font-size: 28px; color: #1a1a2e; margin-bottom: 4px; }
        .header .subtitle { font-size: 14px; color: #666; }
        .cards { display: flex; flex-wrap: wrap; gap: 20px; justify-content: center; margin-bottom: 32px; }
        .card {
            background: #fff; border-radius: 12px; padding: 24px 28px; min-width: 220px; flex: 1; max-width: 280px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08); border-left: 5px solid #0078d4; position: relative;
        }
        .card .label { font-size: 13px; color: #666; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px; }
        .card .value { font-size: 36px; font-weight: 700; color: #1a1a2e; }
        .card .detail { font-size: 12px; color: #888; margin-top: 6px; }
        .card.blue    { border-left-color: #0078d4; }
        .card.red     { border-left-color: #d13438; }
        .card.green   { border-left-color: #107c10; }
        .card.orange  { border-left-color: #ff8c00; }
        .card.purple  { border-left-color: #8764b8; }
        .card.teal    { border-left-color: #00b7c3; }
        .section { margin-bottom: 24px; }
        .section h2 { font-size: 18px; color: #1a1a2e; margin-bottom: 12px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0; text-align: center; }
        .breakdown { display: flex; flex-wrap: wrap; gap: 16px; justify-content: center; }
        .breakdown .card { min-width: 180px; max-width: 260px; padding: 18px 22px; }
        .breakdown .card .value { font-size: 28px; }
        table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.08); margin-top: 12px; }
        th { background: #0078d4; color: #fff; text-align: left; padding: 10px 14px; font-size: 13px; text-transform: uppercase; letter-spacing: 0.3px; }
        td { padding: 9px 14px; font-size: 13px; border-bottom: 1px solid #eee; }
        tr:hover { background: #f5f9ff; }
        td.warn { color: #d13438; font-weight: 600; }
        td.group { color: #8764b8; font-style: italic; }
        .footer { text-align: center; font-size: 12px; color: #999; margin-top: 32px; }
    </style>
</head>
<body>
    <div class="header">
        <h1>Entra ID Privileged Role Memberships</h1>
        <div class="subtitle">Generated: $S_ReportDate | Tenant: $S_TenantId</div>
    </div>

    <div class="cards">
        <div class="card blue">
            <div class="label">Total Privileged Users</div>
            <div class="value">$S_TotalPrivileged</div>
        </div>
        <div class="card green">
            <div class="label">Enabled Accounts</div>
            <div class="value">$S_EnabledCount</div>
            <div class="detail">$([math]::Round(($S_EnabledCount / [math]::Max($S_TotalPrivileged,1)) * 100, 1))% of total</div>
        </div>
        <div class="card red">
            <div class="label">Disabled Accounts</div>
            <div class="value">$S_DisabledCount</div>
        </div>
        <div class="card teal">
            <div class="label">Cloud-Only</div>
            <div class="value">$S_CloudOnlyCount</div>
        </div>
        <div class="card orange">
            <div class="label">On-Prem Synced</div>
            <div class="value">$S_OnPremCount</div>
        </div>
        <div class="card red">
            <div class="label">No Sign-In Recorded</div>
            <div class="value">$S_NoSignInCount</div>
        </div>
        <div class="card purple">
            <div class="label">Assigned via Group</div>
            <div class="value">$S_ViaGroupCount</div>
            <div class="detail">$S_DirectCount direct assignment(s)</div>
        </div>
    </div>

    <div class="section">
        <h2>Role Breakdown (User Count per Role)</h2>
        <div class="breakdown">
$S_RoleCardsHtml
        </div>
    </div>

    <div class="section">
        <h2>All Privileged Users</h2>
        <table>
            <thead>
                <tr><th>Display Name</th><th>UPN</th><th>Domain</th><th>Status</th><th>Licensed</th><th>On-Prem Synced</th><th>User Type</th><th>Last Sign-In</th><th>Days Since</th><th>Assigned Via</th><th>Roles</th></tr>
            </thead>
            <tbody>
$S_TableRows
            </tbody>
        </table>
    </div>

    <div class="footer">
        CSV data: $(Split-Path $S_OutputPath -Leaf) | Report generated by ReportEntraIDRolesMemberships.ps1
    </div>
</body>
</html>
"@

    $S_Html | Out-File -FilePath $S_HtmlPath -Encoding UTF8
    Write-Host "HTML report exported to: $S_HtmlPath" -ForegroundColor Green
}
catch
{
    Write-Error "An error occurred: $_"
}
finally
{
    $S_DisconnectChoice = Read-Host "`nDisconnect from Microsoft Graph? [Y] Yes  [N] Keep session  (Default: N)"
    if ($S_DisconnectChoice -eq 'Y')
    {
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Write-Host "Disconnected." -ForegroundColor Green
    }
    else
    {
        Write-Host "Graph session kept alive." -ForegroundColor Green
    }
}
