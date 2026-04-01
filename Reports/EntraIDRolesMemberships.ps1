#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Identity.DirectoryManagement

<#
.SYNOPSIS
    Reports on users assigned to Administrator or Global Reader directory roles,
    including their last sign-in date. Exports CSV and HTML.

.PARAMETER OutputPath
    Path for the output CSV file. Defaults to a timestamped file in the current directory.

.PARAMETER Test
    When specified, limits processing for quick testing.

.EXAMPLE
    .\EntraIDRolesMemberships.ps1

.EXAMPLE
    .\EntraIDRolesMemberships.ps1 -OutputPath "C:\Reports\RoleMembers.csv"
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

if (-not $OutputPath) {
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path (Get-Location).Path "EntraIDRoleMemberships_$timestamp.csv"
}

$htmlPath = [System.IO.Path]::ChangeExtension($OutputPath, '.html')

# ── Connect to Microsoft Graph ─────────────────────────────────────────────────
$requiredScopes = @(
    'User.Read.All'
    'Group.Read.All'
    'RoleManagement.Read.Directory'
    'AuditLog.Read.All'
)

$existingContext = Get-MgContext
if ($existingContext) {
    Write-Host "Existing Graph session detected:" -ForegroundColor Yellow
    Write-Host "  Account : $($existingContext.Account)" -ForegroundColor Yellow
    Write-Host "  TenantId: $($existingContext.TenantId)" -ForegroundColor Yellow
    Write-Host "  Scopes  : $($existingContext.Scopes -join ', ')" -ForegroundColor Yellow
    Write-Host ""

    $choice = Read-Host "Use existing session? [Y] Yes  [N] Disconnect and reconnect  (Default: Y)"
    if ($choice -eq 'N') {
        Write-Host "Disconnecting existing session..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Write-Host "Reconnecting with required scopes..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome
        Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
    }
    else {
        Write-Host "Using existing Graph session." -ForegroundColor Green
    }
}
else {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes $requiredScopes -NoWelcome
    Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
}

try {
    # ── Get Privileged Directory Roles ─────────────────────────────────────────
    Write-Host "Retrieving directory roles..." -ForegroundColor Cyan
    $DirectoryRoles = Get-MgDirectoryRole -All

    $PrivilegedRoles = $DirectoryRoles | Where-Object {
        $_.DisplayName -like "*Administrator*" -or $_.DisplayName -eq "Global Reader"
    }

    $roleCount = ($PrivilegedRoles | Measure-Object).Count
    Write-Host "Found $roleCount privileged roles." -ForegroundColor Cyan

    # ── Get Members of Each Role (handles both Users and Groups) ────────────────
    Write-Host "Retrieving role members..." -ForegroundColor Cyan
    # Key = UserId, Value = list of @{ Role; ViaGroup }
    $roleMemberMap = @{}
    $groupViaMap   = @{}  # Key = UserId, Value = list of group display names

    foreach ($role in $PrivilegedRoles) {
        $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All
        foreach ($member in $members) {
            $odataType = $member.AdditionalProperties.'@odata.type'

            if ($odataType -eq '#microsoft.graph.group') {
                # This is a group assigned to the role — expand its members
                $groupName = $member.AdditionalProperties.displayName
                Write-Host "  Expanding group '$groupName' in role '$($role.DisplayName)'..." -ForegroundColor Yellow
                try {
                    $groupMembers = Get-MgGroupMember -GroupId $member.Id -All
                }
                catch {
                    Write-Warning "Could not expand group $($member.Id) ($groupName): $_"
                    $groupMembers = @()
                }
                foreach ($gm in $groupMembers) {
                    $gmType = $gm.AdditionalProperties.'@odata.type'
                    if ($gmType -eq '#microsoft.graph.user') {
                        if (-not $roleMemberMap.ContainsKey($gm.Id)) {
                            $roleMemberMap[$gm.Id] = [System.Collections.Generic.List[string]]::new()
                            $groupViaMap[$gm.Id]   = [System.Collections.Generic.List[string]]::new()
                        }
                        $roleMemberMap[$gm.Id].Add($role.DisplayName)
                        $groupViaMap[$gm.Id].Add($groupName)
                    }
                }
            }
            else {
                # Direct user assignment
                if (-not $roleMemberMap.ContainsKey($member.Id)) {
                    $roleMemberMap[$member.Id] = [System.Collections.Generic.List[string]]::new()
                    $groupViaMap[$member.Id]   = [System.Collections.Generic.List[string]]::new()
                }
                $roleMemberMap[$member.Id].Add($role.DisplayName)
            }
        }
    }

    $uniqueUserIds = $roleMemberMap.Keys
    $userCount = ($uniqueUserIds | Measure-Object).Count
    Write-Host "Found $userCount unique privileged users. Retrieving details..." -ForegroundColor Cyan

    # ── Retrieve User Details ──────────────────────────────────────────────────
    $userProperties = @(
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

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $currentUser = 0

    foreach ($userId in $uniqueUserIds) {
        $currentUser++
        Write-Progress -Activity "Processing Privileged Users" -Status "$currentUser of $userCount" -PercentComplete (($currentUser / $userCount) * 100)

        try {
            $user = Get-MgUser -UserId $userId -Property $userProperties
        }
        catch {
            Write-Warning "Could not retrieve user $userId : $_"
            continue
        }

        # ── Last Sign-In ──────────────────────────────────────────────────────
        $lastInteractive    = $user.SignInActivity.LastSignInDateTime
        $lastNonInteractive = $user.SignInActivity.LastNonInteractiveSignInDateTime

        $dates = @($lastInteractive, $lastNonInteractive) | Where-Object { $_ -ne $null }

        if ($dates.Count -eq 0) {
            $lastSignInDate = 'No Sign-In Recorded'
            $daysSinceSignIn = 'N/A'
        }
        else {
            $lastSignIn = ($dates | Sort-Object -Descending | Select-Object -First 1)
            $lastSignInDate = $lastSignIn.ToString('yyyy-MM-dd HH:mm')
            $daysSinceSignIn = [math]::Floor(((Get-Date) - $lastSignIn).TotalDays)
        }

        # ── Account Status ─────────────────────────────────────────────────────
        $accountStatus = if ($user.AccountEnabled) { 'Enabled' } else { 'Disabled' }

        # ── Licensing & Sync ───────────────────────────────────────────────────
        $isLicensed = ($user.AssignedLicenses | Measure-Object).Count -gt 0
        $isOnPremSynced = $user.OnPremisesSyncEnabled -eq $true

        # ── Roles & Group Assignment ─────────────────────────────────────────
        $assignedRoles = ($roleMemberMap[$userId] | Select-Object -Unique) -join '; '
        $viaGroups = ($groupViaMap[$userId] | Select-Object -Unique) -join '; '
        $assignedViaGroup = if ($viaGroups) { $viaGroups } else { 'Direct' }

        # ── Domain ─────────────────────────────────────────────────────────────
        $domain = ($user.UserPrincipalName -split '@')[1]

        # ── Build result row ───────────────────────────────────────────────────
        $results.Add([PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            Domain            = $domain
            AccountStatus     = $accountStatus
            IsLicensed        = $isLicensed
            IsOnPremSynced    = $isOnPremSynced
            UserType          = $user.UserType
            LastSignIn        = $lastSignInDate
            DaysSinceSignIn   = $daysSinceSignIn
            AssignedViaGroup  = $assignedViaGroup
            Roles             = $assignedRoles
        })
    }

    Write-Progress -Activity "Processing Privileged Users" -Completed

    # Sort results by DisplayName
    $results = $results | Sort-Object DisplayName

    # ── Export CSV ─────────────────────────────────────────────────────────────
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nCSV exported to: $OutputPath" -ForegroundColor Green
    Write-Host "Total rows: $($results.Count)" -ForegroundColor Green

    # ── Calculate Statistics ───────────────────────────────────────────────────
    $totalPrivileged = $results.Count
    $enabledCount    = ($results | Where-Object { $_.AccountStatus -eq 'Enabled' }).Count
    $disabledCount   = ($results | Where-Object { $_.AccountStatus -eq 'Disabled' }).Count
    $onPremCount     = ($results | Where-Object { $_.IsOnPremSynced -eq $true }).Count
    $cloudOnlyCount  = $totalPrivileged - $onPremCount
    $noSignInCount   = ($results | Where-Object { $_.LastSignIn -eq 'No Sign-In Recorded' }).Count
    $viaGroupCount   = ($results | Where-Object { $_.AssignedViaGroup -ne 'Direct' }).Count
    $directCount     = $totalPrivileged - $viaGroupCount

    # Role breakdown
    $roleCounts = @{}
    foreach ($r in $results) {
        foreach ($roleName in ($r.Roles -split '; ')) {
            $roleName = $roleName.Trim()
            if ($roleName) {
                if (-not $roleCounts.ContainsKey($roleName)) { $roleCounts[$roleName] = 0 }
                $roleCounts[$roleName]++
            }
        }
    }

    # Build role breakdown cards HTML
    $roleCardsHtml = ($roleCounts.GetEnumerator() | Sort-Object Name | ForEach-Object {
        "            <div class=`"card purple`"><div class=`"label`">$([System.Web.HttpUtility]::HtmlEncode($_.Key))</div><div class=`"value`">$($_.Value)</div></div>"
    }) -join "`n"

    # Build full table rows
    $tableRows = ($results | ForEach-Object {
        $signInClass = if ($_.LastSignIn -eq 'No Sign-In Recorded') { ' class="warn"' } elseif ($_.DaysSinceSignIn -ne 'N/A' -and [int]$_.DaysSinceSignIn -gt 90) { ' class="warn"' } else { '' }
        $statusClass = if ($_.AccountStatus -eq 'Disabled') { ' class="warn"' } else { '' }
        $groupClass = if ($_.AssignedViaGroup -ne 'Direct') { ' class="group"' } else { '' }
        "        <tr><td>$([System.Web.HttpUtility]::HtmlEncode($_.DisplayName))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.UserPrincipalName))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.Domain))</td><td$statusClass>$($_.AccountStatus)</td><td>$($_.IsLicensed)</td><td>$($_.IsOnPremSynced)</td><td>$($_.UserType)</td><td$signInClass>$($_.LastSignIn)</td><td$signInClass>$($_.DaysSinceSignIn)</td><td$groupClass>$([System.Web.HttpUtility]::HtmlEncode($_.AssignedViaGroup))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.Roles))</td></tr>"
    }) -join "`n"

    # ── Generate HTML Report ───────────────────────────────────────────────────
    $reportDate = Get-Date -Format 'dd MMM yyyy HH:mm'
    $tenantId   = (Get-MgContext).TenantId

    $html = @"
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
        <div class="subtitle">Generated: $reportDate | Tenant: $tenantId</div>
    </div>

    <div class="cards">
        <div class="card blue">
            <div class="label">Total Privileged Users</div>
            <div class="value">$totalPrivileged</div>
        </div>
        <div class="card green">
            <div class="label">Enabled Accounts</div>
            <div class="value">$enabledCount</div>
            <div class="detail">$([math]::Round(($enabledCount / [math]::Max($totalPrivileged,1)) * 100, 1))% of total</div>
        </div>
        <div class="card red">
            <div class="label">Disabled Accounts</div>
            <div class="value">$disabledCount</div>
        </div>
        <div class="card teal">
            <div class="label">Cloud-Only</div>
            <div class="value">$cloudOnlyCount</div>
        </div>
        <div class="card orange">
            <div class="label">On-Prem Synced</div>
            <div class="value">$onPremCount</div>
        </div>
        <div class="card red">
            <div class="label">No Sign-In Recorded</div>
            <div class="value">$noSignInCount</div>
        </div>
        <div class="card purple">
            <div class="label">Assigned via Group</div>
            <div class="value">$viaGroupCount</div>
            <div class="detail">$directCount direct assignment(s)</div>
        </div>
    </div>

    <div class="section">
        <h2>Role Breakdown (User Count per Role)</h2>
        <div class="breakdown">
$roleCardsHtml
        </div>
    </div>

    <div class="section">
        <h2>All Privileged Users</h2>
        <table>
            <thead>
                <tr><th>Display Name</th><th>UPN</th><th>Domain</th><th>Status</th><th>Licensed</th><th>On-Prem Synced</th><th>User Type</th><th>Last Sign-In</th><th>Days Since</th><th>Assigned Via</th><th>Roles</th></tr>
            </thead>
            <tbody>
$tableRows
            </tbody>
        </table>
    </div>

    <div class="footer">
        CSV data: $(Split-Path $OutputPath -Leaf) | Report generated by EntraIDRolesMemberships.ps1
    </div>
</body>
</html>
"@

    $html | Out-File -FilePath $htmlPath -Encoding UTF8
    Write-Host "HTML report exported to: $htmlPath" -ForegroundColor Green
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    $disconnectChoice = Read-Host "`nDisconnect from Microsoft Graph? [Y] Yes  [N] Keep session  (Default: N)"
    if ($disconnectChoice -eq 'Y') {
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Write-Host "Disconnected." -ForegroundColor Green
    }
    else {
        Write-Host "Graph session kept alive." -ForegroundColor Green
    }
}