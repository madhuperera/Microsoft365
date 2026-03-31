#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Users, Microsoft.Graph.Reports

<#
.SYNOPSIS
    Generates a CSV report of all member users with MFA registration, active account status,
    licensing, sync status, and group memberships.

.DESCRIPTION
    Connects to Microsoft Graph, retrieves all Member users (userType eq 'Member'),
    checks sign-in activity, authentication methods, license assignment, on-premises sync,
    and group memberships for each user, then exports a CSV report.

.PARAMETER InactiveDays
    Number of days since last sign-in to consider an account inactive.
    If a user has not signed in within this many days, ActiveAccount = False.

.PARAMETER OutputPath
    Path for the output CSV file. Defaults to a timestamped file in the script directory.

.PARAMETER TargetGroupPrefix
    Prefix to match group display names against.
    Defaults to 'EID-SEC-U-A-ROLE-NoMFA:'.
    Any group whose name starts with this prefix will be reported.

.PARAMETER Test
    When specified, only processes the first 10 users for quick testing.

.EXAMPLE
    .\Get-NonMFAReport.ps1 -InactiveDays 90

.EXAMPLE
    .\Get-NonMFAReport.ps1 -InactiveDays 30 -Test
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateRange(1, 3650)]
    [int]$InactiveDays,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$TargetGroupPrefix = 'EID-SEC-U-A-ROLE-NoMFA:',

    [Parameter(Mandatory = $false)]
    [switch]$Test
)

# ── Setup ──────────────────────────────────────────────────────────────────────
$ErrorActionPreference = 'Stop'

if (-not $OutputPath) {
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path (Get-Location).Path "NonMFAReport_$timestamp.csv"
}

$htmlPath = [System.IO.Path]::ChangeExtension($OutputPath, '.html')

$cutoffDate = (Get-Date).AddDays(-$InactiveDays)

# ── Connect to Microsoft Graph ─────────────────────────────────────────────────
$requiredScopes = @(
    'User.Read.All'
    'Group.Read.All'
    'AuditLog.Read.All'
    'UserAuthenticationMethod.Read.All'
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
    # ── Retrieve Member Users ──────────────────────────────────────────────────
    $userProperties = @(
        'Id'
        'DisplayName'
        'UserPrincipalName'
        'Mail'
        'SignInActivity'
        'AccountEnabled'
        'AssignedLicenses'
        'OnPremisesSyncEnabled'
        'UserType'
    ) -join ','

    Write-Host "Retrieving member users..." -ForegroundColor Cyan

    if ($Test) {
        Write-Host "[TEST MODE] Limiting to 10 users." -ForegroundColor Yellow
        $users = Get-MgUser -Filter "userType eq 'Member'" -Property $userProperties -Top 10
    }
    else {
        $users = Get-MgUser -Filter "userType eq 'Member'" -Property $userProperties -All
    }

    $userCount = ($users | Measure-Object).Count
    Write-Host "Found $userCount member users. Processing..." -ForegroundColor Cyan

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $currentUser = 0

    foreach ($user in $users) {
        $currentUser++
        Write-Progress -Activity "Processing Users" -Status "$currentUser of $userCount - $($user.DisplayName)" -PercentComplete (($currentUser / $userCount) * 100)

        # ── Authentication Methods ─────────────────────────────────────────────
        $authMethodError = $false
        try {
            $authMethods = Get-MgUserAuthenticationMethod -UserId $user.Id
        }
        catch {
            if ($_ -match 'accessDenied|403|Forbidden|Authorization failed') {
                Write-Warning "Access denied reading auth methods for $($user.UserPrincipalName) (privileged account?)"
                $authMethodError = $true
            }
            else {
                Write-Warning "Could not retrieve auth methods for $($user.UserPrincipalName): $_"
            }
            $authMethods = @()
        }

        # ── Group Membership (Target Group Prefix) ─────────────────────────────
        try {
            $memberOf = Get-MgUserMemberOf -UserId $user.Id -All
            $allGroupNames = $memberOf | Where-Object {
                $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group'
            } | ForEach-Object { $_.AdditionalProperties.displayName }

            $matchedGroups = @($allGroupNames | Where-Object { $_ -like "$TargetGroupPrefix*" })
            $matchCount = $matchedGroups.Count

            if ($matchCount -eq 0) {
                $groupResult = 'No Group'
            }
            elseif ($matchCount -eq 1) {
                $groupResult = $matchedGroups[0]
            }
            else {
                $groupResult = "WARNING: Multiple groups ($($matchedGroups -join '; '))"
            }
        }
        catch {
            Write-Warning "Could not retrieve groups for $($user.UserPrincipalName): $_"
            $groupResult = 'Error'
        }

        # ── Determine MFA Registration ─────────────────────────────────────────
        $authTypes = $authMethods | ForEach-Object { $_.AdditionalProperties.'@odata.type' }

        $hasModernAuth = $authTypes | Where-Object {
            $_ -in @(
                '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod'
                '#microsoft.graph.fido2AuthenticationMethod'
                '#microsoft.graph.softwareOathAuthenticationMethod'
            )
        }
        $hasLegacyAuth = $authTypes | Where-Object {
            $_ -in @(
                '#microsoft.graph.phoneAuthenticationMethod'
            )
        }

        if ($authMethodError) {
            $mfaRegistered = 'Access Denied (Privileged Account)'
        }
        elseif ($hasModernAuth) {
            $mfaRegistered = 'Modern Auth'
        }
        elseif ($hasLegacyAuth) {
            $mfaRegistered = 'Legacy Auth'
        }
        else {
            $mfaRegistered = 'No MFA'
        }

        # ── Determine Active Account ───────────────────────────────────────────        $lastInteractive    = $null
        $lastNonInteractive = $null
        if (-not $user.AccountEnabled) {
            $activeAccount = 'Disabled'
        }
        else {
            # Get the most recent sign-in from both interactive and non-interactive
            $lastInteractive    = $user.SignInActivity.LastSignInDateTime
            $lastNonInteractive = $user.SignInActivity.LastNonInteractiveSignInDateTime

            $dates = @($lastInteractive, $lastNonInteractive) | Where-Object { $_ -ne $null }

            if ($dates.Count -eq 0) {
                $activeAccount = 'No Sign-In Recorded'
            }
            else {
                $lastSignIn = ($dates | Sort-Object -Descending | Select-Object -First 1)
                if ($lastSignIn -ge $cutoffDate) {
                    $activeAccount = 'Yes'
                }
                else {
                    $daysAgo = [math]::Floor(((Get-Date) - $lastSignIn).TotalDays)
                    $activeAccount = "${daysAgo}+ Days Ago"
                }
            }
        }

        # ── Licensing & Sync ───────────────────────────────────────────────────
        $isLicensed = ($user.AssignedLicenses | Measure-Object).Count -gt 0
        $isOnPremSynced = $user.OnPremisesSyncEnabled -eq $true

        # ── Mail & Domain ──────────────────────────────────────────────────────
        $mail = if ($user.Mail) { $user.Mail } else { 'None' }
        $domain = if ($user.Mail) {
            ($user.Mail -split '@')[1]
        } else {
            ($user.UserPrincipalName -split '@')[1]
        }

        # ── Build result row ───────────────────────────────────────────────────
        $results.Add([PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            Mail              = $mail
            Domain            = $domain
            MFA_Registered    = $mfaRegistered
            ActiveAccount     = $activeAccount
            IsLicensed        = $isLicensed
            IsOnPremSynced    = $isOnPremSynced
            Group             = $groupResult
        })
    }

    Write-Progress -Activity "Processing Users" -Completed

    # ── Export CSV ─────────────────────────────────────────────────────────────
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nCSV exported to: $OutputPath" -ForegroundColor Green
    Write-Host "Total rows: $($results.Count)" -ForegroundColor Green

    # ── Calculate Statistics ───────────────────────────────────────────────────
    $totalMembers   = $results.Count
    $disabledCount  = ($results | Where-Object { $_.ActiveAccount -eq 'Disabled' }).Count
    $enabledCount   = $totalMembers - $disabledCount

    $enabledUsers       = $results | Where-Object { $_.ActiveAccount -ne 'Disabled' }
    $enabledModernAuth  = ($enabledUsers | Where-Object { $_.MFA_Registered -eq 'Modern Auth' }).Count
    $enabledLegacyAuth  = ($enabledUsers | Where-Object { $_.MFA_Registered -eq 'Legacy Auth' }).Count
    $enabledNoMFA       = ($enabledUsers | Where-Object { $_.MFA_Registered -eq 'No MFA' }).Count
    $enabledHasMFA      = $enabledModernAuth + $enabledLegacyAuth

    $noMfaNoGroup = ($enabledUsers | Where-Object {
        $_.MFA_Registered -eq 'No MFA' -and $_.Group -eq 'No Group'
    }).Count

    # Build table rows for No MFA & No Group users
    $noMfaNoGroupUsers = $enabledUsers | Where-Object {
        $_.MFA_Registered -eq 'No MFA' -and $_.Group -eq 'No Group'
    }
    $tableRows = ($noMfaNoGroupUsers | ForEach-Object {
        "        <tr><td>$([System.Web.HttpUtility]::HtmlEncode($_.DisplayName))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.UserPrincipalName))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.Mail))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.Domain))</td><td>$($_.ActiveAccount)</td><td>$($_.IsLicensed)</td><td>$($_.IsOnPremSynced)</td></tr>"
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
    <title>Non-MFA Report</title>
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
        .card.alert   { border-left-color: #d13438; background: #fdf2f2; }
        .card.alert .value { color: #d13438; }
        .section { margin-bottom: 24px; }
        .section h2 { font-size: 18px; color: #1a1a2e; margin-bottom: 12px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0; text-align: center; }
        .breakdown { display: flex; flex-wrap: wrap; gap: 16px; justify-content: center; }
        .breakdown .card { min-width: 180px; max-width: 240px; padding: 18px 22px; }
        .breakdown .card .value { font-size: 28px; }
        table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.08); margin-top: 12px; }
        th { background: #d13438; color: #fff; text-align: left; padding: 10px 14px; font-size: 13px; text-transform: uppercase; letter-spacing: 0.3px; }
        td { padding: 9px 14px; font-size: 13px; border-bottom: 1px solid #eee; }
        tr:hover { background: #fdf2f2; }
        .footer { text-align: center; font-size: 12px; color: #999; margin-top: 32px; }
    </style>
</head>
<body>
    <div class="header">
        <h1>Non-MFA User Report</h1>
        <div class="subtitle">Generated: $reportDate | Tenant: $tenantId | Inactive threshold: $InactiveDays days</div>
    </div>

    <div class="cards">
        <div class="card blue">
            <div class="label">Total Member Users</div>
            <div class="value">$totalMembers</div>
        </div>
        <div class="card green">
            <div class="label">Enabled Accounts</div>
            <div class="value">$enabledCount</div>
            <div class="detail">$([math]::Round(($enabledCount / [math]::Max($totalMembers,1)) * 100, 1))% of total</div>
        </div>
        <div class="card red">
            <div class="label">Disabled Accounts</div>
            <div class="value">$disabledCount</div>
            <div class="detail">$([math]::Round(($disabledCount / [math]::Max($totalMembers,1)) * 100, 1))% of total</div>
        </div>
    </div>

    <div class="section">
        <h2>Enabled Accounts &mdash; MFA Breakdown</h2>
        <div class="breakdown">
            <div class="card green">
                <div class="label">Modern Auth (MFA)</div>
                <div class="value">$enabledModernAuth</div>
                <div class="detail">Authenticator / Passkey / TOTP</div>
            </div>
            <div class="card orange">
                <div class="label">Legacy Auth (MFA)</div>
                <div class="value">$enabledLegacyAuth</div>
                <div class="detail">SMS / Voice</div>
            </div>
            <div class="card purple">
                <div class="label">Total with MFA</div>
                <div class="value">$enabledHasMFA</div>
                <div class="detail">$([math]::Round(($enabledHasMFA / [math]::Max($enabledCount,1)) * 100, 1))% of enabled</div>
            </div>
            <div class="card red">
                <div class="label">No MFA Registered</div>
                <div class="value">$enabledNoMFA</div>
                <div class="detail">$([math]::Round(($enabledNoMFA / [math]::Max($enabledCount,1)) * 100, 1))% of enabled</div>
            </div>
        </div>
    </div>

    <div class="cards">
        <div class="card alert">
            <div class="label">No MFA &amp; No Target Group</div>
            <div class="value">$noMfaNoGroup</div>
            <div class="detail">Enabled accounts with no MFA and not in any $TargetGroupPrefix group &mdash; requires attention</div>
        </div>
    </div>

    <div class="section">
        <h2>Users with No MFA &amp; No Target Group</h2>
        <table>
            <thead>
                <tr><th>Display Name</th><th>UPN</th><th>Mail</th><th>Domain</th><th>Active</th><th>Licensed</th><th>On-Prem Synced</th></tr>
            </thead>
            <tbody>
$tableRows
            </tbody>
        </table>
    </div>

    <div class="footer">
        CSV data: $(Split-Path $OutputPath -Leaf) | Report generated by Get-NonMFAReport.ps1
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
    # ── Disconnect ─────────────────────────────────────────────────────────────
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
