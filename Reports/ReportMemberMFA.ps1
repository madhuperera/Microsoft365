#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Reports

<#
.SYNOPSIS
    Generates a CSV and HTML report of MFA coverage across all Member users
    (Guests are excluded).

.DESCRIPTION
    Connects to Microsoft Graph, retrieves all Member users (userType eq 'Member',
    excluding Guests), then for each user inspects the registered authentication
    methods to classify MFA status (Modern Auth, Legacy Auth, or No MFA).
    Account activity, licensing and on-premises sync state are also captured.
    Results are exported to CSV and a summary HTML dashboard.

.PARAMETER InactiveDays
    Number of days since last sign-in to consider an account inactive.
    If a user has not signed in within this many days, ActiveAccount reports
    how long ago they last signed in.

.PARAMETER OutputPath
    Path for the output CSV file. Defaults to a timestamped file in the
    current working directory.

.PARAMETER Test
    When specified, only processes the first 10 Member users for quick testing.

.EXAMPLE
    .\ReportMemberMFA.ps1 -InactiveDays 90

.EXAMPLE
    .\ReportMemberMFA.ps1 -InactiveDays 30 -Test
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateRange(1, 3650)]
    [int]$InactiveDays,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]$Test
)

# ── Setup ──────────────────────────────────────────────────────────────────────
$ErrorActionPreference = 'Stop'

if (-not $OutputPath) {
    $S_Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path (Get-Location).Path "ReportMemberMFA_$S_Timestamp.csv"
}

$S_HtmlPath = [System.IO.Path]::ChangeExtension($OutputPath, '.html')

$S_CutoffDate = (Get-Date).AddDays(-$InactiveDays)

# ── Connect to Microsoft Graph ─────────────────────────────────────────────────
$S_RequiredGraphScopes = @(
    'User.Read.All'
    'AuditLog.Read.All'
    'UserAuthenticationMethod.Read.All'
)

$S_ExistingContext = Get-MgContext
if ($S_ExistingContext) {
    Write-Host "Existing Graph session detected:" -ForegroundColor Yellow
    Write-Host "  Account : $($S_ExistingContext.Account)" -ForegroundColor Yellow
    Write-Host "  TenantId: $($S_ExistingContext.TenantId)" -ForegroundColor Yellow
    Write-Host "  Scopes  : $($S_ExistingContext.Scopes -join ', ')" -ForegroundColor Yellow
    Write-Host ""

    $S_Choice = Read-Host "Use existing session? [Y] Yes  [N] Disconnect and reconnect  (Default: Y)"
    if ($S_Choice -eq 'N') {
        Write-Host "Disconnecting existing session..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Write-Host "Reconnecting with required scopes..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome
        Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
    }
    else {
        Write-Host "Using existing Graph session." -ForegroundColor Green
    }
}
else {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome
    Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
}

try {
    # ── Retrieve Member Users (Guests excluded) ────────────────────────────────
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

    Write-Host "Retrieving member users (Guests excluded)..." -ForegroundColor Cyan

    if ($Test) {
        Write-Host "[TEST MODE] Limiting to first 10 Member users." -ForegroundColor Yellow
        $users = Get-MgUser -Filter "userType eq 'Member'" -Property $userProperties -Top 10
    }
    else {
        $users = Get-MgUser -Filter "userType eq 'Member'" -Property $userProperties -All
    }

    $userCount = ($users | Measure-Object).Count
    Write-Host "Found $userCount Member users. Processing..." -ForegroundColor Cyan

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

        # ── Determine Active Account ───────────────────────────────────────────
        $lastInteractive    = $null
        $lastNonInteractive = $null
        if (-not $user.AccountEnabled) {
            $activeAccount = 'Disabled'
        }
        else {
            $lastInteractive    = $user.SignInActivity.LastSignInDateTime
            $lastNonInteractive = $user.SignInActivity.LastNonInteractiveSignInDateTime

            $dates = @($lastInteractive, $lastNonInteractive) | Where-Object { $_ -ne $null }

            if ($dates.Count -eq 0) {
                $activeAccount = 'No Sign-In Recorded'
            }
            else {
                $lastSignIn = ($dates | Sort-Object -Descending | Select-Object -First 1)
                if ($lastSignIn -ge $S_CutoffDate) {
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
            UserType          = $user.UserType
            MFA_Registered    = $mfaRegistered
            ActiveAccount     = $activeAccount
            IsLicensed        = $isLicensed
            IsOnPremSynced    = $isOnPremSynced
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

    $coveragePercent = if ($enabledCount -gt 0) {
        [math]::Round(($enabledHasMFA / $enabledCount) * 100, 1)
    } else { 0 }

    # Build table rows for enabled users without MFA
    $noMfaUsers = $enabledUsers | Where-Object { $_.MFA_Registered -eq 'No MFA' }
    $tableRows = ($noMfaUsers | ForEach-Object {
        "        <tr><td>$([System.Web.HttpUtility]::HtmlEncode($_.DisplayName))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.UserPrincipalName))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.Mail))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.Domain))</td><td>$($_.ActiveAccount)</td><td>$($_.IsLicensed)</td><td>$($_.IsOnPremSynced)</td></tr>"
    }) -join "`n"

    # ── Generate HTML Report ───────────────────────────────────────────────────
    $reportDate = Get-Date -Format 'dd MMM yyyy HH:mm'
    $S_TenantId = (Get-MgContext).TenantId

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Member MFA Coverage Report</title>
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
        <h1>Member MFA Coverage Report</h1>
        <div class="subtitle">Generated: $reportDate | Tenant: $S_TenantId | Inactive threshold: $InactiveDays days | Guests excluded</div>
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
        <div class="card purple">
            <div class="label">MFA Coverage (Enabled)</div>
            <div class="value">$coveragePercent%</div>
            <div class="detail">$enabledHasMFA of $enabledCount enabled members</div>
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

    <div class="section">
        <h2>Enabled Member Users without MFA</h2>
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
        CSV data: $(Split-Path $OutputPath -Leaf) | Report generated by ReportMemberMFA.ps1
    </div>
</body>
</html>
"@

    $html | Out-File -FilePath $S_HtmlPath -Encoding UTF8
    Write-Host "HTML report exported to: $S_HtmlPath" -ForegroundColor Green
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    # ── Disconnect ─────────────────────────────────────────────────────────────
    $S_DisconnectChoice = Read-Host "`nDisconnect from Microsoft Graph? [Y] Yes  [N] Keep session  (Default: N)"
    if ($S_DisconnectChoice -eq 'Y') {
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Write-Host "Disconnected." -ForegroundColor Green
    }
    else {
        Write-Host "Graph session kept alive." -ForegroundColor Green
    }
}
