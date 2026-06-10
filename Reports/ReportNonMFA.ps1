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
    .\ReportNonMFA.ps1 -InactiveDays 90

.EXAMPLE
    .\ReportNonMFA.ps1 -InactiveDays 30 -Test
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

if (-not $OutputPath)
{
    $S_Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path (Get-Location).Path "ReportNonMFA_$S_Timestamp.csv"
}

$S_HtmlPath = [System.IO.Path]::ChangeExtension($OutputPath, '.html')

$S_CutoffDate = (Get-Date).AddDays(-$InactiveDays)

# ── Connect to Microsoft Graph ─────────────────────────────────────────────────
$S_RequiredGraphScopes = @(
    'User.Read.All'
    'Group.Read.All'
    'AuditLog.Read.All'
    'UserAuthenticationMethod.Read.All'
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
    # ── Retrieve Member Users ──────────────────────────────────────────────────
    $S_UserProperties = @(
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

    if ($Test)
    {
        Write-Host "[TEST MODE] Limiting to 10 users." -ForegroundColor Yellow
        $S_Users = Get-MgUser -Filter "userType eq 'Member'" -Property $S_UserProperties -Top 10
    }
    else
    {
        $S_Users = Get-MgUser -Filter "userType eq 'Member'" -Property $S_UserProperties -All
    }

    $S_UserCount = ($S_Users | Measure-Object).Count
    Write-Host "Found $S_UserCount member users. Processing..." -ForegroundColor Cyan

    $S_Results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $S_CurrentUser = 0

    foreach ($S_User in $S_Users)
    {
        $S_CurrentUser++
        Write-Progress -Activity "Processing Users" -Status "$S_CurrentUser of $S_UserCount - $($S_User.DisplayName)" -PercentComplete (($S_CurrentUser / $S_UserCount) * 100)

        # ── Authentication Methods ─────────────────────────────────────────────
        $S_AuthMethodError = $false
        try
        {
            $S_AuthMethods = Get-MgUserAuthenticationMethod -UserId $S_User.Id
        }
        catch
        {
            if ($_ -match 'accessDenied|403|Forbidden|Authorization failed')
            {
                Write-Warning "Access denied reading auth methods for $($S_User.UserPrincipalName) (privileged account?)"
                $S_AuthMethodError = $true
            }
            else
            {
                Write-Warning "Could not retrieve auth methods for $($S_User.UserPrincipalName): $_"
            }

            $S_AuthMethods = @()
        }

        # ── Group Membership (Target Group Prefix) ─────────────────────────────
        try
        {
            $S_MemberOf = Get-MgUserMemberOf -UserId $S_User.Id -All
            $S_AllGroupNames = $S_MemberOf | Where-Object {
                $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group'
            } | ForEach-Object { $_.AdditionalProperties.displayName }

            $S_MatchedGroups = @($S_AllGroupNames | Where-Object { $_ -like "$TargetGroupPrefix*" })
            $S_MatchCount = $S_MatchedGroups.Count

            if ($S_MatchCount -eq 0)
            {
                $S_GroupResult = 'No Group'
            }
            elseif ($S_MatchCount -eq 1)
            {
                $S_GroupResult = $S_MatchedGroups[0]
            }
            else
            {
                $S_GroupResult = "WARNING: Multiple groups ($($S_MatchedGroups -join '; '))"
            }
        }
        catch
        {
            Write-Warning "Could not retrieve groups for $($S_User.UserPrincipalName): $_"
            $S_GroupResult = 'Error'
        }

        # ── Determine MFA Registration ─────────────────────────────────────────
        $S_AuthTypes = $S_AuthMethods | ForEach-Object { $_.AdditionalProperties.'@odata.type' }

        $S_HasModernAuth = $S_AuthTypes | Where-Object {
            $_ -in @(
                '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod'
                '#microsoft.graph.fido2AuthenticationMethod'
                '#microsoft.graph.softwareOathAuthenticationMethod'
            )
        }
        $S_HasLegacyAuth = $S_AuthTypes | Where-Object {
            $_ -in @(
                '#microsoft.graph.phoneAuthenticationMethod'
            )
        }

        if ($S_AuthMethodError)
        {
            $S_MfaRegistered = 'Access Denied (Privileged Account)'
        }
        elseif ($S_HasModernAuth)
        {
            $S_MfaRegistered = 'Modern Auth'
        }
        elseif ($S_HasLegacyAuth)
        {
            $S_MfaRegistered = 'Legacy Auth'
        }
        else
        {
            $S_MfaRegistered = 'No MFA'
        }

        # ── Determine Active Account ───────────────────────────────────────────
        $S_LastInteractive = $null
        $S_LastNonInteractive = $null

        if (-not $S_User.AccountEnabled)
        {
            $S_ActiveAccount = 'Disabled'
        }
        else
        {
            # Get the most recent sign-in from both interactive and non-interactive
            $S_LastInteractive = $S_User.SignInActivity.LastSignInDateTime
            $S_LastNonInteractive = $S_User.SignInActivity.LastNonInteractiveSignInDateTime

            $S_Dates = @($S_LastInteractive, $S_LastNonInteractive) | Where-Object { $_ -ne $null }

            if ($S_Dates.Count -eq 0)
            {
                $S_ActiveAccount = 'No Sign-In Recorded'
            }
            else
            {
                $S_LastSignIn = ($S_Dates | Sort-Object -Descending | Select-Object -First 1)
                if ($S_LastSignIn -ge $S_CutoffDate)
                {
                    $S_ActiveAccount = 'Yes'
                }
                else
                {
                    $S_DaysAgo = [math]::Floor(((Get-Date) - $S_LastSignIn).TotalDays)
                    $S_ActiveAccount = "${S_DaysAgo}+ Days Ago"
                }
            }
        }

        # ── Licensing & Sync ───────────────────────────────────────────────────
        $S_IsLicensed = ($S_User.AssignedLicenses | Measure-Object).Count -gt 0
        $S_IsOnPremSynced = $S_User.OnPremisesSyncEnabled -eq $true

        # ── Mail & Domain ──────────────────────────────────────────────────────
        if ($S_User.Mail)
        {
            $S_Mail = $S_User.Mail
            $S_Domain = ($S_User.Mail -split '@')[1]
        }
        else
        {
            $S_Mail = 'None'
            $S_Domain = ($S_User.UserPrincipalName -split '@')[1]
        }

        # ── Build result row ───────────────────────────────────────────────────
        $S_Results.Add([PSCustomObject]@{
            DisplayName       = $S_User.DisplayName
            UserPrincipalName = $S_User.UserPrincipalName
            Mail              = $S_Mail
            Domain            = $S_Domain
            MFA_Registered    = $S_MfaRegistered
            ActiveAccount     = $S_ActiveAccount
            IsLicensed        = $S_IsLicensed
            IsOnPremSynced    = $S_IsOnPremSynced
            Group             = $S_GroupResult
        })
    }

    Write-Progress -Activity "Processing Users" -Completed

    # ── Export CSV ─────────────────────────────────────────────────────────────
    $S_Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nCSV exported to: $OutputPath" -ForegroundColor Green
    Write-Host "Total rows: $($S_Results.Count)" -ForegroundColor Green

    # ── Calculate Statistics ───────────────────────────────────────────────────
    $S_TotalMembers = $S_Results.Count
    $S_DisabledCount = ($S_Results | Where-Object { $_.ActiveAccount -eq 'Disabled' }).Count
    $S_EnabledCount = $S_TotalMembers - $S_DisabledCount

    $S_EnabledUsers = $S_Results | Where-Object { $_.ActiveAccount -ne 'Disabled' }
    $S_EnabledModernAuth = ($S_EnabledUsers | Where-Object { $_.MFA_Registered -eq 'Modern Auth' }).Count
    $S_EnabledLegacyAuth = ($S_EnabledUsers | Where-Object { $_.MFA_Registered -eq 'Legacy Auth' }).Count
    $S_EnabledNoMFA = ($S_EnabledUsers | Where-Object { $_.MFA_Registered -eq 'No MFA' }).Count
    $S_EnabledHasMFA = $S_EnabledModernAuth + $S_EnabledLegacyAuth

    $S_NoMfaNoGroup = ($S_EnabledUsers | Where-Object {
        $_.MFA_Registered -eq 'No MFA' -and $_.Group -eq 'No Group'
    }).Count

    # Build table rows for No MFA & No Group users
    $S_NoMfaNoGroupUsers = $S_EnabledUsers | Where-Object {
        $_.MFA_Registered -eq 'No MFA' -and $_.Group -eq 'No Group'
    }
    $S_TableRows = ($S_NoMfaNoGroupUsers | ForEach-Object {
        "        <tr><td>$([System.Web.HttpUtility]::HtmlEncode($_.DisplayName))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.UserPrincipalName))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.Mail))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.Domain))</td><td>$($_.ActiveAccount)</td><td>$($_.IsLicensed)</td><td>$($_.IsOnPremSynced)</td></tr>"
    }) -join "`n"

    # ── Generate HTML Report ───────────────────────────────────────────────────
    $S_ReportDate = Get-Date -Format 'dd MMM yyyy HH:mm'
    $S_TenantId = (Get-MgContext).TenantId

    $S_Html = @"
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
        <div class="subtitle">Generated: $S_ReportDate | Tenant: $S_TenantId | Inactive threshold: $InactiveDays days</div>
    </div>

    <div class="cards">
        <div class="card blue">
            <div class="label">Total Member Users</div>
            <div class="value">$S_TotalMembers</div>
        </div>
        <div class="card green">
            <div class="label">Enabled Accounts</div>
            <div class="value">$S_EnabledCount</div>
            <div class="detail">$([math]::Round(($S_EnabledCount / [math]::Max($S_TotalMembers,1)) * 100, 1))% of total</div>
        </div>
        <div class="card red">
            <div class="label">Disabled Accounts</div>
            <div class="value">$S_DisabledCount</div>
            <div class="detail">$([math]::Round(($S_DisabledCount / [math]::Max($S_TotalMembers,1)) * 100, 1))% of total</div>
        </div>
    </div>

    <div class="section">
        <h2>Enabled Accounts &mdash; MFA Breakdown</h2>
        <div class="breakdown">
            <div class="card green">
                <div class="label">Modern Auth (MFA)</div>
                <div class="value">$S_EnabledModernAuth</div>
                <div class="detail">Authenticator / Passkey / TOTP</div>
            </div>
            <div class="card orange">
                <div class="label">Legacy Auth (MFA)</div>
                <div class="value">$S_EnabledLegacyAuth</div>
                <div class="detail">SMS / Voice</div>
            </div>
            <div class="card purple">
                <div class="label">Total with MFA</div>
                <div class="value">$S_EnabledHasMFA</div>
                <div class="detail">$([math]::Round(($S_EnabledHasMFA / [math]::Max($S_EnabledCount,1)) * 100, 1))% of enabled</div>
            </div>
            <div class="card red">
                <div class="label">No MFA Registered</div>
                <div class="value">$S_EnabledNoMFA</div>
                <div class="detail">$([math]::Round(($S_EnabledNoMFA / [math]::Max($S_EnabledCount,1)) * 100, 1))% of enabled</div>
            </div>
        </div>
    </div>

    <div class="cards">
        <div class="card alert">
            <div class="label">No MFA &amp; No Target Group</div>
            <div class="value">$S_NoMfaNoGroup</div>
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
$S_TableRows
            </tbody>
        </table>
    </div>

    <div class="footer">
        CSV data: $(Split-Path $OutputPath -Leaf) | Report generated by ReportNonMFA.ps1
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
    # ── Disconnect ─────────────────────────────────────────────────────────────
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
