#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users

<#
.SYNOPSIS
    Reports authentication methods for Entra ID member accounts, classifying each as Modern or Legacy.

.DESCRIPTION
    Connects to Microsoft Graph and retrieves all member user accounts.
    For each user, enumerates registered authentication methods, classifies them as
    Modern Authentication, Legacy Authentication, No MFA, or No Methods, and exports
    results to CSV and HTML. Includes licence status and on-premises sync status.
    Adapted from Tony Redmond's Office365itpros scripts.

.PARAMETER OutputPath
    Path for the output CSV file. Defaults to a timestamped file in the current directory.
    The HTML report is written to the same path with a .html extension.

.EXAMPLE
    .\ReportLegacyAuthenticationMethods.ps1

.EXAMPLE
    .\ReportLegacyAuthenticationMethods.ps1 -OutputPath "C:\Reports\AuthMethods.csv"
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
    $S_OutputPath = Join-Path -Path (Get-Location).Path -ChildPath "ReportLegacyAuthenticationMethods_$S_Timestamp.csv"
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

Write-Host "Finding Azure AD accounts"
[array]$S_Users = Get-MgUser -Filter "userType eq 'Member'" -ConsistencyLevel eventual -CountVariable S_Records -All -Property Id, DisplayName, UserPrincipalName, AccountEnabled, AssignedLicenses, OnPremisesSyncEnabled
if (!($S_Users))
{
    Write-Host "No users found in Azure AD... exiting!"
    break
}

$S_Counter = 0
$S_Report = [System.Collections.Generic.List[Object]]::new()
foreach ($S_User in $S_Users)
{
    $S_Counter++
    Write-Host ("Processing user {0} {1}/{2}." -f $S_User.DisplayName, $S_Counter, $S_Users.Count)
    $S_AuthMethods = Get-MgUserAuthenticationMethod -UserId $S_User.Id
    Start-Sleep -Milliseconds $S_GraphRequestDelayMilliseconds

    $S_ModernTypes = @()
    $S_LegacyTypes = @()
    $S_NonMfaTypes = @()
    $S_ModernOdataTypes = @(
        "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod",
        "#microsoft.graph.fido2AuthenticationMethod",
        "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod"
    )

    foreach ($S_Method in $S_AuthMethods)
    {
        $S_Type = $S_Method.AdditionalProperties['@odata.type']
        switch ($S_Type)
        {
            "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" { $S_ModernTypes += "Microsoft Authenticator" }
            "#microsoft.graph.fido2AuthenticationMethod" { $S_ModernTypes += "Passkey" }
            "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" { $S_ModernTypes += "Passwordless" }
            "#microsoft.graph.phoneAuthenticationMethod" { $S_LegacyTypes += "Phone/SMS" }
            "#microsoft.graph.passwordAuthenticationMethod" { $S_NonMfaTypes += "Password" }
            "#microsoft.graph.emailAuthenticationMethod" { $S_NonMfaTypes += "Email" }
            "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" { $S_NonMfaTypes += "Windows Hello" }
            default { $S_LegacyTypes += $S_Type }
        }
    }

    if ($S_ModernTypes.Count -gt 0)
    {
        $S_DisplayMethod = "Modern Authentication"
        $S_P1 = ($S_ModernTypes + $S_LegacyTypes + $S_NonMfaTypes) -join ", "
    }
    elseif ($S_LegacyTypes.Count -gt 0)
    {
        $S_DisplayMethod = "Legacy Authentication"
        $S_P1 = ($S_LegacyTypes + $S_NonMfaTypes) -join ", "
    }
    else
    {
        if ($S_NonMfaTypes.Count -gt 0)
        {
            $S_DisplayMethod = "No MFA"
            $S_P1 = $S_NonMfaTypes -join ", "
        }
        else
        {
            $S_DisplayMethod = "No Methods"
            $S_P1 = ""
        }
    }

    # Determine license status
    $S_IsLicensed = if ($S_User.AssignedLicenses.Count -gt 0)
    {
        "Yes"
    }
    else
    {
        "No"
    }

    # Determine on-premises sync status
    $S_IsSynced = if ($S_User.OnPremisesSyncEnabled -eq $true)
    {
        "Yes"
    }
    elseif ($S_User.OnPremisesSyncEnabled -eq $false)
    {
        "No"
    }
    else
    {
        "Cloud-Only"
    }

    # Determine account enabled status
    $S_IsEnabled = if ($S_User.AccountEnabled -eq $true)
    {
        "Yes"
    }
    else
    {
        "No"
    }

    $S_ReportLine = [PSCustomObject]@{
        User             = $S_User.DisplayName
        UPN              = $S_User.UserPrincipalName
        Type             = $S_DisplayMethod
        Methods          = $S_P1
        Licensed         = $S_IsLicensed
        OnPremisesSynced = $S_IsSynced
        AccountEnabled   = $S_IsEnabled
        Id               = $S_User.Id
    }
    $S_Report.Add($S_ReportLine)
}

$S_Report = $S_Report | Sort-Object User 

# --- CSV Export ---
$S_Report | Export-Csv -Path $S_OutputPath -NoTypeInformation -Encoding UTF8

# --- HTML Report: Enabled Users MFA Status ---
$S_TotalMembers = $S_Report.Count
$S_TotalEnabled = ($S_Report | Where-Object { $_.AccountEnabled -eq "Yes" }).Count
$S_TotalDisabled = $S_TotalMembers - $S_TotalEnabled
$S_EnabledUsers = $S_Report | Where-Object { $_.AccountEnabled -eq "Yes" }
$S_EnabledLicensed = ($S_EnabledUsers | Where-Object { $_.Licensed -eq "Yes" }).Count

$S_MfaModern = ($S_EnabledUsers | Where-Object { $_.Type -eq "Modern Authentication" }).Count
$S_MfaLegacy = ($S_EnabledUsers | Where-Object { $_.Type -eq "Legacy Authentication" }).Count
$S_MfaNone = ($S_EnabledUsers | Where-Object { $_.Type -eq "No MFA" }).Count
$S_MfaNoMethod = ($S_EnabledUsers | Where-Object { $_.Type -eq "No Methods" }).Count

# Licensed enabled users with No MFA
$S_LicensedNoMfa = ($S_EnabledUsers | Where-Object { $_.Licensed -eq "Yes" -and $_.Type -eq "No MFA" }).Count
# On-prem synced enabled users with No MFA
$S_OnPremNoMfa = ($S_EnabledUsers | Where-Object { $_.OnPremisesSynced -eq "Yes" -and $_.Type -eq "No MFA" }).Count

$S_PieLabels = "'Modern Auth', 'Legacy Auth', 'No MFA', 'No Methods'"
$S_PieData = "{0}, {1}, {2}, {3}" -f $S_MfaModern, $S_MfaLegacy, $S_MfaNone, $S_MfaNoMethod
$S_PieColors = "'#27ae60', '#f39c12', '#e74c3c', '#95a5a6'"

$S_TableRows = ($S_EnabledUsers | Sort-Object User | ForEach-Object {
    $S_TypeClass = switch ($_.Type)
    {
        "Modern Authentication" { "modern" }
        "Legacy Authentication" { "legacy" }
        "No MFA" { "nomfa" }
        default { "nomethod" }
    }
    "<tr><td>{0}</td><td>{1}</td><td><span class='badge {2}'>{3}</span></td><td>{4}</td><td>{5}</td><td>{6}</td></tr>" -f
        [System.Net.WebUtility]::HtmlEncode($_.User),
        [System.Net.WebUtility]::HtmlEncode($_.UPN),
        $S_TypeClass,
        [System.Net.WebUtility]::HtmlEncode($_.Type),
        [System.Net.WebUtility]::HtmlEncode($_.Methods),
        [System.Net.WebUtility]::HtmlEncode($_.Licensed),
        [System.Net.WebUtility]::HtmlEncode($_.OnPremisesSynced)
}) -join "`n"

$S_TenantName = if ($S_ActiveContext.TenantId)
{
    $S_ActiveContext.TenantId
}
else
{
    "Unknown"
}
$S_ReportDate = Get-Date -Format "dd MMM yyyy HH:mm"

$S_Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Authentication Methods Report</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; padding: 30px; }
  .header { background: linear-gradient(135deg, #1a1a2e, #16213e); color: #fff; padding: 30px 40px; border-radius: 12px; margin-bottom: 30px; }
  .header h1 { font-size: 1.6em; margin-bottom: 6px; }
  .header p { font-size: 0.9em; opacity: 0.8; }
  .summary-cards { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .card { background: #fff; border-radius: 10px; padding: 24px 30px; flex: 1; min-width: 180px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
  .card .label { font-size: 0.85em; color: #777; text-transform: uppercase; letter-spacing: 0.5px; }
  .card .value { font-size: 2em; font-weight: 700; margin-top: 6px; }
  .info-cards { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .info-card { background: #fff; border-radius: 10px; padding: 20px 26px; flex: 1; min-width: 200px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); border-left: 5px solid #ccc; }
  .info-card .info-label { font-size: 0.82em; color: #555; margin-bottom: 6px; }
  .info-card .info-value { font-size: 1.8em; font-weight: 700; }
  .info-card.green  { border-left-color: #27ae60; } .info-card.green  .info-value { color: #27ae60; }
  .info-card.amber  { border-left-color: #f39c12; } .info-card.amber  .info-value { color: #f39c12; }
  .info-card.red    { border-left-color: #e74c3c; } .info-card.red    .info-value { color: #e74c3c; }
  .info-card.grey   { border-left-color: #7f8c8d; } .info-card.grey   .info-value { color: #7f8c8d; }
  .info-card.purple { border-left-color: #8e44ad; } .info-card.purple .info-value { color: #8e44ad; }
  .info-card.indigo { border-left-color: #2c3e50; } .info-card.indigo .info-value { color: #2c3e50; }
  .chart-section { background: #fff; border-radius: 10px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 30px; }
  .chart-section h2 { font-size: 1.1em; margin-bottom: 20px; color: #1a1a2e; }
  .chart-container { max-width: 580px; margin: 0 auto; }
  .table-section { background: #fff; border-radius: 10px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 30px; overflow-x: auto; }
  .table-section h2 { font-size: 1.1em; margin-bottom: 16px; color: #1a1a2e; }
  table { width: 100%; border-collapse: collapse; font-size: 0.88em; }
  th { background: #1a1a2e; color: #fff; padding: 12px 14px; text-align: left; position: sticky; top: 0; }
  td { padding: 10px 14px; border-bottom: 1px solid #eee; }
  tr:hover td { background: #f8f9fa; }
  .badge { padding: 4px 10px; border-radius: 12px; font-size: 0.82em; font-weight: 600; color: #fff; }
  .badge.modern   { background: #27ae60; }
  .badge.legacy   { background: #f39c12; }
  .badge.nomfa    { background: #e74c3c; }
  .badge.nomethod { background: #95a5a6; }
  .footer { text-align: center; font-size: 0.8em; color: #999; margin-top: 20px; }
</style>
</head>
<body>

<div class="header">
  <h1>Authentication Methods Report — Enabled Users</h1>
  <p>Tenant: $S_TenantName &nbsp;|&nbsp; Generated: $S_ReportDate &nbsp;|&nbsp; Focus: Enabled member accounts &amp; MFA status</p>
</div>

<div class="summary-cards">
  <div class="card"><div class="label">Total Members</div><div class="value" style="color:#2c3e50;">$S_TotalMembers</div></div>
  <div class="card"><div class="label">Enabled</div><div class="value" style="color:#3498db;">$S_TotalEnabled</div></div>
  <div class="card"><div class="label">Disabled</div><div class="value" style="color:#95a5a6;">$S_TotalDisabled</div></div>
  <div class="card"><div class="label">Enabled &amp; Licensed</div><div class="value" style="color:#27ae60;">$S_EnabledLicensed</div></div>
</div>

<div class="info-cards">
  <div class="info-card green"><div class="info-label">Modern Auth (Authenticator / Passkey)</div><div class="info-value">$S_MfaModern</div></div>
  <div class="info-card amber"><div class="info-label">Legacy Auth (Phone / SMS)</div><div class="info-value">$S_MfaLegacy</div></div>
  <div class="info-card red"><div class="info-label">No MFA (Password Only)</div><div class="info-value">$S_MfaNone</div></div>
  <div class="info-card grey"><div class="info-label">No Methods Registered</div><div class="info-value">$S_MfaNoMethod</div></div>
</div>

<div class="info-cards">
  <div class="info-card purple"><div class="info-label">Licensed &amp; No MFA</div><div class="info-value">$S_LicensedNoMfa</div></div>
  <div class="info-card indigo"><div class="info-label">On-Prem Synced &amp; No MFA</div><div class="info-value">$S_OnPremNoMfa</div></div>
</div>

<div class="chart-section">
  <h2>MFA Status — Enabled Users</h2>
  <div class="chart-container"><canvas id="pieChart"></canvas></div>
</div>

<div class="table-section">
  <h2>Enabled User Details ($S_TotalEnabled users)</h2>
  <table>
    <thead><tr><th>Display Name</th><th>UPN</th><th>MFA Status</th><th>Methods</th><th>Licensed</th><th>On-Prem Synced</th></tr></thead>
    <tbody>
$S_TableRows
    </tbody>
  </table>
</div>

<div class="footer">Report generated by ReportLegacyAuthenticationMethods.ps1</div>
<script>
new Chart(document.getElementById('pieChart'), {
  type: 'pie',
  data: {
    labels: [$S_PieLabels],
    datasets: [{
      data: [$S_PieData],
      backgroundColor: [$S_PieColors],
      borderWidth: 2, borderColor: '#fff'
    }]
  },
  options: {
    responsive: true,
    plugins: {
      legend: { position: 'right', labels: { padding: 24, font: { size: 15 }, boxWidth: 20 } },
      tooltip: {
        callbacks: {
          label: function(ctx) {
            var total = ctx.dataset.data.reduce(function(a,b){ return a+b; }, 0);
            var pct = total > 0 ? ((ctx.parsed / total) * 100).toFixed(1) : 0;
            return ctx.label + ': ' + ctx.parsed + ' (' + pct + '%)';
          }
        }
      }
    }
  }
});
</script>
</body>
</html>
"@

$S_HtmlPath = [System.IO.Path]::ChangeExtension($S_OutputPath, '.html')
$S_Html | Out-File -FilePath $S_HtmlPath -Encoding UTF8

# --- Console Summary ---
Write-Host ""
Write-Host "Authentication Methods Report" -ForegroundColor Cyan
Write-Host "--------------------------------------------"
Write-Host ("Total member users        : {0}" -f $S_TotalMembers)
Write-Host ("Enabled users             : {0}" -f $S_TotalEnabled)
Write-Host ("Disabled users            : {0}" -f $S_TotalDisabled)
Write-Host ("Modern Auth               : {0}" -f $S_MfaModern)
Write-Host ("Legacy Auth               : {0}" -f $S_MfaLegacy)
Write-Host ("No MFA (Password Only)    : {0}" -f $S_MfaNone)
Write-Host ("No Methods                : {0}" -f $S_MfaNoMethod)
Write-Host ("CSV exported to           : {0}" -f $S_OutputPath)
Write-Host ("HTML exported to          : {0}" -f $S_HtmlPath)

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
