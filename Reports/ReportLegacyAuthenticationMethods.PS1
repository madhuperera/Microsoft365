<#
.SYNOPSIS
  Generates a report of authentication methods used by Azure AD (Microsoft 365) accounts.

.DESCRIPTION
  This script connects to Microsoft Graph using the required scopes and retrieves all Azure AD user accounts.
  For each user, it enumerates their registered authentication methods, classifies them as Modern or Legacy, and compiles a report.
  The report includes license status, on-premises sync status, and account enabled status.
  The report is displayed in an Out-GridView window and exported as a CSV file.

.PARAMETER None
  The script does not accept parameters. It operates on all Azure AD users.

.NOTES
  Author: Adapted from Tony Redmond's Office365itpros scripts.
  Source: https://github.com/12Knocksinna/Office365itpros/blob/master/ReportAuthenticationMethods.PS1
  More Info: https://office365itpros.com

.REQUIREMENTS
  - Microsoft Graph PowerShell SDK (Connect-MgGraph, Get-MgUser, Get-MgUserAuthenticationMethod)
  - Appropriate permissions: UserAuthenticationMethod.Read.All, Directory.Read.All, User.Read.All

.OUTPUTS
  - Displays a summary of authentication methods in the console.
  - Shows detailed results in an Out-GridView window.
  - Exports results to 'ReportAuthenticationMethods.csv' in the current directory.

.EXAMPLE
  .\ReportLegacyAuthenticationMethods.PS1

  Connects to Microsoft Graph, processes all users, and exports their authentication methods with license and sync status.

.DISCLAIMER
  Do not use this script in production until you have validated it in a non-production environment.
  Review and adapt the code to meet your organization's requirements.
#>
# ReportAuthenticationMethods.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportAuthenticationMethods.PS1
# A report of the authentication methods for Azure AD licensed accounts

Connect-MgGraph -Scopes UserAuthenticationMethod.Read.All, Directory.Read.All, User.Read.All
# Select-MgProfile Beta

Write-Host "Finding Azure AD accounts"
[array]$Users = Get-MgUser -Filter "userType eq 'Member'" -ConsistencyLevel eventual -CountVariable Records -All -Property Id, DisplayName, UserPrincipalName, AccountEnabled, AssignedLicenses, OnPremisesSyncEnabled
If (!($Users)) { Write-Host "No users found in Azure AD... exiting!"; break }

$i = 0
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $Users)
{
 $i++
 Write-Host ("Processing user {0} {1}/{2}." -f $User.DisplayName, $i, $Users.Count)
$AuthMethods = Get-MgUserAuthenticationMethod -UserId $User.Id

$ModernTypes = @()
$LegacyTypes = @()
$NonMfaTypes = @()
$ModernOdataTypes = @(
  "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod",
  "#microsoft.graph.fido2AuthenticationMethod",
  "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod"
)

foreach ($Method in $AuthMethods) {
  $Type = $Method.AdditionalProperties['@odata.type']
  switch ($Type) {
    "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" { $ModernTypes += "Microsoft Authenticator" }
    "#microsoft.graph.fido2AuthenticationMethod" { $ModernTypes += "Passkey" }
    "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" { $ModernTypes += "Passwordless" }
    "#microsoft.graph.phoneAuthenticationMethod" { $LegacyTypes += "Phone/SMS" }
    "#microsoft.graph.passwordAuthenticationMethod" { $NonMfaTypes += "Password" }
    "#microsoft.graph.emailAuthenticationMethod" { $NonMfaTypes += "Email" }
    "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" { $NonMfaTypes += "Windows Hello" }
    default { $LegacyTypes += $Type }
  }
}

if ($ModernTypes.Count -gt 0) 
{
  $DisplayMethod = "Modern Authentication"
  $P1 = ($ModernTypes + $LegacyTypes + $NonMfaTypes) -join ", "
} 
elseif ($LegacyTypes.Count -gt 0) 
{
  $DisplayMethod = "Legacy Authentication"
  $P1 = ($LegacyTypes + $NonMfaTypes) -join ", "
} 
else 
{
  if ($NonMfaTypes.Count -gt 0) {
    $DisplayMethod = "No MFA"
    $P1 = $NonMfaTypes -join ", "
  } else {
    $DisplayMethod = "No Methods"
    $P1 = ""
  }
}

# Determine license status
$IsLicensed = if ($User.AssignedLicenses.Count -gt 0) { "Yes" } else { "No" }

# Determine on-premises sync status
$IsSynced = if ($User.OnPremisesSyncEnabled -eq $true) { "Yes" } elseif ($User.OnPremisesSyncEnabled -eq $false) { "No" } else { "Cloud-Only" }

# Determine account enabled status
$IsEnabled = if ($User.AccountEnabled -eq $true) { "Yes" } else { "No" }

$ReportLine = [PSCustomObject]@{
  User   = $User.DisplayName
  UPN    = $User.UserPrincipalName
  Type   = $DisplayMethod
  Methods= $P1
  Licensed = $IsLicensed
  OnPremisesSynced = $IsSynced
  AccountEnabled = $IsEnabled
  Id     = $User.Id
}
$Report.Add($ReportLine)

} #End ForEach User
 
   
$Report = $Report | Sort-Object User 

# --- CSV Export ---
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvFile = ".\ReportAuthenticationMethods_{0}.csv" -f $timestamp
$Report | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8

# --- HTML Report: Enabled Users MFA Status ---
$totalMembers     = $Report.Count
$totalEnabled     = ($Report | Where-Object { $_.AccountEnabled -eq "Yes" }).Count
$totalDisabled    = $totalMembers - $totalEnabled
$enabledUsers     = $Report | Where-Object { $_.AccountEnabled -eq "Yes" }
$enabledLicensed  = ($enabledUsers | Where-Object { $_.Licensed -eq "Yes" }).Count

$mfaModern   = ($enabledUsers | Where-Object { $_.Type -eq "Modern Authentication" }).Count
$mfaLegacy   = ($enabledUsers | Where-Object { $_.Type -eq "Legacy Authentication" }).Count
$mfaNone     = ($enabledUsers | Where-Object { $_.Type -eq "No MFA" }).Count
$mfaNoMethod = ($enabledUsers | Where-Object { $_.Type -eq "No Methods" }).Count

# Licensed enabled users with No MFA
$licensedNoMfa = ($enabledUsers | Where-Object { $_.Licensed -eq "Yes" -and $_.Type -eq "No MFA" }).Count
# On-prem synced enabled users with No MFA
$onPremNoMfa   = ($enabledUsers | Where-Object { $_.OnPremisesSynced -eq "Yes" -and $_.Type -eq "No MFA" }).Count

$pieLabels = "'Modern Auth', 'Legacy Auth', 'No MFA', 'No Methods'"
$pieData   = "{0}, {1}, {2}, {3}" -f $mfaModern, $mfaLegacy, $mfaNone, $mfaNoMethod
$pieColors = "'#27ae60', '#f39c12', '#e74c3c', '#95a5a6'"

$tableRows = ($enabledUsers | Sort-Object User | ForEach-Object {
	$typeClass = switch ($_.Type) {
		"Modern Authentication" { "modern" }
		"Legacy Authentication" { "legacy" }
		"No MFA"               { "nomfa" }
		default                { "nomethod" }
	}
	"<tr><td>{0}</td><td>{1}</td><td><span class='badge {2}'>{3}</span></td><td>{4}</td><td>{5}</td><td>{6}</td></tr>" -f
		[System.Net.WebUtility]::HtmlEncode($_.User),
		[System.Net.WebUtility]::HtmlEncode($_.UPN),
		$typeClass,
		[System.Net.WebUtility]::HtmlEncode($_.Type),
		[System.Net.WebUtility]::HtmlEncode($_.Methods),
		[System.Net.WebUtility]::HtmlEncode($_.Licensed),
		[System.Net.WebUtility]::HtmlEncode($_.OnPremisesSynced)
}) -join "`n"

$context = Get-MgContext
$tenantName = if ($context.TenantId) { $context.TenantId } else { "Unknown" }
$reportDate = Get-Date -Format "dd MMM yyyy HH:mm"

$html = @"
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
  <p>Tenant: $tenantName &nbsp;|&nbsp; Generated: $reportDate &nbsp;|&nbsp; Focus: Enabled member accounts &amp; MFA status</p>
</div>

<div class="summary-cards">
  <div class="card"><div class="label">Total Members</div><div class="value" style="color:#2c3e50;">$totalMembers</div></div>
  <div class="card"><div class="label">Enabled</div><div class="value" style="color:#3498db;">$totalEnabled</div></div>
  <div class="card"><div class="label">Disabled</div><div class="value" style="color:#95a5a6;">$totalDisabled</div></div>
  <div class="card"><div class="label">Enabled &amp; Licensed</div><div class="value" style="color:#27ae60;">$enabledLicensed</div></div>
</div>

<div class="info-cards">
  <div class="info-card green"><div class="info-label">Modern Auth (Authenticator / Passkey)</div><div class="info-value">$mfaModern</div></div>
  <div class="info-card amber"><div class="info-label">Legacy Auth (Phone / SMS)</div><div class="info-value">$mfaLegacy</div></div>
  <div class="info-card red"><div class="info-label">No MFA (Password Only)</div><div class="info-value">$mfaNone</div></div>
  <div class="info-card grey"><div class="info-label">No Methods Registered</div><div class="info-value">$mfaNoMethod</div></div>
</div>

<div class="info-cards">
  <div class="info-card purple"><div class="info-label">Licensed &amp; No MFA</div><div class="info-value">$licensedNoMfa</div></div>
  <div class="info-card indigo"><div class="info-label">On-Prem Synced &amp; No MFA</div><div class="info-value">$onPremNoMfa</div></div>
</div>

<div class="chart-section">
  <h2>MFA Status — Enabled Users</h2>
  <div class="chart-container"><canvas id="pieChart"></canvas></div>
</div>

<div class="table-section">
  <h2>Enabled User Details ($totalEnabled users)</h2>
  <table>
    <thead><tr><th>Display Name</th><th>UPN</th><th>MFA Status</th><th>Methods</th><th>Licensed</th><th>On-Prem Synced</th></tr></thead>
    <tbody>
$tableRows
    </tbody>
  </table>
</div>

<div class="footer">Report generated by ReportLegacyAuthenticationMethods.ps1</div>

<script>
new Chart(document.getElementById('pieChart'), {
  type: 'pie',
  data: {
    labels: [$pieLabels],
    datasets: [{
      data: [$pieData],
      backgroundColor: [$pieColors],
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

$htmlFile = ".\ReportAuthenticationMethods_{0}.html" -f $timestamp
$html | Out-File -FilePath $htmlFile -Encoding UTF8

# --- Console Summary ---
Write-Host ""
Write-Host "Authentication Methods Report" -ForegroundColor Cyan
Write-Host "--------------------------------------------"
Write-Host ("Total member users        : {0}" -f $totalMembers)
Write-Host ("Enabled users             : {0}" -f $totalEnabled)
Write-Host ("Disabled users            : {0}" -f $totalDisabled)
Write-Host ("Modern Auth               : {0}" -f $mfaModern)
Write-Host ("Legacy Auth               : {0}" -f $mfaLegacy)
Write-Host ("No MFA (Password Only)    : {0}" -f $mfaNone)
Write-Host ("No Methods                : {0}" -f $mfaNoMethod)
Write-Host ("CSV exported to           : {0}" -f $csvFile)
Write-Host ("HTML exported to          : {0}" -f $htmlFile)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
