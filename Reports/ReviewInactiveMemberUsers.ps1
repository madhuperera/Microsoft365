#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users

<#
.SYNOPSIS
    Reports on member users who have not signed in within a configurable number of days.

.DESCRIPTION
    Connects to Microsoft Graph and retrieves all enabled member user accounts.
    Evaluates sign-in activity and identifies members who have been inactive for longer
    than the specified threshold. Exports results to CSV and HTML.
    When Mode is set to Disable, optionally disables inactive accounts after reporting.

.PARAMETER ReportPath
    Folder or file path for the output report. If a folder is specified, a timestamped
    filename is generated automatically. Defaults to the current directory.

.PARAMETER InactiveDays
    Number of days of inactivity after which a member is considered inactive. Defaults to 90.

.PARAMETER Mode
    Operating mode. ReportOnly (default) only exports a report. Disable additionally
    disables inactive member accounts after reporting.

.PARAMETER SkipIfLastSignInIsNEVER
    When set to $true (default), skips disabling accounts that have never signed in.
    Only applies when Mode is Disable.

.EXAMPLE
    .\ReviewInactiveMemberUsers.ps1

.EXAMPLE
    .\ReviewInactiveMemberUsers.ps1 -InactiveDays 60 -Mode ReportOnly

.EXAMPLE
    .\ReviewInactiveMemberUsers.ps1 -InactiveDays 90 -Mode Disable
#>

[CmdletBinding()]
param (
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[string]$ReportPath,

	[Parameter(Mandatory = $false)]
	[ValidateRange(1, 3650)]
	[int]$InactiveDays = 90,

	[Parameter(Mandatory = $false)]
	[ValidateSet("ReportOnly", "Disable")]
	[string]$Mode = "ReportOnly",

	[Parameter(Mandatory = $false)]
	[bool]$SkipIfLastSignInIsNEVER = $true
)

$ErrorActionPreference = "Stop"

# Scopes required for ReportOnly mode. Disable mode additionally requires User.ReadWrite.All.
$S_RequiredGraphScopes = @(
	'User.Read.All'
	'AuditLog.Read.All'
)

$S_GraphRequestDelayMilliseconds = 5

try
{
	if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users))
	{
		throw "Microsoft.Graph.Users module is not installed. Install it using Install-Module Microsoft.Graph -Scope CurrentUser."
	}

	Import-Module Microsoft.Graph.Users -ErrorAction Stop

	$S_Context = Get-MgContext
	if (-not $S_Context)
	{
		$S_Scopes = if ($Mode -eq "Disable")
		{
			"User.ReadWrite.All", "AuditLog.Read.All"
		}
		else
		{
			$S_RequiredGraphScopes
		}

		Connect-MgGraph -Scopes $S_Scopes -ErrorAction Stop | Out-Null
	}

	$S_CutoffDate = (Get-Date).AddDays(-$InactiveDays)

	$S_Users = Get-MgUser -All `
		-Filter "userType eq 'Member' and accountEnabled eq true" `
		-Property "id,displayName,userPrincipalName,mail,userType,accountEnabled,createdDateTime,signInActivity,assignedLicenses,onPremisesSyncEnabled" `
		-ConsistencyLevel eventual

	$S_Report = foreach ($S_User in $S_Users)
	{
		$S_SignInActivity = $S_User.SignInActivity
		$S_LastInteractive = $S_SignInActivity.lastSignInDateTime
		$S_LastNonInteractive = $S_SignInActivity.lastNonInteractiveSignInDateTime

		$S_LastInteractiveDt =
			if ($S_LastInteractive)
			{
				[datetime]$S_LastInteractive
			}
			else
			{
				$null
			}

		$S_LastNonInteractiveDt =
			if ($S_LastNonInteractive)
			{
				[datetime]$S_LastNonInteractive
			}
			else
			{
				$null
			}

		$S_MostRecent = @($S_LastInteractiveDt, $S_LastNonInteractiveDt) | Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1
		$S_LastSignInAgo =
			if ($S_MostRecent)
			{
				[int]((Get-Date) - $S_MostRecent).TotalDays
			}
			else
			{
				"Never"
			}

		$S_IsInactive = $false
		if (-not $S_LastInteractiveDt -and -not $S_LastNonInteractiveDt)
		{
			$S_IsInactive = $true
		}
		elseif ((-not $S_LastInteractiveDt -or $S_LastInteractiveDt -lt $S_CutoffDate) -and (-not $S_LastNonInteractiveDt -or $S_LastNonInteractiveDt -lt $S_CutoffDate))
		{
			$S_IsInactive = $true
		}

		[pscustomobject]@{
			DisplayName = $S_User.DisplayName
			UserPrincipalName = $S_User.UserPrincipalName
			Mail = $S_User.Mail
			PrimaryDomain =
				if ($S_User.Mail -and $S_User.Mail -match "@")
				{
					($S_User.Mail -split "@", 2)[1]
				}
				else
				{
					$null
				}
			UserType = $S_User.UserType
			AccountEnabled = $S_User.AccountEnabled
			LicenseAssigned = ($S_User.AssignedLicenses.Count -gt 0)
			OnPremisesSyncEnabled = [bool]$S_User.OnPremisesSyncEnabled
			CreatedDateTime = $S_User.CreatedDateTime
			LastInteractiveSignIn = $S_LastInteractiveDt
			LastNonInteractiveSignIn = $S_LastNonInteractiveDt
			LastSignInAgoDays = $S_LastSignInAgo
			Inactive = $S_IsInactive
		}
	}

	$S_InactiveUsers = $S_Report | Where-Object { $_.Inactive }
	$S_TotalMembers = $S_Report.Count
	$S_TotalInactive = $S_InactiveUsers.Count
	$S_PercentInactive =
		if ($S_TotalMembers -gt 0)
		{
			[math]::Round(($S_TotalInactive / $S_TotalMembers) * 100, 2)
		}
		else
		{
			0
		}

	if (-not $ReportPath)
	{
		$ReportPath = (Get-Location).Path
	}

	$S_ReportFolder =
		if (Test-Path $ReportPath -PathType Container)
		{
			$ReportPath
		}
		else
		{
			Split-Path -Parent $ReportPath
		}

	if ($S_ReportFolder -and -not (Test-Path $S_ReportFolder))
	{
		New-Item -ItemType Directory -Path $S_ReportFolder -Force | Out-Null
	}

	$S_Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
	$S_ReportFile =
		if (Test-Path $ReportPath -PathType Container)
		{
			Join-Path $ReportPath ("ReviewInactiveMemberUsers_{0}.csv" -f $S_Timestamp)
		}
		else
		{
			$ReportPath
		}

	$S_InactiveUsers | Sort-Object DisplayName | Export-Csv -Path $S_ReportFile -NoTypeInformation -Encoding UTF8

	# --- HTML Report with Pie Chart: Inactive Member Users by Last Sign-In Age ---
	$S_Bucket1Max = $InactiveDays * 2
	$S_Bucket2Max = 365

	$S_BucketLabel1 = "{0}-{1} Days" -f $InactiveDays, $S_Bucket1Max
	$S_BucketLabel2 = "{0}-{1} Days" -f $S_Bucket1Max, $S_Bucket2Max
	$S_BucketLabel3 = "Over 1 Year"
	$S_BucketLabel4 = "Never Signed In"

	$S_Bucket1Count = 0
	$S_Bucket2Count = 0
	$S_Bucket3Count = 0
	$S_Bucket4Count = 0

	foreach ($S_User in $S_InactiveUsers)
	{
		if ($S_User.LastSignInAgoDays -eq "Never")
		{
			$S_Bucket4Count++
		}
		elseif ([int]$S_User.LastSignInAgoDays -gt $S_Bucket2Max)
		{
			$S_Bucket3Count++
		}
		elseif ([int]$S_User.LastSignInAgoDays -gt $S_Bucket1Max)
		{
			$S_Bucket2Count++
		}
		else
		{
			$S_Bucket1Count++
		}
	}

	# Collapse bucket2 into bucket3 if InactiveDays*2 >= 365
	if ($S_Bucket1Max -ge $S_Bucket2Max)
	{
		$S_BucketLabelsJson = "'{0}', '{1}', '{2}'" -f $S_BucketLabel1, $S_BucketLabel3, $S_BucketLabel4
		$S_BucketDataJson = "{0}, {1}, {2}" -f $S_Bucket1Count, ($S_Bucket2Count + $S_Bucket3Count), $S_Bucket4Count
		$S_BucketColorsJson = "'#3498db', '#e74c3c', '#95a5a6'"
	}
	else
	{
		$S_BucketLabelsJson = "'{0}', '{1}', '{2}', '{3}'" -f $S_BucketLabel1, $S_BucketLabel2, $S_BucketLabel3, $S_BucketLabel4
		$S_BucketDataJson = "{0}, {1}, {2}, {3}" -f $S_Bucket1Count, $S_Bucket2Count, $S_Bucket3Count, $S_Bucket4Count
		$S_BucketColorsJson = "'#3498db', '#f39c12', '#e74c3c', '#95a5a6'"
	}

	$S_TableRows = ($S_InactiveUsers | Sort-Object DisplayName | ForEach-Object {
		$S_SignInAge =
			if ($_.LastSignInAgoDays -eq "Never")
			{
				"Never"
			}
			else
			{
				"{0} days" -f $_.LastSignInAgoDays
			}

		$S_Licensed =
			if ($_.LicenseAssigned)
			{
				"Yes"
			}
			else
			{
				"No"
			}

		$S_SyncEnabled =
			if ($_.OnPremisesSyncEnabled)
			{
				"Yes"
			}
			else
			{
				"No"
			}

		"<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td></tr>" -f
			[System.Net.WebUtility]::HtmlEncode($_.DisplayName),
			[System.Net.WebUtility]::HtmlEncode($_.UserPrincipalName),
			[System.Net.WebUtility]::HtmlEncode($_.Mail),
			[System.Net.WebUtility]::HtmlEncode($S_Licensed),
			[System.Net.WebUtility]::HtmlEncode($S_SyncEnabled),
			[System.Net.WebUtility]::HtmlEncode($S_SignInAge),
			[System.Net.WebUtility]::HtmlEncode($_.CreatedDateTime)
	}) -join "`n"

	# --- Additional stats for info cards ---
	$S_InactiveLicensed = ($S_InactiveUsers | Where-Object { $_.LicenseAssigned }).Count
	$S_InactiveUnlicensed = $S_TotalInactive - $S_InactiveLicensed
	$S_InactiveOnPrem = ($S_InactiveUsers | Where-Object { $_.OnPremisesSyncEnabled }).Count
	$S_InactiveCloudOnly = $S_TotalInactive - $S_InactiveOnPrem
	$S_InactiveNeverSignIn = ($S_InactiveUsers | Where-Object { $_.LastSignInAgoDays -eq "Never" }).Count

	$S_TenantName =
		if ($S_Context.TenantId)
		{
			$S_Context.TenantId
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
<title>Inactive Member Users Report</title>
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
  .card .value.red { color: #e74c3c; }
  .card .value.blue { color: #3498db; }
  .card .value.orange { color: #f39c12; }
  .info-cards { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .info-card { background: #fff; border-radius: 10px; padding: 20px 26px; flex: 1; min-width: 200px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); border-left: 5px solid #ccc; }
  .info-card .info-label { font-size: 0.82em; color: #555; margin-bottom: 6px; }
  .info-card .info-value { font-size: 1.8em; font-weight: 700; }
  .info-card.purple { border-left-color: #8e44ad; }
  .info-card.purple .info-value { color: #8e44ad; }
  .info-card.teal { border-left-color: #16a085; }
  .info-card.teal .info-value { color: #16a085; }
  .info-card.indigo { border-left-color: #2c3e50; }
  .info-card.indigo .info-value { color: #2c3e50; }
  .info-card.coral { border-left-color: #e67e22; }
  .info-card.coral .info-value { color: #e67e22; }
  .info-card.grey { border-left-color: #7f8c8d; }
  .info-card.grey .info-value { color: #7f8c8d; }
  .chart-section { background: #fff; border-radius: 10px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 30px; }
  .chart-section h2 { font-size: 1.1em; margin-bottom: 20px; color: #1a1a2e; }
  .chart-container { max-width: 580px; margin: 0 auto; }
  .table-section { background: #fff; border-radius: 10px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 30px; overflow-x: auto; }
  .table-section h2 { font-size: 1.1em; margin-bottom: 16px; color: #1a1a2e; }
  table { width: 100%; border-collapse: collapse; font-size: 0.88em; }
  th { background: #1a1a2e; color: #fff; padding: 12px 14px; text-align: left; position: sticky; top: 0; }
  td { padding: 10px 14px; border-bottom: 1px solid #eee; }
  tr:hover td { background: #f8f9fa; }
  .footer { text-align: center; font-size: 0.8em; color: #999; margin-top: 20px; }
</style>
</head>
<body>

<div class="header">
  <h1>Inactive Member Users Report</h1>
  <p>Tenant: $S_TenantName &nbsp;|&nbsp; Generated: $S_ReportDate &nbsp;|&nbsp; Inactive Threshold: $InactiveDays days</p>
</div>

<div class="summary-cards">
  <div class="card"><div class="label">Total Enabled Members</div><div class="value blue">$S_TotalMembers</div></div>
  <div class="card"><div class="label">Inactive Members</div><div class="value red">$S_TotalInactive</div></div>
  <div class="card"><div class="label">Active Members</div><div class="value" style="color:#27ae60;">$($S_TotalMembers - $S_TotalInactive)</div></div>
  <div class="card"><div class="label">Inactive %</div><div class="value orange">$S_PercentInactive%</div></div>
</div>

<div class="info-cards">
  <div class="info-card purple"><div class="info-label">Inactive &amp; Licensed</div><div class="info-value">$S_InactiveLicensed</div></div>
  <div class="info-card coral"><div class="info-label">Inactive &amp; Unlicensed</div><div class="info-value">$S_InactiveUnlicensed</div></div>
  <div class="info-card indigo"><div class="info-label">Inactive &amp; On-Prem Synced</div><div class="info-value">$S_InactiveOnPrem</div></div>
  <div class="info-card teal"><div class="info-label">Inactive &amp; Cloud-Only</div><div class="info-value">$S_InactiveCloudOnly</div></div>
  <div class="info-card grey"><div class="info-label">Never Signed In</div><div class="info-value">$S_InactiveNeverSignIn</div></div>
</div>

<div class="chart-section">
  <h2>Inactive Members by Last Sign-In Age</h2>
  <div class="chart-container"><canvas id="pieChart"></canvas></div>
</div>

<div class="table-section">
  <h2>Inactive Member User Details ($S_TotalInactive users)</h2>
  <table>
    <thead><tr><th>Display Name</th><th>UPN</th><th>Mail</th><th>Licensed</th><th>On-Prem Sync</th><th>Last Sign-In</th><th>Created</th></tr></thead>
    <tbody>
$S_TableRows
    </tbody>
  </table>
</div>

<div class="footer">Report generated by ReviewInactiveMemberUsers.ps1</div>

<script>
new Chart(document.getElementById('pieChart'), {
  type: 'pie',
  data: {
    labels: [$S_BucketLabelsJson],
    datasets: [{
      data: [$S_BucketDataJson],
      backgroundColor: [$S_BucketColorsJson],
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

	$S_HtmlReportFile = Join-Path $S_ReportFolder ("ReviewInactiveMemberUsers_{0}.html" -f $S_Timestamp)
	$S_Html | Out-File -FilePath $S_HtmlReportFile -Encoding UTF8

	$S_DisableReport = @()
	if ($Mode -eq "Disable" -and $S_TotalInactive -gt 0)
	{
		$S_DisableCandidates =
			if ($SkipIfLastSignInIsNEVER)
			{
				$S_InactiveUsers | Where-Object { $_.LastSignInAgoDays -ne "Never" }
			}
			else
			{
				$S_InactiveUsers
			}

		$S_TotalToDisable = $S_DisableCandidates.Count
		if ($SkipIfLastSignInIsNEVER)
		{
			Write-Host ("Skip users with never sign-in: {0}" -f ($S_InactiveUsers.Count - $S_TotalToDisable))
		}

		$S_Counter = 0
		foreach ($S_User in $S_DisableCandidates)
		{
			$S_Counter++
			Write-Host ("Processing {0}/{1}: {2}" -f $S_Counter, $S_TotalToDisable, $S_User.UserPrincipalName) -ForegroundColor Yellow
			for ($S_Countdown = 3; $S_Countdown -ge 1; $S_Countdown--)
			{
				Write-Host ("Disabling in {0}..." -f $S_Countdown)
				Start-Sleep -Seconds 1
			}

			$S_DisableStatus = "Success"
			$S_DisableError = $null
			try
			{
				Update-MgUser -UserId $S_User.UserPrincipalName -AccountEnabled:$false -ErrorAction Stop
			}
			catch
			{
				$S_DisableStatus = "Failed"
				$S_DisableError = $_.Exception.Message
			}

			$S_DisableReport += [pscustomobject]@{
				DisplayName       = $S_User.DisplayName
				UserPrincipalName = $S_User.UserPrincipalName
				Disabled          = $S_DisableStatus
				Error             = $S_DisableError
			}
		}
	}

	$S_DisableReportFile = $null
	if ($S_DisableReport.Count -gt 0)
	{
		$S_DisableReportFile = Join-Path $S_ReportFolder ("DisabledMemberUsers_{0}.csv" -f $S_Timestamp)
		$S_DisableReport | Export-Csv -Path $S_DisableReportFile -NoTypeInformation -Encoding UTF8
	}

	Write-Host "Inactive Member Users Report" -ForegroundColor Cyan
	Write-Host "--------------------------------------------"
	Write-Host ("Total enabled member users : {0}" -f $S_TotalMembers)
	Write-Host ("Inactive member users      : {0}" -f $S_TotalInactive)
	Write-Host ("Inactive percentage       : {0}%" -f $S_PercentInactive)
	Write-Host ("Inactive days threshold   : {0}" -f $InactiveDays)
	Write-Host ("CSV report exported to    : {0}" -f $S_ReportFile)
	Write-Host ("HTML report exported to   : {0}" -f $S_HtmlReportFile)
	Write-Host ("Mode                     : {0}" -f $Mode)
	if ($S_DisableReportFile)
	{
		Write-Host ("Disable report exported  : {0}" -f $S_DisableReportFile)
	}

	$S_DisconnectChoice = Read-Host "Disconnect from Microsoft Graph? (Y/N)"
	if ($S_DisconnectChoice -match '^(y|yes)$')
	{
		Disconnect-MgGraph -ErrorAction SilentlyContinue
	}
}
catch
{
	Write-Error $_
	exit 1
}
