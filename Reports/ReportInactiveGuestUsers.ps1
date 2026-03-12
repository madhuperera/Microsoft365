param(
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

try {
	if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
		throw "Microsoft.Graph.Users module is not installed. Install it using Install-Module Microsoft.Graph -Scope CurrentUser."
	}

	Import-Module Microsoft.Graph.Users -ErrorAction Stop

	$context = Get-MgContext
	if (-not $context) {
		$scopes = if ($Mode -eq "Disable") { "User.ReadWrite.All", "AuditLog.Read.All" } else { "User.Read.All", "AuditLog.Read.All" }
		Connect-MgGraph -Scopes $scopes -ErrorAction Stop | Out-Null
	}

	$cutoffDate = (Get-Date).AddDays(-$InactiveDays)

	$users = Get-MgUser -All `
		-Filter "userType eq 'Guest' and accountEnabled eq true" `
		-Property "id,displayName,userPrincipalName,mail,userType,accountEnabled,createdDateTime,signInActivity,externalUserState,externalUserStateChangeDateTime" `
		-ConsistencyLevel eventual

	$report = foreach ($user in $users) {
		$signInActivity = $user.SignInActivity
		$lastInteractive = $signInActivity.lastSignInDateTime
		$lastNonInteractive = $signInActivity.lastNonInteractiveSignInDateTime

		$lastInteractiveDt = if ($lastInteractive) { [datetime]$lastInteractive } else { $null }
		$lastNonInteractiveDt = if ($lastNonInteractive) { [datetime]$lastNonInteractive } else { $null }

		$mostRecent = @($lastInteractiveDt, $lastNonInteractiveDt) | Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1
		$lastSignInAgo = if ($mostRecent) { [int]((Get-Date) - $mostRecent).TotalDays } else { "Never" }

		$isInactive = $false
		if (-not $lastInteractiveDt -and -not $lastNonInteractiveDt) {
			$isInactive = $true
		} elseif ((-not $lastInteractiveDt -or $lastInteractiveDt -lt $cutoffDate) -and (-not $lastNonInteractiveDt -or $lastNonInteractiveDt -lt $cutoffDate)) {
			$isInactive = $true
		}

		[pscustomobject]@{
			DisplayName                 = $user.DisplayName
			UserPrincipalName           = $user.UserPrincipalName
			Mail                        = $user.Mail
			UserType                    = $user.UserType
			AccountEnabled              = $user.AccountEnabled
			CreatedDateTime             = $user.CreatedDateTime
			InvitationStatus            = $user.ExternalUserState
			InvitationStatusChangeDate  = $user.ExternalUserStateChangeDateTime
			LastInteractiveSignIn       = $lastInteractiveDt
			LastNonInteractiveSignIn    = $lastNonInteractiveDt
			LastSignInAgoDays           = $lastSignInAgo
			Inactive                    = $isInactive
		}
	}

	$inactiveUsers = $report | Where-Object { $_.Inactive }
	$totalGuests = $report.Count
	$totalInactive = $inactiveUsers.Count
	$percentInactive = if ($totalGuests -gt 0) { [math]::Round(($totalInactive / $totalGuests) * 100, 2) } else { 0 }

	if (-not $ReportPath) {
		$ReportPath = (Get-Location).Path
	}

	$reportFolder = if (Test-Path $ReportPath -PathType Container) { $ReportPath } else { Split-Path -Parent $ReportPath }
	if ($reportFolder -and -not (Test-Path $reportFolder)) {
		New-Item -ItemType Directory -Path $reportFolder -Force | Out-Null
	}

	$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
	$reportFile = if (Test-Path $ReportPath -PathType Container) {
		Join-Path $ReportPath ("InactiveGuestUsers_{0}.csv" -f $timestamp)
	} else {
		$ReportPath
	}

	$inactiveUsers | Sort-Object DisplayName | Export-Csv -Path $reportFile -NoTypeInformation -Encoding UTF8

	# --- HTML Report with Pie Chart: Inactive Guest Users by Last Sign-In Age ---
	$bucket1Max = $InactiveDays * 2
	$bucket2Max = 365

	$bucketLabel1 = "{0}-{1} Days" -f $InactiveDays, $bucket1Max
	$bucketLabel2 = "{0}-{1} Days" -f $bucket1Max, $bucket2Max
	$bucketLabel3 = "Over 1 Year"
	$bucketLabel4 = "Never Signed In"

	$bucket1Count = 0; $bucket2Count = 0; $bucket3Count = 0; $bucket4Count = 0

	foreach ($u in $inactiveUsers) {
		if ($u.LastSignInAgoDays -eq "Never") {
			$bucket4Count++
		} elseif ([int]$u.LastSignInAgoDays -gt $bucket2Max) {
			$bucket3Count++
		} elseif ([int]$u.LastSignInAgoDays -gt $bucket1Max) {
			$bucket2Count++
		} else {
			$bucket1Count++
		}
	}

	# Collapse bucket2 into bucket3 if InactiveDays*2 >= 365
	if ($bucket1Max -ge $bucket2Max) {
		$bucketLabelsJson = "'{0}', '{1}', '{2}'" -f $bucketLabel1, $bucketLabel3, $bucketLabel4
		$bucketDataJson   = "{0}, {1}, {2}" -f $bucket1Count, ($bucket2Count + $bucket3Count), $bucket4Count
		$bucketColorsJson = "'#3498db', '#e74c3c', '#95a5a6'"
	} else {
		$bucketLabelsJson = "'{0}', '{1}', '{2}', '{3}'" -f $bucketLabel1, $bucketLabel2, $bucketLabel3, $bucketLabel4
		$bucketDataJson   = "{0}, {1}, {2}, {3}" -f $bucket1Count, $bucket2Count, $bucket3Count, $bucket4Count
		$bucketColorsJson = "'#3498db', '#f39c12', '#e74c3c', '#95a5a6'"
	}

	$tableRows = ($inactiveUsers | Sort-Object DisplayName | ForEach-Object {
		$signInAge = if ($_.LastSignInAgoDays -eq "Never") { "Never" } else { "{0} days" -f $_.LastSignInAgoDays }
		$invStatus = if ($_.InvitationStatus) { $_.InvitationStatus } else { "-" }
		"<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td></tr>" -f
			[System.Net.WebUtility]::HtmlEncode($_.DisplayName),
			[System.Net.WebUtility]::HtmlEncode($_.UserPrincipalName),
			[System.Net.WebUtility]::HtmlEncode($_.Mail),
			[System.Net.WebUtility]::HtmlEncode($invStatus),
			[System.Net.WebUtility]::HtmlEncode($signInAge)
	}) -join "`n"

	$tenantName = if ($context.TenantId) { $context.TenantId } else { "Unknown" }
	$reportDate = Get-Date -Format "dd MMM yyyy HH:mm"

	$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Inactive Guest Users Report</title>
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
  <h1>Inactive Guest Users Report</h1>
  <p>Tenant: $tenantName &nbsp;|&nbsp; Generated: $reportDate &nbsp;|&nbsp; Inactive Threshold: $InactiveDays days</p>
</div>

<div class="summary-cards">
  <div class="card"><div class="label">Total Enabled Guests</div><div class="value blue">$totalGuests</div></div>
  <div class="card"><div class="label">Inactive Guests</div><div class="value red">$totalInactive</div></div>
  <div class="card"><div class="label">Active Guests</div><div class="value" style="color:#27ae60;">$($totalGuests - $totalInactive)</div></div>
  <div class="card"><div class="label">Inactive %</div><div class="value orange">$percentInactive%</div></div>
</div>

<div class="chart-section">
  <h2>Inactive Guests by Last Sign-In Age</h2>
  <div class="chart-container"><canvas id="pieChart"></canvas></div>
</div>

<div class="table-section">
  <h2>Inactive Guest User Details ($totalInactive users)</h2>
  <table>
    <thead><tr><th>Display Name</th><th>UPN</th><th>Mail</th><th>Invitation Status</th><th>Last Sign-In</th></tr></thead>
    <tbody>
$tableRows
    </tbody>
  </table>
</div>

<div class="footer">Report generated by ReportInactiveGuestUsers.ps1</div>

<script>
new Chart(document.getElementById('pieChart'), {
  type: 'pie',
  data: {
    labels: [$bucketLabelsJson],
    datasets: [{
      data: [$bucketDataJson],
      backgroundColor: [$bucketColorsJson],
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

	$htmlReportFile = Join-Path $reportFolder ("InactiveGuestUsers_{0}.html" -f $timestamp)
	$html | Out-File -FilePath $htmlReportFile -Encoding UTF8

	$disableReport = @()
	if ($Mode -eq "Disable" -and $totalInactive -gt 0) {
		$disableCandidates = if ($SkipIfLastSignInIsNEVER) {
			$inactiveUsers | Where-Object { $_.LastSignInAgoDays -ne "Never" }
		} else {
			$inactiveUsers
		}

		$totalToDisable = $disableCandidates.Count
		if ($SkipIfLastSignInIsNEVER) {
			Write-Host ("Skip users with never sign-in: {0}" -f ($inactiveUsers.Count - $totalToDisable))
		}

		$counter = 0
		foreach ($user in $disableCandidates) {
			$counter++
			Write-Host ("Processing {0}/{1}: {2}" -f $counter, $totalToDisable, $user.UserPrincipalName) -ForegroundColor Yellow
			for ($countdown = 3; $countdown -ge 1; $countdown--) {
				Write-Host ("Disabling in {0}..." -f $countdown)
				Start-Sleep -Seconds 1
			}

			$disableStatus = "Success"
			$disableError = $null
			try {
				Update-MgUser -UserId $user.UserPrincipalName -AccountEnabled:$false -ErrorAction Stop
			}
			catch {
				$disableStatus = "Failed"
				$disableError = $_.Exception.Message
			}

			$disableReport += [pscustomobject]@{
				DisplayName       = $user.DisplayName
				UserPrincipalName = $user.UserPrincipalName
				Disabled          = $disableStatus
				Error             = $disableError
			}
		}
	}

	$disableReportFile = $null
	if ($disableReport.Count -gt 0) {
		$disableReportFile = Join-Path $reportFolder ("DisabledGuestUsers_{0}.csv" -f $timestamp)
		$disableReport | Export-Csv -Path $disableReportFile -NoTypeInformation -Encoding UTF8
	}

	Write-Host "Inactive Guest Users Report" -ForegroundColor Cyan
	Write-Host "--------------------------------------------"
	Write-Host ("Total enabled guest users : {0}" -f $totalGuests)
	Write-Host ("Inactive guest users      : {0}" -f $totalInactive)
	Write-Host ("Inactive percentage       : {0}%" -f $percentInactive)
	Write-Host ("Inactive days threshold   : {0}" -f $InactiveDays)
	Write-Host ("CSV report exported to    : {0}" -f $reportFile)
	Write-Host ("HTML report exported to   : {0}" -f $htmlReportFile)
	Write-Host ("Mode                     : {0}" -f $Mode)
	if ($disableReportFile) {
		Write-Host ("Disable report exported  : {0}" -f $disableReportFile)
	}

	$disconnectChoice = Read-Host "Disconnect from Microsoft Graph? (Y/N)"
	if ($disconnectChoice -match '^(y|yes)$') {
		Disconnect-MgGraph -ErrorAction SilentlyContinue
	}
}
catch {
	Write-Error $_
	exit 1
}
