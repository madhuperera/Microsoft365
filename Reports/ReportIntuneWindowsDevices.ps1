#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
	Reports on Intune-managed Windows devices grouped by OS generation and build,
	matching the OsGeneration / WinBuild / Total breakdown commonly used for
	Windows servicing reviews.

.DESCRIPTION
	Connects to Microsoft Graph and retrieves all Intune-managed Windows devices.
	For each device the script captures user, model, full OS version, parsed
	OS generation (Windows 10 / Windows 11 / Other Windows), Windows build number,
	last sync date and compliance state. Outputs a CSV file and an HTML report
	showing the OsGeneration + WinBuild + Total table, version spread cards and
	a sortable / filterable device table.

	Optionally accepts minimum supported build numbers for Windows 10 and
	Windows 11. Devices on a build lower than the supplied minimum are flagged
	as "Outdated" in the report.

.PARAMETER MinimumSupportedWindows10Build
	Optional. Minimum supported Windows 10 build number (e.g. 19045 for 22H2).
	Devices with a build lower than this value are flagged as Outdated.

.PARAMETER MinimumSupportedWindows11Build
	Optional. Minimum supported Windows 11 build number (e.g. 22631 for 23H2,
	26100 for 24H2). Devices with a build lower than this value are flagged
	as Outdated.

.PARAMETER ReportPath
	Folder for the output reports. If omitted the current working directory is used.

.EXAMPLE
	.\ReportIntuneWindowsDevices.ps1

.EXAMPLE
	.\ReportIntuneWindowsDevices.ps1 -MinimumSupportedWindows10Build 19045 -MinimumSupportedWindows11Build 22631

.EXAMPLE
	.\ReportIntuneWindowsDevices.ps1 -MinimumSupportedWindows11Build 26100 -ReportPath C:\Reports
#>

[CmdletBinding()]
param(
	[Parameter(Mandatory = $false)]
	[ValidateRange(1, 999999)]
	[int]$MinimumSupportedWindows10Build,

	[Parameter(Mandatory = $false)]
	[ValidateRange(1, 999999)]
	[int]$MinimumSupportedWindows11Build,

	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[string]$ReportPath
)

$ErrorActionPreference = "Stop"

$S_RequiredGraphScopes = @(
	'DeviceManagementManagedDevices.Read.All'
	'Organization.Read.All'
)

try {
	if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
		throw "Microsoft.Graph.Authentication module is not installed. Install it using Install-Module Microsoft.Graph -Scope CurrentUser."
	}
	Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

	# --- Connect to Graph ---
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
			Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome -ErrorAction Stop | Out-Null
		}
	}
	else
	{
		Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome -ErrorAction Stop | Out-Null
	}
	$S_ExistingContext = Get-MgContext

	Write-Host ""
	Write-Host "Active Graph context:" -ForegroundColor Cyan
	Write-Host "  Account    : $($S_ExistingContext.Account)" -ForegroundColor Cyan
	Write-Host "  TenantId   : $($S_ExistingContext.TenantId)" -ForegroundColor Cyan
	Write-Host "  Scopes     : $($S_ExistingContext.Scopes -join ', ')" -ForegroundColor Cyan
	Write-Host ""

	$S_ContextConfirmation = Read-Host "Proceed with this Graph context? [Y] Yes  [N] No  (Default: N)"
	if ([string]::IsNullOrWhiteSpace($S_ContextConfirmation)) { $S_ContextConfirmation = 'N' }
	if ($S_ContextConfirmation.ToUpperInvariant() -ne 'Y') {
		throw "Operation cancelled. Please reconnect to the correct tenant and account, then run again."
	}

	# --- Tenant info ---
	$tenantDisplayName = $null
	try {
		$orgResp = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization' -ErrorAction Stop
		if ($orgResp.value) { $tenantDisplayName = $orgResp.value[0].displayName }
	} catch { }
	if (-not $tenantDisplayName) { $tenantDisplayName = $S_ExistingContext.TenantId }
	$tenantId = if ($S_ExistingContext.TenantId) { $S_ExistingContext.TenantId } else { 'Unknown' }

	# --- Fetch managed Windows devices ---
	Write-Host "Fetching Intune-managed Windows devices..." -ForegroundColor Cyan
	$select = 'id,deviceName,userPrincipalName,userDisplayName,operatingSystem,osVersion,model,manufacturer,enrolledDateTime,lastSyncDateTime,complianceState,managedDeviceOwnerType,joinType,serialNumber'
	$filter = "operatingSystem eq 'Windows'"
	$encodedFilter = [System.Uri]::EscapeDataString($filter)
	$uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=$encodedFilter&`$select=$select&`$top=200"

	$devices = New-Object System.Collections.Generic.List[object]
	do {
		$resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
		if ($resp.value) {
			foreach ($d in $resp.value) { $devices.Add([pscustomobject]$d) | Out-Null }
		}
		$uri = $resp.'@odata.nextLink'
	} while ($uri)

	Write-Host ("  Retrieved {0} Windows devices" -f $devices.Count) -ForegroundColor Green

	# --- Build report rows ---
	# OS version format from Intune is typically "10.0.19045.4046".
	# Win10 and Win11 both report major 10; Win11 is identified by build >= 22000.
	$now = Get-Date
	$report = foreach ($d in $devices) {
		$rawVer = if ($d.osVersion) { [string]$d.osVersion } else { '' }

		$winBuild = $null
		if ($rawVer -match '^\s*\d+\.\d+\.(\d+)') {
			$winBuild = [int]$Matches[1]
		}

		$osGeneration = 'Other Windows'
		if ($null -ne $winBuild) {
			if ($winBuild -ge 22000) { $osGeneration = 'Windows 11' }
			elseif ($winBuild -ge 10240) { $osGeneration = 'Windows 10' }
		}

		$threshold = $null
		if ($osGeneration -eq 'Windows 10' -and $PSBoundParameters.ContainsKey('MinimumSupportedWindows10Build')) {
			$threshold = $MinimumSupportedWindows10Build
		}
		elseif ($osGeneration -eq 'Windows 11' -and $PSBoundParameters.ContainsKey('MinimumSupportedWindows11Build')) {
			$threshold = $MinimumSupportedWindows11Build
		}

		$supportStatus = 'Unknown'
		if ($null -eq $winBuild) {
			$supportStatus = 'Unknown'
		}
		elseif ($null -ne $threshold) {
			if ($winBuild -lt $threshold) { $supportStatus = 'Outdated' } else { $supportStatus = 'Supported' }
		}
		else {
			$supportStatus = 'NoThreshold'
		}

		$lastSync = if ($d.lastSyncDateTime) { [datetime]$d.lastSyncDateTime } else { $null }
		$daysSinceSync = if ($lastSync) { [int]($now - $lastSync).TotalDays } else { $null }

		[pscustomobject]@{
			DeviceName        = $d.deviceName
			User              = if ($d.userDisplayName) { $d.userDisplayName } else { $d.userPrincipalName }
			UserPrincipalName = $d.userPrincipalName
			OsGeneration      = $osGeneration
			OSVersion         = $rawVer
			WinBuild          = $winBuild
			MinimumSupported  = $threshold
			SupportStatus     = $supportStatus
			Manufacturer      = $d.manufacturer
			Model             = $d.model
			Ownership         = $d.managedDeviceOwnerType
			JoinType          = $d.joinType
			ComplianceState   = $d.complianceState
			EnrolledDateTime  = $d.enrolledDateTime
			LastSyncDateTime  = $d.lastSyncDateTime
			DaysSinceLastSync = $daysSinceSync
			SerialNumber      = $d.serialNumber
		}
	}

	# --- Stats ---
	$totalDevices = $report.Count
	$win10Devices = $report | Where-Object { $_.OsGeneration -eq 'Windows 10' }
	$win11Devices = $report | Where-Object { $_.OsGeneration -eq 'Windows 11' }
	$otherDevices = $report | Where-Object { $_.OsGeneration -eq 'Other Windows' }

	$totalWin10  = $win10Devices.Count
	$totalWin11  = $win11Devices.Count
	$totalOther  = $otherDevices.Count

	$outdatedWin10 = ($win10Devices | Where-Object { $_.SupportStatus -eq 'Outdated' }).Count
	$outdatedWin11 = ($win11Devices | Where-Object { $_.SupportStatus -eq 'Outdated' }).Count
	$totalOutdated = $outdatedWin10 + $outdatedWin11

	# --- Build OsGeneration + WinBuild breakdown (matches the Intune-style table) ---
	$generationOrder = @{ 'Other Windows' = 0; 'Windows 10' = 1; 'Windows 11' = 2 }
	$breakdown = $report | Group-Object OsGeneration, WinBuild | ForEach-Object {
		$first = $_.Group[0]
		$gen = $first.OsGeneration
		$build = $first.WinBuild
		$min = $null
		if ($gen -eq 'Windows 10' -and $PSBoundParameters.ContainsKey('MinimumSupportedWindows10Build')) { $min = $MinimumSupportedWindows10Build }
		elseif ($gen -eq 'Windows 11' -and $PSBoundParameters.ContainsKey('MinimumSupportedWindows11Build')) { $min = $MinimumSupportedWindows11Build }
		$outdated = ($null -ne $min -and $null -ne $build -and $build -lt $min)
		[pscustomobject]@{
			OsGeneration = $gen
			WinBuild     = if ($null -eq $build) { 0 } else { $build }
			Total        = $_.Count
			Outdated     = $outdated
		}
	} | Sort-Object @{Expression={$generationOrder[$_.OsGeneration]}}, WinBuild

	# --- Output paths ---
	if (-not $ReportPath) { $ReportPath = (Get-Location).Path }
	$reportFolder = if (Test-Path $ReportPath -PathType Container) { $ReportPath } else { Split-Path -Parent $ReportPath }
	if ($reportFolder -and -not (Test-Path $reportFolder)) {
		New-Item -ItemType Directory -Path $reportFolder -Force | Out-Null
	}
	$S_Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
	$csvFile  = Join-Path $reportFolder ("ReportIntuneWindowsDevices_{0}.csv"  -f $S_Timestamp)
	$breakdownCsv = Join-Path $reportFolder ("ReportIntuneWindowsDevices_Breakdown_{0}.csv" -f $S_Timestamp)
	$htmlFile = Join-Path $reportFolder ("ReportIntuneWindowsDevices_{0}.html" -f $S_Timestamp)

	# --- CSV exports ---
	$report | Sort-Object OsGeneration, WinBuild, DeviceName | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
	$breakdown | Select-Object OsGeneration, WinBuild, Total, Outdated | Export-Csv -Path $breakdownCsv -NoTypeInformation -Encoding UTF8

	# --- HTML helpers ---
	$enc = { param($s) if ($null -eq $s -or $s -eq '') { '-' } else { [System.Net.WebUtility]::HtmlEncode([string]$s) } }
	$reportDate = Get-Date -Format "dd MMM yyyy HH:mm"

	$win10Threshold = if ($PSBoundParameters.ContainsKey('MinimumSupportedWindows10Build')) { $MinimumSupportedWindows10Build } else { $null }
	$win11Threshold = if ($PSBoundParameters.ContainsKey('MinimumSupportedWindows11Build')) { $MinimumSupportedWindows11Build } else { $null }
	$win10ThresholdDisp = if ($null -ne $win10Threshold) { $win10Threshold } else { 'not set' }
	$win11ThresholdDisp = if ($null -ne $win11Threshold) { $win11Threshold } else { 'not set' }

	# Breakdown table rows (the OsGeneration / WinBuild / Total view)
	$breakdownRows = ($breakdown | ForEach-Object {
		$cls = if ($_.Outdated) { ' class="outdated-row"' } else { '' }
		$buildText = if ($_.WinBuild -eq 0) { '-' } else { $_.WinBuild }
		$badge = if ($_.Outdated) { " <span class='badge badge-inactive'>Outdated</span>" } else { '' }
		"<tr$cls><td>$(& $enc $_.OsGeneration)$badge</td><td>$buildText</td><td>$($_.Total)</td></tr>"
	}) -join "`n"

	# Build version spread cards per OsGeneration
	function Build-SpreadCardsHtml {
		param([object[]]$Items, [string]$Label)
		if (-not $Items -or $Items.Count -eq 0) {
			return "<div class='dist-card'><div class='dist-label'>No $Label devices</div><div class='dist-value'>0</div></div>"
		}
		($Items | ForEach-Object {
			$cls = if ($_.Outdated) { 'dist-card outdated' } else { 'dist-card' }
			$badge = if ($_.Outdated) { "<div class='outdated-badge'>Outdated</div>" } else { '' }
			$buildText = if ($_.WinBuild -eq 0) { 'Unknown' } else { $_.WinBuild }
			"<div class='$cls'><div class='dist-label'>$Label $buildText</div><div class='dist-value'>$($_.Total)</div>$badge</div>"
		}) -join "`n"
	}

	$win10Items = $breakdown | Where-Object { $_.OsGeneration -eq 'Windows 10' }
	$win11Items = $breakdown | Where-Object { $_.OsGeneration -eq 'Windows 11' }
	$otherItems = $breakdown | Where-Object { $_.OsGeneration -eq 'Other Windows' }

	$win10CardsHtml = Build-SpreadCardsHtml -Items $win10Items -Label 'Build'
	$win11CardsHtml = Build-SpreadCardsHtml -Items $win11Items -Label 'Build'
	$otherCardsHtml = Build-SpreadCardsHtml -Items $otherItems -Label 'Build'

	# Build chart data JSON
	function ConvertTo-ChartJson {
		param([object[]]$Items)
		if (-not $Items -or $Items.Count -eq 0) { return '{"labels":[],"data":[],"outdated":[]}' }
		$labels = ($Items | ForEach-Object {
			$b = if ($_.WinBuild -eq 0) { 'Unknown' } else { [string]$_.WinBuild }
			'"' + $b + '"'
		}) -join ','
		$counts  = ($Items | ForEach-Object { $_.Total }) -join ','
		$outFlag = ($Items | ForEach-Object { if ($_.Outdated) { 'true' } else { 'false' } }) -join ','
		"{`"labels`":[$labels],`"data`":[$counts],`"outdated`":[$outFlag]}"
	}

	$win10ChartJson = ConvertTo-ChartJson -Items $win10Items
	$win11ChartJson = ConvertTo-ChartJson -Items $win11Items
	$otherChartJson = ConvertTo-ChartJson -Items $otherItems

	# Device table rows
	$tableRows = ($report | Sort-Object OsGeneration, WinBuild, DeviceName | ForEach-Object {
		$daysVal      = if ($null -ne $_.DaysSinceLastSync) { $_.DaysSinceLastSync } else { -1 }
		$lastSyncDisp = if ($_.LastSyncDateTime) { ([datetime]$_.LastSyncDateTime).ToString("dd MMM yyyy") } else { '-' }
		$enrolDisp    = if ($_.EnrolledDateTime) { ([datetime]$_.EnrolledDateTime).ToString("dd MMM yyyy") } else { '-' }
		$statusClass  = switch ($_.SupportStatus) {
			'Outdated'    { 'badge-inactive' }
			'Supported'   { 'badge-active' }
			'NoThreshold' { 'badge-disabled' }
			default        { 'badge-disabled' }
		}
		$statusText = if ($_.SupportStatus -eq 'NoThreshold') { 'No Threshold' } else { $_.SupportStatus }
		$compClass = switch ($_.ComplianceState) {
			'compliant'    { 'badge-active' }
			'noncompliant' { 'badge-inactive' }
			default         { 'badge-disabled' }
		}
		$buildText = if ($null -ne $_.WinBuild) { $_.WinBuild } else { '-' }
		$rowAttr = "data-generation=`"$($_.OsGeneration)`" data-status=`"$($_.SupportStatus)`""
		"<tr $rowAttr>" +
			"<td>$(& $enc $_.DeviceName)</td>" +
			"<td>$(& $enc $_.User)</td>" +
			"<td>$(& $enc $_.OsGeneration)</td>" +
			"<td>$(& $enc $_.OSVersion)</td>" +
			"<td>$buildText</td>" +
			"<td><span class='badge $statusClass'>$statusText</span></td>" +
			"<td>$(& $enc $_.Manufacturer)</td>" +
			"<td>$(& $enc $_.Model)</td>" +
			"<td>$(& $enc $_.Ownership)</td>" +
			"<td>$(& $enc $_.JoinType)</td>" +
			"<td><span class='badge $compClass'>$(& $enc $_.ComplianceState)</span></td>" +
			"<td>$enrolDisp</td>" +
			"<td>$lastSyncDisp</td>" +
			"<td>$(if ($daysVal -ge 0) { "$daysVal days" } else { 'Never' })</td>" +
		"</tr>"
	}) -join "`n"

	$pctOutdated = if ($totalDevices -gt 0) { [math]::Round(($totalOutdated / $totalDevices) * 100, 1) } else { 0 }

	# --- HTML report ---
	$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Intune Windows Devices Report</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; padding: 30px; }
  .header { background: linear-gradient(135deg, #1a1a2e, #16213e); color: #fff; padding: 30px 40px; border-radius: 12px; margin-bottom: 30px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 16px; }
  .header-left h1 { font-size: 1.6em; margin-bottom: 6px; }
  .header-left p { font-size: 0.9em; opacity: 0.85; }
  .header-right { font-size: 0.9em; opacity: 0.9; text-align: right; }
  .header-right strong { color: #ffd166; }

  .section-title { font-size: 1.15em; font-weight: 600; color: #1a1a2e; margin: 0 0 16px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0; }

  .summary-cards { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .card { background: #fff; border-radius: 10px; padding: 24px 30px; flex: 1; min-width: 180px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
  .card .label { font-size: 0.85em; color: #777; text-transform: uppercase; letter-spacing: 0.5px; }
  .card .value { font-size: 2em; font-weight: 700; margin-top: 6px; }
  .card .sub { font-size: 0.78em; color: #999; margin-top: 2px; }

  .dist-section { margin-bottom: 30px; }
  .dist-cards { display: flex; gap: 14px; flex-wrap: wrap; }
  .dist-card { background: #fff; border-radius: 10px; padding: 16px 22px; min-width: 130px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); border-left: 4px solid #3498db; text-align: center; position: relative; }
  .dist-card .dist-label { font-size: 0.8em; color: #555; margin-bottom: 6px; text-transform: uppercase; letter-spacing: 0.3px; }
  .dist-card .dist-value { font-size: 1.6em; font-weight: 700; color: #1a1a2e; }
  .dist-card.outdated { border-left-color: #e74c3c; background: #fff5f5; }
  .dist-card.outdated .dist-value { color: #c0392b; }
  .outdated-badge { display: inline-block; margin-top: 6px; padding: 2px 8px; font-size: 0.7em; font-weight: 700; color: #fff; background: #e74c3c; border-radius: 10px; text-transform: uppercase; letter-spacing: 0.4px; }

  .charts-row { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .chart-section { background: #fff; border-radius: 10px; padding: 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); flex: 1; min-width: 320px; }
  .chart-section h2 { font-size: 1.05em; margin-bottom: 16px; color: #1a1a2e; }
  .chart-container { max-width: 380px; margin: 0 auto; }

  .table-section { background: #fff; border-radius: 10px; padding: 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 30px; overflow-x: auto; }
  .table-section h2 { font-size: 1.1em; margin-bottom: 16px; color: #1a1a2e; }
  .table-controls { display: flex; gap: 14px; margin-bottom: 16px; flex-wrap: wrap; align-items: center; }
  .table-controls input[type="text"] { padding: 8px 14px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; min-width: 260px; }
  .table-controls select { padding: 8px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; background: #fff; }
  .table-controls .count-label { font-size: 0.85em; color: #777; margin-left: auto; }
  table { width: 100%; border-collapse: collapse; font-size: 0.85em; }
  th { background: #1a1a2e; color: #fff; padding: 10px 12px; text-align: left; cursor: pointer; user-select: none; white-space: nowrap; position: sticky; top: 0; }
  th:hover { background: #2c3e50; }
  td { padding: 9px 12px; border-bottom: 1px solid #eee; white-space: nowrap; }
  tr:hover td { background: #f8f9fa; }
  tr.hidden-row { display: none; }
  tr.outdated-row td { background: #fff5f5; }
  tr.outdated-row:hover td { background: #ffe8e8; }

  .badge { padding: 3px 10px; border-radius: 12px; font-size: 0.8em; font-weight: 600; }
  .badge-active { background: #d4edda; color: #155724; }
  .badge-inactive { background: #f8d7da; color: #721c24; }
  .badge-disabled { background: #e2e3e5; color: #495057; }

  .breakdown-table { background: #fff; border-radius: 10px; padding: 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 30px; }
  .breakdown-table table { max-width: 600px; }
  .breakdown-table h2 { font-size: 1.1em; margin-bottom: 16px; color: #1a1a2e; }

  .footer { text-align: center; font-size: 0.8em; color: #999; margin-top: 20px; }
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <h1>Intune Windows Devices Report</h1>
    <p>Tenant: $(& $enc $tenantDisplayName) ($tenantId) &nbsp;|&nbsp; Generated: $reportDate</p>
  </div>
  <div class="header-right">
    Min supported Windows 10 build: <strong>$win10ThresholdDisp</strong><br/>
    Min supported Windows 11 build: <strong>$win11ThresholdDisp</strong>
  </div>
</div>

<!-- OVERVIEW -->
<div class="section-title">Overview</div>
<div class="summary-cards">
  <div class="card"><div class="label">Total Windows Devices</div><div class="value" style="color:#1a1a2e;">$totalDevices</div></div>
  <div class="card"><div class="label">Windows 10</div><div class="value" style="color:#3498db;">$totalWin10</div><div class="sub">$outdatedWin10 outdated</div></div>
  <div class="card"><div class="label">Windows 11</div><div class="value" style="color:#27ae60;">$totalWin11</div><div class="sub">$outdatedWin11 outdated</div></div>
  <div class="card"><div class="label">Other Windows</div><div class="value" style="color:#9b59b6;">$totalOther</div></div>
  <div class="card"><div class="label">Total Outdated</div><div class="value" style="color:#e74c3c;">$totalOutdated</div><div class="sub">$pctOutdated% of total</div></div>
</div>

<!-- BREAKDOWN TABLE (matches the OsGeneration / WinBuild / Total view) -->
<div class="breakdown-table">
  <h2>OsGeneration / WinBuild / Total</h2>
  <table>
    <thead><tr><th>OsGeneration</th><th>WinBuild</th><th>Total</th></tr></thead>
    <tbody>
$breakdownRows
    </tbody>
  </table>
</div>

<!-- VERSION SPREAD CARDS -->
<div class="dist-section">
  <div class="section-title">Windows 10 Build Spread (Min Supported: $win10ThresholdDisp)</div>
  <div class="dist-cards">
$win10CardsHtml
  </div>
</div>

<div class="dist-section">
  <div class="section-title">Windows 11 Build Spread (Min Supported: $win11ThresholdDisp)</div>
  <div class="dist-cards">
$win11CardsHtml
  </div>
</div>

<div class="dist-section">
  <div class="section-title">Other Windows Build Spread</div>
  <div class="dist-cards">
$otherCardsHtml
  </div>
</div>

<!-- CHARTS -->
<div class="charts-row">
  <div class="chart-section"><h2>Windows 10 by Build</h2><div class="chart-container"><canvas id="win10Chart"></canvas></div></div>
  <div class="chart-section"><h2>Windows 11 by Build</h2><div class="chart-container"><canvas id="win11Chart"></canvas></div></div>
  <div class="chart-section"><h2>Other Windows by Build</h2><div class="chart-container"><canvas id="otherChart"></canvas></div></div>
</div>

<!-- DEVICE TABLE -->
<div class="table-section">
  <h2>Device Details</h2>
  <div class="table-controls">
    <input type="text" id="searchBox" placeholder="Search by name, user, model, build..." onkeyup="filterTable()" />
    <select id="generationFilter" onchange="filterTable()">
      <option value="all">All Generations</option>
      <option value="Windows 10">Windows 10</option>
      <option value="Windows 11">Windows 11</option>
      <option value="Other Windows">Other Windows</option>
    </select>
    <select id="statusFilter" onchange="filterTable()">
      <option value="all">All Versions</option>
      <option value="Outdated">Outdated Only</option>
      <option value="Supported">Supported Only</option>
      <option value="NoThreshold">No Threshold</option>
      <option value="Unknown">Unknown</option>
    </select>
    <span class="count-label" id="rowCount"></span>
  </div>
  <table id="deviceTable">
    <thead><tr>
      <th onclick="sortTable(0)">Device Name</th>
      <th onclick="sortTable(1)">User</th>
      <th onclick="sortTable(2)">OS Generation</th>
      <th onclick="sortTable(3)">OS Version</th>
      <th onclick="sortTable(4)">Build</th>
      <th onclick="sortTable(5)">Support Status</th>
      <th onclick="sortTable(6)">Manufacturer</th>
      <th onclick="sortTable(7)">Model</th>
      <th onclick="sortTable(8)">Ownership</th>
      <th onclick="sortTable(9)">Join Type</th>
      <th onclick="sortTable(10)">Compliance</th>
      <th onclick="sortTable(11)">Enrolled</th>
      <th onclick="sortTable(12)">Last Sync</th>
      <th onclick="sortTable(13)">Days Since Sync</th>
    </tr></thead>
    <tbody>
$tableRows
    </tbody>
  </table>
</div>

<div class="footer">Report generated by ReportIntuneWindowsDevices.ps1</div>

<script>
var win10Data = $win10ChartJson;
var win11Data = $win11ChartJson;
var otherData = $otherChartJson;

function buildChart(canvasId, payload, baseColor) {
  if (!payload.labels.length) { return; }
  var bg = payload.outdated.map(function(o){ return o ? '#e74c3c' : baseColor; });
  new Chart(document.getElementById(canvasId), {
    type: 'bar',
    data: { labels: payload.labels, datasets: [{ label: 'Devices', data: payload.data, backgroundColor: bg, borderWidth: 0 }] },
    options: {
      responsive: true,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: function(ctx) {
        var t = ctx.dataset.data.reduce(function(a,b){return a+b;},0);
        var pct = t > 0 ? ((ctx.parsed.y / t) * 100).toFixed(1) : 0;
        var flag = payload.outdated[ctx.dataIndex] ? ' (Outdated)' : '';
        return ctx.parsed.y + ' devices (' + pct + '%)' + flag;
      } } } },
      scales: { y: { beginAtZero: true, ticks: { precision: 0 } } }
    }
  });
}

buildChart('win10Chart', win10Data, '#3498db');
buildChart('win11Chart', win11Data, '#27ae60');
buildChart('otherChart', otherData, '#9b59b6');

function filterTable() {
  var search = document.getElementById('searchBox').value.toLowerCase();
  var generation = document.getElementById('generationFilter').value;
  var status = document.getElementById('statusFilter').value;
  var rows = document.querySelectorAll('#deviceTable tbody tr');
  var visible = 0;
  rows.forEach(function(row) {
    var text = row.textContent.toLowerCase();
    var matchSearch = !search || text.indexOf(search) !== -1;
    var matchGen = generation === 'all' || row.getAttribute('data-generation') === generation;
    var matchStatus = status === 'all' || row.getAttribute('data-status') === status;
    if (matchSearch && matchGen && matchStatus) {
      row.classList.remove('hidden-row');
      visible++;
    } else {
      row.classList.add('hidden-row');
    }
  });
  document.getElementById('rowCount').textContent = 'Showing ' + visible + ' of ' + rows.length + ' devices';
}

var sortDir = {};
function sortTable(col) {
  var tbody = document.getElementById('deviceTable').querySelector('tbody');
  var rows = Array.from(tbody.querySelectorAll('tr'));
  var dir = sortDir[col] === 'asc' ? 'desc' : 'asc';
  sortDir[col] = dir;
  rows.sort(function(a, b) {
    var av = a.cells[col].textContent.trim().toLowerCase();
    var bv = b.cells[col].textContent.trim().toLowerCase();
    var an = parseFloat(av), bn = parseFloat(bv);
    if (!isNaN(an) && !isNaN(bn)) { return dir === 'asc' ? an - bn : bn - an; }
    if (av < bv) return dir === 'asc' ? -1 : 1;
    if (av > bv) return dir === 'asc' ? 1 : -1;
    return 0;
  });
  rows.forEach(function(row) { tbody.appendChild(row); });
}

filterTable();
</script>
</body>
</html>
"@

	$html | Out-File -FilePath $htmlFile -Encoding UTF8

	# --- Console summary ---
	Write-Host ""
	Write-Host "Intune Windows Devices Report" -ForegroundColor Cyan
	Write-Host "--------------------------------------------"
	Write-Host ("Tenant                       : {0} ({1})" -f $tenantDisplayName, $tenantId)
	Write-Host ("Min supported Windows 10     : {0}" -f $win10ThresholdDisp)
	Write-Host ("Min supported Windows 11     : {0}" -f $win11ThresholdDisp)
	Write-Host ("Total Windows devices        : {0}" -f $totalDevices)
	Write-Host ("  Windows 10                 : {0}  (Outdated: {1})" -f $totalWin10, $outdatedWin10) -ForegroundColor Cyan
	Write-Host ("  Windows 11                 : {0}  (Outdated: {1})" -f $totalWin11, $outdatedWin11) -ForegroundColor Green
	Write-Host ("  Other Windows              : {0}" -f $totalOther) -ForegroundColor DarkGray
	Write-Host ("Total outdated               : {0}  ({1}%)" -f $totalOutdated, $pctOutdated) -ForegroundColor Red
	Write-Host ""
	Write-Host "OsGeneration / WinBuild / Total" -ForegroundColor Cyan
	$breakdown | Format-Table OsGeneration, WinBuild, Total, Outdated -AutoSize | Out-String | Write-Host
	Write-Host ("CSV report (devices)         : {0}" -f $csvFile) -ForegroundColor Yellow
	Write-Host ("CSV report (breakdown)       : {0}" -f $breakdownCsv) -ForegroundColor Yellow
	Write-Host ("HTML report                  : {0}" -f $htmlFile) -ForegroundColor Yellow

	$S_DisconnectChoice = Read-Host "Disconnect from Microsoft Graph? (Y/N)"
	if ($S_DisconnectChoice -match '^(y|yes)$') {
		Disconnect-MgGraph -ErrorAction SilentlyContinue
	}
}
catch {
	Write-Error $_
	exit 1
}
