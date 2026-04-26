param(
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[string]$ReportPath,

	[Parameter(Mandatory = $false)]
	[ValidateRange(1, 3650)]
	[int]$InactiveDays = 180
)

$ErrorActionPreference = "Stop"

$S_RequiredGraphScopes = @(
	'Device.Read.All'
	'Organization.Read.All'
)

try {
	if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Identity.DirectoryManagement)) {
		throw "Microsoft.Graph.Identity.DirectoryManagement module is not installed. Install it using Install-Module Microsoft.Graph -Scope CurrentUser."
	}

	Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

	$context = Get-MgContext
	if (-not $context) {
		Connect-MgGraph -Scopes $S_RequiredGraphScopes -ErrorAction Stop | Out-Null
		$context = Get-MgContext
	}

	# --- Resolve tenant display name ---
	$tenantDisplayName = $null
	try {
		$org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
		$tenantDisplayName = $org.DisplayName
	} catch { }
	if (-not $tenantDisplayName) { $tenantDisplayName = $context.TenantId }
	$tenantId = if ($context.TenantId) { $context.TenantId } else { "Unknown" }

	# --- Fetch all Windows devices ---
	$devices = Get-MgDevice -All `
		-Filter "operatingSystem eq 'Windows'" `
		-Property "id,displayName,operatingSystem,operatingSystemVersion,trustType,accountEnabled,approximateLastSignInDateTime,managementType,registrationDateTime,isCompliant" `
		-ErrorAction Stop

	# --- Build report data ---
	$now = Get-Date
	$cutoffDate = $now.AddDays(-$InactiveDays)
	$report = foreach ($device in $devices) {
		$regDt = if ($device.RegistrationDateTime) { [datetime]$device.RegistrationDateTime } else { $null }
		$lastActivityDt = if ($device.ApproximateLastSignInDateTime) { [datetime]$device.ApproximateLastSignInDateTime } else { $null }

		$daysSinceEnrollment = if ($regDt -and $lastActivityDt) { [int]($lastActivityDt - $regDt).TotalDays } else { $null }
		$daysSinceLastActivity = if ($lastActivityDt) { [int]($now - $lastActivityDt).TotalDays } else { $null }

		# Status: Disabled > Inactive > Active
		if (-not $device.AccountEnabled) {
			$status = "Disabled"
		} elseif (-not $lastActivityDt -or $lastActivityDt -lt $cutoffDate) {
			$status = "Inactive"
		} else {
			$status = "Active"
		}

		[pscustomobject]@{
			DisplayName                      = $device.DisplayName
			AccountEnabled                   = $device.AccountEnabled
			OperatingSystem                  = $device.OperatingSystem
			OperatingSystemVersion           = $device.OperatingSystemVersion
			TrustType                        = $device.TrustType
			ManagementType                   = $device.ManagementType
			IsCompliant                      = $device.IsCompliant
			RegistrationDateTime             = $device.RegistrationDateTime
			ApproximateLastSignInDateTime    = $device.ApproximateLastSignInDateTime
			DaysEnrollmentToLastActivity     = $daysSinceEnrollment
			DaysSinceLastActivity            = $daysSinceLastActivity
			Status                           = $status
		}
	}

	# --- Stats ---
	$totalDevices    = $report.Count
	$totalEnabled    = ($report | Where-Object { $_.AccountEnabled }).Count
	$totalDisabled   = $totalDevices - $totalEnabled
	$totalCompliant  = ($report | Where-Object { $_.IsCompliant -eq $true }).Count
	$totalActive     = ($report | Where-Object { $_.Status -eq "Active" }).Count
	$totalInactive   = ($report | Where-Object { $_.Status -eq "Inactive" }).Count
	$percentActive   = if ($totalEnabled -gt 0) { [math]::Round(($totalActive / $totalEnabled) * 100, 1) } else { 0 }
	$percentInactive = if ($totalEnabled -gt 0) { [math]::Round(($totalInactive / $totalEnabled) * 100, 1) } else { 0 }

	$trustSummary = $report | Group-Object TrustType | Sort-Object Count -Descending | ForEach-Object {
		[pscustomobject]@{ TrustType = $_.Name; Count = $_.Count }
	}

	$mgmtSummary = $report | Group-Object ManagementType | Sort-Object Count -Descending | ForEach-Object {
		[pscustomobject]@{ ManagementType = $_.Name; Count = $_.Count }
	}

	# --- File paths ---
	if (-not $ReportPath) {
		$ReportPath = (Get-Location).Path
	}

	$reportFolder = if (Test-Path $ReportPath -PathType Container) { $ReportPath } else { Split-Path -Parent $ReportPath }
	if ($reportFolder -and -not (Test-Path $reportFolder)) {
		New-Item -ItemType Directory -Path $reportFolder -Force | Out-Null
	}

	$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
	$csvFile = if (Test-Path $ReportPath -PathType Container) {
		Join-Path $ReportPath ("AllWindowsDevices_{0}.csv" -f $timestamp)
	} else {
		$ReportPath
	}

	# --- CSV export ---
	$report | Sort-Object DisplayName | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8

	# --- HTML report ---
	$reportDate = Get-Date -Format "dd MMM yyyy HH:mm"

	# Build per-device JSON for client-side threshold recalculation
	$devicesJson = ($report | Sort-Object DisplayName | ForEach-Object {
		$daysVal = if ($null -ne $_.DaysSinceLastActivity) { $_.DaysSinceLastActivity } else { -1 }
		$ena = if ($_.AccountEnabled) { "true" } else { "false" }
		$comp = if ($_.IsCompliant -eq $true) { "true" } elseif ($_.IsCompliant -eq $false) { "false" } else { "null" }
		$tt = if ($_.TrustType) { $_.TrustType } else { "None" }
		$mt = if ($_.ManagementType) { $_.ManagementType } else { "None" }
		'{{"days":{0},"ena":{1},"comp":{2},"tt":"{3}","mt":"{4}"}}' -f $daysVal, $ena, $comp, $tt, $mt
	}) -join ","

	# Build table rows with data attributes
	$tableRows = ($report | Sort-Object DisplayName | ForEach-Object {
		$daysVal = if ($null -ne $_.DaysSinceLastActivity) { $_.DaysSinceLastActivity } else { -1 }
		$enabled = if ($_.AccountEnabled) { "Yes" } else { "No" }
		$compliant = if ($_.IsCompliant -eq $true) { "Yes" } elseif ($_.IsCompliant -eq $false) { "No" } else { "-" }
		$trust = if ($_.TrustType) { [System.Net.WebUtility]::HtmlEncode($_.TrustType) } else { "-" }
		$mgmt = if ($_.ManagementType) { [System.Net.WebUtility]::HtmlEncode($_.ManagementType) } else { "-" }
		$regDate = if ($_.RegistrationDateTime) { ([datetime]$_.RegistrationDateTime).ToString("dd MMM yyyy") } else { "-" }
		$lastActivity = if ($_.ApproximateLastSignInDateTime) { ([datetime]$_.ApproximateLastSignInDateTime).ToString("dd MMM yyyy") } else { "-" }
		$enrollToActivity = if ($null -ne $_.DaysEnrollmentToLastActivity) { "$($_.DaysEnrollmentToLastActivity) days" } else { "-" }
		$sinceActivity = if ($null -ne $_.DaysSinceLastActivity) { "$($_.DaysSinceLastActivity) days" } else { "Never" }
		$enabledClass = if ($_.AccountEnabled) { "active" } else { "disabled" }
		$compliantClass = if ($_.IsCompliant -eq $true) { "active" } elseif ($_.IsCompliant -eq $false) { "inactive" } else { "" }
		$statusClass = switch ($_.Status) { "Active" { "active" } "Inactive" { "inactive" } "Disabled" { "disabled" } }

		"<tr data-days=`"$daysVal`" data-ena=`"$(if ($_.AccountEnabled) { '1' } else { '0' })`"><td>$([System.Net.WebUtility]::HtmlEncode($_.DisplayName))</td><td><span class=`"badge badge-$enabledClass`">$enabled</span></td><td>$([System.Net.WebUtility]::HtmlEncode($_.OperatingSystem))</td><td>$([System.Net.WebUtility]::HtmlEncode($_.OperatingSystemVersion))</td><td>$trust</td><td>$mgmt</td><td><span class=`"badge badge-$compliantClass`">$compliant</span></td><td>$regDate</td><td>$lastActivity</td><td>$enrollToActivity</td><td>$sinceActivity</td><td><span class=`"badge badge-$statusClass`">$($_.Status)</span></td></tr>"
	}) -join "`n"

	$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>All Windows Devices Report</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; padding: 30px; }
  .header { background: linear-gradient(135deg, #1a1a2e, #16213e); color: #fff; padding: 30px 40px; border-radius: 12px; margin-bottom: 30px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 16px; }
  .header-left h1 { font-size: 1.6em; margin-bottom: 6px; }
  .header-left p { font-size: 0.9em; opacity: 0.8; }
  .header-right { display: flex; align-items: center; gap: 10px; }
  .header-right label { font-size: 0.9em; opacity: 0.85; }
  .header-right select { padding: 8px 14px; border: none; border-radius: 6px; font-size: 0.95em; font-weight: 600; background: rgba(255,255,255,0.15); color: #fff; cursor: pointer; }
  .header-right select option { color: #333; background: #fff; }

  .section-title { font-size: 1.15em; font-weight: 600; color: #1a1a2e; margin-bottom: 16px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0; }

  .summary-cards { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .card { background: #fff; border-radius: 10px; padding: 24px 30px; flex: 1; min-width: 180px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
  .card .label { font-size: 0.85em; color: #777; text-transform: uppercase; letter-spacing: 0.5px; }
  .card .value { font-size: 2em; font-weight: 700; margin-top: 6px; }
  .card .sub { font-size: 0.78em; color: #999; margin-top: 2px; }

  .dist-section { margin-bottom: 30px; }
  .dist-cards { display: flex; gap: 16px; flex-wrap: wrap; }
  .dist-card { background: #fff; border-radius: 10px; padding: 18px 24px; min-width: 140px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); border-left: 4px solid #3498db; text-align: center; }
  .dist-card .dist-label { font-size: 0.82em; color: #555; margin-bottom: 6px; text-transform: uppercase; letter-spacing: 0.3px; }
  .dist-card .dist-value { font-size: 1.6em; font-weight: 700; color: #1a1a2e; }

  .charts-row { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .chart-section { background: #fff; border-radius: 10px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); flex: 1; min-width: 360px; }
  .chart-section h2 { font-size: 1.1em; margin-bottom: 20px; color: #1a1a2e; }
  .chart-container { max-width: 400px; margin: 0 auto; }

  .table-section { background: #fff; border-radius: 10px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 30px; overflow-x: auto; }
  .table-section h2 { font-size: 1.1em; margin-bottom: 16px; color: #1a1a2e; }
  .table-controls { display: flex; gap: 14px; margin-bottom: 16px; flex-wrap: wrap; align-items: center; }
  .table-controls input[type="text"] { padding: 8px 14px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; min-width: 260px; }
  .table-controls select { padding: 8px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; background: #fff; }
  .table-controls .count-label { font-size: 0.85em; color: #777; margin-left: auto; }
  table { width: 100%; border-collapse: collapse; font-size: 0.86em; }
  th { background: #1a1a2e; color: #fff; padding: 12px 14px; text-align: left; position: sticky; top: 0; cursor: pointer; user-select: none; white-space: nowrap; }
  th:hover { background: #2c3e50; }
  td { padding: 10px 14px; border-bottom: 1px solid #eee; white-space: nowrap; }
  tr:hover td { background: #f8f9fa; }
  tr.hidden-row { display: none; }

  .badge { padding: 3px 10px; border-radius: 12px; font-size: 0.82em; font-weight: 600; }
  .badge-active { background: #d4edda; color: #155724; }
  .badge-inactive { background: #f8d7da; color: #721c24; }
  .badge-disabled { background: #e2e3e5; color: #495057; }

  .activity-green { color: #155724; font-weight: 600; }
  .activity-amber { color: #856404; font-weight: 600; }
  .activity-red { color: #c0392b; font-weight: 600; }
  .activity-brightred { color: #e74c3c; font-weight: 700; }
  .activity-never { color: #6c757d; font-weight: 600; }

  .footer { text-align: center; font-size: 0.8em; color: #999; margin-top: 20px; }
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <h1>All Windows Devices Report</h1>
    <p>Tenant: $([System.Net.WebUtility]::HtmlEncode($tenantDisplayName)) ($tenantId) &nbsp;|&nbsp; Generated: $reportDate</p>
  </div>
  <div class="header-right">
    <label for="thresholdSelect">Inactive Threshold:</label>
    <select id="thresholdSelect" onchange="applyThreshold()">
      <option value="30" $(if ($InactiveDays -eq 30) { 'selected' })>30 Days</option>
      <option value="60" $(if ($InactiveDays -eq 60) { 'selected' })>60 Days</option>
      <option value="90" $(if ($InactiveDays -eq 90) { 'selected' })>90 Days</option>
      <option value="180" $(if ($InactiveDays -eq 180 -or ($InactiveDays -ne 30 -and $InactiveDays -ne 60 -and $InactiveDays -ne 90 -and $InactiveDays -ne 360)) { 'selected' })>180 Days</option>
      <option value="360" $(if ($InactiveDays -eq 360) { 'selected' })>360 Days</option>
    </select>
  </div>
</div>

<!-- OVERVIEW -->
<div class="section-title">Overview</div>
<div class="summary-cards">
  <div class="card"><div class="label">Total Windows Devices</div><div class="value" style="color:#1a1a2e;" id="cardTotal">$totalDevices</div></div>
  <div class="card"><div class="label">Active</div><div class="value" style="color:#27ae60;" id="cardActive">-</div><div class="sub" id="cardActivePct"></div></div>
  <div class="card"><div class="label">Inactive</div><div class="value" style="color:#e74c3c;" id="cardInactive">-</div><div class="sub" id="cardInactivePct"></div></div>
  <div class="card"><div class="label">Disabled</div><div class="value" style="color:#6c757d;" id="cardDisabled">-</div><div class="sub" id="cardDisabledPct"></div></div>
</div>

<!-- ACTIVE DEVICES: MANAGEMENT TYPE -->
<div class="dist-section">
  <div class="section-title">Active Devices — Management Type</div>
  <div class="dist-cards" id="activeMgmtCards"></div>
</div>

<!-- ACTIVE DEVICES: TRUST TYPE -->
<div class="dist-section">
  <div class="section-title">Active Devices — Trust Type (Join Type)</div>
  <div class="dist-cards" id="activeTrustCards"></div>
</div>

<!-- CHARTS (active devices only) -->
<div class="charts-row">
  <div class="chart-section">
    <h2>Active Devices by Management Type</h2>
    <div class="chart-container"><canvas id="mgmtChart"></canvas></div>
  </div>
  <div class="chart-section">
    <h2>Active Devices by Trust Type</h2>
    <div class="chart-container"><canvas id="trustChart"></canvas></div>
  </div>
</div>

<!-- DEVICE TABLE (all devices) -->
<div class="table-section">
  <h2>Windows Device Details</h2>
  <div class="table-controls">
    <input type="text" id="searchBox" placeholder="Search by name, OS version, trust type..." onkeyup="filterTable()" />
    <select id="statusFilter" onchange="filterTable()">
      <option value="all">All Status</option>
      <option value="active">Active Only</option>
      <option value="inactive">Inactive Only</option>
      <option value="disabled">Disabled Only</option>
    </select>
    <span class="count-label" id="rowCount"></span>
  </div>
  <table id="deviceTable">
    <thead><tr>
      <th onclick="sortTable(0)">Name</th>
      <th onclick="sortTable(1)">Enabled</th>
      <th onclick="sortTable(2)">OS</th>
      <th onclick="sortTable(3)">Version</th>
      <th onclick="sortTable(4)">Join Type</th>
      <th onclick="sortTable(5)">Management Type</th>
      <th onclick="sortTable(6)">Compliant</th>
      <th onclick="sortTable(7)">Registered</th>
      <th onclick="sortTable(8)">Last Activity</th>
      <th onclick="sortTable(9)">Enrollment to Activity</th>
      <th onclick="sortTable(10)">Days Since Activity</th>
      <th onclick="sortTable(11)">Status</th>
    </tr></thead>
    <tbody>
$tableRows
    </tbody>
  </table>
</div>

<div class="footer">Report generated by ReportAllWindowsDevices.ps1</div>

<script>
var deviceData = [$devicesJson];
var chartColors = ['#3498db','#27ae60','#e74c3c','#f39c12','#9b59b6','#1abc9c','#e67e22','#2c3e50','#95a5a6','#d35400'];

var chartOpts = function(pos) {
  return { responsive: true, plugins: { legend: { position: pos || 'right', labels: { padding: 16, font: { size: 13 }, boxWidth: 16 } }, tooltip: { callbacks: { label: function(ctx) { var t = ctx.dataset.data.reduce(function(a,b){return a+b},0); return ctx.label+': '+ctx.parsed+' ('+(t>0?((ctx.parsed/t)*100).toFixed(1):0)+'%)'; } } } } };
};

var mgmtChart = new Chart(document.getElementById('mgmtChart'), { type:'doughnut', data:{ labels:[], datasets:[{ data:[], backgroundColor:[], borderWidth:2, borderColor:'#fff' }] }, options: chartOpts() });
var trustChart = new Chart(document.getElementById('trustChart'), { type:'doughnut', data:{ labels:[], datasets:[{ data:[], backgroundColor:[], borderWidth:2, borderColor:'#fff' }] }, options: chartOpts() });

function pct(n, total) { return total > 0 ? ((n / total) * 100).toFixed(1) : '0.0'; }

function buildDistCards(containerId, distMap) {
  var container = document.getElementById(containerId);
  container.innerHTML = '';
  var keys = Object.keys(distMap).sort(function(a,b){ return distMap[b] - distMap[a]; });
  keys.forEach(function(key) {
    var card = document.createElement('div');
    card.className = 'dist-card';
    card.innerHTML = '<div class="dist-label">' + key + '</div><div class="dist-value">' + distMap[key] + '</div>';
    container.appendChild(card);
  });
}

function updateChart(chart, distMap) {
  var keys = Object.keys(distMap).sort(function(a,b){ return distMap[b] - distMap[a]; });
  chart.data.labels = keys;
  chart.data.datasets[0].data = keys.map(function(k){ return distMap[k]; });
  chart.data.datasets[0].backgroundColor = keys.map(function(_,i){ return chartColors[i % chartColors.length]; });
  chart.update();
}

function applyThreshold() {
  var threshold = parseInt(document.getElementById('thresholdSelect').value);
  var total = deviceData.length;
  var active = 0, inactive = 0, disabled = 0;
  var activeMgmt = {}, activeTrust = {};

  for (var i = 0; i < deviceData.length; i++) {
    var d = deviceData[i];
    if (!d.ena) { disabled++; continue; }
    var isInactive = (d.days === -1) || (d.days >= threshold);
    if (isInactive) {
      inactive++;
    } else {
      active++;
      activeMgmt[d.mt] = (activeMgmt[d.mt] || 0) + 1;
      activeTrust[d.tt] = (activeTrust[d.tt] || 0) + 1;
    }
  }

  var enabled = active + inactive;

  document.getElementById('cardTotal').textContent = total;
  document.getElementById('cardActive').textContent = active;
  document.getElementById('cardActivePct').textContent = pct(active, enabled) + '% of enabled';
  document.getElementById('cardInactive').textContent = inactive;
  document.getElementById('cardInactivePct').textContent = pct(inactive, enabled) + '% of enabled';
  document.getElementById('cardDisabled').textContent = disabled;
  document.getElementById('cardDisabledPct').textContent = pct(disabled, total) + '% of total';

  // Update active device analysis cards
  buildDistCards('activeMgmtCards', activeMgmt);
  buildDistCards('activeTrustCards', activeTrust);

  // Update charts
  updateChart(mgmtChart, activeMgmt);
  updateChart(trustChart, activeTrust);

  // Update table row statuses and color-code Days Since Activity
  var rows = document.querySelectorAll('#deviceTable tbody tr');
  for (var j = 0; j < rows.length; j++) {
    var days = parseInt(rows[j].getAttribute('data-days'));
    var rowEna = rows[j].getAttribute('data-ena');
    var badge = rows[j].cells[11].querySelector('.badge');
    if (rowEna === '0') {
      badge.className = 'badge badge-disabled';
      badge.textContent = 'Disabled';
      rows[j].setAttribute('data-status', 'disabled');
    } else {
      var rowInactive = (days === -1) || (days >= threshold);
      if (rowInactive) {
        badge.className = 'badge badge-inactive';
        badge.textContent = 'Inactive';
        rows[j].setAttribute('data-status', 'inactive');
      } else {
        badge.className = 'badge badge-active';
        badge.textContent = 'Active';
        rows[j].setAttribute('data-status', 'active');
      }
    }
    // Color-code Days Since Activity cell (index 10)
    var activityCell = rows[j].cells[10];
    if (days === -1) {
      activityCell.className = 'activity-never';
    } else if (days < 30) {
      activityCell.className = 'activity-green';
    } else if (days < 60) {
      activityCell.className = 'activity-amber';
    } else if (days < 180) {
      activityCell.className = 'activity-red';
    } else {
      activityCell.className = 'activity-brightred';
    }
  }

  filterTable();
}

function filterTable() {
  var search = document.getElementById('searchBox').value.toLowerCase();
  var status = document.getElementById('statusFilter').value;
  var rows = document.querySelectorAll('#deviceTable tbody tr');
  var visible = 0;
  rows.forEach(function(row) {
    var text = row.textContent.toLowerCase();
    var matchSearch = !search || text.indexOf(search) !== -1;
    var rowStatus = row.getAttribute('data-status');
    var matchStatus = status === 'all' || rowStatus === status;
    if (matchSearch && matchStatus) {
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
    if (!isNaN(an) && !isNaN(bn)) {
      return dir === 'asc' ? an - bn : bn - an;
    }
    if (av < bv) return dir === 'asc' ? -1 : 1;
    if (av > bv) return dir === 'asc' ? 1 : -1;
    return 0;
  });
  rows.forEach(function(row) { tbody.appendChild(row); });
}

// Initial render with default threshold
applyThreshold();
</script>
</body>
</html>
"@

	$htmlReportFile = Join-Path $reportFolder ("AllWindowsDevices_{0}.html" -f $timestamp)
	$html | Out-File -FilePath $htmlReportFile -Encoding UTF8

	# --- Console summary ---
	Write-Host ""
	Write-Host "All Windows Devices Report" -ForegroundColor Cyan
	Write-Host "--------------------------------------------"
	Write-Host ("Tenant                   : {0} ({1})" -f $tenantDisplayName, $tenantId)
	Write-Host ("Total Windows devices    : {0}" -f $totalDevices)
	Write-Host ("Active                   : {0}  ({1}% of enabled)" -f $totalActive, $percentActive) -ForegroundColor Green
	Write-Host ("Inactive                 : {0}  ({1}% of enabled)" -f $totalInactive, $percentInactive) -ForegroundColor Red
	Write-Host ("Disabled                 : {0}" -f $totalDisabled) -ForegroundColor DarkGray
	Write-Host ("Compliant                : {0}" -f $totalCompliant) -ForegroundColor Green
	Write-Host ""
	Write-Host "Trust Types" -ForegroundColor Cyan
	foreach ($tt in $trustSummary) {
		Write-Host ("  {0,-25}: {1}" -f $tt.TrustType, $tt.Count)
	}
	Write-Host ""
	Write-Host "Management Types" -ForegroundColor Cyan
	foreach ($mt in $mgmtSummary) {
		Write-Host ("  {0,-25}: {1}" -f $mt.ManagementType, $mt.Count)
	}
	Write-Host ""
	Write-Host ("Inactive days threshold  : {0}" -f $InactiveDays)
	Write-Host ("Devices in export        : {0}" -f $totalDevices)
	Write-Host ("CSV report               : {0}" -f $csvFile) -ForegroundColor Yellow
	Write-Host ("HTML report              : {0}" -f $htmlReportFile) -ForegroundColor Yellow

	$disconnectChoice = Read-Host "Disconnect from Microsoft Graph? (Y/N)"
	if ($disconnectChoice -match '^(y|yes)$') {
		Disconnect-MgGraph -ErrorAction SilentlyContinue
	}
}
catch {
	Write-Error $_
	exit 1
}
