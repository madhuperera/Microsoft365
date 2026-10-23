#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Reports on mailbox quota usage for all user mailboxes in Exchange Online.

.DESCRIPTION
    Connects to Exchange Online and retrieves mailbox quota statistics for all user mailboxes.
    Calculates the percentage of quota used and flags mailboxes that exceed the warning threshold.
    Exports results to CSV and HTML.

.PARAMETER Threshold
    Percentage of quota usage to use as the warning level. Defaults to 85.

.PARAMETER ReportPath
    Folder or file path for the output report. If a folder is specified, a timestamped
    filename is generated automatically. Defaults to the current directory.

.PARAMETER TestMode
    When specified, processes a random sample of 10 mailboxes instead of all mailboxes.
    Use this to validate the report format before running against the full tenant.

.EXAMPLE
    .\ReportMailboxQuota.ps1

.EXAMPLE
    .\ReportMailboxQuota.ps1 -Threshold 90

.EXAMPLE
    .\ReportMailboxQuota.ps1 -TestMode

.EXAMPLE
    .\ReportMailboxQuota.ps1 -ReportPath "C:\Reports"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 100)]
    [int]$Threshold = 85,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ReportPath,

    [Parameter(Mandatory = $false)]
    [switch]$TestMode
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable))
{
    throw "ExchangeOnlineManagement module is not installed. Install it using: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}

# --- Connect to Exchange Online ---
$S_ExistingSession = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
if ($S_ExistingSession)
{
    Write-Host "Existing Exchange Online session detected:" -ForegroundColor Yellow
    Write-Host "  Account   : $($S_ExistingSession.UserPrincipalName)" -ForegroundColor Yellow
    Write-Host "  TenantId  : $($S_ExistingSession.TenantID)" -ForegroundColor Yellow
    Write-Host ""
    $S_Choice = Read-Host "Use existing session? [Y] Yes  [N] Disconnect and reconnect  (Default: Y)"
    if ($S_Choice -eq 'N')
    {
        Disconnect-ExchangeOnline -Confirm:$false
        Connect-ExchangeOnline -ShowBanner:$false
    }
}
else
{
    Connect-ExchangeOnline -ShowBanner:$false
}

# --- Tenant info ---
$S_TenantDisplayName = $null
$S_TenantId = $null
try
{
    $S_ConnInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
    if ($S_ConnInfo)
    {
        $S_TenantDisplayName = $S_ConnInfo.Organization
        $S_TenantId = $S_ConnInfo.TenantID
    }
}
catch
{
}
if (-not $S_TenantDisplayName)
{
    $S_TenantDisplayName = 'Exchange Online'
}
if (-not $S_TenantId)
{
    $S_TenantId = 'Unknown'
}

# --- Retrieve mailboxes ---
if ($TestMode)
{
    Write-Host "TEST MODE: Retrieving random sample of 10 mailboxes..." -ForegroundColor Cyan
    [array]$S_Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -PropertySet Quota -Properties DisplayName -ResultSize 200 |
        Get-Random -Count 10
}
else
{
    Write-Host "Retrieving all user mailboxes..." -ForegroundColor Cyan
    [array]$S_Mbx = Get-ExoMailbox -RecipientTypeDetails UserMailbox -PropertySet Quota -Properties DisplayName -ResultSize Unlimited
}
Write-Host "  Found $($S_Mbx.Count) mailboxes to process" -ForegroundColor Green

# --- Build report data ---
$S_Report = [System.Collections.Generic.List[PSCustomObject]]::new()
$S_Counter = 0
foreach ($S_MailboxEntry in $S_Mbx)
{
    $S_Counter++
    Write-Host "[$S_Counter/$($S_Mbx.Count)] Processing $($S_MailboxEntry.DisplayName)..." -ForegroundColor Gray

    try
    {
        $S_MbxStats = Get-ExoMailboxStatistics -Identity $S_MailboxEntry.UserPrincipalName -ErrorAction Stop |
            Select-Object ItemCount, TotalItemSize

        # Byte count of quota used
        [int64]$S_QuotaUsed = [convert]::ToInt64(
            ((($S_MbxStats.TotalItemSize.ToString().split('(')[1]).split(')')[0]).split(' ')[0] -replace '[,]', '')
        )

        # Byte count for mailbox quota limit
        [int64]$S_MbxQuota = [convert]::ToInt64(
            ((($S_MailboxEntry.ProhibitSendReceiveQuota.ToString().split('(')[1]).split(')')[0]).split(' ')[0] -replace '[,]', '')
        )

        $S_MbxQuotaGB      = [math]::Round($S_MbxQuota / 1GB, 2)
        $S_QuotaUsedGB     = [math]::Round($S_QuotaUsed / 1GB, 2)
        $S_PercentRaw      = [math]::Round(($S_QuotaUsed / $S_MbxQuota) * 100, 1)
        $S_PercentDisplay  = "$S_PercentRaw%"

        if ($S_PercentRaw -ge $Threshold)
        {
            $S_Status = 'Over Threshold'
            Write-Host "  WARNING: $($S_MailboxEntry.DisplayName) is at $S_PercentDisplay" -ForegroundColor Red
        }
        elseif ($S_PercentRaw -ge ($Threshold - 10))
        {
            $S_Status = 'Approaching'
        }
        else
        {
            $S_Status = 'Healthy'
        }

        $S_ReportLine = [PSCustomObject]@{
            Mailbox        = $S_MailboxEntry.DisplayName
            UPN            = $S_MailboxEntry.UserPrincipalName
            QuotaGB        = $S_MbxQuotaGB
            UsedGB         = $S_QuotaUsedGB
            PercentUsed    = $S_PercentRaw
            PercentDisplay = $S_PercentDisplay
            ItemCount      = $S_MbxStats.ItemCount
            Status         = $S_Status
            Error          = $null
        }
    }
    catch
    {
        Write-Warning "  Could not retrieve stats for $($S_MailboxEntry.DisplayName): $($_.Exception.Message)"
        $S_ReportLine = [PSCustomObject]@{
            Mailbox        = $S_MailboxEntry.DisplayName
            UPN            = $S_MailboxEntry.UserPrincipalName
            QuotaGB        = $null
            UsedGB         = $null
            PercentUsed    = $null
            PercentDisplay = 'Error'
            ItemCount      = $null
            Status         = 'Error'
            Error          = $_.Exception.Message
        }
    }

    $S_Report.Add($S_ReportLine)
}

# --- Stats ---
$S_TotalMailboxes   = $S_Report.Count
$S_TotalHealthy     = @($S_Report | Where-Object { $_.Status -eq 'Healthy' }).Count
$S_TotalApproaching = @($S_Report | Where-Object { $_.Status -eq 'Approaching' }).Count
$S_TotalOver        = @($S_Report | Where-Object { $_.Status -eq 'Over Threshold' }).Count
$S_TotalError       = @($S_Report | Where-Object { $_.Status -eq 'Error' }).Count
$S_ValidEntries     = @($S_Report | Where-Object { $null -ne $_.PercentUsed })
$S_AvgPercent       = if ($S_ValidEntries.Count -gt 0)
{
    [math]::Round(($S_ValidEntries | Measure-Object -Property PercentUsed -Average).Average, 1)
}
else { 0 }

# --- File paths ---
if (-not $ReportPath)
{
    $ReportPath = (Get-Location).Path
}
if (Test-Path $ReportPath -PathType Container)
{
    $S_ReportFolder = $ReportPath
}
else
{
    $S_ReportFolder = Split-Path -Parent $ReportPath
}
if ($S_ReportFolder -and -not (Test-Path $S_ReportFolder))
{
    New-Item -ItemType Directory -Path $S_ReportFolder -Force | Out-Null
}

$S_Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$S_FileBase  = if ($TestMode) { "ReportMailboxQuota_TEST_$S_Timestamp" } else { "ReportMailboxQuota_$S_Timestamp" }

if (Test-Path $ReportPath -PathType Container)
{
    $S_CsvFile  = Join-Path $ReportPath "$S_FileBase.csv"
    $S_HtmlFile = Join-Path $ReportPath "$S_FileBase.html"
}
else
{
    $S_CsvFile  = $ReportPath
    $S_HtmlFile = [System.IO.Path]::ChangeExtension($ReportPath, '.html')
}

# --- CSV export ---
$S_Report | Sort-Object Mailbox | Select-Object Mailbox, UPN, QuotaGB, UsedGB, PercentDisplay, ItemCount, Status, Error |
    Export-Csv -Path $S_CsvFile -NoTypeInformation -Encoding UTF8

# --- HTML report ---
$S_ReportDate = Get-Date -Format 'dd MMM yyyy HH:mm'

$S_TestModeBanner = if ($TestMode)
{
    '<div style="background:#fff3cd;border:1px solid #ffc107;border-radius:8px;padding:12px 20px;margin-bottom:20px;color:#856404;font-size:0.9em;"><strong>Test Mode:</strong> This report was generated from a random sample of 10 mailboxes. Run without <code>-TestMode</code> to report on all mailboxes.</div>'
}
else { '' }

# Build table rows
$S_TableRows = ($S_Report | Sort-Object Mailbox | ForEach-Object {
    $S_Name   = [System.Net.WebUtility]::HtmlEncode($_.Mailbox)
    $S_Upn    = [System.Net.WebUtility]::HtmlEncode($_.UPN)
    $S_Quota  = if ($null -ne $_.QuotaGB) { "$($_.QuotaGB) GB" } else { '-' }
    $S_Used   = if ($null -ne $_.UsedGB) { "$($_.UsedGB) GB" } else { '-' }
    $S_Pct    = $_.PercentDisplay
    $S_Items  = if ($null -ne $_.ItemCount) { $_.ItemCount.ToString('N0') } else { '-' }
    $S_PctRaw = if ($null -ne $_.PercentUsed) { $_.PercentUsed } else { 0 }

    $S_StatusClass = switch ($_.Status)
    {
        'Healthy'        { 'healthy' }
        'Approaching'    { 'approaching' }
        'Over Threshold' { 'overthreshold' }
        'Error'          { 'error' }
        default          { 'healthy' }
    }

    $S_BarColour = switch ($_.Status)
    {
        'Healthy'        { '#27ae60' }
        'Approaching'    { '#f39c12' }
        'Over Threshold' { '#e74c3c' }
        'Error'          { '#bbb' }
        default          { '#27ae60' }
    }

    $S_ProgressBar = if ($null -ne $_.PercentUsed)
    {
        "<div class=`"progress-wrap`"><div class=`"progress-bar`" style=`"width:$($S_PctRaw)%;background:$S_BarColour;`"></div></div>"
    }
    else { '-' }

    "<tr data-pct=`"$S_PctRaw`" data-status=`"$($_.Status)`"><td>$S_Name</td><td class=`"upn`">$S_Upn</td><td>$S_Quota</td><td>$S_Used</td><td>$S_Pct $S_ProgressBar</td><td>$S_Items</td><td><span class=`"badge badge-$S_StatusClass`">$($_.Status)</span></td></tr>"
}) -join "`n"

$S_Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Mailbox Quota Report</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; padding: 30px; }

  .header { background: linear-gradient(135deg, #1a1a2e, #16213e); color: #fff; padding: 30px 40px; border-radius: 12px; margin-bottom: 30px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 16px; }
  .header-left h1 { font-size: 1.6em; margin-bottom: 6px; }
  .header-left p { font-size: 0.9em; opacity: 0.8; }

  .section-title { font-size: 1.15em; font-weight: 600; color: #1a1a2e; margin-bottom: 16px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0; }

  .summary-cards { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .card { background: #fff; border-radius: 10px; padding: 24px 30px; flex: 1; min-width: 150px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
  .card .label { font-size: 0.85em; color: #777; text-transform: uppercase; letter-spacing: 0.5px; }
  .card .value { font-size: 2em; font-weight: 700; margin-top: 6px; }
  .card .sub { font-size: 0.78em; color: #999; margin-top: 2px; }

  .charts-row { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .chart-section { background: #fff; border-radius: 10px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); flex: 1; min-width: 300px; }
  .chart-section h2 { font-size: 1.1em; margin-bottom: 20px; color: #1a1a2e; }
  .chart-container { max-width: 380px; margin: 0 auto; }

  .table-section { background: #fff; border-radius: 10px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 30px; overflow-x: auto; }
  .table-section h2 { font-size: 1.1em; margin-bottom: 16px; color: #1a1a2e; }
  .table-controls { display: flex; gap: 14px; margin-bottom: 16px; flex-wrap: wrap; align-items: center; }
  .table-controls input[type="text"] { padding: 8px 14px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; min-width: 260px; }
  .table-controls select { padding: 8px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; background: #fff; }
  .table-controls .count-label { font-size: 0.85em; color: #777; margin-left: auto; }

  table { width: 100%; border-collapse: collapse; font-size: 0.84em; }
  th { background: #1a1a2e; color: #fff; padding: 12px 14px; text-align: left; position: sticky; top: 0; cursor: pointer; user-select: none; white-space: nowrap; }
  th:hover { background: #2c3e50; }
  td { padding: 10px 14px; border-bottom: 1px solid #eee; white-space: nowrap; }
  tr:hover td { background: #f8f9fa; }
  tr.hidden-row { display: none; }
  .upn { font-family: 'Consolas', monospace; font-size: 0.82em; color: #666; }

  .progress-wrap { width: 100px; height: 6px; background: #eee; border-radius: 3px; display: inline-block; vertical-align: middle; margin-left: 8px; }
  .progress-bar { height: 100%; border-radius: 3px; }

  .badge { padding: 3px 10px; border-radius: 12px; font-size: 0.82em; font-weight: 600; }
  .badge-healthy       { background: #d4edda; color: #155724; }
  .badge-approaching   { background: #fff3cd; color: #856404; }
  .badge-overthreshold { background: #f8d7da; color: #721c24; }
  .badge-error         { background: #e2e3e5; color: #495057; }

  .footer { text-align: center; font-size: 0.8em; color: #999; margin-top: 20px; }
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <h1>Mailbox Quota Report</h1>
    <p>Tenant: $([System.Net.WebUtility]::HtmlEncode($S_TenantDisplayName)) ($S_TenantId) &nbsp;|&nbsp; Generated: $S_ReportDate &nbsp;|&nbsp; Warning Threshold: $Threshold%</p>
  </div>
</div>

$S_TestModeBanner

<!-- OVERVIEW -->
<div class="section-title">Overview</div>
<div class="summary-cards">
  <div class="card"><div class="label">Total Mailboxes</div><div class="value" style="color:#1a1a2e;">$S_TotalMailboxes</div></div>
  <div class="card"><div class="label">Healthy</div><div class="value" style="color:#27ae60;">$S_TotalHealthy</div><div class="sub">Below $([math]::Max(1, $Threshold - 10))% used</div></div>
  <div class="card"><div class="label">Approaching Limit</div><div class="value" style="color:#f39c12;">$S_TotalApproaching</div><div class="sub">$([math]::Max(1, $Threshold - 10))% – $Threshold% used</div></div>
  <div class="card"><div class="label">Over Threshold</div><div class="value" style="color:#e74c3c;">$S_TotalOver</div><div class="sub">Above $Threshold% used</div></div>
  <div class="card"><div class="label">Average Usage</div><div class="value" style="color:#3498db;">$S_AvgPercent%</div><div class="sub">Across all mailboxes</div></div>
  $(if ($S_TotalError -gt 0) { "<div class=`"card`"><div class=`"label`">Errors</div><div class=`"value`" style=`"color:#6c757d;`">$S_TotalError</div><div class=`"sub`">Could not retrieve stats</div></div>" })
</div>

<!-- CHARTS -->
<div class="charts-row">
  <div class="chart-section">
    <h2>Quota Status Distribution</h2>
    <div class="chart-container"><canvas id="statusChart"></canvas></div>
  </div>
  <div class="chart-section">
    <h2>Usage Distribution</h2>
    <div class="chart-container"><canvas id="usageChart"></canvas></div>
  </div>
</div>

<!-- FULL TABLE -->
<div class="table-section">
  <h2>Mailbox Quota Details</h2>
  <div class="table-controls">
    <input type="text" id="searchBox" placeholder="Search by name or UPN..." onkeyup="filterTable()" />
    <select id="statusFilter" onchange="filterTable()">
      <option value="all">All Status</option>
      <option value="Healthy">Healthy</option>
      <option value="Approaching">Approaching</option>
      <option value="Over Threshold">Over Threshold</option>
      $(if ($S_TotalError -gt 0) { '<option value="Error">Error</option>' })
    </select>
    <span class="count-label" id="countLabel">Showing $S_TotalMailboxes of $S_TotalMailboxes mailboxes</span>
  </div>
  <table id="mainTable">
    <thead>
      <tr>
        <th onclick="sortTable(0)">Mailbox &#x25B2;&#x25BC;</th>
        <th onclick="sortTable(1)">UPN &#x25B2;&#x25BC;</th>
        <th onclick="sortTable(2)">Quota (GB) &#x25B2;&#x25BC;</th>
        <th onclick="sortTable(3)">Used (GB) &#x25B2;&#x25BC;</th>
        <th onclick="sortTable(4)">% Used &#x25B2;&#x25BC;</th>
        <th onclick="sortTable(5)">Items &#x25B2;&#x25BC;</th>
        <th>Status</th>
      </tr>
    </thead>
    <tbody id="tableBody">
$S_TableRows
    </tbody>
  </table>
</div>

<div class="footer">Generated by ReportMailboxQuota.ps1 &nbsp;|&nbsp; $S_ReportDate</div>

<script>
// --- Chart: Status Distribution ---
(function() {
  var ctx = document.getElementById('statusChart').getContext('2d');
  new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: ['Healthy', 'Approaching', 'Over Threshold'$(if ($S_TotalError -gt 0) { ", 'Error'" })],
      datasets: [{
        data: [$S_TotalHealthy, $S_TotalApproaching, $S_TotalOver$(if ($S_TotalError -gt 0) { ", $S_TotalError" })],
        backgroundColor: ['#27ae60', '#f39c12', '#e74c3c'$(if ($S_TotalError -gt 0) { ", '#bbb'" })],
        borderWidth: 2,
        borderColor: '#fff'
      }]
    },
    options: { plugins: { legend: { position: 'bottom' } }, cutout: '60%' }
  });
})();

// --- Chart: Usage Distribution histogram ---
(function() {
  var rows = document.querySelectorAll('#tableBody tr');
  var buckets = [0,0,0,0,0]; // 0-20, 20-40, 40-60, 60-80, 80-100
  rows.forEach(function(r) {
    var pct = parseFloat(r.getAttribute('data-pct'));
    if (isNaN(pct)) return;
    if (pct < 20) buckets[0]++;
    else if (pct < 40) buckets[1]++;
    else if (pct < 60) buckets[2]++;
    else if (pct < 80) buckets[3]++;
    else buckets[4]++;
  });
  var ctx = document.getElementById('usageChart').getContext('2d');
  new Chart(ctx, {
    type: 'bar',
    data: {
      labels: ['0-20%', '20-40%', '40-60%', '60-80%', '80-100%'],
      datasets: [{
        label: 'Mailboxes',
        data: buckets,
        backgroundColor: ['#27ae60','#2ecc71','#f39c12','#e67e22','#e74c3c'],
        borderRadius: 4
      }]
    },
    options: {
      plugins: { legend: { display: false } },
      scales: { y: { beginAtZero: true, ticks: { precision: 0 } } }
    }
  });
})();

// --- Filter & search ---
function filterTable() {
  var search = document.getElementById('searchBox').value.toLowerCase();
  var status = document.getElementById('statusFilter').value;
  var rows = document.querySelectorAll('#tableBody tr');
  var visible = 0;
  rows.forEach(function(r) {
    var text = r.textContent.toLowerCase();
    var rowStatus = r.getAttribute('data-status') || '';
    var matchSearch = !search || text.indexOf(search) > -1;
    var matchStatus = status === 'all' || rowStatus === status;
    if (matchSearch && matchStatus) {
      r.classList.remove('hidden-row');
      visible++;
    } else {
      r.classList.add('hidden-row');
    }
  });
  document.getElementById('countLabel').textContent = 'Showing ' + visible + ' of ' + rows.length + ' mailboxes';
}

// --- Sort ---
var sortDir = {};
function sortTable(col) {
  var tbody = document.getElementById('tableBody');
  var rows = Array.from(tbody.querySelectorAll('tr'));
  var dir = sortDir[col] === 'asc' ? 'desc' : 'asc';
  sortDir[col] = dir;
  rows.sort(function(a, b) {
    var ta = a.cells[col] ? a.cells[col].textContent.trim() : '';
    var tb = b.cells[col] ? b.cells[col].textContent.trim() : '';
    var na = parseFloat(ta); var nb = parseFloat(tb);
    if (!isNaN(na) && !isNaN(nb)) { return dir === 'asc' ? na - nb : nb - na; }
    return dir === 'asc' ? ta.localeCompare(tb) : tb.localeCompare(ta);
  });
  rows.forEach(function(r) { tbody.appendChild(r); });
}
</script>
</body>
</html>
"@

$S_Html | Out-File -FilePath $S_HtmlFile -Encoding UTF8

Write-Host ""
Write-Host "Report complete." -ForegroundColor Green
Write-Host "  CSV  : $S_CsvFile"
Write-Host "  HTML : $S_HtmlFile"
Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "  Total mailboxes   : $S_TotalMailboxes"
Write-Host "  Healthy           : $S_TotalHealthy"
Write-Host "  Approaching limit : $S_TotalApproaching"
Write-Host "  Over threshold    : $S_TotalOver" -ForegroundColor $(if ($S_TotalOver -gt 0) { 'Red' } else { 'White' })
if ($S_TotalError -gt 0) { Write-Host "  Errors            : $S_TotalError" -ForegroundColor Yellow }
Write-Host "  Average usage     : $S_AvgPercent%"

