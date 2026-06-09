#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
	Reports on Intune-managed Android and iOS/iPadOS devices and highlights devices
	running OS versions older than the latest supported major version.

.DESCRIPTION
	Connects to Microsoft Graph and retrieves all Intune-managed mobile devices
	(Android, iOS and iPadOS). For each device the script captures user, model,
	OS major version, last sync date, compliance state and whether the major OS
	version is older than the supplied latest-supported major versions for
	Android and iOS/iPadOS. Outputs a CSV file and an HTML report with version
	spread, charts and a sortable / filterable device table.

.PARAMETER LatestSupportedAndroid
	The latest supported Android major version (integer, e.g. 15). Devices with a
	major Android OS version less than this value are flagged as "Outdated".

.PARAMETER LatestSupportedIOS
	The latest supported iOS / iPadOS major version (integer, e.g. 18). Devices
	with a major iOS/iPadOS OS version less than this value are flagged as "Outdated".

.PARAMETER ReportPath
	Folder for the output reports. If omitted the current working directory is used.

.EXAMPLE
	.\ReportIntuneMobileDevices.ps1 -LatestSupportedAndroid 15 -LatestSupportedIOS 18

.EXAMPLE
	.\ReportIntuneMobileDevices.ps1 -LatestSupportedAndroid 14 -LatestSupportedIOS 17 -ReportPath C:\Reports
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateRange(1, 99)]
    [int]$LatestSupportedAndroid,

    [Parameter(Mandatory = $true)]
    [ValidateRange(1, 99)]
    [int]$LatestSupportedIOS,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ReportPath
)

$ErrorActionPreference = "Stop"

$S_ReportPath = $ReportPath

$S_RequiredGraphScopes = @(
    'DeviceManagementManagedDevices.Read.All'
    'Organization.Read.All'
)

try
{
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication))
    {
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
    if ([string]::IsNullOrWhiteSpace($S_ContextConfirmation))
    {
        $S_ContextConfirmation = 'N'
    }
    if ($S_ContextConfirmation.ToUpperInvariant() -ne 'Y')
    {
        throw "Operation cancelled. Please reconnect to the correct tenant and account, then run again."
    }

    # --- Tenant info ---
    $S_TenantDisplayName = $null
    try
    {
        $S_OrgResp = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization' -ErrorAction Stop
        if ($S_OrgResp.value)
        {
            $S_TenantDisplayName = $S_OrgResp.value[0].displayName
        }
    }
    catch
    {
    }
    if (-not $S_TenantDisplayName)
    {
        $S_TenantDisplayName = $S_ExistingContext.TenantId
    }
    $S_TenantId = if ($S_ExistingContext.TenantId)
    {
        $S_ExistingContext.TenantId
    }
    else
    {
        'Unknown'
    }

    # --- Fetch managed devices (Android, iOS, iPadOS) ---
    Write-Host "Fetching Intune-managed Android and iOS/iPadOS devices..." -ForegroundColor Cyan
    $S_Select = 'id,deviceName,userPrincipalName,userDisplayName,operatingSystem,osVersion,model,manufacturer,deviceType,enrolledDateTime,lastSyncDateTime,complianceState,managedDeviceOwnerType,joinType,serialNumber'
    $S_Filter = "(operatingSystem eq 'Android' or operatingSystem eq 'iOS' or operatingSystem eq 'iPadOS')"
    $S_EncodedFilter = [System.Uri]::EscapeDataString($S_Filter)
    $S_Uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=$S_EncodedFilter&`$select=$S_Select&`$top=200"

    $S_Devices = New-Object System.Collections.Generic.List[object]
    do
    {
        $S_Resp = Invoke-MgGraphRequest -Method GET -Uri $S_Uri -ErrorAction Stop
        if ($S_Resp.value)
        {
            foreach ($d in $S_Resp.value)
            {
                $S_Devices.Add([pscustomobject]$d) | Out-Null
            }
        }
        $S_Uri = $S_Resp.'@odata.nextLink'
    } while ($S_Uri)

    Write-Host ("  Retrieved {0} mobile devices" -f $S_Devices.Count) -ForegroundColor Green

    # --- Build report rows ---
    $S_Now = Get-Date
    $S_Report = foreach ($d in $S_Devices)
    {
        $S_Os = $d.operatingSystem
        $S_RawVer = if ($d.osVersion)
        {
            [string]$d.osVersion
        }
        else
        {
            ''
        }

        # Treat iOS devices with iPad models as iPadOS for clarity
        $S_Platform = $S_Os
        if ($S_Os -eq 'iOS' -and ($d.model -match 'iPad' -or $d.deviceType -eq 'iPad'))
        {
            $S_Platform = 'iPadOS'
        }

        $S_MajorVer = $null
        if ($S_RawVer -match '^\s*(\d+)')
        {
            $S_MajorVer = [int]$Matches[1]
        }

        $S_Threshold = switch ($S_Platform)
        {
            'Android' { $LatestSupportedAndroid }
            'iOS' { $LatestSupportedIOS }
            'iPadOS' { $LatestSupportedIOS }
            default { $null }
        }

        $S_SupportStatus = 'Unknown'
        if ($null -ne $S_MajorVer -and $null -ne $S_Threshold)
        {
            if ($S_MajorVer -lt $S_Threshold)
            {
                $S_SupportStatus = 'Outdated'
            }
            else
            {
                $S_SupportStatus = 'Supported'
            }
        }

        $S_LastSync = if ($d.lastSyncDateTime)
        {
            [datetime]$d.lastSyncDateTime
        }
        else
        {
            $null
        }
        $S_DaysSinceSync = if ($S_LastSync)
        {
            [int]($S_Now - $S_LastSync).TotalDays
        }
        else
        {
            $null
        }

        [pscustomobject]@{
            DeviceName        = $d.deviceName
            User              = if ($d.userDisplayName)
            {
                $d.userDisplayName
            }
            else
            {
                $d.userPrincipalName
            }
            UserPrincipalName = $d.userPrincipalName
            Platform          = $S_Platform
            OperatingSystem   = $S_Os
            OSVersion         = $S_RawVer
            MajorVersion      = $S_MajorVer
            LatestSupported   = $S_Threshold
            SupportStatus     = $S_SupportStatus
            Manufacturer      = $d.manufacturer
            Model             = $d.model
            Ownership         = $d.managedDeviceOwnerType
            ComplianceState   = $d.complianceState
            EnrolledDateTime  = $d.enrolledDateTime
            LastSyncDateTime  = $d.lastSyncDateTime
            DaysSinceLastSync = $S_DaysSinceSync
            SerialNumber      = $d.serialNumber
        }
    }

    # --- Stats ---
    $S_TotalDevices = $S_Report.Count
    $S_AndroidDevices = $S_Report | Where-Object { $_.Platform -eq 'Android' }
    $S_IosDevices = $S_Report | Where-Object { $_.Platform -eq 'iOS' }
    $S_IpadDevices = $S_Report | Where-Object { $_.Platform -eq 'iPadOS' }

    $S_TotalAndroid = $S_AndroidDevices.Count
    $S_TotalIos = $S_IosDevices.Count
    $S_TotalIpad = $S_IpadDevices.Count

    $S_OutdatedAndroid = ($S_AndroidDevices | Where-Object { $_.SupportStatus -eq 'Outdated' }).Count
    $S_OutdatedIos = ($S_IosDevices     | Where-Object { $_.SupportStatus -eq 'Outdated' }).Count
    $S_OutdatedIpad = ($S_IpadDevices    | Where-Object { $_.SupportStatus -eq 'Outdated' }).Count
    $S_TotalOutdated = $S_OutdatedAndroid + $S_OutdatedIos + $S_OutdatedIpad

    function Get-VersionSpread
    {
        param([object[]]$Devices, [int]$Threshold)
        $grouped = $Devices | Group-Object MajorVersion | Sort-Object {
            if ([string]::IsNullOrEmpty($_.Name))
            {
                -1
            }
            else
            {
                [int]$_.Name
            }
        } -Descending
        foreach ($g in $grouped)
        {
            $ver = if ([string]::IsNullOrEmpty($g.Name))
            {
                'Unknown'
            }
            else
            {
                $g.Name
            }
            $verNum = if ($ver -eq 'Unknown')
            {
                $null
            }
            else
            {
                [int]$ver
            }
            $outdated = ($null -ne $verNum -and $verNum -lt $Threshold)
            [pscustomobject]@{
                Version  = $ver
                Count    = $g.Count
                Outdated = $outdated
            }
        }
    }

    $S_AndroidSpread = @(Get-VersionSpread -Devices $S_AndroidDevices -Threshold $LatestSupportedAndroid)
    $S_IosSpread = @(Get-VersionSpread -Devices $S_IosDevices     -Threshold $LatestSupportedIOS)
    $S_IpadSpread = @(Get-VersionSpread -Devices $S_IpadDevices    -Threshold $LatestSupportedIOS)

    # --- Output paths ---
    if (-not $S_ReportPath)
    {
        $S_ReportPath = (Get-Location).Path
    }
    $S_ReportFolder = if (Test-Path $S_ReportPath -PathType Container)
    {
        $S_ReportPath
    }
    else
    {
        Split-Path -Parent $S_ReportPath
    }
    if ($S_ReportFolder -and -not (Test-Path $S_ReportFolder))
    {
        New-Item -ItemType Directory -Path $S_ReportFolder -Force | Out-Null
    }
    $S_Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $S_CsvFile = Join-Path $S_ReportFolder ("ReportIntuneMobileDevices_{0}.csv" -f $S_Timestamp)
    $S_HtmlFile = Join-Path $S_ReportFolder ("ReportIntuneMobileDevices_{0}.html" -f $S_Timestamp)

    # --- CSV export ---
    $S_Report | Sort-Object Platform, DeviceName | Export-Csv -Path $S_CsvFile -NoTypeInformation -Encoding UTF8

    # --- HTML helpers ---
    $S_Enc = { param($s) if ($null -eq $s -or $s -eq '') { '-' } else { [System.Net.WebUtility]::HtmlEncode([string]$s) } }
    $S_ReportDate = Get-Date -Format "dd MMM yyyy HH:mm"

    # Build version spread cards HTML
    function Build-SpreadCardsHtml
    {
        param([object[]]$Spread, [string]$PlatformLabel)
        if (-not $Spread -or $Spread.Count -eq 0)
        {
            return "<div class='dist-card'><div class='dist-label'>No $PlatformLabel devices</div><div class='dist-value'>0</div></div>"
        }
        ($Spread | ForEach-Object {
            $cls = if ($_.Outdated)
            {
                'dist-card outdated'
            }
            else
            {
                'dist-card'
            }
            $badge = if ($_.Outdated)
            {
                "<div class='outdated-badge'>Outdated</div>"
            }
            else
            {
                ''
            }
            "<div class='$cls'><div class='dist-label'>$PlatformLabel $($_.Version)</div><div class='dist-value'>$($_.Count)</div>$badge</div>"
        }) -join "`n"
    }

    $S_AndroidCardsHtml = Build-SpreadCardsHtml -Spread $S_AndroidSpread -PlatformLabel 'Android'
    $S_IosCardsHtml = Build-SpreadCardsHtml -Spread $S_IosSpread     -PlatformLabel 'iOS'
    $S_IpadCardsHtml = Build-SpreadCardsHtml -Spread $S_IpadSpread    -PlatformLabel 'iPadOS'

    # Build chart data JSON
    function ConvertTo-ChartJson
    {
        param([object[]]$Spread)
        if (-not $Spread -or $Spread.Count -eq 0)
        {
            return '{"labels":[],"data":[],"outdated":[]}'
        }
        $labels = ($Spread | ForEach-Object { '"' + $_.Version + '"' }) -join ','
        $counts = ($Spread | ForEach-Object { $_.Count }) -join ','
        $outFlag = ($Spread | ForEach-Object { if ($_.Outdated) { 'true' } else { 'false' } }) -join ','
        "{`"labels`":[$labels],`"data`":[$counts],`"outdated`":[$outFlag]}"
    }

    $S_AndroidChartJson = ConvertTo-ChartJson -Spread $S_AndroidSpread
    $S_IosChartJson = ConvertTo-ChartJson -Spread $S_IosSpread
    $S_IpadChartJson = ConvertTo-ChartJson -Spread $S_IpadSpread

    # Build table rows
    $S_TableRows = ($S_Report | Sort-Object Platform, DeviceName | ForEach-Object {
            $S_DaysVal = if ($null -ne $_.DaysSinceLastSync)
            {
                $_.DaysSinceLastSync
            }
            else
            {
                -1
            }
            $S_LastSyncDisp = if ($_.LastSyncDateTime)
            {
                ([datetime]$_.LastSyncDateTime).ToString("dd MMM yyyy")
            }
            else
            {
                '-'
            }
            $S_EnrolDisp = if ($_.EnrolledDateTime)
            {
                ([datetime]$_.EnrolledDateTime).ToString("dd MMM yyyy")
            }
            else
            {
                '-'
            }
            $S_StatusClass = switch ($_.SupportStatus)
            {
                'Outdated' { 'badge-inactive' }
                'Supported' { 'badge-active' }
                default { 'badge-disabled' }
            }
            $S_CompClass = switch ($_.ComplianceState)
            {
                'compliant' { 'badge-active' }
                'noncompliant' { 'badge-inactive' }
                default { 'badge-disabled' }
            }
            $S_RowAttr = "data-platform=`"$($_.Platform)`" data-status=`"$($_.SupportStatus)`""
            "<tr $S_RowAttr>" +
            "<td>$(& $S_Enc $_.DeviceName)</td>" +
            "<td>$(& $S_Enc $_.User)</td>" +
            "<td>$(& $S_Enc $_.Platform)</td>" +
            "<td>$(& $S_Enc $_.OSVersion)</td>" +
            "<td>$(if ($null -ne $_.MajorVersion) { $_.MajorVersion } else { '-' })</td>" +
            "<td><span class='badge $S_StatusClass'>$($_.SupportStatus)</span></td>" +
            "<td>$(& $S_Enc $_.Manufacturer)</td>" +
            "<td>$(& $S_Enc $_.Model)</td>" +
            "<td>$(& $S_Enc $_.Ownership)</td>" +
            "<td><span class='badge $S_CompClass'>$(& $S_Enc $_.ComplianceState)</span></td>" +
            "<td>$S_EnrolDisp</td>" +
            "<td>$S_LastSyncDisp</td>" +
            "<td>$(if ($S_DaysVal -ge 0) { "$S_DaysVal days" } else { 'Never' })</td>" +
            "</tr>"
        }) -join "`n"

    $S_PctOutdated = if ($S_TotalDevices -gt 0)
    {
        [math]::Round(($S_TotalOutdated / $S_TotalDevices) * 100, 1)
    }
    else
    {
        0
    }

    # --- HTML report ---
    $S_Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Intune Mobile Devices Report</title>
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

  .badge { padding: 3px 10px; border-radius: 12px; font-size: 0.8em; font-weight: 600; }
  .badge-active { background: #d4edda; color: #155724; }
  .badge-inactive { background: #f8d7da; color: #721c24; }
  .badge-disabled { background: #e2e3e5; color: #495057; }

  .footer { text-align: center; font-size: 0.8em; color: #999; margin-top: 20px; }
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <h1>Intune Mobile Devices Report</h1>
    <p>Tenant: $(& $S_Enc $S_TenantDisplayName) ($S_TenantId) &nbsp;|&nbsp; Generated: $S_ReportDate</p>
  </div>
  <div class="header-right">
    Latest supported Android: <strong>$LatestSupportedAndroid</strong><br/>
    Latest supported iOS / iPadOS: <strong>$LatestSupportedIOS</strong>
  </div>
</div>

<!-- OVERVIEW -->
<div class="section-title">Overview</div>
<div class="summary-cards">
  <div class="card"><div class="label">Total Mobile Devices</div><div class="value" style="color:#1a1a2e;">$S_TotalDevices</div></div>
  <div class="card"><div class="label">Android</div><div class="value" style="color:#27ae60;">$S_TotalAndroid</div><div class="sub">$S_OutdatedAndroid outdated</div></div>
  <div class="card"><div class="label">iOS</div><div class="value" style="color:#3498db;">$S_TotalIos</div><div class="sub">$S_OutdatedIos outdated</div></div>
  <div class="card"><div class="label">iPadOS</div><div class="value" style="color:#9b59b6;">$S_TotalIpad</div><div class="sub">$S_OutdatedIpad outdated</div></div>
  <div class="card"><div class="label">Total Outdated</div><div class="value" style="color:#e74c3c;">$S_TotalOutdated</div><div class="sub">$S_PctOutdated% of total</div></div>
</div>

<!-- VERSION SPREAD -->
<div class="dist-section">
  <div class="section-title">Android Version Spread (Latest Supported: $LatestSupportedAndroid)</div>
  <div class="dist-cards">
$S_AndroidCardsHtml
  </div>
</div>

<div class="dist-section">
  <div class="section-title">iOS Version Spread (Latest Supported: $LatestSupportedIOS)</div>
  <div class="dist-cards">
$S_IosCardsHtml
  </div>
</div>

<div class="dist-section">
  <div class="section-title">iPadOS Version Spread (Latest Supported: $LatestSupportedIOS)</div>
  <div class="dist-cards">
$S_IpadCardsHtml
  </div>
</div>

<!-- CHARTS -->
<div class="charts-row">
  <div class="chart-section"><h2>Android by Major Version</h2><div class="chart-container"><canvas id="androidChart"></canvas></div></div>
  <div class="chart-section"><h2>iOS by Major Version</h2><div class="chart-container"><canvas id="iosChart"></canvas></div></div>
  <div class="chart-section"><h2>iPadOS by Major Version</h2><div class="chart-container"><canvas id="ipadChart"></canvas></div></div>
</div>

<!-- DEVICE TABLE -->
<div class="table-section">
  <h2>Device Details</h2>
  <div class="table-controls">
    <input type="text" id="searchBox" placeholder="Search by name, user, model, OS version..." onkeyup="filterTable()" />
    <select id="platformFilter" onchange="filterTable()">
      <option value="all">All Platforms</option>
      <option value="Android">Android</option>
      <option value="iOS">iOS</option>
      <option value="iPadOS">iPadOS</option>
    </select>
    <select id="statusFilter" onchange="filterTable()">
      <option value="all">All Versions</option>
      <option value="Outdated">Outdated Only</option>
      <option value="Supported">Supported Only</option>
      <option value="Unknown">Unknown Only</option>
    </select>
    <span class="count-label" id="rowCount"></span>
  </div>
  <table id="deviceTable">
    <thead><tr>
      <th onclick="sortTable(0)">Device Name</th>
      <th onclick="sortTable(1)">User</th>
      <th onclick="sortTable(2)">Platform</th>
      <th onclick="sortTable(3)">OS Version</th>
      <th onclick="sortTable(4)">Major</th>
      <th onclick="sortTable(5)">Support Status</th>
      <th onclick="sortTable(6)">Manufacturer</th>
      <th onclick="sortTable(7)">Model</th>
      <th onclick="sortTable(8)">Ownership</th>
      <th onclick="sortTable(9)">Compliance</th>
      <th onclick="sortTable(10)">Enrolled</th>
      <th onclick="sortTable(11)">Last Sync</th>
      <th onclick="sortTable(12)">Days Since Sync</th>
    </tr></thead>
    <tbody>
$S_TableRows
    </tbody>
  </table>
</div>

<div class="footer">Report generated by ReportIntuneMobileDevices.ps1</div>

<script>
var androidData = $S_AndroidChartJson;
var iosData     = $S_IosChartJson;
var ipadData    = $S_IpadChartJson;

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

buildChart('androidChart', androidData, '#27ae60');
buildChart('iosChart',     iosData,     '#3498db');
buildChart('ipadChart',    ipadData,    '#9b59b6');

function filterTable() {
  var search = document.getElementById('searchBox').value.toLowerCase();
  var platform = document.getElementById('platformFilter').value;
  var status = document.getElementById('statusFilter').value;
  var rows = document.querySelectorAll('#deviceTable tbody tr');
  var visible = 0;
  rows.forEach(function(row) {
    var text = row.textContent.toLowerCase();
    var matchSearch = !search || text.indexOf(search) !== -1;
    var matchPlatform = platform === 'all' || row.getAttribute('data-platform') === platform;
    var matchStatus = status === 'all' || row.getAttribute('data-status') === status;
    if (matchSearch && matchPlatform && matchStatus) {
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

    $S_Html | Out-File -FilePath $S_HtmlFile -Encoding UTF8

    # --- Console summary ---
    Write-Host ""
    Write-Host "Intune Mobile Devices Report" -ForegroundColor Cyan
    Write-Host "--------------------------------------------"
    Write-Host ("Tenant                   : {0} ({1})" -f $S_TenantDisplayName, $S_TenantId)
    Write-Host ("Latest supported Android : {0}" -f $LatestSupportedAndroid)
    Write-Host ("Latest supported iOS     : {0}" -f $LatestSupportedIOS)
    Write-Host ("Total mobile devices     : {0}" -f $S_TotalDevices)
    Write-Host ("  Android                : {0}  (Outdated: {1})" -f $S_TotalAndroid, $S_OutdatedAndroid) -ForegroundColor Green
    Write-Host ("  iOS                    : {0}  (Outdated: {1})" -f $S_TotalIos, $S_OutdatedIos) -ForegroundColor Green
    Write-Host ("  iPadOS                 : {0}  (Outdated: {1})" -f $S_TotalIpad, $S_OutdatedIpad) -ForegroundColor Green
    Write-Host ("Total outdated           : {0}  ({1}%)" -f $S_TotalOutdated, $S_PctOutdated) -ForegroundColor Red
    Write-Host ""
    Write-Host ("CSV report               : {0}" -f $S_CsvFile) -ForegroundColor Yellow
    Write-Host ("HTML report              : {0}" -f $S_HtmlFile) -ForegroundColor Yellow

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
