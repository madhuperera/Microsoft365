#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Reports on applications managed by Microsoft Intune. Filterable by platform.

.DESCRIPTION
    Connects to Microsoft Graph and retrieves all mobile apps managed by Intune.
    For each application, reports the platform, publisher, version, assignment groups,
    and assignment intent. Supports filtering by platform (Windows, Android, iOS, or All).
    Exports results to CSV and HTML.

.PARAMETER Platform
    Platform to filter results by. Accepts All (default), Windows, Android, or iOS.

.PARAMETER ReportPath
    Folder or file path for the output report. If a folder is specified, a timestamped
    filename is generated automatically. Defaults to the current directory.

.EXAMPLE
    .\ReportIntuneApps.ps1

.EXAMPLE
    .\ReportIntuneApps.ps1 -Platform Windows

.EXAMPLE
    .\ReportIntuneApps.ps1 -Platform iOS -ReportPath "C:\Reports"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet("All", "Windows", "Android", "iOS")]
    [string]$Platform = "All",

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ReportPath
)

$ErrorActionPreference = 'Stop'

$S_ReportPath = $ReportPath

$S_RequiredGraphScopes = @(
    'DeviceManagementApps.Read.All'
    'Group.Read.All'
    'Organization.Read.All'
)

$S_GraphRequestDelayMilliseconds = 5

# --- @odata.type maps for platform filtering ---
$S_PlatformTypeMap = @{
    Windows = @(
        '#microsoft.graph.win32LobApp',
        '#microsoft.graph.windowsMobileMSI',
        '#microsoft.graph.windowsStoreApp',
        '#microsoft.graph.microsoftStoreForBusinessApp',
        '#microsoft.graph.winGetApp',
        '#microsoft.graph.windowsUniversalAppX',
        '#microsoft.graph.windowsAppX',
        '#microsoft.graph.windowsWebApp',
        '#microsoft.graph.windowsMicrosoftEdgeApp',
        '#microsoft.graph.officeSuiteApp'
    )
    iOS     = @(
        '#microsoft.graph.iosLobApp',
        '#microsoft.graph.iosStoreApp',
        '#microsoft.graph.iosVppApp',
        '#microsoft.graph.iosWebApp',
        '#microsoft.graph.managedIOSLobApp',
        '#microsoft.graph.managedIOSStoreApp'
    )
    Android = @(
        '#microsoft.graph.androidLobApp',
        '#microsoft.graph.androidStoreApp',
        '#microsoft.graph.androidManagedStoreApp',
        '#microsoft.graph.androidForWorkApp',
        '#microsoft.graph.managedAndroidLobApp',
        '#microsoft.graph.managedAndroidStoreApp'
    )
}

function Get-AppPlatform
{
    param([string]$OdataType)
    foreach ($F_Plat in $S_PlatformTypeMap.Keys)
    {
        if ($S_PlatformTypeMap[$F_Plat] -contains $OdataType)
        {
            return $F_Plat
        }
    }
    return 'Other'
}

function Get-AppVersion
{
    param($App)
    # Try common version property names across app types
    foreach ($F_Prop in @('displayVersion', 'version', 'versionName', 'productVersion', 'identityVersion', 'bundleVersion'))
    {
        if ($App.PSObject.Properties[$F_Prop] -and $App.$F_Prop)
        {
            return [string]$App.$F_Prop
        }
    }
    return ''
}

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
    Write-Host "  Environment: $($S_ExistingContext.Environment)" -ForegroundColor Cyan
    Write-Host "  Scopes     : $($S_ExistingContext.Scopes -join ', ')" -ForegroundColor Cyan
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

    # --- Fetch all mobileApps with assignments (Beta returns more app types) ---
    Write-Host "Fetching Intune mobile apps with assignments..." -ForegroundColor Cyan
    $S_Apps = New-Object System.Collections.Generic.List[object]
    $S_Uri = 'https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?$expand=assignments&$top=100'
    do
    {
        $S_Resp = Invoke-MgGraphRequest -Method GET -Uri $S_Uri -ErrorAction Stop
        if ($S_Resp.value)
        {
            foreach ($a in $S_Resp.value)
            {
                $S_Apps.Add([pscustomobject]$a) | Out-Null
            }
        }
        $S_Uri = $S_Resp.'@odata.nextLink'
    } while ($S_Uri)
    Write-Host ("  Retrieved {0} apps total" -f $S_Apps.Count) -ForegroundColor Green

    # --- Filter by platform ---
    if ($Platform -ne 'All')
    {
        $S_Wanted = $S_PlatformTypeMap[$Platform]
        $S_Apps = $S_Apps | Where-Object { $S_Wanted -contains $_.'@odata.type' }
        Write-Host ("  After {0} filter: {1} apps" -f $Platform, $S_Apps.Count) -ForegroundColor Green
    }

    # --- Resolve group display names referenced by assignments ---
    $S_GroupIds = New-Object System.Collections.Generic.HashSet[string]
    foreach ($S_App in $S_Apps)
    {
        if ($S_App.assignments)
        {
            foreach ($S_Asn in $S_App.assignments)
            {
                if ($S_Asn.target -and $S_Asn.target.groupId)
                {
                    [void]$S_GroupIds.Add([string]$S_Asn.target.groupId)
                }
            }
        }
    }

    $S_GroupLookup = @{}
    if ($S_GroupIds.Count -gt 0)
    {
        Write-Host ("Resolving {0} group display names..." -f $S_GroupIds.Count) -ForegroundColor Cyan
        foreach ($S_Gid in $S_GroupIds)
        {
            try
            {
                $S_G = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/groups/{0}?`$select=id,displayName" -f $S_Gid) -ErrorAction Stop
                $S_GroupLookup[$S_Gid] = $S_G.displayName
            }
            catch
            {
                $S_GroupLookup[$S_Gid] = "(Unknown / Deleted: $S_Gid)"
            }
        }
    }

    function Resolve-AssignmentTarget
    {
        param($Target)
        if (-not $Target)
        {
            return 'Unknown'
        }
        switch ($Target.'@odata.type')
        {
            '#microsoft.graph.allLicensedUsersAssignmentTarget' { return 'All Users' }
            '#microsoft.graph.allDevicesAssignmentTarget' { return 'All Devices' }
            '#microsoft.graph.groupAssignmentTarget' { return $S_GroupLookup[[string]$Target.groupId] }
            '#microsoft.graph.exclusionGroupAssignmentTarget' { return ('EXCLUDE: {0}' -f $S_GroupLookup[[string]$Target.groupId]) }
            default { return ($Target.'@odata.type' -replace '#microsoft.graph.', '') }
        }
    }

    # --- Build report data ---
    $S_Report = foreach ($S_App in $S_Apps)
    {
        $S_Plat = Get-AppPlatform -OdataType $S_App.'@odata.type'
        $S_Ver = Get-AppVersion -App $S_App
        $S_TypeShort = ($S_App.'@odata.type' -replace '#microsoft.graph.', '')

        $S_Required = New-Object System.Collections.Generic.List[string]
        $S_Available = New-Object System.Collections.Generic.List[string]
        $S_Uninstall = New-Object System.Collections.Generic.List[string]

        if ($S_App.assignments)
        {
            foreach ($S_Asn in $S_App.assignments)
            {
                $S_Name = Resolve-AssignmentTarget -Target $S_Asn.target
                switch ($S_Asn.intent)
                {
                    'required' { $S_Required.Add($S_Name)  | Out-Null }
                    'available' { $S_Available.Add($S_Name) | Out-Null }
                    'availableWithoutEnrollment' { $S_Available.Add($S_Name + ' (No Enrollment)') | Out-Null }
                    'uninstall' { $S_Uninstall.Add($S_Name) | Out-Null }
                }
            }
        }

        [pscustomobject]@{
            DisplayName     = $S_App.displayName
            Platform        = $S_Plat
            AppType         = $S_TypeShort
            Version         = $S_Ver
            Publisher       = $S_App.publisher
            IsAssigned      = [bool]$S_App.isAssigned
            PublishingState = $S_App.publishingState
            CreatedDateTime = $S_App.createdDateTime
            LastModified    = $S_App.lastModifiedDateTime
            RequiredCount   = $S_Required.Count
            AvailableCount  = $S_Available.Count
            UninstallCount  = $S_Uninstall.Count
            RequiredGroups  = ($S_Required -join '; ')
            AvailableGroups = ($S_Available -join '; ')
            UninstallGroups = ($S_Uninstall -join '; ')
        }
    }

    $S_Report = $S_Report | Sort-Object Platform, DisplayName

    # --- Stats ---
    $S_TotalApps = $S_Report.Count
    $S_TotalAssigned = ($S_Report | Where-Object { $_.IsAssigned }).Count
    $S_TotalUnassigned = $S_TotalApps - $S_TotalAssigned
    $S_TotalRequiredAsn = ($S_Report | Measure-Object RequiredCount  -Sum).Sum
    $S_TotalAvailableAsn = ($S_Report | Measure-Object AvailableCount -Sum).Sum
    $S_TotalUninstallAsn = ($S_Report | Measure-Object UninstallCount -Sum).Sum

    $S_PlatformSummary = $S_Report | Group-Object Platform | Sort-Object Count -Descending | ForEach-Object {
        [pscustomobject]@{ Platform = $_.Name; Count = $_.Count }
    }
    $S_TypeSummary = $S_Report | Group-Object AppType | Sort-Object Count -Descending | ForEach-Object {
        [pscustomobject]@{ AppType = $_.Name; Count = $_.Count }
    }

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
    $S_Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $S_FileBase = "ReportIntuneApps_{0}_{1}" -f $Platform, $S_Timestamp
    $S_CsvFile = if (Test-Path $S_ReportPath -PathType Container)
    {
        Join-Path $S_ReportPath ("{0}.csv" -f $S_FileBase)
    }
    else
    {
        $S_ReportPath
    }
    $S_HtmlFile = Join-Path $S_ReportFolder ("{0}.html" -f $S_FileBase)

    # --- CSV export ---
    $S_Report | Export-Csv -Path $S_CsvFile -NoTypeInformation -Encoding UTF8

    # --- HTML report ---
    $S_ReportDate = Get-Date -Format 'dd MMM yyyy HH:mm'

    $S_Enc = {
        param($s)
        if ($null -eq $s -or $s -eq '')
        {
            '-'
        }
        else
        {
            [System.Net.WebUtility]::HtmlEncode([string]$s)
        }
    }

    $S_TableRows = ($S_Report | ForEach-Object {
            $S_Created = if ($_.CreatedDateTime)
            {
                ([datetime]$_.CreatedDateTime).ToString('dd MMM yyyy')
            }
            else
            {
                '-'
            }
            $S_Modified = if ($_.LastModified)
            {
                ([datetime]$_.LastModified).ToString('dd MMM yyyy')
            }
            else
            {
                '-'
            }
            $S_AssignedBadge = if ($_.IsAssigned)
            {
                '<span class="badge badge-active">Yes</span>'
            }
            else
            {
                '<span class="badge badge-disabled">No</span>'
            }
            $S_ReqHtml = if ($_.RequiredCount -gt 0)
            {
                (& $S_Enc $_.RequiredGroups)
            }
            else
            {
                '-'
            }
            $S_AvlHtml = if ($_.AvailableCount -gt 0)
            {
                (& $S_Enc $_.AvailableGroups)
            }
            else
            {
                '-'
            }
            $S_UniHtml = if ($_.UninstallCount -gt 0)
            {
                (& $S_Enc $_.UninstallGroups)
            }
            else
            {
                '-'
            }
            "<tr data-platform=`"$($_.Platform)`" data-assigned=`"$([int][bool]$_.IsAssigned)`"><td>$(& $S_Enc $_.DisplayName)</td><td><span class=`"badge badge-platform`">$($_.Platform)</span></td><td>$(& $S_Enc $_.AppType)</td><td>$(& $S_Enc $_.Version)</td><td>$(& $S_Enc $_.Publisher)</td><td>$S_AssignedBadge</td><td>$(& $S_Enc $_.PublishingState)</td><td>$S_Created</td><td>$S_Modified</td><td><span class=`"count-pill count-req`">$($_.RequiredCount)</span></td><td>$S_ReqHtml</td><td><span class=`"count-pill count-avl`">$($_.AvailableCount)</span></td><td>$S_AvlHtml</td><td><span class=`"count-pill count-uni`">$($_.UninstallCount)</span></td><td>$S_UniHtml</td></tr>"
        }) -join "`n"

    $S_PlatformOptions = ($S_PlatformSummary | ForEach-Object { "<option value=`"$($_.Platform)`">$($_.Platform) ($($_.Count))</option>" }) -join "`n      "

    $S_Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Intune Apps Report - $Platform</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; padding: 30px; }
  .header { background: linear-gradient(135deg, #1a1a2e, #16213e); color: #fff; padding: 30px 40px; border-radius: 12px; margin-bottom: 30px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 16px; }
  .header-left h1 { font-size: 1.6em; margin-bottom: 6px; }
  .header-left p { font-size: 0.9em; opacity: 0.85; }
  .header-right .pill { background: rgba(255,255,255,0.15); padding: 8px 16px; border-radius: 20px; font-weight: 600; font-size: 0.9em; }

  .section-title { font-size: 1.15em; font-weight: 600; color: #1a1a2e; margin-bottom: 16px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0; }

  .summary-cards { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .card { background: #fff; border-radius: 10px; padding: 24px 30px; flex: 1; min-width: 170px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
  .card .label { font-size: 0.82em; color: #777; text-transform: uppercase; letter-spacing: 0.5px; }
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
  .table-controls input[type="text"] { padding: 8px 14px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; min-width: 280px; }
  .table-controls select { padding: 8px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; background: #fff; }
  .table-controls .count-label { font-size: 0.85em; color: #777; margin-left: auto; }
  table { width: 100%; border-collapse: collapse; font-size: 0.85em; }
  th { background: #1a1a2e; color: #fff; padding: 12px 14px; text-align: left; position: sticky; top: 0; cursor: pointer; user-select: none; white-space: nowrap; }
  th:hover { background: #2c3e50; }
  td { padding: 10px 14px; border-bottom: 1px solid #eee; vertical-align: top; }
  td.wrap { white-space: normal; max-width: 280px; word-break: break-word; }
  tr:hover td { background: #f8f9fa; }
  tr.hidden-row { display: none; }

  .badge { padding: 3px 10px; border-radius: 12px; font-size: 0.8em; font-weight: 600; white-space: nowrap; }
  .badge-active { background: #d4edda; color: #155724; }
  .badge-disabled { background: #e2e3e5; color: #495057; }
  .badge-platform { background: #d6eaf8; color: #1f4e79; }

  .count-pill { display: inline-block; min-width: 28px; padding: 2px 8px; border-radius: 10px; text-align: center; font-weight: 700; font-size: 0.82em; color: #fff; }
  .count-req { background: #e74c3c; }
  .count-avl { background: #27ae60; }
  .count-uni { background: #f39c12; }

  .footer { text-align: center; font-size: 0.8em; color: #999; margin-top: 20px; }
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <h1>Intune Apps Report</h1>
    <p>Tenant: $(& $S_Enc $S_TenantDisplayName) ($S_TenantId) &nbsp;|&nbsp; Platform Filter: <strong>$Platform</strong> &nbsp;|&nbsp; Generated: $S_ReportDate</p>
  </div>
  <div class="header-right">
    <span class="pill">$S_TotalApps Apps</span>
  </div>
</div>

<!-- OVERVIEW -->
<div class="section-title">Overview</div>
<div class="summary-cards">
  <div class="card"><div class="label">Total Apps</div><div class="value" style="color:#1a1a2e;">$S_TotalApps</div></div>
  <div class="card"><div class="label">Assigned</div><div class="value" style="color:#27ae60;">$S_TotalAssigned</div><div class="sub">$S_TotalUnassigned unassigned</div></div>
  <div class="card"><div class="label">Required Assignments</div><div class="value" style="color:#e74c3c;">$S_TotalRequiredAsn</div></div>
  <div class="card"><div class="label">Available Assignments</div><div class="value" style="color:#27ae60;">$S_TotalAvailableAsn</div></div>
  <div class="card"><div class="label">Uninstall Assignments</div><div class="value" style="color:#f39c12;">$S_TotalUninstallAsn</div></div>
</div>

<!-- PLATFORM DIST -->
<div class="dist-section">
  <div class="section-title">Apps by Platform</div>
  <div class="dist-cards">
$(($S_PlatformSummary | ForEach-Object { "    <div class=`"dist-card`"><div class=`"dist-label`">$($_.Platform)</div><div class=`"dist-value`">$($_.Count)</div></div>" }) -join "`n")
  </div>
</div>

<!-- APP TYPE DIST -->
<div class="dist-section">
  <div class="section-title">Apps by Type</div>
  <div class="dist-cards">
$(($S_TypeSummary | ForEach-Object { "    <div class=`"dist-card`"><div class=`"dist-label`">$($_.AppType)</div><div class=`"dist-value`">$($_.Count)</div></div>" }) -join "`n")
  </div>
</div>

<!-- CHARTS -->
<div class="charts-row">
  <div class="chart-section">
    <h2>Apps by Platform</h2>
    <div class="chart-container"><canvas id="platformChart"></canvas></div>
  </div>
  <div class="chart-section">
    <h2>Assignment Intent Breakdown</h2>
    <div class="chart-container"><canvas id="intentChart"></canvas></div>
  </div>
</div>

<!-- APPS TABLE -->
<div class="table-section">
  <h2>App Details and Group Assignments</h2>
  <div class="table-controls">
    <input type="text" id="searchBox" placeholder="Search by name, publisher, group..." onkeyup="filterTable()" />
    <select id="platformFilter" onchange="filterTable()">
      <option value="all">All Platforms</option>
      $S_PlatformOptions
    </select>
    <select id="assignedFilter" onchange="filterTable()">
      <option value="all">All Apps</option>
      <option value="1">Assigned Only</option>
      <option value="0">Unassigned Only</option>
    </select>
    <span class="count-label" id="rowCount"></span>
  </div>
  <table id="appTable">
    <thead><tr>
      <th onclick="sortTable(0)">App Name</th>
      <th onclick="sortTable(1)">Platform</th>
      <th onclick="sortTable(2)">Type</th>
      <th onclick="sortTable(3)">Version</th>
      <th onclick="sortTable(4)">Publisher</th>
      <th onclick="sortTable(5)">Assigned</th>
      <th onclick="sortTable(6)">State</th>
      <th onclick="sortTable(7)">Created</th>
      <th onclick="sortTable(8)">Modified</th>
      <th onclick="sortTable(9)">#Req</th>
      <th>Required Groups</th>
      <th onclick="sortTable(11)">#Avl</th>
      <th>Available Groups</th>
      <th onclick="sortTable(13)">#Uni</th>
      <th>Uninstall Groups</th>
    </tr></thead>
    <tbody>
$S_TableRows
    </tbody>
  </table>
</div>

<div class="footer">Report generated by ReportIntuneApps.ps1</div>

<script>
var chartColors = ['#3498db','#27ae60','#e74c3c','#f39c12','#9b59b6','#1abc9c','#e67e22','#2c3e50','#95a5a6','#d35400'];

var platformData = { $(($S_PlatformSummary | ForEach-Object { "'$($_.Platform)': $($_.Count)" }) -join ', ') };
var intentData = { 'Required': $S_TotalRequiredAsn, 'Available': $S_TotalAvailableAsn, 'Uninstall': $S_TotalUninstallAsn };

function makeDoughnut(canvasId, dataObj) {
  var keys = Object.keys(dataObj);
  new Chart(document.getElementById(canvasId), {
    type: 'doughnut',
    data: { labels: keys, datasets: [{ data: keys.map(function(k){return dataObj[k];}), backgroundColor: keys.map(function(_,i){return chartColors[i % chartColors.length];}), borderWidth: 2, borderColor: '#fff' }] },
    options: { responsive: true, plugins: { legend: { position: 'right', labels: { padding: 16, font: { size: 13 }, boxWidth: 16 } }, tooltip: { callbacks: { label: function(ctx) { var t = ctx.dataset.data.reduce(function(a,b){return a+b;},0); return ctx.label+': '+ctx.parsed+' ('+(t>0?((ctx.parsed/t)*100).toFixed(1):0)+'%)'; } } } } }
  });
}
makeDoughnut('platformChart', platformData);
makeDoughnut('intentChart', intentData);

// Mark grouped/long-content cells as wrap
document.querySelectorAll('#appTable tbody tr').forEach(function(row){
  [10, 12, 14].forEach(function(i){ if (row.cells[i]) row.cells[i].classList.add('wrap'); });
});

function filterTable() {
  var search = document.getElementById('searchBox').value.toLowerCase();
  var plat = document.getElementById('platformFilter').value;
  var asn = document.getElementById('assignedFilter').value;
  var rows = document.querySelectorAll('#appTable tbody tr');
  var visible = 0;
  rows.forEach(function(row) {
    var text = row.textContent.toLowerCase();
    var rowPlat = row.getAttribute('data-platform');
    var rowAsn = row.getAttribute('data-assigned');
    var matchSearch = !search || text.indexOf(search) !== -1;
    var matchPlat = plat === 'all' || rowPlat === plat;
    var matchAsn = asn === 'all' || rowAsn === asn;
    if (matchSearch && matchPlat && matchAsn) {
      row.classList.remove('hidden-row');
      visible++;
    } else {
      row.classList.add('hidden-row');
    }
  });
  document.getElementById('rowCount').textContent = 'Showing ' + visible + ' of ' + rows.length + ' apps';
}

var sortDir = {};
function sortTable(col) {
  var tbody = document.getElementById('appTable').querySelector('tbody');
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
    Write-Host "Intune Apps Report" -ForegroundColor Cyan
    Write-Host "--------------------------------------------"
    Write-Host ("Tenant                   : {0} ({1})" -f $S_TenantDisplayName, $S_TenantId)
    Write-Host ("Platform filter          : {0}" -f $Platform)
    Write-Host ("Total apps               : {0}" -f $S_TotalApps)
    Write-Host ("Assigned / Unassigned    : {0} / {1}" -f $S_TotalAssigned, $S_TotalUnassigned)
    Write-Host ("Required assignments     : {0}" -f $S_TotalRequiredAsn)  -ForegroundColor Red
    Write-Host ("Available assignments    : {0}" -f $S_TotalAvailableAsn) -ForegroundColor Green
    Write-Host ("Uninstall assignments    : {0}" -f $S_TotalUninstallAsn) -ForegroundColor Yellow
    Write-Host ""
    Write-Host "By Platform" -ForegroundColor Cyan
    foreach ($p in $S_PlatformSummary)
    {
        Write-Host ("  {0,-15}: {1}" -f $p.Platform, $p.Count)
    }
    Write-Host ""
    Write-Host ("CSV report               : {0}" -f $S_CsvFile)  -ForegroundColor Yellow
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
