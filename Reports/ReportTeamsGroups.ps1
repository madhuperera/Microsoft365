#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Teams, Microsoft.Graph.Identity.DirectoryManagement

<#
.SYNOPSIS
    Reports on all Microsoft Teams-enabled groups in the tenant, including
    member count, owner count, and visibility (Public or Private).

.DESCRIPTION
    Connects to Microsoft Graph and enumerates every Microsoft 365 group that
    is Teams-enabled (resourceProvisioningOptions contains "Team"). For each
    group the script collects the display name, mail nickname, visibility,
    creation date, and the number of owners and members, and writes the
    result to CSV in the current location (or to -ReportPath if supplied).

    A small configurable delay is inserted between Graph calls to reduce
    pressure on the service.

.PARAMETER ReportPath
    Optional. Path to a folder or full file path for the CSV output. When a
    folder is supplied (or the parameter is omitted) a timestamped file name
    is generated in that folder. Defaults to the current working location.

.EXAMPLE
    .\ReportTeamsGroups.ps1

    Generates a timestamped CSV in the current directory containing every
    Teams-enabled group with member count, owner count, and visibility.

.EXAMPLE
    .\ReportTeamsGroups.ps1 -ReportPath C:\Reports\TeamsGroups.csv

    Writes the report to the specified CSV file.

.NOTES
    Required Microsoft Graph permissions:
      - Group.Read.All
      - GroupMember.Read.All
      - Team.ReadBasic.All
      - Organization.Read.All (used to resolve the tenant display name in the HTML report)

    Shared as-is without warranty.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ReportPath
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Script-level configuration
# ---------------------------------------------------------------------------

$S_RequiredGraphScopes = @(
    'Group.Read.All'
    'GroupMember.Read.All'
    'Team.ReadBasic.All'
    'Organization.Read.All'
)

$S_GraphRequestDelayMilliseconds       = 5
$S_RequireGraphContextConfirmation     = $true
$S_GraphContextConfirmationDelaySeconds = 10
$S_DisconnectGraphSessionOnExit        = $false

# ---------------------------------------------------------------------------
# Module checks
# ---------------------------------------------------------------------------

try
{
    foreach ($S_Module in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Teams', 'Microsoft.Graph.Identity.DirectoryManagement'))
    {
        if (-not (Get-Module -ListAvailable -Name $S_Module))
        {
            throw "$S_Module module is not installed. Install it using 'Install-Module $S_Module -Scope CurrentUser'."
        }

        Import-Module $S_Module -ErrorAction Stop | Out-Null
    }

    # -----------------------------------------------------------------------
    # Microsoft Graph connection handling
    # -----------------------------------------------------------------------

    $S_Context = Get-MgContext

    if ($S_Context)
    {
        Write-Host "An existing Microsoft Graph context was found:" -ForegroundColor Cyan
        Write-Host ("  Account     : {0}" -f $S_Context.Account)
        Write-Host ("  Tenant ID   : {0}" -f $S_Context.TenantId)
        Write-Host ("  Environment : {0}" -f $S_Context.Environment)
        Write-Host ("  Scopes      : {0}" -f ($S_Context.Scopes -join ', '))

        $S_ContinueChoice = Read-Host "Continue using this Microsoft Graph connection? (Y/N)"
        if ($S_ContinueChoice -notmatch '^(y|yes)$')
        {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            $S_Context = $null
        }
    }

    if (-not $S_Context)
    {
        Connect-MgGraph -Scopes $S_RequiredGraphScopes -ErrorAction Stop | Out-Null
        $S_Context = Get-MgContext

        Write-Host "Connected to Microsoft Graph:" -ForegroundColor Cyan
        Write-Host ("  Account     : {0}" -f $S_Context.Account)
        Write-Host ("  Tenant ID   : {0}" -f $S_Context.TenantId)
        Write-Host ("  Environment : {0}" -f $S_Context.Environment)
        Write-Host ("  Scopes      : {0}" -f ($S_Context.Scopes -join ', '))

        if ($S_RequireGraphContextConfirmation)
        {
            $S_ConfirmChoice = Read-Host "Is this the correct tenant and account to continue? (Y/N)"
            if ($S_ConfirmChoice -notmatch '^(y|yes)$')
            {
                throw "Operator did not confirm the Microsoft Graph context. Aborting."
            }

            Start-Sleep -Seconds $S_GraphContextConfirmationDelaySeconds
        }
    }

    # -----------------------------------------------------------------------
    # Collect Teams-enabled groups
    # -----------------------------------------------------------------------

    Write-Verbose "Retrieving Teams-enabled groups from Microsoft Graph."

    $S_TeamsGroups = Get-MgGroup -All `
        -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" `
        -Property 'id,displayName,mailNickname,visibility,createdDateTime,description' `
        -ConsistencyLevel eventual

    Start-Sleep -Milliseconds $S_GraphRequestDelayMilliseconds

    $S_TotalGroups = ($S_TeamsGroups | Measure-Object).Count
    Write-Host ("Found {0} Teams-enabled groups." -f $S_TotalGroups) -ForegroundColor Cyan

    # -----------------------------------------------------------------------
    # Build the report
    # -----------------------------------------------------------------------

    $S_Report = New-Object System.Collections.Generic.List[object]
    $S_Counter = 0

    foreach ($S_Group in $S_TeamsGroups)
    {
        $S_Counter++

        $S_PercentComplete = if ($S_TotalGroups -gt 0) { [int](($S_Counter / $S_TotalGroups) * 100) } else { 0 }
        Write-Progress -Activity "Collecting owner and member counts" `
            -Status ("[{0}/{1}] {2}" -f $S_Counter, $S_TotalGroups, $S_Group.DisplayName) `
            -PercentComplete $S_PercentComplete

        Write-Verbose ("[{0}/{1}] Processing {2}" -f $S_Counter, $S_TotalGroups, $S_Group.DisplayName)

        $F_OwnerCount  = 0
        $F_MemberCount = 0
        $F_OwnerNames  = @()
        $F_OwnerUpns   = @()

        try
        {
            $F_Owners = Get-MgGroupOwner -GroupId $S_Group.Id -All -ErrorAction Stop
            $F_OwnerCount = ($F_Owners | Measure-Object).Count

            foreach ($F_Owner in $F_Owners)
            {
                $F_OwnerName = $null
                $F_OwnerUpn  = $null
                if ($F_Owner.AdditionalProperties)
                {
                    if ($F_Owner.AdditionalProperties.ContainsKey('displayName'))        { $F_OwnerName = [string]$F_Owner.AdditionalProperties['displayName'] }
                    if ($F_Owner.AdditionalProperties.ContainsKey('userPrincipalName')) { $F_OwnerUpn  = [string]$F_Owner.AdditionalProperties['userPrincipalName'] }
                    if (-not $F_OwnerUpn -and $F_Owner.AdditionalProperties.ContainsKey('mail')) { $F_OwnerUpn = [string]$F_Owner.AdditionalProperties['mail'] }
                }
                if (-not $F_OwnerName) { $F_OwnerName = if ($F_OwnerUpn) { $F_OwnerUpn } else { $F_Owner.Id } }
                $F_OwnerNames += $F_OwnerName
                if ($F_OwnerUpn) { $F_OwnerUpns += $F_OwnerUpn }
            }

            Start-Sleep -Milliseconds $S_GraphRequestDelayMilliseconds
        }
        catch
        {
            Write-Warning ("Failed to retrieve owners for '{0}': {1}" -f $S_Group.DisplayName, $_.Exception.Message)
        }

        try
        {
            $F_Members = Get-MgGroupMember -GroupId $S_Group.Id -All -ErrorAction Stop
            $F_MemberCount = ($F_Members | Measure-Object).Count
            Start-Sleep -Milliseconds $S_GraphRequestDelayMilliseconds
        }
        catch
        {
            Write-Warning ("Failed to retrieve members for '{0}': {1}" -f $S_Group.DisplayName, $_.Exception.Message)
        }

        $S_Report.Add([pscustomobject]@{
            DisplayName     = $S_Group.DisplayName
            MailNickname    = $S_Group.MailNickname
            Visibility      = $S_Group.Visibility
            OwnerCount      = $F_OwnerCount
            OwnerNames      = ($F_OwnerNames -join '; ')
            OwnerUpns       = ($F_OwnerUpns  -join '; ')
            MemberCount     = $F_MemberCount
            CreatedDateTime = $S_Group.CreatedDateTime
            GroupId         = $S_Group.Id
            Description     = $S_Group.Description
        })
    }

    Write-Progress -Activity "Collecting owner and member counts" -Completed

    # -----------------------------------------------------------------------
    # Resolve tenant display name
    # -----------------------------------------------------------------------

    $S_TenantDisplayName = $null
    try
    {
        $S_Org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
        $S_TenantDisplayName = $S_Org.DisplayName
    }
    catch
    {
        Write-Verbose ("Failed to resolve tenant display name: {0}" -f $_.Exception.Message)
    }
    if (-not $S_TenantDisplayName) { $S_TenantDisplayName = $S_Context.TenantId }
    $S_TenantId = if ($S_Context.TenantId) { $S_Context.TenantId } else { 'Unknown' }

    # -----------------------------------------------------------------------
    # Resolve output path
    # -----------------------------------------------------------------------

    if (-not $ReportPath)
    {
        $ReportPath = (Get-Location).Path
    }

    $S_Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

    if (Test-Path -Path $ReportPath -PathType Container)
    {
        $S_OutputPath = Join-Path -Path $ReportPath -ChildPath ("ReportTeamsGroups_{0}.csv" -f $S_Timestamp)
    }
    else
    {
        $S_OutputFolder = Split-Path -Parent $ReportPath
        if ($S_OutputFolder -and -not (Test-Path -Path $S_OutputFolder))
        {
            New-Item -ItemType Directory -Path $S_OutputFolder -Force | Out-Null
        }
        $S_OutputPath = $ReportPath
    }

    $S_HtmlOutputPath = [System.IO.Path]::ChangeExtension($S_OutputPath, '.html')

    $S_Report |
        Sort-Object DisplayName |
        Export-Csv -Path $S_OutputPath -NoTypeInformation -Encoding UTF8

    # -----------------------------------------------------------------------
    # Summary statistics
    # -----------------------------------------------------------------------

    $S_PublicCount        = ($S_Report | Where-Object { $_.Visibility -eq 'Public' }  | Measure-Object).Count
    $S_PrivateCount       = ($S_Report | Where-Object { $_.Visibility -eq 'Private' } | Measure-Object).Count
    $S_NoOwnerCount       = ($S_Report | Where-Object { $_.OwnerCount -eq 0 } | Measure-Object).Count
    $S_SingleOwnerCount   = ($S_Report | Where-Object { $_.OwnerCount -eq 1 } | Measure-Object).Count
    $S_PublicNoOwnerCount = ($S_Report | Where-Object { $_.Visibility -eq 'Public' -and $_.OwnerCount -eq 0 } | Measure-Object).Count

    $S_PublicPercent  = if ($S_TotalGroups -gt 0) { [math]::Round(($S_PublicCount  / $S_TotalGroups) * 100, 1) } else { 0 }
    $S_PrivatePercent = if ($S_TotalGroups -gt 0) { [math]::Round(($S_PrivateCount / $S_TotalGroups) * 100, 1) } else { 0 }

    # -----------------------------------------------------------------------
    # Build HTML report
    # -----------------------------------------------------------------------

    $S_ReportDate = Get-Date -Format 'dd MMM yyyy HH:mm'

    # Owner-count buckets
    $S_OwnerBucket0   = ($S_Report | Where-Object { $_.OwnerCount -eq 0 } | Measure-Object).Count
    $S_OwnerBucket1   = ($S_Report | Where-Object { $_.OwnerCount -eq 1 } | Measure-Object).Count
    $S_OwnerBucket2   = ($S_Report | Where-Object { $_.OwnerCount -eq 2 } | Measure-Object).Count
    $S_OwnerBucket3_5 = ($S_Report | Where-Object { $_.OwnerCount -ge 3 -and $_.OwnerCount -le 5 } | Measure-Object).Count
    $S_OwnerBucket6   = ($S_Report | Where-Object { $_.OwnerCount -ge 6 } | Measure-Object).Count

    # Membership-size buckets
    $S_MemBucket0    = ($S_Report | Where-Object { $_.MemberCount -eq 0 } | Measure-Object).Count
    $S_MemBucket1_10 = ($S_Report | Where-Object { $_.MemberCount -ge 1   -and $_.MemberCount -le 10 }   | Measure-Object).Count
    $S_MemBucket11_50= ($S_Report | Where-Object { $_.MemberCount -ge 11  -and $_.MemberCount -le 50 }   | Measure-Object).Count
    $S_MemBucket51_250 = ($S_Report | Where-Object { $_.MemberCount -ge 51 -and $_.MemberCount -le 250 } | Measure-Object).Count
    $S_MemBucket251  = ($S_Report | Where-Object { $_.MemberCount -ge 251 } | Measure-Object).Count

    $S_TableRows = ($S_Report | Sort-Object DisplayName | ForEach-Object {
        $F_Visibility   = if ($_.Visibility) { $_.Visibility } else { 'Unknown' }
        $F_VisClass     = if ($F_Visibility -eq 'Public') { 'badge-public' } elseif ($F_Visibility -eq 'Private') { 'badge-private' } else { 'badge-disabled' }
        $F_OwnerClass   = if ($_.OwnerCount -eq 0) { 'badge-critical' } elseif ($_.OwnerCount -eq 1) { 'badge-warning' } else { 'badge-active' }
        $F_Created      = if ($_.CreatedDateTime) { ([datetime]$_.CreatedDateTime).ToString('dd MMM yyyy') } else { '-' }
        $F_Description  = if ($_.Description) {
            $F_Desc = [string]$_.Description
            if ($F_Desc.Length -gt 120) { $F_Desc = $F_Desc.Substring(0, 117) + '...' }
            [System.Net.WebUtility]::HtmlEncode($F_Desc)
        } else { '-' }
        $F_RowClass     = if ($F_Visibility -eq 'Public' -and $_.OwnerCount -eq 0) { 'row-critical' } else { '' }
        $F_OwnerRisk    = if ($_.OwnerCount -eq 0) { 'none' } elseif ($_.OwnerCount -eq 1) { 'single' } else { 'ok' }

        # Owner cell: list display names, with UPNs as tooltip; data-owners attribute holds lowercased names+UPNs for the filter
        if ($_.OwnerCount -gt 0 -and $_.OwnerNames)
        {
            $F_OwnerNamesHtml = [System.Net.WebUtility]::HtmlEncode([string]$_.OwnerNames).Replace('; ', '<br>')
            $F_OwnerTitle     = [System.Net.WebUtility]::HtmlEncode([string]$_.OwnerUpns)
            $F_OwnerCell      = "<td title=`"$F_OwnerTitle`">$F_OwnerNamesHtml</td>"
        } else {
            $F_OwnerCell = '<td><em style="color:#c0392b;">(none)</em></td>'
        }
        $F_OwnerSearch = (([string]$_.OwnerNames + ' ' + [string]$_.OwnerUpns)).ToLower()
        $F_OwnerSearchAttr = [System.Net.WebUtility]::HtmlEncode($F_OwnerSearch)

        "<tr class=`"$F_RowClass`" data-vis=`"$($F_Visibility.ToLower())`" data-ownerrisk=`"$F_OwnerRisk`" data-owners=`"$F_OwnerSearchAttr`"><td>$([System.Net.WebUtility]::HtmlEncode($_.DisplayName))</td><td>$([System.Net.WebUtility]::HtmlEncode($_.MailNickname))</td><td><span class=`"badge $F_VisClass`">$F_Visibility</span></td><td><span class=`"badge $F_OwnerClass`">$($_.OwnerCount)</span></td>$F_OwnerCell<td>$($_.MemberCount)</td><td>$F_Created</td><td>$F_Description</td></tr>"
    }) -join "`n"

    $S_Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Teams Groups Report</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; padding: 30px; }
  .header { background: linear-gradient(135deg, #1a1a2e, #16213e); color: #fff; padding: 30px 40px; border-radius: 12px; margin-bottom: 30px; }
  .header h1 { font-size: 1.6em; margin-bottom: 6px; }
  .header p { font-size: 0.9em; opacity: 0.8; }

  .section-title { font-size: 1.15em; font-weight: 600; color: #1a1a2e; margin-bottom: 16px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0; }

  .summary-cards { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .card { background: #fff; border-radius: 10px; padding: 24px 30px; flex: 1; min-width: 180px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
  .card .label { font-size: 0.85em; color: #777; text-transform: uppercase; letter-spacing: 0.5px; }
  .card .value { font-size: 2em; font-weight: 700; margin-top: 6px; }
  .card .sub { font-size: 0.78em; color: #999; margin-top: 2px; }
  .card.critical { border-left: 4px solid #e74c3c; }
  .card.warning  { border-left: 4px solid #f39c12; }

  .dist-section { margin-bottom: 30px; }
  .dist-cards { display: flex; gap: 16px; flex-wrap: wrap; }
  .dist-card { background: #fff; border-radius: 10px; padding: 18px 24px; min-width: 140px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); border-left: 4px solid #3498db; text-align: center; }
  .dist-card.critical { border-left-color: #e74c3c; }
  .dist-card.warning  { border-left-color: #f39c12; }
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
  td { padding: 10px 14px; border-bottom: 1px solid #eee; vertical-align: top; }
  tr:hover td { background: #f8f9fa; }
  tr.hidden-row { display: none; }
  tr.row-critical td { background: #fdecea; }
  tr.row-critical:hover td { background: #fbd6d2; }

  .badge { padding: 3px 10px; border-radius: 12px; font-size: 0.82em; font-weight: 600; }
  .badge-active   { background: #d4edda; color: #155724; }
  .badge-warning  { background: #fff3cd; color: #856404; }
  .badge-critical { background: #f8d7da; color: #721c24; }
  .badge-public   { background: #f8d7da; color: #721c24; }
  .badge-private  { background: #d1ecf1; color: #0c5460; }
  .badge-disabled { background: #e2e3e5; color: #495057; }

  .footer { text-align: center; font-size: 0.8em; color: #999; margin-top: 20px; }
</style>
</head>
<body>

<div class="header">
  <h1>Teams Groups Report</h1>
  <p>Tenant: $([System.Net.WebUtility]::HtmlEncode($S_TenantDisplayName)) ($S_TenantId) &nbsp;|&nbsp; Generated: $S_ReportDate</p>
</div>

<!-- OVERVIEW -->
<div class="section-title">Overview</div>
<div class="summary-cards">
  <div class="card"><div class="label">Total Teams</div><div class="value" style="color:#1a1a2e;">$S_TotalGroups</div></div>
  <div class="card"><div class="label">Public</div><div class="value" style="color:#e74c3c;">$S_PublicCount</div><div class="sub">$S_PublicPercent% of total</div></div>
  <div class="card"><div class="label">Private</div><div class="value" style="color:#3498db;">$S_PrivateCount</div><div class="sub">$S_PrivatePercent% of total</div></div>
  <div class="card critical"><div class="label">Public + No Owner</div><div class="value" style="color:#e74c3c;">$S_PublicNoOwnerCount</div><div class="sub">Critical risk</div></div>
  <div class="card critical"><div class="label">Teams with No Owner</div><div class="value" style="color:#e74c3c;">$S_NoOwnerCount</div></div>
  <div class="card warning"><div class="label">Teams with Single Owner</div><div class="value" style="color:#f39c12;">$S_SingleOwnerCount</div></div>
</div>

<!-- OWNER COUNT DISTRIBUTION -->
<div class="dist-section">
  <div class="section-title">Owner Count Distribution</div>
  <div class="dist-cards">
    <div class="dist-card critical"><div class="dist-label">0 Owners</div><div class="dist-value">$S_OwnerBucket0</div></div>
    <div class="dist-card warning"><div class="dist-label">1 Owner</div><div class="dist-value">$S_OwnerBucket1</div></div>
    <div class="dist-card"><div class="dist-label">2 Owners</div><div class="dist-value">$S_OwnerBucket2</div></div>
    <div class="dist-card"><div class="dist-label">3-5 Owners</div><div class="dist-value">$S_OwnerBucket3_5</div></div>
    <div class="dist-card"><div class="dist-label">6+ Owners</div><div class="dist-value">$S_OwnerBucket6</div></div>
  </div>
</div>

<!-- CHARTS -->
<div class="charts-row">
  <div class="chart-section">
    <h2>Visibility</h2>
    <div class="chart-container"><canvas id="visChart"></canvas></div>
  </div>
  <div class="chart-section">
    <h2>Membership Size</h2>
    <div class="chart-container"><canvas id="memChart"></canvas></div>
  </div>
</div>

<!-- TABLE -->
<div class="table-section">
  <h2>Teams Group Details</h2>
  <div class="table-controls">
    <input type="text" id="searchBox" placeholder="Search by name, mail nickname, description..." onkeyup="filterTable()" />
    <input type="text" id="ownerSearch" placeholder="Filter by owner name or UPN..." onkeyup="filterTable()" />
    <select id="visFilter" onchange="filterTable()">
      <option value="all">All Visibility</option>
      <option value="public">Public Only</option>
      <option value="private">Private Only</option>
    </select>
    <select id="ownerFilter" onchange="filterTable()">
      <option value="all">All Owner Counts</option>
      <option value="none">No Owners</option>
      <option value="single">Single Owner</option>
      <option value="ok">2+ Owners</option>
      <option value="critical">Public + No Owner (critical)</option>
    </select>
    <span class="count-label" id="rowCount"></span>
  </div>
  <table id="teamsTable">
    <thead><tr>
      <th onclick="sortTable(0)">Display Name</th>
      <th onclick="sortTable(1)">Mail Nickname</th>
      <th onclick="sortTable(2)">Visibility</th>
      <th onclick="sortTable(3)">Owners</th>
      <th onclick="sortTable(4)">Owner Names</th>
      <th onclick="sortTable(5)">Members</th>
      <th onclick="sortTable(6)">Created</th>
      <th onclick="sortTable(7)">Description</th>
    </tr></thead>
    <tbody>
$S_TableRows
    </tbody>
  </table>
</div>

<div class="footer">Report generated by ReportTeamsGroups.ps1</div>

<script>
var chartOpts = function(pos) {
  return { responsive: true, plugins: { legend: { position: pos || 'right', labels: { padding: 16, font: { size: 13 }, boxWidth: 16 } }, tooltip: { callbacks: { label: function(ctx) { var t = ctx.dataset.data.reduce(function(a,b){return a+b},0); return ctx.label+': '+ctx.parsed+' ('+(t>0?((ctx.parsed/t)*100).toFixed(1):0)+'%)'; } } } } };
};

new Chart(document.getElementById('visChart'), {
  type: 'doughnut',
  data: {
    labels: ['Public', 'Private'],
    datasets: [{ data: [$S_PublicCount, $S_PrivateCount], backgroundColor: ['#e74c3c', '#3498db'], borderWidth: 2, borderColor: '#fff' }]
  },
  options: chartOpts()
});

new Chart(document.getElementById('memChart'), {
  type: 'doughnut',
  data: {
    labels: ['0', '1-10', '11-50', '51-250', '251+'],
    datasets: [{ data: [$S_MemBucket0, $S_MemBucket1_10, $S_MemBucket11_50, $S_MemBucket51_250, $S_MemBucket251], backgroundColor: ['#95a5a6', '#3498db', '#27ae60', '#f39c12', '#9b59b6'], borderWidth: 2, borderColor: '#fff' }]
  },
  options: chartOpts()
});

function filterTable() {
  var search = document.getElementById('searchBox').value.toLowerCase();
  var ownerSearch = document.getElementById('ownerSearch').value.toLowerCase().trim();
  var vis = document.getElementById('visFilter').value;
  var owner = document.getElementById('ownerFilter').value;
  var rows = document.querySelectorAll('#teamsTable tbody tr');
  var visible = 0;
  rows.forEach(function(row) {
    var text = row.textContent.toLowerCase();
    var rowVis = row.getAttribute('data-vis');
    var rowOwner = row.getAttribute('data-ownerrisk');
    var rowOwners = row.getAttribute('data-owners') || '';
    var matchSearch = !search || text.indexOf(search) !== -1;
    var matchOwnerSearch = !ownerSearch || rowOwners.indexOf(ownerSearch) !== -1;
    var matchVis = vis === 'all' || rowVis === vis;
    var matchOwner;
    if (owner === 'all') {
      matchOwner = true;
    } else if (owner === 'critical') {
      matchOwner = (rowVis === 'public' && rowOwner === 'none');
    } else {
      matchOwner = rowOwner === owner;
    }
    if (matchSearch && matchOwnerSearch && matchVis && matchOwner) {
      row.classList.remove('hidden-row');
      visible++;
    } else {
      row.classList.add('hidden-row');
    }
  });
  document.getElementById('rowCount').textContent = 'Showing ' + visible + ' of ' + rows.length + ' teams';
}

var sortDir = {};
function sortTable(col) {
  var tbody = document.getElementById('teamsTable').querySelector('tbody');
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

filterTable();
</script>
</body>
</html>
"@

    $S_Html | Out-File -FilePath $S_HtmlOutputPath -Encoding UTF8

    # -----------------------------------------------------------------------
    # Console summary
    # -----------------------------------------------------------------------

    Write-Host ""
    Write-Host "Teams Groups Report" -ForegroundColor Cyan
    Write-Host "--------------------------------------------"
    Write-Host ("Total Teams-enabled groups : {0}" -f $S_TotalGroups)
    Write-Host ("Public groups              : {0}" -f $S_PublicCount)
    Write-Host ("Private groups             : {0}" -f $S_PrivateCount)
    Write-Host ("Teams with no owner        : {0}" -f $S_NoOwnerCount)
    Write-Host ("Teams with single owner    : {0}" -f $S_SingleOwnerCount)
    Write-Host ("Public + no owner (critical): {0}" -f $S_PublicNoOwnerCount) -ForegroundColor $(if ($S_PublicNoOwnerCount -gt 0) { 'Red' } else { 'Gray' })
    Write-Host ("CSV report exported to     : {0}" -f $S_OutputPath)
    Write-Host ("HTML report exported to    : {0}" -f $S_HtmlOutputPath)

    # -----------------------------------------------------------------------
    # Disconnection handling
    # -----------------------------------------------------------------------

    if ($S_DisconnectGraphSessionOnExit)
    {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    else
    {
        $S_DisconnectChoice = Read-Host "Disconnect from Microsoft Graph? (Y/N)"
        if ($S_DisconnectChoice -match '^(y|yes)$')
        {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
    }
}
catch
{
    Write-Error $_
    exit 1
}
