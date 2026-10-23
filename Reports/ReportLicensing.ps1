#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement

<#
.SYNOPSIS
    Reports on Microsoft 365 licensing plans (subscribed SKUs) in the tenant.
    Exports CSV and HTML.

.PARAMETER OutputPath
    Path for the output CSV file. Defaults to a timestamped file in the current directory.

.PARAMETER ExcludeFree
    When specified, excludes free and trial SKUs from the report.

.EXAMPLE
    .\ReportLicensing.ps1

.EXAMPLE
    .\ReportLicensing.ps1 -OutputPath "C:\Reports\Licensing.csv"

.EXAMPLE
    .\ReportLicensing.ps1 -ExcludeFree
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]$ExcludeFree
)

# ── Setup ──────────────────────────────────────────────────────────────────────
$ErrorActionPreference = 'Stop'

$S_OutputPath = $OutputPath

if (-not $S_OutputPath)
{
    $S_Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $S_OutputPath = Join-Path (Get-Location).Path "ReportLicensing_$S_Timestamp.csv"
}

$S_HtmlPath = [System.IO.Path]::ChangeExtension($S_OutputPath, '.html')

# ── Connect to Microsoft Graph ─────────────────────────────────────────────────
$S_RequiredGraphScopes = @(
    'Organization.Read.All'
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
        Write-Host "Disconnecting existing session..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Write-Host "Reconnecting with required scopes..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome
        Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
    }
    else
    {
        Write-Host "Using existing Graph session." -ForegroundColor Green
    }
}
else
{
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome
    Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
}

try
{
    # ── Tenant info ────────────────────────────────────────────────────────────
    $S_Context = Get-MgContext
    $S_TenantId = if ($S_Context.TenantId)
    {
        $S_Context.TenantId
    }
    else
    {
        'Unknown'
    }
    $S_TenantDisplayName = $null
    try
    {
        $S_Org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
        $S_TenantDisplayName = $S_Org.DisplayName
    }
    catch
    {
    }
    if (-not $S_TenantDisplayName)
    {
        $S_TenantDisplayName = $S_TenantId
    }

    # ── Download Microsoft SKU reference for friendly names ────────────────────
    $S_SkuDisplayNameLookup = @{}
    $S_MsRefUrl = 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv'
    try
    {
        Write-Host 'Downloading Microsoft SKU reference data...' -ForegroundColor Cyan
        $S_RefCsv = Invoke-WebRequest -Uri $S_MsRefUrl -UseBasicParsing -ErrorAction Stop
        $S_RefText = [System.Text.Encoding]::UTF8.GetString($S_RefCsv.Content)
        $S_RefData = $S_RefText | ConvertFrom-Csv
        foreach ($S_Row in $S_RefData)
        {
            if ($S_Row.String_Id -and -not $S_SkuDisplayNameLookup.ContainsKey($S_Row.String_Id))
            {
                $S_SkuDisplayNameLookup[$S_Row.String_Id] = $S_Row.Product_Display_Name
            }
        }
        Write-Host "  Loaded $($S_SkuDisplayNameLookup.Count) SKU display names" -ForegroundColor Green
    }
    catch
    {
        Write-Warning 'Could not download SKU reference CSV. Friendly names will fall back to SkuPartNumber.'
    }

    # ── Fetch subscribed SKUs ──────────────────────────────────────────────────
    Write-Host 'Fetching subscribed license plans...' -ForegroundColor Cyan
    $S_SubscribedSkus = Get-MgSubscribedSku -All -ErrorAction Stop
    Write-Host "  Found $($S_SubscribedSkus.Count) license plans" -ForegroundColor Green

    # ── Build report data ──────────────────────────────────────────────────────
    $S_Results = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($S_Sku in $S_SubscribedSkus)
    {
        $S_Enabled = $S_Sku.PrepaidUnits.Enabled
        $S_Warning = $S_Sku.PrepaidUnits.Warning
        $S_Suspended = $S_Sku.PrepaidUnits.Suspended
        $S_LockedOut = $S_Sku.PrepaidUnits.LockedOut
        $S_Consumed = $S_Sku.ConsumedUnits
        $S_Available = $S_Enabled - $S_Consumed

        $S_ServicePlanNames = ($S_Sku.ServicePlans | Sort-Object ServicePlanName | ForEach-Object { $_.ServicePlanName }) -join '; '
        $S_ServicePlanCount = ($S_Sku.ServicePlans | Measure-Object).Count

        # Calculate per-SKU utilisation
        $S_UtilisationSkuPct = if ($S_Enabled -gt 0)
        {
            [math]::Round(($S_Consumed / $S_Enabled) * 100, 1)
        }
        else
        {
            0
        }

        # Resolve friendly display name
        $S_DisplayName = $S_SkuDisplayNameLookup[$S_Sku.SkuPartNumber]
        if (-not $S_DisplayName)
        {
            $S_DisplayName = $S_Sku.SkuPartNumber
        }

        # Detect free/trial SKUs from friendly name or SkuPartNumber patterns
        $S_IsFreeOrTrial = (
            $S_DisplayName -match '\b(free|trial|viral)\b' -or
            $S_Sku.SkuPartNumber -match '(FREE|TRIAL|VIRAL)'
        )

        $S_Results.Add([PSCustomObject]@{
            DisplayName      = $S_DisplayName
            SkuPartNumber    = $S_Sku.SkuPartNumber
            SkuId            = $S_Sku.SkuId
            IsFreeOrTrial    = $S_IsFreeOrTrial
            AppliesTo        = $S_Sku.AppliesTo
            CapabilityStatus = $S_Sku.CapabilityStatus
            Enabled          = $S_Enabled
            Consumed         = $S_Consumed
            Available        = $S_Available
            UtilisationPct   = $S_UtilisationSkuPct
            Warning          = $S_Warning
            Suspended        = $S_Suspended
            LockedOut        = $S_LockedOut
            ServicePlanCount = $S_ServicePlanCount
            ServicePlans     = $S_ServicePlanNames
        })
    }

    # Filter out free/trial if requested
    if ($ExcludeFree)
    {
        $S_ExcludedCount = ($S_Results | Where-Object { $_.IsFreeOrTrial }).Count
        $S_Results = [System.Collections.Generic.List[PSCustomObject]]($S_Results | Where-Object { -not $_.IsFreeOrTrial })
        if ($S_ExcludedCount -gt 0)
        {
            Write-Host "  Excluded $S_ExcludedCount free/trial SKU(s)" -ForegroundColor Yellow
        }
    }

    # Sort by DisplayName
    $S_Results = $S_Results | Sort-Object DisplayName

    # ── Export CSV ─────────────────────────────────────────────────────────────
    $S_Results | Export-Csv -Path $S_OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nCSV exported to: $S_OutputPath" -ForegroundColor Green
    Write-Host "Total SKUs: $($S_Results.Count)" -ForegroundColor Green

    # ── Calculate Statistics ───────────────────────────────────────────────────
    $S_TotalSkus = $S_Results.Count
    $S_EnabledSkus = ($S_Results | Where-Object { $_.CapabilityStatus -eq 'Enabled' }).Count
    $S_WarningSkus = ($S_Results | Where-Object { $_.CapabilityStatus -eq 'Warning' }).Count
    $S_SuspendedSkus = ($S_Results | Where-Object { $_.CapabilityStatus -eq 'Suspended' }).Count
    $S_DisabledSkus = ($S_Results | Where-Object { $_.CapabilityStatus -notin @('Enabled', 'Warning', 'Suspended') }).Count
    $S_TotalEnabled = ($S_Results | Measure-Object -Property Enabled -Sum).Sum
    $S_TotalConsumed = ($S_Results | Measure-Object -Property Consumed -Sum).Sum
    $S_TotalAvailable = ($S_Results | Measure-Object -Property Available -Sum).Sum
    $S_UserSkus = ($S_Results | Where-Object { $_.AppliesTo -eq 'User' }).Count
    $S_CompanySkus = ($S_Results | Where-Object { $_.AppliesTo -eq 'Company' }).Count
    $S_OverAllocated = ($S_Results | Where-Object { $_.Available -lt 0 }).Count
    $S_FullyConsumed = ($S_Results | Where-Object { $_.Available -eq 0 -and $_.Enabled -gt 0 }).Count
    $S_FreeTrialSkus = ($S_Results | Where-Object { $_.IsFreeOrTrial }).Count
    $S_PaidSkus = $S_TotalSkus - $S_FreeTrialSkus

    # Calculate utilisation on paid SKUs only (free/trial inflate with 10000+ enabled units)
    $S_PaidResults = $S_Results | Where-Object { -not $_.IsFreeOrTrial }
    $S_PaidEnabled = ($S_PaidResults | Measure-Object -Property Enabled -Sum).Sum
    $S_PaidConsumed = ($S_PaidResults | Measure-Object -Property Consumed -Sum).Sum
    $S_PaidAvailable = ($S_PaidResults | Measure-Object -Property Available -Sum).Sum
    $S_UtilisationPct = if ($S_PaidEnabled -gt 0)
    {
        [math]::Round(($S_PaidConsumed / $S_PaidEnabled) * 100, 1)
    }
    else
    {
        0
    }

    $S_PaidAvailableCardClass = if ($S_PaidAvailable -lt 0)
    {
        'red'
    }
    elseif ($S_PaidAvailable -eq 0)
    {
        'orange'
    }
    else
    {
        'green'
    }
    $S_OverAllocatedCardClass = if ($S_OverAllocated -gt 0)
    {
        'red'
    }
    else
    {
        'green'
    }
    $S_EnabledSkusCardClass = if ($S_EnabledSkus -eq $S_TotalSkus)
    {
        'green'
    }
    else
    {
        'orange'
    }
    $S_EnabledSkusDetail = if ($S_WarningSkus -gt 0)
    {
        "$S_WarningSkus warning"
    }
    else
    {
        'All healthy'
    }

    # ── Console summary ───────────────────────────────────────────────────────
    Write-Host ''
    Write-Host "Licensing Summary — $S_TenantDisplayName" -ForegroundColor Cyan
    Write-Host '--------------------------------------------'
    Write-Host ("  Total SKUs       : {0}" -f $S_TotalSkus)
    Write-Host ("  Enabled SKUs     : {0}" -f $S_EnabledSkus) -ForegroundColor Green
    if ($S_WarningSkus -gt 0)
    {
        Write-Host ("  Warning SKUs     : {0}" -f $S_WarningSkus) -ForegroundColor Yellow
    }
    if ($S_SuspendedSkus -gt 0)
    {
        Write-Host ("  Suspended SKUs   : {0}" -f $S_SuspendedSkus) -ForegroundColor Red
    }
    Write-Host ("  Paid Licenses    : {0} enabled / {1} consumed / {2} available" -f $S_PaidEnabled, $S_PaidConsumed, $S_PaidAvailable)
    Write-Host ("  Utilisation      : {0}% (paid SKUs only)" -f $S_UtilisationPct)
    if ($S_FreeTrialSkus -gt 0)
    {
        Write-Host ("  Free/Trial SKUs  : {0}" -f $S_FreeTrialSkus) -ForegroundColor Yellow
    }
    if ($S_OverAllocated -gt 0)
    {
        Write-Host ("  Over-allocated   : {0} SKU(s)" -f $S_OverAllocated) -ForegroundColor Red
    }
    Write-Host ''

    # ── Build HTML table rows ──────────────────────────────────────────────────
    $S_TableRows = ($S_Results | ForEach-Object {
        $S_StatusBadge = switch ($_.CapabilityStatus)
        {
            'Enabled'   { '<span class="badge badge-enabled">Enabled</span>' }
            'Warning'   { '<span class="badge badge-warning">Warning</span>' }
            'Suspended' { '<span class="badge badge-suspended">Suspended</span>' }
            'LockedOut' { '<span class="badge badge-suspended">LockedOut</span>' }
            'Deleted'   { '<span class="badge badge-suspended">Deleted</span>' }
            default     { "<span class=`"badge`">$([System.Web.HttpUtility]::HtmlEncode($_.CapabilityStatus))</span>" }
        }
        $S_FreeBadge = if ($_.IsFreeOrTrial)
        {
            ' <span class="badge badge-free">Free/Trial</span>'
        }
        else
        {
            ''
        }
        $S_AvailableClass = if ($_.Available -lt 0)
        {
            ' class="warn"'
        }
        elseif ($_.Available -eq 0 -and $_.Enabled -gt 0)
        {
            ' class="warn-amber"'
        }
        else
        {
            ''
        }
        $S_UtilClass = if ($_.Enabled -eq 0)
        {
            ''
        }
        elseif ($_.UtilisationPct -gt 100)
        {
            'util-red'
        }
        elseif ($_.UtilisationPct -ge 90)
        {
            'util-green'
        }
        elseif ($_.UtilisationPct -ge 10)
        {
            'util-mid'
        }
        else
        {
            'util-low'
        }
        $S_NameClass = if ($S_UtilClass)
        {
            " class=`"$S_UtilClass`""
        }
        else
        {
            ''
        }
        $S_UtilTdClass = if ($S_UtilClass)
        {
            " class=`"$S_UtilClass`""
        }
        else
        {
            ''
        }
        $S_UtilDisplay = if ($_.Enabled -eq 0)
        {
            '-'
        }
        else
        {
            "$($_.UtilisationPct)%"
        }
        $S_FreeAttr = if ($_.IsFreeOrTrial)
        {
            ' data-free="true"'
        }
        else
        {
            ' data-free="false"'
        }
        "        <tr$S_FreeAttr><td$S_NameClass>$([System.Web.HttpUtility]::HtmlEncode($_.DisplayName))$S_FreeBadge</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.SkuPartNumber))</td><td>$S_StatusBadge</td><td>$($_.AppliesTo)</td><td>$($_.Enabled)</td><td>$($_.Consumed)</td><td$S_AvailableClass>$($_.Available)</td><td$S_UtilTdClass>$S_UtilDisplay</td><td>$($_.Warning)</td><td>$($_.Suspended)</td><td>$($_.ServicePlanCount)</td></tr>"
    }) -join "`n"

    # ── Generate HTML Report ───────────────────────────────────────────────────
    $S_ReportDate = Get-Date -Format 'dd MMM yyyy HH:mm'

    $S_Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>M365 Licensing Report — $([System.Web.HttpUtility]::HtmlEncode($S_TenantDisplayName))</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; padding: 24px; }
        .header { text-align: center; margin-bottom: 32px; }
        .header h1 { font-size: 28px; color: #1a1a2e; margin-bottom: 4px; }
        .header .subtitle { font-size: 14px; color: #666; }
        .cards { display: flex; flex-wrap: wrap; gap: 20px; justify-content: center; margin-bottom: 32px; }
        .card {
            background: #fff; border-radius: 12px; padding: 24px 28px; min-width: 200px; flex: 1; max-width: 260px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08); border-left: 5px solid #0078d4; position: relative;
        }
        .card .label { font-size: 13px; color: #666; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px; }
        .card .value { font-size: 36px; font-weight: 700; color: #1a1a2e; }
        .card .detail { font-size: 12px; color: #888; margin-top: 6px; }
        .card.blue    { border-left-color: #0078d4; }
        .card.green   { border-left-color: #107c10; }
        .card.red     { border-left-color: #d13438; }
        .card.orange  { border-left-color: #ff8c00; }
        .card.purple  { border-left-color: #8764b8; }
        .card.teal    { border-left-color: #00b7c3; }
        .section { margin-bottom: 24px; }
        .section h2 { font-size: 18px; color: #1a1a2e; margin-bottom: 12px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0; }
        table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.08); margin-top: 12px; }
        th { background: #0078d4; color: #fff; text-align: left; padding: 10px 14px; font-size: 13px; text-transform: uppercase; letter-spacing: 0.3px; cursor: pointer; user-select: none; white-space: nowrap; }
        th:hover { background: #106ebe; }
        th .sort-arrow { font-size: 10px; margin-left: 4px; opacity: 0.5; }
        th.sorted-asc .sort-arrow { opacity: 1; }
        th.sorted-desc .sort-arrow { opacity: 1; }
        td { padding: 9px 14px; font-size: 13px; border-bottom: 1px solid #eee; }
        tr:hover { background: #f5f9ff; }
        td.warn { color: #d13438; font-weight: 600; }
        td.warn-amber { color: #ff8c00; font-weight: 600; }
        td.util-red { background: #e81123; color: #fff; font-weight: 600; }
        td.util-low { background: #c4314b; color: #fff; font-weight: 600; }
        td.util-mid { background: #ca5010; color: #fff; font-weight: 600; }
        td.util-green { background: #107c10; color: #fff; font-weight: 600; }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
        .badge-enabled { background: #dff6dd; color: #107c10; }
        .badge-warning { background: #fff4ce; color: #8a6914; }
        .badge-suspended { background: #fde7e9; color: #d13438; }
        .badge-free { background: #e8e8e8; color: #666; }
        .filter-bar { display: flex; flex-wrap: wrap; gap: 16px; align-items: center; margin-bottom: 16px; }
        .filter-bar label { font-size: 14px; color: #555; cursor: pointer; user-select: none; }
        .filter-bar input[type="checkbox"] { margin-right: 4px; cursor: pointer; }
        .search-box { flex: 1; min-width: 200px; }
        .search-box input { padding: 10px 16px; width: 100%; max-width: 400px; border: 1px solid #ddd; border-radius: 8px; font-size: 14px; }
        .search-box input:focus { outline: none; border-color: #0078d4; box-shadow: 0 0 0 2px rgba(0,120,212,0.2); }
        .footer { text-align: center; font-size: 12px; color: #999; margin-top: 32px; }
    </style>
</head>
<body>
    <div class="header">
        <h1>M365 Licensing Report</h1>
        <div class="subtitle">Tenant: $([System.Web.HttpUtility]::HtmlEncode($S_TenantDisplayName)) ($S_TenantId) | Generated: $S_ReportDate</div>
    </div>

    <div class="cards">
        <div class="card blue">
            <div class="label">Total SKUs</div>
            <div class="value">$S_TotalSkus</div>
            <div class="detail">$S_PaidSkus paid / $S_FreeTrialSkus free or trial</div>
        </div>
        <div class="card green">
            <div class="label">Paid Licenses</div>
            <div class="value">$S_PaidEnabled</div>
            <div class="detail">Enabled across paid plans</div>
        </div>
        <div class="card teal">
            <div class="label">Consumed (Paid)</div>
            <div class="value">$S_PaidConsumed</div>
            <div class="detail">$S_UtilisationPct% utilisation</div>
        </div>
        <div class="card $S_PaidAvailableCardClass">
            <div class="label">Available (Paid)</div>
            <div class="value">$S_PaidAvailable</div>
        </div>
        <div class="card $S_OverAllocatedCardClass">
            <div class="label">Over-Allocated</div>
            <div class="value">$S_OverAllocated</div>
            <div class="detail">SKU(s) exceeding entitlements</div>
        </div>
        <div class="card $S_EnabledSkusCardClass">
            <div class="label">Enabled SKUs</div>
            <div class="value">$S_EnabledSkus</div>
            <div class="detail">$S_EnabledSkusDetail</div>
        </div>
    </div>

    <div class="section">
        <h2>Subscribed License Plans</h2>
        <div class="filter-bar">
            <div class="search-box">
                <input type="text" id="searchInput" placeholder="Filter by name or SKU..." onkeyup="filterTable()">
            </div>
            <label><input type="checkbox" id="hideFree" onchange="filterTable()"> Hide Free / Trial SKUs ($S_FreeTrialSkus)</label>
        </div>
        <table id="skuTable">
            <thead>
                <tr><th onclick="sortTable(0,'text')">License Plan <span class="sort-arrow">&udarr;</span></th><th onclick="sortTable(1,'text')">SKU Part Number <span class="sort-arrow">&udarr;</span></th><th onclick="sortTable(2,'text')">Status <span class="sort-arrow">&udarr;</span></th><th onclick="sortTable(3,'text')">Applies To <span class="sort-arrow">&udarr;</span></th><th onclick="sortTable(4,'num')">Enabled <span class="sort-arrow">&udarr;</span></th><th onclick="sortTable(5,'num')">Consumed <span class="sort-arrow">&udarr;</span></th><th onclick="sortTable(6,'num')">Available <span class="sort-arrow">&udarr;</span></th><th onclick="sortTable(7,'num')">Utilisation <span class="sort-arrow">&udarr;</span></th><th onclick="sortTable(8,'num')">Warning <span class="sort-arrow">&udarr;</span></th><th onclick="sortTable(9,'num')">Suspended <span class="sort-arrow">&udarr;</span></th><th onclick="sortTable(10,'num')">Service Plans <span class="sort-arrow">&udarr;</span></th></tr>
            </thead>
            <tbody>
$S_TableRows
            </tbody>
        </table>
    </div>

    <div class="footer">
        CSV data: $(Split-Path $S_OutputPath -Leaf) | Report generated by ReportLicensing.ps1
    </div>

    <script>
        function filterTable() {
            var input = document.getElementById('searchInput').value.toLowerCase();
            var hideFree = document.getElementById('hideFree').checked;
            var rows = document.querySelectorAll('#skuTable tbody tr');
            rows.forEach(function(row) {
                var displayName = row.cells[0].textContent.toLowerCase();
                var skuPart = row.cells[1].textContent.toLowerCase();
                var isFree = row.getAttribute('data-free') === 'true';
                var matchesSearch = displayName.indexOf(input) > -1 || skuPart.indexOf(input) > -1;
                var matchesFree = !hideFree || !isFree;
                row.style.display = (matchesSearch && matchesFree) ? '' : 'none';
            });
        }

        var currentSortCol = -1;
        var currentSortAsc = true;

        function sortTable(colIndex, type) {
            var table = document.getElementById('skuTable');
            var tbody = table.querySelector('tbody');
            var rows = Array.from(tbody.querySelectorAll('tr'));
            var headers = table.querySelectorAll('th');

            // Toggle direction if same column clicked again
            if (currentSortCol === colIndex) {
                currentSortAsc = !currentSortAsc;
            } else {
                currentSortAsc = true;
                currentSortCol = colIndex;
            }

            // Update header styling
            headers.forEach(function(h) { h.classList.remove('sorted-asc', 'sorted-desc'); });
            headers[colIndex].classList.add(currentSortAsc ? 'sorted-asc' : 'sorted-desc');

            rows.sort(function(a, b) {
                var aVal = a.cells[colIndex].textContent.trim();
                var bVal = b.cells[colIndex].textContent.trim();
                if (type === 'num') {
                    aVal = parseFloat(aVal) || 0;
                    bVal = parseFloat(bVal) || 0;
                    return currentSortAsc ? aVal - bVal : bVal - aVal;
                } else {
                    aVal = aVal.toLowerCase();
                    bVal = bVal.toLowerCase();
                    if (aVal < bVal) return currentSortAsc ? -1 : 1;
                    if (aVal > bVal) return currentSortAsc ? 1 : -1;
                    return 0;
                }
            });

            rows.forEach(function(row) { tbody.appendChild(row); });
        }
    </script>
</body>
</html>
"@

    $S_Html | Out-File -FilePath $S_HtmlPath -Encoding UTF8
    Write-Host "HTML report exported to: $S_HtmlPath" -ForegroundColor Green

    # ── Output file paths ─────────────────────────────────────────────────────
    Write-Host ''
    Write-Host 'Reports:' -ForegroundColor Cyan
    Write-Host "  CSV  : $S_OutputPath" -ForegroundColor Yellow
    Write-Host "  HTML : $S_HtmlPath" -ForegroundColor Yellow
}
catch
{
    Write-Error "An error occurred: $_"
}
finally
{
    $S_DisconnectChoice = Read-Host "`nDisconnect from Microsoft Graph? [Y] Yes  [N] Keep session  (Default: N)"
    if ($S_DisconnectChoice -eq 'Y')
    {
        Write-Host 'Disconnecting from Microsoft Graph...' -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Write-Host 'Disconnected.' -ForegroundColor Green
    }
    else
    {
        Write-Host 'Graph session kept alive.' -ForegroundColor Green
    }
}
