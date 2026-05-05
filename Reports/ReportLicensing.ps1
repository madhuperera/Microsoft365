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

if (-not $OutputPath) {
    $S_Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path (Get-Location).Path "ReportLicensing_$S_Timestamp.csv"
}

$S_HtmlPath = [System.IO.Path]::ChangeExtension($OutputPath, '.html')

# ── Connect to Microsoft Graph ─────────────────────────────────────────────────
$S_RequiredGraphScopes = @(
    'Organization.Read.All'
)

$S_GraphRequestDelayMilliseconds = 5

$S_ExistingContext = Get-MgContext
if ($S_ExistingContext) {
    Write-Host "Existing Graph session detected:" -ForegroundColor Yellow
    Write-Host "  Account : $($S_ExistingContext.Account)" -ForegroundColor Yellow
    Write-Host "  TenantId: $($S_ExistingContext.TenantId)" -ForegroundColor Yellow
    Write-Host "  Scopes  : $($S_ExistingContext.Scopes -join ', ')" -ForegroundColor Yellow
    Write-Host ""

    $choice = Read-Host "Use existing session? [Y] Yes  [N] Disconnect and reconnect  (Default: Y)"
    if ($choice -eq 'N') {
        Write-Host "Disconnecting existing session..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Write-Host "Reconnecting with required scopes..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome
        Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
    }
    else {
        Write-Host "Using existing Graph session." -ForegroundColor Green
    }
}
else {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome
    Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
}

try {
    # ── Tenant info ────────────────────────────────────────────────────────────
    $context = Get-MgContext
    $tenantId = if ($context.TenantId) { $context.TenantId } else { "Unknown" }
    $tenantDisplayName = $null
    try {
        $org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
        $tenantDisplayName = $org.DisplayName
    } catch { }
    if (-not $tenantDisplayName) { $tenantDisplayName = $tenantId }

    # ── Download Microsoft SKU reference for friendly names ────────────────────
    $skuDisplayNameLookup = @{}
    $msRefUrl = 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv'
    try {
        Write-Host "Downloading Microsoft SKU reference data..." -ForegroundColor Cyan
        $refCsv = Invoke-WebRequest -Uri $msRefUrl -UseBasicParsing -ErrorAction Stop
        $refText = [System.Text.Encoding]::UTF8.GetString($refCsv.Content)
        $refData = $refText | ConvertFrom-Csv
        foreach ($row in $refData) {
            if ($row.String_Id -and -not $skuDisplayNameLookup.ContainsKey($row.String_Id)) {
                $skuDisplayNameLookup[$row.String_Id] = $row.Product_Display_Name
            }
        }
        Write-Host "  Loaded $($skuDisplayNameLookup.Count) SKU display names" -ForegroundColor Green
    } catch {
        Write-Warning "Could not download SKU reference CSV. Friendly names will fall back to SkuPartNumber."
    }

    # ── Fetch subscribed SKUs ──────────────────────────────────────────────────
    Write-Host "Fetching subscribed license plans..." -ForegroundColor Cyan
    $subscribedSkus = Get-MgSubscribedSku -All -ErrorAction Stop
    Write-Host "  Found $($subscribedSkus.Count) license plans" -ForegroundColor Green

    # ── Build report data ──────────────────────────────────────────────────────
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($sku in $subscribedSkus) {
        $enabled   = $sku.PrepaidUnits.Enabled
        $warning   = $sku.PrepaidUnits.Warning
        $suspended = $sku.PrepaidUnits.Suspended
        $lockedOut = $sku.PrepaidUnits.LockedOut
        $consumed  = $sku.ConsumedUnits
        $available = $enabled - $consumed

        $servicePlanNames = ($sku.ServicePlans | Sort-Object ServicePlanName | ForEach-Object { $_.ServicePlanName }) -join '; '
        $servicePlanCount = ($sku.ServicePlans | Measure-Object).Count

        # Calculate per-SKU utilisation
        $utilisationSkuPct = if ($enabled -gt 0) { [math]::Round(($consumed / $enabled) * 100, 1) } else { 0 }

        # Resolve friendly display name
        $displayName = $skuDisplayNameLookup[$sku.SkuPartNumber]
        if (-not $displayName) { $displayName = $sku.SkuPartNumber }

        # Detect free/trial SKUs from friendly name or SkuPartNumber patterns
        $isFreeOrTrial = (
            $displayName -match '\b(free|trial|viral)\b' -or
            $sku.SkuPartNumber -match '(FREE|TRIAL|VIRAL)'
        )

        $results.Add([PSCustomObject]@{
            DisplayName      = $displayName
            SkuPartNumber    = $sku.SkuPartNumber
            SkuId            = $sku.SkuId
            IsFreeOrTrial    = $isFreeOrTrial
            AppliesTo        = $sku.AppliesTo
            CapabilityStatus = $sku.CapabilityStatus
            Enabled          = $enabled
            Consumed         = $consumed
            Available        = $available
            UtilisationPct   = $utilisationSkuPct
            Warning          = $warning
            Suspended        = $suspended
            LockedOut        = $lockedOut
            ServicePlanCount = $servicePlanCount
            ServicePlans     = $servicePlanNames
        })
    }

    # Filter out free/trial if requested
    if ($ExcludeFree) {
        $excludedCount = ($results | Where-Object { $_.IsFreeOrTrial }).Count
        $results = [System.Collections.Generic.List[PSCustomObject]]($results | Where-Object { -not $_.IsFreeOrTrial })
        if ($excludedCount -gt 0) {
            Write-Host "  Excluded $excludedCount free/trial SKU(s)" -ForegroundColor Yellow
        }
    }

    # Sort by DisplayName
    $results = $results | Sort-Object DisplayName

    # ── Export CSV ─────────────────────────────────────────────────────────────
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nCSV exported to: $OutputPath" -ForegroundColor Green
    Write-Host "Total SKUs: $($results.Count)" -ForegroundColor Green

    # ── Calculate Statistics ───────────────────────────────────────────────────
    $totalSkus        = $results.Count
    $enabledSkus      = ($results | Where-Object { $_.CapabilityStatus -eq 'Enabled' }).Count
    $warningSkus      = ($results | Where-Object { $_.CapabilityStatus -eq 'Warning' }).Count
    $suspendedSkus    = ($results | Where-Object { $_.CapabilityStatus -eq 'Suspended' }).Count
    $disabledSkus     = ($results | Where-Object { $_.CapabilityStatus -notin @('Enabled', 'Warning', 'Suspended') }).Count
    $totalEnabled     = ($results | Measure-Object -Property Enabled -Sum).Sum
    $totalConsumed    = ($results | Measure-Object -Property Consumed -Sum).Sum
    $totalAvailable   = ($results | Measure-Object -Property Available -Sum).Sum
    $userSkus         = ($results | Where-Object { $_.AppliesTo -eq 'User' }).Count
    $companySkus      = ($results | Where-Object { $_.AppliesTo -eq 'Company' }).Count
    $overAllocated    = ($results | Where-Object { $_.Available -lt 0 }).Count
    $fullyConsumed    = ($results | Where-Object { $_.Available -eq 0 -and $_.Enabled -gt 0 }).Count
    $freeTrialSkus    = ($results | Where-Object { $_.IsFreeOrTrial }).Count
    $paidSkus         = $totalSkus - $freeTrialSkus

    # Calculate utilisation on paid SKUs only (free/trial inflate with 10000+ enabled units)
    $paidResults        = $results | Where-Object { -not $_.IsFreeOrTrial }
    $paidEnabled        = ($paidResults | Measure-Object -Property Enabled -Sum).Sum
    $paidConsumed       = ($paidResults | Measure-Object -Property Consumed -Sum).Sum
    $paidAvailable      = ($paidResults | Measure-Object -Property Available -Sum).Sum
    $utilisationPct     = if ($paidEnabled -gt 0) { [math]::Round(($paidConsumed / $paidEnabled) * 100, 1) } else { 0 }

    # ── Console summary ───────────────────────────────────────────────────────
    Write-Host ""
    Write-Host "Licensing Summary — $tenantDisplayName" -ForegroundColor Cyan
    Write-Host "--------------------------------------------"
    Write-Host ("  Total SKUs       : {0}" -f $totalSkus)
    Write-Host ("  Enabled SKUs     : {0}" -f $enabledSkus) -ForegroundColor Green
    if ($warningSkus -gt 0)  { Write-Host ("  Warning SKUs     : {0}" -f $warningSkus) -ForegroundColor Yellow }
    if ($suspendedSkus -gt 0){ Write-Host ("  Suspended SKUs   : {0}" -f $suspendedSkus) -ForegroundColor Red }
    Write-Host ("  Paid Licenses    : {0} enabled / {1} consumed / {2} available" -f $paidEnabled, $paidConsumed, $paidAvailable)
    Write-Host ("  Utilisation      : {0}% (paid SKUs only)" -f $utilisationPct)
    if ($freeTrialSkus -gt 0) { Write-Host ("  Free/Trial SKUs  : {0}" -f $freeTrialSkus) -ForegroundColor Yellow }
    if ($overAllocated -gt 0) { Write-Host ("  Over-allocated   : {0} SKU(s)" -f $overAllocated) -ForegroundColor Red }
    Write-Host ""

    # ── Build HTML table rows ──────────────────────────────────────────────────
    $tableRows = ($results | ForEach-Object {
        $statusBadge = switch ($_.CapabilityStatus) {
            'Enabled'   { '<span class="badge badge-enabled">Enabled</span>' }
            'Warning'   { '<span class="badge badge-warning">Warning</span>' }
            'Suspended' { '<span class="badge badge-suspended">Suspended</span>' }
            'LockedOut' { '<span class="badge badge-suspended">LockedOut</span>' }
            'Deleted'   { '<span class="badge badge-suspended">Deleted</span>' }
            default     { "<span class=`"badge`">$([System.Web.HttpUtility]::HtmlEncode($_.CapabilityStatus))</span>" }
        }
        $freeBadge = if ($_.IsFreeOrTrial) { ' <span class="badge badge-free">Free/Trial</span>' } else { '' }
        $availableClass = if ($_.Available -lt 0) { ' class="warn"' } elseif ($_.Available -eq 0 -and $_.Enabled -gt 0) { ' class="warn-amber"' } else { '' }
        $utilClass = if ($_.Enabled -eq 0) { '' } elseif ($_.UtilisationPct -gt 100) { 'util-red' } elseif ($_.UtilisationPct -ge 90) { 'util-green' } elseif ($_.UtilisationPct -ge 10) { 'util-mid' } else { 'util-low' }
        $nameClass = if ($utilClass) { " class=`"$utilClass`"" } else { '' }
        $utilTdClass = if ($utilClass) { " class=`"$utilClass`"" } else { '' }
        $utilDisplay = if ($_.Enabled -eq 0) { '-' } else { "$($_.UtilisationPct)%" }
        $freeAttr = if ($_.IsFreeOrTrial) { ' data-free="true"' } else { ' data-free="false"' }
        "        <tr$freeAttr><td$nameClass>$([System.Web.HttpUtility]::HtmlEncode($_.DisplayName))$freeBadge</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.SkuPartNumber))</td><td>$statusBadge</td><td>$($_.AppliesTo)</td><td>$($_.Enabled)</td><td>$($_.Consumed)</td><td$availableClass>$($_.Available)</td><td$utilTdClass>$utilDisplay</td><td>$($_.Warning)</td><td>$($_.Suspended)</td><td>$($_.ServicePlanCount)</td></tr>"
    }) -join "`n"

    # ── Generate HTML Report ───────────────────────────────────────────────────
    $reportDate = Get-Date -Format 'dd MMM yyyy HH:mm'

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>M365 Licensing Report — $([System.Web.HttpUtility]::HtmlEncode($tenantDisplayName))</title>
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
        <div class="subtitle">Tenant: $([System.Web.HttpUtility]::HtmlEncode($tenantDisplayName)) ($tenantId) | Generated: $reportDate</div>
    </div>

    <div class="cards">
        <div class="card blue">
            <div class="label">Total SKUs</div>
            <div class="value">$totalSkus</div>
            <div class="detail">$paidSkus paid / $freeTrialSkus free or trial</div>
        </div>
        <div class="card green">
            <div class="label">Paid Licenses</div>
            <div class="value">$paidEnabled</div>
            <div class="detail">Enabled across paid plans</div>
        </div>
        <div class="card teal">
            <div class="label">Consumed (Paid)</div>
            <div class="value">$paidConsumed</div>
            <div class="detail">$utilisationPct% utilisation</div>
        </div>
        <div class="card $(if ($paidAvailable -lt 0) { 'red' } elseif ($paidAvailable -eq 0) { 'orange' } else { 'green' })">
            <div class="label">Available (Paid)</div>
            <div class="value">$paidAvailable</div>
        </div>
        <div class="card $(if ($overAllocated -gt 0) { 'red' } else { 'green' })">
            <div class="label">Over-Allocated</div>
            <div class="value">$overAllocated</div>
            <div class="detail">SKU(s) exceeding entitlements</div>
        </div>
        <div class="card $(if ($enabledSkus -eq $totalSkus) { 'green' } else { 'orange' })">
            <div class="label">Enabled SKUs</div>
            <div class="value">$enabledSkus</div>
            <div class="detail">$(if ($warningSkus -gt 0) { "$warningSkus warning" } else { 'All healthy' })</div>
        </div>
    </div>

    <div class="section">
        <h2>Subscribed License Plans</h2>
        <div class="filter-bar">
            <div class="search-box">
                <input type="text" id="searchInput" placeholder="Filter by name or SKU..." onkeyup="filterTable()">
            </div>
            <label><input type="checkbox" id="hideFree" onchange="filterTable()"> Hide Free / Trial SKUs ($freeTrialSkus)</label>
        </div>
        <table id="skuTable">
            <thead>
                <tr><th onclick="sortTable(0,'text')">License Plan <span class="sort-arrow">&udarr;</span></th><th onclick="sortTable(1,'text')">SKU Part Number <span class="sort-arrow">&udarr;</span></th><th onclick="sortTable(2,'text')">Status <span class="sort-arrow">&udarr;</span></th><th onclick="sortTable(3,'text')">Applies To <span class="sort-arrow">&udarr;</span></th><th onclick="sortTable(4,'num')">Enabled <span class="sort-arrow">&udarr;</span></th><th onclick="sortTable(5,'num')">Consumed <span class="sort-arrow">&udarr;</span></th><th onclick="sortTable(6,'num')">Available <span class="sort-arrow">&udarr;</span></th><th onclick="sortTable(7,'num')">Utilisation <span class="sort-arrow">&udarr;</span></th><th onclick="sortTable(8,'num')">Warning <span class="sort-arrow">&udarr;</span></th><th onclick="sortTable(9,'num')">Suspended <span class="sort-arrow">&udarr;</span></th><th onclick="sortTable(10,'num')">Service Plans <span class="sort-arrow">&udarr;</span></th></tr>
            </thead>
            <tbody>
$tableRows
            </tbody>
        </table>
    </div>

    <div class="footer">
        CSV data: $(Split-Path $OutputPath -Leaf) | Report generated by ReportLicensing.ps1
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

    $html | Out-File -FilePath $S_HtmlPath -Encoding UTF8
    Write-Host "HTML report exported to: $S_HtmlPath" -ForegroundColor Green

    # ── Output file paths ─────────────────────────────────────────────────────
    Write-Host ""
    Write-Host "Reports:" -ForegroundColor Cyan
    Write-Host "  CSV  : $OutputPath" -ForegroundColor Yellow
    Write-Host "  HTML : $S_HtmlPath" -ForegroundColor Yellow
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    $S_DisconnectChoice = Read-Host "`nDisconnect from Microsoft Graph? [Y] Yes  [N] Keep session  (Default: N)"
    if ($S_DisconnectChoice -eq 'Y') {
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Write-Host "Disconnected." -ForegroundColor Green
    }
    else {
        Write-Host "Graph session kept alive." -ForegroundColor Green
    }
}
