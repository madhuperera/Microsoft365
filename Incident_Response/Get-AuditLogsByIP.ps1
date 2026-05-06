#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Searches the Exchange Online unified audit log for activity from specified IP addresses.

.DESCRIPTION
    Connects to Exchange Online and searches the unified audit log for events matching
    one or more IP addresses within a defined date range. Supports optional filtering
    by user UPN and operation type. Parses operation-specific fields from audit records
    and exports results to CSV and HTML.
    Intended for use during active security investigations.

.PARAMETER IPAddresses
    One or more IP addresses to search for. Accepts multiple values.

.PARAMETER FromDate
    Start date (UTC) for the audit log search window.

.PARAMETER ToDate
    End date (UTC) for the audit log search window.

.PARAMETER UserUPNs
    Optional. Filter results to specific user UPNs. Omit to search all users.

.PARAMETER Operations
    Optional. Filter results to specific audit operations. Omit to return all operations.
    Common IR values: UserLoggedIn, MailItemsAccessed, New-InboxRule, Move,
    MoveToDeletedItems, HardDelete, SoftDelete, Update, Create, SendAs, SendOnBehalf.

.PARAMETER OutputPath
    Folder to save the CSV and HTML reports. The folder will be created if it does not exist.

.EXAMPLE
    .\Get-AuditLogsByIP.ps1 -IPAddresses "1.2.3.4" -FromDate "2025-01-01" -ToDate "2025-01-31" -OutputPath "C:\IR\Output"

.EXAMPLE
    .\Get-AuditLogsByIP.ps1 -IPAddresses "1.2.3.4","5.6.7.8" -FromDate "2025-01-01" -ToDate "2025-01-31" -Operations "MailItemsAccessed","New-InboxRule" -OutputPath "C:\IR\Output"
#>

[CmdletBinding()]
param (
    # One or more attacker IPs, e.g. -IPAddresses "1.2.3.4","5.6.7.8"
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$IPAddresses,

    # Start date (UTC) for the audit log search window.
    [Parameter(Mandatory = $true)]
    [datetime]$FromDate,

    # End date (UTC) for the audit log search window.
    [Parameter(Mandatory = $true)]
    [datetime]$ToDate,

    # Optional: filter to specific user(s). Omit to search ALL users.
    [Parameter(Mandatory = $false)]
    [string[]]$UserUPNs,

    # Optional: filter to specific operation(s). Omit to search ALL operations.
    # Common IR values: UserLoggedIn, MailItemsAccessed, New-InboxRule, Move,
    #   MoveToDeletedItems, HardDelete, SoftDelete, Update, Create, SendAs, SendOnBehalf
    [Parameter(Mandatory = $false)]
    [string[]]$Operations,

    # Folder to save the CSV and HTML reports.
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Connection
# ---------------------------------------------------------------------------

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    throw "ExchangeOnlineManagement module is not installed. Install it using Install-Module ExchangeOnlineManagement -Scope CurrentUser."
}

Import-Module ExchangeOnlineManagement -ErrorAction Stop

$S_ExistingConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
if ($S_ExistingConnection) {
    Write-Host "Existing Exchange Online session detected:" -ForegroundColor Yellow
    Write-Host "  Account     : $($S_ExistingConnection.UserPrincipalName)" -ForegroundColor Yellow
    Write-Host "  Organization: $($S_ExistingConnection.Organization)" -ForegroundColor Yellow
    Write-Host "  TenantId    : $($S_ExistingConnection.TenantId)" -ForegroundColor Yellow
    Write-Host ""

    $S_Choice = Read-Host "Use existing session? [Y] Yes  [N] Disconnect and reconnect  (Default: Y)"
    if ($S_Choice -eq 'N') {
        Write-Host "Disconnecting existing session..." -ForegroundColor Cyan
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        Write-Host "Reconnecting to Exchange Online..." -ForegroundColor Cyan
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Host "Connected to Exchange Online." -ForegroundColor Green
    }
    else {
        Write-Host "Using existing Exchange Online session." -ForegroundColor Green
    }
}
else {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "Connected to Exchange Online." -ForegroundColor Green
}

$S_Context = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
$S_TenantName = if ($S_Context -and $S_Context.Organization)    { [string]$S_Context.Organization }
              elseif ($S_Context -and $S_Context.DelegatedOrganization) { [string]$S_Context.DelegatedOrganization }
              else { 'Unknown' }
$S_TenantId   = if ($S_Context) { [string]$S_Context.TenantId }          else { 'Unknown' }
$S_Account    = if ($S_Context) { [string]$S_Context.UserPrincipalName } else { 'Unknown' }

Write-Host ""
Write-Host ("=" * 70) -ForegroundColor Cyan
Write-Host "  Tenant     : $S_TenantName"  -ForegroundColor Cyan
Write-Host "  Tenant ID  : $S_TenantId"    -ForegroundColor Cyan
Write-Host "  Account    : $S_Account"     -ForegroundColor Cyan
Write-Host ("=" * 70) -ForegroundColor Cyan
Write-Host ""

try {
    # ---------------------------------------------------------------------------
    # Search with pagination (max 5000 records per call, loop until exhausted)
    # ---------------------------------------------------------------------------

    Write-Host "Searching audit logs..." -ForegroundColor Cyan
    Write-Host "  IP(s)       : $($IPAddresses -join ', ')"
    Write-Host "  Date range  : $FromDate  ->  $ToDate"
    if ($UserUPNs)   { Write-Host "  User(s)     : $($UserUPNs -join ', ')" }
    if ($Operations) { Write-Host "  Operation(s): $($Operations -join ', ')" }

    $sessionId  = [System.Guid]::NewGuid().ToString()
    $allRecords = [System.Collections.Generic.List[object]]::new()

    do {
        $searchParams = @{
            StartDate      = $FromDate
            EndDate        = $ToDate
            IPAddresses    = $IPAddresses
            ResultSize     = 5000
            SessionId      = $sessionId
            SessionCommand = 'ReturnLargeSet'
            ErrorAction    = 'Stop'
        }
        if ($UserUPNs)   { $searchParams['UserIds']    = $UserUPNs }
        if ($Operations) { $searchParams['Operations'] = $Operations }

        $batch = Search-UnifiedAuditLog @searchParams
        if ($batch) {
            $allRecords.AddRange([object[]]$batch)
            Write-Host "  Retrieved $($allRecords.Count) record(s) so far..." -ForegroundColor Gray
        }
    } while ($batch -and $batch.Count -eq 5000)

    if ($allRecords.Count -eq 0) {
        Write-Host "`nNo records found for the specified criteria." -ForegroundColor Yellow
        return
    }

    Write-Host "`nTotal records found: $($allRecords.Count)" -ForegroundColor Green

    # ---------------------------------------------------------------------------
    # Parse each record — operation-specific fields + raw fallback
    # ---------------------------------------------------------------------------

    $parsedRecords = foreach ($entry in $allRecords) {
        $auditData = $null
        try { $auditData = $entry.AuditData | ConvertFrom-Json } catch { }

        $row = [ordered]@{
            CreationDate = $entry.CreationDate
            UserId       = $entry.UserIds
            Operation    = $entry.Operations
            ClientIP     = if ($auditData.ClientIP)            { $auditData.ClientIP }
                           elseif ($auditData.ActorIpAddress) { $auditData.ActorIpAddress }
                           else { $null }
            RecordType   = $entry.RecordType
            ResultStatus = $auditData.ResultStatus
            EmailSubject      = $null
            EmailFolder       = $null
            SourceFolder      = $null
            DestinationFolder = $null
            Attachments       = $null
            RuleContent       = $null
            RawAuditData      = $null
        }

        switch ($entry.Operations) {
            { $_ -in 'Update', 'Create' } {
                $item = $auditData.Item
                $row['EmailSubject'] = $item.Subject
                $row['EmailFolder']  = $item.ParentFolder.Path
                $row['Attachments']  = $item.Attachments | ConvertTo-Json -Compress
            }
            { $_ -in 'MoveToDeletedItems', 'HardDelete', 'SoftDelete' } {
                $affected = $auditData.AffectedItems
                $row['EmailSubject'] = $affected.Subject
                $row['EmailFolder']  = $affected.ParentFolder.Path
            }
            'Move' {
                $affected = $auditData.AffectedItems
                $row['EmailSubject']      = $affected.Subject
                $row['SourceFolder']      = $auditData.Folder.Path
                $row['DestinationFolder'] = $auditData.DestFolder.Path
            }
            'New-InboxRule' {
                $row['RuleContent'] = $auditData.Parameters | ConvertTo-Json -Compress
            }
            default {
                $row['RawAuditData'] = $entry.AuditData
            }
        }

        [PSCustomObject]$row
    }

    # ---------------------------------------------------------------------------
    # Export
    # ---------------------------------------------------------------------------

    if (-not (Test-Path -LiteralPath $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    $reportTime = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
    $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $ipTag      = ($IPAddresses -join '_') -replace '[^\w\-]', '-'

    $csvName = "AuditLogsByIP_${ipTag}_${reportTime}.csv"
    $csvPath = Join-Path -Path $OutputPath -ChildPath $csvName
    $parsedRecords | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nCSV report saved to: $csvPath" -ForegroundColor Green

    # ---------------------------------------------------------------------------
    # Build simple HTML report
    # ---------------------------------------------------------------------------

    $opsSummary = $parsedRecords | Group-Object -Property Operation | Sort-Object Count -Descending
    $userSummary = $parsedRecords | Group-Object -Property UserId | Sort-Object Count -Descending | Select-Object -First 20

    $opsRows = ($opsSummary | ForEach-Object {
        "<tr><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.Name))</td><td>$($_.Count)</td></tr>"
    }) -join "`n"

    $userRows = ($userSummary | ForEach-Object {
        "<tr><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.Name))</td><td>$($_.Count)</td></tr>"
    }) -join "`n"

    $detailRows = ($parsedRecords | ForEach-Object {
        $created = if ($_.CreationDate) { ([datetime]$_.CreationDate).ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        "<tr><td>$created</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.UserId))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.Operation))</td><td><code>$([System.Net.WebUtility]::HtmlEncode([string]$_.ClientIP))</code></td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.ResultStatus))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.EmailSubject))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.EmailFolder))</td></tr>"
    }) -join "`n"

    $ipListHtml = ($IPAddresses | ForEach-Object { "<span class='ip-chip'>$([System.Net.WebUtility]::HtmlEncode($_))</span>" }) -join ' '

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Audit Logs by IP Report</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, sans-serif; background: #f0f2f5; color: #333; padding: 30px; }
        .header { background: linear-gradient(135deg, #1a1a2e, #16213e); color: #fff; padding: 24px 32px; border-radius: 10px; margin-bottom: 24px; }
        .header h1 { font-size: 1.7em; margin-bottom: 12px; }
        .header .meta { display: grid; grid-template-columns: max-content 1fr; gap: 6px 14px; font-size: 0.9em; opacity: 0.95; }
        .header .meta .lbl { font-weight: 600; opacity: 0.7; text-transform: uppercase; font-size: 11px; letter-spacing: 0.5px; align-self: center; }
        .header .meta .val { word-break: break-all; }
        .ip-chip { display: inline-block; background: #ffd166; color: #1a1a2e; font-family: Consolas, 'Courier New', monospace; font-weight: 700; font-size: 14px; padding: 4px 10px; border-radius: 4px; margin: 2px 4px 2px 0; }
        .summary-cards { display: flex; gap: 16px; margin-bottom: 24px; flex-wrap: wrap; }
        .card { background: #fff; border-radius: 8px; padding: 18px 22px; flex: 1; min-width: 160px; box-shadow: 0 1px 4px rgba(0,0,0,0.06); border-left: 4px solid #0078d4; }
        .card .label { font-size: 11px; color: #666; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px; }
        .card .value { font-size: 28px; font-weight: 700; color: #1a1a2e; }
        .section { margin-bottom: 24px; }
        .section h2 { font-size: 16px; color: #1a1a2e; margin-bottom: 10px; padding-bottom: 6px; border-bottom: 2px solid #e0e0e0; }
        table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 6px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,0.06); }
        th { background: #0078d4; color: #fff; text-align: left; padding: 10px 12px; font-size: 11px; text-transform: uppercase; letter-spacing: 0.3px; }
        td { padding: 9px 12px; font-size: 12px; border-bottom: 1px solid #eee; vertical-align: top; }
        tbody tr:hover { background: #f5f5f5; }
        code { background: #f0f0f0; padding: 2px 6px; border-radius: 3px; font-family: Consolas, 'Courier New', monospace; font-size: 11px; }
        .footer { text-align: center; font-size: 11px; color: #999; margin-top: 24px; }
    </style>
</head>
<body>
    <div class="header">
        <h1>Audit Logs by IP Report</h1>
        <div class="meta">
            <div class="lbl">Generated</div><div class="val">$reportDate</div>
            <div class="lbl">Tenant</div><div class="val">$([System.Net.WebUtility]::HtmlEncode($S_TenantName))</div>
            <div class="lbl">Tenant ID</div><div class="val"><code style="background:rgba(255,255,255,0.1);color:#fff;">$S_TenantId</code></div>
            <div class="lbl">Period</div><div class="val">$($FromDate.ToString('yyyy-MM-dd HH:mm')) &rarr; $($ToDate.ToString('yyyy-MM-dd HH:mm'))</div>
            <div class="lbl">IP(s)</div><div class="val">$ipListHtml</div>
        </div>
    </div>

    <div class="summary-cards">
        <div class="card">
            <div class="label">Total Records</div>
            <div class="value">$($parsedRecords.Count)</div>
        </div>
        <div class="card">
            <div class="label">Unique Users</div>
            <div class="value">$(($parsedRecords | Select-Object -ExpandProperty UserId -Unique).Count)</div>
        </div>
        <div class="card">
            <div class="label">Unique Operations</div>
            <div class="value">$($opsSummary.Count)</div>
        </div>
        <div class="card">
            <div class="label">Unique Client IPs</div>
            <div class="value">$(($parsedRecords | Select-Object -ExpandProperty ClientIP -Unique | Where-Object { $_ }).Count)</div>
        </div>
    </div>

    <div class="section">
        <h2>Operations Summary</h2>
        <table>
            <thead><tr><th>Operation</th><th>Count</th></tr></thead>
            <tbody>
$opsRows
            </tbody>
        </table>
    </div>

    <div class="section">
        <h2>Top 20 Users</h2>
        <table>
            <thead><tr><th>User</th><th>Count</th></tr></thead>
            <tbody>
$userRows
            </tbody>
        </table>
    </div>

    <div class="section">
        <h2>Audit Log Details</h2>
        <table>
            <thead>
                <tr>
                    <th>Date (UTC)</th>
                    <th>User</th>
                    <th>Operation</th>
                    <th>Client IP</th>
                    <th>Result</th>
                    <th>Email Subject</th>
                    <th>Folder</th>
                </tr>
            </thead>
            <tbody>
$detailRows
            </tbody>
        </table>
    </div>

    <div class="footer">
        CSV data: $(Split-Path $csvPath -Leaf) | Report generated by Get-AuditLogsByIP.ps1
    </div>
</body>
</html>
"@

    $htmlName = "AuditLogsByIP_${ipTag}_${reportTime}.html"
    $S_HtmlPath = Join-Path -Path $OutputPath -ChildPath $htmlName
    $html | Out-File -FilePath $S_HtmlPath -Encoding UTF8
    Write-Host "HTML report saved to: $S_HtmlPath" -ForegroundColor Green

    Write-Host "`nOperations summary:" -ForegroundColor Cyan
    $opsSummary | Select-Object Name, Count | Format-Table -AutoSize
}
finally {
    $S_DisconnectChoice = Read-Host "`nDisconnect from Exchange Online? [Y] Yes  [N] Keep session  (Default: N)"
    if ($S_DisconnectChoice -eq 'Y') {
        Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        Write-Host "Disconnected." -ForegroundColor Green
    }
    else {
        Write-Host "Exchange Online session kept alive." -ForegroundColor Green
    }
}
