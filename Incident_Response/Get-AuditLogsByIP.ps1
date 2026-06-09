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

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement))
{
    throw "ExchangeOnlineManagement module is not installed. Install it using Install-Module ExchangeOnlineManagement -Scope CurrentUser."
}

Import-Module ExchangeOnlineManagement -ErrorAction Stop

$S_ExistingConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
if ($S_ExistingConnection)
{
    Write-Host "Existing Exchange Online session detected:" -ForegroundColor Yellow
    Write-Host "  Account     : $($S_ExistingConnection.UserPrincipalName)" -ForegroundColor Yellow
    Write-Host "  Organization: $($S_ExistingConnection.Organization)" -ForegroundColor Yellow
    Write-Host "  TenantId    : $($S_ExistingConnection.TenantId)" -ForegroundColor Yellow
    Write-Host ""

    $S_Choice = Read-Host "Use existing session? [Y] Yes  [N] Disconnect and reconnect  (Default: Y)"
    if ($S_Choice -eq 'N')
    {
        Write-Host "Disconnecting existing session..." -ForegroundColor Cyan
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        Write-Host "Reconnecting to Exchange Online..." -ForegroundColor Cyan
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Host "Connected to Exchange Online." -ForegroundColor Green
    }
    else
    {
        Write-Host "Using existing Exchange Online session." -ForegroundColor Green
    }
}
else
{
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

try
{
    # ---------------------------------------------------------------------------
    # Search with pagination (max 5000 records per call, loop until exhausted)
    # ---------------------------------------------------------------------------

    Write-Host "Searching audit logs..." -ForegroundColor Cyan
    Write-Host "  IP(s)       : $($IPAddresses -join ', ')"
    Write-Host "  Date range  : $FromDate  ->  $ToDate"
    if ($UserUPNs)
    {
        Write-Host "  User(s)     : $($UserUPNs -join ', ')"
    }

    if ($Operations)
    {
        Write-Host "  Operation(s): $($Operations -join ', ')"
    }

    $S_SessionId = [System.Guid]::NewGuid().ToString()
    $S_AllRecords = [System.Collections.Generic.List[object]]::new()

    do
    {
        $S_SearchParams = @{
            StartDate      = $FromDate
            EndDate        = $ToDate
            IPAddresses    = $IPAddresses
            ResultSize     = 5000
            SessionId      = $S_SessionId
            SessionCommand = 'ReturnLargeSet'
            ErrorAction    = 'Stop'
        }
        if ($UserUPNs)
        {
            $S_SearchParams['UserIds'] = $UserUPNs
        }

        if ($Operations)
        {
            $S_SearchParams['Operations'] = $Operations
        }

        $S_Batch = Search-UnifiedAuditLog @S_SearchParams
        if ($S_Batch)
        {
            $S_AllRecords.AddRange([object[]]$S_Batch)
            Write-Host "  Retrieved $($S_AllRecords.Count) record(s) so far..." -ForegroundColor Gray
        }
    } while ($S_Batch -and $S_Batch.Count -eq 5000)

    if ($S_AllRecords.Count -eq 0)
    {
        Write-Host "`nNo records found for the specified criteria." -ForegroundColor Yellow
        return
    }

    Write-Host "`nTotal records found: $($S_AllRecords.Count)" -ForegroundColor Green

    # ---------------------------------------------------------------------------
    # Parse each record — operation-specific fields + raw fallback
    # ---------------------------------------------------------------------------

    $S_ParsedRecords = foreach ($S_Entry in $S_AllRecords)
    {
        $S_AuditData = $null
        try
        {
            $S_AuditData = $S_Entry.AuditData | ConvertFrom-Json
        }
        catch
        {
        }

        $S_Row = [ordered]@{
            CreationDate      = $S_Entry.CreationDate
            UserId            = $S_Entry.UserIds
            Operation         = $S_Entry.Operations
            ClientIP          = if ($S_AuditData.ClientIP) { $S_AuditData.ClientIP }
                                elseif ($S_AuditData.ActorIpAddress) { $S_AuditData.ActorIpAddress }
                                else { $null }
            RecordType        = $S_Entry.RecordType
            ResultStatus      = $S_AuditData.ResultStatus
            EmailSubject      = $null
            EmailFolder       = $null
            SourceFolder      = $null
            DestinationFolder = $null
            Attachments       = $null
            RuleContent       = $null
            RawAuditData      = $null
        }

        switch ($S_Entry.Operations)
        {
            { $_ -in 'Update', 'Create' } {
                $S_Item = $S_AuditData.Item
                $S_Row['EmailSubject'] = $S_Item.Subject
                $S_Row['EmailFolder'] = $S_Item.ParentFolder.Path
                $S_Row['Attachments'] = $S_Item.Attachments | ConvertTo-Json -Compress
            }
            { $_ -in 'MoveToDeletedItems', 'HardDelete', 'SoftDelete' } {
                $S_Affected = $S_AuditData.AffectedItems
                $S_Row['EmailSubject'] = $S_Affected.Subject
                $S_Row['EmailFolder'] = $S_Affected.ParentFolder.Path
            }
            'Move' {
                $S_Affected = $S_AuditData.AffectedItems
                $S_Row['EmailSubject'] = $S_Affected.Subject
                $S_Row['SourceFolder'] = $S_AuditData.Folder.Path
                $S_Row['DestinationFolder'] = $S_AuditData.DestFolder.Path
            }
            'New-InboxRule' {
                $S_Row['RuleContent'] = $S_AuditData.Parameters | ConvertTo-Json -Compress
            }
            default {
                $S_Row['RawAuditData'] = $S_Entry.AuditData
            }
        }

        [PSCustomObject]$S_Row
    }

    # ---------------------------------------------------------------------------
    # Export
    # ---------------------------------------------------------------------------

    if (-not (Test-Path -LiteralPath $OutputPath))
    {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    $S_ReportTime = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
    $S_ReportDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $S_IpTag = ($IPAddresses -join '_') -replace '[^\w\-]', '-'

    $S_CsvName = "AuditLogsByIP_${S_IpTag}_${S_ReportTime}.csv"
    $S_CsvPath = Join-Path -Path $OutputPath -ChildPath $S_CsvName
    $S_ParsedRecords | Export-Csv -LiteralPath $S_CsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nCSV report saved to: $S_CsvPath" -ForegroundColor Green

    # ---------------------------------------------------------------------------
    # Build simple HTML report
    # ---------------------------------------------------------------------------

    $S_OpsSummary = $S_ParsedRecords | Group-Object -Property Operation | Sort-Object Count -Descending
    $S_UserSummary = $S_ParsedRecords | Group-Object -Property UserId | Sort-Object Count -Descending | Select-Object -First 20

    $S_OpsRows = ($S_OpsSummary | ForEach-Object {
        "<tr><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.Name))</td><td>$($_.Count)</td></tr>"
    }) -join "`n"

    $S_UserRows = ($S_UserSummary | ForEach-Object {
        "<tr><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.Name))</td><td>$($_.Count)</td></tr>"
    }) -join "`n"

    $S_DetailRows = ($S_ParsedRecords | ForEach-Object {
        $S_Created = if ($_.CreationDate) { ([datetime]$_.CreationDate).ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
        "<tr><td>$S_Created</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.UserId))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.Operation))</td><td><code>$([System.Net.WebUtility]::HtmlEncode([string]$_.ClientIP))</code></td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.ResultStatus))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.EmailSubject))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.EmailFolder))</td></tr>"
    }) -join "`n"

    $S_IpListHtml = ($IPAddresses | ForEach-Object { "<span class='ip-chip'>$([System.Net.WebUtility]::HtmlEncode($_))</span>" }) -join ' '

    $S_Html = @"
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
            <div class="lbl">Generated</div><div class="val">$S_ReportDate</div>
            <div class="lbl">Tenant</div><div class="val">$([System.Net.WebUtility]::HtmlEncode($S_TenantName))</div>
            <div class="lbl">Tenant ID</div><div class="val"><code style="background:rgba(255,255,255,0.1);color:#fff;">$S_TenantId</code></div>
            <div class="lbl">Period</div><div class="val">$($FromDate.ToString('yyyy-MM-dd HH:mm')) &rarr; $($ToDate.ToString('yyyy-MM-dd HH:mm'))</div>
            <div class="lbl">IP(s)</div><div class="val">$S_IpListHtml</div>
        </div>
    </div>

    <div class="summary-cards">
        <div class="card">
            <div class="label">Total Records</div>
            <div class="value">$($S_ParsedRecords.Count)</div>
        </div>
        <div class="card">
            <div class="label">Unique Users</div>
            <div class="value">$(($S_ParsedRecords | Select-Object -ExpandProperty UserId -Unique).Count)</div>
        </div>
        <div class="card">
            <div class="label">Unique Operations</div>
            <div class="value">$($S_OpsSummary.Count)</div>
        </div>
        <div class="card">
            <div class="label">Unique Client IPs</div>
            <div class="value">$(($S_ParsedRecords | Select-Object -ExpandProperty ClientIP -Unique | Where-Object { $_ }).Count)</div>
        </div>
    </div>

    <div class="section">
        <h2>Operations Summary</h2>
        <table>
            <thead><tr><th>Operation</th><th>Count</th></tr></thead>
            <tbody>
$S_OpsRows
            </tbody>
        </table>
    </div>

    <div class="section">
        <h2>Top 20 Users</h2>
        <table>
            <thead><tr><th>User</th><th>Count</th></tr></thead>
            <tbody>
$S_UserRows
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
$S_DetailRows
            </tbody>
        </table>
    </div>

    <div class="footer">
        CSV data: $(Split-Path $S_CsvPath -Leaf) | Report generated by Get-AuditLogsByIP.ps1
    </div>
</body>
</html>
"@

    $S_HtmlName = "AuditLogsByIP_${S_IpTag}_${S_ReportTime}.html"
    $S_HtmlPath = Join-Path -Path $OutputPath -ChildPath $S_HtmlName
    $S_Html | Out-File -FilePath $S_HtmlPath -Encoding UTF8
    Write-Host "HTML report saved to: $S_HtmlPath" -ForegroundColor Green

    Write-Host "`nOperations summary:" -ForegroundColor Cyan
    $S_OpsSummary | Select-Object Name, Count | Format-Table -AutoSize
}
finally
{
    $S_DisconnectChoice = Read-Host "`nDisconnect from Exchange Online? [Y] Yes  [N] Keep session  (Default: N)"
    if ($S_DisconnectChoice -eq 'Y')
    {
        Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        Write-Host "Disconnected." -ForegroundColor Green
    }
    else
    {
        Write-Host "Exchange Online session kept alive." -ForegroundColor Green
    }
}
