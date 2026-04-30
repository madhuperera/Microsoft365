#Requires -Modules Microsoft.Graph.Users

<#
.SYNOPSIS
    Reports all interactive sign-in events matching one or more specified IP addresses.

.DESCRIPTION
    Connects to Microsoft Graph and retrieves interactive sign-in events for the specified
    date range. Filters results to only those matching the supplied IP address list.
    For each matched sign-in, reports user, application, resource, location, and status
    information. Exports results to CSV.

.PARAMETER IPAddresses
    One or more IP addresses to match. Accepts multiple values.

.PARAMETER Days
    Number of days to look back from now (UTC). Supported range: 1 to 30.

.PARAMETER OutputPath
    Folder to save the CSV report. The folder will be created if it does not exist.

.EXAMPLE
    .\Get-InteractiveSignInsByIP.ps1 -IPAddresses "1.2.3.4" -Days 7 -OutputPath "C:\IR\Output"

.EXAMPLE
    .\Get-InteractiveSignInsByIP.ps1 -IPAddresses "1.2.3.4","5.6.7.8" -Days 30 -OutputPath "C:\IR\Output"
#>

[CmdletBinding()]
param (
    # One or more IP addresses to match, e.g. -IPAddresses "1.2.3.4","5.6.7.8"
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$IPAddresses,

    # Number of days to look back from now (UTC). Supported range: 1 to 30.
    [Parameter(Mandatory = $true)]
    [ValidateRange(1, 30)]
    [int]$Days,

    # Folder to save the CSV report.
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$S_RequiredGraphScopes = @(
    'AuditLog.Read.All'
)

# ---------------------------------------------------------------------------
# Connection
# ---------------------------------------------------------------------------

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    throw "Microsoft.Graph.Users module is not installed. Install it using Install-Module Microsoft.Graph -Scope CurrentUser."
}

Import-Module Microsoft.Graph.Users -ErrorAction Stop

$context = Get-MgContext -ErrorAction SilentlyContinue
if ($context) {
    Write-Host "Existing Graph session detected:" -ForegroundColor Yellow
    Write-Host "  Account : $($context.Account)" -ForegroundColor Yellow
    Write-Host "  TenantId: $($context.TenantId)" -ForegroundColor Yellow
    Write-Host "  Scopes  : $($context.Scopes -join ', ')" -ForegroundColor Yellow
    Write-Host ""

    $choice = Read-Host "Use existing session? [Y] Yes  [N] Disconnect and reconnect  (Default: Y)"
    if ($choice -eq 'N') {
        Write-Host "Disconnecting existing session..." -ForegroundColor Cyan
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Write-Host "Reconnecting with required scope..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome -ErrorAction Stop | Out-Null
        Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
    }
    else {
        Write-Host "Using existing Graph session." -ForegroundColor Green
    }
}
else {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome -ErrorAction Stop | Out-Null
    Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# Inputs and date range
# ---------------------------------------------------------------------------

try {

$normalizedIPs = $IPAddresses |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ } |
    Select-Object -Unique

if (-not $normalizedIPs -or $normalizedIPs.Count -eq 0) {
    throw 'No valid IP addresses were provided after normalization.'
}

$ipHashSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($ip in $normalizedIPs) {
    [void]$ipHashSet.Add($ip)
}

$toUtc = (Get-Date).ToUniversalTime()
$fromUtc = $toUtc.AddDays(-$Days)

$fromIso = $fromUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')
$toIso = $toUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')

Write-Host "`nSearching interactive sign-ins..." -ForegroundColor Cyan
Write-Host "  IP(s)      : $($normalizedIPs -join ', ')"
Write-Host "  Date range : $fromIso  ->  $toIso"
Write-Host "  Days       : $Days"

function Get-GraphNextLink {
    param(
        [Parameter(Mandatory = $false)]
        [object]$Response
    )

    if (-not $Response) {
        return $null
    }

    if ($Response.PSObject.Properties.Name -contains '@odata.nextLink') {
        return [string]$Response.'@odata.nextLink'
    }

    return $null
}

function Get-GraphPageValues {
    param(
        [Parameter(Mandatory = $false)]
        [object]$Response
    )

    if (-not $Response) {
        return @()
    }

    if ($Response.PSObject.Properties.Name -contains 'value') {
        return @($Response.value)
    }

    return @()
}

# ---------------------------------------------------------------------------
# Query sign-ins with pagination
# ---------------------------------------------------------------------------

$selectFields = @(
    'id',
    'createdDateTime',
    'userDisplayName',
    'userPrincipalName',
    'userId',
    'ipAddress',
    'appDisplayName',
    'clientAppUsed',
    'resourceDisplayName',
    'status',
    'isInteractive',
    'location',
    'conditionalAccessStatus'
) -join ','

# Server-side filter by date and interactive sign-ins.
$filter = "createdDateTime ge $fromIso and createdDateTime le $toIso and isInteractive eq true"
$encodedFilter = [uri]::EscapeDataString($filter)
$encodedSelect = [uri]::EscapeDataString($selectFields)

$nextLink = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$top=1000&`$filter=$encodedFilter&`$select=$encodedSelect"
$allSignIns = [System.Collections.Generic.List[object]]::new()
$fallbackUsed = $false

try {
    while ($nextLink) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextLink -OutputType PSObject
        $pageValues = Get-GraphPageValues -Response $response

        if ($pageValues.Count -gt 0) {
            $allSignIns.AddRange([object[]]$pageValues)
            Write-Host "  Retrieved $($allSignIns.Count) interactive sign-in record(s) so far..." -ForegroundColor Gray
        }

        $nextLink = Get-GraphNextLink -Response $response
        if ($nextLink) {
            [System.Threading.Thread]::Sleep(5)
        }
    }
}
catch {
    Write-Warning "Interactive server-side filter failed. Retrying with date-only filter and local interactive filtering."
    $fallbackUsed = $true
    $allSignIns.Clear()

    $fallbackFilter = "createdDateTime ge $fromIso and createdDateTime le $toIso"
    $encodedFallbackFilter = [uri]::EscapeDataString($fallbackFilter)
    $nextLink = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$top=1000&`$filter=$encodedFallbackFilter&`$select=$encodedSelect"

    while ($nextLink) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextLink -OutputType PSObject
        $pageValues = Get-GraphPageValues -Response $response

        if ($pageValues.Count -gt 0) {
            $allSignIns.AddRange([object[]]$pageValues)
            Write-Host "  Retrieved $($allSignIns.Count) total sign-in record(s) so far..." -ForegroundColor Gray
        }

        $nextLink = Get-GraphNextLink -Response $response
        if ($nextLink) {
            [System.Threading.Thread]::Sleep(5)
        }
    }
}

if ($allSignIns.Count -eq 0) {
    Write-Host "`nNo sign-ins found in the selected time window." -ForegroundColor Yellow
    return
}

# ---------------------------------------------------------------------------
# Filter by interactive + IP list and shape output
# ---------------------------------------------------------------------------

$matched = foreach ($entry in $allSignIns) {
    $isInteractive = $true
    if ($fallbackUsed) {
        $isInteractive = $entry.isInteractive -eq $true
    }

    $entryIp = $entry.ipAddress
    if (-not $isInteractive -or -not $entryIp -or -not $ipHashSet.Contains([string]$entryIp)) {
        continue
    }

    $statusErrorCode = $null
    $statusFailureReason = $null
    if ($entry.status) {
        $statusErrorCode = $entry.status.errorCode
        $statusFailureReason = $entry.status.failureReason
    }

    $city = $null
    $state = $null
    $country = $null
    if ($entry.location) {
        $city = $entry.location.city
        $state = $entry.location.state
        $country = $entry.location.countryOrRegion
    }

    [PSCustomObject]@{
        CreatedDateTime         = $entry.createdDateTime
        UserDisplayName         = $entry.userDisplayName
        UserPrincipalName       = $entry.userPrincipalName
        UserId                  = $entry.userId
        IPAddress               = $entryIp
        AppDisplayName          = $entry.appDisplayName
        ClientAppUsed           = $entry.clientAppUsed
        ResourceDisplayName     = $entry.resourceDisplayName
        IsInteractive           = $entry.isInteractive
        ConditionalAccessStatus = $entry.conditionalAccessStatus
        StatusErrorCode         = $statusErrorCode
        StatusFailureReason     = $statusFailureReason
        City                    = $city
        State                   = $state
        CountryOrRegion         = $country
        SignInId                = $entry.id
    }
}

if (-not $matched) {
    Write-Host "`nNo interactive sign-ins matched the supplied IP list." -ForegroundColor Yellow
    return
}

$matched = $matched | Sort-Object CreatedDateTime

# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$reportTime = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$ipTag = ($normalizedIPs -join '_') -replace '[^\w\-]', '-'
$csvName = "InteractiveSignInsByIP_${ipTag}_${reportTime}.csv"
$csvPath = Join-Path -Path $OutputPath -ChildPath $csvName

$matched | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "`nMatched records: $($matched.Count)" -ForegroundColor Green
Write-Host "Report saved to: $csvPath" -ForegroundColor Green

# Quick on-screen summary per IP.
$summary = $matched |
    Group-Object -Property IPAddress |
    Sort-Object -Property Name |
    Select-Object @{ Name = 'IPAddress'; Expression = { $_.Name } }, @{ Name = 'Count'; Expression = { $_.Count } }

Write-Host "`nSummary by IP:" -ForegroundColor Cyan
$summary | Format-Table -AutoSize

}
finally {
    $disconnectChoice = Read-Host "`nDisconnect from Microsoft Graph? [Y] Yes  [N] Keep session  (Default: N)"
    if ($disconnectChoice -eq 'Y') {
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Write-Host "Disconnected." -ForegroundColor Green
    }
    else {
        Write-Host "Graph session kept alive." -ForegroundColor Green
    }
}
