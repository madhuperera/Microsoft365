#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users

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
param(
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

$S_GraphRequestDelayMilliseconds = 5

# ---------------------------------------------------------------------------
# Connection
# ---------------------------------------------------------------------------

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users))
{
    throw "Microsoft.Graph.Users module is not installed. Install it using Install-Module Microsoft.Graph -Scope CurrentUser."
}

Import-Module Microsoft.Graph.Users -ErrorAction Stop

$S_Context = Get-MgContext -ErrorAction SilentlyContinue
if ($S_Context)
{
    Write-Host "Existing Graph session detected:" -ForegroundColor Yellow
    Write-Host "  Account : $($S_Context.Account)" -ForegroundColor Yellow
    Write-Host "  TenantId: $($S_Context.TenantId)" -ForegroundColor Yellow
    Write-Host "  Scopes  : $($S_Context.Scopes -join ', ')" -ForegroundColor Yellow
    Write-Host ""

    $S_Choice = Read-Host "Use existing session? [Y] Yes  [N] Disconnect and reconnect  (Default: Y)"
    if ($S_Choice -eq 'N')
    {
        Write-Host "Disconnecting existing session..." -ForegroundColor Cyan
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Write-Host "Reconnecting with required scope..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome -ErrorAction Stop | Out-Null
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
    Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome -ErrorAction Stop | Out-Null
    Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# Inputs and date range
# ---------------------------------------------------------------------------

try
{
    $S_NormalizedIPs = $IPAddresses |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ } |
        Select-Object -Unique

    if (-not $S_NormalizedIPs -or $S_NormalizedIPs.Count -eq 0)
    {
        throw 'No valid IP addresses were provided after normalization.'
    }

    $S_IpHashSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($S_Ip in $S_NormalizedIPs)
    {
        [void]$S_IpHashSet.Add($S_Ip)
    }

    $S_ToUtc = (Get-Date).ToUniversalTime()
    $S_FromUtc = $S_ToUtc.AddDays(-$Days)

    $S_FromIso = $S_FromUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')
    $S_ToIso = $S_ToUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')

    Write-Host "`nSearching interactive sign-ins..." -ForegroundColor Cyan
    Write-Host "  IP(s)      : $($S_NormalizedIPs -join ', ')"
    Write-Host "  Date range : $S_FromIso  ->  $S_ToIso"
    Write-Host "  Days       : $Days"

    function Get-GraphNextLink
    {
        param(
            [Parameter(Mandatory = $false)]
            [object]$Response
        )

        if (-not $Response)
        {
            return $null
        }

        if ($Response.PSObject.Properties.Name -contains '@odata.nextLink')
        {
            return [string]$Response.'@odata.nextLink'
        }

        return $null
    }

    function Get-GraphPageValues
    {
        param(
            [Parameter(Mandatory = $false)]
            [object]$Response
        )

        if (-not $Response)
        {
            return @()
        }

        if ($Response.PSObject.Properties.Name -contains 'value')
        {
            return @($Response.value)
        }

        return @()
    }

    # ---------------------------------------------------------------------------
    # Query sign-ins with pagination
    # ---------------------------------------------------------------------------

    $S_SelectFields = @(
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
    $S_Filter = "createdDateTime ge $S_FromIso and createdDateTime le $S_ToIso and isInteractive eq true"
    $S_EncodedFilter = [uri]::EscapeDataString($S_Filter)
    $S_EncodedSelect = [uri]::EscapeDataString($S_SelectFields)

    $S_NextLink = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$top=1000&`$filter=$S_EncodedFilter&`$select=$S_EncodedSelect"
    $S_AllSignIns = [System.Collections.Generic.List[object]]::new()
    $S_FallbackUsed = $false

    try
    {
        while ($S_NextLink)
        {
            $S_Response = Invoke-MgGraphRequest -Method GET -Uri $S_NextLink -OutputType PSObject
            $S_PageValues = Get-GraphPageValues -Response $S_Response

            if ($S_PageValues.Count -gt 0)
            {
                $S_AllSignIns.AddRange([object[]]$S_PageValues)
                Write-Host "  Retrieved $($S_AllSignIns.Count) interactive sign-in record(s) so far..." -ForegroundColor Gray
            }

            $S_NextLink = Get-GraphNextLink -Response $S_Response
            if ($S_NextLink)
            {
                [System.Threading.Thread]::Sleep($S_GraphRequestDelayMilliseconds)
            }
        }
    }
    catch
    {
        Write-Warning "Interactive server-side filter failed. Retrying with date-only filter and local interactive filtering."
        $S_FallbackUsed = $true
        $S_AllSignIns.Clear()

        $S_FallbackFilter = "createdDateTime ge $S_FromIso and createdDateTime le $S_ToIso"
        $S_EncodedFallbackFilter = [uri]::EscapeDataString($S_FallbackFilter)
        $S_NextLink = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$top=1000&`$filter=$S_EncodedFallbackFilter&`$select=$S_EncodedSelect"

        while ($S_NextLink)
        {
            $S_Response = Invoke-MgGraphRequest -Method GET -Uri $S_NextLink -OutputType PSObject
            $S_PageValues = Get-GraphPageValues -Response $S_Response

            if ($S_PageValues.Count -gt 0)
            {
                $S_AllSignIns.AddRange([object[]]$S_PageValues)
                Write-Host "  Retrieved $($S_AllSignIns.Count) total sign-in record(s) so far..." -ForegroundColor Gray
            }

            $S_NextLink = Get-GraphNextLink -Response $S_Response
            if ($S_NextLink)
            {
                [System.Threading.Thread]::Sleep($S_GraphRequestDelayMilliseconds)
            }
        }
    }

    if ($S_AllSignIns.Count -eq 0)
    {
        Write-Host "`nNo sign-ins found in the selected time window." -ForegroundColor Yellow
        return
    }

    # ---------------------------------------------------------------------------
    # Filter by interactive + IP list and shape output
    # ---------------------------------------------------------------------------

    $S_Matched = foreach ($S_Entry in $S_AllSignIns)
    {
        $S_IsInteractive = $true
        if ($S_FallbackUsed)
        {
            $S_IsInteractive = $S_Entry.isInteractive -eq $true
        }

        $S_EntryIp = $S_Entry.ipAddress
        if (-not $S_IsInteractive -or -not $S_EntryIp -or -not $S_IpHashSet.Contains([string]$S_EntryIp))
        {
            continue
        }

        $S_StatusErrorCode = $null
        $S_StatusFailureReason = $null
        if ($S_Entry.status)
        {
            $S_StatusErrorCode = $S_Entry.status.errorCode
            $S_StatusFailureReason = $S_Entry.status.failureReason
        }

        $S_City = $null
        $S_State = $null
        $S_Country = $null
        if ($S_Entry.location)
        {
            $S_City = $S_Entry.location.city
            $S_State = $S_Entry.location.state
            $S_Country = $S_Entry.location.countryOrRegion
        }

        [PSCustomObject]@{
            CreatedDateTime         = $S_Entry.createdDateTime
            UserDisplayName         = $S_Entry.userDisplayName
            UserPrincipalName       = $S_Entry.userPrincipalName
            UserId                  = $S_Entry.userId
            IPAddress               = $S_EntryIp
            AppDisplayName          = $S_Entry.appDisplayName
            ClientAppUsed           = $S_Entry.clientAppUsed
            ResourceDisplayName     = $S_Entry.resourceDisplayName
            IsInteractive           = $S_Entry.isInteractive
            ConditionalAccessStatus = $S_Entry.conditionalAccessStatus
            StatusErrorCode         = $S_StatusErrorCode
            StatusFailureReason     = $S_StatusFailureReason
            City                    = $S_City
            State                   = $S_State
            CountryOrRegion         = $S_Country
            SignInId                = $S_Entry.id
        }
    }

    if (-not $S_Matched)
    {
        Write-Host "`nNo interactive sign-ins matched the supplied IP list." -ForegroundColor Yellow
        return
    }

    $S_Matched = $S_Matched | Sort-Object CreatedDateTime

    # ---------------------------------------------------------------------------
    # Export
    # ---------------------------------------------------------------------------

    if (-not (Test-Path -LiteralPath $OutputPath))
    {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    $S_ReportTime = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
    $S_IpTag = ($S_NormalizedIPs -join '_') -replace '[^\w\-]', '-'
    $S_CsvName = "InteractiveSignInsByIP_${S_IpTag}_${S_ReportTime}.csv"
    $S_CsvPath = Join-Path -Path $OutputPath -ChildPath $S_CsvName

    $S_Matched | Export-Csv -LiteralPath $S_CsvPath -NoTypeInformation -Encoding UTF8

    Write-Host "`nMatched records: $($S_Matched.Count)" -ForegroundColor Green
    Write-Host "Report saved to: $S_CsvPath" -ForegroundColor Green

    # Quick on-screen summary per IP.
    $S_Summary = $S_Matched |
        Group-Object -Property IPAddress |
        Sort-Object -Property Name |
        Select-Object @{ Name = 'IPAddress'; Expression = { $_.Name } }, @{ Name = 'Count'; Expression = { $_.Count } }

    Write-Host "`nSummary by IP:" -ForegroundColor Cyan
    $S_Summary | Format-Table -AutoSize
}
finally
{
    $S_DisconnectChoice = Read-Host "`nDisconnect from Microsoft Graph? [Y] Yes  [N] Keep session  (Default: N)"
    if ($S_DisconnectChoice -eq 'Y')
    {
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Write-Host "Disconnected." -ForegroundColor Green
    }
    else
    {
        Write-Host "Graph session kept alive." -ForegroundColor Green
    }
}
