#Requires -Modules Microsoft.Graph.Users

<#
.SYNOPSIS
    Reports all unique IP addresses seen in interactive sign-in events over a configurable lookback period.

.DESCRIPTION
    Connects to Microsoft Graph and retrieves interactive sign-in events for the specified date range.
    Aggregates results by unique IP address and reports user, application, and geographic information
    for each IP. Exports results to CSV.

.PARAMETER Days
    Number of days to look back from now (UTC). Supported range: 1 to 30.

.PARAMETER OutputPath
    Folder to save the CSV report. The folder will be created if it does not exist.

.PARAMETER ThrottleMs
    Delay in milliseconds between Graph API page calls to reduce throttling. Defaults to 5.

.EXAMPLE
    .\Get-InteractiveSignInUniqueIPs.ps1 -Days 7 -OutputPath "C:\IR\Output"

.EXAMPLE
    .\Get-InteractiveSignInUniqueIPs.ps1 -Days 30 -OutputPath "C:\IR\Output"
#>

[CmdletBinding()]
param (
    # Number of days to look back from now (UTC). Supported range: 1 to 30.
    [Parameter(Mandatory = $true)]
    [ValidateRange(1, 30)]
    [int]$Days,

    # Folder to save the CSV report.
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath,

    # Delay in milliseconds between page calls to reduce API pressure.
    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 1000)]
    [int]$ThrottleMs = 5
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

$context = Get-MgContext -ErrorAction SilentlyContinue
$tenantId = if ($context) { $context.TenantId } else { 'Unknown' }
$tenantName = 'Unknown'
try {
    $orgInfo = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=displayName,verifiedDomains' -OutputType PSObject -ErrorAction Stop
    if ($orgInfo.value -and $orgInfo.value.Count -gt 0) {
        $org = $orgInfo.value[0]
        $tenantName = [string]$org.displayName
        $defaultDomain = ($org.verifiedDomains | Where-Object { $_.isDefault -eq $true } | Select-Object -First 1).name
        if ($defaultDomain) { $tenantName = "$tenantName ($defaultDomain)" }
    }
}
catch {
    Write-Warning "Could not retrieve tenant display name: $($_.Exception.Message)"
}

Write-Host ""
Write-Host ("=" * 70) -ForegroundColor Cyan
Write-Host "  Tenant: $tenantName" -ForegroundColor Cyan
Write-Host "  Tenant ID: $tenantId" -ForegroundColor Cyan
Write-Host ("=" * 70) -ForegroundColor Cyan
Write-Host ""

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

try {
    # ---------------------------------------------------------------------------
    # Inputs and date range
    # ---------------------------------------------------------------------------

    $toUtc = (Get-Date).ToUniversalTime()
    $fromUtc = $toUtc.AddDays(-$Days)

    $fromIso = $fromUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')
    $toIso = $toUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')

    Write-Host "`nSearching interactive sign-ins for unique IP inventory..." -ForegroundColor Cyan
    Write-Host "  Date range : $fromIso  ->  $toIso"
    Write-Host "  Days       : $Days"
    Write-Host "  ThrottleMs : $ThrottleMs"

    # ---------------------------------------------------------------------------
    # Query sign-ins with pagination
    # ---------------------------------------------------------------------------

    $selectFields = @(
        'id',
        'createdDateTime',
        'userPrincipalName',
        'ipAddress',
        'isInteractive',
        'location',
        'status',
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
            if ($nextLink -and $ThrottleMs -gt 0) {
                [System.Threading.Thread]::Sleep($ThrottleMs)
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
            if ($nextLink -and $ThrottleMs -gt 0) {
                [System.Threading.Thread]::Sleep($ThrottleMs)
            }
        }
    }

    if ($allSignIns.Count -eq 0) {
        Write-Host "`nNo sign-ins found in the selected time window." -ForegroundColor Yellow
        return
    }

    # ---------------------------------------------------------------------------
    # Build unique IP inventory with country/location details
    # ---------------------------------------------------------------------------

    $interactiveSignIns = if ($fallbackUsed) {
        @($allSignIns | Where-Object { $_.isInteractive -eq $true })
    }
    else {
        @($allSignIns)
    }

    $signInsWithIp = @(
        $interactiveSignIns | Where-Object {
            $_.PSObject.Properties.Name -contains 'ipAddress' -and
            -not [string]::IsNullOrWhiteSpace([string]$_.ipAddress)
        }
    )

    if ($signInsWithIp.Count -eq 0) {
        Write-Host "`nNo interactive sign-ins with IP addresses were found." -ForegroundColor Yellow
        return
    }

    # --- ISO 3166-1 alpha-2 country code to full name lookup ---
    $countryCodeMap = @{
        'NZ' = 'New Zealand'; 'IN' = 'India'; 'CN' = 'China'; 'KR' = 'South Korea'; 'US' = 'United States'
        'RU' = 'Russia'; 'BR' = 'Brazil'; 'MY' = 'Malaysia'; 'LU' = 'Luxembourg'; 'FI' = 'Finland'
        'DE' = 'Germany'; 'TW' = 'Taiwan'; 'AU' = 'Australia'; 'ET' = 'Ethiopia'; 'GB' = 'United Kingdom'
        'FR' = 'France'; 'JP' = 'Japan'; 'SG' = 'Singapore'; 'CA' = 'Canada'; 'MX' = 'Mexico'
        'ES' = 'Spain'; 'IT' = 'Italy'; 'NL' = 'Netherlands'; 'SE' = 'Sweden'; 'NO' = 'Norway'
        'CH' = 'Switzerland'; 'AT' = 'Austria'; 'BE' = 'Belgium'; 'CZ' = 'Czech Republic'; 'PL' = 'Poland'
        'TR' = 'Turkey'; 'SA' = 'Saudi Arabia'; 'AE' = 'United Arab Emirates'; 'ZA' = 'South Africa'; 'NG' = 'Nigeria'
        'EG' = 'Egypt'; 'IL' = 'Israel'; 'TH' = 'Thailand'; 'VN' = 'Vietnam'; 'ID' = 'Indonesia'
        'PH' = 'Philippines'; 'HK' = 'Hong Kong'; 'MO' = 'Macau'; 'PK' = 'Pakistan'; 'LK' = 'Sri Lanka'
        'BD' = 'Bangladesh'; 'NP' = 'Nepal'; 'KH' = 'Cambodia'; 'LA' = 'Laos'; 'MM' = 'Myanmar'
        'IR' = 'Iran'; 'IQ' = 'Iraq'; 'SY' = 'Syria'; 'JO' = 'Jordan'; 'KW' = 'Kuwait'
        'LB' = 'Lebanon'; 'OM' = 'Oman'; 'QA' = 'Qatar'; 'YE' = 'Yemen'; 'AF' = 'Afghanistan'
        'KZ' = 'Kazakhstan'; 'UZ' = 'Uzbekistan'; 'KG' = 'Kyrgyzstan'; 'TJ' = 'Tajikistan'; 'TM' = 'Turkmenistan'
        'AM' = 'Armenia'; 'AZ' = 'Azerbaijan'; 'GE' = 'Georgia'; 'BY' = 'Belarus'; 'UA' = 'Ukraine'
        'MD' = 'Moldova'; 'RO' = 'Romania'; 'BG' = 'Bulgaria'; 'GR' = 'Greece'; 'CY' = 'Cyprus'
        'PT' = 'Portugal'; 'IE' = 'Ireland'; 'IS' = 'Iceland'; 'DK' = 'Denmark'; 'EE' = 'Estonia'
        'LV' = 'Latvia'; 'LT' = 'Lithuania'; 'HU' = 'Hungary'; 'SK' = 'Slovakia'; 'SI' = 'Slovenia'
        'HR' = 'Croatia'; 'BA' = 'Bosnia and Herzegovina'; 'RS' = 'Serbia'; 'ME' = 'Montenegro'; 'MK' = 'North Macedonia'
        'AL' = 'Albania'; 'XK' = 'Kosovo'; 'MT' = 'Malta'; 'AD' = 'Andorra'; 'MC' = 'Monaco'
        'LI' = 'Liechtenstein'; 'SM' = 'San Marino'; 'VA' = 'Vatican City'; 'AR' = 'Argentina'; 'CL' = 'Chile'
        'CO' = 'Colombia'; 'PE' = 'Peru'; 'VE' = 'Venezuela'; 'EC' = 'Ecuador'; 'BO' = 'Bolivia'
        'PY' = 'Paraguay'; 'UY' = 'Uruguay'; 'GY' = 'Guyana'; 'SR' = 'Suriname'; 'CU' = 'Cuba'
        'DO' = 'Dominican Republic'; 'HT' = 'Haiti'; 'JM' = 'Jamaica'; 'PR' = 'Puerto Rico'; 'TT' = 'Trinidad and Tobago'
        'BS' = 'Bahamas'; 'BB' = 'Barbados'; 'GT' = 'Guatemala'; 'HN' = 'Honduras'; 'SV' = 'El Salvador'
        'NI' = 'Nicaragua'; 'CR' = 'Costa Rica'; 'PA' = 'Panama'; 'BZ' = 'Belize'; 'KE' = 'Kenya'
        'TZ' = 'Tanzania'; 'UG' = 'Uganda'; 'RW' = 'Rwanda'; 'GH' = 'Ghana'; 'CI' = "Cote d'Ivoire"
        'SN' = 'Senegal'; 'CM' = 'Cameroon'; 'AO' = 'Angola'; 'MZ' = 'Mozambique'; 'ZW' = 'Zimbabwe'
        'ZM' = 'Zambia'; 'BW' = 'Botswana'; 'NA' = 'Namibia'; 'MA' = 'Morocco'; 'DZ' = 'Algeria'
        'TN' = 'Tunisia'; 'LY' = 'Libya'; 'SD' = 'Sudan'; 'SO' = 'Somalia'; 'MG' = 'Madagascar'
        'MU' = 'Mauritius'; 'SC' = 'Seychelles'; 'FJ' = 'Fiji'; 'PG' = 'Papua New Guinea'; 'WS' = 'Samoa'
        'TO' = 'Tonga'; 'VU' = 'Vanuatu'; 'NC' = 'New Caledonia'; 'PF' = 'French Polynesia'; 'GU' = 'Guam'
        'VG' = 'British Virgin Islands'; 'KY' = 'Cayman Islands'; 'BM' = 'Bermuda'; 'MV' = 'Maldives'; 'BT' = 'Bhutan'
        'BN' = 'Brunei'; 'TL' = 'Timor-Leste'; 'PS' = 'Palestine'
    }

    $uniqueIpReport = foreach ($group in ($signInsWithIp | Group-Object -Property ipAddress)) {
        $records = @($group.Group)

        $countries = @(
            $records |
                ForEach-Object {
                    if ($_.PSObject.Properties.Name -contains 'location' -and $_.location) {
                        $_.location.countryOrRegion
                    }
                } |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Sort-Object -Unique
        )

        $states = @(
            $records |
                ForEach-Object {
                    if ($_.PSObject.Properties.Name -contains 'location' -and $_.location) {
                        $_.location.state
                    }
                } |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Sort-Object -Unique
        )

        $cities = @(
            $records |
                ForEach-Object {
                    if ($_.PSObject.Properties.Name -contains 'location' -and $_.location) {
                        $_.location.city
                    }
                } |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Sort-Object -Unique
        )

        $users = @(
            $records |
                ForEach-Object { $_.userPrincipalName } |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Sort-Object -Unique
        )

        $timestamps = @(
            $records |
                ForEach-Object {
                    if ($_.createdDateTime) {
                        [datetime]$_.createdDateTime
                    }
                } |
                Sort-Object
        )

        $firstSeen = $null
        $lastSeen = $null
        if ($timestamps.Count -gt 0) {
            $firstSeen = $timestamps[0]
            $lastSeen = $timestamps[$timestamps.Count - 1]
        }

        # Count successful vs failed attempts (errorCode 0 = success)
        # Use simple dot notation - works for both Hashtable and PSObject
        $successCount = 0
        $caSuccessCount = 0
        $caFailureCount = 0
        $caNotAppliedCount = 0
        foreach ($record in $records) {
            # Sign-in status (errorCode 0 = success)
            $errorCode = $null
            try { $errorCode = $record.status.errorCode } catch {}
            if ($null -eq $errorCode) {
                try { $errorCode = $record['status']['errorCode'] } catch {}
            }
            if ($null -ne $errorCode -and [int]$errorCode -eq 0) {
                $successCount++
            }

            # Conditional Access status
            $caStatus = $null
            try { $caStatus = [string]$record.conditionalAccessStatus } catch {}
            if ([string]::IsNullOrWhiteSpace($caStatus)) {
                try { $caStatus = [string]$record['conditionalAccessStatus'] } catch {}
            }
            switch -Regex ($caStatus) {
                '^success$'    { $caSuccessCount++ }
                '^failure$'    { $caFailureCount++ }
                '^notApplied$' { $caNotAppliedCount++ }
            }
        }
        $failedCount = $group.Count - $successCount

        $statusSummary = if ($successCount -gt 0 -and $failedCount -eq 0) {
            'Success'
        } elseif ($failedCount -gt 0 -and $successCount -eq 0) {
            'Failed'
        } else {
            'Mixed'
        }

        $caSummary = if ($caSuccessCount -gt 0 -and $caFailureCount -eq 0) {
            'Success'
        } elseif ($caFailureCount -gt 0 -and $caSuccessCount -eq 0) {
            'Failure'
        } elseif ($caSuccessCount -eq 0 -and $caFailureCount -eq 0 -and $caNotAppliedCount -gt 0) {
            'Not Applied'
        } elseif ($caSuccessCount -gt 0 -or $caFailureCount -gt 0) {
            'Mixed'
        } else {
            'N/A'
        }

        # Resolve country code to full name
        $countryCode = if ($countries.Count -gt 0) { $countries[0] } else { 'Unknown' }
        $countryFullName = if ($countryCodeMap.ContainsKey($countryCode)) { $countryCodeMap[$countryCode] } else { $countryCode }
        $countryDisplay = if ($countryCode -eq 'Unknown') { 'Unknown' } else { "$countryFullName ($countryCode)" }

        [PSCustomObject]@{
            IPAddress        = [string]$group.Name
            SignInCount      = $group.Count
            SuccessCount     = $successCount
            FailedCount      = $failedCount
            Status           = $statusSummary
            CASuccessCount   = $caSuccessCount
            CAFailureCount   = $caFailureCount
            CANotAppliedCount= $caNotAppliedCount
            CAStatus         = $caSummary
            CountryCode      = $countryCode
            CountryFullName  = $countryFullName
            CountryOrRegion  = $countryDisplay
            State            = if ($states.Count -gt 0) { $states -join ' | ' } else { $null }
            City             = if ($cities.Count -gt 0) { $cities -join ' | ' } else { $null }
            UniqueUsers      = $users.Count
            SampleUsers      = ($users | Select-Object -First 5) -join '; '
            FirstSeenUtc     = $firstSeen
            LastSeenUtc      = $lastSeen
        }
    }

    $uniqueIpReport = @($uniqueIpReport) | Sort-Object -Property SignInCount -Descending

    # ---------------------------------------------------------------------------
    # Export
    # ---------------------------------------------------------------------------

    if (-not (Test-Path -LiteralPath $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    $reportTime = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
    $reportDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $csvName = "InteractiveUniqueIPs_${Days}d_${reportTime}.csv"
    $csvPath = Join-Path -Path $OutputPath -ChildPath $csvName

    $uniqueIpReport | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8

    # --- Build HTML report ---
    $totalSignIns = ($uniqueIpReport | Measure-Object -Property SignInCount -Sum).Sum
    $uniqueCountries = @(
        $uniqueIpReport |
            ForEach-Object { $_.CountryOrRegion } |
            Where-Object { $_ -ne 'Unknown' } |
            Sort-Object -Unique
    ).Count

    $tableRows = ($uniqueIpReport | ForEach-Object {
        $riskClass = switch ($_.SignInCount) {
            { $_ -gt 500 } { 'risk-high' }
            { $_ -gt 100 } { 'risk-medium' }
            default { 'risk-low' }
        }
        $statusClass = switch ($_.Status) {
            'Success' { 'status-success' }
            'Failed'  { 'status-failed' }
            default   { 'status-mixed' }
        }
        $caStatusClass = switch ($_.CAStatus) {
            'Success'     { 'status-success' }
            'Failure'     { 'status-failed' }
            'Not Applied' { 'status-na' }
            'Mixed'       { 'status-mixed' }
            default       { 'status-na' }
        }
        $statusText = if ($_.Status -eq 'Mixed') { "Mixed ($($_.SuccessCount)/$($_.FailedCount))" } else { $_.Status }
        "<tr class=`"$riskClass`"><td><code>$([System.Net.WebUtility]::HtmlEncode($_.IPAddress))</code></td><td>$([System.Net.WebUtility]::HtmlEncode($_.CountryOrRegion))</td><td>$([System.Net.WebUtility]::HtmlEncode($_.State -as [string]))</td><td>$([System.Net.WebUtility]::HtmlEncode($_.City -as [string]))</td><td>$($_.SignInCount)</td><td><span class=`"badge $statusClass`">$statusText</span></td><td><span class=`"badge $caStatusClass`">$($_.CAStatus)</span></td><td class=`"col-success`">$($_.SuccessCount)</td><td class=`"col-failed`">$($_.FailedCount)</td><td>$($_.UniqueUsers)</td><td>$($_.FirstSeenUtc.ToString('yyyy-MM-dd HH:mm:ss'))</td><td>$($_.LastSeenUtc.ToString('yyyy-MM-dd HH:mm:ss'))</td></tr>"
    }) -join "`n"

    # --- Build country statistics ---
    $countryStats = foreach ($country in ($uniqueIpReport | Select-Object -ExpandProperty CountryFullName -Unique | Where-Object { $_ -ne 'Unknown' })) {
        $countryIps = @($uniqueIpReport | Where-Object { $_.CountryFullName -eq $country })
        $countryCode = ($countryIps | Select-Object -First 1).CountryCode
        $countrySignIns = ($countryIps | Measure-Object -Property SignInCount -Sum).Sum
        $countrySuccess = ($countryIps | Measure-Object -Property SuccessCount -Sum).Sum
        $countryFailed = ($countryIps | Measure-Object -Property FailedCount -Sum).Sum

        [PSCustomObject]@{
            Country      = $country
            CountryCode  = $countryCode
            IPs          = $countryIps.Count
            SignIns      = $countrySignIns
            Success      = $countrySuccess
            Failed       = $countryFailed
            Users        = ($countryIps | Measure-Object -Property UniqueUsers -Sum).Sum
        }
    }

    $countryStats = @($countryStats) | Sort-Object -Property SignIns -Descending

    $countryTableRows = ($countryStats | ForEach-Object {
        "<tr><td>$([System.Net.WebUtility]::HtmlEncode($_.Country)) <span class=`"country-code`">($([System.Net.WebUtility]::HtmlEncode($_.CountryCode)))</span></td><td>$($_.IPs)</td><td>$($_.SignIns)</td><td class=`"col-success`">$($_.Success)</td><td class=`"col-failed`">$($_.Failed)</td><td>$($_.Users)</td></tr>"
    }) -join "`n"

    $countryFilterOptions = ($uniqueIpReport | Select-Object -ExpandProperty CountryFullName -Unique | Sort-Object | ForEach-Object {
        "<option value=`"$([System.Net.WebUtility]::HtmlEncode($_))`">$([System.Net.WebUtility]::HtmlEncode($_))</option>"
    }) -join "`n"

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Interactive Sign-In Unique IPs Report</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; padding: 30px; }
        .header { background: linear-gradient(135deg, #1a1a2e, #16213e); color: #fff; padding: 30px 40px; border-radius: 12px; margin-bottom: 30px; }
        .header h1 { font-size: 1.8em; margin-bottom: 4px; }
        .header .subtitle { font-size: 0.95em; opacity: 0.85; }
        .summary-cards { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
        .card { background: #fff; border-radius: 10px; padding: 24px 30px; flex: 1; min-width: 180px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); border-left: 5px solid #0078d4; }
        .card.blue { border-left-color: #0078d4; }
        .card.green { border-left-color: #107c10; }
        .card.orange { border-left-color: #ff8c00; }
        .card.red { border-left-color: #d13438; }
        .card .label { font-size: 12px; color: #666; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px; }
        .card .value { font-size: 36px; font-weight: 700; color: #1a1a2e; }
        .card .detail { font-size: 12px; color: #888; margin-top: 6px; }
        .section { margin-bottom: 30px; }
        .section h2 { font-size: 18px; color: #1a1a2e; margin-bottom: 12px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0; }
        .controls { background: #fff; padding: 16px; border-radius: 8px; margin-bottom: 16px; box-shadow: 0 2px 4px rgba(0,0,0,0.04); display: flex; gap: 12px; flex-wrap: wrap; align-items: center; }
        .controls input, .controls select { padding: 8px 12px; border: 1px solid #ddd; border-radius: 4px; font-size: 13px; }
        .controls input:focus, .controls select:focus { outline: none; border-color: #0078d4; box-shadow: 0 0 0 3px rgba(0, 120, 212, 0.1); }
        .controls label { font-size: 13px; font-weight: 600; color: #333; }
        table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
        th { background: #0078d4; color: #fff; text-align: left; padding: 12px 14px; font-size: 12px; text-transform: uppercase; letter-spacing: 0.3px; }
        td { padding: 11px 14px; font-size: 13px; border-bottom: 1px solid #eee; }
        tbody tr:hover { background: #f5f5f5; }
        tbody tr.risk-high { background: #fff8f7; }
        tbody tr.risk-high:hover { background: #ffedea; }
        tbody tr.risk-medium { background: #fffbf5; }
        tbody tr.risk-medium:hover { background: #fff4e6; }
        tbody tr.hidden { display: none; }
        code { background: #f0f0f0; padding: 2px 6px; border-radius: 3px; font-family: 'Courier New', monospace; }
        .badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 600; text-transform: uppercase; }
        .badge.status-success { background: #dff6dd; color: #107c10; }
        .badge.status-failed { background: #fde7e9; color: #d13438; }
        .badge.status-mixed { background: #fff4ce; color: #997a00; }
        .badge.status-na { background: #e6e6e6; color: #666; }
        .col-success { color: #107c10; font-weight: 600; }
        .col-failed { color: #d13438; font-weight: 600; }
        .country-code { color: #888; font-size: 11px; }
        .footer { text-align: center; font-size: 12px; color: #999; margin-top: 32px; }
        .stats-grid { display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 20px; }
        .stats-grid > div { background: #fff; padding: 14px 18px; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.04); flex: 1; min-width: 140px; }
        .stats-grid .label { font-size: 11px; color: #666; text-transform: uppercase; }
        .stats-grid .value { font-size: 20px; font-weight: 700; color: #1a1a2e; margin-top: 4px; }
    </style>
    <script>
        function filterTable() {
            const searchInput = document.getElementById('searchInput');
            const countrySelect = document.getElementById('countryFilter');
            const searchValue = searchInput.value.toLowerCase();
            const countryValue = countrySelect.value;
            const rows = document.querySelectorAll('#ipTable tbody tr');

            let visibleCount = 0;
            rows.forEach(row => {
                const ipCell = row.cells[0].textContent.toLowerCase();
                const countryCell = row.cells[1].textContent;
                
                const matchSearch = searchValue === '' || ipCell.includes(searchValue);
                const matchCountry = countryValue === '' || countryCell.startsWith(countryValue);

                if (matchSearch && matchCountry) {
                    row.classList.remove('hidden');
                    visibleCount++;
                } else {
                    row.classList.add('hidden');
                }
            });

            document.getElementById('resultCount').textContent = visibleCount;
        }

        document.addEventListener('DOMContentLoaded', function() {
            document.getElementById('searchInput').addEventListener('keyup', filterTable);
            document.getElementById('countryFilter').addEventListener('change', filterTable);
        });
    </script>
</head>
<body>
    <div class="header">
        <h1>Interactive Sign-In Unique IPs Report</h1>
        <div class="subtitle">Generated: $reportDate | Tenant: $tenantName | Tenant ID: $tenantId | Period: $Days days | Total IPs: $($uniqueIpReport.Count)</div>
    </div>

    <div class="summary-cards">
        <div class="card blue">
            <div class="label">Unique IP Addresses</div>
            <div class="value">$($uniqueIpReport.Count)</div>
        </div>
        <div class="card green">
            <div class="label">Total Sign-Ins</div>
            <div class="value">$totalSignIns</div>
        </div>
        <div class="card orange">
            <div class="label">Unique Countries</div>
            <div class="value">$uniqueCountries</div>
        </div>
        <div class="card red">
            <div class="label">Avg per IP</div>
            <div class="value">$([math]::Round($totalSignIns / [math]::Max($uniqueIpReport.Count, 1), 1))</div>
            <div class="detail">sign-ins per IP</div>
        </div>
    </div>

    <div class="section">
        <h2>Sign-In Activity by Country</h2>
        <table>
            <thead>
                <tr>
                    <th>Country</th>
                    <th>IP Count</th>
                    <th>Total Sign-Ins</th>
                    <th>Successful</th>
                    <th>Failed</th>
                    <th>Unique Users</th>
                </tr>
            </thead>
            <tbody>
$countryTableRows
            </tbody>
        </table>
    </div>

    <div class="section">
        <h2>IP Geolocation & Sign-In Activity</h2>
        <div class="controls">
            <label for="searchInput">🔍 Search IP:</label>
            <input type="text" id="searchInput" placeholder="Enter IP address...">
            <label for="countryFilter">Filter by Country:</label>
            <select id="countryFilter">
                <option value="">-- All Countries --</option>
$countryFilterOptions
            </select>
            <span id="resultCount" style="margin-left: auto; font-size: 13px; color: #666;"></span>
        </div>
        <table id="ipTable">
            <thead>
                <tr>
                    <th>IP Address</th>
                    <th>Country</th>
                    <th>State</th>
                    <th>City</th>
                    <th>Sign-Ins</th>
                    <th>Status</th>
                    <th>CA Status</th>
                    <th>Success</th>
                    <th>Failed</th>
                    <th>Users</th>
                    <th>First Seen</th>
                    <th>Last Seen</th>
                </tr>
            </thead>
            <tbody>
$tableRows
            </tbody>
        </table>
    </div>

    <div class="footer">
        CSV data: $(Split-Path $csvPath -Leaf) | Report generated by Get-InteractiveSignInUniqueIPs.ps1
    </div>
</body>
</html>
"@

    $htmlName = "InteractiveUniqueIPs_${Days}d_${reportTime}.html"
    $htmlPath = Join-Path -Path $OutputPath -ChildPath $htmlName
    $html | Out-File -FilePath $htmlPath -Encoding UTF8

    Write-Host "`nHTML report saved to: $htmlPath" -ForegroundColor Green
    Write-Host "`nUnique IP addresses: $($uniqueIpReport.Count)" -ForegroundColor Green
    Write-Host "Total sign-ins: $totalSignIns" -ForegroundColor Green

    Write-Host "`nTop IPs by sign-in count:" -ForegroundColor Cyan
    $uniqueIpReport |
        Select-Object -First 20 -Property IPAddress, CountryOrRegion, SignInCount, LastSeenUtc |
        Format-Table -AutoSize
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
