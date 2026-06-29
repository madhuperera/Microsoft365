#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users

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
param(
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

$S_Context = Get-MgContext -ErrorAction SilentlyContinue
if ($S_Context)
{
    $S_TenantId = $S_Context.TenantId
}
else
{
    $S_TenantId = 'Unknown'
}

$S_TenantName = 'Unknown'
try
{
    $S_OrgInfo = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=displayName,verifiedDomains' -OutputType PSObject -ErrorAction Stop
    if ($S_OrgInfo.value -and $S_OrgInfo.value.Count -gt 0)
    {
        $S_Org = $S_OrgInfo.value[0]
        $S_TenantName = [string]$S_Org.displayName
        $S_DefaultDomain = ($S_Org.verifiedDomains | Where-Object { $_.isDefault -eq $true } | Select-Object -First 1).name
        if ($S_DefaultDomain)
        {
            $S_TenantName = "$S_TenantName ($S_DefaultDomain)"
        }
    }
}
catch
{
    Write-Warning "Could not retrieve tenant display name: $($_.Exception.Message)"
}

Write-Host ""
Write-Host ("=" * 70) -ForegroundColor Cyan
Write-Host "  Tenant: $S_TenantName" -ForegroundColor Cyan
Write-Host "  Tenant ID: $S_TenantId" -ForegroundColor Cyan
Write-Host ("=" * 70) -ForegroundColor Cyan
Write-Host ""

function Get-GraphNextLink {
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

function Get-GraphPageValues {
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

try
{
    # ---------------------------------------------------------------------------
    # Inputs and date range
    # ---------------------------------------------------------------------------

    $S_ThrottleMs = $ThrottleMs
    $S_ToUtc = (Get-Date).ToUniversalTime()
    $S_FromUtc = $S_ToUtc.AddDays(-$Days)

    $S_FromIso = $S_FromUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')
    $S_ToIso = $S_ToUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')

    Write-Host "`nSearching interactive sign-ins for unique IP inventory..." -ForegroundColor Cyan
    Write-Host "  Date range : $S_FromIso  ->  $S_ToIso"
    Write-Host "  Days       : $Days"
    Write-Host "  ThrottleMs : $S_ThrottleMs"

    # ---------------------------------------------------------------------------
    # Query sign-ins with pagination
    # ---------------------------------------------------------------------------

    $S_SelectFields = @(
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
            if ($S_NextLink -and $S_ThrottleMs -gt 0)
            {
                [System.Threading.Thread]::Sleep($S_ThrottleMs)
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
            if ($S_NextLink -and $S_ThrottleMs -gt 0)
            {
                [System.Threading.Thread]::Sleep($S_ThrottleMs)
            }
        }
    }

    if ($S_AllSignIns.Count -eq 0)
    {
        Write-Host "`nNo sign-ins found in the selected time window." -ForegroundColor Yellow
        return
    }

    # ---------------------------------------------------------------------------
    # Build unique IP inventory with country/location details
    # ---------------------------------------------------------------------------

    if ($S_FallbackUsed)
    {
        $S_InteractiveSignIns = @($S_AllSignIns | Where-Object { $_.isInteractive -eq $true })
    }
    else
    {
        $S_InteractiveSignIns = @($S_AllSignIns)
    }

    $S_SignInsWithIp = @(
        $S_InteractiveSignIns | Where-Object {
            $_.PSObject.Properties.Name -contains 'ipAddress' -and
            -not [string]::IsNullOrWhiteSpace([string]$_.ipAddress)
        }
    )

    if ($S_SignInsWithIp.Count -eq 0)
    {
        Write-Host "`nNo interactive sign-ins with IP addresses were found." -ForegroundColor Yellow
        return
    }

    # --- ISO 3166-1 alpha-2 country code to full name lookup ---
    $S_CountryCodeMap = @{
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

    $S_UniqueIpReport = foreach ($S_Group in ($S_SignInsWithIp | Group-Object -Property ipAddress))
    {
        $S_Records = @($S_Group.Group)

        $S_Countries = @(
            $S_Records |
                ForEach-Object {
                    if ($_.PSObject.Properties.Name -contains 'location' -and $_.location)
                    {
                        $_.location.countryOrRegion
                    }
                } |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Sort-Object -Unique
        )

        $S_States = @(
            $S_Records |
                ForEach-Object {
                    if ($_.PSObject.Properties.Name -contains 'location' -and $_.location)
                    {
                        $_.location.state
                    }
                } |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Sort-Object -Unique
        )

        $S_Cities = @(
            $S_Records |
                ForEach-Object {
                    if ($_.PSObject.Properties.Name -contains 'location' -and $_.location)
                    {
                        $_.location.city
                    }
                } |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Sort-Object -Unique
        )

        $S_Users = @(
            $S_Records |
                ForEach-Object { $_.userPrincipalName } |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Sort-Object -Unique
        )

        $S_Timestamps = @(
            $S_Records |
                ForEach-Object {
                    if ($_.createdDateTime)
                    {
                        [datetime]$_.createdDateTime
                    }
                } |
                Sort-Object
        )

        $S_FirstSeen = $null
        $S_LastSeen = $null
        if ($S_Timestamps.Count -gt 0)
        {
            $S_FirstSeen = $S_Timestamps[0]
            $S_LastSeen = $S_Timestamps[$S_Timestamps.Count - 1]
        }

        # Count successful vs failed attempts (errorCode 0 = success)
        # Use simple dot notation - works for both Hashtable and PSObject
        $S_SuccessCount = 0
        $S_CaSuccessCount = 0
        $S_CaFailureCount = 0
        $S_CaNotAppliedCount = 0
        foreach ($S_Record in $S_Records)
        {
            # Sign-in status (errorCode 0 = success)
            $S_ErrorCode = $null
            try
            {
                $S_ErrorCode = $S_Record.status.errorCode
            }
            catch
            {
            }

            if ($null -eq $S_ErrorCode)
            {
                try
                {
                    $S_ErrorCode = $S_Record['status']['errorCode']
                }
                catch
                {
                }
            }

            if ($null -ne $S_ErrorCode -and [int]$S_ErrorCode -eq 0)
            {
                $S_SuccessCount++
            }

            # Conditional Access status
            $S_CaStatus = $null
            try
            {
                $S_CaStatus = [string]$S_Record.conditionalAccessStatus
            }
            catch
            {
            }

            if ([string]::IsNullOrWhiteSpace($S_CaStatus))
            {
                try
                {
                    $S_CaStatus = [string]$S_Record['conditionalAccessStatus']
                }
                catch
                {
                }
            }

            switch -Regex ($S_CaStatus)
            {
                '^success$'    { $S_CaSuccessCount++ }
                '^failure$'    { $S_CaFailureCount++ }
                '^notApplied$' { $S_CaNotAppliedCount++ }
            }
        }

        $S_FailedCount = $S_Group.Count - $S_SuccessCount

        if ($S_SuccessCount -gt 0 -and $S_FailedCount -eq 0)
        {
            $S_StatusSummary = 'Success'
        }
        elseif ($S_FailedCount -gt 0 -and $S_SuccessCount -eq 0)
        {
            $S_StatusSummary = 'Failed'
        }
        else
        {
            $S_StatusSummary = 'Mixed'
        }

        if ($S_CaSuccessCount -gt 0 -and $S_CaFailureCount -eq 0)
        {
            $S_CaSummary = 'Success'
        }
        elseif ($S_CaFailureCount -gt 0 -and $S_CaSuccessCount -eq 0)
        {
            $S_CaSummary = 'Failure'
        }
        elseif ($S_CaSuccessCount -eq 0 -and $S_CaFailureCount -eq 0 -and $S_CaNotAppliedCount -gt 0)
        {
            $S_CaSummary = 'Not Applied'
        }
        elseif ($S_CaSuccessCount -gt 0 -or $S_CaFailureCount -gt 0)
        {
            $S_CaSummary = 'Mixed'
        }
        else
        {
            $S_CaSummary = 'N/A'
        }

        # Resolve country code to full name
        if ($S_Countries.Count -gt 0)
        {
            $S_CountryCode = $S_Countries[0]
        }
        else
        {
            $S_CountryCode = 'Unknown'
        }

        if ($S_CountryCodeMap.ContainsKey($S_CountryCode))
        {
            $S_CountryFullName = $S_CountryCodeMap[$S_CountryCode]
        }
        else
        {
            $S_CountryFullName = $S_CountryCode
        }

        if ($S_CountryCode -eq 'Unknown')
        {
            $S_CountryDisplay = 'Unknown'
        }
        else
        {
            $S_CountryDisplay = "$S_CountryFullName ($S_CountryCode)"
        }

        if ($S_States.Count -gt 0)
        {
            $S_StateValue = $S_States -join ' | '
        }
        else
        {
            $S_StateValue = $null
        }

        if ($S_Cities.Count -gt 0)
        {
            $S_CityValue = $S_Cities -join ' | '
        }
        else
        {
            $S_CityValue = $null
        }

        [PSCustomObject]@{
            IPAddress         = [string]$S_Group.Name
            SignInCount       = $S_Group.Count
            SuccessCount      = $S_SuccessCount
            FailedCount       = $S_FailedCount
            Status            = $S_StatusSummary
            CASuccessCount    = $S_CaSuccessCount
            CAFailureCount    = $S_CaFailureCount
            CANotAppliedCount = $S_CaNotAppliedCount
            CAStatus          = $S_CaSummary
            CountryCode       = $S_CountryCode
            CountryFullName   = $S_CountryFullName
            CountryOrRegion   = $S_CountryDisplay
            State             = $S_StateValue
            City              = $S_CityValue
            UniqueUsers       = $S_Users.Count
            SampleUsers       = ($S_Users | Select-Object -First 5) -join '; '
            FirstSeenUtc      = $S_FirstSeen
            LastSeenUtc       = $S_LastSeen
        }
    }

    $S_UniqueIpReport = @($S_UniqueIpReport) | Sort-Object -Property SignInCount -Descending

    # ---------------------------------------------------------------------------
    # Export
    # ---------------------------------------------------------------------------

    if (-not (Test-Path -LiteralPath $OutputPath))
    {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    $S_ReportTime = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
    $S_ReportDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $S_CsvName = "InteractiveUniqueIPs_${Days}d_${S_ReportTime}.csv"
    $S_CsvPath = Join-Path -Path $OutputPath -ChildPath $S_CsvName

    $S_UniqueIpReport | Export-Csv -LiteralPath $S_CsvPath -NoTypeInformation -Encoding UTF8

    # --- Build HTML report ---
    $S_TotalSignIns = ($S_UniqueIpReport | Measure-Object -Property SignInCount -Sum).Sum
    $S_UniqueCountries = @(
        $S_UniqueIpReport |
            ForEach-Object { $_.CountryOrRegion } |
            Where-Object { $_ -ne 'Unknown' } |
            Sort-Object -Unique
    ).Count

    $S_TableRows = ($S_UniqueIpReport | ForEach-Object {
        switch ($_.SignInCount)
        {
            { $_ -gt 500 } { $S_RiskClass = 'risk-high' }
            { $_ -gt 100 } { $S_RiskClass = 'risk-medium' }
            default { $S_RiskClass = 'risk-low' }
        }

        switch ($_.Status)
        {
            'Success' { $S_StatusClass = 'status-success' }
            'Failed'  { $S_StatusClass = 'status-failed' }
            default   { $S_StatusClass = 'status-mixed' }
        }

        switch ($_.CAStatus)
        {
            'Success'     { $S_CaStatusClass = 'status-success' }
            'Failure'     { $S_CaStatusClass = 'status-failed' }
            'Not Applied' { $S_CaStatusClass = 'status-na' }
            'Mixed'       { $S_CaStatusClass = 'status-mixed' }
            default       { $S_CaStatusClass = 'status-na' }
        }

        if ($_.Status -eq 'Mixed')
        {
            $S_StatusText = "Mixed ($($_.SuccessCount)/$($_.FailedCount))"
        }
        else
        {
            $S_StatusText = $_.Status
        }

        "<tr class=`"$S_RiskClass`"><td><code>$([System.Net.WebUtility]::HtmlEncode($_.IPAddress))</code></td><td>$([System.Net.WebUtility]::HtmlEncode($_.CountryOrRegion))</td><td>$([System.Net.WebUtility]::HtmlEncode($_.State -as [string]))</td><td>$([System.Net.WebUtility]::HtmlEncode($_.City -as [string]))</td><td>$($_.SignInCount)</td><td><span class=`"badge $S_StatusClass`">$S_StatusText</span></td><td><span class=`"badge $S_CaStatusClass`">$($_.CAStatus)</span></td><td class=`"col-success`">$($_.SuccessCount)</td><td class=`"col-failed`">$($_.FailedCount)</td><td>$($_.UniqueUsers)</td><td>$($_.FirstSeenUtc.ToString('yyyy-MM-dd HH:mm:ss'))</td><td>$($_.LastSeenUtc.ToString('yyyy-MM-dd HH:mm:ss'))</td></tr>"
    }) -join "`n"

    # --- Build country statistics ---
    $S_CountryStats = foreach ($S_Country in ($S_UniqueIpReport | Select-Object -ExpandProperty CountryFullName -Unique | Where-Object { $_ -ne 'Unknown' }))
    {
        $S_CountryIps = @($S_UniqueIpReport | Where-Object { $_.CountryFullName -eq $S_Country })
        $S_CountryCode = ($S_CountryIps | Select-Object -First 1).CountryCode
        $S_CountrySignIns = ($S_CountryIps | Measure-Object -Property SignInCount -Sum).Sum
        $S_CountrySuccess = ($S_CountryIps | Measure-Object -Property SuccessCount -Sum).Sum
        $S_CountryFailed = ($S_CountryIps | Measure-Object -Property FailedCount -Sum).Sum

        [PSCustomObject]@{
            Country     = $S_Country
            CountryCode = $S_CountryCode
            IPs         = $S_CountryIps.Count
            SignIns     = $S_CountrySignIns
            Success     = $S_CountrySuccess
            Failed      = $S_CountryFailed
            Users       = ($S_CountryIps | Measure-Object -Property UniqueUsers -Sum).Sum
        }
    }

    $S_CountryStats = @($S_CountryStats) | Sort-Object -Property SignIns -Descending

    $S_CountryTableRows = ($S_CountryStats | ForEach-Object {
        "<tr><td>$([System.Net.WebUtility]::HtmlEncode($_.Country)) <span class=`"country-code`">($([System.Net.WebUtility]::HtmlEncode($_.CountryCode)))</span></td><td>$($_.IPs)</td><td>$($_.SignIns)</td><td class=`"col-success`">$($_.Success)</td><td class=`"col-failed`">$($_.Failed)</td><td>$($_.Users)</td></tr>"
    }) -join "`n"

    $S_CountryFilterOptions = ($S_UniqueIpReport | Select-Object -ExpandProperty CountryFullName -Unique | Sort-Object | ForEach-Object {
        "<option value=`"$([System.Net.WebUtility]::HtmlEncode($_))`">$([System.Net.WebUtility]::HtmlEncode($_))</option>"
    }) -join "`n"

    $S_Html = @"
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
        <div class="subtitle">Generated: $S_ReportDate | Tenant: $S_TenantName | Tenant ID: $S_TenantId | Period: $Days days | Total IPs: $($S_UniqueIpReport.Count)</div>
    </div>

    <div class="summary-cards">
        <div class="card blue">
            <div class="label">Unique IP Addresses</div>
            <div class="value">$($S_UniqueIpReport.Count)</div>
        </div>
        <div class="card green">
            <div class="label">Total Sign-Ins</div>
            <div class="value">$S_TotalSignIns</div>
        </div>
        <div class="card orange">
            <div class="label">Unique Countries</div>
            <div class="value">$S_UniqueCountries</div>
        </div>
        <div class="card red">
            <div class="label">Avg per IP</div>
            <div class="value">$([math]::Round($S_TotalSignIns / [math]::Max($S_UniqueIpReport.Count, 1), 1))</div>
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
$S_CountryTableRows
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
$S_CountryFilterOptions
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
$S_TableRows
            </tbody>
        </table>
    </div>

    <div class="footer">
        CSV data: $(Split-Path $S_CsvPath -Leaf) | Report generated by Get-InteractiveSignInUniqueIPs.ps1
    </div>
</body>
</html>
"@

    $S_HtmlName = "InteractiveUniqueIPs_${Days}d_${S_ReportTime}.html"
    $S_HtmlPath = Join-Path -Path $OutputPath -ChildPath $S_HtmlName
    $S_Html | Out-File -FilePath $S_HtmlPath -Encoding UTF8

    Write-Host "`nHTML report saved to: $S_HtmlPath" -ForegroundColor Green
    Write-Host "`nUnique IP addresses: $($S_UniqueIpReport.Count)" -ForegroundColor Green
    Write-Host "Total sign-ins: $S_TotalSignIns" -ForegroundColor Green

    Write-Host "`nTop IPs by sign-in count:" -ForegroundColor Cyan
    $S_UniqueIpReport |
        Select-Object -First 20 -Property IPAddress, CountryOrRegion, SignInCount, LastSeenUtc |
        Format-Table -AutoSize
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
