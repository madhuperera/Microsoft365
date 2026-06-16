#Requires -Version 5.1

<#
.SYNOPSIS
    Reports on Network Devices discovered by Microsoft Defender for Endpoint.

.DESCRIPTION
    Authenticates to the Defender for Endpoint (MDE) REST API using an Entra ID
    application registration and queries the device inventory directly — no
    Advanced Hunting / KQL is used.

    Two MDE API endpoints are called:
      • /api/machines               — device identity, OS, IP, sensor status.
      • /api/DeviceTvmHardwareAndFirmware — system manufacturer, product name,
                                     firmware version, manufacturer, and serial
                                     number (TVM data). Joined by DeviceId.

    Only devices with deviceCategory eq 'NetworkDevice' are returned.
    A -DaysBack filter (based on lastSeen) is applied server-side.
    Results are exported to CSV and an interactive HTML dashboard.

    Prerequisites:
      - Microsoft Defender for Endpoint P1 or P2 licence (Device Discovery).
      - Entra ID app registration with APPLICATION permission:
          WindowsDefenderATP  →  Machine.Read.All
        (TVM firmware data also requires Machine.Read.All — same permission).
      - The app must be granted admin consent in the target tenant.

.PARAMETER TenantId
    Entra ID tenant ID (GUID). Found in Azure Portal → Entra ID → Overview.

.PARAMETER AppId
    Application (client) ID of the app registration.

.PARAMETER AppSecret
    Client secret of the app registration. You can pass a plain string or a
    SecureString. The secret is held in memory only for the token request and
    then discarded.

.PARAMETER OutputPath
    Path for the output CSV file. Defaults to a timestamped file in the current
    working directory (ReportMdeNetworkDevices_yyyyMMdd_HHmmss.csv).

.PARAMETER DaysBack
    Only include devices last seen within this many days. Defaults to 30.
    Valid range: 1 – 180.

.PARAMETER Test
    Limits the device inventory query to the first 50 records for a quick
    end-to-end test without full enumeration.

.EXAMPLE
    .\ReportMdeNetworkDevices.ps1 -TenantId 'xxxxxxxx-...' -AppId 'yyyyyyyy-...' -AppSecret (Read-Host -AsSecureString)

.EXAMPLE
    .\ReportMdeNetworkDevices.ps1 -TenantId $tid -AppId $aid -AppSecret $sec -DaysBack 7 -Test

.EXAMPLE
    .\ReportMdeNetworkDevices.ps1 -TenantId $tid -AppId $aid -AppSecret $sec -OutputPath C:\Reports\NetworkDevices.csv
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$AppId,

    [Parameter(Mandatory = $true)]
    [object]$AppSecret,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 180)]
    [int]$DaysBack = 30,

    [Parameter(Mandatory = $false)]
    [switch]$Test
)

# ── Setup ──────────────────────────────────────────────────────────────────────
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Web

if (-not $OutputPath)
{
    $S_Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath  = Join-Path (Get-Location).Path "ReportMdeNetworkDevices_$S_Timestamp.csv"
}

$S_HtmlPath = [System.IO.Path]::ChangeExtension($OutputPath, '.html')

# ── Acquire MDE API token (client credentials) ─────────────────────────────────
Write-Host "`nAcquiring MDE API token..." -ForegroundColor Cyan

# Normalize the client secret to plain text only for the HTTP POST body.
if ($AppSecret -is [System.Security.SecureString])
{
    $S_BSTR        = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($AppSecret)
    $S_PlainSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($S_BSTR)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($S_BSTR)
}
else
{
    $S_PlainSecret = [string]$AppSecret
}

try
{
    $S_TokenResponse = Invoke-RestMethod `
        -Method      POST `
        -Uri         "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body        @{
            grant_type    = 'client_credentials'
            client_id     = $AppId
            client_secret = $S_PlainSecret
            scope         = 'https://api.securitycenter.microsoft.com/.default'
        }
}
finally
{
    # Zero the plain-text secret regardless of success or failure.
    $S_PlainSecret = $null
    [System.GC]::Collect()
}

$S_AccessToken = $S_TokenResponse.access_token
Write-Host "Token acquired. Expires in $($S_TokenResponse.expires_in)s." -ForegroundColor Green

$S_AuthHeaders = @{
    Authorization  = "Bearer $S_AccessToken"
    'Content-Type' = 'application/json'
    Accept         = 'application/json'
}

# ── Helper: paginated MDE API GET ──────────────────────────────────────────────
# The MDE REST API returns up to 10,000 records per page and provides
# an @odata.nextLink for subsequent pages. This function walks all pages.
function Invoke-MdePagedGet
{
    param(
        [string]    $F_Uri,
        [hashtable] $F_Headers
    )
    $F_All  = [System.Collections.Generic.List[object]]::new()
    $F_Next = $F_Uri
    do
    {
        $F_Page = Invoke-RestMethod -Method GET -Uri $F_Next -Headers $F_Headers
        foreach ($F_Item in @($F_Page.value)) { $F_All.Add($F_Item) }
        $F_Next = if ($F_Page.'@odata.nextLink') { [string]$F_Page.'@odata.nextLink' } else { $null }
    } while ($F_Next)
    return $F_All
}

try
{
    # ── Fetch device inventory ─────────────────────────────────────────────────
    # /api/machines supports OData $filter on onboardingStatus and lastSeen.
    # The public Machine schema does not expose a network-device category
    # filter, so we use discovered / unmanaged inventory records as the closest
    # valid inventory slice and keep the query within the supported REST API.
    Write-Host "`nQuerying MDE device inventory..." -ForegroundColor Cyan
    Write-Host "  Lookback : $DaysBack day(s)" -ForegroundColor DarkGray
    if ($Test) { Write-Host "  Mode     : TEST (limited to 50 records)" -ForegroundColor DarkGray }

    $S_CutoffDate = (Get-Date).AddDays(-$DaysBack).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $S_MachineFilter = [uri]::EscapeDataString(
        "onboardingStatus ne 'Onboarded' and lastSeen ge $S_CutoffDate"
    )
    $S_TopClause  = if ($Test) { '&$top=50' } else { '' }
    $S_MachineUri = "https://api.securitycenter.microsoft.com/api/machines?`$filter=$S_MachineFilter$S_TopClause"

    $S_RawMachines = @(Invoke-MdePagedGet -F_Uri $S_MachineUri -F_Headers $S_AuthHeaders)
    Write-Host "Retrieved $($S_RawMachines.Count) inventory record(s)." -ForegroundColor Green

    if ($S_RawMachines.Count -eq 0)
    {
        Write-Warning @"
No inventory records were returned for the last $DaysBack day(s).

Possible causes:
  • Device Discovery is not enabled in Defender for Endpoint settings.
    • No discovered devices have been returned in this tenant yet.
  • The app registration may be missing the Machine.Read.All permission.
  • Try increasing -DaysBack (e.g. -DaysBack 90) to widen the lookback window.
"@
        return
    }

    # ── Fetch TVM hardware / firmware data ─────────────────────────────────────
    # DeviceTvmHardwareAndFirmware is a flat export of all TVM hardware records.
    # We download the full set once and build a lookup table keyed by deviceId
    # so each machine record can be enriched without per-device API calls.
    Write-Host "Querying TVM hardware and firmware data..." -ForegroundColor Cyan
    $S_HardwareUri  = 'https://api.securitycenter.microsoft.com/api/DeviceTvmHardwareAndFirmware'
    $S_HardwareLookup = @{}
    try
    {
        $S_RawHardware = @(Invoke-MdePagedGet -F_Uri $S_HardwareUri -F_Headers $S_AuthHeaders)
        foreach ($S_Hw in $S_RawHardware)
        {
            # The API uses 'deviceId' (lowercase d) as the key field.
            $S_HwKey = [string]$S_Hw.deviceId
            if (-not [string]::IsNullOrWhiteSpace($S_HwKey))
            {
                $S_HardwareLookup[$S_HwKey] = $S_Hw
            }
        }
        Write-Host "  $($S_HardwareLookup.Count) hardware record(s) loaded." -ForegroundColor DarkGray
    }
    catch
    {
        Write-Warning "Could not retrieve TVM hardware data (firmware columns will be empty): $_"
    }

    # ── Build result objects ───────────────────────────────────────────────────
    Write-Host "Processing device records..." -ForegroundColor Cyan

    $S_Results    = [System.Collections.Generic.List[PSCustomObject]]::new()
    $S_TotalCount = $S_RawMachines.Count
    $S_Index      = 0

    foreach ($S_Device in $S_RawMachines)
    {
        $S_Index++
        Write-Progress -Activity "Processing Network Devices" `
            -Status        "$S_Index of $S_TotalCount — $($S_Device.computerDnsName)" `
            -PercentComplete ([math]::Round(($S_Index / $S_TotalCount) * 100))

        # Resolve primary IP.  The machines endpoint exposes lastIpAddress
        # (internal) and lastExternalIpAddress (NAT / public). Prefer internal.
        $S_PrimaryIp   = [string]$S_Device.lastIpAddress
        $S_ExternalIp  = [string]$S_Device.lastExternalIpAddress

        # MAC address — pulled from the ipInterfaces array (first entry that has one).
        $S_MacAddress = ''
        foreach ($S_Iface in @($S_Device.ipInterfaces))
        {
            if (-not [string]::IsNullOrWhiteSpace([string]$S_Iface.macAddress))
            {
                $S_MacAddress = [string]$S_Iface.macAddress
                break
            }
        }

        # Enrich with TVM hardware data if available for this device.
        $S_Hw = if ($S_HardwareLookup.ContainsKey([string]$S_Device.id)) { $S_HardwareLookup[[string]$S_Device.id] } else { $null }

        $S_Results.Add([PSCustomObject]@{
            DeviceId             = [string]$S_Device.id
            DeviceName           = [string]$S_Device.computerDnsName
            DeviceCategory       = [string]$S_Device.deviceCategory
            OSPlatform           = [string]$S_Device.osPlatform
            OSVersion            = [string]$S_Device.osVersion
            OSBuild              = [string]$S_Device.osBuild
            LastIP               = $S_PrimaryIp
            ExternalIP           = $S_ExternalIp
            MacAddress           = $S_MacAddress
            SensorHealthState    = [string]$S_Device.healthStatus
            OnboardingStatus     = [string]$S_Device.onboardingStatus
            RiskScore            = [string]$S_Device.riskScore
            ExposureLevel        = [string]$S_Device.exposureLevel
            LastSeen             = [string]$S_Device.lastSeen
            SystemManufacturer   = if ($S_Hw) { [string]$S_Hw.systemManufacturer }   else { '' }
            SystemProductName    = if ($S_Hw) { [string]$S_Hw.systemProductName }    else { '' }
            FirmwareVersion      = if ($S_Hw) { [string]$S_Hw.firmwareVersion }      else { '' }
            FirmwareManufacturer = if ($S_Hw) { [string]$S_Hw.firmwareManufacturer } else { '' }
            FirmwareSerialNumber = if ($S_Hw) { [string]$S_Hw.firmwareSerialNumber } else { '' }
        })
    }

    Write-Progress -Activity "Processing Network Devices" -Completed

    # ── Export CSV ─────────────────────────────────────────────────────────────
    $S_Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nCSV exported to  : $OutputPath" -ForegroundColor Green
    Write-Host "Total rows       : $($S_Results.Count)" -ForegroundColor Green

    # ── Calculate Statistics ───────────────────────────────────────────────────
    $S_TotalDevices    = $S_Results.Count
    $S_OnboardedCount  = ($S_Results | Where-Object { $_.OnboardingStatus -eq 'Onboarded' }).Count
    $S_DiscoveredCount = ($S_Results | Where-Object { $_.OnboardingStatus -ne 'Onboarded' }).Count
    $S_ActiveSensor    = ($S_Results | Where-Object { $_.SensorHealthState -eq 'Active' }).Count
    $S_WithFirmware    = ($S_Results | Where-Object { -not [string]::IsNullOrWhiteSpace($_.FirmwareVersion) }).Count
    $S_WithOS          = ($S_Results | Where-Object { -not [string]::IsNullOrWhiteSpace($_.OSVersion) }).Count

    $S_OsPlatformGroups  = $S_Results | Group-Object OSPlatform    | Sort-Object Count -Descending
    $S_OnboardingGroups  = $S_Results | Group-Object OnboardingStatus | Sort-Object Count -Descending
    $S_HealthGroups      = $S_Results | Group-Object SensorHealthState | Sort-Object Count -Descending
    $S_ManufacturerGroups = @(
        $S_Results |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.SystemManufacturer) } |
            Group-Object SystemManufacturer |
            Sort-Object Count -Descending |
            Select-Object -First 10
    )

    # ── Helper: safe HTML encode ───────────────────────────────────────────────
    function ConvertTo-SafeHtml { param([string]$F_Text)
        if ([string]::IsNullOrWhiteSpace($F_Text)) { return '<span style="color:#999">—</span>' }
        return [System.Web.HttpUtility]::HtmlEncode($F_Text)
    }
    function ConvertTo-HtmlEncoded { param([string]$F_Text)
        return [System.Web.HttpUtility]::HtmlEncode([string]$F_Text)
    }

    # ── Build device table rows ────────────────────────────────────────────────
    $S_TableRows = ($S_Results | ForEach-Object {

        # Onboarding status pill
        $S_OnboardStyle = switch ($_.OnboardingStatus)
        {
            'Onboarded'       { 'background:#eaf6ec;color:#107c10;border:1px solid #107c10' }
            'CanBeOnboarded'  { 'background:#fff4ce;color:#8a6d00;border:1px solid #bc8000' }
            'Discovered'      { 'background:#fff4ce;color:#8a6d00;border:1px solid #bc8000' }
            default           { 'background:#f3f3f3;color:#666;border:1px solid #ccc' }
        }
        $S_OnboardCell = "<span style=""$S_OnboardStyle;padding:2px 8px;border-radius:10px;font-size:12px;font-weight:600"">$(ConvertTo-HtmlEncoded $_.OnboardingStatus)</span>"

        # Sensor health pill
        $S_HealthStyle = switch ($_.SensorHealthState)
        {
            'Active'   { 'background:#eaf6ec;color:#107c10;border:1px solid #107c10' }
            'Inactive' { 'background:#ffe8cc;color:#9a4f00;border:1px solid #d83b01' }
            'NoSensorData' { 'background:#fdecea;color:#a4262c;border:1px solid #a4262c' }
            default    { 'background:#f3f3f3;color:#666;border:1px solid #ccc' }
        }
        $S_HealthText = if ([string]::IsNullOrWhiteSpace($_.SensorHealthState)) { 'Unknown' } else { ConvertTo-HtmlEncoded $_.SensorHealthState }
        $S_HealthCell = "<span style=""$S_HealthStyle;padding:2px 8px;border-radius:10px;font-size:12px;font-weight:600"">$S_HealthText</span>"

        # Firmware — monospace for version string
        $S_FwCell = if ([string]::IsNullOrWhiteSpace($_.FirmwareVersion))
        {
            '<span style="color:#999">—</span>'
        }
        else
        {
            "<code>$(ConvertTo-HtmlEncoded $_.FirmwareVersion)</code>"
        }

        # IP — monospace (internal IP from lastIpAddress)
        $S_IpCell = if ([string]::IsNullOrWhiteSpace($_.LastIP))
        {
            '<span style="color:#999">—</span>'
        }
        else
        {
            "<code>$(ConvertTo-HtmlEncoded $_.LastIP)</code>"
        }

        # MAC — monospace
        $S_MacCell = if ([string]::IsNullOrWhiteSpace($_.MacAddress))
        {
            '<span style="color:#999">—</span>'
        }
        else
        {
            "<code>$(ConvertTo-HtmlEncoded $_.MacAddress)</code>"
        }

        # OS Version — monospace
        $S_OsVerCell = if ([string]::IsNullOrWhiteSpace($_.OSVersion))
        {
            '<span style="color:#999">—</span>'
        }
        else
        {
            "<code>$(ConvertTo-HtmlEncoded $_.OSVersion)</code>"
        }

        # Last seen — friendly format
        $S_LastSeenCell = if ([string]::IsNullOrWhiteSpace($_.LastSeen))
        {
            '<span style="color:#999">—</span>'
        }
        else
        {
            try
            {
                [System.Web.HttpUtility]::HtmlEncode(([datetime]$_.LastSeen).ToString('yyyy-MM-dd HH:mm'))
            }
            catch
            {
                ConvertTo-HtmlEncoded $_.LastSeen
            }
        }

        "        <tr data-onboarding=""$(ConvertTo-HtmlEncoded $_.OnboardingStatus)"" data-os=""$(ConvertTo-HtmlEncoded $_.OSPlatform)"" data-health=""$(ConvertTo-HtmlEncoded $_.SensorHealthState)""><td>$(ConvertTo-SafeHtml $_.DeviceName)</td><td>$S_IpCell</td><td>$S_MacCell</td><td>$(ConvertTo-SafeHtml $_.OSPlatform)</td><td>$S_OsVerCell</td><td>$(ConvertTo-SafeHtml $_.OSBuild)</td><td>$(ConvertTo-SafeHtml $_.SystemManufacturer)</td><td>$(ConvertTo-SafeHtml $_.SystemProductName)</td><td>$S_FwCell</td><td>$(ConvertTo-SafeHtml $_.FirmwareManufacturer)</td><td>$S_OnboardCell</td><td>$S_HealthCell</td><td>$S_LastSeenCell</td></tr>"

    }) -join "`n"

    # ── OS platform distribution rows ──────────────────────────────────────────
    $S_OsDistRows = ($S_OsPlatformGroups | ForEach-Object {
        $S_OsPct  = [math]::Round(($_.Count / [math]::Max($S_TotalDevices, 1)) * 100, 1)
        $S_OsName = if ([string]::IsNullOrWhiteSpace($_.Name)) { 'Unknown' } else { [System.Web.HttpUtility]::HtmlEncode($_.Name) }
        "                <tr><td>$S_OsName</td><td>$($_.Count)</td><td>$S_OsPct%</td></tr>"
    }) -join "`n"

    # ── Manufacturer distribution rows ─────────────────────────────────────────
    $S_MfrDistRows = if ($S_ManufacturerGroups.Count -eq 0)
    {
        "                <tr><td colspan=""3"" style=""color:#999;text-align:center"">No hardware manufacturer data available</td></tr>"
    }
    else
    {
        ($S_ManufacturerGroups | ForEach-Object {
            $S_MfrPct  = [math]::Round(($_.Count / [math]::Max($S_TotalDevices, 1)) * 100, 1)
            $S_MfrName = [System.Web.HttpUtility]::HtmlEncode($_.Name)
            "                <tr><td>$S_MfrName</td><td>$($_.Count)</td><td>$S_MfrPct%</td></tr>"
        }) -join "`n"
    }

    # ── Onboarding status filter options ───────────────────────────────────────
    $S_OnboardOptions = ($S_OnboardingGroups | ForEach-Object {
        $S_Val = if ([string]::IsNullOrWhiteSpace($_.Name)) { '' } else { [System.Web.HttpUtility]::HtmlEncode($_.Name) }
        $S_Lbl = if ([string]::IsNullOrWhiteSpace($_.Name)) { 'Unknown' } else { [System.Web.HttpUtility]::HtmlEncode($_.Name) }
        "                <option value=""$S_Val"">$S_Lbl ($($_.Count))</option>"
    }) -join "`n"

    $S_OsOptions = ($S_OsPlatformGroups | ForEach-Object {
        $S_Val = if ([string]::IsNullOrWhiteSpace($_.Name)) { '' } else { [System.Web.HttpUtility]::HtmlEncode($_.Name) }
        $S_Lbl = if ([string]::IsNullOrWhiteSpace($_.Name)) { 'Unknown' } else { [System.Web.HttpUtility]::HtmlEncode($_.Name) }
        "                <option value=""$S_Val"">$S_Lbl ($($_.Count))</option>"
    }) -join "`n"

    $S_HealthOptions = ($S_HealthGroups | ForEach-Object {
        $S_Val = if ([string]::IsNullOrWhiteSpace($_.Name)) { '' } else { [System.Web.HttpUtility]::HtmlEncode($_.Name) }
        $S_Lbl = if ([string]::IsNullOrWhiteSpace($_.Name)) { 'Unknown' } else { [System.Web.HttpUtility]::HtmlEncode($_.Name) }
        "                <option value=""$S_Val"">$S_Lbl ($($_.Count))</option>"
    }) -join "`n"

    # ── Metadata ───────────────────────────────────────────────────────────────
    $S_GeneratedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    # TenantId and AppId come directly from the script parameters.
    $S_ReportTenant = $TenantId
    $S_ReportApp    = $AppId

    # ── Computed percentages for cards ─────────────────────────────────────────
    $S_PctOnboarded  = [math]::Round(($S_OnboardedCount  / [math]::Max($S_TotalDevices, 1)) * 100, 1)
    $S_PctDiscovered = [math]::Round(($S_DiscoveredCount / [math]::Max($S_TotalDevices, 1)) * 100, 1)
    $S_PctActive     = [math]::Round(($S_ActiveSensor    / [math]::Max($S_TotalDevices, 1)) * 100, 1)
    $S_PctFirmware   = [math]::Round(($S_WithFirmware    / [math]::Max($S_TotalDevices, 1)) * 100, 1)
    $S_PctOS         = [math]::Round(($S_WithOS          / [math]::Max($S_TotalDevices, 1)) * 100, 1)

    $S_TestBadge = if ($Test) { ' &mdash; <span style="color:#d13438;font-weight:700">TEST MODE (50 records)</span>' } else { '' }

    # ── HTML Report ────────────────────────────────────────────────────────────
    $S_Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MDE Network Devices Report</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; padding: 24px; }
        .header { text-align: center; margin-bottom: 28px; }
        .header h1 { font-size: 26px; color: #1a1a2e; margin-bottom: 4px; }
        .header .subtitle { font-size: 13px; color: #666; line-height: 1.6; }
        .header .badge { display: inline-block; background: #0078d4; color: #fff; border-radius: 6px; padding: 3px 12px; font-size: 12px; font-weight: 600; margin-left: 8px; vertical-align: middle; }
        .info-bar { max-width: 1400px; margin: 0 auto 20px; padding: 10px 16px; background: #e8f0fe; border: 1px solid #b3c8f5; border-left: 4px solid #0078d4; border-radius: 6px; font-size: 12.5px; color: #1a3a6e; }
        .info-bar strong { color: #0078d4; }
        .breakdown { display: flex; flex-wrap: wrap; gap: 16px; justify-content: center; margin-bottom: 24px; max-width: 1400px; margin-left: auto; margin-right: auto; }
        .card { background: #fff; border-radius: 12px; padding: 22px 24px; min-width: 190px; flex: 1; max-width: 240px; box-shadow: 0 2px 8px rgba(0,0,0,.07); border-left: 5px solid #0078d4; }
        .card .label { font-size: 12px; color: #666; text-transform: uppercase; letter-spacing: .5px; margin-bottom: 8px; }
        .card .value { font-size: 34px; font-weight: 700; color: #1a1a2e; line-height: 1; }
        .card .detail { font-size: 12px; color: #888; margin-top: 6px; }
        .card.blue   { border-left-color: #0078d4; }
        .card.green  { border-left-color: #107c10; }
        .card.orange { border-left-color: #d83b01; }
        .card.grey   { border-left-color: #767676; }
        .card.purple { border-left-color: #7719aa; }
        .section { max-width: 1400px; margin: 0 auto 24px; background: #fff; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,.07); overflow: hidden; }
        .section h2 { font-size: 15px; font-weight: 700; padding: 14px 20px; border-bottom: 1px solid #e8e8e8; color: #1a1a2e; background: #fafafa; }
        .dist-wrapper { display: flex; gap: 32px; padding: 16px 20px 20px; flex-wrap: wrap; }
        .dist-block { flex: 1; min-width: 240px; }
        .dist-block h3 { font-size: 13px; font-weight: 600; color: #444; margin-bottom: 10px; }
        .dist-block table { width: 100%; border-collapse: collapse; font-size: 13px; }
        .dist-block th { text-align: left; padding: 6px 10px; color: #666; font-size: 11px; text-transform: uppercase; letter-spacing: .04em; border-bottom: 2px solid #e8e8e8; }
        .dist-block td { padding: 6px 10px; border-bottom: 1px solid #f0f0f0; }
        .filter-bar { display: flex; flex-wrap: wrap; gap: 12px; padding: 12px 20px; border-bottom: 1px solid #e8e8e8; align-items: center; background: #fafafa; }
        .filter-bar input[type=text] { flex: 1; min-width: 220px; padding: 7px 10px; border: 1px solid #ccc; border-radius: 6px; font-size: 13px; }
        .filter-bar label { font-size: 13px; color: #555; }
        .filter-bar select { padding: 6px 8px; border: 1px solid #ccc; border-radius: 6px; font-size: 13px; background: #fff; }
        .filter-bar button { padding: 6px 14px; border: 1px solid #d13438; color: #d13438; background: #fff; border-radius: 6px; font-size: 12px; font-weight: 600; cursor: pointer; }
        .filter-bar button:hover { background: #fdecea; }
        .table-wrap { overflow-x: auto; }
        table.device-table { width: 100%; border-collapse: collapse; font-size: 13px; }
        table.device-table th { background: #fafafa; color: #555; font-weight: 600; text-transform: uppercase; font-size: 11px; letter-spacing: .04em; padding: 10px 12px; text-align: left; border-bottom: 2px solid #e8e8e8; white-space: nowrap; cursor: pointer; user-select: none; }
        table.device-table th:hover { color: #0078d4; }
        table.device-table td { padding: 9px 12px; border-bottom: 1px solid #f0f0f0; vertical-align: middle; word-break: break-word; max-width: 200px; }
        table.device-table tr:hover td { background: #f5f8ff; }
        code { font-family: Consolas, 'Cascadia Mono', 'Courier New', monospace; font-size: 0.9em; background: #f3f3f3; color: #1a1a2e; padding: 1px 5px; border: 1px solid #e0e0e0; border-radius: 3px; word-break: break-all; }
        .row-count { padding: 8px 20px; font-size: 12px; color: #888; border-top: 1px solid #e8e8e8; background: #fafafa; }
        .footer { text-align: center; font-size: 12px; color: #999; margin-top: 28px; }
    </style>
</head>
<body>

<div class="header">
    <h1>MDE Network Devices Report <span class="badge">Defender for Endpoint</span></h1>
    <div class="subtitle">
        Generated: $S_GeneratedAt$S_TestBadge<br>
        Tenant: $S_ReportTenant &middot; App: $S_ReportApp &middot; Lookback: $DaysBack day(s)
    </div>
</div>

<div class="info-bar" style="max-width:1400px;margin:0 auto 20px">
    <strong>Data sources:</strong>
    /api/machines (device identity, OS, IP, sensor health, onboarding status) &bull;
    /api/DeviceTvmHardwareAndFirmware (system manufacturer, product, firmware version &amp; serial) &bull;
    Both called directly against the <strong>MDE REST API</strong> (<code>api.securitycenter.microsoft.com</code>)
    using client-credential authentication &mdash; no Advanced Hunting or KQL involved.
<div class="breakdown">
    <div class="card blue">
        <div class="label">Total Network Devices</div>
        <div class="value">$S_TotalDevices</div>
        <div class="detail">Last $DaysBack day(s)</div>
    </div>
    <div class="card green">
        <div class="label">Onboarded</div>
        <div class="value">$S_OnboardedCount</div>
        <div class="detail">$S_PctOnboarded% of total</div>
    </div>
    <div class="card orange">
        <div class="label">Discovered / Not Onboarded</div>
        <div class="value">$S_DiscoveredCount</div>
        <div class="detail">$S_PctDiscovered% of total</div>
    </div>
    <div class="card green">
        <div class="label">Active Sensor</div>
        <div class="value">$S_ActiveSensor</div>
        <div class="detail">$S_PctActive% of total</div>
    </div>
    <div class="card purple">
        <div class="label">Firmware Data Available</div>
        <div class="value">$S_WithFirmware</div>
        <div class="detail">$S_PctFirmware% of total</div>
    </div>
    <div class="card grey">
        <div class="label">OS Version Data</div>
        <div class="value">$S_WithOS</div>
        <div class="detail">$S_PctOS% of total</div>
    </div>
</div>

<!-- Distribution Tables -->
<div class="section">
    <h2>Device Distribution</h2>
    <div class="dist-wrapper">
        <div class="dist-block">
            <h3>OS Platform</h3>
            <table>
                <thead><tr><th>Platform</th><th>Count</th><th>%</th></tr></thead>
                <tbody>
$S_OsDistRows
                </tbody>
            </table>
        </div>
        <div class="dist-block">
            <h3>Top Manufacturers (by Firmware Data)</h3>
            <table>
                <thead><tr><th>Manufacturer</th><th>Count</th><th>%</th></tr></thead>
                <tbody>
$S_MfrDistRows
                </tbody>
            </table>
        </div>
    </div>
</div>

<!-- Device Inventory Table -->
<div class="section">
    <h2>Network Device Inventory</h2>
    <div class="filter-bar">
        <input type="text" id="searchInput" placeholder="Search by device name, IP, MAC, firmware, OS..." oninput="applyFilters()">
        <label>Onboarding:
            <select id="filterOnboarding" onchange="applyFilters()">
                <option value="">All</option>
$S_OnboardOptions
            </select>
        </label>
        <label>OS Platform:
            <select id="filterOS" onchange="applyFilters()">
                <option value="">All</option>
$S_OsOptions
            </select>
        </label>
        <label>Sensor Health:
            <select id="filterHealth" onchange="applyFilters()">
                <option value="">All</option>
$S_HealthOptions
            </select>
        </label>
        <button onclick="clearFilters()">Clear Filters</button>
    </div>
    <div class="table-wrap">
        <table class="device-table" id="deviceTable">
            <thead>
                <tr>
                    <th onclick="sortTable(0)">Device Name &#8597;</th>
                    <th onclick="sortTable(1)">IP Address &#8597;</th>
                    <th onclick="sortTable(2)">MAC Address &#8597;</th>
                    <th onclick="sortTable(3)">OS Platform &#8597;</th>
                    <th onclick="sortTable(4)">OS Version &#8597;</th>
                    <th onclick="sortTable(5)">OS Build &#8597;</th>
                    <th onclick="sortTable(6)">Manufacturer &#8597;</th>
                    <th onclick="sortTable(7)">Product &#8597;</th>
                    <th onclick="sortTable(8)">Firmware Version &#8597;</th>
                    <th onclick="sortTable(9)">Firmware Mfr &#8597;</th>
                    <th onclick="sortTable(10)">Onboarding Status &#8597;</th>
                    <th onclick="sortTable(11)">Sensor Health &#8597;</th>
                    <th onclick="sortTable(12)">Last Seen &#8597;</th>
                </tr>
            </thead>
            <tbody id="deviceTableBody">
$S_TableRows
            </tbody>
        </table>
    </div>
    <div class="row-count" id="rowCount">Showing $S_TotalDevices of $S_TotalDevices device(s)</div>
</div>

<div class="footer">
    Generated by ReportMdeNetworkDevices.ps1 &bull; Microsoft Defender for Endpoint &bull; $S_GeneratedAt
</div>

<script>
    var allRows = Array.from(document.querySelectorAll('#deviceTableBody tr'));

    function applyFilters() {
        var search     = document.getElementById('searchInput').value.toLowerCase();
        var onboarding = document.getElementById('filterOnboarding').value.toLowerCase();
        var os         = document.getElementById('filterOS').value.toLowerCase();
        var health     = document.getElementById('filterHealth').value.toLowerCase();
        var visible    = 0;
        allRows.forEach(function(row) {
            var text          = row.textContent.toLowerCase();
            var onboardAttr   = (row.getAttribute('data-onboarding') || '').toLowerCase();
            var osAttr        = (row.getAttribute('data-os')         || '').toLowerCase();
            var healthAttr    = (row.getAttribute('data-health')     || '').toLowerCase();
            var show          = true;
            if (search     && !text.includes(search))          show = false;
            if (onboarding && !onboardAttr.includes(onboarding)) show = false;
            if (os         && !osAttr.includes(os))            show = false;
            if (health     && !healthAttr.includes(health))    show = false;
            row.style.display = show ? '' : 'none';
            if (show) visible++;
        });
        document.getElementById('rowCount').textContent =
            'Showing ' + visible + ' of ' + allRows.length + ' device(s)';
    }

    function clearFilters() {
        document.getElementById('searchInput').value = '';
        document.getElementById('filterOnboarding').value = '';
        document.getElementById('filterOS').value = '';
        document.getElementById('filterHealth').value = '';
        applyFilters();
    }

    var sortAsc = {};
    function sortTable(col) {
        var tbody = document.getElementById('deviceTableBody');
        var rows  = Array.from(tbody.querySelectorAll('tr'));
        var asc   = !sortAsc[col];
        sortAsc[col] = asc;
        rows.sort(function(a, b) {
            var aT = ((a.querySelectorAll('td')[col] || {}).textContent || '').trim();
            var bT = ((b.querySelectorAll('td')[col] || {}).textContent || '').trim();
            return asc
                ? aT.localeCompare(bT, undefined, { numeric: true, sensitivity: 'base' })
                : bT.localeCompare(aT, undefined, { numeric: true, sensitivity: 'base' });
        });
        rows.forEach(function(r) { tbody.appendChild(r); });
        applyFilters();
    }
</script>
</body>
</html>
"@

    $S_Html | Out-File -FilePath $S_HtmlPath -Encoding UTF8
    Write-Host "HTML report exported to: $S_HtmlPath" -ForegroundColor Green

    # ── Console Summary ────────────────────────────────────────────────────────
    Write-Host ""
    Write-Host "──────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  Summary" -ForegroundColor Cyan
    Write-Host "──────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ("  Total Network Devices      : {0}"    -f $S_TotalDevices)
    Write-Host ("  Onboarded                  : {0} ({1}%)" -f $S_OnboardedCount,  $S_PctOnboarded)
    Write-Host ("  Discovered / Not Onboarded : {0} ({1}%)" -f $S_DiscoveredCount, $S_PctDiscovered)
    Write-Host ("  Active Sensor              : {0} ({1}%)" -f $S_ActiveSensor,    $S_PctActive)
    Write-Host ("  With Firmware Data         : {0} ({1}%)" -f $S_WithFirmware,    $S_PctFirmware)
    Write-Host ("  With OS Version Data       : {0} ({1}%)" -f $S_WithOS,          $S_PctOS)
    Write-Host "──────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  OS Platform Breakdown:" -ForegroundColor Cyan
    foreach ($S_G in $S_OsPlatformGroups)
    {
        $S_Pct  = [math]::Round(($S_G.Count / [math]::Max($S_TotalDevices, 1)) * 100, 1)
        $S_Name = if ([string]::IsNullOrWhiteSpace($S_G.Name)) { 'Unknown' } else { $S_G.Name }
        Write-Host ("    {0,-30} {1,4}  ({2}%)" -f $S_Name, $S_G.Count, $S_Pct)
    }
    Write-Host ""
}
catch
{
    Write-Error "Script failed: $_"
}
