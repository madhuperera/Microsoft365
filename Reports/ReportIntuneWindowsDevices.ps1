#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
	Reports on Intune-managed Windows devices grouped by OS generation and build,
	matching the OsGeneration / WinBuild / Total breakdown commonly used for
	Windows servicing reviews.

.DESCRIPTION
	Connects to Microsoft Graph and retrieves all Intune-managed Windows devices.
	For each device the script captures user, model, full OS version, parsed
	OS generation (Windows 10 / Windows 11 / Other Windows), Windows build number,
	last sync date and compliance state. Outputs a CSV file and an HTML report
	showing the OsGeneration + WinBuild + Total table, version spread cards and
	a sortable / filterable device table.

	Optionally accepts minimum supported build numbers for Windows 10 and
	Windows 11. Devices on a build lower than the supplied minimum are flagged
	as "Outdated" in the report.

.PARAMETER MinimumSupportedWindows10Build
	Optional. Minimum supported Windows 10 build number (e.g. 19045 for 22H2).
	Devices with a build lower than this value are flagged as Outdated.

.PARAMETER MinimumSupportedWindows11Build
	Optional. Minimum supported Windows 11 build number (e.g. 22631 for 23H2,
	26100 for 24H2). Devices with a build lower than this value are flagged
	as Outdated.

.PARAMETER ReportPath
	Folder for the output reports. If omitted the current working directory is used.

.EXAMPLE
	.\ReportIntuneWindowsDevices.ps1

.EXAMPLE
	.\ReportIntuneWindowsDevices.ps1 -MinimumSupportedWindows10Build 19045 -MinimumSupportedWindows11Build 22631

.EXAMPLE
	.\ReportIntuneWindowsDevices.ps1 -MinimumSupportedWindows11Build 26100 -ReportPath C:\Reports
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 999999)]
    [int]$MinimumSupportedWindows10Build,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 999999)]
    [int]$MinimumSupportedWindows11Build,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ReportPath
)

$ErrorActionPreference = "Stop"

$S_ReportPath = $ReportPath

$S_RequiredGraphScopes = @(
    'DeviceManagementManagedDevices.Read.All'
    'Organization.Read.All'
)

$S_GraphRequestDelayMilliseconds = 5

try
{
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication))
    {
        throw "Microsoft.Graph.Authentication module is not installed. Install it using Install-Module Microsoft.Graph -Scope CurrentUser."
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    # --- Connect to Graph ---
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
            Disconnect-MgGraph | Out-Null
            Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome -ErrorAction Stop | Out-Null
        }
    }
    else
    {
        Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome -ErrorAction Stop | Out-Null
    }
    $S_ExistingContext = Get-MgContext

    Write-Host ""
    Write-Host "Active Graph context:" -ForegroundColor Cyan
    Write-Host "  Account    : $($S_ExistingContext.Account)" -ForegroundColor Cyan
    Write-Host "  TenantId   : $($S_ExistingContext.TenantId)" -ForegroundColor Cyan
    Write-Host "  Scopes     : $($S_ExistingContext.Scopes -join ', ')" -ForegroundColor Cyan
    Write-Host ""

    $S_ContextConfirmation = Read-Host "Proceed with this Graph context? [Y] Yes  [N] No  (Default: N)"
    if ([string]::IsNullOrWhiteSpace($S_ContextConfirmation))
    {
        $S_ContextConfirmation = 'N'
    }
    if ($S_ContextConfirmation.ToUpperInvariant() -ne 'Y')
    {
        throw "Operation cancelled. Please reconnect to the correct tenant and account, then run again."
    }

    # --- Tenant info ---
    $S_TenantDisplayName = $null
    try
    {
        $S_OrgResp = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization' -ErrorAction Stop
        if ($S_OrgResp.value)
        {
            $S_TenantDisplayName = $S_OrgResp.value[0].displayName
        }
    }
    catch
    {
    }
    if (-not $S_TenantDisplayName)
    {
        $S_TenantDisplayName = $S_ExistingContext.TenantId
    }
    $S_TenantId = if ($S_ExistingContext.TenantId)
    {
        $S_ExistingContext.TenantId
    }
    else
    {
        'Unknown'
    }

    # --- Fetch managed Windows devices ---
    Write-Host "Fetching Intune-managed Windows devices..." -ForegroundColor Cyan
    $S_Select = 'id,deviceName,userPrincipalName,userDisplayName,operatingSystem,osVersion,model,manufacturer,enrolledDateTime,lastSyncDateTime,complianceState,managedDeviceOwnerType,joinType,serialNumber'
    $S_Filter = "operatingSystem eq 'Windows'"
    $S_EncodedFilter = [System.Uri]::EscapeDataString($S_Filter)
    $S_Uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=$S_EncodedFilter&`$select=$S_Select&`$top=200"

    $S_Devices = New-Object System.Collections.Generic.List[object]
    do
    {
        $S_Resp = Invoke-MgGraphRequest -Method GET -Uri $S_Uri -ErrorAction Stop
        if ($S_Resp.value)
        {
            foreach ($S_D in $S_Resp.value)
            {
                $S_Devices.Add([pscustomobject]$S_D) | Out-Null
            }
        }
        $S_Uri = $S_Resp.'@odata.nextLink'
    } while ($S_Uri)

    Write-Host ("  Retrieved {0} Windows devices" -f $S_Devices.Count) -ForegroundColor Green

    # --- Build report rows ---
    # OS version format from Intune is typically "10.0.19045.4046".
    # Win10 and Win11 both report major 10; Win11 is identified by build >= 22000.
    $S_Now = Get-Date
    $S_Report = foreach ($S_D in $S_Devices)
    {
        $S_RawVer = if ($S_D.osVersion)
        {
            [string]$S_D.osVersion
        }
        else
        {
            ''
        }

        $S_WinBuild = $null
        if ($S_RawVer -match '^\s*\d+\.\d+\.(\d+)')
        {
            $S_WinBuild = [int]$Matches[1]
        }

        $S_OsGeneration = 'Other Windows'
        if ($null -ne $S_WinBuild)
        {
            if ($S_WinBuild -ge 22000)
            {
                $S_OsGeneration = 'Windows 11'
            }
            elseif ($S_WinBuild -ge 10240)
            {
                $S_OsGeneration = 'Windows 10'
            }
        }

        $S_Threshold = $null
        if ($S_OsGeneration -eq 'Windows 10' -and $PSBoundParameters.ContainsKey('MinimumSupportedWindows10Build'))
        {
            $S_Threshold = $MinimumSupportedWindows10Build
        }
        elseif ($S_OsGeneration -eq 'Windows 11' -and $PSBoundParameters.ContainsKey('MinimumSupportedWindows11Build'))
        {
            $S_Threshold = $MinimumSupportedWindows11Build
        }

        $S_SupportStatus = 'Unknown'
        if ($null -eq $S_WinBuild)
        {
            $S_SupportStatus = 'Unknown'
        }
        elseif ($null -ne $S_Threshold)
        {
            if ($S_WinBuild -lt $S_Threshold)
            {
                $S_SupportStatus = 'Outdated'
            }
            else
            {
                $S_SupportStatus = 'Supported'
            }
        }
        else
        {
            $S_SupportStatus = 'NoThreshold'
        }

        $S_LastSync = if ($S_D.lastSyncDateTime)
        {
            [datetime]$S_D.lastSyncDateTime
        }
        else
        {
            $null
        }
        $S_DaysSinceSync = if ($S_LastSync)
        {
            [int]($S_Now - $S_LastSync).TotalDays
        }
        else
        {
            $null
        }

        [pscustomobject]@{
            DeviceName        = $S_D.deviceName
            User              = if ($S_D.userDisplayName)
            {
                $S_D.userDisplayName
            }
            else
            {
                $S_D.userPrincipalName
            }
            UserPrincipalName = $S_D.userPrincipalName
            OsGeneration      = $S_OsGeneration
            OSVersion         = $S_RawVer
            WinBuild          = $S_WinBuild
            MinimumSupported  = $S_Threshold
            SupportStatus     = $S_SupportStatus
            Manufacturer      = $S_D.manufacturer
            Model             = $S_D.model
            Ownership         = $S_D.managedDeviceOwnerType
            JoinType          = $S_D.joinType
            ComplianceState   = $S_D.complianceState
            EnrolledDateTime  = $S_D.enrolledDateTime
            LastSyncDateTime  = $S_D.lastSyncDateTime
            DaysSinceLastSync = $S_DaysSinceSync
            SerialNumber      = $S_D.serialNumber
        }
    }

    # --- Stats ---
    $S_TotalDevices = $S_Report.Count
    $S_Win10Devices = $S_Report | Where-Object { $_.OsGeneration -eq 'Windows 10' }
    $S_Win11Devices = $S_Report | Where-Object { $_.OsGeneration -eq 'Windows 11' }
    $S_OtherDevices = $S_Report | Where-Object { $_.OsGeneration -eq 'Other Windows' }

    $S_TotalWin10 = $S_Win10Devices.Count
    $S_TotalWin11 = $S_Win11Devices.Count
    $S_TotalOther = $S_OtherDevices.Count

    $S_OutdatedWin10 = ($S_Win10Devices | Where-Object { $_.SupportStatus -eq 'Outdated' }).Count
    $S_OutdatedWin11 = ($S_Win11Devices | Where-Object { $_.SupportStatus -eq 'Outdated' }).Count
    $S_TotalOutdated = $S_OutdatedWin10 + $S_OutdatedWin11

    # --- Build OsGeneration + WinBuild breakdown (matches the Intune-style table) ---
    $S_GenerationOrder = @{ 'Other Windows' = 0; 'Windows 10' = 1; 'Windows 11' = 2 }
    $S_Breakdown = $S_Report | Group-Object OsGeneration, WinBuild | ForEach-Object {
        $S_First = $_.Group[0]
        $S_Gen = $S_First.OsGeneration
        $S_Build = $S_First.WinBuild
        $S_Min = $null
        if ($S_Gen -eq 'Windows 10' -and $PSBoundParameters.ContainsKey('MinimumSupportedWindows10Build'))
        {
            $S_Min = $MinimumSupportedWindows10Build
        }
        elseif ($S_Gen -eq 'Windows 11' -and $PSBoundParameters.ContainsKey('MinimumSupportedWindows11Build'))
        {
            $S_Min = $MinimumSupportedWindows11Build
        }
        $S_Outdated = ($null -ne $S_Min -and $null -ne $S_Build -and $S_Build -lt $S_Min)
        [pscustomobject]@{
            OsGeneration = $S_Gen
            WinBuild     = if ($null -eq $S_Build)
            {
                0
            }
            else
            {
                $S_Build
            }
            Total        = $_.Count
            Outdated     = $S_Outdated
        }
    } | Sort-Object @{Expression = { $S_GenerationOrder[$_.OsGeneration] } }, WinBuild

    # --- Output paths ---
    if (-not $S_ReportPath)
    {
        $S_ReportPath = (Get-Location).Path
    }
    $S_ReportFolder = if (Test-Path $S_ReportPath -PathType Container)
    {
        $S_ReportPath
    }
    else
    {
        Split-Path -Parent $S_ReportPath
    }
    if ($S_ReportFolder -and -not (Test-Path $S_ReportFolder))
    {
        New-Item -ItemType Directory -Path $S_ReportFolder -Force | Out-Null
    }
    $S_Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $S_CsvFile = Join-Path $S_ReportFolder ("ReportIntuneWindowsDevices_{0}.csv" -f $S_Timestamp)
    $S_BreakdownCsv = Join-Path $S_ReportFolder ("ReportIntuneWindowsDevices_Breakdown_{0}.csv" -f $S_Timestamp)
    $S_HtmlFile = Join-Path $S_ReportFolder ("ReportIntuneWindowsDevices_{0}.html" -f $S_Timestamp)

    # --- CSV exports ---
    $S_Report | Sort-Object OsGeneration, WinBuild, DeviceName | Export-Csv -Path $S_CsvFile -NoTypeInformation -Encoding UTF8
    $S_Breakdown | Select-Object OsGeneration, WinBuild, Total, Outdated | Export-Csv -Path $S_BreakdownCsv -NoTypeInformation -Encoding UTF8

    # --- HTML helpers ---
    $S_Enc = {
        param($s)
        if ($null -eq $s -or $s -eq '')
        {
            '-'
        }
        else
        {
            [System.Net.WebUtility]::HtmlEncode([string]$s)
        }
    }
    $S_ReportDate = Get-Date -Format "dd MMM yyyy HH:mm"

    $S_Win10Threshold = if ($PSBoundParameters.ContainsKey('MinimumSupportedWindows10Build'))
    {
        $MinimumSupportedWindows10Build
    }
    else
    {
        $null
    }
    $S_Win11Threshold = if ($PSBoundParameters.ContainsKey('MinimumSupportedWindows11Build'))
    {
        $MinimumSupportedWindows11Build
    }
    else
    {
        $null
    }
    $S_Win10ThresholdDisp = if ($null -ne $S_Win10Threshold)
    {
        $S_Win10Threshold
    }
    else
    {
        'not set'
    }
    $S_Win11ThresholdDisp = if ($null -ne $S_Win11Threshold)
    {
        $S_Win11Threshold
    }
    else
    {
        'not set'
    }

    # Breakdown table rows (the OsGeneration / WinBuild / Total view)
    $S_BreakdownRows = ($S_Breakdown | ForEach-Object {
            $S_Cls = if ($_.Outdated)
            {
                ' class="outdated-row"'
            }
            else
            {
                ''
            }
            $S_BuildText = if ($_.WinBuild -eq 0)
            {
                '-'
            }
            else
            {
                $_.WinBuild
            }
            $S_Badge = if ($_.Outdated)
            {
                " <span class='badge badge-inactive'>Outdated</span>"
            }
            else
            {
                ''
            }
            "<tr$S_Cls><td>$(& $S_Enc $_.OsGeneration)$S_Badge</td><td>$S_BuildText</td><td>$($_.Total)</td></tr>"
        }) -join "`n"

    # Build version spread cards per OsGeneration
    function Build-SpreadCardsHtml
    {
        param([object[]]$F_Items, [string]$F_Label)
        if (-not $F_Items -or $F_Items.Count -eq 0)
        {
            return "<div class='dist-card'><div class='dist-label'>No $F_Label devices</div><div class='dist-value'>0</div></div>"
        }
        ($F_Items | ForEach-Object {
            $F_Cls = if ($_.Outdated)
            {
                'dist-card outdated'
            }
            else
            {
                'dist-card'
            }
            $F_Badge = if ($_.Outdated)
            {
                "<div class='outdated-badge'>Outdated</div>"
            }
            else
            {
                ''
            }
            $F_BuildText = if ($_.WinBuild -eq 0)
            {
                'Unknown'
            }
            else
            {
                $_.WinBuild
            }
            "<div class='$F_Cls'><div class='dist-label'>$F_Label $F_BuildText</div><div class='dist-value'>$($_.Total)</div>$F_Badge</div>"
        }) -join "`n"
    }

    $S_Win10Items = $S_Breakdown | Where-Object { $_.OsGeneration -eq 'Windows 10' }
    $S_Win11Items = $S_Breakdown | Where-Object { $_.OsGeneration -eq 'Windows 11' }
    $S_OtherItems = $S_Breakdown | Where-Object { $_.OsGeneration -eq 'Other Windows' }

    $S_Win10CardsHtml = Build-SpreadCardsHtml -F_Items $S_Win10Items -F_Label 'Build'
    $S_Win11CardsHtml = Build-SpreadCardsHtml -F_Items $S_Win11Items -F_Label 'Build'
    $S_OtherCardsHtml = Build-SpreadCardsHtml -F_Items $S_OtherItems -F_Label 'Build'

    # Build chart data JSON
    function ConvertTo-ChartJson
    {
        param([object[]]$F_Items)
        if (-not $F_Items -or $F_Items.Count -eq 0)
        {
            return '{"labels":[],"data":[],"outdated":[]}'
        }
        $F_Labels = ($F_Items | ForEach-Object {
                $F_B = if ($_.WinBuild -eq 0)
                {
                    'Unknown'
                }
                else
                {
                    [string]$_.WinBuild
                }
                '"' + $F_B + '"'
            }) -join ','
        $F_Counts = ($F_Items | ForEach-Object { $_.Total }) -join ','
        $F_OutFlag = ($F_Items | ForEach-Object {
                if ($_.Outdated)
                {
                    'true'
                }
                else
                {
                    'false'
                }
            }) -join ','
        "{`"labels`":[$F_Labels],`"data`":[$F_Counts],`"outdated`":[$F_OutFlag]}"
    }

    $S_Win10ChartJson = ConvertTo-ChartJson -F_Items $S_Win10Items
    $S_Win11ChartJson = ConvertTo-ChartJson -F_Items $S_Win11Items
    $S_OtherChartJson = ConvertTo-ChartJson -F_Items $S_OtherItems

    # Device table rows
    $S_TableRows = ($S_Report | Sort-Object OsGeneration, WinBuild, DeviceName | ForEach-Object {
            $S_DaysVal = if ($null -ne $_.DaysSinceLastSync)
            {
                $_.DaysSinceLastSync
            }
            else
            {
                -1
            }
            $S_LastSyncDisp = if ($_.LastSyncDateTime)
            {
                ([datetime]$_.LastSyncDateTime).ToString("dd MMM yyyy")
            }
            else
            {
                '-'
            }
            $S_EnrolDisp = if ($_.EnrolledDateTime)
            {
                ([datetime]$_.EnrolledDateTime).ToString("dd MMM yyyy")
            }
            else
            {
                '-'
            }
            $S_StatusClass = switch ($_.SupportStatus)
            {
                'Outdated' { 'badge-inactive' }
                'Supported' { 'badge-active' }
                'NoThreshold' { 'badge-disabled' }
                default { 'badge-disabled' }
            }
            $S_StatusText = if ($_.SupportStatus -eq 'NoThreshold')
            {
                'No Threshold'
            }
            else
            {
                $_.SupportStatus
            }
            $S_CompClass = switch ($_.ComplianceState)
            {
                'compliant' { 'badge-active' }
                'noncompliant' { 'badge-inactive' }
                default { 'badge-disabled' }
            }
            $S_BuildText = if ($null -ne $_.WinBuild)
            {
                $_.WinBuild
            }
            else
            {
                '-'
            }
            $S_DaysSinceSyncDisplay = if ($S_DaysVal -ge 0)
            {
                "$S_DaysVal days"
            }
            else
            {
                'Never'
            }
            $S_RowAttr = "data-generation=`"$($_.OsGeneration)`" data-status=`"$($_.SupportStatus)`""
            "<tr $S_RowAttr>" +
            "<td>$(& $S_Enc $_.DeviceName)</td>" +
            "<td>$(& $S_Enc $_.User)</td>" +
            "<td>$(& $S_Enc $_.OsGeneration)</td>" +
            "<td>$(& $S_Enc $_.OSVersion)</td>" +
            "<td>$S_BuildText</td>" +
            "<td><span class='badge $S_StatusClass'>$S_StatusText</span></td>" +
            "<td>$(& $S_Enc $_.Manufacturer)</td>" +
            "<td>$(& $S_Enc $_.Model)</td>" +
            "<td>$(& $S_Enc $_.Ownership)</td>" +
            "<td>$(& $S_Enc $_.JoinType)</td>" +
            "<td><span class='badge $S_CompClass'>$(& $S_Enc $_.ComplianceState)</span></td>" +
            "<td>$S_EnrolDisp</td>" +
            "<td>$S_LastSyncDisp</td>" +
            "<td>$S_DaysSinceSyncDisplay</td>" +
            "</tr>"
        }) -join "`n"

    $S_PctOutdated = if ($S_TotalDevices -gt 0)
    {
        [math]::Round(($S_TotalOutdated / $S_TotalDevices) * 100, 1)
    }
    else
    {
        0
    }

    # --- HTML report ---
    $S_Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Intune Windows Devices Report</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; padding: 30px; }
  .header { background: linear-gradient(135deg, #1a1a2e, #16213e); color: #fff; padding: 30px 40px; border-radius: 12px; margin-bottom: 30px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 16px; }
  .header-left h1 { font-size: 1.6em; margin-bottom: 6px; }
  .header-left p { font-size: 0.9em; opacity: 0.85; }
  .header-right { font-size: 0.9em; opacity: 0.9; text-align: right; }
  .header-right strong { color: #ffd166; }

  .section-title { font-size: 1.15em; font-weight: 600; color: #1a1a2e; margin: 0 0 16px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0; }

  .summary-cards { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .card { background: #fff; border-radius: 10px; padding: 24px 30px; flex: 1; min-width: 180px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
  .card .label { font-size: 0.85em; color: #777; text-transform: uppercase; letter-spacing: 0.5px; }
  .card .value { font-size: 2em; font-weight: 700; margin-top: 6px; }
  .card .sub { font-size: 0.78em; color: #999; margin-top: 2px; }

  .dist-section { margin-bottom: 30px; }
  .dist-cards { display: flex; gap: 14px; flex-wrap: wrap; }
  .dist-card { background: #fff; border-radius: 10px; padding: 16px 22px; min-width: 130px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); border-left: 4px solid #3498db; text-align: center; position: relative; }
  .dist-card .dist-label { font-size: 0.8em; color: #555; margin-bottom: 6px; text-transform: uppercase; letter-spacing: 0.3px; }
  .dist-card .dist-value { font-size: 1.6em; font-weight: 700; color: #1a1a2e; }
  .dist-card.outdated { border-left-color: #e74c3c; background: #fff5f5; }
  .dist-card.outdated .dist-value { color: #c0392b; }
  .outdated-badge { display: inline-block; margin-top: 6px; padding: 2px 8px; font-size: 0.7em; font-weight: 700; color: #fff; background: #e74c3c; border-radius: 10px; text-transform: uppercase; letter-spacing: 0.4px; }

  .charts-row { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .chart-section { background: #fff; border-radius: 10px; padding: 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); flex: 1; min-width: 320px; }
  .chart-section h2 { font-size: 1.05em; margin-bottom: 16px; color: #1a1a2e; }
  .chart-container { max-width: 380px; margin: 0 auto; }

  .table-section { background: #fff; border-radius: 10px; padding: 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 30px; overflow-x: auto; }
  .table-section h2 { font-size: 1.1em; margin-bottom: 16px; color: #1a1a2e; }
  .table-controls { display: flex; gap: 14px; margin-bottom: 16px; flex-wrap: wrap; align-items: center; }
  .table-controls input[type="text"] { padding: 8px 14px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; min-width: 260px; }
  .table-controls select { padding: 8px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; background: #fff; }
  .table-controls .count-label { font-size: 0.85em; color: #777; margin-left: auto; }
  table { width: 100%; border-collapse: collapse; font-size: 0.85em; }
  th { background: #1a1a2e; color: #fff; padding: 10px 12px; text-align: left; cursor: pointer; user-select: none; white-space: nowrap; position: sticky; top: 0; }
  th:hover { background: #2c3e50; }
  td { padding: 9px 12px; border-bottom: 1px solid #eee; white-space: nowrap; }
  tr:hover td { background: #f8f9fa; }
  tr.hidden-row { display: none; }
  tr.outdated-row td { background: #fff5f5; }
  tr.outdated-row:hover td { background: #ffe8e8; }

  .badge { padding: 3px 10px; border-radius: 12px; font-size: 0.8em; font-weight: 600; }
  .badge-active { background: #d4edda; color: #155724; }
  .badge-inactive { background: #f8d7da; color: #721c24; }
  .badge-disabled { background: #e2e3e5; color: #495057; }

  .breakdown-table { background: #fff; border-radius: 10px; padding: 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 30px; }
  .breakdown-table table { max-width: 600px; }
  .breakdown-table h2 { font-size: 1.1em; margin-bottom: 16px; color: #1a1a2e; }

  .footer { text-align: center; font-size: 0.8em; color: #999; margin-top: 20px; }
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <h1>Intune Windows Devices Report</h1>
    <p>Tenant: $(& $S_Enc $S_TenantDisplayName) ($S_TenantId) &nbsp;|&nbsp; Generated: $S_ReportDate</p>
  </div>
  <div class="header-right">
    Min supported Windows 10 build: <strong>$S_Win10ThresholdDisp</strong><br/>
    Min supported Windows 11 build: <strong>$S_Win11ThresholdDisp</strong>
  </div>
</div>

<!-- OVERVIEW -->
<div class="section-title">Overview</div>
<div class="summary-cards">
  <div class="card"><div class="label">Total Windows Devices</div><div class="value" style="color:#1a1a2e;">$S_TotalDevices</div></div>
  <div class="card"><div class="label">Windows 10</div><div class="value" style="color:#3498db;">$S_TotalWin10</div><div class="sub">$S_OutdatedWin10 outdated</div></div>
  <div class="card"><div class="label">Windows 11</div><div class="value" style="color:#27ae60;">$S_TotalWin11</div><div class="sub">$S_OutdatedWin11 outdated</div></div>
  <div class="card"><div class="label">Other Windows</div><div class="value" style="color:#9b59b6;">$S_TotalOther</div></div>
  <div class="card"><div class="label">Total Outdated</div><div class="value" style="color:#e74c3c;">$S_TotalOutdated</div><div class="sub">$S_PctOutdated% of total</div></div>
</div>

<!-- BREAKDOWN TABLE (matches the OsGeneration / WinBuild / Total view) -->
<div class="breakdown-table">
  <h2>OsGeneration / WinBuild / Total</h2>
  <table>
    <thead><tr><th>OsGeneration</th><th>WinBuild</th><th>Total</th></tr></thead>
    <tbody>
$S_BreakdownRows
    </tbody>
  </table>
</div>

<!-- VERSION SPREAD CARDS -->
<div class="dist-section">
  <div class="section-title">Windows 10 Build Spread (Min Supported: $S_Win10ThresholdDisp)</div>
  <div class="dist-cards">
$S_Win10CardsHtml
  </div>
</div>

<div class="dist-section">
  <div class="section-title">Windows 11 Build Spread (Min Supported: $S_Win11ThresholdDisp)</div>
  <div class="dist-cards">
$S_Win11CardsHtml
  </div>
</div>

<div class="dist-section">
  <div class="section-title">Other Windows Build Spread</div>
  <div class="dist-cards">
$S_OtherCardsHtml
  </div>
</div>

<!-- CHARTS -->
<div class="charts-row">
  <div class="chart-section"><h2>Windows 10 by Build</h2><div class="chart-container"><canvas id="win10Chart"></canvas></div></div>
  <div class="chart-section"><h2>Windows 11 by Build</h2><div class="chart-container"><canvas id="win11Chart"></canvas></div></div>
  <div class="chart-section"><h2>Other Windows by Build</h2><div class="chart-container"><canvas id="otherChart"></canvas></div></div>
</div>

<!-- DEVICE TABLE -->
<div class="table-section">
  <h2>Device Details</h2>
  <div class="table-controls">
    <input type="text" id="searchBox" placeholder="Search by name, user, model, build..." onkeyup="filterTable()" />
    <select id="generationFilter" onchange="filterTable()">
      <option value="all">All Generations</option>
      <option value="Windows 10">Windows 10</option>
      <option value="Windows 11">Windows 11</option>
      <option value="Other Windows">Other Windows</option>
    </select>
    <select id="statusFilter" onchange="filterTable()">
      <option value="all">All Versions</option>
      <option value="Outdated">Outdated Only</option>
      <option value="Supported">Supported Only</option>
      <option value="NoThreshold">No Threshold</option>
      <option value="Unknown">Unknown</option>
    </select>
    <span class="count-label" id="rowCount"></span>
  </div>
  <table id="deviceTable">
    <thead><tr>
      <th onclick="sortTable(0)">Device Name</th>
      <th onclick="sortTable(1)">User</th>
      <th onclick="sortTable(2)">OS Generation</th>
      <th onclick="sortTable(3)">OS Version</th>
      <th onclick="sortTable(4)">Build</th>
      <th onclick="sortTable(5)">Support Status</th>
      <th onclick="sortTable(6)">Manufacturer</th>
      <th onclick="sortTable(7)">Model</th>
      <th onclick="sortTable(8)">Ownership</th>
      <th onclick="sortTable(9)">Join Type</th>
      <th onclick="sortTable(10)">Compliance</th>
      <th onclick="sortTable(11)">Enrolled</th>
      <th onclick="sortTable(12)">Last Sync</th>
      <th onclick="sortTable(13)">Days Since Sync</th>
    </tr></thead>
    <tbody>
$S_TableRows
    </tbody>
  </table>
</div>

<div class="footer">Report generated by ReportIntuneWindowsDevices.ps1</div>

<script>
var win10Data = $S_Win10ChartJson;
var win11Data = $S_Win11ChartJson;
var otherData = $S_OtherChartJson;

function buildChart(canvasId, payload, baseColor) {
  if (!payload.labels.length) { return; }
  var bg = payload.outdated.map(function(o){ return o ? '#e74c3c' : baseColor; });
  new Chart(document.getElementById(canvasId), {
    type: 'bar',
    data: { labels: payload.labels, datasets: [{ label: 'Devices', data: payload.data, backgroundColor: bg, borderWidth: 0 }] },
    options: {
      responsive: true,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: function(ctx) {
        var t = ctx.dataset.data.reduce(function(a,b){return a+b;},0);
        var pct = t > 0 ? ((ctx.parsed.y / t) * 100).toFixed(1) : 0;
        var flag = payload.outdated[ctx.dataIndex] ? ' (Outdated)' : '';
        return ctx.parsed.y + ' devices (' + pct + '%)' + flag;
      } } } },
      scales: { y: { beginAtZero: true, ticks: { precision: 0 } } }
    }
  });
}

buildChart('win10Chart', win10Data, '#3498db');
buildChart('win11Chart', win11Data, '#27ae60');
buildChart('otherChart', otherData, '#9b59b6');

function filterTable() {
  var search = document.getElementById('searchBox').value.toLowerCase();
  var generation = document.getElementById('generationFilter').value;
  var status = document.getElementById('statusFilter').value;
  var rows = document.querySelectorAll('#deviceTable tbody tr');
  var visible = 0;
  rows.forEach(function(row) {
    var text = row.textContent.toLowerCase();
    var matchSearch = !search || text.indexOf(search) !== -1;
    var matchGen = generation === 'all' || row.getAttribute('data-generation') === generation;
    var matchStatus = status === 'all' || row.getAttribute('data-status') === status;
    if (matchSearch && matchGen && matchStatus) {
      row.classList.remove('hidden-row');
      visible++;
    } else {
      row.classList.add('hidden-row');
    }
  });
  document.getElementById('rowCount').textContent = 'Showing ' + visible + ' of ' + rows.length + ' devices';
}

var sortDir = {};
function sortTable(col) {
  var tbody = document.getElementById('deviceTable').querySelector('tbody');
  var rows = Array.from(tbody.querySelectorAll('tr'));
  var dir = sortDir[col] === 'asc' ? 'desc' : 'asc';
  sortDir[col] = dir;
  rows.sort(function(a, b) {
    var av = a.cells[col].textContent.trim().toLowerCase();
    var bv = b.cells[col].textContent.trim().toLowerCase();
    var an = parseFloat(av), bn = parseFloat(bv);
    if (!isNaN(an) && !isNaN(bn)) { return dir === 'asc' ? an - bn : bn - an; }
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

    $S_Html | Out-File -FilePath $S_HtmlFile -Encoding UTF8

    # --- Console summary ---
    Write-Host ""
    Write-Host "Intune Windows Devices Report" -ForegroundColor Cyan
    Write-Host "--------------------------------------------"
    Write-Host ("Tenant                       : {0} ({1})" -f $S_TenantDisplayName, $S_TenantId)
    Write-Host ("Min supported Windows 10     : {0}" -f $S_Win10ThresholdDisp)
    Write-Host ("Min supported Windows 11     : {0}" -f $S_Win11ThresholdDisp)
    Write-Host ("Total Windows devices        : {0}" -f $S_TotalDevices)
    Write-Host ("  Windows 10                 : {0}  (Outdated: {1})" -f $S_TotalWin10, $S_OutdatedWin10) -ForegroundColor Cyan
    Write-Host ("  Windows 11                 : {0}  (Outdated: {1})" -f $S_TotalWin11, $S_OutdatedWin11) -ForegroundColor Green
    Write-Host ("  Other Windows              : {0}" -f $S_TotalOther) -ForegroundColor DarkGray
    Write-Host ("Total outdated               : {0}  ({1}%)" -f $S_TotalOutdated, $S_PctOutdated) -ForegroundColor Red
    Write-Host ""
    Write-Host "OsGeneration / WinBuild / Total" -ForegroundColor Cyan
    $S_Breakdown | Format-Table OsGeneration, WinBuild, Total, Outdated -AutoSize | Out-String | Write-Host
    Write-Host ("CSV report (devices)         : {0}" -f $S_CsvFile) -ForegroundColor Yellow
    Write-Host ("CSV report (breakdown)       : {0}" -f $S_BreakdownCsv) -ForegroundColor Yellow
    Write-Host ("HTML report                  : {0}" -f $S_HtmlFile) -ForegroundColor Yellow

    $S_DisconnectChoice = Read-Host "Disconnect from Microsoft Graph? (Y/N)"
    if ($S_DisconnectChoice -match '^(y|yes)$')
    {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
}
catch
{
    Write-Error $_
    exit 1
}
