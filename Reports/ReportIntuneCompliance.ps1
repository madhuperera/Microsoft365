#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
	Reports on Intune device compliance across all operating systems, including
	tenant-level compliance settings, overall compliance state, per-policy
	breakdown and an "active devices only" view.

.DESCRIPTION
	Connects to Microsoft Graph and produces an HTML + CSV report covering:
		1. Tenant-level compliance settings (secure-by-default + check-in
		   threshold). The HTML renders these as a call-to-action banner when
		   misconfigured (secureByDefault = false) and as a green status banner
		   when configured correctly.
		2. Tenant-wide compliance state pie chart from the Intune
		   deviceCompliancePolicyDeviceStateSummary endpoint.
		3. A second pie chart limited to "active" devices only (devices that
		   have checked in within the tenant check-in threshold). Devices that
		   haven't checked in within that window are excluded.
		4. Per-policy compliance breakdown table (Compliant / Non-compliant /
		   Error / Conflict / Not applicable / Unknown) tagged with the OS the
		   policy targets.
		5. Per-device per-policy CSV drill-down for offline analysis.

.PARAMETER ReportPath
	Folder for the output reports. If omitted the current working directory
	is used.

.PARAMETER IncludeDeviceStatusDetails
	If specified, queries deviceStatuses for every compliance policy and
	exports a per-device per-policy CSV. This makes one Graph call per policy
	and can be slow on large tenants.

.EXAMPLE
	.\ReportIntuneCompliance.ps1

.EXAMPLE
	.\ReportIntuneCompliance.ps1 -IncludeDeviceStatusDetails -ReportPath C:\Reports
#>

[CmdletBinding()]
param(
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[string]$ReportPath,

	[Parameter(Mandatory = $false)]
	[switch]$IncludeDeviceStatusDetails
)

$ErrorActionPreference = "Stop"

$S_RequiredGraphScopes = @(
	'DeviceManagementConfiguration.Read.All'
	'DeviceManagementManagedDevices.Read.All'
	'DeviceManagementServiceConfig.Read.All'
	'Organization.Read.All'
)

try {
	if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
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
	if ([string]::IsNullOrWhiteSpace($S_ContextConfirmation)) { $S_ContextConfirmation = 'N' }
	if ($S_ContextConfirmation.ToUpperInvariant() -ne 'Y') {
		throw "Operation cancelled. Please reconnect to the correct tenant and account, then run again."
	}

	# --- Tenant info ---
	$tenantDisplayName = $null
	try {
		$orgResp = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization' -ErrorAction Stop
		if ($orgResp.value) { $tenantDisplayName = $orgResp.value[0].displayName }
	} catch { }
	if (-not $tenantDisplayName) { $tenantDisplayName = $S_ExistingContext.TenantId }
	$tenantId = if ($S_ExistingContext.TenantId) { $S_ExistingContext.TenantId } else { 'Unknown' }

	# --- Tenant compliance settings ---
	Write-Host "Reading tenant compliance settings..." -ForegroundColor Cyan
	$secureByDefault = $null
	$checkinThresholdDays = $null
	try {
		$settings = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/beta/deviceManagement/settings' -ErrorAction Stop
		$secureByDefault = $settings.secureByDefault
		$checkinThresholdDays = $settings.deviceComplianceCheckinThresholdDays
	} catch {
		Write-Warning "Failed to read tenant compliance settings: $($_.Exception.Message)"
	}
	if ($null -eq $checkinThresholdDays -or $checkinThresholdDays -le 0) { $checkinThresholdDays = 30 }
	Write-Host ("  secureByDefault                      : {0}" -f $secureByDefault) -ForegroundColor Green
	Write-Host ("  deviceComplianceCheckinThresholdDays : {0}" -f $checkinThresholdDays) -ForegroundColor Green

	# --- Tenant-wide compliance state summary ---
	Write-Host "Fetching tenant-wide compliance summary..." -ForegroundColor Cyan
	$tenantSummary = $null
	try {
		$tenantSummary = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicyDeviceStateSummary' -ErrorAction Stop
	} catch {
		Write-Warning "Failed to read deviceCompliancePolicyDeviceStateSummary: $($_.Exception.Message)"
	}

	$tenantStates = [ordered]@{
		Compliant      = if ($tenantSummary) { [int]$tenantSummary.compliantDeviceCount }     else { 0 }
		NonCompliant   = if ($tenantSummary) { [int]$tenantSummary.nonCompliantDeviceCount }  else { 0 }
		InGracePeriod  = if ($tenantSummary) { [int]$tenantSummary.inGracePeriodCount }       else { 0 }
		Error          = if ($tenantSummary) { [int]$tenantSummary.errorDeviceCount }         else { 0 }
		Conflict       = if ($tenantSummary) { [int]$tenantSummary.conflictDeviceCount }      else { 0 }
		NotApplicable  = if ($tenantSummary) { [int]$tenantSummary.notApplicableDeviceCount } else { 0 }
		Remediated     = if ($tenantSummary) { [int]$tenantSummary.remediatedDeviceCount }    else { 0 }
		ConfigManager  = if ($tenantSummary) { [int]$tenantSummary.configManagerCount }       else { 0 }
		Unknown        = if ($tenantSummary) { [int]$tenantSummary.unknownDeviceCount }       else { 0 }
	}
	$tenantTotal = ($tenantStates.Values | Measure-Object -Sum).Sum

	# --- Managed devices (used for active-devices view) ---
	Write-Host "Fetching managed devices..." -ForegroundColor Cyan
	$select = 'id,deviceName,userPrincipalName,userDisplayName,operatingSystem,osVersion,complianceState,managedDeviceOwnerType,lastSyncDateTime,enrolledDateTime'
	$uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$select=$select&`$top=200"
	$devices = New-Object System.Collections.Generic.List[object]
	do {
		$resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
		if ($resp.value) {
			foreach ($d in $resp.value) { $devices.Add([pscustomobject]$d) | Out-Null }
		}
		$uri = $resp.'@odata.nextLink'
	} while ($uri)
	Write-Host ("  Retrieved {0} managed devices" -f $devices.Count) -ForegroundColor Green

	# --- Active vs stale split based on check-in threshold ---
	$now = Get-Date
	$activeCutoff = $now.AddDays(-[int]$checkinThresholdDays)
	$activeStates = [ordered]@{
		compliant    = 0
		noncompliant = 0
		ingraceperiod = 0
		error        = 0
		conflict     = 0
		notapplicable = 0
		configmanager = 0
		unknown      = 0
	}
	$staleCount = 0
	foreach ($d in $devices) {
		$last = if ($d.lastSyncDateTime) { [datetime]$d.lastSyncDateTime } else { $null }
		if (-not $last -or $last -lt $activeCutoff) { $staleCount++; continue }
		$key = if ($d.complianceState) { ([string]$d.complianceState).ToLowerInvariant() } else { 'unknown' }
		if (-not $activeStates.Contains($key)) { $activeStates[$key] = 0 }
		$activeStates[$key]++
	}
	$activeTotal = ($activeStates.Values | Measure-Object -Sum).Sum

	# --- Compliance policies + status overview ---
	Write-Host "Fetching compliance policies..." -ForegroundColor Cyan
	$policies = New-Object System.Collections.Generic.List[object]
	$pUri = 'https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies?$expand=assignments'
	do {
		$resp = Invoke-MgGraphRequest -Method GET -Uri $pUri -ErrorAction Stop
		if ($resp.value) {
			foreach ($p in $resp.value) { $policies.Add([pscustomobject]$p) | Out-Null }
		}
		$pUri = $resp.'@odata.nextLink'
	} while ($pUri)
	Write-Host ("  Retrieved {0} compliance policies" -f $policies.Count) -ForegroundColor Green

	$osTypeMap = @{
		'#microsoft.graph.windows10CompliancePolicy'              = 'Windows 10/11'
		'#microsoft.graph.windowsPhone81CompliancePolicy'         = 'Windows Phone'
		'#microsoft.graph.windows81CompliancePolicy'              = 'Windows 8.1'
		'#microsoft.graph.iosCompliancePolicy'                    = 'iOS / iPadOS'
		'#microsoft.graph.macOSCompliancePolicy'                  = 'macOS'
		'#microsoft.graph.androidCompliancePolicy'                = 'Android (Device Admin)'
		'#microsoft.graph.androidWorkProfileCompliancePolicy'     = 'Android Work Profile'
		'#microsoft.graph.androidDeviceOwnerCompliancePolicy'     = 'Android (Device Owner)'
		'#microsoft.graph.aospDeviceOwnerCompliancePolicy'        = 'AOSP (Device Owner)'
		'#microsoft.graph.androidForWorkCompliancePolicy'         = 'Android for Work'
		'#microsoft.graph.linuxCompliancePolicy'                  = 'Linux'
	}

	$policyReport = New-Object System.Collections.Generic.List[object]
	$deviceStatusRows = New-Object System.Collections.Generic.List[object]

	# Iterate every policy and pull per-device compliance from the v2 Intune
	# Reports endpoint (getCompliancePolicyDevicesReport). This is the same
	# data source the Intune portal uses for the Compliant / Noncompliant /
	# Others tile, so totals match the UI exactly. The endpoint returns a
	# columnar { Schema:[{Column,PropertyType}], Values:[[...]] } payload.
	$policyIndex = 0
	foreach ($pol in $policies) {
		$policyIndex++
		Write-Host ("  [{0}/{1}] {2}" -f $policyIndex, $policies.Count, $pol.displayName) -ForegroundColor DarkGray

		$odata = $pol.'@odata.type'
		$os = if ($odata -and $osTypeMap.ContainsKey($odata)) { $osTypeMap[$odata] } else { ($odata -replace '#microsoft\.graph\.', '') }

		$statusCounts = [ordered]@{
			Compliant      = 0
			Noncompliant   = 0
			InGracePeriod  = 0
			Conflict       = 0
			Error          = 0
			NotApplicable  = 0
			ConfigManager  = 0
			NotEvaluated   = 0
			RemediatedNoncompliance = 0
			Unknown        = 0
		}

		$top = 1000
		$skip = 0
		$keepFetching = $true
		while ($keepFetching) {
			$body = @{
				filter = "(PolicyId eq '$($pol.id)')"
				skip   = $skip
				top    = $top
				select = @('DeviceId','DeviceName','UPN','UserEmail','UserName','OS','OSDescription','OSVersion','OwnerType','LastContact','ComplianceState','PolicyId','PolicyName','PolicyPlatformType','ReportStatus','DeviceModel','DeviceType','IMEI')
			} | ConvertTo-Json -Depth 5 -Compress

			$resp = $null
			try {
				$resp = Invoke-MgGraphRequest -Method POST `
					-Uri 'https://graph.microsoft.com/beta/deviceManagement/reports/getCompliancePolicyDevicesReport' `
					-ContentType 'application/json' `
					-Body $body `
					-ErrorAction Stop
			} catch {
				Write-Warning ("getCompliancePolicyDevicesReport failed for {0}: {1}" -f $pol.displayName, $_.Exception.Message)
				break
			}

			# Response can come back as a hashtable or a JSON byte stream depending on module version.
			if ($resp -is [byte[]]) {
				$resp = [System.Text.Encoding]::UTF8.GetString($resp) | ConvertFrom-Json
			}

			$schema = $resp.Schema
			$values = $resp.Values
			if (-not $schema -or -not $values -or $values.Count -eq 0) { break }

			# Build column-name -> index map for this page
			$colIdx = @{}
			for ($i = 0; $i -lt $schema.Count; $i++) {
				$cname = if ($schema[$i].Column) { [string]$schema[$i].Column } elseif ($schema[$i].PropertyName) { [string]$schema[$i].PropertyName } else { $null }
				if ($cname) { $colIdx[$cname] = $i }
			}
			$ixState  = $colIdx['ComplianceState']
			$ixDevice = $colIdx['DeviceName']
			$ixUpn    = $colIdx['UPN']
			$ixOs     = $colIdx['OS']
			$ixOsVer  = $colIdx['OSVersion']
			$ixOwner  = $colIdx['OwnerType']
			$ixLast   = $colIdx['LastContact']
			$ixDevId  = $colIdx['DeviceId']
			$ixModel  = $colIdx['DeviceModel']

			foreach ($row in $values) {
				$rawState = if ($null -ne $ixState) { [string]$row[$ixState] } else { 'Unknown' }
				if ([string]::IsNullOrWhiteSpace($rawState)) { $rawState = 'Unknown' }

				# Normalize to a canonical key used in $statusCounts
				$key = switch -Regex ($rawState) {
					'^(?i)compliant$'                { 'Compliant'; break }
					'^(?i)non[- ]?compliant$'        { 'Noncompliant'; break }
					'^(?i)inGracePeriod$'            { 'InGracePeriod'; break }
					'^(?i)in[- ]?grace[- ]?period$'  { 'InGracePeriod'; break }
					'^(?i)conflict$'                 { 'Conflict'; break }
					'^(?i)error$'                    { 'Error'; break }
					'^(?i)not[- ]?applicable$'       { 'NotApplicable'; break }
					'^(?i)configManager$'            { 'ConfigManager'; break }
					'^(?i)not[- ]?evaluated$'        { 'NotEvaluated'; break }
					'^(?i)remediated.*'              { 'RemediatedNoncompliance'; break }
					default                           { 'Unknown' }
				}
				if (-not $statusCounts.Contains($key)) { $statusCounts[$key] = 0 }
				$statusCounts[$key]++

				if ($IncludeDeviceStatusDetails) {
					$deviceStatusRows.Add([pscustomobject]@{
						PolicyName        = $pol.displayName
						PolicyOS          = $os
						DeviceId          = if ($null -ne $ixDevId)  { $row[$ixDevId] }  else { '' }
						DeviceDisplayName = if ($null -ne $ixDevice) { $row[$ixDevice] } else { '' }
						UserPrincipalName = if ($null -ne $ixUpn)    { $row[$ixUpn] }    else { '' }
						OS                = if ($null -ne $ixOs)     { $row[$ixOs] }    else { '' }
						OSVersion         = if ($null -ne $ixOsVer)  { $row[$ixOsVer] } else { '' }
						OwnerType         = if ($null -ne $ixOwner)  { $row[$ixOwner] } else { '' }
						LastContact       = if ($null -ne $ixLast)   { $row[$ixLast] }  else { '' }
						DeviceModel       = if ($null -ne $ixModel)  { $row[$ixModel] } else { '' }
						ComplianceState   = $rawState
						MappedBucket      = $key
					}) | Out-Null
				}
			}

			if ($values.Count -lt $top) { $keepFetching = $false } else { $skip += $top }
		}

		# Map raw statuses to Intune portal tiles:
		#   Compliant  = Compliant + InGracePeriod (portal counts grace as compliant in the bar)
		#   Noncompliant = Noncompliant + RemediatedNoncompliance
		#   Others     = Conflict + Error + NotApplicable + ConfigManager + NotEvaluated + Unknown
		$compliant    = [int]$statusCounts['Compliant'] + [int]$statusCounts['InGracePeriod']
		$nonCompliant = [int]$statusCounts['Noncompliant'] + [int]$statusCounts['RemediatedNoncompliance']
		$others       = [int]$statusCounts['Conflict'] + [int]$statusCounts['Error'] + [int]$statusCounts['NotApplicable'] + [int]$statusCounts['ConfigManager'] + [int]$statusCounts['NotEvaluated'] + [int]$statusCounts['Unknown']
		$total = $compliant + $nonCompliant + $others
		$pctCompliant = if (($compliant + $nonCompliant) -gt 0) { [math]::Round(($compliant / ($compliant + $nonCompliant)) * 100, 1) } else { $null }

		$assignedGroupCount = if ($pol.assignments) { @($pol.assignments).Count } else { 0 }

		$policyReport.Add([pscustomobject]@{
			DisplayName          = $pol.displayName
			OperatingSystem      = $os
			OdataType            = $odata
			AssignmentCount      = $assignedGroupCount
			Compliant            = $compliant
			NonCompliant         = $nonCompliant
			Others               = $others
			RawCompliant         = [int]$statusCounts['Compliant']
			InGracePeriod        = [int]$statusCounts['InGracePeriod']
			Remediated           = [int]$statusCounts['RemediatedNoncompliance']
			InError              = [int]$statusCounts['Error']
			Conflict             = [int]$statusCounts['Conflict']
			NotApplicable        = [int]$statusCounts['NotApplicable']
			NotEvaluated         = [int]$statusCounts['NotEvaluated']
			ConfigManager        = [int]$statusCounts['ConfigManager']
			Unknown              = [int]$statusCounts['Unknown']
			TotalReporting       = $total
			PercentCompliant     = $pctCompliant
			LastModifiedDateTime = $pol.lastModifiedDateTime
			Id                   = $pol.id
		}) | Out-Null
	}

	# --- Output paths ---
	if (-not $ReportPath) { $ReportPath = (Get-Location).Path }
	$reportFolder = if (Test-Path $ReportPath -PathType Container) { $ReportPath } else { Split-Path -Parent $ReportPath }
	if ($reportFolder -and -not (Test-Path $reportFolder)) {
		New-Item -ItemType Directory -Path $reportFolder -Force | Out-Null
	}
	$S_Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
	$policyCsv      = Join-Path $reportFolder ("ReportIntuneCompliance_Policies_{0}.csv"      -f $S_Timestamp)
	$deviceStatusCsv = Join-Path $reportFolder ("ReportIntuneCompliance_DeviceStatuses_{0}.csv" -f $S_Timestamp)
	$htmlFile       = Join-Path $reportFolder ("ReportIntuneCompliance_{0}.html"             -f $S_Timestamp)

	# --- CSV exports ---
	$policyReport | Sort-Object OperatingSystem, DisplayName | Export-Csv -Path $policyCsv -NoTypeInformation -Encoding UTF8
	if ($IncludeDeviceStatusDetails -and $deviceStatusRows.Count -gt 0) {
		$deviceStatusRows | Sort-Object PolicyName, DeviceDisplayName | Export-Csv -Path $deviceStatusCsv -NoTypeInformation -Encoding UTF8
	}

	# --- HTML helpers ---
	$enc = { param($s) if ($null -eq $s -or $s -eq '') { '-' } else { [System.Net.WebUtility]::HtmlEncode([string]$s) } }
	$reportDate = Get-Date -Format "dd MMM yyyy HH:mm"

	# Banner state
	$secureByDefaultDisp = if ($null -eq $secureByDefault) { 'Unknown' } elseif ($secureByDefault) { 'On' } else { 'Off' }
	$bannerClass = if ($secureByDefault -eq $true) { 'banner banner-good' } else { 'banner banner-bad' }
	$bannerTitle = if ($secureByDefault -eq $true) {
		'Tenant compliance settings look good'
	} elseif ($secureByDefault -eq $false) {
		'Action required: Devices without a compliance policy are being marked Compliant'
	} else {
		'Tenant compliance settings could not be read'
	}
	$bannerBody = if ($secureByDefault -eq $true) {
		"Mark devices with no compliance policy assigned as Not compliant is currently <strong>On</strong>. Devices that have not checked in for <strong>$checkinThresholdDays</strong> days will be marked as Not compliant."
	} elseif ($secureByDefault -eq $false) {
		"Mark devices with no compliance policy assigned as is currently <strong>Compliant</strong> (insecure default). Change this to <strong>Not compliant</strong> in Intune > Endpoint security > Device compliance > Compliance policy settings. Current check-in threshold is <strong>$checkinThresholdDays</strong> days."
	} else {
		"Could not read deviceManagement/settings. Verify the signed-in account has the required Graph scopes."
	}

	function ConvertTo-PieJson {
		param([System.Collections.IDictionary]$Map)
		$entries = @()
		foreach ($k in $Map.Keys) {
			if ([int]$Map[$k] -gt 0) { $entries += [pscustomobject]@{ Label = $k; Value = [int]$Map[$k] } }
		}
		if (-not $entries -or $entries.Count -eq 0) { return '{"labels":[],"data":[]}' }
		$labels = ($entries | ForEach-Object { '"' + $_.Label + '"' }) -join ','
		$data   = ($entries | ForEach-Object { $_.Value }) -join ','
		"{`"labels`":[$labels],`"data`":[$data]}"
	}

	# Tenant pie data
	$tenantPieJson = ConvertTo-PieJson -Map $tenantStates
	$activePieJson = ConvertTo-PieJson -Map $activeStates

	# Policy table rows
	$policyRows = ($policyReport | Sort-Object OperatingSystem, DisplayName | ForEach-Object {
		$pct = if ($null -ne $_.PercentCompliant) { ('{0}%' -f $_.PercentCompliant) } else { '-' }
		$lm  = if ($_.LastModifiedDateTime) { ([datetime]$_.LastModifiedDateTime).ToString('dd MMM yyyy') } else { '-' }
		$ncClass = if ($_.NonCompliant -gt 0) { 'cell-bad' } else { '' }
		$cClass  = if ($_.Compliant -gt 0) { 'cell-good' } else { '' }
		$othersTitle = ("Error: {0}, Conflict: {1}, Not Applicable: {2}, Not Evaluated: {3}, ConfigManager: {4}, Unknown: {5}" -f $_.InError, $_.Conflict, $_.NotApplicable, $_.NotEvaluated, $_.ConfigManager, $_.Unknown)
		$rowAttr = "data-os=`"$(& $enc $_.OperatingSystem)`""
		"<tr $rowAttr>" +
			"<td>$(& $enc $_.DisplayName)</td>" +
			"<td>$(& $enc $_.OperatingSystem)</td>" +
			"<td>$($_.AssignmentCount)</td>" +
			"<td class='$cClass' title='Compliant: $($_.RawCompliant), In Grace Period: $($_.InGracePeriod)'>$($_.Compliant)</td>" +
			"<td class='$ncClass' title='Noncompliant: $([int]$_.NonCompliant - [int]$_.Remediated), Remediated: $($_.Remediated)'>$($_.NonCompliant)</td>" +
			"<td title='$othersTitle'>$($_.Others)</td>" +
			"<td>$($_.TotalReporting)</td>" +
			"<td>$pct</td>" +
			"<td>$lm</td>" +
		"</tr>"
	}) -join "`n"

	# OS filter options
	$osOptions = ($policyReport | Select-Object -ExpandProperty OperatingSystem -Unique | Sort-Object | ForEach-Object {
		"<option value=`"$(& $enc $_)`">$(& $enc $_)</option>"
	}) -join "`n"

	$drilldownNote = if ($IncludeDeviceStatusDetails -and $deviceStatusRows.Count -gt 0) {
		"<p style='font-size:0.85em;color:#555;margin-top:8px;'>Per-device drill-down exported to <code>$(& $enc (Split-Path $deviceStatusCsv -Leaf))</code> ($($deviceStatusRows.Count) rows)</p>"
	} elseif ($IncludeDeviceStatusDetails) {
		"<p style='font-size:0.85em;color:#555;margin-top:8px;'>No per-device status rows returned.</p>"
	} else {
		"<p style='font-size:0.85em;color:#777;margin-top:8px;'>Re-run with <code>-IncludeDeviceStatusDetails</code> to also produce a per-device-per-policy CSV drill-down.</p>"
	}

	$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Intune Compliance Report</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; padding: 30px; }
  .header { background: linear-gradient(135deg, #1a1a2e, #16213e); color: #fff; padding: 30px 40px; border-radius: 12px; margin-bottom: 24px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 16px; }
  .header h1 { font-size: 1.6em; margin-bottom: 6px; }
  .header p { font-size: 0.9em; opacity: 0.85; }

  .banner { padding: 22px 28px; border-radius: 12px; margin-bottom: 28px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); border-left: 8px solid; display: flex; gap: 18px; align-items: flex-start; }
  .banner .icon { font-size: 1.8em; line-height: 1; flex-shrink: 0; }
  .banner-good { background: #eaf7ec; border-left-color: #27ae60; color: #155724; }
  .banner-good .icon::before { content: '\2714'; color: #27ae60; }
  .banner-bad { background: #fdecea; border-left-color: #e74c3c; color: #721c24; }
  .banner-bad .icon::before { content: '\26A0'; color: #e74c3c; }
  .banner h2 { font-size: 1.15em; margin-bottom: 6px; }
  .banner p { font-size: 0.92em; line-height: 1.5; }
  .banner code, .banner strong { font-weight: 700; }

  .section-title { font-size: 1.15em; font-weight: 600; color: #1a1a2e; margin: 0 0 16px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0; }

  .summary-cards { display: flex; gap: 20px; margin-bottom: 28px; flex-wrap: wrap; }
  .card { background: #fff; border-radius: 10px; padding: 22px 26px; flex: 1; min-width: 160px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
  .card .label { font-size: 0.82em; color: #777; text-transform: uppercase; letter-spacing: 0.4px; }
  .card .value { font-size: 1.9em; font-weight: 700; margin-top: 6px; }
  .card .sub { font-size: 0.78em; color: #999; margin-top: 2px; }

  .charts-row { display: flex; gap: 20px; margin-bottom: 28px; flex-wrap: wrap; }
  .chart-section { background: #fff; border-radius: 10px; padding: 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); flex: 1; min-width: 360px; }
  .chart-section h2 { font-size: 1.05em; margin-bottom: 4px; color: #1a1a2e; }
  .chart-section .subtitle { font-size: 0.85em; color: #777; margin-bottom: 16px; }
  .chart-container { max-width: 380px; margin: 0 auto; }

  .table-section { background: #fff; border-radius: 10px; padding: 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 28px; overflow-x: auto; }
  .table-section h2 { font-size: 1.1em; margin-bottom: 16px; color: #1a1a2e; }
  .table-controls { display: flex; gap: 14px; margin-bottom: 16px; flex-wrap: wrap; align-items: center; }
  .table-controls input[type="text"] { padding: 8px 14px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; min-width: 260px; }
  .table-controls select { padding: 8px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; background: #fff; }
  .table-controls .count-label { font-size: 0.85em; color: #777; margin-left: auto; }
  table { width: 100%; border-collapse: collapse; font-size: 0.86em; }
  th { background: #1a1a2e; color: #fff; padding: 10px 12px; text-align: left; cursor: pointer; user-select: none; white-space: nowrap; position: sticky; top: 0; }
  th:hover { background: #2c3e50; }
  td { padding: 9px 12px; border-bottom: 1px solid #eee; white-space: nowrap; }
  tr:hover td { background: #f8f9fa; }
  tr.hidden-row { display: none; }
  td.cell-bad { background: #fdecea; color: #721c24; font-weight: 600; }
  td.cell-good { background: #eaf7ec; color: #155724; font-weight: 600; }

  .footer { text-align: center; font-size: 0.8em; color: #999; margin-top: 20px; }
</style>
</head>
<body>

<div class="header">
  <div>
    <h1>Intune Compliance Report</h1>
    <p>Tenant: $(& $enc $tenantDisplayName) ($tenantId) &nbsp;|&nbsp; Generated: $reportDate</p>
  </div>
  <div style="text-align:right;font-size:0.9em;opacity:0.9;">
    Secure by default: <strong>$secureByDefaultDisp</strong><br/>
    Check-in threshold: <strong>$checkinThresholdDays days</strong>
  </div>
</div>

<!-- TENANT COMPLIANCE SETTINGS BANNER -->
<div class="$bannerClass">
  <div class="icon"></div>
  <div>
    <h2>$bannerTitle</h2>
    <p>$bannerBody</p>
  </div>
</div>

<!-- OVERVIEW CARDS -->
<div class="section-title">Overview</div>
<div class="summary-cards">
  <div class="card"><div class="label">Total Reporting Devices</div><div class="value" style="color:#1a1a2e;">$tenantTotal</div></div>
  <div class="card"><div class="label">Compliant</div><div class="value" style="color:#27ae60;">$($tenantStates.Compliant)</div></div>
  <div class="card"><div class="label">Non-compliant</div><div class="value" style="color:#e74c3c;">$($tenantStates.NonCompliant)</div></div>
  <div class="card"><div class="label">In Grace Period</div><div class="value" style="color:#f39c12;">$($tenantStates.InGracePeriod)</div></div>
  <div class="card"><div class="label">Active Devices (last $checkinThresholdDays days)</div><div class="value" style="color:#3498db;">$activeTotal</div><div class="sub">$staleCount stale / not checked in</div></div>
  <div class="card"><div class="label">Total Policies</div><div class="value" style="color:#9b59b6;">$($policyReport.Count)</div></div>
</div>

<!-- CHARTS -->
<div class="charts-row">
  <div class="chart-section">
    <h2>Tenant Compliance State</h2>
    <div class="subtitle">All reporting devices (deviceCompliancePolicyDeviceStateSummary)</div>
    <div class="chart-container"><canvas id="tenantPie"></canvas></div>
  </div>
  <div class="chart-section">
    <h2>Active Devices Compliance State</h2>
    <div class="subtitle">Only devices that checked in within the last $checkinThresholdDays days ($staleCount stale devices excluded)</div>
    <div class="chart-container"><canvas id="activePie"></canvas></div>
  </div>
</div>

<!-- POLICY TABLE -->
<div class="table-section">
  <h2>Compliance Policies</h2>
  <div class="table-controls">
    <input type="text" id="searchBox" placeholder="Search policy name or OS..." onkeyup="filterTable()" />
    <select id="osFilter" onchange="filterTable()">
      <option value="all">All Operating Systems</option>
$osOptions
    </select>
    <span class="count-label" id="rowCount"></span>
  </div>
  <table id="policyTable">
    <thead><tr>
      <th onclick="sortTable(0)">Policy Name</th>
      <th onclick="sortTable(1)">OS</th>
      <th onclick="sortTable(2)">Assignments</th>
      <th onclick="sortTable(3)">Compliant</th>
      <th onclick="sortTable(4)">Non-compliant</th>
      <th onclick="sortTable(5)">Others</th>
      <th onclick="sortTable(6)">Total</th>
      <th onclick="sortTable(7)">% Compliant</th>
      <th onclick="sortTable(8)">Last Modified</th>
    </tr></thead>
    <tbody>
$policyRows
    </tbody>
  </table>
  $drilldownNote
</div>

<div class="footer">Report generated by ReportIntuneCompliance.ps1</div>

<script>
var tenantPieData = $tenantPieJson;
var activePieData = $activePieJson;

var stateColors = {
  'Compliant':     '#27ae60',
  'compliant':     '#27ae60',
  'NonCompliant':  '#e74c3c',
  'noncompliant':  '#e74c3c',
  'InGracePeriod': '#f39c12',
  'ingraceperiod': '#f39c12',
  'Error':         '#c0392b',
  'error':         '#c0392b',
  'Conflict':      '#d35400',
  'conflict':      '#d35400',
  'NotApplicable': '#95a5a6',
  'notapplicable': '#95a5a6',
  'Remediated':    '#1abc9c',
  'remediated':    '#1abc9c',
  'ConfigManager': '#34495e',
  'configmanager': '#34495e',
  'Unknown':       '#7f8c8d',
  'unknown':       '#7f8c8d'
};

function buildPie(canvasId, payload) {
  if (!payload.labels.length || payload.data.every(function(v){return v===0;})) {
    var ctx = document.getElementById(canvasId).getContext('2d');
    ctx.font = '14px Segoe UI'; ctx.fillStyle = '#999';
    ctx.fillText('No data', 10, 30);
    return;
  }
  var bg = payload.labels.map(function(l){ return stateColors[l] || '#bdc3c7'; });
  new Chart(document.getElementById(canvasId), {
    type: 'doughnut',
    data: { labels: payload.labels, datasets: [{ data: payload.data, backgroundColor: bg, borderWidth: 2, borderColor: '#fff' }] },
    options: {
      responsive: true,
      plugins: {
        legend: { position: 'right', labels: { padding: 12, font: { size: 12 }, boxWidth: 14 } },
        tooltip: { callbacks: { label: function(ctx) {
          var t = ctx.dataset.data.reduce(function(a,b){return a+b;},0);
          var pct = t > 0 ? ((ctx.parsed / t) * 100).toFixed(1) : 0;
          return ctx.label + ': ' + ctx.parsed + ' (' + pct + '%)';
        } } }
      }
    }
  });
}

buildPie('tenantPie', tenantPieData);
buildPie('activePie', activePieData);

function filterTable() {
  var search = document.getElementById('searchBox').value.toLowerCase();
  var os = document.getElementById('osFilter').value;
  var rows = document.querySelectorAll('#policyTable tbody tr');
  var visible = 0;
  rows.forEach(function(row) {
    var text = row.textContent.toLowerCase();
    var matchSearch = !search || text.indexOf(search) !== -1;
    var matchOs = os === 'all' || row.getAttribute('data-os') === os;
    if (matchSearch && matchOs) {
      row.classList.remove('hidden-row');
      visible++;
    } else {
      row.classList.add('hidden-row');
    }
  });
  document.getElementById('rowCount').textContent = 'Showing ' + visible + ' of ' + rows.length + ' policies';
}

var sortDir = {};
function sortTable(col) {
  var tbody = document.getElementById('policyTable').querySelector('tbody');
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

	$html | Out-File -FilePath $htmlFile -Encoding UTF8

	# --- Console summary ---
	Write-Host ""
	Write-Host "Intune Compliance Report" -ForegroundColor Cyan
	Write-Host "--------------------------------------------"
	Write-Host ("Tenant                    : {0} ({1})" -f $tenantDisplayName, $tenantId)
	Write-Host ("secureByDefault           : {0}" -f $secureByDefaultDisp) -ForegroundColor $(if ($secureByDefault -eq $true) { 'Green' } else { 'Red' })
	Write-Host ("Check-in threshold (days) : {0}" -f $checkinThresholdDays)
	Write-Host ("Total reporting devices   : {0}" -f $tenantTotal)
	Write-Host ("  Compliant               : {0}" -f $tenantStates.Compliant) -ForegroundColor Green
	Write-Host ("  Non-compliant           : {0}" -f $tenantStates.NonCompliant) -ForegroundColor Red
	Write-Host ("  In Grace Period         : {0}" -f $tenantStates.InGracePeriod) -ForegroundColor Yellow
	Write-Host ("  Error                   : {0}" -f $tenantStates.Error)
	Write-Host ("  Conflict                : {0}" -f $tenantStates.Conflict)
	Write-Host ("  Not Applicable          : {0}" -f $tenantStates.NotApplicable)
	Write-Host ("Active devices            : {0}  (Stale: {1})" -f $activeTotal, $staleCount)
	Write-Host ("Total compliance policies : {0}" -f $policyReport.Count)
	Write-Host ""
	Write-Host ("CSV (policies)            : {0}" -f $policyCsv) -ForegroundColor Yellow
	if ($IncludeDeviceStatusDetails -and $deviceStatusRows.Count -gt 0) {
		Write-Host ("CSV (device statuses)     : {0}" -f $deviceStatusCsv) -ForegroundColor Yellow
	}
	Write-Host ("HTML report               : {0}" -f $htmlFile) -ForegroundColor Yellow

	$S_DisconnectChoice = Read-Host "Disconnect from Microsoft Graph? (Y/N)"
	if ($S_DisconnectChoice -match '^(y|yes)$') {
		Disconnect-MgGraph -ErrorAction SilentlyContinue
	}
}
catch {
	Write-Error $_
	exit 1
}
