param(
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[string]$ReportPath,

	[Parameter(Mandatory = $false)]
	[ValidateRange(1, 3650)]
	[int]$InactiveDays = 90,

	[Parameter(Mandatory = $false)]
	[string[]]$PrimaryLicenseSkus = @(
		"O365_BUSINESS_ESSENTIALS",
		"SMB_BUSINESS_ESSENTIALS",
		"O365_BUSINESS_PREMIUM",
		"SMB_BUSINESS_PREMIUM",
		"SPB",
		"SPE_E3",
		"SPE_E5",
		"SPE_F1",
		"ENTERPRISEPACK",
		"ENTERPRISEPREMIUM",
		"DESKLESSPACK",
		"M365_F1",
		"MICROSOFT_365_FRONTLINE_F3",
		"SMB_BUSINESS",
		"O365_BUSINESS",
		"EXCHANGESTANDARD",
		"EXCHANGEENTERPRISE",
		"STANDARDPACK",
		"STANDARDWOFFPACK",
		"ENTERPRISEWITHSCAL",
		"DEVELOPERPACK",
		"DEVELOPERPACK_E5"
	)
)

$ErrorActionPreference = "Stop"

try {
	if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
		throw "Microsoft.Graph.Users module is not installed. Install it using Install-Module Microsoft.Graph -Scope CurrentUser."
	}

	Import-Module Microsoft.Graph.Users -ErrorAction Stop

	$context = Get-MgContext
	if (-not $context) {
		Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All", "Organization.Read.All" -ErrorAction Stop | Out-Null
		$context = Get-MgContext
	}

	# --- Resolve tenant display name ---
	$tenantDisplayName = $null
	try {
		$org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
		$tenantDisplayName = $org.DisplayName
	} catch { }
	if (-not $tenantDisplayName) { $tenantDisplayName = $context.TenantId }
	$tenantId = if ($context.TenantId) { $context.TenantId } else { "Unknown" }

	$cutoffDate = (Get-Date).AddDays(-$InactiveDays)

	# --- SkuPartNumber -> Display Name mapping ---
	$skuDisplayNames = @{
		"O365_BUSINESS_ESSENTIALS"       = "Microsoft 365 Business Basic"
		"SMB_BUSINESS_ESSENTIALS"        = "Microsoft 365 Business Basic"
		"O365_BUSINESS_PREMIUM"          = "Microsoft 365 Business Standard"
		"SMB_BUSINESS_PREMIUM"           = "Microsoft 365 Business Standard"
		"SPB"                            = "Microsoft 365 Business Premium"
		"SPE_E3"                         = "Microsoft 365 E3"
		"SPE_E5"                         = "Microsoft 365 E5"
		"SPE_F1"                         = "Microsoft 365 F1"
		"M365_F1"                        = "Microsoft 365 F1"
		"MICROSOFT_365_FRONTLINE_F3"     = "Microsoft 365 F3"
		"ENTERPRISEPACK"                 = "Office 365 E3"
		"ENTERPRISEPREMIUM"              = "Office 365 E5"
		"STANDARDPACK"                   = "Office 365 E1"
		"STANDARDWOFFPACK"               = "Office 365 E2"
		"ENTERPRISEWITHSCAL"             = "Office 365 E4"
		"DESKLESSPACK"                   = "Office 365 F3"
		"SMB_BUSINESS"                   = "Microsoft 365 Apps for Business"
		"O365_BUSINESS"                  = "Microsoft 365 Apps for Business"
		"OFFICESUBSCRIPTION"             = "Microsoft 365 Apps for Enterprise"
		"EXCHANGESTANDARD"               = "Exchange Online Plan 1"
		"EXCHANGEENTERPRISE"             = "Exchange Online Plan 2"
		"EXCHANGEDESKLESS"               = "Exchange Online Kiosk"
		"DEVELOPERPACK"                  = "Office 365 E3 Developer"
		"DEVELOPERPACK_E5"               = "Microsoft 365 E5 Developer"
		"Microsoft_365_Copilot"          = "Microsoft 365 Copilot"
		"PROJECTPREMIUM"                 = "Project Plan 5"
		"PROJECTPROFESSIONAL"            = "Project Plan 3"
		"VISIOCLIENT"                    = "Visio Plan 2"
		"POWER_BI_PRO"                   = "Power BI Pro"
		"POWER_BI_STANDARD"              = "Power BI (Free)"
		"PBI_PREMIUM_PER_USER"           = "Power BI Premium Per User"
		"STREAM"                         = "Microsoft Stream"
		"EMS"                            = "Enterprise Mobility + Security E3"
		"EMSPREMIUM"                     = "Enterprise Mobility + Security E5"
		"AAD_PREMIUM"                    = "Microsoft Entra ID P1"
		"AAD_PREMIUM_P2"                 = "Microsoft Entra ID P2"
		"ATP_ENTERPRISE"                 = "Microsoft Defender for Office 365 Plan 1"
		"THREAT_INTELLIGENCE"            = "Microsoft Defender for Office 365 Plan 2"
		"ATA"                            = "Microsoft Defender for Identity"
		"WIN_DEF_ATP"                    = "Microsoft Defender for Endpoint Plan 2"
		"INTUNE_A"                       = "Microsoft Intune Plan 1"
		"RIGHTSMANAGEMENT"               = "Azure Information Protection Plan 1"
		"RIGHTSMANAGEMENT_ADHOC"         = "Rights Management Adhoc"
		"MCOIMP"                         = "Skype for Business Online Plan 1"
		"MCOSTANDARD"                    = "Skype for Business Online Plan 2"
		"PHONESYSTEM_VIRTUALUSER"        = "Microsoft Teams Phone Resource Account"
		"MCOCAP"                         = "Microsoft Teams Shared Devices"
		"MCOEV"                          = "Microsoft Teams Phone Standard"
		"Microsoft_Teams_Audio_Conferencing_select_dial_out" = "Microsoft Teams Audio Conferencing"
		"MEETING_ROOM"                   = "Microsoft Teams Rooms Standard"
		"Teams_Ess"                      = "Microsoft Teams Essentials"
		"TEAMS_EXPLORATORY"              = "Microsoft Teams Exploratory"
		"FLOW_FREE"                      = "Power Automate Free"
		"POWERAPPS_VIRAL"                = "Power Apps Plan 2 Trial"
		"WINDOWS_STORE"                  = "Windows Store for Business"
		"WIN10_PRO_ENT_SUB"              = "Windows 10/11 Enterprise E3"
		"WIN10_VDA_E5"                   = "Windows 10/11 Enterprise E5"
	}

	# --- Build SKU lookup: SkuId -> Display Name ---
	$skuLookup = @{}
	$primarySkuIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
	try {
		$subscribedSkus = Get-MgSubscribedSku -All -ErrorAction Stop
		foreach ($sku in $subscribedSkus) {
			if ($sku.SkuPartNumber -and $sku.SkuId) {
				$displayName = if ($skuDisplayNames.ContainsKey($sku.SkuPartNumber)) {
					$skuDisplayNames[$sku.SkuPartNumber]
				} else {
					$sku.SkuPartNumber
				}
				$skuLookup[$sku.SkuId] = $displayName
				if ($PrimaryLicenseSkus -contains $sku.SkuPartNumber) {
					[void]$primarySkuIds.Add($sku.SkuId)
				}
			}
		}
	} catch {
		Write-Warning "Could not retrieve subscribed SKUs: $($_.Exception.Message)"
	}

	# --- Fetch all member users ---
	$graphFilter = "userType eq 'Member'"

	$users = Get-MgUser -All `
		-Filter $graphFilter `
		-Property "id,displayName,userPrincipalName,mail,userType,accountEnabled,createdDateTime,signInActivity,assignedLicenses,onPremisesSyncEnabled,jobTitle,department" `
		-ConsistencyLevel eventual

	# --- Build report data for ALL users ---
	$report = foreach ($user in $users) {
		$signInActivity = $user.SignInActivity
		$lastInteractive = $signInActivity.lastSignInDateTime
		$lastNonInteractive = $signInActivity.lastNonInteractiveSignInDateTime

		$lastInteractiveDt = if ($lastInteractive) { [datetime]$lastInteractive } else { $null }
		$lastNonInteractiveDt = if ($lastNonInteractive) { [datetime]$lastNonInteractive } else { $null }

		$mostRecent = @($lastInteractiveDt, $lastNonInteractiveDt) | Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1
		$lastSignInAgo = if ($mostRecent) { [int]((Get-Date) - $mostRecent).TotalDays } else { "Never" }

		$isInactive = $false
		if (-not $lastInteractiveDt -and -not $lastNonInteractiveDt) {
			$isInactive = $true
		} elseif ((-not $lastInteractiveDt -or $lastInteractiveDt -lt $cutoffDate) -and (-not $lastNonInteractiveDt -or $lastNonInteractiveDt -lt $cutoffDate)) {
			$isInactive = $true
		}

		$isLicensed = ($user.AssignedLicenses.Count -gt 0)

		# Resolve license names
		$userLicenseNames = @()
		$primaryLicense = $null
		foreach ($assignedLic in $user.AssignedLicenses) {
			$skuId = $assignedLic.SkuId
			$name = if ($skuLookup.ContainsKey($skuId)) { $skuLookup[$skuId] } else { $skuId }
			$userLicenseNames += $name
			if (-not $primaryLicense -and $primarySkuIds.Contains($skuId)) {
				$primaryLicense = $name
			}
		}
		$allLicenses = ($userLicenseNames | Sort-Object) -join " | "
		if (-not $primaryLicense -and $isLicensed) { $primaryLicense = "Unidentified" }
		if (-not $primaryLicense) { $primaryLicense = "None" }

		# Status: Disabled accounts are always "Disabled"; only enabled accounts are Active/Inactive
		if (-not $user.AccountEnabled) {
			$status = "Disabled"
			$isInactive = $false
		} else {
			$status = if ($isInactive) { "Inactive" } else { "Active" }
		}

		[pscustomobject]@{
			DisplayName              = $user.DisplayName
			UserPrincipalName        = $user.UserPrincipalName
			Mail                     = $user.Mail
			PrimaryDomain            = if ($user.Mail -and $user.Mail -match "@") { ($user.Mail -split "@", 2)[1] } else { $null }
			JobTitle                 = $user.JobTitle
			Department               = $user.Department
			UserType                 = $user.UserType
			AccountEnabled           = $user.AccountEnabled
			LicenseAssigned          = $isLicensed
			PrimaryLicense           = $primaryLicense
			AllLicenses              = $allLicenses
			OnPremisesSyncEnabled    = [bool]$user.OnPremisesSyncEnabled
			CreatedDateTime          = $user.CreatedDateTime
			LastInteractiveSignIn    = $lastInteractiveDt
			LastNonInteractiveSignIn = $lastNonInteractiveDt
			LastSignInAgoDays        = $lastSignInAgo
			Status                   = $status
			Inactive                 = $isInactive
		}
	}

	# --- Stats for console output (using PowerShell InactiveDays param) ---
	$totalMembers    = $report.Count
	$totalDisabled   = ($report | Where-Object { -not $_.AccountEnabled }).Count
	$totalEnabled    = $totalMembers - $totalDisabled
	$totalActive     = ($report | Where-Object { $_.AccountEnabled -and -not $_.Inactive }).Count
	$totalInactive   = ($report | Where-Object { $_.AccountEnabled -and $_.Inactive }).Count
	$percentInactive = if ($totalEnabled -gt 0) { [math]::Round(($totalInactive / $totalEnabled) * 100, 1) } else { 0 }
	$percentActive   = if ($totalEnabled -gt 0) { [math]::Round(($totalActive / $totalEnabled) * 100, 1) } else { 0 }
	$totalLicensed   = ($report | Where-Object { $_.LicenseAssigned }).Count
	$totalUnlicensed = $totalMembers - $totalLicensed
	$licensedActive   = ($report | Where-Object { $_.LicenseAssigned -and $_.AccountEnabled -and -not $_.Inactive }).Count
	$licensedInactive = ($report | Where-Object { $_.LicenseAssigned -and $_.AccountEnabled -and $_.Inactive }).Count
	$licensedDisabled = ($report | Where-Object { $_.LicenseAssigned -and -not $_.AccountEnabled }).Count
	$unlicensedActive   = ($report | Where-Object { -not $_.LicenseAssigned -and $_.AccountEnabled -and -not $_.Inactive }).Count
	$unlicensedInactive = ($report | Where-Object { -not $_.LicenseAssigned -and $_.AccountEnabled -and $_.Inactive }).Count
	$unlicensedDisabled = ($report | Where-Object { -not $_.LicenseAssigned -and -not $_.AccountEnabled }).Count
	$totalOnPrem     = ($report | Where-Object { $_.OnPremisesSyncEnabled }).Count
	$totalCloudOnly  = $totalMembers - $totalOnPrem
	$neverSignedIn   = ($report | Where-Object { $_.LastSignInAgoDays -eq "Never" -and $_.AccountEnabled }).Count

	# --- File paths ---
	if (-not $ReportPath) {
		$ReportPath = (Get-Location).Path
	}

	$reportFolder = if (Test-Path $ReportPath -PathType Container) { $ReportPath } else { Split-Path -Parent $ReportPath }
	if ($reportFolder -and -not (Test-Path $reportFolder)) {
		New-Item -ItemType Directory -Path $reportFolder -Force | Out-Null
	}

	$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
	$csvFile = if (Test-Path $ReportPath -PathType Container) {
		Join-Path $ReportPath ("AllMemberUsers_{0}.csv" -f $timestamp)
	} else {
		$ReportPath
	}

	# --- CSV: export ALL users with all columns ---
	$report | Sort-Object DisplayName | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8

	# --- Build per-user JSON for client-side threshold recalculation ---
	$usersJson = ($report | Sort-Object DisplayName | ForEach-Object {
		$daysVal = if ($_.LastSignInAgoDays -eq "Never") { -1 } else { [int]$_.LastSignInAgoDays }
		$lic = if ($_.LicenseAssigned) { "true" } else { "false" }
		$onp = if ($_.OnPremisesSyncEnabled) { "true" } else { "false" }
		$ena = if ($_.AccountEnabled) { "true" } else { "false" }
		'{{"days":{0},"lic":{1},"onp":{2},"ena":{3}}}' -f $daysVal, $lic, $onp, $ena
	}) -join ","

	# --- Build HTML table rows (all users, status computed at default threshold) ---
	$tableRows = ($report | Sort-Object DisplayName | ForEach-Object {
		$daysVal = if ($_.LastSignInAgoDays -eq "Never") { -1 } else { [int]$_.LastSignInAgoDays }
		$signInAge = if ($_.LastSignInAgoDays -eq "Never") { "Never" } else { "{0} days" -f $_.LastSignInAgoDays }
		$licensed = if ($_.LicenseAssigned) { "Yes" } else { "No" }
		$syncEnabled = if ($_.OnPremisesSyncEnabled) { "Yes" } else { "No" }
		$enabled = if ($_.AccountEnabled) { "Yes" } else { "No" }
		$statusClass = switch ($_.Status) { "Active" { "active" } "Inactive" { "inactive" } "Disabled" { "disabled" } }
		"<tr data-days=`"{0}`" data-lic=`"{1}`" data-onp=`"{2}`" data-ena=`"{3}`"><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td><td><span class=`"badge badge-{10}`">{11}</span></td><td>{12}</td><td>{13}</td><td>{14}</td><td>{15}</td><td title=`"{16}`">{16}</td></tr>" -f
			$daysVal,
			$(if ($_.LicenseAssigned) { "1" } else { "0" }),
			$(if ($_.OnPremisesSyncEnabled) { "1" } else { "0" }),
			$(if ($_.AccountEnabled) { "1" } else { "0" }),
			[System.Net.WebUtility]::HtmlEncode($_.DisplayName),
			[System.Net.WebUtility]::HtmlEncode($_.UserPrincipalName),
			[System.Net.WebUtility]::HtmlEncode($_.Mail),
			[System.Net.WebUtility]::HtmlEncode($_.Department),
			[System.Net.WebUtility]::HtmlEncode($_.PrimaryLicense),
			[System.Net.WebUtility]::HtmlEncode($licensed),
			$statusClass,
			[System.Net.WebUtility]::HtmlEncode($_.Status),
			[System.Net.WebUtility]::HtmlEncode($enabled),
			[System.Net.WebUtility]::HtmlEncode($syncEnabled),
			[System.Net.WebUtility]::HtmlEncode($signInAge),
			[System.Net.WebUtility]::HtmlEncode($_.CreatedDateTime),
			[System.Net.WebUtility]::HtmlEncode($_.AllLicenses)
	}) -join "`n"

	$reportDate = Get-Date -Format "dd MMM yyyy HH:mm"
	$totalCount = $report.Count

	$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>All Member Users Report</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; padding: 30px; }
  .header { background: linear-gradient(135deg, #1a1a2e, #16213e); color: #fff; padding: 30px 40px; border-radius: 12px; margin-bottom: 30px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 16px; }
  .header-left h1 { font-size: 1.6em; margin-bottom: 6px; }
  .header-left p { font-size: 0.9em; opacity: 0.8; }
  .header-right { display: flex; align-items: center; gap: 10px; }
  .header-right label { font-size: 0.9em; opacity: 0.85; }
  .header-right select { padding: 8px 14px; border: none; border-radius: 6px; font-size: 0.95em; font-weight: 600; background: rgba(255,255,255,0.15); color: #fff; cursor: pointer; }
  .header-right select option { color: #333; background: #fff; }

  .summary-cards { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .card { background: #fff; border-radius: 10px; padding: 24px 30px; flex: 1; min-width: 180px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
  .card .label { font-size: 0.85em; color: #777; text-transform: uppercase; letter-spacing: 0.5px; }
  .card .value { font-size: 2em; font-weight: 700; margin-top: 6px; }
  .card .sub { font-size: 0.78em; color: #999; margin-top: 2px; }

  .section-title { font-size: 1.15em; font-weight: 600; color: #1a1a2e; margin-bottom: 16px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0; }

  .info-cards { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .info-card { background: #fff; border-radius: 10px; padding: 20px 26px; flex: 1; min-width: 200px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); border-left: 5px solid #ccc; }
  .info-card .info-label { font-size: 0.82em; color: #555; margin-bottom: 6px; }
  .info-card .info-value { font-size: 1.8em; font-weight: 700; }
  .info-card .info-sub { font-size: 0.75em; color: #999; margin-top: 2px; }
  .info-card.green { border-left-color: #27ae60; } .info-card.green .info-value { color: #27ae60; }
  .info-card.red { border-left-color: #e74c3c; } .info-card.red .info-value { color: #e74c3c; }
  .info-card.purple { border-left-color: #8e44ad; } .info-card.purple .info-value { color: #8e44ad; }
  .info-card.teal { border-left-color: #16a085; } .info-card.teal .info-value { color: #16a085; }
  .info-card.indigo { border-left-color: #2c3e50; } .info-card.indigo .info-value { color: #2c3e50; }
  .info-card.coral { border-left-color: #e67e22; } .info-card.coral .info-value { color: #e67e22; }
  .info-card.grey { border-left-color: #7f8c8d; } .info-card.grey .info-value { color: #7f8c8d; }
  .info-card.blue { border-left-color: #3498db; } .info-card.blue .info-value { color: #3498db; }

  .charts-row { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .chart-section { background: #fff; border-radius: 10px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); flex: 1; min-width: 360px; }
  .chart-section h2 { font-size: 1.1em; margin-bottom: 20px; color: #1a1a2e; }
  .chart-container { max-width: 450px; margin: 0 auto; }

  .badge { padding: 3px 10px; border-radius: 12px; font-size: 0.82em; font-weight: 600; }
  .badge-active { background: #d4edda; color: #155724; }
  .badge-inactive { background: #f8d7da; color: #721c24; }
  .badge-disabled { background: #e2e3e5; color: #495057; }

  .table-controls { display: flex; gap: 14px; margin-bottom: 16px; flex-wrap: wrap; align-items: center; }
  .table-controls input[type="text"] { padding: 8px 14px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; min-width: 260px; }
  .table-controls select { padding: 8px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; background: #fff; }
  .table-controls .count-label { font-size: 0.85em; color: #777; margin-left: auto; }

  .table-section { background: #fff; border-radius: 10px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 30px; overflow-x: auto; }
  .table-section h2 { font-size: 1.1em; margin-bottom: 16px; color: #1a1a2e; }
  table { width: 100%; border-collapse: collapse; font-size: 0.86em; }
  th { background: #1a1a2e; color: #fff; padding: 12px 14px; text-align: left; position: sticky; top: 0; cursor: pointer; user-select: none; white-space: nowrap; }
  th:hover { background: #2c3e50; }
  td { padding: 10px 14px; border-bottom: 1px solid #eee; }
  tr:hover td { background: #f8f9fa; }
  tr.hidden-row { display: none; }

  .footer { text-align: center; font-size: 0.8em; color: #999; margin-top: 20px; }
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <h1>All Member Users Report</h1>
    <p>Tenant: $([System.Net.WebUtility]::HtmlEncode($tenantDisplayName)) ($tenantId) &nbsp;|&nbsp; Generated: $reportDate</p>
  </div>
  <div class="header-right">
    <label for="thresholdSelect">Inactive Threshold:</label>
    <select id="thresholdSelect" onchange="applyThreshold()">
      <option value="30" $(if ($InactiveDays -eq 30) { 'selected' })>30 Days</option>
      <option value="60" $(if ($InactiveDays -eq 60) { 'selected' })>60 Days</option>
      <option value="90" $(if ($InactiveDays -eq 90 -or ($InactiveDays -ne 30 -and $InactiveDays -ne 60 -and $InactiveDays -ne 180)) { 'selected' })>90 Days</option>
      <option value="180" $(if ($InactiveDays -eq 180) { 'selected' })>180 Days</option>
    </select>
  </div>
</div>

<!-- OVERVIEW -->
<div class="section-title">Overview</div>
<div class="summary-cards">
  <div class="card"><div class="label">Total Members</div><div class="value" style="color:#1a1a2e;" id="cardTotal">$totalCount</div></div>
  <div class="card"><div class="label">Active</div><div class="value" style="color:#27ae60;" id="cardActive">-</div><div class="sub" id="cardActivePct"></div></div>
  <div class="card"><div class="label">Inactive</div><div class="value" style="color:#e74c3c;" id="cardInactive">-</div><div class="sub" id="cardInactivePct"></div></div>
  <div class="card"><div class="label">Disabled</div><div class="value" style="color:#6c757d;" id="cardDisabled">-</div><div class="sub" id="cardDisabledPct"></div></div>
  <div class="card"><div class="label">Never Signed In</div><div class="value" style="color:#7f8c8d;" id="cardNever">-</div><div class="sub">(enabled only)</div></div>
</div>

<!-- LICENSED USERS -->
<div class="section-title">Licensed Users</div>
<div class="info-cards">
  <div class="info-card blue"><div class="info-label">Total Licensed</div><div class="info-value" id="cardLicensed">-</div></div>
  <div class="info-card green"><div class="info-label">Licensed &amp; Active</div><div class="info-value" id="cardLicActive">-</div><div class="info-sub" id="cardLicActivePct"></div></div>
  <div class="info-card red"><div class="info-label">Licensed &amp; Inactive</div><div class="info-value" id="cardLicInactive">-</div><div class="info-sub" id="cardLicInactivePct"></div></div>
  <div class="info-card grey"><div class="info-label">Licensed &amp; Disabled</div><div class="info-value" id="cardLicDisabled">-</div><div class="info-sub" id="cardLicDisabledPct"></div></div>
</div>

<!-- UNLICENSED USERS -->
<div class="section-title">Unlicensed Users</div>
<div class="info-cards">
  <div class="info-card indigo"><div class="info-label">Total Unlicensed</div><div class="info-value" id="cardUnlicensed">-</div></div>
  <div class="info-card teal"><div class="info-label">Unlicensed &amp; Active</div><div class="info-value" id="cardUnlicActive">-</div><div class="info-sub" id="cardUnlicActivePct"></div></div>
  <div class="info-card coral"><div class="info-label">Unlicensed &amp; Inactive</div><div class="info-value" id="cardUnlicInactive">-</div><div class="info-sub" id="cardUnlicInactivePct"></div></div>
  <div class="info-card grey"><div class="info-label">Unlicensed &amp; Disabled</div><div class="info-value" id="cardUnlicDisabled">-</div><div class="info-sub" id="cardUnlicDisabledPct"></div></div>
</div>

<!-- IDENTITY SOURCE -->
<div class="section-title">Identity Source</div>
<div class="info-cards">
  <div class="info-card indigo"><div class="info-label">On-Prem Synced</div><div class="info-value" id="cardOnPrem">-</div></div>
  <div class="info-card purple"><div class="info-label">Cloud-Only</div><div class="info-value" id="cardCloud">-</div></div>
</div>

<!-- CHARTS -->
<div class="charts-row">
  <div class="chart-section">
    <h2>Active vs Inactive Members</h2>
    <div class="chart-container"><canvas id="activityChart"></canvas></div>
  </div>
  <div class="chart-section">
    <h2>Licensed Users: Active vs Inactive</h2>
    <div class="chart-container"><canvas id="licensedChart"></canvas></div>
  </div>
</div>

<div class="charts-row">
  <div class="chart-section">
    <h2>Licensed vs Unlicensed</h2>
    <div class="chart-container"><canvas id="licenseChart"></canvas></div>
  </div>
  <div class="chart-section">
    <h2>Inactive Members by Last Sign-In Age</h2>
    <div class="chart-container"><canvas id="ageBucketChart"></canvas></div>
  </div>
</div>

<!-- TABLE -->
<div class="table-section">
  <h2>Member User Details</h2>
  <div class="table-controls">
    <input type="text" id="searchBox" placeholder="Search by name, UPN, mail, department..." onkeyup="filterTable()" />
    <select id="statusFilter" onchange="filterTable()">
      <option value="all">All Status</option>
      <option value="active">Active Only</option>
      <option value="inactive">Inactive Only</option>
      <option value="disabled">Disabled Only</option>
    </select>
    <select id="licenseFilter" onchange="filterTable()">
      <option value="all">All License</option>
      <option value="yes">Licensed Only</option>
      <option value="no">Unlicensed Only</option>
    </select>
    <span class="count-label" id="rowCount"></span>
  </div>
  <table id="userTable">
    <thead><tr>
      <th onclick="sortTable(0)">Display Name</th>
      <th onclick="sortTable(1)">UPN</th>
      <th onclick="sortTable(2)">Mail</th>
      <th onclick="sortTable(3)">Department</th>
      <th onclick="sortTable(4)">Primary License</th>
      <th onclick="sortTable(5)">Licensed</th>
      <th onclick="sortTable(6)">Status</th>
      <th onclick="sortTable(7)">Enabled</th>
      <th onclick="sortTable(8)">On-Prem Sync</th>
      <th onclick="sortTable(9)">Last Sign-In</th>
      <th onclick="sortTable(10)">Created</th>
      <th onclick="sortTable(11)">All Licenses</th>
    </tr></thead>
    <tbody>
$tableRows
    </tbody>
  </table>
</div>

<div class="footer">Report generated by ReportAllMemberUsers.ps1</div>

<script>
// --- Embedded user data for client-side threshold recalculation ---
var userData = [$usersJson];

var chartOpts = function(pos) {
  return { responsive: true, plugins: { legend: { position: pos || 'right', labels: { padding: 20, font: { size: 14 }, boxWidth: 18 } }, tooltip: { callbacks: { label: function(ctx) { var t = ctx.dataset.data.reduce(function(a,b){return a+b},0); return ctx.label+': '+ctx.parsed+' ('+(t>0?((ctx.parsed/t)*100).toFixed(1):0)+'%)'; } } } } };
};

var activityChart = new Chart(document.getElementById('activityChart'), { type:'doughnut', data:{ labels:['Active','Inactive','Disabled'], datasets:[{ data:[0,0,0], backgroundColor:['#27ae60','#e74c3c','#6c757d'], borderWidth:2, borderColor:'#fff' }] }, options: chartOpts() });
var licensedChart = new Chart(document.getElementById('licensedChart'), { type:'doughnut', data:{ labels:['Licensed Active','Licensed Inactive'], datasets:[{ data:[0,0], backgroundColor:['#27ae60','#e74c3c'], borderWidth:2, borderColor:'#fff' }] }, options: chartOpts() });
var licenseChart = new Chart(document.getElementById('licenseChart'), { type:'doughnut', data:{ labels:['Licensed','Unlicensed'], datasets:[{ data:[0,0], backgroundColor:['#3498db','#95a5a6'], borderWidth:2, borderColor:'#fff' }] }, options: chartOpts() });
var ageBucketChart = new Chart(document.getElementById('ageBucketChart'), { type:'pie', data:{ labels:[], datasets:[{ data:[], backgroundColor:[], borderWidth:2, borderColor:'#fff' }] }, options: chartOpts() });

function pct(n, total) { return total > 0 ? ((n / total) * 100).toFixed(1) : '0.0'; }

function applyThreshold() {
  var threshold = parseInt(document.getElementById('thresholdSelect').value);
  var total = userData.length;
  var active = 0, inactive = 0, disabled = 0, never = 0;
  var licensed = 0, unlicensed = 0;
  var licActive = 0, licInactive = 0, licDisabled = 0, unlicActive = 0, unlicInactive = 0, unlicDisabled = 0;
  var onPrem = 0, cloud = 0;
  var b1Max = threshold * 2, b2Max = 365;
  var b1 = 0, b2 = 0, b3 = 0, b4 = 0;

  for (var i = 0; i < userData.length; i++) {
    var u = userData[i];
    if (u.lic) licensed++; else unlicensed++;
    if (u.onp) onPrem++; else cloud++;

    if (!u.ena) {
      disabled++;
      if (u.lic) licDisabled++; else unlicDisabled++;
      continue;
    }

    // Only enabled accounts are evaluated for active/inactive
    var isInactive = (u.days === -1) || (u.days >= threshold);
    if (u.days === -1) never++;

    if (isInactive) {
      inactive++;
      if (u.lic) licInactive++; else unlicInactive++;
      if (u.days === -1) b4++;
      else if (u.days > b2Max) b3++;
      else if (u.days > b1Max) b2++;
      else b1++;
    } else {
      active++;
      if (u.lic) licActive++; else unlicActive++;
    }
  }

  var enabled = active + inactive;

  // Update cards
  document.getElementById('cardTotal').textContent = total;
  document.getElementById('cardActive').textContent = active;
  document.getElementById('cardActivePct').textContent = pct(active, enabled) + '% of enabled';
  document.getElementById('cardInactive').textContent = inactive;
  document.getElementById('cardInactivePct').textContent = pct(inactive, enabled) + '% of enabled';
  document.getElementById('cardDisabled').textContent = disabled;
  document.getElementById('cardDisabledPct').textContent = pct(disabled, total) + '% of members';
  document.getElementById('cardNever').textContent = never;

  document.getElementById('cardLicensed').textContent = licensed;
  document.getElementById('cardLicActive').textContent = licActive;
  document.getElementById('cardLicActivePct').textContent = pct(licActive, licensed) + '% of licensed';
  document.getElementById('cardLicInactive').textContent = licInactive;
  document.getElementById('cardLicInactivePct').textContent = pct(licInactive, licensed) + '% of licensed';
  document.getElementById('cardLicDisabled').textContent = licDisabled;
  document.getElementById('cardLicDisabledPct').textContent = pct(licDisabled, licensed) + '% of licensed';

  document.getElementById('cardUnlicensed').textContent = unlicensed;
  document.getElementById('cardUnlicActive').textContent = unlicActive;
  document.getElementById('cardUnlicActivePct').textContent = pct(unlicActive, unlicensed) + '% of unlicensed';
  document.getElementById('cardUnlicInactive').textContent = unlicInactive;
  document.getElementById('cardUnlicInactivePct').textContent = pct(unlicInactive, unlicensed) + '% of unlicensed';
  document.getElementById('cardUnlicDisabled').textContent = unlicDisabled;
  document.getElementById('cardUnlicDisabledPct').textContent = pct(unlicDisabled, unlicensed) + '% of unlicensed';

  document.getElementById('cardOnPrem').textContent = onPrem;
  document.getElementById('cardCloud').textContent = cloud;

  // Update table row statuses
  var rows = document.querySelectorAll('#userTable tbody tr');
  for (var j = 0; j < rows.length; j++) {
    var days = parseInt(rows[j].getAttribute('data-days'));
    var rowEna = rows[j].getAttribute('data-ena');
    var badge = rows[j].cells[6].querySelector('.badge');
    if (rowEna === '0') {
      badge.className = 'badge badge-disabled';
      badge.textContent = 'Disabled';
      rows[j].setAttribute('data-status', 'disabled');
    } else {
      var rowInactive = (days === -1) || (days >= threshold);
      if (rowInactive) {
        badge.className = 'badge badge-inactive';
        badge.textContent = 'Inactive';
        rows[j].setAttribute('data-status', 'inactive');
      } else {
        badge.className = 'badge badge-active';
        badge.textContent = 'Active';
        rows[j].setAttribute('data-status', 'active');
      }
    }
  }

  // Update charts
  activityChart.data.datasets[0].data = [active, inactive, disabled];
  activityChart.update();

  licensedChart.data.datasets[0].data = [licActive, licInactive];
  licensedChart.update();

  licenseChart.data.datasets[0].data = [licensed, unlicensed];
  licenseChart.update();

  // Age bucket chart
  var bucketLabels, bucketData, bucketColors;
  if (b1Max >= b2Max) {
    bucketLabels = [threshold+'-'+b1Max+' Days', 'Over 1 Year', 'Never Signed In'];
    bucketData = [b1, b2 + b3, b4];
    bucketColors = ['#3498db', '#e74c3c', '#95a5a6'];
  } else {
    bucketLabels = [threshold+'-'+b1Max+' Days', b1Max+'-'+b2Max+' Days', 'Over 1 Year', 'Never Signed In'];
    bucketData = [b1, b2, b3, b4];
    bucketColors = ['#3498db', '#f39c12', '#e74c3c', '#95a5a6'];
  }
  ageBucketChart.data.labels = bucketLabels;
  ageBucketChart.data.datasets[0].data = bucketData;
  ageBucketChart.data.datasets[0].backgroundColor = bucketColors;
  ageBucketChart.update();

  filterTable();
}

function filterTable() {
  var search = document.getElementById('searchBox').value.toLowerCase();
  var status = document.getElementById('statusFilter').value;
  var license = document.getElementById('licenseFilter').value;
  var rows = document.querySelectorAll('#userTable tbody tr');
  var visible = 0;
  rows.forEach(function(row) {
    var text = row.textContent.toLowerCase();
    var matchSearch = !search || text.indexOf(search) !== -1;
    var rowStatus = row.getAttribute('data-status');
    var rowLic = row.getAttribute('data-lic');
    var matchStatus = status === 'all' || rowStatus === status;
    var matchLicense = license === 'all' || (license === 'yes' && rowLic === '1') || (license === 'no' && rowLic === '0');
    if (matchSearch && matchStatus && matchLicense) {
      row.classList.remove('hidden-row');
      visible++;
    } else {
      row.classList.add('hidden-row');
    }
  });
  document.getElementById('rowCount').textContent = 'Showing ' + visible + ' of ' + rows.length + ' users';
}

var sortDir = {};
function sortTable(col) {
  var tbody = document.getElementById('userTable').querySelector('tbody');
  var rows = Array.from(tbody.querySelectorAll('tr'));
  var dir = sortDir[col] === 'asc' ? 'desc' : 'asc';
  sortDir[col] = dir;
  rows.sort(function(a, b) {
    var av = a.cells[col].textContent.trim().toLowerCase();
    var bv = b.cells[col].textContent.trim().toLowerCase();
    if (av < bv) return dir === 'asc' ? -1 : 1;
    if (av > bv) return dir === 'asc' ? 1 : -1;
    return 0;
  });
  rows.forEach(function(row) { tbody.appendChild(row); });
}

// Initial render with default threshold
applyThreshold();
</script>
</body>
</html>
"@

	$htmlReportFile = Join-Path $reportFolder ("AllMemberUsers_{0}.html" -f $timestamp)
	$html | Out-File -FilePath $htmlReportFile -Encoding UTF8

	# --- Console summary ---
	Write-Host ""
	Write-Host "All Member Users Report" -ForegroundColor Cyan
	Write-Host "--------------------------------------------"
	Write-Host ("Tenant                   : {0} ({1})" -f $tenantDisplayName, $tenantId)
	Write-Host ("Total members            : {0}" -f $totalMembers)
	Write-Host ("Active members           : {0}  ({1}% of enabled)" -f $totalActive, $percentActive) -ForegroundColor Green
	Write-Host ("Inactive members         : {0}  ({1}% of enabled)" -f $totalInactive, $percentInactive) -ForegroundColor Red
	Write-Host ("Disabled members         : {0}" -f $totalDisabled) -ForegroundColor DarkGray
	Write-Host ("Never signed in          : {0}  (enabled only)" -f $neverSignedIn)
	Write-Host ""
	Write-Host "Licensed Users" -ForegroundColor Cyan
	Write-Host ("  Total licensed         : {0}" -f $totalLicensed)
	Write-Host ("  Licensed & active      : {0}" -f $licensedActive) -ForegroundColor Green
	Write-Host ("  Licensed & inactive    : {0}" -f $licensedInactive) -ForegroundColor Red
	Write-Host ("  Licensed & disabled    : {0}" -f $licensedDisabled) -ForegroundColor DarkGray
	Write-Host ""
	Write-Host "Unlicensed Users" -ForegroundColor Cyan
	Write-Host ("  Total unlicensed       : {0}" -f $totalUnlicensed)
	Write-Host ("  Unlicensed & active    : {0}" -f $unlicensedActive) -ForegroundColor Green
	Write-Host ("  Unlicensed & inactive  : {0}" -f $unlicensedInactive) -ForegroundColor Red
	Write-Host ("  Unlicensed & disabled  : {0}" -f $unlicensedDisabled) -ForegroundColor DarkGray
	Write-Host ""
	Write-Host "Identity Source" -ForegroundColor Cyan
	Write-Host ("  On-prem synced         : {0}" -f $totalOnPrem)
	Write-Host ("  Cloud-only             : {0}" -f $totalCloudOnly)
	Write-Host ""
	Write-Host ("Inactive days threshold  : {0}" -f $InactiveDays)
	Write-Host ("Users in export          : {0}" -f $totalCount)
	Write-Host ("CSV report               : {0}" -f $csvFile) -ForegroundColor Yellow
	Write-Host ("HTML report              : {0}" -f $htmlReportFile) -ForegroundColor Yellow

	$disconnectChoice = Read-Host "Disconnect from Microsoft Graph? (Y/N)"
	if ($disconnectChoice -match '^(y|yes)$') {
		Disconnect-MgGraph -ErrorAction SilentlyContinue
	}
}
catch {
	Write-Error $_
	exit 1
}
