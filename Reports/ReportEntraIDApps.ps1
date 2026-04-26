param(
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[string]$ReportPath,

	[Parameter(Mandatory = $false)]
	[ValidateRange(1, 3650)]
	[int]$InactiveDays = 180
)

$ErrorActionPreference = "Stop"

$S_RequiredGraphScopes = @(
	'Application.Read.All'
	'AuditLog.Read.All'
	'Organization.Read.All'
)

try {
	# --- Module check ---
	$requiredModules = @('Microsoft.Graph.Applications', 'Microsoft.Graph.Identity.DirectoryManagement')
	foreach ($mod in $requiredModules) {
		if (-not (Get-Module -ListAvailable -Name $mod)) {
			throw "$mod module is not installed. Install it using Install-Module Microsoft.Graph -Scope CurrentUser."
		}
	}
	Import-Module Microsoft.Graph.Applications -ErrorAction Stop
	Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

	# --- Connect to Graph ---
	$context = Get-MgContext
	if (-not $context) {
		Connect-MgGraph -Scopes $S_RequiredGraphScopes -ErrorAction Stop | Out-Null
		$context = Get-MgContext
	}

	# --- Tenant info ---
	$tenantDisplayName = $null
	try {
		$org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
		$tenantDisplayName = $org.DisplayName
	} catch { }
	if (-not $tenantDisplayName) { $tenantDisplayName = $context.TenantId }
	$tenantId = if ($context.TenantId) { $context.TenantId } else { "Unknown" }

	# --- Fetch all Service Principals (Enterprise Applications) ---
	Write-Host "Fetching all enterprise applications (service principals)..." -ForegroundColor Cyan
	$servicePrincipals = Get-MgServicePrincipal -All `
		-Property "id,appId,displayName,servicePrincipalType,accountEnabled,appOwnerOrganizationId,signInAudience,createdDateTime,passwordCredentials,keyCredentials,signInActivity,tags,notes" `
		-ErrorAction Stop
	Write-Host "  Found $($servicePrincipals.Count) service principals" -ForegroundColor Green

	# --- Fetch App Registrations to identify which SPs have a local app reg ---
	Write-Host "Fetching app registrations for cross-reference..." -ForegroundColor Cyan
	$appRegLookup = @{}
	try {
		$appRegistrations = Get-MgApplication -All `
			-Property "id,appId,displayName,passwordCredentials,keyCredentials" `
			-ErrorAction Stop
		foreach ($ar in $appRegistrations) {
			if ($ar.AppId) { $appRegLookup[$ar.AppId] = $ar }
		}
		Write-Host "  Found $($appRegistrations.Count) app registrations" -ForegroundColor Green
	} catch {
		Write-Warning "Could not fetch app registrations. HasAppRegistration column will be unavailable."
	}

	# --- Resolve Microsoft Graph permission names ---
	Write-Host "Resolving Graph permission names..." -ForegroundColor Cyan
	$graphAppId = '00000003-0000-0000-c000-000000000000'
	$graphAppRoles = @{}
	$graphSpnId = $null
	try {
		$graphSpn = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'" -Property "id,appRoles" -ErrorAction Stop
		$graphSpnId = $graphSpn.Id
		foreach ($role in $graphSpn.AppRoles) {
			$graphAppRoles[$role.Id] = $role.Value
		}
	} catch {
		Write-Warning "Could not resolve Microsoft Graph permission names."
	}

	# --- Fetch all granted Graph app role assignments (efficient single call) ---
	Write-Host "Fetching granted Microsoft Graph permissions..." -ForegroundColor Cyan
	$grantedPermsLookup = @{}
	if ($graphSpnId) {
		try {
			$graphAssignments = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $graphSpnId -All -ErrorAction Stop
			foreach ($assignment in $graphAssignments) {
				$principalId = $assignment.PrincipalId
				$roleName = $graphAppRoles[$assignment.AppRoleId]
				if ($roleName) {
					if (-not $grantedPermsLookup.ContainsKey($principalId)) {
						$grantedPermsLookup[$principalId] = [System.Collections.Generic.List[string]]::new()
					}
					$grantedPermsLookup[$principalId].Add($roleName)
				}
			}
			Write-Host "  Found $($graphAssignments.Count) granted Graph permissions" -ForegroundColor Green
		} catch {
			Write-Warning "Could not fetch Graph app role assignments. Permission data may be incomplete."
		}
	}

	# --- Define high-privilege application permissions ---
	$highPrivilegePermissions = @(
		'Directory.ReadWrite.All',
		'RoleManagement.ReadWrite.Directory',
		'Application.ReadWrite.All',
		'AppRoleAssignment.ReadWrite.All',
		'Mail.ReadWrite',
		'Mail.Send',
		'MailboxSettings.ReadWrite',
		'Files.ReadWrite.All',
		'Sites.ReadWrite.All',
		'Sites.FullControl.All',
		'User.ReadWrite.All',
		'Group.ReadWrite.All',
		'GroupMember.ReadWrite.All',
		'Policy.ReadWrite.ConditionalAccess',
		'UserAuthenticationMethod.ReadWrite.All',
		'Chat.ReadWrite.All',
		'ChannelMessage.Send',
		'TeamSettings.ReadWrite.All'
	)

	# Microsoft's tenant ID for first-party app detection
	$microsoftTenantId = 'f8cdef31-a31e-4b4a-93e4-5f571e91255a'

	# --- Build report data ---
	Write-Host "Building report data for $($servicePrincipals.Count) enterprise applications..." -ForegroundColor Cyan
	$now = Get-Date
	$cutoffDate = $now.AddDays(-$InactiveDays)
	$expiringThresholdDate = $now.AddDays(30)

	$report = foreach ($sp in $servicePrincipals) {
		# Microsoft first-party detection
		$isMicrosoft = ($sp.AppOwnerOrganizationId -eq $microsoftTenantId)

		# Has app registration in this tenant?
		$hasAppReg = $appRegLookup.ContainsKey($sp.AppId)
		$linkedAppReg = if ($hasAppReg) { $appRegLookup[$sp.AppId] } else { $null }

		# --- Credentials (from SP + linked App Reg) ---
		$allCreds = @()
		# Credentials on the Service Principal itself (SAML certs, etc.)
		if ($sp.PasswordCredentials) {
			foreach ($pc in $sp.PasswordCredentials) {
				$allCreds += [pscustomobject]@{ Type = 'Secret'; EndDateTime = $pc.EndDateTime; Source = 'SP' }
			}
		}
		if ($sp.KeyCredentials) {
			foreach ($kc in $sp.KeyCredentials) {
				$allCreds += [pscustomobject]@{ Type = 'Certificate'; EndDateTime = $kc.EndDateTime; Source = 'SP' }
			}
		}
		# Credentials on the linked App Registration
		if ($linkedAppReg) {
			if ($linkedAppReg.PasswordCredentials) {
				foreach ($pc in $linkedAppReg.PasswordCredentials) {
					$allCreds += [pscustomobject]@{ Type = 'Secret'; EndDateTime = $pc.EndDateTime; Source = 'AppReg' }
				}
			}
			if ($linkedAppReg.KeyCredentials) {
				foreach ($kc in $linkedAppReg.KeyCredentials) {
					$allCreds += [pscustomobject]@{ Type = 'Certificate'; EndDateTime = $kc.EndDateTime; Source = 'AppReg' }
				}
			}
		}

		$secretCount = @($allCreds | Where-Object { $_.Type -eq 'Secret' }).Count
		$certCount = @($allCreds | Where-Object { $_.Type -eq 'Certificate' }).Count

		$earliestExpiry = $null
		$daysUntilExpiry = $null
		$expiredCount = 0
		$expiringSoonCount = 0

		foreach ($cred in $allCreds) {
			$expDt = if ($cred.EndDateTime) { [datetime]$cred.EndDateTime } else { $null }
			if ($expDt) {
				if ($null -eq $earliestExpiry -or $expDt -lt $earliestExpiry) {
					$earliestExpiry = $expDt
				}
				if ($expDt -lt $now) {
					$expiredCount++
				} elseif ($expDt -lt $expiringThresholdDate) {
					$expiringSoonCount++
				}
			}
		}

		if ($earliestExpiry) {
			$daysUntilExpiry = [int]($earliestExpiry - $now).TotalDays
		}

		if ($allCreds.Count -eq 0) {
			$credentialStatus = "No Credentials"
		} elseif ($expiredCount -eq $allCreds.Count) {
			$credentialStatus = "Critical"
		} elseif ($expiredCount -gt 0) {
			$credentialStatus = "Warning"
		} elseif ($expiringSoonCount -gt 0) {
			$credentialStatus = "Expiring Soon"
		} else {
			$credentialStatus = "Healthy"
		}

		# --- Granted Graph permissions (actually consented, not just configured) ---
		$grantedPerms = if ($grantedPermsLookup.ContainsKey($sp.Id)) { $grantedPermsLookup[$sp.Id] } else { @() }
		$highPrivPerms = @($grantedPerms | Where-Object { $_ -in $highPrivilegePermissions })
		$isHighPrivilege = $highPrivPerms.Count -gt 0

		# --- Sign-in activity ---
		$lastSignIn = $null
		$daysSinceActivity = $null
		if ($sp.SignInActivity) {
			$lastSignIn = $sp.SignInActivity.LastSignInDateTime
			if (-not $lastSignIn) {
				$lastSignIn = $sp.SignInActivity.LastNonInteractiveSignInDateTime
			}
		}
		if ($lastSignIn) {
			$daysSinceActivity = [int]($now - ([datetime]$lastSignIn)).TotalDays
		}

		# --- Status (Disabled > Inactive > Active) ---
		if (-not $sp.AccountEnabled) {
			$status = "Disabled"
		} elseif ($lastSignIn -and ([datetime]$lastSignIn) -ge $cutoffDate) {
			$status = "Active"
		} else {
			$status = "Inactive"
		}

		# --- SP Type friendly name ---
		$spTypeName = switch ($sp.ServicePrincipalType) {
			'Application' { 'Application' }
			'ManagedIdentity' { 'Managed Identity' }
			'Legacy' { 'Legacy' }
			'SocialIdp' { 'Social IdP' }
			default { if ($sp.ServicePrincipalType) { $sp.ServicePrincipalType } else { 'Unknown' } }
		}

		[pscustomobject]@{
			DisplayName             = $sp.DisplayName
			AppId                   = $sp.AppId
			ObjectId                = $sp.Id
			ServicePrincipalType    = $spTypeName
			AccountEnabled          = $sp.AccountEnabled
			IsMicrosoft             = $isMicrosoft
			HasAppRegistration      = $hasAppReg
			CreatedDateTime         = $sp.CreatedDateTime
			SecretCount             = $secretCount
			CertificateCount        = $certCount
			EarliestExpiry          = $earliestExpiry
			DaysUntilExpiry         = $daysUntilExpiry
			CredentialStatus        = $credentialStatus
			GrantedPermissionCount  = $grantedPerms.Count
			IsHighPrivilege         = $isHighPrivilege
			HighPrivPermissions     = ($highPrivPerms -join ", ")
			AllGrantedPermissions   = ($grantedPerms -join ", ")
			LastSignIn              = $lastSignIn
			DaysSinceActivity       = $daysSinceActivity
			Status                  = $status
		}
	}

	# --- Stats ---
	$totalApps       = @($report).Count
	$totalEnabled    = @($report | Where-Object { $_.AccountEnabled }).Count
	$totalDisabled   = @($report | Where-Object { -not $_.AccountEnabled }).Count
	$totalActive     = @($report | Where-Object { $_.Status -eq "Active" }).Count
	$totalInactive   = @($report | Where-Object { $_.Status -eq "Inactive" }).Count
	$totalHighPriv   = @($report | Where-Object { $_.IsHighPrivilege }).Count
	$totalMicrosoft  = @($report | Where-Object { $_.IsMicrosoft }).Count
	$totalWithAppReg = @($report | Where-Object { $_.HasAppRegistration }).Count
	$totalCritical   = @($report | Where-Object { $_.CredentialStatus -eq "Critical" }).Count
	$totalWarning    = @($report | Where-Object { $_.CredentialStatus -eq "Warning" }).Count
	$totalExpSoon    = @($report | Where-Object { $_.CredentialStatus -eq "Expiring Soon" }).Count
	$totalHealthy    = @($report | Where-Object { $_.CredentialStatus -eq "Healthy" }).Count
	$totalNoCreds    = @($report | Where-Object { $_.CredentialStatus -eq "No Credentials" }).Count

	$spTypeSummary = $report | Group-Object ServicePrincipalType | Sort-Object Count -Descending | ForEach-Object {
		[pscustomobject]@{ Type = $_.Name; Count = $_.Count }
	}

	# --- File paths ---
	if (-not $ReportPath) { $ReportPath = (Get-Location).Path }
	$reportFolder = if (Test-Path $ReportPath -PathType Container) { $ReportPath } else { Split-Path -Parent $ReportPath }
	if ($reportFolder -and -not (Test-Path $reportFolder)) {
		New-Item -ItemType Directory -Path $reportFolder -Force | Out-Null
	}

	$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
	$csvFile = if (Test-Path $ReportPath -PathType Container) {
		Join-Path $ReportPath ("EntraIDApps_{0}.csv" -f $timestamp)
	} else { $ReportPath }

	# --- CSV export ---
	$report | Sort-Object DisplayName | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8

	# --- HTML report ---
	$reportDate = Get-Date -Format "dd MMM yyyy HH:mm"

	# Build per-app JSON for client-side threshold recalculation
	$appsJson = ($report | Sort-Object DisplayName | ForEach-Object {
		$daysVal = if ($null -ne $_.DaysSinceActivity) { $_.DaysSinceActivity } else { -1 }
		$hp = if ($_.IsHighPrivilege) { "true" } else { "false" }
		$ms = if ($_.IsMicrosoft) { "true" } else { "false" }
		$en = if ($_.AccountEnabled) { "true" } else { "false" }
		$ar = if ($_.HasAppRegistration) { "true" } else { "false" }
		$cs = ($_.CredentialStatus) -replace '"', '\"'
		$spt = ($_.ServicePrincipalType) -replace '"', '\"'
		'{{"days":{0},"hp":{1},"ms":{2},"en":{3},"ar":{4},"cs":"{5}","spt":"{6}"}}' -f $daysVal, $hp, $ms, $en, $ar, $cs, $spt
	}) -join ","

	# Build table rows
	$tableRows = ($report | Sort-Object DisplayName | ForEach-Object {
		$daysVal = if ($null -ne $_.DaysSinceActivity) { $_.DaysSinceActivity } else { -1 }
		$enabledVal = if ($_.AccountEnabled) { "1" } else { "0" }
		$msVal = if ($_.IsMicrosoft) { "1" } else { "0" }
		$appName = [System.Net.WebUtility]::HtmlEncode($_.DisplayName)
		$appIdEnc = [System.Net.WebUtility]::HtmlEncode($_.AppId)
		$spType = [System.Net.WebUtility]::HtmlEncode($_.ServicePrincipalType)
		$enabled = if ($_.AccountEnabled) { '<span class="badge badge-active">Yes</span>' } else { '<span class="badge badge-disabled">Disabled</span>' }
		$msBadge = if ($_.IsMicrosoft) { '<span class="badge badge-ms">Microsoft</span>' } else { '<span class="badge badge-thirdparty">3rd Party</span>' }
		$arBadge = if ($_.HasAppRegistration) { '<span class="badge badge-active">Yes</span>' } else { '<span class="badge badge-nocreds">No</span>' }
		$created = if ($_.CreatedDateTime) { ([datetime]$_.CreatedDateTime).ToString("dd MMM yyyy") } else { "-" }
		$creds = if ($_.SecretCount -eq 0 -and $_.CertificateCount -eq 0) { "None" } else { "{0}S / {1}C" -f $_.SecretCount, $_.CertificateCount }
		$expiry = if ($_.EarliestExpiry) { ([datetime]$_.EarliestExpiry).ToString("dd MMM yyyy") } else { "-" }
		$daysUntil = if ($null -ne $_.DaysUntilExpiry) { "$($_.DaysUntilExpiry) days" } else { "-" }
		$credStatusClass = switch ($_.CredentialStatus) {
			"Healthy" { "healthy" } "Expiring Soon" { "expiring" } "Warning" { "warning" } "Critical" { "critical" } "No Credentials" { "nocreds" }
		}
		$highPrivBadge = if ($_.IsHighPrivilege) { '<span class="badge badge-highpriv">Yes</span>' } else { '<span class="badge badge-lowpriv">No</span>' }
		$lastSignIn = if ($_.LastSignIn) { ([datetime]$_.LastSignIn).ToString("dd MMM yyyy") } else { "-" }
		$sinceActivity = if ($null -ne $_.DaysSinceActivity) { "$($_.DaysSinceActivity) days" } else { "Never" }
		$statusClass = switch ($_.Status) { "Active" { "active" } "Inactive" { "inactive" } "Disabled" { "disabled" } }

		"<tr data-days=`"$daysVal`" data-ena=`"$enabledVal`" data-ms=`"$msVal`"><td>$appName</td><td class=`"app-id`">$appIdEnc</td><td>$spType</td><td>$enabled</td><td>$msBadge</td><td>$arBadge</td><td>$created</td><td>$creds</td><td>$expiry</td><td class=`"cred-days`">$daysUntil</td><td><span class=`"badge badge-$credStatusClass`">$($_.CredentialStatus)</span></td><td>$($_.GrantedPermissionCount)</td><td>$highPrivBadge</td><td>$lastSignIn</td><td>$sinceActivity</td><td><span class=`"badge badge-$statusClass`">$($_.Status)</span></td></tr>"
	}) -join "`n"

	# Build high-privilege apps detail rows
	$highPrivApps = $report | Where-Object { $_.IsHighPrivilege } | Sort-Object DisplayName
	$highPrivRows = if ($highPrivApps) {
		($highPrivApps | ForEach-Object {
			$appName = [System.Net.WebUtility]::HtmlEncode($_.DisplayName)
			$perms = [System.Net.WebUtility]::HtmlEncode($_.HighPrivPermissions)
			$msBadge = if ($_.IsMicrosoft) { '<span class="badge badge-ms">Microsoft</span>' } else { '<span class="badge badge-thirdparty">3rd Party</span>' }
			$credBadgeClass = switch ($_.CredentialStatus) { "Healthy" { "healthy" } "Expiring Soon" { "expiring" } "Warning" { "warning" } "Critical" { "critical" } "No Credentials" { "nocreds" } }
			$statusClass = switch ($_.Status) { "Active" { "active" } "Inactive" { "inactive" } "Disabled" { "disabled" } }
			"<tr><td>$appName</td><td class=`"app-id`">$([System.Net.WebUtility]::HtmlEncode($_.AppId))</td><td>$msBadge</td><td class=`"perm-list`">$perms</td><td><span class=`"badge badge-$credBadgeClass`">$($_.CredentialStatus)</span></td><td><span class=`"badge badge-$statusClass`">$($_.Status)</span></td></tr>"
		}) -join "`n"
	} else {
		"<tr><td colspan=`"6`" style=`"text-align:center;color:#999;`">No highly privileged applications found</td></tr>"
	}

	$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Entra ID Enterprise Applications Report</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; padding: 30px; }
  .header { background: linear-gradient(135deg, #1a1a2e, #16213e); color: #fff; padding: 30px 40px; border-radius: 12px; margin-bottom: 30px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 16px; }
  .header-left h1 { font-size: 1.6em; margin-bottom: 6px; }
  .header-left p { font-size: 0.9em; opacity: 0.8; }
  .header-right { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
  .header-right label { font-size: 0.9em; opacity: 0.85; }
  .header-right select { padding: 8px 14px; border: none; border-radius: 6px; font-size: 0.95em; font-weight: 600; background: rgba(255,255,255,0.15); color: #fff; cursor: pointer; }
  .header-right select option { color: #333; background: #fff; }
  .toggle-label { display: flex; align-items: center; gap: 6px; cursor: pointer; font-size: 0.9em; opacity: 0.85; }
  .toggle-label input[type="checkbox"] { width: 16px; height: 16px; cursor: pointer; }

  .section-title { font-size: 1.15em; font-weight: 600; color: #1a1a2e; margin-bottom: 16px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0; }

  .summary-cards { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .card { background: #fff; border-radius: 10px; padding: 24px 30px; flex: 1; min-width: 160px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
  .card .label { font-size: 0.85em; color: #777; text-transform: uppercase; letter-spacing: 0.5px; }
  .card .value { font-size: 2em; font-weight: 700; margin-top: 6px; }
  .card .sub { font-size: 0.78em; color: #999; margin-top: 2px; }

  .dist-section { margin-bottom: 30px; }
  .dist-cards { display: flex; gap: 16px; flex-wrap: wrap; }
  .dist-card { background: #fff; border-radius: 10px; padding: 18px 24px; min-width: 140px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); border-left: 4px solid #3498db; text-align: center; }
  .dist-card .dist-label { font-size: 0.82em; color: #555; margin-bottom: 6px; text-transform: uppercase; letter-spacing: 0.3px; }
  .dist-card .dist-value { font-size: 1.6em; font-weight: 700; color: #1a1a2e; }

  .charts-row { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .chart-section { background: #fff; border-radius: 10px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); flex: 1; min-width: 340px; }
  .chart-section h2 { font-size: 1.1em; margin-bottom: 20px; color: #1a1a2e; }
  .chart-container { max-width: 400px; margin: 0 auto; }

  .table-section { background: #fff; border-radius: 10px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 30px; overflow-x: auto; }
  .table-section h2 { font-size: 1.1em; margin-bottom: 16px; color: #1a1a2e; }
  .table-controls { display: flex; gap: 14px; margin-bottom: 16px; flex-wrap: wrap; align-items: center; }
  .table-controls input[type="text"] { padding: 8px 14px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; min-width: 260px; }
  .table-controls select { padding: 8px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; background: #fff; }
  .table-controls .count-label { font-size: 0.85em; color: #777; margin-left: auto; }

  table { width: 100%; border-collapse: collapse; font-size: 0.84em; }
  th { background: #1a1a2e; color: #fff; padding: 12px 14px; text-align: left; position: sticky; top: 0; cursor: pointer; user-select: none; white-space: nowrap; }
  th:hover { background: #2c3e50; }
  td { padding: 10px 14px; border-bottom: 1px solid #eee; white-space: nowrap; }
  tr:hover td { background: #f8f9fa; }
  tr.hidden-row { display: none; }
  .app-id { font-family: 'Consolas', monospace; font-size: 0.82em; color: #666; }
  .perm-list { white-space: normal; max-width: 400px; font-size: 0.82em; color: #c0392b; }

  .badge { padding: 3px 10px; border-radius: 12px; font-size: 0.82em; font-weight: 600; }
  .badge-active { background: #d4edda; color: #155724; }
  .badge-inactive { background: #f8d7da; color: #721c24; }
  .badge-disabled { background: #e2e3e5; color: #495057; }
  .badge-healthy { background: #d4edda; color: #155724; }
  .badge-expiring { background: #fff3cd; color: #856404; }
  .badge-warning { background: #ffe0b2; color: #e65100; }
  .badge-critical { background: #f8d7da; color: #721c24; }
  .badge-nocreds { background: #e2e3e5; color: #495057; }
  .badge-highpriv { background: #f8d7da; color: #721c24; }
  .badge-lowpriv { background: #e2e3e5; color: #6c757d; }
  .badge-ms { background: #d1ecf1; color: #0c5460; }
  .badge-thirdparty { background: #e8daef; color: #6c3483; }

  .activity-green { color: #155724; font-weight: 600; }
  .activity-amber { color: #856404; font-weight: 600; }
  .activity-red { color: #c0392b; font-weight: 600; }
  .activity-brightred { color: #e74c3c; font-weight: 700; }
  .activity-never { color: #6c757d; font-weight: 600; }

  .cred-green { color: #155724; font-weight: 600; }
  .cred-amber { color: #856404; font-weight: 600; }
  .cred-red { color: #c0392b; font-weight: 600; }
  .cred-expired { color: #e74c3c; font-weight: 700; }
  .cred-none { color: #6c757d; }

  .highpriv-table { margin-top: 12px; }
  .highpriv-table th { font-size: 0.84em; padding: 10px 12px; }
  .highpriv-table td { font-size: 0.84em; padding: 8px 12px; }

  .footer { text-align: center; font-size: 0.8em; color: #999; margin-top: 20px; }
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <h1>Entra ID Enterprise Applications Report</h1>
    <p>Tenant: $([System.Net.WebUtility]::HtmlEncode($tenantDisplayName)) ($tenantId) &nbsp;|&nbsp; Generated: $reportDate &nbsp;|&nbsp; Total: $totalApps ($totalMicrosoft Microsoft, $($totalApps - $totalMicrosoft) third-party)</p>
  </div>
  <div class="header-right">
    <label class="toggle-label"><input type="checkbox" id="hideMsApps" onchange="applyThreshold()"> Hide Microsoft Apps</label>
    <label for="thresholdSelect">Inactive Threshold:</label>
    <select id="thresholdSelect" onchange="applyThreshold()">
      <option value="30" $(if ($InactiveDays -eq 30) { 'selected' })>30 Days</option>
      <option value="60" $(if ($InactiveDays -eq 60) { 'selected' })>60 Days</option>
      <option value="90" $(if ($InactiveDays -eq 90) { 'selected' })>90 Days</option>
      <option value="180" $(if ($InactiveDays -eq 180 -or ($InactiveDays -ne 30 -and $InactiveDays -ne 60 -and $InactiveDays -ne 90 -and $InactiveDays -ne 360)) { 'selected' })>180 Days</option>
      <option value="360" $(if ($InactiveDays -eq 360) { 'selected' })>360 Days</option>
    </select>
  </div>
</div>

<!-- OVERVIEW -->
<div class="section-title">Overview</div>
<div class="summary-cards">
  <div class="card"><div class="label">Total Enterprise Apps</div><div class="value" style="color:#1a1a2e;" id="cardTotal">-</div><div class="sub" id="cardTotalSub"></div></div>
  <div class="card"><div class="label">Active</div><div class="value" style="color:#27ae60;" id="cardActive">-</div><div class="sub" id="cardActivePct"></div></div>
  <div class="card"><div class="label">Inactive</div><div class="value" style="color:#e74c3c;" id="cardInactive">-</div><div class="sub" id="cardInactivePct"></div></div>
  <div class="card"><div class="label">Disabled</div><div class="value" style="color:#6c757d;" id="cardDisabled">-</div><div class="sub" id="cardDisabledPct"></div></div>
  <div class="card"><div class="label">With App Registration</div><div class="value" style="color:#3498db;" id="cardAppReg">-</div></div>
  <div class="card"><div class="label">High Privilege</div><div class="value" style="color:#c0392b;" id="cardHighPriv">-</div></div>
</div>

<!-- CREDENTIAL HEALTH -->
<div class="dist-section">
  <div class="section-title">Credential Health</div>
  <div class="dist-cards" id="credCards"></div>
</div>

<!-- CHARTS -->
<div class="charts-row">
  <div class="chart-section">
    <h2>Credential Status Distribution</h2>
    <div class="chart-container"><canvas id="credChart"></canvas></div>
  </div>
  <div class="chart-section">
    <h2>Application Activity</h2>
    <div class="chart-container"><canvas id="activityChart"></canvas></div>
  </div>
  <div class="chart-section">
    <h2>Application Types</h2>
    <div class="chart-container"><canvas id="typeChart"></canvas></div>
  </div>
</div>

<!-- HIGHLY PRIVILEGED APPS -->
<div class="table-section">
  <h2>Highly Privileged Applications — Granted Graph Permissions ($totalHighPriv)</h2>
  <p style="font-size:0.88em;color:#777;margin-bottom:12px;">Applications with high-privilege Microsoft Graph permissions actually granted (admin consented) — not just configured</p>
  <table class="highpriv-table">
    <thead><tr>
      <th>Application Name</th>
      <th>App (Client) ID</th>
      <th>Owner</th>
      <th>High Privilege Permissions (Granted)</th>
      <th>Credential Status</th>
      <th>Activity Status</th>
    </tr></thead>
    <tbody>
$highPrivRows
    </tbody>
  </table>
</div>

<!-- FULL TABLE -->
<div class="table-section">
  <h2>All Enterprise Application Details</h2>
  <div class="table-controls">
    <input type="text" id="searchBox" placeholder="Search by name, app ID, type..." onkeyup="filterTable()" />
    <select id="statusFilter" onchange="filterTable()">
      <option value="all">All Status</option>
      <option value="active">Active Only</option>
      <option value="inactive">Inactive Only</option>
      <option value="disabled">Disabled Only</option>
    </select>
    <select id="credFilter" onchange="filterTable()">
      <option value="all">All Credential Status</option>
      <option value="critical">Critical</option>
      <option value="warning">Warning</option>
      <option value="expiring">Expiring Soon</option>
      <option value="healthy">Healthy</option>
      <option value="nocreds">No Credentials</option>
    </select>
    <select id="privFilter" onchange="filterTable()">
      <option value="all">All Privilege</option>
      <option value="high">High Privilege Only</option>
      <option value="standard">Standard Only</option>
    </select>
    <select id="ownerFilter" onchange="filterTable()">
      <option value="all">All Owners</option>
      <option value="thirdparty">Third-Party Only</option>
      <option value="microsoft">Microsoft Only</option>
    </select>
    <span class="count-label" id="rowCount"></span>
  </div>
  <table id="appTable">
    <thead><tr>
      <th onclick="sortTable(0)">App Name</th>
      <th onclick="sortTable(1)">App (Client) ID</th>
      <th onclick="sortTable(2)">Type</th>
      <th onclick="sortTable(3)">Enabled</th>
      <th onclick="sortTable(4)">Owner</th>
      <th onclick="sortTable(5)">App Reg</th>
      <th onclick="sortTable(6)">Created</th>
      <th onclick="sortTable(7)">Secrets / Certs</th>
      <th onclick="sortTable(8)">Earliest Expiry</th>
      <th onclick="sortTable(9)">Days Until Expiry</th>
      <th onclick="sortTable(10)">Credential Status</th>
      <th onclick="sortTable(11)">Granted Perms</th>
      <th onclick="sortTable(12)">High Privilege</th>
      <th onclick="sortTable(13)">Last Sign-in</th>
      <th onclick="sortTable(14)">Days Since Activity</th>
      <th onclick="sortTable(15)">Status</th>
    </tr></thead>
    <tbody>
$tableRows
    </tbody>
  </table>
</div>

<div class="footer">Report generated by ReportEntraIDApps.ps1</div>

<script>
var appData = [$appsJson];
var chartColors = ['#3498db','#27ae60','#e74c3c','#f39c12','#9b59b6','#1abc9c','#e67e22','#2c3e50','#95a5a6','#d35400'];
var credColorMap = { 'Healthy':'#27ae60', 'Expiring Soon':'#f39c12', 'Warning':'#e67e22', 'Critical':'#e74c3c', 'No Credentials':'#95a5a6' };

var chartOpts = function(pos) {
  return { responsive: true, plugins: { legend: { position: pos || 'right', labels: { padding: 14, font: { size: 12 }, boxWidth: 14 } }, tooltip: { callbacks: { label: function(ctx) { var t = ctx.dataset.data.reduce(function(a,b){return a+b},0); return ctx.label+': '+ctx.parsed+' ('+(t>0?((ctx.parsed/t)*100).toFixed(1):0)+'%)'; } } } } };
};

var credChart = new Chart(document.getElementById('credChart'), { type:'doughnut', data:{ labels:[], datasets:[{ data:[], backgroundColor:[], borderWidth:2, borderColor:'#fff' }] }, options: chartOpts() });
var activityChart = new Chart(document.getElementById('activityChart'), { type:'doughnut', data:{ labels:[], datasets:[{ data:[], backgroundColor:[], borderWidth:2, borderColor:'#fff' }] }, options: chartOpts() });
var typeChart = new Chart(document.getElementById('typeChart'), { type:'doughnut', data:{ labels:[], datasets:[{ data:[], backgroundColor:[], borderWidth:2, borderColor:'#fff' }] }, options: chartOpts() });

function pct(n, total) { return total > 0 ? ((n / total) * 100).toFixed(1) : '0.0'; }

function applyThreshold() {
  var threshold = parseInt(document.getElementById('thresholdSelect').value);
  var hideMs = document.getElementById('hideMsApps').checked;

  var filtered = hideMs ? appData.filter(function(d) { return !d.ms; }) : appData;
  var total = filtered.length;
  var active = 0, inactive = 0, disabled = 0, highPriv = 0, withAppReg = 0;
  var credCounts = {};
  var typeCounts = {};

  for (var i = 0; i < filtered.length; i++) {
    var d = filtered[i];
    if (!d.en) {
      disabled++;
    } else {
      var isInactive = (d.days === -1) || (d.days >= threshold);
      if (isInactive) { inactive++; } else { active++; }
    }
    if (d.hp) highPriv++;
    if (d.ar) withAppReg++;
    credCounts[d.cs] = (credCounts[d.cs] || 0) + 1;
    typeCounts[d.spt] = (typeCounts[d.spt] || 0) + 1;
  }

  var enabled = active + inactive;

  document.getElementById('cardTotal').textContent = total;
  document.getElementById('cardTotalSub').textContent = hideMs ? 'Microsoft apps hidden' : 'All apps';
  document.getElementById('cardActive').textContent = active;
  document.getElementById('cardActivePct').textContent = pct(active, enabled) + '% of enabled';
  document.getElementById('cardInactive').textContent = inactive;
  document.getElementById('cardInactivePct').textContent = pct(inactive, enabled) + '% of enabled';
  document.getElementById('cardDisabled').textContent = disabled;
  document.getElementById('cardDisabledPct').textContent = pct(disabled, total) + '% of total';
  document.getElementById('cardAppReg').textContent = withAppReg;
  document.getElementById('cardHighPriv').textContent = highPriv;

  // Credential cards
  var credContainer = document.getElementById('credCards');
  credContainer.innerHTML = '';
  var credOrder = ['Critical','Warning','Expiring Soon','Healthy','No Credentials'];
  var borderColors = { 'Critical':'#e74c3c', 'Warning':'#e67e22', 'Expiring Soon':'#f39c12', 'Healthy':'#27ae60', 'No Credentials':'#95a5a6' };
  credOrder.forEach(function(key) {
    if (credCounts[key]) {
      var card = document.createElement('div');
      card.className = 'dist-card';
      card.style.borderLeftColor = borderColors[key] || '#3498db';
      card.innerHTML = '<div class="dist-label">' + key + '</div><div class="dist-value">' + credCounts[key] + '</div>';
      credContainer.appendChild(card);
    }
  });

  // Credential chart
  var cLabels = [], cData = [], cColors = [];
  credOrder.forEach(function(key) {
    if (credCounts[key]) { cLabels.push(key); cData.push(credCounts[key]); cColors.push(credColorMap[key] || '#95a5a6'); }
  });
  credChart.data.labels = cLabels;
  credChart.data.datasets[0].data = cData;
  credChart.data.datasets[0].backgroundColor = cColors;
  credChart.update();

  // Activity chart
  activityChart.data.labels = ['Active', 'Inactive', 'Disabled'];
  activityChart.data.datasets[0].data = [active, inactive, disabled];
  activityChart.data.datasets[0].backgroundColor = ['#27ae60', '#e74c3c', '#95a5a6'];
  activityChart.update();

  // Type chart
  var tKeys = Object.keys(typeCounts).sort(function(a,b){ return typeCounts[b] - typeCounts[a]; });
  typeChart.data.labels = tKeys;
  typeChart.data.datasets[0].data = tKeys.map(function(k){ return typeCounts[k]; });
  typeChart.data.datasets[0].backgroundColor = tKeys.map(function(_,i){ return chartColors[i % chartColors.length]; });
  typeChart.update();

  // Update table row statuses and color-code columns
  var rows = document.querySelectorAll('#appTable tbody tr');
  for (var j = 0; j < rows.length; j++) {
    var days = parseInt(rows[j].getAttribute('data-days'));
    var rowEna = rows[j].getAttribute('data-ena');
    var badge = rows[j].cells[15].querySelector('.badge');
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
    // Color-code Days Since Activity (cell 14)
    var activityCell = rows[j].cells[14];
    if (days === -1) {
      activityCell.className = 'activity-never';
    } else if (days < 30) {
      activityCell.className = 'activity-green';
    } else if (days < 60) {
      activityCell.className = 'activity-amber';
    } else if (days < 180) {
      activityCell.className = 'activity-red';
    } else {
      activityCell.className = 'activity-brightred';
    }
    // Color-code Days Until Expiry (cell 9)
    var credDaysCell = rows[j].cells[9];
    var credDaysText = credDaysCell.textContent.trim();
    if (credDaysText === '-') {
      credDaysCell.className = 'cred-none';
    } else {
      var credDays = parseInt(credDaysText);
      if (credDays < 0) {
        credDaysCell.className = 'cred-expired';
      } else if (credDays <= 30) {
        credDaysCell.className = 'cred-red';
      } else if (credDays <= 90) {
        credDaysCell.className = 'cred-amber';
      } else {
        credDaysCell.className = 'cred-green';
      }
    }
  }

  filterTable();
}

function filterTable() {
  var search = document.getElementById('searchBox').value.toLowerCase();
  var status = document.getElementById('statusFilter').value;
  var cred = document.getElementById('credFilter').value;
  var priv = document.getElementById('privFilter').value;
  var owner = document.getElementById('ownerFilter').value;
  var hideMs = document.getElementById('hideMsApps').checked;
  var rows = document.querySelectorAll('#appTable tbody tr');
  var visible = 0;
  rows.forEach(function(row) {
    var text = row.textContent.toLowerCase();
    var matchSearch = !search || text.indexOf(search) !== -1;
    var rowStatus = row.getAttribute('data-status');
    var matchStatus = status === 'all' || rowStatus === status;
    var credBadge = row.cells[10].querySelector('.badge');
    var credStatus = credBadge ? credBadge.textContent.trim().toLowerCase() : '';
    var credMap = { 'critical':'critical', 'warning':'warning', 'expiring soon':'expiring', 'healthy':'healthy', 'no credentials':'nocreds' };
    var matchCred = cred === 'all' || credMap[credStatus] === cred;
    var privBadge = row.cells[12].querySelector('.badge');
    var privText = privBadge ? privBadge.textContent.trim().toLowerCase() : '';
    var matchPriv = priv === 'all' || (priv === 'high' && privText === 'yes') || (priv === 'standard' && privText === 'no');
    var isMs = row.getAttribute('data-ms') === '1';
    var matchOwner = owner === 'all' || (owner === 'microsoft' && isMs) || (owner === 'thirdparty' && !isMs);
    var matchHideMs = !hideMs || !isMs;
    if (matchSearch && matchStatus && matchCred && matchPriv && matchOwner && matchHideMs) {
      row.classList.remove('hidden-row');
      visible++;
    } else {
      row.classList.add('hidden-row');
    }
  });
  document.getElementById('rowCount').textContent = 'Showing ' + visible + ' of ' + rows.length + ' applications';
}

var sortDir = {};
function sortTable(col) {
  var tbody = document.getElementById('appTable').querySelector('tbody');
  var rows = Array.from(tbody.querySelectorAll('tr'));
  var dir = sortDir[col] === 'asc' ? 'desc' : 'asc';
  sortDir[col] = dir;
  rows.sort(function(a, b) {
    var av = a.cells[col].textContent.trim().toLowerCase();
    var bv = b.cells[col].textContent.trim().toLowerCase();
    var an = parseFloat(av), bn = parseFloat(bv);
    if (!isNaN(an) && !isNaN(bn)) {
      return dir === 'asc' ? an - bn : bn - an;
    }
    if (av < bv) return dir === 'asc' ? -1 : 1;
    if (av > bv) return dir === 'asc' ? 1 : -1;
    return 0;
  });
  rows.forEach(function(row) { tbody.appendChild(row); });
}

// Initial render
applyThreshold();
</script>
</body>
</html>
"@

	$htmlReportFile = Join-Path $reportFolder ("EntraIDApps_{0}.html" -f $timestamp)
	$html | Out-File -FilePath $htmlReportFile -Encoding UTF8

	# --- Console summary ---
	Write-Host ""
	Write-Host "Entra ID Enterprise Applications Report" -ForegroundColor Cyan
	Write-Host "--------------------------------------------"
	Write-Host ("Tenant                   : {0} ({1})" -f $tenantDisplayName, $tenantId)
	Write-Host ("Total enterprise apps    : {0}" -f $totalApps)
	Write-Host ("  Microsoft first-party  : {0}" -f $totalMicrosoft) -ForegroundColor DarkGray
	Write-Host ("  Third-party / custom   : {0}" -f ($totalApps - $totalMicrosoft))
	Write-Host ("  With app registration  : {0}" -f $totalWithAppReg)
	Write-Host ""
	Write-Host "Activity" -ForegroundColor Cyan
	Write-Host ("  Active                 : {0}" -f $totalActive) -ForegroundColor Green
	Write-Host ("  Inactive               : {0}" -f $totalInactive) -ForegroundColor Red
	Write-Host ("  Disabled               : {0}" -f $totalDisabled) -ForegroundColor DarkGray
	Write-Host ""
	Write-Host "Application Types" -ForegroundColor Cyan
	foreach ($spt in $spTypeSummary) {
		Write-Host ("  {0,-25}: {1}" -f $spt.Type, $spt.Count)
	}
	Write-Host ""
	Write-Host "Credential Health" -ForegroundColor Cyan
	Write-Host ("  Healthy                : {0}" -f $totalHealthy) -ForegroundColor Green
	Write-Host ("  Expiring Soon (30d)    : {0}" -f $totalExpSoon) -ForegroundColor Yellow
	Write-Host ("  Warning (some expired) : {0}" -f $totalWarning) -ForegroundColor DarkYellow
	Write-Host ("  Critical (all expired) : {0}" -f $totalCritical) -ForegroundColor Red
	Write-Host ("  No Credentials         : {0}" -f $totalNoCreds) -ForegroundColor DarkGray
	Write-Host ""
	Write-Host "High Privilege Applications (Granted Graph Permissions)" -ForegroundColor Cyan
	Write-Host ("  Count                  : {0}" -f $totalHighPriv) -ForegroundColor Red
	if ($highPrivApps -and $highPrivApps.Count -gt 0) {
		foreach ($hp in $highPrivApps) {
			Write-Host ("  - {0}" -f $hp.DisplayName)
			Write-Host ("    Granted: {0}" -f $hp.HighPrivPermissions) -ForegroundColor DarkYellow
		}
	}
	Write-Host ""
	Write-Host ("Inactive days threshold  : {0}" -f $InactiveDays)
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
