#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Applications, Microsoft.Graph.Identity.DirectoryManagement

<#
.SYNOPSIS
    Reports on enterprise applications (service principals) in Entra ID, including
    sign-in activity and credential expiry.

.DESCRIPTION
    Connects to Microsoft Graph and retrieves all service principals (enterprise applications).
    Cross-references app registrations to identify which service principals have a local
    app registration. Resolves Microsoft Graph permission names for each application.
    Exports results to CSV and HTML.
    Supports an inactivity threshold to flag applications that have not been used recently.

.PARAMETER ReportPath
    Folder or file path for the output report. If a folder is specified, a timestamped
    filename is generated automatically. Defaults to the current directory.

.PARAMETER InactiveDays
    Number of days since last sign-in to consider an application inactive. Defaults to 180.

.EXAMPLE
    .\ReportEntraIDApps.ps1

.EXAMPLE
    .\ReportEntraIDApps.ps1 -InactiveDays 90
#>

[CmdletBinding()]
param(
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[string]$ReportPath,

	[Parameter(Mandatory = $false)]
	[ValidateRange(1, 3650)]
	[int]$InactiveDays = 180
)

$ErrorActionPreference = 'Stop'

$S_RequiredGraphScopes = @(
	'Application.Read.All'
	'AuditLog.Read.All'
	'Organization.Read.All'
)

$S_GraphRequestDelayMilliseconds = 5

try
{
	# --- Module check ---
	$S_RequiredModules = @('Microsoft.Graph.Applications', 'Microsoft.Graph.Identity.DirectoryManagement')
	foreach ($S_Mod in $S_RequiredModules)
	{
		if (-not (Get-Module -ListAvailable -Name $S_Mod))
		{
			throw "$S_Mod module is not installed. Install it using Install-Module Microsoft.Graph -Scope CurrentUser."
		}
	}
	Import-Module Microsoft.Graph.Applications -ErrorAction Stop
	Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

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
			Connect-MgGraph -Scopes $S_RequiredGraphScopes -ErrorAction Stop | Out-Null
		}
	}
	else
	{
		Connect-MgGraph -Scopes $S_RequiredGraphScopes -ErrorAction Stop | Out-Null
	}
	$S_ExistingContext = Get-MgContext

	# --- Tenant info ---
	$S_TenantDisplayName = $null
	try
	{
		$S_Org = Get-MgOrganization -ErrorAction Stop | Select-Object -First 1
		$S_TenantDisplayName = $S_Org.DisplayName
	}
	catch
	{
	}
	if (-not $S_TenantDisplayName)
	{
		$S_TenantDisplayName = $S_ExistingContext.TenantId
	}
	if ($S_ExistingContext.TenantId)
	{
		$S_TenantId = $S_ExistingContext.TenantId
	}
	else
	{
		$S_TenantId = 'Unknown'
	}

	# --- Fetch all Service Principals (Enterprise Applications) ---
	Write-Host "Fetching all enterprise applications (service principals)..." -ForegroundColor Cyan
	$S_ServicePrincipals = Get-MgServicePrincipal -All `
		-Property "id,appId,displayName,servicePrincipalType,accountEnabled,appOwnerOrganizationId,signInAudience,createdDateTime,passwordCredentials,keyCredentials,signInActivity,tags,notes" `
		-ErrorAction Stop
	Write-Host "  Found $($S_ServicePrincipals.Count) service principals" -ForegroundColor Green

	# --- Fetch App Registrations to identify which SPs have a local app reg ---
	Write-Host "Fetching app registrations for cross-reference..." -ForegroundColor Cyan
	$S_AppRegLookup = @{}
	try
	{
		$S_AppRegistrations = Get-MgApplication -All `
			-Property "id,appId,displayName,passwordCredentials,keyCredentials" `
			-ErrorAction Stop
		foreach ($S_Ar in $S_AppRegistrations)
		{
			if ($S_Ar.AppId)
			{
				$S_AppRegLookup[$S_Ar.AppId] = $S_Ar
			}
		}
		Write-Host "  Found $($S_AppRegistrations.Count) app registrations" -ForegroundColor Green
	}
	catch
	{
		Write-Warning "Could not fetch app registrations. HasAppRegistration column will be unavailable."
	}

	# --- Resolve Microsoft Graph permission names ---
	Write-Host "Resolving Graph permission names..." -ForegroundColor Cyan
	$S_GraphAppId = '00000003-0000-0000-c000-000000000000'
	$S_GraphAppRoles = @{}
	$S_GraphSpnId = $null
	try
	{
		$S_GraphSpn = Get-MgServicePrincipal -Filter "appId eq '$S_GraphAppId'" -Property "id,appRoles" -ErrorAction Stop
		$S_GraphSpnId = $S_GraphSpn.Id
		foreach ($S_Role in $S_GraphSpn.AppRoles)
		{
			$S_GraphAppRoles[$S_Role.Id] = $S_Role.Value
		}
	}
	catch
	{
		Write-Warning "Could not resolve Microsoft Graph permission names."
	}

	# --- Fetch all granted Graph app role assignments (efficient single call) ---
	Write-Host "Fetching granted Microsoft Graph permissions..." -ForegroundColor Cyan
	$S_GrantedPermsLookup = @{}
	if ($S_GraphSpnId)
	{
		try
		{
			$S_GraphAssignments = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $S_GraphSpnId -All -ErrorAction Stop
			foreach ($S_Assignment in $S_GraphAssignments)
			{
				$S_PrincipalId = $S_Assignment.PrincipalId
				$S_RoleName = $S_GraphAppRoles[$S_Assignment.AppRoleId]
				if ($S_RoleName)
				{
					if (-not $S_GrantedPermsLookup.ContainsKey($S_PrincipalId))
					{
						$S_GrantedPermsLookup[$S_PrincipalId] = [System.Collections.Generic.List[string]]::new()
					}
					$S_GrantedPermsLookup[$S_PrincipalId].Add($S_RoleName)
				}
			}
			Write-Host "  Found $($S_GraphAssignments.Count) granted Graph permissions" -ForegroundColor Green
		}
		catch
		{
			Write-Warning "Could not fetch Graph app role assignments. Permission data may be incomplete."
		}
	}

	# --- Fetch enhanced SP sign-in activity (beta report) ---
	# servicePrincipal.signInActivity only covers user-driven sign-ins, so daemons,
	# managed identities and SAML-only apps look perpetually inactive. The dedicated
	# report endpoint exposes delegated + application-auth (client-credential) buckets.
	Write-Host "Fetching service principal sign-in activity report (beta)..." -ForegroundColor Cyan
	$S_SignInLookup = @{}
	try
	{
		$S_Uri = 'https://graph.microsoft.com/beta/reports/servicePrincipalSignInActivities?$top=999'
		do
		{
			$S_Resp = Invoke-MgGraphRequest -Method GET -Uri $S_Uri -OutputType PSObject -ErrorAction Stop
			foreach ($S_Entry in $S_Resp.value)
			{
				if (-not $S_Entry.appId)
				{
					continue
				}

				$S_Delegated = $null
				if ($S_Entry.delegatedClientSignInActivity)
				{
					$S_Delegated = $S_Entry.delegatedClientSignInActivity.lastSignInDateTime
				}
				if (-not $S_Delegated -and $S_Entry.delegatedResourceSignInActivity)
				{
					$S_Delegated = $S_Entry.delegatedResourceSignInActivity.lastSignInDateTime
				}

				$S_AppOnly = $null
				if ($S_Entry.applicationAuthenticationClientSignInActivity)
				{
					$S_AppOnly = $S_Entry.applicationAuthenticationClientSignInActivity.lastSignInDateTime
				}
				if (-not $S_AppOnly -and $S_Entry.applicationAuthenticationResourceSignInActivity)
				{
					$S_AppOnly = $S_Entry.applicationAuthenticationResourceSignInActivity.lastSignInDateTime
				}

				$S_Latest = $null
				if ($S_Entry.lastSignInActivity -and $S_Entry.lastSignInActivity.lastSignInDateTime)
				{
					$S_Latest = $S_Entry.lastSignInActivity.lastSignInDateTime
				}
				else
				{
					foreach ($S_D in @($S_Delegated, $S_AppOnly))
					{
						if ($S_D -and (-not $S_Latest -or [datetime]$S_D -gt [datetime]$S_Latest))
						{
							$S_Latest = $S_D
						}
					}
				}

				if ($S_Delegated -and $S_AppOnly)
				{
					$S_Source = 'Both'
				}
				elseif ($S_Delegated)
				{
					$S_Source = 'Delegated'
				}
				elseif ($S_AppOnly)
				{
					$S_Source = 'AppOnly'
				}
				else
				{
					$S_Source = 'None'
				}

				$S_SignInLookup[$S_Entry.appId] = [pscustomobject]@{
					LastSignIn = $S_Latest
					Delegated  = $S_Delegated
					AppOnly    = $S_AppOnly
					Source     = $S_Source
				}
			}
			$S_Uri = $S_Resp.'@odata.nextLink'
		}
		while ($S_Uri)
		Write-Host "  Found activity for $($S_SignInLookup.Count) service principals" -ForegroundColor Green
	}
	catch
	{
		Write-Warning "Could not fetch servicePrincipalSignInActivities (beta). Falling back to signInActivity only. $($_.Exception.Message)"
	}

	# --- Define high-privilege application permissions ---
	$S_HighPrivilegePermissions = @(
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
	$S_MicrosoftTenantId = 'f8cdef31-a31e-4b4a-93e4-5f571e91255a'

	# --- Build report data ---
	Write-Host "Building report data for $($S_ServicePrincipals.Count) enterprise applications..." -ForegroundColor Cyan
	$S_Now = Get-Date
	$S_CutoffDate = $S_Now.AddDays(-$InactiveDays)
	$S_ExpiringThresholdDate = $S_Now.AddDays(30)

	$S_Report = foreach ($S_Sp in $S_ServicePrincipals)
	{
		# Microsoft first-party detection
		$S_IsMicrosoft = ($S_Sp.AppOwnerOrganizationId -eq $S_MicrosoftTenantId)

		# Has app registration in this tenant?
		$S_HasAppReg = $S_AppRegLookup.ContainsKey($S_Sp.AppId)
		if ($S_HasAppReg)
		{
			$S_LinkedAppReg = $S_AppRegLookup[$S_Sp.AppId]
		}
		else
		{
			$S_LinkedAppReg = $null
		}

		# --- Credentials (from SP + linked App Reg) ---
		$S_AllCreds = @()
		# Credentials on the Service Principal itself (SAML certs, etc.)
		if ($S_Sp.PasswordCredentials)
		{
			foreach ($S_Pc in $S_Sp.PasswordCredentials)
			{
				$S_AllCreds += [pscustomobject]@{ Type = 'Secret'; EndDateTime = $S_Pc.EndDateTime; Source = 'SP' }
			}
		}
		if ($S_Sp.KeyCredentials)
		{
			foreach ($S_Kc in $S_Sp.KeyCredentials)
			{
				$S_AllCreds += [pscustomobject]@{ Type = 'Certificate'; EndDateTime = $S_Kc.EndDateTime; Source = 'SP' }
			}
		}
		# Credentials on the linked App Registration
		if ($S_LinkedAppReg)
		{
			if ($S_LinkedAppReg.PasswordCredentials)
			{
				foreach ($S_Pc in $S_LinkedAppReg.PasswordCredentials)
				{
					$S_AllCreds += [pscustomobject]@{ Type = 'Secret'; EndDateTime = $S_Pc.EndDateTime; Source = 'AppReg' }
				}
			}
			if ($S_LinkedAppReg.KeyCredentials)
			{
				foreach ($S_Kc in $S_LinkedAppReg.KeyCredentials)
				{
					$S_AllCreds += [pscustomobject]@{ Type = 'Certificate'; EndDateTime = $S_Kc.EndDateTime; Source = 'AppReg' }
				}
			}
		}

		$S_SecretCount = @($S_AllCreds | Where-Object { $_.Type -eq 'Secret' }).Count
		$S_CertCount = @($S_AllCreds | Where-Object { $_.Type -eq 'Certificate' }).Count

		$S_EarliestExpiry = $null
		$S_DaysUntilExpiry = $null
		$S_ExpiredCount = 0
		$S_ExpiringSoonCount = 0

		foreach ($S_Cred in $S_AllCreds)
		{
			if ($S_Cred.EndDateTime)
			{
				$S_ExpDt = [datetime]$S_Cred.EndDateTime
			}
			else
			{
				$S_ExpDt = $null
			}
			if ($S_ExpDt)
			{
				if ($null -eq $S_EarliestExpiry -or $S_ExpDt -lt $S_EarliestExpiry)
				{
					$S_EarliestExpiry = $S_ExpDt
				}
				if ($S_ExpDt -lt $S_Now)
				{
					$S_ExpiredCount++
				}
				elseif ($S_ExpDt -lt $S_ExpiringThresholdDate)
				{
					$S_ExpiringSoonCount++
				}
			}
		}

		if ($S_EarliestExpiry)
		{
			$S_DaysUntilExpiry = [int]($S_EarliestExpiry - $S_Now).TotalDays
		}

		if ($S_AllCreds.Count -eq 0)
		{
			$S_CredentialStatus = 'No Credentials'
		}
		elseif ($S_ExpiredCount -eq $S_AllCreds.Count)
		{
			$S_CredentialStatus = 'Critical'
		}
		elseif ($S_ExpiredCount -gt 0)
		{
			$S_CredentialStatus = 'Warning'
		}
		elseif ($S_ExpiringSoonCount -gt 0)
		{
			$S_CredentialStatus = 'Expiring Soon'
		}
		else
		{
			$S_CredentialStatus = 'Healthy'
		}

		# --- Granted Graph permissions (actually consented, not just configured) ---
		if ($S_GrantedPermsLookup.ContainsKey($S_Sp.Id))
		{
			$S_GrantedPerms = $S_GrantedPermsLookup[$S_Sp.Id]
		}
		else
		{
			$S_GrantedPerms = @()
		}
		$S_HighPrivPerms = @($S_GrantedPerms | Where-Object { $_ -in $S_HighPrivilegePermissions })
		$S_IsHighPrivilege = $S_HighPrivPerms.Count -gt 0

		# --- Sign-in activity ---
		# Activity assessment is only meaningful for non-Microsoft apps. Microsoft
		# first-party apps are managed by Microsoft and frequently lack tenant-side
		# sign-in records, so we leave their activity fields empty.
		$S_LastSignIn = $null
		$S_LastDelegated = $null
		$S_LastAppOnly = $null
		if ($S_IsMicrosoft)
		{
			$S_ActivitySource = 'N/A'
		}
		else
		{
			$S_ActivitySource = 'None'
		}
		$S_DaysSinceActivity = $null

		if (-not $S_IsMicrosoft)
		{
			# Primary: dedicated SP sign-in activity report (covers delegated + app-only)
			if ($S_SignInLookup.ContainsKey($S_Sp.AppId))
			{
				$S_Entry = $S_SignInLookup[$S_Sp.AppId]
				$S_LastSignIn = $S_Entry.LastSignIn
				$S_LastDelegated = $S_Entry.Delegated
				$S_LastAppOnly = $S_Entry.AppOnly
				$S_ActivitySource = $S_Entry.Source
			}
			# Fallback: legacy signInActivity property if the report had no entry
			if (-not $S_LastSignIn -and $S_Sp.SignInActivity)
			{
				$S_LastSignIn = $S_Sp.SignInActivity.LastSignInDateTime
				if (-not $S_LastSignIn)
				{
					$S_LastSignIn = $S_Sp.SignInActivity.LastNonInteractiveSignInDateTime
				}
				if ($S_LastSignIn)
				{
					$S_LastDelegated = $S_LastSignIn
					$S_ActivitySource = 'Delegated'
				}
			}
			if ($S_LastSignIn)
			{
				$S_DaysSinceActivity = [int]($S_Now - ([datetime]$S_LastSignIn)).TotalDays
			}
		}

		# --- Status (Disabled > Active > Inactive) ---
		# Microsoft first-party apps are not assessed for activity; treat enabled MS
		# apps as Active so they don't pollute the inactive bucket.
		if (-not $S_Sp.AccountEnabled)
		{
			$S_Status = 'Disabled'
		}
		elseif ($S_IsMicrosoft)
		{
			$S_Status = 'Active'
		}
		elseif ($S_LastSignIn -and ([datetime]$S_LastSignIn) -ge $S_CutoffDate)
		{
			$S_Status = 'Active'
		}
		else
		{
			$S_Status = 'Inactive'
		}

		# --- SP Type friendly name ---
		$S_SpTypeName = switch ($S_Sp.ServicePrincipalType)
		{
			'Application' { 'Application' }
			'ManagedIdentity' { 'Managed Identity' }
			'Legacy' { 'Legacy' }
			'SocialIdp' { 'Social IdP' }
			default {
				if ($S_Sp.ServicePrincipalType)
				{
					$S_Sp.ServicePrincipalType
				}
				else
				{
					'Unknown'
				}
			}
		}

		[pscustomobject]@{
			DisplayName             = $S_Sp.DisplayName
			AppId                   = $S_Sp.AppId
			ObjectId                = $S_Sp.Id
			ServicePrincipalType    = $S_SpTypeName
			AccountEnabled          = $S_Sp.AccountEnabled
			IsMicrosoft             = $S_IsMicrosoft
			HasAppRegistration      = $S_HasAppReg
			CreatedDateTime         = $S_Sp.CreatedDateTime
			SecretCount             = $S_SecretCount
			CertificateCount        = $S_CertCount
			EarliestExpiry          = $S_EarliestExpiry
			DaysUntilExpiry         = $S_DaysUntilExpiry
			CredentialStatus        = $S_CredentialStatus
			GrantedPermissionCount  = $S_GrantedPerms.Count
			IsHighPrivilege         = $S_IsHighPrivilege
			HighPrivPermissions     = ($S_HighPrivPerms -join ', ')
			AllGrantedPermissions   = ($S_GrantedPerms -join ', ')
			LastSignIn              = $S_LastSignIn
			LastDelegatedSignIn     = $S_LastDelegated
			LastAppOnlySignIn       = $S_LastAppOnly
			ActivitySource          = $S_ActivitySource
			DaysSinceActivity       = $S_DaysSinceActivity
			Status                  = $S_Status
		}
	}

	# --- Stats ---
	$S_TotalApps       = @($S_Report).Count
	$S_TotalEnabled    = @($S_Report | Where-Object { $_.AccountEnabled }).Count
	$S_TotalDisabled   = @($S_Report | Where-Object { -not $_.AccountEnabled }).Count
	$S_TotalActive     = @($S_Report | Where-Object { $_.Status -eq 'Active' }).Count
	$S_TotalInactive   = @($S_Report | Where-Object { $_.Status -eq 'Inactive' }).Count
	$S_TotalHighPriv   = @($S_Report | Where-Object { $_.IsHighPrivilege }).Count
	$S_TotalMicrosoft  = @($S_Report | Where-Object { $_.IsMicrosoft }).Count
	$S_TotalWithAppReg = @($S_Report | Where-Object { $_.HasAppRegistration }).Count
	$S_TotalCritical   = @($S_Report | Where-Object { $_.CredentialStatus -eq 'Critical' }).Count
	$S_TotalWarning    = @($S_Report | Where-Object { $_.CredentialStatus -eq 'Warning' }).Count
	$S_TotalExpSoon    = @($S_Report | Where-Object { $_.CredentialStatus -eq 'Expiring Soon' }).Count
	$S_TotalHealthy    = @($S_Report | Where-Object { $_.CredentialStatus -eq 'Healthy' }).Count
	$S_TotalNoCreds    = @($S_Report | Where-Object { $_.CredentialStatus -eq 'No Credentials' }).Count

	$S_SpTypeSummary = $S_Report | Group-Object ServicePrincipalType | Sort-Object Count -Descending | ForEach-Object {
		[pscustomobject]@{ Type = $_.Name; Count = $_.Count }
	}

	# --- File paths ---
	if (-not $ReportPath)
	{
		$ReportPath = (Get-Location).Path
	}
	if (Test-Path $ReportPath -PathType Container)
	{
		$S_ReportFolder = $ReportPath
	}
	else
	{
		$S_ReportFolder = Split-Path -Parent $ReportPath
	}
	if ($S_ReportFolder -and -not (Test-Path $S_ReportFolder))
	{
		New-Item -ItemType Directory -Path $S_ReportFolder -Force | Out-Null
	}

	$S_Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
	if (Test-Path $ReportPath -PathType Container)
	{
		$S_CsvFile = Join-Path $ReportPath ("ReportEntraIDApps_{0}.csv" -f $S_Timestamp)
	}
	else
	{
		$S_CsvFile = $ReportPath
	}

	# --- CSV export ---
	$S_Report | Sort-Object DisplayName | Export-Csv -Path $S_CsvFile -NoTypeInformation -Encoding UTF8

	# --- HTML report ---
	$S_ReportDate = Get-Date -Format 'dd MMM yyyy HH:mm'

	# Build per-app JSON for client-side threshold recalculation
	$S_AppsJson = ($S_Report | Sort-Object DisplayName | ForEach-Object {
		if ($null -ne $_.DaysSinceActivity)
		{
			$S_DaysVal = $_.DaysSinceActivity
		}
		else
		{
			$S_DaysVal = -1
		}
		if ($_.IsHighPrivilege)
		{
			$S_Hp = 'true'
		}
		else
		{
			$S_Hp = 'false'
		}
		if ($_.IsMicrosoft)
		{
			$S_Ms = 'true'
		}
		else
		{
			$S_Ms = 'false'
		}
		if ($_.AccountEnabled)
		{
			$S_En = 'true'
		}
		else
		{
			$S_En = 'false'
		}
		if ($_.HasAppRegistration)
		{
			$S_Ar = 'true'
		}
		else
		{
			$S_Ar = 'false'
		}
		$S_Cs = ($_.CredentialStatus) -replace '"', '\"'
		$S_Spt = ($_.ServicePrincipalType) -replace '"', '\"'
		'{{"days":{0},"hp":{1},"ms":{2},"en":{3},"ar":{4},"cs":"{5}","spt":"{6}"}}' -f $S_DaysVal, $S_Hp, $S_Ms, $S_En, $S_Ar, $S_Cs, $S_Spt
	}) -join ','

	# Build table rows
	$S_TableRows = ($S_Report | Sort-Object DisplayName | ForEach-Object {
		if ($null -ne $_.DaysSinceActivity)
		{
			$S_DaysVal = $_.DaysSinceActivity
		}
		else
		{
			$S_DaysVal = -1
		}
		if ($_.AccountEnabled)
		{
			$S_EnabledVal = '1'
		}
		else
		{
			$S_EnabledVal = '0'
		}
		if ($_.IsMicrosoft)
		{
			$S_MsVal = '1'
		}
		else
		{
			$S_MsVal = '0'
		}
		$S_AppName = [System.Net.WebUtility]::HtmlEncode($_.DisplayName)
		$S_AppIdEnc = [System.Net.WebUtility]::HtmlEncode($_.AppId)
		$S_SpType = [System.Net.WebUtility]::HtmlEncode($_.ServicePrincipalType)
		if ($_.AccountEnabled)
		{
			$S_Enabled = '<span class="badge badge-active">Yes</span>'
		}
		else
		{
			$S_Enabled = '<span class="badge badge-disabled">Disabled</span>'
		}
		if ($_.IsMicrosoft)
		{
			$S_MsBadge = '<span class="badge badge-ms">Microsoft</span>'
		}
		else
		{
			$S_MsBadge = '<span class="badge badge-thirdparty">3rd Party</span>'
		}
		if ($_.HasAppRegistration)
		{
			$S_ArBadge = '<span class="badge badge-active">Yes</span>'
		}
		else
		{
			$S_ArBadge = '<span class="badge badge-nocreds">No</span>'
		}
		if ($_.CreatedDateTime)
		{
			$S_Created = ([datetime]$_.CreatedDateTime).ToString('dd MMM yyyy')
		}
		else
		{
			$S_Created = '-'
		}
		if ($_.SecretCount -eq 0 -and $_.CertificateCount -eq 0)
		{
			$S_Creds = 'None'
		}
		else
		{
			$S_Creds = '{0}S / {1}C' -f $_.SecretCount, $_.CertificateCount
		}
		if ($_.EarliestExpiry)
		{
			$S_Expiry = ([datetime]$_.EarliestExpiry).ToString('dd MMM yyyy')
		}
		else
		{
			$S_Expiry = '-'
		}
		if ($null -ne $_.DaysUntilExpiry)
		{
			$S_DaysUntil = "$($_.DaysUntilExpiry) days"
		}
		else
		{
			$S_DaysUntil = '-'
		}
		$S_CredStatusClass = switch ($_.CredentialStatus)
		{
			'Healthy' { 'healthy' }
			'Expiring Soon' { 'expiring' }
			'Warning' { 'warning' }
			'Critical' { 'critical' }
			'No Credentials' { 'nocreds' }
		}
		if ($_.IsHighPrivilege)
		{
			$S_HighPrivBadge = '<span class="badge badge-highpriv">Yes</span>'
		}
		else
		{
			$S_HighPrivBadge = '<span class="badge badge-lowpriv">No</span>'
		}
		if ($_.LastSignIn)
		{
			$S_LastSignIn = ([datetime]$_.LastSignIn).ToString('dd MMM yyyy')
		}
		else
		{
			$S_LastSignIn = '-'
		}
		if ($null -ne $_.DaysSinceActivity)
		{
			$S_SinceActivity = "$($_.DaysSinceActivity) days"
		}
		else
		{
			$S_SinceActivity = 'Never'
		}
		$S_StatusClass = switch ($_.Status)
		{
			'Active' { 'active' }
			'Inactive' { 'inactive' }
			'Disabled' { 'disabled' }
		}

		"<tr data-days=`"$S_DaysVal`" data-ena=`"$S_EnabledVal`" data-ms=`"$S_MsVal`"><td>$S_AppName</td><td class=`"app-id`">$S_AppIdEnc</td><td>$S_SpType</td><td>$S_Enabled</td><td>$S_MsBadge</td><td>$S_ArBadge</td><td>$S_Created</td><td>$S_Creds</td><td>$S_Expiry</td><td class=`"cred-days`">$S_DaysUntil</td><td><span class=`"badge badge-$S_CredStatusClass`">$($_.CredentialStatus)</span></td><td>$($_.GrantedPermissionCount)</td><td>$S_HighPrivBadge</td><td>$S_LastSignIn</td><td>$S_SinceActivity</td><td><span class=`"badge badge-$S_StatusClass`">$($_.Status)</span></td></tr>"
	}) -join "`n"

	# Build high-privilege apps detail rows
	$S_HighPrivApps = $S_Report | Where-Object { $_.IsHighPrivilege } | Sort-Object DisplayName
	if ($S_HighPrivApps)
	{
		$S_HighPrivRows = ($S_HighPrivApps | ForEach-Object {
			$S_AppName = [System.Net.WebUtility]::HtmlEncode($_.DisplayName)
			$S_Perms = [System.Net.WebUtility]::HtmlEncode($_.HighPrivPermissions)
			if ($_.IsMicrosoft)
			{
				$S_MsBadge = '<span class="badge badge-ms">Microsoft</span>'
			}
			else
			{
				$S_MsBadge = '<span class="badge badge-thirdparty">3rd Party</span>'
			}
			$S_CredBadgeClass = switch ($_.CredentialStatus)
			{
				'Healthy' { 'healthy' }
				'Expiring Soon' { 'expiring' }
				'Warning' { 'warning' }
				'Critical' { 'critical' }
				'No Credentials' { 'nocreds' }
			}
			$S_StatusClass = switch ($_.Status)
			{
				'Active' { 'active' }
				'Inactive' { 'inactive' }
				'Disabled' { 'disabled' }
			}
			"<tr><td>$S_AppName</td><td class=`"app-id`">$([System.Net.WebUtility]::HtmlEncode($_.AppId))</td><td>$S_MsBadge</td><td class=`"perm-list`">$S_Perms</td><td><span class=`"badge badge-$S_CredBadgeClass`">$($_.CredentialStatus)</span></td><td><span class=`"badge badge-$S_StatusClass`">$($_.Status)</span></td></tr>"
		}) -join "`n"
	}
	else
	{
		$S_HighPrivRows = "<tr><td colspan=`"6`" style=`"text-align:center;color:#999;`">No highly privileged applications found</td></tr>"
	}

	$S_Html = @"
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
    <p>Tenant: $([System.Net.WebUtility]::HtmlEncode($S_TenantDisplayName)) ($S_TenantId) &nbsp;|&nbsp; Generated: $S_ReportDate &nbsp;|&nbsp; Total: $S_TotalApps ($S_TotalMicrosoft Microsoft, $($S_TotalApps - $S_TotalMicrosoft) third-party)</p>
  </div>
  <div class="header-right">
    <label class="toggle-label"><input type="checkbox" id="hideMsApps" onchange="applyThreshold()" checked> Hide Microsoft Apps</label>
    <label for="thresholdSelect">Inactive Threshold:</label>
    <select id="thresholdSelect" onchange="applyThreshold()">
      <option value="30" $(if ($InactiveDays -eq 30)
      { 'selected' })>30 Days</option>
      <option value="60" $(if ($InactiveDays -eq 60)
      { 'selected' })>60 Days</option>
      <option value="90" $(if ($InactiveDays -eq 90)
      { 'selected' })>90 Days</option>
      <option value="180" $(if ($InactiveDays -eq 180 -or ($InactiveDays -ne 30 -and $InactiveDays -ne 60 -and $InactiveDays -ne 90 -and $InactiveDays -ne 360))
      { 'selected' })>180 Days</option>
      <option value="360" $(if ($InactiveDays -eq 360)
      { 'selected' })>360 Days</option>
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
  <h2>Highly Privileged Applications — Granted Graph Permissions ($S_TotalHighPriv)</h2>
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
$S_HighPrivRows
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
$S_TableRows
    </tbody>
  </table>
</div>

<div class="footer">Report generated by ReportEntraIDApps.ps1</div>

<script>
var appData = [$S_AppsJson];
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

	$S_HtmlReportFile = Join-Path $S_ReportFolder ("ReportEntraIDApps_{0}.html" -f $S_Timestamp)
	$S_Html | Out-File -FilePath $S_HtmlReportFile -Encoding UTF8

	# --- Console summary ---
	Write-Host ""
	Write-Host "Entra ID Enterprise Applications Report" -ForegroundColor Cyan
	Write-Host "--------------------------------------------"
	Write-Host ("Tenant                   : {0} ({1})" -f $S_TenantDisplayName, $S_TenantId)
	Write-Host ("Total enterprise apps    : {0}" -f $S_TotalApps)
	Write-Host ("  Microsoft first-party  : {0}" -f $S_TotalMicrosoft) -ForegroundColor DarkGray
	Write-Host ("  Third-party / custom   : {0}" -f ($S_TotalApps - $S_TotalMicrosoft))
	Write-Host ("  With app registration  : {0}" -f $S_TotalWithAppReg)
	Write-Host ""
	Write-Host "Activity" -ForegroundColor Cyan
	Write-Host ("  Active                 : {0}" -f $S_TotalActive) -ForegroundColor Green
	Write-Host ("  Inactive               : {0}" -f $S_TotalInactive) -ForegroundColor Red
	Write-Host ("  Disabled               : {0}" -f $S_TotalDisabled) -ForegroundColor DarkGray
	Write-Host ""
	Write-Host "Application Types" -ForegroundColor Cyan
	foreach ($S_Spt in $S_SpTypeSummary)
	{
		Write-Host ("  {0,-25}: {1}" -f $S_Spt.Type, $S_Spt.Count)
	}
	Write-Host ""
	Write-Host "Credential Health" -ForegroundColor Cyan
	Write-Host ("  Healthy                : {0}" -f $S_TotalHealthy) -ForegroundColor Green
	Write-Host ("  Expiring Soon (30d)    : {0}" -f $S_TotalExpSoon) -ForegroundColor Yellow
	Write-Host ("  Warning (some expired) : {0}" -f $S_TotalWarning) -ForegroundColor DarkYellow
	Write-Host ("  Critical (all expired) : {0}" -f $S_TotalCritical) -ForegroundColor Red
	Write-Host ("  No Credentials         : {0}" -f $S_TotalNoCreds) -ForegroundColor DarkGray
	Write-Host ""
	Write-Host "High Privilege Applications (Granted Graph Permissions)" -ForegroundColor Cyan
	Write-Host ("  Count                  : {0}" -f $S_TotalHighPriv) -ForegroundColor Red
	if ($S_HighPrivApps -and $S_HighPrivApps.Count -gt 0)
	{
		foreach ($S_Hp in $S_HighPrivApps)
		{
			Write-Host ("  - {0}" -f $S_Hp.DisplayName)
			Write-Host ("    Granted: {0}" -f $S_Hp.HighPrivPermissions) -ForegroundColor DarkYellow
		}
	}
	Write-Host ""
	Write-Host ("Inactive days threshold  : {0}" -f $InactiveDays)
	Write-Host ("CSV report               : {0}" -f $S_CsvFile) -ForegroundColor Yellow
	Write-Host ("HTML report              : {0}" -f $S_HtmlReportFile) -ForegroundColor Yellow

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
