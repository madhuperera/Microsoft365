param(
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[string]$ReportPath,

	[Parameter(Mandatory = $false)]
	[ValidateRange(1, 3650)]
	[int]$InactiveDays = 90,

	[Parameter(Mandatory = $false)]
	[ValidateSet("ReportOnly", "Disable")]
	[string]$Mode = "ReportOnly",

	[Parameter(Mandatory = $false)]
	[bool]$SkipIfLastSignInIsNEVER = $true
)

$ErrorActionPreference = "Stop"

try {
	if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
		throw "Microsoft.Graph.Users module is not installed. Install it using Install-Module Microsoft.Graph -Scope CurrentUser."
	}

	Import-Module Microsoft.Graph.Users -ErrorAction Stop

	$context = Get-MgContext
	if (-not $context) {
		$scopes = if ($Mode -eq "Disable") { "User.ReadWrite.All", "AuditLog.Read.All" } else { "User.Read.All", "AuditLog.Read.All" }
		Connect-MgGraph -Scopes $scopes -ErrorAction Stop | Out-Null
	}

	$cutoffDate = (Get-Date).AddDays(-$InactiveDays)

	$users = Get-MgUser -All `
		-Filter "userType eq 'Member' and accountEnabled eq true" `
		-Property "id,displayName,userPrincipalName,mail,userType,accountEnabled,createdDateTime,signInActivity,assignedLicenses,onPremisesSyncEnabled" `
		-ConsistencyLevel eventual

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

		[pscustomobject]@{
			DisplayName                 = $user.DisplayName
			UserPrincipalName           = $user.UserPrincipalName
			Mail                        = $user.Mail
			PrimaryDomain               = if ($user.Mail -and $user.Mail -match "@") { ($user.Mail -split "@", 2)[1] } else { $null }
			UserType                    = $user.UserType
			AccountEnabled              = $user.AccountEnabled
			LicenseAssigned             = ($user.AssignedLicenses.Count -gt 0)
			OnPremisesSyncEnabled       = [bool]$user.OnPremisesSyncEnabled
			CreatedDateTime             = $user.CreatedDateTime
			LastInteractiveSignIn       = $lastInteractiveDt
			LastNonInteractiveSignIn    = $lastNonInteractiveDt
			LastSignInAgoDays           = $lastSignInAgo
			Inactive                    = $isInactive
		}
	}

	$inactiveUsers = $report | Where-Object { $_.Inactive }
	$totalMembers = $report.Count
	$totalInactive = $inactiveUsers.Count
	$percentInactive = if ($totalMembers -gt 0) { [math]::Round(($totalInactive / $totalMembers) * 100, 2) } else { 0 }

	if (-not $ReportPath) {
		$ReportPath = (Get-Location).Path
	}

	$reportFolder = if (Test-Path $ReportPath -PathType Container) { $ReportPath } else { Split-Path -Parent $ReportPath }
	if ($reportFolder -and -not (Test-Path $reportFolder)) {
		New-Item -ItemType Directory -Path $reportFolder -Force | Out-Null
	}

	$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
	$reportFile = if (Test-Path $ReportPath -PathType Container) {
		Join-Path $ReportPath ("InactiveMemberUsers_{0}.csv" -f $timestamp)
	} else {
		$ReportPath
	}

	$inactiveUsers | Sort-Object DisplayName | Export-Csv -Path $reportFile -NoTypeInformation -Encoding UTF8

	$disableReport = @()
	if ($Mode -eq "Disable" -and $totalInactive -gt 0) {
		$disableCandidates = if ($SkipIfLastSignInIsNEVER) {
			$inactiveUsers | Where-Object { $_.LastSignInAgoDays -ne "Never" }
		} else {
			$inactiveUsers
		}

		$totalToDisable = $disableCandidates.Count
		if ($SkipIfLastSignInIsNEVER) {
			Write-Host ("Skip users with never sign-in: {0}" -f ($inactiveUsers.Count - $totalToDisable))
		}

		$counter = 0
		foreach ($user in $disableCandidates) {
			$counter++
			Write-Host ("Processing {0}/{1}: {2}" -f $counter, $totalToDisable, $user.UserPrincipalName) -ForegroundColor Yellow
			for ($countdown = 3; $countdown -ge 1; $countdown--) {
				Write-Host ("Disabling in {0}..." -f $countdown)
				Start-Sleep -Seconds 1
			}

			$disableStatus = "Success"
			$disableError = $null
			try {
				Update-MgUser -UserId $user.UserPrincipalName -AccountEnabled:$false -ErrorAction Stop
			}
			catch {
				$disableStatus = "Failed"
				$disableError = $_.Exception.Message
			}

			$disableReport += [pscustomobject]@{
				DisplayName       = $user.DisplayName
				UserPrincipalName = $user.UserPrincipalName
				Disabled          = $disableStatus
				Error             = $disableError
			}
		}
	}

	$disableReportFile = $null
	if ($disableReport.Count -gt 0) {
		$disableReportFile = Join-Path $reportFolder ("DisabledMemberUsers_{0}.csv" -f $timestamp)
		$disableReport | Export-Csv -Path $disableReportFile -NoTypeInformation -Encoding UTF8
	}

	Write-Host "Inactive Member Users Report" -ForegroundColor Cyan
	Write-Host "--------------------------------------------"
	Write-Host ("Total enabled member users : {0}" -f $totalMembers)
	Write-Host ("Inactive member users      : {0}" -f $totalInactive)
	Write-Host ("Inactive percentage       : {0}%" -f $percentInactive)
	Write-Host ("Inactive days threshold   : {0}" -f $InactiveDays)
	Write-Host ("Report exported to        : {0}" -f $reportFile)
	Write-Host ("Mode                     : {0}" -f $Mode)
	if ($disableReportFile) {
		Write-Host ("Disable report exported  : {0}" -f $disableReportFile)
	}

	$disconnectChoice = Read-Host "Disconnect from Microsoft Graph? (Y/N)"
	if ($disconnectChoice -match '^(y|yes)$') {
		Disconnect-MgGraph -ErrorAction SilentlyContinue
	}
}
catch {
	Write-Error $_
	exit 1
}
