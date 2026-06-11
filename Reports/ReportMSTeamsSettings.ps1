#Requires -Modules MicrosoftTeams

<#
.SYNOPSIS
    Reports on a set of tenant-wide Microsoft Teams governance and security
    settings, and lists any Teams that have no owner.

.DESCRIPTION
    Connects to Microsoft Teams (MicrosoftTeams PowerShell module) and reads
    the effective organization-level configuration to report on the following
    items. Each item shows its actual tenant value and a Compliant /
    Not Compliant verdict against the required hardened state:

      * Approved third-party cloud storage providers
        (Get-CsTeamsClientConfiguration)
      * Teams that have no owner
        (Get-Team / Get-TeamUser)
      * Communication with unmanaged ("personal"/consumer) Teams users disabled
        (Get-CsTenantFederationConfiguration -> AllowTeamsConsumer)
      * External (unmanaged) Teams users allowed to initiate conversations
        (Get-CsTenantFederationConfiguration -> AllowTeamsConsumerInbound)
      * Communication with trial-only Teams tenants restricted
        (Get-CsTenantFederationConfiguration -> ExternalAccessWithTrialTenants)
      * App permission policies - Teams apps controlled
        (Get-CsTeamsAppPermissionPolicy -Identity Global)
      * Anonymous / dial-in callers allowed to start meetings
        (Get-CsTeamsMeetingPolicy -Identity Global -> AllowAnonymousUsersToStartMeeting)
      * Meeting lobby - only organization users bypass the lobby
        (Get-CsTeamsMeetingPolicy -Identity Global -> AutoAdmittedUsers)
      * Phone (PSTN) dial-in users can bypass the lobby
        (Get-CsTeamsMeetingPolicy -Identity Global -> AllowPSTNUsersToBypassLobby)

    The script is read-only; it does not make any tenant changes. Results are
    written to a CSV (settings summary), a second CSV (Teams with no owner) and
    a standalone interactive HTML report. Output is timestamped and placed in
    the current location (or under -ReportPath when supplied).

.PARAMETER ReportPath
    Optional. Path to a folder or full file path for the CSV output. When a
    folder is supplied (or the parameter is omitted) timestamped file names are
    generated in that folder. Defaults to the current working location.

.EXAMPLE
    .\ReportMSTeamsSettings.ps1

    Generates timestamped CSV and HTML reports in the current directory.

.EXAMPLE
    .\ReportMSTeamsSettings.ps1 -ReportPath C:\Reports

    Writes the reports into C:\Reports.

.NOTES
    Required module:
      - MicrosoftTeams (Install-Module MicrosoftTeams -Scope CurrentUser)

    Required role (read-only is sufficient):
      - Global Reader, Teams Administrator, or Teams Communications Administrator

    Note: Get-CsTeamsAppPermissionPolicy only returns data on tenants that have
    not been migrated to App Centric Management (ACM) / Unified App Management
    (UAM). On migrated tenants this item is reported as "Not available".

    Shared as-is without warranty.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ReportPath
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Script-level configuration
# ---------------------------------------------------------------------------

$S_RequiredModules                      = @('MicrosoftTeams')
$S_TeamsRequestDelayMilliseconds        = 5
$S_RequireContextConfirmation           = $true
$S_ContextConfirmationDelaySeconds      = 10
$S_DisconnectSessionOnExit              = $false

# Third-party cloud storage providers exposed by Get-CsTeamsClientConfiguration.
# Key = friendly name, Value = property name on the configuration object.
$S_CloudStorageProviders = [ordered]@{
    'DropBox'     = 'AllowDropBox'
    'Box'         = 'AllowBox'
    'GoogleDrive' = 'AllowGoogleDrive'
    'ShareFile'   = 'AllowShareFile'
    'Egnyte'      = 'AllowEgnyte'
}

try
{
    # -----------------------------------------------------------------------
    # Module checks
    # -----------------------------------------------------------------------

    foreach ($S_Module in $S_RequiredModules)
    {
        if (-not (Get-Module -ListAvailable -Name $S_Module))
        {
            throw "$S_Module module is not installed. Install it using 'Install-Module $S_Module -Scope CurrentUser'."
        }

        Import-Module $S_Module -ErrorAction Stop | Out-Null
    }

    # -----------------------------------------------------------------------
    # Microsoft Teams connection handling
    # -----------------------------------------------------------------------

    Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Cyan
    $S_Connection = Connect-MicrosoftTeams -ErrorAction Stop

    $S_Account  = if ($S_Connection.Account) { $S_Connection.Account.Id } else { $S_Connection.Account }
    $S_TenantId = if ($S_Connection.Tenant) { $S_Connection.Tenant.Id } else { $S_Connection.TenantId }

    Write-Host "Connected to Microsoft Teams:" -ForegroundColor Cyan
    Write-Host ("  Account     : {0}" -f $S_Account)
    Write-Host ("  Tenant ID   : {0}" -f $S_TenantId)

    if ($S_RequireContextConfirmation)
    {
        $S_ConfirmChoice = Read-Host "Is this the correct tenant and account to continue? (Y/N)"
        if ($S_ConfirmChoice -notmatch '^(y|yes)$')
        {
            Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null
            throw "Operator did not confirm the Microsoft Teams context. Aborting."
        }

        Start-Sleep -Seconds $S_ContextConfirmationDelaySeconds
    }

    # Resolve a friendly tenant display name (best effort).
    $S_TenantDisplayName = $null
    try
    {
        $S_Tenant = Get-CsTenant -ErrorAction Stop
        $S_TenantDisplayName = $S_Tenant.DisplayName
        if (-not $S_TenantId) { $S_TenantId = $S_Tenant.TenantId }
    }
    catch
    {
        Write-Verbose ("Failed to resolve tenant display name: {0}" -f $_.Exception.Message)
    }
    if (-not $S_TenantDisplayName) { $S_TenantDisplayName = if ($S_TenantId) { $S_TenantId } else { 'Unknown' } }
    if (-not $S_TenantId) { $S_TenantId = 'Unknown' }

    # -----------------------------------------------------------------------
    # Collect tenant-wide configuration
    # -----------------------------------------------------------------------

    Write-Host ""
    Write-Host "Reading Teams tenant configuration..." -ForegroundColor Cyan

    Write-Verbose "Retrieving Teams client configuration (cloud storage)."
    $S_ClientConfig = Get-CsTeamsClientConfiguration -Identity Global -ErrorAction Stop
    Start-Sleep -Milliseconds $S_TeamsRequestDelayMilliseconds

    Write-Verbose "Retrieving tenant federation configuration (external access)."
    $S_Federation = Get-CsTenantFederationConfiguration -ErrorAction Stop
    Start-Sleep -Milliseconds $S_TeamsRequestDelayMilliseconds

    Write-Verbose "Retrieving global Teams meeting policy (lobby / anonymous)."
    $S_MeetingPolicy = Get-CsTeamsMeetingPolicy -Identity Global -ErrorAction Stop
    Start-Sleep -Milliseconds $S_TeamsRequestDelayMilliseconds

    Write-Verbose "Retrieving global Teams app permission policy."
    $S_AppPolicy = $null
    $S_AppPolicyAvailable = $true
    try
    {
        $S_AppPolicy = Get-CsTeamsAppPermissionPolicy -Identity Global -ErrorAction Stop
    }
    catch
    {
        # Cmdlet returns no data on ACM/UAM-migrated tenants.
        $S_AppPolicyAvailable = $false
        Write-Verbose ("App permission policy not available: {0}" -f $_.Exception.Message)
    }
    Start-Sleep -Milliseconds $S_TeamsRequestDelayMilliseconds

    # -----------------------------------------------------------------------
    # Collect Teams without owners
    # -----------------------------------------------------------------------

    Write-Host "Enumerating Teams and checking ownership..." -ForegroundColor Cyan

    $S_AllTeams = @(Get-Team -ErrorAction Stop)
    Start-Sleep -Milliseconds $S_TeamsRequestDelayMilliseconds

    $S_TotalTeams = ($S_AllTeams | Measure-Object).Count
    $S_OwnerlessTeams = New-Object System.Collections.Generic.List[object]
    $S_Counter = 0

    foreach ($S_Team in $S_AllTeams)
    {
        $S_Counter++

        $S_PercentComplete = if ($S_TotalTeams -gt 0) { [int](($S_Counter / $S_TotalTeams) * 100) } else { 0 }
        Write-Progress -Activity "Checking Team ownership" `
            -Status ("[{0}/{1}] {2}" -f $S_Counter, $S_TotalTeams, $S_Team.DisplayName) `
            -PercentComplete $S_PercentComplete

        $F_OwnerCount = 0
        try
        {
            $F_Owners = @(Get-TeamUser -GroupId $S_Team.GroupId -Role Owner -ErrorAction Stop)
            $F_OwnerCount = ($F_Owners | Measure-Object).Count
            Start-Sleep -Milliseconds $S_TeamsRequestDelayMilliseconds
        }
        catch
        {
            Write-Warning ("Failed to retrieve owners for '{0}': {1}" -f $S_Team.DisplayName, $_.Exception.Message)
            continue
        }

        if ($F_OwnerCount -eq 0)
        {
            $S_OwnerlessTeams.Add([pscustomobject]@{
                DisplayName  = $S_Team.DisplayName
                MailNickname = $S_Team.MailNickName
                Visibility   = $S_Team.Visibility
                Archived     = $S_Team.Archived
                GroupId      = $S_Team.GroupId
            })
        }
    }

    Write-Progress -Activity "Checking Team ownership" -Completed

    $S_OwnerlessCount = ($S_OwnerlessTeams | Measure-Object).Count

    # -----------------------------------------------------------------------
    # Evaluate each setting for compliance
    # -----------------------------------------------------------------------
    # Compliance values: 'Compliant'     (meets the required hardened state)
    #                    'Not Compliant' (deviates from the required state)
    #                    'Unknown'       (state could not be determined)
    #
    # Each entry records the ACTUAL value read from the tenant and the
    # compliance verdict derived from the requirements below.

    $S_Settings = New-Object System.Collections.Generic.List[object]

    # 1. Approved cloud storage providers
    #    Compliant requirement: ALL third-party cloud storage providers disabled.
    $S_EnabledStorage = foreach ($F_Provider in $S_CloudStorageProviders.Keys)
    {
        $F_PropertyName = $S_CloudStorageProviders[$F_Provider]
        if ($S_ClientConfig.$F_PropertyName -eq $true) { $F_Provider }
    }
    $S_EnabledStorage = @($S_EnabledStorage)
    $S_StorageValue   = if ($S_EnabledStorage.Count -gt 0) { "Enabled: " + ($S_EnabledStorage -join ', ') } else { 'All disabled' }
    $S_Settings.Add([pscustomobject]@{
        Category    = 'Cloud Storage'
        Setting     = 'Third-party cloud storage providers'
        ActualValue = $S_StorageValue
        Compliance  = $(if ($S_EnabledStorage.Count -eq 0) { 'Compliant' } else { 'Not Compliant' })
        Requirement = 'All third-party cloud storage providers must be disabled'
    })

    # 2. Teams without owners
    #    Compliant requirement: no Team without an owner.
    $S_Settings.Add([pscustomobject]@{
        Category    = 'Ownership'
        Setting     = 'Teams without an owner'
        ActualValue = "$S_OwnerlessCount of $S_TotalTeams Teams have no owner"
        Compliance  = $(if ($S_OwnerlessCount -eq 0) { 'Compliant' } else { 'Not Compliant' })
        Requirement = 'No Team may exist without an owner'
    })

    # 3. Communication with unmanaged Teams users disabled
    #    Compliant requirement: AllowTeamsConsumer = False.
    $S_Settings.Add([pscustomobject]@{
        Category    = 'External Access'
        Setting     = 'Communication with unmanaged Teams users'
        ActualValue = "AllowTeamsConsumer = $([string]$S_Federation.AllowTeamsConsumer)"
        Compliance  = $(if ($S_Federation.AllowTeamsConsumer -eq $false) { 'Compliant' } else { 'Not Compliant' })
        Requirement = 'Communication with unmanaged Teams users must be disabled'
    })

    # 4. External Teams users allowed to initiate conversations
    #    Compliant requirement: unmanaged users cannot initiate, i.e. consumer
    #    access is off OR inbound is off.
    $S_ExternalCanInitiate = ($S_Federation.AllowTeamsConsumer -eq $true -and $S_Federation.AllowTeamsConsumerInbound -eq $true)
    $S_Settings.Add([pscustomobject]@{
        Category    = 'External Access'
        Setting     = 'External Teams users can initiate conversations'
        ActualValue = "AllowTeamsConsumerInbound = $([string]$S_Federation.AllowTeamsConsumerInbound)"
        Compliance  = $(if ($S_ExternalCanInitiate) { 'Not Compliant' } else { 'Compliant' })
        Requirement = 'External Teams users must not be allowed to initiate conversations'
    })

    # 5. Communication with trial Teams tenants disabled
    #    Compliant requirement: ExternalAccessWithTrialTenants = Blocked.
    $S_Settings.Add([pscustomobject]@{
        Category    = 'External Access'
        Setting     = 'Communication with trial-only Teams tenants'
        ActualValue = "ExternalAccessWithTrialTenants = $([string]$S_Federation.ExternalAccessWithTrialTenants)"
        Compliance  = $(if ([string]$S_Federation.ExternalAccessWithTrialTenants -eq 'Blocked') { 'Compliant' } else { 'Not Compliant' })
        Requirement = 'Communication with trial-only Teams tenants must be disabled (Blocked)'
    })

    # 6. App permission policies - users cannot install apps without admin approval
    #    Compliant requirement: apps are restricted to an admin-approved allow
    #    list (no catalog left as an open BlockedAppList = allow-all).
    if ($S_AppPolicyAvailable -and $S_AppPolicy)
    {
        $S_AppTypes = @($S_AppPolicy.DefaultCatalogAppsType, $S_AppPolicy.GlobalCatalogAppsType, $S_AppPolicy.PrivateCatalogAppsType)
        # Allow-all (non-compliant) when any catalog is an open BlockedAppList.
        $S_AppsOpen      = @($S_AppTypes | Where-Object { $_ -eq 'BlockedAppList' }).Count -gt 0
        $S_AppValue      = "Default=$($S_AppPolicy.DefaultCatalogAppsType); Global=$($S_AppPolicy.GlobalCatalogAppsType); Private=$($S_AppPolicy.PrivateCatalogAppsType)"
        $S_AppCompliance = $(if ($S_AppsOpen) { 'Not Compliant' } else { 'Compliant' })
    }
    else
    {
        $S_AppValue      = 'Not available (App Centric Management / UAM)'
        $S_AppCompliance = 'Unknown'
    }
    $S_Settings.Add([pscustomobject]@{
        Category    = 'Apps'
        Setting     = 'App permission policies'
        ActualValue = $S_AppValue
        Compliance  = $S_AppCompliance
        Requirement = 'Users must not be able to install apps unless an admin approves them'
    })

    # 7. Anonymous / dial-in callers allowed to start meetings
    #    Compliant requirement: AllowAnonymousUsersToStartMeeting = False.
    $S_Settings.Add([pscustomobject]@{
        Category    = 'Meetings'
        Setting     = 'Anonymous / dial-in callers can start meetings'
        ActualValue = "AllowAnonymousUsersToStartMeeting = $([string]$S_MeetingPolicy.AllowAnonymousUsersToStartMeeting)"
        Compliance  = $(if ($S_MeetingPolicy.AllowAnonymousUsersToStartMeeting -eq $false) { 'Compliant' } else { 'Not Compliant' })
        Requirement = 'AllowAnonymousUsersToStartMeeting must be disabled'
    })

    # 8. Meeting lobby - auto-admitted users must not include guests
    #    Compliant requirement: AutoAdmittedUsers is a restricted value that
    #    excludes guests / external users.
    $S_AutoAdmit = [string]$S_MeetingPolicy.AutoAdmittedUsers
    $S_AutoAdmitCompliantValues = @('EveryoneInCompanyExcludingGuests', 'OrganizerOnly', 'InvitedUsers')
    $S_Settings.Add([pscustomobject]@{
        Category    = 'Meetings'
        Setting     = 'Meeting lobby - auto-admitted users'
        ActualValue = "AutoAdmittedUsers = $S_AutoAdmit"
        Compliance  = $(if ($S_AutoAdmit -in $S_AutoAdmitCompliantValues) { 'Compliant' } else { 'Not Compliant' })
        Requirement = "AutoAdmittedUsers must be a restricted value that excludes guests ($($S_AutoAdmitCompliantValues -join ', '))"
    })

    # 9. Phone (PSTN) dial-in users can bypass the lobby
    #    Compliant requirement: AllowPSTNUsersToBypassLobby = False.
    $S_Settings.Add([pscustomobject]@{
        Category    = 'Meetings'
        Setting     = 'Phone dial-in users can bypass the lobby'
        ActualValue = "AllowPSTNUsersToBypassLobby = $([string]$S_MeetingPolicy.AllowPSTNUsersToBypassLobby)"
        Compliance  = $(if ($S_MeetingPolicy.AllowPSTNUsersToBypassLobby -eq $false) { 'Compliant' } else { 'Not Compliant' })
        Requirement = 'AllowPSTNUsersToBypassLobby must be disabled'
    })

    # -----------------------------------------------------------------------
    # Resolve output paths
    # -----------------------------------------------------------------------

    if (-not $ReportPath)
    {
        $ReportPath = (Get-Location).Path
    }

    $S_Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

    if (Test-Path -Path $ReportPath -PathType Container)
    {
        $S_OutputPath = Join-Path -Path $ReportPath -ChildPath ("ReportMSTeamsSettings_{0}.csv" -f $S_Timestamp)
    }
    else
    {
        $S_OutputFolder = Split-Path -Parent $ReportPath
        if ($S_OutputFolder -and -not (Test-Path -Path $S_OutputFolder))
        {
            New-Item -ItemType Directory -Path $S_OutputFolder -Force | Out-Null
        }
        $S_OutputPath = $ReportPath
    }

    $S_OwnerlessOutputPath = [System.IO.Path]::ChangeExtension($S_OutputPath, $null).TrimEnd('.') + '_NoOwner.csv'
    $S_HtmlOutputPath      = [System.IO.Path]::ChangeExtension($S_OutputPath, '.html')

    $S_Settings |
        Select-Object Category, Setting, ActualValue, Compliance, Requirement |
        Export-Csv -Path $S_OutputPath -NoTypeInformation -Encoding UTF8

    if ($S_OwnerlessCount -gt 0)
    {
        $S_OwnerlessTeams |
            Sort-Object DisplayName |
            Export-Csv -Path $S_OwnerlessOutputPath -NoTypeInformation -Encoding UTF8
    }

    # -----------------------------------------------------------------------
    # Summary statistics
    # -----------------------------------------------------------------------

    $S_CompliantCount    = ($S_Settings | Where-Object { $_.Compliance -eq 'Compliant' }     | Measure-Object).Count
    $S_NotCompliantCount = ($S_Settings | Where-Object { $_.Compliance -eq 'Not Compliant' } | Measure-Object).Count
    $S_UnknownCount      = ($S_Settings | Where-Object { $_.Compliance -eq 'Unknown' }       | Measure-Object).Count
    $S_OverallStatus     = if ($S_NotCompliantCount -eq 0 -and $S_UnknownCount -eq 0) { 'Compliant' } else { 'Not Compliant' }
    $S_OverallCardClass  = if ($S_OverallStatus -eq 'Compliant') { 'good' } else { 'critical' }
    $S_OverallColor      = if ($S_OverallStatus -eq 'Compliant') { '#27ae60' } else { '#e74c3c' }

    # -----------------------------------------------------------------------
    # Build HTML report
    # -----------------------------------------------------------------------

    $S_ReportDate = Get-Date -Format 'dd MMM yyyy HH:mm'

    $S_SettingRows = ($S_Settings | ForEach-Object {
        $F_ComplianceClass = switch ($_.Compliance)
        {
            'Compliant'     { 'badge-active' }
            'Not Compliant' { 'badge-critical' }
            default         { 'badge-disabled' }
        }
        $F_RowClass    = if ($_.Compliance -eq 'Not Compliant') { 'row-critical' } else { '' }
        $F_ComplianceKey = ($_.Compliance -replace '\s', '').ToLower()

        "<tr class=`"$F_RowClass`" data-compliance=`"$F_ComplianceKey`"><td>$([System.Net.WebUtility]::HtmlEncode($_.Setting))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.ActualValue))</td><td><span class=`"badge $F_ComplianceClass`">$($_.Compliance)</span></td><td>$([System.Net.WebUtility]::HtmlEncode($_.Category))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.Requirement))</td></tr>"
    }) -join "`n"

    if ($S_OwnerlessCount -gt 0)
    {
        $S_OwnerlessRows = ($S_OwnerlessTeams | Sort-Object DisplayName | ForEach-Object {
            $F_Vis      = if ($_.Visibility) { $_.Visibility } else { 'Unknown' }
            $F_Archived = if ($_.Archived) { 'Yes' } else { 'No' }
            "<tr><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.DisplayName))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.MailNickname))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$F_Vis))</td><td>$F_Archived</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.GroupId))</td></tr>"
        }) -join "`n"
    }
    else
    {
        $S_OwnerlessRows = "<tr><td colspan=`"5`" style=`"text-align:center;color:#27ae60;`">All Teams have at least one owner.</td></tr>"
    }

    $S_Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Microsoft Teams Settings Report</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; padding: 30px; }
  .header { background: linear-gradient(135deg, #1a1a2e, #16213e); color: #fff; padding: 30px 40px; border-radius: 12px; margin-bottom: 30px; }
  .header h1 { font-size: 1.6em; margin-bottom: 6px; }
  .header p { font-size: 0.9em; opacity: 0.8; }

  .section-title { font-size: 1.15em; font-weight: 600; color: #1a1a2e; margin-bottom: 16px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0; }

  .summary-cards { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .card { background: #fff; border-radius: 10px; padding: 24px 30px; flex: 1; min-width: 160px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
  .card .label { font-size: 0.85em; color: #777; text-transform: uppercase; letter-spacing: 0.5px; }
  .card .value { font-size: 2em; font-weight: 700; margin-top: 6px; }
  .card.critical { border-left: 4px solid #e74c3c; }
  .card.good     { border-left: 4px solid #27ae60; }

  .charts-row { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .chart-section { background: #fff; border-radius: 10px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); flex: 1; min-width: 360px; }
  .chart-section h2 { font-size: 1.1em; margin-bottom: 20px; color: #1a1a2e; }
  .chart-container { max-width: 360px; margin: 0 auto; }

  .table-section { background: #fff; border-radius: 10px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 30px; overflow-x: auto; }
  .table-section h2 { font-size: 1.1em; margin-bottom: 16px; color: #1a1a2e; }
  .table-controls { display: flex; gap: 14px; margin-bottom: 16px; flex-wrap: wrap; align-items: center; }
  .table-controls input[type="text"] { padding: 8px 14px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; min-width: 260px; }
  .table-controls select { padding: 8px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; background: #fff; }
  .table-controls .count-label { font-size: 0.85em; color: #777; margin-left: auto; }
  table { width: 100%; border-collapse: collapse; font-size: 0.86em; }
  th { background: #1a1a2e; color: #fff; padding: 12px 14px; text-align: left; position: sticky; top: 0; cursor: pointer; user-select: none; white-space: nowrap; }
  th:hover { background: #2c3e50; }
  td { padding: 10px 14px; border-bottom: 1px solid #eee; vertical-align: top; }
  tr:hover td { background: #f8f9fa; }
  tr.hidden-row { display: none; }
  tr.row-critical td { background: #fdecea; }
  tr.row-critical:hover td { background: #fbd6d2; }

  .badge { padding: 3px 10px; border-radius: 12px; font-size: 0.82em; font-weight: 600; }
  .badge-active   { background: #d4edda; color: #155724; }
  .badge-critical { background: #f8d7da; color: #721c24; }
  .badge-disabled { background: #e2e3e5; color: #495057; }

  .footer { text-align: center; font-size: 0.8em; color: #999; margin-top: 20px; }
</style>
</head>
<body>

<div class="header">
  <h1>Microsoft Teams Settings Report</h1>
  <p>Tenant: $([System.Net.WebUtility]::HtmlEncode($S_TenantDisplayName)) ($S_TenantId) &nbsp;|&nbsp; Generated: $S_ReportDate</p>
</div>

<!-- OVERVIEW -->
<div class="section-title">Overview</div>
<div class="summary-cards">
  <div class="card $S_OverallCardClass"><div class="label">Overall Status</div><div class="value" style="color:$S_OverallColor;">$S_OverallStatus</div></div>
  <div class="card"><div class="label">Settings Checked</div><div class="value" style="color:#1a1a2e;">$($S_Settings.Count)</div></div>
  <div class="card good"><div class="label">Compliant</div><div class="value" style="color:#27ae60;">$S_CompliantCount</div></div>
  <div class="card critical"><div class="label">Not Compliant</div><div class="value" style="color:#e74c3c;">$S_NotCompliantCount</div></div>
  <div class="card"><div class="label">Unknown</div><div class="value" style="color:#7f8c8d;">$S_UnknownCount</div></div>
  <div class="card critical"><div class="label">Teams Without Owner</div><div class="value" style="color:#e74c3c;">$S_OwnerlessCount</div></div>
</div>

<!-- CHART -->
<div class="charts-row">
  <div class="chart-section">
    <h2>Compliance Summary</h2>
    <div class="chart-container"><canvas id="complianceChart"></canvas></div>
  </div>
</div>

<!-- SETTINGS TABLE -->
<div class="table-section">
  <h2>Teams Compliance Settings</h2>
  <div class="table-controls">
    <input type="text" id="searchBox" placeholder="Search settings..." onkeyup="filterTable()" />
    <select id="complianceFilter" onchange="filterTable()">
      <option value="all">All Statuses</option>
      <option value="notcompliant">Not Compliant</option>
      <option value="compliant">Compliant</option>
      <option value="unknown">Unknown</option>
    </select>
    <span class="count-label" id="rowCount"></span>
  </div>
  <table id="settingsTable">
    <thead><tr>
      <th onclick="sortTable(0)">Setting</th>
      <th onclick="sortTable(1)">Actual Value</th>
      <th onclick="sortTable(2)">Compliance</th>
      <th onclick="sortTable(3)">Category</th>
      <th onclick="sortTable(4)">Requirement</th>
    </tr></thead>
    <tbody>
$S_SettingRows
    </tbody>
  </table>
</div>

<!-- TEAMS WITHOUT OWNERS -->
<div class="table-section">
  <h2>Teams Without Owners ($S_OwnerlessCount)</h2>
  <table id="ownerlessTable">
    <thead><tr>
      <th>Display Name</th>
      <th>Mail Nickname</th>
      <th>Visibility</th>
      <th>Archived</th>
      <th>Group Id</th>
    </tr></thead>
    <tbody>
$S_OwnerlessRows
    </tbody>
  </table>
</div>

<div class="footer">Report generated by ReportMSTeamsSettings.ps1</div>

<script>
new Chart(document.getElementById('complianceChart'), {
  type: 'doughnut',
  data: {
    labels: ['Compliant', 'Not Compliant', 'Unknown'],
    datasets: [{ data: [$S_CompliantCount, $S_NotCompliantCount, $S_UnknownCount], backgroundColor: ['#27ae60', '#e74c3c', '#95a5a6'], borderWidth: 2, borderColor: '#fff' }]
  },
  options: { responsive: true, plugins: { legend: { position: 'right', labels: { padding: 16, font: { size: 13 }, boxWidth: 16 } } } }
});

function filterTable() {
  var search = document.getElementById('searchBox').value.toLowerCase();
  var compliance = document.getElementById('complianceFilter').value;
  var rows = document.querySelectorAll('#settingsTable tbody tr');
  var visible = 0;
  rows.forEach(function(row) {
    var text = row.textContent.toLowerCase();
    var rowCompliance = row.getAttribute('data-compliance');
    var matchSearch = !search || text.indexOf(search) !== -1;
    var matchCompliance = compliance === 'all' || rowCompliance === compliance;
    if (matchSearch && matchCompliance) {
      row.classList.remove('hidden-row');
      visible++;
    } else {
      row.classList.add('hidden-row');
    }
  });
  document.getElementById('rowCount').textContent = 'Showing ' + visible + ' of ' + rows.length + ' settings';
}

var sortDir = {};
function sortTable(col) {
  var tbody = document.getElementById('settingsTable').querySelector('tbody');
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

filterTable();
</script>
</body>
</html>
"@

    $S_Html | Out-File -FilePath $S_HtmlOutputPath -Encoding UTF8

    # -----------------------------------------------------------------------
    # Console summary
    # -----------------------------------------------------------------------

    Write-Host ""
    Write-Host "Microsoft Teams Settings Report" -ForegroundColor Cyan
    Write-Host "--------------------------------------------"
    foreach ($S_Item in $S_Settings)
    {
        $F_Color = switch ($S_Item.Compliance)
        {
            'Compliant'     { 'Green' }
            'Not Compliant' { 'Red' }
            default         { 'Gray' }
        }
        Write-Host ("[{0}] {1}: {2}" -f $S_Item.Compliance, $S_Item.Setting, $S_Item.ActualValue) -ForegroundColor $F_Color
    }
    Write-Host "--------------------------------------------"
    Write-Host ("Overall status : {0}" -f $S_OverallStatus) -ForegroundColor $(if ($S_OverallStatus -eq 'Compliant') { 'Green' } else { 'Red' })
    Write-Host ("Compliant      : {0}" -f $S_CompliantCount) -ForegroundColor Green
    Write-Host ("Not compliant  : {0}" -f $S_NotCompliantCount) -ForegroundColor $(if ($S_NotCompliantCount -gt 0) { 'Red' } else { 'Gray' })
    Write-Host ("Unknown        : {0}" -f $S_UnknownCount)
    Write-Host ("Teams w/o owner: {0}" -f $S_OwnerlessCount) -ForegroundColor $(if ($S_OwnerlessCount -gt 0) { 'Red' } else { 'Gray' })
    Write-Host ("CSV report exported to  : {0}" -f $S_OutputPath)
    if ($S_OwnerlessCount -gt 0)
    {
        Write-Host ("No-owner CSV exported to: {0}" -f $S_OwnerlessOutputPath)
    }
    Write-Host ("HTML report exported to : {0}" -f $S_HtmlOutputPath)

    # -----------------------------------------------------------------------
    # Disconnection handling
    # -----------------------------------------------------------------------

    if ($S_DisconnectSessionOnExit)
    {
        Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null
    }
    else
    {
        $S_DisconnectChoice = Read-Host "Disconnect from Microsoft Teams? (Y/N)"
        if ($S_DisconnectChoice -match '^(y|yes)$')
        {
            Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null
        }
    }
}
catch
{
    Write-Error $_
    exit 1
}
