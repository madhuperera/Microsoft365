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

$ErrorActionPreference = 'Stop'

$S_ReportPath = $ReportPath

$S_RequiredGraphScopes = @(
    'DeviceManagementConfiguration.Read.All'
    'DeviceManagementManagedDevices.Read.All'
    'DeviceManagementServiceConfig.Read.All'
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

    # --- Tenant compliance settings ---
    Write-Host "Reading tenant compliance settings..." -ForegroundColor Cyan
    $S_SecureByDefault = $null
    $S_CheckinThresholdDays = $null
    try
    {
        $S_Settings = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/beta/deviceManagement/settings' -ErrorAction Stop
        $S_SecureByDefault = $S_Settings.secureByDefault
        $S_CheckinThresholdDays = $S_Settings.deviceComplianceCheckinThresholdDays
    }
    catch
    {
        Write-Warning "Failed to read tenant compliance settings: $($_.Exception.Message)"
    }
    if ($null -eq $S_CheckinThresholdDays -or $S_CheckinThresholdDays -le 0)
    {
        $S_CheckinThresholdDays = 30
    }
    Write-Host ("  secureByDefault                      : {0}" -f $S_SecureByDefault) -ForegroundColor Green
    Write-Host ("  deviceComplianceCheckinThresholdDays : {0}" -f $S_CheckinThresholdDays) -ForegroundColor Green

    # --- Tenant-wide compliance state summary ---
    Write-Host "Fetching tenant-wide compliance summary..." -ForegroundColor Cyan
    $S_TenantSummary = $null
    try
    {
        $S_TenantSummary = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicyDeviceStateSummary' -ErrorAction Stop
    }
    catch
    {
        Write-Warning "Failed to read deviceCompliancePolicyDeviceStateSummary: $($_.Exception.Message)"
    }

    $S_TenantStates = [ordered]@{
        Compliant     = if ($S_TenantSummary)
        {
            [int]$S_TenantSummary.compliantDeviceCount
        }
        else
        {
            0
        }
        NonCompliant  = if ($S_TenantSummary)
        {
            [int]$S_TenantSummary.nonCompliantDeviceCount
        }
        else
        {
            0
        }
        InGracePeriod = if ($S_TenantSummary)
        {
            [int]$S_TenantSummary.inGracePeriodCount
        }
        else
        {
            0
        }
        Error         = if ($S_TenantSummary)
        {
            [int]$S_TenantSummary.errorDeviceCount
        }
        else
        {
            0
        }
        Conflict      = if ($S_TenantSummary)
        {
            [int]$S_TenantSummary.conflictDeviceCount
        }
        else
        {
            0
        }
        NotApplicable = if ($S_TenantSummary)
        {
            [int]$S_TenantSummary.notApplicableDeviceCount
        }
        else
        {
            0
        }
        Remediated    = if ($S_TenantSummary)
        {
            [int]$S_TenantSummary.remediatedDeviceCount
        }
        else
        {
            0
        }
        ConfigManager = if ($S_TenantSummary)
        {
            [int]$S_TenantSummary.configManagerCount
        }
        else
        {
            0
        }
        Unknown       = if ($S_TenantSummary)
        {
            [int]$S_TenantSummary.unknownDeviceCount
        }
        else
        {
            0
        }
    }
    $S_TenantTotal = ($S_TenantStates.Values | Measure-Object -Sum).Sum

    # --- Managed devices (used for active-devices view) ---
    Write-Host "Fetching managed devices..." -ForegroundColor Cyan
    $S_Select = 'id,deviceName,userPrincipalName,userDisplayName,operatingSystem,osVersion,complianceState,managedDeviceOwnerType,lastSyncDateTime,enrolledDateTime'
    $S_Uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$select=$S_Select&`$top=200"
    $S_Devices = New-Object System.Collections.Generic.List[object]
    do
    {
        $S_Resp = Invoke-MgGraphRequest -Method GET -Uri $S_Uri -ErrorAction Stop
        if ($S_Resp.value)
        {
            foreach ($d in $S_Resp.value)
            {
                $S_Devices.Add([pscustomobject]$d) | Out-Null
            }
        }
        $S_Uri = $S_Resp.'@odata.nextLink'
    } while ($S_Uri)
    Write-Host ("  Retrieved {0} managed devices" -f $S_Devices.Count) -ForegroundColor Green

    # --- Active vs stale split based on check-in threshold ---
    $S_Now = Get-Date
    $S_ActiveCutoff = $S_Now.AddDays( - [int]$S_CheckinThresholdDays)
    $S_ActiveStates = [ordered]@{
        compliant     = 0
        noncompliant  = 0
        ingraceperiod = 0
        error         = 0
        conflict      = 0
        notapplicable = 0
        configmanager = 0
        unknown       = 0
    }
    $S_StaleCount = 0
    foreach ($d in $S_Devices)
    {
        $S_Last = if ($d.lastSyncDateTime)
        {
            [datetime]$d.lastSyncDateTime
        }
        else
        {
            $null
        }
        if (-not $S_Last -or $S_Last -lt $S_ActiveCutoff)
        {
            $S_StaleCount++; continue
        }
        $S_Key = if ($d.complianceState)
        {
            ([string]$d.complianceState).ToLowerInvariant()
        }
        else
        {
            'unknown'
        }
        if (-not $S_ActiveStates.Contains($S_Key))
        {
            $S_ActiveStates[$S_Key] = 0
        }
        $S_ActiveStates[$S_Key]++
    }
    $S_ActiveTotal = ($S_ActiveStates.Values | Measure-Object -Sum).Sum

    # --- Compliance policies + status overview ---
    Write-Host "Fetching compliance policies..." -ForegroundColor Cyan
    $S_Policies = New-Object System.Collections.Generic.List[object]
    $S_PUri = 'https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies?$expand=assignments'
    do
    {
        $S_Resp = Invoke-MgGraphRequest -Method GET -Uri $S_PUri -ErrorAction Stop
        if ($S_Resp.value)
        {
            foreach ($p in $S_Resp.value)
            {
                $S_Policies.Add([pscustomobject]$p) | Out-Null
            }
        }
        $S_PUri = $S_Resp.'@odata.nextLink'
    } while ($S_PUri)
    Write-Host ("  Retrieved {0} compliance policies" -f $S_Policies.Count) -ForegroundColor Green

    $S_OsTypeMap = @{
        '#microsoft.graph.windows10CompliancePolicy'          = 'Windows 10/11'
        '#microsoft.graph.windowsPhone81CompliancePolicy'     = 'Windows Phone'
        '#microsoft.graph.windows81CompliancePolicy'          = 'Windows 8.1'
        '#microsoft.graph.iosCompliancePolicy'                = 'iOS / iPadOS'
        '#microsoft.graph.macOSCompliancePolicy'              = 'macOS'
        '#microsoft.graph.androidCompliancePolicy'            = 'Android (Device Admin)'
        '#microsoft.graph.androidWorkProfileCompliancePolicy' = 'Android Work Profile'
        '#microsoft.graph.androidDeviceOwnerCompliancePolicy' = 'Android (Device Owner)'
        '#microsoft.graph.aospDeviceOwnerCompliancePolicy'    = 'AOSP (Device Owner)'
        '#microsoft.graph.androidForWorkCompliancePolicy'     = 'Android for Work'
        '#microsoft.graph.linuxCompliancePolicy'              = 'Linux'
    }

    $S_PolicyReport = New-Object System.Collections.Generic.List[object]
    $S_DeviceStatusRows = New-Object System.Collections.Generic.List[object]

    # Iterate every policy and pull per-device compliance from the v2 Intune
    # Reports endpoint (getCompliancePolicyDevicesReport). This is the same
    # data source the Intune portal uses for the Compliant / Noncompliant /
    # Others tile, so totals match the UI exactly. The endpoint returns a
    # columnar { Schema:[{Column,PropertyType}], Values:[[...]] } payload.
    $S_PolicyIndex = 0
    foreach ($pol in $S_Policies)
    {
        $S_PolicyIndex++
        Write-Host ("  [{0}/{1}] {2}" -f $S_PolicyIndex, $S_Policies.Count, $pol.displayName) -ForegroundColor DarkGray

        $S_Odata = $pol.'@odata.type'
        $S_Os = if ($S_Odata -and $S_OsTypeMap.ContainsKey($S_Odata))
        {
            $S_OsTypeMap[$S_Odata]
        }
        else
        {
            ($S_Odata -replace '#microsoft\.graph\.', '')
        }

        $S_StatusCounts = [ordered]@{
            Compliant               = 0
            Noncompliant            = 0
            InGracePeriod           = 0
            Conflict                = 0
            Error                   = 0
            NotApplicable           = 0
            ConfigManager           = 0
            NotEvaluated            = 0
            RemediatedNoncompliance = 0
            Unknown                 = 0
        }

        $S_Top = 1000
        $S_Skip = 0
        $S_KeepFetching = $true
        while ($S_KeepFetching)
        {
            $S_Body = @{
                filter = "(PolicyId eq '$($pol.id)')"
                skip   = $S_Skip
                top    = $S_Top
                select = @('DeviceId', 'DeviceName', 'UPN', 'UserEmail', 'UserName', 'OS', 'OSDescription', 'OSVersion', 'OwnerType', 'LastContact', 'ComplianceState', 'PolicyId', 'PolicyName', 'PolicyPlatformType', 'ReportStatus', 'DeviceModel', 'DeviceType', 'IMEI')
            } | ConvertTo-Json -Depth 5 -Compress

            $S_Resp = $null
            try
            {
                $S_Resp = Invoke-MgGraphRequest -Method POST `
                    -Uri 'https://graph.microsoft.com/beta/deviceManagement/reports/getCompliancePolicyDevicesReport' `
                    -ContentType 'application/json' `
                    -Body $S_Body `
                    -ErrorAction Stop
            }
            catch
            {
                Write-Warning ("getCompliancePolicyDevicesReport failed for {0}: {1}" -f $pol.displayName, $_.Exception.Message)
                break
            }

            # Response can come back as a hashtable or a JSON byte stream depending on module version.
            if ($S_Resp -is [byte[]])
            {
                $S_Resp = [System.Text.Encoding]::UTF8.GetString($S_Resp) | ConvertFrom-Json
            }

            $S_Schema = $S_Resp.Schema
            $S_Values = $S_Resp.Values
            if (-not $S_Schema -or -not $S_Values -or $S_Values.Count -eq 0)
            {
                break
            }

            # Build column-name -> index map for this page
            $S_ColIdx = @{}
            for ($S_I = 0; $S_I -lt $S_Schema.Count; $S_I++)
            {
                $S_Cname = if ($S_Schema[$S_I].Column)
                {
                    [string]$S_Schema[$S_I].Column
                }
                elseif ($S_Schema[$S_I].PropertyName)
                {
                    [string]$S_Schema[$S_I].PropertyName
                }
                else
                {
                    $null
                }
                if ($S_Cname)
                {
                    $S_ColIdx[$S_Cname] = $S_I
                }
            }
            $S_IxState = $S_ColIdx['ComplianceState']
            $S_IxDevice = $S_ColIdx['DeviceName']
            $S_IxUpn = $S_ColIdx['UPN']
            $S_IxOs = $S_ColIdx['OS']
            $S_IxOsVer = $S_ColIdx['OSVersion']
            $S_IxOwner = $S_ColIdx['OwnerType']
            $S_IxLast = $S_ColIdx['LastContact']
            $S_IxDevId = $S_ColIdx['DeviceId']
            $S_IxModel = $S_ColIdx['DeviceModel']

            foreach ($row in $S_Values)
            {
                $S_RawState = if ($null -ne $S_IxState)
                {
                    [string]$row[$S_IxState]
                }
                else
                {
                    'Unknown'
                }
                if ([string]::IsNullOrWhiteSpace($S_RawState))
                {
                    $S_RawState = 'Unknown'
                }

                # Normalize to a canonical key used in $statusCounts
                $S_Key = switch -Regex ($S_RawState)
                {
                    '^(?i)compliant$' { 'Compliant'; break }
                    '^(?i)non[- ]?compliant$' { 'Noncompliant'; break }
                    '^(?i)inGracePeriod$' { 'InGracePeriod'; break }
                    '^(?i)in[- ]?grace[- ]?period$' { 'InGracePeriod'; break }
                    '^(?i)conflict$' { 'Conflict'; break }
                    '^(?i)error$' { 'Error'; break }
                    '^(?i)not[- ]?applicable$' { 'NotApplicable'; break }
                    '^(?i)configManager$' { 'ConfigManager'; break }
                    '^(?i)not[- ]?evaluated$' { 'NotEvaluated'; break }
                    '^(?i)remediated.*' { 'RemediatedNoncompliance'; break }
                    default { 'Unknown' }
                }
                if (-not $S_StatusCounts.Contains($S_Key))
                {
                    $S_StatusCounts[$S_Key] = 0
                }
                $S_StatusCounts[$S_Key]++

                if ($IncludeDeviceStatusDetails)
                {
                    $S_DeviceStatusRows.Add([pscustomobject]@{
                            PolicyName        = $pol.displayName
                            PolicyOS          = $S_Os
                            DeviceId          = if ($null -ne $S_IxDevId)
                            {
                                $row[$S_IxDevId]
                            }
                            else
                            {
                                ''
                            }
                            DeviceDisplayName = if ($null -ne $S_IxDevice)
                            {
                                $row[$S_IxDevice]
                            }
                            else
                            {
                                ''
                            }
                            UserPrincipalName = if ($null -ne $S_IxUpn)
                            {
                                $row[$S_IxUpn]
                            }
                            else
                            {
                                ''
                            }
                            OS                = if ($null -ne $S_IxOs)
                            {
                                $row[$S_IxOs]
                            }
                            else
                            {
                                ''
                            }
                            OSVersion         = if ($null -ne $S_IxOsVer)
                            {
                                $row[$S_IxOsVer]
                            }
                            else
                            {
                                ''
                            }
                            OwnerType         = if ($null -ne $S_IxOwner)
                            {
                                $row[$S_IxOwner]
                            }
                            else
                            {
                                ''
                            }
                            LastContact       = if ($null -ne $S_IxLast)
                            {
                                $row[$S_IxLast]
                            }
                            else
                            {
                                ''
                            }
                            DeviceModel       = if ($null -ne $S_IxModel)
                            {
                                $row[$S_IxModel]
                            }
                            else
                            {
                                ''
                            }
                            ComplianceState   = $S_RawState
                            MappedBucket      = $S_Key
                        }) | Out-Null
                }
            }

            if ($S_Values.Count -lt $S_Top)
            {
                $S_KeepFetching = $false
            }
            else
            {
                $S_Skip += $S_Top
            }
        }

        # Map raw statuses to Intune portal tiles:
        #   Compliant  = Compliant + InGracePeriod (portal counts grace as compliant in the bar)
        #   Noncompliant = Noncompliant + RemediatedNoncompliance
        #   Others     = Conflict + Error + NotApplicable + ConfigManager + NotEvaluated + Unknown
        $S_Compliant = [int]$S_StatusCounts['Compliant'] + [int]$S_StatusCounts['InGracePeriod']
        $S_NonCompliant = [int]$S_StatusCounts['Noncompliant'] + [int]$S_StatusCounts['RemediatedNoncompliance']
        $S_Others = [int]$S_StatusCounts['Conflict'] + [int]$S_StatusCounts['Error'] + [int]$S_StatusCounts['NotApplicable'] + [int]$S_StatusCounts['ConfigManager'] + [int]$S_StatusCounts['NotEvaluated'] + [int]$S_StatusCounts['Unknown']
        $S_Total = $S_Compliant + $S_NonCompliant + $S_Others
        $S_PctCompliant = if (($S_Compliant + $S_NonCompliant) -gt 0)
        {
            [math]::Round(($S_Compliant / ($S_Compliant + $S_NonCompliant)) * 100, 1)
        }
        else
        {
            $null
        }

        $S_AssignedGroupCount = if ($pol.assignments)
        {
            @($pol.assignments).Count
        }
        else
        {
            0
        }

        $S_PolicyReport.Add([pscustomobject]@{
                DisplayName          = $pol.displayName
                OperatingSystem      = $S_Os
                OdataType            = $S_Odata
                AssignmentCount      = $S_AssignedGroupCount
                Compliant            = $S_Compliant
                NonCompliant         = $S_NonCompliant
                Others               = $S_Others
                RawCompliant         = [int]$S_StatusCounts['Compliant']
                InGracePeriod        = [int]$S_StatusCounts['InGracePeriod']
                Remediated           = [int]$S_StatusCounts['RemediatedNoncompliance']
                InError              = [int]$S_StatusCounts['Error']
                Conflict             = [int]$S_StatusCounts['Conflict']
                NotApplicable        = [int]$S_StatusCounts['NotApplicable']
                NotEvaluated         = [int]$S_StatusCounts['NotEvaluated']
                ConfigManager        = [int]$S_StatusCounts['ConfigManager']
                Unknown              = [int]$S_StatusCounts['Unknown']
                TotalReporting       = $S_Total
                PercentCompliant     = $S_PctCompliant
                LastModifiedDateTime = $pol.lastModifiedDateTime
                Id                   = $pol.id
            }) | Out-Null
    }

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
    $S_PolicyCsv = Join-Path $S_ReportFolder ("ReportIntuneCompliance_Policies_{0}.csv" -f $S_Timestamp)
    $S_DeviceStatusCsv = Join-Path $S_ReportFolder ("ReportIntuneCompliance_DeviceStatuses_{0}.csv" -f $S_Timestamp)
    $S_HtmlFile = Join-Path $S_ReportFolder ("ReportIntuneCompliance_{0}.html" -f $S_Timestamp)

    # --- CSV exports ---
    $S_PolicyReport | Sort-Object OperatingSystem, DisplayName | Export-Csv -Path $S_PolicyCsv -NoTypeInformation -Encoding UTF8
    if ($IncludeDeviceStatusDetails -and $S_DeviceStatusRows.Count -gt 0)
    {
        $S_DeviceStatusRows | Sort-Object PolicyName, DeviceDisplayName | Export-Csv -Path $S_DeviceStatusCsv -NoTypeInformation -Encoding UTF8
    }

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

    # Banner state
    $S_SecureByDefaultDisp = if ($null -eq $S_SecureByDefault)
    {
        'Unknown'
    }
    elseif ($S_SecureByDefault)
    {
        'On'
    }
    else
    {
        'Off'
    }
    $S_BannerClass = if ($S_SecureByDefault -eq $true)
    {
        'banner banner-good'
    }
    else
    {
        'banner banner-bad'
    }
    $S_BannerTitle = if ($S_SecureByDefault -eq $true)
    {
        'Tenant compliance settings look good'
    }
    elseif ($S_SecureByDefault -eq $false)
    {
        'Action required: Devices without a compliance policy are being marked Compliant'
    }
    else
    {
        'Tenant compliance settings could not be read'
    }
    $S_BannerBody = if ($S_SecureByDefault -eq $true)
    {
        "Mark devices with no compliance policy assigned as Not compliant is currently <strong>On</strong>. Devices that have not checked in for <strong>$S_CheckinThresholdDays</strong> days will be marked as Not compliant."
    }
    elseif ($S_SecureByDefault -eq $false)
    {
        "Mark devices with no compliance policy assigned as is currently <strong>Compliant</strong> (insecure default). Change this to <strong>Not compliant</strong> in Intune > Endpoint security > Device compliance > Compliance policy settings. Current check-in threshold is <strong>$S_CheckinThresholdDays</strong> days."
    }
    else
    {
        "Could not read deviceManagement/settings. Verify the signed-in account has the required Graph scopes."
    }

    function ConvertTo-PieJson
    {
        param([System.Collections.IDictionary]$Map)
        $F_Entries = @()
        foreach ($F_Key in $Map.Keys)
        {
            if ([int]$Map[$F_Key] -gt 0)
            {
                $F_Entries += [pscustomobject]@{ Label = $F_Key; Value = [int]$Map[$F_Key] }
            }
        }
        if (-not $F_Entries -or $F_Entries.Count -eq 0)
        {
            return '{"labels":[],"data":[]}'
        }
        $F_Labels = ($F_Entries | ForEach-Object { '"' + $_.Label + '"' }) -join ','
        $F_Data = ($F_Entries | ForEach-Object { $_.Value }) -join ','
        "{`"labels`":[$F_Labels],`"data`":[$F_Data]}"
    }

    # Tenant pie data
    $S_TenantPieJson = ConvertTo-PieJson -Map $S_TenantStates
    $S_ActivePieJson = ConvertTo-PieJson -Map $S_ActiveStates

    # Policy table rows
    $S_PolicyRows = ($S_PolicyReport | Sort-Object OperatingSystem, DisplayName | ForEach-Object {
            $S_Pct = if ($null -ne $_.PercentCompliant)
            {
                ('{0}%' -f $_.PercentCompliant)
            }
            else
            {
                '-'
            }
            $S_Lm = if ($_.LastModifiedDateTime)
            {
                ([datetime]$_.LastModifiedDateTime).ToString('dd MMM yyyy')
            }
            else
            {
                '-'
            }
            $S_NcClass = if ($_.NonCompliant -gt 0)
            {
                'cell-bad'
            }
            else
            {
                ''
            }
            $S_CClass = if ($_.Compliant -gt 0)
            {
                'cell-good'
            }
            else
            {
                ''
            }
            $S_OthersTitle = ("Error: {0}, Conflict: {1}, Not Applicable: {2}, Not Evaluated: {3}, ConfigManager: {4}, Unknown: {5}" -f $_.InError, $_.Conflict, $_.NotApplicable, $_.NotEvaluated, $_.ConfigManager, $_.Unknown)
            $S_RowAttr = "data-os=`"$(& $S_Enc $_.OperatingSystem)`""
            "<tr $S_RowAttr>" +
            "<td>$(& $S_Enc $_.DisplayName)</td>" +
            "<td>$(& $S_Enc $_.OperatingSystem)</td>" +
            "<td>$($_.AssignmentCount)</td>" +
            "<td class='$S_CClass' title='Compliant: $($_.RawCompliant), In Grace Period: $($_.InGracePeriod)'>$($_.Compliant)</td>" +
            "<td class='$S_NcClass' title='Noncompliant: $([int]$_.NonCompliant - [int]$_.Remediated), Remediated: $($_.Remediated)'>$($_.NonCompliant)</td>" +
            "<td title='$S_OthersTitle'>$($_.Others)</td>" +
            "<td>$($_.TotalReporting)</td>" +
            "<td>$S_Pct</td>" +
            "<td>$S_Lm</td>" +
            "</tr>"
        }) -join "`n"

    # OS filter options
    $S_OsOptions = ($S_PolicyReport | Select-Object -ExpandProperty OperatingSystem -Unique | Sort-Object | ForEach-Object {
            "<option value=`"$(& $S_Enc $_)`">$(& $S_Enc $_)</option>"
        }) -join "`n"

    $S_DrilldownNote = if ($IncludeDeviceStatusDetails -and $S_DeviceStatusRows.Count -gt 0)
    {
        "<p style='font-size:0.85em;color:#555;margin-top:8px;'>Per-device drill-down exported to <code>$(& $S_Enc (Split-Path $S_DeviceStatusCsv -Leaf))</code> ($($S_DeviceStatusRows.Count) rows)</p>"
    }
    elseif ($IncludeDeviceStatusDetails)
    {
        "<p style='font-size:0.85em;color:#555;margin-top:8px;'>No per-device status rows returned.</p>"
    }
    else
    {
        "<p style='font-size:0.85em;color:#777;margin-top:8px;'>Re-run with <code>-IncludeDeviceStatusDetails</code> to also produce a per-device-per-policy CSV drill-down.</p>"
    }

    $S_Html = @"
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
    <p>Tenant: $(& $S_Enc $S_TenantDisplayName) ($S_TenantId) &nbsp;|&nbsp; Generated: $S_ReportDate</p>
  </div>
  <div style="text-align:right;font-size:0.9em;opacity:0.9;">
    Secure by default: <strong>$S_SecureByDefaultDisp</strong><br/>
    Check-in threshold: <strong>$S_CheckinThresholdDays days</strong>
  </div>
</div>

<!-- TENANT COMPLIANCE SETTINGS BANNER -->
<div class="$S_BannerClass">
  <div class="icon"></div>
  <div>
    <h2>$S_BannerTitle</h2>
    <p>$S_BannerBody</p>
  </div>
</div>

<!-- OVERVIEW CARDS -->
<div class="section-title">Overview</div>
<div class="summary-cards">
  <div class="card"><div class="label">Total Reporting Devices</div><div class="value" style="color:#1a1a2e;">$S_TenantTotal</div></div>
  <div class="card"><div class="label">Compliant</div><div class="value" style="color:#27ae60;">$($S_TenantStates.Compliant)</div></div>
  <div class="card"><div class="label">Non-compliant</div><div class="value" style="color:#e74c3c;">$($S_TenantStates.NonCompliant)</div></div>
  <div class="card"><div class="label">In Grace Period</div><div class="value" style="color:#f39c12;">$($S_TenantStates.InGracePeriod)</div></div>
  <div class="card"><div class="label">Active Devices (last $S_CheckinThresholdDays days)</div><div class="value" style="color:#3498db;">$S_ActiveTotal</div><div class="sub">$S_StaleCount stale / not checked in</div></div>
  <div class="card"><div class="label">Total Policies</div><div class="value" style="color:#9b59b6;">$($S_PolicyReport.Count)</div></div>
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
    <div class="subtitle">Only devices that checked in within the last $S_CheckinThresholdDays days ($S_StaleCount stale devices excluded)</div>
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
$S_OsOptions
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
$S_PolicyRows
    </tbody>
  </table>
  $S_DrilldownNote
</div>

<div class="footer">Report generated by ReportIntuneCompliance.ps1</div>

<script>
var tenantPieData = $S_TenantPieJson;
var activePieData = $S_ActivePieJson;

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

    $S_Html | Out-File -FilePath $S_HtmlFile -Encoding UTF8

    # --- Console summary ---
    Write-Host ""
    Write-Host "Intune Compliance Report" -ForegroundColor Cyan
    Write-Host "--------------------------------------------"
    Write-Host ("Tenant                    : {0} ({1})" -f $S_TenantDisplayName, $S_TenantId)
    $S_SecureByDefaultColour = if ($S_SecureByDefault -eq $true)
    {
        'Green'
    }
    else
    {
        'Red'
    }
    Write-Host ("secureByDefault           : {0}" -f $S_SecureByDefaultDisp) -ForegroundColor $S_SecureByDefaultColour
    Write-Host ("Check-in threshold (days) : {0}" -f $S_CheckinThresholdDays)
    Write-Host ("Total reporting devices   : {0}" -f $S_TenantTotal)
    Write-Host ("  Compliant               : {0}" -f $S_TenantStates.Compliant) -ForegroundColor Green
    Write-Host ("  Non-compliant           : {0}" -f $S_TenantStates.NonCompliant) -ForegroundColor Red
    Write-Host ("  In Grace Period         : {0}" -f $S_TenantStates.InGracePeriod) -ForegroundColor Yellow
    Write-Host ("  Error                   : {0}" -f $S_TenantStates.Error)
    Write-Host ("  Conflict                : {0}" -f $S_TenantStates.Conflict)
    Write-Host ("  Not Applicable          : {0}" -f $S_TenantStates.NotApplicable)
    Write-Host ("Active devices            : {0}  (Stale: {1})" -f $S_ActiveTotal, $S_StaleCount)
    Write-Host ("Total compliance policies : {0}" -f $S_PolicyReport.Count)
    Write-Host ""
    Write-Host ("CSV (policies)            : {0}" -f $S_PolicyCsv) -ForegroundColor Yellow
    if ($IncludeDeviceStatusDetails -and $S_DeviceStatusRows.Count -gt 0)
    {
        Write-Host ("CSV (device statuses)     : {0}" -f $S_DeviceStatusCsv) -ForegroundColor Yellow
    }
    Write-Host ("HTML report               : {0}" -f $S_HtmlFile) -ForegroundColor Yellow

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
