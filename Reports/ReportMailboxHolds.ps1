#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Reports on the holds and retention policies that cover Exchange Online mailboxes.

.DESCRIPTION
    Connects to Exchange Online (and, unless -SkipPolicyResolution is used, to the
    Security & Compliance PowerShell endpoint via Connect-IPPSSession) and reports the
    hold state of each mailbox: Litigation Hold, Microsoft 365 (Purview) retention
    policies, eDiscovery case holds, legacy In-Place Holds, retention-label holds
    (ComplianceTagHoldApplied) and delay holds.

    Each mailbox's InPlaceHolds GUIDs are decoded (prefix + action suffix) and matched
    against Get-RetentionCompliancePolicy, Get-CaseHoldPolicy and
    Get-AppRetentionCompliancePolicy so the report shows the friendly policy NAME that
    covers each mailbox, and (in the policy-centric view) the mailboxes each policy covers.

    Results are exported to CSV and to an interactive HTML report with two views:
    a mailbox-centric table and a policy-centric coverage table.

.PARAMETER MailboxScope
    Active : active user mailboxes only (default).
    All    : active mailboxes plus inactive mailboxes and soft-deleted mailboxes.

.PARAMETER TestMode
    Processes a random sample (see -SampleSize, default 10) instead of every mailbox.
    Use this to validate the report format before running against the full tenant.

.PARAMETER SampleSize
    Number of mailboxes sampled when -TestMode is specified. Defaults to 10.

.PARAMETER ReportPath
    Folder or file path for the output report. If a folder is specified, a timestamped
    filename is generated automatically. Defaults to the current directory.

.PARAMETER SkipPolicyResolution
    Skips the Security & Compliance (Connect-IPPSSession) connection. Hold GUIDs are still
    decoded by type but are shown as raw GUIDs instead of being resolved to policy names.

.EXAMPLE
    .\ReportMailboxHolds.ps1

.EXAMPLE
    .\ReportMailboxHolds.ps1 -TestMode

.EXAMPLE
    .\ReportMailboxHolds.ps1 -MailboxScope All -ReportPath "C:\Reports"

.EXAMPLE
    .\ReportMailboxHolds.ps1 -SkipPolicyResolution

.NOTES
    Hold decoding follows Microsoft Learn "Identify Exchange mailbox hold types in
    eDiscovery" (https://learn.microsoft.com/purview/edisc-hold-types-mailboxes).
    Read-only: the script only runs Get-* / Connect-* cmdlets.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('Active', 'All')]
    [string]$MailboxScope = 'Active',

    [Parameter(Mandatory = $false)]
    [switch]$TestMode,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 500)]
    [int]$SampleSize = 10,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ReportPath,

    [Parameter(Mandatory = $false)]
    [switch]$SkipPolicyResolution
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable))
{
    throw "ExchangeOnlineManagement module is not installed. Install it using: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

# Reduce any hold value / policy identity to a canonical 32-char lowercase GUID
# so InPlaceHolds entries and policy objects can be keyed against each other.
function ConvertTo-NormalisedGuid
{
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $S_Hex = ($Value -replace '[^0-9a-fA-F]', '')
    if ($S_Hex.Length -lt 32) { return $null }
    try
    {
        return ([guid]$S_Hex.Substring(0, 32)).ToString('N').ToLowerInvariant()
    }
    catch
    {
        return $null
    }
}

# Summarise a policy's deployment/distribution state from -DistributionDetail output.
# DistributionStatus is the headline (Success = fully distributed); DistributionResults
# carries per-location error detail when it is not Success.
function Get-PolicyDistribution
{
    param($Policy)

    $S_Status = "$($Policy.DistributionStatus)".Trim()
    $S_Detail = $null
    if ($S_Status -and $S_Status -ne 'Success')
    {
        try
        {
            $S_Parts = @()
            foreach ($S_Res in @($Policy.DistributionResults))
            {
                if (-not $S_Res) { continue }
                $S_Loc = "$($S_Res.Location)".Trim()
                $S_St  = "$($S_Res.Status)".Trim()
                if ($S_St -and $S_St -ne 'Success')
                {
                    $S_Parts += ((@($S_Loc, $S_St) | Where-Object { $_ }) -join ': ')
                }
            }
            if ($S_Parts.Count -gt 0) { $S_Detail = ($S_Parts | Select-Object -Unique) -join '; ' }
        }
        catch { }
    }
    [PSCustomObject]@{ Status = $S_Status; Detail = $S_Detail }
}

# Decode a single InPlaceHolds entry into a structured hold record and resolve its name.
function Resolve-HoldEntry
{
    param(
        [string]$Raw,
        [hashtable]$RetentionHash,
        [hashtable]$CaseHash,
        [hashtable]$AppHash
    )

    $S_Value = $Raw.Trim()
    $S_Excluded = $false
    if ($S_Value.StartsWith('-'))
    {
        $S_Excluded = $true
        $S_Value = $S_Value.Substring(1)
    }

    # Action suffix (:1 delete/label-publish, :2 hold-only, :3 retain-then-delete)
    $S_Action = $null
    if ($S_Value -match ':(\d+)\s*$')
    {
        switch ($Matches[1])
        {
            '1'     { $S_Action = 'Delete / label' }
            '2'     { $S_Action = 'Hold only' }
            '3'     { $S_Action = 'Retain then delete' }
            default { $S_Action = "Action $($Matches[1])" }
        }
        $S_Value = $S_Value -replace ':\d+\s*$', ''
    }

    # Prefix -> hold type + scope
    $S_Type = 'LegacyInPlaceHold'
    $S_Scope = $null
    $S_GuidPart = $S_Value
    if ($S_Value -match '^UniH')     { $S_Type = 'eDiscoveryCaseHold'; $S_GuidPart = $S_Value.Substring(4) }
    elseif ($S_Value -match '^cld')  { $S_Type = 'eDiscoveryCloudHold'; $S_GuidPart = $S_Value.Substring(3) }
    elseif ($S_Value -match '^mbx')  { $S_Type = 'RetentionPolicy'; $S_Scope = 'Exchange mailbox'; $S_GuidPart = $S_Value.Substring(3) }
    elseif ($S_Value -match '^skp')  { $S_Type = 'RetentionPolicy'; $S_Scope = 'Skype / Teams'; $S_GuidPart = $S_Value.Substring(3) }
    elseif ($S_Value -match '^grp')  { $S_Type = 'RetentionPolicy'; $S_Scope = 'M365 Group'; $S_GuidPart = $S_Value.Substring(3) }

    $S_Norm = ConvertTo-NormalisedGuid $S_GuidPart

    # Resolve name. Try the type's natural table first, then fall back to the others.
    $S_Name = $null
    $S_DistStatus = $null
    $S_DistDetail = $null
    if ($S_Norm)
    {
        $S_Order = switch ($S_Type)
        {
            'eDiscoveryCaseHold'  { @($CaseHash, $RetentionHash, $AppHash) }
            'eDiscoveryCloudHold' { @($CaseHash, $RetentionHash, $AppHash) }
            'RetentionPolicy'     { @($RetentionHash, $AppHash, $CaseHash) }
            default               { @($RetentionHash, $CaseHash, $AppHash) }
        }
        foreach ($S_Table in $S_Order)
        {
            if ($S_Table -and $S_Table.ContainsKey($S_Norm))
            {
                $S_Match = $S_Table[$S_Norm]
                $S_Name = $S_Match.Name
                if ($S_Match.Type -and $S_Type -eq 'LegacyInPlaceHold') { $S_Type = $S_Match.Type }
                if (-not $S_Scope -and $S_Match.Scope) { $S_Scope = $S_Match.Scope }
                $S_DistStatus = $S_Match.DistributionStatus
                $S_DistDetail = $S_Match.DistributionDetail
                break
            }
        }
    }

    if (-not $S_Name)
    {
        $S_Name = if ($S_Norm) { "(unresolved) $S_Norm" } else { $Raw }
    }

    [PSCustomObject]@{
        Type               = $S_Type
        Scope              = $S_Scope
        Name               = $S_Name
        Guid               = $S_Norm
        Excluded           = $S_Excluded
        Action             = $S_Action
        Raw                = $Raw
        OrgWide            = $false
        DistributionStatus = $S_DistStatus
        DistributionDetail = $S_DistDetail
    }
}

# ---------------------------------------------------------------------------
# Connect to Exchange Online
# ---------------------------------------------------------------------------
$S_ExoSession = Get-ConnectionInformation -ErrorAction SilentlyContinue |
    Where-Object { $_.State -eq 'Connected' -and $_.ConnectionUri -notmatch 'compliance|protection' } |
    Select-Object -First 1
if ($S_ExoSession)
{
    Write-Host "Existing Exchange Online session detected:" -ForegroundColor Yellow
    Write-Host "  Account   : $($S_ExoSession.UserPrincipalName)" -ForegroundColor Yellow
    Write-Host "  TenantId  : $($S_ExoSession.TenantID)" -ForegroundColor Yellow
    Write-Host ""
    $S_Choice = Read-Host "Use existing Exchange Online session? [Y] Yes  [N] Disconnect and reconnect  (Default: Y)"
    if ($S_Choice -eq 'N')
    {
        Disconnect-ExchangeOnline -Confirm:$false
        Connect-ExchangeOnline -ShowBanner:$false
    }
}
else
{
    Connect-ExchangeOnline -ShowBanner:$false
}

# --- Tenant info ---
$S_TenantDisplayName = $null
$S_TenantId = $null
try
{
    $S_ConnInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue |
        Where-Object { $_.State -eq 'Connected' -and $_.ConnectionUri -notmatch 'compliance|protection' } |
        Select-Object -First 1
    if ($S_ConnInfo)
    {
        $S_TenantDisplayName = $S_ConnInfo.Organization
        $S_TenantId = $S_ConnInfo.TenantID
    }
}
catch { }
if (-not $S_TenantDisplayName) { $S_TenantDisplayName = 'Exchange Online' }
if (-not $S_TenantId) { $S_TenantId = 'Unknown' }

# ---------------------------------------------------------------------------
# Connect to Security & Compliance PowerShell (for policy-name resolution)
# ---------------------------------------------------------------------------
$S_PolicyResolutionEnabled = -not $SkipPolicyResolution
if ($S_PolicyResolutionEnabled)
{
    $S_IppsSession = Get-ConnectionInformation -ErrorAction SilentlyContinue |
        Where-Object { $_.State -eq 'Connected' -and $_.ConnectionUri -match 'compliance|protection' } |
        Select-Object -First 1
    try
    {
        if ($S_IppsSession)
        {
            Write-Host ""
            Write-Host "Existing Security & Compliance session detected:" -ForegroundColor Yellow
            Write-Host "  Account   : $($S_IppsSession.UserPrincipalName)" -ForegroundColor Yellow
            Write-Host ""
            $S_Choice = Read-Host "Use existing Security & Compliance session? [Y] Yes  [N] Disconnect and reconnect  (Default: Y)"
            if ($S_Choice -eq 'N')
            {
                Disconnect-ExchangeOnline -Confirm:$false
                Connect-IPPSSession -ShowBanner:$false
            }
        }
        else
        {
            Write-Host ""
            Write-Host "Connecting to Security & Compliance PowerShell (Connect-IPPSSession) for policy-name resolution..." -ForegroundColor Cyan
            Connect-IPPSSession -ShowBanner:$false
        }
    }
    catch
    {
        Write-Warning "Could not connect to Security & Compliance PowerShell: $($_.Exception.Message)"
        Write-Warning "Continuing without policy-name resolution; hold GUIDs will be shown unresolved."
        $S_PolicyResolutionEnabled = $false
    }
}
else
{
    Write-Host ""
    Write-Host "SkipPolicyResolution specified: hold GUIDs will not be resolved to policy names." -ForegroundColor Yellow
}

# ---------------------------------------------------------------------------
# Build policy lookup tables from Security & Compliance
# ---------------------------------------------------------------------------
$S_RetentionHash = @{}
$S_CaseHash = @{}
$S_AppHash = @{}
$S_OrgWidePolicies = [System.Collections.Generic.List[PSCustomObject]]::new()

if ($S_PolicyResolutionEnabled)
{
    Write-Host "Retrieving retention & hold policies from Security & Compliance..." -ForegroundColor Cyan

    try
    {
        foreach ($S_Pol in @(Get-RetentionCompliancePolicy -DistributionDetail -ErrorAction Stop))
        {
            $S_Key = ConvertTo-NormalisedGuid $S_Pol.Guid.ToString()
            $S_Dist = Get-PolicyDistribution $S_Pol
            if ($S_Key) { $S_RetentionHash[$S_Key] = [PSCustomObject]@{ Name = $S_Pol.Name; Type = 'RetentionPolicy'; Scope = $null; DistributionStatus = $S_Dist.Status; DistributionDetail = $S_Dist.Detail } }
        }
        Write-Host "  Retention policies : $($S_RetentionHash.Count)" -ForegroundColor Green
    }
    catch { Write-Warning "  Get-RetentionCompliancePolicy failed: $($_.Exception.Message)" }

    try
    {
        # Get-CaseHoldPolicy must be scoped to a case in most tenants; enumerate all
        # eDiscovery cases and collect their hold policies. Fall back to a bare call.
        $S_CasePolicies = [System.Collections.Generic.List[object]]::new()
        $S_Cases = @()
        try { $S_Cases = @(Get-ComplianceCase -ErrorAction Stop) } catch { }
        foreach ($S_Case in $S_Cases)
        {
            try
            {
                foreach ($S_Pol in @(Get-CaseHoldPolicy -Case $S_Case.Identity -ErrorAction Stop))
                {
                    $S_Pol | Add-Member -NotePropertyName '_CaseName' -NotePropertyValue $S_Case.Name -Force
                    $S_CasePolicies.Add($S_Pol)
                }
            }
            catch { }
        }
        if ($S_CasePolicies.Count -eq 0)
        {
            try { foreach ($S_Pol in @(Get-CaseHoldPolicy -ErrorAction Stop)) { $S_CasePolicies.Add($S_Pol) } } catch { }
        }
        foreach ($S_Pol in $S_CasePolicies)
        {
            $S_Key = ConvertTo-NormalisedGuid $S_Pol.Guid.ToString()
            $S_ScopeText = if ($S_Pol._CaseName) { "Case: $($S_Pol._CaseName)" } else { 'eDiscovery case' }
            $S_Dist = Get-PolicyDistribution $S_Pol
            if ($S_Key) { $S_CaseHash[$S_Key] = [PSCustomObject]@{ Name = $S_Pol.Name; Type = 'eDiscoveryCaseHold'; Scope = $S_ScopeText; DistributionStatus = $S_Dist.Status; DistributionDetail = $S_Dist.Detail } }
        }
        Write-Host "  eDiscovery holds   : $($S_CaseHash.Count)" -ForegroundColor Green
    }
    catch { Write-Warning "  Get-CaseHoldPolicy enumeration failed: $($_.Exception.Message)" }

    try
    {
        foreach ($S_Pol in @(Get-AppRetentionCompliancePolicy -DistributionDetail -ErrorAction Stop))
        {
            $S_Key = ConvertTo-NormalisedGuid $S_Pol.Guid.ToString()
            $S_Dist = Get-PolicyDistribution $S_Pol
            if ($S_Key) { $S_AppHash[$S_Key] = [PSCustomObject]@{ Name = $S_Pol.Name; Type = 'AppRetentionPolicy'; Scope = ($S_Pol.Applications -join ', '); DistributionStatus = $S_Dist.Status; DistributionDetail = $S_Dist.Detail } }
        }
        Write-Host "  App retention pol. : $($S_AppHash.Count)" -ForegroundColor Green
    }
    catch { Write-Warning "  Get-AppRetentionCompliancePolicy failed: $($_.Exception.Message)" }
}

# ---------------------------------------------------------------------------
# Organization-wide (entire-location) retention policies.
# Per Microsoft Learn, a retention policy scoped to the whole location is NOT
# stamped on each mailbox's InPlaceHolds; it only appears in Get-OrganizationConfig
# and applies to every mailbox except those that carry a -mbx<guid> exclusion.
# Get-OrganizationConfig is an Exchange Online cmdlet, so this runs regardless of
# whether policy-name resolution (Connect-IPPSSession) succeeded.
# ---------------------------------------------------------------------------
try
{
    $S_OrgConfig = Get-OrganizationConfig -ErrorAction Stop
    foreach ($S_Raw in @($S_OrgConfig.InPlaceHolds))
    {
        if ([string]::IsNullOrWhiteSpace($S_Raw)) { continue }
        $S_OrgWidePolicies.Add((Resolve-HoldEntry -Raw $S_Raw -RetentionHash $S_RetentionHash -CaseHash $S_CaseHash -AppHash $S_AppHash))
    }
    Write-Host "  Org-wide holds     : $($S_OrgWidePolicies.Count)" -ForegroundColor Green
}
catch { Write-Warning "  Get-OrganizationConfig failed: $($_.Exception.Message)" }

# Org-wide policies that cover USER mailboxes: mbx (Exchange mailbox / 1xN Teams chats)
# and skp (Skype). grp policies target group mailboxes, not user mailboxes, so they are
# excluded from per-mailbox coverage (still shown in the org-wide summary section).
$S_OrgWideMailboxPolicies = @(
    $S_OrgWidePolicies | Where-Object { $_.Guid -and -not $_.Excluded -and $_.Scope -ne 'M365 Group' }
)

# ---------------------------------------------------------------------------
# Retrieve mailboxes
# ---------------------------------------------------------------------------
$S_Props = @(
    'DisplayName', 'UserPrincipalName', 'PrimarySmtpAddress', 'RecipientTypeDetails',
    'IsInactiveMailbox', 'ExternalDirectoryObjectId', 'ExchangeGuid',
    'LitigationHoldEnabled', 'LitigationHoldDate', 'LitigationHoldOwner', 'LitigationHoldDuration',
    'RetentionHoldEnabled', 'RetentionPolicy', 'InPlaceHolds',
    'ComplianceTagHoldApplied', 'DelayHoldApplied'
)
# Note: DelayReleaseHoldApplied is not exposed by Get-ExoMailbox -Properties, so it is
# not collected here; DelayHoldApplied still captures the common Outlook/email delay hold.

$S_AllMailboxes = [System.Collections.Generic.List[object]]::new()
$S_SeenKeys = [System.Collections.Generic.HashSet[string]]::new()

function Add-Mailboxes
{
    param([array]$Mailboxes, [string]$State)

    foreach ($S_Mb in $Mailboxes)
    {
        $S_Key = if ($S_Mb.ExchangeGuid) { $S_Mb.ExchangeGuid.ToString() } elseif ($S_Mb.ExternalDirectoryObjectId) { $S_Mb.ExternalDirectoryObjectId.ToString() } else { $S_Mb.UserPrincipalName }
        if ($S_Key -and -not $S_SeenKeys.Add($S_Key)) { continue }
        $S_Mb | Add-Member -NotePropertyName 'MailboxState' -NotePropertyValue $State -Force
        $S_AllMailboxes.Add($S_Mb)
    }
}

Write-Host ""
Write-Host "Retrieving mailboxes (scope: $MailboxScope)..." -ForegroundColor Cyan

Add-Mailboxes -State 'Active' -Mailboxes @(
    Get-ExoMailbox -RecipientTypeDetails UserMailbox -Properties $S_Props -ResultSize Unlimited
)

if ($MailboxScope -eq 'All')
{
    try
    {
        Add-Mailboxes -State 'Inactive' -Mailboxes @(
            Get-ExoMailbox -InactiveMailboxOnly -Properties $S_Props -ResultSize Unlimited -ErrorAction Stop
        )
    }
    catch { Write-Warning "Could not retrieve inactive mailboxes: $($_.Exception.Message)" }

    try
    {
        Add-Mailboxes -State 'SoftDeleted' -Mailboxes @(
            Get-ExoMailbox -SoftDeletedMailbox -Properties $S_Props -ResultSize Unlimited -ErrorAction Stop
        )
    }
    catch { Write-Warning "Could not retrieve soft-deleted mailboxes: $($_.Exception.Message)" }
}

[array]$S_Mbx = $S_AllMailboxes
if ($TestMode)
{
    Write-Host "TEST MODE: sampling $SampleSize of $($S_Mbx.Count) mailboxes..." -ForegroundColor Cyan
    [array]$S_Mbx = $S_Mbx | Get-Random -Count ([math]::Min($SampleSize, $S_Mbx.Count))
}
Write-Host "  Found $($S_Mbx.Count) mailboxes to process" -ForegroundColor Green

# ---------------------------------------------------------------------------
# Build report data
# ---------------------------------------------------------------------------
$S_Report = [System.Collections.Generic.List[PSCustomObject]]::new()
$S_Counter = 0
foreach ($S_Mb in $S_Mbx)
{
    $S_Counter++
    Write-Host "[$S_Counter/$($S_Mbx.Count)] $($S_Mb.DisplayName)..." -ForegroundColor Gray

    try
    {
        $S_Holds = [System.Collections.Generic.List[PSCustomObject]]::new()

        if ($S_Mb.LitigationHoldEnabled)
        {
            $S_LitName = 'Litigation Hold'
            if ($S_Mb.LitigationHoldDuration -and "$($S_Mb.LitigationHoldDuration)" -ne 'Unlimited')
            {
                $S_LitName = "Litigation Hold ($($S_Mb.LitigationHoldDuration))"
            }
            $S_Holds.Add([PSCustomObject]@{ Type = 'LitigationHold'; Scope = 'Whole mailbox'; Name = $S_LitName; Guid = $null; Excluded = $false; Action = 'Hold only'; Raw = 'LitigationHoldEnabled'; OrgWide = $false; DistributionStatus = $null; DistributionDetail = $null })
        }

        foreach ($S_Raw in @($S_Mb.InPlaceHolds))
        {
            if ([string]::IsNullOrWhiteSpace($S_Raw)) { continue }
            $S_Holds.Add((Resolve-HoldEntry -Raw $S_Raw -RetentionHash $S_RetentionHash -CaseHash $S_CaseHash -AppHash $S_AppHash))
        }

        # Apply organization-wide (entire-location) retention policies. These are not
        # stamped on the mailbox, so add them here unless the mailbox is explicitly
        # excluded (-mbx<guid>) or already carries the policy as an explicit hold.
        $S_ExcludedGuids = [System.Collections.Generic.HashSet[string]]::new()
        $S_PresentGuids  = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($S_H in $S_Holds)
        {
            if (-not $S_H.Guid) { continue }
            if ($S_H.Excluded) { [void]$S_ExcludedGuids.Add($S_H.Guid) }
            else { [void]$S_PresentGuids.Add($S_H.Guid) }
        }
        foreach ($S_Ow in $S_OrgWideMailboxPolicies)
        {
            if ($S_ExcludedGuids.Contains($S_Ow.Guid) -or $S_PresentGuids.Contains($S_Ow.Guid)) { continue }
            $S_Holds.Add([PSCustomObject]@{
                Type               = $S_Ow.Type
                Scope              = $S_Ow.Scope
                Name               = $S_Ow.Name
                Guid               = $S_Ow.Guid
                Excluded           = $false
                Action             = $S_Ow.Action
                Raw                = $S_Ow.Raw
                OrgWide            = $true
                DistributionStatus = $S_Ow.DistributionStatus
                DistributionDetail = $S_Ow.DistributionDetail
            })
        }

        $S_ActiveHolds = @($S_Holds | Where-Object { -not $_.Excluded })
        $S_OnHold = ($S_ActiveHolds.Count -gt 0) -or
                    $S_Mb.LitigationHoldEnabled -or
                    $S_Mb.ComplianceTagHoldApplied -or
                    $S_Mb.DelayHoldApplied -or
                    $S_Mb.DelayReleaseHoldApplied

        $S_HoldTypes = @($S_ActiveHolds | Select-Object -ExpandProperty Type -Unique)
        if ($S_Mb.ComplianceTagHoldApplied) { $S_HoldTypes += 'RetentionLabelHold' }
        if ($S_Mb.DelayHoldApplied -or $S_Mb.DelayReleaseHoldApplied) { $S_HoldTypes += 'DelayHold' }
        $S_HoldTypes = @($S_HoldTypes | Select-Object -Unique)

        $S_PolicyNames = @($S_ActiveHolds |
            Where-Object { $_.Type -ne 'LitigationHold' } |
            Select-Object -ExpandProperty Name -Unique)

        $S_ReportLine = [PSCustomObject]@{
            Mailbox                 = $S_Mb.DisplayName
            UPN                     = $S_Mb.UserPrincipalName
            PrimarySmtp             = "$($S_Mb.PrimarySmtpAddress)"
            State                   = $S_Mb.MailboxState
            RecipientTypeDetails    = "$($S_Mb.RecipientTypeDetails)"
            OnHold                  = [bool]$S_OnHold
            LitigationHoldEnabled   = [bool]$S_Mb.LitigationHoldEnabled
            LitigationHoldDuration  = "$($S_Mb.LitigationHoldDuration)"
            RetentionHoldEnabled    = [bool]$S_Mb.RetentionHoldEnabled
            MrmRetentionPolicy      = "$($S_Mb.RetentionPolicy)"
            ComplianceTagHoldApplied = [bool]$S_Mb.ComplianceTagHoldApplied
            DelayHoldApplied        = [bool]$S_Mb.DelayHoldApplied
            DelayReleaseHoldApplied = [bool]$S_Mb.DelayReleaseHoldApplied
            HoldCount               = $S_ActiveHolds.Count
            HoldTypes               = $S_HoldTypes
            HoldTypesSummary        = ($S_HoldTypes -join '; ')
            PolicyNames             = ($S_PolicyNames -join '; ')
            Holds                   = $S_Holds
            RawInPlaceHolds         = (@($S_Mb.InPlaceHolds) -join '; ')
            Error                   = $null
        }
    }
    catch
    {
        Write-Warning "  Could not process $($S_Mb.DisplayName): $($_.Exception.Message)"
        $S_ReportLine = [PSCustomObject]@{
            Mailbox                 = $S_Mb.DisplayName
            UPN                     = $S_Mb.UserPrincipalName
            PrimarySmtp             = "$($S_Mb.PrimarySmtpAddress)"
            State                   = $S_Mb.MailboxState
            RecipientTypeDetails    = "$($S_Mb.RecipientTypeDetails)"
            OnHold                  = $false
            LitigationHoldEnabled   = $false
            LitigationHoldDuration  = ''
            RetentionHoldEnabled    = $false
            MrmRetentionPolicy      = ''
            ComplianceTagHoldApplied = $false
            DelayHoldApplied        = $false
            DelayReleaseHoldApplied = $false
            HoldCount               = 0
            HoldTypes               = @()
            HoldTypesSummary        = ''
            PolicyNames             = ''
            Holds                   = @()
            RawInPlaceHolds         = ''
            Error                   = $_.Exception.Message
        }
    }

    $S_Report.Add($S_ReportLine)
}

# ---------------------------------------------------------------------------
# Stats
# ---------------------------------------------------------------------------
$S_Total          = $S_Report.Count
$S_OnHoldCount    = @($S_Report | Where-Object { $_.OnHold }).Count
$S_NoHoldCount    = @($S_Report | Where-Object { -not $_.OnHold -and -not $_.Error }).Count
$S_LitCount       = @($S_Report | Where-Object { $_.LitigationHoldEnabled }).Count
$S_RetCount       = @($S_Report | Where-Object { $_.HoldTypes -contains 'RetentionPolicy' }).Count
$S_EdiscCount     = @($S_Report | Where-Object { $_.HoldTypes -contains 'eDiscoveryCaseHold' -or $_.HoldTypes -contains 'eDiscoveryCloudHold' }).Count
$S_DelayCount     = @($S_Report | Where-Object { $_.DelayHoldApplied -or $_.DelayReleaseHoldApplied }).Count
$S_ErrorCount     = @($S_Report | Where-Object { $_.Error }).Count

# --- Hold-type counts for the bar chart ---
$S_TypeLabels = @('LitigationHold', 'RetentionPolicy', 'eDiscoveryCaseHold', 'AppRetentionPolicy', 'LegacyInPlaceHold', 'RetentionLabelHold', 'DelayHold')
$S_TypeCounts = foreach ($S_T in $S_TypeLabels) { @($S_Report | Where-Object { $_.HoldTypes -contains $S_T }).Count }

# --- Policy-centric pivot (covered / excluded mailboxes per resolved policy) ---
$S_PolicyPivot = @{}
foreach ($S_Line in $S_Report)
{
    foreach ($S_Hold in @($S_Line.Holds))
    {
        if ($S_Hold.Type -eq 'LitigationHold') { continue }
        $S_PolKey = "$($S_Hold.Type)|$($S_Hold.Name)|$($S_Hold.Scope)"
        if (-not $S_PolicyPivot.ContainsKey($S_PolKey))
        {
            $S_PolicyPivot[$S_PolKey] = [PSCustomObject]@{
                Name               = $S_Hold.Name
                Type               = $S_Hold.Type
                Scope              = $S_Hold.Scope
                DistributionStatus = $S_Hold.DistributionStatus
                DistributionDetail = $S_Hold.DistributionDetail
                Covered            = [System.Collections.Generic.List[string]]::new()
                Excluded           = [System.Collections.Generic.List[string]]::new()
            }
        }
        if ($S_Hold.Excluded) { $S_PolicyPivot[$S_PolKey].Excluded.Add($S_Line.Mailbox) }
        else { $S_PolicyPivot[$S_PolKey].Covered.Add($S_Line.Mailbox) }
    }
}
$S_Policies = @($S_PolicyPivot.Values | Sort-Object @{ Expression = { $_.Covered.Count }; Descending = $true }, Name)

# --- All resolved retention / app-retention policies with their distribution status ---
# Answers "is each retention policy fully distributed or not", independent of mailbox coverage.
$S_AllPolicyList = @(@($S_RetentionHash.Values) + @($S_AppHash.Values) | Sort-Object Name)
$S_DistIncomplete = @($S_AllPolicyList | Where-Object { $_.DistributionStatus -and $_.DistributionStatus -ne 'Success' })
$S_DistIncompleteCount = $S_DistIncomplete.Count

# ---------------------------------------------------------------------------
# File paths
# ---------------------------------------------------------------------------
if (-not $ReportPath) { $ReportPath = (Get-Location).Path }
if (Test-Path $ReportPath -PathType Container) { $S_ReportFolder = $ReportPath }
else { $S_ReportFolder = Split-Path -Parent $ReportPath }
if ($S_ReportFolder -and -not (Test-Path $S_ReportFolder))
{
    New-Item -ItemType Directory -Path $S_ReportFolder -Force | Out-Null
}

$S_Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$S_FileBase  = if ($TestMode) { "ReportMailboxHolds_TEST_$S_Timestamp" } else { "ReportMailboxHolds_$S_Timestamp" }

if (Test-Path $ReportPath -PathType Container)
{
    $S_CsvFile  = Join-Path $ReportPath "$S_FileBase.csv"
    $S_HtmlFile = Join-Path $ReportPath "$S_FileBase.html"
}
else
{
    $S_CsvFile  = $ReportPath
    $S_HtmlFile = [System.IO.Path]::ChangeExtension($ReportPath, '.html')
}

# ---------------------------------------------------------------------------
# CSV export
# ---------------------------------------------------------------------------
$S_Report | Sort-Object Mailbox |
    Select-Object Mailbox, UPN, PrimarySmtp, State, RecipientTypeDetails, OnHold,
        LitigationHoldEnabled, LitigationHoldDuration, RetentionHoldEnabled, MrmRetentionPolicy,
        ComplianceTagHoldApplied, DelayHoldApplied, DelayReleaseHoldApplied,
        HoldCount, HoldTypesSummary, PolicyNames, RawInPlaceHolds, Error |
    Export-Csv -Path $S_CsvFile -NoTypeInformation -Encoding UTF8

# ---------------------------------------------------------------------------
# HTML report
# ---------------------------------------------------------------------------
$S_ReportDate = Get-Date -Format 'dd MMM yyyy HH:mm'

function Format-Badge
{
    param([string]$Type, [bool]$Excluded)
    $S_Map = @{
        'LitigationHold'      = @{ cls = 'lit'; text = 'Litigation' }
        'RetentionPolicy'     = @{ cls = 'ret'; text = 'Retention' }
        'eDiscoveryCaseHold'  = @{ cls = 'edisc'; text = 'eDiscovery' }
        'eDiscoveryCloudHold' = @{ cls = 'edisc'; text = 'eDiscovery (cloud)' }
        'AppRetentionPolicy'  = @{ cls = 'app'; text = 'App retention' }
        'LegacyInPlaceHold'   = @{ cls = 'legacy'; text = 'In-Place (legacy)' }
        'RetentionLabelHold'  = @{ cls = 'label'; text = 'Retention label' }
        'DelayHold'           = @{ cls = 'delay'; text = 'Delay hold' }
    }
    $S_Def = if ($S_Map.ContainsKey($Type)) { $S_Map[$Type] } else { @{ cls = 'legacy'; text = $Type } }
    $S_Txt = $S_Def.text
    if ($Excluded) { $S_Txt = "Excluded: $S_Txt" }
    "<span class=`"badge badge-$($S_Def.cls)$(if ($Excluded) { ' badge-excluded' })`">$([System.Net.WebUtility]::HtmlEncode($S_Txt))</span>"
}

# Render a distribution/deployment status as a colour-coded badge (Success = complete).
function Format-Distribution
{
    param([string]$Status, [string]$Detail)

    if ([string]::IsNullOrWhiteSpace($Status)) { return '<span class="muted">-</span>' }
    if ($Status -eq 'Success') { return '<span class="badge badge-none">Complete</span>' }

    $S_Cls = if ($Status -match 'Pending|Processing|Sync') { 'label' } else { 'onhold' }
    $S_Text = "Incomplete: $Status"
    $S_Title = if ($Detail) { " title=`"$([System.Net.WebUtility]::HtmlEncode($Detail))`"" } else { '' }
    "<span class=`"badge badge-$S_Cls`"$S_Title>$([System.Net.WebUtility]::HtmlEncode($S_Text))</span>"
}

$S_TestModeBanner = if ($TestMode)
{
    "<div class=`"notice`"><strong>Test Mode:</strong> This report was generated from a random sample of $($S_Mbx.Count) mailboxes. Run without <code>-TestMode</code> to report on all mailboxes in scope.</div>"
}
else { '' }

$S_PolicyBanner = if (-not $S_PolicyResolutionEnabled)
{
    '<div class="notice notice-warn"><strong>Policy names not resolved:</strong> the Security &amp; Compliance connection was skipped or unavailable, so hold GUIDs are shown unresolved.</div>'
}
else { '' }

# Org-wide policy note
$S_OrgWideRows = ''
if ($S_OrgWidePolicies.Count -gt 0)
{
    $S_OrgWideRows = (@($S_OrgWidePolicies) | ForEach-Object {
        $S_N = [System.Net.WebUtility]::HtmlEncode($_.Name)
        $S_Badge = Format-Badge -Type $_.Type -Excluded $_.Excluded
        $S_ScopeSpan = ''
        if ($_.Scope)
        {
            $S_ScopeEnc = [System.Net.WebUtility]::HtmlEncode($_.Scope)
            $S_ScopeSpan = "<span class=`"muted`">($S_ScopeEnc)</span>"
        }
        "<li>$S_Badge $S_N $S_ScopeSpan</li>"
    }) -join "`n"
    $S_OrgWideRows = @"
<div class="table-section">
  <h2>Organization-wide retention policies</h2>
  <p class="muted" style="margin-bottom:12px;">These entire-location policies are <strong>not stamped</strong> on individual mailbox objects (they only appear in <code>Get-OrganizationConfig</code>), yet they apply to <strong>every</strong> in-scope mailbox except those explicitly excluded. Their coverage is expanded into the per-mailbox rows and the policy-coverage view below. Group-scoped (<code>grp</code>) policies are listed here for reference but target group mailboxes, not user mailboxes.</p>
  <ul class="orglist">$S_OrgWideRows</ul>
</div>
"@
}

# --- Mailbox-centric table rows ---
$S_MailboxRows = ($S_Report | Sort-Object Mailbox | ForEach-Object {
    $S_Name  = [System.Net.WebUtility]::HtmlEncode($_.Mailbox)
    $S_Upn   = [System.Net.WebUtility]::HtmlEncode($_.UPN)
    $S_St    = [System.Net.WebUtility]::HtmlEncode($_.State)

    $S_Badges = if ($_.Error)
    {
        '<span class="badge badge-error">Error</span>'
    }
    elseif ($_.HoldTypes.Count -eq 0)
    {
        '<span class="badge badge-none">No hold</span>'
    }
    else
    {
        (@($_.HoldTypes) | ForEach-Object { Format-Badge -Type $_ -Excluded $false }) -join ' '
    }

    $S_PolCell = if ($_.PolicyNames) { [System.Net.WebUtility]::HtmlEncode($_.PolicyNames) } else { '<span class="muted">-</span>' }
    $S_Mrm     = if ($_.MrmRetentionPolicy) { [System.Net.WebUtility]::HtmlEncode($_.MrmRetentionPolicy) } else { '<span class="muted">-</span>' }
    $S_OnHold  = if ($_.OnHold) { '<span class="badge badge-onhold">On hold</span>' } else { '<span class="badge badge-none">No</span>' }

    # Litigation Hold duration: shown for every mailbox regardless of LitigationHoldEnabled,
    # since LitigationHoldEnabled stays False for org-wide retention coverage. The value
    # renders as an EnhancedTimeSpan ("2555.00:00:00") or "Unlimited".
    $S_LdRaw = "$($_.LitigationHoldDuration)".Trim()
    $S_LitDur = if ($S_LdRaw -match '^(\d+)\.\d{2}:\d{2}:\d{2}') { "$($Matches[1]) days" }
        elseif ([string]::IsNullOrWhiteSpace($S_LdRaw)) { '<span class="muted">-</span>' }
        else { [System.Net.WebUtility]::HtmlEncode($S_LdRaw) }

    $S_DataTypes    = ([System.Net.WebUtility]::HtmlEncode((@($_.HoldTypes) -join ' ')))
    $S_DataPolicies = ([System.Net.WebUtility]::HtmlEncode($_.PolicyNames))

    "<tr data-holdtypes=`"$S_DataTypes`" data-state=`"$S_St`" data-policies=`"$S_DataPolicies`" data-onhold=`"$($_.OnHold)`"><td>$S_Name</td><td class=`"upn`">$S_Upn</td><td>$S_St</td><td>$S_OnHold</td><td>$S_Badges</td><td>$S_PolCell</td><td>$S_Mrm</td><td>$S_LitDur</td></tr>"
}) -join "`n"

# --- Policy-centric table rows ---
$S_PolicyRows = ($S_Policies | ForEach-Object {
    $S_N     = [System.Net.WebUtility]::HtmlEncode($_.Name)
    $S_Badge = Format-Badge -Type $_.Type -Excluded $false
    $S_Sc    = if ($_.Scope) { [System.Net.WebUtility]::HtmlEncode($_.Scope) } else { '<span class="muted">-</span>' }
    $S_CovList = (@($_.Covered | Sort-Object) | ForEach-Object { [System.Net.WebUtility]::HtmlEncode($_) }) -join ', '
    if (-not $S_CovList) { $S_CovList = '<span class="muted">-</span>' }
    $S_ExcCount = $_.Excluded.Count
    $S_Dist = Format-Distribution -Status $_.DistributionStatus -Detail $_.DistributionDetail

    "<tr data-policytype=`"$([System.Net.WebUtility]::HtmlEncode($_.Type))`"><td>$S_N</td><td>$S_Badge</td><td>$S_Sc</td><td>$S_Dist</td><td>$($_.Covered.Count)</td><td>$S_ExcCount</td><td class=`"covered`">$S_CovList</td></tr>"
}) -join "`n"

# --- Policy distribution table rows (all resolved retention / app-retention policies) ---
$S_DistRows = ($S_AllPolicyList | ForEach-Object {
    $S_N     = [System.Net.WebUtility]::HtmlEncode($_.Name)
    $S_Badge = Format-Badge -Type $_.Type -Excluded $false
    $S_Dist  = Format-Distribution -Status $_.DistributionStatus -Detail $_.DistributionDetail
    $S_DetailCell = if ($_.DistributionDetail) { [System.Net.WebUtility]::HtmlEncode($_.DistributionDetail) } else { '<span class="muted">-</span>' }
    "<tr><td>$S_N</td><td>$S_Badge</td><td>$S_Dist</td><td>$S_DetailCell</td></tr>"
}) -join "`n"

if (-not $S_MailboxRows) { $S_MailboxRows = '<tr><td colspan="8" class="muted">No mailboxes found.</td></tr>' }
if (-not $S_PolicyRows)  { $S_PolicyRows  = '<tr><td colspan="7" class="muted">No policy-based holds matched any mailbox in scope.</td></tr>' }
if (-not $S_DistRows)    { $S_DistRows    = '<tr><td colspan="4" class="muted">No retention policies resolved (Security &amp; Compliance not connected).</td></tr>' }

$S_ErrorCard = if ($S_ErrorCount -gt 0) { "<div class=`"card`"><div class=`"label`">Errors</div><div class=`"value`" style=`"color:#6c757d;`">$S_ErrorCount</div><div class=`"sub`">Could not process</div></div>" } else { '' }

$S_Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Mailbox Holds &amp; Retention Report</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; padding: 30px; }

  .header { background: linear-gradient(135deg, #1a1a2e, #16213e); color: #fff; padding: 30px 40px; border-radius: 12px; margin-bottom: 30px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 16px; }
  .header-left h1 { font-size: 1.6em; margin-bottom: 6px; }
  .header-left p { font-size: 0.9em; opacity: 0.8; }

  .section-title { font-size: 1.15em; font-weight: 600; color: #1a1a2e; margin-bottom: 16px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0; }

  .notice { background:#fff3cd;border:1px solid #ffc107;border-radius:8px;padding:12px 20px;margin-bottom:20px;color:#856404;font-size:0.9em; }
  .notice-warn { background:#f8d7da;border-color:#f5c2c7;color:#721c24; }

  .summary-cards { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .card { background: #fff; border-radius: 10px; padding: 24px 30px; flex: 1; min-width: 140px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
  .card .label { font-size: 0.85em; color: #777; text-transform: uppercase; letter-spacing: 0.5px; }
  .card .value { font-size: 2em; font-weight: 700; margin-top: 6px; }
  .card .sub { font-size: 0.78em; color: #999; margin-top: 2px; }

  .charts-row { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .chart-section { background: #fff; border-radius: 10px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); flex: 1; min-width: 300px; }
  .chart-section h2 { font-size: 1.1em; margin-bottom: 20px; color: #1a1a2e; }
  .chart-container { max-width: 380px; margin: 0 auto; }

  .table-section { background: #fff; border-radius: 10px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 30px; overflow-x: auto; }
  .table-section h2 { font-size: 1.1em; margin-bottom: 16px; color: #1a1a2e; }
  .table-controls { display: flex; gap: 14px; margin-bottom: 16px; flex-wrap: wrap; align-items: center; }
  .table-controls input[type="text"] { padding: 8px 14px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; min-width: 240px; }
  .table-controls select { padding: 8px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; background: #fff; }
  .table-controls .count-label { font-size: 0.85em; color: #777; margin-left: auto; }

  table { width: 100%; border-collapse: collapse; font-size: 0.84em; }
  th { background: #1a1a2e; color: #fff; padding: 12px 14px; text-align: left; position: sticky; top: 0; cursor: pointer; user-select: none; white-space: nowrap; }
  th:hover { background: #2c3e50; }
  td { padding: 10px 14px; border-bottom: 1px solid #eee; vertical-align: top; }
  tr:hover td { background: #f8f9fa; }
  tr.hidden-row { display: none; }
  .upn { font-family: 'Consolas', monospace; font-size: 0.82em; color: #666; }
  .covered { font-size: 0.9em; color: #444; max-width: 520px; white-space: normal; }
  .muted { color: #aaa; }
  .orglist { list-style: none; display: flex; flex-direction: column; gap: 8px; }

  .badge { padding: 3px 10px; border-radius: 12px; font-size: 0.8em; font-weight: 600; display: inline-block; margin: 1px 0; }
  .badge-lit    { background: #e7d9ff; color: #5b21b6; }
  .badge-ret    { background: #d4edff; color: #0b5394; }
  .badge-edisc  { background: #ffe0cc; color: #9a3412; }
  .badge-app    { background: #d1f5e0; color: #12694a; }
  .badge-legacy { background: #e2e3e5; color: #495057; }
  .badge-label  { background: #fff3cd; color: #856404; }
  .badge-delay  { background: #fde2e1; color: #9a1f1a; }
  .badge-excluded { opacity: 0.55; text-decoration: line-through; }
  .badge-onhold { background: #f8d7da; color: #721c24; }
  .badge-none   { background: #d4edda; color: #155724; }
  .badge-error  { background: #e2e3e5; color: #495057; }

  .footer { text-align: center; font-size: 0.8em; color: #999; margin-top: 20px; }
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <h1>Mailbox Holds &amp; Retention Report</h1>
    <p>Tenant: $([System.Net.WebUtility]::HtmlEncode($S_TenantDisplayName)) ($S_TenantId) &nbsp;|&nbsp; Generated: $S_ReportDate &nbsp;|&nbsp; Scope: $MailboxScope</p>
  </div>
</div>

$S_TestModeBanner
$S_PolicyBanner

<!-- OVERVIEW -->
<div class="section-title">Overview</div>
<div class="summary-cards">
  <div class="card"><div class="label">Total Mailboxes</div><div class="value" style="color:#1a1a2e;">$S_Total</div></div>
  <div class="card"><div class="label">On Hold</div><div class="value" style="color:#e74c3c;">$S_OnHoldCount</div></div>
  <div class="card"><div class="label">No Hold</div><div class="value" style="color:#27ae60;">$S_NoHoldCount</div></div>
  <div class="card"><div class="label">Litigation Hold</div><div class="value" style="color:#7c3aed;">$S_LitCount</div></div>
  <div class="card"><div class="label">Retention Policy</div><div class="value" style="color:#0b5394;">$S_RetCount</div></div>
  <div class="card"><div class="label">eDiscovery</div><div class="value" style="color:#9a3412;">$S_EdiscCount</div></div>
  <div class="card"><div class="label">Delay Hold</div><div class="value" style="color:#9a1f1a;">$S_DelayCount</div></div>
  <div class="card"><div class="label">Distribution Incomplete</div><div class="value" style="color:$(if ($S_DistIncompleteCount -gt 0) { '#e67e22' } else { '#27ae60' });">$S_DistIncompleteCount</div><div class="sub">of $($S_AllPolicyList.Count) retention policies</div></div>
  $S_ErrorCard
</div>

<!-- CHARTS -->
<div class="charts-row">
  <div class="chart-section">
    <h2>Hold Status</h2>
    <div class="chart-container"><canvas id="statusChart"></canvas></div>
  </div>
  <div class="chart-section">
    <h2>Mailboxes by Hold Type</h2>
    <div class="chart-container"><canvas id="typeChart"></canvas></div>
  </div>
</div>

$S_OrgWideRows

<!-- VIEW 1: MAILBOX-CENTRIC -->
<div class="table-section">
  <h2>View 1 &mdash; Mailboxes and the holds covering them</h2>
  <div class="table-controls">
    <input type="text" id="mbxSearch" placeholder="Search name / UPN / policy..." onkeyup="filterMbx()" />
    <select id="typeFilter" onchange="filterMbx()">
      <option value="all">All hold types</option>
      <option value="LitigationHold">Litigation Hold</option>
      <option value="RetentionPolicy">Retention Policy</option>
      <option value="eDiscoveryCaseHold">eDiscovery Case Hold</option>
      <option value="eDiscoveryCloudHold">eDiscovery Cloud Hold</option>
      <option value="AppRetentionPolicy">App Retention Policy</option>
      <option value="LegacyInPlaceHold">Legacy In-Place Hold</option>
      <option value="RetentionLabelHold">Retention Label Hold</option>
      <option value="DelayHold">Delay Hold</option>
    </select>
    <select id="stateFilter" onchange="filterMbx()">
      <option value="all">All states</option>
      <option value="Active">Active</option>
      <option value="Inactive">Inactive</option>
      <option value="SoftDeleted">SoftDeleted</option>
    </select>
    <select id="onholdFilter" onchange="filterMbx()">
      <option value="all">On hold: any</option>
      <option value="True">On hold only</option>
      <option value="False">Not on hold</option>
    </select>
    <span class="count-label" id="mbxCount">Showing $S_Total of $S_Total mailboxes</span>
  </div>
  <table id="mbxTable">
    <thead>
      <tr>
        <th onclick="sortTable('mbxTable', 0)">Mailbox &#x25B2;&#x25BC;</th>
        <th onclick="sortTable('mbxTable', 1)">UPN &#x25B2;&#x25BC;</th>
        <th onclick="sortTable('mbxTable', 2)">State &#x25B2;&#x25BC;</th>
        <th onclick="sortTable('mbxTable', 3)">On Hold &#x25B2;&#x25BC;</th>
        <th>Hold Types</th>
        <th onclick="sortTable('mbxTable', 5)">Retention / Hold Policies &#x25B2;&#x25BC;</th>
        <th onclick="sortTable('mbxTable', 6)">MRM Policy &#x25B2;&#x25BC;</th>
        <th onclick="sortTable('mbxTable', 7)">Litigation Duration &#x25B2;&#x25BC;</th>
      </tr>
    </thead>
    <tbody id="mbxBody">
$S_MailboxRows
    </tbody>
  </table>
</div>

<!-- VIEW 2: POLICY-CENTRIC -->
<div class="table-section">
  <h2>View 2 &mdash; Policies and the mailboxes they cover</h2>
  <div class="table-controls">
    <input type="text" id="polSearch" placeholder="Search policy name / mailbox..." onkeyup="filterPol()" />
    <span class="count-label" id="polCount">$($S_Policies.Count) policies</span>
  </div>
  <table id="polTable">
    <thead>
      <tr>
        <th onclick="sortTable('polTable', 0)">Policy &#x25B2;&#x25BC;</th>
        <th>Type</th>
        <th onclick="sortTable('polTable', 2)">Scope &#x25B2;&#x25BC;</th>
        <th onclick="sortTable('polTable', 3)">Distribution &#x25B2;&#x25BC;</th>
        <th onclick="sortTable('polTable', 4)">Covered &#x25B2;&#x25BC;</th>
        <th onclick="sortTable('polTable', 5)">Excluded &#x25B2;&#x25BC;</th>
        <th>Covered mailboxes</th>
      </tr>
    </thead>
    <tbody id="polBody">
$S_PolicyRows
    </tbody>
  </table>
</div>

<!-- POLICY DISTRIBUTION -->
<div class="table-section">
  <h2>Retention policy distribution status</h2>
  <p class="muted" style="margin-bottom:12px;">Every resolved Microsoft Purview retention / app-retention policy and whether its deployment to all locations is <strong>complete</strong> (<code>DistributionStatus = Success</code>) or still pending/errored. Hover an incomplete badge for the per-location detail.</p>
  <div class="table-controls">
    <input type="text" id="distSearch" placeholder="Search policy name..." onkeyup="filterDist()" />
    <span class="count-label" id="distCount">$($S_AllPolicyList.Count) policies &nbsp;|&nbsp; $S_DistIncompleteCount not fully distributed</span>
  </div>
  <table id="distTable">
    <thead>
      <tr>
        <th onclick="sortTable('distTable', 0)">Policy &#x25B2;&#x25BC;</th>
        <th>Type</th>
        <th onclick="sortTable('distTable', 2)">Distribution &#x25B2;&#x25BC;</th>
        <th>Detail</th>
      </tr>
    </thead>
    <tbody id="distBody">
$S_DistRows
    </tbody>
  </table>
</div>

<div class="footer">Generated by ReportMailboxHolds.ps1 &nbsp;|&nbsp; $S_ReportDate</div>

<script>
// --- Chart: hold status ---
(function() {
  var ctx = document.getElementById('statusChart').getContext('2d');
  new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: ['On Hold', 'No Hold'$(if ($S_ErrorCount -gt 0) { ", 'Error'" })],
      datasets: [{
        data: [$S_OnHoldCount, $S_NoHoldCount$(if ($S_ErrorCount -gt 0) { ", $S_ErrorCount" })],
        backgroundColor: ['#e74c3c', '#27ae60'$(if ($S_ErrorCount -gt 0) { ", '#bbb'" })],
        borderWidth: 2, borderColor: '#fff'
      }]
    },
    options: { plugins: { legend: { position: 'bottom' } }, cutout: '60%' }
  });
})();

// --- Chart: hold types ---
(function() {
  var ctx = document.getElementById('typeChart').getContext('2d');
  new Chart(ctx, {
    type: 'bar',
    data: {
      labels: ['Litigation','Retention','eDiscovery','App ret.','Legacy','Ret. label','Delay'],
      datasets: [{
        label: 'Mailboxes',
        data: [$($S_TypeCounts -join ', ')],
        backgroundColor: ['#7c3aed','#0b5394','#9a3412','#12694a','#6c757d','#856404','#9a1f1a'],
        borderRadius: 4
      }]
    },
    options: {
      indexAxis: 'y',
      plugins: { legend: { display: false } },
      scales: { x: { beginAtZero: true, ticks: { precision: 0 } } }
    }
  });
})();

// --- Filter: mailbox table ---
function filterMbx() {
  var search = document.getElementById('mbxSearch').value.toLowerCase();
  var type = document.getElementById('typeFilter').value;
  var state = document.getElementById('stateFilter').value;
  var onhold = document.getElementById('onholdFilter').value;
  var rows = document.querySelectorAll('#mbxBody tr');
  var visible = 0;
  rows.forEach(function(r) {
    var text = r.textContent.toLowerCase();
    var types = r.getAttribute('data-holdtypes') || '';
    var rState = r.getAttribute('data-state') || '';
    var rOnhold = r.getAttribute('data-onhold') || '';
    var matchSearch = !search || text.indexOf(search) > -1;
    var matchType = type === 'all' || types.split(' ').indexOf(type) > -1;
    var matchState = state === 'all' || rState === state;
    var matchOnhold = onhold === 'all' || rOnhold === onhold;
    if (matchSearch && matchType && matchState && matchOnhold) { r.classList.remove('hidden-row'); visible++; }
    else { r.classList.add('hidden-row'); }
  });
  document.getElementById('mbxCount').textContent = 'Showing ' + visible + ' of ' + rows.length + ' mailboxes';
}

// --- Filter: policy table ---
function filterPol() {
  var search = document.getElementById('polSearch').value.toLowerCase();
  var rows = document.querySelectorAll('#polBody tr');
  var visible = 0;
  rows.forEach(function(r) {
    var text = r.textContent.toLowerCase();
    if (!search || text.indexOf(search) > -1) { r.classList.remove('hidden-row'); visible++; }
    else { r.classList.add('hidden-row'); }
  });
  document.getElementById('polCount').textContent = visible + ' policies';
}

// --- Filter: distribution table ---
function filterDist() {
  var search = document.getElementById('distSearch').value.toLowerCase();
  var rows = document.querySelectorAll('#distBody tr');
  var visible = 0;
  rows.forEach(function(r) {
    var text = r.textContent.toLowerCase();
    if (!search || text.indexOf(search) > -1) { r.classList.remove('hidden-row'); visible++; }
    else { r.classList.add('hidden-row'); }
  });
  document.getElementById('distCount').textContent = visible + ' policies';
}

// --- Sort ---
var sortDir = {};
function sortTable(tableId, col) {
  var tbody = document.getElementById(tableId).querySelector('tbody');
  var rows = Array.from(tbody.querySelectorAll('tr'));
  var key = tableId + '-' + col;
  var dir = sortDir[key] === 'asc' ? 'desc' : 'asc';
  sortDir[key] = dir;
  rows.sort(function(a, b) {
    var ta = a.cells[col] ? a.cells[col].textContent.trim() : '';
    var tb = b.cells[col] ? b.cells[col].textContent.trim() : '';
    var na = parseFloat(ta); var nb = parseFloat(tb);
    if (!isNaN(na) && !isNaN(nb)) { return dir === 'asc' ? na - nb : nb - na; }
    return dir === 'asc' ? ta.localeCompare(tb) : tb.localeCompare(ta);
  });
  rows.forEach(function(r) { tbody.appendChild(r); });
}
</script>
</body>
</html>
"@

$S_Html | Out-File -FilePath $S_HtmlFile -Encoding UTF8

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "Report complete." -ForegroundColor Green
Write-Host "  CSV  : $S_CsvFile"
Write-Host "  HTML : $S_HtmlFile"
Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "  Total mailboxes    : $S_Total"
Write-Host "  On hold            : $S_OnHoldCount" -ForegroundColor $(if ($S_OnHoldCount -gt 0) { 'Yellow' } else { 'White' })
Write-Host "  No hold            : $S_NoHoldCount"
Write-Host "  Litigation hold    : $S_LitCount"
Write-Host "  Retention policy   : $S_RetCount"
Write-Host "  eDiscovery hold    : $S_EdiscCount"
Write-Host "  Delay hold         : $S_DelayCount"
if ($S_ErrorCount -gt 0) { Write-Host "  Errors             : $S_ErrorCount" -ForegroundColor Yellow }
Write-Host "  Resolved policies  : $($S_Policies.Count)"
