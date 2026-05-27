#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Reports, Microsoft.Graph.Identity.SignIns

<#
.SYNOPSIS
    Generates a CSV and HTML report of MFA coverage across all Member users
    (Guests are excluded).

.DESCRIPTION
    Connects to Microsoft Graph, retrieves all Member users (userType eq 'Member',
    excluding Guests), then for each user inspects the registered authentication
    methods to classify MFA status (Modern Auth, Legacy Auth, or No MFA).
    Account activity, licensing and on-premises sync state are also captured.
    Results are exported to CSV and a summary HTML dashboard.

.PARAMETER InactiveDays
    Number of days since last sign-in to consider an account inactive.
    If a user has not signed in within this many days, ActiveAccount reports
    how long ago they last signed in.

.PARAMETER OutputPath
    Path for the output CSV file. Defaults to a timestamped file in the
    current working directory.

.PARAMETER Test
    When specified, only processes the first 10 Member users for quick testing.

.EXAMPLE
    .\ReportMemberMFA.ps1 -InactiveDays 90

.EXAMPLE
    .\ReportMemberMFA.ps1 -InactiveDays 30 -Test
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateRange(1, 3650)]
    [int]$InactiveDays,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]$Test
)

# ── Setup ──────────────────────────────────────────────────────────────────────
$ErrorActionPreference = 'Stop'

if (-not $OutputPath) {
    $S_Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputPath = Join-Path (Get-Location).Path "ReportMemberMFA_$S_Timestamp.csv"
}

$S_HtmlPath = [System.IO.Path]::ChangeExtension($OutputPath, '.html')

$S_CutoffDate = (Get-Date).AddDays(-$InactiveDays)

# ── Connect to Microsoft Graph ─────────────────────────────────────────────────
$S_RequiredGraphScopes = @(
    'User.Read.All'
    'AuditLog.Read.All'
    'UserAuthenticationMethod.Read.All'
    'Policy.Read.All'
)

$S_ExistingContext = Get-MgContext
if ($S_ExistingContext) {
    Write-Host "Existing Graph session detected:" -ForegroundColor Yellow
    Write-Host "  Account : $($S_ExistingContext.Account)" -ForegroundColor Yellow
    Write-Host "  TenantId: $($S_ExistingContext.TenantId)" -ForegroundColor Yellow
    Write-Host "  Scopes  : $($S_ExistingContext.Scopes -join ', ')" -ForegroundColor Yellow
    Write-Host ""

    $S_Choice = Read-Host "Use existing session? [Y] Yes  [N] Disconnect and reconnect  (Default: Y)"
    if ($S_Choice -eq 'N') {
        Write-Host "Disconnecting existing session..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Write-Host "Reconnecting with required scopes..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome
        Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
    }
    else {
        Write-Host "Using existing Graph session." -ForegroundColor Green
    }
}
else {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome
    Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
}

try {
    # ── Check Security Defaults Status ─────────────────────────────────────────
    Write-Host "Checking Security Defaults status..." -ForegroundColor Cyan
    $Script:S_SecurityDefaultsEnabled = $null
    try {
        $S_SecDefaultsPolicy = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy'
        $Script:S_SecurityDefaultsEnabled = [bool]$S_SecDefaultsPolicy.isEnabled
        Write-Host "Security Defaults Enabled: $Script:S_SecurityDefaultsEnabled" -ForegroundColor (if ($Script:S_SecurityDefaultsEnabled) { 'Red' } else { 'Green' })
    }
    catch {
        Write-Warning "Could not retrieve Security Defaults policy: $_"
    }

    # ── Conditional Access — MFA-enforcing policies (only when Security Defaults is OFF) ─────────
    $Script:S_MfaCaPolicies = @()
    if ($Script:S_SecurityDefaultsEnabled -eq $false) {
        Write-Host "Retrieving enabled Conditional Access policies..." -ForegroundColor Cyan
        try {
            $S_EnabledCaPolicies = Get-MgIdentityConditionalAccessPolicy -All `
                -Filter "State eq 'enabled'" `
                -Property "id,displayName,state,grantControls,conditions"

            # Keep only policies whose grant controls require MFA (built-in) OR an Authentication Strength.
            # NOTE: The Graph SDK returns a non-null but empty AuthenticationStrength placeholder on
            # policies that don't actually configure one (e.g. Sign-In Frequency, Block policies).
            # We must validate that the AuthenticationStrength.Id is populated.
            $S_FilteredCaPolicies = @(
                $S_EnabledCaPolicies | Where-Object {
                    $S_Grant = $_.GrantControls
                    if ($null -eq $S_Grant) { return $false }
                    $S_HasMfaBuiltIn = ($S_Grant.BuiltInControls -contains 'mfa')
                    $S_HasAuthStrength = (
                        $null -ne $S_Grant.AuthenticationStrength -and
                        -not [string]::IsNullOrWhiteSpace([string]$S_Grant.AuthenticationStrength.Id)
                    )
                    $S_HasMfaBuiltIn -or $S_HasAuthStrength
                }
            )

            # ── Resolver caches (avoid repeat Graph calls) ─────────────────────────
            $S_UserCache  = @{}
            $S_GroupCache = @{}
            $S_RoleCache  = @{}

            function Resolve-CaUserId {
                param([string]$Id)
                if ([string]::IsNullOrWhiteSpace($Id)) { return $null }
                if ($Id -in @('All', 'None', 'GuestsOrExternalUsers')) { return $Id }
                if ($S_UserCache.ContainsKey($Id)) { return $S_UserCache[$Id] }
                try {
                    $u = Get-MgUser -UserId $Id -Property Id,DisplayName,UserPrincipalName -ErrorAction Stop
                    $label = "$($u.DisplayName) <$($u.UserPrincipalName)>"
                }
                catch { $label = "<unresolved:$Id>" }
                $S_UserCache[$Id] = $label
                return $label
            }

            function Resolve-CaGroupId {
                param([string]$Id)
                if ([string]::IsNullOrWhiteSpace($Id)) { return $null }
                if ($S_GroupCache.ContainsKey($Id)) { return $S_GroupCache[$Id] }
                try {
                    $g = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$Id`?`$select=id,displayName" -ErrorAction Stop
                    $label = "$($g.displayName) [group]"
                }
                catch { $label = "<unresolved:$Id>" }
                $S_GroupCache[$Id] = $label
                return $label
            }

            function Resolve-CaRoleId {
                param([string]$Id)
                if ([string]::IsNullOrWhiteSpace($Id)) { return $null }
                if ($S_RoleCache.ContainsKey($Id)) { return $S_RoleCache[$Id] }
                try {
                    $r = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/directoryRoleTemplates/$Id" -ErrorAction Stop
                    $label = "$($r.displayName) [role]"
                }
                catch { $label = "<unresolved:$Id>" }
                $S_RoleCache[$Id] = $label
                return $label
            }

            $S_LocationCache = @{}
            function Resolve-CaLocationId {
                param([string]$Id)
                if ([string]::IsNullOrWhiteSpace($Id)) { return $null }
                if ($Id -in @('All', 'AllTrusted', 'MultiFactorAuthentication')) { return $Id }
                if ($S_LocationCache.ContainsKey($Id)) { return $S_LocationCache[$Id] }
                try {
                    $l = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations/$Id" -ErrorAction Stop
                    $label = "$($l.displayName) [location]"
                }
                catch { $label = "<unresolved:$Id>" }
                $S_LocationCache[$Id] = $label
                return $label
            }

            # ── Project enriched policy objects ────────────────────────────────────
            # Enrich EVERY MFA-enforcing policy first (regardless of audience).
            # The audience filter (Members vs Guests) is applied AFTER enrichment
            # so that a future guest-focused script can reuse the same projection
            # by simply flipping the filter predicate from TargetsMembers to
            # TargetsGuests — no rewrite of the resolver/enrichment code required.
            $S_EnrichedCaPolicies = @(
                foreach ($p in $S_FilteredCaPolicies) {
                    $apps  = $p.Conditions.Applications
                    $usrs  = $p.Conditions.Users                    
                    $plat  = $p.Conditions.Platforms
                    $locs  = $p.Conditions.Locations
                    $S_GrantTypes = @()
                    if ($p.GrantControls.BuiltInControls -contains 'mfa') { $S_GrantTypes += 'MFA' }
                    $S_HasAuthStr = (
                        $null -ne $p.GrantControls.AuthenticationStrength -and
                        -not [string]::IsNullOrWhiteSpace([string]$p.GrantControls.AuthenticationStrength.Id)
                    )
                    if ($S_HasAuthStr) { $S_GrantTypes += 'AuthStrength' }

                    # ── Grant operator & companion controls ──────────────────────
                    # Operator is 'AND' or 'OR'. With OR + other controls present,
                    # a user can satisfy the policy WITHOUT MFA (e.g. compliant
                    # device alone) — a potential MFA gap we must flag.
                    $S_GrantOperator = $p.GrantControls.Operator   # 'AND' | 'OR' | $null
                    $S_AllBuiltIn    = @($p.GrantControls.BuiltInControls)
                    $S_OtherBuiltIn  = @($S_AllBuiltIn | Where-Object { $_ -ne 'mfa' })

                    # Companion controls = everything in the grant besides MFA / AuthStrength
                    $S_CompanionControls = @($S_OtherBuiltIn)

                    # Bypassable if: operator is OR and there is at least one
                    # other grant control that could satisfy the policy alone.
                    $S_MfaIsBypassable = (
                        $S_GrantOperator -eq 'OR' -and
                        $S_CompanionControls.Count -gt 0
                    )

                    # ── Workload (application) coverage tier ─────────────────────
                    # Classify how broadly the policy covers cloud apps. Used by
                    # the gap engine to weight whether a user is meaningfully
                    # protected for everyday sign-ins.
                    #   • Full    → IncludeApplications contains 'All'
                    #   • Data    → IncludeApplications contains 'Office365'
                    #               (Exchange / SharePoint / Teams / etc. — the
                    #               primary data plane in most M365 tenants)
                    #   • Partial → specific app GUIDs only, user-action-only,
                    #               or empty include scope
                    $S_IncAppsRaw = @($apps.IncludeApplications)
                    $S_WorkloadCoverage = if ($S_IncAppsRaw -contains 'All') {
                        'Full'
                    } elseif ($S_IncAppsRaw -contains 'Office365') {
                        'Data'
                    } else {
                        'Partial'
                    }

                    # ── Other-conditions posture ───────────────────────────────
                    # Summarise the "Other Conditions" tab in one value. Used
                    # by the gap engine together with WorkloadCoverage.
                    #   • RiskBased    → policy is gated on sign-in or user risk
                    #                    — only fires on risky sessions, so it is
                    #                    intentionally NOT an everyday MFA control.
                    #                    Takes precedence (the headline trait).
                    #   • Constrained  → narrowing conditions present that prevent
                    #                    the policy from firing on every sign-in:
                    #                    excluded apps, platform scope, location
                    #                    scope, or ClientAppTypes is anything other
                    #                    than 'all' / empty.
                    #   • Unrestricted → nothing in the Conditions tab narrows when
                    #                    the policy fires — the cleanest shape.
                    # NOTE: user/group/role excludes live in Assignments, not
                    # Conditions, so they are deliberately NOT considered here.
                    $S_HasRisk = (
                        @($p.Conditions.SignInRiskLevels | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0 -or
                        @($p.Conditions.UserRiskLevels   | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0
                    )
                    $S_ClientAppsRaw = @($p.Conditions.ClientAppTypes | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                    $S_ClientAppsIsAll = (
                        $S_ClientAppsRaw.Count -eq 0 -or
                        ($S_ClientAppsRaw.Count -eq 1 -and $S_ClientAppsRaw[0] -eq 'all')
                    )
                    $S_HasExcludedApps  = @($apps.ExcludeApplications | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0
                    # An Include* array is "Any" (not narrowing) when it is empty
                    # or contains the single sentinel 'all'/'All'. Graph returns
                    # IncludePlatforms = ['all'] for "Any device platform" and
                    # IncludeLocations = ['All'] for "Any location" — both of
                    # which are the default scopes and must NOT be flagged.
                    # We also strip nulls/blanks because `@($plat.IncludePlatforms)`
                    # becomes `@($null)` (Count=1) when the Platforms condition
                    # block itself is absent.
                    $S_IncPlat = @($plat.IncludePlatforms | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                    $S_IncLoc  = @($locs.IncludeLocations | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                    $S_ExcPlat = @($plat.ExcludePlatforms | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                    $S_ExcLoc  = @($locs.ExcludeLocations | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                    $S_PlatIncludeIsAny = (
                        $S_IncPlat.Count -eq 0 -or
                        ($S_IncPlat.Count -eq 1 -and $S_IncPlat[0] -in @('all','All'))
                    )
                    $S_LocIncludeIsAny = (
                        $S_IncLoc.Count -eq 0 -or
                        ($S_IncLoc.Count -eq 1 -and $S_IncLoc[0] -in @('all','All'))
                    )
                    $S_HasPlatformScope = (
                        -not $S_PlatIncludeIsAny -or
                        $S_ExcPlat.Count -gt 0
                    )
                    $S_HasLocationScope = (
                        -not $S_LocIncludeIsAny -or
                        $S_ExcLoc.Count -gt 0
                    )
                    $S_HasNarrowing = (
                        -not $S_ClientAppsIsAll -or
                        $S_HasExcludedApps -or
                        $S_HasPlatformScope -or
                        $S_HasLocationScope
                    )
                    $S_ConditionsPosture = if ($S_HasRisk) {
                        'RiskBased'
                    } elseif ($S_HasNarrowing) {
                        'Constrained'
                    } else {
                        'Unrestricted'
                    }

                    # ── Enforcement tier (gap-engine eligibility) ─────────────
                    # Pre-computed verdict the per-user MFA gap engine will use
                    # to decide which CA policies to count as meaningful MFA
                    # enforcement for an everyday sign-in. Anything classified
                    # 'Ignored' is still surfaced in the HTML table but will
                    # NOT contribute to per-user MFA coverage in the next phase.
                    #   • Ideal      → Full coverage + Unrestricted conditions
                    #   • Acceptable → (Full|Data) coverage + (Unrestricted|Constrained) conditions
                    #                  (excluding the Ideal combo)
                    #   • Ignored    → Partial coverage, or RiskBased posture
                    $S_EnforcementTier = if ($S_WorkloadCoverage -eq 'Full' -and $S_ConditionsPosture -eq 'Unrestricted') {
                        'Ideal'
                    } elseif ($S_WorkloadCoverage -in @('Full','Data') -and $S_ConditionsPosture -in @('Unrestricted','Constrained')) {
                        'Acceptable'
                    } else {
                        'Ignored'
                    }

                    # ── Audience classification ──────────────────────────────────
                    # Determine whether this policy can apply to Member users,
                    # Guest users, or both. Rules:
                    #   • IncludeUsers = 'All'        → both audiences
                    #   • Specific user GUIDs         → assume both (a GUID may be
                    #     a member or a guest; we don't pre-resolve userType here)
                    #   • IncludeGroups               → assume both (group can mix
                    #     member + guest members)
                    #   • IncludeRoles                → assume both (directory
                    #     roles CAN be assigned to B2B guests, so a role-targeted
                    #     policy applies to whichever members/guests hold the role)
                    #   • IncludeGuestsOrExternalUsers (new model w/ GuestOrExternalUserTypes)
                    #                                 → Guests only
                    #   • Legacy 'GuestsOrExternalUsers' token in IncludeUsers
                    #                                 → Guests only
                    $S_IncUsersRaw   = @($usrs.IncludeUsers)
                    $S_HasUsersAll   = ($S_IncUsersRaw -contains 'All')
                    $S_HasUsersGuestToken = ($S_IncUsersRaw -contains 'GuestsOrExternalUsers')
                    $S_HasSpecificUsers   = @(
                        $S_IncUsersRaw | Where-Object {
                            $_ -and $_ -notin @('All', 'None', 'GuestsOrExternalUsers')
                        }
                    ).Count -gt 0
                    $S_HasIncGroups  = @($usrs.IncludeGroups).Count -gt 0
                    $S_HasIncRoles   = @($usrs.IncludeRoles).Count  -gt 0
                    $S_IncGuestObj   = $usrs.IncludeGuestsOrExternalUsers
                    $S_IncGuestTypes = @()
                    if ($null -ne $S_IncGuestObj -and -not [string]::IsNullOrWhiteSpace([string]$S_IncGuestObj.GuestOrExternalUserTypes)) {
                        $S_IncGuestTypes = @(([string]$S_IncGuestObj.GuestOrExternalUserTypes -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                    }
                    $S_HasIncGuestSpec = $S_IncGuestTypes.Count -gt 0

                    $S_TargetsMembers = (
                        $S_HasUsersAll -or
                        $S_HasSpecificUsers -or
                        $S_HasIncGroups -or
                        $S_HasIncRoles
                    )
                    $S_TargetsGuests = (
                        $S_HasUsersAll -or
                        $S_HasSpecificUsers -or
                        $S_HasIncGroups -or
                        $S_HasIncRoles -or
                        $S_HasIncGuestSpec -or
                        $S_HasUsersGuestToken
                    )

                    [PSCustomObject]@{
                        Id                         = $p.Id
                        DisplayName                = $p.DisplayName
                        State                      = $p.State
                        GrantType                  = ($S_GrantTypes -join ' + ')
                        AuthenticationStrengthId   = $p.GrantControls.AuthenticationStrength.Id
                        AuthenticationStrengthName = $p.GrantControls.AuthenticationStrength.DisplayName
                        GrantOperator              = $S_GrantOperator
                        AllBuiltInControls         = $S_AllBuiltIn
                        CompanionControls          = $S_CompanionControls
                        HasAuthenticationStrength  = $S_HasAuthStr
                        MfaIsBypassable            = $S_MfaIsBypassable
                        WorkloadCoverage           = $S_WorkloadCoverage
                        ConditionsPosture          = $S_ConditionsPosture
                        EnforcementTier            = $S_EnforcementTier
                        TargetsMembers             = $S_TargetsMembers
                        TargetsGuests              = $S_TargetsGuests
                        IncludeApplications        = @($apps.IncludeApplications)
                        ExcludeApplications        = @($apps.ExcludeApplications)
                        IncludeUserActions         = @($apps.IncludeUserActions)
                        IncludeUsers               = @($usrs.IncludeUsers  | ForEach-Object { Resolve-CaUserId  $_ })
                        ExcludeUsers               = @($usrs.ExcludeUsers  | ForEach-Object { Resolve-CaUserId  $_ })
                        IncludeGroups              = @($usrs.IncludeGroups | ForEach-Object { Resolve-CaGroupId $_ })
                        ExcludeGroups              = @($usrs.ExcludeGroups | ForEach-Object { Resolve-CaGroupId $_ })
                        IncludeRoles               = @($usrs.IncludeRoles  | ForEach-Object { Resolve-CaRoleId  $_ })
                        ExcludeRoles               = @($usrs.ExcludeRoles  | ForEach-Object { Resolve-CaRoleId  $_ })
                        IncludeGuestsOrExternalUserTypes = $S_IncGuestTypes
                        IncludePlatforms           = @($plat.IncludePlatforms | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                        ExcludePlatforms           = @($plat.ExcludePlatforms | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                        IncludeLocations           = @($locs.IncludeLocations | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { Resolve-CaLocationId $_ })
                        ExcludeLocations           = @($locs.ExcludeLocations | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { Resolve-CaLocationId $_ })
                        ClientAppTypes             = @($p.Conditions.ClientAppTypes | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                        SignInRiskLevels           = @($p.Conditions.SignInRiskLevels | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                        UserRiskLevels             = @($p.Conditions.UserRiskLevels | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                    }
                }
            )

            # ── Audience filter ────────────────────────────────────────────────────
            # This script reports on MEMBER MFA coverage, so we keep only the
            # policies that can apply to Member users. A future guest-targeted
            # script should re-use the enrichment above and swap this predicate
            # to `$_.TargetsGuests`.
            $S_GuestOnlyPolicies = @($S_EnrichedCaPolicies | Where-Object { -not $_.TargetsMembers })
            $Script:S_MfaCaPolicies = @($S_EnrichedCaPolicies | Where-Object { $_.TargetsMembers })

            Write-Host ("Found {0} enabled CA policy(ies) requiring MFA or Authentication Strength; {1} apply to Members, {2} are guest-only (excluded)." -f `
                $S_EnrichedCaPolicies.Count, $Script:S_MfaCaPolicies.Count, $S_GuestOnlyPolicies.Count) -ForegroundColor Green

            # ── Enforcement-tier summary (member-scoped) ────────────────────
            $S_TierCounts = $Script:S_MfaCaPolicies | Group-Object EnforcementTier -AsHashTable -AsString
            $S_IdealN      = if ($S_TierCounts -and $S_TierCounts['Ideal'])      { @($S_TierCounts['Ideal']).Count }      else { 0 }
            $S_AcceptableN = if ($S_TierCounts -and $S_TierCounts['Acceptable']) { @($S_TierCounts['Acceptable']).Count } else { 0 }
            $S_IgnoredN    = if ($S_TierCounts -and $S_TierCounts['Ignored'])    { @($S_TierCounts['Ignored']).Count }    else { 0 }
            Write-Host ("  Enforcement tiers — Ideal: {0}, Acceptable: {1}, Ignored: {2} (gap engine will skip Ignored)." -f `
                $S_IdealN, $S_AcceptableN, $S_IgnoredN) -ForegroundColor Cyan
        }
        catch {
            Write-Warning "Could not retrieve Conditional Access policies: $_"
        }
    }
    else {
        Write-Host "Skipping Conditional Access policy query (Security Defaults is enabled or unknown)." -ForegroundColor DarkGray
    }

    # ── Authentication Strength policies referenced by kept CA policies ────────────────────────
    # We deliberately only resolve the AuthenticationStrengths that are actually
    # wired into one of the MFA-enforcing CA policies — anything else in the
    # tenant is noise for gap-analysis purposes.
    $Script:S_ReferencedAuthStrengths = @{}   # Keyed by Id for O(1) lookup later
    $S_ReferencedAuthStrengthIds = @(
        $Script:S_MfaCaPolicies |
            Where-Object { $_.HasAuthenticationStrength -and $_.AuthenticationStrengthId } |
            Select-Object -ExpandProperty AuthenticationStrengthId -Unique
    )

    if ($S_ReferencedAuthStrengthIds.Count -gt 0) {
        Write-Host ("Resolving {0} referenced Authentication Strength policy(ies)..." -f $S_ReferencedAuthStrengthIds.Count) -ForegroundColor Cyan

        # Single-factor combination tokens that do NOT satisfy MFA on their own.
        $S_SingleFactorTokens = @('password', 'federatedSingleFactor')

        foreach ($S_AuthStrId in $S_ReferencedAuthStrengthIds) {
            try {
                $a = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/policies/authenticationStrengthPolicies/$S_AuthStrId"

                $combos = @($a.allowedCombinations)

                # MFA-capable if at least one combo has >1 factor, or its single
                # token is not in the single-factor list.
                $mfaCombos = @(
                    $combos | Where-Object {
                        $parts = ($_ -split ',') | ForEach-Object { $_.Trim() }
                        if ($parts.Count -gt 1) { return $true }
                        return ($parts[0] -notin $S_SingleFactorTokens)
                    }
                )

                $Script:S_ReferencedAuthStrengths[$S_AuthStrId] = [PSCustomObject]@{
                    Id                    = $a.id
                    DisplayName           = $a.displayName
                    Description           = $a.description
                    PolicyType            = $a.policyType        # builtIn | custom
                    RequirementsSatisfied = $a.requirementsSatisfied
                    AllowedCombinations   = $combos
                    MfaCombinations       = $mfaCombos
                    IsMfaCapable          = ($mfaCombos.Count -gt 0)
                }
            }
            catch {
                Write-Warning "Could not retrieve Authentication Strength '$S_AuthStrId': $_"
                $Script:S_ReferencedAuthStrengths[$S_AuthStrId] = [PSCustomObject]@{
                    Id                    = $S_AuthStrId
                    DisplayName           = '<unresolved>'
                    Description           = $null
                    PolicyType            = $null
                    RequirementsSatisfied = $null
                    AllowedCombinations   = @()
                    MfaCombinations       = @()
                    IsMfaCapable          = $null
                }
            }
        }

        $mfaCapableCount    = @($Script:S_ReferencedAuthStrengths.Values | Where-Object { $_.IsMfaCapable -eq $true  }).Count
        $notMfaCapableCount = @($Script:S_ReferencedAuthStrengths.Values | Where-Object { $_.IsMfaCapable -eq $false }).Count
        Write-Host ("Resolved {0} Authentication Strength policy(ies): {1} MFA-capable, {2} not MFA-capable." -f `
            $Script:S_ReferencedAuthStrengths.Count, $mfaCapableCount, $notMfaCapableCount) -ForegroundColor Green

        if ($notMfaCapableCount -gt 0) {
            Write-Warning "One or more CA policies reference an Authentication Strength that is NOT MFA-capable."
        }
    }
    else {
        Write-Host "No CA policies reference an Authentication Strength — skipping." -ForegroundColor DarkGray
    }

    # ── Retrieve Member Users (Guests excluded) ────────────────────────────────
    $userProperties = @(
        'Id'
        'DisplayName'
        'UserPrincipalName'
        'Mail'
        'SignInActivity'
        'AccountEnabled'
        'AssignedLicenses'
        'OnPremisesSyncEnabled'
        'UserType'
    ) -join ','

    Write-Host "Retrieving member users (Guests excluded)..." -ForegroundColor Cyan

    if ($Test) {
        Write-Host "[TEST MODE] Limiting to first 10 Member users." -ForegroundColor Yellow
        $users = Get-MgUser -Filter "userType eq 'Member'" -Property $userProperties -Top 10
    }
    else {
        $users = Get-MgUser -Filter "userType eq 'Member'" -Property $userProperties -All
    }

    $userCount = ($users | Measure-Object).Count
    Write-Host "Found $userCount Member users. Processing..." -ForegroundColor Cyan

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $currentUser = 0

    foreach ($user in $users) {
        $currentUser++
        Write-Progress -Activity "Processing Users" -Status "$currentUser of $userCount - $($user.DisplayName)" -PercentComplete (($currentUser / $userCount) * 100)

        # ── Authentication Methods ─────────────────────────────────────────────
        $authMethodError = $false
        try {
            $authMethods = Get-MgUserAuthenticationMethod -UserId $user.Id
        }
        catch {
            if ($_ -match 'accessDenied|403|Forbidden|Authorization failed') {
                Write-Warning "Access denied reading auth methods for $($user.UserPrincipalName) (privileged account?)"
                $authMethodError = $true
            }
            else {
                Write-Warning "Could not retrieve auth methods for $($user.UserPrincipalName): $_"
            }
            $authMethods = @()
        }

        # ── Determine MFA Registration ─────────────────────────────────────────
        $authTypes = $authMethods | ForEach-Object { $_.AdditionalProperties.'@odata.type' }

        $hasModernAuth = $authTypes | Where-Object {
            $_ -in @(
                '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod'
                '#microsoft.graph.fido2AuthenticationMethod'
                '#microsoft.graph.softwareOathAuthenticationMethod'
            )
        }
        $hasLegacyAuth = $authTypes | Where-Object {
            $_ -in @(
                '#microsoft.graph.phoneAuthenticationMethod'
            )
        }

        if ($authMethodError) {
            $mfaRegistered = 'Access Denied (Privileged Account)'
        }
        elseif ($hasModernAuth) {
            $mfaRegistered = 'Modern Auth'
        }
        elseif ($hasLegacyAuth) {
            $mfaRegistered = 'Legacy Auth'
        }
        else {
            $mfaRegistered = 'No MFA'
        }

        # ── Determine Active Account ───────────────────────────────────────────
        $lastInteractive    = $null
        $lastNonInteractive = $null
        if (-not $user.AccountEnabled) {
            $activeAccount = 'Disabled'
        }
        else {
            $lastInteractive    = $user.SignInActivity.LastSignInDateTime
            $lastNonInteractive = $user.SignInActivity.LastNonInteractiveSignInDateTime

            $dates = @($lastInteractive, $lastNonInteractive) | Where-Object { $_ -ne $null }

            if ($dates.Count -eq 0) {
                $activeAccount = 'No Sign-In Recorded'
            }
            else {
                $lastSignIn = ($dates | Sort-Object -Descending | Select-Object -First 1)
                if ($lastSignIn -ge $S_CutoffDate) {
                    $activeAccount = 'Yes'
                }
                else {
                    $daysAgo = [math]::Floor(((Get-Date) - $lastSignIn).TotalDays)
                    $activeAccount = "${daysAgo}+ Days Ago"
                }
            }
        }

        # ── Licensing & Sync ───────────────────────────────────────────────────
        $isLicensed = ($user.AssignedLicenses | Measure-Object).Count -gt 0
        $isOnPremSynced = $user.OnPremisesSyncEnabled -eq $true

        # ── Mail & Domain ──────────────────────────────────────────────────────
        $mail = if ($user.Mail) { $user.Mail } else { 'None' }
        $domain = if ($user.Mail) {
            ($user.Mail -split '@')[1]
        } else {
            ($user.UserPrincipalName -split '@')[1]
        }

        # ── Build result row ───────────────────────────────────────────────────
        $results.Add([PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            Mail              = $mail
            Domain            = $domain
            UserType          = $user.UserType
            MFA_Registered    = $mfaRegistered
            ActiveAccount     = $activeAccount
            IsLicensed        = $isLicensed
            IsOnPremSynced    = $isOnPremSynced
        })
    }

    Write-Progress -Activity "Processing Users" -Completed

    # ── Export CSV ─────────────────────────────────────────────────────────────
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nCSV exported to: $OutputPath" -ForegroundColor Green
    Write-Host "Total rows: $($results.Count)" -ForegroundColor Green

    # ── Calculate Statistics ───────────────────────────────────────────────────
    $totalMembers   = $results.Count
    $disabledCount  = ($results | Where-Object { $_.ActiveAccount -eq 'Disabled' }).Count
    $enabledCount   = $totalMembers - $disabledCount

    $enabledUsers       = $results | Where-Object { $_.ActiveAccount -ne 'Disabled' }
    $enabledModernAuth  = ($enabledUsers | Where-Object { $_.MFA_Registered -eq 'Modern Auth' }).Count
    $enabledLegacyAuth  = ($enabledUsers | Where-Object { $_.MFA_Registered -eq 'Legacy Auth' }).Count
    $enabledNoMFA       = ($enabledUsers | Where-Object { $_.MFA_Registered -eq 'No MFA' }).Count
    $enabledHasMFA      = $enabledModernAuth + $enabledLegacyAuth

    $coveragePercent = if ($enabledCount -gt 0) {
        [math]::Round(($enabledHasMFA / $enabledCount) * 100, 1)
    } else { 0 }

    # Build table rows for enabled users without MFA
    $noMfaUsers = $enabledUsers | Where-Object { $_.MFA_Registered -eq 'No MFA' }
    $tableRows = ($noMfaUsers | ForEach-Object {
        "        <tr><td>$([System.Web.HttpUtility]::HtmlEncode($_.DisplayName))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.UserPrincipalName))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.Mail))</td><td>$([System.Web.HttpUtility]::HtmlEncode($_.Domain))</td><td>$($_.ActiveAccount)</td><td>$($_.IsLicensed)</td><td>$($_.IsOnPremSynced)</td></tr>"
    }) -join "`n"

    # ── Build CA Policy table rows ──────────────────────────────────────
    function Format-CaList {
        param([object[]]$Items)
        if ($null -eq $Items -or $Items.Count -eq 0) { return '<span style="color:#999">—</span>' }
        ($Items | ForEach-Object { [System.Web.HttpUtility]::HtmlEncode([string]$_) }) -join '<br>'
    }

    $S_CaPolicyRows = if ($Script:S_MfaCaPolicies.Count -eq 0) {
        '        <tr><td colspan="9" style="text-align:center;color:#999">No MFA-enforcing Conditional Access policies were found.</td></tr>'
    }
    else {
        ($Script:S_MfaCaPolicies | ForEach-Object {
            $S_Included = @()
            $S_Included += $_.IncludeUsers
            $S_Included += $_.IncludeGroups
            $S_Included += $_.IncludeRoles
            $S_Excluded = @()
            $S_Excluded += $_.ExcludeUsers
            $S_Excluded += $_.ExcludeGroups
            $S_Excluded += $_.ExcludeRoles

            $S_Workloads = @()
            $S_Workloads += $_.IncludeApplications
            if ($_.IncludeUserActions.Count -gt 0) {
                $S_Workloads += ($_.IncludeUserActions | ForEach-Object { "action: $_" })
            }
            if ($_.ExcludeApplications.Count -gt 0) {
                $S_Workloads += ($_.ExcludeApplications | ForEach-Object { "exclude: $_" })
            }
            # Coverage-tier badge prepended to the Workloads cell
            $S_CoverageColor = switch ($_.WorkloadCoverage) {
                'Full'    { '#107c10' }   # green
                'Data'    { '#0078d4' }   # blue
                default   { '#ff8c00' }   # orange (Partial)
            }
            $S_CoverageBadge = "<span style=""display:inline-block;padding:2px 8px;border-radius:10px;background:$S_CoverageColor;color:#fff;font-size:11px;font-weight:600;margin-bottom:4px"">$($_.WorkloadCoverage) coverage</span>"
            $S_WorkloadsHtml = $S_CoverageBadge + '<br>' + (Format-CaList $S_Workloads)

            $S_GrantParts = @()
            if ($_.AllBuiltInControls.Count -gt 0) { $S_GrantParts += ($_.AllBuiltInControls -join " $($_.GrantOperator) ") }
            if ($_.HasAuthenticationStrength) {
                $S_AsName = if ($_.AuthenticationStrengthName) { $_.AuthenticationStrengthName } else { '<unknown>' }
                $S_GrantParts += "AuthStrength: $S_AsName"
            }
            $S_GrantText = ($S_GrantParts -join '<br>')
            if ($_.MfaIsBypassable) {
                $S_GrantText += '<br><span style="color:#d13438;font-weight:600">⚠ MFA bypassable (OR with companion controls)</span>'
            }

            $S_OtherTargeting = @()
            if ($_.IncludePlatforms.Count -gt 0) { $S_OtherTargeting += "Platforms incl: $($_.IncludePlatforms -join ', ')" }
            if ($_.ExcludePlatforms.Count -gt 0) { $S_OtherTargeting += "Platforms excl: $($_.ExcludePlatforms -join ', ')" }
            if ($_.IncludeLocations.Count -gt 0) { $S_OtherTargeting += "Locations incl: $($_.IncludeLocations -join ', ')" }
            if ($_.ExcludeLocations.Count -gt 0) { $S_OtherTargeting += "Locations excl: $($_.ExcludeLocations -join ', ')" }
            if ($_.ClientAppTypes.Count   -gt 0) { $S_OtherTargeting += "Client apps: $($_.ClientAppTypes -join ', ')" }
            if ($_.SignInRiskLevels.Count -gt 0) { $S_OtherTargeting += "Sign-in risk: $($_.SignInRiskLevels -join ', ')" }
            if ($_.UserRiskLevels.Count   -gt 0) { $S_OtherTargeting += "User risk: $($_.UserRiskLevels -join ', ')" }
            # Conditions-posture badge prepended to the Other Conditions cell.
            # RiskBased = neutral gray (de-emphasised — gap engine ignores these).
            $S_PostureColor = switch ($_.ConditionsPosture) {
                'Unrestricted' { '#107c10' }   # green
                'Constrained'  { '#ff8c00' }   # amber
                default        { '#6b6b6b' }   # neutral gray (RiskBased)
            }
            $S_PostureBadge = "<span style=""display:inline-block;padding:2px 8px;border-radius:10px;background:$S_PostureColor;color:#fff;font-size:11px;font-weight:600;margin-bottom:4px"">$($_.ConditionsPosture)</span>"
            $S_OtherInner   = if ($S_OtherTargeting.Count -gt 0) { Format-CaList $S_OtherTargeting } else { '<span style="color:#999">—</span>' }
            $S_OtherText    = $S_PostureBadge + '<br>' + $S_OtherInner

            # Enforcement-tier badge under the policy name. Gray = Ignored by gap engine.
            $S_TierColor = switch ($_.EnforcementTier) {
                'Ideal'      { '#107c10' }   # green
                'Acceptable' { '#ff8c00' }   # amber
                default      { '#6b6b6b' }   # neutral gray (Ignored)
            }
            # Place the tier badge ABOVE the name for consistency with Workloads and Conditions columns.
            $S_TierBadge = "<span style=""display:inline-block;padding:2px 8px;border-radius:10px;background:$S_TierColor;color:#fff;font-size:11px;font-weight:600;margin-bottom:4px"">$($_.EnforcementTier)</span>"
            $S_NameCell  = $S_TierBadge + '<br>' + [System.Web.HttpUtility]::HtmlEncode($_.DisplayName)

            # data-* attributes drive the client-side filter dropdowns above the table.
            "        <tr data-tier=""$($_.EnforcementTier)"" data-coverage=""$($_.WorkloadCoverage)"" data-posture=""$($_.ConditionsPosture)""><td>$S_NameCell</td>" +
            "<td>$($_.State)</td>" +
            "<td>$S_WorkloadsHtml</td>" +
            "<td>$(Format-CaList $S_Included)</td>" +
            "<td>$(Format-CaList $S_Excluded)</td>" +
            "<td>$S_GrantText</td>" +
            "<td>$S_OtherText</td></tr>"
        }) -join "`n"
    }

    # ── Build Authentication Strength table rows ──────────────────────────────────
    $S_AuthStrengthRows = if ($null -eq $Script:S_ReferencedAuthStrengths -or $Script:S_ReferencedAuthStrengths.Count -eq 0) {
        '        <tr><td colspan="6" style="text-align:center;color:#999">No Authentication Strengths referenced by the kept CA policies.</td></tr>'
    }
    else {
        ($Script:S_ReferencedAuthStrengths.Values | ForEach-Object {
            $S_MfaText = switch ($_.IsMfaCapable) {
                $true   { '<span style="color:#107c10;font-weight:600">Yes</span>' }
                $false  { '<span style="color:#d13438;font-weight:600">No</span>' }
                default { '<span style="color:#999">Unknown</span>' }
            }
            "        <tr><td>$([System.Web.HttpUtility]::HtmlEncode($_.DisplayName))</td>" +
            "<td>$($_.PolicyType)</td>" +
            "<td>$S_MfaText</td>" +
            "<td>$(Format-CaList $_.AllowedCombinations)</td>" +
            "<td>$(Format-CaList $_.MfaCombinations)</td>" +
            "<td>$([System.Web.HttpUtility]::HtmlEncode([string]$_.Description))</td></tr>"
        }) -join "`n"
    }

    # ── Generate HTML Report ───────────────────────────────────────────────────
    $reportDate = Get-Date -Format 'dd MMM yyyy HH:mm'
    $S_TenantId = (Get-MgContext).TenantId

    if ($null -eq $Script:S_SecurityDefaultsEnabled) {
        $S_SecDefaultsText  = 'Security Defaults: Unknown'
        $S_SecDefaultsClass = 'unknown'
    }
    elseif ($Script:S_SecurityDefaultsEnabled) {
        $S_SecDefaultsText  = 'Security Defaults: ENABLED'
        $S_SecDefaultsClass = 'enabled'
    }
    else {
        $S_SecDefaultsText  = 'Security Defaults: DISABLED'
        $S_SecDefaultsClass = 'disabled'
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Member MFA Coverage Report</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; padding: 24px; }
        .header { text-align: center; margin-bottom: 32px; }
        .header h1 { font-size: 28px; color: #1a1a2e; margin-bottom: 4px; }
        .header .subtitle { font-size: 14px; color: #666; }
        .cards { display: flex; flex-wrap: wrap; gap: 20px; justify-content: center; margin-bottom: 32px; }
        .card {
            background: #fff; border-radius: 12px; padding: 24px 28px; min-width: 220px; flex: 1; max-width: 280px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08); border-left: 5px solid #0078d4; position: relative;
        }
        .card .label { font-size: 13px; color: #666; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px; }
        .card .value { font-size: 36px; font-weight: 700; color: #1a1a2e; }
        .card .detail { font-size: 12px; color: #888; margin-top: 6px; }
        .card.blue    { border-left-color: #0078d4; }
        .card.red     { border-left-color: #d13438; }
        .card.green   { border-left-color: #107c10; }
        .card.orange  { border-left-color: #ff8c00; }
        .card.purple  { border-left-color: #8764b8; }
        .card.alert   { border-left-color: #d13438; background: #fdf2f2; }
        .card.alert .value { color: #d13438; }
        .section { margin-bottom: 24px; }
        .section h2 { font-size: 18px; color: #1a1a2e; margin-bottom: 12px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0; text-align: center; }
        .breakdown { display: flex; flex-wrap: wrap; gap: 16px; justify-content: center; }
        .breakdown .card { min-width: 180px; max-width: 240px; padding: 18px 22px; }
        .breakdown .card .value { font-size: 28px; }
        table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.08); margin-top: 12px; }
        th { background: #d13438; color: #fff; text-align: left; padding: 10px 14px; font-size: 13px; text-transform: uppercase; letter-spacing: 0.3px; }
        td { padding: 9px 14px; font-size: 13px; border-bottom: 1px solid #eee; vertical-align: top; }
        tr:hover { background: #fdf2f2; }
        .footer { text-align: center; font-size: 12px; color: #999; margin-top: 32px; }
        .secdefaults { text-align: center; font-size: 14px; font-weight: 700; padding: 10px 16px; border-radius: 6px; margin: 0 auto 20px; max-width: 480px; }
        .secdefaults.enabled  { background: #fdecea; color: #d13438; border: 1px solid #d13438; }
        .secdefaults.disabled { background: #eaf6ec; color: #107c10; border: 1px solid #107c10; }
        .secdefaults.unknown  { background: #f3f3f3; color: #666;    border: 1px solid #999; }
        .ca-filters { display: flex; flex-wrap: wrap; gap: 12px; justify-content: center; align-items: center; margin: 8px 0 4px; font-size: 13px; }
        .ca-filters label { font-weight: 600; color: #555; margin-right: 4px; }
        .ca-filters select { padding: 4px 8px; border: 1px solid #ccc; border-radius: 6px; background: #fff; font-size: 13px; cursor: pointer; }
        .ca-filters button { padding: 4px 12px; border: 1px solid #d13438; background: #fff; color: #d13438; border-radius: 6px; font-size: 12px; font-weight: 600; cursor: pointer; }
        .ca-filters button:hover { background: #fdecea; }
    </style>
    <script>
        function caApplyFilters() {
            var tier     = document.getElementById('caFilterTier').value;
            var coverage = document.getElementById('caFilterCoverage').value;
            var posture  = document.getElementById('caFilterPosture').value;
            var rows = document.querySelectorAll('#caPolicyTable tbody tr');
            var visible = 0;
            rows.forEach(function (r) {
                if (!r.hasAttribute('data-tier')) { return; } // skip empty-state placeholder
                var ok = (tier === '' || r.getAttribute('data-tier') === tier)
                      && (coverage === '' || r.getAttribute('data-coverage') === coverage)
                      && (posture === '' || r.getAttribute('data-posture') === posture);
                r.style.display = ok ? '' : 'none';
                if (ok) { visible++; }
            });
            var count = document.getElementById('caFilterCount');
            if (count) { count.textContent = visible + ' policy(ies) shown'; }
        }
        function caResetFilters() {
            document.getElementById('caFilterTier').value = '';
            document.getElementById('caFilterCoverage').value = '';
            document.getElementById('caFilterPosture').value = '';
            caApplyFilters();
        }
        document.addEventListener('DOMContentLoaded', caApplyFilters);
    </script>
</head>
<body>
    <div class="header">
        <h1>Member MFA Coverage Report</h1>
        <div class="subtitle">Generated: $reportDate | Tenant: $S_TenantId | Inactive threshold: $InactiveDays days | Guests excluded</div>
    </div>

    <div class="secdefaults $S_SecDefaultsClass">$S_SecDefaultsText</div>

    <div class="cards">
        <div class="card blue">
            <div class="label">Total Member Users</div>
            <div class="value">$totalMembers</div>
        </div>
        <div class="card green">
            <div class="label">Enabled Accounts</div>
            <div class="value">$enabledCount</div>
            <div class="detail">$([math]::Round(($enabledCount / [math]::Max($totalMembers,1)) * 100, 1))% of total</div>
        </div>
        <div class="card red">
            <div class="label">Disabled Accounts</div>
            <div class="value">$disabledCount</div>
            <div class="detail">$([math]::Round(($disabledCount / [math]::Max($totalMembers,1)) * 100, 1))% of total</div>
        </div>
        <div class="card purple">
            <div class="label">MFA Coverage (Enabled)</div>
            <div class="value">$coveragePercent%</div>
            <div class="detail">$enabledHasMFA of $enabledCount enabled members</div>
        </div>
    </div>

    <div class="section">
        <h2>Enabled Accounts &mdash; MFA Breakdown</h2>
        <div class="breakdown">
            <div class="card green">
                <div class="label">Modern Auth (MFA)</div>
                <div class="value">$enabledModernAuth</div>
                <div class="detail">Authenticator / Passkey / TOTP</div>
            </div>
            <div class="card orange">
                <div class="label">Legacy Auth (MFA)</div>
                <div class="value">$enabledLegacyAuth</div>
                <div class="detail">SMS / Voice</div>
            </div>
            <div class="card purple">
                <div class="label">Total with MFA</div>
                <div class="value">$enabledHasMFA</div>
                <div class="detail">$([math]::Round(($enabledHasMFA / [math]::Max($enabledCount,1)) * 100, 1))% of enabled</div>
            </div>
            <div class="card red">
                <div class="label">No MFA Registered</div>
                <div class="value">$enabledNoMFA</div>
                <div class="detail">$([math]::Round(($enabledNoMFA / [math]::Max($enabledCount,1)) * 100, 1))% of enabled</div>
            </div>
        </div>
    </div>

    <div class="section">
        <h2>Enabled Member Users without MFA</h2>
        <table>
            <thead>
                <tr><th>Display Name</th><th>UPN</th><th>Mail</th><th>Domain</th><th>Active</th><th>Licensed</th><th>On-Prem Synced</th></tr>
            </thead>
            <tbody>
$tableRows
            </tbody>
        </table>
    </div>

    <div class="section">
        <h2>MFA-Enforcing Conditional Access Policies</h2>
        <div class="ca-filters">
            <label for="caFilterTier">Tier:</label>
            <select id="caFilterTier" onchange="caApplyFilters()">
                <option value="">All</option>
                <option value="Ideal">Ideal</option>
                <option value="Acceptable">Acceptable</option>
                <option value="Ignored">Ignored</option>
            </select>
            <label for="caFilterCoverage">Workloads:</label>
            <select id="caFilterCoverage" onchange="caApplyFilters()">
                <option value="">All</option>
                <option value="Full">Full coverage</option>
                <option value="Data">Data coverage</option>
                <option value="Partial">Partial coverage</option>
            </select>
            <label for="caFilterPosture">Conditions:</label>
            <select id="caFilterPosture" onchange="caApplyFilters()">
                <option value="">All</option>
                <option value="Unrestricted">Unrestricted</option>
                <option value="Constrained">Constrained</option>
                <option value="RiskBased">RiskBased</option>
            </select>
            <button type="button" onclick="caResetFilters()">Reset</button>
            <span id="caFilterCount" style="color:#666;font-size:12px;"></span>
        </div>
        <table id="caPolicyTable">
            <thead>
                <tr><th>Name</th><th>State</th><th>Workloads</th><th>Included</th><th>Excluded</th><th>Grants</th><th>Other Conditions</th></tr>
            </thead>
            <tbody>
$S_CaPolicyRows
            </tbody>
        </table>
    </div>

    <div class="section">
        <h2>Authentication Strengths Referenced by CA Policies</h2>
        <table>
            <thead>
                <tr><th>Name</th><th>Type</th><th>MFA Capable</th><th>Allowed Combinations</th><th>MFA Combinations</th><th>Description</th></tr>
            </thead>
            <tbody>
$S_AuthStrengthRows
            </tbody>
        </table>
    </div>

    <div class="footer">
        CSV data: $(Split-Path $OutputPath -Leaf) | Report generated by ReportMemberMFA.ps1
    </div>
</body>
</html>
"@

    $html | Out-File -FilePath $S_HtmlPath -Encoding UTF8
    Write-Host "HTML report exported to: $S_HtmlPath" -ForegroundColor Green
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    # ── Disconnect ─────────────────────────────────────────────────────────────
    $S_DisconnectChoice = Read-Host "`nDisconnect from Microsoft Graph? [Y] Yes  [N] Keep session  (Default: N)"
    if ($S_DisconnectChoice -eq 'Y') {
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Write-Host "Disconnected." -ForegroundColor Green
    }
    else {
        Write-Host "Graph session kept alive." -ForegroundColor Green
    }
}
