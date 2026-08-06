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

if (-not $OutputPath)
{
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

$S_GraphRequestDelayMilliseconds = 5

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
        Write-Host "Disconnecting existing session..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Write-Host "Reconnecting with required scopes..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome
        Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
    }
    else
    {
        Write-Host "Using existing Graph session." -ForegroundColor Green
    }
}
else
{
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome
    Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
}
$S_ActiveContext = Get-MgContext
Write-Host ""
Write-Host "Active Graph context:" -ForegroundColor Cyan
Write-Host "  Account    : $($S_ActiveContext.Account)" -ForegroundColor Cyan
Write-Host "  TenantId   : $($S_ActiveContext.TenantId)" -ForegroundColor Cyan
Write-Host "  Environment: $($S_ActiveContext.Environment)" -ForegroundColor Cyan
Write-Host ""

$S_ContextConfirmation = Read-Host "Proceed with this Graph context? [Y] Yes  [N] No  (Default: N)"
if ([string]::IsNullOrWhiteSpace($S_ContextConfirmation))
{
    $S_ContextConfirmation = 'N'
}
else
{
    $S_ContextConfirmation = $S_ContextConfirmation.ToUpperInvariant()
}
if ($S_ContextConfirmation -ne 'Y')
{
    throw "Operation cancelled. Please reconnect to the correct tenant and account, then run again."
}

try
{
    # ── Check Security Defaults Status ─────────────────────────────────────────
    Write-Host "Checking Security Defaults status..." -ForegroundColor Cyan
    $Script:S_SecurityDefaultsEnabled = $null
    try
    {
        $S_SecDefaultsPolicy = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy'
        $Script:S_SecurityDefaultsEnabled = [bool]$S_SecDefaultsPolicy.isEnabled
        Write-Host "Security Defaults Enabled: $Script:S_SecurityDefaultsEnabled" -ForegroundColor $(if ($Script:S_SecurityDefaultsEnabled)
            {
                'Red' } else
            {
                'Green' })
    }
    catch
    {
        Write-Warning "Could not retrieve Security Defaults policy: $_"
    }

    # ── Conditional Access — MFA-enforcing policies (only when Security Defaults is OFF) ─────────
    $Script:S_MfaCaPolicies     = @()
    $Script:S_UserExclusionMap  = @{}   # UserId -> List[string] of Ideal-policy ReportIds excluding the user
    if ($Script:S_SecurityDefaultsEnabled -eq $false)
    {
        Write-Host "Retrieving enabled Conditional Access policies..." -ForegroundColor Cyan
        try
        {
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
                    if ($null -eq $S_Grant)
                    {
                        return $false }
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

            function Resolve-CaUserId
            {
                param([string]$F_Id)
                if ([string]::IsNullOrWhiteSpace($F_Id))
                {
                    return $null }
                if ($F_Id -in @('All', 'None', 'GuestsOrExternalUsers'))
                {
                    return $F_Id }
                if ($S_UserCache.ContainsKey($F_Id))
                {
                    return $S_UserCache[$F_Id] }
                try
                {
                    $F_U = Get-MgUser -UserId $F_Id -Property Id,DisplayName,UserPrincipalName -ErrorAction Stop
                    $F_Label = "$($F_U.DisplayName) <$($F_U.UserPrincipalName)>"
                }
                catch
                {
                    $F_Label = "<unresolved:$F_Id>" }
                $S_UserCache[$F_Id] = $F_Label
                return $F_Label
            }

            function Resolve-CaGroupId
            {
                param([string]$F_Id)
                if ([string]::IsNullOrWhiteSpace($F_Id))
                {
                    return $null }
                if ($S_GroupCache.ContainsKey($F_Id))
                {
                    return $S_GroupCache[$F_Id] }
                try
                {
                    $F_G = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$F_Id`?`$select=id,displayName" -ErrorAction Stop
                    $F_Label = "$($F_G.displayName) [group]"
                }
                catch
                {
                    $F_Label = "<unresolved:$F_Id>" }
                $S_GroupCache[$F_Id] = $F_Label
                return $F_Label
            }

            function Resolve-CaRoleId
            {
                param([string]$F_Id)
                if ([string]::IsNullOrWhiteSpace($F_Id))
                {
                    return $null }
                if ($S_RoleCache.ContainsKey($F_Id))
                {
                    return $S_RoleCache[$F_Id] }
                try
                {
                    $F_R = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/directoryRoleTemplates/$F_Id" -ErrorAction Stop
                    $F_Label = "$($F_R.displayName) [role]"
                }
                catch
                {
                    $F_Label = "<unresolved:$F_Id>" }
                $S_RoleCache[$F_Id] = $F_Label
                return $F_Label
            }

            $S_LocationCache = @{}
            function Resolve-CaLocationId
            {
                param([string]$F_Id)
                if ([string]::IsNullOrWhiteSpace($F_Id))
                {
                    return $null }
                if ($F_Id -in @('All', 'AllTrusted', 'MultiFactorAuthentication'))
                {
                    return $F_Id }
                if ($S_LocationCache.ContainsKey($F_Id))
                {
                    return $S_LocationCache[$F_Id] }
                try
                {
                    $F_L = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations/$F_Id" -ErrorAction Stop
                    $F_Label = "$($F_L.displayName) [location]"
                }
                catch
                {
                    $F_Label = "<unresolved:$F_Id>" }
                $S_LocationCache[$F_Id] = $F_Label
                return $F_Label
            }

            # ── Project enriched policy objects ────────────────────────────────────
            # Enrich EVERY MFA-enforcing policy first (regardless of audience).
            # The audience filter (Members vs Guests) is applied AFTER enrichment
            # so that a future guest-focused script can reuse the same projection
            # by simply flipping the filter predicate from TargetsMembers to
            # TargetsGuests — no rewrite of the resolver/enrichment code required.
            $S_EnrichedCaPolicies = @(
                foreach ($S_P in $S_FilteredCaPolicies)
                {
                    $S_Apps  = $S_P.Conditions.Applications
                    $S_Usrs  = $S_P.Conditions.Users                    
                    $S_Plat  = $S_P.Conditions.Platforms
                    $S_Locs  = $S_P.Conditions.Locations
                    $S_GrantTypes = @()
                    if ($S_P.GrantControls.BuiltInControls -contains 'mfa')
                    {
                        $S_GrantTypes += 'MFA' }
                    $S_HasAuthStr = (
                        $null -ne $S_P.GrantControls.AuthenticationStrength -and
                        -not [string]::IsNullOrWhiteSpace([string]$S_P.GrantControls.AuthenticationStrength.Id)
                    )
                    if ($S_HasAuthStr)
                    {
                        $S_GrantTypes += 'AuthStrength' }

                    # ── Grant operator & companion controls ──────────────────────
                    # Operator is 'AND' or 'OR'. With OR + other controls present,
                    # a user can satisfy the policy WITHOUT MFA (e.g. compliant
                    # device alone) — a potential MFA gap we must flag.
                    $S_GrantOperator = $S_P.GrantControls.Operator   # 'AND' | 'OR' | $null
                    $S_AllBuiltIn    = @($S_P.GrantControls.BuiltInControls)
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
                    $S_IncAppsRaw = @($S_Apps.IncludeApplications)
                    $S_WorkloadCoverage = if ($S_IncAppsRaw -contains 'All')
                    {
                        'Full'
                    } elseif ($S_IncAppsRaw -contains 'Office365')
                    {
                        'Data'
                    } else
                    {
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
                        @($S_P.Conditions.SignInRiskLevels | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0 -or
                        @($S_P.Conditions.UserRiskLevels   | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0
                    )
                    $S_ClientAppsRaw = @($S_P.Conditions.ClientAppTypes | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                    $S_ClientAppsIsAll = (
                        $S_ClientAppsRaw.Count -eq 0 -or
                        ($S_ClientAppsRaw.Count -eq 1 -and $S_ClientAppsRaw[0] -eq 'all')
                    )
                    $S_HasExcludedApps  = @($S_Apps.ExcludeApplications | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0
                    # An Include* array is "Any" (not narrowing) when it is empty
                    # or contains the single sentinel 'all'/'All'. Graph returns
                    # IncludePlatforms = ['all'] for "Any device platform" and
                    # IncludeLocations = ['All'] for "Any location" — both of
                    # which are the default scopes and must NOT be flagged.
                    # We also strip nulls/blanks because `@($S_Plat.IncludePlatforms)`
                    # becomes `@($null)` (Count=1) when the Platforms condition
                    # block itself is absent.
                    $S_IncPlat = @($S_Plat.IncludePlatforms | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                    $S_IncLoc  = @($S_Locs.IncludeLocations | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                    $S_ExcPlat = @($S_Plat.ExcludePlatforms | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                    $S_ExcLoc  = @($S_Locs.ExcludeLocations | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
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
                    $S_ConditionsPosture = if ($S_HasRisk)
                    {
                        'RiskBased'
                    } elseif ($S_HasNarrowing)
                    {
                        'Constrained'
                    } else
                    {
                        'Unrestricted'
                    }

                    # EnforcementTier is computed AFTER Persona (below) because
                    # 'Ideal' requires a broad audience (AllUsers / Internal) and
                    # Guests are excluded from tiering entirely.

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
                    $S_IncUsersRaw   = @($S_Usrs.IncludeUsers)
                    $S_HasUsersAll   = ($S_IncUsersRaw -contains 'All')
                    $S_HasUsersGuestToken = ($S_IncUsersRaw -contains 'GuestsOrExternalUsers')
                    $S_HasSpecificUsers   = @(
                        $S_IncUsersRaw | Where-Object {
                            $_ -and $_ -notin @('All', 'None', 'GuestsOrExternalUsers')
                        }
                    ).Count -gt 0
                    $S_HasIncGroups  = @($S_Usrs.IncludeGroups).Count -gt 0
                    $S_HasIncRoles   = @($S_Usrs.IncludeRoles).Count  -gt 0
                    $S_IncGuestObj   = $S_Usrs.IncludeGuestsOrExternalUsers
                    $S_IncGuestTypes = @()
                    if ($null -ne $S_IncGuestObj -and -not [string]::IsNullOrWhiteSpace([string]$S_IncGuestObj.GuestOrExternalUserTypes))
                    {
                        $S_IncGuestTypes = @(([string]$S_IncGuestObj.GuestOrExternalUserTypes -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                    }
                    $S_HasIncGuestSpec = $S_IncGuestTypes.Count -gt 0

                    # Mirror for the EXCLUDE side — used by Persona classification
                    # to detect AllUsers policies that explicitly carve guests out.
                    $S_ExcGuestObj   = $S_Usrs.ExcludeGuestsOrExternalUsers
                    $S_ExcGuestTypes = @()
                    if ($null -ne $S_ExcGuestObj -and -not [string]::IsNullOrWhiteSpace([string]$S_ExcGuestObj.GuestOrExternalUserTypes))
                    {
                        $S_ExcGuestTypes = @(([string]$S_ExcGuestObj.GuestOrExternalUserTypes -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                    }
                    $S_HasExcGuestSpec = $S_ExcGuestTypes.Count -gt 0
                    $S_HasExcRoles     = @($S_Usrs.ExcludeRoles).Count -gt 0

                    # ── Persona ─────────────────────────────────────────────────
                    # Classify the AUDIENCE shape of the policy. Conditions like
                    # apps/locations/risk are NOT considered here — they live in
                    # WorkloadCoverage / ConditionsPosture. Persona is purely
                    # about who the policy is targeting.
                    #   • AllUsers — IncludeUsers='All' with no admin/guest carve-out
                    #                (user/group excludes don't change this verdict)
                    #   • Internal — IncludeUsers='All' AND (ExcludeRoles OR ExcludeGuests)
                    #   • Admins   — role-only include scope
                    #   • Guests   — guest-only include scope
                    #   • Targeted — specific users and/or groups only
                    #   • Mixed    — combinations that don't cleanly match above
                    $S_Persona = if ($S_HasUsersAll)
                    {
                        if ($S_HasExcRoles -or $S_HasExcGuestSpec)
                        {
                            'Internal' } else
                        {
                            'AllUsers' }
                    }
                    elseif ($S_HasIncRoles -and -not $S_HasSpecificUsers -and -not $S_HasIncGroups -and -not $S_HasIncGuestSpec -and -not $S_HasUsersGuestToken)
                    {
                        'Admins'
                    }
                    elseif (($S_HasIncGuestSpec -or $S_HasUsersGuestToken) -and -not $S_HasSpecificUsers -and -not $S_HasIncGroups -and -not $S_HasIncRoles)
                    {
                        'Guests'
                    }
                    elseif (($S_HasSpecificUsers -or $S_HasIncGroups) -and -not $S_HasIncRoles -and -not $S_HasIncGuestSpec -and -not $S_HasUsersGuestToken)
                    {
                        'Targeted'
                    }
                    else
                    {
                        'Mixed'
                    }

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

                    # ── Enforcement tier (gap-engine eligibility) ─────────────
                    # Pre-computed verdict the per-user MFA gap engine will use
                    # to decide which CA policies to count as meaningful MFA
                    # enforcement for an everyday Member sign-in. Persona-aware:
                    #   • Ideal      → Full coverage + Unrestricted conditions
                    #                  AND Persona ∈ (AllUsers, Internal)
                    #   • Ignored    → Persona = Guests (out of scope for member MFA)
                    #                  OR Partial coverage / RiskBased posture
                    #   • Acceptable → everything in between (Full|Data) +
                    #                  (Unrestricted|Constrained), non-Guest persona
                    $S_EnforcementTier = if ($S_Persona -eq 'Guests')
                    {
                        'Ignored'
                    } elseif ($S_WorkloadCoverage -eq 'Full' -and $S_ConditionsPosture -eq 'Unrestricted' -and $S_Persona -in @('AllUsers','Internal'))
                    {
                        'Ideal'
                    } elseif ($S_WorkloadCoverage -in @('Full','Data') -and $S_ConditionsPosture -in @('Unrestricted','Constrained'))
                    {
                        'Acceptable'
                    } else
                    {
                        'Ignored'
                    }

                    [PSCustomObject]@{
                        Id                         = $S_P.Id
                        DisplayName                = $S_P.DisplayName
                        State                      = $S_P.State
                        GrantType                  = ($S_GrantTypes -join ' + ')
                        AuthenticationStrengthId   = $S_P.GrantControls.AuthenticationStrength.Id
                        AuthenticationStrengthName = $S_P.GrantControls.AuthenticationStrength.DisplayName
                        GrantOperator              = $S_GrantOperator
                        AllBuiltInControls         = $S_AllBuiltIn
                        CompanionControls          = $S_CompanionControls
                        HasAuthenticationStrength  = $S_HasAuthStr
                        MfaIsBypassable            = $S_MfaIsBypassable
                        WorkloadCoverage           = $S_WorkloadCoverage
                        ConditionsPosture          = $S_ConditionsPosture
                        EnforcementTier            = $S_EnforcementTier
                        Persona                    = $S_Persona
                        TargetsMembers             = $S_TargetsMembers
                        TargetsGuests              = $S_TargetsGuests
                        IncludeApplications        = @($S_Apps.IncludeApplications)
                        ExcludeApplications        = @($S_Apps.ExcludeApplications)
                        IncludeUserActions         = @($S_Apps.IncludeUserActions)
                        IncludeUsers               = @($S_Usrs.IncludeUsers  | ForEach-Object { Resolve-CaUserId  $_ })
                        ExcludeUsers               = @($S_Usrs.ExcludeUsers  | ForEach-Object { Resolve-CaUserId  $_ })
                        ExcludeUserIds             = @($S_Usrs.ExcludeUsers  | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                        IncludeGroups              = @($S_Usrs.IncludeGroups | ForEach-Object { Resolve-CaGroupId $_ })
                        ExcludeGroups              = @($S_Usrs.ExcludeGroups | ForEach-Object { Resolve-CaGroupId $_ })
                        ExcludeGroupIds            = @($S_Usrs.ExcludeGroups | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                        IncludeRoles               = @($S_Usrs.IncludeRoles  | ForEach-Object { Resolve-CaRoleId  $_ })
                        ExcludeRoles               = @($S_Usrs.ExcludeRoles  | ForEach-Object { Resolve-CaRoleId  $_ })
                        IncludeGuestsOrExternalUserTypes = $S_IncGuestTypes
                        ExcludeGuestsOrExternalUserTypes = $S_ExcGuestTypes
                        IncludePlatforms           = @($S_Plat.IncludePlatforms | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                        ExcludePlatforms           = @($S_Plat.ExcludePlatforms | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                        IncludeLocations           = @($S_Locs.IncludeLocations | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { Resolve-CaLocationId $_ })
                        ExcludeLocations           = @($S_Locs.ExcludeLocations | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { Resolve-CaLocationId $_ })
                        ClientAppTypes             = @($S_P.Conditions.ClientAppTypes | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                        SignInRiskLevels           = @($S_P.Conditions.SignInRiskLevels | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                        UserRiskLevels             = @($S_P.Conditions.UserRiskLevels | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
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
            $S_IdealN      = if ($S_TierCounts -and $S_TierCounts['Ideal'])
            {
                @($S_TierCounts['Ideal']).Count }      else
            {
                0 }
            $S_AcceptableN = if ($S_TierCounts -and $S_TierCounts['Acceptable'])
            {
                @($S_TierCounts['Acceptable']).Count } else
            {
                0 }
            $S_IgnoredN    = if ($S_TierCounts -and $S_TierCounts['Ignored'])
            {
                @($S_TierCounts['Ignored']).Count }    else
            {
                0 }
            Write-Host ("  Enforcement tiers — Ideal: {0}, Acceptable: {1}, Ignored: {2} (gap engine will skip Ignored)." -f `
                    $S_IdealN, $S_AcceptableN, $S_IgnoredN) -ForegroundColor Cyan

            # ── Persona summary (audience shape) ───────────────────────
            $S_PersonaCounts = $Script:S_MfaCaPolicies | Group-Object Persona -AsHashTable -AsString
            $S_PersonaParts = foreach ($S_PName in 'AllUsers','Internal','Admins','Guests','Targeted','Mixed')
            {
                $S_PN = if ($S_PersonaCounts -and $S_PersonaCounts[$S_PName])
                {
                    @($S_PersonaCounts[$S_PName]).Count } else
                {
                    0 }
                "{0}: {1}" -f $S_PName, $S_PN
            }
            Write-Host ("  Personas — " + ($S_PersonaParts -join ', ')) -ForegroundColor Cyan

            # ── Assign Report IDs (sequential, zero-padded 3 digits) ─────────
            # Stable label per policy for cross-referencing from the per-user
            # exclusion column (e.g. CaExclusionTags = "001, 003").
            $S_ReportIdx = 0
            foreach ($S_Pol in $Script:S_MfaCaPolicies)
            {
                $S_ReportIdx++
                $S_Pol | Add-Member -NotePropertyName ReportId -NotePropertyValue ('{0:D3}' -f $S_ReportIdx) -Force
            }

            # ── Per-user exclusion resolution (Ideal policies only) ─────────
            # For each Ideal-tier policy, expand:
            #   • ExcludeUsers  (direct user GUIDs)
            #   • ExcludeGroups (transitive members — nested groups handled
            #                     server-side by Graph's transitiveMembers)
            # Producing a map: UserId -> @(ReportId,...). Only Member users that
            # are AccountEnabled are recorded (matches the user-base filter used
            # later when iterating $S_Users). PIM-eligible role members and
            # PIM-for-Groups eligible assignees are NOT included — transitiveMembers
            # only returns active members. That gap will be addressed by a future
            # ReportAdminMFA.ps1 pipeline.
            $S_IdealPolicies = @($Script:S_MfaCaPolicies | Where-Object { $_.EnforcementTier -eq 'Ideal' })
            if ($S_IdealPolicies.Count -gt 0)
            {
                Write-Host ("Resolving exclusions for {0} Ideal policy(ies)..." -f $S_IdealPolicies.Count) -ForegroundColor Cyan
                $S_GroupMemberCache = @{}   # GroupId -> List[string] of member UserIds
                foreach ($S_IP in $S_IdealPolicies)
                {
                    $S_ExclUserIds = New-Object 'System.Collections.Generic.HashSet[string]'

                    # Direct user exclusions
                    foreach ($S_Uid in @($S_IP.ExcludeUserIds | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }))
                    {
                        [void]$S_ExclUserIds.Add([string]$S_Uid)
                    }

                    # Group exclusions — transitively expanded
                    foreach ($S_Gid in @($S_IP.ExcludeGroupIds | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }))
                    {
                        if (-not $S_GroupMemberCache.ContainsKey($S_Gid))
                        {
                            $S_Members = New-Object 'System.Collections.Generic.List[string]'
                            try
                            {
                                $S_Uri = "https://graph.microsoft.com/v1.0/groups/$S_Gid/transitiveMembers/microsoft.graph.user`?`$select=id,userType,accountEnabled&`$top=999"
                                do
                                {
                                    $S_Resp = Invoke-MgGraphRequest -Method GET -Uri $S_Uri -ErrorAction Stop
                                    foreach ($S_M in @($S_Resp.value))
                                    {
                                        if ($S_M.userType -eq 'Member' -and $S_M.accountEnabled -eq $true)
                                        {
                                            $S_Members.Add([string]$S_M.id)
                                        }
                                    }
                                    $S_Uri = $S_Resp.'@odata.nextLink'
                                } while ($S_Uri)
                            }
                            catch
                            {
                                Write-Warning ("Could not expand group {0} for policy {1}: {2}" -f $S_Gid, $S_IP.ReportId, $_)
                            }
                            $S_GroupMemberCache[$S_Gid] = $S_Members
                        }
                        foreach ($S_Mid in $S_GroupMemberCache[$S_Gid])
                        {
                            [void]$S_ExclUserIds.Add($S_Mid)
                        }
                    }

                    foreach ($S_Uid in $S_ExclUserIds)
                    {
                        if (-not $Script:S_UserExclusionMap.ContainsKey($S_Uid))
                        {
                            $Script:S_UserExclusionMap[$S_Uid] = New-Object 'System.Collections.Generic.List[string]'
                        }
                        $Script:S_UserExclusionMap[$S_Uid].Add($S_IP.ReportId)
                    }
                }
                Write-Host ("  {0} user(s) tagged with Ideal-policy exclusions." -f $Script:S_UserExclusionMap.Count) -ForegroundColor Cyan
            }
        }
        catch
        {
            Write-Warning "Could not retrieve Conditional Access policies: $_"
        }
    }
    else
    {
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

    if ($S_ReferencedAuthStrengthIds.Count -gt 0)
    {
        Write-Host ("Resolving {0} referenced Authentication Strength policy(ies)..." -f $S_ReferencedAuthStrengthIds.Count) -ForegroundColor Cyan

        # Single-factor combination tokens that do NOT satisfy MFA on their own.
        $S_SingleFactorTokens = @('password', 'federatedSingleFactor')

        foreach ($S_AuthStrId in $S_ReferencedAuthStrengthIds)
        {
            try
            {
                $S_A = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/policies/authenticationStrengthPolicies/$S_AuthStrId"

                $S_Combos = @($S_A.allowedCombinations)

                # MFA-capable if at least one combo has >1 factor, or its single
                # token is not in the single-factor list.
                $S_MfaCombos = @(
                    $S_Combos | Where-Object {
                        $S_Parts = ($_ -split ',') | ForEach-Object { $_.Trim() }
                        if ($S_Parts.Count -gt 1)
                        {
                            return $true }
                        return ($S_Parts[0] -notin $S_SingleFactorTokens)
                    }
                )

                $Script:S_ReferencedAuthStrengths[$S_AuthStrId] = [PSCustomObject]@{
                    Id                    = $S_A.id
                    DisplayName           = $S_A.displayName
                    Description           = $S_A.description
                    PolicyType            = $S_A.policyType        # builtIn | custom
                    RequirementsSatisfied = $S_A.requirementsSatisfied
                    AllowedCombinations   = $S_Combos
                    MfaCombinations       = $S_MfaCombos
                    IsMfaCapable          = ($S_MfaCombos.Count -gt 0)
                }
            }
            catch
            {
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

        $S_MfaCapableCount    = @($Script:S_ReferencedAuthStrengths.Values | Where-Object { $_.IsMfaCapable -eq $true  }).Count
        $S_NotMfaCapableCount = @($Script:S_ReferencedAuthStrengths.Values | Where-Object { $_.IsMfaCapable -eq $false }).Count
        Write-Host ("Resolved {0} Authentication Strength policy(ies): {1} MFA-capable, {2} not MFA-capable." -f `
                $Script:S_ReferencedAuthStrengths.Count, $S_MfaCapableCount, $S_NotMfaCapableCount) -ForegroundColor Green

        if ($S_NotMfaCapableCount -gt 0)
        {
            Write-Warning "One or more CA policies reference an Authentication Strength that is NOT MFA-capable."
        }
    }
    else
    {
        Write-Host "No CA policies reference an Authentication Strength — skipping." -ForegroundColor DarkGray
    }

    # ── Retrieve Member Users (Guests excluded) ────────────────────────────────
    $S_UserProperties = @(
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

    if ($Test)
    {
        Write-Host "[TEST MODE] Selecting 10 random Member users." -ForegroundColor Yellow
        $S_AllMembers = Get-MgUser -Filter "userType eq 'Member'" -Property $S_UserProperties -All
        $S_Users = @($S_AllMembers | Get-Random -Count ([Math]::Min(10, ($S_AllMembers | Measure-Object).Count)))
    }
    else
    {
        $S_Users = Get-MgUser -Filter "userType eq 'Member'" -Property $S_UserProperties -All
    }

    $S_UserCount = ($S_Users | Measure-Object).Count
    Write-Host "Found $S_UserCount Member users. Processing..." -ForegroundColor Cyan

    $S_Results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $S_CurrentUser = 0

    foreach ($S_User in $S_Users)
    {
        $S_CurrentUser++
        Write-Progress -Activity "Processing Users" -Status "$S_CurrentUser of $S_UserCount - $($S_User.DisplayName)" -PercentComplete (($S_CurrentUser / $S_UserCount) * 100)

        # ── Authentication Methods ─────────────────────────────────────────────
        $S_AuthMethodError = $false
        try
        {
            $S_AuthMethods = Get-MgUserAuthenticationMethod -UserId $S_User.Id -ErrorAction Stop
        }
        catch
        {
            Write-Warning "Could not retrieve auth methods for $($S_User.UserPrincipalName): $($_.Exception.Message)"
            $S_AuthMethodError = $true
            $S_AuthMethods = @()
        }

        # ── Determine MFA Registration ─────────────────────────────────────────
        $S_AuthTypes = $S_AuthMethods | ForEach-Object { $_.AdditionalProperties.'@odata.type' }

        $S_HasModernAuth = $S_AuthTypes | Where-Object {
            $_ -in @(
                '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod'
                '#microsoft.graph.fido2AuthenticationMethod'
                '#microsoft.graph.softwareOathAuthenticationMethod'
            )
        }
        $S_HasLegacyAuth = $S_AuthTypes | Where-Object {
            $_ -in @(
                '#microsoft.graph.phoneAuthenticationMethod'
            )
        }

        if ($S_AuthMethodError)
        {
            $S_MfaRegistered = 'Access Denied (Privileged Account)'
        }
        elseif ($S_HasModernAuth)
        {
            $S_MfaRegistered = 'Modern Auth'
        }
        elseif ($S_HasLegacyAuth)
        {
            $S_MfaRegistered = 'Legacy Auth'
        }
        else
        {
            $S_MfaRegistered = 'No MFA'
        }

        # ── Determine Active Account ───────────────────────────────────────────
        $S_LastInteractive    = $null
        $S_LastNonInteractive = $null
        if (-not $S_User.AccountEnabled)
        {
            $S_ActiveAccount = 'Disabled'
        }
        else
        {
            $S_LastInteractive    = $S_User.SignInActivity.LastSignInDateTime
            $S_LastNonInteractive = $S_User.SignInActivity.LastNonInteractiveSignInDateTime

            $S_Dates = @($S_LastInteractive, $S_LastNonInteractive) | Where-Object { $_ -ne $null }

            if ($S_Dates.Count -eq 0)
            {
                $S_ActiveAccount = 'No Sign-In Recorded'
            }
            else
            {
                $S_LastSignIn = ($S_Dates | Sort-Object -Descending | Select-Object -First 1)
                if ($S_LastSignIn -ge $S_CutoffDate)
                {
                    $S_ActiveAccount = 'Yes'
                }
                else
                {
                    $S_DaysAgo = [math]::Floor(((Get-Date) - $S_LastSignIn).TotalDays)
                    $S_ActiveAccount = "${S_DaysAgo}+ Days Ago"
                }
            }
        }

        # ── Licensing & Sync ───────────────────────────────────────────────────
        $S_IsLicensed = ($S_User.AssignedLicenses | Measure-Object).Count -gt 0
        $S_IsOnPremSynced = $S_User.OnPremisesSyncEnabled -eq $true

        # ── Mail & Domain ──────────────────────────────────────────────────────
        $S_Mail = if ($S_User.Mail)
        {
            $S_User.Mail } else
        {
            'None' }
        $S_Domain = if ($S_User.Mail)
        {
            ($S_User.Mail -split '@')[1]
        } else
        {
            ($S_User.UserPrincipalName -split '@')[1]
        }

        # ── CA Exclusion Tags ──────────────────────────────────────────────────
        # Comma-joined list of Ideal-policy ReportIds (e.g. "001, 003") that
        # exclude this user via ExcludeUsers or transitively via ExcludeGroups.
        # Empty when the user is not on any Ideal exclusion list.
        $S_CaExclusionTags = if ($Script:S_UserExclusionMap.ContainsKey($S_User.Id))
        {
            (@($Script:S_UserExclusionMap[$S_User.Id]) | Sort-Object -Unique) -join ', '
        } else
        {
            '' }

        # ── CA Coverage (Ideal-policy perspective) ─────────────────────────────
        # Full    : >=1 Ideal policy exists AND user is on zero Ideal exclusion lists
        # Partial : user is excluded from at least one Ideal policy (Phase 2 will
        #           split Partial vs None once we evaluate per-policy inclusion)
        # None    : no Ideal CA policy exists in the tenant
        $S_CaCoverage = if (($S_IdealPolicies | Measure-Object).Count -eq 0)
        {
            'None'
        }
        elseif ([string]::IsNullOrWhiteSpace($S_CaExclusionTags))
        {
            'Full'
        }
        else
        {
            'Partial'
        }

        # ── MFA Posture & Risk Score ───────────────────────────────────────────
        # Posture / Risk matrix (locked):
        #   Modern Auth + Full    -> Fully Compliant     (0)
        #   Legacy Auth + Full    -> Weak Factor         (2)
        #   Modern Auth + Partial -> Coverage Gap        (3)
        #   Legacy Auth + Partial -> Weak & Gap          (5)
        #   Modern Auth + None    -> Unenforced          (6)
        #   Legacy Auth + None    -> Weak & Unenforced   (7)
        #   No MFA      + Full    -> At Risk             (8)
        #   No MFA      + Partial -> At Risk             (9)
        #   No MFA      + None    -> Critical            (10)
        #   Unknown     + any     -> Unknown             (blank)
        # v3 change (Phase 1): "No MFA + Partial" is bucketed under At Risk
        # rather than Coverage Gap. The user has no registered MFA method at
        # all, so the threat shape is the same as "No MFA + Full". The risk
        # score is preserved at 9 (vs 8) so sort order still distinguishes
        # the two states, but the surfaced label, colour, and remediation
        # copy treat both with the same urgency. Coverage Gap is now reserved
        # for "Modern Auth + Partial" only — users who already have strong
        # MFA and merely sit in an exclusion list of an Ideal policy.
        # Disabled accounts always score 0 (not a live attack surface);
        # Posture is still computed so cleanup candidates remain visible.
        $S_PostureKey = "$S_MfaRegistered|$S_CaCoverage"
        $S_MfaPosture = switch ($S_PostureKey)
        {
            'Modern Auth|Full'
            {
                'Fully Compliant' }
            'Modern Auth|Partial'
            {
                'Coverage Gap' }
            'Modern Auth|None'
            {
                'Unenforced' }
            'Legacy Auth|Full'
            {
                'Weak Factor' }
            'Legacy Auth|Partial'
            {
                'Weak & Gap' }
            'Legacy Auth|None'
            {
                'Weak & Unenforced' }
            'No MFA|Full'
            {
                'At Risk' }
            'No MFA|Partial'
            {
                'At Risk' }
            'No MFA|None'
            {
                'Critical' }
            default
            {
                'Unknown' }
        }
        $S_RiskScoreRaw = switch ($S_PostureKey)
        {
            'Modern Auth|Full'
            {
                0 }
            'Legacy Auth|Full'
            {
                2 }
            'Modern Auth|Partial'
            {
                3 }
            'Legacy Auth|Partial'
            {
                5 }
            'Modern Auth|None'
            {
                6 }
            'Legacy Auth|None'
            {
                7 }
            'No MFA|Full'
            {
                8 }
            'No MFA|Partial'
            {
                9 }
            'No MFA|None'
            {
                10 }
            default
            {
                $null }
        }
        $S_RiskScore = if ($S_ActiveAccount -eq 'Disabled')
        {
            0 } else
        {
            $S_RiskScoreRaw }

        # ── Build result row ───────────────────────────────────────────────────
        $S_Results.Add([PSCustomObject]@{
                DisplayName       = $S_User.DisplayName
                UserPrincipalName = $S_User.UserPrincipalName
                Mail              = $S_Mail
                Domain            = $S_Domain
                UserType          = $S_User.UserType
                MFA_Registered    = $S_MfaRegistered
                ActiveAccount     = $S_ActiveAccount
                IsLicensed        = $S_IsLicensed
                IsOnPremSynced    = $S_IsOnPremSynced
                CaExclusionTags   = $S_CaExclusionTags
                CaCoverage        = $S_CaCoverage
                MfaPosture        = $S_MfaPosture
                RiskScore         = $S_RiskScore
            })
    }

    Write-Progress -Activity "Processing Users" -Completed

    # ── Export CSV ─────────────────────────────────────────────────────────────
    $S_Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nCSV exported to: $OutputPath" -ForegroundColor Green
    Write-Host "Total rows: $($S_Results.Count)" -ForegroundColor Green

    # ── Calculate Statistics ───────────────────────────────────────────────────
    $S_TotalMembers   = $S_Results.Count
    $S_DisabledCount  = ($S_Results | Where-Object { $_.ActiveAccount -eq 'Disabled' }).Count
    $S_EnabledCount   = $S_TotalMembers - $S_DisabledCount

    # Active / Inactive split within enabled accounts. ActiveAccount values:
    #   'Yes'                 -> Active (signed in within InactiveDays)
    #   'Disabled'            -> Disabled
    #   'No Sign-In Recorded' -> Inactive (never observed)
    #   <date string>         -> Inactive (last sign-in older than threshold)
    $S_ActiveCount        = ($S_Results | Where-Object { $_.ActiveAccount -eq 'Yes' }).Count
    $S_InactiveCount      = $S_EnabledCount - $S_ActiveCount

    $S_EnabledUsers       = $S_Results | Where-Object { $_.ActiveAccount -ne 'Disabled' }
    $S_EnabledModernAuth  = ($S_EnabledUsers | Where-Object { $_.MFA_Registered -eq 'Modern Auth' }).Count
    $S_EnabledLegacyAuth  = ($S_EnabledUsers | Where-Object { $_.MFA_Registered -eq 'Legacy Auth' }).Count
    $S_EnabledNoMFA       = ($S_EnabledUsers | Where-Object { $_.MFA_Registered -eq 'No MFA' }).Count
    $S_EnabledAccessDenied= ($S_EnabledUsers | Where-Object { $_.MFA_Registered -eq 'Access Denied (Privileged Account)' }).Count
    $S_EnabledHasMFA      = $S_EnabledModernAuth + $S_EnabledLegacyAuth

    $S_CoveragePercent = if ($S_EnabledCount -gt 0)
    {
        [math]::Round(($S_EnabledHasMFA / $S_EnabledCount) * 100, 1)
    } else
    {
        0 }

    # Posture distribution across enabled accounts (matches the 9-value taxonomy
    # rendered as pills in the user table). Disabled accounts are excluded so
    # the numbers reflect the live attack surface.
    $S_PostureBuckets = @('Fully Compliant','Weak Factor','Coverage Gap','Unenforced','Weak & Gap','Weak & Unenforced','At Risk','Critical','Unknown')
    $S_PostureCounts  = @{}
    foreach ($S_PB in $S_PostureBuckets)
    {
        $S_PostureCounts[$S_PB] = 0 }
    foreach ($S_EU in @($S_EnabledUsers))
    {
        $S_PKey = if ([string]::IsNullOrWhiteSpace([string]$S_EU.MfaPosture))
        {
            'Unknown' } else
        {
            [string]$S_EU.MfaPosture }
        if ($S_PostureCounts.ContainsKey($S_PKey))
        {
            $S_PostureCounts[$S_PKey]++ } else
        {
            $S_PostureCounts['Unknown']++ }
    }

    # Build table rows for ALL Member users (filterable client-side)
    $S_TableRows = ($S_Results | ForEach-Object {
            $S_TagsCell      = if ([string]::IsNullOrWhiteSpace($_.CaExclusionTags))
            {
                '<span style="color:#999">—</span>' } else
            {
                [System.Web.HttpUtility]::HtmlEncode($_.CaExclusionTags) }
            $S_TagsBucket    = if ([string]::IsNullOrWhiteSpace($_.CaExclusionTags))
            {
                'None' } else
            {
                'Has' }
            $S_ActiveBucket  = if ($_.ActiveAccount -eq 'Disabled')
            {
                'Disabled' } elseif ($_.ActiveAccount -eq 'Yes')
            {
                'Active' } elseif ($_.ActiveAccount -eq 'No Sign-In Recorded')
            {
                'NoSignIn' } else
            {
                'Stale' }
            $S_CoverageClass = switch ($_.CaCoverage)
            {
                'Full'
                {
                    'background:#eaf6ec;color:#107c10;border:1px solid #107c10' }
                'Partial'
                {
                    'background:#fff4ce;color:#8a6d00;border:1px solid #c0a000' }
                default
                {
                    'background:#f3f3f3;color:#666;border:1px solid #999' }
            }
            $S_CoverageCell  = "<span style=""$S_CoverageClass;padding:2px 8px;border-radius:10px;font-size:12px;font-weight:600"">$([System.Web.HttpUtility]::HtmlEncode([string]$_.CaCoverage))</span>"
            $S_PostureClass  = switch ($_.MfaPosture)
            {
                'Fully Compliant'
                {
                    'background:#eaf6ec;color:#107c10;border:1px solid #107c10' }
                'Weak Factor'
                {
                    'background:#fff4ce;color:#8a6d00;border:1px solid #c0a000' }
                'Coverage Gap'
                {
                    'background:#fff4ce;color:#8a6d00;border:1px solid #c0a000' }
                'Unenforced'
                {
                    'background:#ffe8cc;color:#9a4f00;border:1px solid #d83b01' }
                'Weak & Gap'
                {
                    'background:#ffe8cc;color:#9a4f00;border:1px solid #d83b01' }
                'Weak & Unenforced'
                {
                    'background:#ffe8cc;color:#9a4f00;border:1px solid #d83b01' }
                'At Risk'
                {
                    'background:#fdecea;color:#a4262c;border:1px solid #a4262c' }
                'Critical'
                {
                    'background:#5c0a12;color:#fff;border:1px solid #3d0008' }
                default
                {
                    'background:#f3f3f3;color:#666;border:1px solid #999' }
            }
            $S_PostureCell   = "<span style=""$S_PostureClass;padding:2px 8px;border-radius:10px;font-size:12px;font-weight:600"">$([System.Web.HttpUtility]::HtmlEncode([string]$_.MfaPosture))</span>"
            # Risk pill: numeric, colour-banded; '—' when Unknown
            $S_RiskNum       = if ($null -eq $_.RiskScore)
            {
                -1 } else
            {
                [int]$_.RiskScore }
            $S_RiskClass     = if     ($S_RiskNum -lt 0)
            {
                'background:#f3f3f3;color:#666;border:1px solid #999' }
            elseif ($S_RiskNum -le 2)
            {
                'background:#eaf6ec;color:#107c10;border:1px solid #107c10' }
            elseif ($S_RiskNum -le 4)
            {
                'background:#f3f9d8;color:#5a6b14;border:1px solid #7a8c1a' }
            elseif ($S_RiskNum -le 6)
            {
                'background:#fff4ce;color:#8a6d00;border:1px solid #bc8000' }
            elseif ($S_RiskNum -le 8)
            {
                'background:#ffe8cc;color:#9a4f00;border:1px solid #d83b01' }
            else
            {
                'background:#fdecea;color:#a4262c;border:1px solid #a4262c' }
            $S_RiskText      = if ($S_RiskNum -lt 0)
            {
                '—' } else
            {
                "$S_RiskNum" }
            $S_RiskCell      = "<span style=""$S_RiskClass;padding:2px 8px;border-radius:10px;font-size:12px;font-weight:700"">$S_RiskText</span>"
            # Technical identifiers (UPN, Mail, Domain) are wrapped in <code> per the
            # HTML documentation rules — semantic + monospace + light grey background.
            $S_UpnEnc    = [System.Web.HttpUtility]::HtmlEncode([string]$_.UserPrincipalName)
            $S_MailEnc   = [System.Web.HttpUtility]::HtmlEncode([string]$_.Mail)
            $S_DomainEnc = [System.Web.HttpUtility]::HtmlEncode([string]$_.Domain)
            $S_UpnCell    = if ([string]::IsNullOrWhiteSpace([string]$_.UserPrincipalName))
            {
                '<span style="color:#999">—</span>' } else
            {
                "<code>$S_UpnEnc</code>" }
            $S_MailCell   = if ([string]::IsNullOrWhiteSpace([string]$_.Mail))
            {
                '<span style="color:#999">—</span>' } else
            {
                "<code>$S_MailEnc</code>" }
            $S_DomainCell = if ([string]::IsNullOrWhiteSpace([string]$_.Domain))
            {
                '<span style="color:#999">—</span>' } else
            {
                "<code>$S_DomainEnc</code>" }
            "        <tr data-auth=""$($_.MFA_Registered)"" data-active=""$S_ActiveBucket"" data-licensed=""$($_.IsLicensed)"" data-onprem=""$($_.IsOnPremSynced)"" data-exclusions=""$S_TagsBucket"" data-cacoverage=""$($_.CaCoverage)"" data-posture=""$($_.MfaPosture)"" data-risk=""$S_RiskNum""><td>$([System.Web.HttpUtility]::HtmlEncode($_.DisplayName))</td><td>$S_UpnCell</td><td>$S_MailCell</td><td>$S_DomainCell</td><td>$([System.Web.HttpUtility]::HtmlEncode([string]$_.MFA_Registered))</td><td>$S_CoverageCell</td><td>$S_PostureCell</td><td style=""text-align:center"">$S_RiskCell</td><td>$([System.Web.HttpUtility]::HtmlEncode([string]$_.ActiveAccount))</td><td>$($_.IsLicensed)</td><td>$($_.IsOnPremSynced)</td><td>$S_TagsCell</td></tr>"
        }) -join "`n"

    # ── Build CA Policy table rows ──────────────────────────────────────
    function Format-CaList
    {
        param([object[]]$F_Items)
        if ($null -eq $F_Items -or $F_Items.Count -eq 0)
        {
            return '<span style="color:#999">—</span>' }
        ($F_Items | ForEach-Object { [System.Web.HttpUtility]::HtmlEncode([string]$_) }) -join '<br>'
    }

    $S_CaPolicyRows = if ($Script:S_MfaCaPolicies.Count -eq 0)
    {
        '        <tr><td colspan="9" style="text-align:center;color:#999">No MFA-enforcing Conditional Access policies were found.</td></tr>'
    }
    else
    {
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
            if ($_.IncludeUserActions.Count -gt 0)
            {
                $S_Workloads += ($_.IncludeUserActions | ForEach-Object { "action: $_" })
            }
            if ($_.ExcludeApplications.Count -gt 0)
            {
                $S_Workloads += ($_.ExcludeApplications | ForEach-Object { "exclude: $_" })
            }
            # Coverage-tier badge prepended to the Workloads cell
            $S_CoverageColor = switch ($_.WorkloadCoverage)
            {
                'Full'
                {
                    '#107c10' }   # green
                'Data'
                {
                    '#0078d4' }   # blue
                default
                {
                    '#ff8c00' }   # orange (Partial)
            }
            $S_CoverageBadge = "<span style=""display:inline-block;padding:2px 8px;border-radius:10px;background:$S_CoverageColor;color:#fff;font-size:11px;font-weight:600;margin-bottom:4px"">$($_.WorkloadCoverage) coverage</span>"
            $S_WorkloadsHtml = $S_CoverageBadge + '<br>' + (Format-CaList $S_Workloads)

            $S_GrantParts = @()
            if ($_.AllBuiltInControls.Count -gt 0)
            {
                $S_GrantParts += ($_.AllBuiltInControls -join " $($_.GrantOperator) ") }
            if ($_.HasAuthenticationStrength)
            {
                $S_AsName = if ($_.AuthenticationStrengthName)
                {
                    $_.AuthenticationStrengthName } else
                {
                    '<unknown>' }
                $S_GrantParts += "AuthStrength: $S_AsName"
            }
            $S_GrantText = ($S_GrantParts -join '<br>')
            if ($_.MfaIsBypassable)
            {
                $S_GrantText += '<br><span style="color:#d13438;font-weight:600">⚠ MFA bypassable (OR with companion controls)</span>'
            }

            $S_OtherTargeting = @()
            if ($_.IncludePlatforms.Count -gt 0)
            {
                $S_OtherTargeting += "Platforms incl: $($_.IncludePlatforms -join ', ')" }
            if ($_.ExcludePlatforms.Count -gt 0)
            {
                $S_OtherTargeting += "Platforms excl: $($_.ExcludePlatforms -join ', ')" }
            if ($_.IncludeLocations.Count -gt 0)
            {
                $S_OtherTargeting += "Locations incl: $($_.IncludeLocations -join ', ')" }
            if ($_.ExcludeLocations.Count -gt 0)
            {
                $S_OtherTargeting += "Locations excl: $($_.ExcludeLocations -join ', ')" }
            if ($_.ClientAppTypes.Count   -gt 0)
            {
                $S_OtherTargeting += "Client apps: $($_.ClientAppTypes -join ', ')" }
            if ($_.SignInRiskLevels.Count -gt 0)
            {
                $S_OtherTargeting += "Sign-in risk: $($_.SignInRiskLevels -join ', ')" }
            if ($_.UserRiskLevels.Count   -gt 0)
            {
                $S_OtherTargeting += "User risk: $($_.UserRiskLevels -join ', ')" }
            # Conditions-posture badge prepended to the Other Conditions cell.
            # RiskBased = neutral gray (de-emphasised — gap engine ignores these).
            $S_PostureColor = switch ($_.ConditionsPosture)
            {
                'Unrestricted'
                {
                    '#107c10' }   # green
                'Constrained'
                {
                    '#ff8c00' }   # amber
                default
                {
                    '#6b6b6b' }   # neutral gray (RiskBased)
            }
            $S_PostureBadge = "<span style=""display:inline-block;padding:2px 8px;border-radius:10px;background:$S_PostureColor;color:#fff;font-size:11px;font-weight:600;margin-bottom:4px"">$($_.ConditionsPosture)</span>"
            $S_OtherInner   = if ($S_OtherTargeting.Count -gt 0)
            {
                Format-CaList $S_OtherTargeting } else
            {
                '<span style="color:#999">—</span>' }
            $S_OtherText    = $S_PostureBadge + '<br>' + $S_OtherInner

            # Enforcement-tier badge under the policy name. Gray = Ignored by gap engine.
            $S_TierColor = switch ($_.EnforcementTier)
            {
                'Ideal'
                {
                    '#107c10' }   # green
                'Acceptable'
                {
                    '#ff8c00' }   # amber
                default
                {
                    '#6b6b6b' }   # neutral gray (Ignored)
            }
            # Place the tier badge ABOVE the name for consistency with Workloads and Conditions columns.
            $S_TierBadge = "<span style=""display:inline-block;padding:2px 8px;border-radius:10px;background:$S_TierColor;color:#fff;font-size:11px;font-weight:600;margin-bottom:4px"">$($_.EnforcementTier)</span>"
            # Policy DisplayName wrapped in <code> as a technical identifier.
            $S_NameCell  = $S_TierBadge + '<br><code>' + [System.Web.HttpUtility]::HtmlEncode($_.DisplayName) + '</code>'

            # Persona badge — audience shape (used by the gap engine to skip
            # Admin-only policies; Guests already filtered upstream).
            $S_PersonaColor = switch ($_.Persona)
            {
                'AllUsers'
                {
                    '#0078d4' }   # blue
                'Internal'
                {
                    '#107c10' }   # green
                'Admins'
                {
                    '#8764b8' }   # purple
                'Guests'
                {
                    '#ff8c00' }   # orange
                'Targeted'
                {
                    '#6b6b6b' }   # gray
                default
                {
                    '#d13438' }   # red (Mixed — sanity flag)
            }
            $S_PersonaBadge = "<span style=""display:inline-block;padding:2px 8px;border-radius:10px;background:$S_PersonaColor;color:#fff;font-size:11px;font-weight:600"">$($_.Persona)</span>"

            # ReportId badge — monospace pill so the per-user CA Exclusions
            # tags (e.g. "001, 003") can be cross-referenced at a glance.
            $S_IdCell = "<span style=""display:inline-block;padding:2px 8px;border-radius:6px;background:#1f1f1f;color:#fff;font-family:Consolas,monospace;font-size:12px;font-weight:700"">$([System.Web.HttpUtility]::HtmlEncode([string]$_.ReportId))</span>"

            # data-* attributes drive the client-side filter dropdowns above the table.
            "        <tr data-reportid=""$($_.ReportId)"" data-tier=""$($_.EnforcementTier)"" data-coverage=""$($_.WorkloadCoverage)"" data-posture=""$($_.ConditionsPosture)"" data-persona=""$($_.Persona)""><td style=""text-align:center"">$S_IdCell</td><td>$S_NameCell</td>" +
            "<td>$($_.State)</td>" +
            "<td>$S_PersonaBadge</td>" +
            "<td>$S_WorkloadsHtml</td>" +
            "<td>$(Format-CaList $S_Included)</td>" +
            "<td>$(Format-CaList $S_Excluded)</td>" +
            "<td>$S_GrantText</td>" +
            "<td>$S_OtherText</td></tr>"
        }) -join "`n"
    }

    # ── Build Authentication Strength table rows ──────────────────────────────────
    $S_AuthStrengthRows = if ($null -eq $Script:S_ReferencedAuthStrengths -or $Script:S_ReferencedAuthStrengths.Count -eq 0)
    {
        '        <tr><td colspan="6" style="text-align:center;color:#999">No Authentication Strengths referenced by the kept CA policies.</td></tr>'
    }
    else
    {
        ($Script:S_ReferencedAuthStrengths.Values | ForEach-Object {
            $S_MfaText = switch ($_.IsMfaCapable)
            {
                $true
                {
                    '<span style="color:#107c10;font-weight:600">Yes</span>' }
                $false
                {
                    '<span style="color:#d13438;font-weight:600">No</span>' }
                default
                {
                    '<span style="color:#999">Unknown</span>' }
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
    $S_ReportDate = Get-Date -Format 'dd MMM yyyy HH:mm'
    $S_TenantId = (Get-MgContext).TenantId
    # Tenant display name + primary (initial *.onmicrosoft.com) domain via the
    # Organization endpoint. Fail-soft so a transient Graph hiccup doesn't kill
    # the report.
    $S_TenantName        = $null
    $S_TenantPrimaryDom  = $null
    try
    {
        $S_OrgResp = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=displayName,verifiedDomains'
        $S_Org     = @($S_OrgResp.value)[0]
        if ($S_Org)
        {
            $S_TenantName       = [string]$S_Org.displayName
            $S_TenantPrimaryDom = ($S_Org.verifiedDomains | Where-Object { $_.isInitial }) | Select-Object -First 1 -ExpandProperty name
        }
    }
    catch
    {
        Write-Warning "Could not retrieve tenant display name: $_"
    }
    $S_TenantLabel = if ($S_TenantName -and $S_TenantPrimaryDom)
    {
        "$S_TenantName ($S_TenantPrimaryDom &middot; $S_TenantId)" }
    elseif ($S_TenantName)
    {
        "$S_TenantName ($S_TenantId)" }
    else
    {
        "$S_TenantId" }

    if ($null -eq $Script:S_SecurityDefaultsEnabled)
    {
        $S_SecDefaultsText  = 'Security Defaults: Unknown'
        $S_SecDefaultsClass = 'unknown'
    }
    elseif ($Script:S_SecurityDefaultsEnabled)
    {
        $S_SecDefaultsText  = 'Security Defaults: ENABLED'
        $S_SecDefaultsClass = 'enabled'
    }
    else
    {
        $S_SecDefaultsText  = 'Security Defaults: DISABLED'
        $S_SecDefaultsClass = 'disabled'
    }

    $S_Html = @"
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
        .disclaimer { max-width: 880px; margin: 14px auto 0; padding: 10px 16px; background: #fff8e1; border: 1px solid #e0b400; border-left: 4px solid #bc8000; border-radius: 6px; color: #5a4500; font-size: 12.5px; line-height: 1.5; text-align: left; }
        .disclaimer strong { color: #8a6d00; }
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
        /* Posture cards — colour-banded to match the table pills. */
        .card.posture-fully       { border-left-color: #107c10; }
        .card.posture-fully .value{ color: #107c10; }
        .card.posture-weak        { border-left-color: #7a8c1a; }
        .card.posture-weak .value { color: #5a6b14; }
        .card.posture-gap         { border-left-color: #bc8000; }
        .card.posture-gap .value  { color: #8a6d00; }
        .card.posture-unenforced  { border-left-color: #d83b01; }
        .card.posture-unenforced .value { color: #9a4f00; }
        .card.posture-weakgap     { border-left-color: #d83b01; }
        .card.posture-weakgap .value    { color: #9a4f00; }
        .card.posture-weakunenf   { border-left-color: #d83b01; }
        .card.posture-weakunenf .value  { color: #9a4f00; }
        .card.posture-atrisk      { border-left-color: #a4262c; }
        .card.posture-atrisk .value     { color: #a4262c; }
        .card.posture-critical    { border-left-color: #5c0a12; background: #fdecea; }
        .card.posture-critical .value   { color: #5c0a12; }
        .card.posture-unknown     { border-left-color: #999; }
        .card.posture-unknown .value    { color: #666; }
        .section { margin-bottom: 24px; }
        .section h2 { font-size: 18px; color: #1a1a2e; margin-bottom: 12px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0; text-align: center; }
        .breakdown { display: flex; flex-wrap: wrap; gap: 16px; justify-content: center; }
        .breakdown .card { min-width: 180px; max-width: 240px; padding: 18px 22px; }
        .breakdown .card .value { font-size: 28px; }
        table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.08); margin-top: 12px; border: 1px solid #ccc; }
        thead { border-bottom: 2px solid #a52428; }
        th { background: #d13438; color: #fff; text-align: left; padding: 10px 14px; font-size: 13px; text-transform: uppercase; letter-spacing: 0.3px; }
        td { padding: 9px 14px; font-size: 13px; border-bottom: 1px solid #eee; vertical-align: top; }
        tbody tr:nth-child(even) { background: #fafafa; }
        tbody tr:last-child td { border-bottom: 1px solid #ccc; }
        tr:hover { background: #fdf2f2; }
        /* Inline code styling for technical identifiers (UPN, domain, policy/group names). */
        code { font-family: Consolas, 'Cascadia Mono', 'Courier New', monospace; font-size: 0.92em; background: #f3f3f3; color: #1a1a2e; padding: 1px 5px; border: 1px solid #e0e0e0; border-radius: 3px; word-break: break-all; }
        /* Table of Contents under the disclaimer. */
        .toc { max-width: 880px; margin: 14px auto 0; padding: 10px 16px; background: #fafafa; border: 1px solid #e0e0e0; border-radius: 6px; font-size: 13px; }
        .toc strong { display: block; margin-bottom: 4px; color: #1a1a2e; }
        .toc ol { list-style: decimal inside; margin: 0; padding: 0; columns: 2; column-gap: 24px; }
        .toc ol li { padding: 2px 0; }
        .toc a { color: #0078d4; text-decoration: none; }
        .toc a:hover { text-decoration: underline; }
        .footer { text-align: center; font-size: 12px; color: #999; margin-top: 32px; }
        .secdefaults { text-align: center; font-size: 14px; font-weight: 700; padding: 10px 16px; border-radius: 6px; margin: 0 auto 20px; max-width: 480px; }
        .secdefaults.enabled  { background: #fdecea; color: #d13438; border: 1px solid #d13438; }
        .secdefaults.disabled { background: #eaf6ec; color: #107c10; border: 1px solid #107c10; }
        .secdefaults.unknown  { background: #f3f3f3; color: #666;    border: 1px solid #999; }
        .ca-filters, .user-filters { display: flex; flex-wrap: wrap; gap: 12px; justify-content: center; align-items: center; margin: 8px 0 4px; font-size: 13px; }
        .ca-filters label, .user-filters label { font-weight: 600; color: #555; margin-right: 4px; }
        .ca-filters select, .user-filters select { padding: 4px 8px; border: 1px solid #ccc; border-radius: 6px; background: #fff; font-size: 13px; cursor: pointer; }
        .ca-filters button, .user-filters button { padding: 4px 12px; border: 1px solid #d13438; background: #fff; color: #d13438; border-radius: 6px; font-size: 12px; font-weight: 600; cursor: pointer; }
        .ca-filters button:hover, .user-filters button:hover { background: #fdecea; }

        /* Snapshot bar — fixed top-right "Save filtered HTML" button. */
        .snapshot-bar { position: fixed; top: 12px; right: 12px; z-index: 9999; display: flex; gap: 8px; }
        .snapshot-bar button { padding: 8px 14px; background: #0078d4; color: #fff; border: 1px solid #005a9e; border-radius: 6px; font-size: 13px; font-weight: 600; cursor: pointer; box-shadow: 0 2px 6px rgba(0,0,0,0.18); }
        .snapshot-bar button:hover { background: #106ebe; }

        /* ── Print / PDF layout ─────────────────────────────────────────
           Per design: filters STAY visible in print so a filtered subset
           saved to PDF still shows the active filter selections as evidence.
        */
        @page { size: A4; margin: 14mm 12mm; }
        @media print {
            body { background: #fff; padding: 0; color: #000; }
            .header h1 { font-size: 22px; }
            .header .subtitle { font-size: 11px; }
            .disclaimer, .toc { box-shadow: none; }
            .card, table { box-shadow: none !important; }
            tr:hover { background: transparent !important; }
            /* Repeat table headers on every continuation page. */
            thead { display: table-header-group; }
            tfoot { display: table-footer-group; }
            /* Keep small blocks intact across page breaks. */
            .disclaimer, .toc, .secdefaults, .card, .breakdown .card { page-break-inside: avoid; break-inside: avoid; }
            /* Don't orphan a section heading at the bottom of a page. */
            .section h2 { page-break-after: avoid; break-after: avoid-page; }
            /* Keep filter controls visible; just freeze their colours for print. */
            .ca-filters select, .user-filters select,
            .ca-filters input, .user-filters input,
            .ca-filters button, .user-filters button {
                background: #fff !important;
                -webkit-print-color-adjust: exact; print-color-adjust: exact;
            }
            /* Force coloured pills / badges to render their background ink. */
            .card, .secdefaults, th, code, span[style*="background"] {
                -webkit-print-color-adjust: exact; print-color-adjust: exact;
            }
            /* Hide the snapshot button when printing. */
            .snapshot-bar { display: none !important; }
        }
    </style>
    <script>
        function caApplyFilters() {
            var reportId = document.getElementById('caFilterReportId').value.trim().toLowerCase();
            var tier     = document.getElementById('caFilterTier').value;
            var coverage = document.getElementById('caFilterCoverage').value;
            var posture  = document.getElementById('caFilterPosture').value;
            var persona  = document.getElementById('caFilterPersona').value;
            var rows = document.querySelectorAll('#caPolicyTable tbody tr');
            var visible = 0;
            rows.forEach(function (r) {
                if (!r.hasAttribute('data-tier')) { return; } // skip empty-state placeholder
                var rid = (r.getAttribute('data-reportid') || '').toLowerCase();
                var ok = (reportId === '' || rid.indexOf(reportId) !== -1)
                      && (tier === '' || r.getAttribute('data-tier') === tier)
                      && (coverage === '' || r.getAttribute('data-coverage') === coverage)
                      && (posture === '' || r.getAttribute('data-posture') === posture)
                      && (persona === '' || r.getAttribute('data-persona') === persona);
                r.style.display = ok ? '' : 'none';
                if (ok) { visible++; }
            });
            var count = document.getElementById('caFilterCount');
            if (count) { count.textContent = visible + ' policy(ies) shown'; }
        }
        function caResetFilters() {
            document.getElementById('caFilterReportId').value = '';
            document.getElementById('caFilterTier').value = '';
            document.getElementById('caFilterCoverage').value = '';
            document.getElementById('caFilterPosture').value = '';
            document.getElementById('caFilterPersona').value = '';
            caApplyFilters();
        }
        document.addEventListener('DOMContentLoaded', caApplyFilters);

        /* ── Save filtered HTML ───────────────────────────────────────────────────────────
           Bakes the current filter state (select values, input values, hidden
           rows) into the DOM via real attributes, then offers the resulting
           document as a downloadable .html file. The saved snapshot reopens
           with exactly the same filter state and the same visible rows.
        */
        function saveFilteredHtml() {
            try {
                // Re-apply filters first so row display styles are current.
                if (typeof caApplyFilters === 'function')   { caApplyFilters(); }
                if (typeof userApplyFilters === 'function') { userApplyFilters(); }

                var clone = document.documentElement.cloneNode(true);

                // Persist <select> selections.
                clone.querySelectorAll('select').forEach(function (origSel) {
                    var liveSel = document.getElementById(origSel.id);
                    if (!liveSel) { return; }
                    Array.prototype.forEach.call(origSel.options, function (opt) {
                        if (opt.value === liveSel.value) { opt.setAttribute('selected', 'selected'); }
                        else                              { opt.removeAttribute('selected'); }
                    });
                });
                // Persist <input> values.
                clone.querySelectorAll('input').forEach(function (origInp) {
                    var liveInp = document.getElementById(origInp.id);
                    if (!liveInp) { return; }
                    origInp.setAttribute('value', liveInp.value);
                });
                // Persist row visibility (display:none for filtered-out rows).
                ['#userTable tbody tr', '#caPolicyTable tbody tr'].forEach(function (sel) {
                    var liveRows = document.querySelectorAll(sel);
                    var cloneRows = clone.querySelectorAll(sel);
                    liveRows.forEach(function (lr, i) {
                        if (!cloneRows[i]) { return; }
                        if (lr.style.display === 'none') {
                            cloneRows[i].style.display = 'none';
                        } else {
                            cloneRows[i].style.removeProperty('display');
                        }
                    });
                });
                // Stamp the snapshot moment into the subtitle.
                var stamp = clone.querySelector('.header .subtitle');
                if (stamp) {
                    var now = new Date();
                    var pad = function (n) { return (n < 10 ? '0' : '') + n; };
                    var ts = now.getFullYear() + '-' + pad(now.getMonth() + 1) + '-' + pad(now.getDate()) + ' ' + pad(now.getHours()) + ':' + pad(now.getMinutes());
                    stamp.innerHTML += ' &middot; <em>Filtered snapshot saved ' + ts + '</em>';
                }

                var html = '<!DOCTYPE html>\n' + clone.outerHTML;
                var blob = new Blob([html], { type: 'text/html;charset=utf-8' });
                var url  = URL.createObjectURL(blob);
                var a    = document.createElement('a');
                var ts2  = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 16);
                a.href     = url;
                a.download = 'MemberMFAReport-filtered-' + ts2 + '.html';
                document.body.appendChild(a);
                a.click();
                setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(url); }, 100);
            } catch (e) {
                alert('Snapshot failed: ' + e.message);
            }
        }
    </script>
</head>
<body>
    <div class="snapshot-bar">
        <button type="button" onclick="saveFilteredHtml()" title="Save the current view (with filters applied) as a standalone HTML file">⬇ Save filtered HTML</button>
    </div>
    <div class="header">
        <h1>Member MFA Coverage Report</h1>
        <div class="subtitle">Generated: $S_ReportDate | Tenant: $S_TenantLabel | Inactive threshold: $InactiveDays days | Guests excluded</div>
        <div class="disclaimer">
            <strong>MFA Reporting Note</strong><br><br>
            Reporting on Multi-Factor Authentication (MFA) should be treated as an indicative assessment rather than definitive audit evidence. MFA coverage can be affected by several configuration areas, including registered authentication methods, legacy per-user MFA settings, Conditional Access policy scope, authentication strengths, exclusions, break-glass accounts, guest access, and other exception paths.
            <br><br>
            This report is intended to help identify potential gaps, highlight areas requiring review, and support prioritisation of follow-up actions. It should not be relied upon as a complete or audit-grade confirmation that MFA is enforced for every user and access scenario.
        </div>
        <div class="toc" aria-label="Table of contents">
            <strong>Contents</strong>
            <ol>
                <li><a href="#sec-1">Accounts</a></li>
                <li><a href="#sec-2">Auth Methods &mdash; Enabled Accounts</a></li>
                <li><a href="#sec-3">MFA Posture &mdash; Enabled Accounts</a></li>
                <li><a href="#sec-4">Member Users</a></li>
                <li><a href="#sec-5">MFA-Enforcing Conditional Access Policies</a></li>
            </ol>
        </div>
    </div>

    <div class="secdefaults $S_SecDefaultsClass">$S_SecDefaultsText</div>

    <div class="section">
        <h2 id="sec-1">1. Accounts</h2>
        <div class="breakdown">
            <div class="card blue">
                <div class="label">Total Members</div>
                <div class="value">$S_TotalMembers</div>
                <div class="detail">Guests excluded</div>
            </div>
            <div class="card green">
                <div class="label">Active</div>
                <div class="value">$S_ActiveCount</div>
                <div class="detail">$([math]::Round(($S_ActiveCount / [math]::Max($S_TotalMembers,1)) * 100, 1))% of total &middot; signed in &lt;${InactiveDays}d</div>
            </div>
            <div class="card orange">
                <div class="label">Inactive</div>
                <div class="value">$S_InactiveCount</div>
                <div class="detail">$([math]::Round(($S_InactiveCount / [math]::Max($S_TotalMembers,1)) * 100, 1))% of total &middot; stale or no sign-in</div>
            </div>
            <div class="card red">
                <div class="label">Disabled</div>
                <div class="value">$S_DisabledCount</div>
                <div class="detail">$([math]::Round(($S_DisabledCount / [math]::Max($S_TotalMembers,1)) * 100, 1))% of total</div>
            </div>
        </div>
    </div>

    <div class="section">
        <h2 id="sec-2">2. Auth Methods &mdash; Enabled Accounts</h2>
        <div class="breakdown">
            <div class="card green">
                <div class="label">Modern Auth</div>
                <div class="value">$S_EnabledModernAuth</div>
                <div class="detail">$([math]::Round(($S_EnabledModernAuth / [math]::Max($S_EnabledCount,1)) * 100, 1))% of enabled &middot; Authenticator / Passkey / TOTP</div>
            </div>
            <div class="card orange">
                <div class="label">Legacy Auth</div>
                <div class="value">$S_EnabledLegacyAuth</div>
                <div class="detail">$([math]::Round(($S_EnabledLegacyAuth / [math]::Max($S_EnabledCount,1)) * 100, 1))% of enabled &middot; SMS / Voice</div>
            </div>
            <div class="card red">
                <div class="label">No MFA</div>
                <div class="value">$S_EnabledNoMFA</div>
                <div class="detail">$([math]::Round(($S_EnabledNoMFA / [math]::Max($S_EnabledCount,1)) * 100, 1))% of enabled &middot; no method registered</div>
            </div>
            <div class="card purple">
                <div class="label">Access Denied</div>
                <div class="value">$S_EnabledAccessDenied</div>
                <div class="detail">Privileged accounts &middot; methods not readable</div>
            </div>
        </div>
    </div>

    <div class="section">
        <h2 id="sec-3">3. MFA Posture &mdash; Enabled Accounts</h2>
        <div class="breakdown">
            <div class="card posture-fully">
                <div class="label">Fully Compliant</div>
                <div class="value">$($S_PostureCounts['Fully Compliant'])</div>
                <div class="detail">Modern auth + Full CA coverage</div>
            </div>
            <div class="card posture-weak">
                <div class="label">Weak Factor</div>
                <div class="value">$($S_PostureCounts['Weak Factor'])</div>
                <div class="detail">Legacy auth + Full CA coverage</div>
            </div>
            <div class="card posture-gap">
                <div class="label">Coverage Gap</div>
                <div class="value">$($S_PostureCounts['Coverage Gap'])</div>
                <div class="detail">Modern auth + Partial CA coverage</div>
            </div>
            <div class="card posture-weakgap">
                <div class="label">Weak &amp; Gap</div>
                <div class="value">$($S_PostureCounts['Weak & Gap'])</div>
                <div class="detail">Legacy auth + Partial CA coverage</div>
            </div>
            <div class="card posture-unenforced">
                <div class="label">Unenforced</div>
                <div class="value">$($S_PostureCounts['Unenforced'])</div>
                <div class="detail">Modern auth + No CA coverage</div>
            </div>
            <div class="card posture-weakunenf">
                <div class="label">Weak &amp; Unenforced</div>
                <div class="value">$($S_PostureCounts['Weak & Unenforced'])</div>
                <div class="detail">Legacy auth + No CA coverage</div>
            </div>
            <div class="card posture-atrisk">
                <div class="label">At Risk</div>
                <div class="value">$($S_PostureCounts['At Risk'])</div>
                <div class="detail">No MFA + Full CA coverage</div>
            </div>
            <div class="card posture-critical">
                <div class="label">Critical</div>
                <div class="value">$($S_PostureCounts['Critical'])</div>
                <div class="detail">No MFA + No CA coverage</div>
            </div>
            <div class="card posture-unknown">
                <div class="label">Unknown</div>
                <div class="value">$($S_PostureCounts['Unknown'])</div>
                <div class="detail">Posture not determined</div>
            </div>
        </div>
    </div>

    <div class="section">
        <h2 id="sec-4">4. Member Users</h2>
        <div class="user-filters">
            <label for="userFltAuth">Auth Method:</label>
            <select id="userFltAuth" onchange="userApplyFilters()">
                <option value="">All</option>
                <option value="Modern Auth">Modern Auth</option>
                <option value="Legacy Auth">Legacy Auth</option>
                <option value="No MFA">No MFA</option>
                <option value="Access Denied (Privileged Account)">Access Denied</option>
            </select>
            <label for="userFltCoverage">CA Coverage:</label>
            <select id="userFltCoverage" onchange="userApplyFilters()">
                <option value="">All</option>
                <option value="Full">Full</option>
                <option value="Partial">Partial</option>
                <option value="None">None</option>
            </select>
            <label for="userFltPosture">MFA Posture:</label>
            <select id="userFltPosture" onchange="userApplyFilters()">
                <option value="">All</option>
                <option value="Fully Compliant">Fully Compliant</option>
                <option value="Weak Factor">Weak Factor</option>
                <option value="Coverage Gap">Coverage Gap</option>
                <option value="Unenforced">Unenforced</option>
                <option value="Weak &amp; Gap">Weak &amp; Gap</option>
                <option value="Weak &amp; Unenforced">Weak &amp; Unenforced</option>
                <option value="At Risk">At Risk</option>
                <option value="Critical">Critical</option>
                <option value="Unknown">Unknown</option>
            </select>
            <label for="userFltRisk">Min Risk:</label>
            <select id="userFltRisk" onchange="userApplyFilters()">
                <option value="">Any</option>
                <option value="3">&ge; 3</option>
                <option value="5">&ge; 5</option>
                <option value="7">&ge; 7</option>
                <option value="9">&ge; 9</option>
            </select>
            <label for="userFltActive">Active:</label>
            <select id="userFltActive" onchange="userApplyFilters()">
                <option value="">All</option>
                <option value="Active">Active (signed in &lt;${InactiveDays}d)</option>
                <option value="Stale">Stale</option>
                <option value="NoSignIn">No Sign-In Recorded</option>
                <option value="Disabled">Disabled</option>
            </select>
            <label for="userFltLicensed">Licensed:</label>
            <select id="userFltLicensed" onchange="userApplyFilters()">
                <option value="">All</option>
                <option value="True">True</option>
                <option value="False">False</option>
            </select>
            <label for="userFltOnPrem">On-Prem Synced:</label>
            <select id="userFltOnPrem" onchange="userApplyFilters()">
                <option value="">All</option>
                <option value="True">True</option>
                <option value="False">False</option>
            </select>
            <label for="userFltExclusions">CA Exclusions:</label>
            <select id="userFltExclusions" onchange="userApplyFilters()">
                <option value="">All</option>
                <option value="Has">Has tags</option>
                <option value="None">None</option>
            </select>
            <button type="button" onclick="userResetFilters()">Reset</button>
            <span id="userFilterCount" style="color:#666;font-size:12px;"></span>
        </div>
        <script>
        var userSortDir = 1; // 1 asc, -1 desc
        function userApplyFilters() {
            var fau = document.getElementById('userFltAuth').value;
            var fc  = document.getElementById('userFltCoverage').value;
            var fp  = document.getElementById('userFltPosture').value;
            var frv = document.getElementById('userFltRisk').value;
            var fr  = frv === '' ? null : parseInt(frv, 10);
            var fa  = document.getElementById('userFltActive').value;
            var fl  = document.getElementById('userFltLicensed').value;
            var fo  = document.getElementById('userFltOnPrem').value;
            var fx  = document.getElementById('userFltExclusions').value;
            var rows = document.querySelectorAll('#userTable tbody tr');
            var shown = 0;
            rows.forEach(function(r){
                var risk = parseInt(r.dataset.risk, 10);
                var ok = (!fau || r.dataset.auth === fau)
                      && (!fc || r.dataset.cacoverage === fc)
                      && (!fp || r.dataset.posture === fp)
                      && (fr === null || (risk >= 0 && risk >= fr))
                      && (!fa || r.dataset.active === fa)
                      && (!fl || r.dataset.licensed === fl)
                      && (!fo || r.dataset.onprem === fo)
                      && (!fx || r.dataset.exclusions === fx);
                r.style.display = ok ? '' : 'none';
                if (ok) shown++;
            });
            document.getElementById('userFilterCount').textContent = shown + ' user(s) shown';
        }
        function userResetFilters() {
            document.getElementById('userFltAuth').value = '';
            document.getElementById('userFltCoverage').value = '';
            document.getElementById('userFltPosture').value = '';
            document.getElementById('userFltRisk').value = '';
            document.getElementById('userFltActive').value = '';
            document.getElementById('userFltLicensed').value = '';
            document.getElementById('userFltOnPrem').value = '';
            document.getElementById('userFltExclusions').value = '';
            userApplyFilters();
        }
        function userSortByRisk() {
            var tbody = document.querySelector('#userTable tbody');
            var rows = Array.prototype.slice.call(tbody.querySelectorAll('tr'));
            userSortDir = -userSortDir;
            rows.sort(function(a, b){
                var ra = parseInt(a.dataset.risk, 10); if (isNaN(ra) || ra < 0) ra = -1;
                var rb = parseInt(b.dataset.risk, 10); if (isNaN(rb) || rb < 0) rb = -1;
                return (ra - rb) * userSortDir;
            });
            rows.forEach(function(r){ tbody.appendChild(r); });
        }
        document.addEventListener('DOMContentLoaded', userApplyFilters);
        </script>
        <table id="userTable">
            <thead>
                <tr><th>Display Name</th><th>UPN</th><th>Mail</th><th>Domain</th><th>Auth Method</th><th>CA Coverage</th><th>MFA Posture</th><th style="cursor:pointer" onclick="userSortByRisk()" title="Click to sort">Risk &#x21C5;</th><th>Active</th><th>Licensed</th><th>On-Prem Synced</th><th>CA Exclusions</th></tr>
            </thead>
            <tbody>
$S_TableRows
            </tbody>
        </table>
    </div>

    <div class="section">
        <h2 id="sec-5">5. MFA-Enforcing Conditional Access Policies</h2>
        <div class="ca-filters">
            <label for="caFilterReportId">ID:</label>
            <input type="text" id="caFilterReportId" oninput="caApplyFilters()" placeholder="e.g. 001" style="padding:4px 8px;border:1px solid #ccc;border-radius:6px;font-size:13px;width:80px;font-family:Consolas,monospace">
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
            <label for="caFilterPersona">Persona:</label>
            <select id="caFilterPersona" onchange="caApplyFilters()">
                <option value="">All</option>
                <option value="AllUsers">AllUsers</option>
                <option value="Internal">Internal</option>
                <option value="Admins">Admins</option>
                <option value="Guests">Guests</option>
                <option value="Targeted">Targeted</option>
                <option value="Mixed">Mixed</option>
            </select>
            <button type="button" onclick="caResetFilters()">Reset</button>
            <span id="caFilterCount" style="color:#666;font-size:12px;"></span>
        </div>
        <table id="caPolicyTable">
            <thead>
                <tr><th>ID</th><th>Name</th><th>State</th><th>Persona</th><th>Workloads</th><th>Included</th><th>Excluded</th><th>Grants</th><th>Other Conditions</th></tr>
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

    $S_Html | Out-File -FilePath $S_HtmlPath -Encoding UTF8
    Write-Host "HTML report exported to: $S_HtmlPath" -ForegroundColor Green

    # ── Summary HTML (header + disclaimer + 3 card sections only; no tables) ──
    # Reuses the same <style> block so the look is identical. The snapshot bar,
    # TOC, table sections, and per-table scripts are omitted because there is
    # nothing to filter / save in the summary view.
    $S_SummaryPath = [System.IO.Path]::Combine(
        [System.IO.Path]::GetDirectoryName($S_HtmlPath),
        [System.IO.Path]::GetFileNameWithoutExtension($S_HtmlPath) + '-Summary.html'
    )
    $S_SummaryHtml = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Member MFA Coverage Report &mdash; Summary</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; padding: 24px; }
        .header { text-align: center; margin-bottom: 32px; }
        .header h1 { font-size: 28px; color: #1a1a2e; margin-bottom: 4px; }
        .header .subtitle { font-size: 14px; color: #666; }
        .disclaimer { max-width: 880px; margin: 14px auto 0; padding: 10px 16px; background: #fff8e1; border: 1px solid #e0b400; border-left: 4px solid #bc8000; border-radius: 6px; color: #5a4500; font-size: 12.5px; line-height: 1.5; text-align: left; }
        .disclaimer strong { color: #8a6d00; }
        .card { background: #fff; border-radius: 12px; padding: 24px 28px; min-width: 220px; flex: 1; max-width: 280px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); border-left: 5px solid #0078d4; position: relative; }
        .card .label { font-size: 13px; color: #666; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px; }
        .card .value { font-size: 36px; font-weight: 700; color: #1a1a2e; }
        .card .detail { font-size: 12px; color: #888; margin-top: 6px; }
        .card.blue    { border-left-color: #0078d4; }
        .card.red     { border-left-color: #d13438; }
        .card.green   { border-left-color: #107c10; }
        .card.orange  { border-left-color: #ff8c00; }
        .card.purple  { border-left-color: #8764b8; }
        .card.posture-fully       { border-left-color: #107c10; }
        .card.posture-fully .value{ color: #107c10; }
        .card.posture-weak        { border-left-color: #7a8c1a; }
        .card.posture-weak .value { color: #5a6b14; }
        .card.posture-gap         { border-left-color: #bc8000; }
        .card.posture-gap .value  { color: #8a6d00; }
        .card.posture-unenforced  { border-left-color: #d83b01; }
        .card.posture-unenforced .value { color: #9a4f00; }
        .card.posture-weakgap     { border-left-color: #d83b01; }
        .card.posture-weakgap .value    { color: #9a4f00; }
        .card.posture-weakunenf   { border-left-color: #d83b01; }
        .card.posture-weakunenf .value  { color: #9a4f00; }
        .card.posture-atrisk      { border-left-color: #a4262c; }
        .card.posture-atrisk .value     { color: #a4262c; }
        .card.posture-critical    { border-left-color: #5c0a12; background: #fdecea; }
        .card.posture-critical .value   { color: #5c0a12; }
        .card.posture-unknown     { border-left-color: #999; }
        .card.posture-unknown .value    { color: #666; }
        .section { margin-bottom: 24px; }
        .section h2 { font-size: 18px; color: #1a1a2e; margin-bottom: 12px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0; text-align: center; }
        .breakdown { display: flex; flex-wrap: wrap; gap: 16px; justify-content: center; }
        .breakdown .card { min-width: 180px; max-width: 240px; padding: 18px 22px; }
        .breakdown .card .value { font-size: 28px; }
        .secdefaults { text-align: center; font-size: 14px; font-weight: 700; padding: 10px 16px; border-radius: 6px; margin: 0 auto 20px; max-width: 480px; }
        .secdefaults.enabled  { background: #fdecea; color: #d13438; border: 1px solid #d13438; }
        .secdefaults.disabled { background: #eaf6ec; color: #107c10; border: 1px solid #107c10; }
        .secdefaults.unknown  { background: #f3f3f3; color: #666;    border: 1px solid #999; }
        .footer { text-align: center; font-size: 12px; color: #999; margin-top: 32px; }
        @page { size: A4; margin: 14mm 12mm; }
        @media print {
            body { background: #fff; padding: 0; color: #000; }
            .card { box-shadow: none !important; }
            .disclaimer, .secdefaults, .card, .breakdown .card { page-break-inside: avoid; break-inside: avoid; }
            .section h2 { page-break-after: avoid; break-after: avoid-page; }
            .card, .secdefaults { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>Member MFA Coverage Report &mdash; Summary</h1>
        <div class="subtitle">Generated: $S_ReportDate | Tenant: $S_TenantLabel | Inactive threshold: $InactiveDays days | Guests excluded</div>
        <div class="disclaimer">
            <strong>MFA Reporting Note</strong><br><br>
            Reporting on Multi-Factor Authentication (MFA) should be treated as an indicative assessment rather than definitive audit evidence. MFA coverage can be affected by several configuration areas, including registered authentication methods, legacy per-user MFA settings, Conditional Access policy scope, authentication strengths, exclusions, break-glass accounts, guest access, and other exception paths.
            <br><br>
            This report is intended to help identify potential gaps, highlight areas requiring review, and support prioritisation of follow-up actions. It should not be relied upon as a complete or audit-grade confirmation that MFA is enforced for every user and access scenario.
        </div>
    </div>

    <div class="secdefaults $S_SecDefaultsClass">$S_SecDefaultsText</div>

    <div class="section">
        <h2>Accounts</h2>
        <div class="breakdown">
            <div class="card blue">
                <div class="label">Total Members</div>
                <div class="value">$S_TotalMembers</div>
                <div class="detail">Guests excluded</div>
            </div>
            <div class="card green">
                <div class="label">Active</div>
                <div class="value">$S_ActiveCount</div>
                <div class="detail">$([math]::Round(($S_ActiveCount / [math]::Max($S_TotalMembers,1)) * 100, 1))% of total &middot; signed in &lt;${InactiveDays}d</div>
            </div>
            <div class="card orange">
                <div class="label">Inactive</div>
                <div class="value">$S_InactiveCount</div>
                <div class="detail">$([math]::Round(($S_InactiveCount / [math]::Max($S_TotalMembers,1)) * 100, 1))% of total &middot; stale or no sign-in</div>
            </div>
            <div class="card red">
                <div class="label">Disabled</div>
                <div class="value">$S_DisabledCount</div>
                <div class="detail">$([math]::Round(($S_DisabledCount / [math]::Max($S_TotalMembers,1)) * 100, 1))% of total</div>
            </div>
        </div>
    </div>

    <div class="section">
        <h2>Auth Methods &mdash; Enabled Accounts</h2>
        <div class="breakdown">
            <div class="card green">
                <div class="label">Modern Auth</div>
                <div class="value">$S_EnabledModernAuth</div>
                <div class="detail">$([math]::Round(($S_EnabledModernAuth / [math]::Max($S_EnabledCount,1)) * 100, 1))% of enabled &middot; Authenticator / Passkey / TOTP</div>
            </div>
            <div class="card orange">
                <div class="label">Legacy Auth</div>
                <div class="value">$S_EnabledLegacyAuth</div>
                <div class="detail">$([math]::Round(($S_EnabledLegacyAuth / [math]::Max($S_EnabledCount,1)) * 100, 1))% of enabled &middot; SMS / Voice</div>
            </div>
            <div class="card red">
                <div class="label">No MFA</div>
                <div class="value">$S_EnabledNoMFA</div>
                <div class="detail">$([math]::Round(($S_EnabledNoMFA / [math]::Max($S_EnabledCount,1)) * 100, 1))% of enabled &middot; no method registered</div>
            </div>
            <div class="card purple">
                <div class="label">Access Denied</div>
                <div class="value">$S_EnabledAccessDenied</div>
                <div class="detail">Privileged accounts &middot; methods not readable</div>
            </div>
        </div>
    </div>

    <div class="section">
        <h2>MFA Posture &mdash; Enabled Accounts</h2>
        <div class="breakdown">
            <div class="card posture-fully">
                <div class="label">Fully Compliant</div>
                <div class="value">$($S_PostureCounts['Fully Compliant'])</div>
                <div class="detail">Modern auth + Full CA coverage</div>
            </div>
            <div class="card posture-weak">
                <div class="label">Weak Factor</div>
                <div class="value">$($S_PostureCounts['Weak Factor'])</div>
                <div class="detail">Legacy auth + Full CA coverage</div>
            </div>
            <div class="card posture-gap">
                <div class="label">Coverage Gap</div>
                <div class="value">$($S_PostureCounts['Coverage Gap'])</div>
                <div class="detail">Modern auth + Partial CA coverage</div>
            </div>
            <div class="card posture-weakgap">
                <div class="label">Weak &amp; Gap</div>
                <div class="value">$($S_PostureCounts['Weak & Gap'])</div>
                <div class="detail">Legacy auth + Partial CA coverage</div>
            </div>
            <div class="card posture-unenforced">
                <div class="label">Unenforced</div>
                <div class="value">$($S_PostureCounts['Unenforced'])</div>
                <div class="detail">Modern auth + No CA coverage</div>
            </div>
            <div class="card posture-weakunenf">
                <div class="label">Weak &amp; Unenforced</div>
                <div class="value">$($S_PostureCounts['Weak & Unenforced'])</div>
                <div class="detail">Legacy auth + No CA coverage</div>
            </div>
            <div class="card posture-atrisk">
                <div class="label">At Risk</div>
                <div class="value">$($S_PostureCounts['At Risk'])</div>
                <div class="detail">No MFA + Full CA coverage</div>
            </div>
            <div class="card posture-critical">
                <div class="label">Critical</div>
                <div class="value">$($S_PostureCounts['Critical'])</div>
                <div class="detail">No MFA + No CA coverage</div>
            </div>
            <div class="card posture-unknown">
                <div class="label">Unknown</div>
                <div class="value">$($S_PostureCounts['Unknown'])</div>
                <div class="detail">Posture not determined</div>
            </div>
        </div>
    </div>

    <div class="footer">
        Summary view &middot; full report: $(Split-Path $S_HtmlPath -Leaf) &middot; generated by ReportMemberMFA.ps1
    </div>
</body>
</html>
"@

    $S_SummaryHtml | Out-File -FilePath $S_SummaryPath -Encoding UTF8
    Write-Host "Summary HTML exported to: $S_SummaryPath" -ForegroundColor Green
}
catch
{
    Write-Error "An error occurred: $_ at $($_.InvocationInfo.ScriptName):$($_.InvocationInfo.ScriptLineNumber)  -> $($_.InvocationInfo.Line.Trim())"
}
finally
{
    # ── Disconnect ─────────────────────────────────────────────────────────────
    $S_DisconnectChoice = Read-Host "`nDisconnect from Microsoft Graph? [Y] Yes  [N] Keep session  (Default: N)"
    if ($S_DisconnectChoice -eq 'Y')
    {
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null
        Write-Host "Disconnected." -ForegroundColor Green
    }
    else
    {
        Write-Host "Graph session kept alive." -ForegroundColor Green
    }
}
