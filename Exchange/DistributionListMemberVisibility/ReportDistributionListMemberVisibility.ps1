#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Reports every distribution list in Exchange Online together with its members, showing which
    members are hidden from the address lists, and optionally unhides them.

.DESCRIPTION
    Reporting and remediation are deliberately kept apart in this script.

    By default the script is READ-ONLY. It enumerates the selected group types, resolves each
    group's membership, and records for every member whether HiddenFromAddressListsEnabled is
    set - that is, whether the member is hidden from the global address list and from every
    other address list. Hidden members still receive mail sent to the list; they simply cannot
    be found in the address book, which is why a distribution list can look correct while its
    members are invisible to the people who need to find them.

    Nothing in the tenant is changed unless -Unhide is supplied. With -Unhide the script sets
    HiddenFromAddressListsEnabled to $false on each hidden member - once per unique member,
    however many lists it appears in - using the Set-* cmdlet that matches that member's
    recipient type. Every attempt and its outcome is carried into the report, so the report is
    both the before picture and the record of what was changed.

    Not every member can be unhidden the same way, and the report says which route each one
    needs:

      Exchange Online   The Set-* cmdlet for the recipient type - Set-Mailbox, Set-MailUser,
                        Set-DistributionGroup and so on. The ordinary case.
      Microsoft Graph   Entra ID guest users. A guest's address-list visibility is governed by
                        the showInAddressList property on the Entra ID user object rather than by
                        Exchange, so the change runs as Update-MgUser -ShowInAddressList against
                        Entra ID. Guests invited through the invitation manager have the property
                        set to false from the moment they are invited, which is why so many of
                        them turn up hidden without anyone having hidden them.
      On-premises AD    Directory-synchronized objects. Exchange Online refuses the write, so the
                        change has to be made in Active Directory and synchronized.

    Results are written to a single CSV at group-and-member grain, and to a self-contained HTML
    report. The CSV carries the underlying grain; the HTML carries the roll-ups computed from
    it. The HTML makes no external requests - the chart is hand-rolled inline SVG and there is
    no CDN reference - so the report opens on an isolated machine. It carries no branding.

.PARAMETER GroupType
    Which group types to report on. Accepts multiple values.

      Distribution        Distribution groups          (MailUniversalDistributionGroup)
      MailEnabledSecurity Mail-enabled security groups  (MailUniversalSecurityGroup)
      Dynamic             Dynamic distribution groups
      M365Group           Microsoft 365 Groups          (GroupMailbox)
      All                 Every type above

    Defaults to Distribution and MailEnabledSecurity, which together are what
    Get-DistributionGroup returns. Dynamic and M365Group are opt-in because dynamic membership
    is evaluated live and can be very large, and because Microsoft 365 Groups are not
    distribution lists even though they are mail-enabled.

.PARAMETER Group
    One or more group names, aliases or SMTP addresses to limit the report to. Anything not
    found is reported as a warning rather than failing the run. Omit to report on every group of
    the selected types.

.PARAMETER Unhide
    MAKES CHANGES. Sets HiddenFromAddressListsEnabled to $false on every hidden member found,
    once per unique member. Without this switch the script only reports.

    Supports -WhatIf and -Confirm. Before the batch runs, the script prints what it is about to
    change and asks for confirmation unless -Force is supplied.

.PARAMETER UnhideGroups
    MAKES CHANGES. Also unhides the group objects themselves where the list is hidden from the
    address lists. Separate from -Unhide because hiding a list and hiding its members are
    independent decisions and are often deliberate for different reasons.

.PARAMETER Member
    Limits the change to these members. Accepts an SMTP address, alias, display name or Entra ID
    object id, and wildcards (-Member *@fabrikam.com). Only meaningful with -Unhide: the report
    still covers every list and member in scope, and members outside this selection are recorded
    as "Not selected" rather than quietly dropped, so the report still explains why a hidden
    member was left hidden.

    Anything that matches nothing is reported as a warning rather than failing the run.

.PARAMETER ExcludeMember
    The inverse: unhide every hidden member in scope except these. Same matching rules as -Member,
    and the two can be combined - -Member selects the pool, then -ExcludeMember removes from it.

.PARAMETER SelectMembers
    Lists the hidden members found and asks which to unhide. Accepts "all", individual numbers,
    ranges ("2-6"), and exclusions ("all -3,-9" for everything except 3 and 9), so "unhide all of
    these except those two" does not have to be typed out as addresses. Cannot be combined with
    -Member or -ExcludeMember. The selection is itself the confirmation, so no second batch prompt
    follows it.

.PARAMETER SkipGraph
    Leaves Entra ID guest users alone. Guests cannot be unhidden with Set-MailUser - the property
    that governs their visibility lives on the Entra ID user object - so the change goes through
    Microsoft Graph instead. Without this switch the script connects to Graph when, and only when,
    there is at least one hidden guest to change; with it, guests are reported and skipped.

.PARAMETER IncludeDirSynced
    Attempts the change on directory-synchronized objects too. By default they are skipped and
    reported as needing an on-premises change, because in an Exchange hybrid the Exchange
    attributes of a synchronized object are owned by on-premises Active Directory and the write
    is refused in Exchange Online. See the hybrid note under .NOTES.

.PARAMETER Force
    Skips the batch confirmation prompt before changes are made. -WhatIf still takes precedence.

.PARAMETER NoPrefetch
    Resolves each member with an individual Get-Recipient call instead of pre-fetching the whole
    recipient directory once. Slower per member but avoids one tenant-wide enumeration, so it
    wins on a very large tenant that has only a handful of lists in scope.

.PARAMETER TestMode
    Processes a random sample of the groups in scope (see -SampleSize) instead of all of them.
    Use this to validate the report format before committing to a full tenant run. -TestMode
    cannot be combined with -Unhide or -UnhideGroups: remediating a random sample is never what
    anyone wants.

.PARAMETER SampleSize
    Number of groups sampled when -TestMode is specified. Defaults to 10.

.PARAMETER ReportPath
    Folder or file path for the output. If a folder is given, timestamped filenames are
    generated inside it. Defaults to the current directory.

.PARAMETER Format
    Which outputs to produce: Csv, Html or Both. Defaults to Both.

.PARAMETER DelayMilliseconds
    Pause between per-group membership calls, to stay clear of service-side throttling. Defaults
    to 0. Raise it if a large run starts returning throttling errors.

.EXAMPLE
    .\ReportDistributionListMemberVisibility.ps1

    Read-only. Every distribution group and mail-enabled security group, with each member's
    address-list visibility. Writes a CSV and an HTML report to the current folder.

.EXAMPLE
    .\ReportDistributionListMemberVisibility.ps1 -TestMode

    Same, against a random sample of 10 groups, to check the output before a full run.

.EXAMPLE
    .\ReportDistributionListMemberVisibility.ps1 -GroupType All -ReportPath C:\Reports

    Read-only, widened to dynamic distribution groups and Microsoft 365 Groups as well.

.EXAMPLE
    .\ReportDistributionListMemberVisibility.ps1 -Unhide -WhatIf

    Shows exactly which members would be unhidden, and changes nothing.

.EXAMPLE
    .\ReportDistributionListMemberVisibility.ps1 -Group "All Staff" -Unhide

    Unhides every hidden member of one list, after confirmation.

.EXAMPLE
    .\ReportDistributionListMemberVisibility.ps1 -Unhide -UnhideGroups -Force

    Unhides every hidden member of every list in scope, and the hidden lists themselves, without
    the batch confirmation prompt.

.EXAMPLE
    .\ReportDistributionListMemberVisibility.ps1 -Unhide -Member ada@contoso.com,grace@contoso.com

    Unhides two named members and leaves every other hidden member alone.

.EXAMPLE
    .\ReportDistributionListMemberVisibility.ps1 -Unhide -ExcludeMember ceo@contoso.com,chair@contoso.com

    Unhides every hidden member in scope except those two, who stay hidden deliberately.

.EXAMPLE
    .\ReportDistributionListMemberVisibility.ps1 -Unhide -Member *@fabrikam.com

    Unhides the hidden guests from one partner domain, and nothing else.

.EXAMPLE
    .\ReportDistributionListMemberVisibility.ps1 -Group "All Staff" -Unhide -SelectMembers

    Lists the hidden members of one list and asks which to unhide.

.NOTES
    Required module : ExchangeOnlineManagement.
                      Microsoft.Graph.Authentication and Microsoft.Graph.Users as well, but only
                      to unhide Entra ID guest users. They are checked for at the point they are
                      needed rather than required up front, so a read-only run - and a run with no
                      hidden guests in it - needs neither installed.
    Required role   : View-Only Recipients for the report. Mail Recipients (or Recipient
                      Management) for -Unhide and -UnhideGroups. User Administrator, or the
                      User.ReadWrite.All Graph scope, to unhide guest users.

    ENTRA ID GUEST USERS
    --------------------
    A guest (RecipientTypeDetails GuestMailUser) is an Entra ID object that Exchange only surfaces
    as a mail user. Its address-list visibility is the showInAddressList property on the Entra ID
    user, and guests invited through the invitation manager get showInAddressList = false at
    invitation time - so guests appear hidden by default, without anyone having hidden them. The
    change therefore runs through Microsoft Graph:

      Update-MgUser -UserId <entra object id> -ShowInAddressList:$true

    Updating showInAddressList also updates the Exchange Online hide-from-address-lists setting.
    The two are nevertheless stored separately, and where they disagree the Exchange setting wins,
    so after the Graph call this script also clears HiddenFromAddressListsEnabled on the mail user
    when Exchange still has it set. That leaves no contradiction behind for the next person to
    find. The Graph user id is the recipient's ExternalDirectoryObjectId.

      Reference: https://learn.microsoft.com/graph/api/resources/user
                 https://learn.microsoft.com/entra/identity/multi-tenant-organizations/multi-tenant-organization-known-issues

    THE MSAL CLASH, AND WHY GRAPH CONNECTS FIRST
    --------------------------------------------
    ExchangeOnlineManagement and Microsoft.Graph.Authentication ship different, incompatible
    copies of Microsoft.Identity.Client and Microsoft.IdentityModel.Abstractions. Only one
    version of each can be loaded per process, so the first module to AUTHENTICATE wins, and
    the second one's sign-in then dies with:

      Method not found: '!0 Microsoft.Identity.Client.BaseAbstractApplicationBuilder`1
      .WithLogging(Microsoft.IdentityModel.Abstractions.IIdentityLogger, Boolean)'

    because that overload's signature references a type from the other module's copy.

    Updating both modules does NOT fix this. As of October 2026 the newest releases still
    disagree: ExchangeOnlineManagement 3.10.1 carries MSAL 4.83.1 with Abstractions 8.19.2,
    while Microsoft.Graph.Authentication 2.41.2 carries MSAL 4.90.0 with Abstractions 8.18.0.

    So when -Unhide could touch a guest, this script authenticates to Graph BEFORE connecting to
    Exchange Online. Two consequences worth knowing:

      - An -Unhide run without -SkipGraph asks for a Graph sign-in up front, even if it turns out
        there are no hidden guests. Use -SkipGraph to skip that when none are expected.
      - If you connect to Exchange Online yourself before running this script, the Graph half is
        already lost for that session. The run still completes: guests are reported as needing
        Graph, and the companion script described below is written so a fresh session can finish.

    COMPANION GUEST SCRIPT
    ----------------------
    Whenever hidden guests are left unchanged - no -Unhide, -SkipGraph, or a failed Graph
    connection - a <report name>_GuestRemediation.ps1 is written next to the report. It carries
    the affected guests as data and does the Graph work before touching Exchange, so running it
    in a fresh session finishes the job without anyone rebuilding the command list by hand. It
    supports -WhatIf.

    HYBRID / DIRECTORY-SYNCHRONIZED OBJECTS
    ---------------------------------------
    In an Exchange hybrid, the Exchange attributes of a directory-synchronized object are owned
    by on-premises Active Directory, and Exchange Online refuses the write. For those objects the
    change has to be made on-premises instead - clear msExchHideFromAddressLists (or run
    Set-RemoteMailbox -HiddenFromAddressListsEnabled $false on the last Exchange server) and let
    it synchronize. The report flags every such member so the on-premises work list is explicit,
    and -Unhide skips them unless -IncludeDirSynced is given. Note also that the hide attribute
    only synchronizes when the object's mailNickname (alias) is populated; an empty alias is the
    usual reason an on-premises change never appears in the cloud.

      Reference: https://learn.microsoft.com/exchange/hybrid-deployment/enable-exchange-attributes-cloud-management
                 https://learn.microsoft.com/troubleshoot/exchange/administration/changes-to-msexchangehiddenfromaddresslist-not-updated-against-recipient

    Tenants that have transferred Exchange-attribute Source of Authority to the cloud can be
    changed in Exchange Online normally; use -IncludeDirSynced there.

    WHAT THIS REPORT DOES NOT COVER
    -------------------------------
      - Address book policies. A member that is visible in the GAL can still be invisible to a
        given user through an ABP. Visibility here means HiddenFromAddressListsEnabled only.
      - Microsoft 365 Groups also carry HiddenFromExchangeClientsEnabled, which hides the group
        from Outlook independently of the address lists. Get-Recipient does not return it, so it
        is neither reported nor changed; check it with Get-UnifiedGroup if a group stays hidden
        in Outlook after being unhidden here.
      - Members that are not mail-enabled recipients never appear in distribution list
        membership, so they cannot appear here either.

    Reference: https://learn.microsoft.com/exchange/address-books/address-lists/manage-address-lists
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('Distribution', 'MailEnabledSecurity', 'Dynamic', 'M365Group', 'All')]
    [string[]]$GroupType = @('Distribution', 'MailEnabledSecurity'),

    [Parameter(Mandatory = $false)]
    [string[]]$Group,

    [Parameter(Mandatory = $false)]
    [switch]$Unhide,

    [Parameter(Mandatory = $false)]
    [switch]$UnhideGroups,

    [Parameter(Mandatory = $false)]
    [string[]]$Member,

    [Parameter(Mandatory = $false)]
    [string[]]$ExcludeMember,

    [Parameter(Mandatory = $false)]
    [switch]$SelectMembers,

    [Parameter(Mandatory = $false)]
    [switch]$SkipGraph,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeDirSynced,

    [Parameter(Mandatory = $false)]
    [switch]$Force,

    [Parameter(Mandatory = $false)]
    [switch]$NoPrefetch,

    [Parameter(Mandatory = $false)]
    [switch]$TestMode,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 500)]
    [int]$SampleSize = 10,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ReportPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Csv', 'Html', 'Both')]
    [string]$Format = 'Both',

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 10000)]
    [int]$DelayMilliseconds = 0
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable))
{
    throw "ExchangeOnlineManagement module is not installed. Install it using: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}

# Targeting switches that silently do nothing are worse than an error, so they are refused
# outright on a read-only run rather than being accepted and ignored.
if (($Member -or $ExcludeMember -or $SelectMembers) -and -not $Unhide)
{
    throw "-Member, -ExcludeMember and -SelectMembers narrow which members -Unhide changes, so they do nothing on a read-only run. Add -Unhide, or drop them. To narrow the report itself, use -Group."
}

if ($SelectMembers -and ($Member -or $ExcludeMember))
{
    throw "-SelectMembers chooses members interactively, so it cannot be combined with -Member or -ExcludeMember. Use one or the other."
}

if ($TestMode -and ($Unhide -or $UnhideGroups))
{
    throw "-TestMode cannot be combined with -Unhide or -UnhideGroups. Remediating a random sample of groups leaves the tenant in a half-changed state. Run the report with -TestMode first, then run -Unhide without it."
}

# ---------------------------------------------------------------------------
# Scope expansion
# ---------------------------------------------------------------------------
$S_Types = if ($GroupType -contains 'All')
{
    @('Distribution', 'MailEnabledSecurity', 'Dynamic', 'M365Group')
}
else
{
    @($GroupType | Select-Object -Unique)
}

# Only needed to unhide guests, and only requested at that point.
$S_RequiredGraphScopes = @('User.ReadWrite.All')

# $null = Graph not attempted yet, $true / $false = the outcome of the one attempt made.
$S_GraphReady = $null

# RecipientTypeDetails values that Get-DistributionGroup returns, keyed by the friendly scope name.
$S_StaticTypeMap = @{
    'Distribution'        = @('MailUniversalDistributionGroup', 'MailNonUniversalGroup', 'RoomList')
    'MailEnabledSecurity' = @('MailUniversalSecurityGroup')
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

# HTML-encode, tolerating nulls, so every value interpolated into the report is safe.
function Format-Html
{
    param([object]$Value)

    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode("$Value")
}

# Read a property that may not exist on the object the service returned, rather than assuming
# a schema. Get-UnifiedGroupLinks and Get-DistributionGroupMember return narrower objects than
# Get-Recipient does.
function Get-PropertyOrNull
{
    param($Object, [string]$Name)

    if ($null -eq $Object) { return $null }
    $S_Prop = $Object.PSObject.Properties[$Name]
    if ($null -eq $S_Prop) { return $null }
    return $S_Prop.Value
}

# Coerce HiddenFromAddressListsEnabled / IsDirSynced into $true, $false or $null (unknown).
# A missing property and a $false value mean very different things in this report, so they are
# never collapsed together.
function ConvertTo-NullableBool
{
    param($Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [bool]) { return $Value }
    $S_Text = "$Value".Trim()
    if ([string]::IsNullOrWhiteSpace($S_Text)) { return $null }
    switch ($S_Text)
    {
        'True'  { return $true }
        'False' { return $false }
        '1'     { return $true }
        '0'     { return $false }
        default { return $null }
    }
}

# The Set-* cmdlet that owns HiddenFromAddressListsEnabled for a given recipient type. There is
# no single cmdlet for this: Exchange splits it by recipient type, and calling the wrong one
# fails outright.
#   https://learn.microsoft.com/exchange/address-books/address-lists/manage-address-lists
function Get-VisibilityCmdlet
{
    param([string]$RecipientTypeDetails)

    $S_Type = "$RecipientTypeDetails".Trim()
    switch -Regex ($S_Type)
    {
        '^(UserMailbox|SharedMailbox|RoomMailbox|EquipmentMailbox|SchedulingMailbox|TeamMailbox|LinkedMailbox|LegacyMailbox)$'
        {
            return 'Set-Mailbox'
        }
        '^(MailUser|RemoteUserMailbox|RemoteSharedMailbox|RemoteRoomMailbox|RemoteEquipmentMailbox)$'
        {
            # Remote* types are mail users as far as Exchange Online is concerned. GuestMailUser
            # is deliberately absent: a guest is an Entra ID object and goes through Graph.
            return 'Set-MailUser'
        }
        '^MailContact$'
        {
            return 'Set-MailContact'
        }
        '^(MailUniversalDistributionGroup|MailUniversalSecurityGroup|MailNonUniversalGroup|RoomList)$'
        {
            return 'Set-DistributionGroup'
        }
        '^DynamicDistributionGroup$'
        {
            return 'Set-DynamicDistributionGroup'
        }
        '^GroupMailbox$'
        {
            return 'Set-UnifiedGroup'
        }
        '^(PublicFolder|MailUniversalPublicFolder)$'
        {
            return 'Set-MailPublicFolder'
        }
        default
        {
            return $null
        }
    }
}

# Which mechanism can actually clear the flag on this recipient. Guests are tested first and
# deliberately beat the directory-synced test: a guest's visibility lives on the Entra ID object,
# which Graph can write whatever the sync state of the rest of the object.
function Get-ChangePath
{
    param([string]$RecipientTypeDetails, $IsDirSynced)

    if ("$RecipientTypeDetails".Trim() -eq 'GuestMailUser') { return 'Graph' }
    if ($IsDirSynced -eq $true) { return 'OnPremises' }
    if (Get-VisibilityCmdlet -RecipientTypeDetails $RecipientTypeDetails) { return 'ExchangeOnline' }
    return 'Unsupported'
}

function Get-ChangePathLabel
{
    param([string]$ChangePath)

    switch ($ChangePath)
    {
        'Graph'          { 'Microsoft Graph (guest)' }
        'OnPremises'     { 'On-premises AD' }
        'ExchangeOnline' { 'Exchange Online' }
        default          { 'Unsupported' }
    }
}

# Match a member against one user-supplied token, on any of the identifiers someone would
# reasonably paste out of the report: address, alias, display name or Entra object id. A token
# containing * or ? is treated as a wildcard, so a whole partner domain can be named at once.
function Test-MemberToken
{
    param($Member, [string]$Token)

    $S_T = "$Token".Trim()
    if (-not $S_T) { return $false }
    $S_Wild = $S_T -match '[\*\?]'

    foreach ($S_Cand in @($Member.PrimarySmtpAddress, $Member.Alias, $Member.DisplayName, $Member.GraphUserId, $Member.Key))
    {
        if (-not $S_Cand) { continue }
        $S_Val = "$S_Cand".Trim()
        if ($S_Wild) { if ($S_Val -ilike $S_T) { return $true } }
        elseif ($S_Val -ieq $S_T) { return $true }
    }
    return $false
}

# Console picker over the hidden members. It exists because the usual ask is neither "unhide
# everything" nor "unhide this one address", but "unhide all of these except those two" - which is
# tedious and error-prone to express as a list of addresses copied out of a report.
function Select-HiddenMembers
{
    param($HiddenMembers)

    $S_List = @($HiddenMembers)
    if ($S_List.Count -eq 0) { return @() }

    Write-Host ""
    Write-Host "Hidden members in scope:" -ForegroundColor Cyan
    for ($S_i = 0; $S_i -lt $S_List.Count; $S_i++)
    {
        $S_M     = $S_List[$S_i]
        $S_Num   = ($S_i + 1).ToString().PadLeft(4)
        $S_Addr  = if ($S_M.PrimarySmtpAddress) { $S_M.PrimarySmtpAddress } else { '(no address)' }
        $S_Route = switch ($S_M.ChangePath)
        {
            'Graph'          { 'Graph (guest)' }
            'OnPremises'     { 'on-premises AD' }
            'ExchangeOnline' { 'Exchange Online' }
            default          { 'unsupported' }
        }
        Write-Host "$S_Num. $($S_M.DisplayName)  <$S_Addr>  [$($S_M.RecipientTypeDetails), via $S_Route]"
    }

    Write-Host ""
    Write-Host "Choose what to unhide:" -ForegroundColor Cyan
    Write-Host "  all          every member listed above"
    Write-Host "  1,4,7        only these"
    Write-Host "  2-6          a range"
    Write-Host "  all -3,-9    everything except these"
    Write-Host "  none         cancel, change nothing"
    Write-Host ""

    # Bounded on purpose: an unbounded prompt loop spins forever in a non-interactive session,
    # where Read-Host returns empty immediately rather than waiting for a person.
    $S_MaxTries = 5
    $S_Tries    = 0

    while ($true)
    {
        $S_Tries++
        if ($S_Tries -gt $S_MaxTries)
        {
            Write-Warning "No usable selection after $S_MaxTries attempts; treating this as a cancellation. Nothing will be changed. If this is a non-interactive session, use -Member or -ExcludeMember instead of -SelectMembers."
            return $null
        }

        $S_Answer = "$(Read-Host 'Selection')".Trim()
        if (-not $S_Answer)
        {
            Write-Host "Enter a selection, or 'none' to cancel. ($($S_MaxTries - $S_Tries) attempt(s) left)" -ForegroundColor Yellow
            continue
        }
        if ($S_Answer -ieq 'none' -or $S_Answer -ieq 'q') { return $null }

        $S_Include = [System.Collections.Generic.List[int]]::new()
        $S_Exclude = [System.Collections.Generic.List[int]]::new()
        $S_All     = $false
        $S_Bad     = $null

        foreach ($S_Tok in @($S_Answer -split '[,\s]+' | Where-Object { $_ }))
        {
            if ($S_Tok -ieq 'all')             { $S_All = $true; continue }
            if ($S_Tok -match '^-(\d+)$')      { $S_Exclude.Add([int]$Matches[1]); continue }
            if ($S_Tok -match '^(\d+)-(\d+)$') { foreach ($S_N in ([int]$Matches[1])..([int]$Matches[2])) { $S_Include.Add($S_N) }; continue }
            if ($S_Tok -match '^\d+$')         { $S_Include.Add([int]$S_Tok); continue }
            $S_Bad = $S_Tok
            break
        }

        if ($S_Bad)
        {
            Write-Host "Could not read '$S_Bad'. Use numbers, ranges like 2-6, 'all', or -3 to exclude." -ForegroundColor Yellow
            continue
        }

        $S_Range = @(@($S_Include) + @($S_Exclude) | Where-Object { $_ -lt 1 -or $_ -gt $S_List.Count } | Sort-Object -Unique)
        if ($S_Range.Count -gt 0)
        {
            Write-Host "Out of range: $($S_Range -join ', '). Pick between 1 and $($S_List.Count)." -ForegroundColor Yellow
            continue
        }

        $S_Chosen = if ($S_All) { @(1..$S_List.Count) } else { @($S_Include) }
        $S_Chosen = @($S_Chosen | Where-Object { $S_Exclude -notcontains $_ } | Sort-Object -Unique)

        if ($S_Chosen.Count -eq 0)
        {
            Write-Host "That selects nothing. Enter 'none' to cancel, or pick at least one." -ForegroundColor Yellow
            continue
        }

        $S_Result = @($S_Chosen | ForEach-Object { $S_List[$_ - 1] })
        Write-Host ""
        Write-Host "Selected $($S_Result.Count) of $($S_List.Count):" -ForegroundColor Green
        foreach ($S_R in $S_Result) { Write-Host "  $($S_R.DisplayName) <$($S_R.PrimarySmtpAddress)>" -ForegroundColor Green }
        if ($S_Result.Count -lt $S_List.Count)
        {
            Write-Host "Leaving hidden: $($S_List.Count - $S_Result.Count)" -ForegroundColor Yellow
        }
        Write-Host ""
        if ((Read-Host "Proceed with this selection? [Y] Yes  [N] Choose again  (Default: Y)") -ieq 'N')
        {
            $S_Tries = 0   # a valid parse followed by a deliberate re-pick is not a failed attempt
            continue
        }
        return $S_Result
    }
}

# Connect to Microsoft Graph, once, at the point a guest actually needs changing. Returns whether
# Graph is usable; the result is cached so a failure is reported once rather than per guest.
function Initialize-GraphConnection
{
    if ($null -ne $script:S_GraphReady) { return $script:S_GraphReady }
    $script:S_GraphReady = $false

    foreach ($S_Mod in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users'))
    {
        if (-not (Get-Module -Name $S_Mod -ListAvailable))
        {
            Write-Warning "Microsoft Graph module '$S_Mod' is not installed, so Entra ID guest users cannot be unhidden. Install it with: Install-Module $S_Mod -Scope CurrentUser"
            return $false
        }
    }

    try
    {
        $S_Ctx = Get-MgContext -ErrorAction SilentlyContinue
        if ($S_Ctx)
        {
            Write-Host ""
            Write-Host "Existing Microsoft Graph session detected:" -ForegroundColor Yellow
            Write-Host "  Account : $($S_Ctx.Account)" -ForegroundColor Yellow
            Write-Host "  TenantId: $($S_Ctx.TenantId)" -ForegroundColor Yellow
            Write-Host "  Scopes  : $($S_Ctx.Scopes -join ', ')" -ForegroundColor Yellow
            Write-Host ""

            $S_Missing = @($S_RequiredGraphScopes | Where-Object { $S_Ctx.Scopes -notcontains $_ })
            if ($S_Missing.Count -gt 0)
            {
                Write-Host "This session lacks the scope needed to update guest users ($($S_Missing -join ', ')). Reconnecting..." -ForegroundColor Cyan
                Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome -ErrorAction Stop
            }
            else
            {
                $S_Choice = Read-Host "Use this Graph session to update guest users? [Y] Yes  [N] Disconnect and reconnect  (Default: Y)"
                if ($S_Choice -eq 'N')
                {
                    Disconnect-MgGraph | Out-Null
                    Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome -ErrorAction Stop
                }
            }
        }
        else
        {
            Write-Host ""
            Write-Host "Connecting to Microsoft Graph to update Entra ID guest users (scope: $($S_RequiredGraphScopes -join ', '))..." -ForegroundColor Cyan
            Connect-MgGraph -Scopes $S_RequiredGraphScopes -NoWelcome -ErrorAction Stop
        }

        $S_Ctx = Get-MgContext -ErrorAction SilentlyContinue
        if (-not $S_Ctx) { throw 'No Microsoft Graph context after connecting.' }

        # Graph has its own session, so it can be pointed at a different tenant than the Exchange
        # session. Unhiding guests in the wrong tenant would be silent, so it is checked.
        if ($S_TenantId -and $S_TenantId -ne 'Unknown' -and "$($S_Ctx.TenantId)" -ne "$S_TenantId")
        {
            Write-Warning "Microsoft Graph is connected to tenant $($S_Ctx.TenantId) but Exchange Online is connected to $S_TenantId. Guest users will NOT be updated. Reconnect Graph to the same tenant and re-run."
            return $false
        }

        Write-Host "  Graph connected: $($S_Ctx.Account) ($($S_Ctx.TenantId))" -ForegroundColor Green
        $script:S_GraphReady = $true
        return $true
    }
    catch
    {
        $S_Msg = "$($_.Exception.Message)"
        Write-Warning "Could not connect to Microsoft Graph: $S_Msg"

        # The MSAL clash is worth naming explicitly, because the raw error ("Method not found:
        # ...WithLogging...") gives no hint of the cause and sends people hunting for a Graph
        # permissions problem that is not there.
        if ($S_Msg -match 'Method not found' -or $S_Msg -match 'Microsoft\.Identity\.Client' -or $S_Msg -match 'IIdentityLogger' -or $S_Msg -match 'IdentityModel')
        {
            Write-Warning ''
            Write-Warning 'This looks like the MSAL assembly clash between ExchangeOnlineManagement and'
            Write-Warning 'Microsoft.Graph.Authentication. The two modules ship different, incompatible'
            Write-Warning 'Microsoft.Identity.Client / Microsoft.IdentityModel.Abstractions versions, and'
            Write-Warning 'whichever one authenticates FIRST in a session wins; the second then fails.'
            Write-Warning 'Updating both modules does not fix it - the newest releases still disagree.'
            Write-Warning ''
            Write-Warning 'Ways forward, cheapest first:'
            Write-Warning '  1. Run the companion guest script this run writes, in a FRESH session. It does'
            Write-Warning '     the Graph work before touching Exchange, which is the order that works.'
            Write-Warning '  2. Or, in a fresh session, run Connect-MgGraph -Scopes User.ReadWrite.All'
            Write-Warning '     BEFORE this script, so Graph authenticates first.'
            Write-Warning '  3. Or re-run with -SkipGraph and handle the guests separately.'
        }
        Write-Warning 'Entra ID guest users will NOT be updated by this run.'
        return $false
    }
}

# Friendly label for a group's RecipientTypeDetails, used in the report and its filters.
function Get-GroupTypeLabel
{
    param([string]$RecipientTypeDetails)

    switch ("$RecipientTypeDetails".Trim())
    {
        'MailUniversalDistributionGroup' { 'Distribution list' }
        'MailUniversalSecurityGroup'     { 'Mail-enabled security' }
        'MailNonUniversalGroup'          { 'Distribution list (non-universal)' }
        'RoomList'                       { 'Room list' }
        'DynamicDistributionGroup'       { 'Dynamic distribution list' }
        'GroupMailbox'                   { 'Microsoft 365 Group' }
        default                          { if ($RecipientTypeDetails) { "$RecipientTypeDetails" } else { 'Group' } }
    }
}

# Hand-rolled horizontal bar chart. Inline SVG keeps the report self-contained: no CDN, no
# outbound request when the file is opened.
function New-BarChartSvg
{
    param(
        [string[]]$Labels,
        [int[]]$Values,
        [int]$Width = 620
    )

    if ($Labels.Count -eq 0) { return '<p class="muted">Nothing to chart.</p>' }

    $S_RowHeight  = 28
    $S_LabelWidth = 210
    $S_BarMax     = $Width - $S_LabelWidth - 60
    $S_Height     = ($Labels.Count * $S_RowHeight) + 8
    $S_Max        = ($Values | Measure-Object -Maximum).Maximum
    if ($S_Max -le 0) { $S_Max = 1 }

    $S_Parts = [System.Collections.Generic.List[string]]::new()
    $S_Parts.Add("<svg viewBox=`"0 0 $Width $S_Height`" width=`"100%`" height=`"$S_Height`" role=`"img`" class=`"barchart`">")
    for ($S_i = 0; $S_i -lt $Labels.Count; $S_i++)
    {
        $S_Y     = ($S_i * $S_RowHeight) + 4
        $S_Val   = [int]$Values[$S_i]
        $S_BarW  = [Math]::Max(1, [int](($S_Val / $S_Max) * $S_BarMax))
        $S_Lbl   = Format-Html $Labels[$S_i]
        $S_TextY = $S_Y + 14
        $S_Parts.Add("<text x=`"0`" y=`"$S_TextY`" class=`"bar-label`">$S_Lbl</text>")
        $S_Parts.Add("<rect x=`"$S_LabelWidth`" y=`"$S_Y`" width=`"$S_BarW`" height=`"18`" rx=`"3`" class=`"bar`"></rect>")
        $S_Parts.Add("<text x=`"$($S_LabelWidth + $S_BarW + 8)`" y=`"$S_TextY`" class=`"bar-value`">$S_Val</text>")
    }
    $S_Parts.Add('</svg>')
    return ($S_Parts -join "`n")
}

# ---------------------------------------------------------------------------
# Connect to Microsoft Graph FIRST, when guests might need changing
# ---------------------------------------------------------------------------
# Order is load-bearing, not stylistic. ExchangeOnlineManagement and
# Microsoft.Graph.Authentication ship different, incompatible Microsoft.Identity.Client and
# Microsoft.IdentityModel.Abstractions versions, and the first module to authenticate in a
# session loads its pair for the whole process - after which the other one's auth fails with
# "Method not found: ...WithLogging(IIdentityLogger, Boolean)". Updating both modules does not
# fix it; their newest releases still disagree. So when a guest change is possible, Graph goes
# first. -SkipGraph skips this entirely, and a read-only or -WhatIf run never needs it.
if ($Unhide -and -not $SkipGraph -and -not $WhatIfPreference)
{
    Write-Host "Connecting to Microsoft Graph before Exchange Online." -ForegroundColor Cyan
    Write-Host "  The two modules ship incompatible MSAL versions and the first to authenticate wins," -ForegroundColor Gray
    Write-Host "  so Graph must go first for Entra ID guest users to be changeable in this session." -ForegroundColor Gray
    Write-Host "  Use -SkipGraph to skip this if no guests are involved." -ForegroundColor Gray
    Initialize-GraphConnection | Out-Null
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

# The Graph preflight ran before the Exchange tenant was known, so the two tenants are compared
# here instead. Changing guests in the wrong tenant would be silent, so a mismatch disables Graph
# rather than being reported after the fact.
if ($S_GraphReady -eq $true -and $S_TenantId -ne 'Unknown')
{
    $S_GraphCtx = Get-MgContext -ErrorAction SilentlyContinue
    if ($S_GraphCtx -and "$($S_GraphCtx.TenantId)" -ne "$S_TenantId")
    {
        Write-Warning "Microsoft Graph is connected to tenant $($S_GraphCtx.TenantId) but Exchange Online is connected to $S_TenantId. Guest users will NOT be updated. Reconnect Graph to the same tenant and re-run."
        $S_GraphReady = $false
    }
}

# ---------------------------------------------------------------------------
# Recipient directory
# ---------------------------------------------------------------------------
# Membership calls return thin objects: Get-UnifiedGroupLinks does not carry
# HiddenFromAddressListsEnabled at all, and nothing but Get-Recipient reliably carries
# IsDirSynced. One tenant-wide enumeration, indexed locally, costs far fewer round trips than
# one Get-Recipient per member - but -NoPrefetch exists for the tenant where that trade goes the
# other way.
$S_RecipientIndex = @{}
$S_LookupMisses   = @{}

function Add-RecipientToIndex
{
    param($Recipient)

    if ($null -eq $Recipient) { return }
    foreach ($S_KeyName in @('ExternalDirectoryObjectId', 'PrimarySmtpAddress', 'DistinguishedName', 'Guid', 'Identity', 'Alias'))
    {
        $S_KeyVal = "$(Get-PropertyOrNull -Object $Recipient -Name $S_KeyName)".Trim().ToLowerInvariant()
        if ($S_KeyVal -and -not $S_RecipientIndex.ContainsKey($S_KeyVal))
        {
            $S_RecipientIndex[$S_KeyVal] = $Recipient
        }
    }
}

if (-not $NoPrefetch)
{
    Write-Host ""
    Write-Host "Pre-fetching the recipient directory (one enumeration, indexed locally)..." -ForegroundColor Cyan
    $S_AllRecipients = @(Get-Recipient -ResultSize Unlimited)
    foreach ($S_R in $S_AllRecipients) { Add-RecipientToIndex -Recipient $S_R }

    # Microsoft 365 Groups are excluded from Get-Recipient unless asked for by name.
    try
    {
        $S_GroupMailboxes = @(Get-Recipient -RecipientTypeDetails GroupMailbox -ResultSize Unlimited)
        foreach ($S_R in $S_GroupMailboxes) { Add-RecipientToIndex -Recipient $S_R }
    }
    catch
    {
        Write-Warning "Could not enumerate Microsoft 365 Groups as recipients: $($_.Exception.Message)"
        $S_GroupMailboxes = @()
    }

    Write-Host "  Recipients indexed : $($S_AllRecipients.Count + $S_GroupMailboxes.Count)" -ForegroundColor Green
}

# Resolve one membership entry to a full recipient object, by index then by targeted lookup.
function Resolve-Recipient
{
    param($Member)

    $S_Keys = @()
    foreach ($S_KeyName in @('ExternalDirectoryObjectId', 'PrimarySmtpAddress', 'DistinguishedName', 'Guid', 'Identity', 'Alias'))
    {
        $S_KeyVal = "$(Get-PropertyOrNull -Object $Member -Name $S_KeyName)".Trim()
        if ($S_KeyVal) { $S_Keys += $S_KeyVal }
    }

    foreach ($S_Key in $S_Keys)
    {
        $S_Lower = $S_Key.ToLowerInvariant()
        if ($S_RecipientIndex.ContainsKey($S_Lower)) { return $S_RecipientIndex[$S_Lower] }
    }

    # Not indexed (or -NoPrefetch): one targeted call, cached both ways so a member that appears
    # in twenty lists is still only looked up once.
    foreach ($S_Key in $S_Keys)
    {
        $S_Lower = $S_Key.ToLowerInvariant()
        if ($S_LookupMisses.ContainsKey($S_Lower)) { continue }
        try
        {
            $S_Found = Get-Recipient -Identity $S_Key -ErrorAction Stop | Select-Object -First 1
            if ($S_Found)
            {
                Add-RecipientToIndex -Recipient $S_Found
                return $S_Found
            }
        }
        catch
        {
            $S_LookupMisses[$S_Lower] = $true
        }
    }
    return $null
}

# ---------------------------------------------------------------------------
# Retrieve groups
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "Retrieving groups (types: $($S_Types -join ', '))..." -ForegroundColor Cyan

$S_RawGroups = [System.Collections.Generic.List[PSCustomObject]]::new()
$S_NotFound  = [System.Collections.Generic.List[string]]::new()

# Wrap each source cmdlet's output in a common shape, so the membership loop does not care
# where a group came from.
function Add-GroupRecord
{
    param($Source, [string]$Kind)

    $S_Rtd = "$(Get-PropertyOrNull -Object $Source -Name 'RecipientTypeDetails')".Trim()
    if (-not $S_Rtd)
    {
        $S_Rtd = switch ($Kind)
        {
            'Dynamic'   { 'DynamicDistributionGroup' }
            'M365Group' { 'GroupMailbox' }
            default     { 'MailUniversalDistributionGroup' }
        }
    }

    $S_RawGroups.Add([PSCustomObject]@{
        Kind                 = $Kind
        Identity             = "$(Get-PropertyOrNull -Object $Source -Name 'Identity')"
        Name                 = "$(Get-PropertyOrNull -Object $Source -Name 'Name')"
        DisplayName          = "$(Get-PropertyOrNull -Object $Source -Name 'DisplayName')"
        Alias                = "$(Get-PropertyOrNull -Object $Source -Name 'Alias')"
        PrimarySmtpAddress   = "$(Get-PropertyOrNull -Object $Source -Name 'PrimarySmtpAddress')"
        RecipientTypeDetails = $S_Rtd
        Hidden               = ConvertTo-NullableBool (Get-PropertyOrNull -Object $Source -Name 'HiddenFromAddressListsEnabled')
        IsDirSynced          = ConvertTo-NullableBool (Get-PropertyOrNull -Object $Source -Name 'IsDirSynced')
        RecipientFilter      = "$(Get-PropertyOrNull -Object $Source -Name 'RecipientFilter')"
        Raw                  = $Source
    })
}

if ($Group)
{
    # Explicit list: resolve each name against every source cmdlet it could belong to.
    foreach ($S_Wanted in $Group)
    {
        $S_Resolved = $false

        if (($S_Types -contains 'Distribution') -or ($S_Types -contains 'MailEnabledSecurity'))
        {
            try
            {
                $S_G = Get-DistributionGroup -Identity $S_Wanted -ErrorAction Stop
                foreach ($S_One in @($S_G)) { Add-GroupRecord -Source $S_One -Kind 'Static' }
                $S_Resolved = $true
            }
            catch { }
        }
        if (-not $S_Resolved -and ($S_Types -contains 'Dynamic'))
        {
            try
            {
                $S_G = Get-DynamicDistributionGroup -Identity $S_Wanted -ErrorAction Stop
                foreach ($S_One in @($S_G)) { Add-GroupRecord -Source $S_One -Kind 'Dynamic' }
                $S_Resolved = $true
            }
            catch { }
        }
        if (-not $S_Resolved -and ($S_Types -contains 'M365Group'))
        {
            try
            {
                $S_G = Get-UnifiedGroup -Identity $S_Wanted -ErrorAction Stop
                foreach ($S_One in @($S_G)) { Add-GroupRecord -Source $S_One -Kind 'M365Group' }
                $S_Resolved = $true
            }
            catch { }
        }

        if (-not $S_Resolved)
        {
            Write-Warning "Group not found in the selected group types: $S_Wanted"
            $S_NotFound.Add($S_Wanted)
        }
    }
}
else
{
    if (($S_Types -contains 'Distribution') -or ($S_Types -contains 'MailEnabledSecurity'))
    {
        # One call returns both distribution groups and mail-enabled security groups; filter
        # locally rather than guessing at server-side filter syntax.
        $S_Wanted = @()
        foreach ($S_T in @('Distribution', 'MailEnabledSecurity'))
        {
            if ($S_Types -contains $S_T) { $S_Wanted += $S_StaticTypeMap[$S_T] }
        }
        foreach ($S_G in @(Get-DistributionGroup -ResultSize Unlimited))
        {
            if ($S_Wanted -contains "$($S_G.RecipientTypeDetails)") { Add-GroupRecord -Source $S_G -Kind 'Static' }
        }
    }
    if ($S_Types -contains 'Dynamic')
    {
        foreach ($S_G in @(Get-DynamicDistributionGroup -ResultSize Unlimited)) { Add-GroupRecord -Source $S_G -Kind 'Dynamic' }
    }
    if ($S_Types -contains 'M365Group')
    {
        foreach ($S_G in @(Get-UnifiedGroup -ResultSize Unlimited)) { Add-GroupRecord -Source $S_G -Kind 'M365Group' }
    }
}

$S_GroupsInScope = @($S_RawGroups | Sort-Object DisplayName, Name)

if ($TestMode -and $S_GroupsInScope.Count -gt $SampleSize)
{
    Write-Host "TEST MODE: sampling $SampleSize of $($S_GroupsInScope.Count) groups..." -ForegroundColor Cyan
    $S_GroupsInScope = @($S_GroupsInScope | Get-Random -Count $SampleSize | Sort-Object DisplayName, Name)
}

Write-Host "  Found $($S_GroupsInScope.Count) groups to process" -ForegroundColor Green

if ($S_GroupsInScope.Count -eq 0)
{
    Write-Warning "No groups matched the selected scope. Nothing to report."
    return
}

# ---------------------------------------------------------------------------
# Walk membership
# ---------------------------------------------------------------------------
# One record per unique member, keyed so a member appearing in many lists is resolved and
# (later) changed exactly once.
$S_Members = @{}
$S_GroupReport = [System.Collections.Generic.List[PSCustomObject]]::new()

$S_Counter = 0
foreach ($S_G in $S_GroupsInScope)
{
    $S_Counter++
    $S_GroupLabel = if ($S_G.DisplayName) { $S_G.DisplayName } else { $S_G.Name }
    Write-Progress -Activity "Reading distribution list membership" -Status "$S_GroupLabel" -PercentComplete (($S_Counter / $S_GroupsInScope.Count) * 100)
    Write-Host "[$S_Counter/$($S_GroupsInScope.Count)] $S_GroupLabel..." -ForegroundColor Gray

    $S_RawMembers = @()
    $S_GroupError = $null
    try
    {
        switch ($S_G.Kind)
        {
            'Dynamic'
            {
                if ([string]::IsNullOrWhiteSpace($S_G.RecipientFilter))
                {
                    $S_GroupError = 'Dynamic group has no recipient filter; membership cannot be previewed.'
                }
                else
                {
                    $S_RawMembers = @(Get-Recipient -RecipientPreviewFilter $S_G.RecipientFilter -ResultSize Unlimited -ErrorAction Stop)
                }
            }
            'M365Group'
            {
                $S_RawMembers = @(Get-UnifiedGroupLinks -Identity $S_G.Identity -LinkType Members -ResultSize Unlimited -ErrorAction Stop)
            }
            default
            {
                $S_RawMembers = @(Get-DistributionGroupMember -Identity $S_G.Identity -ResultSize Unlimited -ErrorAction Stop)
            }
        }
    }
    catch
    {
        $S_GroupError = $_.Exception.Message
        Write-Warning "  Could not read members of '$S_GroupLabel': $($_.Exception.Message)"
    }

    $S_MemberKeys   = [System.Collections.Generic.List[string]]::new()
    $S_HiddenInThis = 0

    foreach ($S_M in $S_RawMembers)
    {
        # Prefer the resolved recipient; fall back to the membership object where a member is
        # not resolvable (deleted mid-run, or outside the caller's recipient scope).
        $S_Rec = Resolve-Recipient -Member $S_M
        $S_Use = if ($S_Rec) { $S_Rec } else { $S_M }

        $S_Key = "$(Get-PropertyOrNull -Object $S_Use -Name 'ExternalDirectoryObjectId')".Trim()
        if (-not $S_Key) { $S_Key = "$(Get-PropertyOrNull -Object $S_Use -Name 'PrimarySmtpAddress')".Trim() }
        if (-not $S_Key) { $S_Key = "$(Get-PropertyOrNull -Object $S_Use -Name 'DistinguishedName')".Trim() }
        if (-not $S_Key) { $S_Key = "$(Get-PropertyOrNull -Object $S_Use -Name 'Identity')".Trim() }
        if (-not $S_Key) { continue }
        $S_Key = $S_Key.ToLowerInvariant()

        if (-not $S_Members.ContainsKey($S_Key))
        {
            $S_Hidden = ConvertTo-NullableBool (Get-PropertyOrNull -Object $S_Use -Name 'HiddenFromAddressListsEnabled')
            $S_Rtd    = "$(Get-PropertyOrNull -Object $S_Use -Name 'RecipientTypeDetails')".Trim()
            $S_Smtp   = "$(Get-PropertyOrNull -Object $S_Use -Name 'PrimarySmtpAddress')".Trim()
            $S_Disp   = "$(Get-PropertyOrNull -Object $S_Use -Name 'DisplayName')".Trim()
            if (-not $S_Disp) { $S_Disp = "$(Get-PropertyOrNull -Object $S_Use -Name 'Name')".Trim() }
            $S_Sync   = ConvertTo-NullableBool (Get-PropertyOrNull -Object $S_Use -Name 'IsDirSynced')
            $S_Eid    = "$(Get-PropertyOrNull -Object $S_Use -Name 'ExternalDirectoryObjectId')".Trim()
            $S_Path   = Get-ChangePath -RecipientTypeDetails $S_Rtd -IsDirSynced $S_Sync

            $S_Members[$S_Key] = [PSCustomObject]@{
                Key                  = $S_Key
                DisplayName          = $S_Disp
                PrimarySmtpAddress   = $S_Smtp
                Alias                = "$(Get-PropertyOrNull -Object $S_Use -Name 'Alias')".Trim()
                RecipientTypeDetails = if ($S_Rtd) { $S_Rtd } else { 'Unknown' }
                Hidden               = $S_Hidden
                IsDirSynced          = $S_Sync
                ChangePath           = $S_Path
                ChangePathLabel      = Get-ChangePathLabel $S_Path
                # Entra ID object id, which is the Graph user id for a guest.
                GraphUserId          = $S_Eid
                Resolved             = [bool]$S_Rec
                # Identity for the Set-* call. Prefer the immutable directory id, then SMTP.
                SetIdentity          = $(
                    $S_Id = $S_Eid
                    if (-not $S_Id) { $S_Id = $S_Smtp }
                    if (-not $S_Id) { $S_Id = "$(Get-PropertyOrNull -Object $S_Use -Name 'Identity')".Trim() }
                    $S_Id
                )
                Groups               = [System.Collections.Generic.List[string]]::new()
                Action               = 'Not attempted'
                ActionDetail         = $null
            }
        }

        $S_Entry = $S_Members[$S_Key]
        if (-not $S_Entry.Groups.Contains($S_GroupLabel)) { $S_Entry.Groups.Add($S_GroupLabel) }
        $S_MemberKeys.Add($S_Key)
        if ($S_Entry.Hidden -eq $true) { $S_HiddenInThis++ }
    }

    $S_GroupReport.Add([PSCustomObject]@{
        DisplayName          = $S_GroupLabel
        Name                 = $S_G.Name
        PrimarySmtpAddress   = $S_G.PrimarySmtpAddress
        Alias                = $S_G.Alias
        Kind                 = $S_G.Kind
        RecipientTypeDetails = $S_G.RecipientTypeDetails
        TypeLabel            = Get-GroupTypeLabel $S_G.RecipientTypeDetails
        GroupHidden          = $S_G.Hidden
        GroupIsDirSynced     = $S_G.IsDirSynced
        Identity             = $S_G.Identity
        MemberKeys           = $S_MemberKeys
        MemberCount          = $S_MemberKeys.Count
        HiddenMemberCount    = $S_HiddenInThis
        Error                = $S_GroupError
        GroupAction          = 'Not attempted'
        GroupActionDetail    = $null
    })

    if ($DelayMilliseconds -gt 0) { Start-Sleep -Milliseconds $DelayMilliseconds }
}
Write-Progress -Activity "Reading distribution list membership" -Completed

$S_HiddenMembers = @($S_Members.Values | Where-Object { $_.Hidden -eq $true } | Sort-Object DisplayName, PrimarySmtpAddress)
$S_HiddenGroups  = @($S_GroupReport | Where-Object { $_.GroupHidden -eq $true })

# ---------------------------------------------------------------------------
# Remediation
# ---------------------------------------------------------------------------
$S_ChangesMade = $false

if ($Unhide -or $UnhideGroups)
{
    # Partition before touching anything, so the operator sees the real work list - including
    # what will be skipped and why - before agreeing to it.
    $S_ToChange = [System.Collections.Generic.List[PSCustomObject]]::new()

    # ---- Member selection -------------------------------------------------
    # Narrows which hidden members get changed. The report still covers everything in scope;
    # deselected members are recorded as 'Not selected' further down rather than dropped, so the
    # report still answers "why is this one still hidden".
    $S_Selected    = [System.Collections.Generic.HashSet[string]]::new()
    $S_PickedByHand = $false
    $S_Unmatched    = [System.Collections.Generic.List[string]]::new()

    if ($Unhide)
    {
        if ($SelectMembers)
        {
            $S_Picked = Select-HiddenMembers -HiddenMembers $S_HiddenMembers
            if ($null -eq $S_Picked)
            {
                Write-Host "Selection cancelled: no members will be changed." -ForegroundColor Yellow
            }
            else
            {
                foreach ($S_One in @($S_Picked)) { [void]$S_Selected.Add($S_One.Key) }
            }
            $S_PickedByHand = $true
        }
        elseif ($Member -or $ExcludeMember)
        {
            $S_Pool = @($S_HiddenMembers)

            if ($Member)
            {
                $S_Kept = [System.Collections.Generic.List[PSCustomObject]]::new()
                foreach ($S_Tok in $Member)
                {
                    $S_Hits = @($S_Pool | Where-Object { Test-MemberToken -Member $_ -Token $S_Tok })
                    if ($S_Hits.Count -eq 0)
                    {
                        Write-Warning "-Member '$S_Tok' matched no hidden member in scope."
                        $S_Unmatched.Add($S_Tok)
                    }
                    foreach ($S_Hit in $S_Hits) { if (-not $S_Kept.Contains($S_Hit)) { $S_Kept.Add($S_Hit) } }
                }
                $S_Pool = @($S_Kept)
            }

            if ($ExcludeMember)
            {
                foreach ($S_Tok in $ExcludeMember)
                {
                    $S_Hits = @($S_Pool | Where-Object { Test-MemberToken -Member $_ -Token $S_Tok })
                    if ($S_Hits.Count -eq 0)
                    {
                        Write-Warning "-ExcludeMember '$S_Tok' matched no hidden member in scope."
                        $S_Unmatched.Add($S_Tok)
                    }
                }
                $S_Pool = @($S_Pool | Where-Object {
                    $S_Keep = $true
                    foreach ($S_Tok in $ExcludeMember)
                    {
                        if (Test-MemberToken -Member $_ -Token $S_Tok) { $S_Keep = $false; break }
                    }
                    $S_Keep
                })
            }

            foreach ($S_One in $S_Pool) { [void]$S_Selected.Add($S_One.Key) }
        }
        else
        {
            foreach ($S_One in $S_HiddenMembers) { [void]$S_Selected.Add($S_One.Key) }
        }
    }

    $S_Targeted = $SelectMembers -or [bool]$Member -or [bool]$ExcludeMember

    if ($Unhide)
    {
        foreach ($S_Mem in $S_HiddenMembers)
        {
            if (-not $S_Selected.Contains($S_Mem.Key))
            {
                $S_Mem.Action = 'Not selected'
                $S_Mem.ActionDetail = 'Left hidden on purpose: this member was not included by the selection for this run.'
                continue
            }

            switch ($S_Mem.ChangePath)
            {
                'Graph'
                {
                    if ($SkipGraph)
                    {
                        $S_Mem.Action = 'Needs Microsoft Graph'
                        $S_Mem.ActionDetail = 'Entra ID guest user, skipped because -SkipGraph was specified. Unhide it with: Update-MgUser -UserId <entra object id> -ShowInAddressList:$true'
                    }
                    elseif (-not $S_Mem.GraphUserId)
                    {
                        $S_Mem.Action = 'Skipped'
                        $S_Mem.ActionDetail = 'Entra ID guest user with no ExternalDirectoryObjectId, so there is no Graph user id to target.'
                    }
                    else
                    {
                        $S_ToChange.Add([PSCustomObject]@{ Kind = 'Member'; Target = $S_Mem; Cmdlet = 'Update-MgUser'; Path = 'Graph' })
                    }
                }

                'OnPremises'
                {
                    if (-not $IncludeDirSynced)
                    {
                        $S_Mem.Action = 'Needs on-premises change'
                        $S_Mem.ActionDetail = 'Directory-synchronized object: Exchange Online refuses the write. Clear msExchHideFromAddressLists on-premises, or re-run with -IncludeDirSynced if Exchange-attribute Source of Authority has been moved to the cloud.'
                    }
                    else
                    {
                        $S_Cmdlet = Get-VisibilityCmdlet -RecipientTypeDetails $S_Mem.RecipientTypeDetails
                        if ($S_Cmdlet -and $S_Mem.SetIdentity)
                        {
                            $S_ToChange.Add([PSCustomObject]@{ Kind = 'Member'; Target = $S_Mem; Cmdlet = $S_Cmdlet; Path = 'ExchangeOnline' })
                        }
                        else
                        {
                            $S_Mem.Action = 'Skipped'
                            $S_Mem.ActionDetail = "No Set-* cmdlet covers recipient type '$($S_Mem.RecipientTypeDetails)'."
                        }
                    }
                }

                'ExchangeOnline'
                {
                    $S_Cmdlet = Get-VisibilityCmdlet -RecipientTypeDetails $S_Mem.RecipientTypeDetails
                    if (-not $S_Cmdlet)
                    {
                        $S_Mem.Action = 'Skipped'
                        $S_Mem.ActionDetail = "No Set-* cmdlet covers recipient type '$($S_Mem.RecipientTypeDetails)'."
                    }
                    elseif (-not $S_Mem.SetIdentity)
                    {
                        $S_Mem.Action = 'Skipped'
                        $S_Mem.ActionDetail = 'No usable identity for the change.'
                    }
                    else
                    {
                        $S_ToChange.Add([PSCustomObject]@{ Kind = 'Member'; Target = $S_Mem; Cmdlet = $S_Cmdlet; Path = 'ExchangeOnline' })
                    }
                }

                default
                {
                    $S_Mem.Action = 'Skipped'
                    $S_Mem.ActionDetail = "No supported route covers recipient type '$($S_Mem.RecipientTypeDetails)'."
                }
            }
        }
    }

    if ($UnhideGroups)
    {
        foreach ($S_Grp in $S_HiddenGroups)
        {
            $S_Cmdlet = Get-VisibilityCmdlet -RecipientTypeDetails $S_Grp.RecipientTypeDetails
            if (-not $S_Cmdlet)
            {
                $S_Grp.GroupAction = 'Skipped'
                $S_Grp.GroupActionDetail = "No Set-* cmdlet covers recipient type '$($S_Grp.RecipientTypeDetails)'."
                continue
            }
            if ($S_Grp.GroupIsDirSynced -eq $true -and -not $IncludeDirSynced)
            {
                $S_Grp.GroupAction = 'Needs on-premises change'
                $S_Grp.GroupActionDetail = 'Directory-synchronized group: clear msExchHideFromAddressLists on-premises.'
                continue
            }
            $S_ToChange.Add([PSCustomObject]@{ Kind = 'Group'; Target = $S_Grp; Cmdlet = $S_Cmdlet; Path = 'ExchangeOnline' })
        }
    }

    $S_MemberChanges = @($S_ToChange | Where-Object { $_.Kind -eq 'Member' }).Count
    $S_GroupChanges  = @($S_ToChange | Where-Object { $_.Kind -eq 'Group' }).Count
    $S_GraphChanges  = @($S_ToChange | Where-Object { $_.Kind -eq 'Member' -and $_.Path -eq 'Graph' }).Count
    $S_ExoChanges    = $S_MemberChanges - $S_GraphChanges
    $S_SkippedSync   = @($S_HiddenMembers | Where-Object { $_.Action -eq 'Needs on-premises change' }).Count
    $S_SkippedGraph  = @($S_HiddenMembers | Where-Object { $_.Action -eq 'Needs Microsoft Graph' }).Count

    Write-Host ""
    Write-Host "Remediation plan:" -ForegroundColor Cyan
    Write-Host "  Hidden members found         : $($S_HiddenMembers.Count)"
    if ($S_Targeted)
    {
        $S_Deselected = @($S_HiddenMembers | Where-Object { $_.Action -eq 'Not selected' }).Count
        Write-Host "  Selected for change          : $($S_Selected.Count)" -ForegroundColor Cyan
        Write-Host "  Left hidden (not selected)   : $S_Deselected" -ForegroundColor Yellow
    }
    Write-Host "  Members to unhide            : $S_MemberChanges"
    Write-Host "    via Exchange Online        : $S_ExoChanges"
    Write-Host "    via Microsoft Graph (guest): $S_GraphChanges"
    if ($UnhideGroups) { Write-Host "  Hidden groups to unhide      : $S_GroupChanges" }
    if ($S_SkippedSync -gt 0)
    {
        Write-Host "  Skipped (directory-synced)   : $S_SkippedSync" -ForegroundColor Yellow
        Write-Host "    These need an on-premises change; they are listed in the report." -ForegroundColor Yellow
    }
    if ($S_SkippedGraph -gt 0)
    {
        Write-Host "  Skipped (guest, -SkipGraph)  : $S_SkippedGraph" -ForegroundColor Yellow
    }
    Write-Host ""

    $S_Proceed = $true
    if ($S_ToChange.Count -eq 0)
    {
        Write-Host "Nothing to change." -ForegroundColor Green
        $S_Proceed = $false
    }
    elseif (-not $Force -and -not $WhatIfPreference -and -not $S_PickedByHand)
    {
        $S_Answer = Read-Host "Set HiddenFromAddressListsEnabled = False on $($S_ToChange.Count) object(s)? [Y] Yes  [N] No  (Default: N)"
        if ($S_Answer -ne 'Y')
        {
            Write-Host "Cancelled. No changes made; the report still describes what was found." -ForegroundColor Yellow
            $S_Proceed = $false
        }
    }

    if ($S_Proceed)
    {
        $S_Done = 0
        foreach ($S_Change in $S_ToChange)
        {
            $S_Done++
            $S_Obj = $S_Change.Target
            $S_Label = if ($S_Change.Kind -eq 'Member')
            {
                if ($S_Obj.PrimarySmtpAddress) { "$($S_Obj.DisplayName) <$($S_Obj.PrimarySmtpAddress)>" } else { "$($S_Obj.DisplayName)" }
            }
            else
            {
                "$($S_Obj.DisplayName) <$($S_Obj.PrimarySmtpAddress)>"
            }
            $S_Id = if ($S_Change.Kind -eq 'Member') { $S_Obj.SetIdentity } else { $S_Obj.Identity }

            Write-Progress -Activity "Unhiding from address lists" -Status $S_Label -PercentComplete (($S_Done / $S_ToChange.Count) * 100)

            $S_Would = if ($S_Change.Path -eq 'Graph')
            {
                "Update-MgUser -UserId '$($S_Obj.GraphUserId)' -ShowInAddressList:`$true"
            }
            else
            {
                "$($S_Change.Cmdlet) -Identity '$S_Id' -HiddenFromAddressListsEnabled `$false"
            }

            if (-not $PSCmdlet.ShouldProcess($S_Label, $S_Would))
            {
                if ($S_Change.Kind -eq 'Member')
                {
                    $S_Obj.Action = 'WhatIf'
                    $S_Obj.ActionDetail = "Would run: $S_Would"
                }
                else
                {
                    $S_Obj.GroupAction = 'WhatIf'
                    $S_Obj.GroupActionDetail = "Would run: $S_Would"
                }
                continue
            }

            try
            {
                if ($S_Change.Path -eq 'Graph')
                {
                    # Entra ID guest. showInAddressList on the Entra object is the setting that
                    # matters; Exchange keeps its own flag, and where the two disagree Exchange
                    # wins, so the Exchange flag is cleared afterwards rather than left to
                    # contradict the directory.
                    if (-not (Initialize-GraphConnection))
                    {
                        $S_Obj.Action = 'Needs Microsoft Graph'
                        $S_Obj.ActionDetail = 'Microsoft Graph was not available or not usable for this tenant. Unhide with: Update-MgUser -UserId <entra object id> -ShowInAddressList:$true'
                        Write-Warning "  [$S_Done/$($S_ToChange.Count)] Skipped (no Graph): $S_Label"
                        continue
                    }

                    Update-MgUser -UserId $S_Obj.GraphUserId -ShowInAddressList:$true -ErrorAction Stop
                    $S_ChangesMade = $true

                    $S_ExoNote = ''
                    if ($S_Obj.SetIdentity)
                    {
                        try
                        {
                            Set-MailUser -Identity $S_Obj.SetIdentity -HiddenFromAddressListsEnabled $false -ErrorAction Stop
                            $S_ExoNote = ' Exchange Online flag cleared as well, so the two settings agree.'
                        }
                        catch
                        {
                            $S_ExoNote = " Exchange Online flag could not be cleared: $($_.Exception.Message)"
                        }
                    }

                    $S_Obj.Action = 'Unhidden'
                    $S_Obj.ActionDetail = "Update-MgUser set showInAddressList = true on the Entra ID guest.$S_ExoNote"
                    $S_Obj.Hidden = $false
                    Write-Host "  [$S_Done/$($S_ToChange.Count)] Unhidden (guest, via Graph): $S_Label" -ForegroundColor Green
                    continue
                }

                & $S_Change.Cmdlet -Identity $S_Id -HiddenFromAddressListsEnabled $false -ErrorAction Stop
                $S_ChangesMade = $true
                if ($S_Change.Kind -eq 'Member')
                {
                    $S_Obj.Action = 'Unhidden'
                    $S_Obj.ActionDetail = "$($S_Change.Cmdlet) succeeded."
                    $S_Obj.Hidden = $false
                }
                else
                {
                    $S_Obj.GroupAction = 'Unhidden'
                    $S_Obj.GroupActionDetail = "$($S_Change.Cmdlet) succeeded."
                    $S_Obj.GroupHidden = $false
                }
                Write-Host "  [$S_Done/$($S_ToChange.Count)] Unhidden: $S_Label" -ForegroundColor Green
            }
            catch
            {
                if ($S_Change.Kind -eq 'Member')
                {
                    $S_Obj.Action = 'Failed'
                    $S_Obj.ActionDetail = $_.Exception.Message
                }
                else
                {
                    $S_Obj.GroupAction = 'Failed'
                    $S_Obj.GroupActionDetail = $_.Exception.Message
                }
                Write-Warning "  [$S_Done/$($S_ToChange.Count)] Failed: $S_Label - $($_.Exception.Message)"
            }
        }
        Write-Progress -Activity "Unhiding from address lists" -Completed
    }
}

# ---------------------------------------------------------------------------
# Stats
# ---------------------------------------------------------------------------
# Recomputed after remediation so the report reflects the tenant as it now stands, with the
# Action columns carrying the history of how it got there.
$S_TotalGroups       = $S_GroupReport.Count
$S_TotalMemberships  = ($S_GroupReport | Measure-Object -Property MemberCount -Sum).Sum
if (-not $S_TotalMemberships) { $S_TotalMemberships = 0 }
$S_UniqueMembers     = $S_Members.Count
$S_HiddenNow         = @($S_Members.Values | Where-Object { $_.Hidden -eq $true })
$S_HiddenNowCount    = $S_HiddenNow.Count
$S_VisibleCount      = @($S_Members.Values | Where-Object { $_.Hidden -eq $false }).Count
$S_UnknownCount      = @($S_Members.Values | Where-Object { $null -eq $_.Hidden }).Count
$S_DirSyncedHidden   = @($S_HiddenNow | Where-Object { $_.IsDirSynced -eq $true }).Count
$S_UnhiddenCount     = @($S_Members.Values | Where-Object { $_.Action -eq 'Unhidden' }).Count
$S_FailedCount       = @($S_Members.Values | Where-Object { $_.Action -eq 'Failed' }).Count
$S_OnPremCount       = @($S_Members.Values | Where-Object { $_.Action -eq 'Needs on-premises change' }).Count
$S_GraphNeededCount  = @($S_Members.Values | Where-Object { $_.Action -eq 'Needs Microsoft Graph' }).Count
$S_NotSelectedCount  = @($S_Members.Values | Where-Object { $_.Action -eq 'Not selected' }).Count
$S_GuestHiddenCount  = @($S_HiddenNow | Where-Object { $_.ChangePath -eq 'Graph' }).Count
$S_HiddenGroupsNow   = @($S_GroupReport | Where-Object { $_.GroupHidden -eq $true }).Count
$S_GroupErrors       = @($S_GroupReport | Where-Object { $_.Error }).Count

# Per-list hidden counts have to be recomputed rather than decremented: one member sits in many
# lists, so unhiding it moves the count of every list it belongs to.
foreach ($S_Grp in $S_GroupReport)
{
    $S_H = 0
    foreach ($S_Key in $S_Grp.MemberKeys)
    {
        if ($S_Members.ContainsKey($S_Key) -and $S_Members[$S_Key].Hidden -eq $true) { $S_H++ }
    }
    $S_Grp.HiddenMemberCount = $S_H
}
$S_GroupsWithHidden = @($S_GroupReport | Where-Object { $_.HiddenMemberCount -gt 0 }).Count

# Hidden members by recipient type, for the chart.
$S_TypeGroups = @($S_HiddenNow | Group-Object RecipientTypeDetails | Sort-Object Count -Descending)
$S_ChartLabels = @($S_TypeGroups | ForEach-Object { $_.Name })
$S_ChartValues = @($S_TypeGroups | ForEach-Object { [int]$_.Count })

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
$S_FileBase  = if ($TestMode) { "ReportDistributionListMemberVisibility_TEST_$S_Timestamp" } else { "ReportDistributionListMemberVisibility_$S_Timestamp" }

if (Test-Path $ReportPath -PathType Container)
{
    $S_CsvFile   = Join-Path $ReportPath "$S_FileBase.csv"
    $S_HtmlFile  = Join-Path $ReportPath "$S_FileBase.html"
    $S_GuestFile = Join-Path $ReportPath "$($S_FileBase)_GuestRemediation.ps1"
}
else
{
    $S_CsvFile   = $ReportPath
    $S_HtmlFile  = [System.IO.Path]::ChangeExtension($ReportPath, '.html')
    $S_GuestFile = [System.IO.Path]::ChangeExtension($ReportPath, '.GuestRemediation.ps1')
}

# ---------------------------------------------------------------------------
# Companion guest remediation script
# ---------------------------------------------------------------------------
# Written whenever hidden guests are left unchanged - because -Unhide was not asked for, because
# -SkipGraph was, or because the Graph connection failed on the MSAL clash. The generated script
# does the Graph work BEFORE touching Exchange, which is the ordering that survives that clash, so
# a fresh session can finish the job without anyone having to reconstruct the command list.
$S_GuestTodo = @($S_Members.Values |
    Where-Object { $_.ChangePath -eq 'Graph' -and $_.Hidden -eq $true -and $_.GraphUserId } |
    Sort-Object DisplayName)
$S_GuestScriptWritten = $false

if ($S_GuestTodo.Count -gt 0 -and $Format -in @('Csv', 'Html', 'Both'))
{
    $S_GuestRows = (@($S_GuestTodo) | ForEach-Object {
        $S_D = "$($_.DisplayName)" -replace "'", "''"
        $S_A = "$($_.PrimarySmtpAddress)" -replace "'", "''"
        $S_I = "$($_.GraphUserId)" -replace "'", "''"
        "    [PSCustomObject]@{ Id = '$S_I'; Display = '$S_D'; Address = '$S_A' }"
    }) -join "`n"

    $S_GuestScript = @"
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users, ExchangeOnlineManagement

<#
.SYNOPSIS
    Unhides $($S_GuestTodo.Count) Entra ID guest user(s) from the address lists.

.DESCRIPTION
    Generated by ReportDistributionListMemberVisibility.ps1 on $(Get-Date -Format 'dd MMM yyyy HH:mm')
    for tenant $S_TenantDisplayName ($S_TenantId).

    RUN THIS IN A FRESH POWERSHELL SESSION, with no Exchange Online connection already open.

    A guest's address-list visibility is the showInAddressList property on the Entra ID user
    object, not an Exchange setting, so it is changed through Microsoft Graph. This script
    authenticates to Graph BEFORE connecting to Exchange Online, deliberately:
    ExchangeOnlineManagement and Microsoft.Graph.Authentication ship different, incompatible
    Microsoft.Identity.Client / Microsoft.IdentityModel.Abstractions versions, and whichever
    authenticates first in a session wins. Connect to Exchange first and the Graph call fails
    with "Method not found: ...WithLogging(IIdentityLogger, Boolean)".

    The Exchange pass at the end clears HiddenFromAddressListsEnabled on the same guests.
    Updating showInAddressList already updates the Exchange setting, but the two are stored
    separately and the Exchange one wins where they disagree, so both are set.

.NOTES
    Required: User Administrator (or the User.ReadWrite.All Graph scope), and Mail Recipients
    in Exchange Online. Supports -WhatIf.
#>

[CmdletBinding(SupportsShouldProcess = `$true)]
param()

`$ErrorActionPreference = 'Stop'

`$S_Guests = @(
$S_GuestRows
)

Write-Host "Unhiding `$(`$S_Guests.Count) guest user(s) from the address lists." -ForegroundColor Cyan

# --- 1. Microsoft Graph first. The order matters; see .DESCRIPTION. ---
Write-Host ""
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes 'User.ReadWrite.All' -NoWelcome

`$S_Ok     = [System.Collections.Generic.List[string]]::new()
`$S_Failed = [System.Collections.Generic.List[string]]::new()

foreach (`$S_G in `$S_Guests)
{
    `$S_Label = "`$(`$S_G.Display) <`$(`$S_G.Address)>"
    if (-not `$PSCmdlet.ShouldProcess(`$S_Label, "Update-MgUser -ShowInAddressList:`$true")) { continue }
    try
    {
        Update-MgUser -UserId `$S_G.Id -ShowInAddressList:`$true -ErrorAction Stop
        Write-Host "  Graph OK : `$S_Label" -ForegroundColor Green
        `$S_Ok.Add(`$S_G.Id)
    }
    catch
    {
        Write-Warning "  Graph FAILED : `$S_Label - `$(`$_.Exception.Message)"
        `$S_Failed.Add(`$S_Label)
    }
}

# --- 2. Then Exchange Online, to clear its own flag on the same guests. ---
if (`$S_Ok.Count -gt 0)
{
    Write-Host ""
    Write-Host "Connecting to Exchange Online to clear the Exchange flag..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:`$false

    foreach (`$S_G in `$S_Guests)
    {
        if (-not `$S_Ok.Contains(`$S_G.Id)) { continue }
        `$S_Label = "`$(`$S_G.Display) <`$(`$S_G.Address)>"
        if (-not `$PSCmdlet.ShouldProcess(`$S_Label, "Set-MailUser -HiddenFromAddressListsEnabled `$false")) { continue }
        try
        {
            Set-MailUser -Identity `$S_G.Id -HiddenFromAddressListsEnabled `$false -ErrorAction Stop
            Write-Host "  Exchange OK : `$S_Label" -ForegroundColor Green
        }
        catch
        {
            Write-Warning "  Exchange flag not cleared : `$S_Label - `$(`$_.Exception.Message)"
        }
    }
}

Write-Host ""
Write-Host "Done. Unhidden via Graph: `$(`$S_Ok.Count). Failed: `$(`$S_Failed.Count)." -ForegroundColor Cyan
if (`$S_Failed.Count -gt 0) { `$S_Failed | ForEach-Object { Write-Host "  `$_" -ForegroundColor Yellow } }
"@

    $S_GuestScript | Out-File -FilePath $S_GuestFile -Encoding UTF8
    $S_GuestScriptWritten = $true
}

# ---------------------------------------------------------------------------
# CSV export - one row per group-and-member pair
# ---------------------------------------------------------------------------
if ($Format -in @('Csv', 'Both'))
{
    $S_CsvRows = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($S_Grp in $S_GroupReport)
    {
        if ($S_Grp.MemberKeys.Count -eq 0)
        {
            $S_CsvRows.Add([PSCustomObject]@{
                GroupDisplayName     = $S_Grp.DisplayName
                GroupPrimarySmtp     = $S_Grp.PrimarySmtpAddress
                GroupType            = $S_Grp.TypeLabel
                GroupHidden          = $S_Grp.GroupHidden
                GroupIsDirSynced     = $S_Grp.GroupIsDirSynced
                GroupMemberCount     = 0
                GroupAction          = $S_Grp.GroupAction
                MemberDisplayName    = $null
                MemberPrimarySmtp    = $null
                MemberAlias          = $null
                MemberRecipientType  = $null
                MemberHidden         = $null
                MemberIsDirSynced    = $null
                MemberChangeVia      = $null
                MemberAction         = $null
                MemberActionDetail   = $null
                GroupError           = $S_Grp.Error
            })
            continue
        }

        foreach ($S_Key in $S_Grp.MemberKeys)
        {
            if (-not $S_Members.ContainsKey($S_Key)) { continue }
            $S_Mem = $S_Members[$S_Key]
            $S_CsvRows.Add([PSCustomObject]@{
                GroupDisplayName     = $S_Grp.DisplayName
                GroupPrimarySmtp     = $S_Grp.PrimarySmtpAddress
                GroupType            = $S_Grp.TypeLabel
                GroupHidden          = $S_Grp.GroupHidden
                GroupIsDirSynced     = $S_Grp.GroupIsDirSynced
                GroupMemberCount     = $S_Grp.MemberCount
                GroupAction          = $S_Grp.GroupAction
                MemberDisplayName    = $S_Mem.DisplayName
                MemberPrimarySmtp    = $S_Mem.PrimarySmtpAddress
                MemberAlias          = $S_Mem.Alias
                MemberRecipientType  = $S_Mem.RecipientTypeDetails
                MemberHidden         = $S_Mem.Hidden
                MemberIsDirSynced    = $S_Mem.IsDirSynced
                MemberChangeVia      = $S_Mem.ChangePathLabel
                MemberAction         = $S_Mem.Action
                MemberActionDetail   = $S_Mem.ActionDetail
                GroupError           = $S_Grp.Error
            })
        }
    }

    $S_CsvRows | Sort-Object GroupDisplayName, MemberDisplayName |
        Export-Csv -Path $S_CsvFile -NoTypeInformation -Encoding UTF8
}

# ---------------------------------------------------------------------------
# HTML report
# ---------------------------------------------------------------------------
if ($Format -in @('Html', 'Both'))
{

$S_ReportDate = Get-Date -Format 'dd MMM yyyy HH:mm'

# Tri-state visibility badge: hidden, visible, or genuinely unknown.
function Format-HiddenBadge
{
    param($Hidden)

    if ($Hidden -eq $true)  { return '<span class="badge badge-hidden">Hidden</span>' }
    if ($Hidden -eq $false) { return '<span class="badge badge-visible">Visible</span>' }
    return '<span class="badge badge-unknown">Unknown</span>'
}

function Format-ActionBadge
{
    param([string]$Action, [string]$Detail)

    if ([string]::IsNullOrWhiteSpace($Action) -or $Action -eq 'Not attempted') { return '<span class="muted">-</span>' }
    $S_Cls = switch ($Action)
    {
        'Unhidden'                 { 'ok' }
        'Failed'                   { 'bad' }
        'Needs on-premises change' { 'warn' }
        'Needs Microsoft Graph'    { 'warn' }
        'Not selected'             { 'neutral' }
        'WhatIf'                   { 'info' }
        default                    { 'neutral' }
    }
    $S_Title = if ($Detail) { " title=`"$(Format-Html $Detail)`"" } else { '' }
    return "<span class=`"badge badge-$S_Cls`"$S_Title>$(Format-Html $Action)</span>"
}

# Which mechanism can clear the flag. This is the column an operator plans the work from, so it
# replaces a plain cloud/synced badge: the sync state only matters because of what it implies here.
function Format-PathBadge
{
    param([string]$ChangePath, [string]$Label)

    $S_Cls = switch ($ChangePath)
    {
        'Graph'          { 'info' }
        'OnPremises'     { 'warn' }
        'ExchangeOnline' { 'neutral' }
        default          { 'bad' }
    }
    $S_Title = switch ($ChangePath)
    {
        'Graph'          { 'Entra ID guest user: unhide with Update-MgUser -ShowInAddressList, not Set-MailUser.' }
        'OnPremises'     { 'Directory-synchronized: clear msExchHideFromAddressLists in on-premises Active Directory.' }
        'ExchangeOnline' { 'Unhide with the Set-* cmdlet for this recipient type.' }
        default          { 'No supported route for this recipient type.' }
    }
    return "<span class=`"badge badge-$S_Cls`" title=`"$(Format-Html $S_Title)`">$(Format-Html $Label)</span>"
}

$S_Banners = [System.Collections.Generic.List[string]]::new()
if ($TestMode)
{
    $S_Banners.Add("<div class=`"notice`"><strong>Test mode.</strong> Generated from a random sample of $($S_GroupsInScope.Count) groups. Run without <code>-TestMode</code> to cover every group in scope.</div>")
}
if (-not ($Unhide -or $UnhideGroups))
{
    $S_Banners.Add('<div class="notice notice-info"><strong>Read-only run.</strong> Nothing in the tenant was changed. Re-run with <code>-Unhide</code> to clear <code>HiddenFromAddressListsEnabled</code> on the hidden members listed below, or with <code>-Unhide -WhatIf</code> to preview the change first.</div>')
}
elseif ($WhatIfPreference)
{
    $S_Banners.Add('<div class="notice notice-info"><strong>WhatIf run.</strong> The changes below were previewed only. Nothing in the tenant was changed.</div>')
}
elseif ($S_ChangesMade)
{
    $S_Banners.Add("<div class=`"notice notice-ok`"><strong>Changes applied.</strong> $S_UnhiddenCount object(s) had <code>HiddenFromAddressListsEnabled</code> set to <code>False</code> during this run. The tables below show the state after those changes; the Action column records what was done.</div>")
}
if ($S_GuestHiddenCount -gt 0 -and -not ($Unhide -or $UnhideGroups))
{
    $S_Banners.Add("<div class=`"notice notice-info`"><strong>$S_GuestHiddenCount hidden member(s) are Entra ID guest users.</strong> A guest's address-list visibility is the <code>showInAddressList</code> property on the Entra ID user object rather than an Exchange setting, so <code>Set-MailUser</code> is the wrong tool for them - the change runs as <code>Update-MgUser -ShowInAddressList</code> through Microsoft Graph. Guests invited through the invitation manager start out with it set to <code>false</code>, so they appear hidden without anyone having hidden them. This script takes that route automatically; filter the member table by <em>Microsoft Graph (guest)</em> to see them.</div>")
}
if ($S_NotSelectedCount -gt 0)
{
    $S_Banners.Add("<div class=`"notice notice-info`"><strong>Targeted run.</strong> $S_UnhiddenCount member(s) were unhidden and <strong>$S_NotSelectedCount were deliberately left hidden</strong>, having not been included by the member selection for this run. They carry a <em>Not selected</em> action below, so this report is not a record of everything that needed doing - only of what was asked for.</div>")
}
if ($S_GuestScriptWritten)
{
    $S_Banners.Add("<div class=`"notice notice-info`"><strong>A companion script was written for the $($S_GuestTodo.Count) hidden guest user(s).</strong> <code>$(Format-Html (Split-Path -Leaf $S_GuestFile))</code>, alongside this report, unhides them through Microsoft Graph. <strong>Run it in a fresh PowerShell session</strong> with no Exchange Online connection open: the Exchange and Graph modules ship incompatible MSAL versions, and whichever authenticates first in a session wins, so the generated script does the Graph work before touching Exchange.</div>")
}
if ($S_GraphNeededCount -gt 0)
{
    $S_Banners.Add("<div class=`"notice`"><strong>$S_GraphNeededCount guest user(s) were not changed.</strong> Microsoft Graph was skipped or unavailable. Unhide each with <code>Update-MgUser -UserId &lt;entra object id&gt; -ShowInAddressList:`$true</code>, or re-run without <code>-SkipGraph</code>.</div>")
}
if ($S_OnPremCount -gt 0)
{
    $S_Banners.Add("<div class=`"notice`"><strong>$S_OnPremCount member(s) need an on-premises change.</strong> They are directory-synchronized, so Exchange Online will not accept the write. Clear <code>msExchHideFromAddressLists</code> in Active Directory (and confirm the object's alias / <code>mailNickname</code> is populated, or the change will not synchronize). Filter the member table by <em>On-premises AD</em> for the work list.</div>")
}
if ($S_FailedCount -gt 0)
{
    $S_Banners.Add("<div class=`"notice notice-bad`"><strong>$S_FailedCount change(s) failed.</strong> Hover the Failed badge in the member table for the error returned by Exchange Online.</div>")
}
if ($S_NotFound.Count -gt 0)
{
    $S_Banners.Add("<div class=`"notice`"><strong>Not found:</strong> $(Format-Html ($S_NotFound -join ', ')). These were requested with <code>-Group</code> but do not exist in the selected group types.</div>")
}
if ($S_GroupErrors -gt 0)
{
    $S_Banners.Add("<div class=`"notice`"><strong>$S_GroupErrors group(s) could not be read.</strong> Their membership is unknown and their rows show a count of zero. See the <code>GroupError</code> column in the CSV.</div>")
}
$S_BannerHtml = ($S_Banners -join "`n")

# --- Group table ---
$S_GroupRows = (@($S_GroupReport | Sort-Object @{ Expression = { $_.HiddenMemberCount }; Descending = $true }, DisplayName) | ForEach-Object {
    $S_Grp = $_
    $S_MemberList = if ($S_Grp.MemberKeys.Count -eq 0)
    {
        if ($S_Grp.Error) { "<span class=`"muted`">Could not be read: $(Format-Html $S_Grp.Error)</span>" }
        else { '<span class="muted">No members.</span>' }
    }
    else
    {
        $S_Items = @(foreach ($S_Key in $S_Grp.MemberKeys)
        {
            if (-not $S_Members.ContainsKey($S_Key)) { continue }
            $S_Mem = $S_Members[$S_Key]
            $S_Mark = if ($S_Mem.Hidden -eq $true) { ' class="mem-hidden"' } elseif ($null -eq $S_Mem.Hidden) { ' class="mem-unknown"' } else { '' }
            $S_Smtp = if ($S_Mem.PrimarySmtpAddress) { " &lt;$(Format-Html $S_Mem.PrimarySmtpAddress)&gt;" } else { '' }
            $S_Tag  = if ($S_Mem.Hidden -eq $true) { ' <span class="mini">hidden</span>' } elseif ($null -eq $S_Mem.Hidden) { ' <span class="mini">unknown</span>' } else { '' }
            "<li$S_Mark>$(Format-Html $S_Mem.DisplayName)<span class=`"mono`">$S_Smtp</span>$S_Tag</li>"
        })
        $S_Summary = "$($S_Grp.MemberCount) member$(if ($S_Grp.MemberCount -ne 1) { 's' })"
        if ($S_Grp.HiddenMemberCount -gt 0) { $S_Summary += ", $($S_Grp.HiddenMemberCount) hidden" }
        "<details><summary>$S_Summary</summary><ul class=`"memlist`">$($S_Items -join '')</ul></details>"
    }

    $S_HiddenCell = if ($S_Grp.HiddenMemberCount -gt 0) { "<strong class=`"num-bad`">$($S_Grp.HiddenMemberCount)</strong>" } else { '0' }
    $S_Data = "$(if ($S_Grp.HiddenMemberCount -gt 0) { 'hashidden' } else { 'nohidden' })"

    "<tr data-hidden=`"$S_Data`" data-type=`"$(Format-Html $S_Grp.TypeLabel)`"><td>$(Format-Html $S_Grp.DisplayName)</td><td class=`"mono`">$(Format-Html $S_Grp.PrimarySmtpAddress)</td><td>$(Format-Html $S_Grp.TypeLabel)</td><td>$(Format-HiddenBadge $S_Grp.GroupHidden)</td><td>$(Format-ActionBadge -Action $S_Grp.GroupAction -Detail $S_Grp.GroupActionDetail)</td><td class=`"num`">$($S_Grp.MemberCount)</td><td class=`"num`">$S_HiddenCell</td><td class=`"memcell`">$S_MemberList</td></tr>"
}) -join "`n"

# --- Member table (one row per unique member) ---
$S_MemberRows = (@($S_Members.Values | Sort-Object @{ Expression = { if ($_.Hidden -eq $true) { 0 } elseif ($null -eq $_.Hidden) { 1 } else { 2 } } }, DisplayName) | ForEach-Object {
    $S_Mem = $_
    $S_GroupCell = (@($S_Mem.Groups | Sort-Object) | ForEach-Object { Format-Html $_ }) -join ', '
    if (-not $S_GroupCell) { $S_GroupCell = '<span class="muted">-</span>' }
    $S_State = if ($S_Mem.Hidden -eq $true) { 'hidden' } elseif ($S_Mem.Hidden -eq $false) { 'visible' } else { 'unknown' }

    "<tr data-state=`"$S_State`" data-path=`"$($S_Mem.ChangePath)`" data-rtype=`"$(Format-Html $S_Mem.RecipientTypeDetails)`"><td>$(Format-Html $S_Mem.DisplayName)</td><td class=`"mono`">$(Format-Html $S_Mem.PrimarySmtpAddress)</td><td>$(Format-Html $S_Mem.RecipientTypeDetails)</td><td>$(Format-HiddenBadge $S_Mem.Hidden)</td><td>$(Format-PathBadge -ChangePath $S_Mem.ChangePath -Label $S_Mem.ChangePathLabel)</td><td>$(Format-ActionBadge -Action $S_Mem.Action -Detail $S_Mem.ActionDetail)</td><td class=`"num`">$($S_Mem.Groups.Count)</td><td class=`"grouplist`">$S_GroupCell</td></tr>"
}) -join "`n"

if (-not $S_GroupRows)  { $S_GroupRows  = '<tr><td colspan="8" class="muted">No groups found.</td></tr>' }
if (-not $S_MemberRows) { $S_MemberRows = '<tr><td colspan="8" class="muted">No members found.</td></tr>' }

# Recipient-type options for the member filter, built from what is actually present.
$S_RtypeOptions = (@($S_Members.Values | ForEach-Object { $_.RecipientTypeDetails } | Sort-Object -Unique) | ForEach-Object {
    "<option value=`"$(Format-Html $_)`">$(Format-Html $_)</option>"
}) -join ''

$S_TypeFilterOptions = (@($S_GroupReport | ForEach-Object { $_.TypeLabel } | Sort-Object -Unique) | ForEach-Object {
    "<option value=`"$(Format-Html $_)`">$(Format-Html $_)</option>"
}) -join ''

$S_Chart = New-BarChartSvg -Labels $S_ChartLabels -Values $S_ChartValues

$S_HiddenPct = if ($S_UniqueMembers -gt 0) { [Math]::Round(($S_HiddenNowCount / $S_UniqueMembers) * 100, 1) } else { 0 }

$S_ActionCard = if ($Unhide -or $UnhideGroups)
{
    "<div class=`"card`"><div class=`"label`">Unhidden this run</div><div class=`"value`">$S_UnhiddenCount</div><div class=`"sub`">$S_FailedCount failed, $S_OnPremCount need on-premises</div></div>"
}
else { '' }

$S_Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Distribution List Member Visibility</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: -apple-system, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; background: #f5f6f7; color: #24292f; padding: 28px; font-size: 15px; }
  .wrap { max-width: 1480px; margin: 0 auto; }

  .header { background: #fff; border: 1px solid #d8dbdf; border-radius: 8px; padding: 22px 26px; margin-bottom: 22px; }
  .header h1 { font-size: 1.35em; font-weight: 600; margin-bottom: 6px; }
  .header p { font-size: 0.86em; color: #57606a; }

  .section-title { font-size: 1.0em; font-weight: 600; margin: 0 0 14px 0; padding-bottom: 7px; border-bottom: 1px solid #d8dbdf; }

  .notice { background: #fff8e5; border: 1px solid #e4d08a; border-left-width: 4px; border-radius: 6px; padding: 11px 16px; margin-bottom: 12px; font-size: 0.86em; color: #4a3c10; }
  .notice-info { background: #eef4fb; border-color: #a8c3e0; color: #1f3b57; }
  .notice-ok   { background: #eef7f0; border-color: #a6ccae; color: #1f4428; }
  .notice-bad  { background: #fdf0ef; border-color: #e0aba6; color: #5c1f1a; }
  .notice code { background: rgba(0,0,0,0.05); padding: 1px 5px; border-radius: 3px; font-size: 0.95em; }

  .summary-cards { display: flex; gap: 14px; margin-bottom: 22px; flex-wrap: wrap; }
  .card { background: #fff; border: 1px solid #d8dbdf; border-radius: 8px; padding: 16px 20px; flex: 1; min-width: 150px; }
  .card .label { font-size: 0.74em; color: #57606a; text-transform: uppercase; letter-spacing: 0.4px; }
  .card .value { font-size: 1.8em; font-weight: 600; margin-top: 4px; }
  .card .sub { font-size: 0.74em; color: #8b949e; margin-top: 2px; }

  .panel { background: #fff; border: 1px solid #d8dbdf; border-radius: 8px; padding: 22px 24px; margin-bottom: 22px; }
  .panel h2 { font-size: 1.0em; font-weight: 600; margin-bottom: 14px; }
  .panel > p.note { font-size: 0.84em; color: #57606a; margin-bottom: 14px; }

  .barchart { display: block; max-width: 640px; }
  .bar { fill: #5a6b7d; }
  .bar-label { font-size: 12px; fill: #24292f; }
  .bar-value { font-size: 12px; fill: #57606a; }

  .controls { display: flex; gap: 10px; margin-bottom: 14px; flex-wrap: wrap; align-items: center; }
  .controls input[type="text"] { padding: 7px 11px; border: 1px solid #ced3d8; border-radius: 5px; font-size: 0.86em; min-width: 250px; font-family: inherit; }
  .controls select { padding: 7px 10px; border: 1px solid #ced3d8; border-radius: 5px; font-size: 0.86em; background: #fff; font-family: inherit; }
  .controls .count { font-size: 0.82em; color: #57606a; margin-left: auto; }

  .tablescroll { overflow-x: auto; }
  table { width: 100%; border-collapse: collapse; font-size: 0.84em; }
  th { background: #eff1f3; color: #24292f; padding: 9px 11px; text-align: left; border-bottom: 1px solid #d8dbdf; cursor: pointer; user-select: none; white-space: nowrap; font-weight: 600; }
  th:hover { background: #e4e7ea; }
  th.nosort { cursor: default; }
  th.nosort:hover { background: #eff1f3; }
  td { padding: 8px 11px; border-bottom: 1px solid #eceef0; vertical-align: top; }
  tr:hover td { background: #fafbfc; }
  tr.hidden-row { display: none; }
  td.num { text-align: right; font-variant-numeric: tabular-nums; }
  .num-bad { color: #9a2e26; }
  .mono { font-family: ui-monospace, 'Consolas', monospace; font-size: 0.94em; color: #57606a; }
  .muted { color: #8b949e; }
  .grouplist { max-width: 420px; white-space: normal; color: #444c56; }
  .memcell { max-width: 460px; }

  details summary { cursor: pointer; font-size: 0.95em; color: #3a5a7d; }
  ul.memlist { list-style: none; margin-top: 7px; display: flex; flex-direction: column; gap: 3px; }
  ul.memlist li { padding: 2px 0; }
  li.mem-hidden { color: #9a2e26; }
  li.mem-unknown { color: #8b6d1f; }
  .mini { font-size: 0.8em; text-transform: uppercase; letter-spacing: 0.4px; color: #9a2e26; }
  li.mem-unknown .mini { color: #8b6d1f; }

  .badge { padding: 2px 9px; border-radius: 10px; font-size: 0.82em; font-weight: 600; display: inline-block; white-space: nowrap; }
  .badge-hidden  { background: #fbe3e1; color: #8a271f; }
  .badge-visible { background: #e3f1e6; color: #1f4428; }
  .badge-unknown { background: #eceef0; color: #57606a; }
  .badge-ok      { background: #e3f1e6; color: #1f4428; }
  .badge-bad     { background: #fbe3e1; color: #8a271f; }
  .badge-warn    { background: #fdf3dc; color: #6b5110; }
  .badge-info    { background: #e7eff8; color: #1f3b57; }
  .badge-neutral { background: #eceef0; color: #57606a; }

  .footer { text-align: center; font-size: 0.78em; color: #8b949e; margin-top: 18px; }

  @media (max-width: 700px) {
    body { padding: 14px; }
    .card { min-width: 120px; }
  }
</style>
</head>
<body>
<div class="wrap">

<div class="header">
  <h1>Distribution List Member Visibility</h1>
  <p>Tenant: $(Format-Html $S_TenantDisplayName) ($(Format-Html $S_TenantId)) &nbsp;|&nbsp; Generated: $S_ReportDate &nbsp;|&nbsp; Scope: $(Format-Html ($S_Types -join ', '))$(if ($TestMode) { ' &nbsp;|&nbsp; Test mode' })$(if ($Unhide -or $UnhideGroups) { if ($WhatIfPreference) { ' &nbsp;|&nbsp; Unhide (WhatIf)' } else { ' &nbsp;|&nbsp; Unhide applied' } } else { ' &nbsp;|&nbsp; Read-only' })</p>
</div>

$S_BannerHtml

<div class="section-title">Overview</div>
<div class="summary-cards">
  <div class="card"><div class="label">Groups</div><div class="value">$S_TotalGroups</div><div class="sub">$S_GroupsWithHidden with hidden members</div></div>
  <div class="card"><div class="label">Memberships</div><div class="value">$S_TotalMemberships</div><div class="sub">group-and-member pairs</div></div>
  <div class="card"><div class="label">Unique members</div><div class="value">$S_UniqueMembers</div><div class="sub">$S_VisibleCount visible$(if ($S_UnknownCount -gt 0) { ", $S_UnknownCount unknown" })</div></div>
  <div class="card"><div class="label">Hidden members</div><div class="value" style="color:$(if ($S_HiddenNowCount -gt 0) { '#9a2e26' } else { '#1f4428' });">$S_HiddenNowCount</div><div class="sub">$S_HiddenPct% of unique members</div></div>
  <div class="card"><div class="label">Directory-synced hidden</div><div class="value" style="color:$(if ($S_DirSyncedHidden -gt 0) { '#6b5110' } else { '#1f4428' });">$S_DirSyncedHidden</div><div class="sub">change these on-premises</div></div>
  <div class="card"><div class="label">Hidden guests</div><div class="value" style="color:$(if ($S_GuestHiddenCount -gt 0) { '#1f3b57' } else { '#1f4428' });">$S_GuestHiddenCount</div><div class="sub">change via Microsoft Graph</div></div>
  <div class="card"><div class="label">Hidden groups</div><div class="value">$S_HiddenGroupsNow</div><div class="sub">the lists themselves</div></div>
  $S_ActionCard
</div>

<div class="panel">
  <h2>Hidden members by recipient type</h2>
  <p class="note">Unique members only, so a member in several lists is counted once. The recipient type determines which <code>Set-*</code> cmdlet can clear the flag, which is why it is worth seeing the split.</p>
  $S_Chart
</div>

<div class="panel">
  <h2>Distribution lists and their members</h2>
  <p class="note">Expand a row's member count to see the membership. Hidden members are marked in red.</p>
  <div class="controls">
    <input type="text" id="grpSearch" placeholder="Search list name, address or member..." onkeyup="filterGroups()" />
    <select id="grpHiddenFilter" onchange="filterGroups()">
      <option value="all">All lists</option>
      <option value="hashidden">With hidden members</option>
      <option value="nohidden">No hidden members</option>
    </select>
    <select id="grpTypeFilter" onchange="filterGroups()">
      <option value="all">All group types</option>
      $S_TypeFilterOptions
    </select>
    <span class="count" id="grpCount">Showing $S_TotalGroups of $S_TotalGroups lists</span>
  </div>
  <div class="tablescroll">
  <table id="grpTable">
    <thead>
      <tr>
        <th onclick="sortTable('grpTable', 0)">List</th>
        <th onclick="sortTable('grpTable', 1)">Address</th>
        <th onclick="sortTable('grpTable', 2)">Type</th>
        <th onclick="sortTable('grpTable', 3)">List hidden</th>
        <th onclick="sortTable('grpTable', 4)">Action</th>
        <th onclick="sortTable('grpTable', 5)">Members</th>
        <th onclick="sortTable('grpTable', 6)">Hidden</th>
        <th class="nosort">Membership</th>
      </tr>
    </thead>
    <tbody id="grpBody">
$S_GroupRows
    </tbody>
  </table>
  </div>
</div>

<div class="panel">
  <h2>Members</h2>
  <p class="note">One row per unique member across every list in scope. <strong>Hidden</strong> means <code>HiddenFromAddressListsEnabled</code> is <code>True</code>: the member still receives list mail but cannot be found in the global address list. <strong>Change via</strong> is the route that can actually clear it: the Exchange <code>Set-*</code> cmdlet for most recipients, Microsoft Graph for Entra ID guests, and on-premises Active Directory for directory-synchronized objects.</p>
  <div class="controls">
    <input type="text" id="memSearch" placeholder="Search member, address or list..." onkeyup="filterMembers()" />
    <select id="memStateFilter" onchange="filterMembers()">
      <option value="hidden">Hidden only</option>
      <option value="all">All members</option>
      <option value="visible">Visible only</option>
      <option value="unknown">Unknown only</option>
    </select>
    <select id="memPathFilter" onchange="filterMembers()">
      <option value="all">Any change route</option>
      <option value="ExchangeOnline">Exchange Online</option>
      <option value="Graph">Microsoft Graph (guest)</option>
      <option value="OnPremises">On-premises AD</option>
      <option value="Unsupported">Unsupported</option>
    </select>
    <select id="memTypeFilter" onchange="filterMembers()">
      <option value="all">All recipient types</option>
      $S_RtypeOptions
    </select>
    <span class="count" id="memCount"></span>
  </div>
  <div class="tablescroll">
  <table id="memTable">
    <thead>
      <tr>
        <th onclick="sortTable('memTable', 0)">Member</th>
        <th onclick="sortTable('memTable', 1)">Address</th>
        <th onclick="sortTable('memTable', 2)">Recipient type</th>
        <th onclick="sortTable('memTable', 3)">Address list</th>
        <th onclick="sortTable('memTable', 4)">Change via</th>
        <th onclick="sortTable('memTable', 5)">Action</th>
        <th onclick="sortTable('memTable', 6)">Lists</th>
        <th class="nosort">Member of</th>
      </tr>
    </thead>
    <tbody id="memBody">
$S_MemberRows
    </tbody>
  </table>
  </div>
</div>

<div class="footer">Generated by ReportDistributionListMemberVisibility.ps1 &nbsp;|&nbsp; $S_ReportDate</div>

</div>
<script>
function filterGroups() {
  var search = document.getElementById('grpSearch').value.toLowerCase();
  var hid = document.getElementById('grpHiddenFilter').value;
  var typ = document.getElementById('grpTypeFilter').value;
  var rows = document.querySelectorAll('#grpBody tr');
  var visible = 0;
  rows.forEach(function(r) {
    var text = r.textContent.toLowerCase();
    var rHid = r.getAttribute('data-hidden') || '';
    var rTyp = r.getAttribute('data-type') || '';
    var ok = (!search || text.indexOf(search) > -1) &&
             (hid === 'all' || rHid === hid) &&
             (typ === 'all' || rTyp === typ);
    if (ok) { r.classList.remove('hidden-row'); visible++; } else { r.classList.add('hidden-row'); }
  });
  document.getElementById('grpCount').textContent = 'Showing ' + visible + ' of ' + rows.length + ' lists';
}

function filterMembers() {
  var search = document.getElementById('memSearch').value.toLowerCase();
  var state = document.getElementById('memStateFilter').value;
  var path = document.getElementById('memPathFilter').value;
  var typ = document.getElementById('memTypeFilter').value;
  var rows = document.querySelectorAll('#memBody tr');
  var visible = 0;
  rows.forEach(function(r) {
    var text = r.textContent.toLowerCase();
    var rState = r.getAttribute('data-state') || '';
    var rPath = r.getAttribute('data-path') || '';
    var rTyp = r.getAttribute('data-rtype') || '';
    var ok = (!search || text.indexOf(search) > -1) &&
             (state === 'all' || rState === state) &&
             (path === 'all' || rPath === path) &&
             (typ === 'all' || rTyp === typ);
    if (ok) { r.classList.remove('hidden-row'); visible++; } else { r.classList.add('hidden-row'); }
  });
  document.getElementById('memCount').textContent = 'Showing ' + visible + ' of ' + rows.length + ' members';
}

var sortDir = {};
function sortTable(tableId, col) {
  var tbody = document.getElementById(tableId).querySelector('tbody');
  var rows = Array.prototype.slice.call(tbody.querySelectorAll('tr'));
  var key = tableId + '-' + col;
  var dir = sortDir[key] === 'asc' ? 'desc' : 'asc';
  sortDir[key] = dir;
  rows.sort(function(a, b) {
    var ta = a.cells[col] ? a.cells[col].textContent.trim() : '';
    var tb = b.cells[col] ? b.cells[col].textContent.trim() : '';
    var na = parseFloat(ta), nb = parseFloat(tb);
    if (!isNaN(na) && !isNaN(nb)) { return dir === 'asc' ? na - nb : nb - na; }
    return dir === 'asc' ? ta.localeCompare(tb) : tb.localeCompare(ta);
  });
  rows.forEach(function(r) { tbody.appendChild(r); });
}

// The member table opens on Hidden only - that is the list anyone acts on - so the initial
// filter has to run on load rather than waiting for a change event.
filterMembers();
</script>
</body>
</html>
"@

$S_Html | Out-File -FilePath $S_HtmlFile -Encoding UTF8

}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "Report complete." -ForegroundColor Green
if ($Format -in @('Csv', 'Both'))  { Write-Host "  CSV  : $S_CsvFile" }
if ($Format -in @('Html', 'Both')) { Write-Host "  HTML : $S_HtmlFile" }
if ($S_GuestScriptWritten)
{
    Write-Host "  GUEST SCRIPT : $S_GuestFile" -ForegroundColor Cyan
    Write-Host "    $($S_GuestTodo.Count) hidden guest(s). Run it in a FRESH session (no Exchange connection open)." -ForegroundColor Cyan
}
Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "  Groups in scope          : $S_TotalGroups"
Write-Host "  Memberships              : $S_TotalMemberships"
Write-Host "  Unique members           : $S_UniqueMembers"
Write-Host "  Hidden from address lists: $S_HiddenNowCount" -ForegroundColor $(if ($S_HiddenNowCount -gt 0) { 'Yellow' } else { 'Green' })
Write-Host "  Visible                  : $S_VisibleCount"
if ($S_UnknownCount -gt 0)   { Write-Host "  Visibility unknown       : $S_UnknownCount" -ForegroundColor Yellow }
Write-Host "  Hidden guests (Entra ID) : $S_GuestHiddenCount"
Write-Host "  Hidden groups            : $S_HiddenGroupsNow"
if ($Unhide -or $UnhideGroups)
{
    Write-Host ""
    Write-Host "  Unhidden this run        : $S_UnhiddenCount" -ForegroundColor Green
    if ($S_FailedCount -gt 0) { Write-Host "  Failed                   : $S_FailedCount" -ForegroundColor Red }
    if ($S_OnPremCount -gt 0)      { Write-Host "  Need on-premises change  : $S_OnPremCount" -ForegroundColor Yellow }
    if ($S_GraphNeededCount -gt 0) { Write-Host "  Need Microsoft Graph     : $S_GraphNeededCount" -ForegroundColor Yellow }
    if ($S_NotSelectedCount -gt 0) { Write-Host "  Left hidden (not selected): $S_NotSelectedCount" }
}
elseif ($S_HiddenNowCount -gt 0)
{
    Write-Host ""
    Write-Host "  Read-only run: nothing was changed. Re-run with -Unhide to clear the flag," -ForegroundColor Cyan
    Write-Host "  or -Unhide -WhatIf to preview the change first." -ForegroundColor Cyan
}
Write-Host ""
