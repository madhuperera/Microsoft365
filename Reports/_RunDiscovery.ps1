#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Orchestrates a Microsoft 365 discovery by running a managed list of the
    Report*.ps1 / Review*.ps1 scripts in this folder against a single shared,
    read-only Microsoft Graph connection.

.DESCRIPTION
    This script is the single entry point for a discovery engagement. It:

      1. Establishes the shared read-only Microsoft Graph context once, by
         invoking _ReadOnlyConnectionScript.ps1 (only when at least one selected
         job needs Graph). Child scripts then reuse that context instead of
         prompting for consent individually.

      2. Runs a manageable, in-script list of report jobs ($S_DiscoveryJobs).
         Each job names the child script, whether it is enabled, which
         connection surface it needs (Graph / Exchange / Teams / Other), the
         name of that child's output parameter, its output subfolder, and a
         splatted hashtable of any parameters the child script supports.

      3. Lets you choose which group of scripts to run at launch with -Scope
         (All / Graph / Exchange / Teams / Other; multiple values allowed).

      4. Writes each script's output into its own subfolder under one
         timestamped discovery folder, plus a transcript log and a run summary.

    Child scripts are invoked with the call operator (&), so they run in the
    same process (sharing the Process-scoped Graph token) but in a child
    variable scope (so their $S_* variables do not leak into or collide with
    this orchestrator).

    IMPORTANT NOTES
      * Exchange Online and Microsoft Teams jobs manage their own sign-in
        (the child scripts call Connect-ExchangeOnline / Connect-MicrosoftTeams
        internally). Expect interactive sign-in prompts for those.
      * Several child scripts contain their own Read-Host prompts (confirm
        context, proceed, disconnect). This orchestrator does not suppress
        them, so a run is semi-attended.
      * No existing script in this folder is modified by this orchestrator.

.PARAMETER Scope
    Which connection group(s) of jobs to run. One or more of:
    All, Graph, Exchange, Teams, Other. Default: All.

.PARAMETER OutputRoot
    Base folder for this discovery's output. Defaults to a timestamped
    .\Discovery_yyyyMMdd_HHmmss folder in the current location.

.PARAMETER Include
    Run only the jobs whose Name is in this list (overrides each job's Enabled
    flag). Still subject to -Scope and -Exclude.

.PARAMETER Exclude
    Skip the jobs whose Name is in this list.

.PARAMETER ListOnly
    Dry run. Prints the resolved plan (which jobs would run, and where their
    output would go) without connecting to anything or running any child script.

.PARAMETER Force
    Passed through to _ReadOnlyConnectionScript.ps1 to force a fresh Graph
    connection with the read-only scope set.

.PARAMETER StopOnError
    Abort the whole run if a job fails. By default a failed job is recorded and
    the remaining jobs continue.

.EXAMPLE
    .\_RunDiscovery.ps1 -ListOnly

    Shows the default plan (all enabled jobs) without running anything.

.EXAMPLE
    .\_RunDiscovery.ps1 -Scope Graph

    Runs only the enabled Graph jobs.

.EXAMPLE
    .\_RunDiscovery.ps1 -Scope Graph,Exchange -OutputRoot C:\Discovery\Contoso

    Runs the enabled Graph and Exchange jobs, writing everything under the
    given folder.

.EXAMPLE
    .\_RunDiscovery.ps1 -Include NonMFA,MemberMFA

    Runs two normally-disabled jobs by name (edit their Parameters in the
    manifest first to supply the mandatory values).

.NOTES
    Edit the $S_DiscoveryJobs array below to manage which reports run and with
    what parameters. Set Enabled = $true/$false, or pass extra parameters in
    the Parameters hashtable (they are splatted onto the child script).
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('All', 'Graph', 'Exchange', 'Teams', 'Other')]
    [string[]]$Scope = 'All',

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputRoot,

    [Parameter(Mandatory = $false)]
    [string[]]$Include,

    [Parameter(Mandatory = $false)]
    [string[]]$Exclude,

    [Parameter(Mandatory = $false)]
    [switch]$ListOnly,

    [Parameter(Mandatory = $false)]
    [switch]$Force,

    [Parameter(Mandatory = $false)]
    [switch]$StopOnError
)

$ErrorActionPreference = 'Stop'

# ===========================================================================
# Discovery job manifest — THE LIST YOU MANAGE
#
# One entry per report. Fields:
#   Name        Short id for the job; also the default output subfolder name.
#   Script      The child script filename in this folder.
#   Enabled     $true to run by default; $false to skip (override with -Include).
#   Connection  Graph | Exchange | Teams | Other  (drives -Scope filtering and
#               whether the shared Graph connector is invoked).
#   OutputParam Informational only — the child's output parameter name
#               ('OutputPath' / 'ReportPath' / $null). Output is routed by
#               running the script from inside its SubFolder (working
#               directory), not by injecting this parameter, because these
#               parameters mean different things across scripts (file vs
#               folder). To force a specific path, put it in Parameters below.
#   SubFolder   Output subfolder under -OutputRoot (defaults to Name if $null).
#               The script runs with this as its working directory, so its
#               output lands here using the script's own naming.
#   Parameters  Hashtable of extra parameters splatted onto the child script.
#               Use this to supply mandatory values, tweak behaviour, or pin an
#               explicit output path/file.
#
# Every report is listed here — nothing is commented out. Enabled is the only
# on/off switch: set Enabled = $true to include a job in a default run, or
# $false to park it in the list for easy re-enabling later. A disabled job can
# still be forced for a single run with -Include <Name>.
#
# Note: some jobs require mandatory child-script parameters (supplied in their
# Parameters hashtable below) — e.g. InactiveDays, the Intune version baselines,
# and CalendarPermissions' DistributionGroupName. Keep those populated.
# ===========================================================================
$S_DiscoveryJobs = @(

    # ----- Microsoft Graph (sorted by SubFolder; unset SubFolder last) -----
    @{ Name = 'AllWindowsDevices';    Script = 'ReportAllWindowsDevices.ps1';                 Enabled = $true;  Connection = 'Graph';    OutputParam = 'ReportPath'; SubFolder = "Device Management";  Parameters = @{ InactiveDays = 90 } }
    @{ Name = 'Domains';              Script = 'ReportDomains_v2.ps1';                        Enabled = $true;  Connection = 'Graph';    OutputParam = 'OutputPath'; SubFolder = "Domains";            Parameters = @{} }
    @{ Name = 'EntraIDApps';          Script = 'ReportEntraIDApps.ps1';                       Enabled = $true;  Connection = 'Graph';    OutputParam = 'ReportPath'; SubFolder = "Entra ID";           Parameters = @{ InactiveDays = 90 } }
    @{ Name = 'IntuneMobileDevices';  Script = 'ReportIntuneMobileDevices.ps1';               Enabled = $true;  Connection = 'Graph';    OutputParam = 'ReportPath'; SubFolder = "Intune";             Parameters = @{ LatestSupportedAndroid = 14; LatestSupportedIOS = 15 } }
    @{ Name = 'IntuneWindowsDevices'; Script = 'ReportIntuneWindowsDevices.ps1';              Enabled = $true;  Connection = 'Graph';    OutputParam = 'ReportPath'; SubFolder = "Intune";             Parameters = @{ MinimumSupportedWindows11Build = 26100 } }
    @{ Name = 'Licensing';            Script = 'ReportLicensing.ps1';                         Enabled = $true;  Connection = 'Graph';    OutputParam = 'OutputPath'; SubFolder = "Licensing";          Parameters = @{} }
    @{ Name = 'TeamsGroups';          Script = 'ReportTeamsGroups.ps1';                       Enabled = $true;  Connection = 'Graph';    OutputParam = 'ReportPath'; SubFolder = "MS Teams";           Parameters = @{} }
    @{ Name = 'AllMemberUsers';       Script = 'ReportAllMemberUsers.ps1';                    Enabled = $true;  Connection = 'Graph';    OutputParam = 'ReportPath'; SubFolder = "User Identity";      Parameters = @{ InactiveDays = 90 } }
    @{ Name = 'EntraIDRoles';         Script = 'ReportEntraIDRolesMemberships.ps1';           Enabled = $true;  Connection = 'Graph';    OutputParam = 'OutputPath'; SubFolder = "User Identity";      Parameters = @{} }
    @{ Name = 'InactiveGuests';       Script = 'ReviewInactiveGuestUsers.ps1';                Enabled = $true;  Connection = 'Graph';    OutputParam = 'ReportPath'; SubFolder = "User Identity";      Parameters = @{ InactiveDays = 90 } }
    @{ Name = 'InactiveMembers';      Script = 'ReviewInactiveMemberUsers.ps1';               Enabled = $true;  Connection = 'Graph';    OutputParam = 'ReportPath'; SubFolder = "User Identity";      Parameters = @{ InactiveDays = 90 } }
    @{ Name = 'MemberMFA';            Script = 'ReportMemberMFA_v4.ps1';                      Enabled = $true;  Connection = 'Graph';    OutputParam = 'OutputPath'; SubFolder = "User Identity";      Parameters = @{ InactiveDays = 90 } }
    @{ Name = 'AADAuthMethods';       Script = 'ReportAADAuthenticationMethods.ps1';          Enabled = $false; Connection = 'Graph';    OutputParam = 'OutputPath'; SubFolder = $null;               Parameters = @{} }
    @{ Name = 'AuthMethods';          Script = 'ReportAuthenticationMethods_v2.ps1';          Enabled = $false; Connection = 'Graph';    OutputParam = 'OutputPath'; SubFolder = $null;               Parameters = @{} }
    @{ Name = 'IntuneApps';           Script = 'ReportIntuneApps.ps1';                        Enabled = $false; Connection = 'Graph';    OutputParam = 'ReportPath'; SubFolder = $null;               Parameters = @{} }
    @{ Name = 'IntuneCompliance';     Script = 'ReportIntuneCompliance.ps1';                  Enabled = $false; Connection = 'Graph';    OutputParam = 'ReportPath'; SubFolder = $null;               Parameters = @{} }
    @{ Name = 'LegacyAuthGuests';     Script = 'ReportLegacyAuthenticationMethodsGuests.ps1'; Enabled = $false; Connection = 'Graph';    OutputParam = 'OutputPath'; SubFolder = $null;               Parameters = @{} }
    @{ Name = 'LegacyAuthMethods';    Script = 'ReportLegacyAuthenticationMethods.ps1';       Enabled = $false; Connection = 'Graph';    OutputParam = 'OutputPath'; SubFolder = $null;               Parameters = @{} }
    @{ Name = 'NonMFA';               Script = 'ReportNonMFA.ps1';                            Enabled = $false; Connection = 'Graph';    OutputParam = 'OutputPath'; SubFolder = $null;               Parameters = @{ InactiveDays = 90 } }
    @{ Name = 'UsersWithManagers';    Script = 'ReportUsersWithManagers.ps1';                 Enabled = $false; Connection = 'Graph';    OutputParam = 'OutputPath'; SubFolder = $null;               Parameters = @{} }

    # ----- Exchange Online (sorted by SubFolder; child scripts self-connect)
    @{ Name = 'MailboxQuota';         Script = 'ReportMailboxQuota.ps1';                      Enabled = $true;  Connection = 'Exchange'; OutputParam = 'ReportPath'; SubFolder = "M365 Admin Centre";  Parameters = @{} }
    @{ Name = 'CalendarPermissions';  Script = 'ReportCalendarPermissions.ps1';               Enabled = $false; Connection = 'Exchange'; OutputParam = $null;        SubFolder = $null;               Parameters = @{ DistributionGroupName = 'All Staff' } }
    @{ Name = 'DKIM';                 Script = 'ReportDkimRecords_v2.ps1';                    Enabled = $false; Connection = 'Exchange'; OutputParam = $null;        SubFolder = $null;               Parameters = @{} }
    @{ Name = 'DMARC';                Script = 'ReportDmarcRecords_v2.ps1';                   Enabled = $false; Connection = 'Exchange'; OutputParam = $null;        SubFolder = $null;               Parameters = @{} }
    @{ Name = 'SPF';                  Script = 'ReportSPFRecords_v2.ps1';                     Enabled = $false; Connection = 'Exchange'; OutputParam = $null;        SubFolder = $null;               Parameters = @{} }
    @{ Name = 'UnusedMailboxes';      Script = 'ReportUnusedExoMailboxes_v2.ps1';             Enabled = $false; Connection = 'Exchange'; OutputParam = 'OutputPath'; SubFolder = $null;               Parameters = @{} }

    # ----- Microsoft Teams (child script self-connects) --------------------
    @{ Name = 'TeamsSettings';        Script = 'ReportMSTeamsSettings.ps1';                   Enabled = $true;  Connection = 'Teams';    OutputParam = 'ReportPath'; SubFolder = "MS Teams";           Parameters = @{} }

    # ----- Other (Windows PowerShell 5.1 only) -----------------------------
    @{ Name = 'MdeNetworkDevices';    Script = 'ReportMdeNetworkDevices.ps1';                 Enabled = $false; Connection = 'Other';    OutputParam = 'OutputPath'; SubFolder = $null;               Parameters = @{} }
)

# ===========================================================================
# Resolve paths and timestamp
# ===========================================================================
$S_ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$S_RunTimestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

if (-not $OutputRoot)
{
    $OutputRoot = Join-Path -Path (Get-Location).Path -ChildPath "Discovery_$S_RunTimestamp"
}

$S_ConnectorPath = Join-Path -Path $S_ScriptRoot -ChildPath '_ReadOnlyConnectionScript.ps1'

# ===========================================================================
# Select which jobs will run
#   effective-enabled = (-Include ? Name-in-Include : Enabled)
#   AND connection in -Scope  AND  Name not in -Exclude
# ===========================================================================
$S_ScopeAll = ($Scope -contains 'All')

$S_SelectedJobs = foreach ($S_Job in $S_DiscoveryJobs)
{
    if (-not $S_Job.SubFolder) { $S_Job.SubFolder = $S_Job.Name }

    $S_EffEnabled = if ($Include) { $Include -contains $S_Job.Name } else { [bool]$S_Job.Enabled }
    if (-not $S_EffEnabled) { continue }

    if (-not $S_ScopeAll -and ($Scope -notcontains $S_Job.Connection)) { continue }

    if ($Exclude -and ($Exclude -contains $S_Job.Name)) { continue }

    $S_Job
}

$S_SelectedJobs = @($S_SelectedJobs)

if ($S_SelectedJobs.Count -eq 0)
{
    Write-Warning "No jobs matched the current selection (Scope: $($Scope -join ', ')). Nothing to do."
    return
}

# ===========================================================================
# Print the resolved plan
# ===========================================================================
Write-Host ""
Write-Host "Microsoft 365 Discovery — run plan" -ForegroundColor Cyan
Write-Host ("  Scope       : {0}" -f ($Scope -join ', '))
Write-Host ("  Output root : {0}" -f $OutputRoot)
Write-Host ("  Jobs        : {0}" -f $S_SelectedJobs.Count)
Write-Host ""

$S_SelectedJobs |
    Select-Object @{ N = 'Name'; E = { $_.Name } },
                  @{ N = 'Connection'; E = { $_.Connection } },
                  @{ N = 'Script'; E = { $_.Script } },
                  @{ N = 'SubFolder'; E = { $_.SubFolder } } |
    Format-Table -AutoSize | Out-Host

$S_NeedsGraph = @($S_SelectedJobs | Where-Object { $_.Connection -eq 'Graph' }).Count -gt 0
$S_NeedsExo   = @($S_SelectedJobs | Where-Object { $_.Connection -eq 'Exchange' }).Count -gt 0
$S_NeedsTeams = @($S_SelectedJobs | Where-Object { $_.Connection -eq 'Teams' }).Count -gt 0

if ($S_NeedsExo -or $S_NeedsTeams)
{
    Write-Host "Note: Exchange Online / Teams jobs manage their own sign-in; expect additional prompts." -ForegroundColor Yellow
}
Write-Host "Note: some child scripts have their own interactive prompts (confirm / proceed / disconnect)." -ForegroundColor Yellow
Write-Host ""

if ($ListOnly)
{
    Write-Host "-ListOnly specified: no connection made and no scripts run." -ForegroundColor Green
    return
}

# ===========================================================================
# Prepare output root and transcript
# ===========================================================================
if (-not (Test-Path -Path $OutputRoot -PathType Container))
{
    New-Item -Path $OutputRoot -ItemType Directory -Force | Out-Null
}

$S_TranscriptPath = Join-Path -Path $OutputRoot -ChildPath '_RunDiscovery.log'
$S_TranscriptStarted = $false
try
{
    Start-Transcript -Path $S_TranscriptPath -Append | Out-Null
    $S_TranscriptStarted = $true
}
catch
{
    Write-Warning "Could not start transcript ($($_.Exception.Message)); continuing without it."
}

# ===========================================================================
# Establish the shared read-only Microsoft Graph connection (once)
# ===========================================================================
if ($S_NeedsGraph)
{
    if (-not (Test-Path -Path $S_ConnectorPath -PathType Leaf))
    {
        if ($S_TranscriptStarted) { Stop-Transcript | Out-Null }
        throw "Graph connector not found: $S_ConnectorPath"
    }

    Write-Host "Establishing shared read-only Microsoft Graph connection..." -ForegroundColor Cyan
    try
    {
        if ($Force) { & $S_ConnectorPath -Force } else { & $S_ConnectorPath }
    }
    catch
    {
        if ($S_TranscriptStarted) { Stop-Transcript | Out-Null }
        throw "Failed to establish the shared Microsoft Graph connection: $($_.Exception.Message)"
    }
}

# ===========================================================================
# Run the selected jobs
# ===========================================================================
$S_Results = [System.Collections.Generic.List[object]]::new()
$S_JobIndex = 0

foreach ($S_Job in $S_SelectedJobs)
{
    $S_JobIndex++
    Write-Host ""
    Write-Host ("[{0}/{1}] {2}  ({3})" -f $S_JobIndex, $S_SelectedJobs.Count, $S_Job.Name, $S_Job.Connection) -ForegroundColor Cyan
    Write-Host ("        Script: {0}" -f $S_Job.Script)

    $S_ScriptPath = Join-Path -Path $S_ScriptRoot -ChildPath $S_Job.Script
    $S_SubFolderPath = Join-Path -Path $OutputRoot -ChildPath $S_Job.SubFolder

    $S_Status = 'Succeeded'
    $S_ErrorMessage = $null
    $S_Duration = [TimeSpan]::Zero

    if (-not (Test-Path -Path $S_ScriptPath -PathType Leaf))
    {
        $S_Status = 'Skipped'
        $S_ErrorMessage = "Script not found: $S_ScriptPath"
        Write-Warning $S_ErrorMessage
    }
    else
    {
        if (-not (Test-Path -Path $S_SubFolderPath -PathType Container))
        {
            New-Item -Path $S_SubFolderPath -ItemType Directory -Force | Out-Null
        }

        # Clone the manifest Parameters so we never mutate the source hashtable.
        $S_Params = @{}
        foreach ($S_Key in $S_Job.Parameters.Keys) { $S_Params[$S_Key] = $S_Job.Parameters[$S_Key] }

        # Output routing is handled by the working directory, NOT by injecting a
        # path into the child's output parameter. Every report script defaults
        # its output to (Get-Location) when no path is supplied, so running it
        # from inside the subfolder (Push-Location below) lets each script use
        # its own correct file/folder handling and native file naming.
        #
        # Injecting a path here is unsafe because the OutputPath/ReportPath
        # parameters are overloaded across scripts: some expect a CSV file, some
        # expect a folder, and some (e.g. ReportDomains_v2) treat an unrecognised
        # value as a NEW folder to create — which produced a nested
        # "Domains\Domains_<ts>.csv\" directory. If the operator supplies an
        # explicit path in the manifest Parameters, it is still honoured.
        Push-Location -Path $S_SubFolderPath
        try
        {
            $S_Duration = Measure-Command {
                & $S_ScriptPath @S_Params
            }
            Write-Host ("        Done in {0:c}" -f $S_Duration) -ForegroundColor Green
        }
        catch
        {
            $S_Status = 'Failed'
            $S_ErrorMessage = $_.Exception.Message
            Write-Warning ("        {0} failed: {1}" -f $S_Job.Name, $S_ErrorMessage)
        }
        finally
        {
            Pop-Location
        }
    }

    $S_Results.Add([PSCustomObject]@{
        Name         = $S_Job.Name
        Connection   = $S_Job.Connection
        Status       = $S_Status
        Duration     = $S_Duration.ToString('c')
        OutputFolder = $S_SubFolderPath
        Error        = $S_ErrorMessage
    })

    if ($S_Status -eq 'Failed' -and $StopOnError)
    {
        Write-Warning "-StopOnError specified; aborting remaining jobs."
        break
    }
}

# ===========================================================================
# Summary
# ===========================================================================
Write-Host ""
Write-Host "Discovery summary" -ForegroundColor Cyan
Write-Host "-----------------"
$S_Results | Format-Table -AutoSize | Out-Host

$S_SummaryCsv = Join-Path -Path $OutputRoot -ChildPath '_RunDiscovery_Summary.csv'
try
{
    $S_Results | Export-Csv -Path $S_SummaryCsv -NoTypeInformation -Encoding UTF8
    Write-Host ("Summary written to: {0}" -f $S_SummaryCsv) -ForegroundColor Green
}
catch
{
    Write-Warning "Could not write summary CSV: $($_.Exception.Message)"
}

$S_Failed = @($S_Results | Where-Object { $_.Status -eq 'Failed' }).Count
$S_Skipped = @($S_Results | Where-Object { $_.Status -eq 'Skipped' }).Count
Write-Host ("Output root: {0}" -f $OutputRoot) -ForegroundColor Green
Write-Host ("Completed: {0} succeeded, {1} failed, {2} skipped." -f `
    @($S_Results | Where-Object { $_.Status -eq 'Succeeded' }).Count, $S_Failed, $S_Skipped) -ForegroundColor Green

if ($S_TranscriptStarted) { Stop-Transcript | Out-Null }
