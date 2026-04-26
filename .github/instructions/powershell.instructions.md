---
applyTo: "**/*.ps1,**/*.psm1"
---

# PowerShell scripting conventions for the Microsoft 365 repository

Apply these rules to every `.ps1` and `.psm1` file in this repository.

These rules reflect the preferred repository style. When editing an existing script, match the closest peer script where practical unless the requested change requires a safer or clearer pattern.

## Required script header

1. Place `#Requires -Modules <comma-separated module list>` on line 1 where the script depends on PowerShell modules.

2. For Microsoft Graph scripts, pin the exact Microsoft Graph sub-modules used.

   Preferred examples:

   - `Microsoft.Graph.Authentication`
   - `Microsoft.Graph.Users`
   - `Microsoft.Graph.Groups`
   - `Microsoft.Graph.Reports`
   - `Microsoft.Graph.Identity.DirectoryManagement`

   Do not use the umbrella `Microsoft.Graph` module unless there is a specific repository-supported reason.

3. For Exchange Online scripts, use `ExchangeOnlineManagement`.

4. Include comment-based help with:

   - `.SYNOPSIS`
   - `.DESCRIPTION`
   - `.PARAMETER <Name>` for every parameter
   - At least one `.EXAMPLE`

5. Place `[CmdletBinding()]` immediately above `param(...)`.

## Parameters

- Use one `[Parameter()]` attribute per parameter and declare `Mandatory` explicitly.
- Use specific types such as `[datetime]`, `[int]`, `[string]`, `[string[]]`, `[bool]`, and `[switch]`.
- Do not use untyped `$args` or generic `[object]` unless there is a clear technical reason.
- Validate inputs with `ValidateNotNullOrEmpty`, `ValidateRange`, `ValidateSet`, or `ValidatePattern` where appropriate.
- For long-running tenant-wide scripts, include a `[switch]$Test` parameter or equivalent safety option where practical.
- Default `-OutputPath` values should use a timestamped filename in the current location where the script produces report files.

Example pattern:

~~~powershell
$S_Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$S_DefaultOutputPath = Join-Path -Path (Get-Location).Path -ChildPath "ReportName_$S_Timestamp.csv"
~~~

When emitting an HTML report alongside a CSV, derive the path with:

~~~powershell
[System.IO.Path]::ChangeExtension($OutputPath, '.html')
~~~

## Variable naming

Use the following variable prefixes:

- `S_` for script-level variables.
- `F_` for function-level variables.

Examples:

~~~powershell
$S_OutputPath = $OutputPath
$F_User = $User
~~~

Do not rename unrelated existing variables unless the requested task includes style alignment.

## Formatting

- Place opening braces on a new line.
- Keep indentation consistent and easy to review.
- Group logic under clear section banners.
- Preserve existing section banner style when editing an existing script.

Example:

~~~powershell
if ($S_ShouldRun)
{
    Write-Verbose "Starting operation."
}
~~~

## Error handling and logging

- Set `$ErrorActionPreference = 'Stop'` directly after `param(...)`.
- Use `try`, `catch`, and `finally` where failure handling or cleanup matters.
- Use `Write-Verbose` or `Write-Information` for progress.
- Use `Write-Warning` for recoverable issues.
- Use `throw` for fatal issues.
- Reserve `Write-Host` for human-targeted summaries or interactive prompts.
- Avoid `Write-Error -ErrorAction SilentlyContinue` patterns because they hide problems from the operator.

## Microsoft Graph permissions

When a script uses Microsoft Graph, the required Graph permissions must be declared near the beginning of the script using this script-level variable:

~~~powershell
$S_RequiredGraphScopes = @(
    "User.Read.All"
)
~~~

The variable must be easy to find and consistently named so the repository can be scanned later to identify all Graph permissions required across scripts.

Do not hide Graph permissions deep inside functions unless there is a strong technical reason.

If additional permissions are required by a specific function, they must still be added to `$S_RequiredGraphScopes`.

Only request scopes the script actually uses.

Prefer read-only scopes such as `*.Read.All` over write scopes such as `*.ReadWrite.All` whenever the workflow is read-only.

Never add broad Graph scopes such as `Directory.ReadWrite.All` just in case.

## Microsoft Graph request delay

Scripts that call Microsoft Graph must include a small configurable delay between Graph calls.

Use this script-level variable with a default value of 5 milliseconds:

~~~powershell
$S_GraphRequestDelayMilliseconds = 5
~~~

Where practical, Graph calls should respect this delay:

~~~powershell
Start-Sleep -Milliseconds $S_GraphRequestDelayMilliseconds
~~~

The delay must be controlled at the script variable level so it can be increased later if throttling or service-side limits become an issue.

## Microsoft Graph connection handling

Scripts that connect to Microsoft Graph must use a safe connection process.

The script should first check whether an existing Microsoft Graph context is already available.

Use `Get-MgContext` to inspect the current context where practical.

If an existing context is found, the script must display enough context information for the operator to confirm the tenant and account before continuing.

Where available, show:

- Account
- Tenant ID
- Environment
- Scopes

The script must then ask whether to continue using the existing Graph connection.

If no existing Graph connection is available, the script must prompt for a new connection using the scopes defined in `$S_RequiredGraphScopes`.

After connecting, the script must display the active Graph context and ask the operator to confirm that the correct tenant and account are being used before continuing.

For interactive scripts, an explicit confirmation prompt is preferred.

For automation scripts, avoid interactive prompts unless the script is specifically designed for attended execution.

Use script-level variables to control confirmation behaviour where practical:

~~~powershell
$S_RequireGraphContextConfirmation = $true
$S_GraphContextConfirmationDelaySeconds = 10
~~~

## Microsoft Graph disconnection handling

At the end of an interactive script that uses Microsoft Graph, the script must ask whether the operator wants to disconnect from the current Graph session.

The script should not automatically disconnect unless that behaviour is clearly intended and documented.

For automation scripts, disconnection behaviour should be controlled by a script-level variable:

~~~powershell
$S_DisconnectGraphSessionOnExit = $false
~~~

## Exchange Online authentication

For Exchange Online scripts:

- Verify the `ExchangeOnlineManagement` module is present where practical.
- Throw a clear install hint if the module is missing.
- Use `Connect-ExchangeOnline -ShowBanner:$false` unless the existing script has a specific reason not to.
- Do not embed credentials, client secrets, certificate thumbprints, tenant IDs, or UPNs in the script body or examples.

## Output and reporting

- Scripts that generate reports should write output to CSV where practical.
- HTML report output may be added where it improves readability for operators or clients.
- Do not change output schema unexpectedly unless requested.
- If changing output columns, document the change in the pull request summary.
- Keep report column names clear and operationally useful.

## File names

- New cmdlet-style scripts should use `Verb-Noun.ps1` with approved PowerShell verbs.
- When editing legacy report scripts that already use names such as `Report*.ps1` or `Verify*.ps1`, keep the existing name unless the task is specifically to rename or restructure files.
- Do not combine a rename with behavioural changes unless requested.

### Reporting scripts must use the `Report` prefix

Any script whose primary purpose is to produce a report (CSV, HTML, JSON, or
console summary describing tenant state) must use the `Report` prefix in its
filename.

- Preferred form: `Report<Subject>.ps1` using PascalCase, no separators
  (for example `ReportInactiveGuestUsers.ps1`, `ReportLicensing.ps1`,
  `ReportCalendarPermissions.ps1`).
- Do not use `Get-*Report.ps1`, `Verify*.ps1`, or bare nouns such as
  `MailboxQuota.ps1` for new reporting scripts. Reserve `Get-*` for scripts
  that return objects to the pipeline for further processing rather than
  producing a finished report file.
- Do not use underscores or hyphens between `Report` and the subject
  (`Report_CalendarPermissions.ps1` is not preferred — use
  `ReportCalendarPermissions.ps1`).
- Existing scripts that do not yet follow this rule must not be renamed as
  part of an unrelated change. Renames are tracked separately so git history
  and any external references remain easy to follow.
- When this convention is applied to existing files in a dedicated rename
  pass, update the corresponding `.SYNOPSIS`, `.EXAMPLE`, and any README
  references in the same change.

### Reporting scripts must live in the top-level `Reports/` folder

All new reporting scripts for every Microsoft 365 workload must be placed in
the single top-level `Reports/` folder. This rule is non-negotiable.

- Do not create per-workload report folders such as `EntraID/Reports/`,
  `Exchange/Reports/`, `Intune/Reports/`, `DNS/Reports/`, or similar.
- Do not move reporting scripts into workload folders to "group" them.
  Group reports through the filename prefix instead (for example
  `ReportEntraID*`, `ReportExo*`, `ReportIntune*`, `ReportDns*`).
- Workload folders such as `Exchange/`, `Configurations/`, and
  `Incident_Response/` are reserved for non-reporting assets such as
  configuration data, transport-rule HTML, remediation actions, and
  reactive incident-response tooling.

## Documentation expectations for PowerShell scripts

Where practical, each script should document:

- Purpose
- Required modules
- Required Graph permissions, if applicable
- Required Exchange Online permissions, if applicable
- Inputs
- Outputs
- Usage example
- Validation steps
- Known limitations

Do not invent permissions, modules, usage examples, or validation steps that are not supported by the script or repository content.
