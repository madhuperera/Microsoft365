# Reports

Read-only Microsoft 365 reporting scripts for tenant discovery and review.

Individual scripts are catalogued in the [repository README](../README.md#script-catalogue).
This note covers the two "glue" scripts that tie the folder together:

| Script | Role |
|--------|------|
| [`_ReadOnlyConnectionScript.ps1`](_ReadOnlyConnectionScript.ps1) | Establishes **one** shared read-only Microsoft Graph session for the whole folder. |
| [`_RunDiscovery.ps1`](_RunDiscovery.ps1) | Orchestrator — runs a managed list of the report scripts against that shared session in a single pass. |

> **Provided as-is, without warranty.** Review each script before running it. Most
> scripts are read-only, but always confirm the active tenant and account first.

---

## `_RunDiscovery.ps1` — discovery orchestrator

Runs many of the `Report*.ps1` / `Review*.ps1` scripts in one go: connects to Graph
once via `_ReadOnlyConnectionScript.ps1`, then executes each enabled report and writes
every script's output into its own subfolder under a single timestamped discovery folder.

### Quick start

```powershell
# From a PowerShell 7+ session opened in this folder:

# 1. Preview what would run (no connection, nothing executed)
.\_RunDiscovery.ps1 -ListOnly

# 2. Run everything enabled in the manifest
.\_RunDiscovery.ps1

# 3. Run only a connection group (multiple allowed)
.\_RunDiscovery.ps1 -Scope Graph
.\_RunDiscovery.ps1 -Scope Graph,Exchange
```

### Parameters

| Parameter | Purpose |
|-----------|---------|
| `-Scope` | Which connection group(s) to run: `All` (default), `Graph`, `Exchange`, `Teams`, `Other`. Accepts multiple values. |
| `-OutputRoot` | Base output folder. Defaults to `.\Discovery_yyyyMMdd_HHmmss`. |
| `-Include` | Run only these job **Names** (overrides each job's `Enabled` flag). Still subject to `-Scope`/`-Exclude`. |
| `-Exclude` | Skip these job Names. |
| `-ListOnly` | Dry run — print the resolved plan without connecting or running anything. |
| `-Force` | Passed through to `_ReadOnlyConnectionScript.ps1` to force a fresh Graph connection. |
| `-StopOnError` | Abort the run if a job fails. By default a failed job is recorded and the rest continue. |

### Managing which reports run

The list lives in the `$S_DiscoveryJobs` array near the top of the script — one line per
report. Edit it to control the run:

- Set `Enabled = $true` / `$false` to include or skip a report.
- Put any parameter the child script supports in its `Parameters = @{ ... }` hashtable;
  it is splatted straight onto the script (e.g. `@{ InactiveDays = 90 }`).
- `SubFolder` overrides the output subfolder (defaults to the job `Name`).

Each job also records its `Connection` (`Graph` / `Exchange` / `Teams` / `Other`), which
drives `-Scope`, and its `OutputParam` (`OutputPath` = a file, `ReportPath` = a folder, or
`$null`), which is how the orchestrator routes each script's output into the right subfolder.

### Output layout

```
Discovery_20260709_102514/
├── _RunDiscovery.log            # full run transcript
├── _RunDiscovery_Summary.csv    # per-job status / duration / output folder / error
├── AuthMethods/
├── Licensing/
├── SPF/
└── ...                          # one subfolder per job that ran
```

### Jobs disabled by default

These ship `Enabled = $false` so a default run never stalls. Edit their `Parameters`
placeholders (or use `-Include`) before enabling:

| Job | Why disabled |
|-----|--------------|
| `NonMFA`, `MemberMFA` | Require `-InactiveDays`. |
| `IntuneMobileDevices` | Requires `-LatestSupportedAndroid` / `-LatestSupportedIOS`. |
| `CalendarPermissions` | Requires `-DistributionGroupName`. |
| `MdeNetworkDevices` | Windows PowerShell 5.1 only. |

### Notes

- **Semi-attended:** Exchange Online and Teams jobs manage their own sign-in, and several
  child scripts have their own `Read-Host` prompts (confirm context / proceed / disconnect).
  Expect to answer prompts during a run.
- **Run from a PowerShell session**, not `pwsh -File`. Launching externally with
  `-File ... -Scope Graph,Exchange` passes the value as one literal string and fails
  validation; use `pwsh -Command './_RunDiscovery.ps1 -Scope Graph,Exchange'` if you must
  invoke it from outside a session.
- **`_v2` scripts** are the cross-platform variants (DNS-over-HTTPS instead of the
  Windows-only `Resolve-DnsName`, guarded `Out-GridView`) and run on macOS/Linux/Windows.
  The orchestrator references the latest variant of each report.

---

## `_ReadOnlyConnectionScript.ps1` — shared Graph connection

Establishes a single read-only Microsoft Graph session (only `*.Read.All` /
`*.ReadBasic.All` scopes) so the report scripts reuse one context instead of prompting for
consent each time. `_RunDiscovery.ps1` calls it automatically when a Graph job is in scope.
To use it on its own:

```powershell
. .\_ReadOnlyConnectionScript.ps1        # dot-source once; -Force to reconnect
```
