# Microsoft 365 Administration and Security Tooling

<div>
<table>
<tr>
<td><strong>Purpose</strong></td>
<td>PowerShell scripts and supporting assets for Microsoft 365 tenant administration, reporting, and incident response.</td>
</tr>
<tr>
<td><strong>Audience</strong></td>
<td>Microsoft 365 administrators and security practitioners.</td>
</tr>
<tr>
<td><strong>Licence</strong></td>
<td>See <a href="LICENSE">LICENSE</a>.</td>
</tr>
</table>
</div>

> **These scripts are shared as-is without any warranty.**
> Review each script carefully before running it in your environment.
> Many scripts require elevated Microsoft 365 permissions.
> Test against a non-production tenant or in report-only mode wherever possible.

---

## Contents

- [Repository overview](#repository-overview)
- [Quick start](#quick-start)
- [Folder structure](#folder-structure)
- [Prerequisites](#prerequisites)
- [Script catalogue](#script-catalogue)
  - [Reports — Entra ID and Identity](#reports--entra-id-and-identity)
  - [Reports — Exchange Online](#reports--exchange-online)
  - [Reports — Intune](#reports--intune)
  - [Reports — Licensing](#reports--licensing)
  - [Reports — DNS and email security](#reports--dns-and-email-security)
  - [Reports — Microsoft Teams](#reports--microsoft-teams)
  - [Exchange — Calendar permission management](#exchange--calendar-permission-management)
  - [Exchange — Transport rules](#exchange--transport-rules)
  - [Exchange — Anti-malware policy](#exchange--anti-malware-policy)
  - [Incident response](#incident-response)
- [Safe usage notes](#safe-usage-notes)
- [Naming conventions and repository standards](#naming-conventions-and-repository-standards)
- [Maintenance guidance](#maintenance-guidance)
- [Feedback and contact](#feedback-and-contact)

---

## Repository overview

This repository contains administrator and security tooling for Microsoft 365 services. The primary content is PowerShell scripts that report on, validate, or take action on Microsoft 365 tenant configuration. Supporting assets such as HTML templates, policy configuration files, and reference lists are also included.

Scripts cover the following Microsoft 365 workloads:

- Entra ID (Azure Active Directory) — users, guests, devices, roles, and application registrations
- Exchange Online — mailboxes, calendar permissions, and DNS email security records
- Microsoft Teams — Teams-enabled groups and tenant-wide governance or security settings
- Microsoft Intune — managed application inventory, device compliance, and Windows/mobile device posture reporting
- Microsoft 365 licensing — plan and SKU reporting
- Incident response — audit log analysis and sign-in investigation

---

## Quick start

1. Install the required PowerShell modules listed in [Prerequisites](#prerequisites).
2. Open PowerShell in the repository root folder.
3. For Graph reporting scripts, run [`Reports/_ReadOnlyConnectionScript.ps1`](Reports/_ReadOnlyConnectionScript.ps1) once to establish a read-only Graph session.
4. Run the script you need from the [Script catalogue](#script-catalogue).
5. Confirm expected output (CSV, HTML, or console summary) has been generated.

---

## Folder structure

```
Microsoft365/
├── .github/                     Repository automation and contributor guidance
│   ├── workflows/               Scheduled Copilot documentation/compliance checks
│   └── instructions/            Repository coding and documentation instruction files
├── Reports/                     All reporting scripts across all workloads
├── Exchange/
│   ├── AntiMalwarePolicy/       Anti-malware policy reference files
│   ├── ManageCalendarPermissionsViaGroups/   Bulk calendar permission scripts
│   └── TransportRules/          Transport rule supporting assets (HTML templates)
├── Incident_Response/           Incident response and investigation scripts
└── README.md
```

### Folder descriptions

| Folder | Purpose |
|--------|---------|
| `.github/` | Repository workflow automation and instruction files used to guide maintenance tasks. |
| `Reports/` | All reporting scripts for every Microsoft 365 workload. Scripts produce CSV or HTML output suitable for review or client delivery. |
| `Exchange/` | Exchange Online operational scripts and supporting assets. Workload-specific non-reporting tooling lives here. |
| `Exchange/AntiMalwarePolicy/` | Reference files for Microsoft Defender anti-malware policy configuration, including a curated list of file types to block. |
| `Exchange/ManageCalendarPermissionsViaGroups/` | Scripts for bulk management of mailbox calendar permissions via distribution group membership. |
| `Exchange/TransportRules/` | HTML templates and assets used in Exchange Online transport rules. |
| `Incident_Response/` | Reactive investigation scripts for use during a security incident. Scripts query audit logs and sign-in data to support analysis. |

---

## Prerequisites

### Shared connection helper

[`Reports/_ReadOnlyConnectionScript.ps1`](Reports/_ReadOnlyConnectionScript.ps1) is a shared helper that establishes a single read-only Microsoft Graph session for the entire `Reports/` folder. Run it once at the start of a reporting session (dot-source it with `. .\_ReadOnlyConnectionScript.ps1`) so that subsequent report scripts can reuse the same context without being prompted for consent on each run. Pass `-Force` to disconnect any existing session and reconnect automatically.

### PowerShell modules

Install the required modules from the PowerShell Gallery before running any script.

```powershell
# Microsoft Graph (install individual sub-modules as required per script)
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Users -Scope CurrentUser
Install-Module Microsoft.Graph.Groups -Scope CurrentUser
Install-Module Microsoft.Graph.Reports -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
Install-Module Microsoft.Graph.Applications -Scope CurrentUser
Install-Module Microsoft.Graph.Teams -Scope CurrentUser

# Exchange Online Management
Install-Module ExchangeOnlineManagement -Scope CurrentUser

# Microsoft Teams Management
Install-Module MicrosoftTeams -Scope CurrentUser
```

### Permissions

Required permissions vary per script. High-level requirements are listed in the script catalogue below. Always check the individual script header for the exact scopes or roles required.

Most reporting scripts require **read-only** Microsoft Graph delegated permissions, Exchange Online read access, or Microsoft Teams read access. Scripts that can disable accounts require additional write permissions.

---

## Script catalogue

### Reports — Entra ID and Identity

All scripts in this section are located in [`Reports/`](Reports/).

<details>
<summary>View Entra ID and Identity scripts</summary>

| Script | Description | Key permissions / modules |
|--------|-------------|--------------------------|
| [`ReportAllMemberUsers.ps1`](Reports/ReportAllMemberUsers.ps1) | Reports on all member users in Entra ID including account status, last sign-in, licensing, and on-premises sync state. Supports an inactivity threshold parameter. Exports to CSV and HTML. | `Microsoft.Graph.Users` |
| [`ReportAllWindowsDevices.ps1`](Reports/ReportAllWindowsDevices.ps1) | Reports on all Windows devices registered in Entra ID, including registration state and last activity. Supports an inactivity threshold parameter. Exports to CSV and HTML. | `Microsoft.Graph.Identity.DirectoryManagement` |
| [`ReportAuthenticationMethods.ps1`](Reports/ReportAuthenticationMethods.ps1) | Reports authentication methods registered for licensed Entra ID member accounts. Exports to CSV. | `UserAuthenticationMethod.Read.All`, `Directory.Read.All`, `User.Read.All` |
| [`ReportEntraIDApps.ps1`](Reports/ReportEntraIDApps.ps1) | Reports on enterprise applications (service principals) in Entra ID, including sign-in activity. Cross-references app registrations to identify which service principals have a local app registration. Supports an inactivity threshold parameter. Exports to CSV and HTML. | `Microsoft.Graph.Applications`, `Microsoft.Graph.Identity.DirectoryManagement` |
| [`ReportEntraIDRolesMemberships.ps1`](Reports/ReportEntraIDRolesMemberships.ps1) | Reports on users assigned to administrator or Global Reader directory roles, including last sign-in date. Exports to CSV and HTML. | `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users`, `Microsoft.Graph.Groups`, `Microsoft.Graph.Identity.DirectoryManagement` |
| [`ReviewInactiveGuestUsers.ps1`](Reports/ReviewInactiveGuestUsers.ps1) | Reports on guest users who have not signed in within a configurable number of days. Supports a `Disable` mode to disable accounts after reporting. Exports to CSV and HTML. | `Microsoft.Graph.Users` |
| [`ReviewInactiveMemberUsers.ps1`](Reports/ReviewInactiveMemberUsers.ps1) | Reports on member users who have not signed in within a configurable number of days. Supports a `Disable` mode to disable accounts after reporting. Exports to CSV and HTML. | `Microsoft.Graph.Users` |
| [`ReportLegacyAuthenticationMethods.ps1`](Reports/ReportLegacyAuthenticationMethods.ps1) | Reports authentication methods for all Entra ID member accounts, classifying each method as Modern or Legacy. Includes licence status and on-premises sync status. Exports to CSV and HTML. | `UserAuthenticationMethod.Read.All`, `Directory.Read.All`, `User.Read.All` |
| [`ReportLegacyAuthenticationMethodsGuests.ps1`](Reports/ReportLegacyAuthenticationMethodsGuests.ps1) | Reports authentication methods for Entra ID guest accounts, classifying each method as Modern or Legacy. Exports to CSV. | `UserAuthenticationMethod.Read.All`, `Directory.Read.All`, `User.Read.All` |
| [`ReportUsersWithManagers.ps1`](Reports/ReportUsersWithManagers.ps1) | Exports Entra ID users who have a manager assigned, including the manager's display name and email. Exports to CSV. | `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users` |
| [`ReportNonMFA.ps1`](Reports/ReportNonMFA.ps1) | Generates a report of member users with MFA registration status, account activity, licensing, on-premises sync state, and group memberships. Exports to CSV and HTML. | `Microsoft.Graph.Authentication`, `Microsoft.Graph.Groups`, `Microsoft.Graph.Users`, `Microsoft.Graph.Reports` |
| [`ReportMemberMFA_v1.ps1`](Reports/ReportMemberMFA_v1.ps1) | Reports MFA coverage across all Member users (Guests excluded), classifying each user as Modern Auth, Legacy Auth, or No MFA. Supports a `-Test` switch to run against the first 10 Member users only. Exports to CSV and HTML. | `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users`, `Microsoft.Graph.Reports`, `Microsoft.Graph.Identity.SignIns`; `User.Read.All`, `AuditLog.Read.All`, `UserAuthenticationMethod.Read.All`, `Policy.Read.All` |
| [`ReportMemberMFA_v2.ps1`](Reports/ReportMemberMFA_v2.ps1) | Enhanced MFA coverage report for all Member users. In addition to classifying each user by authentication method (Modern Auth, Legacy Auth, or No MFA), queries and analyses Conditional Access policies that enforce MFA to determine per-user CA coverage. Supports a `-Test` switch. Exports to CSV and HTML. See [`ReportMemberMFA-Guide.html`](Reports/ReportMemberMFA-Guide.html) for a guide to interpreting the report output. | `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users`, `Microsoft.Graph.Reports`, `Microsoft.Graph.Identity.SignIns`; `Policy.Read.All` |
| [`ReportMemberMFA_v3.ps1`](Reports/ReportMemberMFA_v3.ps1) | Legacy versioned member MFA coverage report retained in the repository. Like v1, it reports MFA coverage across Member users (Guests excluded) and classifies each user as Modern Auth, Legacy Auth, or No MFA with inactivity-day filtering. Exports to CSV and HTML. | `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users`, `Microsoft.Graph.Reports`, `Microsoft.Graph.Identity.SignIns`; `User.Read.All`, `AuditLog.Read.All`, `UserAuthenticationMethod.Read.All`, `Policy.Read.All` |
| [`ReportMemberMFA_v4.ps1`](Reports/ReportMemberMFA_v4.ps1) | Reports MFA coverage across Member users with an `-IncludeDisabled` option to include disabled member accounts in the assessment. Supports `-Test`. Exports to CSV and HTML. | `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users`, `Microsoft.Graph.Reports`, `Microsoft.Graph.Identity.SignIns`; `User.Read.All`, `AuditLog.Read.All`, `UserAuthenticationMethod.Read.All`, `Policy.Read.All` |
| [`ReportAADAuthenticationMethods.ps1`](Reports/ReportAADAuthenticationMethods.ps1) | Reports Windows Hello for Business authentication method registrations for enabled Entra ID users. Exports to CSV. | `UserAuthenticationMethod.Read.All` |

</details>

---

### Reports — Exchange Online

All scripts in this section are located in [`Reports/`](Reports/).

<details>
<summary>View Exchange Online reporting scripts</summary>

| Script | Description | Key permissions / modules |
|--------|-------------|--------------------------|
| [`ReportMailboxQuota.ps1`](Reports/ReportMailboxQuota.ps1) | Reports on mailbox quota usage for all user mailboxes. Accepts a warning threshold percentage parameter. Exports to CSV. | `ExchangeOnlineManagement` |
| [`ReportCalendarPermissions.ps1`](Reports/ReportCalendarPermissions.ps1) | Reports on calendar permissions for all members of a specified distribution group. Outputs to the console. | `ExchangeOnlineManagement` |
| [`ReportUnusedExoMailboxes.ps1`](Reports/ReportUnusedExoMailboxes.ps1) | Finds and reports on Exchange Online mailboxes that show no recent activity, cross-referencing Entra ID sign-in data. Exports to CSV. | `ExchangeOnlineManagement`, `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users` |

</details>

---

### Reports — Intune

All scripts in this section are located in [`Reports/`](Reports/).

<details>
<summary>View Intune reporting scripts</summary>

| Script | Description | Key permissions / modules |
|--------|-------------|--------------------------|
| [`ReportIntuneApps.ps1`](Reports/ReportIntuneApps.ps1) | Reports on applications managed by Microsoft Intune. Filterable by platform (Windows, Android, iOS, or All). Exports to CSV and HTML. | `Microsoft.Graph.Authentication`; `DeviceManagementApps.Read.All`, `Group.Read.All`, `Organization.Read.All` |
| [`ReportIntuneCompliance.ps1`](Reports/ReportIntuneCompliance.ps1) | Reports tenant Intune device compliance posture including tenant compliance settings, overall and active-device compliance summaries, and per-policy compliance breakdowns. Supports optional per-device status detail export. Exports to CSV and HTML. | `Microsoft.Graph.Authentication`; `DeviceManagementConfiguration.Read.All`, `DeviceManagementManagedDevices.Read.All`, `DeviceManagementServiceConfig.Read.All`, `Organization.Read.All` |
| [`ReportIntuneMobileDevices.ps1`](Reports/ReportIntuneMobileDevices.ps1) | Reports Intune-managed Android and iOS/iPadOS devices, including OS version age against supplied supported major versions. Exports to CSV and HTML. | `Microsoft.Graph.Authentication`; `DeviceManagementManagedDevices.Read.All`, `Organization.Read.All` |
| [`ReportIntuneWindowsDevices.ps1`](Reports/ReportIntuneWindowsDevices.ps1) | Reports Intune-managed Windows devices by OS generation and build, with optional minimum supported build thresholds to flag outdated devices. Exports to CSV and HTML. | `Microsoft.Graph.Authentication`; `DeviceManagementManagedDevices.Read.All`, `Organization.Read.All` |

</details>

---

### Reports — Licensing

All scripts in this section are located in [`Reports/`](Reports/).

<details>
<summary>View licensing reporting scripts</summary>

| Script | Description | Key permissions / modules |
|--------|-------------|--------------------------|
| [`ReportLicensing.ps1`](Reports/ReportLicensing.ps1) | Reports on Microsoft 365 licensing plans (subscribed SKUs) in the tenant, including total, assigned, and available seat counts. Supports an option to exclude free and trial SKUs. Exports to CSV and HTML. | `Microsoft.Graph.Authentication`, `Microsoft.Graph.Identity.DirectoryManagement` |

</details>

---

### Reports — DNS and email security

All scripts in this section are located in [`Reports/`](Reports/).

<details>
<summary>View DNS and email security verification scripts</summary>

| Script | Description | Key permissions / modules |
|--------|-------------|--------------------------|
| [`ReportDkimRecords.ps1`](Reports/ReportDkimRecords.ps1) | Queries DKIM DNS records for all accepted domains configured in Exchange Online. Uses the Cloudflare DNS resolver (1.1.1.1). Outputs results to the console. | `ExchangeOnlineManagement` |
| [`ReportDmarcRecords.ps1`](Reports/ReportDmarcRecords.ps1) | Queries DMARC DNS records for all accepted domains configured in Exchange Online. Uses the Cloudflare DNS resolver (1.1.1.1). Outputs results to the console. | `ExchangeOnlineManagement` |
| [`ReportDomains.ps1`](Reports/ReportDomains.ps1) | Generates a comprehensive domain and DNS HTML report for all verified Microsoft 365 tenant domains. Collects authoritative name servers, SPF, MX, DKIM (selector1/selector2), and DMARC records. DNS is resolved against authoritative name servers first, falling back to Cloudflare (1.1.1.1) then Google (8.8.8.8). Supports optional RDAP/WHOIS registrar lookup. Exports to HTML. | `Microsoft.Graph.Authentication`, `Microsoft.Graph.Identity.DirectoryManagement`; `Domain.Read.All`, `Organization.Read.All` |
| [`ReportSPFRecords.ps1`](Reports/ReportSPFRecords.ps1) | Queries SPF DNS records for all accepted domains configured in Exchange Online. Uses the Cloudflare DNS resolver (1.1.1.1). Outputs results to the console. | `ExchangeOnlineManagement` |

</details>

---

### Reports — Microsoft Teams

All scripts in this section are located in [`Reports/`](Reports/).

<details>
<summary>View Microsoft Teams reporting scripts</summary>

| Script | Description | Key permissions / modules |
|--------|-------------|--------------------------|
| [`ReportMSTeamsSettings.ps1`](Reports/ReportMSTeamsSettings.ps1) | Reports tenant-wide Microsoft Teams governance and security settings and lists any Teams that have no owner. Exports CSV and HTML reports. | `MicrosoftTeams` — Global Reader, Teams Administrator, or Teams Communications Administrator |
| [`ReportTeamsGroups.ps1`](Reports/ReportTeamsGroups.ps1) | Reports on all Microsoft Teams-enabled groups in the tenant, including display name, mail nickname, visibility (Public/Private), creation date, and owner and member counts. Exports to CSV. | `Microsoft.Graph.Authentication`, `Microsoft.Graph.Groups`, `Microsoft.Graph.Teams`, `Microsoft.Graph.Identity.DirectoryManagement`; `Group.Read.All`, `GroupMember.Read.All`, `Team.ReadBasic.All` |

</details>

---

### Exchange — Calendar permission management

Scripts are located in [`Exchange/ManageCalendarPermissionsViaGroups/`](Exchange/ManageCalendarPermissionsViaGroups/).

These scripts allow bulk management of mailbox calendar permissions based on distribution group membership. They require an active Exchange Online session.

<details>
<summary>View calendar permission management scripts</summary>

| Script | Description | Key permissions / modules |
|--------|-------------|--------------------------|
| [`Add-AccountToCalendarPermissions.ps1`](Exchange/ManageCalendarPermissionsViaGroups/Add-AccountToCalendarPermissions.ps1) | Grants a specified account the given calendar permission level on the calendars of all members of a distribution group. | `ExchangeOnlineManagement` |
| [`Add-PermissionsToAll.ps1`](Exchange/ManageCalendarPermissionsViaGroups/Add-PermissionsToAll.ps1) | Grants a specified staff member access to the calendars of all user mailboxes in the tenant. | `ExchangeOnlineManagement` |
| [`Remove-AccountPermissions.ps1`](Exchange/ManageCalendarPermissionsViaGroups/Remove-AccountPermissions.ps1) | Removes calendar permissions for a specified account from all members of a distribution group. | `ExchangeOnlineManagement` |

</details>

---

### Exchange — Transport rules

Assets are located in [`Exchange/TransportRules/`](Exchange/TransportRules/).

| File | Description |
|------|-------------|
| [`QuarantineDisclaimer.html`](Exchange/TransportRules/QuarantineDisclaimer.html) | HTML template for use in an Exchange Online transport rule quarantine disclaimer notification. |

---

### Exchange — Anti-malware policy

Assets are located in [`Exchange/AntiMalwarePolicy/`](Exchange/AntiMalwarePolicy/).

| File | Description |
|------|-------------|
| [`ListOfFileTypesToBlock.txt`](Exchange/AntiMalwarePolicy/ListOfFileTypesToBlock.txt) | A reference list of file extensions to block in a Microsoft Defender for Office 365 anti-malware policy. |

---

### Incident response

Scripts are located in [`Incident_Response/`](Incident_Response/).

These scripts are intended for use during active security investigations. They require elevated permissions and should be run by an administrator with appropriate access to Exchange Online audit logs or Entra ID sign-in data.

<details>
<summary>View incident response scripts</summary>

| Script | Description | Key permissions / modules |
|--------|-------------|--------------------------|
| [`Get-AuditLogsByIP.ps1`](Incident_Response/Get-AuditLogsByIP.ps1) | Searches the Exchange Online unified audit log for activity originating from one or more specified IP addresses within a defined time window. Supports filtering by user UPN and operation type. Exports to CSV and HTML. | `ExchangeOnlineManagement` — requires audit log access |
| [`Get-InteractiveSignInUniqueIPs.ps1`](Incident_Response/Get-InteractiveSignInUniqueIPs.ps1) | Reports all unique IP addresses seen in interactive sign-in events over a configurable lookback period (1–30 days). Exports to CSV. | `Microsoft.Graph.Users` — requires sign-in log access |
| [`Get-InteractiveSignInsByIP.ps1`](Incident_Response/Get-InteractiveSignInsByIP.ps1) | Reports all interactive sign-in events matching one or more specified IP addresses over a configurable lookback period (1–30 days). Exports to CSV. | `Microsoft.Graph.Users` — requires sign-in log access |

</details>

---

## Safe usage notes

- **Scripts are provided as-is without warranty.** Review each script's content before running it in your environment.
- **Run in report-only mode first.** Scripts that support a `Mode` parameter (such as `ReviewInactiveGuestUsers.ps1` and `ReviewInactiveMemberUsers.ps1`) default to `ReportOnly`. Always review the report output before switching to `Disable` mode.
- **Validate the active tenant.** Scripts that connect to Microsoft Graph will display the active account and tenant before proceeding. Confirm these match your intended target before continuing.
- **Do not embed credentials.** Never store tenant IDs, client secrets, certificate thumbprints, or user credentials in script files or commit them to this repository.
- **Permissions.** Request only the permissions each script actually requires. Avoid broad scopes such as `Directory.ReadWrite.All`.
- **Incident response scripts.** Audit log and sign-in data queries are subject to Microsoft 365 retention limits. Ensure your tenant has the appropriate Purview or Entra ID licence for the log retention window you need.

---

## Validation and testing

- There is currently no dedicated automated test suite in this repository.
- Validate documentation updates by checking:
  - relative links resolve correctly from `README.md`;
  - script and folder references match the current repository structure;
  - module/scope references in documentation match script headers where stated.

---

## Troubleshooting

- **Module import failures:** Install missing modules from [Prerequisites](#prerequisites) in the current PowerShell user scope.
- **Graph connection prompt loops:** Re-run [`Reports/_ReadOnlyConnectionScript.ps1`](Reports/_ReadOnlyConnectionScript.ps1) and use `-Force` if you need to reset the existing context.
- **Missing output files:** Verify write access to the current working directory or provide the relevant output path parameter where supported.
- **No data returned:** Confirm the connected account has the required delegated permissions and that the selected time range is valid for available audit/sign-in retention.

---

## Known limitations

- Script output and available properties depend on Microsoft 365 service APIs and tenant licensing.
- Audit log and sign-in investigations are constrained by Microsoft 365 and Entra ID retention windows.
- Some script permission notes in this README are high-level; always treat each script header as the authoritative source.

---

## Naming conventions and repository standards

### Script naming

| Pattern | Purpose | Example |
|---------|---------|---------|
| `Report<Subject>.ps1` | Strictly read-only scripts that produce a CSV, HTML, or console summary without making any tenant changes | `ReportLicensing.ps1`, `ReportEntraIDApps.ps1` |
| `Review<Subject>.ps1` | Scripts that can both report on and make changes to tenant configuration (for example, via a `Disable` or `Remediate` mode) | `ReviewInactiveGuestUsers.ps1`, `ReviewInactiveMemberUsers.ps1` |
| `Verb-Noun.ps1` | Action scripts using approved PowerShell verbs | `Get-AuditLogsByIP.ps1`, `Add-AccountToCalendarPermissions.ps1` |

`ReviewInactiveGuestUsers.ps1` and `ReviewInactiveMemberUsers.ps1` use the `Review` prefix because they include a `Disable` mode that can modify tenant configuration.

Several scripts in this repository previously used older naming patterns (`AADAuthenticationMethods.ps1`, `MailboxQuota.ps1`, `UnusedExoMailboxes.PS1`, `Report_CalendarPermissions.ps1`, `VerifyDkimRecords.ps1`, `VerifyDmarcRecords.ps1`, `VerifySPFRecords.ps1`, `Get-NonMFAReport.ps1`). These have been renamed in a dedicated rename pass to align with the `Report<Subject>.ps1` convention. See git history for the old names.

### Folder placement rules

| Asset type | Correct location |
|-----------|-----------------|
| Reporting scripts (any workload) | `Reports/` |
| Exchange Online action scripts | `Exchange/` |
| Exchange Online supporting assets | `Exchange/<subfolder>/` |
| Defender / policy configuration assets | `Exchange/<subfolder>/` |
| Incident response scripts | `Incident_Response/` |

Reporting scripts for all workloads must live in the single top-level `Reports/` folder. Do not create per-workload report subfolders. Use filename prefixes (such as `ReportEntraID*`, `ReportExo*`, `ReportIntune*`) to group reports within that folder.

### Variable naming

Scripts use the following prefix convention for variable scope:

| Prefix | Scope |
|--------|-------|
| `S_` | Script-level variables |
| `F_` | Function-level variables |

### Script header requirements

All scripts should include:

- `#Requires -Modules <module list>` on line 1 (where modules are required)
- Comment-based help with `.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER`, and at least one `.EXAMPLE`
- `[CmdletBinding()]` immediately above `param(...)`
- `$ErrorActionPreference = 'Stop'` immediately after `param(...)`
- `$S_RequiredGraphScopes` variable listing all Graph scopes used (for Microsoft Graph scripts)

---

## Maintenance guidance

When adding a new script to this repository:

1. **Determine the correct folder.** Reporting scripts go in `Reports/`. Action scripts go in the appropriate workload folder (`Exchange/`, `Incident_Response/`, etc.).
2. **Use the correct filename.** New read-only reporting scripts must use the `Report<Subject>.ps1` format. Scripts that can both report and make tenant changes must use `Review<Subject>.ps1`. Action scripts must use `Verb-Noun.ps1` with an approved PowerShell verb.
3. **Include a complete script header.** Add `#Requires`, comment-based help, `[CmdletBinding()]`, and all required parameters with validation attributes.
4. **Declare Graph scopes.** For any Microsoft Graph script, declare all required scopes in `$S_RequiredGraphScopes`.
5. **Update this README.** Add the new script to the appropriate table in the [Script catalogue](#script-catalogue) section. Include the script name as a relative link, a concise description, and the key modules or permissions required.
6. **Avoid breaking changes.** Do not change output column names or parameter names without noting the change in the pull request summary.

When renaming existing scripts to align with the current naming convention, make the rename a dedicated change and update all README references in the same commit.

---

## Feedback and contact

Constructive feedback is welcome. Scripts are maintained in personal time and may not be updated immediately.

If you have a question or a script idea, please raise a GitHub issue or get in touch via:

- [LinkedIn](https://www.linkedin.com/in/madhuperera/)
- [Twitter / X](https://twitter.com/madhu_perera)
