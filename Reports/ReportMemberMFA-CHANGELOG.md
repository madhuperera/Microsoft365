# Changelog — Report: Member MFA

All notable changes to the production script
[`ReportMemberMFA_prod.ps1`](ReportMemberMFA_prod.ps1) and its companion guide
[`ReportMemberMFA-Guide.html`](ReportMemberMFA-Guide.html). New work is done in
[`ReportMemberMFA_dev.ps1`](ReportMemberMFA_dev.ps1) and promoted with
[`_PromoteReport.ps1`](_PromoteReport.ps1). Older revisions are kept as frozen
reference copies (`_v1`, `_v2`, …) alongside the production script.

Format: [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).
Versioning: [SemVer](https://semver.org/) — `MAJOR.MINOR.PATCH`.

Release tags are pushed as `report-member-mfa/vX.Y.Z` so this repository can
host changelogs and tags for many reports without collisions.

---

## [Unreleased]

### Changed
- **Two-name lifecycle (`_prod` / `_dev`).** Release 1.3.0 is now published
  as [`ReportMemberMFA_prod.ps1`](ReportMemberMFA_prod.ps1). It is identical
  to [`ReportMemberMFA_v4.ps1`](ReportMemberMFA_v4.ps1) apart from the
  filename used in the three help examples. Other scripts and projects should
  reference the `_prod` filename only. `_RunDiscovery.ps1` now calls `_prod`.
- [`ReportMemberMFA_v4.ps1`](ReportMemberMFA_v4.ps1) is unchanged and is now
  a frozen reference copy.
- [`ReportMemberMFA_dev.ps1`](ReportMemberMFA_dev.ps1) starts as a copy of
  1.3.0 (help examples renamed only) and is where new features are built.

### Added
- [`_PromoteReport.ps1`](_PromoteReport.ps1) promotes `_dev` to `_prod` and
  retires the outgoing `_prod` as the next frozen `_vN` snapshot.

### In development (`ReportMemberMFA_dev.ps1` only, not yet promoted)
- **Passkey flag.** New `HasPasskey` CSV column, appended after `RiskScore`
  (`True` / `False` / `Unknown` when methods are not readable). It is `True`
  when the user has a passkey (FIDO2) registered: a FIDO2 security key, a
  passkey in Microsoft Authenticator, or a synced passkey. `MFA_Registered`
  is unchanged, so these users still report as `Modern Auth`. The HTML report
  gains a **Passkey** column at the end of the user table, a **Passkey**
  filter, and a **Passkey Registered** card in the Auth Methods section of
  the full and Summary reports.
- **Authentication Strength check for Ideal policies.** A policy that uses
  *Require authentication strength* instead of *Require multifactor
  authentication* was already eligible for the Ideal tier, but the strength
  itself was not inspected. The strength is now resolved before tiering. The
  policy is Ideal only when every combination the strength allows is
  multifactor. If the strength allows a single-factor combination the policy
  is tiered Ignored and flagged in the Grants column; if the strength cannot
  be read the policy is capped at Acceptable. Policies that use *Require
  multifactor authentication* are unaffected.

### Known issues
- See [`ReportMemberMFA-KNOWN-ISSUES.md`](ReportMemberMFA-KNOWN-ISSUES.md) for
  open planning items.

---

## [1.3.0] — 2026-06-03

### Changed
- **Enabled-only by default.** The Graph user query now pushes
  `accountEnabled eq true` into `$filter`, so Disabled accounts are not
  fetched and not enumerated for authentication methods on a default run.
  This is the single largest performance win on tenants with many stale
  accounts. The summary cards already counted enabled accounts only, so
  this is now the natural default.
- **User table defaults to `Enabled (any)`.** The Active filter on the user
  table now starts on a new **Enabled (any)** option (everything except
  Disabled). The previous `All` behaviour is still available as
  **All (incl. Disabled)**. Picking a posture filter from the table now
  matches the corresponding summary card on the first render.
- **Disabled card renders as `N/A` by default.** When the script runs
  enabled-only, the Disabled card on the main report and the Summary file
  shows **N/A** with a tooltip *"Disabled accounts were not fetched in
  this run. Pass `-IncludeDisabled` to populate this card."*

### Added
- **`-IncludeDisabled` switch.** Opt-in flag that restores v3 behaviour:
  Disabled accounts are fetched, their auth methods enumerated, the
  Disabled card and the table populate with real values, and the user
  table shows every account by default.
  - Released as [`ReportMemberMFA_v4.ps1`](ReportMemberMFA_v4.ps1), now
    published as [`ReportMemberMFA_prod.ps1`](ReportMemberMFA_prod.ps1).
  - [`ReportMemberMFA_v3.ps1`](ReportMemberMFA_v3.ps1) frozen as reference.

---

## [1.2.0] — 2026-06-03

### Changed
- **Posture mapping (Phase 1):** `No MFA + Partial` now maps to **At Risk**
  instead of **Coverage Gap**. The user has no registered MFA method at all,
  so the threat shape is the same as `No MFA + Full`. The risk score remains
  **9** (vs **8** for `No MFA + Full`) so sort order still distinguishes the
  two states, but the surfaced posture label, colour, and remediation copy
  treat both with the same urgency. **Coverage Gap is now reserved for
  `Modern Auth + Partial` (Risk Level 3) only** — users who already have
  strong MFA and merely sit in an exclusion list of an Ideal policy.
  - [`ReportMemberMFA_v3.ps1`](ReportMemberMFA_v3.ps1) is the active script.
  - [`ReportMemberMFA_v2.ps1`](ReportMemberMFA_v2.ps1) frozen as reference.

---

## [1.1.0] — 2026-05-29

### Added
- **Save filtered HTML snapshot** button on the full report. Bakes the current
  filter state (selects, text inputs, hidden rows) into a cloned DOM and
  downloads it as a standalone `*-filtered-<timestamp>.html` so a filtered view
  can be archived or printed to PDF.
- **Companion summary file** (`<report>-Summary.html`) — Tenant Details,
  Disclaimer, and the three card sections only. No user/CA tables. Intended
  for client-facing summary delivery.
- **Companion guide** [`ReportMemberMFA-Guide.html`](ReportMemberMFA-Guide.html)
  — plain-English explanation of every summary card, the CA Policy Framework
  (Ideal / Acceptable / Ignored), the posture matrix, the evaluation flow
  chart, and a Frequently Asked Questions section. Targeted at non-technical
  readers.

### Changed
- HTML report restyled per the repository HTML documentation rules
  (`html-documentation-instructions.md`):
  - A4 print stylesheet (`@page A4 14mm 12mm`) with page-break protection on
    cards, TOC, disclaimer, and security defaults sections.
  - Filters retained in print output so a filtered PDF is achievable —
    `select`/`input`/`button` background colours frozen via
    `print-color-adjust: exact`; only hover effects are stripped.
  - Numbered TOC, single CSS block, table border + zebra striping +
    sticky-but-printable thead (`border-bottom: 2px #a52428`).
  - All technical IDs (UPN, mail address, domain, policy display name) wrapped
    in `<code>` for visual separation.

### Frozen
- Original `ReportMemberMFA.ps1` renamed to `ReportMemberMFA_v1.ps1` as a
  reference copy. No functional changes.

---

## [1.0.0] — Initial release

- First release of `ReportMemberMFA.ps1` (now `ReportMemberMFA_v1.ps1`).
- CA Policy Framework Scoring v1.0:
  - Tiers: **Ideal**, **Acceptable**, **Ignored** (only Ideal counts toward
    user coverage in v1.0).
  - 9-bucket posture matrix: Fully Compliant, Weak Factor, Coverage Gap,
    Unenforced, Weak & Gap, Weak & Unenforced, At Risk, Critical, Unknown.
  - Risk levels: 0, 2, 3, 5, 6, 7, 8, 9, 10, Unknown.
- Personas: Internal, Admins, Guests (Guests excluded from the report).
- HTML output written via `Out-File -Encoding UTF8` from a single embedded
  here-string.
