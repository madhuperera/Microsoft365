# Changelog — Report: Member MFA

All notable changes to the active `ReportMemberMFA_vN.ps1` script and its
companion guide [`ReportMemberMFA-Guide.html`](ReportMemberMFA-Guide.html).
Older revisions are kept as frozen reference copies (`_v1`, `_v2`, …)
alongside the active version.

Format: [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).
Versioning: [SemVer](https://semver.org/) — `MAJOR.MINOR.PATCH`.

Release tags are pushed as `report-member-mfa/vX.Y.Z` so this repository can
host changelogs and tags for many reports without collisions.

---

## [Unreleased]

### Planned for [1.3.0] / `ReportMemberMFA_v4.ps1`
- **Enabled-only by default + `-IncludeDisabled` opt-in.** Push
  `accountEnabled eq true` into the initial Graph `$filter` so Disabled
  accounts are never fetched and never enumerated for authentication
  methods — the largest single cost in the run today on tenants with many
  stale accounts. The summary cards already count enabled accounts only, so
  this becomes the natural default. The **Disabled** card renders as **N/A**
  with a tooltip explaining that disabled accounts were not enumerated in
  this run. Operators who want the full picture pass `-IncludeDisabled`
  to restore the v3 behaviour (fetch and enumerate everything, Disabled
  card populated).

### Known issues
- See [`ReportMemberMFA-KNOWN-ISSUES.md`](ReportMemberMFA-KNOWN-ISSUES.md) for
  open planning items.

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
