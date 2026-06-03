# Known issues — Report: Member MFA

Source of truth: GitHub Issues on
[`madhuperera/Microsoft365`](https://github.com/madhuperera/Microsoft365/issues)
with label `report:member-mfa`.

This file is a snapshot for offline readers. When an issue is filed, link it
here. When it is closed, move the line to the **Resolved** section with the
release tag that fixed it.

Files in scope:

- [`ReportMemberMFA_v4.ps1`](ReportMemberMFA_v4.ps1) — current working version
- [`ReportMemberMFA_v3.ps1`](ReportMemberMFA_v3.ps1) — frozen v3 reference
- [`ReportMemberMFA_v2.ps1`](ReportMemberMFA_v2.ps1) — frozen v2 reference
- [`ReportMemberMFA_v1.ps1`](ReportMemberMFA_v1.ps1) — frozen v1 reference
- [`ReportMemberMFA-Guide.html`](ReportMemberMFA-Guide.html) — companion guide

---

## Open

### Scoring / posture matrix

- [ ] **MFA Posture naming clash.** Internal labels A/B/C/D from earlier
  drafts versus the public posture names (Fully Compliant, Weak Factor, …).
  Final naming convention to be locked.

### Coverage logic

- [ ] **Phase 2 — per-user MFA gap engine.** Today, `caCoverage = Full` is
  determined by *absence from any Ideal-policy exclusion list*. It does not
  verify whether the user is in the policy's *include* scope. A user excluded
  from every Ideal policy by virtue of not being included in the first place
  would still score Full. Work item: build a per-user inclusion evaluator that
  walks `users.includeUsers`, `groups.includeGroups`, and role-based scopes
  before declaring Full coverage.

### Permissions

- [ ] **Admins surface as `Unknown` / Access Denied.** Reading authentication
  methods for privileged accounts requires Graph permissions the running
  identity may not hold. Document the exact scope set required (likely
  `UserAuthenticationMethod.Read.All` in addition to the standard read-only
  set) and surface a clearer hint in the report when this happens.

### Documentation / clarity

- [ ] **Cards-vs-risk-score mismatch reported by client.** User flagged that
  certain card counts don't appear to line up with risk-score totals in their
  own data. Specific reproduction case still owed by the user before this can
  be triaged.

---

## Resolved

- **User table defaults to enabled accounts** — Resolved in **1.3.0**
  ([`ReportMemberMFA_v4.ps1`](ReportMemberMFA_v4.ps1)). The Active filter on
  the user table starts on a new **Enabled (any)** option, so picking a
  posture lines up with the matching summary card on first render. The old
  `All` behaviour is still available as **All (incl. Disabled)**.
- **Enabled-only by default + `-IncludeDisabled` opt-in** — Resolved in
  **1.3.0** ([`ReportMemberMFA_v4.ps1`](ReportMemberMFA_v4.ps1)). The script
  pushes `accountEnabled eq true` into the Graph `$filter` by default so
  Disabled accounts are never fetched. The Disabled card renders as
  **N/A** unless `-IncludeDisabled` is passed.
- **Coverage Gap card conflated two postures** — Resolved in **1.2.0**
  ([`ReportMemberMFA_v3.ps1`](ReportMemberMFA_v3.ps1)). `No MFA + Partial`
  (Risk Level 9) is now reported as **At Risk**, matching `No MFA + Full`
  (Risk Level 8). Coverage Gap is now reserved for `Modern Auth + Partial`
  (Risk Level 3) only. See
  [`ReportMemberMFA-CHANGELOG.md`](ReportMemberMFA-CHANGELOG.md).
