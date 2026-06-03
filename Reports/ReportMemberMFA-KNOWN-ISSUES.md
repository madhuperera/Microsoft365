# Known issues — Report: Member MFA

Source of truth: GitHub Issues on
[`madhuperera/Microsoft365`](https://github.com/madhuperera/Microsoft365/issues)
with label `report:member-mfa`.

This file is a snapshot for offline readers. When an issue is filed, link it
here. When it is closed, move the line to the **Resolved** section with the
release tag that fixed it.

Files in scope:

- [`ReportMemberMFA_v3.ps1`](ReportMemberMFA_v3.ps1) — current working version
- [`ReportMemberMFA_v2.ps1`](ReportMemberMFA_v2.ps1) — frozen v2 reference
- [`ReportMemberMFA_v1.ps1`](ReportMemberMFA_v1.ps1) — frozen v1 reference
- [`ReportMemberMFA-Guide.html`](ReportMemberMFA-Guide.html) — companion guide

---

## Open

### User experience

- [ ] **User table should default to enabled accounts only.** Today the user
  table renders every member account (Enabled and Disabled) by default, while
  the summary cards always count enabled accounts only. As a result, picking
  a posture (e.g. **At Risk**) from the table filter does not produce a row
  count that matches the corresponding card. Proposed change: filter the user
  table to `Active = Enabled` on initial render, with a clearly labelled
  control to show Disabled accounts as well. Once shipped, picking any posture
  filter will line up exactly with the matching card.

### Performance

- [ ] **Enabled-only by default + `-IncludeDisabled` opt-in (planned for v4).**
  Today the script calls `Get-MgUser -Filter "userType eq 'Member'"`, which
  returns Enabled *and* Disabled accounts, then issues one
  `/users/{id}/authentication/methods` Graph call per user. On tenants with
  many stale Disabled accounts this is the single biggest cost in the run.
  Proposed change for [`ReportMemberMFA_v4.ps1`](ReportMemberMFA_v4.ps1):
  the script defaults to **enabled accounts only** by pushing
  `accountEnabled eq true` into the Graph filter — Disabled accounts are
  never fetched and never enumerated for auth methods. Because the summary
  cards already count enabled accounts only, this is the natural default.
  The **Disabled** card on the report renders as **N/A** with a tooltip
  explaining that disabled accounts were not enumerated in this run.
  Operators who want the full picture pass an opt-in `-IncludeDisabled`
  switch, which restores the v3 behaviour (Disabled accounts fetched, auth
  methods enumerated, Disabled card populated).

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

- **Coverage Gap card conflated two postures** — Resolved in **1.2.0**
  ([`ReportMemberMFA_v3.ps1`](ReportMemberMFA_v3.ps1)). `No MFA + Partial`
  (Risk Level 9) is now reported as **At Risk**, matching `No MFA + Full`
  (Risk Level 8). Coverage Gap is now reserved for `Modern Auth + Partial`
  (Risk Level 3) only. See
  [`ReportMemberMFA-CHANGELOG.md`](ReportMemberMFA-CHANGELOG.md).
