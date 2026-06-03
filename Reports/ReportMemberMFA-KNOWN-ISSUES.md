# Known issues — Report: Member MFA

Source of truth: GitHub Issues on
[`madhuperera/Microsoft365`](https://github.com/madhuperera/Microsoft365/issues)
with label `report:member-mfa`.

This file is a snapshot for offline readers. When an issue is filed, link it
here. When it is closed, move the line to the **Resolved** section with the
release tag that fixed it.

Files in scope:

- [`ReportMemberMFA_v2.ps1`](ReportMemberMFA_v2.ps1) — current working version
- [`ReportMemberMFA_v1.ps1`](ReportMemberMFA_v1.ps1) — frozen v1 reference
- [`ReportMemberMFA-Guide.html`](ReportMemberMFA-Guide.html) — companion guide

---

## Open

### Scoring / posture matrix

- [ ] **Coverage Gap card conflates two postures.**
  `No MFA + Partial` (Risk Level 9) and `Modern Auth + Partial` (Risk Level 3)
  both currently surface as **Coverage Gap**, despite having very different
  threat profiles. The `No MFA + Partial` group is closer to **At Risk** /
  **Critical**. Decision pending: split into a new bucket, remap to At Risk,
  or remap to Critical.
  See [`ReportMemberMFA_v2.ps1`](ReportMemberMFA_v2.ps1) lines ~781–820.

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

_No items yet._
