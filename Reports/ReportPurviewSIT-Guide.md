# ReportPurviewSIT.ps1 — User Guide

Reports which Microsoft Purview **Sensitive Information Types** (SITs) exist across Exchange Online,
OneDrive, SharePoint and Teams, where they live, and how confident the detections are.

Read-only. Produces a CSV and a self-contained HTML report.

> **Field reference for the output files is a separate document:**
> [`ReportPurviewSIT-OutputSchema.md`](ReportPurviewSIT-OutputSchema.md) — what every column means and
> whether it is a Purview fact, our arithmetic, or our opinion.

---

## Contents

1. [Quick start](#1-quick-start)
2. [Requirements](#2-requirements)
3. [How it works](#3-how-it-works)
4. [The two modes](#4-the-two-modes)
5. [Parameter reference](#5-parameter-reference)
6. [Choosing which types to report on](#6-choosing-which-types-to-report-on)
7. [Examples](#7-examples)
8. [Long runs, checkpoints and resuming](#8-long-runs-checkpoints-and-resuming)
9. [Verifying the assumptions](#9-verifying-the-assumptions)
10. [Troubleshooting](#10-troubleshooting)
11. [Safety, privacy and what leaves the tenant](#11-safety-privacy-and-what-leaves-the-tenant)

---

## 1. Quick start

```powershell
# See what a default run would do. No connection, works offline.
.\ReportPurviewSIT.ps1 -ListOnly

# Default run: US core identifiers, item counts across all four workloads. Minutes.
.\ReportPurviewSIT.ps1

# Item-level detail with confidence bands. Slower - one call per page.
.\ReportPurviewSIT.ps1 -Mode Full -View us-core -Workload ODB
```

Always start with `-ListOnly`. It resolves your selection against the taxonomy and prints exactly
which types would be queried, how many API calls that is, and a rough duration — without connecting to
anything.

---

## 2. Requirements

| Requirement | Detail |
|---|---|
| PowerShell | 7+ preferred. Windows PowerShell 5.1 supported. |
| Module | `ExchangeOnlineManagement` — `Install-Module ExchangeOnlineManagement -Scope CurrentUser` |
| Role | **Content Explorer List Viewer**, in the Microsoft Purview compliance portal |
| Licensing | E5 for 113 types; **Priva additionally, for named entities in Exchange** — see below |

`Content Explorer **Content** Viewer` is **not** required. This script reads counts and locations, and
never previews item content.

### Licensing changes what you can see

An unlicensed type returns **zero rather than an error**, which is indistinguishable from genuinely
having none. Two rules apply:

| Trait | Requirement | Scope |
|---|---|---|
| `requires-e5` (113 types) | Microsoft 365 E5 or equivalent compliance add-on | All credential scanning and all named entity types |
| `exo-requires-priva` (60 types) | Privacy Service Plan (Microsoft Priva), **in addition to E5** | Named entities in **Exchange at rest only** — SharePoint, OneDrive and Teams unaffected |

The script reads these rules from the taxonomy and flags a workload where every type carrying one of
these traits returned zero while other types found data. It also tells you whether those same types
work in another workload, which distinguishes "no E5 at all" from "E5 but no Priva for Exchange".

Note that *Advanced classification scanning and protection* — a tenant setting often assumed relevant
— applies only to bundled types in **Endpoint DLP** policies. It has no bearing on a Content Explorer
census.

---

## 3. How it works

```
  ┌─ Load taxonomy ────────── embedded 327-SIT tag file (or -TaxonomyPath)
  │
  ├─ Resolve selection ────── -SensitiveInfoType > -Tag > -View, then -Jurisdiction, then -ExcludeTag
  │                            └─ -ListOnly stops here. No connection made.
  │
  ├─ Connect ──────────────── Connect-IPPSSession (reuses an existing session if you confirm)
  │
  ├─ Reconcile with tenant ── Get-DlpSensitiveInformationType is authoritative
  │                            ├─ selected but absent from tenant  -> skipped, listed in Data Quality
  │                            └─ in tenant but not in taxonomy    -> listed as untagged
  │
  ├─ Collect ──────────────── Census: 1 call per SIT per workload
  │                            Full:   1 call per page, records flattened to item x SIT rows
  │                            └─ checkpoint written as it goes
  │
  └─ Output ───────────────── CSV (underlying grain) + HTML (roll-ups, charts, caveats)
```

**The taxonomy is a classification layer, never a substitute for the tenant.** The script always asks
the tenant what types exist. The bundled taxonomy only adds tags to them. This is why custom SITs
still work — they simply have empty tag columns.

**Why a taxonomy at all.** A type belongs to several independent dimensions at once. `U.S. Physical
Addresses` is simultaneously a US type, a contact type, an indirect identifier, and a noisy named
entity requiring E5. Forcing it into one category loses whichever facts you did not pick. Tags keep all
of them, so jurisdiction becomes a filter rather than a separate file.

---

## 4. The two modes

| | **Census** (default) | **Full** |
|---|---|---|
| What it reads | An aggregate count per type per workload | Every item record, paged |
| API calls | 1 per SIT per workload | 1 per **page** |
| Typical duration | Minutes | Hours for a broad selection |
| CSV grain | one row per SIT per workload | one row per **item per SIT** |
| Confidence bands | ✗ not possible | ✓ |
| Item names, labels, dates | ✗ not possible | ✓ |
| Containers affected | ✗ | ✓ |
| Resumable | ✓ | ✓ |

**Census answers "where should I look".** Full answers "what exactly is there".

Census physically cannot produce confidence bands or item detail — it never retrieves item records.
There is no approximation available; the data simply is not in the response.

### A note on confidence filtering

The script **never** passes `-ConfidenceLevel`, in either mode. Content Explorer's `TotalCount` is not
scoped by confidence, so a filtered call returns a confidence-blind total alongside a filtered record
set. Reporting that as "high confidence hits" would be wrong.

To get high-confidence figures, run `-Mode Full` and use the `HighMatches` column. That derives the
bands from each item's own match data, which is both correct and more informative — you get all three
bands from one pass instead of three filtered passes that cannot be combined.

---

## 5. Parameter reference

### Mode and output

| Parameter | Type | Values | Default | Purpose |
|---|---|---|---|---|
| `-Mode` | string | `Census`, `Full` | `Census` | Counts only, or full item-level export. |
| `-Format` | string | `Csv`, `Html`, `Both` | `Both` | Which files to write. |
| `-ReportPath` | string | folder or file path | current directory | Folder gets timestamped names; a file path is used as given. |
| `-IncludeItemNames` | switch | — | off | Write real file names and email subjects. **See §11.** |

### Selecting sensitive information types

| Parameter | Type | Values | Default | Purpose |
|---|---|---|---|---|
| `-View` | string[] | see §6.1 | `us-core` | Named tag filters. Multiple views are combined as a union. |
| `-Tag` | string[] | `dimension:value` | — | Ad hoc filter, **AND** semantics. Overrides `-View`. |
| `-Jurisdiction` | string[] | ISO 3166-1 alpha-2, `EU`, `GLOBAL` | — | Country filter, **OR** semantics. Narrows whatever else applied. |
| `-ExcludeTag` | string[] | `dimension:value` | — | Drop any type carrying one of these. Applied **last**. |
| `-SensitiveInfoType` | string[] | exact type names | — | Explicit list. Overrides everything above. |

### Scope and volume

| Parameter | Type | Values | Default | Purpose |
|---|---|---|---|---|
| `-Workload` | string[] | `EXO`, `ODB`, `SPO`, `Teams`, `All` | `All` | Where to look. `All` wins if present. |
| `-TestMode` | switch | — | off | Sample a random subset of the selected types. |
| `-SampleSize` | int | 1–500 | `10` | How many types `-TestMode` samples. |
| `-ListOnly` | switch | — | off | Dry run. **No connection, no API call.** |
| `-DelayMilliseconds` | int | 0–10000 | `200` | Pause between calls, to stay clear of throttling. |

### Long runs

| Parameter | Type | Values | Default | Purpose |
|---|---|---|---|---|
| `-CheckpointPath` | string | file path | alongside the report | Where progress is recorded. |
| `-ResumeFrom` | string | file path | — | Resume from a checkpoint. Successes reused, failures retried. |

### Taxonomy

| Parameter | Type | Default | Purpose |
|---|---|---|---|
| `-TaxonomyPath` | string | embedded copy | Use an external taxonomy JSON instead of the built-in 327-type file. |

### Assumption verification

| Parameter | Type | Values | Default | Purpose |
|---|---|---|---|---|
| `-VerifyAssumptions` | switch | — | off | Run the two probes instead of reporting, then exit. See §9. |
| `-VerifySensitiveInfoType` | string | exact type name | auto-selected | Which type to probe with. |
| `-VerifyMaxPages` | int | 1–1000 | `50` | Page ceiling per workload during a probe. |

Standard common parameters (`-Verbose`, `-ErrorAction`, …) are supported.

---

## 6. Choosing which types to report on

### 6.1 Views — the easy way

| View | Types | Use for |
|---|---|---|
| `us-core` | **7** | Direct US identifiers. Default, and the recommended DLP starting point. |
| `us-extended` | **11** | All US personal data including indirect identifiers, noisiest sources removed. Excludes named entities, so it shrank when that trait expanded. |
| `us-all-pii` | **28** | Everything personal including named entities. **Expect high false positives.** |
| `health` | **22** | Clinical identifiers and terminology, any jurisdiction. |
| `credentials` | **53** | Keys, tokens, secrets. A security finding, not a privacy finding. |
| `all` | **327** | Everything. |

```powershell
.\ReportPurviewSIT.ps1 -View us-extended
.\ReportPurviewSIT.ps1 -View us-core,health      # union: 28 types
```

### 6.2 Tags — the flexible way

Tags are namespaced `dimension:value`. Every type carries exactly one `class`, one `pii` and one
`tier`, at least one `jurisdiction`, and zero or more `trait`.

**`class`** — the kind of data

| Value | Types |
|---|---|
| `identity` | 162 |
| `credential` | 53 |
| `contact` | 51 |
| `financial` | 39 |
| `health` | 22 |

**`pii`** — does it identify a person on its own?

| Value | Types | Meaning |
|---|---|---|
| `direct` | 203 | Identifies an individual alone. |
| `indirect` | 64 | Personal only with context — addresses, diagnosis codes. |
| `none` | 60 | **Not personal data.** Routing numbers, company IDs, secrets. |

**`tier`** — reporting priority

| Value | Types |
|---|---|
| `core` | 169 |
| `extended` | 98 |
| `excluded` | 53 |
| `corroborating` | 7 |

**`trait`** — detection characteristics

| Value | Types | Why it matters |
|---|---|---|
| `requires-e5` | 113 | Needs E5 or an equivalent compliance add-on. **Returns zero, not an error, without it.** |
| `exo-requires-priva` | 60 | Named entity. Exchange at-rest classification **additionally** needs the Privacy Service Plan (Microsoft Priva). SharePoint, OneDrive and Teams are unaffected. |
| `named-entity` | 60 | Dictionary-backed. High recall, **high noise**. |
| `high-volume` | 7 | Expect very large match counts. |
| `weak-pattern` | 3 | Low precision without corroboration. |
| `checksum-validated` | 2 | Luhn or similar. **High precision — trust these.** |
| `corroborating` | 2 | Supports another type rather than standing alone. |
| `bundle` | 1 | Wraps other types. Do not report alongside its members. |
| `keyword-corroborated` | 1 | Needs nearby keywords for high confidence. |
| `format-varies-by-state` | 1 | Sub-national format variation. |
| `pairs-with:<guid>` | 1 | Precision improves alongside the referenced type. |

### `exo-requires-priva` — the trap that looks like a clean result

Named entity types (`All Full Names`, `All Physical Addresses`, the medical terminology types) are
classified at rest in Exchange **only if the tenant holds the Privacy Service Plan**, on top of E5.
Without it, Exchange returns **zero** for all 60 of them — no error, no warning.

Observed on a production tenant: a named entity type returned **thousands of items in Teams and zero
in Exchange**, while ordinary pattern types returned plenty in Exchange. Names had not vanished from
the mail; the classification simply was not licensed for that workload.

The script detects this pattern and says so — see §10. `us-core` and `us-extended` exclude named
entities anyway, so default runs are unaffected; it bites on `us-all-pii`, `health` and `all`.

**`jurisdiction`** — 64 values. The largest:

`GLOBAL` (70), `US` (12), `JP` (9), `AU` (8), `FR` (8), `AT` (7), `CA` (7), `HU` (7), `PL` (7),
`GB` (7), `EU` (6), `DE` (6), `GR` (6), `NL` (6), `NZ` (6), `ES` (6) …

…plus 48 more, down to single-type entries like `MX`, `QA`, `SA`, `TH`. Any ISO 3166-1 alpha-2 code
present in the taxonomy works.

```powershell
# AND semantics: must carry BOTH tags
.\ReportPurviewSIT.ps1 -Tag class:health,pii:direct        # 9 types
```

### 6.3 Precedence — which parameter wins

```
-SensitiveInfoType   beats   -Tag   beats   -View
                                     |
                     -Jurisdiction narrows whichever applied  (OR across the listed countries)
                                     |
                     -ExcludeTag applied last, always
```

**One special case worth knowing.** `-Jurisdiction` on its own selects from **all** types, not from the
default `us-core` view — otherwise `-Jurisdiction NZ` would intersect with a US-only view and return
nothing. Name a view explicitly and the strict intersection applies:

```powershell
.\ReportPurviewSIT.ps1 -ListOnly -Jurisdiction NZ                  # 6 types  (base widened to all)
.\ReportPurviewSIT.ps1 -ListOnly -View us-core -Jurisdiction NZ    # 0 - errors, and correctly so
```

### 6.4 One specific type

```powershell
.\ReportPurviewSIT.ps1 -SensitiveInfoType 'U.S. Social Security Number (SSN)'
```

The name must match your tenant **exactly** — no wildcards. Find it with:

```powershell
Get-DlpSensitiveInformationType | Where-Object Name -like '*Social Security*' | Select-Object Name, Id
```

Watch the punctuation: `U.S. / U.K. Passport Number`, `U.S. Driver's License Number` (the apostrophe
needs double quotes in PowerShell), `U.S. Individual Taxpayer Identification Number (ITIN)`.

This is also the **only** way to report on a custom SIT, since custom types are not in the taxonomy and
so cannot be reached by any view or tag.

---

## 7. Examples

### Getting oriented

```powershell
# What would a full-taxonomy census cost? Offline, instant.
.\ReportPurviewSIT.ps1 -ListOnly -View all

# List every type in a view, with its class and tier
.\ReportPurviewSIT.ps1 -ListOnly -View us-extended

# Find a type name
.\ReportPurviewSIT.ps1 -ListOnly -View all | Select-String 'Passport'
```

### Everyday census runs

```powershell
# Default: US core identifiers, all workloads
.\ReportPurviewSIT.ps1

# Broader US sweep, into a specific folder
.\ReportPurviewSIT.ps1 -View us-extended -ReportPath C:\Reports\ClientA

# Mail and chat only
.\ReportPurviewSIT.ps1 -View us-all-pii -Workload EXO,Teams

# Security rather than privacy: keys and secrets in document stores
.\ReportPurviewSIT.ps1 -View credentials -Workload ODB,SPO

# Whole estate, everything except secrets  (274 types)
.\ReportPurviewSIT.ps1 -View all -ExcludeTag class:credential
```

### Non-US and multinational

```powershell
# All New Zealand types  (6)
.\ReportPurviewSIT.ps1 -Jurisdiction NZ

# UK estate  (7)
.\ReportPurviewSIT.ps1 -Jurisdiction GB

# Two jurisdictions, direct identifiers only, minus the noisy named entities
.\ReportPurviewSIT.ps1 -Jurisdiction US,NZ -Tag pii:direct -ExcludeTag trait:named-entity

# Australian core identifiers  (4)
.\ReportPurviewSIT.ps1 -Tag tier:core -Jurisdiction AU

# EU-wide and global direct identifiers  (9)
.\ReportPurviewSIT.ps1 -Jurisdiction EU,GLOBAL -Tag pii:direct
```

### Cutting noise

```powershell
# Drop low-precision detections
.\ReportPurviewSIT.ps1 -View us-extended -ExcludeTag trait:weak-pattern     # 9 types

# Only checksum-validated types - the highest-precision findings available
.\ReportPurviewSIT.ps1 -Tag trait:checksum-validated

# Health data, direct identifiers only  (9)
.\ReportPurviewSIT.ps1 -Tag class:health,pii:direct
```

### Full item-level export

```powershell
# Prove the output shape on three types first - cheapest possible real run
.\ReportPurviewSIT.ps1 -Mode Full -View us-core -Workload ODB -TestMode -SampleSize 3

# One type, one workload. The sensible first real Full run.
.\ReportPurviewSIT.ps1 -Mode Full -SensitiveInfoType 'U.S. Social Security Number (SSN)' -Workload ODB

# US personal data in OneDrive, with confidence bands
.\ReportPurviewSIT.ps1 -Mode Full -View us-extended -Workload ODB

# Include real file names and subjects - handle the output carefully
.\ReportPurviewSIT.ps1 -Mode Full -View us-core -Workload EXO -IncludeItemNames
```

### Practical workflow for a new client

```powershell
# 1. Cost it, offline
.\ReportPurviewSIT.ps1 -ListOnly -View us-extended

# 2. Census the estate - minutes - to find where anything actually is
.\ReportPurviewSIT.ps1 -View us-extended -ReportPath .\ClientA

# 3. Read the census, then Full-export only the workloads that had findings
.\ReportPurviewSIT.ps1 -Mode Full -View us-extended -Workload ODB,SPO -ReportPath .\ClientA
```

Census first, Full second. It turns a speculative multi-hour export into a targeted one.

---

## 8. Long runs, checkpoints and resuming

The Security & Compliance token expires after roughly **one hour**. Anything broad will cross that.

**What the script does automatically**

- Writes a checkpoint as it goes, atomically — an interrupted run leaves a complete file, never a
  truncated one.
- Detects an expired session, reconnects, and retries.
- Retries transient failures with exponential backoff.
- Records a failure per SIT and workload rather than aborting the whole run.
- Warns up front if the estimated duration exceeds the token lifetime.

**Resuming**

```powershell
.\ReportPurviewSIT.ps1 -View all -ResumeFrom .\ReportPurviewSIT_20260802_141500.checkpoint.json
```

Successful results are reused; **recorded failures are retried**. If the checkpoint could not be
written, the run says so in red — do not rely on `-ResumeFrom` after that.

**A limitation to know about.** `-ListOnly` estimates as though the run were a census, so its figure is
**wrong for `-Mode Full`** — it reports the SIT count, not the page count. Size a Full run by censusing
first: item count ÷ 1000 ≈ pages.

---

## 9. Verifying the assumptions

Two behaviours the tool depends on. Both are recorded in the script header and surfaced in the HTML
appendix.

| | Assumption | Status | Consequence |
|---|---|---|---|
| **A1** | `TotalCount` is complete, not sampled | **Workload-dependent — not settled** | Holds in Exchange. **Differs in Teams.** It also counts *records*, not distinct items — see below. |
| **A2** | Exchange and Teams return disjoint item sets | **Weak evidence for** | If wrong, cross-workload totals **double count** Teams chat. |

**A1 result — Exchange holds, Teams does not.** In Exchange, census `TotalCount` exactly equalled the
records paged back, so sampling does **not** truncate the total there. In Teams the total exceeded the
records paged back by roughly 40%, with paging ending naturally rather than being cut short.

Either the Teams census overstates or the Teams paged export is incomplete — those implicate opposite
modes, so **do not quote Teams figures from either until it is explained**. Other workloads are fine.

Both probes also showed `TotalCount` counting folder-level occurrences rather than items: a message
filed in two folders of one mailbox counts twice. So **census figures run around 1% above the true
item count**, and census and Full legitimately disagree. Full is the precise one.

**A2 result.** A substantial Exchange set against a single Teams item, no overlap. That argues against
*systematic* overlap — had Teams chat been surfacing under Exchange as a rule, that item should have
appeared — but one sample cannot rule out *occasional* overlap. The caveat stays as a precaution.

The probe returns `WEAK` below 25 items per workload and `HOLDS` above it. A single overlapping ID is
proof at any sample size, so `FAILS` is always conclusive; only a negative result needs volume.

**A tenant with little Teams content cannot settle A2 at all**, and that is a legitimate outcome. To
find out whether yours can, census Teams on its own — 327 calls, around 15 minutes — and look at the
largest count:

```powershell
.\ReportPurviewSIT.ps1 -View all -Workload Teams
```

If nothing reaches 25, leave the caveat in place rather than inferring from thin data. Otherwise probe
with the type that scored highest:

```powershell
.\ReportPurviewSIT.ps1 -VerifyAssumptions -View us-extended
```

It pages a probe type, compares the census count against an exhaustive export (A1), and intersects the
EXO and Teams message-ID sets (A2). Verdicts are `HOLDS`, `FAILS`, `DIFFERS` or `INCONCLUSIVE` — it
reports inconclusive rather than guessing when it hits the page ceiling.

Read-only, and it reads only identifiers — never item names or content. The findings file holds counts
and verdicts only. Update the assumption register in the script header afterwards.

---

## 10. Troubleshooting

| Symptom | Cause and fix |
|---|---|
| `None of the selected SITs exist in this tenant` | Usually a misspelled `-SensitiveInfoType`. Names must match exactly; check with `Get-DlpSensitiveInformationType`. |
| `No SITs matched the selection` | Tag typo, or an impossible intersection such as `-View us-core -Jurisdiction NZ`. Run `-ListOnly -View all` to see what exists. |
| `View 'x' is not defined` | The error lists the valid view ids. |
| Everything returns 0 | Check the `Traits` column for `requires-e5`; unlicensed types return zero, not an error. Also confirm you hold **Content Explorer List Viewer**. |
| Run stops around the one-hour mark | Token expiry. Re-run with `-ResumeFrom <checkpoint>`. |
| `Checkpoint: FAILED TO WRITE` | Output folder not writable. Fix before a long run — resume will not work. |
| Types missing from the report | They were never queried. Absence from the output says nothing about them — widen the selection. |
| A row shows `err` rather than a number | That call failed. It is **unknown, not zero**. Listed under Data Quality; retried on resume. |
| Full mode is taking hours | Expected — one call per page. Census first, then Full only where findings exist. |
| Prompted to sign in mid-run | Token expired and the script reconnected. Normal on long runs. |

---

## 11. Safety, privacy and what leaves the tenant

**Read-only.** Only `Connect-*`, `Get-*` and `Export-*` cmdlets are used. Nothing in the tenant is
created, modified or deleted, and no policy or configuration is touched.

**No egress.** Nothing leaves the tenant except to your local filesystem. No telemetry, no analytics.
The HTML report makes **no external requests** — charts are inline SVG rather than a CDN-hosted
library — so it can be opened on an isolated machine and will not phone home.

**The output files are sensitive.** A Full export tells a reader exactly where sensitive data sits.

- **`ItemName` is redacted by default** to a stable hash such as `[redacted:7fb89142342a]`. Identical
  names still group together, but nothing is disclosed.
- **`-IncludeItemNames` writes the real values.** For Exchange those are **email subject lines**, which
  routinely restate the sensitive content itself — *"SSN for John Smith"*. Use it only when you need to
  chase specific items, and treat the file as containing personal data in plain text.
- Every pivot the detail file is designed for works fine on redacted data.

Store and transmit the output at the same classification as the data it describes.

---

## Related files

| File | Purpose |
|---|---|
| `ReportPurviewSIT.ps1` | The script. Comment-based help via `Get-Help .\ReportPurviewSIT.ps1 -Full`. |
| `ReportPurviewSIT-OutputSchema.md` | Field-by-field reference for every output file. |
| `sit-taxonomy.json` | The 327-type tag taxonomy. Embedded in the script; supply your own with `-TaxonomyPath`. |
