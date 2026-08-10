# ReportPurviewSIT — Output Schema

Field reference for every file `ReportPurviewSIT.ps1` writes.

The organising question this document answers is **where each number came from**. A column that
Microsoft Purview returned is a fact about your tenant. A column this script calculated is an
interpretation. A column that came from the bundled taxonomy is an editorial opinion held by this
project, not by Microsoft. Those three things look identical in a spreadsheet, so every field below is
labelled.

---

## Provenance labels

| Label | Meaning |
|---|---|
| **Purview** | Returned by `Export-ContentExplorerData` or `Get-DlpSensitiveInformationType` and passed through unchanged, apart from being renamed or reformatted. A fact about your tenant. |
| **Derived** | Calculated by this script from Purview data. Reproducible arithmetic, but arithmetic *we* chose. The formula is given for each one. |
| **Taxonomy** | This project's classification of the sensitive information type, from the embedded `sit-taxonomy.json`. **Microsoft does not publish these.** They are our judgement and you may disagree with any of them. |

A blunt summary: **Purview columns are evidence. Derived columns are analysis. Taxonomy columns are
opinion.** If a client challenges a figure, the first two are defensible from the API, the third is a
conversation about how you chose to categorise things.

---

## Files produced

| File | Written when | Grain |
|---|---|---|
| `ReportPurviewSIT_<timestamp>.csv` | always | **Full mode:** one row per item per SIT. **Census mode:** one row per SIT per workload. |
| `ReportPurviewSIT_<timestamp>.html` | `-Format Html\|Both` | Roll-ups, charts, caveats. Self-contained, no external requests. |
| `ReportPurviewSIT_<timestamp>.checkpoint.json` | always | Progress record for `-ResumeFrom`. |
| `ReportPurviewSIT_<timestamp>.assumptions.json` | `-VerifyAssumptions` only | Verdicts from the two assumption probes. |

Roll-ups appear **only** in the HTML. The CSV deliberately carries just the underlying grain, so there
is one source of truth rather than two.

---

# 1. Full mode CSV — `-Mode Full`

One row per item per sensitive information type detected on it. An item carrying four types produces
four rows; its item-level fields repeat across them. That repetition is correct for a fact table and
is what makes the file pivot cleanly.

## 1.1 Item identity and location

| Column | Source | Meaning |
|---|---|---|
| `Workload` | **Purview** | API field `Location`. One of `EXO`, `ODB`, `SPO`, `Teams`. |
| `Container` | **Purview** | API field `FileSourceUrl`. **Meaning changes by workload** — see 1.6. |
| `ItemId` | **Purview** | API field `FileUrl`. **Meaning changes by workload** — see 1.6. Not always a navigable URL. |
| `ItemName` | **Purview**, transformed by default | API field `FileName`. Redacted unless `-IncludeItemNames` — see 1.5. |
| `ItemKey` | **Derived** | Deduplication key. See below. |

### `ItemKey` — why it exists

`ItemId` is **not** a unique key for mail and chat. For Exchange and Teams it is a *message ID*, and
the same message appears once per mailbox holding a copy. An email to eight internal recipients yields
up to nine rows sharing one `ItemId`, differing only by `Container`.

```
ItemKey = ItemId                    for ODB and SPO   (one path, one file)
ItemKey = ItemId + "|" + Container  for EXO and Teams (one message, per mailbox copy)
```

This column exists so you never have to remember that rule. It is computed by the script, not returned
by the API.

**Three different questions, three different columns:**

| Question | Count distinct on |
|---|---|
| How many copies exist across the estate | `ItemKey` |
| How many distinct emails or files contain this | `ItemId` |
| How many mailboxes or sites are affected | `Container` |

These can differ by an order of magnitude — on synthetic test data the same Exchange result gave
roughly three times as many copies as distinct messages, across a handful of mailboxes. Always state
which one you are quoting.

## 1.2 The sensitive information type

| Column | Source | Meaning |
|---|---|---|
| `SensitiveInfoType` | **Purview** | Display name as your tenant reports it via `Get-DlpSensitiveInformationType`. |
| `SitId` | **Purview** | The type's GUID, from `SensitiveInfoTypesData`. Stable across all tenants for built-in types. |

`SitId` is the reliable join key. Names occasionally change; GUIDs do not. Note that a handful of
built-in types return **uppercase** GUIDs — compare case-insensitively, or those rows will silently
fail to match in Excel `VLOOKUP` or Power Query.

## 1.3 Taxonomy classification — *this project's opinion, not Microsoft's*

These six columns exist so you can filter and group by something more useful than a type name. **None
of them come from Microsoft.** They are assigned in `sit-taxonomy.json`, which covers all 327 built-in
types. Custom types in your tenant have no entry, so these columns are **empty** for them — that is
expected, not a fault.

| Column | Source | Values | Meaning |
|---|---|---|---|
| `Class` | **Taxonomy** | `identity`, `financial`, `health`, `contact`, `technical`, `credential` | The kind of data described. |
| `Pii` | **Taxonomy** | `direct`, `indirect`, `none` | Whether it identifies a person **on its own**. |
| `Tier` | **Taxonomy** | `core`, `extended`, `corroborating`, `excluded` | Reporting priority within its jurisdiction. |
| `Jurisdictions` | **Taxonomy** | ISO 3166-1 alpha-2, `EU`, or `GLOBAL`; `; ` separated | Where the type is legally meaningful. More than one is normal. |
| `Traits` | **Taxonomy** | `; ` separated, may be empty | Detection characteristics affecting noise and licensing. |
| `RecommendedConfidence` | **Taxonomy** | `High`, `Medium`, `Low` | The threshold we suggest for this type. **Advisory only — it does not filter anything and no figure in this file respects it.** |

### `Pii` is the column that prevents the most common reporting error

- `direct` — identifies an individual alone. A social security number.
- `indirect` — personal only in combination or with context. A street address, a diagnosis code.
- `none` — **not personal data at all.** Bank routing numbers, company identifiers, API keys.

Filtering to `Pii = direct` is usually what someone means by "how much PII do we have". Reporting an
ICD-10 code or an ABA routing number as though it were an identifier is the mistake this column
exists to stop.

### `Traits` worth knowing

| Trait | Consequence |
|---|---|
| `named-entity` | Dictionary-backed. High recall, **high false positives**. 60 types. |
| `requires-e5` | Needs E5 or an equivalent compliance add-on. Returns **zero, not an error**, without it — indistinguishable from genuinely having none. 113 types. |
| `exo-requires-priva` | Named entity. Exchange at-rest classification **additionally** needs the Privacy Service Plan (Priva). Without it Exchange returns zero for all 60 of these while SharePoint, OneDrive and Teams work normally. |
| `high-volume` | Expect very large match counts. |
| `weak-pattern` | Low precision without corroboration. |
| `checksum-validated` | Validated by checksum, e.g. Luhn. High precision — trust these. |
| `keyword-corroborated` | Needs nearby keywords to reach high confidence. |
| `bundle` | Wraps other types. Do not report alongside its members; you will double count. |
| `pairs-with:<guid>` | Precision improves when found with the referenced type. |

## 1.4 Confidence bands — **all Derived**

This is the most important section to understand, because the raw API values are counter-intuitive.

### What Microsoft actually returns

Each item carries a `SensitiveInfoTypesData` JSON blob, one entry per detected type:

```json
{ "Id": "a44669fe-...", "LowConfidenceMatch": 524, "MediumConfidenceMatch": 512, "HighConfidenceMatch": 504 }
```

These are **cumulative thresholds, not buckets.** Each is the count of matches at or above that
confidence level, so `Low >= Medium >= High` always holds. The example above is **not** 524+512+504 =
1,540 matches. It is **524 matches**, of which 504 are high confidence.

**Adding the three raw fields roughly triples the true total.** This is the single easiest way to
produce a badly wrong number from Content Explorer data.

### What this file gives you instead

Three **mutually exclusive** bands that can safely be summed:

| Column | Source | Formula | Meaning |
|---|---|---|---|
| `HighMatches` | **Derived** | `HighConfidenceMatch` | Matches at high confidence. |
| `MediumBandMatches` | **Derived** | `MediumConfidenceMatch − HighConfidenceMatch` | Matches at medium but **not** high. |
| `LowBandMatches` | **Derived** | `LowConfidenceMatch − MediumConfidenceMatch` | Matches at low but **not** medium. |
| `TotalMatches` | **Derived** | `HighMatches + MediumBandMatches + LowBandMatches` | All matches at any confidence. |

`HighMatches` is the only band that equals its raw field — the top threshold and the top bucket are
the same thing. The other two are subtractions.

**Guaranteed invariant, on every row without exception:**

```
HighMatches + MediumBandMatches + LowBandMatches = TotalMatches
```

In normal operation `TotalMatches` also equals the raw `LowConfidenceMatch`. It is computed from the
bands rather than copied from `Low` so that the invariant holds even when the source data is odd —
see `BandsAdjusted`.

### These are matches, not items

A single document holding forty credit card numbers contributes **1** to any item count and **40** to
`TotalMatches`. Do not describe match counts as "records affected".

### `MatchesTruncated`

| Column | Source | Type |
|---|---|---|
| `MatchesTruncated` | **Derived** | Boolean — `True` / `False` |

Set when the raw `LowConfidenceMatch` or `MediumConfidenceMatch` is **exactly 500**, which appears to
be a per-item per-type cap. When `True`, treat that row's match counts as **"at least" the value
shown**, not exact.

This is a **heuristic, not a documented limit.** A value of 524 has also been observed, so 500 is not
a hard global cap. It is flagged rather than asserted.

It is a separate boolean **by design**: writing `500+` into `TotalMatches` would turn a numeric column
into text and break every pivot in the file. The number stays a number; the caveat lives in its own
column.

### `BandsAdjusted`

| Column | Source | Type |
|---|---|---|
| `BandsAdjusted` | **Derived** | Boolean — `True` / `False` |

`True` when the API returned an ordering that should be impossible (`Low >= Medium >= High` did not
hold), which would produce a negative band. Those bands are clamped to zero and `TotalMatches`
recomputed from them, so the invariant survives.

On a `True` row, `TotalMatches` will **not** equal the raw `LowConfidenceMatch`. Expect this to be
zero rows. If it is not, treat those rows with suspicion and check whether the cumulative-threshold
behaviour still holds.

## 1.5 `ItemName` and redaction

`ItemName` is the API's `FileName` field, and it is the most sensitive column in the file.

| Workload | What `FileName` actually contains |
|---|---|
| ODB / SPO | The file name |
| EXO | **The email subject line** |
| Teams | **Always empty** |

Email subjects routinely restate the very content being searched for — *"SSN for John Smith"*, *"Wire
details, account 1234"*. So an export can contain personal data in plain text even though it never
touched a message body. File names in OneDrive and SharePoint do the same, less often.

**Default behaviour: redacted.** `ItemName` contains a stable truncated SHA-256 hash:

```
[redacted:7fb89142342a]
```

Identical names produce identical hashes, so you can still answer "how many distinct subjects" and
group by name, without any name being disclosed.

**`-IncludeItemNames` writes the real values.** Only use it when you need to chase down specific items,
and treat the resulting file as containing personal data: do not email it, do not paste rows into
chat, and store it accordingly.

Every pivot in section 3 works fine on redacted data.

## 1.6 The same columns mean different things per workload

The API returns the **same twelve properties for every workload**; only their meaning changes, because
the field names are file-oriented and get reused for mail. This is the most common source of
misreading a Content Explorer export.

| Column | ODB / SPO | EXO | Teams |
|---|---|---|---|
| `Container` | Site URL | **Mailbox address** | **Mailbox address** |
| `ItemId` | Full file path | **Message ID** (not a link) | **Message ID** (not a link) |
| `ItemName` | File name | **Email subject** | **Always empty** |

Two consequences:

- **Teams rows carry no human-readable label at all.** No subject, author, team or channel. Teams is
  reportable as *volume and exposure* only. Retrieving an actual message needs eDiscovery or Content
  Search using the message ID. Set that expectation before anyone asks "show me one".
- **Teams files are not in Teams.** Files shared in a Teams chat live in SharePoint and appear under
  `SPO`. A `Teams` row is a chat or channel *message*.

## 1.7 Labels and dates

| Column | Source | Meaning |
|---|---|---|
| `SensitivityLabel` | **Purview** | Applied sensitivity label. Empty when none. |
| `RetentionLabel` | **Purview** | Applied retention label. Empty when none. |
| `LastModifiedTime` | **Purview**, reformatted | Last modification, as `yyyy-MM-dd HH:mm:ss`. Useful for finding stale, unowned content. |

Empty label columns mean *no label applied*, not *unknown*.

## 1.8 What is deliberately not in the file

| Not included | Why |
|---|---|
| Raw `LowConfidenceMatch` / `MediumConfidenceMatch` | Cumulative and dangerous to sum. The derived bands replace them; `BandsAdjusted` flags any anomaly. |
| `UserCreated` / `UserModified` | Returned by the API but **empty for EXO and Teams**, so there is no sender attribution. `Container` (the mailbox) is the only custodian signal. Answering "who sent this" needs message trace or eDiscovery. |
| `TrainableClassifiers` | Out of scope; this tool reports sensitive information types. |
| Types that were not queried | See below. |
| Matched content itself | Never retrieved. This tool reads counts and locations only. |

### Only types you actually queried appear

Records are **item-centric**: querying one type returns every item carrying it, and each record lists
*all* classifications on that item — commonly 50 or more. It is tempting to keep those extra types
since they arrive free, but doing so would be wrong.

If you queried SSN, an item might also report a Polish national ID. You would then hold only the
subset of Polish ID items that *happen to also contain an SSN* — a biased sample presented as if it
were coverage. So rows are emitted only for the types in your selection, and every row in the file has
complete coverage for its type.

To report on a type, put it in scope: `-View`, `-Tag`, `-Jurisdiction`, or `-SensitiveInfoType`.

---

# 2. Census mode CSV — default

Census reads only an aggregate count per type per workload. It never retrieves item records, so it
**cannot** produce confidence bands, labels, dates or item identity. Those need `-Mode Full`.

| Column | Source | Meaning |
|---|---|---|
| `SitName` | **Purview** | Type display name from the tenant. |
| `SitId` | **Purview** | Type GUID. |
| `Workload` | **Purview** | `EXO`, `ODB`, `SPO` or `Teams`. |
| `TotalCount` | **Purview** | Items carrying this type in this workload. See caveats below. |
| `Class`, `Pii`, `Tier`, `Jurisdictions`, `Traits`, `RecommendedConfidence` | **Taxonomy** | As section 1.3. |
| `Error` | **Derived** | Populated when the call failed. |

### `TotalCount` caveats

- **It counts matches once per item, not per match.** One document with forty card numbers
  contributes one, not forty.
- **It counts records, not distinct items — so it is an upper bound.** Observed against a production
  tenant, August 2026: in Exchange, `TotalCount` matched an exhaustive paged export exactly, so it is
  *not* truncated by sampling there. But those records collapsed by just under 1% on deduplication.
  The export walks folders while the container field only identifies the mailbox, so one message
  filed in both Inbox and Archive is two records and one item. Use `-Mode Full` when you need exact
  item counts.
- **In Teams the two did not agree at all** — see A1 in section 4.
- **Census and Full will disagree, legitimately.** For the same type and workload, census reports
  occurrences and Full reports distinct items. Full is the more precise of the two. This is not a bug.
- **It is not scoped by confidence.** The API ignores a confidence filter when computing it, so the
  script never passes one. Every census figure means *"carrying this type at any confidence"*. There
  is no honest way to get a high-confidence figure from a census.
- **An `Error` row is not zero.** It is unknown. Do not treat a blank as a nil finding.

### `Error` is populated — now what

The type and workload pair has **no figure**. Re-run with `-ResumeFrom <checkpoint>`; recorded
failures are retried while successes are reused.

---

# 3. Using the Full CSV in Excel

The long format pivots directly. Insert a PivotTable over the whole sheet.

**Exposure by location and type**

```
Rows    : Container
Columns : SensitiveInfoType
Values  : Sum of HighMatches
```

**Defensible headline figure** — filter `Pii = direct`, then:

```
Values  : Distinct Count of ItemKey
```

Use `HighMatches` rather than `TotalMatches` for anything client-facing; low-confidence matches on
`weak-pattern` types generate most of the noise.

**Worst-affected mailboxes or sites**

```
Rows    : Container
Values  : Distinct Count of ItemKey     (descending)
```

**Stale sensitive content** — add `LastModifiedTime` grouped by year.

### Rules for any figure you quote

1. **Never sum across types for an item count.** One document carries many types, so a sum across
   types counts it repeatedly. Distinct count of `ItemKey` is the item measure.
2. **Summing matches across types is fine.** Match counts are per type per item and do not overlap.
3. **Summing across workloads** is fine for ODB, SPO and EXO. See the EXO/Teams caveat below.
4. **State which item measure you used** — copies, distinct messages, or containers.
5. **Exclude `MatchesTruncated = True` rows, or label the figure a floor.**

---

# 4. Load-bearing assumptions

Recorded in the script's header as an assumption register, and reproduced here because they affect how
much weight the numbers carry. Settle them with `-VerifyAssumptions` against a tenant.

### A1 — `TotalCount` is complete, not sampled — **workload-dependent, not settled**

Two probes against one production tenant, August 2026, disagree:

| Workload | Probe type | Result |
|---|---|---|
| **Exchange** | pattern-based | `TotalCount` **equalled** the records paged back. Records exceeded distinct items by under 1%. **HOLDS.** |
| **Teams** | named entity | `TotalCount` **exceeded** the records paged back by roughly 40%, with paging ending on its own rather than hitting the page ceiling. **DIFFERS.** |

In Exchange the assumption holds: the total is not truncated by the *"5 folders or page size,
whichever completes first"* sampling, so census figures there are not understated.

Teams is unresolved, and the two possible readings damage **opposite** outputs:

- the Teams census **overstates** → Census mode inflates Teams figures
- the Teams paged export is **incomplete** → Full mode understates Teams figures

**Until that is explained, do not quote Teams figures from either mode.** Other workloads are
unaffected. The probe now tests every workload in scope and reports the worst verdict, precisely
because an earlier single-workload test reported "confirmed" and missed this.

Separately, and seen in both probes: `TotalCount` **counts records rather than distinct items**, so
census figures are an upper bound and Census and Full legitimately disagree by around 1%.

### A2 — Exchange and Teams item sets are disjoint — **weak evidence for**

Probed August 2026 on a pattern-based type: a substantial Exchange set against a **single** Teams
item, with **no overlap**. The reference tenant held almost no Teams content carrying any sensitive
information type, so a larger sample was not available.

Two failure modes, affected differently by that sample size:

| Failure mode | Would it damage a report? | What the probe showed |
|---|---|---|
| **Systematic** — Teams chat surfaces under Exchange as a rule | Yes, materially | Argued **against**. Had it been happening, that Teams item should have appeared in the Exchange set. It did not. |
| **Occasional** — a minority of messages surface twice | Marginally | **Cannot be ruled out** at one sample. |

So the caveat remains, but as a precaution rather than a live suspicion.

The probe reports `WEAK` below 25 items per workload and `HOLDS` above it. Note the asymmetry: **a
single overlapping ID is proof of overlap at any sample size**, so a `FAILS` verdict is always
conclusive. Only the negative result needs volume.

Teams chat is stored in Exchange mailboxes, and both workloads return a mailbox address as the
container, so the same message may be reachable under both. If it is, any total that sums EXO and Teams
double counts.

Until settled, when both are in scope the run prints a warning and the HTML labels its total column
*"upper bound"*, with a matching caveat in its scope section. The CSV is unaffected either way: it
holds no cross-workload total, so there is nothing in it to overstate. **Per-workload figures are
unaffected too** — use those for anything you intend to quote, and derive any total yourself from
`Workload` if you accept the caveat.

---

# 5. What a nil result does and does not mean

A type with no rows means **nothing was detected within Purview's classification capability and the
scope of that run**. It is not a guarantee of absence. Specifically:

- **Encrypted and otherwise uninspectable items are not classified.**
- **Licensing changes what is detectable.** `requires-e5` types return zero without E5, and
  `exo-requires-priva` types return zero **in Exchange only** without the Privacy Service Plan. Both
  look identical to genuinely having none. Check `Traits` before reporting a clean result — the
  report's Data Quality section flags this pattern automatically when it detects it.
- **Types outside your selection were never queried.** Their absence says nothing whatsoever.
- **Content Explorer reflects what has been indexed**, which lags recent changes.

---

# 6. Supporting files

## `*.checkpoint.json`

Progress record, rewritten as the run proceeds and used by `-ResumeFrom`. Written atomically, so an
interrupted run leaves a complete file rather than a truncated one.

```json
{ "Generated": "2026-08-02T20:15:14.000+12:00",
  "Rows": [ { "SitName": "...", "SitId": "...", "Workload": "EXO", "TotalCount": 437, "Error": null } ] }
```

Rows with a non-null `Error` are **retried** on resume, not reused. Contains no item data.

## `*.assumptions.json`

Written by `-VerifyAssumptions` only. Verdicts are `HOLDS`, `FAILS`, `DIFFERS` or `INCONCLUSIVE`, each
with a plain-language `Detail`. Contains **counts and verdicts only** — no identifiers, no item names,
no content.

---

# 7. Read-only and no egress

The script calls only `Connect-*`, `Get-*` and `Export-*` cmdlets. Nothing in the tenant is created,
modified or deleted, and no configuration is changed.

Nothing leaves the tenant except to your local filesystem. The HTML report makes **no external
requests** — charts are inline SVG, not a charting library from a CDN — so it can be opened on an
isolated machine and will not phone home. This matters because the report describes where personal
data lives.

**The output files are a different matter.** A Full export tells a reader exactly where sensitive data
sits, and with `-IncludeItemNames` it contains personal data in plain text. Handle and store the
output at the same classification as the data it describes.
