#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Reports which Microsoft Purview Sensitive Information Types (SITs) are present across Exchange
    Online, OneDrive, SharePoint and Teams, and how many items carry each one.

.DESCRIPTION
    Connects to Security & Compliance PowerShell (Connect-IPPSSession) and runs a census over
    Microsoft Purview Content Explorer: one Export-ContentExplorerData call per selected SIT per
    selected workload with -PageSize 1, reading only the TotalCount from the summary object.

    This answers "which sensitive data types exist and where do they live" for a whole tenant without
    exporting item-level records, which keeps a tenant-wide run to minutes rather than hours.

    SIT selection is driven by an embedded tag taxonomy covering all 327 Microsoft built-in SITs.
    Because built-in SIT GUIDs are identical in every tenant, the taxonomy is portable. Tags are
    namespaced 'dimension:value' across five dimensions - jurisdiction, class, pii, tier and trait -
    so a SIT can be selected by country, by data class, by priority tier, or by any combination,
    rather than being forced into a single category. Named tag filters ship as "views".

    Two modes. Census (default) reads item counts only and is fast. Full pages the item records back
    and produces the item-level detail file, at the cost of one API call per page rather than one per
    SIT per workload.

    Results are written to a single CSV and to a self-contained HTML report. The division of labour is
    deliberate: the HTML carries every roll-up and cross-tab, and the CSV carries only the underlying
    grain that the roll-ups were computed from. Shipping the roll-ups as CSVs too would duplicate the
    HTML and invite two sources of truth.
    The HTML makes no external requests: charts are hand-rolled inline SVG, so the report can be
    opened on an isolated machine and never phones home. This matters because the report describes
    the location of real personal data.

    READ-ONLY. The script calls only Connect-*, Get-* and Export-* cmdlets. Nothing in the tenant is
    created, modified or deleted.

    NO EGRESS. Nothing leaves the tenant except to the local filesystem. There is no telemetry and no
    outbound network call beyond the Microsoft 365 endpoints the modules already use.

.PARAMETER Mode
    Census : item counts only, one API call per SIT per workload. Fast, and answers "which sensitive
             data types exist and where". Cannot produce confidence bands, labels or item detail,
             because it never retrieves item records. This is the default.
    Full   : pages the item records back and writes the item-level detail CSV, one row per item per
             sensitive information type. One API call per page, so a broad selection is a multi-hour
             run. Checkpointed and resumable.

.PARAMETER IncludeItemNames
    Writes real item names to the detail CSV. Off by default, and deliberately so: for Exchange the
    item name is the email subject line, which routinely restates the sensitive content itself
    ("SSN for John Smith", "Wire details, account 1234"), and file names in OneDrive and SharePoint
    frequently do the same. Without this switch the ItemName column carries a stable truncated hash
    instead, which still groups identical names together but discloses nothing. Every pivot the
    detail file is designed for works without real names; only chasing down one specific item needs
    them, which is a narrower and more deliberate act.

.PARAMETER View
    One or more named tag filters from the taxonomy. Supplied views:

      us-core       Direct US identifiers. Default reporting set and recommended DLP starting point
      us-extended   All US personal data including indirect identifiers, noisiest sources removed
      us-all-pii    Everything personal including named entities. Expect high false positives
      health        Clinical identifiers and terminology, any jurisdiction
      credentials   Security finding rather than privacy finding
      all           Every SIT known to the tenant

    Multiple views are combined as a union. Defaults to us-core.

.PARAMETER Tag
    Ad hoc tag filter using AND semantics: a SIT must carry every tag listed. Overrides -View.
    Example: -Tag jurisdiction:US,pii:direct

.PARAMETER Jurisdiction
    Convenience filter for one or more jurisdictions, combined with OR semantics. Accepts ISO 3166-1
    alpha-2 country codes, EU, or GLOBAL, with or without the 'jurisdiction:' prefix. Applied on top
    of -View or -Tag. Example: -Jurisdiction US,NZ

.PARAMETER ExcludeTag
    Tags to exclude. A SIT carrying any of these is dropped. Applied last, after every other
    selection rule. Example: -ExcludeTag trait:named-entity,trait:high-volume

.PARAMETER SensitiveInfoType
    Explicit SIT names to report on. Overrides -View, -Tag and -Jurisdiction. Names must match the
    tenant exactly as returned by Get-DlpSensitiveInformationType.

.PARAMETER Workload
    Which workloads to census: EXO, ODB, SPO, Teams, or All. Accepts multiple values. Defaults to All.

    Note that Teams files live in SharePoint and surface under SPO. A Teams census covers chat and
    channel messages only.

.PARAMETER ReportPath
    Folder or file path for the output. If a folder is given, timestamped filenames are generated
    inside it. Defaults to the current directory.

.PARAMETER Format
    Which outputs to produce: Csv, Html or Both. Defaults to Both.

.PARAMETER TaxonomyPath
    Path to an external taxonomy JSON file, replacing the embedded copy. Use this to add a
    jurisdiction or retag SITs without editing the script.

.PARAMETER TestMode
    Censuses a random sample of the selected SITs (see -SampleSize) instead of all of them. Use this
    to validate the report format before committing to a full tenant run.

.PARAMETER SampleSize
    Number of SITs sampled when -TestMode is specified. Defaults to 10.

.PARAMETER ListOnly
    Dry run. Resolves the SIT selection from the taxonomy, prints the SITs that would be queried, the
    number of API calls and a rough duration estimate, then exits. Makes no connection and no API
    call, so it works offline.

.PARAMETER DelayMilliseconds
    Pause between Content Explorer calls, to stay clear of service-side throttling. Defaults to 200.

.PARAMETER CheckpointPath
    Path for the checkpoint file. Defaults to a file alongside the report. The checkpoint is written
    as the run progresses so a failed or expired run can be resumed with -ResumeFrom.

.PARAMETER ResumeFrom
    Path to a checkpoint file from an earlier run. SIT and workload pairs already recorded are
    skipped and their counts reused.

.EXAMPLE
    .\ReportPurviewSIT.ps1

    Default run: US core direct identifiers, censused across all four workloads.

.EXAMPLE
    .\ReportPurviewSIT.ps1 -ListOnly -View all

    Shows how many API calls a full-taxonomy census would make and roughly how long it would take,
    without connecting to anything.

.EXAMPLE
    .\ReportPurviewSIT.ps1 -View us-all-pii -Workload EXO,Teams -ReportPath C:\Reports

    Everything personal in a US estate, mail and chat only.

.EXAMPLE
    .\ReportPurviewSIT.ps1 -Jurisdiction US,NZ -Tag pii:direct -ExcludeTag trait:named-entity

    Every direct personal identifier for two jurisdictions, with the noisy named-entity types removed.

.EXAMPLE
    .\ReportPurviewSIT.ps1 -View all -ResumeFrom .\ReportPurviewSIT_20260802_141500.checkpoint.json

    Resumes a full-taxonomy census that stopped when the Security & Compliance token expired.

.NOTES
    Required role : Content Explorer List Viewer (Microsoft Purview compliance portal).
                    Content Explorer Content Viewer is NOT required - this script reads counts only
                    and never previews item content.
    Required module : ExchangeOnlineManagement

    Licensing affects results. Several SITs, particularly the named-entity types, need E5 or the
    Information Protection and Governance add-on. On an E3 tenant those SITs return zero rather than
    an error, which is indistinguishable from genuinely having none. The report flags this.

    Counting caveats, stated here because they are the main way a report like this misleads:

      - TotalCount is NOT scoped by confidence level, so this script never applies -ConfidenceLevel.
        Every figure it reports is "items carrying this SIT at any confidence".
      - TotalCount is a count of ITEMS, not of matches. One item holding forty credit card numbers
        counts once.
      - Item counts must not be summed across SITs. A single document commonly carries dozens of
        SITs, so a sum double counts. Per-SIT and per-workload figures are exact; any roll-up across
        SITs is labelled as a sum of overlapping counts, not a distinct item count.
      - Summing across workloads IS valid, because the four workloads are disjoint.

    Confidence-band reporting and per-item detail require a full paged export rather than a census,
    and are a planned addition. This script deliberately reports only what a census can support.

    ASSUMPTION REGISTER
    -------------------
    Two load-bearing assumptions have NOT been confirmed against a live tenant. Both are recorded here
    so the next reader knows what is established and what is merely believed. Run
    -VerifyAssumptions against a real tenant to settle them, then update the status lines below.

      A1. TotalCount is complete, not sampled.       STATUS: WORKLOAD-DEPENDENT. NOT SETTLED.
          Two probes against one production tenant, August 2026, disagree:

            Pattern-based SIT / EXO                 several pages, ceiling not hit
                TotalCount EQUALLED records paged back   -> HOLDS
                records exceeded distinct items by <1%

            Named entity SIT / Teams                several pages, ceiling not hit
                TotalCount EXCEEDED records paged back
                by roughly 40%                           -> DIFFERS
                records exceeded distinct items by <1%

          Exchange matched exactly. Teams returned little over half of what its TotalCount claimed,
          and paging ended on its own rather than hitting the page ceiling, so it is not a truncation
          artefact of this script. Either the Teams census overstates, or the Teams paged export is
          incomplete. Both readings are damaging and they point at opposite outputs:
              - census overstates  -> Census mode inflates Teams figures
              - paging incomplete  -> Full mode UNDERSTATES Teams figures
          UNTIL THIS IS UNDERSTOOD, DO NOT QUOTE TEAMS FIGURES FROM EITHER MODE.

          This is why the probe now tests EVERY workload in scope and takes the worst verdict. An
          earlier version tested one workload and reported it as if it settled the assumption
          everywhere, which is how the Teams discrepancy went unnoticed.

          Separately, and confirmed in both probes: TOTALCOUNT COUNTS RECORDS, NOT DISTINCT ITEMS.
          Records sharing both FileUrl and FileSourceUrl - the same item returned more than once,
          because the export walks FOLDERS while FileSourceUrl only identifies the mailbox or site -
          collapse on deduplication. A message filed in Inbox and Archive is two records, one item.
          So census figures are an upper bound, and Census and Full legitimately disagree by around
          0.5-1% even where A1 holds.

      A2. EXO and Teams item sets are disjoint.        STATUS: WEAK EVIDENCE FOR, August 2026
          Probed on a pattern-based SIT: a substantial Exchange set against a SINGLE Teams item, with
          no overlap between them. The reference tenant held almost no Teams content carrying any
          SIT, so a larger sample was not available there.

          Read that result carefully, because the two failure modes differ:
            - SYSTEMATIC overlap - Teams chat surfacing under EXO as a rule - is the one that would
              materially damage a report. If it were happening, that single Teams item should have
              appeared in the Exchange set. It did not, which argues against it.
            - OCCASIONAL overlap affecting a minority of messages cannot be ruled out at n=1.

          So the caveat stays, but it is a precaution rather than a live suspicion. The probe reports
          verdict WEAK below 25 items per side, and only HOLDS above it. Note that a single
          overlapping ID is proof of overlap at any sample size, so a FAILS verdict is always
          conclusive; only the negative result needs volume.

          The open question: Teams chat is stored in Exchange mailboxes, and both the EXO and Teams
          record schemas return a mailbox address in FileSourceUrl, so the same message may be
          reachable under both workloads. If it is, ItemsTotal double counts Teams chat whenever EXO
          and Teams are both in scope. Until settled, ItemsTotalBasis reports 'may-overlap-exo-teams'
          for those runs and the HTML carries a matching caveat.

.PARAMETER VerifyAssumptions
    Runs the two assumption probes above against the connected tenant instead of producing a report,
    then exits. Read-only, and deliberately privacy-preserving: it pages item records in order to
    count and compare identifiers, but never writes, prints or persists item names, subjects or
    match content. Only counts reach the console and the findings file.

.PARAMETER VerifySensitiveInfoType
    SIT to probe with -VerifyAssumptions. Defaults to automatically choosing one with a workable,
    non-zero item count in the workloads being tested.

.PARAMETER VerifyMaxPages
    Page ceiling per workload during -VerifyAssumptions, so a probe against a large tenant stays
    bounded. Defaults to 50. A probe that hits the ceiling reports inconclusive rather than guessing.

    Reference:
      Export-ContentExplorerData - https://learn.microsoft.com/powershell/module/exchangepowershell/export-contentexplorerdata
      Content Explorer permissions - https://learn.microsoft.com/purview/microsoft-365-compliance-center-permissions
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('Census', 'Full')]
    [string]$Mode = 'Census',

    [Parameter(Mandatory = $false)]
    [switch]$IncludeItemNames,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string[]]$View = @('us-core'),

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string[]]$Tag,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string[]]$Jurisdiction,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string[]]$ExcludeTag,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string[]]$SensitiveInfoType,

    [Parameter(Mandatory = $false)]
    [ValidateSet('EXO', 'ODB', 'SPO', 'Teams', 'All')]
    [string[]]$Workload = @('All'),

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ReportPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Csv', 'Html', 'Both')]
    [string]$Format = 'Both',

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$TaxonomyPath,

    [Parameter(Mandatory = $false)]
    [switch]$TestMode,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 500)]
    [int]$SampleSize = 10,

    [Parameter(Mandatory = $false)]
    [switch]$ListOnly,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 10000)]
    [int]$DelayMilliseconds = 200,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$CheckpointPath,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ResumeFrom,

    [Parameter(Mandatory = $false)]
    [switch]$VerifyAssumptions,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$VerifySensitiveInfoType,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 1000)]
    [int]$VerifyMaxPages = 50
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Script-level settings
# ---------------------------------------------------------------------------

# Roles the operator needs in the Microsoft Purview compliance portal. Declared at script level so
# the repository can be scanned for the permissions every script requires.
$S_RequiredPurviewRoles = @(
    'Content Explorer List Viewer'
)

# Pause between Content Explorer calls. Surfaced as -DelayMilliseconds; held in a script variable so
# it can be raised centrally if the service starts throttling.
$S_RequestDelayMilliseconds = $DelayMilliseconds

# Checkpoint is flushed to disk every this many completed SIT/workload pairs.
$S_CheckpointEvery = 10

# Nominal seconds per Content Explorer call, used only for the -ListOnly duration estimate before any
# real timing is available. Once the run starts, the estimate is replaced by the measured average.
$S_NominalSecondsPerCall = 2.5

# Security & Compliance tokens expire after roughly an hour. Warn when a run is likely to cross that.
$S_TokenLifetimeSeconds = 3600

$S_MaxRetriesPerCall = 2

# Page size for Full mode. 1000 balances call count against response size; the service caps at 10000.
$S_FullPageSize = 1000

# ---------------------------------------------------------------------------
# Embedded taxonomy
#
# All 327 Microsoft built-in SITs, tagged. Built-in SIT GUIDs are identical in every tenant (all are
# Publisher = Microsoft Corporation, IsOutOfBox = True), which is what makes shipping this worthwhile.
# Stored minified and ASCII-escaped so the blob is encoding-safe on both PowerShell 7 and Windows
# PowerShell 5.1. Override it with -TaxonomyPath rather than editing it here.
# ---------------------------------------------------------------------------
$S_EmbeddedTaxonomyJson = @'
{"schemaVersion":"2.0","description":"Tag-based classification of Microsoft Purview built-in sensitive information types. GUIDs are Microsoft built-in identifiers and are identical across all tenants. Tenant custom SITs will not appear here and are reported as untagged at runtime.","sourceReference":"https://learn.microsoft.com/purview/sit-sensitive-information-type-entity-definitions","generated":"2026-08-02","totalSits":327,"tagDimensions":{"jurisdiction":{"description":"ISO 3166-1 alpha-2 country code, EU, or GLOBAL where the type is not country specific. A SIT may carry more than one.","example":"jurisdiction:US"},"class":{"description":"Kind of data described.","values":{"identity":"Government or organisational identifiers for a person","financial":"Accounts, cards, tax and banking identifiers","health":"Clinical data, medical identifiers and terminology","contact":"Physical addresses and contact details","technical":"Network and device identifiers","credential":"Keys, tokens, passwords and connection strings"}},"pii":{"description":"Whether the type identifies a person by itself.","values":{"direct":"Identifies an individual on its own","indirect":"Personal only in combination or with context","none":"Not personal data. Organisational identifiers, codes and secrets"}},"tier":{"description":"Reporting priority within its jurisdiction.","values":{"core":"Primary identifier. Default reporting set","extended":"Genuine PII, secondary priority","corroborating":"Raises confidence in another type rather than standing alone","excluded":"Not a privacy finding. Credentials and secrets"}},"trait":{"description":"Detection characteristics affecting noise and licensing.","values":{"named-entity":"Dictionary-backed bundle. High recall, high noise","requires-e5":"Requires a Microsoft 365 E5 licence (or equivalent compliance add-on). Applies to all credential scanning SITs and all named entity SITs.","high-volume":"Expect very large match counts","weak-pattern":"Low precision without corroboration","checksum-validated":"Validated by checksum such as Luhn. High precision","keyword-corroborated":"Requires nearby keywords to reach high confidence","corroborating":"Intended to support another SIT","bundle":"Wraps other SITs. Do not select alongside its members","format-varies-by-state":"Sub-national format variation","pairs-with:<guid>":"Precision improves when combined with the referenced SIT","exo-requires-priva":"Named entity. At-rest classification in Exchange Online additionally requires the Privacy Service Plan (Microsoft Priva). Without it, Exchange returns zero for this SIT, which is indistinguishable from a genuine nil result."}}},"views":[{"id":"us-core","name":"US Core PII","description":"Direct personal identifiers for a US estate. Default reporting set and recommended DLP starting point.","filter":{"allOf":[{"anyOf":["jurisdiction:US","jurisdiction:GLOBAL"]},"pii:direct","tier:core"]}},{"id":"us-extended","name":"US Extended PII","description":"All US personal data including indirect identifiers, with the highest noise sources removed.","filter":{"allOf":[{"anyOf":["jurisdiction:US","jurisdiction:GLOBAL"]},{"anyOf":["pii:direct","pii:indirect"]}],"noneOf":["class:credential","trait:named-entity","trait:high-volume"]}},{"id":"us-all-pii","name":"US All PII","description":"Everything personal for a US estate including noisy named entities. Expect high false positive rates.","filter":{"allOf":[{"anyOf":["jurisdiction:US","jurisdiction:GLOBAL"]},{"anyOf":["pii:direct","pii:indirect"]}],"noneOf":["class:credential"]}},{"id":"health","name":"Health data","description":"Clinical identifiers and terminology, any jurisdiction.","filter":{"allOf":["class:health"]}},{"id":"credentials","name":"Credentials and secrets","description":"Security finding rather than a privacy finding.","filter":{"allOf":["class:credential"]}},{"id":"all","name":"Everything","description":"Every SIT known to the tenant.","filter":{}}],"sits":[{"id":"cb353f78-2b72-4c3c-8827-92ebe4f69fdf","name":"ABA Routing Number","recommendedConfidence":"Medium","tags":["class:financial","jurisdiction:US","pii:none","tier:corroborating","trait:corroborating"],"note":"American Bankers Association routing number. US specific. Corroborates U.S. Bank Account Number rather than standing alone."},{"id":"55e47f22-402c-486f-9779-f82430d2eee4","name":"ASP.NET Machine Key","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"ba147552-2ed1-4a8f-a441-981e94ab1895","name":"All Credential Types","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"50b8b56b-4ef8-44c2-a924-03374f5831ce","name":"All Full Names","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:GLOBAL","pii:direct","tier:extended","trait:exo-requires-priva","trait:high-volume","trait:named-entity","trait:requires-e5","trait:weak-pattern"],"note":"Matches almost any two capitalised tokens. Expect very high volume."},{"id":"065bdd91-ef07-40d3-b8a4-0aea722eaa49","name":"All Medical Terms And Conditions","recommendedConfidence":"High","tags":["class:health","jurisdiction:GLOBAL","pii:indirect","tier:extended","trait:bundle","trait:exo-requires-priva","trait:high-volume","trait:named-entity","trait:requires-e5"],"note":"Bundle of the medical term SITs. Selecting bundle and members together double counts."},{"id":"8548332d-6d71-41f8-97db-cc3b5fa544e6","name":"All Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:GLOBAL","pii:indirect","tier:extended","trait:exo-requires-priva","trait:high-volume","trait:named-entity","trait:requires-e5"]},{"id":"881b8266-702b-4ebc-9c27-ea38c3765ab5","name":"Amazon S3 Client Secret Access Key","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"eefbb00e-8282-433c-8620-8f1da3bffdb2","name":"Argentina National Identity (DNI) Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:AR","pii:direct","tier:core"]},{"id":"98da3da1-9199-4571-b7c4-b6522980b507","name":"Argentina Unique Tax Identification Key (CUIT/CUIL)","recommendedConfidence":"Medium","tags":["class:financial","jurisdiction:AR","pii:direct","tier:core"]},{"id":"74a54de9-2a30-4aa0-a8aa-3d9327fc07c7","name":"Australia Bank Account Number","recommendedConfidence":"Medium","tags":["class:financial","jurisdiction:AU","pii:direct","tier:core"]},{"id":"1cbbc8f5-9216-4392-9eb5-5ac2298d1356","name":"Australia Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:AU","pii:direct","tier:core"]},{"id":"104a99a0-3d3b-4542-a40d-ab0b9e1efe63","name":"Australia Medical Account Number","recommendedConfidence":"High","tags":["class:health","jurisdiction:AU","pii:direct","tier:extended"]},{"id":"29869db6-602d-4853-ab93-3484f905df50","name":"Australia Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:AU","pii:direct","tier:core"]},{"id":"d1d18c85-1203-46f5-b32f-2d6309de4e5b","name":"Australia Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:AU","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"e29bc95f-ff70-4a37-aa01-04d17360a4c5","name":"Australia Tax File Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:AU","pii:direct","tier:core"]},{"id":"76e83b3b-01ee-4530-aced-e667a6609f49","name":"Australian Business Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:AU","pii:none","tier:corroborating"]},{"id":"b1fba4f7-7b3e-4bb9-8f9a-9366df604dbb","name":"Australian Company Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:AU","pii:none","tier:corroborating"]},{"id":"682f18ce-44eb-482b-8198-2bcb96a0761e","name":"Austria Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:AT","pii:direct","tier:core"]},{"id":"5ec06c3b-007e-4820-8343-7ff73b889735","name":"Austria Identity Card","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:AT","pii:direct","tier:core"]},{"id":"1c96ae4e-303b-447d-86c7-77113ac266bf","name":"Austria Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:AT","pii:direct","tier:core"]},{"id":"286d9b20-eb7a-4ad8-a7fa-b7eccc0eade5","name":"Austria Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:AT","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"6896a906-86c9-4d19-a2da-6e43ccd19b7b","name":"Austria Social Security Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:AT","pii:direct","tier:core"]},{"id":"4fd58d22-af28-4451-b18a-6f722430a56d","name":"Austria Tax Identification Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:AT","pii:direct","tier:core"]},{"id":"b6a3eda2-c56c-4b69-a5f7-dca34db00f48","name":"Austria Value Added Tax (VAT) Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:AT","pii:direct","tier:extended"]},{"id":"ad0a2311-d91f-468c-bd00-925f8991d3e1","name":"Azure App Service Deployment Password","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"933284b2-12cf-4bd1-86df-c5300681ab79","name":"Azure Batch Shared Access Key","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"37e2bc99-1346-4a68-bd16-267e784711d1","name":"Azure Bot Framework Secret Key","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"03ffa580-cd25-4df2-9a3e-fb9622cdeac1","name":"Azure Bot Service App Secret","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"aa9b9abb-0c21-4e27-8393-cc1a8c9c407f","name":"Azure COSMOS DB Account Access Key","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"954d61cd-e0d0-4466-9bfe-f54b4e9714df","name":"Azure Cognitive Search API Key","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"91704742-2d51-4c3e-928b-f0600b55edb6","name":"Azure Cognitive Service Key","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"c5dcd46a-068f-4e19-a093-c91b4e19055a","name":"Azure Container Registry Access Key","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"b70cb1d1-3c98-436f-b63e-f61f3503e407","name":"Azure Databricks Personal Access Token","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"2d9471ab-72e3-4512-99ee-652482362107","name":"Azure DevOps App Secret","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"04749030-d045-4a81-905a-dd4e48f880a2","name":"Azure DevOps Personal Access Token","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"0f587d92-eb28-44a9-bd1c-90f2892b47aa","name":"Azure DocumentDB Auth Key","recommendedConfidence":"High","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"ccb6a7c2-1693-487b-8912-90788d74b906","name":"Azure EventGrid Access Key","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"f03131f9-b7f2-4ec7-856d-39e814614334","name":"Azure Function Master / API Key","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"ce1a126d-186f-4700-8c0c-486157b953fd","name":"Azure IAAS Database Connection String and Azure SQL Connection String","recommendedConfidence":"High","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"0b34bec3-d5d6-4974-b7b0-dcdb5c90c29d","name":"Azure IoT Connection String","recommendedConfidence":"High","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"a24bd725-5c89-4705-a406-6c3cf7535e45","name":"Azure IoT Shared Access Key","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"04a87f1b-b1fe-4f2f-8898-4d32d968372a","name":"Azure Logic App Shared Access Signature","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"209902e2-e787-4c1f-9fd3-41943cb02e95","name":"Azure Machine Learning Web Service API Key","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"93e77273-df14-4d35-ab29-77c606b6589f","name":"Azure Maps Subscription Key","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"75f4cc8a-a68e-49e5-89ce-fa8f03d286a5","name":"Azure Publish Setting Password","recommendedConfidence":"High","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"095a7e6c-efd8-46d5-af7b-5298d53a49fc","name":"Azure Redis Cache Connection String","recommendedConfidence":"High","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"b96e960c-b899-4fbf-9a38-5b8e91c1179a","name":"Azure Redis Cache Connection String Password","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"4d235014-e564-47f4-a6fb-6ebb4a826834","name":"Azure SAS","recommendedConfidence":"High","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"52974858-9b9c-47bb-99d0-67fa1968c3c5","name":"Azure SQL Connection String","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"b9a6578f-a83f-4fcd-bf44-2130bae49a6f","name":"Azure Service Bus Connection String","recommendedConfidence":"High","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"f4ac21a7-a3d9-4291-9094-f54eb94bff5c","name":"Azure Service Bus Shared Access Signature","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"287cbe68-26fb-42cf-993a-78890b46a1fa","name":"Azure Shared Access Key / Web Hook Token","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"158bc166-89fe-44ff-bb4c-0704f8b3f2dd","name":"Azure SignalR Access Key","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"5500fb81-47f7-471d-9de3-2b781487342d","name":"Azure Storage Account Access Key","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"c7bc98e8-551a-4c35-a92d-d2c8cda714a7","name":"Azure Storage Account Key","recommendedConfidence":"High","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"7ff41bd0-5419-4523-91d6-383b3a37f084","name":"Azure Storage Account Key (Generic)","recommendedConfidence":"High","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"7a47fadd-6969-4e19-af75-33e6d584d1be","name":"Azure Storage Account Shared Access Signature","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"f522138f-ba06-4c00-8db0-dea5c22f6f58","name":"Azure Storage Account Shared Access Signature for High Risk Resources","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"9b9c114b-1ee1-4ce2-a712-d0117d57dbeb","name":"Azure Subscription Management Certificate","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"d89fd329-9324-433c-b687-2c37bd5166f3","name":"Belgium Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:BE","pii:direct","tier:core"]},{"id":"fb969c9e-0fd1-4b18-8091-a2123c5e6a54","name":"Belgium National Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:BE","pii:direct","tier:core"]},{"id":"d7b1315b-21ca-4774-a32a-596010ff78fd","name":"Belgium Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:BE","pii:direct","tier:core"]},{"id":"08e0fc0e-def1-49c3-bc4f-f432650d9138","name":"Belgium Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:BE","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"85b5b3c3-f2de-4ae8-ac46-fd3cb38bf9ed","name":"Belgium Value Added Tax Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:BE","pii:direct","tier:extended"]},{"id":"fc564df4-201b-4405-b6da-54a2501c198f","name":"Blood Test Terms","recommendedConfidence":"High","tags":["class:health","jurisdiction:GLOBAL","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"321c43ef-82f4-4b68-be93-179f54b42dab","name":"Brand Medication Names","recommendedConfidence":"High","tags":["class:health","jurisdiction:GLOBAL","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"78e09124-f2c3-4656-b32a-c1a132cd2711","name":"Brazil CPF Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:BR","pii:direct","tier:core"]},{"id":"9b58b5cd-5e90-4df6-b34f-1ebcc88ceae4","name":"Brazil Legal Entity Number (CNPJ)","recommendedConfidence":"High","tags":["class:identity","jurisdiction:BR","pii:none","tier:corroborating"]},{"id":"486de900-db70-41b3-a886-abdf25af119c","name":"Brazil National ID Card (RG)","recommendedConfidence":"High","tags":["class:identity","jurisdiction:BR","pii:direct","tier:core"]},{"id":"b37568b5-499c-416b-8785-f72b791fcc5a","name":"Brazil Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:BR","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"66d39258-94c2-43b2-804b-aa312258e54b","name":"Bulgaria Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:BG","pii:direct","tier:core"]},{"id":"f7172b82-c588-4216-845e-4e54e397f29a","name":"Bulgaria Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:BG","pii:direct","tier:core"]},{"id":"a6fc6e44-c29b-49cf-814f-75178ef9ac30","name":"Bulgaria Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:BG","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"100d58b1-0a35-4fb1-aa89-e4a86fb53fcc","name":"Bulgaria Uniform Civil Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:BG","pii:direct","tier:core"]},{"id":"552e814c-cb50-4d94-bbaa-bb1d1ffb34de","name":"Canada Bank Account Number","recommendedConfidence":"Medium","tags":["class:financial","jurisdiction:CA","pii:direct","tier:core"]},{"id":"37186abb-8e48-4800-ad3c-e3d1610b3db0","name":"Canada Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:CA","pii:direct","tier:core"]},{"id":"59c0bf39-7fab-482c-af25-00faa4384c94","name":"Canada Health Service Number","recommendedConfidence":"Medium","tags":["class:health","jurisdiction:CA","pii:direct","tier:extended"]},{"id":"14d0db8b-498a-43ed-9fca-f6097ae687eb","name":"Canada Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:CA","pii:direct","tier:core"]},{"id":"722e12ac-c89a-4ec8-a1b7-fea3469f89db","name":"Canada Personal Health Identification Number (PHIN)","recommendedConfidence":"Medium","tags":["class:health","jurisdiction:CA","pii:direct","tier:core"]},{"id":"4de48f8f-afc8-438d-9e7f-787188178d8c","name":"Canada Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:CA","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"a2f29c85-ecb8-4514-a610-364790c0773e","name":"Canada Social Insurance Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:CA","pii:direct","tier:extended"]},{"id":"4e979794-49a0-407e-a0b9-2c536937b925","name":"Chile Identity Card Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:CL","pii:direct","tier:core"]},{"id":"c37d5e84-bc2e-458a-be24-6a69b23562f7","name":"China Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:CN","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"c92daa86-2d16-4871-901f-816b3f554fc1","name":"China Resident Identity Card (PRC) Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:CN","pii:direct","tier:core"]},{"id":"64ab8b8c-cb50-4e2c-97e5-2a1699a850e2","name":"Client Secret / Api Key","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"060c591a-0fd7-4598-a715-8a94b283dfcf","name":"Colombia National ID","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:CO","pii:direct","tier:core"]},{"id":"df275e7b-40d0-4519-b668-59cee178df65","name":"Colombia Tax Identification Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:CO","pii:direct","tier:core"]},{"id":"50842eb7-edc8-4019-85dd-5a5c1f2bb085","name":"Credit Card Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:GLOBAL","pii:direct","tier:core","trait:checksum-validated"],"note":"Luhn-validated. Most reliable financial SIT."},{"id":"005b3ef1-47dd-4e68-bb02-c6db484d00f2","name":"Croatia Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:HR","pii:direct","tier:core"]},{"id":"ff12f884-c20a-4189-b185-34c8e7258d47","name":"Croatia Identity Card Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:HR","pii:direct","tier:core"]},{"id":"7d7a729d-32d8-4204-8d01-d5e6a6c25581","name":"Croatia Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:HR","pii:direct","tier:core"]},{"id":"31983b6d-db95-4eb2-a630-b44bd091968d","name":"Croatia Personal Identification (OIB) Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:HR","pii:direct","tier:core"]},{"id":"34cf4670-97df-4752-97bc-f6a7b144f774","name":"Croatia Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:HR","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"356fa104-f9ac-4aff-a0e4-2e6e65ea06c4","name":"Cyprus Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:CY","pii:direct","tier:core"]},{"id":"3ba8afe5-7a6c-4929-8247-0001b6878438","name":"Cyprus Identity Card","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:CY","pii:direct","tier:core"]},{"id":"9193e2e8-7f8c-43c1-a274-ac40d651936f","name":"Cyprus Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:CY","pii:direct","tier:core"]},{"id":"1e87347d-47f8-40e7-9d9c-d54097d2f598","name":"Cyprus Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:CY","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"40e64bd9-55f3-4a09-9bd6-1db18dced9dd","name":"Cyprus Tax Identification Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:CY","pii:direct","tier:core"]},{"id":"86b40d3b-d8ea-4c36-aab0-ef9416a6769c","name":"Czech Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:CZ","pii:direct","tier:core"]},{"id":"60c0725a-4eb6-455b-9dda-05d8a7396497","name":"Czech Personal Identity Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:CZ","pii:direct","tier:core"]},{"id":"7bcd8ce8-5e92-4bbe-bc92-fa669f0369fa","name":"Czech Republic Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:CZ","pii:direct","tier:core"]},{"id":"75571732-c941-4cf3-8edc-db5a7bbc5e03","name":"Czech Republic Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:CZ","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"98a95812-6203-451a-a220-d39870ebef0e","name":"Denmark Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:DK","pii:direct","tier:core"]},{"id":"25e8c47e-e6fe-4884-a211-74898f8c0196","name":"Denmark Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:DK","pii:direct","tier:core"]},{"id":"6c4f2fef-56e1-4c00-8093-88d7a01cf460","name":"Denmark Personal Identification Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:DK","pii:direct","tier:core"]},{"id":"06f4c2d4-77c2-415f-8495-75127051a61c","name":"Denmark Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:DK","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"17066377-466d-43ff-997f-c9240414021c","name":"Diseases","recommendedConfidence":"High","tags":["class:health","jurisdiction:GLOBAL","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"9a5445ad-406e-43eb-8bd7-cac17ab6d0e4","name":"Drug Enforcement Agency (DEA) Number","recommendedConfidence":"High","tags":["class:health","jurisdiction:US","pii:direct","tier:extended"],"note":"US prescriber identifier."},{"id":"0e9b3178-9678-47dd-a509-37222ca96b42","name":"EU Debit Card Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:EU","pii:direct","tier:extended"]},{"id":"b8fe86d1-c056-453b-bfaa-9fe698699ecc","name":"EU Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:EU","pii:direct","tier:core"]},{"id":"419f449f-6d9d-4be1-a154-b531f7a91b41","name":"EU National Identification Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:EU","pii:direct","tier:core"]},{"id":"21883626-6245-4f3d-9b61-5cbb43e625ee","name":"EU Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:EU","pii:direct","tier:core"]},{"id":"d24e32a4-c0bb-4ba8-899d-6303b95742d9","name":"EU Social Security Number (SSN) or Equivalent ID","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:EU","pii:direct","tier:core"]},{"id":"e09c07d3-66e5-4783-989d-49ac62748f5f","name":"EU Tax Identification Number (TIN)","recommendedConfidence":"Medium","tags":["class:financial","jurisdiction:EU","pii:direct","tier:core"]},{"id":"bc99c8a0-95ea-410e-9b7c-05e51b02e469","name":"Ecuador Unique Identification Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:EC","pii:direct","tier:core"]},{"id":"51da8171-da70-4cc1-9d65-055a59ca4f83","name":"Estonia Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:EE","pii:direct","tier:core"]},{"id":"61f7073a-509e-425b-a754-bc01bb5d5b8c","name":"Estonia Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:EE","pii:direct","tier:core"]},{"id":"bfb26de6-dad5-4d48-ab72-4789cdd0654c","name":"Estonia Personal Identification Code","recommendedConfidence":"High","tags":["class:identity","jurisdiction:EE","pii:direct","tier:core"]},{"id":"026cadfb-df9c-4633-8949-6dd9f9f4eb01","name":"Estonia Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:EE","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"bb3b27a3-79bd-4ac4-81a7-f9fca3c7d1a7","name":"Finland Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:FI","pii:direct","tier:core"]},{"id":"60f75aed-81bf-4625-89b0-0846b9248ee7","name":"Finland European Health Insurance Number","recommendedConfidence":"Medium","tags":["class:health","jurisdiction:FI","pii:direct","tier:extended"]},{"id":"338fd995-4cb5-4f87-ad35-79bd1dd926c1","name":"Finland National ID","recommendedConfidence":"High","tags":["class:identity","jurisdiction:FI","pii:direct","tier:core"]},{"id":"d1685ac3-1d3a-40f8-8198-32ef5669c7a5","name":"Finland Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:FI","pii:direct","tier:core"]},{"id":"a7417d1f-eee9-4f88-8f43-e94e35677ff8","name":"Finland Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:FI","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"18e55a36-a01b-4b0f-943d-dc10282a1824","name":"France Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:FR","pii:direct","tier:core"]},{"id":"9bc2069e-76df-4ff9-ac02-2f519469e236","name":"France Health Insurance Number","recommendedConfidence":"Medium","tags":["class:health","jurisdiction:FR","pii:direct","tier:extended"]},{"id":"f741ac74-1bc0-4665-b69b-f0c7f927c0c4","name":"France National ID Card (CNI)","recommendedConfidence":"Low","tags":["class:identity","jurisdiction:FR","pii:direct","tier:core"]},{"id":"3008b884-8c8c-4cd8-a289-99f34fc7ff5d","name":"France Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:FR","pii:direct","tier:core"]},{"id":"674c3b99-2a5f-4698-ab29-81828328da9b","name":"France Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:FR","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"71f62b97-efe0-4aa1-aa49-e14de253619d","name":"France Social Security Number (INSEE)","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:FR","pii:direct","tier:core"]},{"id":"ed59e77e-171d-442c-9ec1-88e2ebcb5b0a","name":"France Tax Identification Number (num\u00e9ro SPI.)","recommendedConfidence":"High","tags":["class:financial","jurisdiction:FR","pii:direct","tier:core"]},{"id":"949121e6-ad9f-4379-8731-710342baea78","name":"France Value Added Tax Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:FR","pii:direct","tier:extended"]},{"id":"24a90d49-519b-4689-aa13-f35c47cc40c3","name":"General Password","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"664cfcd3-afdb-4fe2-b417-c2040ddd960d","name":"General Symmetric Key","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"d82f1c56-8eb9-4cf2-b279-43872774c7f6","name":"Generic Medication Names","recommendedConfidence":"High","tags":["class:health","jurisdiction:GLOBAL","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"91da9335-1edb-45b7-a95f-5fe41a16c63c","name":"German Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:DE","pii:direct","tier:core"]},{"id":"2e3da144-d42b-47ed-b123-fbf78604e52c","name":"German Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:DE","pii:direct","tier:core"]},{"id":"e577372f-c42e-47a0-9d85-bebed1c237d4","name":"Germany Identity Card Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:DE","pii:direct","tier:core"]},{"id":"d703beef-45ec-497f-bdcb-74e220e764b7","name":"Germany Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:DE","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"43316a89-9880-40cf-b980-04bc7eefcec5","name":"Germany Tax Identification Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:DE","pii:direct","tier:core"]},{"id":"db177eb2-8811-4842-bffc-128c14aa219f","name":"Germany Value Added Tax Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:DE","pii:direct","tier:extended"]},{"id":"927cd1e7-8377-4df3-9bc3-fabdaef94787","name":"GitHub Personal Access Token","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"9bea20ec-764d-4987-88a1-093bbbacadd2","name":"Google API key","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"7a2200b5-aacf-4e3c-ab36-136d3e68b7da","name":"Greece Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:GR","pii:direct","tier:core"]},{"id":"82568215-1da1-46d3-874a-d2294d81b5ac","name":"Greece National ID Card","recommendedConfidence":"High","tags":["class:identity","jurisdiction:GR","pii:direct","tier:core"]},{"id":"7e65eb47-cdf9-4f52-8f90-2a27d5ee67e3","name":"Greece Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:GR","pii:direct","tier:core"]},{"id":"ec2e5fea-da77-4fb5-bc6f-75cd7642bac3","name":"Greece Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:GR","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"e39b03f4-50ea-41ae-af7a-a4b9539596ad","name":"Greece Social Security Number (AMKA)","recommendedConfidence":"High","tags":["class:identity","jurisdiction:GR","pii:direct","tier:core"]},{"id":"15a54a5a-53d4-4080-ad43-a2a4fe1d3bf7","name":"Greek Tax identification Number","recommendedConfidence":"Medium","tags":["class:financial","jurisdiction:GR","pii:direct","tier:core"]},{"id":"9135827f-a27c-4d6d-ba8c-69e166414b87","name":"Greenland Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:GL","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"e63c28a7-ad29-4c17-a41a-3d2a0b70fd9c","name":"Hong Kong Identity Card (HKID) Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:HK","pii:direct","tier:core"]},{"id":"193c16a6-58cb-4d8a-adf0-7cbad55e0734","name":"Http Authorization Header","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"0de78315-9537-47f5-95ab-b3e77eba3993","name":"Hungarian Social Security Number (TAJ)","recommendedConfidence":"High","tags":["class:identity","jurisdiction:HU","pii:direct","tier:core"]},{"id":"976349a0-683b-477a-90f8-ff0a220d5592","name":"Hungarian Value Added Tax Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:HU","pii:direct","tier:extended"]},{"id":"9d31c46b-6e6b-444c-aeb1-6dd7e604bb24","name":"Hungary Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:HU","pii:direct","tier:core"]},{"id":"5b483910-9aa7-4c99-9917-f4001464bda7","name":"Hungary Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:HU","pii:direct","tier:core"]},{"id":"7b5cc218-7046-47d9-80c9-f325b50896ca","name":"Hungary Personal Identification Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:HU","pii:direct","tier:core"]},{"id":"4dcc3216-d23f-40a0-a1ee-80ff322f599e","name":"Hungary Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:HU","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"ede42eb4-59d9-49eb-9603-d7853fbda91d","name":"Hungary Tax identification Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:HU","pii:direct","tier:core"]},{"id":"1daa4ad5-e2dd-4ca4-a788-54722c09efb2","name":"IP Address","recommendedConfidence":"High","tags":["class:technical","jurisdiction:GLOBAL","pii:indirect","tier:extended","trait:high-volume"]},{"id":"a7dd5e5f-e7f9-4626-a2c6-86a8cb6830d2","name":"IP Address v4","recommendedConfidence":"Medium","tags":["class:technical","jurisdiction:GLOBAL","pii:indirect","tier:extended","trait:high-volume"]},{"id":"3f691089-7413-4926-ab3b-3c5ea8a1c17e","name":"IP Address v6","recommendedConfidence":"Medium","tags":["class:technical","jurisdiction:GLOBAL","pii:indirect","tier:extended","trait:high-volume"]},{"id":"dfb6646c-d21b-4318-8def-e55af142d881","name":"Iceland Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:IS","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"88fd9212-9d63-451d-9c74-05a13ac6732a","name":"Impairments Listed In The U.S. Disability Evaluation Under Social Security","recommendedConfidence":"High","tags":["class:health","jurisdiction:US","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"680788a3-53b6-455a-b891-c38cd76dc917","name":"India Driver's License Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:IN","pii:direct","tier:core"]},{"id":"9f5a721c-2fd2-446a-a27e-0c02fbe4630c","name":"India GST Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:IN","pii:direct","tier:extended"]},{"id":"2602bfee-9bb0-47a5-a7a6-2bf3053e2804","name":"India Permanent Account Number (PAN)","recommendedConfidence":"High","tags":["class:identity","jurisdiction:IN","pii:direct","tier:core"]},{"id":"1ca46b29-76f5-4f46-9383-cfa15e91048f","name":"India Unique Identification (Aadhaar) Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:IN","pii:direct","tier:core"]},{"id":"646d643f-5228-4408-acc8-f2e81a6df897","name":"India Voter Id Card","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:IN","pii:direct","tier:extended"]},{"id":"67d902ee-c3ce-44a2-81ac-e9746e7d60a1","name":"Indonesia Drivers License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:ID","pii:direct","tier:core"]},{"id":"da68fdb0-f383-4981-8c86-82689d3b7d55","name":"Indonesia Identity Card (KTP) Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:ID","pii:direct","tier:core"]},{"id":"d0001c83-e72f-4360-98d3-f5a41dc5a380","name":"Indonesia Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:ID","pii:direct","tier:core"]},{"id":"e7dc4711-11b7-4cb0-b88b-2c394a771f0e","name":"International Banking Account Number (IBAN)","recommendedConfidence":"High","tags":["class:financial","jurisdiction:GLOBAL","pii:direct","tier:extended","trait:checksum-validated"]},{"id":"3356946c-6bb7-449b-b253-6ffa419c0ce7","name":"International Classification of Diseases (ICD-10-CM)","recommendedConfidence":"High","tags":["class:health","jurisdiction:US","pii:indirect","tier:extended"],"note":"Clinical Modification is the US adaptation of ICD-10. Health data but not a personal identifier alone."},{"id":"fa3f9c74-ee07-4c52-b5f2-085d6b2c0ec4","name":"International Classification of Diseases (ICD-9-CM)","recommendedConfidence":"High","tags":["class:health","jurisdiction:US","pii:indirect","tier:extended"],"note":"Legacy US diagnosis code set. Still present in archived records."},{"id":"e01bccd9-eb4d-414f-ace1-e9b6a4c4a2ca","name":"Ireland Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:IE","pii:direct","tier:core"]},{"id":"a2130f27-9ee2-4103-84f9-a6b1ee7d0cbf","name":"Ireland Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:IE","pii:direct","tier:core"]},{"id":"1cdb674d-c19a-4fcf-9f4b-7f56cc87345a","name":"Ireland Personal Public Service (PPS) Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:IE","pii:direct","tier:extended"]},{"id":"651646de-f4dd-4e7c-91ad-8679d8049bfa","name":"Ireland Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:IE","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"7d08b2ff-a0b9-437f-957c-aeddbf9b2b25","name":"Israel Bank Account Number","recommendedConfidence":"Medium","tags":["class:financial","jurisdiction:IL","pii:direct","tier:core"]},{"id":"e05881f5-1db1-418c-89aa-a3ac5c5277ee","name":"Israel National ID","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:IL","pii:direct","tier:core"]},{"id":"97d6244f-9157-41bd-8e0c-9d669a5c4d71","name":"Italy Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:IT","pii:direct","tier:core"]},{"id":"4cd79172-8da9-4ff5-9188-98b1e7e2eca6","name":"Italy Fiscal Code","recommendedConfidence":"High","tags":["class:identity","jurisdiction:IT","pii:direct","tier:core"]},{"id":"39811019-4750-445f-b26d-4c0e6c431544","name":"Italy Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:IT","pii:direct","tier:core"]},{"id":"3270dbe0-d63b-4dde-8aa4-bc6f19f8a802","name":"Italy Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:IT","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"26a8cc07-2283-4a2a-ab1d-4ab643c4c67f","name":"Italy Value Added Tax Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:IT","pii:direct","tier:extended"]},{"id":"d354f95b-96ee-4b80-80bc-4377312b55bc","name":"Japan Bank Account Number","recommendedConfidence":"Medium","tags":["class:financial","jurisdiction:JP","pii:direct","tier:core"]},{"id":"c6011143-d087-451c-8313-7f6d4aed2270","name":"Japan Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:JP","pii:direct","tier:core"]},{"id":"75177310-1a09-4613-bf6d-833aae3743f8","name":"Japan Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:JP","pii:direct","tier:core"]},{"id":"19289d79-b636-40b3-a92b-817ec7133531","name":"Japan Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:JP","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"01c1209b-6389-4faf-a5f8-3f7e13899652","name":"Japan Resident Registration Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:JP","pii:direct","tier:extended"]},{"id":"c840e719-0896-45bb-84fd-1ed5c95e45ff","name":"Japan Social Insurance Number (SIN)","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:JP","pii:direct","tier:extended"]},{"id":"9e0eaf79-ff20-4ffb-b3e4-e7368d5db6ff","name":"Japanese My Number Corporate","recommendedConfidence":"High","tags":["class:identity","jurisdiction:JP","pii:none","tier:corroborating"]},{"id":"98da8e66-7299-4ebd-9f82-c871ab37d3ef","name":"Japanese My Number Personal","recommendedConfidence":"High","tags":["class:identity","jurisdiction:JP","pii:direct","tier:core"]},{"id":"ac36fef2-a289-4e2c-bb48-b02366e89fc0","name":"Japanese Residence Card Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:JP","pii:direct","tier:extended"]},{"id":"56967ee6-421b-412a-8076-653d97164620","name":"Korea Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:KR","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"ee05bb9c-7b87-42e1-9987-446b243245d5","name":"Lab Test Terms","recommendedConfidence":"High","tags":["class:health","jurisdiction:GLOBAL","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"ec996de0-30f2-46b1-b192-4d2ff8805fa7","name":"Latvia Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:LV","pii:direct","tier:core"]},{"id":"23ae25ec-cc28-421b-b77a-3054eadf1ede","name":"Latvia Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:LV","pii:direct","tier:core"]},{"id":"03fcf763-27c2-49ed-9422-2641c6c895c9","name":"Latvia Personal Code","recommendedConfidence":"High","tags":["class:identity","jurisdiction:LV","pii:direct","tier:core"]},{"id":"92a71ddd-1acc-4aca-ac8e-94db01841e1c","name":"Latvia Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:LV","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"83127ddf-b7c7-4938-8c81-8b08d65bbb5f","name":"Liechtenstein Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:LI","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"7740d4dd-4f07-452b-b039-6c8dab0ba5bc","name":"Lifestyles That Relate To Medical Conditions","recommendedConfidence":"High","tags":["class:health","jurisdiction:GLOBAL","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"86f7628b-e0f4-4dc3-9fbc-e4300e4c7d78","name":"Lithuania Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:LT","pii:direct","tier:core"]},{"id":"1b79900f-047b-4c3f-846f-7d73b5534bce","name":"Lithuania Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:LT","pii:direct","tier:core"]},{"id":"cd6d3786-8ec3-4524-a2cf-1e0095379171","name":"Lithuania Personal Code","recommendedConfidence":"High","tags":["class:identity","jurisdiction:LT","pii:direct","tier:core"]},{"id":"cdde1960-c72b-4705-b56f-606570fc33bc","name":"Lithuania Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:LT","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"89daf717-1544-4860-9a2e-fc9166dd8852","name":"Luxembourg Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:LU","pii:direct","tier:core"]},{"id":"aaf661ed-29ec-426d-8bf9-880cad298ebb","name":"Luxembourg National Identification Number (Natural persons)","recommendedConfidence":"High","tags":["class:identity","jurisdiction:LU","pii:direct","tier:core"]},{"id":"84bffa3a-d805-4788-a613-b1e4df3804cf","name":"Luxembourg National Identification Number (Non-natural persons)","recommendedConfidence":"High","tags":["class:identity","jurisdiction:LU","pii:direct","tier:core"]},{"id":"81d5c027-bed9-4421-91a0-3b2e55b3eb85","name":"Luxembourg Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:LU","pii:direct","tier:core"]},{"id":"1626e616-84a4-401b-b062-d12f6ad470a6","name":"Luxembourg Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:LU","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"7f0e921c-9677-435b-aba2-bb8f1013c749","name":"Malaysia Identity Card Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:MY","pii:direct","tier:core"]},{"id":"62ea2a27-798e-4de3-a31f-a54c5d02ff37","name":"Malaysia Passport Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:MY","pii:direct","tier:core"]},{"id":"a3bdaa4a-8371-4735-8fa5-56ee0fb4afc4","name":"Malta Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:MT","pii:direct","tier:core"]},{"id":"854b36b3-a388-4ac8-a4ec-677c2b5e4356","name":"Malta Identity Card Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:MT","pii:direct","tier:core"]},{"id":"b2b21198-48f9-4d13-b2a5-03969bff0fb8","name":"Malta Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:MT","pii:direct","tier:core"]},{"id":"e18f8702-ef56-4db5-be2b-98c205dc2e3e","name":"Malta Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:MT","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"ec830c63-65f4-45d0-9d8c-910dc8334b20","name":"Malta Tax ID Number","recommendedConfidence":"Medium","tags":["class:financial","jurisdiction:MT","pii:direct","tier:extended"]},{"id":"a430ac76-0141-4378-b5fd-293e65890d48","name":"Medical Specialties","recommendedConfidence":"High","tags":["class:health","jurisdiction:GLOBAL","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"f753a286-f5cc-47e6-a592-4be25fd02591","name":"Medicare Beneficiary Identifier (MBI) card","recommendedConfidence":"Medium","tags":["class:health","jurisdiction:US","pii:direct","tier:core"],"note":"US Medicare identifier."},{"id":"e905ad4d-5a74-406d-bf36-b1efca798af4","name":"Mexico Unique Population Registry Code (CURP)","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:MX","pii:direct","tier:core"]},{"id":"c1ce0e56-014f-40c0-a268-e702117a60e3","name":"Microsoft Bing Maps Key","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"f61444f3-b87e-4c16-b835-b588033db6b7","name":"Microsoft Entra Client Access Token","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"d820a2f1-46ef-4220-957e-6eb9abc1aadb","name":"Microsoft Entra Client Secret","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"b7d05856-6d0a-4488-b4a8-1b5c48bab75a","name":"Microsoft Entra User Credentials","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"c5f54253-ef7e-44f6-a578-440ed67e946d","name":"Netherlands Citizen's Service (BSN) Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:NL","pii:direct","tier:core"]},{"id":"6247fbea-ab80-4be5-8233-308b7c031401","name":"Netherlands Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:NL","pii:direct","tier:core"]},{"id":"61786727-bafd-45f6-94d9-888d815e228e","name":"Netherlands Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:NL","pii:direct","tier:core"]},{"id":"b3650e47-f2dc-4d29-a4fc-1b9154005c07","name":"Netherlands Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:NL","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"01f42a64-eba7-4892-a67b-398237e4ade2","name":"Netherlands Tax Identification Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:NL","pii:direct","tier:core"]},{"id":"4f320d9b-4972-41ae-b337-88d499bb1ade","name":"Netherlands Value Added Tax Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:NL","pii:direct","tier:extended"]},{"id":"1924b377-d287-49c9-a737-cfe7a8a2615a","name":"New Zealand Driver License Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:NZ","pii:direct","tier:core"]},{"id":"dd0fe2bc-7bcf-455f-bac1-83b1e3eb25fd","name":"New Zealand Inland Revenue number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:NZ","pii:direct","tier:core"]},{"id":"2b71c1c8-d14e-4430-82dc-fd1ed6bf05c7","name":"New Zealand Ministry of Health Number","recommendedConfidence":"High","tags":["class:health","jurisdiction:NZ","pii:direct","tier:extended"]},{"id":"0fe7f6b1-b287-421c-9286-3a5fb7bc208c","name":"New Zealand Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:NZ","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"20f3c48d-4ac1-4cd2-86bd-34ecc1826e9d","name":"New Zealand Social Welfare Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:NZ","pii:direct","tier:extended"]},{"id":"1a97fc2b-dd2f-48f1-bc4e-2ddf25813956","name":"New Zealand bank account number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:NZ","pii:direct","tier:core"]},{"id":"d4c8a798-e9f2-4bd3-9652-500d24080fc3","name":"Norway Identity Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:NO","pii:direct","tier:core"]},{"id":"8751dcac-e02a-490a-bdf2-aff90ce09f84","name":"Norway Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:NO","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"3257da62-3d98-44ca-a923-6f8b0cd94b7f","name":"Philippines National ID","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:PH","pii:direct","tier:core"]},{"id":"6fa57f91-314a-4561-8248-7ab921957448","name":"Philippines Passport Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:PH","pii:direct","tier:core"]},{"id":"019b39dd-8c25-4765-91a3-d9c6baf3c3b3","name":"Philippines Unified Multi-Purpose ID Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:PH","pii:direct","tier:extended"]},{"id":"24d51f99-ee9e-4060-a077-cae58cab1ee4","name":"Poland Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:PL","pii:direct","tier:core"]},{"id":"25e64989-ed5d-40ca-a939-6c14183bb7bf","name":"Poland Identity Card","recommendedConfidence":"High","tags":["class:identity","jurisdiction:PL","pii:direct","tier:core"]},{"id":"e3aaf206-4297-412f-9e06-ba8487e22456","name":"Poland National ID (PESEL)","recommendedConfidence":"High","tags":["class:identity","jurisdiction:PL","pii:direct","tier:core"]},{"id":"03937fb5-d2b6-4487-b61f-0f8bff7c3517","name":"Poland Passport","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:PL","pii:direct","tier:core"]},{"id":"ed728544-c58d-470d-9bde-ee6ec901f38d","name":"Poland Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:PL","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"1ff28b4d-40f2-49e9-b677-9606a88e2bca","name":"Poland Tax Identification Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:PL","pii:direct","tier:core"]},{"id":"fc87b421-f437-4f8b-b739-29a735ead0d9","name":"Polish REGON Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:PL","pii:none","tier:corroborating"]},{"id":"91a7ece2-add4-4986-9a15-c84544d81ecd","name":"Portugal Citizen Card Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:PT","pii:direct","tier:core"]},{"id":"977f1e5a-2c33-4bcc-b516-95bb275cff23","name":"Portugal Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:PT","pii:direct","tier:core"]},{"id":"080a52fd-a7bc-431e-b54d-51f08f59db11","name":"Portugal Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:PT","pii:direct","tier:core"]},{"id":"65372402-3131-4f1e-9983-4439841d1f15","name":"Portugal Tax Identification Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:PT","pii:direct","tier:core"]},{"id":"eda085a4-1f9b-46ca-8f1f-885ef394a476","name":"Portuguese Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:PT","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"52b1b60e-a4be-4b5a-a67b-6f9bbb7811da","name":"Qatari ID Card Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:QA","pii:direct","tier:extended"]},{"id":"b5511ace-2fd8-4ae4-b6fc-c7c6e4689e3c","name":"Romania Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:RO","pii:direct","tier:core"]},{"id":"5d31b90c-7fe2-4a76-a14b-767b8fd19d6c","name":"Romania Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:RO","pii:direct","tier:core"]},{"id":"eb5fa399-fe28-4c67-8188-d63a616ed89c","name":"Romania Personal Numerical Code (CNP)","recommendedConfidence":"High","tags":["class:identity","jurisdiction:RO","pii:direct","tier:extended"]},{"id":"e3a1f6fc-92f9-4b1c-b3fb-94b182cdc3ed","name":"Romania Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:RO","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"515c522b-ce04-4a40-816b-8226f26c65e7","name":"Russia Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:RU","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"76ec2f5d-cedb-48e1-8070-1998794af445","name":"Russian Passport Number (Domestic)","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:RU","pii:direct","tier:core"]},{"id":"ac5f4878-75e4-4b82-af2d-02e13ea9f411","name":"Russian Passport Number (International)","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:RU","pii:direct","tier:core"]},{"id":"81145f47-868a-464c-956a-eed7696fdcdf","name":"Russian Taxpayer Identification Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:RU","pii:direct","tier:core"]},{"id":"e76b6205-d3cb-46f2-bd63-c90153f2f97d","name":"SQL Server Connection String","recommendedConfidence":"High","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"cb2ab58c-9cb8-4c81-baf8-a4e106791df4","name":"SWIFT Code","recommendedConfidence":"Medium","tags":["class:financial","jurisdiction:GLOBAL","pii:none","tier:corroborating","trait:corroborating"]},{"id":"8c5a0ba8-404a-41a3-8871-746aa21ee6c0","name":"Saudi Arabia National ID","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:SA","pii:direct","tier:core"]},{"id":"beaa54d7-ec6a-4160-a37b-f3916b68934b","name":"Singapore Driving License Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:SG","pii:direct","tier:extended"]},{"id":"cead390a-dd83-4856-9751-fb6dc98c34da","name":"Singapore National Registration Identity Card (NRIC) Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:SG","pii:direct","tier:core"]},{"id":"4081ae73-c4ec-405d-932c-8b36ae35b20c","name":"Singapore Passport Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:SG","pii:direct","tier:core"]},{"id":"6dee3dcd-0946-48a8-b49e-e3596b671fc5","name":"Singapore Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:SG","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"1ac53bfc-5956-49ae-b4b6-7087bf0cfc88","name":"Slack Access Token","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"14240c22-b6de-4ce5-a90b-137f74252513","name":"Slovakia Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:SK","pii:direct","tier:core"]},{"id":"238e1f08-d80e-4793-af33-9b57918335b7","name":"Slovakia Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:SK","pii:direct","tier:core"]},{"id":"951c26b7-3b35-4f73-924b-15dd599cb9ab","name":"Slovakia Personal Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:SK","pii:direct","tier:core"]},{"id":"1b342635-f7dc-4e15-a288-3f7e02c93169","name":"Slovakia Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:SK","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"d5bc089a-f2ee-433d-a6b1-5c253051d6f2","name":"Slovenia Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:SI","pii:direct","tier:core"]},{"id":"235b7976-7bbe-4df5-bb40-08678e749d1a","name":"Slovenia Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:SI","pii:direct","tier:core"]},{"id":"0b76dc4d-cf09-4174-b334-fc72e1b38bb9","name":"Slovenia Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:SI","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"e47b071e-c352-4d70-8241-8c215ad65505","name":"Slovenia Tax Identification Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:SI","pii:direct","tier:core"]},{"id":"68948b27-803d-41e4-adf1-13e05eb541bb","name":"Slovenia Unique Master Citizen Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:SI","pii:direct","tier:core"]},{"id":"e2adf7cb-8ea6-4048-a2ed-d89eb65f2780","name":"South Africa Identification Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:ZA","pii:direct","tier:core"]},{"id":"0311f67f-d239-4f57-bde7-bd76253576eb","name":"South Africa Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:ZA","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"3da0b024-4c08-406b-9f9e-de96accd0209","name":"South Korea Driver's License Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:KR","pii:direct","tier:core"]},{"id":"cb725748-0f69-4282-8a03-e9b222a10a18","name":"South Korea Passport Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:KR","pii:direct","tier:core"]},{"id":"5b802e18-ba80-44c4-bc83-bf2ad36ae36a","name":"South Korea Resident Registration Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:KR","pii:direct","tier:extended"]},{"id":"8e6251b9-47b4-40e8-a42b-0f80876be192","name":"Spain DNI","recommendedConfidence":"High","tags":["class:identity","jurisdiction:ES","pii:direct","tier:core"]},{"id":"d5a82922-b501-4f40-8868-341321146aa2","name":"Spain Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:ES","pii:direct","tier:core"]},{"id":"d17a57de-9fa5-4e9f-85d3-85c26d89686e","name":"Spain Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:ES","pii:direct","tier:core"]},{"id":"36c81f62-73f6-4f62-8350-2a809cb27ff6","name":"Spain Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:ES","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"5df987c0-8eae-4bce-ace7-b316347f3070","name":"Spain Social Security Number (SSN)","recommendedConfidence":"High","tags":["class:identity","jurisdiction:ES","pii:direct","tier:core"]},{"id":"10f0d113-b0e1-47dc-872a-a4f45b9376a3","name":"Spain Tax Identification Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:ES","pii:direct","tier:core"]},{"id":"daee7349-84f5-48ec-81c9-f6fafc7c976f","name":"Surgical Procedures","recommendedConfidence":"High","tags":["class:health","jurisdiction:GLOBAL","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"70088720-90dd-47f5-805e-5525f3567391","name":"Sweden Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:SE","pii:direct","tier:core"]},{"id":"f69aaf40-79be-4fac-8f05-fd1910d272c8","name":"Sweden National ID","recommendedConfidence":"High","tags":["class:identity","jurisdiction:SE","pii:direct","tier:core"]},{"id":"ba4e7456-55a9-4d89-9140-c33673553526","name":"Sweden Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:SE","pii:direct","tier:core"]},{"id":"99d64fe7-8168-4cb8-9e8f-882444b4375a","name":"Sweden Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:SE","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"139acba0-a5bc-4fbb-876d-f7a493ae8a40","name":"Sweden Tax Identification Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:SE","pii:direct","tier:core"]},{"id":"277cfa4b-6eaa-4a1b-9492-099dec849971","name":"Swiss Social Security Number AHV","recommendedConfidence":"High","tags":["class:identity","jurisdiction:CH","pii:direct","tier:core"]},{"id":"0741710b-c0b5-400c-98a7-864942ac89be","name":"Switzerland Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:CH","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"4c7bfc34-8dd1-421d-8fb7-6c6182c2af03","name":"Taiwan National ID","recommendedConfidence":"High","tags":["class:identity","jurisdiction:TW","pii:direct","tier:core"]},{"id":"e7251cb4-4c2c-41df-963e-924eb3dae04a","name":"Taiwan Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:TW","pii:direct","tier:core"]},{"id":"7964a3bf-18fd-4e6c-a04e-d72f0cf4941c","name":"Taiwan Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:TW","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"48269fec-05ea-46ea-b326-f5623a58c6e9","name":"Taiwan Resident Certificate (ARC/TARC)","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:TW","pii:none","tier:excluded","trait:requires-e5"]},{"id":"44ca9e86-ead7-4c5d-884a-e2eaa401515e","name":"Thai Population Identification Code","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:TH","pii:direct","tier:extended"]},{"id":"cf1d79ec-3162-49a8-9bcd-3cbeda2a3439","name":"Turkey Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:TR","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"fb621f20-3876-4cfc-acec-8c8e73ca32c7","name":"Turkish National Identification number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:TR","pii:direct","tier:core"]},{"id":"f6dc2d17-3549-41e2-af29-ae1846ae9542","name":"Types Of Medication","recommendedConfidence":"High","tags":["class:health","jurisdiction:GLOBAL","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"f93de4be-d94c-40df-a8be-461738047551","name":"U.K. Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:GB","pii:direct","tier:core"]},{"id":"a3eea206-dc0c-4f06-9e22-aa1be3059963","name":"U.K. Electoral Roll Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:GB","pii:direct","tier:extended"]},{"id":"3192014e-2a16-44e9-aa69-4b20375c9a78","name":"U.K. National Health Service Number","recommendedConfidence":"High","tags":["class:health","jurisdiction:GB","pii:direct","tier:extended"]},{"id":"16c07343-c26f-49d2-a987-3daf717e94cc","name":"U.K. National Insurance Number (NINO)","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:GB","pii:direct","tier:extended"]},{"id":"accaf4c2-fb54-40f7-aea8-db0e36a2e9eb","name":"U.K. Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:GB","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"ad4a8116-0db8-439a-b545-6d967642f0ec","name":"U.K. Unique Taxpayer Reference Number","recommendedConfidence":"High","tags":["class:financial","jurisdiction:GB","pii:direct","tier:core"]},{"id":"178ec42a-18b4-47cc-85c7-d62c92fd67f8","name":"U.S. / U.K. Passport Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:GB","jurisdiction:US","pii:direct","tier:core"],"note":"Single SIT covers both US and UK passports."},{"id":"a2ce32a8-f935-4bb6-8e96-2a5157672e2c","name":"U.S. Bank Account Number","recommendedConfidence":"Medium","tags":["class:financial","jurisdiction:US","pii:direct","tier:core","trait:pairs-with:cb353f78-2b72-4c3c-8827-92ebe4f69fdf","trait:weak-pattern"],"note":"Weak pattern alone. Pair with ABA Routing Number to raise precision."},{"id":"dfeb356f-61cd-459e-bf0f-7c6d28b458c6","name":"U.S. Driver's License Number","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:US","pii:direct","tier:core","trait:format-varies-by-state","trait:weak-pattern"],"note":"State formats vary widely. Noisy without keyword proximity."},{"id":"e55e2a32-f92d-4985-a35d-a0b269eb687b","name":"U.S. Individual Taxpayer Identification Number (ITIN)","recommendedConfidence":"Medium","tags":["class:financial","jurisdiction:US","pii:direct","tier:core"],"note":"Taxpayer ID for non-SSN holders. Include wherever SSN is included."},{"id":"44aa44f2-63d1-41df-af0d-970283ac41e2","name":"U.S. Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:US","pii:indirect","tier:extended","trait:exo-requires-priva","trait:high-volume","trait:named-entity","trait:requires-e5"],"note":"Named entity. High volume and requires advanced classification licensing."},{"id":"a44669fe-0d48-453d-a9b1-2cc83f2cba77","name":"U.S. Social Security Number (SSN)","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:US","pii:direct","tier:core","trait:keyword-corroborated"],"note":"Highest-priority US identifier. Strong pattern with keyword corroboration."},{"id":"853a8051-ad90-417c-9345-c73ac0adc1c3","name":"UAE Identity Card Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:AE","pii:direct","tier:core"]},{"id":"111b5e63-6f4e-47c6-99bc-dd95625a5acf","name":"UAE Passport Number","recommendedConfidence":"High","tags":["class:identity","jurisdiction:AE","pii:direct","tier:core"]},{"id":"1817a540-221f-4459-9202-3bd78b81d803","name":"Ukraine Passport Number (Domestic)","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:UA","pii:direct","tier:core"]},{"id":"cfbe032d-22e0-4f28-ab68-d66e9641f1e2","name":"Ukraine Passport Number (International)","recommendedConfidence":"Medium","tags":["class:identity","jurisdiction:UA","pii:direct","tier:core"]},{"id":"97bbd986-e9a8-47e9-91d0-3eb20fdb978f","name":"Ukraine Physical Addresses","recommendedConfidence":"Medium","tags":["class:contact","jurisdiction:UA","pii:indirect","tier:extended","trait:exo-requires-priva","trait:named-entity","trait:requires-e5"]},{"id":"2de77d45-9d47-4ce2-990d-18e26b487194","name":"User Login Credentials","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]},{"id":"417b111d-1909-4457-b686-d6470b29a7ec","name":"X.509 Certificate Private Key","recommendedConfidence":"Medium","tags":["class:credential","jurisdiction:GLOBAL","pii:none","tier:excluded","trait:requires-e5"]}],"licensing":{"note":"Licensing materially changes what a scan can detect. A SIT that is not licensed returns zero rather than an error, so an unlicensed type is indistinguishable from a genuine nil result. Always confirm tenant licensing before interpreting any zero.","rules":[{"applies":"trait:requires-e5","requirement":"Microsoft 365 E5, or an equivalent compliance add-on","affects":"All credential scanning SITs and all named entity SITs","source":"https://learn.microsoft.com/purview/sit-sensitive-information-type-learn-about#licensing"},{"applies":"trait:exo-requires-priva","requirement":"Privacy Service Plan (Microsoft Priva), in addition to E5","affects":"Named entity SITs evaluated against Exchange Online data at rest only. SharePoint, OneDrive and Teams are unaffected","source":"https://learn.microsoft.com/purview/sit-named-entities-use#unified-dlp-support-for-named-entities"}],"notApplicable":[{"item":"Advanced classification scanning and protection","detail":"A tenant setting required only for bundled SITs in ENDPOINT DLP policies. Not required for at-rest classification of Exchange, SharePoint, OneDrive or Teams, and therefore not relevant to a Content Explorer census.","source":"https://learn.microsoft.com/purview/sit-sensitive-information-type-learn-about"}]}}
'@

# ---------------------------------------------------------------------------
# Taxonomy helpers
# ---------------------------------------------------------------------------

# Load and sanity check a taxonomy document. Tag values are deliberately open: an unrecognised tag is
# not an error, so contributors can extend the taxonomy without touching code. What IS checked is tag
# cardinality, because a SIT missing its class/pii/tier tag will silently vanish from every view.
function Import-SitTaxonomy
{
    param(
        [Parameter(Mandatory = $true)][string]$Json,
        [Parameter(Mandatory = $true)][string]$Source
    )

    try
    {
        $F_Doc = $Json | ConvertFrom-Json
    }
    catch
    {
        throw "Taxonomy at '$Source' is not valid JSON: $($_.Exception.Message)"
    }

    if (-not $F_Doc.sits)
    {
        throw "Taxonomy at '$Source' contains no 'sits' array."
    }

    $F_Schema = "$($F_Doc.schemaVersion)"
    if ($F_Schema -and $F_Schema -notmatch '^2\.')
    {
        Write-Warning "Taxonomy schema version '$F_Schema' was not written for this script (expected 2.x). Continuing, but check the tag dimensions and view filter grammar."
    }

    # Cardinality: exactly one class / pii / tier, at least one jurisdiction.
    $F_Issues = [System.Collections.Generic.List[string]]::new()
    foreach ($F_Sit in $F_Doc.sits)
    {
        $F_Tags = @($F_Sit.tags)
        foreach ($F_Dim in @('class', 'pii', 'tier'))
        {
            $F_Count = @($F_Tags | Where-Object { $_ -like "$F_Dim`:*" }).Count
            if ($F_Count -ne 1)
            {
                $F_Issues.Add("$($F_Sit.name): expected exactly one '$F_Dim' tag, found $F_Count")
            }
        }
        if (@($F_Tags | Where-Object { $_ -like 'jurisdiction:*' }).Count -lt 1)
        {
            $F_Issues.Add("$($F_Sit.name): no 'jurisdiction' tag")
        }
    }

    if ($F_Issues.Count -gt 0)
    {
        Write-Warning "Taxonomy at '$Source' has $($F_Issues.Count) tag cardinality issue(s). Affected SITs may not appear in any view."
        foreach ($F_Issue in $F_Issues) { Write-Verbose "  $F_Issue" }
    }

    [PSCustomObject]@{
        Document = $F_Doc
        Source   = $Source
        SitCount = @($F_Doc.sits).Count
        Issues   = $F_Issues
    }
}

# Evaluate one filter entry against a SIT's tags. An entry is either a plain tag string, or an object
# carrying an 'anyOf' array which matches if any of its tags is present. This is the whole grammar -
# it is deliberately not a general expression language.
function Test-SitFilterEntry
{
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Tags,
        [Parameter(Mandatory = $true)]$Entry
    )

    if ($Entry -is [string])
    {
        # -contains on strings is case-insensitive, which is what we want.
        return ($Tags -contains $Entry)
    }

    if ($Entry.PSObject.Properties.Name -contains 'anyOf')
    {
        foreach ($F_Any in @($Entry.anyOf))
        {
            if ($Tags -contains $F_Any) { return $true }
        }
        return $false
    }

    Write-Warning "Unrecognised taxonomy filter entry; ignoring it: $($Entry | ConvertTo-Json -Compress)"
    return $false
}

# Apply an allOf / noneOf filter. An absent or empty clause passes, so an empty filter matches
# everything - which is how the 'all' view is expressed.
function Test-SitFilter
{
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Tags,
        [Parameter(Mandatory = $true)]$Filter
    )

    if (-not $Filter) { return $true }

    # An absent clause must yield no entries. @($null) is a one-element array containing $null, not an
    # empty array, so the null has to be filtered out explicitly or every absent clause iterates once.
    foreach ($F_Entry in @($Filter.allOf | Where-Object { $null -ne $_ }))
    {
        if (-not (Test-SitFilterEntry -Tags $Tags -Entry $F_Entry)) { return $false }
    }

    foreach ($F_Entry in @($Filter.noneOf | Where-Object { $null -ne $_ }))
    {
        if (Test-SitFilterEntry -Tags $Tags -Entry $F_Entry) { return $false }
    }

    return $true
}

# Normalise a jurisdiction argument to a full tag. Accepts 'US' or 'jurisdiction:US'.
function ConvertTo-JurisdictionTag
{
    param([Parameter(Mandatory = $true)][string]$Value)

    $F_Trimmed = $Value.Trim()
    if ($F_Trimmed -like 'jurisdiction:*') { return $F_Trimmed }
    return "jurisdiction:$($F_Trimmed.ToUpperInvariant())"
}

# Resolve the SIT selection.
#
# Precedence, highest first: -SensitiveInfoType, then -Tag, then -View. -Jurisdiction narrows
# whichever of those applied, and -ExcludeTag is applied last in every case.
function Resolve-SitSelection
{
    param(
        [Parameter(Mandatory = $true)]$Taxonomy,
        [string[]]$ViewIds,
        [string[]]$Tags,
        [string[]]$Jurisdictions,
        [string[]]$ExcludeTags,
        [string[]]$Names
    )

    $F_AllSits = @($Taxonomy.Document.sits)
    $F_Selected = @()
    $F_Basis = ''

    if ($Names)
    {
        # Explicit names. Anything not in the taxonomy is still carried through as a selection, so a
        # custom or brand-new tenant SIT can be reported on; it is simply untagged.
        $F_Basis = "explicit SIT names ($($Names.Count))"
        foreach ($F_Name in $Names)
        {
            $F_Match = $F_AllSits | Where-Object { $_.name -eq $F_Name } | Select-Object -First 1
            if ($F_Match)
            {
                $F_Selected += $F_Match
            }
            else
            {
                $F_Selected += [PSCustomObject]@{
                    id                    = $null
                    name                  = $F_Name
                    recommendedConfidence = $null
                    tags                  = @()
                    note                  = 'Not present in the taxonomy.'
                }
            }
        }
    }
    elseif ($Tags)
    {
        $F_Basis = "tag filter ($($Tags -join ' AND '))"
        $F_Filter = [PSCustomObject]@{ allOf = $Tags; noneOf = @() }
        $F_Selected = @($F_AllSits | Where-Object { Test-SitFilter -Tags @($_.tags) -Filter $F_Filter })
    }
    else
    {
        $F_Basis = "view(s): $($ViewIds -join ', ')"
        $F_Seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($F_ViewId in $ViewIds)
        {
            $F_View = $Taxonomy.Document.views | Where-Object { $_.id -eq $F_ViewId } | Select-Object -First 1
            if (-not $F_View)
            {
                $F_Available = (@($Taxonomy.Document.views | ForEach-Object { $_.id }) -join ', ')
                throw "View '$F_ViewId' is not defined in the taxonomy. Available views: $F_Available"
            }
            foreach ($F_Sit in $F_AllSits)
            {
                if (-not (Test-SitFilter -Tags @($F_Sit.tags) -Filter $F_View.filter)) { continue }
                if ($F_Seen.Add("$($F_Sit.id)")) { $F_Selected += $F_Sit }
            }
        }
    }

    # -Jurisdiction narrows the result with OR semantics across the listed jurisdictions.
    if ($Jurisdictions -and -not $Names)
    {
        $F_JTags = @($Jurisdictions | ForEach-Object { ConvertTo-JurisdictionTag $_ })
        $F_Basis = "$F_Basis, jurisdiction in ($($F_JTags -join ' OR '))"
        $F_Selected = @($F_Selected | Where-Object {
            $F_SitTags = @($_.tags)
            $F_Hit = $false
            foreach ($F_JTag in $F_JTags) { if ($F_SitTags -contains $F_JTag) { $F_Hit = $true; break } }
            $F_Hit
        })
    }

    # -ExcludeTag always applies last.
    if ($ExcludeTags)
    {
        $F_Basis = "$F_Basis, excluding ($($ExcludeTags -join ', '))"
        $F_Selected = @($F_Selected | Where-Object {
            $F_SitTags = @($_.tags)
            $F_Hit = $false
            foreach ($F_XTag in $ExcludeTags) { if ($F_SitTags -contains $F_XTag) { $F_Hit = $true; break } }
            -not $F_Hit
        })
    }

    [PSCustomObject]@{
        Sits  = @($F_Selected | Sort-Object name)
        Basis = $F_Basis
    }
}

# Pull the value of a single-cardinality dimension off a tag list, for grouping and display.
function Get-TagValue
{
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Tags,
        [Parameter(Mandatory = $true)][string]$Dimension
    )

    foreach ($F_Tag in $Tags)
    {
        if ($F_Tag -like "$Dimension`:*") { return $F_Tag.Substring($Dimension.Length + 1) }
    }
    return $null
}

# All values of a multi-cardinality dimension.
function Get-TagValues
{
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Tags,
        [Parameter(Mandatory = $true)][string]$Dimension
    )

    @($Tags | Where-Object { $_ -like "$Dimension`:*" } | ForEach-Object { $_.Substring($Dimension.Length + 1) })
}

# ---------------------------------------------------------------------------
# Connection handling
# ---------------------------------------------------------------------------

# The Security & Compliance endpoint is reached through the same module as Exchange Online, so the two
# sessions are told apart by their ConnectionUri.
function Get-IppsConnection
{
    Get-ConnectionInformation -ErrorAction SilentlyContinue |
        Where-Object { $_.State -eq 'Connected' -and $_.ConnectionUri -match 'compliance|protection' } |
        Select-Object -First 1
}

function Connect-PurviewComplianceSession
{
    param([switch]$Silent)

    $F_Session = Get-IppsConnection
    if ($F_Session -and -not $Silent)
    {
        Write-Host "Existing Security & Compliance session detected:" -ForegroundColor Yellow
        Write-Host "  Account   : $($F_Session.UserPrincipalName)" -ForegroundColor Yellow
        Write-Host "  TenantId  : $($F_Session.TenantID)" -ForegroundColor Yellow
        Write-Host ""
        $F_Choice = Read-Host "Use existing Security & Compliance session? [Y] Yes  [N] Disconnect and reconnect  (Default: Y)"
        if ($F_Choice -eq 'N')
        {
            Disconnect-ExchangeOnline -Confirm:$false
            Connect-IPPSSession -ShowBanner:$false
        }
    }
    elseif (-not $F_Session)
    {
        if (-not $Silent)
        {
            Write-Host "Connecting to Security & Compliance PowerShell (Connect-IPPSSession)..." -ForegroundColor Cyan
        }
        Connect-IPPSSession -ShowBanner:$false
    }

    Get-IppsConnection
}

# Is this failure worth reconnecting for?
#
# Getting this wrong costs both ways. A false positive burns a reconnect (which can surface an
# interactive sign-in prompt mid-run) and retries a genuine service fault as though it were an auth
# problem, burying the real cause. A false negative fails a run that could have recovered.
#
# So state is the primary test, not text: if there is no live Security & Compliance session any more,
# it is an auth failure whatever the message said. Message matching is only the fallback for the case
# where the token has expired but the session object still reports Connected, and the patterns are
# deliberately specific - an unanchored '401' matches item counts, GUIDs and URLs.
function Test-AuthFailure
{
    param([Parameter(Mandatory = $true)]$ErrorRecord)

    if (-not (Get-IppsConnection)) { return $true }

    $F_Text = "$($ErrorRecord.Exception.Message) $($ErrorRecord.FullyQualifiedErrorId)"

    $F_AuthPatterns = @(
        'access token has expired'
        'token (has )?expired'
        'session (has )?expired'
        '\bexpired\b[^.]{0,40}\btoken\b'
        '\btoken\b[^.]{0,40}\bexpired\b'
        '\bunauthori[sz]ed\b'
        '\bunauthenticated\b'
        'authentication (failed|error)'
        'access is denied'
        'access denied'
        're-?authenticat'
        'Connect-IPPSSession'
        '\(401\)'
        '\b401\s+unauthori[sz]ed\b'
        '\(403\)'
        'invalid[_ ]grant'
        'credentials? (have )?expired'
    )

    foreach ($F_Pattern in $F_AuthPatterns)
    {
        if ($F_Text -imatch $F_Pattern) { return $true }
    }

    return $false
}

# ---------------------------------------------------------------------------
# Census
# ---------------------------------------------------------------------------

# One SIT in one workload. Reads TotalCount from the summary object only.
#
# -PageSize 1 is the smallest request that still returns the summary, and -ConfidenceLevel is
# deliberately never passed: TotalCount is scoped by TagName, TagType, Workload and SiteUrl /
# UserPrincipalName, but NOT by confidence. Passing a confidence filter here would return a
# confidence-blind TotalCount alongside a filtered record set, and reporting that number as
# "high confidence items" would be wrong.
function Get-SitWorkloadCount
{
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$WorkloadCode,
        [Parameter(Mandatory = $true)][int]$MaxRetries
    )

    $F_Attempt = 0
    while ($true)
    {
        $F_Attempt++
        try
        {
            $F_Result = Export-ContentExplorerData -TagType SensitiveInformationType -TagName $Name -Workload $WorkloadCode -PageSize 1 -ErrorAction Stop

            # The cmdlet returns an array whose element [0] is the summary. A tag with no matches can
            # come back as the summary alone, or as $null on some service responses.
            if (-not $F_Result)
            {
                return [PSCustomObject]@{ TotalCount = 0; MorePagesAvailable = $false; Error = $null }
            }

            $F_Summary = @($F_Result)[0]
            $F_Count = 0
            if ($null -ne $F_Summary.TotalCount) { $F_Count = [long]$F_Summary.TotalCount }

            return [PSCustomObject]@{
                TotalCount         = $F_Count
                MorePagesAvailable = [bool]$F_Summary.MorePagesAvailable
                Error              = $null
            }
        }
        catch
        {
            if ((Test-AuthFailure -ErrorRecord $_) -and $F_Attempt -le $MaxRetries)
            {
                Write-Warning "  Session appears to have expired. Reconnecting and retrying '$Name' / $WorkloadCode..."
                try
                {
                    Connect-PurviewComplianceSession -Silent | Out-Null
                }
                catch
                {
                    return [PSCustomObject]@{ TotalCount = $null; MorePagesAvailable = $false; Error = "Reconnect failed: $($_.Exception.Message)" }
                }
                continue
            }

            if ($F_Attempt -le $MaxRetries)
            {
                Start-Sleep -Seconds ([math]::Min(30, [math]::Pow(2, $F_Attempt)))
                continue
            }

            return [PSCustomObject]@{ TotalCount = $null; MorePagesAvailable = $false; Error = $_.Exception.Message }
        }
    }
}

# Written after every few results, so it must never be observed half-written. Writing straight to the
# destination means a process killed mid-write leaves a truncated file - which is exactly the failure
# checkpointing exists to survive, and -ResumeFrom parses this file on startup. So: write a temp
# sibling, flush it, then swap it into place. A kill leaves either the previous complete checkpoint or
# the new complete one, never a partial.
function Save-Checkpoint
{
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$Rows
    )

    $F_Temp = "$Path.writing"
    try
    {
        $F_Payload = [PSCustomObject]@{
            Generated = (Get-Date).ToString('o')
            Rows      = @($Rows | Select-Object SitName, SitId, Workload, TotalCount, Error)
        }

        # WriteAllText completes the file before the swap; Out-File would leave the handle open.
        # UTF8Encoding($false) so no byte order mark: this file is machine-read, by -ResumeFrom and
        # potentially by other tooling, and a BOM breaks strict JSON parsers.
        [System.IO.File]::WriteAllText($F_Temp, ($F_Payload | ConvertTo-Json -Depth 5), [System.Text.UTF8Encoding]::new($false))

        if (Test-Path -LiteralPath $Path)
        {
            # Replace is atomic where the filesystem supports it. The third argument is the backup
            # path and must be a genuine null to skip the backup - passing $null hands .NET an empty
            # string instead, which throws and silently leaves the previous checkpoint in place.
            [System.IO.File]::Replace($F_Temp, $Path, [NullString]::Value)
        }
        else
        {
            [System.IO.File]::Move($F_Temp, $Path)
        }

        return $true
    }
    catch
    {
        Write-Warning "Could not write checkpoint to '$Path': $($_.Exception.Message)"
        if (Test-Path -LiteralPath $F_Temp)
        {
            try { Remove-Item -LiteralPath $F_Temp -Force -ErrorAction Stop } catch { }
        }
        return $false
    }
}

# ---------------------------------------------------------------------------
# Full mode: item-level collection
# ---------------------------------------------------------------------------

# One SHA256 instance for the run rather than one per row.
$S_Sha256 = [System.Security.Cryptography.SHA256]::Create()

# Item names are the single most sensitive field in the export. For Exchange the name IS the subject
# line, and subjects routinely restate the content the scan was looking for. A truncated hash keeps
# rows with identical names groupable - "how many distinct subjects" is still answerable - while
# disclosing nothing about what they said.
function Get-DisplayItemName
{
    param([AllowNull()][string]$Value, [bool]$Reveal)

    if ([string]::IsNullOrEmpty($Value)) { return '' }
    if ($Reveal) { return $Value }

    $F_Bytes = $S_Sha256.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Value))
    $F_Hex = ($F_Bytes[0..5] | ForEach-Object { $_.ToString('x2') }) -join ''
    return "[redacted:$F_Hex]"
}

# Deduplication key. FileUrl alone is a genuine unique key for OneDrive and SharePoint - one path, one
# file. For Exchange and Teams it is a MESSAGE ID, and the same message appears once per mailbox
# holding a copy, so an email to eight internal recipients yields up to nine records sharing one
# FileUrl and differing only by FileSourceUrl. Deduplicating those on FileUrl alone would collapse
# genuinely separate exposures.
function Get-ItemDedupKey
{
    param([string]$WorkloadCode, [string]$ItemId, [string]$Container)

    if ($WorkloadCode -eq 'EXO' -or $WorkloadCode -eq 'Teams') { return "$ItemId|$Container" }
    return $ItemId
}

# Turn the three cumulative API counts into three mutually exclusive bands.
#
# LowConfidenceMatch / MediumConfidenceMatch / HighConfidenceMatch are THRESHOLDS, not buckets: each
# is the count of matches at or above that confidence, so Low >= Medium >= High always holds and
# summing the raw three roughly triples the true total.
#
# Clamping matters for the invariant. If the service ever returns an ordering violation, a negative
# band would make the three columns stop summing to the total. So bands are clamped at zero and the
# total is then computed FROM the bands, which keeps High + MediumBand + LowBand = Total true on every
# row without exception. In the normal case that total is identical to the raw LowConfidenceMatch.
# BandsAdjusted flags the rows where it is not.
function Get-ConfidenceBands
{
    param(
        [Parameter(Mandatory = $true)]$Entry
    )

    $F_Low = 0; $F_Med = 0; $F_High = 0
    if ($null -ne $Entry.LowConfidenceMatch) { $F_Low = [long]$Entry.LowConfidenceMatch }
    if ($null -ne $Entry.MediumConfidenceMatch) { $F_Med = [long]$Entry.MediumConfidenceMatch }
    if ($null -ne $Entry.HighConfidenceMatch) { $F_High = [long]$Entry.HighConfidenceMatch }

    $F_HighBand = $F_High
    $F_MedBand = $F_Med - $F_High
    $F_LowBand = $F_Low - $F_Med

    $F_Adjusted = $false
    if ($F_HighBand -lt 0) { $F_HighBand = 0; $F_Adjusted = $true }
    if ($F_MedBand -lt 0) { $F_MedBand = 0; $F_Adjusted = $true }
    if ($F_LowBand -lt 0) { $F_LowBand = 0; $F_Adjusted = $true }

    # Exactly 500 in Low or Medium appears to be a per-item per-SIT cap. A value of 524 has also been
    # observed, so it is not a hard global cap - hence a flag rather than a claim. The flag is a
    # separate boolean and never leaks into the numeric columns, which would break pivots.
    $F_Truncated = ($F_Low -eq 500 -or $F_Med -eq 500)

    [PSCustomObject]@{
        HighMatches       = $F_HighBand
        MediumBandMatches = $F_MedBand
        LowBandMatches    = $F_LowBand
        TotalMatches      = ($F_HighBand + $F_MedBand + $F_LowBand)
        Truncated         = $F_Truncated
        BandsAdjusted     = $F_Adjusted
    }
}

# Page one SIT in one workload, flattening each record into item x SIT rows as it goes rather than
# holding the raw records. Rows are appended to the shared sink; the shared seen-set deduplicates.
#
# Two things this has to get right, both from the export schema:
#   - Records are ITEM-centric, not tag-centric. Querying one SIT returns every item carrying it, and
#     each record lists ALL classifications on that item - commonly 50+. So a single query produces
#     rows for many SITs, and the same item arrives again from other SIT queries. Hence the seen-set.
#   - Rows are emitted ONLY for SITs that were actually queried. An item returned by the SSN query
#     might also carry a Polish national ID, but since that SIT was never queried we have only the
#     subset of its items that happen to also hold an SSN. Emitting it would imply coverage the run
#     does not have. Restricting to queried SITs keeps every row's coverage complete.
function Invoke-SitItemCollection
{
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$WorkloadCode,
        [Parameter(Mandatory = $true)][hashtable]$QueriedSits,
        [Parameter(Mandatory = $true)]$RowSink,
        [Parameter(Mandatory = $true)]$SeenKeys,
        [Parameter(Mandatory = $true)][int]$MaxRetries,
        [Parameter(Mandatory = $true)][int]$DelayMs,
        [Parameter(Mandatory = $true)][bool]$RevealNames
    )

    $F_Cookie = $null
    $F_Pages = 0
    $F_Records = 0
    $F_NewRows = 0
    $F_Error = $null
    $F_Attempt = 0

    do
    {
        try
        {
            if ($F_Cookie)
            {
                $F_Result = Export-ContentExplorerData -TagType SensitiveInformationType -TagName $Name -Workload $WorkloadCode -PageSize $script:S_FullPageSize -PageCookie $F_Cookie -ErrorAction Stop
            }
            else
            {
                $F_Result = Export-ContentExplorerData -TagType SensitiveInformationType -TagName $Name -Workload $WorkloadCode -PageSize $script:S_FullPageSize -ErrorAction Stop
            }
            $F_Attempt = 0
        }
        catch
        {
            $F_Attempt++
            if ((Test-AuthFailure -ErrorRecord $_) -and $F_Attempt -le $MaxRetries)
            {
                Write-Warning "  Session appears to have expired. Reconnecting and resuming '$Name' / $WorkloadCode from the last page cookie..."
                try { Connect-PurviewComplianceSession -Silent | Out-Null }
                catch { $F_Error = "Reconnect failed: $($_.Exception.Message)"; break }
                continue
            }
            if ($F_Attempt -le $MaxRetries)
            {
                Start-Sleep -Seconds ([math]::Min(30, [math]::Pow(2, $F_Attempt)))
                continue
            }
            $F_Error = $_.Exception.Message
            break
        }

        $F_Pages++
        if (-not $F_Result) { break }
        $F_Summary = @($F_Result)[0]

        # Guard the slice: $F_Result[1..0] when RecordsReturned is 0 resolves as elements 1 and 0
        # reversed, injecting the summary object into the record set.
        if ($F_Summary.RecordsReturned -gt 0)
        {
            foreach ($F_Rec in @($F_Result[1..$F_Summary.RecordsReturned]))
            {
                $F_Records++

                $F_ItemId = "$($F_Rec.FileUrl)"
                $F_Container = "$($F_Rec.FileSourceUrl)"
                $F_Key = Get-ItemDedupKey -WorkloadCode $WorkloadCode -ItemId $F_ItemId -Container $F_Container

                $F_Data = $null
                if ($F_Rec.SensitiveInfoTypesData)
                {
                    try { $F_Data = $F_Rec.SensitiveInfoTypesData | ConvertFrom-Json }
                    catch { continue }
                }
                if (-not $F_Data) { continue }

                $F_Modified = ''
                if ($F_Rec.LastModifiedTime)
                {
                    try { $F_Modified = ([datetime]$F_Rec.LastModifiedTime).ToString('yyyy-MM-dd HH:mm:ss') } catch { $F_Modified = "$($F_Rec.LastModifiedTime)" }
                }
                $F_DisplayName = Get-DisplayItemName -Value "$($F_Rec.FileName)" -Reveal $RevealNames

                foreach ($F_Entry in @($F_Data))
                {
                    $F_SitId = "$($F_Entry.Id)"
                    if (-not $F_SitId) { continue }
                    if (-not $QueriedSits.ContainsKey($F_SitId)) { continue }

                    # The same item arrives from every SIT query that matches it, so this pair may
                    # already be recorded.
                    if (-not $SeenKeys.Add("$F_Key|$F_SitId")) { continue }

                    $F_Sit = $QueriedSits[$F_SitId]
                    $F_Bands = Get-ConfidenceBands -Entry $F_Entry

                    $RowSink.Add([PSCustomObject]@{
                        Workload              = $WorkloadCode
                        Container             = $F_Container
                        ItemId                = $F_ItemId
                        ItemName              = $F_DisplayName
                        ItemKey               = $F_Key
                        SensitiveInfoType     = $F_Sit.Name
                        SitId                 = $F_SitId
                        Class                 = (Get-TagValue -Tags $F_Sit.Tags -Dimension 'class')
                        Pii                   = (Get-TagValue -Tags $F_Sit.Tags -Dimension 'pii')
                        Tier                  = (Get-TagValue -Tags $F_Sit.Tags -Dimension 'tier')
                        Jurisdictions         = ((Get-TagValues -Tags $F_Sit.Tags -Dimension 'jurisdiction') -join '; ')
                        Traits                = ((Get-TagValues -Tags $F_Sit.Tags -Dimension 'trait') -join '; ')
                        RecommendedConfidence = $F_Sit.RecommendedConfidence
                        HighMatches           = $F_Bands.HighMatches
                        MediumBandMatches     = $F_Bands.MediumBandMatches
                        LowBandMatches        = $F_Bands.LowBandMatches
                        TotalMatches          = $F_Bands.TotalMatches
                        MatchesTruncated      = $F_Bands.Truncated
                        BandsAdjusted         = $F_Bands.BandsAdjusted
                        SensitivityLabel      = "$($F_Rec.SensitivityLabel)"
                        RetentionLabel        = "$($F_Rec.RetentionLabel)"
                        LastModifiedTime      = $F_Modified
                    })
                    $F_NewRows++
                }
            }
        }

        $F_PrevCookie = $F_Cookie
        $F_Cookie = "$($F_Summary.PageCookie)"
        $F_More = [bool]$F_Summary.MorePagesAvailable

        # Loop on MorePagesAvailable and PageCookie only, NEVER on RecordsReturned. A page returns data
        # from 5 folders or PageSize, whichever completes first, so an empty page while more remain is
        # normal; stopping there would silently truncate the export.
        if ($F_More -and $F_Cookie -and $F_PrevCookie -and $F_Cookie -eq $F_PrevCookie)
        {
            $F_Error = "Paging stalled: the service returned the same PageCookie twice while reporting MorePagesAvailable."
            break
        }

        if ($DelayMs -gt 0) { Start-Sleep -Milliseconds $DelayMs }
    }
    while ($F_More -and $F_Cookie)

    [PSCustomObject]@{
        Pages   = $F_Pages
        Records = $F_Records
        NewRows = $F_NewRows
        Error   = $F_Error
    }
}

# ---------------------------------------------------------------------------
# Assumption verification (-VerifyAssumptions)
#
# Settles the two entries in the assumption register against a live tenant. Read-only, and handles
# item records without ever surfacing their content: only FileUrl and FileSourceUrl are touched, and
# only to build hash sets for counting. FileName is never read, so EXO subject lines - which routinely
# restate the sensitive content - cannot leak into the console or the findings file.
# ---------------------------------------------------------------------------

# Page a SIT in one workload to exhaustion and return identifier sets. Loops on MorePagesAvailable and
# PageCookie only, never on RecordsReturned: a page can legitimately come back empty while more pages
# remain, and stopping there would silently undercount.
function Get-ItemKeySet
{
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$WorkloadCode,
        [Parameter(Mandatory = $true)][int]$MaxPages,
        [Parameter(Mandatory = $true)][int]$DelayMs
    )

    $F_ItemIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $F_Compound = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $F_Containers = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    $F_Cookie = $null
    $F_Pages = 0
    $F_RecordsSeen = 0
    $F_TotalCount = $null
    $F_Truncated = $false
    $F_Error = $null

    try
    {
        do
        {
            $F_Pages++
            if ($F_Pages -gt $MaxPages) { $F_Truncated = $true; break }

            if ($F_Cookie)
            {
                $F_Result = Export-ContentExplorerData -TagType SensitiveInformationType -TagName $Name -Workload $WorkloadCode -PageSize 1000 -PageCookie $F_Cookie -ErrorAction Stop
            }
            else
            {
                $F_Result = Export-ContentExplorerData -TagType SensitiveInformationType -TagName $Name -Workload $WorkloadCode -PageSize 1000 -ErrorAction Stop
            }

            if (-not $F_Result) { break }
            $F_Summary = @($F_Result)[0]
            if ($null -eq $F_TotalCount -and $null -ne $F_Summary.TotalCount) { $F_TotalCount = [long]$F_Summary.TotalCount }

            # Guard the slice. $F_Result[1..0] with RecordsReturned = 0 resolves as elements 1 and 0
            # reversed, injecting the summary object into the record set.
            if ($F_Summary.RecordsReturned -gt 0)
            {
                foreach ($F_Rec in @($F_Result[1..$F_Summary.RecordsReturned]))
                {
                    $F_RecordsSeen++
                    $F_Url = "$($F_Rec.FileUrl)"
                    $F_Src = "$($F_Rec.FileSourceUrl)"
                    if ($F_Url) { [void]$F_ItemIds.Add($F_Url) }
                    if ($F_Src) { [void]$F_Containers.Add($F_Src) }
                    [void]$F_Compound.Add("$F_Url|$F_Src")
                }
            }

            $F_PrevCookie = $F_Cookie
            $F_Cookie = "$($F_Summary.PageCookie)"
            $F_More = [bool]$F_Summary.MorePagesAvailable

            # Stall guard. Looping on MorePagesAvailable alone is correct, but it trusts the service to
            # advance the cookie. If it ever hands back the same cookie while still claiming more pages,
            # this loop would spin until the page ceiling and report a false result. Stop and say so.
            if ($F_More -and $F_Cookie -and $F_PrevCookie -and $F_Cookie -eq $F_PrevCookie)
            {
                $F_Error = "Paging stalled: the service returned the same PageCookie twice while reporting MorePagesAvailable. Collected $F_RecordsSeen record(s) before stopping."
                break
            }

            if ($DelayMs -gt 0) { Start-Sleep -Milliseconds $DelayMs }
        }
        while ($F_More -and $F_Cookie)
    }
    catch
    {
        $F_Error = $_.Exception.Message
    }

    [PSCustomObject]@{
        Workload        = $WorkloadCode
        TotalCount      = $F_TotalCount
        Pages           = $F_Pages
        RecordsSeen     = $F_RecordsSeen
        DistinctItemIds = $F_ItemIds
        DistinctCopies  = $F_Compound
        Containers      = $F_Containers
        Truncated       = $F_Truncated
        Error           = $F_Error
    }
}

function Invoke-AssumptionVerification
{
    param(
        [Parameter(Mandatory = $true)]$Candidates,
        [string]$PreferredSit,
        [Parameter(Mandatory = $true)][int]$MaxPages,
        [Parameter(Mandatory = $true)][int]$DelayMs,
        [Parameter(Mandatory = $true)][string]$FindingsPath,
        [Parameter(Mandatory = $true)][string[]]$Workloads
    )

    Write-Host ""
    Write-Host "Assumption verification" -ForegroundColor Cyan
    Write-Host "-----------------------" -ForegroundColor Cyan
    Write-Host "Read-only. Item records are paged to compare identifiers; item names, subjects and" -ForegroundColor DarkGray
    Write-Host "match content are never read, printed or written." -ForegroundColor DarkGray
    Write-Host ""

    # --- Choose a probe SIT -----------------------------------------------------------------
    # Wanted: non-zero in EXO and Teams (so A2 is testable at all) and small enough to page inside
    # the ceiling. A census call per candidate is cheap - TotalCount only, no paging.
    $F_Probe = $null
    $F_ExoCount = 0
    $F_TeamsCount = 0

    if ($PreferredSit)
    {
        $F_Probe = $PreferredSit
        # Distinguish "returned zero" from "the call failed". Treating a failure as an absence would
        # let a broken query masquerade as a clean finding.
        $F_ExoResp = Get-SitWorkloadCount -Name $F_Probe -WorkloadCode 'EXO' -MaxRetries 1
        $F_TeamsResp = Get-SitWorkloadCount -Name $F_Probe -WorkloadCode 'Teams' -MaxRetries 1
        if ($F_ExoResp.Error) { Write-Warning "EXO count for '$F_Probe' FAILED: $($F_ExoResp.Error). Treating as unavailable, not as zero." }
        if ($F_TeamsResp.Error) { Write-Warning "Teams count for '$F_Probe' FAILED: $($F_TeamsResp.Error). Treating as unavailable, not as zero." }
        $F_ExoCount = $F_ExoResp.TotalCount
        $F_TeamsCount = $F_TeamsResp.TotalCount
        if ($null -eq $F_ExoCount) { $F_ExoCount = 0 }
        if ($null -eq $F_TeamsCount) { $F_TeamsCount = 0 }
        Write-Host "Probe SIT (specified): $F_Probe" -ForegroundColor Green
    }
    else
    {
        Write-Host "Selecting a probe SIT (census calls only)..." -ForegroundColor Cyan
        $F_Ceiling = $MaxPages * 1000

        # A2 compares two sets by intersecting them, so it is only meaningful if BOTH sides carry
        # enough items for an empty intersection to mean something. An earlier version accepted any
        # count above zero and duly picked a SIT with a single Teams item against a large EXO set,
        # which reported "no overlap" on evidence that could not have shown one. Hence a floor.
        $F_MinForA2 = 25

        $F_Fallback = $null; $F_FallbackE = 0; $F_FallbackT = 0
        foreach ($F_Cand in $Candidates)
        {
            $F_E = (Get-SitWorkloadCount -Name $F_Cand.Name -WorkloadCode 'EXO' -MaxRetries 1).TotalCount
            $F_T = (Get-SitWorkloadCount -Name $F_Cand.Name -WorkloadCode 'Teams' -MaxRetries 1).TotalCount
            if ($null -eq $F_E) { $F_E = 0 }
            if ($null -eq $F_T) { $F_T = 0 }
            Write-Host ("  {0,-52} EXO {1,8}  Teams {2,8}" -f $F_Cand.Name, $F_E, $F_T) -ForegroundColor DarkGray

            if ($F_E -ge $F_MinForA2 -and $F_T -ge $F_MinForA2 -and $F_E -lt $F_Ceiling -and $F_T -lt $F_Ceiling)
            {
                $F_Probe = $F_Cand.Name; $F_ExoCount = $F_E; $F_TeamsCount = $F_T
                break
            }
            # Otherwise keep the first SIT with any usable data, as an A1-only fallback.
            if (-not $F_Fallback -and (($F_E -gt 0 -and $F_E -lt $F_Ceiling) -or ($F_T -gt 0 -and $F_T -lt $F_Ceiling)))
            {
                $F_Fallback = $F_Cand.Name; $F_FallbackE = $F_E; $F_FallbackT = $F_T
            }
        }

        if (-not $F_Probe -and $F_Fallback)
        {
            $F_Probe = $F_Fallback; $F_ExoCount = $F_FallbackE; $F_TeamsCount = $F_FallbackT
            Write-Host ""
            Write-Warning "No SIT in this selection has at least $F_MinForA2 items in BOTH Exchange and Teams, so A2 cannot be tested meaningfully. Proceeding with '$F_Probe' for A1 only."
            Write-Warning "To test A2, name a type with real Teams volume using -VerifySensitiveInfoType, or widen the selection with -View all."
        }
    }

    if (-not $F_Probe)
    {
        Write-Warning "No SIT in the selection returned a workable non-zero count. Nothing to probe."
        Write-Warning "Re-run with -View all, or name a SIT with -VerifySensitiveInfoType."
        return
    }

    Write-Host ""
    Write-Host "Probe SIT: $F_Probe  (EXO $F_ExoCount, Teams $F_TeamsCount)" -ForegroundColor Green

    # --- A1: is TotalCount complete, or sampled? ---------------------------------------------
    #
    # Tested PER WORKLOAD, not once. An earlier version probed a single workload and reported the
    # result as though it settled the assumption everywhere; a later run then found EXO matching
    # exactly while Teams was 41% short. The behaviour is evidently workload-dependent, so each one
    # in scope gets its own verdict and the overall result is the worst of them.
    Write-Host ""
    Write-Host "A1  TotalCount completeness, per workload" -ForegroundColor Cyan
    Write-Host "    Comparing each -PageSize 1 TotalCount against an exhaustive paged export." -ForegroundColor DarkGray
    Write-Host ""

    $F_A1Results = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($F_Wl in $Workloads)
    {
        $F_Resp = Get-SitWorkloadCount -Name $F_Probe -WorkloadCode $F_Wl -MaxRetries 1
        if ($F_Resp.Error)
        {
            Write-Host ("  {0,-6} census call FAILED: {1}" -f $F_Wl, $F_Resp.Error) -ForegroundColor Red
            $F_A1Results.Add([PSCustomObject]@{ Workload = $F_Wl; CensusTotalCount = $null; PagedRecords = $null; PagedDistinct = $null; Pages = 0; HitPageCeiling = $false; Verdict = 'ERROR'; Detail = "Census call failed: $($F_Resp.Error)"; GrainNote = '' })
            continue
        }

        $F_Cc = $F_Resp.TotalCount
        if ($null -eq $F_Cc) { $F_Cc = 0 }
        if ($F_Cc -le 0)
        {
            Write-Host ("  {0,-6} TotalCount 0 - nothing to compare" -f $F_Wl) -ForegroundColor DarkGray
            $F_A1Results.Add([PSCustomObject]@{ Workload = $F_Wl; CensusTotalCount = 0; PagedRecords = $null; PagedDistinct = $null; Pages = 0; HitPageCeiling = $false; Verdict = 'NO-DATA'; Detail = 'TotalCount is zero in this workload, so completeness cannot be tested here.'; GrainNote = '' })
            continue
        }

        $F_Pg = Get-ItemKeySet -Name $F_Probe -WorkloadCode $F_Wl -MaxPages $MaxPages -DelayMs $DelayMs
        $F_V = 'INCONCLUSIVE'; $F_D = ''; $F_G = ''

        if ($F_Pg.Error)
        {
            $F_D = "Paged export failed: $($F_Pg.Error)"
        }
        elseif ($F_Pg.Truncated)
        {
            $F_D = "Hit the $MaxPages page ceiling before exhausting the result set. Raise -VerifyMaxPages and re-run."
        }
        else
        {
            if ($F_Cc -eq $F_Pg.RecordsSeen)
            {
                $F_V = 'HOLDS'
                $F_D = "TotalCount ($F_Cc) equals the records paged back ($($F_Pg.RecordsSeen)). Not truncated by folder sampling."
            }
            elseif ($F_Cc -lt $F_Pg.RecordsSeen)
            {
                $F_V = 'FAILS'
                $F_D = "TotalCount ($F_Cc) is LOWER than the records paged back ($($F_Pg.RecordsSeen)). Census figures understate this workload."
            }
            else
            {
                $F_Short = $F_Cc - $F_Pg.RecordsSeen
                $F_Pct = [math]::Round(($F_Short / $F_Cc) * 100, 1)
                $F_V = 'DIFFERS'
                $F_D = "TotalCount ($F_Cc) exceeds the records paged back ($($F_Pg.RecordsSeen)) by $F_Short ($F_Pct%), and paging ended on its own rather than hitting the ceiling. Either the census overstates this workload or the paged export is incomplete - the two disagree and neither should be quoted for this workload until it is understood."
            }

            if ($F_Pg.DistinctCopies.Count -lt $F_Pg.RecordsSeen)
            {
                $F_Gap = $F_Pg.RecordsSeen - $F_Pg.DistinctCopies.Count
                $F_GPct = if ($F_Pg.DistinctCopies.Count -gt 0) { [math]::Round(($F_Gap / $F_Pg.DistinctCopies.Count) * 100, 2) } else { 0 }
                $F_G = "$($F_Pg.RecordsSeen) records deduplicate to $($F_Pg.DistinctCopies.Count) distinct items ($F_Gap duplicate(s), $F_GPct% inflation), so counts here are occurrences rather than items."
            }
        }

        $F_Colour = switch ($F_V) { 'HOLDS' { 'Green' } 'FAILS' { 'Red' } 'DIFFERS' { 'Red' } default { 'Yellow' } }
        Write-Host ("  {0,-6} TotalCount {1,8}   records {2,8}   distinct {3,8}   {4}" -f $F_Wl, $F_Cc, $F_Pg.RecordsSeen, $F_Pg.DistinctCopies.Count, $F_V) -ForegroundColor $F_Colour
        if ($F_D) { Write-Host "         $F_D" -ForegroundColor DarkGray }
        if ($F_G) { Write-Host "         $F_G" -ForegroundColor DarkYellow }

        $F_A1Results.Add([PSCustomObject]@{
            Workload = $F_Wl; CensusTotalCount = $F_Cc
            PagedRecords = $F_Pg.RecordsSeen; PagedDistinct = $F_Pg.DistinctCopies.Count
            Pages = $F_Pg.Pages; HitPageCeiling = $F_Pg.Truncated
            Verdict = $F_V; Detail = $F_D; GrainNote = $F_G
        })
    }

    # Overall verdict is the worst observed. One workload disagreeing is enough to withhold the claim.
    $F_Tested = @($F_A1Results | Where-Object { $_.Verdict -in @('HOLDS', 'FAILS', 'DIFFERS') })
    $F_A1Verdict = if (@($F_Tested | Where-Object { $_.Verdict -eq 'FAILS' }).Count -gt 0) { 'FAILS' }
        elseif (@($F_Tested | Where-Object { $_.Verdict -eq 'DIFFERS' }).Count -gt 0) { 'DIFFERS' }
        elseif ($F_Tested.Count -gt 0) { 'HOLDS' }
        else { 'INCONCLUSIVE' }

    $F_A1Detail = if ($F_Tested.Count -eq 0)
    {
        'No workload returned data for this probe SIT, so completeness could not be tested. Choose a different SIT.'
    }
    else
    {
        "Tested $($F_Tested.Count) workload(s): " + (($F_Tested | ForEach-Object { "$($_.Workload)=$($_.Verdict)" }) -join ', ') + ". The assumption holds only where every tested workload reports HOLDS."
    }

    Write-Host ""
    Write-Host "    Overall A1 verdict : $F_A1Verdict" -ForegroundColor $(if ($F_A1Verdict -eq 'HOLDS') { 'Green' } elseif ($F_A1Verdict -eq 'INCONCLUSIVE') { 'Yellow' } else { 'Red' })
    Write-Host "    $F_A1Detail"

    # Keep a representative paged set for A2 to reuse, so it does not re-page the same workload.
    $F_A1Workload = if ($F_ExoCount -gt 0) { 'EXO' } else { 'Teams' }
    $F_Paged = $null


    # --- A2: do EXO and Teams return the same messages? ---------------------------------------
    Write-Host ""
    Write-Host "A2  EXO / Teams disjointness" -ForegroundColor Cyan
    Write-Host "    Intersecting the FileUrl (message ID) sets returned by each workload." -ForegroundColor DarkGray

    $F_A2Verdict = 'INCONCLUSIVE'
    $F_A2Detail = ''
    $F_Overlap = 0
    $F_ExoSet = $null
    $F_TeamsSet = $null

    if ($F_ExoCount -le 0 -or $F_TeamsCount -le 0)
    {
        $F_A2Detail = "Probe SIT has no data in both EXO ($F_ExoCount) and Teams ($F_TeamsCount), so the sets cannot be compared. Name a SIT present in both with -VerifySensitiveInfoType."
    }
    else
    {
        $F_ExoSet = Get-ItemKeySet -Name $F_Probe -WorkloadCode 'EXO' -MaxPages $MaxPages -DelayMs $DelayMs
        $F_TeamsSet = Get-ItemKeySet -Name $F_Probe -WorkloadCode 'Teams' -MaxPages $MaxPages -DelayMs $DelayMs

        if ($F_ExoSet.Error -or $F_TeamsSet.Error)
        {
            $F_A2Detail = "Paged export failed: $($F_ExoSet.Error) $($F_TeamsSet.Error)"
        }
        elseif ($F_ExoSet.Truncated -or $F_TeamsSet.Truncated)
        {
            $F_A2Detail = "Hit the $MaxPages page ceiling. A non-empty intersection would still be proof of overlap, but an empty one proves nothing. Raise -VerifyMaxPages."
        }
        else
        {
            $F_Intersect = [System.Collections.Generic.HashSet[string]]::new($F_ExoSet.DistinctItemIds, [StringComparer]::OrdinalIgnoreCase)
            $F_Intersect.IntersectWith($F_TeamsSet.DistinctItemIds)
            $F_Overlap = $F_Intersect.Count

            if ($F_Overlap -eq 0)
            {
                # An empty intersection only means something if both sides were big enough to have
                # produced one. One Teams item against a thousand EXO items cannot demonstrate
                # disjointness, and reporting HOLDS there would be a false pass.
                $F_Smallest = [math]::Min($F_ExoSet.DistinctItemIds.Count, $F_TeamsSet.DistinctItemIds.Count)
                if ($F_Smallest -lt 25)
                {
                    # Not worthless, but not conclusive either, and the distinction matters.
                    #
                    # The overlap that would actually damage a report is SYSTEMATIC: if Teams chat
                    # surfaces under EXO as a rule, then every Teams item should also appear in the EXO
                    # set, and even one Teams item failing to do so argues against it. What a small
                    # sample cannot rule out is OCCASIONAL overlap affecting a minority of messages.
                    $F_A2Verdict = 'WEAK'
                    $F_A2Detail = "No overlap found, but the smaller workload returned only $F_Smallest item(s) ($($F_ExoSet.DistinctItemIds.Count) EXO, $($F_TeamsSet.DistinctItemIds.Count) Teams). Read this as evidence AGAINST systematic overlap - if Teams items appeared under Exchange as a rule, these would have - but too small to rule out occasional overlap. Re-probe with a type carrying at least 25 items in both workloads to settle it."
                }
                else
                {
                    $F_A2Verdict = 'HOLDS'
                    $F_A2Detail = "No message ID appears under both workloads ($($F_ExoSet.DistinctItemIds.Count) EXO, $($F_TeamsSet.DistinctItemIds.Count) Teams). Summing EXO and Teams does not double count."
                }
            }
            else
            {
                $F_A2Verdict = 'FAILS'
                $F_A2Detail = "$F_Overlap message ID(s) appear under BOTH workloads. ItemsTotal double counts them whenever EXO and Teams are both in scope. Report EXO and Teams separately, or deduplicate on FileUrl across the two."
            }
        }
    }

    Write-Host "    Overlapping message IDs : $F_Overlap"
    Write-Host "    Verdict                 : $F_A2Verdict" -ForegroundColor $(if ($F_A2Verdict -eq 'HOLDS') { 'Green' } elseif ($F_A2Verdict -eq 'FAILS') { 'Red' } elseif ($F_A2Verdict -eq 'WEAK') { 'DarkYellow' } else { 'Yellow' })
    Write-Host "    $F_A2Detail"

    # --- Findings file. Counts and verdicts only; no identifiers, no item content. -------------
    $F_Findings = [PSCustomObject]@{
        Generated       = (Get-Date).ToString('o')
        ProbeSit        = $F_Probe
        PageCeiling     = $MaxPages
        A1_TotalCount   = [PSCustomObject]@{
            Assumption        = 'TotalCount at -PageSize 1 is complete, not truncated by folder sampling'
            ComparedAgainst   = 'PagedRecords (the population TotalCount counts), not PagedDistinct'
            OverallVerdict    = $F_A1Verdict
            Detail            = $F_A1Detail
            PerWorkload       = @($F_A1Results)
        }
        A2_Disjointness = [PSCustomObject]@{
            Assumption        = 'EXO and Teams return disjoint item sets, so ItemsTotal may sum them'
            ExoDistinctItems  = if ($F_ExoSet) { $F_ExoSet.DistinctItemIds.Count } else { $null }
            TeamsDistinctItems = if ($F_TeamsSet) { $F_TeamsSet.DistinctItemIds.Count } else { $null }
            OverlappingItems  = $F_Overlap
            HitPageCeiling    = if ($F_ExoSet -and $F_TeamsSet) { ($F_ExoSet.Truncated -or $F_TeamsSet.Truncated) } else { $null }
            Verdict           = $F_A2Verdict
            Detail            = $F_A2Detail
        }
    }

    try
    {
        # BOM-less, for the same reason as the checkpoint: this file is meant to be parsed.
        [System.IO.File]::WriteAllText($FindingsPath, ($F_Findings | ConvertTo-Json -Depth 6), [System.Text.UTF8Encoding]::new($false))
        Write-Host ""
        Write-Host "Findings written to: $FindingsPath" -ForegroundColor Green
    }
    catch
    {
        Write-Warning "Could not write findings to '$FindingsPath': $($_.Exception.Message)"
    }

    Write-Host ""
    Write-Host "Update the ASSUMPTION REGISTER in this script's header with these verdicts so the next" -ForegroundColor Yellow
    Write-Host "reader knows they are established rather than assumed." -ForegroundColor Yellow
}

# ---------------------------------------------------------------------------
# Inline SVG charts
#
# Hand-rolled rather than a charting library, because the HTML report must make no external request:
# it describes where personal data lives, and is often opened on a machine that should not be calling
# a CDN. Two shapes cover everything this report needs.
# ---------------------------------------------------------------------------

function ConvertTo-HtmlText
{
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return '' }
    [System.Net.WebUtility]::HtmlEncode($Text)
}

# SVG text does not wrap or clip to a box, so an over-long chart label runs off the left edge of the
# viewBox and gets cut. Truncate to what the label gutter can hold; the full text stays in the bar's
# <title> tooltip. Roughly 6.2px per character at 12px Segoe UI.
function Format-ChartLabel
{
    param(
        [Parameter(Mandatory = $true)][AllowNull()][string]$Label,
        [Parameter(Mandatory = $true)][int]$AvailableWidth
    )

    $F_MaxChars = [math]::Max(8, [int](($AvailableWidth - 8) / 6.2))
    if ($null -eq $Label) { return '' }
    if ($Label.Length -le $F_MaxChars) { return $Label }
    return ($Label.Substring(0, $F_MaxChars - 1) + [char]0x2026)
}

# Horizontal bar chart. Horizontal because SIT names are long and would be unreadable rotated.
function New-SvgBarChart
{
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Rows,   # Label, Value
        [string]$Colour = '#0b5394',
        [int]$LabelWidth = 260,
        [int]$Width = 900
    )

    if (-not $Rows -or $Rows.Count -eq 0)
    {
        return '<p class="muted">No data to chart.</p>'
    }

    $F_RowHeight = 26
    $F_BarHeight = 16
    $F_ValueWidth = 90
    $F_PlotWidth = $Width - $LabelWidth - $F_ValueWidth
    $F_Height = ($Rows.Count * $F_RowHeight) + 8

    $F_Max = 0
    foreach ($F_Row in $Rows) { if ($F_Row.Value -gt $F_Max) { $F_Max = $F_Row.Value } }
    if ($F_Max -le 0) { $F_Max = 1 }

    $F_Sb = [System.Text.StringBuilder]::new()
    [void]$F_Sb.AppendLine("<svg class=`"chart`" viewBox=`"0 0 $Width $F_Height`" width=`"100%`" height=`"$F_Height`" role=`"img`">")

    $F_Y = 4
    foreach ($F_Row in $Rows)
    {
        $F_BarWidth = [math]::Max(1, [int](($F_Row.Value / $F_Max) * $F_PlotWidth))
        $F_Label = ConvertTo-HtmlText (Format-ChartLabel -Label $F_Row.Label -AvailableWidth $LabelWidth)
        $F_Full = ConvertTo-HtmlText $F_Row.Label
        $F_TextY = $F_Y + 12

        [void]$F_Sb.AppendLine("  <text x=`"$($LabelWidth - 8)`" y=`"$F_TextY`" text-anchor=`"end`" class=`"svg-label`">$F_Label<title>$F_Full</title></text>")
        [void]$F_Sb.AppendLine("  <rect x=`"$LabelWidth`" y=`"$F_Y`" width=`"$F_BarWidth`" height=`"$F_BarHeight`" rx=`"2`" fill=`"$Colour`"><title>$F_Full - $('{0:N0}' -f $F_Row.Value)</title></rect>")
        [void]$F_Sb.AppendLine("  <text x=`"$($LabelWidth + $F_BarWidth + 8)`" y=`"$F_TextY`" class=`"svg-value`">$('{0:N0}' -f $F_Row.Value)</text>")

        $F_Y += $F_RowHeight
    }

    [void]$F_Sb.AppendLine('</svg>')
    $F_Sb.ToString()
}

# Stacked horizontal bars, one segment per workload. Stacking is legitimate here because the four
# workloads are disjoint: an item lives in exactly one of them.
function New-SvgStackedBarChart
{
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Rows,   # Label, Values[]
        [Parameter(Mandatory = $true)][string[]]$SeriesNames,
        [Parameter(Mandatory = $true)][string[]]$SeriesColours,
        [int]$LabelWidth = 260,
        [int]$Width = 900
    )

    if (-not $Rows -or $Rows.Count -eq 0)
    {
        return '<p class="muted">No data to chart.</p>'
    }

    $F_RowHeight = 26
    $F_BarHeight = 16
    $F_ValueWidth = 90
    $F_PlotWidth = $Width - $LabelWidth - $F_ValueWidth
    $F_Height = ($Rows.Count * $F_RowHeight) + 34

    $F_Max = 0
    foreach ($F_Row in $Rows)
    {
        $F_Sum = 0
        foreach ($F_V in $F_Row.Values) { $F_Sum += $F_V }
        if ($F_Sum -gt $F_Max) { $F_Max = $F_Sum }
    }
    if ($F_Max -le 0) { $F_Max = 1 }

    $F_Sb = [System.Text.StringBuilder]::new()
    [void]$F_Sb.AppendLine("<svg class=`"chart`" viewBox=`"0 0 $Width $F_Height`" width=`"100%`" height=`"$F_Height`" role=`"img`">")

    # Legend
    $F_LegendX = $LabelWidth
    for ($F_I = 0; $F_I -lt $SeriesNames.Count; $F_I++)
    {
        [void]$F_Sb.AppendLine("  <rect x=`"$F_LegendX`" y=`"4`" width=`"10`" height=`"10`" rx=`"2`" fill=`"$($SeriesColours[$F_I])`" />")
        [void]$F_Sb.AppendLine("  <text x=`"$($F_LegendX + 15)`" y=`"13`" class=`"svg-legend`">$(ConvertTo-HtmlText $SeriesNames[$F_I])</text>")
        $F_LegendX += 15 + ($SeriesNames[$F_I].Length * 7) + 16
    }

    $F_Y = 26
    foreach ($F_Row in $Rows)
    {
        $F_Label = ConvertTo-HtmlText (Format-ChartLabel -Label $F_Row.Label -AvailableWidth $LabelWidth)
        $F_Full = ConvertTo-HtmlText $F_Row.Label
        $F_TextY = $F_Y + 12
        [void]$F_Sb.AppendLine("  <text x=`"$($LabelWidth - 8)`" y=`"$F_TextY`" text-anchor=`"end`" class=`"svg-label`">$F_Label<title>$F_Full</title></text>")

        $F_X = $LabelWidth
        $F_Total = 0
        for ($F_I = 0; $F_I -lt $SeriesNames.Count; $F_I++)
        {
            $F_Value = $F_Row.Values[$F_I]
            $F_Total += $F_Value
            if ($F_Value -le 0) { continue }

            $F_SegWidth = [math]::Max(1, [int](($F_Value / $F_Max) * $F_PlotWidth))
            [void]$F_Sb.AppendLine("  <rect x=`"$F_X`" y=`"$F_Y`" width=`"$F_SegWidth`" height=`"$F_BarHeight`" fill=`"$($SeriesColours[$F_I])`"><title>$F_Full - $(ConvertTo-HtmlText $SeriesNames[$F_I]): $('{0:N0}' -f $F_Value)</title></rect>")
            $F_X += $F_SegWidth
        }

        [void]$F_Sb.AppendLine("  <text x=`"$($F_X + 8)`" y=`"$F_TextY`" class=`"svg-value`">$('{0:N0}' -f $F_Total)</text>")
        $F_Y += $F_RowHeight
    }

    [void]$F_Sb.AppendLine('</svg>')
    $F_Sb.ToString()
}

# ---------------------------------------------------------------------------
# Resolve taxonomy and selection
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "Purview Sensitive Information Type census" -ForegroundColor Cyan
Write-Host "-----------------------------------------" -ForegroundColor Cyan

if ($TaxonomyPath)
{
    if (-not (Test-Path -LiteralPath $TaxonomyPath -PathType Leaf))
    {
        throw "Taxonomy file not found: $TaxonomyPath"
    }
    $S_Taxonomy = Import-SitTaxonomy -Json (Get-Content -LiteralPath $TaxonomyPath -Raw) -Source $TaxonomyPath
}
else
{
    $S_Taxonomy = Import-SitTaxonomy -Json $S_EmbeddedTaxonomyJson -Source 'embedded'
}

Write-Host "Taxonomy         : $($S_Taxonomy.SitCount) SITs from $($S_Taxonomy.Source)" -ForegroundColor Green

# -Jurisdiction on its own reads as "everything for this country". Left alone it would narrow the
# DEFAULT view instead, and since us-core is US-only, -Jurisdiction NZ would select nothing at all.
# So widen the base to every SIT unless the caller actually named a view or another selector.
$S_ViewIds = $View
if ($Jurisdiction -and -not $PSBoundParameters.ContainsKey('View') -and -not $Tag -and -not $SensitiveInfoType)
{
    $S_ViewIds = @('all')
    Write-Verbose "-Jurisdiction was supplied without -View or -Tag; selecting from all SITs rather than the default us-core view."
}

$S_Selection = Resolve-SitSelection -Taxonomy $S_Taxonomy `
    -ViewIds $S_ViewIds -Tags $Tag -Jurisdictions $Jurisdiction -ExcludeTags $ExcludeTag -Names $SensitiveInfoType

[array]$S_SelectedSits = $S_Selection.Sits
if ($S_SelectedSits.Count -eq 0)
{
    throw "No SITs matched the selection ($($S_Selection.Basis)). Check the view id, tag spelling, or run with -ListOnly -View all to see what is available."
}

# Resolve the workload list. 'All' wins if present.
if ($Workload -contains 'All') { [array]$S_Workloads = @('EXO', 'ODB', 'SPO', 'Teams') }
else { [array]$S_Workloads = @($Workload | Select-Object -Unique) }

if ($TestMode)
{
    $S_SampleCount = [math]::Min($SampleSize, $S_SelectedSits.Count)
    Write-Host "TEST MODE        : sampling $S_SampleCount of $($S_SelectedSits.Count) selected SITs" -ForegroundColor Yellow
    [array]$S_SelectedSits = @($S_SelectedSits | Get-Random -Count $S_SampleCount | Sort-Object name)
}

$S_PlannedCalls = $S_SelectedSits.Count * $S_Workloads.Count

Write-Host "Selection        : $($S_Selection.Basis)" -ForegroundColor Green
Write-Host "SITs selected    : $($S_SelectedSits.Count)" -ForegroundColor Green
Write-Host "Workloads        : $($S_Workloads -join ', ')" -ForegroundColor Green
Write-Host "API calls        : $S_PlannedCalls" -ForegroundColor Green

$S_EstimateSeconds = $S_PlannedCalls * ($S_NominalSecondsPerCall + ($S_RequestDelayMilliseconds / 1000))
$S_EstimateText = [timespan]::FromSeconds($S_EstimateSeconds).ToString('hh\:mm\:ss')
Write-Host "Rough duration   : ~$S_EstimateText (estimate only)" -ForegroundColor Green

if ($S_EstimateSeconds -gt $S_TokenLifetimeSeconds)
{
    Write-Warning "This run is estimated to exceed the ~1 hour Security & Compliance token lifetime."
    Write-Warning "The script reconnects and retries automatically, and writes a checkpoint as it goes. If it does fail outright, restart it with -ResumeFrom <checkpoint>."
}

if ($ListOnly)
{
    Write-Host ""
    Write-Host "SITs that would be queried:" -ForegroundColor Cyan
    foreach ($S_Sit in $S_SelectedSits)
    {
        $S_TierValue = Get-TagValue -Tags @($S_Sit.tags) -Dimension 'tier'
        $S_ClassValue = Get-TagValue -Tags @($S_Sit.tags) -Dimension 'class'
        Write-Host ("  {0,-58} {1,-14} {2}" -f $S_Sit.name, $S_ClassValue, $S_TierValue)
    }
    Write-Host ""
    Write-Host "Dry run only - nothing was connected to and no data was read." -ForegroundColor Yellow
    return
}

# ---------------------------------------------------------------------------
# Connect and preflight
# ---------------------------------------------------------------------------
Write-Host ""
$S_Connection = Connect-PurviewComplianceSession

$S_TenantName = 'Unknown tenant'
$S_TenantId = 'Unknown'
$S_Account = 'Unknown'
if ($S_Connection)
{
    if ($S_Connection.Organization) { $S_TenantName = $S_Connection.Organization }
    if ($S_Connection.TenantID) { $S_TenantId = $S_Connection.TenantID }
    if ($S_Connection.UserPrincipalName) { $S_Account = $S_Connection.UserPrincipalName }
}

Write-Host ""
Write-Host "Connected to     : $S_TenantName ($S_TenantId)" -ForegroundColor Green
Write-Host "Account          : $S_Account" -ForegroundColor Green

# Authoritative SIT list from the tenant. The taxonomy is a classification layer over this, never a
# substitute for it: Microsoft adds SITs, and tenants hold custom ones the taxonomy has never seen.
Write-Host ""
Write-Host "Retrieving tenant SIT list (Get-DlpSensitiveInformationType)..." -ForegroundColor Cyan

$S_TenantSitById = @{}      # PowerShell hashtables are case-insensitive, which matters: a handful of
$S_TenantSitByName = @{}    # built-in SITs return uppercase GUIDs.
try
{
    foreach ($S_TSit in @(Get-DlpSensitiveInformationType -ErrorAction Stop))
    {
        $S_TId = "$($S_TSit.Id)"
        if ($S_TId) { $S_TenantSitById[$S_TId] = $S_TSit }
        if ($S_TSit.Name) { $S_TenantSitByName[[string]$S_TSit.Name] = $S_TSit }
    }
    Write-Host "  Tenant SITs    : $($S_TenantSitByName.Count)" -ForegroundColor Green
}
catch
{
    throw "Could not retrieve the tenant SIT list: $($_.Exception.Message). Confirm the account holds the '$($S_RequiredPurviewRoles -join "', '")' role in the Microsoft Purview compliance portal."
}

# Split the selection into what the tenant actually has, and what it does not.
$S_Runnable = [System.Collections.Generic.List[PSCustomObject]]::new()
$S_MissingFromTenant = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($S_Sit in $S_SelectedSits)
{
    $S_Tenant = $null
    if ($S_Sit.id -and $S_TenantSitById.ContainsKey("$($S_Sit.id)")) { $S_Tenant = $S_TenantSitById["$($S_Sit.id)"] }
    elseif ($S_TenantSitByName.ContainsKey("$($S_Sit.name)")) { $S_Tenant = $S_TenantSitByName["$($S_Sit.name)"] }

    if ($S_Tenant)
    {
        # Always query using the tenant's own name - Export-ContentExplorerData matches on TagName,
        # and Microsoft occasionally renames a SIT while keeping its GUID.
        $S_Runnable.Add([PSCustomObject]@{
            Name                  = [string]$S_Tenant.Name
            Id                    = "$($S_Tenant.Id)"
            TaxonomyName          = $S_Sit.name
            Tags                  = @($S_Sit.tags)
            RecommendedConfidence = $S_Sit.recommendedConfidence
            Note                  = $S_Sit.note
        })
    }
    else
    {
        Write-Verbose "Skipping '$($S_Sit.name)' - not present in this tenant."
        $S_MissingFromTenant.Add([PSCustomObject]@{ Name = $S_Sit.name; Id = $S_Sit.id })
    }
}

if ($S_MissingFromTenant.Count -gt 0)
{
    Write-Warning "$($S_MissingFromTenant.Count) selected SIT(s) are not present in this tenant and will be skipped. They are listed in the report's data quality section."
}

if ($S_Runnable.Count -eq 0)
{
    throw "None of the selected SITs exist in this tenant. Nothing to report on."
}

# Tenant SITs the taxonomy has never seen. Never dropped silently - they are surfaced in the report so
# a custom SIT is visible even though it cannot be selected by tag.
$S_TaxonomyIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($S_Sit in @($S_Taxonomy.Document.sits)) { [void]$S_TaxonomyIds.Add("$($S_Sit.id)") }

$S_UntaggedTenantSits = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($S_Key in $S_TenantSitByName.Keys)
{
    $S_TSit = $S_TenantSitByName[$S_Key]
    if (-not $S_TaxonomyIds.Contains("$($S_TSit.Id)"))
    {
        $S_UntaggedTenantSits.Add([PSCustomObject]@{
            Name      = [string]$S_TSit.Name
            Id        = "$($S_TSit.Id)"
            Publisher = "$($S_TSit.Publisher)"
        })
    }
}

# ---------------------------------------------------------------------------
# Output paths
# ---------------------------------------------------------------------------
if (-not $ReportPath) { $ReportPath = (Get-Location).Path }

$S_Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$S_FileBase = if ($TestMode) { "ReportPurviewSIT_TEST_$S_Timestamp" } else { "ReportPurviewSIT_$S_Timestamp" }

if (Test-Path -LiteralPath $ReportPath -PathType Container)
{
    $S_ReportFolder = $ReportPath
    $S_CsvFile = Join-Path $ReportPath "$S_FileBase.csv"
    $S_HtmlFile = Join-Path $ReportPath "$S_FileBase.html"
}
else
{
    $S_ReportFolder = Split-Path -Parent $ReportPath
    if (-not $S_ReportFolder) { $S_ReportFolder = (Get-Location).Path }
    $S_CsvFile = $ReportPath
    $S_HtmlFile = [System.IO.Path]::ChangeExtension($ReportPath, '.html')
    $S_FileBase = [System.IO.Path]::GetFileNameWithoutExtension($ReportPath)
}

if ($S_ReportFolder -and -not (Test-Path -LiteralPath $S_ReportFolder))
{
    New-Item -ItemType Directory -Path $S_ReportFolder -Force | Out-Null
}

# Fail before the run rather than after it, if the folder is not writable.
try
{
    $S_WriteProbe = Join-Path $S_ReportFolder ".writetest_$([guid]::NewGuid().ToString('N')).tmp"
    [System.IO.File]::WriteAllText($S_WriteProbe, 'x')
    Remove-Item -LiteralPath $S_WriteProbe -Force
}
catch
{
    throw "Output folder is not writable: $S_ReportFolder ($($_.Exception.Message))"
}

if (-not $CheckpointPath) { $CheckpointPath = Join-Path $S_ReportFolder "$($S_FileBase).checkpoint.json" }

# ---------------------------------------------------------------------------
# Assumption verification mode - probes instead of reports, then exits.
# ---------------------------------------------------------------------------
if ($VerifyAssumptions)
{
    Invoke-AssumptionVerification -Candidates $S_Runnable `
        -PreferredSit $VerifySensitiveInfoType `
        -MaxPages $VerifyMaxPages `
        -DelayMs $S_RequestDelayMilliseconds `
        -Workloads $S_Workloads `
        -FindingsPath (Join-Path $S_ReportFolder "$($S_FileBase).assumptions.json")
    return
}

# ---------------------------------------------------------------------------
# Resume
# ---------------------------------------------------------------------------
$S_Completed = @{}   # "sitname|workload" -> TotalCount
if ($ResumeFrom)
{
    if (-not (Test-Path -LiteralPath $ResumeFrom -PathType Leaf))
    {
        throw "Checkpoint file not found: $ResumeFrom"
    }
    try
    {
        $S_Checkpoint = Get-Content -LiteralPath $ResumeFrom -Raw | ConvertFrom-Json
        foreach ($S_Row in @($S_Checkpoint.Rows))
        {
            # Only reuse successful results. A row that errored last time is retried.
            if ($null -eq $S_Row.TotalCount -or $S_Row.Error) { continue }
            $S_Completed["$($S_Row.SitName)|$($S_Row.Workload)"] = [long]$S_Row.TotalCount
        }
        Write-Host ""
        Write-Host "Resuming from    : $ResumeFrom ($($S_Completed.Count) result(s) reused)" -ForegroundColor Yellow
    }
    catch
    {
        throw "Could not read checkpoint '$ResumeFrom': $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# Run the census
# ---------------------------------------------------------------------------
Write-Host ""
if ($Mode -eq 'Full')
{
    Write-Host "Running FULL item-level export over $($S_Runnable.Count) SIT(s) x $($S_Workloads.Count) workload(s)..." -ForegroundColor Cyan
    Write-Host "One API call per page rather than per SIT, so this is substantially longer than a census." -ForegroundColor DarkGray
}
else
{
    Write-Host "Running census over $($S_Runnable.Count) SIT(s) x $($S_Workloads.Count) workload(s)..." -ForegroundColor Cyan
}
Write-Host ""

$S_Results = [System.Collections.Generic.List[PSCustomObject]]::new()
$S_Failures = [System.Collections.Generic.List[PSCustomObject]]::new()

# Full-mode state. The row sink holds one entry per item per SIT; the seen-set deduplicates, because
# the same item is returned by every SIT query it matches. The lookup is keyed by SIT GUID and is
# case-insensitive by virtue of being a PowerShell hashtable - a handful of built-in SITs return
# uppercase GUIDs, and SensitiveInfoTypesData casing need not match the SIT list's.
$S_DetailRows = [System.Collections.Generic.List[PSCustomObject]]::new()
$S_SeenItemSit = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$S_QueriedSitLookup = @{}
foreach ($S_Sit in $S_Runnable) { $S_QueriedSitLookup[$S_Sit.Id] = $S_Sit }

$S_TotalCalls = $S_Runnable.Count * $S_Workloads.Count
$S_CallIndex = 0
$S_SinceCheckpoint = 0
$S_Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$S_RealCalls = 0

foreach ($S_Sit in $S_Runnable)
{
    foreach ($S_Wl in $S_Workloads)
    {
        $S_CallIndex++
        $S_Key = "$($S_Sit.Name)|$S_Wl"

        if ($S_Completed.ContainsKey($S_Key))
        {
            $S_Count = $S_Completed[$S_Key]
            $S_Err = $null
        }
        elseif ($Mode -eq 'Full')
        {
            # Full mode counts items for real: page the records, flatten them, and count the distinct
            # items this SIT was actually found on. That figure replaces TotalCount as the per-SIT
            # per-workload count, so every downstream roll-up is computed from observed items rather
            # than from an aggregate the tool never verified.
            $S_Before = $S_DetailRows.Count
            $S_Response = Invoke-SitItemCollection -Name $S_Sit.Name -WorkloadCode $S_Wl `
                -QueriedSits $S_QueriedSitLookup -RowSink $S_DetailRows -SeenKeys $S_SeenItemSit `
                -MaxRetries $S_MaxRetriesPerCall -DelayMs $S_RequestDelayMilliseconds -RevealNames ([bool]$IncludeItemNames)

            $S_Err = $S_Response.Error
            $S_RealCalls += $S_Response.Pages

            # Count distinct items carrying THIS SIT in THIS workload, across all rows gathered so far
            # - the item may have arrived from another SIT's query.
            $S_Count = @($S_DetailRows | Where-Object { $_.SitId -eq $S_Sit.Id -and $_.Workload -eq $S_Wl }).Count

            if ($S_Err)
            {
                $S_Failures.Add([PSCustomObject]@{ SitName = $S_Sit.Name; Workload = $S_Wl; Error = $S_Err })
                $S_Count = $null
            }
        }
        else
        {
            $S_Response = Get-SitWorkloadCount -Name $S_Sit.Name -WorkloadCode $S_Wl -MaxRetries $S_MaxRetriesPerCall
            $S_Count = $S_Response.TotalCount
            $S_Err = $S_Response.Error
            $S_RealCalls++

            if ($S_Err)
            {
                $S_Failures.Add([PSCustomObject]@{ SitName = $S_Sit.Name; Workload = $S_Wl; Error = $S_Err })
            }

            if ($S_RequestDelayMilliseconds -gt 0) { Start-Sleep -Milliseconds $S_RequestDelayMilliseconds }
        }

        $S_Results.Add([PSCustomObject]@{
            SitName               = $S_Sit.Name
            SitId                 = $S_Sit.Id
            Workload              = $S_Wl
            TotalCount            = $S_Count
            Error                 = $S_Err
            Tags                  = $S_Sit.Tags
            Class                 = (Get-TagValue -Tags $S_Sit.Tags -Dimension 'class')
            Pii                   = (Get-TagValue -Tags $S_Sit.Tags -Dimension 'pii')
            Tier                  = (Get-TagValue -Tags $S_Sit.Tags -Dimension 'tier')
            Jurisdictions         = (Get-TagValues -Tags $S_Sit.Tags -Dimension 'jurisdiction')
            Traits                = (Get-TagValues -Tags $S_Sit.Tags -Dimension 'trait')
            RecommendedConfidence = $S_Sit.RecommendedConfidence
            Note                  = $S_Sit.Note
        })

        # Progress. Once a few real calls are timed, the remaining estimate uses the measured rate.
        $S_Eta = ''
        if ($S_RealCalls -ge 3)
        {
            $S_PerCall = $S_Stopwatch.Elapsed.TotalSeconds / $S_RealCalls
            $S_Remaining = [timespan]::FromSeconds($S_PerCall * ($S_TotalCalls - $S_CallIndex))
            $S_Eta = "  ETA $($S_Remaining.ToString('hh\:mm\:ss'))"
        }

        $S_CountText = if ($null -eq $S_Count) { 'ERROR' } else { '{0:N0}' -f $S_Count }
        $S_Colour = if ($null -eq $S_Count) { 'Red' } elseif ($S_Count -gt 0) { 'Yellow' } else { 'DarkGray' }
        Write-Host ("[{0}/{1}] {2,-52} {3,-6} {4,10}{5}" -f $S_CallIndex, $S_TotalCalls, $S_Sit.Name, $S_Wl, $S_CountText, $S_Eta) -ForegroundColor $S_Colour

        $S_SinceCheckpoint++
        if ($S_SinceCheckpoint -ge $S_CheckpointEvery)
        {
            $S_CheckpointOk = Save-Checkpoint -Path $CheckpointPath -Rows $S_Results
            $S_SinceCheckpoint = 0
        }
    }
}

$S_Stopwatch.Stop()
$S_CheckpointOk = Save-Checkpoint -Path $CheckpointPath -Rows $S_Results

Write-Host ""
Write-Host "Census complete in $($S_Stopwatch.Elapsed.ToString('hh\:mm\:ss'))." -ForegroundColor Green

# ---------------------------------------------------------------------------
# Aggregate
#
# Every roll-up below is either per-SIT or per-workload. Summing item counts ACROSS SITs is not a
# distinct item count, because one document routinely carries dozens of SITs - so wherever such a sum
# appears it is named 'SumOfItemCounts' and captioned, never presented as "items".
#
# Summing ACROSS WORKLOADS is a separate question. ODB, SPO and EXO are distinct stores, so those are
# safe to add. Teams is not established: Teams chat is stored in Exchange mailboxes and both schemas
# put a mailbox address in FileSourceUrl, so a message may be reachable under EXO and Teams both. That
# has not been confirmed against a tenant (assumption A2 in the header), so when EXO and Teams are
# both in scope, ItemsTotal is reported as possibly overlapping rather than as a clean total.
# Run -VerifyAssumptions to settle it.
# ---------------------------------------------------------------------------
$S_ExoTeamsBothInScope = ($S_Workloads -contains 'EXO') -and ($S_Workloads -contains 'Teams')
$S_TotalBasis = if ($S_ExoTeamsBothInScope) { 'may-overlap-exo-teams' } else { 'exact' }

if ($S_ExoTeamsBothInScope)
{
    Write-Warning "EXO and Teams are both in scope. Teams chat is stored in Exchange mailboxes, so ItemsTotal may count the same message twice. This is unverified (assumption A2) - use the per-workload columns for figures you intend to quote, or run -VerifyAssumptions to settle it."
}

# Per SIT, across workloads. Valid to sum: the four workloads are disjoint.
$S_BySit = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($S_Group in ($S_Results | Group-Object SitName))
{
    $S_Rows = @($S_Group.Group)
    $S_First = $S_Rows[0]

    $S_Line = [ordered]@{
        SensitiveInfoType     = $S_Group.Name
        SitId                 = $S_First.SitId
        Class                 = $S_First.Class
        Pii                   = $S_First.Pii
        Tier                  = $S_First.Tier
        Jurisdictions         = (@($S_First.Jurisdictions) -join '; ')
        Traits                = (@($S_First.Traits) -join '; ')
        RecommendedConfidence = $S_First.RecommendedConfidence
    }
    foreach ($S_Wl in $S_Workloads)
    {
        $S_Row = $S_Rows | Where-Object { $_.Workload -eq $S_Wl } | Select-Object -First 1
        $S_Line["Items_$S_Wl"] = if ($S_Row -and $null -ne $S_Row.TotalCount) { $S_Row.TotalCount } else { $null }
    }
    $S_Line['ItemsTotal'] = (@($S_Rows | Where-Object { $null -ne $_.TotalCount } | ForEach-Object { $_.TotalCount }) | Measure-Object -Sum).Sum
    if ($null -eq $S_Line['ItemsTotal']) { $S_Line['ItemsTotal'] = 0 }
    # Says whether ItemsTotal is a clean sum of disjoint stores, or may double count Teams chat that
    # also surfaces under EXO. Consumed by the HTML, which labels the total column accordingly.
    $S_Line['ItemsTotalBasis'] = $S_TotalBasis
    $S_Line['Errors'] = @($S_Rows | Where-Object { $_.Error }).Count
    $S_Line['Note'] = $S_First.Note

    $S_BySit.Add([PSCustomObject]$S_Line)
}
[array]$S_BySit = @($S_BySit | Sort-Object -Property ItemsTotal -Descending)

# Per tag value, over the four single/multi-cardinality dimensions worth reporting.
$S_ByTag = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($S_Dim in @('class', 'pii', 'tier', 'jurisdiction'))
{
    $S_Buckets = @{}
    foreach ($S_Line in $S_BySit)
    {
        $S_Values = switch ($S_Dim)
        {
            'class'        { @($S_Line.Class) }
            'pii'          { @($S_Line.Pii) }
            'tier'         { @($S_Line.Tier) }
            'jurisdiction' { @($S_Line.Jurisdictions -split ';\s*' | Where-Object { $_ }) }
        }
        foreach ($S_Value in $S_Values)
        {
            if (-not $S_Value) { $S_Value = '(untagged)' }
            if (-not $S_Buckets.ContainsKey($S_Value)) { $S_Buckets[$S_Value] = [System.Collections.Generic.List[PSCustomObject]]::new() }
            $S_Buckets[$S_Value].Add($S_Line)
        }
    }

    foreach ($S_Value in $S_Buckets.Keys)
    {
        $S_Lines = @($S_Buckets[$S_Value])
        $S_Entry = [ordered]@{
            Dimension        = $S_Dim
            TagValue         = $S_Value
            SitsSelected     = $S_Lines.Count
            SitsWithFindings = @($S_Lines | Where-Object { $_.ItemsTotal -gt 0 }).Count
        }
        foreach ($S_Wl in $S_Workloads)
        {
            $S_Sum = (@($S_Lines | ForEach-Object { $_."Items_$S_Wl" } | Where-Object { $null -ne $_ }) | Measure-Object -Sum).Sum
            $S_Entry["SumOfItemCounts_$S_Wl"] = if ($null -eq $S_Sum) { 0 } else { $S_Sum }
        }
        $S_TotalSum = (@($S_Lines | ForEach-Object { $_.ItemsTotal }) | Measure-Object -Sum).Sum
        $S_Entry['SumOfItemCounts'] = if ($null -eq $S_TotalSum) { 0 } else { $S_TotalSum }
        $S_ByTag.Add([PSCustomObject]$S_Entry)
    }
}
[array]$S_ByTag = @($S_ByTag | Sort-Object Dimension, @{ Expression = 'SumOfItemCounts'; Descending = $true })

# Per workload.
$S_ByWorkload = foreach ($S_Wl in $S_Workloads)
{
    $S_Rows = @($S_Results | Where-Object { $_.Workload -eq $S_Wl })
    $S_Sum = (@($S_Rows | Where-Object { $null -ne $_.TotalCount } | ForEach-Object { $_.TotalCount }) | Measure-Object -Sum).Sum
    [PSCustomObject]@{
        Workload         = $S_Wl
        SitsWithFindings = @($S_Rows | Where-Object { $_.TotalCount -gt 0 }).Count
        SumOfItemCounts  = if ($null -eq $S_Sum) { 0 } else { $S_Sum }
        Errors           = @($S_Rows | Where-Object { $_.Error }).Count
    }
}

# Full-mode metrics, computed from observed items rather than from an aggregate count.
$S_DistinctItems = 0
$S_ContainersAffected = 0
$S_HighMatchTotal = 0
$S_MediumBandTotal = 0
$S_LowBandTotal = 0
$S_TruncatedRows = 0
$S_AdjustedRows = 0
$S_ItemsWithHigh = 0

if ($Mode -eq 'Full' -and $S_DetailRows.Count -gt 0)
{
    $S_DistinctItems = @($S_DetailRows | ForEach-Object { $_.ItemKey } | Select-Object -Unique).Count
    $S_ContainersAffected = @($S_DetailRows | ForEach-Object { $_.Container } | Where-Object { $_ } | Select-Object -Unique).Count
    $S_HighMatchTotal = (@($S_DetailRows | ForEach-Object { $_.HighMatches }) | Measure-Object -Sum).Sum
    $S_MediumBandTotal = (@($S_DetailRows | ForEach-Object { $_.MediumBandMatches }) | Measure-Object -Sum).Sum
    $S_LowBandTotal = (@($S_DetailRows | ForEach-Object { $_.LowBandMatches }) | Measure-Object -Sum).Sum
    $S_TruncatedRows = @($S_DetailRows | Where-Object { $_.MatchesTruncated }).Count
    $S_AdjustedRows = @($S_DetailRows | Where-Object { $_.BandsAdjusted }).Count
    $S_ItemsWithHigh = @($S_DetailRows | Where-Object { $_.HighMatches -gt 0 } | ForEach-Object { $_.ItemKey } | Select-Object -Unique).Count

    foreach ($S_V in @('S_HighMatchTotal', 'S_MediumBandTotal', 'S_LowBandTotal'))
    {
        if ($null -eq (Get-Variable -Name $S_V -ValueOnly)) { Set-Variable -Name $S_V -Value 0 }
    }

    if ($S_AdjustedRows -gt 0)
    {
        Write-Warning "$S_AdjustedRows row(s) had a confidence ordering the API is not expected to produce (Low >= Medium >= High did not hold). Their bands were clamped at zero and their total recomputed from the bands, so the three bands still sum to the total. They are flagged BandsAdjusted in the CSV."
    }
}

# Detect a LICENSING gap, as opposed to genuinely clean data.
#
# An unlicensed SIT returns zero rather than an error, so it is indistinguishable from a real nil
# result - the single most damaging way this report could mislead. Confirmed on the reference tenant:
# a named entity SIT matching almost any two capitalised tokens returned thousands of items in Teams
# and ZERO in Exchange, while ordinary pattern types returned plenty in Exchange. Names do not vanish
# from email. Named entity SITs need the Privacy Service Plan for Exchange at-rest classification
# specifically, and that plan was absent.
#
# The rules are read from the taxonomy's licensing block rather than hardcoded, so adding a rule is a
# JSON edit and not a code change. Each rule names the trait it applies to; a workload where every SIT
# carrying that trait returned zero, while other SITs found data, is flagged against that rule.
$S_CoverageGaps = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($S_Rule in @($S_Taxonomy.Document.licensing.rules))
{
    $S_RuleTrait = "$($S_Rule.applies)"
    if (-not $S_RuleTrait) { continue }
    $S_RuleTrait = $S_RuleTrait -replace '^trait:', ''

    foreach ($S_Wl in $S_Workloads)
    {
        $S_WlRows = @($S_Results | Where-Object { $_.Workload -eq $S_Wl -and -not $_.Error })
        if ($S_WlRows.Count -eq 0) { continue }

        $S_Affected = @($S_WlRows | Where-Object { @($_.Traits) -contains $S_RuleTrait })
        $S_Others = @($S_WlRows | Where-Object { -not (@($_.Traits) -contains $S_RuleTrait) })
        if ($S_Affected.Count -eq 0) { continue }

        $S_AffectedHits = @($S_Affected | Where-Object { $_.TotalCount -gt 0 }).Count
        $S_OtherHits = @($S_Others | Where-Object { $_.TotalCount -gt 0 }).Count

        if ($S_AffectedHits -eq 0 -and $S_OtherHits -gt 0)
        {
            # Do these same types return data in another workload? If they do, the licence is
            # evidently held tenant-wide and a blanket rule cannot explain a single-workload zero -
            # which points at the workload-specific rule instead. This is what separates "no E5" from
            # "E5 but no Priva for Exchange at rest", and reporting both without it misleads.
            $S_Elsewhere = @($S_Results | Where-Object {
                $_.Workload -ne $S_Wl -and -not $_.Error -and $_.TotalCount -gt 0 -and (@($_.Traits) -contains $S_RuleTrait)
            } | ForEach-Object { $_.Workload } | Select-Object -Unique)

            $S_CoverageGaps.Add([PSCustomObject]@{
                Workload       = $S_Wl
                Trait          = $S_RuleTrait
                AffectedSits   = $S_Affected.Count
                OtherHits      = $S_OtherHits
                Requirement    = "$($S_Rule.requirement)"
                Affects        = "$($S_Rule.affects)"
                Source         = "$($S_Rule.source)"
                WorksElsewhere = $S_Elsewhere
            })
        }
    }
}

foreach ($S_Gap in $S_CoverageGaps)
{
    $S_GapMsg = "$($S_Gap.Workload): all $($S_Gap.AffectedSits) SIT(s) tagged '$($S_Gap.Trait)' returned zero, while $($S_Gap.OtherHits) other type(s) found data. Licence required: $($S_Gap.Requirement)."
    if ($S_Gap.WorksElsewhere.Count -gt 0)
    {
        $S_GapMsg += " NOTE: these same types DO return data in $($S_Gap.WorksElsewhere -join ', '), so that licence is evidently held and this gap is specific to $($S_Gap.Workload) - look to the workload-specific requirement rather than this one."
    }
    $S_GapMsg += " Either way the zeros are an artefact, not evidence of clean data - do not report those types as clean in this workload."
    Write-Warning $S_GapMsg
}

$S_SitsWithFindings = @($S_BySit | Where-Object { $_.ItemsTotal -gt 0 })
$S_SitsClean = @($S_BySit | Where-Object { $_.ItemsTotal -eq 0 -and $_.Errors -eq 0 })
$S_DirectPiiSits = @($S_SitsWithFindings | Where-Object { $_.Pii -eq 'direct' })
$S_GrandSum = (@($S_BySit | ForEach-Object { $_.ItemsTotal }) | Measure-Object -Sum).Sum
if ($null -eq $S_GrandSum) { $S_GrandSum = 0 }

# ---------------------------------------------------------------------------
# CSV output
# ---------------------------------------------------------------------------
# One CSV, carrying the grain the report was computed from and nothing else. Roll-ups live in the
# HTML; exporting them here as well would duplicate that and create a second source of truth.
#
# Long format on purpose. It pivots directly in Excel - rows as Container, columns as
# SensitiveInfoType, values as sum of HighMatches - so any cross-tab can be produced from the file
# without re-running the tool. A wide layout with one column per SIT is one pivot away from this, and
# would change its header with every -View, so no two runs would be comparable.
if ($Format -in @('Csv', 'Both'))
{
    if ($Mode -eq 'Full')
    {
        $S_DetailRows |
            Sort-Object Workload, Container, ItemId, SensitiveInfoType |
            Select-Object Workload, Container, ItemId, ItemName, ItemKey,
                SensitiveInfoType, SitId, Class, Pii, Tier, Jurisdictions, Traits, RecommendedConfidence,
                HighMatches, MediumBandMatches, LowBandMatches, TotalMatches,
                MatchesTruncated, BandsAdjusted,
                SensitivityLabel, RetentionLabel, LastModifiedTime |
            Export-Csv -Path $S_CsvFile -NoTypeInformation -Encoding UTF8
    }
    else
    {
        # Census cannot produce item detail - it never retrieves item records - so its underlying
        # grain is the SIT x workload count itself.
        $S_Results |
            Select-Object SitName, SitId, Workload, TotalCount, Class, Pii, Tier,
                @{ n = 'Jurisdictions'; e = { @($_.Jurisdictions) -join '; ' } },
                @{ n = 'Traits'; e = { @($_.Traits) -join '; ' } },
                RecommendedConfidence, Error |
            Sort-Object SitName, Workload |
            Export-Csv -Path $S_CsvFile -NoTypeInformation -Encoding UTF8
    }
}

# ---------------------------------------------------------------------------
# HTML report
# ---------------------------------------------------------------------------
if ($Format -in @('Html', 'Both'))
{

$S_ReportDate = Get-Date -Format 'dd MMM yyyy HH:mm'
$S_WorkloadColours = @('#0b5394', '#12694a', '#9a3412', '#5b21b6')

# Chart: top SITs by item count. Per-SIT totals are exact.
$S_TopSitRows = @($S_SitsWithFindings | Select-Object -First 15 | ForEach-Object {
    [PSCustomObject]@{ Label = $_.SensitiveInfoType; Value = [double]$_.ItemsTotal }
})
$S_TopSitChart = New-SvgBarChart -Rows $S_TopSitRows -Colour '#0b5394'

# Chart: class roll-up, split by workload.
$S_ClassRows = @($S_ByTag | Where-Object { $_.Dimension -eq 'class' -and $_.SumOfItemCounts -gt 0 } |
    Sort-Object -Property SumOfItemCounts -Descending | ForEach-Object {
        $S_Entry = $_
        [PSCustomObject]@{
            Label  = $S_Entry.TagValue
            Values = @($S_Workloads | ForEach-Object { [double]$S_Entry."SumOfItemCounts_$_" })
        }
    })
$S_ClassChart = New-SvgStackedBarChart -Rows $S_ClassRows -SeriesNames $S_Workloads `
    -SeriesColours @($S_WorkloadColours[0..($S_Workloads.Count - 1)])

# Per-SIT table markup. Named for the table it builds, not "detail" - $S_DetailRows is the item-level
# data set and the two must not collide.
$S_SitTableRows = (@($S_BySit) | ForEach-Object {
    $S_Line = $_
    $S_Cells = foreach ($S_Wl in $S_Workloads)
    {
        $S_Val = $S_Line."Items_$S_Wl"
        if ($null -eq $S_Val) { '<td class="num muted">err</td>' }
        elseif ($S_Val -eq 0) { '<td class="num muted">0</td>' }
        else { "<td class=`"num strong`">$('{0:N0}' -f $S_Val)</td>" }
    }

    $S_PiiBadge = switch ($S_Line.Pii)
    {
        'direct'   { '<span class="badge badge-direct">direct</span>' }
        'indirect' { '<span class="badge badge-indirect">indirect</span>' }
        'none'     { '<span class="badge badge-nonpii">not personal</span>' }
        default    { '<span class="badge badge-untagged">untagged</span>' }
    }

    $S_TraitText = if ($S_Line.Traits) { "<div class=`"traits`">$(ConvertTo-HtmlText $S_Line.Traits)</div>" } else { '' }
    $S_Found = if ($S_Line.ItemsTotal -gt 0) { 'yes' } else { 'no' }

    "<tr data-class=`"$(ConvertTo-HtmlText $S_Line.Class)`" data-pii=`"$(ConvertTo-HtmlText $S_Line.Pii)`" data-tier=`"$(ConvertTo-HtmlText $S_Line.Tier)`" data-found=`"$S_Found`">" +
    "<td><strong>$(ConvertTo-HtmlText $S_Line.SensitiveInfoType)</strong>$S_TraitText</td>" +
    "<td>$(ConvertTo-HtmlText $S_Line.Class)</td>" +
    "<td>$S_PiiBadge</td>" +
    "<td>$(ConvertTo-HtmlText $S_Line.Tier)</td>" +
    "<td>$(ConvertTo-HtmlText $S_Line.Jurisdictions)</td>" +
    ($S_Cells -join '') +
    "<td class=`"num strong`">$('{0:N0}' -f $S_Line.ItemsTotal)</td>" +
    "</tr>"
}) -join "`n"
if (-not $S_SitTableRows) { $S_SitTableRows = "<tr><td colspan=`"$(6 + $S_Workloads.Count)`" class=`"muted`">No SITs were queried.</td></tr>" }

$S_WorkloadHeaders = (@($S_Workloads | ForEach-Object { "<th class=`"num`">$(ConvertTo-HtmlText $_)</th>" }) -join '')

# By-workload table.
$S_WorkloadRows = (@($S_ByWorkload) | ForEach-Object {
    "<tr><td><strong>$(ConvertTo-HtmlText $_.Workload)</strong></td><td class=`"num`">$($_.SitsWithFindings)</td><td class=`"num`">$('{0:N0}' -f $_.SumOfItemCounts)</td><td class=`"num`">$($_.Errors)</td></tr>"
}) -join "`n"

# Data quality tables.
$S_MissingRows = (@($S_MissingFromTenant) | ForEach-Object {
    "<tr><td>$(ConvertTo-HtmlText $_.Name)</td><td class=`"mono`">$(ConvertTo-HtmlText $_.Id)</td></tr>"
}) -join "`n"
if (-not $S_MissingRows) { $S_MissingRows = '<tr><td colspan="2" class="muted">Every selected SIT exists in this tenant.</td></tr>' }

$S_UntaggedRows = (@($S_UntaggedTenantSits | Sort-Object Name) | ForEach-Object {
    "<tr><td>$(ConvertTo-HtmlText $_.Name)</td><td class=`"mono`">$(ConvertTo-HtmlText $_.Id)</td><td>$(ConvertTo-HtmlText $_.Publisher)</td></tr>"
}) -join "`n"
if (-not $S_UntaggedRows) { $S_UntaggedRows = '<tr><td colspan="3" class="muted">Every SIT in this tenant is covered by the taxonomy.</td></tr>' }

$S_FailureRows = (@($S_Failures) | ForEach-Object {
    "<tr><td>$(ConvertTo-HtmlText $_.SitName)</td><td>$(ConvertTo-HtmlText $_.Workload)</td><td>$(ConvertTo-HtmlText $_.Error)</td></tr>"
}) -join "`n"
if (-not $S_FailureRows) { $S_FailureRows = '<tr><td colspan="3" class="muted">No calls failed.</td></tr>' }

# Caveat bullets that only apply to some runs.
$S_OverlapCaveat = if ($S_ExoTeamsBothInScope)
{
    '<li><strong>The EXO and Teams columns may overlap.</strong> Teams chat is stored in Exchange mailboxes, and both workloads return a mailbox address as the container, so the same message may be counted under each. Whether they actually overlap has not been confirmed against a tenant. Until it is, treat the per-workload columns as the quotable figures and the Total column as an upper bound. The CSV records this as <code>ItemsTotalBasis = may-overlap-exo-teams</code>.</li>'
}
else { '' }

# A1 held in Exchange but not in Teams (see the assumption register). Either way TotalCount counts
# records rather than distinct items, which only affects Census mode - it reports TotalCount directly.
# Full mode counts observed distinct items and needs no such caveat.
$S_TotalCountCaveat = if ($Mode -eq 'Full')
{
    '<li><strong>Item counts here are distinct items</strong>, each counted by observing the item rather than trusting an aggregate. A census of the same types reports slightly higher figures, because the census total counts folder-level occurrences &mdash; an item filed in two folders counts twice there and once here.</li>'
}
else
{
    '<li><strong>These are occurrence counts, not distinct items.</strong> The total counts records rather than items, and the export walks folders while the container field only identifies the mailbox or site, so an item filed in two folders counts twice. Treat these as an <strong>upper bound</strong>. Run <code>-Mode Full</code> for exact item counts.</li>'
}

# Teams figures are under an open question on the reference tenant: its census total and its paged
# export disagreed by 41%, in a direction that means one of the two modes is wrong about Teams.
$S_TeamsCaveat = if ($S_Workloads -contains 'Teams')
{
    '<li><strong>Treat Teams figures as provisional.</strong> On the reference tenant, Teams was the one workload where the census total and an exhaustive item-level export disagreed sharply &mdash; the total claimed roughly twice what paging returned, with paging ending naturally rather than being cut short. Exchange, by contrast, matched exactly. Until that is explained, Teams numbers from either mode should be treated as indicative and not quoted as findings. Other workloads are unaffected.</li>'
}
else { '' }

$S_TotalHeader = if ($S_ExoTeamsBothInScope) { 'Total (upper bound)' } else { 'Total' }

# How many of the SITs actually queried carry each licensing-gated trait, so the caveat is specific
# to this run rather than a generic warning the reader has to translate.
$S_LicensingSummary = ''
$S_LicParts = @()
foreach ($S_Rule in @($S_Taxonomy.Document.licensing.rules))
{
    $S_RTrait = ("$($S_Rule.applies)" -replace '^trait:', '')
    if (-not $S_RTrait) { continue }
    $S_RCount = @($S_BySit | Where-Object { ($_.Traits -split ';\s*') -contains $S_RTrait }).Count
    if ($S_RCount -gt 0)
    {
        $S_LicParts += "$S_RCount of the $($S_BySit.Count) types in this run need $([System.Net.WebUtility]::HtmlEncode("$($S_Rule.requirement)"))"
    }
}
if ($S_LicParts.Count -gt 0) { $S_LicensingSummary = "In this run: " + ($S_LicParts -join '; ') + '.' }

# Confidence breakdown. Full mode only - a census never retrieves the per-item match data these bands
# are derived from, so there is nothing honest to show in Census mode.
$S_ConfidenceSection = ''
if ($Mode -eq 'Full' -and $S_DetailRows.Count -gt 0)
{
    $S_BandRows = @()
    foreach ($S_G in ($S_DetailRows | Group-Object SensitiveInfoType))
    {
        $S_H = (@($S_G.Group | ForEach-Object { $_.HighMatches }) | Measure-Object -Sum).Sum
        $S_M = (@($S_G.Group | ForEach-Object { $_.MediumBandMatches }) | Measure-Object -Sum).Sum
        $S_L = (@($S_G.Group | ForEach-Object { $_.LowBandMatches }) | Measure-Object -Sum).Sum
        if ($null -eq $S_H) { $S_H = 0 }; if ($null -eq $S_M) { $S_M = 0 }; if ($null -eq $S_L) { $S_L = 0 }
        if (($S_H + $S_M + $S_L) -le 0) { continue }
        $S_BandRows += [PSCustomObject]@{ Label = $S_G.Name; Values = @([double]$S_H, [double]$S_M, [double]$S_L); Sort = ($S_H + $S_M + $S_L) }
    }
    $S_BandRows = @($S_BandRows | Sort-Object -Property Sort -Descending | Select-Object -First 20)

    $S_BandChart = New-SvgStackedBarChart -Rows $S_BandRows `
        -SeriesNames @('High', 'Medium band', 'Low band') `
        -SeriesColours @('#0b5394', '#4a90d9', '#bcd8f0')

    $S_TruncNote = if ($S_TruncatedRows -gt 0)
    {
        "<p class=`"caption`"><strong>$S_TruncatedRows row(s)</strong> sit exactly at the apparent per-item per-type match cap. Their true match counts may be higher than shown, so these bars are a floor rather than an exact figure. Those rows carry <code>MatchesTruncated = True</code> in the CSV.</p>"
    }
    else { '' }

    $S_ConfidenceSection = @"
<div class="chart-section">
  <h2>Confidence breakdown by type</h2>
  <div class="caption">Matches split into three <strong>mutually exclusive</strong> bands. The API returns cumulative thresholds &mdash; each count includes everything above it &mdash; so these are derived: High is reported as-is, Medium band is Medium minus High, Low band is Low minus Medium. Adding the raw API figures instead would roughly triple the true total. Counts are matches, not items: one document can hold many.</div>
  $S_BandChart
  $S_TruncNote
</div>
"@
}

# The detail CSV is the sensitive artefact, so say plainly what is in it.
$S_CsvNote = if ($Mode -eq 'Full')
{
    $S_NameState = if ($IncludeItemNames)
    {
        'Item names were written <strong>in full</strong>, which for Exchange means real email subject lines. Treat that file as containing personal data in plain text.'
    }
    else
    {
        'Item names were redacted to a stable hash, so identical names still group together but none are disclosed. Re-run with <code>-IncludeItemNames</code> if you need the real values.'
    }
    "<div class=`"notice`"><strong>Companion detail file.</strong> A CSV accompanies this report at one row per item per sensitive information type, carrying the confidence bands, labels and locations behind every figure here. $S_NameState</div>"
}
else { '' }

$S_TestBanner = if ($TestMode)
{
    "<div class=`"notice`"><strong>Test mode:</strong> this report covers a random sample of $($S_Runnable.Count) SIT(s), not the full selection. Re-run without <code>-TestMode</code> for complete figures.</div>"
}
else { '' }

$S_FailureBanner = if ($S_Failures.Count -gt 0)
{
    "<div class=`"notice notice-warn`"><strong>$($S_Failures.Count) call(s) failed.</strong> The affected SIT and workload combinations have no figure in this report and are listed under Data quality. Treat their counts as unknown, not as zero.</div>"
}
else { '' }

# Reproducibility appendix. Rebuild the invocation from the bound parameters.
$S_ParamLines = foreach ($S_P in $PSBoundParameters.GetEnumerator() | Sort-Object Key)
{
    $S_Val = if ($S_P.Value -is [switch]) { '(switch)' } else { (@($S_P.Value) -join ', ') }
    "<tr><td class=`"mono`">-$(ConvertTo-HtmlText $S_P.Key)</td><td class=`"mono`">$(ConvertTo-HtmlText $S_Val)</td></tr>"
}
$S_ParamRows = ($S_ParamLines -join "`n")
if (-not $S_ParamRows) { $S_ParamRows = '<tr><td colspan="2" class="muted">Run with no parameters (all defaults).</td></tr>' }

$S_Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Purview Sensitive Information Type Census</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; padding: 30px; }

  .header { background: linear-gradient(135deg, #1a1a2e, #16213e); color: #fff; padding: 30px 40px; border-radius: 12px; margin-bottom: 30px; }
  .header h1 { font-size: 1.6em; margin-bottom: 6px; }
  .header p { font-size: 0.9em; opacity: 0.85; line-height: 1.6; }

  .section-title { font-size: 1.15em; font-weight: 600; color: #1a1a2e; margin-bottom: 16px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0; }

  .notice { background:#fff3cd;border:1px solid #ffc107;border-radius:8px;padding:12px 20px;margin-bottom:20px;color:#856404;font-size:0.9em; }
  .notice-warn { background:#f8d7da;border-color:#f5c2c7;color:#721c24; }

  .caveats { background:#fff; border-left:4px solid #0b5394; border-radius:8px; padding:20px 26px; margin-bottom:30px; box-shadow:0 2px 8px rgba(0,0,0,0.06); }
  .caveats h2 { font-size:1.05em; color:#1a1a2e; margin-bottom:12px; }
  .caveats ul { margin-left:18px; font-size:0.88em; line-height:1.75; color:#444; }
  .caveats strong { color:#1a1a2e; }

  .summary-cards { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
  .card { background: #fff; border-radius: 10px; padding: 24px 30px; flex: 1; min-width: 150px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
  .card .label { font-size: 0.85em; color: #777; text-transform: uppercase; letter-spacing: 0.5px; }
  .card .value { font-size: 2em; font-weight: 700; margin-top: 6px; }
  .card .sub { font-size: 0.78em; color: #999; margin-top: 2px; }

  .chart-section { background: #fff; border-radius: 10px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 30px; overflow-x: auto; }
  .chart-section h2 { font-size: 1.1em; margin-bottom: 6px; color: #1a1a2e; }
  .chart-section .caption { font-size: 0.82em; color: #777; margin-bottom: 18px; }
  svg.chart { display:block; min-width: 640px; }
  .svg-label { font: 12px 'Segoe UI', sans-serif; fill: #444; }
  .svg-value { font: 11px 'Segoe UI', sans-serif; fill: #666; }
  .svg-legend { font: 11px 'Segoe UI', sans-serif; fill: #555; }

  .table-section { background: #fff; border-radius: 10px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom: 30px; overflow-x: auto; }
  .table-section h2 { font-size: 1.1em; margin-bottom: 6px; color: #1a1a2e; }
  .table-section .caption { font-size: 0.82em; color: #777; margin-bottom: 16px; }
  .table-controls { display: flex; gap: 14px; margin-bottom: 16px; flex-wrap: wrap; align-items: center; }
  .table-controls input[type="text"] { padding: 8px 14px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; min-width: 240px; }
  .table-controls select { padding: 8px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 0.88em; background: #fff; }
  .table-controls .count-label { font-size: 0.85em; color: #777; margin-left: auto; }

  table { width: 100%; border-collapse: collapse; font-size: 0.84em; }
  th { background: #1a1a2e; color: #fff; padding: 12px 14px; text-align: left; cursor: pointer; user-select: none; white-space: nowrap; }
  th.num { text-align: right; }
  th:hover { background: #2c3e50; }
  td { padding: 10px 14px; border-bottom: 1px solid #eee; vertical-align: top; }
  td.num { text-align: right; font-variant-numeric: tabular-nums; }
  td.strong { font-weight: 600; color: #1a1a2e; }
  tr:hover td { background: #f8f9fa; }
  tr.hidden-row { display: none; }
  .mono { font-family: 'Consolas', monospace; font-size: 0.85em; color: #666; }
  .muted { color: #aaa; }
  .traits { font-size: 0.88em; color: #999; margin-top: 3px; }

  .badge { padding: 3px 10px; border-radius: 12px; font-size: 0.8em; font-weight: 600; display: inline-block; }
  .badge-direct   { background: #f8d7da; color: #721c24; }
  .badge-indirect { background: #fff3cd; color: #856404; }
  .badge-nonpii   { background: #e2e3e5; color: #495057; }
  .badge-untagged { background: #e7d9ff; color: #5b21b6; }

  .footer { text-align: center; font-size: 0.8em; color: #999; margin-top: 20px; }

  @media print {
    body { background: #fff; padding: 0; }
    .table-section, .chart-section, .card, .caveats { box-shadow: none; border: 1px solid #ddd; page-break-inside: avoid; }
    .table-controls { display: none; }
  }
</style>
</head>
<body>

<div class="header">
  <h1>Purview Sensitive Information Type Census</h1>
  <p>
    Tenant: $(ConvertTo-HtmlText $S_TenantName) ($(ConvertTo-HtmlText $S_TenantId)) &nbsp;|&nbsp; Generated: $S_ReportDate<br>
    Mode: $(if ($Mode -eq 'Full') { 'Full item-level export' } else { 'Census (item counts only)' }) &nbsp;|&nbsp; Workloads: $(ConvertTo-HtmlText ($S_Workloads -join ', '))<br>
    Selection: $(ConvertTo-HtmlText $S_Selection.Basis)
  </p>
</div>

$S_TestBanner
$S_FailureBanner
$S_CsvNote

<div class="caveats">
  <h2>Scope and caveats &mdash; read before quoting any figure</h2>
  <ul>
    <li><strong>Counts are items, not matches.</strong> One document holding forty credit card numbers counts once.</li>
    <li><strong>Do not add the per-SIT numbers together.</strong> A single document routinely carries dozens of SITs, so a sum across SITs counts the same item repeatedly. Any total across SITs in this report is labelled &ldquo;sum of overlapping counts&rdquo; for that reason.</li>
    $S_TotalCountCaveat
    $S_TeamsCaveat
    $S_OverlapCaveat
    <li><strong>No confidence filter was applied.</strong> Content Explorer's total is not scoped by confidence level, so filtering would have produced a misleading number. Every figure here means &ldquo;items carrying this type at any confidence&rdquo;. High-confidence figures need a full item-level export.</li>
    <li><strong>A nil result is not a guarantee of absence.</strong> It means nothing was detected within Purview's classification capability and the scope of this run. Encrypted and otherwise uninspectable items are not classified.</li>
    <li><strong>Licensing changes what is detectable.</strong> $(ConvertTo-HtmlText "$($S_Taxonomy.Document.licensing.note)")$(if ($S_LicensingSummary) { " $S_LicensingSummary" })</li>
    <li><strong>Teams covers chat and channel messages only.</strong> Files shared in Teams are stored in SharePoint and appear under SPO. Teams records also carry no subject, author, team or channel, so Teams is reportable as volume and exposure only &mdash; identifying a specific message needs eDiscovery.</li>
    <li><strong>Scope of this run:</strong> $($S_Runnable.Count) sensitive information type(s) across $(ConvertTo-HtmlText ($S_Workloads -join ', ')). Types outside the selection were not queried and their absence here says nothing about them.</li>
  </ul>
</div>

<div class="section-title">Executive summary</div>
<div class="summary-cards">
  <div class="card"><div class="label">Types queried</div><div class="value" style="color:#1a1a2e;">$($S_Runnable.Count)</div><div class="sub">of $($S_Taxonomy.SitCount) in taxonomy</div></div>
  <div class="card"><div class="label">Types found</div><div class="value" style="color:#e74c3c;">$($S_SitsWithFindings.Count)</div><div class="sub">at least one item</div></div>
  <div class="card"><div class="label">Direct identifiers found</div><div class="value" style="color:#c0392b;">$($S_DirectPiiSits.Count)</div><div class="sub">identify a person alone</div></div>
  <div class="card"><div class="label">Types not detected</div><div class="value" style="color:#27ae60;">$($S_SitsClean.Count)</div><div class="sub">zero items</div></div>
  <div class="card"><div class="label">Sum of item counts</div><div class="value" style="color:#0b5394;">$('{0:N0}' -f $S_GrandSum)</div><div class="sub">overlapping &mdash; not distinct items</div></div>
</div>

<div class="chart-section">
  <h2>Where the exposure is &mdash; top types by item count</h2>
  <div class="caption">Items containing each type, summed across the workloads in scope. Each bar is an exact figure; the bars must not be added to each other.</div>
  $S_TopSitChart
</div>

<div class="chart-section">
  <h2>Data class by workload</h2>
  <div class="caption">Sum of overlapping item counts per data class, segmented by workload. Use this for relative shape, not as a distinct item count.</div>
  $S_ClassChart
</div>

$S_ConfidenceSection

<div class="table-section">
  <h2>By workload</h2>
  <div class="caption">How many of the queried types were found in each workload.</div>
  <table id="wlTable">
    <thead>
      <tr>
        <th onclick="sortTable('wlTable', 0)">Workload &#x25B2;&#x25BC;</th>
        <th class="num" onclick="sortTable('wlTable', 1)">Types found &#x25B2;&#x25BC;</th>
        <th class="num" onclick="sortTable('wlTable', 2)">Sum of item counts &#x25B2;&#x25BC;</th>
        <th class="num" onclick="sortTable('wlTable', 3)">Errors &#x25B2;&#x25BC;</th>
      </tr>
    </thead>
    <tbody>
$S_WorkloadRows
    </tbody>
  </table>
</div>

<div class="table-section">
  <h2>Every type queried</h2>
  <div class="caption">One row per sensitive information type. Per-type and per-workload figures are exact item counts.</div>
  <div class="table-controls">
    <input type="text" id="sitSearch" placeholder="Search type / jurisdiction / trait..." onkeyup="filterSits()" />
    <select id="foundFilter" onchange="filterSits()">
      <option value="all">Found and not found</option>
      <option value="yes">Found only</option>
      <option value="no">Not detected only</option>
    </select>
    <select id="piiFilter" onchange="filterSits()">
      <option value="all">All PII classes</option>
      <option value="direct">Direct identifiers</option>
      <option value="indirect">Indirect identifiers</option>
      <option value="none">Not personal data</option>
    </select>
    <select id="tierFilter" onchange="filterSits()">
      <option value="all">All tiers</option>
      <option value="core">Core</option>
      <option value="extended">Extended</option>
      <option value="corroborating">Corroborating</option>
      <option value="excluded">Excluded</option>
    </select>
    <span class="count-label" id="sitCount">Showing $($S_BySit.Count) of $($S_BySit.Count) types</span>
  </div>
  <table id="sitTable">
    <thead>
      <tr>
        <th onclick="sortTable('sitTable', 0)">Sensitive information type &#x25B2;&#x25BC;</th>
        <th onclick="sortTable('sitTable', 1)">Class &#x25B2;&#x25BC;</th>
        <th onclick="sortTable('sitTable', 2)">PII &#x25B2;&#x25BC;</th>
        <th onclick="sortTable('sitTable', 3)">Tier &#x25B2;&#x25BC;</th>
        <th onclick="sortTable('sitTable', 4)">Jurisdiction &#x25B2;&#x25BC;</th>
$S_WorkloadHeaders
        <th class="num" onclick="sortTable('sitTable', $(5 + $S_Workloads.Count))">$S_TotalHeader &#x25B2;&#x25BC;</th>
      </tr>
    </thead>
    <tbody id="sitBody">
$S_SitTableRows
    </tbody>
  </table>
</div>

<div class="table-section">
  <h2>Data quality</h2>
  <div class="caption">What this run could not see, and what it saw but could not classify. Read this before treating the report as complete.</div>

  $(if ($S_CoverageGaps.Count -gt 0)
  {
    $S_GapRows = (@($S_CoverageGaps) | ForEach-Object {
        $S_Src = if ($_.Source) { "<br><span class=`"muted`" style=`"font-size:0.9em;`">$(ConvertTo-HtmlText $_.Source)</span>" } else { '' }
        $S_Elsew = if ($_.WorksElsewhere.Count -gt 0)
        {
            "<br><span style=`"color:#856404;`">These types <strong>do</strong> return data in $(ConvertTo-HtmlText ($_.WorksElsewhere -join ', ')), so this licence is evidently held and the gap is specific to $(ConvertTo-HtmlText $_.Workload).</span>"
        }
        else { '' }
        "<tr><td><strong>$(ConvertTo-HtmlText $_.Workload)</strong></td><td><code>$(ConvertTo-HtmlText $_.Trait)</code></td><td class=`"num`">$($_.AffectedSits)</td><td class=`"num`">$($_.OtherHits)</td><td><strong>$(ConvertTo-HtmlText $_.Requirement)</strong><br>$(ConvertTo-HtmlText $_.Affects)$S_Elsew$S_Src</td></tr>"
    }) -join "`n"
    @"
  <div class="notice notice-warn" style="margin-bottom:18px;"><strong>Licensing gap &mdash; the zeros below are not evidence of clean data.</strong> Every sensitive information type carrying the listed trait returned zero in that workload, while other types found plenty. An unlicensed type returns zero rather than an error, so it looks identical to a genuine nil result. Obtain the licence named below, or exclude those types from the scope and say so, but do not report them as clean.</div>
  <table><thead><tr><th>Workload</th><th>Trait</th><th class="num">Types affected</th><th class="num">Other types with findings</th><th>Licence required</th></tr></thead><tbody>
$S_GapRows
  </tbody></table>
"@
  })

  <h3 style="font-size:0.95em;margin:18px 0 8px;color:#1a1a2e;">Calls that failed ($($S_Failures.Count))</h3>
  <table><thead><tr><th>Type</th><th>Workload</th><th>Error</th></tr></thead><tbody>
$S_FailureRows
  </tbody></table>

  <h3 style="font-size:0.95em;margin:24px 0 8px;color:#1a1a2e;">Selected but absent from this tenant ($($S_MissingFromTenant.Count))</h3>
  <table><thead><tr><th>Type</th><th>Taxonomy GUID</th></tr></thead><tbody>
$S_MissingRows
  </tbody></table>

  <h3 style="font-size:0.95em;margin:24px 0 8px;color:#1a1a2e;">Present in the tenant but not in the taxonomy ($($S_UntaggedTenantSits.Count))</h3>
  <p class="caption">Custom types, or built-ins added by Microsoft since the taxonomy was written. They carry no tags, so no view or tag filter can select them. Query them by name with <code>-SensitiveInfoType</code>.</p>
  <table><thead><tr><th>Type</th><th>GUID</th><th>Publisher</th></tr></thead><tbody>
$S_UntaggedRows
  </tbody></table>
</div>

<div class="table-section">
  <h2>Appendix &mdash; run parameters and assumptions</h2>
  <div class="caption">Everything needed to reproduce this report. Taxonomy source: $(ConvertTo-HtmlText $S_Taxonomy.Source). Duration: $($S_Stopwatch.Elapsed.ToString('hh\:mm\:ss')). Calls made: $S_RealCalls of $S_TotalCalls planned.</div>

  <h3 style="font-size:0.95em;margin:0 0 8px;color:#1a1a2e;">Load-bearing assumptions</h3>
  <p class="caption">Two behaviours this report depends on have not been confirmed against a live tenant. Run the tool with <code>-VerifyAssumptions</code> to settle them.</p>
  <table><thead><tr><th>Assumption</th><th>Status</th><th>If wrong</th></tr></thead><tbody>
    <tr><td><strong>A1.</strong> <code>TotalCount</code> from a single-record call is complete, not truncated by sampling</td><td><span class="badge badge-direct">workload-dependent &mdash; not settled</span></td><td>Exchange matched an exhaustive export exactly. Teams did not &mdash; its total claimed roughly twice what paging returned. $(if ($S_Workloads -contains 'Teams') { 'Teams is in scope for this run &mdash; see the scope notes.' } else { 'Teams is not in scope for this run.' })</td></tr>
    <tr><td><strong>A2.</strong> Exchange and Teams return disjoint sets of items, so the Total column may sum them</td><td><span class="badge badge-indirect">unverified</span></td><td>$(if ($S_ExoTeamsBothInScope) { 'The Total column may double count Teams chat. Per-workload columns are unaffected.' } else { 'Not engaged by this run &mdash; Exchange and Teams were not both in scope.' })</td></tr>
  </tbody></table>

  <h3 style="font-size:0.95em;margin:24px 0 8px;color:#1a1a2e;">Parameters</h3>
  <table><thead><tr><th>Parameter</th><th>Value</th></tr></thead><tbody>
$S_ParamRows
  </tbody></table>
</div>

<div class="footer">Generated by ReportPurviewSIT.ps1 &nbsp;|&nbsp; $S_ReportDate &nbsp;|&nbsp; Read-only. No data left the tenant except to this file.</div>

<script>
function filterSits() {
  var search = document.getElementById('sitSearch').value.toLowerCase();
  var found = document.getElementById('foundFilter').value;
  var pii = document.getElementById('piiFilter').value;
  var tier = document.getElementById('tierFilter').value;
  var rows = document.querySelectorAll('#sitBody tr');
  var visible = 0;
  rows.forEach(function(r) {
    var text = r.textContent.toLowerCase();
    var matchSearch = !search || text.indexOf(search) > -1;
    var matchFound = found === 'all' || (r.getAttribute('data-found') || '') === found;
    var matchPii = pii === 'all' || (r.getAttribute('data-pii') || '') === pii;
    var matchTier = tier === 'all' || (r.getAttribute('data-tier') || '') === tier;
    if (matchSearch && matchFound && matchPii && matchTier) { r.classList.remove('hidden-row'); visible++; }
    else { r.classList.add('hidden-row'); }
  });
  document.getElementById('sitCount').textContent = 'Showing ' + visible + ' of ' + rows.length + ' types';
}

var sortDir = {};
function sortTable(tableId, col) {
  var tbody = document.getElementById(tableId).querySelector('tbody');
  var rows = Array.from(tbody.querySelectorAll('tr'));
  var key = tableId + '-' + col;
  var dir = sortDir[key] === 'asc' ? 'desc' : 'asc';
  sortDir[key] = dir;
  rows.sort(function(a, b) {
    var ta = a.cells[col] ? a.cells[col].textContent.trim() : '';
    var tb = b.cells[col] ? b.cells[col].textContent.trim() : '';
    var na = parseFloat(ta.replace(/,/g, ''));
    var nb = parseFloat(tb.replace(/,/g, ''));
    if (!isNaN(na) && !isNaN(nb)) { return dir === 'asc' ? na - nb : nb - na; }
    return dir === 'asc' ? ta.localeCompare(tb) : tb.localeCompare(ta);
  });
  rows.forEach(function(r) { tbody.appendChild(r); });
}
</script>
</body>
</html>
"@

$S_Html | Out-File -FilePath $S_HtmlFile -Encoding UTF8

}

# ---------------------------------------------------------------------------
# Console summary
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "Report complete." -ForegroundColor Green
if ($Format -in @('Csv', 'Both'))
{
    if ($Mode -eq 'Full')
    {
        Write-Host "  CSV (item x SIT) : $S_CsvFile  ($('{0:N0}' -f $S_DetailRows.Count) rows)"
    }
    else
    {
        Write-Host "  CSV (SIT x load) : $S_CsvFile"
    }
}
if ($Format -in @('Html', 'Both'))
{
    Write-Host "  HTML             : $S_HtmlFile"
}
if ($S_CheckpointOk)
{
    Write-Host "  Checkpoint     : $CheckpointPath"
}
else
{
    # Loud, because a stale checkpoint looks identical to a good one and only bites on resume.
    Write-Host "  Checkpoint     : FAILED TO WRITE - do not rely on -ResumeFrom for this run" -ForegroundColor Red
}
Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "  Types queried            : $($S_Runnable.Count)"
Write-Host "  Types with findings      : $($S_SitsWithFindings.Count)" -ForegroundColor $(if ($S_SitsWithFindings.Count -gt 0) { 'Yellow' } else { 'White' })
Write-Host "  Direct identifiers found : $($S_DirectPiiSits.Count)" -ForegroundColor $(if ($S_DirectPiiSits.Count -gt 0) { 'Yellow' } else { 'White' })
Write-Host "  Types not detected       : $($S_SitsClean.Count)"
Write-Host "  Not present in tenant    : $($S_MissingFromTenant.Count)"
Write-Host "  Untagged tenant SITs     : $($S_UntaggedTenantSits.Count)"
if ($Mode -eq 'Full')
{
    Write-Host "  Item x SIT rows          : $('{0:N0}' -f $S_DetailRows.Count)"
    Write-Host "  Distinct items           : $('{0:N0}' -f $S_DistinctItems)"
    Write-Host "  Containers affected      : $('{0:N0}' -f $S_ContainersAffected)"
    Write-Host "  High-confidence matches  : $('{0:N0}' -f $S_HighMatchTotal)"
    if ($S_TruncatedRows -gt 0)
    {
        Write-Host "  Rows at the match cap    : $S_TruncatedRows (MatchesTruncated = True; treat those counts as at least the value shown)" -ForegroundColor Yellow
    }
    if ($S_AdjustedRows -gt 0)
    {
        Write-Host "  Rows with band anomalies : $S_AdjustedRows (BandsAdjusted = True)" -ForegroundColor Yellow
    }
    if (-not $IncludeItemNames)
    {
        Write-Host "  Item names               : redacted to a stable hash. Use -IncludeItemNames to write real names." -ForegroundColor DarkGray
    }
    else
    {
        Write-Host "  Item names               : WRITTEN IN FULL - this CSV contains email subjects and file names" -ForegroundColor Yellow
    }
}
if ($S_Failures.Count -gt 0)
{
    Write-Host "  Failed calls             : $($S_Failures.Count)" -ForegroundColor Red
}
Write-Host ""
Write-Host "Item counts are per type. Do not add them across types - one item commonly carries many." -ForegroundColor DarkGray
