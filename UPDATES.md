# M365 License Optimization Report - Version History

## v0.3.1 - 23/02/2026

### Enhancement — LOA Rule Pack / Manual Audit Checklist
New optional `-RulePackPath` parameter loads `docs/LOA_RulePack_M365.json` (declarative rule pack).
Filters MANUAL-level rules and appends a **Manual Audit Checklist** section to the summary TXT
with step-by-step portal checks and Microsoft Learn doc references. Gracefully skips if file
not found. AUTO/SEMI rules are informational only — existing hardcoded detections remain authoritative.

### Enhancement — Archive Status & Auto-Expanding Archive Columns
Extended `Get-EXOMailbox` fetch to include `ArchiveStatus` and `AutoExpandingArchiveEnabled` properties.
Two new CSV columns: `Archive Status` (Active/None) and `Auto-Expanding Archive` (True/False).
Informational only — no new recommendations, enables future archive-based detection.

### New Detection — Suite Inversion (Bug #1)
The E5 upgrade check required 2+ paid add-ons, masking cases where a single heavy E5 mini-suite
(Identity Threat Protection €12/mo, or Information Protection Compliance €12/mo) plus E3 already
exceeds full E5 price. Lowered threshold from ≥2 to ≥1 paid add-ons. The cost comparison
(`$delta > 0`) naturally prevents false positives.

**New tag:** `SUITE INVERSION` — fires when E3 + add-ons > E5 cost. Existing `E5 CONSOLIDATION`
tag continues for cases where E5 is not strictly cheaper but offers value consolidation.
Confidence: High.

### Enhanced Detection — AI Add-On Overlap (Bug #2)
The existing `AI OVERLAP REVIEW` blanket-flagged all users with both Teams Premium and Copilot.
Now uses `$teamsMeetingsOrganized` to split into two tiers:
- `AI ADD-ON OVERLAP` (High) — 0 meetings organized → Teams Premium is definitively redundant
- `AI OVERLAP REVIEW` (Review) — >0 meetings → manual check needed for webinar features

### Fixed Detection — E5 Data Hoarder / Inactive Hold (Bug #3)
The disabled-account + litigation-hold message incorrectly stated "License MUST be retained to
maintain hold." This is factually wrong — Microsoft documentation confirms that removing the
license creates a free Inactive Mailbox that retains ALL content and holds indefinitely.

**Fix:** Two-tier detection replacing the old single message:
- `E5 DATA HOARDER` (High) — expensive suite (E3/E5/etc.) on disabled+held account → remove license
- `INACTIVE HOLD` (High) — cheaper license on disabled+held account → remove license
- Also fixed the same factual error in the shared mailbox litigation hold message.
- Fixed `$disabledCostAcc` regex to include new tags (costs would otherwise silently drop from Tier 1).

### New Detection — Seeded Visio Overlap (Bug #5)
Microsoft includes the lightweight "Visio in Microsoft 365" web app natively in E1/E3/E5 suites.
Users with standalone Visio Plan 1 (€4.70/mo) alongside a qualifying suite likely don't need it
for viewing/light editing.

**New tag:** `SEEDED VISIO OVERLAP` — fires when user has Visio Plan 1 + E1/E3/E5 suite and
SharePoint file activity < 5 (proxy for light usage). Confidence: Medium.

## v0.3.0 - 23/02/2026

### Infrastructure — UPN Case-Insensitive Normalization

All UPN-keyed lookup hashtables were using case-sensitive plain hashtables. Graph API and CSV
reports can return inconsistent UPN casing across paginated responses (e.g., "User@Domain.com"
on page 1 vs "user@domain.com" on page 3), and trailing whitespace can appear in CSV exports.
This caused silent data loss in every join — activity data would fail to match user records,
producing false "zero usage" / "no activity" results.

**Fix:** Systematic normalization across ~15 UPN insertion points:

- `Build-UPNLookup` (10 report lookups: ActiveUser, M365App, Email, Teams, OneDrive, SharePoint,
  Mailbox, EmailApp, ODUsage, TeamsDevice) — keys normalized with `.ToString().Trim().ToLower()`
- `$lkpActivations` build loop — key normalized with `.ToString().Trim().ToLower()`
- `$lkpUserObj`, `$lkpAssignedLicenses`, `$userLicenseMap` — user fetch loop normalizes `$upnKey`
- `$idToUpn` — now stores normalized lowercase UPN (cascades to PIM lookups)
- `$lkpSignIn`, `$lkpLicAssignment` — beta API fetch normalizes `$uUpnKey`
- `$lkpMailboxType`, `$lkpLitigationHold` — **BUG FIX**: was computing `$mbxUpnLower` but
  inserting with raw `$mbxUpn` — now correctly uses `$mbxUpnLower`
- `$lkpSmtpToUpn` — value now stores normalized lowercase UPN
- `$lkpMdoCoverageByUpn` — `$covUpn` key normalized
- `$lkpAdminRoles` — `$memberUpn` key normalized with `.ToString().Trim().ToLower()`

**Impact:** Every join (Email/OneDrive/Teams/AppPlatform/Activation/Mailbox/SignIn/AdminRoles)
now matches correctly regardless of source casing. Eliminates a class of false "no activity"
results without touching any recommendation logic.

## v0.2.1 - 23/02/2026

### Bug Fix — Multi-PC Gate Overcounting Windows Activations

The frontline blocker's `$winActTotal` summed the 'Windows' column across ALL activation rows
per user — including Visio, Project, and other Product Types. A user with Office on 1 PC and
Visio on 1 PC would show `$winActTotal = 2`, incorrectly blocking the F3 downgrade even though
Office is only on one device.

**Fix:** Added `Product Type -match 'Office|Microsoft 365 Apps|M365 Apps'` filter to only count
Office/M365 Apps activation rows when computing the multi-PC gate.

### Bug Fix — Confidence Scoring Ignores Missing Data Sources

Activity-based recommendations (NO ACTIVITY, SHELFWARE, Frontline Candidate, Business Downgrade,
etc.) could receive High or Medium confidence even when key upstream data sources were missing.
For example, if EmailActivity report failed to download, a user with `$em = $null` could receive
a "No Activity" recommendation with High confidence — even though the data was simply absent.

**Fix:** Added post-hoc confidence downgrade after the main recConfidence mapping. If the
recommendation category is activity-based and `$missingDataSources` includes any of
EmailActivity, TeamsActivity, OneDriveActivity, or M365AppPlatform, confidence is forced to
"Review". Deterministic categories (disabled account, duplicate, dormant, etc.) are unaffected.

### New Detection — Bundle Inefficiency (Basic + Apps Frankenstein)

**Tag:** `BUNDLE INEFFICIENCY` | **Confidence:** High

Detects users with Business Basic (email+Teams, €6/mo) + Apps for Business (desktop apps,
€11.50/mo) = €17.50/mo. M365 Business Standard (€12.50/mo) natively includes both. Consolidation
saves €5/mo per user and simplifies license management. Handles all alias SKU variants.

### New Detection — Premium Add-On Waste (Visio/Project Desktop → Plan 1)

**Tag:** `PREMIUM ADD-ON WASTE` | **Confidence:** Medium

Detects users with desktop-tier Visio Plan 2 (€14.60/mo) or Project Plan 3/5 (€29.30–€52.80/mo)
who pass the shelfware check (have product activation) but show NO Windows/Mac desktop activation
for that specific product. These users access via mobile/web only — the expensive desktop-tier SKU
is overkill. Recommends downgrade to web-only Plan 1 equivalent (Visio Plan 1 €4.70/mo, Project
Online Essentials €7/mo).

Integrated into the existing shelfware block for clean flow: no activation → SHELFWARE,
activation but no desktop → PREMIUM ADD-ON WASTE, has desktop activation → passes both checks.

### New Detection — Teams Phone Right-Sizing (Shared Device Waste)

**Tag:** `TEAMS PHONE RIGHT-SIZING` | **Confidence:** High

Detects non-human accounts (Shared Mailbox, Room Mailbox, Equipment Mailbox) assigned full Teams
Phone Standard (MCOEV, €8/mo) when they only need the cheaper Teams Shared Devices license
(MCOCAP, €2.50/mo) designed for common area phones, lobby devices, and conference rooms.
Saves €5.50/mo per device. Added MCOCAP pricing to built-in price table.

---

## v0.2.0 - 23/02/2026

### Architecture Fix — Floating-Point Finance Trap

Switched all currency values and accumulators from `[double]` to `[decimal]` to eliminate
IEEE 754 floating-point drift on large tenants. When adding `€12.50 + €3.50` tens of thousands
of times, `[double]` eventually produces `142589.499999999` instead of `142589.50`. This
undermines trust in the "Total Annual Spend" figure presented to CFOs.

**Changes:**
- `Get-SkuMonthlyPrice` now returns `[decimal]` (single gateway for all pricing lookups)
- All cost accumulators use `[decimal]` type constraint (`[decimal]$totalMonthlySpendAcc = 0`)
- Per-user `$userMonthlyCost` uses `[decimal]` type constraint
- JSON/CSV pricing overrides cast to `[decimal]` instead of `[double]`
- Cost-by-dimension running dictionaries (`$deptCostDict`, `$countryCostDict`, `$companyCostDict`)
  and recommendation distribution accumulator use `[decimal]`
- Delta report cost parsing uses `[decimal]` instead of `[double]`

### Architecture Fix — MDO Group Timeout Circuit Breaker

Added a circuit breaker to `Get-GroupMemberUpns` to prevent timeout on massive groups
(100k+ members). Security teams often apply Safe Links/Attachments policies to "All Employees"
dynamic groups. Expanding these groups via `/transitiveMembers` can take 10+ minutes with
100+ paginated API calls, causing throttling and potential script timeouts.

**Fix:** Before expanding a group, the function now checks the member count via a single
`$count` API call. If the group exceeds 10,000 members, it returns an `[ALL_TENANT]` marker
instead of expanding. The marker propagates through the MDO coverage chain:
`Get-GroupMemberUpns` → `Resolve-IdentityToMailboxSmtps` → `Add-MdoCoverage` → per-user
evaluation. All users in the tenant are treated as covered by that policy.

### New Detection — A La Carte Waste (Kiosk + Desktop Apps Clash)

**Tag:** `A LA CARTE WASTE` | **Confidence:** High

Detects users with Exchange Kiosk (€1/mo, 2 GB mailbox) + standalone M365 Apps for Enterprise
(€13.90/mo) = €14.90/mo. Consolidating into M365 Business Standard (€12.50/mo) saves money
AND massively upgrades the user: 50 GB mailbox, 1 TB OneDrive, Teams desktop.

### New Detection — Redundant Archive on Shared Mailbox

**Tag:** `REDUNDANT ARCHIVE` | **Confidence:** High

Detects shared mailboxes assigned both Exchange Plan 2 and standalone Exchange Online Archiving.
Plan 2 natively includes auto-expanding archives — the EOA add-on (€3/mo) is 100% redundant.
Junior admins often panic-buy EOA when a shared mailbox approaches 100 GB, not realizing the
archive capability is already there.

Also added `"EXCHANGEENTERPRISE" = @("EXCHANGE_ARCHIVE")` to `$suiteIncludes` so the duplicate
detection engine catches EXO Plan 2 + standalone EOA for ALL users (not just shared mailboxes).

---

## v0.1.1 - 23/02/2026

### Bug Fix — MDO Deserialized Object Enumeration Failure

Fixed three `foreach` loops in the MDO policy evaluation section (Preset Security Policies,
Safe Links rules, Safe Attachments rules) that threw `ArgumentTransformationMetadataException`
under `Set-StrictMode -Version Latest` on PowerShell 7.

**Root cause:** EXO implicit remoting returns `Deserialized.*` objects. When a cmdlet returns
a single result, the `foreach` statement's `GetEnumerator()` call fails because the
deserialized type doesn't implement `IEnumerable` properly. PowerShell falls back to type
conversion attempts (including `Int32`), which throws:
```
Cannot convert the "Strict Preset Security Policy" value of type
"Deserialized.Microsoft.Exchange.Management.SystemConfigurationTasks.ATPProtectionPolicyRule"
to type "System.Int32".
```

**Fix:** Replaced `foreach ($r in @(Get-*Rule))` with pipeline-based `Get-*Rule | ForEach-Object`.
The pipeline handles single/multiple/null results correctly for all object types. Changed
`continue` → `return` inside the `ForEach-Object` blocks (which acts as `continue` in pipeline context).

**Affected lines:** ~2588-2639 (MDO policy evaluation in Section 7b)

---

## v0.1.0 - 23/02/2026

### Detection Engine Expansion — 40+ Optimization Scenarios

Massive expansion of the recommendation engine with 25+ new detection scenarios, granular
capability analysis, and architectural improvements to duplicate detection. The script now
covers the full spectrum of M365 license waste: from pure waste (dormant, disabled) through
right-sizing (frontline, web-only) to advanced behavioral analysis (background sync ghosts,
automation accounts, cold storage).

#### New Detection Scenarios — Tier 1: Pure Waste

- **GUEST ACCOUNT WASTE** — External B2B guest accounts (#EXT#) holding paid licenses.
  Guests can access Teams/SharePoint using their home tenant's license for free.
  Confidence: High.
- **NON-HUMAN ACCOUNT WASTE** — Service principals, room/equipment mailboxes, or accounts
  with naming patterns (svc-, app-, test-, noreply@, etc.) holding full user licenses.
  Confidence: High.
- **VIRAL/EXPLORATORY** — Users on free/viral SKUs (Power BI Free, Teams Exploratory, etc.)
  cluttering the license inventory. Recommends cleanup or conversion. Confidence: High.
- **INACTIVE MAILBOX (FREE)** — Unlicensed mailbox on Litigation Hold silently converts to
  a free Inactive Mailbox. Split from the previous blanket "data purged in 30 days" warning
  to avoid unnecessary panic. Confidence: High.
- **BACKGROUND SYNC ONLY** — Users with zero interactive M365 activity but OneDrive file
  sync running in the background (abandoned laptop syncing). Slips past standard "no activity"
  detection because sync counts as activity. Confidence: High.

#### New Detection Scenarios — Tier 2: Right-Sizing

- **TEAMS UNBUNDLING** — Users on legacy bundled suites with zero Teams activity. Switch to
  the newer "Without Teams" (EEA/Global) SKU equivalent to save the bundled Teams cost.
  Confidence: Medium.
- **FRONTLINE ADD-ON BLOAT** — F-series users whose bolted-on add-ons (Exchange Plan 2,
  Entra P2, Power BI Pro, etc.) push total cost above Business Premium or E3. Recommends
  consolidating to a full suite. Confidence: Medium.
- **FRONTLINE BLOCKED** — User qualifies for frontline downgrade on activity, but has
  Office activated on 2+ Windows devices. F3 only provides VDI shared-device rights —
  downgrading would deactivate Office on all dedicated PCs. Confidence: High.
- **FRONTLINE CANDIDATE (HIGH CONFIDENCE)** — Enhanced existing frontline detection: when
  `$activatedPlatforms` is empty (zero device activations ever), the recommendation is
  prefixed with "(HIGH CONFIDENCE)" as proof the user has no corporate PC/Mac.
- **F3→F1 MICRO-DOWNGRADE** — F3 users whose only activity is Teams mobile/web. F1 covers
  Teams + basic web apps at ~50% the cost. Confidence: Medium.
- **STANDALONE DESKTOP APP WASTE** — Users with a standalone M365 Apps subscription who
  never use desktop Office (web/mobile only). Confidence: Medium.
- **OVER-LICENSED ARCHIVE** — Users on E3/E5 with zero interactive activity whose only value
  is mailbox retention. Could use a cheaper Exchange Online Plan 2 (archive-only) license.
  Confidence: Review.

#### New Detection Scenarios — Activity & Behavioral

- **HEAVY EXTERNAL SHARER** — Users sending/sharing more than 50% of their content
  externally. Flags potential DLP review rather than license action. Confidence: Review.
- **EXPENSIVE COLD STORAGE** — Zero activity but significant data stored (mailbox > 10 GB
  or OneDrive > 50 GB). You're paying full suite price just for storage. Confidence: High.
- **MDM/MAM WASTE** — User holds Intune/EMS entitlement but shows 100% web-only access
  (no desktop or mobile app usage). Intune has nothing to manage. Confidence: Medium.
- **AUTOMATION ACCOUNT** — Admin account with no interactive sign-in but recent
  non-interactive sign-in. Service/automation account — not truly dormant. Replaces
  false "DORMANT ADMIN RISK" alerts. Confidence: High.
- **LEGACY SERVICE ACCOUNT** — Accounts using only POP3/IMAP4/SMTP legacy protocols on
  premium suites. Likely a service account that only needs Exchange Online Plan 1.
  Confidence: High.

#### New Detection Scenarios — Shelfware & Add-Ons

- **INTUNE SUITE WASTE** — Intune Suite add-on assigned but zero advanced features used
  (Remote Help, Tunnel, etc.). Base Intune Plan 1 from E3/E5 is sufficient. Confidence: Medium.
- **WINDOWS LICENSE WASTE** — Standalone Windows E3/E5 assigned alongside a suite that
  already includes it (M365 E3/E5). Caught by enhanced duplicate detection. Confidence: High.

#### Capability Analysis & Security Coverage

- **Granular capability flags** — Per-user boolean flags for MdoP1/P2, MdeP1/P2,
  DlpEmailFiles/Teams/Endpoint, Mdi, MdcApps, Xdr via the `Merge-UserCapabilities` function
  and `$planCapabilities` table (~130 SKU entries).
- **Equivalence classes** — `$hasFullDefenderStack`, `$hasFullPurviewStack` with
  short-circuit upsell logic: only recommends the next tier if the user is "one feature away"
  from full coverage.
- **New columns**: `SecurityCoverageLevel` (None/Basic/Advanced/E5-equivalent) and
  `ComplianceCoverageLevel` (None/Basic/Advanced/E5-equivalent) for heat-map analysis.

#### Duplicate Detection Engine Improvements

- **Bundle consolidation** — Detects users paying for Office 365 + EMS + Windows individually
  instead of a unified M365 bundle (E3 and E5 tiers). Also detects Entra P2 + Entra Governance
  overlap recommending Entra Suite.
- **2-hop alias resolution** — `$skuCoverageAliases` map handles service plan ID variations
  (e.g., `EXCHANGE_S_STANDARD` ↔ `EXCHANGESTANDARD`) so the duplicate engine catches coverage
  that was previously missed due to Microsoft's inconsistent naming.
- **AI overlap detection** — Flags Teams Premium + Copilot overlap (both include intelligent
  meeting recap).
- **Entra Suite overlap detection** — Upgraded from generic "LICENSING CHECK" to specific
  "ENTRA SUITE OVERLAP" recommendation.
- **EOA overlap** — Exchange Online Archiving added to `$suiteIncludes` for E3/E5 suites so
  standalone EOA is caught as a duplicate.

#### Logic Flaw Fixes

- **"No Desktop Apps" False Positive** — No longer flags web-only SKU users (Business Basic,
  etc.) to "consider web-only license" — gated on `$hasDesktopAppEntitlement`.
- **Audio Conferencing Price** — MCOMEETADV price corrected from €2.50 to €0.00 (free since 2023).
- **Teams Premium Shelfware Threshold** — Changed from `-eq 0` to `-lt 3` to reduce noise
  for a €10/mo SKU.
- **Teams Premium Organizer Fix** — Shelfware check now uses `Meetings Organized Count`
  instead of total `Meeting Count`. Premium features are organizer-driven.
- **E5 Step-Up vs. Shelfware** — E5 UPGRADE OPPORTUNITY now appended with a caveat to verify
  each add-on is actively used before committing to a permanent E5 step-up.
- **Copilot False Dormancy** — Copilot inactive message updated to warn that web-based
  Copilot Chat (copilot.microsoft.com) doesn't generate standard app telemetry. Added
  `COPILOT REVIEW` caveat to the NO ACTIVITY block for Copilot holders.
- **Power Platform False NO ACTIVITY** — Power Automate/Power Apps run server-side without
  M365 app telemetry. Appended `POWER PLATFORM REVIEW` caveat.

#### New CSV Columns

| Column | Description |
|--------|-------------|
| `Last Non-Interactive Sign-In` | Last non-interactive sign-in date (automation/service accounts) |
| `Days Since Non-Interactive Sign-In` | Days since last non-interactive sign-in |
| `Teams: Meetings Organized` | Meetings organized count (vs total meetings attended) |

#### Executive Summary & Tiered Savings Model

- **`$totalIdentifiedWaste`** = dormant + disabled + deleted + noActivity + shelfware + sharedMbx
- **`$tier1Waste`** = `$totalIdentifiedWaste` + duplicate cost (immediate: remove license)
- **`$tier2Savings`** = frontline + businessBasic + exoPlan2 + e5Upgrade + bundleConsolidation
  (right-sizing: downgrade SKU delta)
- **`$totalMoneyOnTable`** = `$tier1Waste` + `$tier2Savings`
- Output: text block in summary, `M365_ExecutiveSummary_*.csv`, "Executive Summary" Excel sheet

#### Documentation

- **`docs/Detection_Scenarios.md`** — Comprehensive reference of all 40+ detection scenarios
  with memorable names (Frankenstein Frontline, Ghost in the Machine, Phantom Sync, etc.),
  tags, confidence levels, plain-English explanations, and real-world scenarios.
- **`docs/Logic_Flaws.md`** — All 13 logic flaws tracked and resolved.
- **`docs/ideas.md`** — Feature ideas tracked and implemented.
- **`docs/bugs.md`** — Bug reports tracked and resolved.

---

## v0.0.9 - 18/02/2026

### Hardening, Module Safety & Parameter UX

Three improvements for production reliability, locked-down environments, and parameter usability.

#### Parallel Download Hardening

- **Fully qualified type names**: `[RunspaceFactory]` → `[System.Management.Automation.Runspaces.RunspaceFactory]`
  and `[PowerShell]` → `[System.Management.Automation.PowerShell]` in the parallel download section.
  Prevents "Unable to find type" errors on hosts where the shorthand type accelerators are not loaded
  (e.g., constrained language mode, some remote sessions).
- **`$PSScriptRoot` guard**: `$PSScriptRoot` is empty when the script is dot-sourced, run from ISE, or
  invoked via ScriptBlock. The JSON data path now falls back to `(Get-Location).Path` if `$PSScriptRoot`
  is empty.
- **UTF-8 BOM stripping + `Invoke-WebRequest`**: Most parallel-downloaded reports returned 0 rows
  because Graph report blob storage prepends a UTF-8 BOM (`U+FEFF`) to CSV content.
  `Invoke-RestMethod` preserved the BOM in the string, causing the first CSV column header to be
  corrupted (e.g. `\xFEFFReport Refresh Date` instead of `Report Refresh Date`). `ConvertFrom-Csv`
  then produced objects with unrecognizable property names → effectively 0 usable rows. Fixed by:
  1. Switching from `Invoke-RestMethod` to `Invoke-WebRequest -UseBasicParsing` — prevents
     auto-parsing that can corrupt CSV content if content-type headers are ambiguous.
  2. Stripping the BOM character from the first position of the response content before
     piping to `ConvertFrom-Csv`. Handles both Unicode BOM (`U+FEFF`) and raw UTF-8 BOM
     bytes decoded as Latin-1 (`ï»¿`), which varies between PS 5.1 and PS 7.
- **Defensive `Build-UPNLookup`**: Before iterating rows, validates that the expected UPN
  column exists in the first row's property names. If missing (corrupted headers, empty data),
  logs a warning with the actual column names found and returns an empty hashtable instead of
  throwing. Same guard added to the activations multi-row aggregation loop.

#### `-AutoInstallModules` Switch

Module installation is no longer automatic. By default, missing required modules cause the script to
**stop with a clear error** listing the exact `Install-Module` commands needed:

```
Missing required modules:
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
    Install-Module Microsoft.Graph.Users -Scope CurrentUser

Install the modules above, or re-run with -AutoInstallModules to install automatically.
```

- **Safe default**: No modules are installed unless `-AutoInstallModules` is explicitly set.
  This is critical for locked-down servers, CI pipelines, and multi-admin environments where
  auto-installing modules may violate policy.
- **Optional modules unaffected**: `ExchangeOnlineManagement` and `ImportExcel` already had
  graceful degradation — they continue to warn and skip when absent.

#### Parameter UX & Safety Improvements

- **`SupportsShouldProcess`**: Added to `[CmdletBinding()]`. The `-UnhideUserData` privacy
  toggle now respects `-WhatIf` and `-Confirm`, allowing tenants to preview the change before
  it modifies the organization's report privacy setting.
- **`HelpMessage`** attributes on all parameters — `Get-Help` now shows meaningful descriptions
  for every parameter.
- **`ValidateRange`** on all threshold parameters:
  - High thresholds: `[ValidateRange(1, [int]::MaxValue)]` (must be positive)
  - Low thresholds: `[ValidateRange(0, [int]::MaxValue)]` (non-negative)
  - `InactiveSignInDays`: `[ValidateRange(1, 365)]`
  - Cross-validation: script throws if any High threshold ≤ its corresponding Low threshold.
- **`ValidateScript`** on file path parameters (`PricingCsvPath`, `SkuDataPath`, `PriorReportPath`):
  `{ Test-Path $_ -PathType Leaf }` — catches bad paths at parameter binding time with a clear
  PowerShell validation error instead of failing later during execution.
- **`OutputFolder` auto-creation**: Moved to script startup (before any work begins). The later
  redundant `Test-Path` check before streaming is removed.

---

## v0.0.8 - 18/02/2026

### EXO Independence, Streaming Export & Delta Report

Three improvements addressing the final architectural gaps for large-tenant production use.

#### Feature 1: `-SkipEXO` Switch

Adds a `-SkipEXO` switch parameter that prevents the script from connecting to Exchange
Online entirely. When active:

- **Mailbox type detection** (Shared/Room/Equipment) is skipped — no Graph API equivalent exists
- **Litigation hold detection** is skipped — EXO-only feature
- **MDO policy coverage** (Safe Links/Attachments) is skipped — EXO cmdlets only
- All affected lookups degrade gracefully (empty hashtables, already guarded with `.ContainsKey()`)
- Clear warnings logged to both console and summary TXT with `$script:skippedDataWarnings`
- The script runs fully on Graph API alone, producing all other recommendations without error

This was confirmed through research: Microsoft Graph's `getMailboxUsageDetail` report does
**not** include a "Recipient Type" column in either v1.0 or beta (GitHub issue #1581 was
closed without adding it). Mailbox type detection is exclusively an EXO PowerShell capability.

#### Feature 2: Streaming CSV Export & Memory Optimization

Eliminates the `$mergedRows` List and `$userServicePlanRows` List entirely. On a 50k-user
tenant, this saves approximately **2+ GB of RAM**:

- **Main CSV**: Streamed row-by-row via `System.IO.StreamWriter` during the merge loop
  (BOM-less UTF8). Each PSCustomObject is constructed, written, then discarded.
- **Service Plan CSV**: Also streamed during the license mapping loop.
- **Summary statistics**: All 27 counters + 6 cost accumulators computed inline during the
  merge loop (no post-loop pass needed).
- **Cost breakdowns**: Running dictionaries (`$deptCostDict`, `$countryCostDict`,
  `$companyCostDict`) replace the `Group-Object` pipelines that previously scanned all rows.
- **Recommendation distribution**: Running counter hashtable replaces the `Group-Object`
  pipeline with regex matching.
- **Intensity cross-tab**: Running dictionary replaces the `Where-Object | Group-Object`
  pipeline.
- **Excel generation**: Reads the streamed CSV back via `Import-Csv | Export-Excel`.
- **`Safe-Sum` function**: Removed — no longer needed (all aggregation is dictionary-based).
- **`ConvertTo-CsvLine` helper**: New function using `[string[]]` array + `-join ','` for
  proper CSV escaping (avoids `StringBuilder.Append()` overload ambiguity in PowerShell).
- **StreamWriter disposal**: Wrapped in `try/finally` to ensure cleanup on error or Ctrl+C.
- **Bug fix**: Renamed `$teamsNoDesktop` counter to `$teamsNoDesktopCount` to avoid collision
  with the per-user boolean variable of the same name.
- **Bug fix**: Renamed `$pid` to `$principalId` in PIM role processing loops — `$PID` is a
  read-only PowerShell automatic variable (current process ID).

#### Suite Inclusion Expansion & Security Check Accuracy

Complete rewrite of `$suiteIncludes` and security SKU arrays against the authoritative
Microsoft licensing reference:
https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference

- **Rebuilt `$suiteIncludes` from Microsoft reference** (28 entries, up from 17):
  - **Removed 4 phantom SKUs** that don't exist in Microsoft's reference:
    `MICROSOFT365_E3`, `MICROSOFT365_E5`, `MICROSOFT365_E5_NOPSTNCONF`, `M365_BUSINESS_PREMIUM`
  - **Fixed `SPE_E3`**: Now correctly includes `INTUNE_A` and `AAD_PREMIUM` (confirmed by
    Microsoft reference — M365 E3 includes Intune and Entra P1)
  - **Fixed `SPE_E5` / `SPE_E5_NOPSTNCONF`**: Added `ATA`, `ADALLOM_S_STANDALONE`, `WIN_DEF_ATP`
    (Defender for Identity, Cloud App Security, Defender for Endpoint)
  - **Fixed `EMSPREMIUM`**: Corrected `ADALLOM_STANDALONE` to `ADALLOM_S_STANDALONE` per reference
  - **Added `SPE_E5_CALLINGMINUTES`** (M365 E5 with Calling Minutes)
  - **Added EEA "no Teams" variants**: `Microsoft_365_E5_(no_Teams)`,
    `O365_w/o_Teams_Bundle_M5`, `Office_365_w/o_Teams_Bundle_Business_Premium`,
    `Microsoft_365_ Business_ Premium_(no Teams)` (EU regulation variants)
  - **Added Education A3**: `M365EDU_A3_FACULTY`, `M365EDU_A3_STUDENT`
  - **Added Education A5 use benefit**: `M365EDU_A5_STUUSEBNFT`
  - **Added `DEVELOPERPACK_E5`** (E5 Developer without Windows/Audio Conf)
  - **Added security add-ons**: `IDENTITY_THREAT_PROTECTION` (E5 Security),
    `IDENTITY_THREAT_PROTECTION_FOR_EMS_E5`, `SPE_F5_SEC` (F5 Security),
    `SPE_F5_SECCOMP` (F5 Security + Compliance), `M365_SECURITY_COMPLIANCE_FOR_FLW`
  - **Added `M365_F1_COMM`** (M365 F1 alternate variant)
  - **Added `MDE_SMB`** to `SPB` (Business Premium includes Defender for Business)
  - **Fixed EDU A5 entries**: Added `MCOEV`, `MCOMEETADV`, `WIN_DEF_ATP`, `ATA`,
    `ADALLOM_S_STANDALONE` (all confirmed in Microsoft reference)
- **Fixed `$hasDefender` / `$hasDefenderP2` / `$hasPurview`**: Now use `$effectiveSkuSet`
  instead of raw `$userSkuList`. Prevents false "SECURITY GAP" on suite holders.
- **Fixed `$anyDefenderSku` / `$defenderP2Skus`**: Added `ADALLOM_S_STANDALONE` (in addition
  to `ADALLOM_STANDALONE`), `IDENTITY_THREAT_PROTECTION_FOR_EMS_E5`
- **Fixed `$businessPremiumSkus`**: Updated to include EEA "no Teams" variants
- **Fixed `$businessNoSecurity`**: Removed phantom SKUs, added EEA "no Teams" Business
  Basic/Standard variants
- **Fixed `$e3Suites`**: Replaced phantom `MICROSOFT365_E3` with `M365EDU_A3_FACULTY`,
  `M365EDU_A3_STUDENT`
- **Downloaded `M365_ServicePlanReference.csv`**: Microsoft's authoritative SKU reference
  file saved locally for future cross-reference

#### License Assignment Error Detection

The `licenseAssignmentStates` beta API already returns `state` and `error` fields but the
script was not using them. Now detects group-based licensing failures:

- **New `License Errors` column** in the main CSV, between License Groups and Last License
  Change. Shows errors like "CountViolation" (insufficient seats), "MutuallyExclusiveViolation"
  (conflicting plans), "DependencyViolation", etc.
- **`LICENSING ERROR` recommendation** generated for any user with assignment errors
- **`$licenseErrors` counter** in summary TXT under LICENSING COMPLIANCE section
- **`License Error` recommendation category** for the distribution chart
- Ref: https://learn.microsoft.com/en-us/entra/identity/users/licensing-powershell-graph-examples

#### Group License Inventory

New section that queries all Entra ID groups with assigned licenses:

- **Data fetched**: Group name, ID, membership type (Dynamic/Assigned), member count,
  assigned SKUs with disabled plan counts
- **Summary TXT**: Group licensing inventory section listing all licensing groups
- **Excel worksheet**: "Group Licensing" sheet with table formatting
- Uses `Get-MgGroup -All` with `AssignedLicenses` property + member count via `$count` endpoint

#### Disabled Account Detection Fix

The "Disabled accounts (licensed)" counter was always reporting 0 because the Graph user
query filtered `accountEnabled eq true` by default, excluding disabled users entirely.

- **Always fetch all users**: Removed the `accountEnabled eq true` Graph filter. Disabled
  accounts are now always retrieved so the waste metric works correctly.
- **Merge loop skip**: Without `-IncludeDisabledAccounts`, disabled users that have **no
  license** are skipped (nothing to flag). Disabled+licensed users are always processed
  through the full recommendation engine for waste detection.
- **`-IncludeDisabledAccounts`**: Meaning updated — now controls whether disabled+unlicensed
  users appear in the report, not whether disabled users are fetched at all.
- **Console feedback**: User retrieval message now shows the disabled account count.

#### Feature 3: Delta Report via `-PriorReportPath`

Adds a `-PriorReportPath` parameter that accepts a previous run's main CSV for delta analysis:

- **Schema-agnostic**: Uses `$row.PSObject.Properties.Name` safe access for every column.
  Works with CSVs from any prior version — missing columns produce empty values, no errors.
- **Delta CSV** (`M365_LicenseDelta_*.csv`): 26 columns comparing each user across runs:
  - Change Type (New / Removed / Existing)
  - License changes (prior vs current SKUs)
  - Cost delta (EUR, per user + aggregate)
  - Recommendation category changes
  - Dormancy shifts (Became Dormant / Became Active)
  - Service intensity changes (Exchange, Teams, OneDrive, SharePoint)
  - Account status changes
- **Copilot adoption tracking**: New Copilot license holders + inactive Copilot users
- **Waste addressed**: Tracks prior waste categories that transitioned to non-waste
- **Delta summary**: Appended to the main summary TXT with user changes, cost trends,
  recommendation changes, dormancy shifts, and Copilot adoption metrics
- **Excel worksheet**: "Delta Analysis" sheet with conditional formatting (green = New users,
  red = Removed users, EUR formatting on cost columns)
- **Dashboard metrics**: Delta section added to the Dashboard sheet when prior report is provided

#### Runtime Bug Fixes

Two runtime errors fixed that occurred under `Set-StrictMode -Version Latest`:

- **Parallel download "Unable to index into System.Boolean"**: The `$reportDefs` array of
  arrays was being flattened by PowerShell's `@()` operator. Inner arrays like
  `@("name", "report", $true)` were merged into a single flat array, so `foreach ($def in
  $reportDefs)` iterated individual strings/booleans instead of sub-arrays, and `$def[1]`
  on a boolean threw. Fixed by prefixing each inner array with the comma operator:
  `,@("name", "report", $true)` to prevent unrolling.
- **MDO `.Count` property errors** (Preset Security Policy, Safe Links, Safe Attachments):
  `Get-ScopedMailboxSmtpsFromRule`, `Resolve-IdentityToMailboxSmtps`, and
  `Get-GroupMemberUpns` returned `HashSet[string]` objects via `return $set`. PowerShell's
  pipeline enumeration unwraps collections — an empty HashSet becomes `$null`, a
  single-item HashSet becomes a bare string. Under strict mode, `$scope.Count` on `$null`
  or a string throws. Fixed by using `return ,$set` (comma operator) on all return
  statements in these three functions, which prevents pipeline unwrapping.

---

## v0.0.7 - 18/02/2026

### Compliance Safety & Summary Performance

Adds litigation hold compliance check to prevent dangerous license removal recommendations,
and consolidates 33+ separate array scans into a single-pass summary statistics loop.

#### Litigation Hold Compliance Check

Shared mailbox recommendations now check `LitigationHoldEnabled` before advising license
removal. Previously, the script could recommend removing a license from a shared mailbox
that was under Litigation Hold — removing the license disables the hold and may allow data
deletion, creating a critical compliance risk.

- **EXO fetch**: Added `LitigationHoldEnabled` to `Get-EXOMailbox -Properties`
- **Lookup**: New `$lkpLitigationHold` hashtable for per-user hold status
- **Recommendation logic**: Litigation hold is checked FIRST in the shared mailbox branch:
  - Hold active → "license required to maintain hold. Do NOT remove license"
  - No hold, under 50 GB → "does not require a user license. Remove user license"
  - No hold, over 50 GB → "requires a license to retain auto-expanding archive"
- **New column**: `Litigation Hold` (True/False) in main report
- **New category**: "Litigation Hold" in recommendation distribution and `$recCategory`
- **Summary stat**: "Litigation Hold (active)" count in ACCOUNT & ROLE FLAGS section

#### Single-Pass Summary Statistics

Replaced 27 individual `@($mergedRows | Where-Object {...}).Count` calls and 6 `Safe-Sum`
pipeline passes with a single `foreach ($r in $mergedRows)` loop that accumulates all
counters and cost sums in one iteration:

- **Before**: 33+ full-array scans via `Where-Object` + `Measure-Object` pipelines,
  each iterating the entire `$mergedRows` collection. On a 50k-user tenant this means
  ~1.65 million comparisons.
- **After**: Single `foreach` loop with property checks and `-match` tests per row.
  All 27 counters and 6 cost accumulators computed in one pass (~50k iterations total).
- **Retained**: `Safe-Sum` helper and `Group-Object` pipelines for cost-by-department /
  country / company breakdowns (these require grouping and cannot be pre-computed).
- **Typical improvement**: ~5-10× faster summary computation on large tenants.

---

## v0.0.6 - 18/02/2026

### Resilience & Memory Improvements

Adds Graph API throttle handling, unified error tracking, and memory optimisation for large
tenants.

#### Graph API Retry with Exponential Backoff

New `Invoke-GraphWithRetry` helper wraps all paginated Graph API calls with automatic retry
on HTTP 429 (throttled) and 5xx (transient server) errors:

- Respects `Retry-After` response header when present
- Exponential backoff: 2s, 4s, 8s, 16s, 32s (max 5 retries, capped at 120s)
- Falls back to re-throw on non-retryable errors or exhausted retries
- Applied to all 8+ paginated loops (subscriptions, beta users, PIM, CA, MDO group members)
  and the `Download-GraphReport` sequential fallback functions

#### Unified Error Tracking (`$script:skippedDataWarnings`)

All data collection `try/catch` blocks now log skipped data sources to a script-scoped list.
The summary report includes a new "DATA COLLECTION WARNINGS" section at the end that surfaces
every error encountered during the run, so users know which columns may be incomplete:

- Subscriptions, sign-in activity, mailbox types, PIM, Conditional Access, admin roles
- MDO policy evaluation (Built-in Protection, Preset Policies, Safe Links, Safe Attachments)
- Previously 4 bare `catch { }` blocks in MDO silently swallowed errors — now all log warnings

#### Memory Optimisation

- **`$lkpUserObj` slimmed** — Previously stored full `MgUser` objects (~30+ properties each).
  Now stores only the 5 properties actually used: `UserType`, `Department`, `CompanyName`,
  `Country`, `AccountEnabled`. On a 50k-user tenant this saves ~200+ MB of hashtable memory.
- **`$allUsers` released** — After building all lookups and UPN sets, the full user array is
  set to `$null` and garbage collected. The license mapping loop (which needs `AssignedLicenses`)
  runs before this point, so no data is lost.

---

## v0.0.5 - 18/02/2026

### Safety, Performance & Maintainability

Three architectural improvements: crash-safe cleanup, parallel report downloads, and
externalized SKU data.

#### try...finally Safety Wrapper

The main execution logic (Sections 3-7: report downloads through Excel export) is now wrapped
in a `try...finally` block. The `finally` block guarantees that:
- The report privacy setting (`displayConcealedNames`) is restored to `$true` even if the
  script crashes, throws an unhandled error, or is interrupted with Ctrl+C.
- Exchange Online and Microsoft Graph sessions are cleanly disconnected.

Previously, a script failure between the privacy toggle (Section 2) and cleanup (end of script)
would leave the tenant with user names visible in reports indefinitely.

#### Parallel Report Downloads (Runspace Pool)

The 11 Graph usage report downloads now run in parallel via a .NET runspace pool. This works
on both PowerShell 5.1 and 7+.

- Extracts the bearer token from the active MgGraph session
- Creates 11 parallel runspaces, each using `Invoke-RestMethod` with the token
- Automatic fallback to sequential downloads if token extraction fails
- Reports timing: `Downloaded 11 reports in X.Xs (parallel|sequential)`
- Typical improvement: 70-80% reduction in download phase time on large tenants

#### Externalized SKU Data (`M365SkuData.json`)

SKU friendly names (~65 entries) and monthly EUR prices (~63 entries) can now be maintained
in an external JSON file instead of editing the script source.

- **File:** `M365SkuData.json` in the same directory as the script
- **Format:** `{ "skuFriendlyNames": { "SKU": "Name", ... }, "skuMonthlyPricesEUR": { "SKU": 0.00, ... } }`
- **Loading order:** Built-in defaults → JSON overrides → `-PricingCsvPath` overrides (last wins)
- **Portable mode:** If JSON file is absent, built-in defaults are used silently
- **Explicit path:** Use `-SkuDataPath` to point to a JSON file in a different location

#### New Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-SkuDataPath` | `M365SkuData.json` (same dir) | Path to external SKU names + prices JSON |

---

## v0.0.4 - 18/02/2026

### Evidence-Driven License Compliance Checks (PIM, Risk-based CA, MDO)

Replaces generic "maybe you need P2" upsells with evidence-driven checks that detect **actual
usage** of Entra ID P2 and Defender for Office 365 features. Flags users who are covered by
PIM role assignments, risk-based Conditional Access policies, or Defender for Office 365 Safe
Links/Attachments rules but lack the corresponding license entitlement.

#### New Data Collection Steps

- **[7b/12] MDO policy coverage** — Evaluates four layers of Defender for Office 365 protection:
  Built-in Protection, Preset Security Policies, custom Safe Links rules, and custom Safe
  Attachments rules. Resolves scoped mailboxes via group expansion and SMTP-to-UPN mapping.
  Output: per-user policy coverage sources.
- **[8b/12] PIM & risk-based Conditional Access** — Fetches PIM role eligibility and assignment
  schedule instances from `/v1.0/roleManagement/directory/` endpoints. Fetches Conditional
  Access policies and filters for those with `signInRiskLevels` or `userRiskLevels` conditions.
  Resolves scoped users via group expansion. Output: per-user PIM roles and risk-based CA
  policy names.

#### New Recommendations

- **LICENSING CHECK — PIM eligibility** — User has PIM-eligible roles but no Entra ID P2 /
  Entra Governance entitlement in their effective SKU set. Only fires when PIM role assignments
  are actually present.
- **LICENSING CHECK — Risk-based CA** — User is explicitly targeted by a risk-based Conditional
  Access policy (scoped, not tenant-wide "All Users") but no Entra ID P2 entitlement found.
  Avoids noise from policies that target all users.
- **LICENSING CHECK — MDO policy scope** — Mailbox appears in scope of non-built-in (custom or
  preset) Safe Links/Attachments rules but no Defender for Office 365 entitlement found.
  Common gap with Exchange Plan 1 and shared mailboxes.

#### New Columns

| Column | Description |
|--------|-------------|
| `PIM Eligible Roles` | Semicolon-separated PIM-eligible role names |
| `PIM Active Roles` | Semicolon-separated PIM-active role assignments |
| `Risk-based CA Policies` | Names of risk-based CA policies covering this user |
| `MDO Policy Coverage` | Defender for Office 365 policy sources covering this mailbox |

#### Effective Entitlement Expansion

Per-user `$effectiveSkuSet` now expands suite SKUs (e.g. E3 → includes AAD_PREMIUM, E5 →
includes AAD_PREMIUM_P2 + ATP_ENTERPRISE) before checking entitlement. This prevents false
positives where a user has the required feature via suite inclusion rather than a standalone SKU.

#### Other Changes

- **Graph scope**: Added `Policy.Read.All` (required for Conditional Access policy reading).
- **EXO fetch**: Added `PrimarySmtpAddress` to mailbox properties for SMTP-to-UPN resolution.
- **Lookup maps**: `$idToUpn` / `$upnToId` for resolving PIM principal IDs and CA user scopes.
- **SMTP lookups**: `$lkpMailboxPrimarySmtp` and `$lkpSmtpToUpn` for MDO rule scope resolution.
- **Recommendation category**: Added "Licensing Check" to `$recCategory` classification and
  recommendation distribution grouping.
- **Summary**: New "LICENSING COMPLIANCE (evidence-driven)" section with licensing check count.

---

## v0.0.3 - 18/02/2026

### License Costs (EUR), Excel Output, Cost Breakdowns, Last License Change

Adds financial visibility with built-in EUR pricing, cost breakdowns by department/country/company,
last license change tracking, and a structured multi-worksheet Excel workbook with pivot-style
tables, charts, and conditional formatting.

#### New Features

- **Built-in EUR pricing table** - ~55 common M365 SKUs with monthly list prices in EUR.
  Override with `-PricingCsvPath` pointing to a CSV with `SkuPartNumber,MonthlyPriceEUR` columns.
- **Per-user cost columns** - `Monthly License Cost (EUR)` and `Annual License Cost (EUR)`
  computed from assigned SKUs. Cost figures appended to waste-category recommendations
  (Disabled Account, Dormant, No Activity, Shelfware, Shared Mailbox, Overlapping).
- **Cost analysis in summary** - Total spend, per-category waste breakdown with EUR amounts,
  right-sizing potential, and cost tables by Department/Country/Company (top 15 each).
- **Last License Change** - Extracted from `licenseAssignmentStates.lastUpdatedDateTime`
  (beta API), tracks the most recent license modification date per user.
- **Department, Company, Country** - Added to user fetch (`CompanyName`, `Country`) and
  included as columns in the main report for cost-center analysis.
- **Recommendation Category** - New column classifying each user's primary recommendation
  into categories (Dormant, Disabled Account, Shelfware, OK, etc.) for pivot table use.
- **SKU Inventory pricing** - `Monthly Unit Price (EUR)` and `Annual Total Cost (EUR)`
  columns added to the SKU inventory export.

#### Excel Workbook Output (requires ImportExcel module)

Single `.xlsx` file with 9 worksheets:

| # | Worksheet | Content |
|---|-----------|---------|
| 1 | Dashboard | Key metrics, waste breakdown, recommendation distribution pie chart |
| 2 | User Report | Full per-user data with conditional formatting (data bars on cost, RAG on intensity, red on dormant, yellow on non-OK recommendations) |
| 3 | Service Plans | Granular SKU/service plan per user |
| 4 | SKU Inventory | Tenant license inventory with pricing and expiry highlighting |
| 5 | Cost by Department | Department-level cost aggregation with bar chart |
| 6 | Cost by Country | Country-level cost aggregation with bar chart |
| 7 | Cost by Company | Company-level cost aggregation |
| 8 | Recommendations | Category-level summary with costs and pie chart |
| 9 | Intensity Analysis | Exchange vs Teams intensity cross-tab |

- **Optional dependency**: `ImportExcel` module. If not installed, CSVs are still produced.
  Use `-NoExcel` to skip Excel output even when the module is available.

#### New Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-PricingCsvPath` | (none) | CSV with custom EUR prices per SKU |
| `-NoExcel` | Off | Skip Excel workbook generation |

#### Other Changes

- Step count increased from 11 to 12 (Excel export is step 12).
- 5 output files when ImportExcel is available (CSV + XLSX), 4 files without.

---

## v0.0.2 - 18/02/2026

### Big-Tenant Performance Overhaul

Eliminated the per-user `Get-MgUserLicenseDetail` API call loop that made the script
unusable on large tenants. A 100k-user tenant previously took ~28 hours; now completes
in ~15 minutes.

#### Changes

- **Bulk license resolution** - Replaced N individual `Get-MgUserLicenseDetail` calls
  with in-memory mapping using `Get-MgUser -Property AssignedLicenses` (bulk) +
  `Get-MgSubscribedSku` (already fetched). Service plan status is derived from each
  user's `DisabledPlans` list against the tenant's SKU service plan definitions.
- **Merged beta API loops** - Combined the sign-in activity and license assignment
  states queries into a single paginated beta endpoint call, halving the number of
  Graph API pages fetched.
- **Subscription pagination** - `/directory/subscriptions` now follows `@odata.nextLink`
  for tenants with many subscriptions.
- **Strict-mode fixes** - Fixed `$resp.value` -> `$resp['value']` and
  `$s.assignedByGroup` -> `$s['assignedByGroup']` for Graph API hashtable responses.
- **Numeric parsing fix** - Replaced broken `-replace '[^\d]','0'` pattern (which
  turned `"1,234"` into `"10234"`) with `Parse-NumericField` / `Parse-DoubleField`
  helper functions that correctly strip non-numeric characters.
- **Step count** - Reduced from 12 to 11 steps (merged loops).

#### Trade-offs

- Service plan `ProvisioningStatus` in the detail CSV now shows `Success` / `Disabled`
  only. Rare transient states (`PendingActivation`, `PendingProvisioning`) appear as
  `Success`. These resolve within hours and the main optimization report does not use
  service plan status.

---

## v0.0.1 - Initial Release

### Features

- 11 Graph Reports API endpoints for usage data
- Sign-in activity via beta API
- License assignment states (direct vs group-based, overlap detection)
- Subscription lifecycle (expiry dates, status)
- Admin role assignments
- Mailbox type detection (User/Shared/Room/Equipment) via EXO
- 10 optimization checks: disabled accounts, overlapping assignments, duplicate suite
  coverage, E3-to-E5 consolidation, shelfware detection, Teams Phone without calling
  plan, Copilot adoption monitoring, Power BI Pro with Premium Capacity, EXO Plan 2
  downgrade, frontline right-sizing
- Security and compliance upsell recommendations
- Business 300-seat limit warnings
- 4 output files: main CSV, service plan detail CSV, SKU inventory CSV, summary TXT
