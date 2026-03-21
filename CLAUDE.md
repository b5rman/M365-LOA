<!-- AUTO-MANAGED: project-description -->
## Overview

**M365 License Optimization Audit (M365-LOA)** — A PowerShell-based tool that audits Microsoft 365 tenants, identifies license waste, and generates actionable optimization recommendations with an interactive HTML dashboard.

Key capabilities: 148+ detection scenarios, 11 Graph API report endpoints, Exchange Online integration, certificate-based auth, delta analysis, Excel + HTML output.

<!-- END AUTO-MANAGED -->

<!-- AUTO-MANAGED: build-commands -->
## Build & Development Commands

```bash
# Run the main audit (interactive auth)
pwsh -File Get-M365LicenseOptimizationReport.ps1

# Run with certificate auth
pwsh -File Get-M365LicenseOptimizationReport.ps1 -ClientId <id> -TenantId <id> -CertificateThumbprint <thumb>

# Generate the HTML heatmap dashboard from audit output
pwsh -File Get-M365LicenseHeatmap.ps1 -OutputFolder ./output

# Syntax check (PowerShell AST parser)
pwsh -File syntaxcheck.ps1

# Generate rule pack reference doc (Node.js)
node generate-rulepack-doc.js
```

No formal test suite — validation is manual via syntax checking and live tenant runs.

<!-- END AUTO-MANAGED -->

<!-- AUTO-MANAGED: architecture -->
## Architecture

```
├── Get-M365LicenseOptimizationReport.ps1  # Main audit script (~7970 lines)
│   ├── Auth (interactive / certificate)
│   ├── Graph API data collection (11 endpoints + sign-in + license detail)
│   ├── Exchange Online collection (mailbox type, litigation hold, MDO)
│   ├── Data merge & enrichment (intensity, usage, admin roles)
│   ├── Recommendation engine (148+ checks via rule pack + inline logic)
│   └── Output: CSV, Excel workbook (15 tabs), executive summary
│
├── Get-M365LicenseHeatmap.ps1             # HTML dashboard generator (~1620 lines)
│   ├── Reads CSV output from main script
│   └── Produces standalone HTML with 6 tabs:
│       Overview, By Category, By User, By SKU, Workload Matrix, License Groups
│
├── LOA_RulePack_M365.json                 # Externalized detection rules (~4220 lines)
├── M365SkuData.json                       # SKU metadata: friendly names, included plans
├── M365SkuPricing.csv                     # Vendor CSP pricing (yearly/12 = monthly EUR)
│
├── App registration/                      # Customer-facing setup scripts (subtree)
├── output/                                # Generated reports land here
└── README.md                              # User-facing documentation
```

**Data flow**: Graph API + EXO → merged per-user hashtable → recommendation engine → CSV/Excel → Heatmap HTML

<!-- END AUTO-MANAGED -->

<!-- AUTO-MANAGED: conventions -->
## Code Conventions

- **Language**: PowerShell 7+ with `Set-StrictMode -Version Latest` (main script) / `-Version 2` (heatmap)
- **Naming**: PascalCase for functions (`Get-M365LicenseHeatmap`), `$camelCase` for local variables, `$PascalCase` for parameters
- **Currency**: Always use `[decimal]` for money values; pricing sourced from vendor CSP yearly÷12
- **UPN handling**: Normalize to lowercase `.ToLower()` before any lookups
- **Collections**: Wrap in `@()` to guarantee array; use `[System.Collections.Generic.HashSet[string]]` for dedup
- **EXO objects**: Deserialized — use `for`/`foreach` loops only (no pipeline `.Where()`)
- **SKU names**: Never match on display name alone; use `SkuPartNumber` as primary key
- **Graph CSV booleans**: Parse `Yes`/`No` strings, not `True`/`False`
- **BOM**: Graph API CSVs have UTF-8 BOM — handle with `Get-Content -Encoding UTF8`
- **Auth**: Never disable WAM; use plain `Connect-MgGraph` / `Connect-ExchangeOnline`
- **Language tone**: Customer-facing, advisory — "Review" not "Risk", never recommend account deletion (only license removal)
- **Hashtable access in StrictMode**: Use `.ContainsKey()` before accessing; `.property` syntax on `[hashtable]` throws in StrictMode

<!-- END AUTO-MANAGED -->

<!-- AUTO-MANAGED: patterns -->
## Detected Patterns

- **Recommendation format**: `"CATEGORY: description | SAVINGS: €X.XX/yr"` or `"LICENSING CHECK: description | COST: €X.XX/yr"`
- **Tiered savings**: Tier 0 (full annual cost), Tier 1 (partial/review), with explicit `$tier0Categories` and `$tier1Savings` mappings
- **MDO dedup**: `$sharedMbxHandledMdo` flag prevents redundant MDO licensing checks on shared mailboxes
- **PIM P2 + CA P1 dedup**: P1 check suppressed when P2 already recommended (P2 is superset)
- **Compliance cost**: Separate from savings — `Get-EstimatedComplianceCost` extracts from LICENSING CHECK segments
- **Heatmap tiles**: Dynamic tile generation from category regex, with per-row matching and modal drill-down
- **Rule pack externalization**: Detection rules in JSON, loaded at runtime, merged with inline checks
- **Cloud PC recency guard**: `$cpcRecentConnection` — skip all Cloud PC recommendations if `$lkpCloudPcDaysSinceSignIn[$upn] -lt 14`; prevents false flags on actively-used Cloud PCs
- **FRONTLINE RESCUE vs FRONTLINE CANDIDATE**: FRONTLINE RESCUE = user has active archive mailbox, can downgrade but not to F3; FRONTLINE CANDIDATE = standard web/mobile-only downgrade path — distinct tags, distinct counters

<!-- END AUTO-MANAGED -->

<!-- AUTO-MANAGED: git-insights -->
## Git Insights

- Repository: `b5rman/M365-LOA` on GitHub
- Active development since v0.3.1, currently at v0.5.6
- Commit style: imperative mood, concise ("Fix X", "Add Y", "Update Z")
- Version bumps noted in commit messages when significant

<!-- END AUTO-MANAGED -->

<!-- MANUAL -->
## Workflow Rules

- **After every change**: Run `syntaxcheck.ps1` → prove the fix is correct → update README.md if needed → commit when asked
- **"Update repo" command**: Push all changes to `origin/main` + update README.md + bump version number
- **Evaluate before fixing**: When a "bug" is reported, critically assess whether it's actually a bug, an acknowledged design limitation, or a feature request. Check existing comments and logic before touching code.
- **No Unicode emojis** in summary output — causes encoding issues. Use `[WARN]` prefix instead.

## Domain Knowledge

- **MDO P1 in E3**: Microsoft Defender for Office 365 Plan 1 (ATP_ENTERPRISE) is now included in both O365 E3 and M365 E3 (2025+). A standalone ATP_ENTERPRISE alongside E3 is a duplicate, not an E5-level add-on. Only THREAT_INTELLIGENCE (MDO P2) counts toward E5 upgrade.
- **ADALLOM_STANDALONE ≠ ADALLOM_O365**: ADALLOM_STANDALONE = full Defender for Cloud Apps; ADALLOM_O365 = lighter "Office 365 Cloud App Security" — different products, never treat as the same.
- **Litigation Hold**: Removing a license from a mailbox on Litigation Hold does NOT purge data. Microsoft auto-creates a free Inactive Mailbox that retains all content indefinitely.
- **M365 Maps** (m365maps.com) is the source of truth for suite includes and plan capabilities. Microsoft's own licensing CSV has data quality issues.
- **Graph API permissions**: `ReportSettings.ReadWrite.All` controls UPN privacy toggle (not `Organization.ReadWrite.All`). Existing customer app registrations may need this added manually + admin consent.
- **Compliance costs ≠ savings**: LICENSING CHECK recommendations contain amounts that must be spent. The heatmap `$costCategories` HashSet returns €0 savings for these — mixing them inflates totals (~26% on Xerius).
- **SKU pricing**: Vendor CSP annual commitment ÷ 12, ex-VAT EUR. Source: xSP pricelist XLSX. Always use yearly entry, never monthly commitment.
- **LOA Rule Pack**: AUTO rules are declarative documentation only — detection logic stays in native PowerShell for performance. Regenerate `LOA_RulePack_M365_Reference.docx` after rule text changes (`node generate-rulepack-doc.js`).

## Customer-Facing Language Rules

- **Never recommend account deletion** — only license removal. Account lifecycle is out of scope.
- **Advisory tone only**: "Consider removing" / "Review whether" / "The license can be safely removed" — never "Remove license" / "Disable immediately" / "Reclaim".
- **Labeling**: "Review" not "Risk", "Quick Wins" not "Pure Waste", "Right-Sizing Opportunities" not just "Right-Sizing".
- **Exclude-from-policy alternative**: All CA P1, CA P2, and MDO licensing check recommendations must include: "Alternatively, exclude this user/mailbox from the policy to avoid the compliance cost."

## JSON Externalization (M365SkuData.json)

| JSON Section | Script Variable | Type |
|---|---|---|
| `skuFriendlyNames` (648) | `$skuFriendlyNames` | hashtable |
| `suiteIncludes` (59) | `$suiteIncludes` | hashtable of string arrays |
| `planCapabilities` (23) | `$planCapabilities` | hashtable of hashtables |
| `planCapabilityAliases` (44) | `$planCapabilityAliases` | hashtable |
| `addOnBundles` (13) | `$addOnBundles` | HashSet[string] |
| `premiumSuites` (25) | `$premiumSuites` | string array |
| `skuCoverageAliases` (9) | `$skuCoverageAliases` | hashtable |
| `skuMonthlyPricesEUR` | `$skuMonthlyPrices` | hashtable (decimal) |

Script has minimal inline fallbacks (~3-6 entries each) so it runs with warnings if JSON is missing.

## Key Variables

- `$userSkuList` = subscription-level SKU part numbers (e.g., "SPE_E3") — use for SKU-level checks
- `$effectiveSkuSet` = service-plan-level enabled plans (e.g., "EXCHANGESTANDARD") — use for capability checks
- `$consolidatedPlans` = Dictionary[string,PSCustomObject] keyed on "upn|sku" — for Excel Service Plans tab
- `$recCategory` = primary recommendation category (first match in elseif chain — ORDER MATTERS)

## Key Helper Functions

- `Invoke-GraphWithRetry` — Exponential backoff for Graph API (429/5xx), all paginated calls
- `Import-CsvStripBom` — Strips UTF-8 BOM from Graph CSV downloads
- `Get-SkuMonthlyPrice` — Returns `[decimal]`; single gateway for all pricing
- `Resolve-SkuFriendlyName` — SKU part number → display name via `$skuFriendlyNames`
- `Get-PlanCapabilities` — SKU → capability profile (with alias fallback)
- `Merge-UserCapabilities` — OR-merges capabilities across all user SKUs
- `Get-EstimatedSavings` / `Get-EstimatedComplianceCost` — Heatmap savings/cost extraction

## Adding New Recommendations (5 touchpoints)

1. Counter variable declaration (~line 3020+)
2. Recommendation logic + `$recommendations.Add("TAG — message")` in per-user loop
3. `$recCategory` mapping (~line 5190+) — order matters, first match wins
4. `$recConfidence` level (~line 5290+) — High/Medium/Review
5. Counter increment (~line 5500+) — use `$recCategory` for accurate counts, NOT pattern matching on `$rec`

**Counter accuracy rule**: A user can have multiple recommendations but only ONE primary category (first elseif match). Pattern matching on `$rec` inflates counts when secondary recommendations exist.

## Adding New Cost Accumulators

- Declare as `[decimal]$xxxAcc = 0` (~line 3040+)
- Increment inline: `$xxxAcc += $calculatedAmount` inside per-user loop
- Finalize after loop: `$xxxFinal = [math]::Round($xxxAcc, 2)` (~line 5700+)
- If Executive Summary item: update tier formula, text block, CSV rows, Excel data arrays

## Executive Summary Tier Model

- `$totalIdentifiedWaste` = dormant+disabled+noActivity+shelfware+copilotReclaim+sharedMbx (NEVER modify — backward compat)
- `$tier1Waste` = $totalIdentifiedWaste + $duplicateCost (Quick Wins: remove license)
- `$tier2Savings` = frontline+businessBasic+exoPlan2+e5Upgrade+bundleConsolidation (Right-Sizing Opportunities)
- `$totalMoneyOnTable` = $tier1Waste + $tier2Savings

## Duplicate Detection Engine

- `$suiteIncludes` maps parent SKU → component service plans (59 suites in JSON)
- Pass 1: standalone SKU matches parent's component (with 2-hop alias via `$skuCoverageAliases`)
- Pass 2: child suite fully covered by parent suite
- `$suiteValidationSkipRx` suppresses warnings for non-suite products (Cloud PC, Dynamics, Copilot, Viva, etc.)

## MDO Evaluation

- All MDO rule evaluation uses `for` loops with index access (`$rules[$ri]`) — NEVER `foreach`/`ForEach-Object` on deserialized EXO objects
- `Get-RulePropArray` converts deserialized property values to `string[]` explicitly
- Per-rule try/catch so one bad rule doesn't fail the entire section
- BuiltInProtection coverage uses `$lkpMailboxPrimarySmtp.Values` (not `$allMailboxes` which is nulled for GC)
- No SKU gate — MDO policies are always scanned when EXO is connected

## PIM Role Detection

- `roleEligibilityScheduleInstances` → true PIM-eligible roles (user must activate)
- `roleAssignmentScheduleInstances` → filtered to `assignmentType = "Activated"` only; permanent/direct (`"Assigned"`) skipped to avoid false P2 licensing flags

## Graph API Column Name Gotchas

- Teams Device Usage: column is `'Used Android Phone'` NOT `'Used Android'`
- OneDrive Usage: uses `'Owner Principal Name'` instead of `'User Principal Name'`
- SharePoint: `'Modified File Count'` does NOT exist — use `'Viewed Or Edited File Count'`
- Safe column access: `if ('ColName' -in $row.PSObject.Properties.Name)` before accessing
- SharePoint file activity serves as proxy for Visio usage (no direct Visio report in Graph API)

## SKU Name Gotchas

- Visio Plan 1 web-only: `VISIOONLINE_PLAN1` (not `VISIOWEB`)
- Visio Plan 2 desktop: `VISIOCLIENT` or `VISIO_PLAN2_DEPT` (departmental)
- Project Plan 3: `PROJECTPROFESSIONAL`, Project Plan 5: `PROJECTPREMIUM`
- Common wrong names: `M365_ADVANCED_SECURITY` (actual: `IDENTITY_THREAT_PROTECTION`), `M365_ADVANCED_COMPLIANCE` (actual: `INFORMATION_PROTECTION_COMPLIANCE`)

## Coding Gotchas

- **Culture-invariant dates**: Always use `[System.Globalization.CultureInfo]::InvariantCulture` as third arg to `ParseExact`
- **Hot-loop performance**: Use `HashSet.Contains()` with `foreach`, not `Where-Object` pipelines, for per-user × 150K iterations
- **`$yrMatches`** not `$matches` — avoids PS automatic variable conflict in heatmap regex
- **Single-quoted here-string** `@'...'@` for static JS in heatmap — prevents PS expanding `${r}`, `${g}`, `${b}` template literals
- **`\u20AC`** for € in PS regex (avoids UTF-8 encoding issues)
- **`\u2014`** for em-dash in RecKey patterns (works in both .NET and JS regex engines)
- **`To-JsonString`** uses `ConvertTo-Json -InputObject` (not pipeline) to prevent single-element array unwrapping
- **`formatRec()` JS regex**: label pattern must be `[A-Za-z][A-Za-z0-9 /\-]+?` not `[A-Z][A-Z0-9 ...]` — lowercase-prefixed labels (e.g. "mobile apps only") are silently skipped by all-uppercase character class
- **`renderUserTable()` maxSav**: use `Math.max(...data.map(u => u.Savings||0), 1)` not `data[0].Savings` — sort order after filtering is not guaranteed, `data[0]` may not be the highest-savings user

## Heatmap Tile Pipeline (Critical Design)

- Tile count AND savings computed from **per-row matching only** (CatKey/RecKey regex against per-user CSV rows)
- Executive Summary CSV used ONLY for header KPIs (Total Spend, Savings Potential), NOT for tile numbers
- This guarantees tile count = drill-down count (what tile says = what you see when clicked)
- Users can appear in multiple tiles (tiles = finding lenses, not exclusive buckets)
- Savings: teal bubbles (`savings-cell`) in tables, green text in modals
- Compliance cost: peach bubbles (`compcost-cell`) in tables, plain peach text in drill-down modals
- Separate sort state for Workload Usage Matrix (`capSortCol`/`capSortAsc`) to avoid conflicts with tab 2
- **CatKey precision**: "Copilot At Risk" tile uses `copilot.watchlist` (not `at.risk|copilot.*risk`); Power BI add-ons require `power.bi` in CatKey alongside `pbi` — over-broad patterns cause wrong-category matches
- **Modal back-state**: `showUserDetail()` must delete both `backTile` AND `backSku` from `mb.dataset`; `showCapUserModal()` must call `clearBackState()` before `showUserDetail()` — stale backSku causes wrong back-navigation from Workload Matrix tab

## Memory Optimization

- EXO mailbox processing nulls `$allMailboxes[$i]` after building lookups to allow GC
- `[System.GC]::Collect()` called after the loop
- MDO BuiltInProtection reads from `$lkpMailboxPrimarySmtp.Values` instead (persistent data)
- Category tab generation: `$allCsvRows` freed after tab creation
- Runspace pool: wrapped in try/finally for guaranteed disposal

## Operational Notes

- **Xerius** (`E:\Xerius\`) is a real customer tenant — always sync changes to both locations when fixing main scripts.
- **Shared Mailbox decision tree**: Litigation Hold → safe to remove (Inactive Mailbox) → Size unknown → review → Over 50 GB → needs EXO Plan 2 → Active archive → retain license → MDO scope without entitlement → needs license → Under 50 GB, no holds/MDO/archive → license not needed.
- **EXO token expiry**: Graph API calls (steps 1–6) can take minutes on large tenants, expiring the EXO token before step 7. The labeled retry loop `exoMbxRetry` handles this automatically.
- **SKU seat counts**: `Total = Enabled + Warning` (Warning seats are usable). `Available = (Enabled + Warning) - ConsumedUnits`. Pool waste detection has no threshold — all paid SKUs with unassigned seats are included.

<!-- END MANUAL -->
