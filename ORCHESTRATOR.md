# ORCHESTRATOR.md — Living Context for M365-LOA

> This file is maintained by the orchestrator thread. It holds architecture, conventions, decisions, known risks, and current state so that context survives compaction and subagents can do good work.

## Repo Identity

- **Repo**: `b5rman/M365-LOA` on GitHub, branch `main`
- **Current Version**: v0.5.6 (Mapping v1.3, RecommendationLogic v1.2.1)
- **Purpose**: M365 License Optimization Audit — PowerShell tool that audits Microsoft 365 tenants, identifies license waste, and generates actionable recommendations with an interactive HTML dashboard.

## Architecture

### Core Production Files

| File | Lines | Role |
|------|-------|------|
| `Get-M365LicenseOptimizationReport.ps1` | 7,977 | Main audit engine |
| `Get-M365LicenseHeatmap.ps1` | 1,614 | HTML dashboard generator |
| `LOA_RulePack_M365.json` | 4,219 | 146 detection rules (declarative metadata; logic in PowerShell) |
| `M365SkuData.json` | 1,355 | 651 SKU names, 62 suite includes, 23 capabilities, 46 aliases |
| `M365SkuPricing.csv` | 198 rows | EUR monthly pricing (vendor CSP yearly÷12) |

### Data Flow

```
Graph API (11 endpoints) + EXO cmdlets
  → 30+ UPN-keyed lookup hashtables
    → per-user recommendation engine (148+ checks, 80+ counters)
      → streaming CSV + Excel workbook (15 tabs)
        → Heatmap HTML (6 tabs, embedded JS/CSS)
```

### Main Script Sections

| Section | Lines | Content |
|---------|-------|---------|
| Parameters & Metadata | 1–270 | 20+ params (auth, thresholds, paths) |
| Helper Functions (16) | 311–894 | Invoke-GraphWithRetry, Import-CsvStripBom, Get-SkuMonthlyPrice, etc. |
| Auth & Connection | 900–1058 | Graph + EXO, cert auth |
| Unhide UPNs | 1059–1091 | ReportSettings.ReadWrite.All toggle |
| Graph Downloads | 1092–1376 | 11 parallel reports (runspace pool) |
| Build Lookups | 1377–1455 | Index CSVs by UPN |
| SKU Inventory | 1456–2216 | License assignments, disabled plans, duplicate detection |
| MDO Coverage | 2217–2654 | Safe Links/Attachments/AntiPhish, group expansion |
| PIM & CA | 2655–3007 | Entra ID P2 signals |
| Counters & Accumulators | 3008–3420 | 80+ counters, [decimal] cost accumulators |
| Recommendation Engine | 3420–6218 | Per-user loop with 148+ checks |
| Category Mapping | 5751–5845 | 95 elseif conditions (ORDER MATTERS — first match wins) |
| Export (CSV/Excel/Summary) | 6229–6917 | Executive summary, workbook |
| Delta Report | 6919–7140 | Prior vs current comparison |

### Heatmap Script Structure

| Section | Lines | Content |
|---------|-------|---------|
| CSV loading & data processing | 47–525 | Parse users, compute per-SKU waste, tile definitions, KPIs |
| PowerShell functions | 102–530 | Parse-Decimal, Get-EstimatedSavings, Get-EstimatedComplianceCost, To-JsonString |
| CSS (dark theme) | 587–697 | JetBrains Mono + Sora fonts, navy/purple/teal/peach palette |
| HTML structure | 698–864 | KPI boxes, 6 tab panels |
| JS data injection | 866–874 | USERS, SKUS, TILES, CAP_USERS, POOL_SKUS, GROUPS, ALL_CATS |
| JS functions (25) | 876–1601 | Rendering, modals, sorting, formatting, back-navigation |
| File output | 1602–1614 | UTF8 write |

### Supporting Files

| File | Purpose |
|------|---------|
| `syntaxcheck.ps1` | AST parser — validates 4 scripts (ONLY validation mechanism) |
| `generate-rulepack-doc.js` (757 lines) | JSON → Word doc (npm `docx` package) |
| `_extract_sku_names.ps1` | Refresh M365SkuData.json from Microsoft CSV |
| `App registration/LOA-App-Registration-Setup.ps1` | Customer Azure AD app setup |
| `README.md` (447 lines) | User-facing documentation |

## Conventions

### PowerShell
- PowerShell 7+, `Set-StrictMode -Version Latest` (main) / `-Version 2` (heatmap)
- PascalCase functions, `$camelCase` locals, `$PascalCase` parameters
- `[decimal]` for all currency — never float/double
- `[System.Globalization.CultureInfo]::InvariantCulture` for date parsing
- UPN: `.ToLower()` before any lookup
- `@()` to guarantee array; `[System.Collections.Generic.HashSet[string]]` for dedup
- EXO objects: `for`/`foreach` only (no pipeline `.Where()` — deserialized objects)
- `.ContainsKey()` before hashtable access (StrictMode requirement)
- `$yrMatches` not `$matches` (avoid PS automatic variable conflict)
- Single-quoted here-string `@'...'@` for static JS (prevents PS expanding `${r}`, `${g}`, `${b}`)

### Naming & Tone
- Customer-facing: "Review" not "Risk", "Quick Wins" not "Pure Waste"
- Never recommend account deletion — only license removal
- Advisory: "Consider removing" / "Review whether" — never "Remove" / "Disable"
- All CA P1/P2 and MDO licensing checks include exclude-from-policy alternative

### SKU Rules
- Match on `SkuPartNumber`, never display name alone
- Graph CSV booleans: `Yes`/`No` strings, not `True`/`False`
- UTF-8 BOM in Graph CSVs — `Import-CsvStripBom` handles this
- `\u20AC` for € in regex; `\u2014` for em-dash in RecKey patterns

### Workflow
- After every change: `syntaxcheck.ps1` → prove fix → update README if needed
- Commit style: imperative mood ("Fix X", "Add Y")
- No Unicode emojis in output (encoding issues) — use `[WARN]` prefix

## Fragile Areas (Ranked by Impact)

1. **Category mapping order** (~line 5751): 95 elseif conditions — first match = primary category. Reordering breaks counter accuracy.
2. **Cost accumulators**: Must be `[decimal]` — IEEE 754 drift loses €10k+ on large tenants (150k users).
3. **Duplicate detection**: Multi-pass with 2-hop alias via `$skuCoverageAliases`. Alias chain order matters.
4. **MDO group expansion**: >10k members → `[ALL_TENANT]` marker mixed with per-SMTP entries in same dict.
5. **Cloud PC recency guard**: `$cpcRecentConnection` skips ALL CPC recs if sign-in < 14 days.
6. **Heatmap `formatRec()` regex**: Labels must start `[A-Za-z]` — lowercase-prefixed silently skipped.
7. **Modal back-navigation**: Stale `backTile`/`backSku` dataset → wrong back-nav. `clearBackState()` critical.
8. **Disabled plans intersection**: Must process ALL user SKUs before finalizing (ExceptWith).
9. **PIM P2 + CA P1 dedup**: P1 suppressed when P2 recommended — implicit in elseif order.
10. **Tier 1 savings categories**: 8 categories previously inflated by ~26% (fixed v0.5.3). Guard against regression.
11. **`$suiteValidationSkipRx`** (~line 1536): Excludes Cloud PC, Dynamics, Viva from suite validation. Adding new product families requires updating this regex.
12. **Delta report column access**: `$getSafe` lambda handles missing columns. Column name changes break silently.

## Key Helper Functions

| Function | Line | Purpose |
|----------|------|---------|
| `Invoke-GraphWithRetry` | 378 | Exponential backoff for Graph API (429/5xx), auto-pagination |
| `Import-CsvStripBom` | 464 | Strips UTF-8 BOM from Graph CSV |
| `Get-SkuMonthlyPrice` | 643 | Returns `[decimal]`; single pricing gateway |
| `Resolve-SkuFriendlyName` | 535 | SKU part number → display name |
| `Get-PlanCapabilities` | 836 | SKU → capability profile (with alias fallback) |
| `Merge-UserCapabilities` | 849 | OR-merge capabilities across all user SKUs |
| `Build-UPNLookup` | 1385 | Creates hashtable index from CSV rows |
| `Get-GroupMemberUpns` | 2261 | Expands group → member UPNs (10k circuit breaker) |
| `Parse-Decimal` | Heatmap:116 | EUR amount parsing (European + standard formats) |
| `Get-EstimatedSavings` | Heatmap:134 | Extract savings from recommendation text |
| `Get-EstimatedComplianceCost` | Heatmap:155 | Extract compliance cost from LICENSING CHECK segments |

## Adding New Recommendations (5 Touchpoints)

1. Counter variable declaration (~line 3020+)
2. Recommendation logic + `$recommendations.Add("TAG — message")` in per-user loop
3. `$recCategory` mapping (~line 5190+) — order matters, first match wins
4. `$recConfidence` level (~line 5290+) — High/Medium/Review
5. Counter increment (~line 5500+) — use `$recCategory`, NOT pattern matching on `$rec`

## Decisions Log

*(Updated as decisions are made in orchestrator thread)*

| Date | Decision | Rationale |
|------|----------|-----------|
| 2026-03-21 | Orchestrator thread established | Living memory for compaction resilience |

## Current State

- **Version**: v0.5.6
- **Last commit**: pending — Harden fragile areas: fix 25 bugs across 5 review cycles
- **No test suite**: Validation = `syntaxcheck.ps1` + live tenant runs
- **No test suite**: Validation = `syntaxcheck.ps1` + live tenant runs
- **Customer reference tenant**: Xerius (`E:\Xerius\`)
