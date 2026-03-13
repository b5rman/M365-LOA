# Ideas — Tracking

> Feature ideas to evaluate. Move to "Implemented" section once done.

## Open

(No open items — all evaluated and either implemented or deferred.)

### Deferred

3) **”Unknown vs Zero” sentinel for usage totals** — Superseded by the confidence scoring post-hoc downgrade implemented in v0.2.1 (Bug #3 fix). When key data sources are missing, activity-based recommendations are automatically downgraded to “Review” confidence. The UPN normalization fix (v0.3.0) also eliminates a major source of false “missing row” scenarios. Revisit only if false zeros persist after these fixes.

### Already Implemented (no change needed)

- **Respect Retry-After** — Already implemented via `Invoke-GraphWithRetry` which parses Retry-After headers.
- **End-of-run health summary** — Already implemented: `$script:skippedDataWarnings` tracks all missing data sources, displayed in the summary output and logged.
- **Defender policy scope** — Already implemented: MDO policy scope is fetched (4 policy types: Built-in, Preset, Safe Links, Safe Attachments), mapped per-UPN via `$lkpMdoCoverageByUpn`, exported as `MDO Policy Coverage` column, and compared against entitlements with `LICENSING CHECK` recommendations for unlicensed mailboxes in scope.
- **Mailbox type + hold flags** — Already implemented: `RecipientTypeDetails` (User/Shared/Room/Equipment), `LitigationHoldEnabled`, and `Has Archive` (from Graph report) all fetched and exported. v0.3.1 added `ArchiveStatus` and `AutoExpandingArchiveEnabled` from `Get-EXOMailbox`.

## Implemented

13. **LOA Rule Pack — Manual Audit Checklist** — New `-RulePackPath` parameter loads `docs/LOA_RulePack_M365.json` declarative rule pack. Filters MANUAL-level rules (portal checks the script cannot automate: Teams Rooms caps, Resource Accounts, Intune enrollment, Power BI sharing, audit retention) and appends a "Manual Audit Checklist" section to the summary TXT with step-by-step instructions and Microsoft Learn doc references. AUTO/SEMI rules remain informational — existing hardcoded detections are authoritative. v0.3.1.

12. **Archive Status & Auto-Expanding Archive Columns** — Extended `Get-EXOMailbox` fetch to include `ArchiveStatus` and `AutoExpandingArchiveEnabled` properties. Two new informational CSV columns: `Archive Status` (Active/None) and `Auto-Expanding Archive` (True/False). Enables future archive-based detection. v0.3.1.

10. **Case-Insensitive UPN Normalization** — All ~15 UPN-keyed lookup hashtables were case-sensitive plain hashtables. Graph API/CSV reports return inconsistent casing across pages (e.g., "User@Domain.com" vs "user@domain.com") and occasional trailing whitespace, causing silent data loss in every join. Fixed: all UPN keys normalized with `.ToString().Trim().ToLower()` at insertion time across Build-UPNLookup, user fetch, beta API fetch, EXO mailbox fetch, admin roles, PIM roles, MDO coverage, and activations. Also fixed a specific bug where `$lkpMailboxType`/`$lkpLitigationHold` computed `$mbxUpnLower` but inserted using raw `$mbxUpn`. v0.3.0.

11. **Multi-PC Gate Office-Only Filter** — Already implemented in v0.2.1 (Bug #6 in bugs.md). The `$winActTotal` computation now filters activation rows to `Product Type -match 'Office|Microsoft 365 Apps|M365 Apps'` before summing Windows counts.

1. **"Frankenstein Frontline" Worker** — Detects F-series users whose bolted-on add-ons (Exchange Plan 2, Entra P2, Power BI Pro, etc.) push their total cost above Business Premium or E3. Recommends consolidating to a full suite to save money and remove Frontline restrictions (2 GB mailbox, 10.9" screen cap). Tag: `FRONTLINE ADD-ON BLOAT`. Confidence: Medium.

2. **"Device-less" Web User (Zero Hardware Footprint)** — Enhanced the existing `FRONTLINE CANDIDATE` detection: when `$activatedPlatforms` is empty (zero device activations ever), the recommendation is prefixed with `(HIGH CONFIDENCE)`. Proves the user literally has no corporate PC/Mac — strongest possible signal for a web-only downgrade. Confidence: High.

3. **"Kiosk + Desktop Apps" Clash** — Exchange Kiosk (€1/mo, 2 GB mailbox) + M365 Apps for Enterprise (€13.90/mo) = €14.90/mo. M365 Business Standard (€12.50/mo) is cheaper AND upgrades mailbox to 50 GB + 1 TB OneDrive. Fixed: added `A LA CARTE WASTE` detection that computes combined a-la-carte cost vs Business Standard price. Tag: `A LA CARTE WASTE`. Confidence: High. v0.2.0.

4. **Over-Licensed Shared Mailbox (Archive Edition)** — Shared mailboxes with Exchange Plan 2 + standalone Exchange Online Archiving. Plan 2 natively includes auto-expanding archives, making EOA (€3/mo) 100% redundant. Fixed: added `REDUNDANT ARCHIVE` detection for shared mailboxes, plus `$suiteIncludes` mapping `"EXCHANGEENTERPRISE" = @("EXCHANGE_ARCHIVE")` so the duplicate engine catches it for all users. Tag: `REDUNDANT ARCHIVE`. Confidence: High. v0.2.0.

5. **"Basic + Apps" Frankenstein** — Business Basic (€6/mo) + Apps for Business (€11.50/mo) = €17.50/mo. Business Standard (€12.50/mo) includes both natively. Fixed: added `BUNDLE INEFFICIENCY` detection comparing combined a-la-carte cost vs Business Standard price. Handles all alias SKU variants (O365_BUSINESS_ESSENTIALS, SMB_BUSINESS_ESSENTIALS, M365_BUSINESS_BASIC + O365_BUSINESS, SMB_BUSINESS). Tag: `BUNDLE INEFFICIENCY`. Confidence: High. v0.2.1.

6. **Business Premium Security Redundancy** — Already handled by the existing duplicate detection engine. SPB's `$suiteIncludes` maps ATP_ENTERPRISE and INTUNE_A as components, so Pass 1 of the duplicate engine flags standalone add-ons of these when SPB is present. EMS E3 is correctly NOT flagged because SPB lacks RIGHTSMANAGEMENT (AIP P1). No code change needed — engine already covers this.

7. **Visio/Project "Plan 1" Downgrade** — Desktop-tier Visio (€14.60/mo) or Project (€29.30–€52.80/mo) assigned, but product-specific activation data shows NO Windows/Mac desktop activation — user only accesses via mobile/web. Fixed: added `PREMIUM ADD-ON WASTE` detection integrated into the existing shelfware block. Checks per-product activation rows for Windows/Mac counts; if both are zero, recommends downgrade to web-only Plan 1 equivalent (Visio Plan 1 €4.70/mo, Project Essentials €7/mo). Tag: `PREMIUM ADD-ON WASTE`. Confidence: Medium. v0.2.1.

8. **Power BI Pro vs. PPU Double-Licensing** — Already implemented at the PBI Pro overlap check section. Users with both `POWER_BI_PREMIUM_PER_USER` and `POWER_BI_PRO` are flagged with `POWER BI PRO REVIEW — PPU is a superset of Pro`. No code change needed.

9. **Teams Phone "Standard" vs "Shared Device" Waste** — Non-human accounts (Shared/Room/Equipment mailboxes) with full Teams Phone Standard (MCOEV, €8/mo) instead of cheaper Teams Shared Devices (MCOCAP, €2.50/mo). Fixed: added `TEAMS PHONE RIGHT-SIZING` detection for non-human accounts with standalone MCOEV. Added MCOCAP pricing (€2.50/mo) to built-in price table. Tag: `TEAMS PHONE RIGHT-SIZING`. Confidence: High. v0.2.1.
