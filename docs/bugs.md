# Bugs — Tracking

> Bug reports and structural improvements. Move to "Implemented" section once done.

## Open

(No open items.)

## Implemented

8. **E5 Step-Up / Suite Inversion** — E5 upgrade check required ≥2 paid add-ons, masking suite inversions where E3 + 1 heavy add-on already exceeds E5 price. Fixed: lowered threshold from ≥2 to ≥1, added `SUITE INVERSION` tag for delta > 0 (mathematically cheaper to upgrade). Existing `E5 CONSOLIDATION` retained for delta ≤ 0. Tag: `SUITE INVERSION`. Confidence: High. v0.3.1.

9. **Teams Premium ↔ Copilot AI Overlap** — `AI OVERLAP REVIEW` was blanket-applied without checking meeting organizer activity. Fixed: split into two tiers using `$teamsMeetingsOrganized`. 0 meetings → `AI ADD-ON OVERLAP` (definitive removal, High confidence). >0 meetings → `AI OVERLAP REVIEW` (manual check, Review confidence). v0.3.1.

10. **E5 Data Hoarder / Inactive Hold (Factual Error)** — Disabled+held account message stated "License MUST be retained to maintain hold" — factually wrong. Microsoft Inactive Mailboxes retain ALL content and holds indefinitely without a license. Fixed: two-tier detection: expensive suite → `E5 DATA HOARDER`, cheaper license → `INACTIVE HOLD`, both with correct messaging. Also fixed same error in shared mailbox hold message. Fixed `$disabledCostAcc` regex to include new tags. v0.3.1.

11. **Visio Plan 1 Seeded Redundancy** — Microsoft includes "Visio in M365" web app natively in E1/E3/E5. Users with standalone Visio Plan 1 (€4.70/mo) alongside qualifying suite don't need it for viewing/light editing. Fixed: added `SEEDED VISIO OVERLAP` detection using SharePoint file activity < 5 as proxy for light usage. Tag: `SEEDED VISIO OVERLAP`. Confidence: Medium. v0.3.1.

12. **EU/Global Teams Unbundling** — Already implemented in v0.2.0 as `TEAMS UNBUNDLING` detection. No additional changes needed.

## Previously Implemented

1. **A-La-Carte Bundle Consolidation** — Users paying for Office 365 E3 + EMS E3 + Windows E3 individually instead of a unified M365 E3 bundle. Fixed: added `BUNDLE CONSOLIDATION` detection that checks for the three pillars (O365 + EMS + Windows) at both E3 and E5 tiers, plus Entra Suite overlap (Entra P2 + Entra Governance → Entra Suite). Tag: `BUNDLE CONSOLIDATION`. Confidence: High.

2. **Global Teams Unbundling** — Users on legacy bundled suites (M365 E3/E5, O365 E1/E3, Business Basic/Standard/Premium) with zero Teams activity are paying the "Teams tax." Microsoft now offers cheaper "Without Teams" SKU equivalents. Fixed: added `TEAMS UNBUNDLING` detection that flags users with a bundled suite and `$teamsTotal -eq 0`. Tag: `TEAMS UNBUNDLING`. Confidence: Medium.

3. **B2B Guest User Over-Licensing** — External guest accounts (#EXT#) accidentally assigned paid M365 licenses. Guests don't need a license in the host tenant for Teams chat, meetings, or SharePoint co-authoring. Fixed: added `GUEST ACCOUNT WASTE` detection checking `UserType -eq 'Guest'` or UPN containing `#EXT#` with `$userAnnualCost -gt 0`. Tag: `GUEST ACCOUNT WASTE`. Confidence: High.

4. **Floating-Point Finance Trap** — All currency values and cost accumulators used `[double]`, causing IEEE 754 floating-point drift when summing prices across thousands of users (e.g., €142,589.499999999 instead of €142,589.50). Fixed: switched `Get-SkuMonthlyPrice` to return `[decimal]`, all cost accumulators to `[decimal]` type constraints, JSON/CSV pricing loaders to `[decimal]` casts. v0.2.0.

5. **MDO Group Timeout Circuit Breaker** — `Get-GroupMemberUpns` could timeout on massive groups (100k+ "All Employees" dynamic groups) used in Defender for Office 365 policies, causing 100+ paginated API calls and heavy throttling. Fixed: added `$count` check before expansion; groups >10,000 members return `[ALL_TENANT]` marker that propagates through the MDO coverage chain to cover all users. v0.2.0.

6. **Multi-PC Gate Overcounting Windows Activations** — `$winActTotal` summed the 'Windows' column across ALL activation rows (Office, Visio, Project, etc.). A user with Office on 1 PC + Visio on 1 PC = winActTotal=2, incorrectly blocking the frontline downgrade. Fixed: filter activation rows to `Product Type -match 'Office|Microsoft 365 Apps|M365 Apps'` before summing. v0.2.1.

7. **Confidence Scoring Ignores Missing Data Sources** — Activity-based recommendations (NO ACTIVITY, SHELFWARE, Frontline Candidate, etc.) could receive High or Medium confidence even when key data sources (EmailActivity, TeamsActivity, OneDriveActivity, M365AppPlatform) were missing. Fixed: added post-hoc confidence downgrade — if recommendation is activity-based and `$missingDataSources` includes key sources, confidence is forced to "Review". v0.2.1.
