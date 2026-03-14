# M365 License Optimization Report

A PowerShell script that audits Microsoft 365 license usage across a tenant and produces
actionable optimization recommendations. Designed for consultants and IT admins running
license reviews on tenants from 100 to 100,000+ users.

## What It Does

Pulls data from 11 Graph usage-report endpoints, sign-in activity, license assignments,
mailbox types, and admin roles — merges everything by UPN and generates per-user
recommendations.

### Data Sources

| Source | API | Purpose |
|--------|-----|---------|
| Office 365 Active User Detail | Graph Reports v1.0 | Last-activity dates per service |
| M365 App User Detail | Graph Reports v1.0 | Desktop/web/mobile app usage per app |
| Office 365 Activations | Graph Reports v1.0 | Platform activation counts |
| Email Activity | Graph Reports v1.0 | Send/receive/read counts |
| Teams User Activity | Graph Reports v1.0 | Chat/call/meeting counts |
| OneDrive Activity | Graph Reports v1.0 | File view/modify/sync/share counts |
| SharePoint Activity | Graph Reports v1.0 | File and page activity |
| Mailbox Usage | Graph Reports v1.0 | Mailbox size, item count, quota |
| Email App Usage | Graph Reports v1.0 | Which email clients per user |
| OneDrive Usage | Graph Reports v1.0 | Per-user storage and file count |
| Teams Device Usage | Graph Reports v1.0 | Which platforms for Teams |
| Sign-in Activity | Graph Beta | Last interactive/non-interactive sign-in |
| License Assignment States | Graph Beta | Direct vs group-based assignment |
| Subscription Lifecycle | Graph v1.0 | Expiry dates and SKU status |
| Admin Roles | Graph v1.0 | Directory role assignments |
| Mailbox Types + Holds | Exchange Online | User/Shared/Room/Equipment, Litigation Hold, Archive Status, Auto-Expanding Archive |
| Assigned Licenses | Graph v1.0 | SKU IDs and disabled plans per user |
| Subscribed SKUs | Graph v1.0 | Tenant license inventory |

### Optimization Checks (52+ Scenarios)

#### Tier 1 — Pure Waste (remove license immediately)
| # | Check | Description |
|---|-------|-------------|
| 1 | **Disabled account licensed** | Sign-in blocked but license still assigned |
| 2 | **Dormant account** | No sign-in within the configurable lookback window |
| 3 | **No activity** | Signed in but zero usage across all M365 workloads |
| 4 | **Deleted/soft-deleted** | Account removed but license still consuming a seat |
| 5 | **Shared mailbox waste** | Shared/room/equipment mailbox with a paid user license |
| 6 | **Guest account waste** | External B2B guest (#EXT#) holding a paid license |
| 7 | **Non-human account waste** | Service accounts (svc-, app-, noreply@, etc.) on full user licenses |
| 8 | **Viral/exploratory cleanup** | Free/trial SKUs (Power BI Free, Teams Exploratory) cluttering inventory |
| 9 | **Overlapping assignments** | Same SKU assigned both directly and via group |
| 10 | **Duplicate suite coverage** | Standalone SKU already included in an assigned suite (2-hop alias resolution) |
| 11 | **E5 Data Hoarder** | Disabled account on Litigation Hold with expensive suite — license is NOT needed for the hold |
| 12 | **Inactive Hold** | Disabled account on Litigation Hold with cheaper license — same safe removal applies |
| 13 | **Inactive mailbox (free)** | Unlicensed mailbox on Litigation Hold — auto-converts to free Inactive Mailbox |
| 12 | **Background sync only** | Zero interactive activity but OneDrive syncing in background (ghost device) |

#### Tier 2 — Right-Sizing (downgrade SKU to save delta)
| # | Check | Description |
|---|-------|-------------|
| 13 | **Frontline candidate** | E3/E5 user who only uses web/mobile apps — downgrade to F3 |
| 14 | **Frontline (high confidence)** | Same as above + zero device activations (no PC/Mac footprint) |
| 15 | **Frontline blocked** | Qualifies on activity but has Office on 2+ PCs (F3 VDI-only would break) |
| 16 | **F3→F1 micro-downgrade** | F3 user whose only activity is Teams mobile/web — F1 is ~50% cheaper |
| 17 | **Frontline add-on bloat** | F-series base + bolted-on add-ons exceeds Business Premium or E3 price |
| 18 | **Business Basic candidate** | Desktop suite user who only uses web apps — downgrade to Business Basic |
| 19 | **EXO Plan 2 downgrade** | Mailbox under 50 GB on Plan 2 — Plan 1 may suffice |
| 20 | **Suite Inversion** | E3 + E5-included add-on(s) costs MORE than full E5 — mathematically cheaper to upgrade |
| 21 | **E3→E5 consolidation** | E3 with E5-included add-ons — may be cheaper as E5 (with shelfware caveat) |
| 21 | **Bundle consolidation** | O365 + EMS + Windows individually — consolidate to M365 E3/E5 bundle |
| 22 | **Teams unbundling** | Bundled suite user with zero Teams activity — switch to "Without Teams" SKU |
| 23 | **Standalone desktop app waste** | M365 Apps subscription but never uses desktop Office (web/mobile only) |
| 24 | **A la carte waste** | Exchange Kiosk + standalone M365 Apps costs more than Business Standard |
| 25 | **Redundant archive** | Shared mailbox with EXO Plan 2 + standalone EOA — Plan 2 already includes archives |
| 26 | **Bundle inefficiency** | Business Basic + Apps for Business costs more than Business Standard |
| 27 | **Seeded Visio overlap** | Visio Plan 1 alongside E1/E3/E5 — suite includes native Visio web app |
| 28 | **Premium add-on waste** | Desktop-tier Visio/Project but product-specific activation shows web/mobile only |
| 28 | **Teams Phone right-sizing** | Non-human account (shared/room/equipment) with full Teams Phone — use Shared Devices |

#### Activity & Behavioral Analysis
| # | Check | Description |
|---|-------|-------------|
| 29 | **Shelfware** | Visio, Project, Power BI Pro, Teams Phone, Teams Premium, Copilot — licensed but inactive |
| 30 | **Power BI Pro with Premium Capacity** | Downgrade to Free if only consuming, not publishing |
| 31 | **Teams Phone without calling plan** | Phone System assigned but no PSTN route configured |
| 32 | **Copilot adoption** | Licensed but inactive (with web Copilot Chat telemetry caveat) |
| 33 | **Expensive cold storage** | Zero activity but large mailbox (>10 GB) or OneDrive (>50 GB) |
| 31 | **MDM/MAM waste** | Intune/EMS entitlement but 100% web-only access — nothing to manage |
| 35 | **Intune Suite waste** | Intune Suite add-on with zero advanced feature usage |
| 36 | **Heavy external sharer** | >50% content shared externally — DLP review flag |
| 37 | **Over-licensed archive** | Zero interactive activity, mailbox-only value — cheaper archive license exists |

#### Administrative & Compliance
| # | Check | Description |
|---|-------|-------------|
| 38 | **Dormant admin risk** | Admin account with no sign-in (interactive or non-interactive) |
| 39 | **Automation account** | Admin with non-interactive sign-in only — service/automation, not truly dormant |
| 40 | **Legacy service account** | POP3/IMAP4/SMTP-only access on premium suite |
| 41 | **Licensing error** | Group-based licensing failure (CountViolation, MutuallyExclusive, etc.) |
| 42 | **Litigation hold (shared mbx)** | Shared mailbox under Litigation Hold — license NOT needed, safe to remove |

#### Security & Compliance Coverage
| # | Check | Description |
|---|-------|-------------|
| 43 | **Security gap / Defender upsell** | Granular coverage analysis (MdoP1/P2, MdeP1/P2, Mdi, MdcApps, Xdr) |
| 44 | **Compliance gap / Purview upsell** | DLP Email+Files, DLP Teams, DLP Endpoint coverage |
| 45 | **PIM/CA licensing check** | PIM-eligible roles or risk-based CA policies without Entra P2 |
| 46 | **MDO policy licensing check** | In scope of Safe Links/Attachments rules without MDO license |
| 47 | **Entra Suite overlap** | Entra P2 + Entra Governance individually — consolidate to Entra Suite |
| 48 | **AI add-on overlap** | Teams Premium + Copilot + 0 meetings organized — Premium definitively redundant |
| 49 | **AI overlap review** | Teams Premium + Copilot + active organizer — verify webinar feature need |

## Requirements

### PowerShell Modules

| Module | Required | Purpose |
|--------|----------|---------|
| Microsoft.Graph.Authentication | Yes | Graph connection |
| Microsoft.Graph.Users | Yes | User and license data |
| Microsoft.Graph.Reports | Yes | Usage report downloads |
| Microsoft.Graph.Identity.DirectoryManagement | Yes | SKU inventory, admin roles |
| ExchangeOnlineManagement | Optional | Mailbox type detection |
| ImportExcel | Optional | Excel workbook output (.xlsx) |

Missing required modules are installed automatically. If `ExchangeOnlineManagement` is
not present, mailbox type detection is skipped gracefully. If `ImportExcel` is not present,
Excel output is skipped (CSVs still produced).

### Graph API Permissions

| Permission | Required For |
|------------|-------------|
| User.Read.All | User properties and assigned licenses |
| Reports.Read.All | All 11 usage reports |
| Organization.Read.All | Subscribed SKUs, subscription lifecycle |
| AuditLog.Read.All | Sign-in activity (beta) |
| RoleManagement.Read.Directory | Admin role assignments |
| Group.Read.All | Resolving license group names |
| Organization.ReadWrite.All | Only if using `-UnhideUserData` |

## Setup

### Option 1 — App Registration (recommended for production)

Run `LOA-App-Registration-Setup.ps1` to create a dedicated app registration with
certificate-based auth. This creates a `LOA-Connection.json` file that the main
script auto-detects — no parameters needed.

```powershell
# One-time setup (requires Global Admin)
.\App registration\LOA-App-Registration-Setup.ps1

# Then just run the report — connection config is auto-detected
.\Get-M365LicenseOptimizationReport.ps1
```

### Option 2 — Interactive login (quick ad-hoc runs)

```powershell
# Signs in interactively via browser — no setup required
.\Get-M365LicenseOptimizationReport.ps1
```

## Usage

```powershell
# Basic run (180-day lookback, output in current directory)
.\Get-M365LicenseOptimizationReport.ps1

# Certificate auth with explicit parameters
.\Get-M365LicenseOptimizationReport.ps1 -ClientId "xxx" -TenantId "yyy" -CertificateThumbprint "zzz"

# 90-day lookback, custom output folder
.\Get-M365LicenseOptimizationReport.ps1 -ReportPeriod D90 -OutputFolder "C:\Reports"

# Unhide hashed UPNs in usage reports (requires Organization.ReadWrite.All)
.\Get-M365LicenseOptimizationReport.ps1 -UnhideUserData

# Run without Exchange Online (Graph-only mode)
.\Get-M365LicenseOptimizationReport.ps1 -SkipEXO

# Delta report comparing against a previous run
.\Get-M365LicenseOptimizationReport.ps1 -PriorReportPath "C:\Reports\M365_LicenseOptimization_20260201.csv"

# Custom pricing CSV (default: M365SkuPricing.csv alongside script)
.\Get-M365LicenseOptimizationReport.ps1 -PricingCsvPath "C:\Pricing\custom_prices.csv"

# Reduce parallel downloads to avoid throttling on busy tenants
.\Get-M365LicenseOptimizationReport.ps1 -MaxParallel 2

# Include disabled+unlicensed accounts, skip Excel output
.\Get-M365LicenseOptimizationReport.ps1 -IncludeDisabledAccounts -NoExcel
```

### Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-ReportPeriod` | D180 | Lookback window: D7, D30, D90, D180 |
| `-OutputFolder` | Current dir | Folder for output files (auto-created if missing) |
| `-IncludeDisabledAccounts` | Off | Include disabled+unlicensed accounts in output |
| `-UnhideUserData` | Off | Temporarily unhide UPNs in reports (supports `-WhatIf`) |
| `-SkipEXO` | Off | Run without Exchange Online (skips mailbox type/litigation hold/MDO) |
| `-AutoInstallModules` | Off | Auto-install missing required modules (safe default: stops with instructions) |
| `-PriorReportPath` | (none) | Previous run's main CSV for delta analysis |
| `-SkuDataPath` | `M365SkuData.json` | External SKU names + prices JSON file |
| `-PricingCsvPath` | `M365SkuPricing.csv` | CSV with `SkuPartNumber,MonthlyPriceEUR` columns (default file ships with script) |
| `-MaxParallel` | 4 | Max concurrent Graph API report downloads (1-11) |
| `-RulePackPath` | `docs/LOA_RulePack_M365.json` | LOA rule pack with manual audit checklist rules and doc refs |
| `-NoExcel` | Off | Skip Excel workbook even if ImportExcel is installed |
| `-ClientId` | (auto) | App registration client ID (auto-detected from LOA-Connection.json) |
| `-TenantId` | (auto) | Tenant ID for certificate auth |
| `-CertificateThumbprint` | (auto) | Certificate thumbprint for app-only auth |
| `-CertificatePath` | (none) | Path to .pfx file (alternative to thumbprint) |
| `-CertificatePassword` | (none) | SecureString password for .pfx file |
| `-ExchangeHighThreshold` | 500 | Emails above this = High intensity |
| `-ExchangeLowThreshold` | 50 | Emails below this = Low intensity |
| `-TeamsHighThreshold` | 200 | Teams actions above this = High |
| `-TeamsLowThreshold` | 20 | Teams actions below this = Low |
| `-OneDriveHighThreshold` | 100 | OneDrive actions above this = High |
| `-OneDriveLowThreshold` | 10 | OneDrive actions below this = Low |
| `-SharePointHighThreshold` | 100 | SharePoint actions above this = High |
| `-SharePointLowThreshold` | 10 | SharePoint actions below this = Low |
| `-InactiveSignInDays` | 90 | Days without sign-in to flag as dormant (1-365) |

## Output Files

The script generates up to 7 files with a timestamp suffix:

| # | File | Description |
|---|------|-------------|
| 1 | **M365_LicenseOptimization_{ts}.csv** | Main per-user report: usage data, platform flags, intensity scores, capability levels, cost, and recommendation text |
| 2 | **M365_ServicePlanDetail_{ts}.csv** | Granular SKU and service plan breakdown per user with provisioning status |
| 3 | **M365_SkuInventory_{ts}.csv** | Tenant-level license inventory with friendly names, consumed/available counts, pricing, subscription status, and expiry dates |
| 4 | **M365_OptimizationSummary_{ts}.txt** | Human-readable summary: executive summary with tiered savings model, cost analysis, recommendation distribution, data collection warnings, manual audit checklist (from LOA rule pack) |
| 5 | **M365_ExecutiveSummary_{ts}.csv** | Tier 1 (pure waste) + Tier 2 (right-sizing) savings breakdown for executive reporting |
| 6 | **M365_LicenseOptimization_{ts}.xlsx** | *(if ImportExcel installed)* Excel workbook with 11 worksheets (see below) |
| 7 | **M365_LicenseDelta_{ts}.csv** | *(if `-PriorReportPath` provided)* Delta analysis: user changes, cost trends, recommendation shifts, dormancy/Copilot adoption tracking |

#### Excel Worksheets

| # | Sheet | Content |
|---|-------|---------|
| 1 | Executive Summary | Tiered savings model, key metrics |
| 2 | Dashboard | Waste breakdown, recommendation pie chart, delta metrics (if prior report) |
| 3 | User Report | Full per-user data with conditional formatting |
| 4 | Service Plans | Granular SKU/service plan per user |
| 5 | SKU Inventory | Tenant license inventory with pricing and expiry highlighting |
| 6 | Group Licensing | Entra ID groups with assigned licenses, member counts, disabled plans |
| 7 | Cost by Department | Department-level cost aggregation with bar chart |
| 8 | Cost by Country | Country-level cost aggregation with bar chart |
| 9 | Cost by Company | Company-level cost aggregation |
| 10 | Recommendations | Category-level summary with costs and pie chart |
| 11 | Delta Analysis | *(if `-PriorReportPath`)* User changes with conditional formatting |

## Performance

| Tenant Size | Approximate Runtime |
|-------------|-------------------|
| 1,000 users | ~2 minutes |
| 10,000 users | ~5 minutes |
| 100,000 users | ~15 minutes |

Runtime is dominated by Graph report downloads and Exchange Online queries. License
processing is done entirely in-memory with no per-user API calls.

## Updating SKU Data

`M365SkuData.json` contains SKU friendly names mapped from Microsoft's licensing reference.
The main script warns if the file is older than 90 days. To refresh:

1. Download the latest CSV from [Microsoft's licensing reference](https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference)
2. Save as `ms_licensing_reference.csv` in the script directory
3. Run `_extract_sku_names.ps1`

```powershell
.\_extract_sku_names.ps1
# Or specify a custom CSV path:
.\_extract_sku_names.ps1 -CsvPath "C:\Downloads\licensing_reference.csv"
```

## Project Structure

```
Get-M365LicenseOptimizationReport.ps1    # Main report script
M365SkuData.json                         # SKU friendly names + prices
_extract_sku_names.ps1                   # Regenerates M365SkuData.json from MS CSV
README.md                                # This file
App registration\
  LOA-App-Registration-Setup.ps1         # App registration + certificate setup
```

## Notes

- Usage data has ~48 hour reporting latency from Microsoft
- If UPNs appear as hashes, re-run with `-UnhideUserData`
- The `-UnhideUserData` flag temporarily changes a tenant-wide setting and restores it
  after the report completes
- Sign-in activity uses the beta Graph API and requires `AuditLog.Read.All`
- Mailbox type detection requires `ExchangeOnlineManagement` module (optional)
- The script uses `Set-StrictMode -Version Latest` for reliability

## Version

Current: **v0.3.3**
