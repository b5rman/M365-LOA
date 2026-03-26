# M365 License Optimization Report

A PowerShell script that audits Microsoft 365 license usage across a tenant and produces
actionable optimization assessments. Designed for consultants and IT admins running
license reviews on tenants from 100 to 100,000+ users.

## What It Does

Pulls data from 13 Graph usage-report endpoints (including Cloud PC), sign-in activity,
license assignments, mailbox types, and admin roles — merges everything by UPN and
generates per-user assessments.

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
| Mailbox Types + Holds | Exchange Online | User/Shared/Room/Equipment, Litigation Hold, Archive Status, Auto-Expanding Archive, Forwarding rules |
| Copilot Usage Detail | Graph Beta | Per-user Copilot activity across M365 apps |
| Cloud PC Provisioning | Graph v1.0 | Provisioned Windows 365 Cloud PCs per user |
| Cloud PC Usage Reports | Graph Beta | Remote connection hours over 90 days (aggregated) |
| Assigned Licenses | Graph v1.0 | SKU IDs and disabled plans per user |
| Subscribed SKUs | Graph v1.0 | Tenant license inventory |

### Optimization Checks (149 Scenarios + Compliance Cost Tracking)

#### Tier 0 — Unlicensed & Non-Human Accounts
| # | Check | Description |
|---|-------|-------------|
| 1 | **Shared mailbox (unlicensed)** | Shared mailbox under 50 GB — no license required |
| 2 | **Shared mailbox storage warning** | Unlicensed shared mailbox approaching 50 GB limit — will need a license if it exceeds the cap |
| 3 | **Room/Equipment mailbox** | No user license needed for room/equipment accounts |
| 4 | **Unlicensed user mailbox** | User mailbox with no license assigned — data will be purged after 30 days |
| 5 | **Inactive mailbox (free)** | Unlicensed mailbox on Litigation Hold — Microsoft auto-converts this to a free Inactive Mailbox that preserves all content |
| 6 | **Unlicensed data risk (OneDrive)** | Unlicensed user with OneDrive data — Microsoft will purge the data after 30 days without a license |
| 7 | **Unlicensed data risk (combined)** | Unlicensed user with mailbox and/or OneDrive data — Microsoft purges both after 30 days without a license |
| 8 | **MDO policy gap (unlicensed)** | Unlicensed mailbox in scope of Defender for Office 365 policies — needs a license for MDO coverage to apply |
| 9 | **Guest account waste** | External B2B guest user (#EXT#) holding a paid license — guests are covered by the 1:5 Entra ID member-to-guest ratio |
| 10 | **Guest user (free SKU)** | Guest user with a free license assigned — no financial impact, informational only |
| 11 | **Automation account (unlicensed)** | Unlicensed service/sync account (e.g. AD Connect, sync_*) — no license cost, consider converting to Workload Identity |
| 12 | **Dormant admin risk (unlicensed)** | Enabled unlicensed admin with no sign-in — security risk even without license cost |
| 13 | **PIM licensing check (unlicensed)** | Unlicensed user with Privileged Identity Management role assignments — requires Entra ID P2 license |
| 14 | **CA P1 licensing check (unlicensed, scoped)** | Unlicensed user targeted by a Conditional Access policy — requires Entra ID P1 license |
| 15 | **CA P1 licensing check (unlicensed, tenant-wide)** | Unlicensed user included in tenant-wide Conditional Access policies — requires Entra ID P1 license |

#### Tier 1 — Quick Wins
| # | Check | Description |
|---|-------|-------------|
| 16 | **E5 Data Hoarder** | Disabled account on Litigation Hold with expensive suite — license is NOT needed for the hold |
| 17 | **Inactive Hold** | Disabled account on Litigation Hold with cheaper license — same safe removal applies |
| 18 | **Disabled account review** | Disabled account still licensed — cannot verify Litigation Hold status without Exchange Online connection; check manually before removing |
| 19 | **Disabled shared mailbox (MDO)** | Disabled shared mailbox protected by Defender for Office 365 policies — downgrade to Exchange Plan 2 or MDO add-on instead of full removal |
| 20 | **Disabled account licensed** | Sign-in is blocked and no Litigation Hold detected — remove license to stop billing |
| 21 | **Disabled account (free SKU)** | Disabled account with a free license — no financial impact but consider removing for tenant hygiene |
| 22 | **Shared mailbox on Litigation Hold** | Shared mailbox on Litigation Hold — license is NOT needed to maintain the hold, safe to remove |
| 23 | **Shared mailbox review** | Shared mailbox size is unknown (usage report missing) — cannot safely recommend removal without verifying size |
| 24 | **Shared mailbox over 50 GB** | Shared mailbox exceeds the free 50 GB limit — requires Exchange Online Plan 2 or a suite license to support the larger mailbox |
| 25 | **Shared mailbox (MDO coverage)** | Shared mailbox under 50 GB but protected by Defender for Office 365 policies — downgrade to Exchange Plan 2 or MDO add-on instead of full removal |
| 26 | **Shared mailbox waste** | Shared mailbox under 50 GB with a paid user license — shared mailboxes under 50 GB do not require a license |
| 27 | **Room/Equipment mailbox waste** | Room/equipment with full user license instead of a Teams Rooms SKU — only needs Teams Rooms license |
| 28 | **Non-human account waste** | Service accounts (svc-, app-, noreply@, etc.) holding expensive user licenses — consider Workload Identity or remove |
| 29 | **Dormant account** | User has not signed in within the configurable lookback window (default 30 days) — license is likely wasted |
| 30 | **No activity** | User signed in but has zero usage across all M365 workloads (Exchange, Teams, OneDrive, SharePoint) — license not being used |
| 31 | **Background sync only** | Zero interactive activity but OneDrive syncing in background — likely an abandoned device, not a real user |
| 32 | **Forwarding mailbox waste** | Dormant user whose mailbox only auto-forwards to another address — a license is not needed just for forwarding |
| 33 | **Overlapping assignments** | Same license SKU assigned both directly and via group-based licensing — remove the direct assignment |
| 34 | **Duplicate suite coverage** | Standalone license already included in the user's suite — paying twice for the same capability |
| 35 | **Viral/exploratory cleanup** | Self-service free/trial licenses (Power BI Free, Teams Exploratory, etc.) alongside a paid suite — remove to clean up inventory |

#### Tier 2 — Right-Sizing Opportunities
| # | Check | Description |
|---|-------|-------------|
| 36 | **Suite Inversion** | E3 plus two or more E5-included add-ons costs more than a full E5 license — upgrading to E5 saves money |
| 37 | **E3→E5 consolidation** | E3 with E5-included add-ons at similar cost to E5 — consolidating to E5 simplifies management (verify unused E5 features are acceptable) |
| 38 | **Bundle consolidation (E3)** | Office 365 E3 + EMS E3 + Windows E3 purchased separately — cheaper as a single M365 E3 bundle |
| 39 | **Bundle consolidation (E5)** | Office 365 E5 + EMS E5 + Windows E5 purchased separately — cheaper as a single M365 E5 bundle |
| 40 | **Teams unbundling** | Suite includes Teams but user has zero Teams activity — switch to the "without Teams" SKU variant to reduce cost |
| 41 | **Windows license waste** | Standalone Windows E3/E5 assigned but user has no Windows device activations — only uses Mac or mobile |
| 42 | **Seeded Visio overlap** | Visio Plan 1 assigned but E3/E5 suite already includes the Visio web app — Plan 1 is redundant |
| 43 | **Frontline candidate** | E3/E5 user who only uses web and mobile apps — eligible for much cheaper F3 Frontline license |
| 44 | **Frontline (high confidence)** | E3/E5 user with web/mobile only usage and zero desktop device activations — strong candidate for F3 downgrade |
| 44b | **Frontline CPC guard** | Web/mobile only but has Cloud PC Enterprise license — recommends E3 instead of F3 to maintain CPC prerequisites (Intune, Entra P1, Windows Enterprise E3) |
| 45 | **Frontline blocked (archive)** | Web/mobile only usage but mailbox has an archive — F3 does not support archive mailboxes |
| 46 | **Frontline rescue** | Archive blocks F3 downgrade — E1 or Business Basic supports archive and costs less than E3/E5 |
| 47 | **Frontline blocked (multi-PC)** | Web/mobile only usage but Office is activated on 2+ PCs — F3 limits desktop apps to shared/VDI devices only |
| 48 | **Frontline review** | E3/E5 user may qualify for F3 but app platform usage data is unavailable — manual review needed |
| 49 | **F3→F1 micro-downgrade** | F3 user whose only activity is Teams on mobile/web — F1 provides Teams access at approximately half the cost |
| 50 | **Frontline add-on bloat** | F1/F3 base license + multiple add-ons costs more than Business Premium or E3 — consolidate to a single suite |
| 51 | **Frontline compliance breach** | F1/F3 user with desktop Office app activations (Windows/Mac) — F-series does NOT include desktop apps, this is a licensing compliance violation flagged in Microsoft audits |
| 52 | **Business Basic candidate** | Business Standard user who only uses web/mobile apps — downgrade to Business Basic |
| 53 | **Business Basic review** | Business Standard user may qualify for Basic but app usage data is unavailable — manual review needed |
| 54 | **E1→Business Basic arbitrage** | Office 365 E1 user on a tenant under 300 seats — Business Basic offers the same features at a lower price |
| 55 | **O365 E3→E1 downgrade** | Office 365 E3 user with no desktop app usage and mailbox under 50 GB — Office 365 E1 provides sufficient capability |
| 56 | **E3→Business Premium** | M365 E3 user on a tenant under 300 seats with mailbox under 50 GB — Business Premium is cheaper and includes security features |
| 57 | **Business Premium inversion** | Business Standard + security/compliance add-ons costs more than Business Premium — upgrade saves money and adds Intune + Entra P1 |
| 58 | **Business Premium security review** | Business Premium already includes Defender for Business — verify the additional Defender Suite add-on is justified |
| 59 | **E5 Voice waste** | M365 E5 user with zero Teams calls and meetings — swap to E5 without Audio Conferencing to save on telephony cost |
| 60 | **Apps Enterprise→Business** | M365 Apps for Enterprise on a tenant under 300 seats — the identical Apps for Business SKU is cheaper |
| 61 | **Standalone desktop app waste** | M365 Apps for Enterprise/Business assigned but user never uses desktop Office — only web/mobile access detected |
| 62 | **A la carte waste (Kiosk)** | Exchange Kiosk + M365 Apps purchased separately costs more than a single Business Standard license |
| 63 | **A la carte waste (Plan 1)** | Exchange Plan 1 + M365 Apps purchased separately costs more than a single Business Standard license |
| 64 | **Bundle inefficiency** | Business Basic + Apps for Business purchased separately costs more than a single Business Standard license |
| 65 | **EXO Plan 2 downgrade** | Mailbox under 50 GB on Exchange Plan 2 — Exchange Plan 1 is cheaper and provides up to 50 GB |
| 66 | **EXO Plan 2 review** | Exchange Plan 2 assigned but usage data is unavailable — manual review needed before downgrading |
| 67 | **EXO Plan 1→Kiosk** | Standalone Exchange Plan 1 but web-only access and <2 GB mailbox — Kiosk is 75% cheaper. Includes compliance note when user is covered by CA/MDO policies |
| 68 | **OneDrive Plan 2→Plan 1** | Standalone OneDrive Plan 2 (unlimited storage) assigned but user stores less than 900 GB — Plan 1 with 1 TB cap is sufficient |
| 69 | **Entra P2→P1 downgrade** | Standalone Entra ID P2 assigned to non-admin who does not use PIM or risk-based Conditional Access — Entra ID P1 is sufficient |
| 70 | **Over-licensed archive** | Exchange Online Archiving add-on on a small mailbox — archive is underused and cheaper options exist |
| 71 | **Redundant archive** | Exchange Plan 2 + standalone Exchange Online Archiving — Plan 2 already includes archiving natively |
| 72 | **Teams Phone right-sizing** | Shared mailbox or room account with full Teams Phone Standard — switch to cheaper Teams Shared Devices license |
| 73 | **PBI PPU add-on waste** | Full Power BI Premium Per User license assigned but user already gets Pro from their suite — switch to the cheaper PPU add-on |

#### Activity & Behavioral Analysis
| # | Check | Description |
|---|-------|-------------|
| 74 | **Shelfware (Teams Premium)** | Teams Premium assigned but user organized fewer than 3 meetings — not getting value from the add-on |
| 75 | **Shelfware (Visio/Project)** | Visio or Project desktop license assigned but zero product activation detected — user never opened the app |
| 76 | **Premium add-on waste** | Visio or Project desktop license assigned but user only activates on web/mobile — downgrade to cheaper web plan |
| 77 | **Shelfware review (web-only)** | Expensive license assigned but only web access detected and no desktop telemetry available — manual review needed |
| 78 | **Shelfware (generic)** | Expensive standalone license with zero app activity in the reporting period — consider reclaiming |
| 79 | **Teams Phone without calling plan** | Teams Phone System assigned but no Microsoft Calling Plan or Operator Connect configured — user cannot make external calls |
| 80 | **Calling plan shelfware** | Paid PSTN Calling Plan assigned but user made zero Teams calls in the reporting period — remove the calling plan |
| 81 | **Legacy service account** | Account only uses legacy protocols (POP3/IMAP4/SMTP) but holds a premium suite license — does not need a full license |
| 82 | **Copilot prerequisite missing** | Microsoft 365 Copilot assigned but user lacks the required base license (E3/E5/Business Standard/Premium) — Copilot will not function |
| 83 | **Copilot reclaim (no activity)** | Copilot assigned but zero Copilot usage and zero M365 workload activity — strong candidate for immediate reclaim |
| 84 | **Copilot watchlist** | Copilot assigned with zero Copilot usage but user actively uses M365 workloads — monitor adoption before reclaiming |
| 85 | **Copilot reclaim (not in report)** | Copilot assigned but user does not appear in Copilot usage report and has no M365 activity — reclaim |
| 86 | **Copilot watchlist (not in report)** | Copilot assigned, user not in Copilot usage report but actively uses M365 workloads — monitor adoption |
| 87 | **Copilot reclaim (report unavailable)** | Copilot assigned, usage report unavailable, and user has no M365 workload activity — reclaim |
| 88 | **Copilot active** | Copilot assigned and actively used across M365 apps — no action needed |
| 89 | **Security Copilot** | Microsoft Security Copilot assigned — verify the security operations team is actively using it |
| 90 | **Copilot Studio (inactive)** | Copilot Studio license assigned but user has never signed in — verify if the license is still needed |
| 91 | **Copilot Studio (active)** | Copilot Studio license assigned to active user — admin/developer tool, no action needed |
| 92 | **AI add-on overlap** | User has both Teams Premium and Copilot but organizes no meetings — Teams Premium is redundant, remove it |
| 93 | **AI overlap review** | User has both Teams Premium and Copilot and actively organizes meetings — verify if webinar/town hall features justify keeping both |
| 94 | **Power BI Pro with Premium Capacity** | Power BI Pro assigned but tenant has Premium Capacity — users only consuming reports can use Free tier instead |
| 95 | **Power BI Pro + PPU overlap** | User has both Power BI Pro and Power BI Premium Per User — Pro is redundant and can be removed |
| 96 | **Expensive cold storage** | Zero M365 activity but user has a large mailbox (>10 GB) or OneDrive (>50 GB) — paying for storage only, consider archiving |
| 97 | **Forwarding mailbox review** | User has low email activity but mailbox auto-forwards to another address — verify if the mailbox is still needed |
| 98 | **Intune shelfware** | Standalone Intune or EMS license assigned but user has zero enrolled devices in Intune — device management entirely unused |
| 99 | **MDM/MAM waste** | Standalone Intune or EMS license assigned but user only accesses M365 via web browser — no devices to manage |
| 100 | **Heavy external sharer** | More than 50% of user's content is shared externally — review Data Loss Prevention (DLP) policies |

#### Security & Compliance Coverage
| # | Check | Description |
|---|-------|-------------|
| 101 | **Security gap (Business)** | Business SKU user without endpoint or email threat protection — no Defender for Endpoint or Defender for Office 365 |
| 102 | **Defender Suite upsell** | User has partial Microsoft Defender coverage — specific missing components identified (MDO, MDE, MDI, MDCA, or XDR) |
| 103 | **Defender Suite upsell (E3)** | E3 user with partial Defender coverage — full Defender for E3 bundle available as upgrade path |
| 104 | **Purview upsell (full Defender)** | User has full Defender stack but no Microsoft Purview compliance coverage — DLP and information protection gap |
| 105 | **Purview upsell (no coverage)** | User has no compliance coverage — no DLP for email, Teams, or endpoints detected |
| 106 | **Purview upsell (E3)** | E3 user with Defender coverage but no Purview — compliance gap identified |
| 107 | **High risk sharing** | Heavy external sharing activity without DLP or compliance controls in place — data exfiltration risk |
| 108 | **Entra Suite overlap** | User has both Entra ID P2 and Entra ID Governance separately — paying for overlapping capabilities |
| 109 | **Bundle consolidation (Entra)** | Entra ID P2 + Governance purchased separately — cheaper as a single Entra Suite license |
| 110 | **Entra Suite review** | Entra Suite assigned alongside E3/E5 — verify bundled Entra P1/P2 features don't already cover the need |
| 111 | **Intune Suite waste** | Intune Suite add-on on E3/E5 user — Remote Help, Advanced Analytics, and EPM are now included in E3/E5 (late 2025) |

#### Administrative & Compliance
| # | Check | Description |
|---|-------|-------------|
| 112 | **Admin license review** | Admin account holding a full productivity suite (E3/E5) — admins should use a security-focused SKU only |
| 113 | **PIM licensing check (licensed)** | Admin uses Privileged Identity Management but does not have an Entra ID P2 license — required for PIM; P2 is a superset of P1 and also covers CA requirements |
| 114 | **Risk-based CA licensing check** | User is in scope of a risk-based Conditional Access policy — requires Entra ID P2 for risk detection to function |
| 115 | **CA P1 licensing check (scoped)** | User is targeted by a Conditional Access policy but does not have Entra ID P1 — required for CA enforcement |
| 116 | **CA P1 licensing check (tenant-wide)** | User is included in tenant-wide Conditional Access policies but does not have Entra ID P1 |
| 117 | **MDO policy licensing check** | Mailbox is in scope of Defender for Office 365 Safe Links/Attachments/Anti-Phishing rules but lacks the MDO entitlement |
| 118 | **Licensing error** | Group-based license assignment has failed (e.g. CountViolation, MutuallyExclusive) — user may not have expected licenses |
| 119 | **Duplicate review** | License appears redundant with an assigned suite but could not be fully confirmed — manual review recommended |
| 120 | **Duplicate coverage** | Standalone license is fully covered by the user's suite — remove the standalone to stop double-paying |
| 121 | **Trial license expiry** | Trial subscription approaching expiry — convert to paid or remove before it expires |
| 122 | **License capacity queue** | User is in queue waiting for a license seat — purchase additional seats or free up existing assignments |

#### Dormancy & Automation Detection
| # | Check | Description |
|---|-------|-------------|
| 123 | **Dormant account** | User has not signed in interactively for more than the configured threshold (default 30 days) — license may be wasted |
| 124 | **Automation account (dormant admin)** | Admin has no interactive sign-in but has recent non-interactive (API/service) sign-in — this is a service/automation account, not truly dormant |
| 125 | **Automation account (non-admin)** | Non-admin user has no interactive sign-in but has recent non-interactive sign-in — likely a service or automation account |
| 126 | **Automation account (UPN/role pattern)** | Account identified as infrastructure/sync by Directory Sync role or UPN pattern (sync_*, adsync*, svc_*, service_*) — consider converting to Workload Identity |
| 127 | **Dormant admin risk** | Admin account with no sign-in at all (interactive or non-interactive) — security risk and potential license waste |
| 128 | **Automation account (never signed in)** | Service account pattern detected (Directory Sync role or UPN) with no sign-in on record — verify if the license is still needed |
| 129 | **Never signed in** | Licensed user has never signed in — license may have been assigned but never used |
| 130 | **Forwarding mailbox waste (dormant)** | Dormant or never-signed-in user whose mailbox only auto-forwards to another address — license not needed for forwarding |
| 131 | **Forwarding mailbox review** | User has low email activity and mailbox auto-forwards — verify if the forwarding mailbox is still needed |

#### Usage Observations (standalone licenses only)
| # | Check | Description |
|---|-------|-------------|
| 132 | **No desktop apps** | User on standalone license only uses web/mobile apps — no desktop Office installations detected |
| 133 | **Mobile apps only** | User on standalone license only accesses M365 from mobile devices — potential Frontline (F-license) candidate |
| 134 | **No M365 app activity** | User on standalone license has zero desktop, web, and mobile app activity in the reporting period |
| 135 | **No Outlook desktop** | User has Exchange entitlement but only uses Outlook on the web or mobile — no Outlook desktop client detected |
| 136 | **Teams web-only** | User accesses Teams only via web browser, no desktop client — potential Frontline candidate |
| 137 | **Low Exchange usage** | User has Exchange entitlement but email send/receive volume is below the configured low threshold |
| 138 | **Low Teams usage** | User has Teams entitlement but chat/call/meeting activity is below the configured low threshold |
| 139 | **Low OneDrive usage** | User has OneDrive entitlement but file activity is below the configured low threshold |

#### Operational Review
| # | Check | Description |
|---|-------|-------------|
| 140 | **Mailbox storage warning (Plan 1)** | Mailbox approaching the 50 GB Exchange Plan 1 limit — mail flow will stop when the quota is reached |
| 141 | **Mailbox storage warning (Plan 2)** | Mailbox approaching the 100 GB Exchange Plan 2 limit — upgrade or archive needed before quota is hit |
| 142 | **Mailbox storage warning (Kiosk)** | Exchange Kiosk mailbox approaching its 2 GB limit — mail flow will stop at cap |
| 143 | **OneDrive storage warning** | OneDrive approaching the 1 TB storage limit on Business/E1 plans — file sync will stop at cap |
| 144 | **Data gap** | License SKU not recognized in the reference data — update M365SkuData.json to include this SKU for accurate analysis |

#### Cloud PC Optimization
| # | Check | Description |
|---|-------|-------------|
| 145 | **Dormant Cloud PC** | Windows 365 Cloud PC provisioned but 0 hours connected in 90 days — license waste |
| 146 | **Cloud PC review** | Windows 365 Cloud PC with less than 10 hours connected in 90 days — low usage, review need |

#### Tenant-Level Optimization
| # | Check | Description |
|---|-------|-------------|
| 147 | **Unassigned license pool waste** | Purchased license seats sitting unassigned in tenant inventory — costing more than €500/year and over 5% of the pool |
| 148 | **Teams Rooms Basic vs Pro** | Paying for Teams Rooms Pro licenses when the tenant has 25 or fewer rooms — qualifies for the free Teams Rooms Basic tier |

## Requirements

### PowerShell Modules

| Module | Required | Purpose |
|--------|----------|---------|
| Microsoft.Graph.Authentication | Yes | Graph connection |
| Microsoft.Graph.Users | Yes | User and license data |
| Microsoft.Graph.Reports | Yes | Usage report downloads |
| Microsoft.Graph.Identity.DirectoryManagement | Yes | SKU inventory, admin roles |
| ExchangeOnlineManagement | Optional | Mailbox type, shared mailbox, and MDO policy detection — without it the script skips mailbox type checks and MDO evaluation |
| ImportExcel | Optional | Excel workbook output (.xlsx) — without it the script produces CSV/TXT only |

By default, the script **stops with a clear error** listing the exact `Install-Module`
commands needed if any required modules are missing. Use `-AutoInstallModules` to install
them automatically instead.

```powershell
# Safe default: script stops and tells you what to install
.\M365-LOA.ps1

# Auto-install missing required modules (CurrentUser scope)
.\M365-LOA.ps1 -AutoInstallModules
```

### Graph API Permissions

| Permission | Required For |
|------------|-------------|
| User.Read.All | User properties and assigned licenses |
| Reports.Read.All | All 11 usage reports |
| Organization.Read.All | Subscribed SKUs, subscription lifecycle |
| AuditLog.Read.All | Sign-in activity (beta) |
| RoleManagement.Read.Directory | Admin role assignments |
| Group.Read.All | Resolving license group names |
| Policy.Read.All | Conditional Access policies (risk-based CA detection) |
| DeviceManagementManagedDevices.Read.All | Enrolled device count (Intune shelfware detection) |
| CloudPC.Read.All | Windows 365 Cloud PC provisioning and usage reports |
| ReportSettings.ReadWrite.All | Microsoft 365 usage reports hash UPNs by default; this permission flips the tenant setting to show real UPNs so the script can match usage to users. The customer must manually re-enable hashing after the assessment if desired |

## Setup

### Option 1 — App Registration (recommended for production)

Run `LOA-App-Registration-Setup.ps1` to create a dedicated app registration with
certificate-based auth. This creates a `LOA-Connection.json` file that the main
script auto-detects — no parameters needed.

```powershell
# One-time setup (requires Global Admin)
.\App registration\LOA-App-Registration-Setup.ps1

# Then just run the report — connection config is auto-detected
.\M365-LOA.ps1
```

### Option 2 — Interactive login (quick ad-hoc runs)

```powershell
# Signs in interactively via browser — no setup required
.\M365-LOA.ps1
```

## Usage

```powershell
# Basic run (30-day lookback, output in current directory)
.\M365-LOA.ps1

# Certificate auth with explicit parameters
.\M365-LOA.ps1 -ClientId "xxx" -TenantId "yyy" -CertificateThumbprint "zzz"

# Custom lookback and output folder
.\M365-LOA.ps1 -InactiveSignInDays 30 -OutputFolder "C:\Reports"

# Keep UPNs hashed in usage reports (default unhides them)
.\M365-LOA.ps1 -KeepHashedUPNs

# Run without Exchange Online (Graph-only mode)
.\M365-LOA.ps1 -SkipEXO

# Delta report comparing against a previous run
.\M365-LOA.ps1 -PriorReportPath "C:\Reports\M365_LicenseOptimization_20260201.csv"

# Custom pricing CSV (default: M365SkuPricing.csv alongside script)
.\M365-LOA.ps1 -PricingCsvPath "C:\Pricing\custom_prices.csv"

# Reduce parallel downloads to avoid throttling on busy tenants
.\M365-LOA.ps1 -MaxParallel 2

# Include disabled+unlicensed accounts, skip Excel output
.\M365-LOA.ps1 -IncludeDisabledAccounts -NoExcel
```

### Parameters

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-ReportPeriod` | D90 | Lookback window: D7, D30, D90, D180 |
| `-OutputFolder` | Current dir | Folder for output files (auto-created if missing) |
| `-IncludeDisabledAccounts` | Off | Include disabled+unlicensed accounts in output |
| `-KeepHashedUPNs` | Off | Skip unhiding UPNs in reports (default unhides; supports `-WhatIf`) |
| `-SkipEXO` | Off | Run without Exchange Online (skips mailbox type/litigation hold/MDO) |
| `-AutoInstallModules` | Off | Auto-install missing required modules (safe default: stops with instructions) |
| `-PriorReportPath` | (none) | Previous run's main CSV for delta analysis |
| `-SkuDataPath` | `M365SkuData.json` | External SKU names + prices JSON file |
| `-PricingCsvPath` | `M365SkuPricing.csv` | CSV with `SkuPartNumber,MonthlyPriceEUR` columns (default file ships with script) |
| `-MaxParallel` | 6 | Max concurrent Graph API report downloads (1-11) |
| `-RulePackPath` | `LOA_RulePack_M365.json` | LOA rule pack with manual audit checklist rules and doc refs |
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
| `-InactiveSignInDays` | 30 | Days without sign-in to flag as dormant (1-365) |

## Output Files

The full pipeline generates up to 11 files with a timestamp suffix:

| # | File | Description |
|---|------|-------------|
| 1 | **M365_LicenseOptimization_{ts}.csv** | Main per-user report: usage data, platform flags, intensity scores, capability levels, cost, assessment text, pre-computed `Estimated Annual Savings (EUR)` and `Estimated Compliance Cost (EUR)` columns |
| 2 | **M365_ServicePlanDetail_{ts}.csv** | Granular SKU and service plan breakdown per user with provisioning status |
| 3 | **M365_SkuInventory_{ts}.csv** | Tenant-level license inventory with friendly names, consumed/available counts, pricing, subscription status, and expiry dates |
| 4 | **M365_OptimizationSummary_{ts}.txt** | Human-readable summary: executive summary with tiered savings model, cost analysis, assessment distribution, Copilot reclaim pipeline breakdown, data collection warnings, manual audit checklist (from LOA rule pack) |
| 5 | **M365_ExecutiveSummary_{ts}.csv** | 9-tier executive summary: Tier 1 quick wins, Tier 2 right-sizing, Tenant optimization, Product flags, Copilot pipeline, Operational Review, Licensing Compliance (CA/MDO/PIM breakdown), Security & Compliance coverage, and Security Posture distribution |
| 6 | **M365_LicenseOptimization_{ts}.xlsx** | Excel workbook with up to 15 worksheets (see below) |
| 7 | **M365_LicenseDelta_{ts}.csv** | *(if `-PriorReportPath` provided)* Delta analysis: user changes, cost trends, assessment shifts, dormancy/Copilot adoption tracking |
| 8 | **M365_LicenseGroups_{ts}.csv** | Entra ID licensing groups with membership type, member count, and assigned SKUs |
| 9 | **M365_OptimizationHeatmap_{ts}.html** | *(generated by M365-Heatmap.ps1)* Interactive HTML dashboard with up to 8 tabs: Overview tiles, Assessments by Category, Assessments by User, Assessments by SKU, License Groups, Workload Usage Matrix, Subscription Alerts (conditional), Copilot Adoption (conditional). Includes a Decision Framework overlay, a Tile Reference Legend with all categories grouped by priority tier, and tenant display name in the header. |
| 10 | **M365_AssessmentReport_{ts}.docx** | *(generated by M365-AssessmentDoc.ps1)* Branded Word assessment report with executive financial summary, tiered findings, compliance cost callout, and disclaimers |
| 11 | **M365_LicenseAssessment_{ts}.xlsx** | *(generated by M365-SummaryXLS.ps1)* Customer-facing Excel workbook — post-processed copy with assessment language, category coloring, rebuilt tabs by primary category, and zero-cost/informational rows removed from summary |

#### Excel Worksheets

| # | Sheet | Content |
|---|-------|---------|
| 1 | Executive Summary | Tiered savings model, key metrics, unassigned license inventory |
| 2 | User Report | Full per-user data with conditional formatting |
| 3 | Admin Review | *(if matching users)* Users flagged for admin review — color-coded blue tab |
| 4 | Dormant | *(if matching users)* Dormant accounts with no sign-in — red tab |
| 5 | Disabled Account | *(if matching users)* Disabled accounts with paid licenses — red tab |
| 6 | Guest Account Waste | *(if matching users)* Guest users with paid licenses — yellow tab |
| 7 | Shared Mailbox | *(if matching users)* Shared mailboxes with license assessments — green tab |
| 8 | Automation Account | *(if matching users)* Service/sync accounts — grey tab |
| 9 | Never Signed In | *(if matching users)* Accounts that have never signed in — red tab |
| 10 | Service Plans | Granular SKU/service plan per user |
| 11 | SKU Inventory | Tenant license inventory with pricing and expiry highlighting |
| 12 | Group Licensing | *(if group-based licensing detected)* Entra ID groups with assigned licenses, member counts, disabled plans |
| 13 | Cost by Department | Department-level cost aggregation with bar chart |
| 14 | Assessments | Category-level summary with costs and pie chart |
| 15 | Intensity Analysis | *(if usage data available)* Per-user workload intensity cross-tab (Exchange, Teams, OneDrive, SharePoint) |

## Performance

| Tenant Size | Approximate Runtime |
|-------------|-------------------|
| 1,000 users | ~2 minutes |
| 10,000 users | ~5 minutes |
| 100,000 users | ~15 minutes |

Runtime is dominated by Graph report downloads and Exchange Online queries. License
processing is done entirely in-memory with no per-user API calls.

## Updating SKU Data

`M365SkuData.json` is the single source of truth for SKU reference data: friendly names,
suite-to-service-plan mappings, plan capabilities, capability aliases, add-on bundles,
premium suite lists, and coverage aliases. The main script warns if the file is older
than 90 days. To refresh SKU friendly names:

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
M365-LOA.ps1    # Main audit script (~8400 lines)
M365-Heatmap.ps1               # Interactive HTML dashboard generator
M365-AssessmentDoc.ps1             # Word assessment report generator (calls generate-report.js)
generate-report.js                       # Node.js docx builder (requires: npm install docx)
M365-SummaryXLS.ps1                   # XLSX post-processor (customer-facing copy)
M365SkuData.json                         # SKU reference data (names, suite maps, capabilities, aliases)
M365SkuPricing.csv                       # SKU monthly prices (EUR) — editable CSV
LOA_RulePack_M365.json                   # Manual audit checklist rules and documentation refs
assets/                                  # Brand assets (cover/header images, not tracked in git)
App registration\
  LOA-App-Registration-Setup.ps1         # App registration + certificate setup
```

## Notes

- Usage data has ~48 hour reporting latency from Microsoft
- UPNs are unhidden by default — pass `-KeepHashedUPNs` to skip this step
- Unhiding UPNs changes a tenant-wide setting — the customer must manually re-enable hashing in M365 Admin Center after the assessment if desired
- Sign-in activity uses the beta Graph API and requires `AuditLog.Read.All`
- Mailbox type detection requires `ExchangeOnlineManagement` module (optional)
- The script uses `Set-StrictMode -Version Latest` for reliability

## Version

Current: **v0.7**

---

## License

Copyright (c) 2024-2026 B5rman. All rights reserved. See [LICENSE](LICENSE) for details.
