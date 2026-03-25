# ========================================================
# M365 License Optimization Assessment
# Version : 1.1.0
# Author  : Bruno Vijverman
# Reads the CSV output from M365-LOA.ps1
# and generates a standalone HTML heatmap dashboard.
# ========================================================

<#
.SYNOPSIS
    Generates an interactive HTML heatmap dashboard from M365 License Optimization report output.

.DESCRIPTION
    Reads the CSV files produced by M365-LOA.ps1 and creates a
    self-contained HTML file with three visual views:
      1. Potential Savings by User — quick-win tiles + sortable user table with recommendation popup
      2. Potential Savings by SKU  — horizontal bar chart of waste per license type
      3. Over/Under-Licensed — per-user workload capability matrix

.PARAMETER OutputFolder
    Folder containing the M365_LicenseOptimization_*.csv files. Defaults to .\output

.PARAMETER ReportCsv
    Full path to a specific M365_LicenseOptimization_*.csv file.

.PARAMETER SummaryCsv
    Full path to a specific M365_ExecutiveSummary_*.csv file.

.PARAMETER HtmlOutput
    Path for the generated HTML file.

.EXAMPLE
    .\M365-Heatmap.ps1
    .\M365-Heatmap.ps1 -OutputFolder "C:\Reports\Contoso"
#>

[CmdletBinding()]
param(
    [string]$OutputFolder = ".\output",
    [string]$ReportCsv   = "",
    [string]$SummaryCsv  = "",
    [string]$HtmlOutput  = ""
)

Set-StrictMode -Version 2

# ── Resolve input files ──────────────────────────────────────────────────────
if (-not $ReportCsv) {
    $candidates = @(Get-ChildItem -Path $OutputFolder -Filter "M365_LicenseOptimization_*.csv" -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -notmatch "SkuInventory|ServicePlan" } |
                    Sort-Object LastWriteTime -Descending)
    if (-not $candidates) {
        Write-Error "No M365_LicenseOptimization_*.csv found in '$OutputFolder'. Use -ReportCsv to specify a file."
        exit 1
    }
    $ReportCsv = $candidates[0].FullName
}

if (-not (Test-Path $ReportCsv)) {
    Write-Error "Report CSV not found: $ReportCsv"
    exit 1
}

if (-not $SummaryCsv) {
    $folder = Split-Path $ReportCsv
    $sum = @(Get-ChildItem -Path $folder -Filter "M365_ExecutiveSummary_*.csv" -ErrorAction SilentlyContinue |
             Sort-Object LastWriteTime -Descending)
    if ($sum) { $SummaryCsv = $sum[0].FullName }
}

if (-not $HtmlOutput) {
    $base = [System.IO.Path]::GetFileNameWithoutExtension($ReportCsv) -replace '^M365_LicenseOptimization_', ''
    $HtmlOutput = Join-Path (Split-Path $ReportCsv) "M365_OptimizationHeatmap_$base.html"
}

Write-Host "Reading report : $ReportCsv" -ForegroundColor Cyan
if ($SummaryCsv) { Write-Host "Reading summary: $SummaryCsv" -ForegroundColor Cyan }

# ── Load CSVs ────────────────────────────────────────────────────────────────
$rows = @(Import-Csv -Path $ReportCsv -Encoding UTF8)
Write-Host "  Loaded $($rows.Count) user rows" -ForegroundColor Gray

$summaryRows = @()
if ($SummaryCsv -and (Test-Path $SummaryCsv)) {
    $summaryRows = @(Import-Csv -Path $SummaryCsv -Encoding UTF8)
}

# ── Auto-discover License Groups CSV ─────────────────────────────────────────
$groupRows = @()
$groupsCsv = @(Get-ChildItem -Path (Split-Path $ReportCsv) -Filter "M365_LicenseGroups_*.csv" -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending)
if ($groupsCsv) {
    $groupRows = @(Import-Csv -Path $groupsCsv[0].FullName -Encoding UTF8)
    Write-Host "  Loaded $($groupRows.Count) licensing group(s)" -ForegroundColor Gray
}

# ── Auto-discover SKU Inventory CSV (for group table context) ────────────────
$skuInvLookup = @{}
$skuInvCsv = @(Get-ChildItem -Path (Split-Path $ReportCsv) -Filter "M365_SkuInventory_*.csv" -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending)
if ($skuInvCsv) {
    $skuInvRows = @(Import-Csv -Path $skuInvCsv[0].FullName -Encoding UTF8)
    foreach ($si in $skuInvRows) {
        $skuInvLookup[$si.'SkuPartNumber'] = [PSCustomObject]@{
            Total    = [int]($si.'Total' -replace '\D','')
            Consumed = [int]($si.'ConsumedUnits' -replace '\D','')
        }
    }
    Write-Host "  Loaded $($skuInvRows.Count) SKU inventory row(s)" -ForegroundColor Gray
}

# ── Disclaimer text (hardcoded — not tenant-specific) ────────────────────────
$disclaimer1 = 'All cost figures are indicative estimates based on public Microsoft list prices. — Actual costs may differ due to EA/CSP/volume pricing.'
$disclaimer2 = 'Copilot usage and Cloud PC analytics rely on Microsoft Graph BETA APIs — These sections may show limited results until the API becomes generally available.'
$disclaimer3 = 'All assessments are advisory. — Assessment scenarios should be validated before making any license changes.'
$disclaimer4 = 'Usage data is based on the last 90 days of Microsoft 365 activity reports. — Users on leave or seasonal workers may appear inactive.'

# ── Helper: parse EUR amount ─────────────────────────────────────────────────
# Handles both European (16560,0 / 1.234,56) and standard (1,234.56) formats
function Parse-Decimal([string]$s) {
    $s = ($s -replace '[\u20AC ]','').Trim()
    if (-not $s) { return [decimal]0 }
    $v = [decimal]0
    # European format: ends with comma + 1-3 digits (e.g. "16560,0" or "1.234,56")
    if ($s -match '^[\d.]+,\d{1,2}$') {
        $s2 = $s -replace '\.','' -replace ',','.'
        if ([decimal]::TryParse($s2,[System.Globalization.NumberStyles]::Any,
            [System.Globalization.CultureInfo]::InvariantCulture,[ref]$v)) { return $v }
    }
    # Standard format: commas are thousands separators
    $s = $s -replace ',',''
    if ([decimal]::TryParse($s,[System.Globalization.NumberStyles]::Any,
        [System.Globalization.CultureInfo]::InvariantCulture,[ref]$v)) { return $v }
    return [decimal]0
}

# ── Savings & compliance cost are now read directly from CSV columns ──────────
# 'Estimated Annual Savings (EUR)' and 'Estimated Compliance Cost (EUR)' are pre-computed
# by M365-LOA.ps1 per user. This eliminates fragile regex extraction from recommendation text.
# The deprecated functions below are retained only for backward compatibility with older CSVs
# that lack the new columns — in that case, the heatmap falls back to €0 (safe default).

function _DEPRECATED_Get-EstimatedSavings([string]$category,[decimal]$annualCost,[string]$recommendation) {
    # Tier 1 categories normally return full annual cost (license removal).
    # Exception: right-sized NO ACTIVITY recs contain "Potential savings: €X.XX/yr"
    # instead of "Annual cost" — use the net savings, not the full license cost.
    # IMPORTANT: Only check the PRIMARY recommendation segment (before the first pipe)
    # to avoid secondary recs (e.g. E5 VOICE REVIEW) overriding the full savings.
    # Also promote to Tier 1 when the rec contains a secondary Tier-1 tag (e.g. DORMANT)
    # that indicates the full license is reclaimable, even if the primary category is
    # more specific (e.g. "Dormant Cloud PC" with secondary "DORMANT — ...").
    $isTier1 = $tier1.Contains($category) -or
               ($recommendation -match '(^|\| )(DORMANT|DISABLED ACCOUNT|DISABLED SHARED|NO ACTIVITY|NEVER SIGNED IN) -')
    if ($isTier1) {
        $primarySegment = ($recommendation -split '\|')[0]
        if ($primarySegment -match 'Potential savings:.*\u20AC([\d.,]+)/yr') {
            return Parse-Decimal $Matches[1]
        }
        return $annualCost
    }
    # Strip LICENSING CHECK segments — those amounts are compliance costs (licenses
    # to ADD), not savings. Without this, a Room/Equipment or other non-cost category
    # with LICENSING CHECK findings would show compliance cost as "savings".
    $recForSavings = $recommendation -replace 'LICENSING CHECK[^|]*', ''
    [decimal]$total = 0
    # Pattern 1: explicit /yr amounts  (e.g. "Saves\u20AC3,20/mo (\u20AC38,40/yr)")
    foreach ($m in [regex]::Matches($recForSavings, '\u20AC([\d.,]+)/yr')) {
        $total += Parse-Decimal $m.Groups[1].Value
    }
    # Pattern 2: Annual waste / overlap cost / savings without /yr suffix
    # Covers: "Annual waste: €X.XX" (A La Carte), "Annual overlap cost: €X.XX" (Duplicate Coverage),
    #          "Annual savings: €X.XX" (AI Add-On Overlap, PBI Pro Review, Archive, etc.)
    foreach ($m in [regex]::Matches($recForSavings, 'Annual (?:waste|overlap cost|savings): \u20AC([\d.,]+)')) {
        $total += Parse-Decimal $m.Groups[1].Value
    }
    # Pattern 3: "Annual cost: €X.XX" (no /yr suffix) — ONLY from item-specific tags.
    # Tags like DORMANT, AUTOMATION ACCOUNT, NEVER SIGNED IN cite the full user license
    # cost for context — these are NOT independent savings. Only these tags cite a specific
    # add-on/subscription cost that represents genuine additional savings:
    $itemCostTags = 'INACTIVE ADD-ON|DORMANT CLOUD PC|CLOUD PC REVIEW|WINDOWS LICENSE REVIEW|FORWARDING MAILBOX REVIEW'
    foreach ($m in [regex]::Matches($recForSavings, "(?:^|\| )\s*(?:$itemCostTags)[^|]*Annual cost: \u20AC([\d.,]+)")) {
        $total += Parse-Decimal $m.Groups[1].Value
    }
    return $total
}

# ── Per-user compliance cost estimation (€ amounts inside LICENSING CHECK segments) ──
function Get-EstimatedComplianceCost([string]$recommendation) {
    [decimal]$total = 0
    foreach ($m in [regex]::Matches($recommendation, 'LICENSING CHECK[^|]*')) {
        foreach ($yr in [regex]::Matches($m.Value, '\u20AC([\d.,]+)/yr')) {
            $total += Parse-Decimal $yr.Groups[1].Value
        }
    }
    # Shared mailbox MDO compliance: when a shared mailbox is in scope of MDO policies,
    # the MDO P1 add-on cost is a compliance requirement (not optional savings).
    # The cost is embedded in SHARED MAILBOX / DISABLED SHARED MAILBOX segments.
    foreach ($m in [regex]::Matches($recommendation, '(?:DISABLED )?SHARED MAILBOX[^|]*Defender for Office[^|]*')) {
        foreach ($yr in [regex]::Matches($m.Value, '\u20AC([\d.,]+)/yr')) {
            $total += Parse-Decimal $yr.Groups[1].Value
        }
    }
    return $total
}

# ── Build per-user heatmap data ───────────────────────────────────────────────
Write-Host "  Processing users..." -ForegroundColor Gray
# ── Recommendation prefix → category mapping (for secondary tags) ────────────
$_recPrefixMap = [ordered]@{
    # Order matters — longer/more-specific prefixes must come before shorter ones
    # because the matching loop uses ^$prefix and breaks on first match.
    'INACTIVE HOLD WITH LICENSE' = 'Inactive Hold With License'
    'INACTIVE HOLD'       = 'Inactive Hold'
    'INACTIVE MAILBOX'    = 'Inactive Mailbox'
    'DISABLED SHARED MAILBOX' = 'Disabled Account'
    'DISABLED ACCOUNT'    = 'Disabled Account'
    'SHARED MAILBOX REVIEW' = 'Shared Mailbox Review'
    'SHARED MAILBOX'      = 'Shared Mailbox'
    'OVERLAPPING LICENSE' = 'Overlapping License'
    'DUPLICATE COVERAGE'  = 'Duplicate Coverage'
    'DUPLICATE REVIEW'    = 'Duplicate Review'
    'SUITE INVERSION'     = 'Suite Inversion'
    'E5 CONSOLIDATION'    = 'E5 Upgrade'
    'BUNDLE CONSOLIDATION'= 'Bundle Consolidation'
    'BUNDLE OPPORTUNITY'  = 'Bundle Opportunity'
    'ENTRA SUITE OVERLAP' = 'Entra Suite Overlap'
    'ENTRA P2 DOWNGRADE'  = 'Entra P2 Downgrade'
    'INTUNE SUITE OVERLAP'= 'Intune Suite Overlap'
    'INTUNE REVIEW'       = 'Intune Review'
    'LICENSING CHECK'     = 'Licensing Compliance Gap'
    'LICENSING ERROR'     = 'License Error'
    'DORMANT ADMIN REVIEW'= 'Dormant Admin Review'
    'DORMANT ADMIN'       = 'Dormant Admin Review'
    'DORMANT CLOUD PC'    = 'Dormant Cloud PC'
    'STALE SIGN-IN'       = 'Stale Sign-In'
    'DORMANT'             = 'Dormant'
    'CLOUD PC REVIEW'     = 'Cloud PC Review'
    'AUTOMATION ACCOUNT'  = 'Automation Account'
    'LEGACY SERVICE ACCOUNT' = 'Legacy Service Account'
    'NEVER SIGNED IN'     = 'Never Signed In'
    'NO ACTIVITY'         = 'No Activity'
    'BACKGROUND SYNC ONLY'= 'Background Sync Only'
    'PREMIUM ADD-ON REVIEW'= 'Premium Add-On Review'
    'REDUNDANT ARCHIVE'   = 'Redundant Archive'
    'OVER-LICENSED ARCHIVE'= 'Over-Licensed Archive'
    'TEAMS UNBUNDLING'    = 'Teams Unbundling'
    'TEAMS PHONE RIGHT-SIZING' = 'Teams Phone Right-Sizing'
    'TEAMS PHONE REVIEW'  = 'Teams Phone Review'
    'E5 VOICE'            = 'E5 Voice Review'
    'EXO PLAN 2 REVIEW'   = 'EXO Plan 2 Review'
    'EXO PLAN 2'          = 'EXO Plan 2 Downgrade'
    'EXCHANGE KIOSK'      = 'Exchange Kiosk Downgrade'
    'FORWARDING MAILBOX'  = 'Forwarding Mailbox Review'
    'AI ADD-ON OVERLAP'   = 'AI Add-On Overlap'
    'AI OVERLAP REVIEW'   = 'AI Overlap Review'
    'COPILOT PREREQUISITE'= 'Copilot Prerequisite'
    'COPILOT RECLAIM'     = 'Copilot Reclaim'
    'COPILOT WATCHLIST'   = 'Copilot Watchlist'
    'COPILOT STUDIO'      = 'Copilot Studio'
    'COPILOT'             = 'Copilot'
    'CALLING PLAN REVIEW' = 'Calling Plan Review'
    'POWER BI PRO REVIEW' = 'Power BI Pro Review'
    'PBI PPU OVERLAP'     = 'PBI PPU Overlap'
    'APP ARBITRAGE'       = 'App Arbitrage'
    'SEEDED VISIO OVERLAP'= 'Seeded Visio Overlap'
    'GUEST ACCOUNT REVIEW'= 'Guest Account Review'
    'GUEST ACCOUNT'       = 'Guest User'
    'NON-HUMAN'           = 'Non-Human Account Review'
    'FREE LICENSE'        = 'Free License Overlap'
    'WINDOWS LICENSE'     = 'Windows License Review'
    'FRONTLINE ADD-ON STACKING' = 'Frontline Add-On Stacking'
    'FRONTLINE RESCUE'    = 'Frontline Rescue'
    'FRONTLINE CANDIDATE' = 'Frontline Candidate'
    'FRONTLINE REVIEW'    = 'Frontline Review'
    'FRONTLINE'           = 'Frontline Review'
    'INACTIVE ADD-ON REVIEW' = 'Inactive Add-On Review'
    'INACTIVE ADD-ON'     = 'Inactive Add-On'
    'ONEDRIVE PLAN 2 REVIEW' = 'OneDrive Plan 2 Review'
    'ONEDRIVE STORAGE WARNING' = 'OneDrive Storage Warning'
    'MAILBOX STORAGE'     = 'Mailbox Storage Warning'
    'EXPENSIVE COLD'      = 'Expensive Cold Storage'
    'BUSINESS PREMIUM INVERSION' = 'Business Premium Inversion'
    'BUSINESS BASIC CANDIDATE' = 'Business Downgrade'
    'E1 DOWNGRADE'        = 'E1 to Business Basic'
    'DATA GAP'            = 'Data Gap'
    'UNLICENSED WITH DATA'= 'Unlicensed With Data'
    'SECURITY GAP'        = 'Security Gap'
    'DEFENDER COVERAGE'   = 'Defender Coverage Review'
    'ADMIN'               = 'Admin Review'
}

$userData = foreach ($r in $rows) {
    $cat  = $r.'Recommendation Category'
    $cost = Parse-Decimal $r.'Annual License Cost (EUR)'
    $rec  = $r.'Recommendation'
    if ($cat -eq 'No Findings' -or $cat -eq '' -or $cat -eq 'Unlicensed' -or $cat -eq 'Copilot Active') { continue }
    $savings  = if ($r.'Estimated Annual Savings (EUR)') { Parse-Decimal $r.'Estimated Annual Savings (EUR)' } else { [decimal]0 }
    $compCost = if ($r.'Estimated Compliance Cost (EUR)') { Parse-Decimal $r.'Estimated Compliance Cost (EUR)' } else { [decimal]0 }

    # Extract secondary categories from | separated recommendation segments
    $secondaryCats = @()
    if ($rec -match '\|') {
        $segments = @($rec -split '\s*\|\s*')
        foreach ($seg in $segments) {
            foreach ($prefix in $_recPrefixMap.Keys) {
                if ($seg -match "^$prefix") {
                    $mapped = $_recPrefixMap[$prefix]
                    if ($mapped -ne $cat -and $mapped -notin $secondaryCats) {
                        $secondaryCats += $mapped
                    }
                    break
                }
            }
        }
    }

    [PSCustomObject]@{
        UPN      = $r.'User Principal Name'
        Name     = $r.'Display Name'
        Dept     = if ($r.'Department') { $r.'Department' } else { '(No Department)' }
        Cost     = [math]::Round($cost, 2)
        Savings  = [math]::Round($savings, 2)
        CompCost = [math]::Round($compCost, 2)
        Category = $cat
        Tags     = $secondaryCats
        Licenses = $r.'License Friendly Names'
        Rec      = $rec
        AdminPriv = if ($r.'Admin Privilege Level') { $r.'Admin Privilege Level' } else { '' }
        CpApps   = if ($r.'Copilot Active Apps') { $r.'Copilot Active Apps' } else { '' }
    }
}
$userData = @($userData)
Write-Host "  $($userData.Count) users with savings opportunities" -ForegroundColor Gray

# ── SKU waste rollup ──────────────────────────────────────────────────────────
# Categories where savings are attributable to a specific SKU family, not the full portfolio.
# For these, attribute 100% to matching SKUs; fall back to equal split if no match.
$categorySkuPattern = @{
    # CPC-specific savings → Windows 365 / Cloud PC SKUs only
    'Dormant Cloud PC'       = '(?i)Windows 365|Cloud PC'
    'Cloud PC Review'        = '(?i)Windows 365|Cloud PC'
    # Copilot savings → Copilot SKU only
    'Copilot Reclaim'        = '(?i)Copilot'
    # Add-on savings → specific add-on SKU only
    'Inactive Add-On'        = '(?i)Visio|Project|Planner|Power BI|Teams Premium'
    'Inactive Add-On Review' = '(?i)Visio|Project|Planner|Power BI|Teams Premium'
    'Premium Add-On Review'  = '(?i)Visio|Project|Planner|Power BI|Teams Premium'
    # Duplicate/overlap savings → M365/O365/standalone product SKUs (not CPC)
    'Duplicate Coverage'     = '(?i)Microsoft 365|Office 365|Business|Exchange|SharePoint|Visio|Project|Power BI|Entra|Intune|Defender|Teams Premium'
    'Overlapping License'    = '(?i)Microsoft 365|Office 365|Business|Exchange|SharePoint|Visio|Project|Power BI|Entra|Intune|Defender|Teams Premium'
    # Suite-specific savings → M365/O365 suite SKUs only (not CPC, not add-ons)
    'Teams Unbundling'       = '(?i)Microsoft 365|Office 365|Business'
    'E5 Voice Review'        = '(?i)Microsoft 365 E5|Office 365 E5'
    'Frontline Candidate'    = '(?i)Microsoft 365|Office 365|Business'
    'Frontline Rescue'       = '(?i)Microsoft 365|Office 365|Business|Exchange'
    'EXO Plan 2 Downgrade'   = '(?i)Exchange'
    'Exchange Kiosk Downgrade' = '(?i)Exchange'
}
$skuRollup = @{}
foreach ($u in $userData) {
    $allLicenses = @($u.Licenses -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    # Only attribute waste to paid SKUs — free/zero-cost licenses don't contribute to savings
    $licenses = @($allLicenses | Where-Object { $_ -notmatch '(?i)\bFree\b|\bTrial\b|\bDeveloper\b|\bExploratory\b|\bAdhoc\b|\bCredits\b|\bResource Account\b|\bClipchamp\b|\bTeams Rooms Basic\b' })
    if (-not $licenses) { $licenses = $allLicenses }  # fallback if all are free
    if (-not $licenses) { $licenses = @('Unknown') }
    # Targeted SKU attribution: for categories with savings tied to a specific product,
    # attribute 100% to matching SKUs only; fall back to equal split if no match
    $targetLicenses = $licenses
    if ($categorySkuPattern.ContainsKey($u.Category)) {
        $matched = @($licenses | Where-Object { $_ -match $categorySkuPattern[$u.Category] })
        if ($matched) { $targetLicenses = $matched }
    }
    $share = [math]::Round($u.Savings / $targetLicenses.Count, 2)
    if ($share -le 0) { continue }
    foreach ($lic in $targetLicenses) {
        if (-not $skuRollup.ContainsKey($lic)) {
            $skuRollup[$lic] = @{ License=$lic; Waste=[decimal]0; Users=0; Categories=@{} }
        }
        $skuRollup[$lic].Waste += $share
        $skuRollup[$lic].Users++
        $c = $u.Category
        if (-not $skuRollup[$lic].Categories.ContainsKey($c)) { $skuRollup[$lic].Categories[$c] = [decimal]0 }
        $skuRollup[$lic].Categories[$c] += $share
    }
}
$skuData = @($skuRollup.Values | Sort-Object Waste -Descending | Select-Object -First 20)

# ── Quick-win category tiles ──────────────────────────────────────────────────
# CatKey  : regex applied to per-user 'Recommendation Category' field
# RecKey  : regex applied to per-user 'Recommendation' text (catches cross-category findings)
# Tile count and savings are ALWAYS computed from per-row matching (same users shown on click)
$tileDefs = @(
    # ── Tier 1: License review (full license cost reclaimable) ───────────────
    [PSCustomObject]@{ Label='Dormant Accounts';       Desc='No sign-in >30 days';             CatKey='^dormant$';                        RecKey='';                    Color='#3ddad7'; Tier=1 }
    [PSCustomObject]@{ Label='Disabled Accounts';      Desc='Sign-in blocked';                  CatKey='^disabled';                        RecKey='';                    Color='#3ddad7'; Tier=1 }
    [PSCustomObject]@{ Label='Never Signed In';        Desc='No interactive sign-in on record'; CatKey='never.signed';                     RecKey='NEVER SIGNED IN';     Color='#3ddad7'; Tier=1 }
    [PSCustomObject]@{ Label='Zero M365 Usage';        Desc='No app activity in period';        CatKey='no.activity|zero.*usage';          RecKey='NO ACTIVITY detected'; Color='#3ddad7'; Tier=1 }
    [PSCustomObject]@{ Label='Admin Review';           Desc='Admin with productivity license';  CatKey='^admin review$';                   RecKey='';                    Color='#3ddad7'; Tier=1 }
    [PSCustomObject]@{ Label='Shared Mailbox';         Desc='Licensed shared mailbox, review'; CatKey='shared.mailbox';                RecKey='';                    Color='#3ddad7'; Tier=1 }
    [PSCustomObject]@{ Label='Guest w/ Paid Licenses'; Desc='B2B guest holding a paid license'; CatKey='guest';                            RecKey='';                    Color='#3ddad7'; Tier=1 }
    [PSCustomObject]@{ Label='Automation Accounts';    Desc='Service/automation account';       CatKey='^automation.account$';             RecKey='';                    Color='#3ddad7'; Tier=1 }
    [PSCustomObject]@{ Label='Dormant Admin Accounts'; Desc='Admin with no sign-in detected';   CatKey='dormant.admin';                    RecKey='';                    Color='#3ddad7'; Tier=1 }
    [PSCustomObject]@{ Label='Dormant Cloud PC';       Desc='0 hours connected in 90 days';     CatKey='^dormant.cloud.pc$';               RecKey='';                    Color='#3ddad7'; Tier=1 }
    [PSCustomObject]@{ Label='Cloud PC Review';        Desc='< 10 hrs connected in 90 days';    CatKey='^cloud.pc.review$';                RecKey='';                    Color='#3ddad7'; Tier=1 }
    [PSCustomObject]@{ Label='Inactive Products';       Desc='No activation detected';          CatKey='^inactive.add-on$|^visio|^project|^power.bi.pro|^pbi.ppu'; RecKey=''; Color='#3ddad7'; Tier=1 }
    [PSCustomObject]@{ Label='Product Review';          Desc='Web-only, verify usage';          CatKey='^inactive.add-on.review$';                         RecKey=''; Color='#5b8def'; Tier=2 }
    [PSCustomObject]@{ Label='Right-Sizing Opportunities'; Desc='Desktop unused, web/mobile only'; CatKey='^premium.add-on.review$';                          RecKey=''; Color='#5b8def'; Tier=2 }
    [PSCustomObject]@{ Label='Copilot Reclaim';        Desc='Zero usage & zero readiness';      CatKey='^copilot.reclaim$';               RecKey='(^|\| )COPILOT RECLAIM';     Color='#3ddad7'; Tier=1 }
    [PSCustomObject]@{ Label='Expensive Cold Storage'; Desc='E5 retained only for archive/hold'; CatKey='expensive.cold';                  RecKey='EXPENSIVE COLD';       Color='#3ddad7'; Tier=1 }
    [PSCustomObject]@{ Label='Background Sync Only';   Desc='Zero interactive activity, OneDrive syncing'; CatKey='background.sync';        RecKey='BACKGROUND SYNC';      Color='#3ddad7'; Tier=1 }

    # ── Tier 2: Right-sizing (partial savings via downgrade/swap) ────────────
    [PSCustomObject]@{ Label='Duplicate Coverage';     Desc='Standalone covered by suite';      CatKey='^duplicate.coverage$';             RecKey='';                    Color='#5b8def'; Tier=2 }
    [PSCustomObject]@{ Label='Duplicate Review';       Desc='Possible duplicate, needs review'; CatKey='^duplicate.review$';               RecKey='';                    Color='#5b8def'; Tier=2 }
    [PSCustomObject]@{ Label='Overlapping License';    Desc='Same license via multiple paths';  CatKey='overlapping';                      RecKey='';                    Color='#5b8def'; Tier=2 }
    [PSCustomObject]@{ Label='Standalone Licenses';    Desc='Standalone included in suite';     CatKey='standalone';                       RecKey='';                    Color='#5b8def'; Tier=2 }
    [PSCustomObject]@{ Label='Teams Unbundling';       Desc='Suite bundles Teams, no usage';    CatKey='teams.unbundling';                 RecKey='';                    Color='#5b8def'; Tier=2 }
    [PSCustomObject]@{ Label='E5 Voice Review';         Desc='E5 with no calling/conferencing';  CatKey='e5.voice';                         RecKey='';                    Color='#5b8def'; Tier=2 }
    [PSCustomObject]@{ Label='Bundle Opportunity';      Desc='Standalone apps cheaper as suite'; CatKey='bundle.opportunity';               RecKey='';                    Color='#5b8def'; Tier=2 }
    [PSCustomObject]@{ Label='Exchange Kiosk Downgrade'; Desc='Web-only usage, <2 GB mailbox';  CatKey='exchange.kiosk';                   RecKey='EXCHANGE KIOSK';       Color='#5b8def'; Tier=2 }
    [PSCustomObject]@{ Label='Forwarding Mailbox Review'; Desc='Mailbox forwarding all mail';    CatKey='forwarding.mailbox.review';        RecKey='FORWARDING MAILBOX';   Color='#5b8def'; Tier=2 }
    [PSCustomObject]@{ Label='Copilot At Risk';        Desc='Zero usage, active in M365';       CatKey='copilot.watchlist';            RecKey='(^|\| )COPILOT WATCHLIST';   Color='#5b8def'; Tier=2 }

    # ── Tier 3: Compliance & review (no direct savings) ──────────────────────
    [PSCustomObject]@{ Label='Licensing Compliance';   Desc='Policy/entitlement gap detected';  CatKey='licensing.compliance|compliance.gap'; RecKey='';                  Color='#ff9f80'; Tier=3 }
    [PSCustomObject]@{ Label='Data Gap';               Desc='Unknown SKU, incomplete analysis'; CatKey='data.gap';                         RecKey='';                    Color='#ff9f80'; Tier=3 }
    [PSCustomObject]@{ Label='Mailbox Storage Warning'; Desc='Mailbox near capacity limit';     CatKey='mailbox.storage';                  RecKey='';                    Color='#ff9f80'; Tier=3 }
    [PSCustomObject]@{ Label='Unlicensed With Data';   Desc='No license but has mailbox data';  CatKey='unlicensed.with.data';             RecKey='';                    Color='#ff9f80'; Tier=3 }
    [PSCustomObject]@{ Label='Free License Overlap';    Desc='Self-service trial/free licenses'; CatKey='free.license';                     RecKey='';                    Color='#ff9f80'; Tier=3 }
    [PSCustomObject]@{ Label='Windows License Review';  Desc='Windows E3/E5 with no sign-in';    CatKey='windows.license';                  RecKey='';                    Color='#ff9f80'; Tier=3 }
)

# ── Dynamic tile generation: catch any category not covered by a well-known tile ─
$_skipCats = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
@('No Findings','','Unlicensed','Copilot Active') | ForEach-Object { [void]$_skipCats.Add($_) }

# Meaningful subtitles for categories that don't have a predefined tile
$_autoTileDesc = @{
    'Frontline Candidate'              = 'E3/E5 user, web/mobile only'
    'Frontline Review'                 = 'May qualify for F3 downgrade'
    'Frontline Blocked'                = 'Frontline profile, blocker detected'
    'Frontline Rescue'                 = 'Archive blocks F3, E1/Basic viable'
    'Frontline Add-On Stacking'        = 'F-series + add-ons exceed E3 cost'
    'Suite Inversion'                  = 'E3 + add-ons cost more than E5'
    'E5 Upgrade'                       = 'E3 + add-ons, E5 simplifies'
    'EXO Plan 2 Downgrade'             = 'Plan 2 user under 50 GB'
    'EXO Plan 2 Review'                = 'Plan 2, usage data missing'
    'Entra P2 Downgrade'               = 'Standalone P2, no PIM/risk-CA usage'
    'Entra Suite Overlap'              = 'Entra add-ons covered by suite'
    'Intune Suite Overlap'             = 'Intune add-ons covered by suite'
    'Intune Review'                    = 'Intune assigned, 0 enrolled devices'
    'E1 to Business Basic'             = 'E1 eligible for cheaper Business Basic'
    'E3 to Business Premium'           = 'E3 eligible for Business Premium'
    'O365 E3 to E1'                    = 'O365 E3 with no desktop app usage'
    'Business Downgrade'               = 'Premium suite, low feature usage'
    'Business Review'                  = 'Possible downgrade, needs review'
    'Business Premium Inversion'       = 'Business Standard + add-ons exceed Premium'
    'Business Premium Security Review' = 'Premium security features unused'
    'Bundle Consolidation'             = 'Multiple SKUs replaceable by one suite'
    'App Arbitrage'                    = 'Standalone apps cheaper than suite'
    'F3 to F1 Downgrade'              = 'F3 with zero email/OneDrive usage'
    'Inactive Hold With License'       = 'Disabled + hold, license not needed'
    'Inactive Hold'                    = 'Disabled + hold, free SKU sufficient'
    'Inactive Mailbox'                 = 'Unlicensed inactive mailbox'
    'Litigation Hold'                  = 'Active hold on mailbox'
    'Room/Equipment'                   = 'Room or equipment mailbox'
    'Non-Human Account Review'         = 'Shared/room with premium suite'
    'Legacy Service Account'           = 'Legacy service account pattern'
    'Stale Sign-In'                    = 'No sign-in but has M365 activity'
    'No Desktop'                       = 'Web-only Office usage detected'
    'Mobile Only'                      = 'Mobile-only Office usage detected'
    'Copilot'                          = 'Copilot license holder'
    'Copilot Watchlist'                = 'Copilot usage declining'
    'Copilot Prerequisite'             = 'Missing prerequisite for Copilot'
    'Copilot Studio'                   = 'Copilot Studio license holder'
    'AI Add-On Overlap'                = 'AI add-on covered by existing suite'
    'AI Overlap Review'                = 'Possible AI add-on overlap'
    'Teams Phone Review'               = 'Phone license, no calling activity'
    'Teams Phone Right-Sizing'         = 'Phone plan exceeds actual usage'
    'Calling Plan Review'              = 'Calling plan with low utilisation'
    'Power BI Pro Review'              = 'Power BI Pro with zero report views'
    'OneDrive Plan 2 Review'           = 'OD Plan 2, under 1 TB usage'
    'OneDrive Storage Warning'         = 'OneDrive approaching capacity'
    'Over-Licensed Archive'            = 'Archive on plan exceeding needs'
    'Redundant Archive'                = 'Archive add-on covered by suite'
    'Seeded Visio Overlap'             = 'Seeded Visio in suite, standalone too'
    'PBI PPU Overlap'                  = 'PPU overlaps with existing license'
    'Trial License'                    = 'Trial/preview SKU still assigned'
    'License Error'                    = 'License assignment error detected'
    'Cloud License Error'              = 'Cloud licensing sync failure'
    'License Capacity'                 = 'SKU approaching seat limit'
    'External Sharing Review'          = 'External sharing enabled, verify need'
    'Security Gap'                     = 'Missing security coverage detected'
    'Defender Coverage Review'         = 'Defender coverage incomplete'
    'Compliance Coverage Review'       = 'Compliance coverage incomplete'
    'Partial Optimization'             = 'Minor optimization opportunity'
}

$liveCategories = @($rows | ForEach-Object { $_.'Recommendation Category' } |
    Where-Object { $_ -and -not $_skipCats.Contains($_) } | Sort-Object -Unique)

foreach ($cat in $liveCategories) {
    $covered = $false
    foreach ($def in $tileDefs) {
        if ($def.CatKey -and $cat -match $def.CatKey) { $covered = $true; break }
    }
    if (-not $covered) {
        $tileDefs += [PSCustomObject]@{
            Label  = $cat
            Desc   = if ($_autoTileDesc.ContainsKey($cat)) { $_autoTileDesc[$cat] } else { 'Auto-detected category' }
            CatKey = '^' + [regex]::Escape($cat) + '$'
            RecKey = ''
            Color  = '#5b8def'
            Tier   = 2
        }
        Write-Host "    + Auto-tile: $cat" -ForegroundColor DarkGray
    }
}

$tileData = @($tileDefs | ForEach-Object {
    $def = $_

    # Per-row matching: CatKey matches by primary category (for savings),
    # RecKey matches by recommendation text (for display only — no savings).
    # This ensures each user's savings appear in exactly one tile while
    # secondary findings remain visible in drill-down.
    $primaryMatched = @($rows | Where-Object {
        $cat = $_.'Recommendation Category'
        if ($_skipCats.Contains($cat)) { return $false }
        if ($def.CatKey -and $cat -match $def.CatKey) { return $true }
        return $false
    })
    $secondaryMatched = @()
    if ($def.RecKey) {
        $primaryUpns = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($pm in $primaryMatched) { [void]$primaryUpns.Add($pm.'User Principal Name') }
        $secondaryMatched = @($rows | Where-Object {
            $cat = $_.'Recommendation Category'
            if ($_skipCats.Contains($cat)) { return $false }
            if ($primaryUpns.Contains($_.'User Principal Name')) { return $false }
            if ($_.'Recommendation' -match $def.RecKey) { return $true }
            return $false
        })
    }
    $matched = @($primaryMatched) + @($secondaryMatched)

    # Savings: only sum from primary matches (prevents double-counting across tiles)
    $tileAmount = [decimal]0
    foreach ($m in $primaryMatched) {
        $tileAmount += if ($m.'Estimated Annual Savings (EUR)') { Parse-Decimal $m.'Estimated Annual Savings (EUR)' } else { [decimal]0 }
    }

    [PSCustomObject]@{
        label   = $def.Label
        desc    = $def.Desc
        color   = $def.Color
        key     = $def.CatKey
        recKey  = $def.RecKey
        users   = $matched.Count
        savings = [math]::Round($tileAmount, 0)
        tier    = if ($def.PSObject.Properties['Tier']) { $def.Tier } else { 2 }
    }
})

# Sort tiles: active findings first (by savings desc), zero-count tiles at end
$tileData = @($tileData |
    Sort-Object @{Expression={if ($_.users -gt 0 -or $_.savings -gt 0) { 0 } else { 1 }}},
               @{Expression='savings'; Descending=$true},
               @{Expression='users'; Descending=$true})

# ── Unassigned Licenses tile (Pool waste — no per-user rows) ─────────────────
$poolDataRows = @($summaryRows | Where-Object { $_.'Tier' -eq 'Pool' -and $_.'Category' -notmatch 'TOTAL' })
$unassignedWaste = [decimal]0
$unassignedSKUs  = 0
foreach ($pr in $poolDataRows) {
    $w = Parse-Decimal $pr.'Annual Amount (EUR)'
    if ($w -gt 0) { $unassignedWaste += $w; $unassignedSKUs++ }
}
$poolTile = [PSCustomObject]@{
    label   = 'Unassigned Licenses'
    desc    = 'Licenses not assigned to any user'
    color   = '#3ddad7'
    key     = ''
    recKey  = ''
    users   = $unassignedSKUs
    savings = [math]::Round($unassignedWaste, 0)
    tier    = 1
}
$tileData = @($poolTile) + @($tileData)

# ── Per-user capability data (Tab 3) ─────────────────────────────────────────
$rowsByUpn = @{}
foreach ($r in $rows) {
    $upn = $r.'User Principal Name'
    if ($upn -and -not $rowsByUpn.ContainsKey($upn)) { $rowsByUpn[$upn] = $r }
}

$capUsers = @($userData | Sort-Object Savings -Descending | ForEach-Object {
    $u = $_
    if ($rowsByUpn.ContainsKey($u.UPN)) {
        $r = $rowsByUpn[$u.UPN]
        [PSCustomObject]@{
            n   = $u.Name
            upn = $u.UPN
            cat  = $u.Category
            tags = $u.Tags
            sav  = $u.Savings
            comp = $u.CompCost
            ex   = ($r.'Has Exchange License'   -eq 'True')
            exU = ($r.'Exchange Intensity'     -and $r.'Exchange Intensity'   -notmatch '^(Low|None|)$')
            exI = if ($r.'Exchange Intensity') { $r.'Exchange Intensity' } else { '' }
            tm  = ($r.'Has Teams License'      -eq 'True')
            tmU = ($r.'Teams Intensity'        -and $r.'Teams Intensity'      -notmatch '^(Low|None|)$')
            tmI = if ($r.'Teams Intensity') { $r.'Teams Intensity' } else { '' }
            dt  = ($r.'No Desktop Apps'        -ne 'True' -and ($r.'Has Exchange License' -eq 'True' -or $r.'Has Teams License' -eq 'True' -or $r.'Has OneDrive License' -eq 'True' -or $r.'Has SharePoint License' -eq 'True'))
            dtU = ($r.'Uses Desktop Apps'      -eq 'True')
            od  = ($r.'Has OneDrive License'   -eq 'True')
            odU = ($r.'OneDrive Intensity'     -and $r.'OneDrive Intensity'   -notmatch '^(Low|None|)$')
            odI = if ($r.'OneDrive Intensity') { $r.'OneDrive Intensity' } else { '' }
            sp  = ($r.'Has SharePoint License' -eq 'True')
            spU = ($r.'SharePoint Intensity'   -and $r.'SharePoint Intensity' -notmatch '^(Low|None|)$')
            spI = if ($r.'SharePoint Intensity') { $r.'SharePoint Intensity' } else { '' }
            co  = [bool]($r.'License Friendly Names' -match 'Copilot')
            coU = [bool]($r.'Copilot Active Apps'    -and $r.'Copilot Active Apps' -ne '')
            adm = if ($r.'Admin Privilege Level') { $r.'Admin Privilege Level' } else { '' }
        }
    }
})

# ── KPI extraction from executive summary ─────────────────────────────────────
$kpiTotalSpend  = [decimal]0
$kpiSavingsPot  = [decimal]0
$kpiTotalUsers  = $rows.Count
$kpiWithRec     = @($rows | Where-Object { $_.'Recommendation Category' -ne 'No Findings' -and $_.'Recommendation Category' -ne '' -and $_.'Recommendation Category' -ne 'Unlicensed' -and $_.'Recommendation Category' -ne 'Copilot Active' }).Count

if ($summaryRows) {
    $ovTotalSpend = $summaryRows | Where-Object { $_.'Category' -match 'Total Annual M365 Spend' }
    if ($ovTotalSpend) { $kpiTotalSpend = Parse-Decimal ($ovTotalSpend | Select-Object -First 1).'Annual Amount (EUR)' }
}
if ($kpiTotalSpend -eq 0) {
    $kpiTotalSpend = ($rows | ForEach-Object { Parse-Decimal $_.'Annual License Cost (EUR)' } | Measure-Object -Sum).Sum
}
# Derive headline savings from tile sums — single source of truth for drill-down consistency
# Includes pool waste tile (unassigned licenses) which has no per-user rows in $userData
$kpiSavingsPot = [decimal]($tileData | Measure-Object -Property savings -Sum).Sum
$kpiSavingsPct = if ($kpiTotalSpend -gt 0) { [math]::Round($kpiSavingsPot / $kpiTotalSpend * 100, 1) } else { 0 }
$kpiCompCost  = [decimal]($userData | Measure-Object -Property CompCost -Sum).Sum
$kpiCompUsers = @($userData | Where-Object { $_.CompCost -gt 0 }).Count

# ── JSON helpers ──────────────────────────────────────────────────────────────
function To-JsonString([object]$obj) {
    return (ConvertTo-Json -InputObject $obj -Depth 5 -Compress)
}

# ── Prepare JS data ───────────────────────────────────────────────────────────
$topUsers = @($userData | Sort-Object Savings -Descending |
    Select-Object Name, UPN, Dept, Cost, Savings, CompCost, Category, Tags, Licenses, Rec, AdminPriv, CpApps)

$skuJs = @($skuData | ForEach-Object {
    $catArr = @($_.Categories.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 5 |
        ForEach-Object { @{ cat=$_.Key; val=[math]::Round($_.Value,2) } })
    [PSCustomObject]@{
        lic   = $_.License
        waste = [math]::Round($_.Waste, 2)
        users = $_.Users
        cats  = $catArr
    }
})

$poolSkuRows = @($summaryRows | Where-Object { $_.'Tier' -eq 'Pool' -and $_.'Category' -notmatch 'TOTAL' } | ForEach-Object {
    [PSCustomObject]@{
        sku        = ($_.'Category' -replace '\s*\(\d+/\d+\s+unassigned\)\s*$','').Trim()
        unassigned = [int](($_.'Users' -replace '\D','') -replace '^$','0')
        waste      = [math]::Round((Parse-Decimal $_.'Annual Amount (EUR)'), 2)
    }
})

# Pre-compute per-SKU: total group members and overlaps (for direct-only calculation)
$_skuGroupTotals = @{}
foreach ($gr in $groupRows) {
    $sk = ($gr.'Assigned Licenses' -replace '\s*\(\d+ plans? disabled\)', '').Trim()
    if (-not $_skuGroupTotals.ContainsKey($sk)) { $_skuGroupTotals[$sk] = @{ Members = 0; Overlap = 0 } }
    $_skuGroupTotals[$sk].Members += [int]($gr.'Member Count' -replace '\D','')
    $_skuGroupTotals[$sk].Overlap += [int]($gr.'Also Direct' -replace '\D','')
}

$groupJs = @($groupRows | ForEach-Object {
    $rawSku   = $_.'Assigned Licenses'
    # Strip plan-disabled annotations like "SPE_E5 (1 plans disabled)" → "SPE_E5"
    $skuClean = ($rawSku -replace '\s*\(\d+ plans? disabled\)', '').Trim()
    $skuSeats = ''
    $directOnly = -1  # -1 = unknown
    if ($skuInvLookup.ContainsKey($skuClean)) {
        $inv = $skuInvLookup[$skuClean]
        $skuSeats = "$($inv.Consumed)/$($inv.Total)"
        # Direct-only = consumed - (group members across all groups for this SKU)
        # Member count already includes overlap users (they are group members who ALSO have direct)
        if ($_skuGroupTotals.ContainsKey($skuClean)) {
            $directOnly = [Math]::Max(0, $inv.Consumed - $_skuGroupTotals[$skuClean].Members)
        }
    }
    [PSCustomObject]@{
        name      = $_.'Group Name'
        type      = $_.'Membership Type'
        members   = [int]($_.'Member Count' -replace '\D','')
        direct    = [int]($_.'Also Direct' -replace '\D','')
        skus      = $rawSku
        skuSeats  = $skuSeats
        directOnly = $directOnly
        count     = [int]($_.'License Count' -replace '\D','')
    }
} | Sort-Object { $_.members } -Descending)

# ── Subscription lifecycle alerts from exec CSV ──────────────────────────────
$subAlerts = @($summaryRows | Where-Object { $_.'Tier' -eq 'Sub' } | ForEach-Object {
    $daysStr = $_.'Pct of Spend' -replace '[^-\d]',''
    [PSCustomObject]@{
        sku    = $_.'Category'
        seats  = $_.'Users'          # e.g. "487/497"
        status = $_.'Annual Amount (EUR)'    # "Suspended", "Warning", "Enabled"
        days   = if ($daysStr -ne '') { [int]$daysStr } else { 0 }  # guard empty-string coercion
    }
} | Sort-Object days)

# ── Department list for filter ───────────────────────────────────────────────
$allDepts = @($userData | Where-Object { $_.Dept -and $_.Dept -ne '(No Department)' } | ForEach-Object { $_.Dept } | Sort-Object -Unique)

# ── Copilot ROI data from exec CSV ───────────────────────────────────────────
$copilotRoi = @{}
foreach ($cr in @($summaryRows | Where-Object { $_.'Tier' -eq 'Copilot' })) {
    $copilotRoi[$cr.'Category'] = $cr.'Users'
}
# Copilot per-app adoption (from userData)
$copilotHolders = @($rows | Where-Object { $_.'Assigned Licenses' -match 'Microsoft_365_Copilot|Copilot' -and $_.'Assigned Licenses' -ne '[UNLICENSED]' })
$cpAppStats = @{ Teams=0; Word=0; Excel=0; PowerPoint=0; Outlook=0; OneNote=0; Loop=0; Chat=0 }
foreach ($ch in $copilotHolders) {
    $apps = $ch.'Copilot Active Apps'
    if ($apps) {
        if ($apps -match 'Teams')      { $cpAppStats.Teams++ }
        if ($apps -match 'Word')       { $cpAppStats.Word++ }
        if ($apps -match 'Excel')      { $cpAppStats.Excel++ }
        if ($apps -match 'PowerPoint') { $cpAppStats.PowerPoint++ }
        if ($apps -match 'Outlook')    { $cpAppStats.Outlook++ }
        if ($apps -match 'OneNote')    { $cpAppStats.OneNote++ }
        if ($apps -match 'Loop')       { $cpAppStats.Loop++ }
        if ($apps -match 'Chat')       { $cpAppStats.Chat++ }
    }
}
$cpHolderJs = @($copilotHolders | ForEach-Object {
    [PSCustomObject]@{
        Name   = $_.'Display Name'
        UPN    = $_.'User Principal Name'
        Dept   = if ($_.'Department') { $_.'Department' } else { '' }
        CpApps = if ($_.'Copilot Active Apps') { $_.'Copilot Active Apps' } else { '' }
        AdminPriv = if ($_.'Admin Privilege Level') { $_.'Admin Privilege Level' } else { '' }
    }
})
$jsCpHolders = To-JsonString $cpHolderJs

$copilotRoiData = [PSCustomObject]@{
    total     = $copilotHolders.Count
    active    = if ($copilotRoi.ContainsKey('Active Users (Copilot usage detected)')) { [int]$copilotRoi['Active Users (Copilot usage detected)'] } else { 0 }
    watchlist = if ($copilotRoi.ContainsKey('Watchlist (zero Copilot, active M365)')) { [int]$copilotRoi['Watchlist (zero Copilot, active M365)'] } else { 0 }
    reclaim   = if ($copilotRoi.ContainsKey('Reclaim (zero Copilot, zero M365)'))    { [int]$copilotRoi['Reclaim (zero Copilot, zero M365)'] }    else { 0 }
    apps      = $cpAppStats
}

$jsTopUsers  = To-JsonString $topUsers
$jsSkuData   = To-JsonString $skuJs
$jsTileData  = To-JsonString $tileData
$jsCapUsers  = To-JsonString $capUsers
$jsPoolSkus  = To-JsonString $poolSkuRows
$jsAllCats   = To-JsonString @($userData | ForEach-Object { $_.Category } | Sort-Object -Unique)
$jsGroups    = To-JsonString $groupJs
$jsSubAlerts = To-JsonString $subAlerts
$jsDepts     = To-JsonString $allDepts
$jsCopilotRoi = To-JsonString $copilotRoiData

$reportDate = (Get-Item $ReportCsv).LastWriteTime.ToString("dd MMM yyyy")
$genDate    = (Get-Date).ToString("dd MMM yyyy HH:mm")

# ── HTML generation ───────────────────────────────────────────────────────────
Write-Host "  Building HTML..." -ForegroundColor Gray

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>M365 License Optimization Assessment</title>
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;600;700&family=Sora:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
:root{
  --p-dark-navy:#111125;--p-purple:#48349a;--p-teal:#3ddad7;--p-pink:#ef6ea7;--p-peach:#ff9f80;--p-steel:#5b89b6;
  --navy-surface:#181835;--navy-card:#1e1e42;--navy-border:#2a2a55;
  --purple-dim:#362878;--purple-glow:rgba(72,52,154,.35);
  --teal-dim:rgba(61,218,215,.12);--teal-glow:rgba(61,218,215,.25);
  --pink-dim:rgba(239,110,167,.12);--pink-glow:rgba(239,110,167,.25);
  --peach-dim:rgba(255,159,128,.12);--steel-dim:rgba(91,137,182,.12);
  --text-primary:#eeeef5;--text-secondary:#9898b8;--text-dim:#6a6a8e;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Sora',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:var(--p-dark-navy);color:var(--text-primary);font-size:14px;line-height:1.6}
body::after{content:'';position:fixed;inset:0;pointer-events:none;z-index:9999;opacity:.025;background-image:url("data:image/svg+xml,%3Csvg viewBox='0 0 256 256' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.85' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)'/%3E%3C/svg%3E");background-size:128px 128px}
header{position:relative;background:var(--navy-surface);color:#fff;padding:32px 32px 24px;overflow:hidden}
header::before{content:'';position:absolute;inset:0;background:radial-gradient(ellipse 700px 500px at 25% 20%,var(--purple-glow) 0%,transparent 70%),radial-gradient(ellipse 500px 400px at 75% 70%,rgba(61,218,215,.08) 0%,transparent 70%),radial-gradient(ellipse 400px 300px at 50% 90%,rgba(239,110,167,.05) 0%,transparent 70%);pointer-events:none}
header h1{font-family:'Sora',sans-serif;font-size:24px;font-weight:700;letter-spacing:.3px;position:relative}
header p{margin-top:4px;color:var(--text-secondary);font-size:13px;position:relative}
.disclaimer{font-size:12px;color:var(--text-secondary);margin-top:14px;font-weight:500;position:relative;overflow:hidden;white-space:nowrap}
.disclaimer-track{display:inline-flex;animation:tickerScroll 100s linear infinite}
.disclaimer-track span{padding-right:20em;flex-shrink:0}
@keyframes tickerScroll{0%{transform:translateX(0)}100%{transform:translateX(-50%)}}
.kpis{display:flex;gap:16px;margin-top:22px;flex-wrap:wrap;position:relative}
.kpi{background:var(--navy-card);border:1px solid var(--navy-border);border-radius:12px;padding:16px 22px;min-width:160px;flex:1}
.kpi .label{font-size:11px;color:var(--text-dim);text-transform:uppercase;letter-spacing:.5px}
.kpi .value{font-family:'JetBrains Mono',monospace;font-size:26px;font-weight:700;margin-top:4px}
.kpi .sub{font-size:11px;color:var(--text-dim);margin-top:2px}
.kpi.alert .value{color:var(--p-pink)}
.kpi.good .value{color:var(--p-teal)}
.tabs{display:flex;gap:0;padding:0 32px;background:var(--navy-surface);border-bottom:1px solid var(--navy-border);position:sticky;top:0;z-index:10;box-shadow:0 2px 12px rgba(0,0,0,.3)}
.tab-btn{padding:14px 24px;cursor:pointer;font-size:13px;font-weight:500;color:var(--text-dim);border:none;background:none;border-bottom:3px solid transparent;margin-bottom:-1px;transition:all .2s;font-family:'Sora',sans-serif}
.tab-btn:hover{color:var(--text-primary)}
.tab-btn.active{color:var(--p-teal);border-bottom-color:var(--p-teal);font-weight:600}
.panel{display:none;padding:24px 32px}
.panel.active{display:block}
h2{font-family:'Sora',sans-serif;font-size:16px;font-weight:600;color:var(--p-teal);margin-bottom:4px}
.section-desc{font-size:12px;color:var(--text-dim);margin-bottom:20px}
.card{background:var(--navy-card);border:1px solid var(--navy-border);border-radius:12px;padding:20px;margin-bottom:20px;box-shadow:0 4px 20px rgba(0,0,0,.2)}
.card h3{font-size:14px;font-weight:600;margin-bottom:14px;color:var(--text-primary)}
/* Dashboard tiles */
.dash-section-title{font-family:'Sora',sans-serif;font-size:15px;font-weight:600;color:var(--p-teal);margin-bottom:16px;padding-top:4px}
.bd-row{display:flex;align-items:center;gap:10px;padding:6px 0;border-bottom:1px solid var(--navy-border)}
.bd-row:hover{background:rgba(255,255,255,.03);border-radius:4px}
.bd-label{width:180px;font-size:13px;color:var(--text-secondary);flex-shrink:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.bd-track{flex:1;height:22px;background:rgba(255,255,255,.06);border-radius:4px;overflow:hidden}
.bd-fill{height:100%;border-radius:4px;transition:width .3s ease}
.bd-amt{width:90px;text-align:right;font-size:13px;font-weight:600;color:var(--text-primary);font-family:'JetBrains Mono',monospace;flex-shrink:0}
.tier-legend{font-size:12px;font-weight:400;color:var(--text-secondary);margin-left:auto;display:flex;align-items:center;white-space:nowrap}
.tier-dot{display:inline-block;width:10px;height:10px;border-radius:50%;vertical-align:middle;margin-right:5px}
.dt-tier-dot{position:absolute;top:8px;right:8px;width:8px;height:8px;border-radius:50%;opacity:.85}
.dash-tiles-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:14px;margin-bottom:24px}
.dash-tile{background:var(--navy-card);border:1px solid var(--navy-border);border-radius:12px;padding:20px 22px;cursor:pointer;transition:all .18s;border-top:4px solid transparent;box-shadow:0 2px 10px rgba(0,0,0,.2);position:relative;overflow:hidden}
.dash-tile:hover{transform:translateY(-3px);box-shadow:0 8px 28px rgba(0,0,0,.35);border-color:rgba(255,255,255,.08)}
.dash-tile.dt-active{box-shadow:0 8px 32px rgba(0,0,0,.4)}
.dash-tile.dt-zero{opacity:.35}
.dash-tile .dt-label{font-size:13px;font-weight:600;margin-bottom:3px;color:var(--text-primary)}
.dash-tile .dt-desc{font-size:11px;color:var(--text-dim);margin-bottom:14px;line-height:1.4}
.dash-tile .dt-count{font-family:'JetBrains Mono',monospace;font-size:32px;font-weight:700;line-height:1}
.dash-tile .dt-savings{font-family:'JetBrains Mono',monospace;font-size:13px;font-weight:600;margin-top:5px}
.dash-tile .dt-bar{height:3px;border-radius:2px;margin-top:14px;opacity:.35}
/* Small tile cards (legacy, kept for compat) */
.tiles-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:12px;margin-bottom:24px}
.tile-card{background:var(--navy-card);border:1px solid var(--navy-border);border-radius:10px;padding:16px 18px;cursor:pointer;transition:all .18s;border-left:4px solid transparent;box-shadow:0 2px 8px rgba(0,0,0,.2);position:relative}
.tile-card:hover{transform:translateY(-2px);box-shadow:0 6px 20px rgba(0,0,0,.3)}
.tile-card.t-active{box-shadow:0 6px 24px rgba(0,0,0,.35)}
.tile-card .t-label{font-size:12px;font-weight:600;margin-bottom:2px;color:var(--text-primary)}
.tile-card .t-desc{font-size:11px;color:var(--text-dim);margin-bottom:10px}
.tile-card .t-count{font-family:'JetBrains Mono',monospace;font-size:26px;font-weight:700;line-height:1.1}
.tile-card .t-savings{font-size:12px;font-weight:500;margin-top:3px;opacity:.85}
.tile-card .t-zero{opacity:.35}
/* User table */
.tbl-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:13px}
th{background:var(--navy-surface);color:var(--text-secondary);font-weight:600;padding:10px 12px;text-align:left;border-bottom:1px solid var(--navy-border);white-space:nowrap;cursor:pointer;user-select:none}
th:hover{background:var(--purple-dim)}
th .sort-icon{font-size:10px;margin-left:4px;opacity:.4}
th.sorted .sort-icon{opacity:1}
td{padding:9px 12px;border-bottom:1px solid rgba(255,255,255,.04);vertical-align:middle;color:var(--text-primary)}
tr.clickable-row{cursor:pointer}
tr.clickable-row:hover td{background:rgba(61,218,215,.06)}
.savings-cell{font-family:'JetBrains Mono',monospace;font-weight:600;border-radius:4px;padding:3px 8px;display:inline-block;font-size:12px}
.compcost-cell{font-family:'JetBrains Mono',monospace;font-weight:600;border-radius:4px;padding:3px 8px;display:inline-block;font-size:12px;background:var(--p-peach);color:#0a1628}
.cat-badge{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:500;background:var(--purple-dim);color:#d0d0e8}
/* SKU bars */
.sku-row{display:flex;align-items:center;gap:12px;margin-bottom:10px}
.sku-name{width:220px;font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex-shrink:0;color:var(--text-secondary)}
.sku-bar-wrap{flex:1;background:rgba(255,255,255,.06);border-radius:4px;height:22px;overflow:hidden;position:relative}
.sku-bar{height:100%;border-radius:4px;display:flex;align-items:center;padding-left:8px;font-size:11px;font-weight:600;color:#fff;transition:width .6s ease}
.sku-amount{width:90px;font-size:12px;font-weight:600;text-align:right;flex-shrink:0;font-family:'JetBrains Mono',monospace}
.sku-users{width:60px;font-size:11px;color:var(--text-dim);text-align:right;flex-shrink:0}
/* Capability matrix */
.cap-table{width:100%;border-collapse:collapse;font-size:13px}
.cap-table th{text-align:center}
.cap-table th.user-col{text-align:left;min-width:160px}
.cap-table td{text-align:center}
.cap-table td.user-name{text-align:left}
.cap-cell{border-radius:4px;padding:3px 4px;font-size:10px;font-weight:600;display:inline-block;min-width:42px}
/* Modal */
.modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.65);z-index:1000;align-items:center;justify-content:center}
.modal-overlay.open{display:flex}
.modal-box{background:var(--navy-card);border:1px solid var(--navy-border);border-radius:14px;padding:28px 32px;max-width:1400px;width:95%;max-height:92vh;overflow-y:auto;position:relative;box-shadow:0 24px 80px rgba(0,0,0,.5)}
.modal-close{position:absolute;top:14px;right:18px;border:none;background:none;font-size:22px;cursor:pointer;color:var(--text-dim);line-height:1;padding:2px 6px;border-radius:4px}
.modal-close:hover{background:rgba(255,255,255,.08);color:var(--text-primary)}
.modal-field{margin-bottom:14px}
.modal-field .mf-label{font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:var(--text-dim);margin-bottom:3px;font-weight:600}
.modal-field .mf-value{font-size:13px;color:var(--text-primary);line-height:1.5}
.modal-rec{background:var(--navy-surface);border:1px solid var(--navy-border);border-radius:8px;padding:14px;font-size:13px;line-height:1.6;color:var(--text-secondary);white-space:pre-wrap;word-break:break-word}
.modal-divider{border:none;border-top:1px solid var(--navy-border);margin:16px 0}
/* Utilities */
.text-muted{color:var(--text-dim)}
.filter-row{display:flex;gap:12px;margin-bottom:16px;align-items:center;flex-wrap:wrap}
.filter-row input,.filter-row select{padding:7px 12px;border:1px solid var(--navy-border);border-radius:6px;font-size:13px;outline:none;background:var(--navy-surface);color:var(--text-primary)}
.filter-row input:focus,.filter-row select:focus{border-color:var(--p-teal)}
.filter-row input::placeholder{color:var(--text-dim)}
.badge-count{background:var(--p-purple);color:#fff;border-radius:10px;padding:1px 8px;font-size:11px;margin-left:6px}
/* Welcome overlay */
.welcome-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.72);z-index:2000;align-items:center;justify-content:center}
.welcome-overlay.open{display:flex}
.welcome-box{background:var(--navy-card);border:1px solid var(--navy-border);border-radius:14px;padding:36px 40px;max-width:820px;width:92%;max-height:88vh;overflow-y:auto;position:relative;box-shadow:0 24px 80px rgba(0,0,0,.5)}
.welcome-box h2{margin:0 0 6px;font-size:20px;color:var(--p-teal)}
.welcome-box h3{margin:18px 0 8px;font-size:15px;color:var(--text-primary)}
.welcome-box p,.welcome-box li{font-size:13px;line-height:1.7;color:var(--text-secondary)}
.welcome-box ul{margin:4px 0 0 18px;padding:0}
.welcome-box li{margin-bottom:2px}
.welcome-paths{display:flex;gap:16px;margin:14px 0}
.welcome-path{flex:1;background:var(--navy-surface);border:1px solid var(--navy-border);border-radius:10px;padding:16px}
.welcome-path h4{margin:0 0 8px;font-size:13px;font-weight:700}
.welcome-path.save h4{color:var(--p-teal)}
.welcome-path.cost h4{color:var(--p-peach)}
.welcome-path ul{margin:4px 0 0 14px}
.welcome-path li{font-size:12px;line-height:1.6}
.welcome-example{background:var(--navy-surface);border:1px solid var(--navy-border);border-radius:10px;padding:16px;margin:12px 0}
.welcome-example h4{margin:0 0 10px;font-size:13px;color:var(--text-primary)}
.welcome-example .ex-cols{display:flex;gap:14px}
.welcome-example .ex-col{flex:1;font-size:12px;line-height:1.6;color:var(--text-secondary)}
.welcome-example .ex-col strong{display:block;margin-bottom:4px;font-size:12px}
.welcome-example .ex-col.save strong{color:var(--p-teal)}
.welcome-example .ex-col.cost strong{color:var(--p-peach)}
.welcome-example .ex-result{font-weight:700;margin-top:6px}
.welcome-example .ex-result.save{color:var(--p-teal)}
.welcome-example .ex-result.cost{color:var(--p-peach)}
.welcome-footer{display:flex;align-items:center;justify-content:space-between;margin-top:20px;padding-top:16px;border-top:1px solid var(--navy-border)}
.welcome-footer label{font-size:12px;color:var(--text-dim);cursor:pointer;display:flex;align-items:center;gap:6px}
.welcome-footer input[type=checkbox]{accent-color:var(--p-teal)}
.welcome-btn{background:var(--p-teal);color:var(--navy-bg);border:none;border-radius:8px;padding:10px 28px;font-size:13px;font-weight:700;cursor:pointer;letter-spacing:.3px}
.welcome-btn:hover{filter:brightness(1.1)}
.info-btn{background:none;border:1px solid var(--navy-border);border-radius:6px;padding:4px 10px;font-size:12px;color:var(--text-dim);cursor:pointer;margin-left:8px;vertical-align:middle}
.info-btn:hover{border-color:var(--p-teal);color:var(--p-teal)}
@media(max-width:700px){.welcome-paths,.welcome-example .ex-cols{flex-direction:column}}
@media print{.tabs{position:static}.panel{display:block!important;page-break-before:always}.panel:first-of-type{page-break-before:auto}.modal-overlay{display:none!important}.welcome-overlay{display:none!important}}
</style>
</head>
<body>

<!-- ── Modal ─────────────────────────────────────────────────────────────── -->
<div class="modal-overlay" id="modal-overlay" onclick="if(event.target===this)closeModal()">
  <div class="modal-box" id="modal-box">
    <button class="modal-close" onclick="event.stopPropagation();closeModal()">&#215;</button>
    <div id="modal-content"></div>
  </div>
</div>

<!-- ── Welcome overlay ──────────────────────────────────────────────────── -->
<div class="welcome-overlay" id="welcome-overlay" onclick="if(event.target===this)closeWelcome()">
  <div class="welcome-box">
    <h2>Understanding This Assessment</h2>
    <p>This assessment presents <strong>options, not decisions</strong>. Every user flagged in the analysis sits at a decision point where one path leads to a cost saving and the other leads to a compliance investment. The financial outcome depends entirely on the decisions your organization makes for each user.</p>
    <p>The two headline figures in this report are <strong>not additive</strong>. They represent the outer bounds of a decision matrix:</p>
    <div class="welcome-paths">
      <div class="welcome-path save">
        <h4>PATH A &mdash; Optimize &amp; Save</h4>
        <ul>
          <li>Remove or downgrade unused licenses</li>
          <li>Convert mailboxes, descope from policies</li>
          <li>Consolidate overlapping SKUs</li>
          <li>Deprovision dormant resources</li>
        </ul>
      </div>
      <div class="welcome-path cost">
        <h4>PATH B &mdash; Remediate &amp; Comply</h4>
        <ul>
          <li>Add missing licenses to close compliance gaps</li>
          <li>Upgrade SKUs to match policy requirements</li>
          <li>Maintain full coverage for security posture</li>
          <li>Accept cost to preserve current architecture</li>
        </ul>
      </div>
    </div>
    <p>In practice, most organizations apply a mix: optimizing some users while remediating others. The actual financial outcome is unique to your organization and will emerge from the decisions made on a per-user or per-group basis.</p>

    <div class="welcome-example">
      <h4>Worked Example: Exchange Online Plan 1 User</h4>
      <p style="font-size:12px;color:var(--text-dim);margin:0 0 10px">A user holds Exchange Online Plan 1 (&euro;42/yr) and is in scope of Conditional Access and Defender for Office 365 policies, but has neither entitlement assigned.</p>
      <div class="ex-cols">
        <div class="ex-col save">
          <strong>Option A &mdash; Optimize &amp; Save</strong>
          &bull; Remove Exchange Online Plan 1<br>
          &bull; Convert mailbox to Shared Mailbox<br>
          &bull; Remove from MDO &amp; CA scope<br>
          &bull; Disable sign-in if no longer active<br>
          <div class="ex-result save">&rarr; &minus;&euro;42/yr saved</div>
        </div>
        <div class="ex-col cost">
          <strong>Option B &mdash; Remediate &amp; Comply</strong>
          &bull; Keep Exchange Online Plan 1<br>
          &bull; Add Entra ID P1 (&euro;5.40/mo) for CA<br>
          &bull; Add MDO P1 (&euro;1.80/mo) for Defender<br>
          &bull; User is now fully compliant<br>
          <div class="ex-result cost">&rarr; +&euro;86/yr additional cost</div>
        </div>
      </div>
    </div>
    <p style="font-size:12px;color:var(--text-dim);margin-top:10px">Same user. Same data. Two very different financial outcomes. Every finding in this report carries this duality. We present the data and options; the choices are yours.</p>

    <div class="welcome-footer">
      <label><input type="checkbox" id="welcome-hide-cb"> Don't show this again</label>
      <button class="welcome-btn" onclick="closeWelcome()">View Assessment</button>
    </div>
  </div>
</div>

<header>
  <h1>M365 License Optimization Assessment <button class="info-btn" onclick="showWelcome()" title="Understanding this assessment">&#9432; Guide</button></h1>
  <p>$reportDate</p>
  <div class="kpis">
    <div class="kpi">
      <div class="label">Total Users</div>
      <div class="value">$kpiTotalUsers</div>
      <div class="sub">in scope</div>
    </div>
    <div class="kpi alert">
      <div class="label">With Assessments</div>
      <div class="value">$kpiWithRec</div>
      <div class="sub">$([math]::Round($kpiWithRec / [math]::Max($kpiTotalUsers,1) * 100, 0))% of users</div>
    </div>
    <div class="kpi">
      <div class="label">Total Annual Spend</div>
      <div class="value">&euro;$([string]::Format('{0:N0}', $kpiTotalSpend))/yr</div>
      <div class="sub">licensed users</div>
    </div>
    <div class="kpi good">
      <div class="label">Potential Annual Savings</div>
      <div class="value">&euro;$([string]::Format('{0:N0}', $kpiSavingsPot))/yr</div>
      <div class="sub">$kpiSavingsPct% of annual spend</div>
    </div>
$(if ($kpiCompCost -gt 0) {
    "    <div class=`"kpi`">
      <div class=`"label`">Potential Compliance Costs</div>
      <div class=`"value`" style=`"color:var(--p-peach)`">&euro;$([string]::Format('{0:N0}', $kpiCompCost))/yr</div>
      <div class=`"sub`">across $kpiCompUsers user$(if ($kpiCompUsers -ne 1) {'s'})</div>
    </div>"
})
  </div>
  <div class="disclaimer"><div class="disclaimer-track"><span>$disclaimer1</span><span>$disclaimer2</span><span>$disclaimer3</span><span>$disclaimer4</span><span>$disclaimer1</span><span>$disclaimer2</span><span>$disclaimer3</span><span>$disclaimer4</span></div></div>
</header>

<div class="tabs">
  <button class="tab-btn active" onclick="showTab(0)">&#9733; Overview</button>
  <button class="tab-btn"        onclick="showTab(1)">&#128202; Assessments by Category</button>
  <button class="tab-btn"        onclick="showTab(2)">&#128176; Assessments by User</button>
  <button class="tab-btn"        onclick="showTab(3)">&#128230; Assessments by SKU</button>
  <button class="tab-btn"        onclick="showTab(4)">&#128274; License Groups</button>
  <button class="tab-btn"        onclick="showTab(5)"><span style="color:#fff">&#9776;</span> Workload Usage Matrix</button>
  <button class="tab-btn"        onclick="showTab(6)" id="sub-alerts-tab-btn" style="display:none"><span style="color:#e53e3e">&#9888;</span> Subscription Alerts</button>
  <button class="tab-btn"        onclick="showTab(7)" id="copilot-tab-btn" style="display:none">&#129302; Copilot Adoption</button>
</div>

<!-- TAB 0: DASHBOARD OVERVIEW -->
<div class="panel active" id="panel-0">
  <div class="dash-section-title" style="display:flex;align-items:center;flex-wrap:wrap;gap:8px">
    Quick Wins
    <span class="tier-legend">
      <span class="tier-dot" style="background:#ef6ea7"></span> License review
      <span class="tier-dot" style="background:#3ddad7;margin-left:12px"></span> Right-sizing
      <span class="tier-dot" style="background:#ff9f80;margin-left:12px"></span> Compliance review
    </span>
  </div>
  <div id="dash-tiles" class="dash-tiles-grid"></div>
</div>

<!-- TAB 1: SAVINGS BY CATEGORY -->
<div class="panel" id="panel-1">
  <div class="card">
    <h3>Assessments by Category</h3>
    <div id="dash-breakdown" style="margin-top:12px"></div>
  </div>
</div>

<!-- TAB 2: SAVINGS BY USER -->
<div class="panel" id="panel-2">
  <div class="card">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:8px">
      <h3 style="margin:0">All Users with Assessments <span class="badge-count" id="user-count"></span></h3>
      <div class="filter-row" style="margin-bottom:0;gap:8px">
        <input type="text" id="user-filter" placeholder="Filter by name / UPN" oninput="renderUserTable()" style="max-width:300px">
        <select id="dept-filter" onchange="renderUserTable()"><option value="">All departments</option></select>
        <select id="cat-filter" onchange="renderUserTable()"><option value="">All categories</option></select>
      </div>
    </div>
    <div class="tbl-wrap">
      <table id="user-table">
        <thead>
          <tr>
            <th onclick="sortTable('Name')"     data-col="Name">     Name <span class="sort-icon">&#9660;</span></th>
            <th onclick="sortTable('Dept')"     data-col="Dept">     Department <span class="sort-icon">&#9660;</span></th>
            <th onclick="sortTable('Savings')"  data-col="Savings">  Savings/yr <span class="sort-icon">&#9660;</span></th>
            <th onclick="sortTable('CompCost')" data-col="CompCost"> Cost/yr <span class="sort-icon">&#9660;</span></th>
            <th onclick="sortTable('Licenses')" data-col="Licenses"> Licenses <span class="sort-icon">&#9660;</span></th>
            <th onclick="sortTable('Category')" data-col="Category"> Category <span class="sort-icon">&#9660;</span></th>
          </tr>
        </thead>
        <tbody id="user-tbody"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- TAB 3: SAVINGS BY SKU -->
<div class="panel" id="panel-3">
  <div class="card">
    <h3>License Optimization Opportunities by SKU</h3>
    <p class="section-desc">Estimated optimization potential attributed to each license type. Hover a bar for category breakdown.</p>
    <div id="sku-chart"></div>
  </div>
</div>

<!-- TAB 4: LICENSE GROUPS -->
<div class="panel" id="panel-4">
  <div class="card">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:8px">
      <h3 style="margin:0">Entra ID Licensing Groups <span class="badge-count" id="group-count"></span></h3>
      <div class="filter-row" style="margin-bottom:0;gap:8px">
        <input type="text" id="group-filter" placeholder="Filter by group name or SKU" oninput="renderGroupTable()" style="max-width:300px">
      </div>
    </div>
    <div class="tbl-wrap">
      <table>
        <thead>
          <tr>
            <th style="text-align:left">Group Name</th>
            <th>Type</th>
            <th>Members</th>
            <th>Direct Assigned</th>
            <th>Tenant Pool</th>
            <th style="text-align:left">Assigned Licenses</th>
          </tr>
        </thead>
        <tbody id="group-tbody"></tbody>
      </table>
    </div>
    <p style="font-size:11px;color:#6a6a8e;margin-top:10px" id="group-empty"></p>
  </div>
</div>

<!-- TAB 5: WORKLOAD USAGE MATRIX (per user) -->
<div class="panel" id="panel-5">
  <div class="card">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:8px">
      <h3 style="margin:0">Provisioned vs Used</h3>
      <div class="filter-row" style="margin-bottom:0;gap:8px">
        <input type="text" id="cap-filter" placeholder="Filter by name / UPN" oninput="renderCapMatrix()" style="max-width:300px">
      </div>
    </div>
    <div style="font-size:11px;color:#6a6a8e;margin-bottom:8px">
      <span style="display:inline-block;width:10px;height:10px;background:#ef6ea7;border-radius:3px;vertical-align:middle"></span> Low / Unused &nbsp;
      <span style="display:inline-block;width:10px;height:10px;background:#5b89b6;border-radius:3px;vertical-align:middle"></span> No license &nbsp;
      <span style="display:inline-block;width:10px;height:10px;background:rgba(61,218,215,.20);border-radius:3px;vertical-align:middle"></span> Medium &nbsp;
      <span style="display:inline-block;width:10px;height:10px;background:rgba(61,218,215,.35);border-radius:3px;vertical-align:middle"></span> High &nbsp;
      <span style="display:inline-block;width:10px;height:10px;background:rgba(255,255,255,.06);border-radius:3px;vertical-align:middle;border:1px solid #2a2a55"></span> N/A
    </div>
    <div class="tbl-wrap">
      <table class="cap-table" id="cap-table">
        <thead>
          <tr>
            <th class="user-col"  onclick="sortCapTable('n')"    data-capcol="n">User <span class="sort-icon">&#9660;</span></th>
            <th onclick="sortCapTable('sav')"  data-capcol="sav">Savings/yr <span class="sort-icon">&#9660;</span></th>
            <th onclick="sortCapTable('comp')" data-capcol="comp">Cost/yr <span class="sort-icon">&#9660;</span></th>
            <th>Exchange</th>
            <th>Teams</th>
            <th>Desktop</th>
            <th>OneDrive</th>
            <th>SharePoint</th>
            <th>Copilot</th>
            <th style="text-align:left" onclick="sortCapTable('cat')" data-capcol="cat">Category <span class="sort-icon">&#9660;</span></th>
          </tr>
        </thead>
        <tbody id="cap-tbody"></tbody>
      </table>
    </div>
    <p style="font-size:11px;color:#6a6a8e;margin-top:10px">Showing all non-OK users by potential savings.</p>
  </div>
</div>

<!-- TAB 6: SUBSCRIPTION ALERTS -->
<div class="panel" id="panel-6">
  <div id="sub-alerts-section"></div>
</div>

<!-- TAB 7: COPILOT ADOPTION & ROI -->
<div class="panel" id="panel-7">
  <div id="copilot-roi-section"></div>
</div>

<script>
const USERS     = $jsTopUsers;
const SKUS      = $jsSkuData;
const TILES     = $jsTileData;
const CAP_USERS = $jsCapUsers;
const POOL_SKUS = $jsPoolSkus;
const GROUPS    = $jsGroups;
const ALL_CATS  = $jsAllCats;
const SUB_ALERTS = $jsSubAlerts;
const DEPTS      = $jsDepts;
const COPILOT_ROI = $jsCopilotRoi;
const CP_HOLDERS  = $jsCpHolders;
"@

$html += @'

// ── Tab switching ─────────────────────────────────────────────────────────────
function showTab(idx) {
  document.querySelectorAll('.panel').forEach((p,i) => p.classList.toggle('active', i===idx));
  document.querySelectorAll('.tab-btn').forEach((b,i) => b.classList.toggle('active', i===idx));
}

// ── Color helpers ─────────────────────────────────────────────────────────────
function savingsColor(val, max) {
  if (max === 0) return 'rgba(255,255,255,.06)';
  return '#3ddad7';
}
function savingsTextColor(val, max) {
  if (max === 0) return '#6a6a8e';
  return '#0a1628';
}
function capCellUser(prov, used, intensity) {
  if (!prov && !used) return { bg:'rgba(255,255,255,.06)', text:'#6a6a8e', label:'N/A' };
  if (prov && used) {
    const lbl = intensity || 'Active';
    if (lbl === 'High')   return { bg:'rgba(61,218,215,.35)',  text:'#3ddad7', label:'High' };
    if (lbl === 'Medium') return { bg:'rgba(61,218,215,.20)',  text:'#3ddad7', label:'Medium' };
    return { bg:'rgba(61,218,215,.25)', text:'#3ddad7', label:lbl };
  }
  if (prov && !used) {
    const lbl = intensity === 'Low' ? 'Low' : 'Unused';
    return { bg:'rgba(239,110,167,.25)', text:'#ef6ea7', label:lbl };
  }
  return { bg:'rgba(91,137,182,.25)', text:'#5b89b6', label:'No license' };
}
function fmtEur(v) { return '\u20ac' + Number(v).toLocaleString('en-GB', {minimumFractionDigits:0,maximumFractionDigits:0}); }
function escHtml(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function admBadge(level) {
  if (!level) return '';
  const c = level === 'High' ? '#ff9f80' : '#6a6a8e';
  const t = level === 'High' ? 'High-privilege admin' : 'Low-privilege admin';
  return ' <span title="'+t+'" style="display:inline-block;font-size:9px;font-weight:700;color:'+c+';border:1px solid '+c+';border-radius:3px;padding:0 3px;vertical-align:middle;margin-left:4px">admin</span>';
}
function renderTags(tags) {
  if (!tags || !tags.length) return '';
  return tags.map(t => '<span style="display:inline-block;padding:1px 6px;border-radius:10px;font-size:10px;font-weight:500;background:rgba(255,159,128,.15);color:#ff9f80;margin-left:4px;white-space:nowrap">'+escHtml(t)+'</span>').join('');
}
function cleanBody(s) {
  // Replace em-dashes and double hyphens with commas
  s = s.replace(/\s*\u2014\s*/g, ', ');
  s = s.replace(/\s*--\s*/g, ', ');
  // Replace semicolons with commas
  s = s.replace(/;\s*/g, ', ');
  // Clean up double commas
  s = s.replace(/,\s*,/g, ',');
  // Trim leading comma after label extraction
  s = s.replace(/^,\s*/, '');
  return s;
}
// ── Subscription Alerts ──────────────────────────────────────────────────────
function renderSubAlerts() {
  const box = document.getElementById('sub-alerts-section');
  if (!box || !SUB_ALERTS || !SUB_ALERTS.length) return;
  // Show the Subscription Alerts tab button
  const tabBtn = document.getElementById('sub-alerts-tab-btn');
  if (tabBtn) tabBtn.style.display = '';
  const rows = SUB_ALERTS.map(s => {
    const isExp   = s.status === 'Suspended' || s.days < 0;
    const isWarn  = s.status === 'Warning';
    const color   = isExp ? '#ef6ea7' : isWarn ? '#ff9f80' : '#3ddad7';
    const bgColor = isExp ? 'rgba(239,110,167,.10)' : isWarn ? 'rgba(255,159,128,.10)' : 'rgba(61,218,215,.08)';
    const badge   = isExp ? 'Suspended' : isWarn ? 'Warning' : 'Expiring';
    const daysAbs = Math.abs(s.days);
    const daysText = s.days < 0 ? `expired ${daysAbs}d ago` : `${daysAbs}d remaining`;
    return `<div style="display:flex;align-items:center;gap:12px;padding:8px 14px;background:${bgColor};border-left:3px solid ${color};border-radius:6px;margin-bottom:6px">
      <span style="font-weight:600;color:${color};font-size:11px;min-width:70px;text-transform:uppercase">${badge}</span>
      <span style="flex:1;font-size:13px">${escHtml(s.sku)}</span>
      <span style="font-size:12px;color:var(--text-secondary)">${escHtml(s.seats)} seats</span>
      <span style="font-size:12px;font-weight:600;color:${color}">${daysText}</span>
    </div>`;
  }).join('');
  box.innerHTML = `<div class="card" style="margin-bottom:20px">
    <h3 style="color:var(--p-pink)">Subscription Alerts <span style="font-size:12px;font-weight:400;color:#6a6a8e;margin-left:8px">${SUB_ALERTS.length} SKU(s) require attention</span></h3>
    <div style="margin-top:12px">${rows}</div>
  </div>`;
}

// ── Copilot ROI Section ──────────────────────────────────────────────────────
function renderCopilotRoi() {
  const box = document.getElementById('copilot-roi-section');
  if (!box || !COPILOT_ROI || COPILOT_ROI.total === 0) return;
  const d = COPILOT_ROI;
  const adoptionPct = d.total > 0 ? Math.round(d.active / d.total * 100) : 0;
  const costPerUser = d.total > 0 ? Math.round(26 * 12 / 1) : 312; // €26/mo standard
  const totalInvest = d.total * 312;
  const activeInvest = d.active * 312;
  // App adoption bars
  const apps = d.apps || {};
  const appList = [
    {name:'Teams',l:apps.Teams||0},{name:'Outlook',l:apps.Outlook||0},
    {name:'Word',l:apps.Word||0},{name:'Excel',l:apps.Excel||0},
    {name:'PowerPoint',l:apps.PowerPoint||0},{name:'OneNote',l:apps.OneNote||0},
    {name:'Copilot Chat',l:apps.Chat||0},{name:'Loop',l:apps.Loop||0}
  ].sort((a,b) => b.l - a.l);
  const maxApp = Math.max(...appList.map(a => a.l), 1);
  const appBars = appList.map(a => {
    const pct = Math.round(a.l / d.total * 100);
    const w   = Math.round(a.l / maxApp * 100);
    const notUsing = d.total - a.l;
    const appKey = a.name === 'Copilot Chat' ? 'Chat' : a.name;
    const clk = notUsing > 0 ? "showCopilotAppGap('" + appKey + "')" : '';
    const cur = notUsing > 0 ? 'pointer' : 'default';
    const ttl = notUsing > 0 ? notUsing + ' user(s) with no Copilot ' + a.name + ' activity \u2014 click to view' : 'All holders active';
    return '<div class="bd-row" style="cursor:' + cur + '" onclick="' + clk + '" title="' + ttl + '">'
      + '<div class="bd-label">' + a.name + '</div>'
      + '<div class="bd-track"><div class="bd-fill" style="width:' + w + '%;background:linear-gradient(90deg,#48349a,#3ddad7)"></div></div>'
      + '<div class="bd-amt">' + a.l + '/' + d.total + '</div>'
      + '</div>';
  }).join('');

  // Pipeline donut-like summary
  const segments = [
    {label:'Active',count:d.active,color:'#3ddad7'},
    {label:'Watchlist',count:d.watchlist,color:'#ff9f80'},
    {label:'Reclaim',count:d.reclaim,color:'#ef6ea7'}
  ].filter(s => s.count > 0);
  const segHtml = segments.map(s =>
    `<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px">
      <div style="width:14px;height:14px;border-radius:50%;background:${s.color};flex-shrink:0"></div>
      <span style="font-size:13px;min-width:80px">${s.label}</span>
      <span style="font-family:'JetBrains Mono',monospace;font-weight:600;font-size:15px">${s.count}</span>
      <span style="font-size:11px;color:#6a6a8e">(${d.total>0?Math.round(s.count/d.total*100):0}%)</span>
    </div>`
  ).join('');

  // Show the Copilot tab button
  const tabBtn = document.getElementById('copilot-tab-btn');
  if (tabBtn) tabBtn.style.display = '';

  box.innerHTML = `<div class="card">
    <h3 style="color:var(--p-teal)">Copilot Adoption &amp; ROI
      <span style="font-size:12px;font-weight:400;color:#6a6a8e;margin-left:8px">${d.total} license holders</span>
    </h3>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-top:16px">
      <div>
        <div style="font-size:12px;color:var(--text-dim);text-transform:uppercase;letter-spacing:.4px;margin-bottom:12px">Adoption Pipeline</div>
        <div style="display:flex;align-items:baseline;gap:8px;margin-bottom:14px">
          <span style="font-family:'JetBrains Mono',monospace;font-size:36px;font-weight:700;color:var(--p-teal)">${adoptionPct}%</span>
          <span style="font-size:12px;color:var(--text-secondary)">adoption rate</span>
        </div>
        ${segHtml}
        <div style="margin-top:14px;padding:10px 14px;background:rgba(61,218,215,.06);border-radius:8px;font-size:12px;color:var(--text-secondary)">
          Annual investment: <span style="font-family:'JetBrains Mono',monospace;font-weight:600;color:var(--text-primary)">${fmtEur(totalInvest)}</span>
          &nbsp;&bull;&nbsp;Cost per active user: <span style="font-family:'JetBrains Mono',monospace;font-weight:600;color:var(--text-primary)">${d.active>0?fmtEur(Math.round(totalInvest/d.active)):'\u2014'}/yr</span>
        </div>
      </div>
      <div>
        <div style="font-size:12px;color:var(--text-dim);text-transform:uppercase;letter-spacing:.4px;margin-bottom:12px">App Penetration (across ${d.total} holders)</div>
        ${appBars}
        <div style="margin-top:10px;font-size:10px;color:#6a6a8e">Click any bar to see who is not using that app</div>
      </div>
    </div>
  </div>`;
}

function showCopilotAppGap(appName) {
  if (!CP_HOLDERS || CP_HOLDERS.length === 0) return;
  const notUsing = CP_HOLDERS.filter(u => !u.CpApps || !u.CpApps.match(new RegExp(appName, 'i')));
  if (notUsing.length === 0) return;
  clearBackState();
  const displayName = appName === 'Chat' ? 'Copilot Chat' : appName;
  let html = '<h2 style="color:var(--p-teal);margin-bottom:4px">No Copilot ' + escHtml(displayName) + ' Activity</h2>';
  html += '<div style="color:var(--text-secondary);font-size:13px;margin-bottom:16px">' + notUsing.length + ' of ' + CP_HOLDERS.length + ' Copilot holder(s) have no Copilot ' + escHtml(displayName) + ' activity in D90</div>';
  html += '<table style="width:100%;border-collapse:collapse;font-size:13px"><thead><tr style="background:#181835;border-bottom:1px solid #2a2a55"><th style="text-align:left;padding:10px 12px;font-weight:600;color:#9898b8">Name</th><th style="text-align:left;padding:10px 12px;font-weight:600;color:#9898b8">Department</th><th style="text-align:left;padding:10px 12px;font-weight:600;color:#9898b8">Active In</th></tr></thead><tbody>';
  notUsing.sort((a,b) => (a.Name||'').localeCompare(b.Name||'')).forEach(u => {
    const activeIn = u.CpApps ? u.CpApps.replace(/;\s*/g, ' - ') : '<span style="color:#6a6a8e">None</span>';
    const hasDetail = USERS.some(x => x.UPN === u.UPN);
    const click = hasDetail ? 'onclick="showCopilotUserDetail(\'' + (u.UPN||'').replace(/'/g,"\\'") + '\')"' : '';
    const cursor = hasDetail ? 'cursor:pointer' : 'cursor:default';
    html += '<tr style="border-bottom:1px solid rgba(255,255,255,.04);' + cursor + '" ' + (hasDetail ? 'title="Click for full assessment"' : '') + ' ' + click + '>'
      + '<td style="padding:10px 12px">' + escHtml(u.Name||u.UPN) + admBadge(u.AdminPriv) + '</td>'
      + '<td style="padding:10px 12px">' + escHtml(u.Dept||'') + '</td>'
      + '<td style="padding:10px 12px;font-size:11px;color:#9898b8">' + activeIn + '</td>'
      + '</tr>';
  });
  html += '</tbody></table>';
  document.getElementById('modal-content').innerHTML = html;
  document.getElementById('modal-overlay').classList.add('open');
}

function showCopilotUserDetail(upn) {
  const u = USERS.find(x => x.UPN === upn);
  if (!u) return;
  showUserDetail(u);
}

function extractAmount(body) {
  // Pull out trailing cost/savings amounts into a styled badge
  // Patterns ordered from most specific to most general; first match wins
  const patterns = [
    /\.\s*((?:Potential |Estimated )?(?:savings|cost|waste):\s*\u20AC[\d.,]+\/mo\s*\(\u20AC[\d.,]+\/yr\))\.?$/i,
    /\.\s*(Annual (?:waste|cost|overlap cost|savings|compliance cost|licensing cost):\s*\u20AC[\d.,]+(?:\/yr)?)\.?$/i,
    /\.\s*(Saves?\s*\u20AC[\d.,]+\/mo\s*\(\u20AC[\d.,]+\/yr\))\.?$/i,
    /\.\s*(Estimated compliance cost:\s*\u20AC[\d.,]+(?:\/yr)?)\.?$/i,
    /\.\s*((?:Annual |Estimated )?(?:cost|savings):\s*\u20AC[\d.,]+(?:\/yr)?)\.?$/i
  ];
  for (const rx of patterns) {
    const am = body.match(rx);
    if (am) {
      const cleaned = body.replace(rx, '.').replace(/\.\s*\.$/, '.').trim();
      return { body: cleaned, amount: am[1] };
    }
  }
  return { body: body, amount: null };
}
// Label → border + badge colors
const REC_LABEL_COLORS = {
  // Pink — dormant / inactive / never signed in
  'DORMANT':              { border:'#ef6ea7', bg:'rgba(239,110,167,.12)', text:'#ef6ea7' },
  'DORMANT ADMIN REVIEW': { border:'#ef6ea7', bg:'rgba(239,110,167,.12)', text:'#ef6ea7' },
  'DORMANT CLOUD PC':     { border:'#ef6ea7', bg:'rgba(239,110,167,.12)', text:'#ef6ea7' },
  'STALE SIGN-IN':        { border:'#ef6ea7', bg:'rgba(239,110,167,.12)', text:'#ef6ea7' },
  'NEVER SIGNED IN':      { border:'#ef6ea7', bg:'rgba(239,110,167,.12)', text:'#ef6ea7' },
  'INACTIVE HOLD':        { border:'#ef6ea7', bg:'rgba(239,110,167,.12)', text:'#ef6ea7' },
  'INACTIVE MAILBOX':     { border:'#ef6ea7', bg:'rgba(239,110,167,.12)', text:'#ef6ea7' },
  // Peach — disabled / compliance / licensing
  'DISABLED ACCOUNT':     { border:'#ff9f80', bg:'rgba(255,159,128,.12)', text:'#ff9f80' },
  'DISABLED SHARED MAILBOX': { border:'#ff9f80', bg:'rgba(255,159,128,.12)', text:'#ff9f80' },
  'LICENSING CHECK':      { border:'#ff9f80', bg:'rgba(255,159,128,.12)', text:'#ff9f80' },
  'LICENSING ERROR':      { border:'#ff9f80', bg:'rgba(255,159,128,.12)', text:'#ff9f80' },
  'CLOUD PC REVIEW':      { border:'#ff9f80', bg:'rgba(255,159,128,.12)', text:'#ff9f80' },
  'SECURITY GAP':         { border:'#ef6ea7', bg:'rgba(239,110,167,.12)', text:'#ef6ea7' },
  // Amber — activity warnings
  'NO ACTIVITY':          { border:'#f59f00', bg:'rgba(245,159,0,.12)',   text:'#f59f00' },
  'BACKGROUND SYNC ONLY': { border:'#f59f00', bg:'rgba(245,159,0,.12)',   text:'#f59f00' },
  'EXPENSIVE COLD STORAGE': { border:'#f59f00', bg:'rgba(245,159,0,.12)', text:'#f59f00' },
  'FORWARDING MAILBOX':   { border:'#f59f00', bg:'rgba(245,159,0,.12)',   text:'#f59f00' },
  // Green — shared mailbox / overlap
  'SHARED MAILBOX':       { border:'#2f9e44', bg:'rgba(47,158,68,.12)',   text:'#2f9e44' },
  'SHARED MAILBOX REVIEW': { border:'#2f9e44', bg:'rgba(47,158,68,.12)', text:'#2f9e44' },
  // Blue — duplicates / overlaps / right-sizing
  'DUPLICATE COVERAGE':   { border:'#5c7cfa', bg:'rgba(92,124,250,.12)', text:'#5c7cfa' },
  'DUPLICATE REVIEW':     { border:'#5c7cfa', bg:'rgba(92,124,250,.12)', text:'#5c7cfa' },
  'OVERLAPPING LICENSE':  { border:'#748ffc', bg:'rgba(116,143,252,.12)', text:'#748ffc' },
  'SUITE INVERSION':      { border:'#748ffc', bg:'rgba(116,143,252,.12)', text:'#748ffc' },
  'BUNDLE CONSOLIDATION': { border:'#748ffc', bg:'rgba(116,143,252,.12)', text:'#748ffc' },
  'BUNDLE OPPORTUNITY':   { border:'#748ffc', bg:'rgba(116,143,252,.12)', text:'#748ffc' },
  // Steel blue — admin / automation / service accounts
  'ADMIN':                { border:'#5b89b6', bg:'rgba(91,137,182,.12)', text:'#5b89b6' },
  'AUTOMATION ACCOUNT':   { border:'#5b89b6', bg:'rgba(91,137,182,.12)', text:'#5b89b6' },
  'NON-HUMAN ACCOUNT REVIEW': { border:'#5b89b6', bg:'rgba(91,137,182,.12)', text:'#5b89b6' },
  'LEGACY SERVICE ACCOUNT': { border:'#5b89b6', bg:'rgba(91,137,182,.12)', text:'#5b89b6' },
  'GUEST ACCOUNT':        { border:'#5b89b6', bg:'rgba(91,137,182,.12)', text:'#5b89b6' },
  // Teal — Copilot
  'COPILOT RECLAIM':      { border:'#0c8599', bg:'rgba(12,133,153,.12)', text:'#0c8599' },
  'COPILOT WATCHLIST':    { border:'#0c8599', bg:'rgba(12,133,153,.12)', text:'#0c8599' },
  'COPILOT PREREQUISITE': { border:'#0c8599', bg:'rgba(12,133,153,.12)', text:'#0c8599' },
  'COPILOT STUDIO':       { border:'#3ddad7', bg:'rgba(61,218,215,.12)', text:'#3ddad7' },
  // Purple — frontline / right-sizing
  'FRONTLINE CANDIDATE':  { border:'#7950f2', bg:'rgba(121,80,242,.12)', text:'#7950f2' },
  'FRONTLINE RESCUE':     { border:'#7950f2', bg:'rgba(121,80,242,.12)', text:'#7950f2' },
  'FRONTLINE ADD-ON STACKING': { border:'#7950f2', bg:'rgba(121,80,242,.12)', text:'#7950f2' },
  'INACTIVE ADD-ON':      { border:'#7950f2', bg:'rgba(121,80,242,.12)', text:'#7950f2' },
  'INACTIVE ADD-ON REVIEW': { border:'#7950f2', bg:'rgba(121,80,242,.12)', text:'#7950f2' },
  // Data / info
  'DATA GAP':             { border:'#868e96', bg:'rgba(134,142,150,.12)', text:'#868e96' },
};
function getLabelStyle(label) {
  const uc = label.toUpperCase();
  for (const [k,v] of Object.entries(REC_LABEL_COLORS)) { if (uc === k) return v; }
  // Partial match for compound labels like "DUPLICATE COVERAGE"
  for (const [k,v] of Object.entries(REC_LABEL_COLORS)) { if (uc.startsWith(k)) return v; }
  return { border:'#48349a', bg:'rgba(72,52,154,.15)', text:'#9898b8' };
}
function styleNotes(html) {
  // Style NOTE:, SECURITY:, CAUTION:, IMPORTANT: as colored callout blocks
  return html
    .replace(/\bNOTE:\s*/gi, '<div style="margin-top:6px;padding:5px 8px;background:rgba(91,137,182,.12);border-left:3px solid #5b89b6;border-radius:0 4px 4px 0;font-size:11.5px;color:#5b89b6;line-height:1.5"><strong>Note:</strong> ')
    .replace(/\bSECURITY:\s*/gi, '<div style="margin-top:6px;padding:5px 8px;background:rgba(239,110,167,.12);border-left:3px solid #ef6ea7;border-radius:0 4px 4px 0;font-size:11.5px;color:#ef6ea7;line-height:1.5"><strong>Security:</strong> ')
    .replace(/\bCAUTION:\s*/gi, '<div style="margin-top:6px;padding:5px 8px;background:rgba(255,159,128,.12);border-left:3px solid #ff9f80;border-radius:0 4px 4px 0;font-size:11.5px;color:#ff9f80;line-height:1.5"><strong>Caution:</strong> ')
    .replace(/\bIMPORTANT:\s*/gi, '<div style="margin-top:6px;padding:5px 8px;background:rgba(255,159,128,.12);border-left:3px solid #ff9f80;border-radius:0 4px 4px 0;font-size:11.5px;color:#ff9f80;line-height:1.5"><strong>Important:</strong> ')
    // Close the div: match from the callout open tag to end of string or next callout/pipe
    .replace(/(<div style="margin-top:6px[^>]*><strong>\w+:<\/strong>\s*)([\s\S]*?)(?=<div style="margin-top:6px|$)/g, '$1$2</div>');
}
function formatRec(raw) {
  if (!raw) return '<span style="color:#6a6a8e">No assessment text available.</span>';
  const parts = raw.split(' | ').filter(p => p.trim() && !p.trim().startsWith('COPILOT ACTIVE'));
  if (parts.length === 0) return escHtml(raw);
  const items = parts.map(p => {
    const m = p.match(/^([A-Za-z][A-Za-z0-9 /\-_.&]+?)(?:\s*\((?:[^()]*|\([^()]*\))*\))?\s*(?:,\s*)?\u2014\s*(.+)/);
    if (m) {
      const label = m[1].trim();
      const ls = getLabelStyle(label);
      const cleaned = cleanBody(m[2].trim());
      const { body, amount } = extractAmount(cleaned);
      // Split body at NOTE:/SECURITY:/CAUTION: for styled callouts
      const bodyHtml = styleNotes(escHtml(body));
      const amountHtml = amount
        ? '<div style="margin-top:5px;display:inline-block;background:rgba(255,159,128,.15);color:#ff9f80;font-size:11px;font-weight:600;padding:2px 8px;border-radius:3px">' + escHtml(amount) + '</div>'
        : '';
      return '<li style="margin-bottom:12px;padding:8px 10px;background:#181835;border-radius:6px;border-left:3px solid ' + ls.border + '">'
        + '<span style="display:inline-block;background:' + ls.bg + ';color:' + ls.text + ';font-size:10px;font-weight:700;padding:2px 6px;border-radius:3px;margin-bottom:4px;letter-spacing:0.3px">' + escHtml(label) + '</span>'
        + '<br><span style="color:#9898b8;line-height:1.6;font-size:12.5px">' + bodyHtml + '</span>'
        + amountHtml
        + '</li>';
    }
    const cleaned = cleanBody(p.trim());
    const { body, amount } = extractAmount(cleaned);
    const bodyHtml = styleNotes(escHtml(body));
    const amountHtml = amount
      ? '<div style="margin-top:5px;display:inline-block;background:rgba(255,159,128,.15);color:#ff9f80;font-size:11px;font-weight:600;padding:2px 8px;border-radius:3px">' + escHtml(amount) + '</div>'
      : '';
    return '<li style="margin-bottom:12px;padding:8px 10px;background:#181835;border-radius:6px;border-left:3px solid #5b89b6">'
      + '<span style="color:#9898b8;line-height:1.6;font-size:12.5px">' + bodyHtml + '</span>'
      + amountHtml
      + '</li>';
  });
  return '<ul style="list-style:none;padding:0;margin:0">' + items.join('') + '</ul>';
}

// ── Dashboard rendering ───────────────────────────────────────────────────────
function renderDashboard() {
  const grid = document.getElementById('dash-tiles');
  const tierColors = {1:'#ef6ea7',2:'#3ddad7',3:'#ff9f80'};
  grid.innerHTML = TILES.map((t,i) => {
    const hasData = t.users > 0 || t.savings > 0;
    const accent = '#5b89b6';
    const dotColor = tierColors[t.tier] || accent;
    return `<div class="dash-tile${hasData ? '' : ' dt-zero'}" style="border-top-color:${accent}" onclick="clickTile(${i})">
      <span class="dt-tier-dot" style="background:${dotColor}"></span>
      <div class="dt-label">${escHtml(t.label)}</div>
      <div class="dt-desc">${escHtml(t.desc)}</div>
      <div class="dt-count" style="color:${hasData ? '#3ddad7' : '#6a6a8e'}">${t.users}</div>
      <div class="dt-savings" style="color:${hasData ? '#3ddad7' : '#6a6a8e'}">${t.savings > 0 ? fmtEur(t.savings)+'/yr' : '\u2014'}</div>
      <div class="dt-bar" style="background:${accent}"></div>
    </div>`;
  }).join('');

  const breakdown = document.getElementById('dash-breakdown');
  const tilesWithSavings = TILES.filter(t => t.savings > 0).sort((a,b) => b.savings - a.savings);
  if (!tilesWithSavings.length) { breakdown.innerHTML = '<p class="text-muted">No potential savings data available.</p>'; return; }
  const maxSav = tilesWithSavings[0].savings;
  const bars = tilesWithSavings.map(t => {
    const pct = (t.savings / maxSav * 100).toFixed(1);
    const idx = TILES.indexOf(t);
    return `<div class="bd-row" onclick="clickTile(${idx})" style="cursor:pointer">
      <div class="bd-label">${escHtml(t.label)}</div>
      <div class="bd-track"><div class="bd-fill" style="width:${pct}%;background:#5b89b6"></div></div>
      <div class="bd-amt">${fmtEur(t.savings)}</div>
    </div>`;
  }).join('');
  breakdown.innerHTML = bars;
}

// ── Tile click → modal ──────────────────────────────────────────────────────
let tileModalUsers = [];
let activeTileIdx = -1;

function clickTile(idx) {
  const t = TILES[idx];
  if (!t.users && !t.savings) return;
  if (!t.key && !t.recKey) { activeTileIdx = -1; showPoolModal(); return; }
  showTileModal(idx);
}

function clearBackState() {
  const mb = document.getElementById('modal-box');
  mb.style.cursor = ''; delete mb.dataset.backTile; delete mb.dataset.backSku;
}

function showTileModal(idx) {
  activeTileIdx = idx;
  clearBackState();
  const t = TILES[idx];
  const tileRx    = t.key    ? new RegExp(t.key, 'i')    : null;
  const tileRecRx = t.recKey ? new RegExp(t.recKey, 'i') : null;
  const matched = USERS.filter(u =>
    (tileRx && tileRx.test(u.Category)) ||
    (tileRecRx && tileRecRx.test(u.Rec||''))
  ).sort((a,b) => b.Savings - a.Savings);
  tileModalUsers = matched;
  const totalSav = matched.reduce((s,u) => s + u.Savings, 0);
  const totalComp = matched.reduce((s,u) => s + (u.CompCost||0), 0);
  const tableRows = matched.map((u, i) =>
    `<tr style="border-bottom:1px solid rgba(255,255,255,.04);cursor:pointer" onclick="showTileUserDetail(${i})" title="Click for full assessment">
      <td style="padding:10px 12px"><div style="font-weight:500">${escHtml(u.Name||u.UPN)}${admBadge(u.AdminPriv)}</div><div style="font-size:11px;color:#6a6a8e">${escHtml(u.UPN||'')}</div></td>
      <td style="padding:10px 12px">${escHtml(u.Dept||'')}</td>
      <td style="padding:10px 12px"><span class="cat-badge">${escHtml(u.Category||'')}</span>${renderTags(u.Tags)}</td>
      <td style="padding:10px 12px;text-align:right">${fmtEur(u.Cost)}</td>
      <td style="padding:10px 12px;text-align:right;font-weight:600;color:#2f9e44">${fmtEur(u.Savings)}</td>
      <td style="padding:10px 12px;text-align:right;color:var(--p-peach);font-weight:600">${u.CompCost > 0 ? fmtEur(u.CompCost) : ''}</td>
    </tr>`
  ).join('');
  document.getElementById('modal-content').innerHTML = `
    <h2 style="font-size:17px;color:${t.color};margin-bottom:16px">${escHtml(t.label)}</h2>
    <table style="width:100%;border-collapse:collapse;font-size:13px">
      <thead>
        <tr style="background:#181835;border-bottom:1px solid #2a2a55">
          <th style="text-align:left;padding:10px 12px;font-weight:600;color:#9898b8">User</th>
          <th style="text-align:left;padding:10px 12px;font-weight:600;color:#9898b8">Department</th>
          <th style="text-align:left;padding:10px 12px;font-weight:600;color:#9898b8">Category</th>
          <th style="text-align:right;padding:10px 12px;font-weight:600;color:#9898b8">Annual Cost</th>
          <th style="text-align:right;padding:10px 12px;font-weight:600;color:#9898b8">Savings/yr</th>
          <th style="text-align:right;padding:10px 12px;font-weight:600;color:var(--p-peach)">Cost/yr</th>
        </tr>
      </thead>
      <tbody>${tableRows || '<tr><td colspan="6" style="padding:16px;text-align:center;color:#6a6a8e">No matching users found</td></tr>'}</tbody>
    </table>
    ${totalSav > 0 ? `<div style="margin-top:12px;text-align:right;font-size:13px;font-weight:700;color:#2f9e44">Total potential savings: ${fmtEur(totalSav)}/yr</div>` : ''}
    ${totalComp > 0 ? `<div style="margin-top:4px;text-align:right;font-size:13px;font-weight:700;color:var(--p-peach)">Total potential compliance cost: ${fmtEur(totalComp)}/yr</div>` : ''}
    <div style="margin-top:8px;font-size:11px;color:#6a6a8e">Click any row to view the full assessment.</div>`;
  document.getElementById('modal-overlay').classList.add('open');
}

function showTileUserDetail(idx) {
  const u = tileModalUsers[idx];
  if (!u) return;
  showUserDetail(u);
}

function showPoolModal() {
  clearBackState();
  const rows = POOL_SKUS.filter(s => s.waste > 0);
  const total = rows.reduce((s,r) => s + r.waste, 0);
  const tableRows = rows.map(s =>
    `<tr style="border-bottom:1px solid rgba(255,255,255,.04)">
      <td style="padding:10px 12px;font-weight:500">${escHtml(s.sku)}</td>
      <td style="padding:10px 12px;text-align:right;color:#9898b8">${s.unassigned.toLocaleString()}</td>
      <td style="padding:10px 12px;text-align:right;font-weight:600;color:#ef6ea7">${fmtEur(s.waste)}/yr</td>
    </tr>`
  ).join('');
  document.getElementById('modal-content').innerHTML = `
    <h2 style="font-size:17px;color:#9898b8;margin-bottom:4px">Unassigned Licenses</h2>
    <div style="font-size:12px;color:#6a6a8e;margin-bottom:16px">Paid licenses in the tenant pool with unassigned seats generating waste</div>
    <table style="width:100%;border-collapse:collapse;font-size:13px">
      <thead>
        <tr style="background:#181835;border-bottom:1px solid #2a2a55">
          <th style="text-align:left;padding:10px 12px;font-weight:600;color:#9898b8">License SKU</th>
          <th style="text-align:right;padding:10px 12px;font-weight:600;color:#9898b8">Unassigned Seats</th>
          <th style="text-align:right;padding:10px 12px;font-weight:600;color:#9898b8">Annual Waste</th>
        </tr>
      </thead>
      <tbody>${tableRows || '<tr><td colspan="3" style="padding:16px;text-align:center;color:#6a6a8e">No unassigned license waste found</td></tr>'}</tbody>
    </table>
    ${rows.length > 1 ? `<div style="margin-top:12px;text-align:right;font-size:13px;font-weight:700;color:#ef6ea7">Total: ${fmtEur(total)}/yr</div>` : ''}`;
  document.getElementById('modal-overlay').classList.add('open');
}

// ── TAB 1: User table ─────────────────────────────────────────────────────────
let sortCol = 'Savings', sortAsc = false;
let filteredData = [];
let capFilteredData = [];
let capSortCol = 'sav', capSortAsc = false;

(function() {
  const sel = document.getElementById('cat-filter');
  ALL_CATS.forEach(c => {
    const o = document.createElement('option');
    o.value = c;
    o.textContent = c;
    sel.appendChild(o);
  });
  // Populate department filter
  const dSel = document.getElementById('dept-filter');
  if (dSel) {
    DEPTS.forEach(d => {
      const o = document.createElement('option');
      o.value = d;
      o.textContent = d;
      dSel.appendChild(o);
    });
  }
  // Render subscription alerts
  renderSubAlerts();
  // Render Copilot ROI
  renderCopilotRoi();
})();

function sortTable(col) {
  if (sortCol === col) sortAsc = !sortAsc; else { sortCol = col; sortAsc = (col !== 'Savings' && col !== 'Cost'); }
  document.querySelectorAll('th[data-col]').forEach(th => th.classList.toggle('sorted', th.dataset.col === col));
  renderUserTable();
}

function renderUserTable() {
  const q    = document.getElementById('user-filter').value.toLowerCase();
  const cat  = document.getElementById('cat-filter').value;
  const dept = document.getElementById('dept-filter') ? document.getElementById('dept-filter').value : '';
  const isFiltered = q || cat || dept;
  let data = USERS.filter(u => {
    // All users with recommendations are shown — zero-impact rows included for completeness
    if (q && !(u.Name||'').toLowerCase().includes(q) && !(u.UPN||'').toLowerCase().includes(q) && !(u.Dept||'').toLowerCase().includes(q)) return false;
    if (cat && u.Category !== cat) return false;
    if (dept && (u.Dept||'') !== dept) return false;
    return true;
  });
  data.sort((a,b) => {
    let va = a[sortCol] != null ? a[sortCol] : '', vb = b[sortCol] != null ? b[sortCol] : '';
    if (typeof va === 'number') return sortAsc ? va-vb : vb-va;
    va = String(va).toLowerCase(); vb = String(vb).toLowerCase();
    return sortAsc ? va.localeCompare(vb) : vb.localeCompare(va);
  });
  filteredData = data;
  document.getElementById('user-count').textContent = data.length;
  const maxSav = data.length ? Math.max(...data.map(u => u.Savings||0), 1) : 1;
  document.getElementById('user-tbody').innerHTML = data.map((u,i) => {
    const bg = savingsColor(u.Savings, maxSav);
    const tc = savingsTextColor(u.Savings, maxSav);
    return `<tr class="clickable-row" onclick="showUserModal(${i})">
      <td><div style="font-weight:500">${escHtml(u.Name||u.UPN)}${admBadge(u.AdminPriv)}</div><div style="font-size:11px;color:#6a6a8e">${escHtml(u.UPN||'')}</div></td>
      <td>${escHtml(u.Dept||'')}</td>
      <td><span class="savings-cell" style="background:${bg};color:${tc}">${fmtEur(u.Savings)}</span></td>
      <td>${u.CompCost > 0 ? `<span class="compcost-cell">${fmtEur(u.CompCost)}</span>` : ''}</td>
      <td style="font-size:11px;color:#9898b8;max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${escHtml(u.Licenses||'')}">${escHtml((u.Licenses||'').replace(/;/g,', '))}</td>
      <td><span class="cat-badge">${escHtml(u.Category||'')}</span>${renderTags(u.Tags)}</td>
    </tr>`;
  }).join('') || '<tr><td colspan="6" style="text-align:center;padding:20px;color:#6a6a8e">No matching users</td></tr>';
}

// ── Modal ─────────────────────────────────────────────────────────────────────
function showUserModal(idx) {
  const u = filteredData[idx];
  if (!u) return;
  activeTileIdx = -1;
  showUserDetail(u);
}

function showUserDetail(u) {
  const licenses = (u.Licenses||'').split(';').map(l => l.trim()).filter(l => l);
  const licHtml = licenses.length
    ? '<ul style="padding-left:16px;margin:0">' + licenses.map(l => `<li>${escHtml(l)}</li>`).join('') + '</ul>'
    : '<span class="text-muted">—</span>';
  const mc = document.getElementById('modal-content');
  const mb = document.getElementById('modal-box');
  mc.innerHTML = `
    <h2 style="font-size:17px;color:#3ddad7;margin-bottom:4px">${escHtml(u.Name||u.UPN)}${admBadge(u.AdminPriv)}</h2>
    <div style="font-size:12px;color:#6a6a8e;margin-bottom:16px">${escHtml(u.UPN||'')}</div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:16px">
      <div class="modal-field">
        <div class="mf-label">Department</div>
        <div class="mf-value">${escHtml(u.Dept||'—')}</div>
      </div>
      <div class="modal-field">
        <div class="mf-label">Category</div>
        <div class="mf-value"><span class="cat-badge">${escHtml(u.Category||'')}</span>${renderTags(u.Tags)}</div>
      </div>
      <div class="modal-field">
        <div class="mf-label">Annual License Cost</div>
        <div class="mf-value" style="font-weight:600">${fmtEur(u.Cost)}</div>
      </div>
      <div class="modal-field">
        <div class="mf-label">Est. Potential Savings/yr</div>
        <div class="mf-value" style="font-weight:700;color:#2f9e44">${fmtEur(u.Savings)}</div>
      </div>
      ${u.CompCost > 0 ? `<div class="modal-field"></div><div class="modal-field">
        <div class="mf-label">Est. Potential Compliance Cost/yr</div>
        <div class="mf-value" style="font-weight:700;color:var(--p-peach)">${fmtEur(u.CompCost)}</div>
      </div>` : ''}
    </div>
    <hr class="modal-divider">
    <div class="modal-field">
      <div class="mf-label">Assigned Licenses</div>
      <div class="mf-value" style="margin-top:4px">${licHtml}</div>
    </div>
    <hr class="modal-divider">
    <div class="modal-field">
      <div class="mf-label">Assessments</div>
      <div class="modal-rec" style="margin-top:6px">${formatRec(u.Rec)}</div>
    </div>`;
  if (activeTileIdx >= 0) {
    setTimeout(function() { mb.style.cursor = 'pointer'; mb.dataset.backTile = activeTileIdx; }, 0);
  } else if (mb.dataset.backSku !== undefined) {
    const skuBack = mb.dataset.backSku;
    setTimeout(function() { mb.style.cursor = 'pointer'; mb.dataset.backSku = skuBack; }, 0);
  } else {
    mb.style.cursor = '';
    delete mb.dataset.backTile;
    delete mb.dataset.backSku;
  }
  document.getElementById('modal-overlay').classList.add('open');
}

function closeModal() {
  clearBackState();
  document.getElementById('modal-overlay').classList.remove('open');
}

document.addEventListener('keydown', function(e) { if (e.key === 'Escape') closeModal(); });
document.getElementById('modal-box').addEventListener('click', function(e) {
  if (e.target.closest('tr') || e.target.closest('a')) return;
  const bt = this.dataset.backTile;
  if (bt != null) { var ti = parseInt(bt, 10); if (!isNaN(ti)) { e.stopPropagation(); showTileModal(ti); return; } }
  const bs = this.dataset.backSku;
  if (bs != null) { var si = parseInt(bs, 10); if (!isNaN(si)) { e.stopPropagation(); showSkuModal(si); } }
});

// ── TAB 2: SKU chart ──────────────────────────────────────────────────────────
const CAT_COLORS = {
  // Longer/more-specific keys MUST come before shorter ones
  // because catColor() uses .includes() which is a substring match.
  // Palette: teal #3ddad7, blue #5b8def, peach #ff9f80, pink #ef6ea7,
  //          navy #5b89b6, muted #6a6a8e, plus tinted variants per category.
  'dormant admin':        '#ef6ea7',       // pink
  'dormant cloud':        '#d94070',       // deep pink
  'stale sign-in':        '#ff9f80',       // peach
  'dormant':              '#ef6ea7',       // pink
  'disabled':             '#ff8a65',       // warm peach
  'no activity':          '#ffb380',       // light peach
  'zero':                 '#ffb380',       // light peach
  'never signed':         '#f0a070',       // amber peach
  'inactive hold':        '#ff8a65',       // warm peach
  'inactive add-on review': '#7a8fc7',     // soft blue
  'inactive add-on':      '#5b8def',       // blue
  'inactive mailbox':     '#5b8def',       // blue
  'inactive':             '#5b8def',       // blue
  'shared mailbox':       '#3ddad7',       // teal
  'duplicate coverage':   '#5b8def',       // blue
  'duplicate review':     '#7a8fc7',       // soft blue
  'duplicate':            '#5b8def',       // blue
  'standalone':           '#6da0e0',       // mid blue
  'overlapping':          '#7a8fc7',       // soft blue
  'teams unbundling':     '#2ec4b6',       // deep teal
  'e5 voice':             '#8b7ed8',       // soft purple
  'e5 data':              '#5b89b6',       // navy
  'a la carte':           '#ff8a65',       // warm peach
  'frontline':            '#3ddad7',       // teal
  'data gap':             '#6a6a8e',       // muted
  'mailbox storage':      '#ffb380',       // light peach
  'viral license':        '#d94070',       // deep pink
  'windows license':      '#8b7ed8',       // soft purple
  'unlicensed with data': '#ef6ea7',       // pink
  'compliance':           '#d94070',       // deep pink
  'copilot reclaim':      '#5b89b6',       // navy
  'copilot at risk':      '#2ec4b6',       // deep teal
  'copilot':              '#5b89b6',       // navy
  'reclaim':              '#5b89b6',       // navy
  'at risk':              '#2ec4b6',       // deep teal
  'add-on':               '#8b7ed8',       // soft purple
  'visio':                '#8b7ed8',       // soft purple
  'project':              '#8b7ed8',       // soft purple
  'pbi':                  '#8b7ed8',       // soft purple
  'guest':                '#a07ed8',       // purple
  'non-human':            '#a07ed8',       // purple
  'automation':           '#2ec4b6',       // deep teal
  'admin review':         '#5b89b6',       // navy
  'admin':                '#5b89b6',       // navy
  'default':              '#6a6a8e'        // muted
};

function catColor(catName) {
  const lower = (catName||'').toLowerCase();
  for (const key of Object.keys(CAT_COLORS)) {
    if (lower.includes(key)) return CAT_COLORS[key];
  }
  return CAT_COLORS['default'];
}

function showSkuCatModal(skuIdx, segIdx) {
  // Resolve category name from segment index (avoids JS string injection)
  const s0 = SKUS[skuIdx];
  const cats0 = Array.isArray(s0.cats) ? s0.cats : (s0.cats ? [s0.cats] : []);
  const assignedTotal0 = cats0.reduce((sum, c) => sum + (c.val || 0), 0);
  const allSegs0 = [...cats0];
  if (Math.max(0, Math.round((s0.waste - assignedTotal0) * 100) / 100) > 0.01) allSegs0.push({ cat: 'Unassigned', val: 0 });
  const catName = (allSegs0[segIdx] || {}).cat || 'Unknown';
  activeTileIdx = -1;
  clearBackState();
  const s = SKUS[skuIdx];
  const skuName = s.lic;
  // For "Unassigned" category, show pool data instead of user list
  if (catName === 'Unassigned') {
    const poolMatch = POOL_SKUS.filter(p => p.sku === skuName);
    if (poolMatch.length) {
      const p = poolMatch[0];
      document.getElementById('modal-content').innerHTML = `
        <h3 style="margin-bottom:4px">${escHtml(skuName)} \u2014 Unassigned Seats</h3>
        <p style="color:#6a6a8e;margin-bottom:16px">${p.unassigned} unassigned seat(s) \u2022 Annual waste: ${fmtEur(p.waste)}/yr</p>
        <p style="font-size:13px;color:#9898b8">These are paid license seats in the tenant pool that are not assigned to any user. Consider reducing the subscription quantity at renewal or assigning them to users who need them.</p>`;
    } else {
      document.getElementById('modal-content').innerHTML = `
        <h3 style="margin-bottom:4px">${escHtml(skuName)} \u2014 Unassigned Seats</h3>
        <p style="color:#6a6a8e">No detailed pool data available for this SKU.</p>`;
    }
    const mb = document.getElementById('modal-box');
    mb.dataset.backSku = skuIdx;
    mb.style.cursor = 'pointer';
    document.getElementById('modal-overlay').classList.add('open');
    return;
  }
  // Filter users who have this SKU AND match this category (primary or secondary)
  const catLower = catName.toLowerCase();
  const matched = USERS.filter(u => {
    const hasLic = (u.Licenses||'').split(';').some(l => l.trim() === skuName);
    if (!hasLic) return false;
    if ((u.Category||'').toLowerCase() === catLower) return true;
    if (u.Tags && u.Tags.some(t => t.toLowerCase() === catLower)) return true;
    // Also check recommendation text for the category pattern
    if ((u.Rec||'').toLowerCase().includes(catLower)) return true;
    return false;
  }).sort((a,b) => b.Savings - a.Savings);
  tileModalUsers = matched;
  const totalSav = matched.reduce((sum,u) => sum + u.Savings, 0);
  const totalComp = matched.reduce((sum,u) => sum + (u.CompCost||0), 0);
  const tableRows = matched.map((u, i) =>
    `<tr style="border-bottom:1px solid rgba(255,255,255,.04);cursor:pointer" onclick="showTileUserDetail(${i})" title="Click for full assessment">
      <td style="padding:10px 12px"><div style="font-weight:500">${escHtml(u.Name||u.UPN)}${admBadge(u.AdminPriv)}</div><div style="font-size:11px;color:#6a6a8e">${escHtml(u.UPN||'')}</div></td>
      <td style="padding:10px 12px">${escHtml(u.Dept||'')}</td>
      <td style="padding:10px 12px"><span class="cat-badge">${escHtml(u.Category||'')}</span>${renderTags(u.Tags)}</td>
      <td style="padding:10px 12px;text-align:right">${fmtEur(u.Cost)}</td>
      <td style="padding:10px 12px;text-align:right;font-weight:600;color:#2f9e44">${fmtEur(u.Savings)}</td>
      <td style="padding:10px 12px;text-align:right;color:var(--p-peach);font-weight:600">${u.CompCost > 0 ? fmtEur(u.CompCost) : ''}</td>
    </tr>`
  ).join('');
  const color = catColor(catName);
  document.getElementById('modal-content').innerHTML = `
    <h3 style="margin-bottom:4px">${escHtml(skuName)}</h3>
    <p style="color:#6a6a8e;margin-bottom:16px"><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:${color};vertical-align:middle;margin-right:4px"></span>${escHtml(catName)} \u2022 ${matched.length} user(s) \u2022 Potential savings: ${fmtEur(totalSav)}/yr</p>
    <table style="width:100%;border-collapse:collapse;font-size:13px">
      <thead><tr style="background:#181835;font-size:12px;color:#9898b8">
        <th style="text-align:left;padding:8px 12px">User</th>
        <th style="text-align:left;padding:8px 12px">Department</th>
        <th style="text-align:left;padding:8px 12px">Category</th>
        <th style="text-align:right;padding:8px 12px">License Cost</th>
        <th style="text-align:right;padding:8px 12px">Potential Savings</th>
        <th style="text-align:right;padding:8px 12px;color:var(--p-peach)">Cost/yr</th>
      </tr></thead>
      <tbody>${tableRows || '<tr><td colspan="6" style="padding:16px;text-align:center;color:#6a6a8e">No matching users found</td></tr>'}</tbody>
    </table>
    ${totalSav > 0 ? `<div style="margin-top:12px;text-align:right;font-size:13px;font-weight:700;color:#2f9e44">Total potential savings: ${fmtEur(totalSav)}/yr</div>` : ''}
    ${totalComp > 0 ? `<div style="margin-top:4px;text-align:right;font-size:13px;font-weight:700;color:var(--p-peach)">Total potential compliance cost: ${fmtEur(totalComp)}/yr</div>` : ''}`;
  const mb = document.getElementById('modal-box');
  mb.dataset.backSku = skuIdx;
  mb.style.cursor = 'pointer';
  document.getElementById('modal-overlay').classList.add('open');
}

function showSkuModal(skuIdx) {
  activeTileIdx = -1;
  clearBackState();
  const s = SKUS[skuIdx];
  const skuName = s.lic;
  // Match users who have this SKU in their license list
  const matched = USERS.filter(u => (u.Licenses||'').split(';').some(l => l.trim() === skuName))
    .sort((a,b) => b.Savings - a.Savings);
  tileModalUsers = matched;
  const totalSav = matched.reduce((sum,u) => sum + u.Savings, 0);
  const totalComp = matched.reduce((sum,u) => sum + (u.CompCost||0), 0);
  const tableRows = matched.map((u, i) =>
    `<tr style="border-bottom:1px solid rgba(255,255,255,.04);cursor:pointer" onclick="showTileUserDetail(${i})" title="Click for full assessment">
      <td style="padding:10px 12px"><div style="font-weight:500">${escHtml(u.Name||u.UPN)}${admBadge(u.AdminPriv)}</div><div style="font-size:11px;color:#6a6a8e">${escHtml(u.UPN||'')}</div></td>
      <td style="padding:10px 12px">${escHtml(u.Dept||'')}</td>
      <td style="padding:10px 12px"><span class="cat-badge">${escHtml(u.Category||'')}</span>${renderTags(u.Tags)}</td>
      <td style="padding:10px 12px;text-align:right">${fmtEur(u.Cost)}</td>
      <td style="padding:10px 12px;text-align:right;font-weight:600;color:#2f9e44">${fmtEur(u.Savings)}</td>
      <td style="padding:10px 12px;text-align:right;color:var(--p-peach);font-weight:600">${u.CompCost > 0 ? fmtEur(u.CompCost) : ''}</td>
    </tr>`
  ).join('');
  const mc = document.getElementById('modal-content');
  mc.innerHTML = `
    <h3 style="margin-bottom:4px">${escHtml(skuName)}</h3>
    <p style="color:#6a6a8e;margin-bottom:16px">${matched.length} user(s) with assessments \u2022 Potential savings: ${fmtEur(totalSav)}/yr \u2022 Total waste: ${fmtEur(s.waste)}/yr</p>
    <table style="width:100%;border-collapse:collapse">
      <thead><tr style="background:#181835;font-size:12px;color:#9898b8">
        <th style="text-align:left;padding:8px 12px">User</th>
        <th style="text-align:left;padding:8px 12px">Department</th>
        <th style="text-align:left;padding:8px 12px">Category</th>
        <th style="text-align:right;padding:8px 12px">License Cost</th>
        <th style="text-align:right;padding:8px 12px">Potential Savings</th>
        <th style="text-align:right;padding:8px 12px;color:var(--p-peach)">Cost/yr</th>
      </tr></thead>
      <tbody>${tableRows}</tbody>
    </table>
    ${totalSav > 0 ? `<div style="margin-top:12px;text-align:right;font-size:13px;font-weight:700;color:#2f9e44">Total potential savings: ${fmtEur(totalSav)}/yr</div>` : ''}
    ${totalComp > 0 ? `<div style="margin-top:4px;text-align:right;font-size:13px;font-weight:700;color:var(--p-peach)">Total potential compliance cost: ${fmtEur(totalComp)}/yr</div>` : ''}`;
  // Set up back navigation from user detail
  const mb = document.getElementById('modal-box');
  mb.dataset.backSku = skuIdx;
  mb.style.cursor = 'pointer';
  document.getElementById('modal-overlay').classList.add('open');
}

function renderSkuChart() {
  const chart = document.getElementById('sku-chart');
  if (!Array.isArray(SKUS) || !SKUS.length) { chart.innerHTML = '<p class="text-muted">No SKU data available.</p>'; return; }
  const maxWaste = SKUS[0].waste;

  // Collect all category names for the legend
  const allCats = [];
  SKUS.forEach(s => {
    const cats = Array.isArray(s.cats) ? s.cats : (s.cats ? [s.cats] : []);
    cats.forEach(c => { if (c.cat && !allCats.includes(c.cat)) allCats.push(c.cat); });
  });

  // Add unassigned to legend if any SKU has it
  const hasUnassigned = SKUS.some(s => {
    const cats = Array.isArray(s.cats) ? s.cats : (s.cats ? [s.cats] : []);
    return Math.max(0, s.waste - cats.reduce((sum,c) => sum+(c.val||0), 0)) > 0.01;
  });
  const legendCats = hasUnassigned ? [...allCats, 'Unassigned'] : allCats;
  const legendHtml = '<div style="display:flex;flex-wrap:wrap;gap:10px;margin-bottom:18px">' +
    legendCats.map(c => {
      const color = c === 'Unassigned' ? '#ced4da' : catColor(c);
      return `<span style="display:inline-flex;align-items:center;gap:5px;font-size:11px;color:#9898b8"><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:${color}"></span>${escHtml(c)}</span>`;
    }).join('') +
    '</div>';

  const barsHtml = SKUS.map(s => {
    const cats = Array.isArray(s.cats) ? s.cats : (s.cats ? [s.cats] : []);
    const barPct = maxWaste > 0 ? (s.waste / maxWaste * 100) : 0;
    // Build stacked segments
    const assignedTotal = cats.reduce((sum, c) => sum + (c.val || 0), 0);
    const unassigned = Math.max(0, Math.round((s.waste - assignedTotal) * 100) / 100);
    const allSegs = [...cats];
    if (unassigned > 0.01) allSegs.push({ cat: 'Unassigned', val: unassigned });
    const segments = allSegs.map(c => {
      const segPct = s.waste > 0 ? (c.val / s.waste * 100).toFixed(2) : 0;
      const color = c.cat === 'Unassigned' ? '#ced4da' : catColor(c.cat);
      const tip = `${c.cat}: ${fmtEur(c.val)} \u2014 click to view users`;
      const skuI = SKUS.indexOf(s);
      const segI = allSegs.indexOf(c);
      return `<div title="${escHtml(tip)}" onclick="event.stopPropagation();showSkuCatModal(${skuI},${segI})" style="width:${segPct}%;background:${color};height:100%;display:inline-block;vertical-align:top;cursor:pointer;transition:opacity .15s" onmouseenter="this.style.opacity='.75'" onmouseleave="this.style.opacity='1'"></div>`;
    }).join('');
    const tipTxt = `${s.lic}: ${fmtEur(s.waste)}\n` + allSegs.map(c => `${c.cat}: ${fmtEur(c.val)}`).join('\n');
    const skuIdx = SKUS.indexOf(s);
    return `<div class="sku-row" title="${escHtml(tipTxt)}" onclick="showSkuModal(${skuIdx})" style="cursor:pointer">
      <div class="sku-name" title="${escHtml(s.lic)}">${escHtml(s.lic)}</div>
      <div class="sku-bar-wrap" style="position:relative">
        <div style="width:${barPct.toFixed(1)}%;height:100%;display:flex;overflow:hidden;border-radius:4px">${segments}</div>
      </div>
      <div class="sku-amount">${fmtEur(s.waste)}</div>
      <div class="sku-users">${s.users} users</div>
    </div>`;
  }).join('');

  chart.innerHTML = legendHtml + barsHtml;
}

// ── TAB 3: Per-user capability matrix ────────────────────────────────────────
const CAP_KEYS = [
  {k:'ex',kU:'exU',kI:'exI',label:'Exchange'},
  {k:'tm',kU:'tmU',kI:'tmI',label:'Teams'},
  {k:'dt',kU:'dtU',kI:null, label:'Desktop'},
  {k:'od',kU:'odU',kI:'odI',label:'OneDrive'},
  {k:'sp',kU:'spU',kI:'spI',label:'SharePoint'},
  {k:'co',kU:'coU',kI:null, label:'Copilot'}
];

function sortCapTable(col) {
  if (capSortCol === col) capSortAsc = !capSortAsc; else { capSortCol = col; capSortAsc = (col !== 'sav' && col !== 'comp'); }
  document.querySelectorAll('th[data-capcol]').forEach(th => th.classList.toggle('sorted', th.dataset.capcol === col));
  renderCapMatrix();
}

function renderCapMatrix() {
  const tbody = document.getElementById('cap-tbody');
  const capArr = Array.isArray(CAP_USERS) ? CAP_USERS : (CAP_USERS ? [CAP_USERS] : []);
  if (!capArr.length) {
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:20px;color:#6a6a8e">No data</td></tr>';
    return;
  }
  const q = (document.getElementById('cap-filter').value || '').toLowerCase();
  let data = q ? capArr.filter(u => (u.n||'').toLowerCase().includes(q) || (u.upn||'').toLowerCase().includes(q)) : [...capArr];
  data.sort((a,b) => {
    let va = a[capSortCol] != null ? a[capSortCol] : '', vb = b[capSortCol] != null ? b[capSortCol] : '';
    if (typeof va === 'number') return capSortAsc ? va-vb : vb-va;
    va = String(va).toLowerCase(); vb = String(vb).toLowerCase();
    return capSortAsc ? va.localeCompare(vb) : vb.localeCompare(va);
  });
  capFilteredData = data;
  const maxSav = data.length ? Math.max(...data.map(u => u.sav||0), 1) : 1;
  tbody.innerHTML = data.map((u, idx) => {
    const bg = savingsColor(u.sav||0, maxSav);
    const tc = savingsTextColor(u.sav||0, maxSav);
    const cells = CAP_KEYS.map(ck => {
      const intensity = ck.kI ? (u[ck.kI]||'') : '';
      const c = capCellUser(u[ck.k], u[ck.kU], intensity);
      const tip = ck.label + ': ' + (u[ck.k] ? 'provisioned' : 'not provisioned') + ' / ' + (u[ck.kU] ? 'in use' : 'not in use') + (intensity ? ' (' + intensity + ')' : '');
      return `<td title="${escHtml(tip)}"><span class="cap-cell" style="background:${c.bg};color:${c.text}">${c.label}</span></td>`;
    }).join('');
    return `<tr class="clickable-row" onclick="showCapUserModal(${idx})">
      <td class="user-name"><div style="font-weight:500;white-space:nowrap">${escHtml(u.n||u.upn||'')}${admBadge(u.adm)}</div><div style="font-size:10px;color:#6a6a8e;white-space:nowrap">${escHtml(u.upn||'')}</div></td>
      <td><span class="savings-cell" style="background:${bg};color:${tc}">${fmtEur(u.sav||0)}</span></td>
      <td>${u.comp > 0 ? `<span class="compcost-cell">${fmtEur(u.comp)}</span>` : ''}</td>
      ${cells}
      <td style="text-align:left"><span class="cat-badge" style="white-space:nowrap">${escHtml(u.cat||'')}</span>${renderTags(u.tags)}</td>
    </tr>`;
  }).join('') || '<tr><td colspan="10" style="text-align:center;padding:20px;color:#6a6a8e">No matching users</td></tr>';
}

function showCapUserModal(idx) {
  const cu = capFilteredData[idx];
  if (!cu) return;
  const u = USERS.find(x => x.UPN === cu.upn);
  if (u) { activeTileIdx = -1; clearBackState(); showUserDetail(u); }
}

// ── License Groups tab ────────────────────────────────────────────────────────
function renderGroupTable() {
  const tbody = document.getElementById('group-tbody');
  const countEl = document.getElementById('group-count');
  const emptyEl = document.getElementById('group-empty');
  const filter = (document.getElementById('group-filter').value || '').toLowerCase();
  const filtered = GROUPS.filter(g => {
    if (!filter) return true;
    return (g.name||'').toLowerCase().includes(filter) || (g.skus||'').toLowerCase().includes(filter);
  });
  countEl.textContent = filtered.length;
  if (!filtered.length) {
    tbody.innerHTML = '';
    emptyEl.textContent = GROUPS.length ? 'No groups match the filter.' : 'No group-based licensing detected in this tenant.';
    return;
  }
  emptyEl.textContent = '';
  const typeBadge = t => {
    const isDyn = (t||'').toLowerCase() === 'dynamic';
    const bg  = isDyn ? 'rgba(239,110,167,.15)' : 'rgba(61,218,215,.15)';
    const col = isDyn ? '#ef6ea7' : '#3ddad7';
    return '<span style="display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:500;background:'+bg+';color:'+col+'">'+escHtml(t)+'</span>';
  };
  const skuBadges = s => (s||'').split('; ').map(sku =>
    '<span style="display:inline-block;padding:1px 6px;border-radius:3px;font-size:11px;background:rgba(255,255,255,.06);color:#9898b8;margin:1px 2px">'+escHtml(sku.trim())+'</span>'
  ).join(' ');
  tbody.innerHTML = filtered.map(g =>
    '<tr><td style="text-align:left;font-weight:500">'+escHtml(g.name)+'</td>'
    +'<td>'+typeBadge(g.type)+'</td>'
    +'<td style="text-align:center;font-weight:600">'+g.members+'</td>'
    +'<td style="text-align:center">'+(g.directOnly > 0 ? g.directOnly : g.directOnly === 0 ? '0' : '')+'</td>'
    +'<td style="text-align:center;color:#9898b8">'+(g.skuSeats||'')+'</td>'
    +'<td style="text-align:left">'+skuBadges(g.skus)+'</td></tr>'
  ).join('');
}

// ── Welcome overlay ──────────────────────────────────────────────────────────
function showWelcome() {
  document.getElementById('welcome-overlay').classList.add('open');
}
function closeWelcome() {
  document.getElementById('welcome-overlay').classList.remove('open');
  if (document.getElementById('welcome-hide-cb').checked) {
    try { localStorage.setItem('m365loa_welcome_dismissed', '1'); } catch(e) {}
  }
}
(function() {
  try {
    if (!localStorage.getItem('m365loa_welcome_dismissed')) showWelcome();
  } catch(e) { showWelcome(); }
})();

// ── Init ──────────────────────────────────────────────────────────────────────
renderDashboard();
renderUserTable();
renderSkuChart();
renderCapMatrix();
renderGroupTable();
</script>
</body>
</html>
'@

# ── Write output ──────────────────────────────────────────────────────────────
[System.IO.File]::WriteAllText($HtmlOutput, $html, [System.Text.Encoding]::UTF8)
$size = [math]::Round((Get-Item $HtmlOutput).Length / 1KB, 1)
Write-Host ""
Write-Host "  Heatmap dashboard written:" -ForegroundColor Green
Write-Host "  $HtmlOutput ($size KB)" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Open in your browser to present to the customer." -ForegroundColor Gray
