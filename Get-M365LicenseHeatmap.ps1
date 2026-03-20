# ========================================================
# M365 License Optimization Dashboard
# Version : 1.1.0
# Author  : Bruno Vijverman
# Reads the CSV output from Get-M365LicenseOptimizationReport.ps1
# and generates a standalone HTML heatmap dashboard.
# ========================================================

<#
.SYNOPSIS
    Generates an interactive HTML heatmap dashboard from M365 License Optimization report output.

.DESCRIPTION
    Reads the CSV files produced by Get-M365LicenseOptimizationReport.ps1 and creates a
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
    .\Get-M365LicenseHeatmap.ps1
    .\Get-M365LicenseHeatmap.ps1 -OutputFolder "C:\Reports\Contoso"
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

# ── Disclaimer text (from executive summary or hardcoded fallback) ───────────
$disclaimerRow = $summaryRows | Where-Object { $_.'Category' -match 'DISCLAIMER' } | Select-Object -First 1
$disclaimerText = if ($disclaimerRow) { ($disclaimerRow.'Category' -replace '^DISCLAIMER:\s*','').Trim() }
              else { 'All cost figures are indicative estimates based on public Microsoft list prices (EUR). Actual costs may differ due to EA/CSP/volume pricing. Verify against your invoice.' }

# ── Tier-1 categories (full license cost = reclaimable savings) ──────────────
$tier1 = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
@('Dormant','Disabled Account','Inactive Hold With License','Inactive Hold','No Activity',
  'Shared Mailbox','Never Signed In','Guest Account Review','Guest User','Non-Human Account Review',
  'Admin Review','Automation Account','Dormant Admin Review','Legacy Service Account',
  'Dormant Cloud PC','Inactive Add-On','Inactive Add-On Review','Free License Overlap') | ForEach-Object { [void]$tier1.Add($_) }

# ── Cost categories (amounts in recommendations are costs, NOT savings) ──────
# These categories flag users who NEED additional licenses — the €/yr in the text
# is the estimated compliance cost, not a potential savings.
$costCategories = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
@('Licensing Compliance Gap','Overlapping License') | ForEach-Object { [void]$costCategories.Add($_) }

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

# ── Per-user savings estimation ──────────────────────────────────────────────
function Get-EstimatedSavings([string]$category,[decimal]$annualCost,[string]$recommendation) {
    if ($costCategories.Contains($category)) { return [decimal]0 }
    if ($tier1.Contains($category)) { return $annualCost }
    [decimal]$total = 0
    # Pattern 1: explicit /yr amounts  (e.g. "Saves \u20AC3,20/mo (\u20AC38,40/yr)")
    foreach ($m in [regex]::Matches($recommendation, '\u20AC([\d.,]+)/yr')) {
        $total += Parse-Decimal $m.Groups[1].Value
    }
    # Pattern 2: Annual waste (Duplicate Coverage / A La Carte)
    foreach ($m in [regex]::Matches($recommendation, 'Annual waste: \u20AC([\d.,]+)')) {
        $total += Parse-Decimal $m.Groups[1].Value
    }
    # Pattern 3: removed — Overlapping License is now a zero-savings hygiene category
    return $total
}

# ── Build per-user heatmap data ───────────────────────────────────────────────
Write-Host "  Processing users..." -ForegroundColor Gray
# ── Recommendation prefix → category mapping (for secondary tags) ────────────
$_recPrefixMap = [ordered]@{
    'INACTIVE HOLD WITH LICENSE' = 'Inactive Hold With License'
    'INACTIVE HOLD'       = 'Inactive Hold'
    'DISABLED ACCOUNT'    = 'Disabled Account'
    'DISABLED SHARED'     = 'Disabled Account'
    'SHARED MAILBOX'      = 'Shared Mailbox'
    'OVERLAPPING LICENSE' = 'Overlapping License'
    'DUPLICATE COVERAGE'  = 'Duplicate Coverage'
    'LICENSING CHECK'     = 'Licensing Compliance Gap'
    'ADMIN'               = 'Admin Review'
    'DORMANT ADMIN'       = 'Dormant Admin Review'
    'DORMANT CLOUD PC'    = 'Dormant Cloud PC'
    'CLOUD PC REVIEW'     = 'Cloud PC Review'
    'AUTOMATION ACCOUNT'  = 'Automation Account'
    'PREMIUM ADD-ON REVIEW'= 'Premium Add-On Review'
    'TEAMS UNBUNDLING'    = 'Teams Unbundling'
    'E5 VOICE'            = 'E5 Voice Review'
    'EXCHANGE KIOSK'      = 'Exchange Kiosk Downgrade'
    'FORWARDING MAILBOX'  = 'Forwarding Mailbox'
    'COPILOT ACTIVE'      = 'Copilot Active'
    'COPILOT RECLAIM'     = 'Copilot Reclaim'
    'COPILOT WATCHLIST'   = 'Copilot Watchlist'
    'GUEST ACCOUNT'       = 'Guest User'
    'NON-HUMAN'           = 'Non-Human Account'
    'FREE LICENSE'        = 'Free License Overlap'
    'WINDOWS LICENSE'     = 'Windows License Review'
    'FRONTLINE'           = 'Frontline Review'
    'INACTIVE ADD-ON REVIEW' = 'Inactive Add-On Review'
    'INACTIVE ADD-ON'     = 'Inactive Add-On'
    'MAILBOX STORAGE'     = 'Mailbox Storage Warning'
    'EXPENSIVE COLD'      = 'Expensive Cold Storage'
}

$userData = foreach ($r in $rows) {
    $cat  = $r.'Recommendation Category'
    $cost = Parse-Decimal $r.'Annual License Cost (EUR)'
    $rec  = $r.'Recommendation'
    if ($cat -eq 'OK' -or $cat -eq '' -or $cat -eq 'Unlicensed') { continue }
    $savings = Get-EstimatedSavings $cat $cost $rec

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
        Category = $cat
        Tags     = $secondaryCats
        Licenses = $r.'License Friendly Names'
        Rec      = $rec
    }
}
$userData = @($userData)
Write-Host "  $($userData.Count) users with savings opportunities" -ForegroundColor Gray

# ── SKU waste rollup ──────────────────────────────────────────────────────────
$skuRollup = @{}
foreach ($u in $userData) {
    $allLicenses = @($u.Licenses -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    # Only attribute waste to paid SKUs — free licenses don't contribute to savings
    $licenses = @($allLicenses | Where-Object { $_ -notmatch '(?i)\bFree\b|\bTrial\b' })
    if (-not $licenses) { $licenses = $allLicenses }  # fallback if all are free
    if (-not $licenses) { $licenses = @('Unknown') }
    $share = [math]::Round($u.Savings / $licenses.Count, 2)
    if ($share -le 0) { continue }
    foreach ($lic in $licenses) {
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
    # ── Tier 1: User-level waste (full license cost reclaimable) ─────────────
    [PSCustomObject]@{ Label='Dormant Accounts';       Desc='No sign-in >30 days';             CatKey='^dormant$';                        RecKey='';                    Color='#3ddad7' }
    [PSCustomObject]@{ Label='Disabled Accounts';      Desc='Sign-in blocked';                  CatKey='disabled';                         RecKey='';                    Color='#3ddad7' }
    [PSCustomObject]@{ Label='Never Signed In';        Desc='No interactive sign-in on record'; CatKey='never.signed';                     RecKey='NEVER SIGNED IN';     Color='#3ddad7' }
    [PSCustomObject]@{ Label='Zero M365 Usage';        Desc='No app activity in period';        CatKey='no.activity|zero.*usage';          RecKey='NO ACTIVITY detected'; Color='#3ddad7' }
    [PSCustomObject]@{ Label='Admin Review';           Desc='Admin with productivity license';  CatKey='^admin review$';                   RecKey='';                    Color='#3ddad7' }
    [PSCustomObject]@{ Label='Shared Mailbox';         Desc='No license needed under 50 GB';    CatKey='shared.mailbox';                   RecKey='';                    Color='#3ddad7' }
    [PSCustomObject]@{ Label='Guest w/ Paid Licenses'; Desc='B2B guest holding a paid license'; CatKey='guest';                            RecKey='';                    Color='#3ddad7' }
    [PSCustomObject]@{ Label='Automation Accounts';    Desc='Service/automation account';       CatKey='^automation.account$';             RecKey='';                    Color='#3ddad7' }
    [PSCustomObject]@{ Label='Dormant Admin Accounts'; Desc='Admin with no sign-in detected';   CatKey='dormant.admin';                    RecKey='';                    Color='#3ddad7' }

    # ── Cloud PC Utilization ──────────────────────────────────────────────────
    [PSCustomObject]@{ Label='Dormant Cloud PC';       Desc='0 hours connected in 90 days';     CatKey='^dormant.cloud.pc$';               RecKey='';                    Color='#3ddad7' }
    [PSCustomObject]@{ Label='Cloud PC Review';        Desc='< 10 hrs connected in 90 days';    CatKey='^cloud.pc.review$';                RecKey='';                    Color='#3ddad7' }

    # ── Tier 2: License optimization (partial savings) ───────────────────────
    [PSCustomObject]@{ Label='Duplicate Coverage';     Desc='Standalone covered by suite';      CatKey='^duplicate.coverage$';             RecKey='';                    Color='#3ddad7' }
    [PSCustomObject]@{ Label='Duplicate Review';       Desc='Possible duplicate, needs review'; CatKey='^duplicate.review$';               RecKey='';                    Color='#3ddad7' }
    [PSCustomObject]@{ Label='Overlapping License';    Desc='Same license via multiple paths';  CatKey='overlapping';                      RecKey='';                    Color='#3ddad7' }
    [PSCustomObject]@{ Label='Standalone Licenses';    Desc='Standalone included in suite';     CatKey='standalone';                       RecKey='';                    Color='#3ddad7' }
    [PSCustomObject]@{ Label='Teams Unbundling';       Desc='Suite bundles Teams, no usage';    CatKey='teams.unbundling';                 RecKey='';                    Color='#3ddad7' }
    [PSCustomObject]@{ Label='E5 Voice Review';         Desc='E5 with no calling/conferencing';  CatKey='e5.voice';                         RecKey='';                    Color='#3ddad7' }
    [PSCustomObject]@{ Label='Bundle Opportunity';      Desc='Standalone apps cheaper as suite'; CatKey='bundle.opportunity';               RecKey='';                    Color='#3ddad7' }

    # ── Tier 3: Review categories ────────────────────────────────────────────
    [PSCustomObject]@{ Label='Licensing Compliance';   Desc='Policy/entitlement gap detected';  CatKey='licensing.compliance|compliance.gap'; RecKey='';                  Color='#3ddad7' }
[PSCustomObject]@{ Label='Data Gap';               Desc='Unknown SKU, incomplete analysis'; CatKey='data.gap';                         RecKey='';                    Color='#3ddad7' }
    [PSCustomObject]@{ Label='Mailbox Storage Warning'; Desc='Mailbox near capacity limit';     CatKey='mailbox.storage';                  RecKey='';                    Color='#3ddad7' }
    [PSCustomObject]@{ Label='Unlicensed With Data';   Desc='No license but has mailbox data';  CatKey='unlicensed.with.data';             RecKey='';                    Color='#3ddad7' }

    # ── Add-on & Copilot ─────────────────────────────────────────────────────
    [PSCustomObject]@{ Label='Unused Premium Add-Ons'; Desc='Visio / Project / PBI Pro';        CatKey='add.on|visio|project|pbi';        RecKey='';                    Color='#3ddad7' }
    [PSCustomObject]@{ Label='Copilot Reclaim';        Desc='Zero usage & zero readiness';      CatKey='reclaim';                         RecKey='';                    Color='#3ddad7' }
    [PSCustomObject]@{ Label='Copilot At Risk';        Desc='Zero usage, active in M365';       CatKey='at.risk|copilot.*risk';            RecKey='';                    Color='#3ddad7' }

    # ── Exchange / Mailbox ───────────────────────────────────────────────────
    [PSCustomObject]@{ Label='Exchange Kiosk Downgrade'; Desc='Web-only usage, <2 GB mailbox';  CatKey='exchange.kiosk';                   RecKey='EXCHANGE KIOSK';       Color='#3ddad7' }
    [PSCustomObject]@{ Label='Forwarding Mailbox Review'; Desc='Mailbox forwarding all mail';    CatKey='forwarding.mailbox.review';        RecKey='FORWARDING MAILBOX';   Color='#3ddad7' }
    [PSCustomObject]@{ Label='Expensive Cold Storage'; Desc='E5 retained only for archive/hold'; CatKey='expensive.cold';                  RecKey='EXPENSIVE COLD';       Color='#3ddad7' }

    # ── Activity / Sync ─────────────────────────────────────────────────────
    [PSCustomObject]@{ Label='Background Sync Only';   Desc='Zero interactive activity, OneDrive syncing'; CatKey='background.sync';        RecKey='BACKGROUND SYNC';      Color='#3ddad7' }

    # ── Cleanup ──────────────────────────────────────────────────────────────
    [PSCustomObject]@{ Label='Free License Overlap';    Desc='Self-service trial/free licenses'; CatKey='free.license';                     RecKey='';                    Color='#3ddad7' }
    [PSCustomObject]@{ Label='Windows License Review';  Desc='Windows E3/E5 with no sign-in';    CatKey='windows.license';                  RecKey='';                    Color='#3ddad7' }
)

# ── Dynamic tile generation: catch any category not covered by a well-known tile ─
$_skipCats = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
@('OK','','Unlicensed') | ForEach-Object { [void]$_skipCats.Add($_) }

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
            Desc   = 'Auto-detected category'
            CatKey = '^' + [regex]::Escape($cat) + '$'
            RecKey = ''
            Color  = '#3ddad7'
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
        $tileAmount += Get-EstimatedSavings $m.'Recommendation Category' (Parse-Decimal $m.'Annual License Cost (EUR)') $m.'Recommendation'
    }

    [PSCustomObject]@{
        label   = $def.Label
        desc    = $def.Desc
        color   = $def.Color
        key     = $def.CatKey
        recKey  = $def.RecKey
        users   = $matched.Count
        savings = [math]::Round($tileAmount, 0)
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
            cat = $u.Category
            sav = $u.Savings
            ex  = ($r.'Has Exchange License'   -eq 'True')
            exU = ($r.'Exchange Intensity'     -and $r.'Exchange Intensity'   -notmatch '^(Low|None|)$')
            tm  = ($r.'Has Teams License'      -eq 'True')
            tmU = ($r.'Teams Intensity'        -and $r.'Teams Intensity'      -notmatch '^(Low|None|)$')
            dt  = ($r.'No Desktop Apps'        -ne 'True')
            dtU = ($r.'Uses Desktop Apps'      -eq 'True')
            od  = ($r.'Has OneDrive License'   -eq 'True')
            odU = ($r.'OneDrive Intensity'     -and $r.'OneDrive Intensity'   -notmatch '^(Low|None|)$')
            sp  = ($r.'Has SharePoint License' -eq 'True')
            spU = ($r.'SharePoint Intensity'   -and $r.'SharePoint Intensity' -notmatch '^(Low|None|)$')
            co  = [bool]($r.'License Friendly Names' -match 'Copilot')
            coU = [bool]($r.'Copilot Active Apps'    -and $r.'Copilot Active Apps' -ne '')
        }
    }
})

# ── KPI extraction from executive summary ─────────────────────────────────────
$kpiTotalSpend  = [decimal]0
$kpiSavingsPot  = [decimal]0
$kpiTotalUsers  = $rows.Count
$kpiWithRec     = @($rows | Where-Object { $_.'Recommendation Category' -ne 'OK' -and $_.'Recommendation Category' -ne '' -and $_.'Recommendation Category' -ne 'Unlicensed' }).Count

if ($summaryRows) {
    $ovTotalSpend = $summaryRows | Where-Object { $_.'Category' -match 'Total Annual M365 Spend' }
    if ($ovTotalSpend) { $kpiTotalSpend = Parse-Decimal ($ovTotalSpend | Select-Object -First 1).'Annual Amount (EUR)' }
}
if ($kpiTotalSpend -eq 0) {
    $kpiTotalSpend = ($rows | ForEach-Object { Parse-Decimal $_.'Annual License Cost (EUR)' } | Measure-Object -Sum).Sum
}
# Derive headline savings from tile sums — single source of truth for drill-down consistency
$kpiSavingsPot = [decimal]($tileData | Measure-Object -Property savings -Sum).Sum
$kpiSavingsPct = if ($kpiTotalSpend -gt 0) { [math]::Round($kpiSavingsPot / $kpiTotalSpend * 100, 1) } else { 0 }

# ── JSON helpers ──────────────────────────────────────────────────────────────
function To-JsonString([object]$obj) {
    return (ConvertTo-Json -InputObject $obj -Depth 5 -Compress)
}

# ── Prepare JS data ───────────────────────────────────────────────────────────
$topUsers = @($userData | Sort-Object Savings -Descending |
    Select-Object Name, UPN, Dept, Cost, Savings, Category, Tags, Licenses, Rec)

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

$groupJs = @($groupRows | ForEach-Object {
    [PSCustomObject]@{
        name    = $_.'Group Name'
        type    = $_.'Membership Type'
        members = [int]($_.'Member Count' -replace '\D','')
        skus    = $_.'Assigned Licenses'
        count   = [int]($_.'License Count' -replace '\D','')
    }
} | Sort-Object { $_.members } -Descending)

$jsTopUsers  = To-JsonString $topUsers
$jsSkuData   = To-JsonString $skuJs
$jsTileData  = To-JsonString $tileData
$jsCapUsers  = To-JsonString $capUsers
$jsPoolSkus  = To-JsonString $poolSkuRows
$jsAllCats   = To-JsonString @($userData | ForEach-Object { $_.Category } | Sort-Object -Unique)
$jsGroups    = To-JsonString $groupJs

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
<title>M365 License Optimization Dashboard</title>
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
.disclaimer{font-size:12px;color:var(--text-secondary);text-align:center;margin-top:14px;font-weight:500;position:relative;animation:pulseGlow 5s ease-in-out infinite}
@keyframes pulseGlow{0%,100%{opacity:.3;text-shadow:none}50%{opacity:1;text-shadow:0 0 8px rgba(255,170,0,.35)}}
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
.cat-badge{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:500;background:var(--purple-dim);color:#d0d0e8}
/* SKU bars */
.sku-row{display:flex;align-items:center;gap:12px;margin-bottom:10px}
.sku-name{width:220px;font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex-shrink:0;color:var(--text-secondary)}
.sku-bar-wrap{flex:1;background:rgba(255,255,255,.06);border-radius:4px;height:22px;overflow:hidden;position:relative}
.sku-bar{height:100%;border-radius:4px;display:flex;align-items:center;padding-left:8px;font-size:11px;font-weight:600;color:#fff;transition:width .6s ease}
.sku-amount{width:90px;font-size:12px;font-weight:600;text-align:right;flex-shrink:0;font-family:'JetBrains Mono',monospace}
.sku-users{width:60px;font-size:11px;color:var(--text-dim);text-align:right;flex-shrink:0}
/* Capability matrix */
.cap-table{width:100%;border-collapse:collapse;font-size:12px}
.cap-table th{background:var(--navy-surface);padding:8px 6px;text-align:center;font-size:11px;font-weight:600;color:var(--text-secondary);border:1px solid var(--navy-border)}
.cap-table th.user-col{text-align:left;padding-left:12px;min-width:160px}
.cap-table td{padding:4px 4px;border:1px solid rgba(255,255,255,.04);text-align:center;vertical-align:middle}
.cap-table td.user-name{text-align:left;padding-left:12px}
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
@media print{.tabs{position:static}.panel{display:block!important;page-break-before:always}.panel:first-of-type{page-break-before:auto}.modal-overlay{display:none!important}}
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

<header>
  <h1>M365 License Optimization Dashboard</h1>
  <p>$reportDate</p>
  <div class="kpis">
    <div class="kpi">
      <div class="label">Total Users</div>
      <div class="value">$kpiTotalUsers</div>
      <div class="sub">in scope</div>
    </div>
    <div class="kpi alert">
      <div class="label">With Recommendations</div>
      <div class="value">$kpiWithRec</div>
      <div class="sub">$([math]::Round($kpiWithRec / [math]::Max($kpiTotalUsers,1) * 100, 0))% of users</div>
    </div>
    <div class="kpi">
      <div class="label">Total Annual Spend</div>
      <div class="value">&euro;$([string]::Format('{0:N0}', $kpiTotalSpend))</div>
      <div class="sub">licensed users</div>
    </div>
    <div class="kpi good">
      <div class="label">Potential Annual Savings</div>
      <div class="value">&euro;$([string]::Format('{0:N0}', $kpiSavingsPot))</div>
      <div class="sub">$kpiSavingsPct% of annual spend</div>
    </div>
  </div>
  <div class="disclaimer">$disclaimerText</div>
</header>

<div class="tabs">
  <button class="tab-btn active" onclick="showTab(0)">&#9733; Overview</button>
  <button class="tab-btn"        onclick="showTab(1)">&#128202; Potential Savings by Category</button>
  <button class="tab-btn"        onclick="showTab(2)">&#128176; Potential Savings by User</button>
  <button class="tab-btn"        onclick="showTab(3)">&#128230; Potential Savings by SKU</button>
  <button class="tab-btn"        onclick="showTab(4)">&#128309; Over/Under-Licensed</button>
  <button class="tab-btn"        onclick="showTab(5)">&#128274; License Groups</button>
</div>

<!-- TAB 0: DASHBOARD OVERVIEW -->
<div class="panel active" id="panel-0">
  <div class="dash-section-title">Quick Win Categories <span style="font-size:12px;font-weight:400;color:#6a6a8e;margin-left:8px">Click a tile to drill into affected users</span></div>
  <div id="dash-tiles" class="dash-tiles-grid"></div>
</div>

<!-- TAB 1: SAVINGS BY CATEGORY -->
<div class="panel" id="panel-1">
  <div class="card">
    <h3>Potential Savings by Category</h3>
    <div id="dash-breakdown" style="margin-top:12px"></div>
  </div>
</div>

<!-- TAB 2: SAVINGS BY USER -->
<div class="panel" id="panel-2">
  <div class="card">
    <h3>All Users by Potential Savings <span class="badge-count" id="user-count"></span></h3>
    <div class="filter-row">
      <input type="text" id="user-filter" placeholder="Filter by name / UPN / department&#8230;" oninput="renderUserTable()" style="flex:1;min-width:200px">
      <select id="cat-filter" onchange="renderUserTable()"><option value="">All categories</option></select>
    </div>
    <p style="font-size:11px;color:#6a6a8e;margin-bottom:12px">Click any row to view the full recommendation.</p>
    <div class="tbl-wrap">
      <table id="user-table">
        <thead>
          <tr>
            <th onclick="sortTable('Name')"     data-col="Name">     Name <span class="sort-icon">&#9660;</span></th>
            <th onclick="sortTable('Dept')"     data-col="Dept">     Department <span class="sort-icon">&#9660;</span></th>
            <th onclick="sortTable('Savings')"  data-col="Savings">  Est. Potential Savings/yr <span class="sort-icon">&#9660;</span></th>
            <th onclick="sortTable('Cost')"     data-col="Cost">     Annual Cost <span class="sort-icon">&#9660;</span></th>
            <th onclick="sortTable('Category')" data-col="Category"> Category <span class="sort-icon">&#9660;</span></th>
            <th>Licenses</th>
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
    <h3>Potential License Waste by SKU</h3>
    <p class="section-desc">Total estimated potential savings attributed to each license type. Hover a bar for category breakdown.</p>
    <div id="sku-chart"></div>
  </div>
</div>

<!-- TAB 4: OVER/UNDER-LICENSED (per user) -->
<div class="panel" id="panel-4">
  <div class="card">
    <h3>Provisioned vs Used</h3>
    <p class="section-desc">
      <span style="display:inline-block;width:12px;height:12px;background:rgba(255,107,107,.7);border-radius:3px;vertical-align:middle"></span> Licensed but not active &nbsp;
      <span style="display:inline-block;width:12px;height:12px;background:rgba(116,192,252,.7);border-radius:3px;vertical-align:middle"></span> Active without license &nbsp;
      <span style="display:inline-block;width:12px;height:12px;background:rgba(81,207,102,.7);border-radius:3px;vertical-align:middle"></span> Active &nbsp;
      <span style="display:inline-block;width:12px;height:12px;background:rgba(255,255,255,.06);border-radius:3px;vertical-align:middle;border:1px solid #2a2a55"></span> Not Applicable
    </p>
    <div class="filter-row">
      <input type="text" id="cap-filter" placeholder="Filter by name / UPN&#8230;" oninput="renderCapMatrix()" style="flex:1;min-width:200px">
    </div>
    <div class="tbl-wrap">
      <table class="cap-table" id="cap-table">
        <thead>
          <tr>
            <th class="user-col">User</th>
            <th style="text-align:left">Category</th>
            <th>Est. Potential Savings</th>
            <th>Exchange</th>
            <th>Teams</th>
            <th>Desktop</th>
            <th>OneDrive</th>
            <th>SharePoint</th>
            <th>Copilot</th>
          </tr>
        </thead>
        <tbody id="cap-tbody"></tbody>
      </table>
    </div>
    <p style="font-size:11px;color:#6a6a8e;margin-top:10px">Showing all non-OK users by potential savings.</p>
  </div>
</div>

<!-- TAB 5: LICENSE GROUPS -->
<div class="panel" id="panel-5">
  <div class="card">
    <h3>Entra ID Licensing Groups <span class="badge-count" id="group-count"></span></h3>
    <p class="section-desc">Groups with licenses assigned via Entra ID group-based licensing. Shows membership type (Dynamic rule or manually Assigned) and the SKUs distributed through each group.</p>
    <div class="filter-row">
      <input type="text" id="group-filter" placeholder="Filter by group name or SKU&#8230;" oninput="renderGroupTable()" style="flex:1;min-width:200px">
    </div>
    <div class="tbl-wrap">
      <table>
        <thead>
          <tr>
            <th style="text-align:left">Group Name</th>
            <th>Type</th>
            <th>Members</th>
            <th>SKUs</th>
            <th style="text-align:left">Assigned Licenses</th>
          </tr>
        </thead>
        <tbody id="group-tbody"></tbody>
      </table>
    </div>
    <p style="font-size:11px;color:#6a6a8e;margin-top:10px" id="group-empty"></p>
  </div>
</div>

<script>
const USERS     = $jsTopUsers;
const SKUS      = $jsSkuData;
const TILES     = $jsTileData;
const CAP_USERS = $jsCapUsers;
const POOL_SKUS = $jsPoolSkus;
const GROUPS    = $jsGroups;
const ALL_CATS  = $jsAllCats;
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
function capCellUser(prov, used) {
  if (!prov && !used) return { bg:'rgba(255,255,255,.06)', text:'#6a6a8e', label:'N/A' };
  if (prov && used)   return { bg:'rgba(61,218,215,.25)', text:'#3ddad7', label:'Active' };
  if (prov && !used)  return { bg:'rgba(239,110,167,.25)', text:'#ef6ea7', label:'Unused' };
  return { bg:'rgba(91,137,182,.25)', text:'#5b89b6', label:'No license' };
}
function fmtEur(v) { return '\u20ac' + Number(v).toLocaleString('en-GB', {minimumFractionDigits:0,maximumFractionDigits:0}); }
function escHtml(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
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
function extractAmount(body) {
  // Pull out trailing cost/savings amounts into a styled badge
  const patterns = [
    /\.\s*(Potential savings:\s*\u20AC[\d.,]+\/mo\s*\(\u20AC[\d.,]+\/yr\))\.?$/i,
    /\.\s*(Annual (?:waste|cost|overlap cost|savings):\s*\u20AC[\d.,]+)\.?$/i,
    /\.\s*(Annual cost:\s*\u20AC[\d.,]+)\.?$/i,
    /\.\s*(Saves?\s*\u20AC[\d.,]+\/mo\s*\(\u20AC[\d.,]+\/yr\))\.?$/i
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
  'DORMANT':          { border:'#ef6ea7', bg:'rgba(239,110,167,.12)', text:'#ef6ea7' },
  'DORMANT ADMIN RISK':{ border:'#ef6ea7', bg:'rgba(239,110,167,.12)', text:'#ef6ea7' },
  'DORMANT CLOUD PC': { border:'#ef6ea7', bg:'rgba(239,110,167,.12)', text:'#ef6ea7' },
  'DISABLED ACCOUNT': { border:'#ff9f80', bg:'rgba(255,159,128,.12)', text:'#ff9f80' },
  'DISABLED SHARED MAILBOX':{ border:'#ff9f80', bg:'rgba(255,159,128,.12)', text:'#ff9f80' },
  'SECURITY GAP':     { border:'#ef6ea7', bg:'rgba(239,110,167,.12)', text:'#ef6ea7' },
  'LICENSING CHECK':  { border:'#ff9f80', bg:'rgba(255,159,128,.12)', text:'#ff9f80' },
  'AUTOMATION ACCOUNT':{ border:'#5b89b6', bg:'rgba(91,137,182,.12)', text:'#5b89b6' },
  'ADMIN':            { border:'#5b89b6', bg:'rgba(91,137,182,.12)', text:'#5b89b6' },
  'CLOUD PC REVIEW':  { border:'#ff9f80', bg:'rgba(255,159,128,.12)', text:'#ff9f80' },
  'COPILOT ACTIVE':   { border:'#3ddad7', bg:'rgba(61,218,215,.12)', text:'#3ddad7' },
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
  if (!raw) return '<span style="color:#6a6a8e">No recommendation text available.</span>';
  const parts = raw.split(' | ').filter(p => p.trim());
  if (parts.length === 0) return escHtml(raw);
  const items = parts.map(p => {
    const m = p.match(/^([A-Z][A-Z0-9 /\-]+?)(?:\s*\((?:[^()]*|\([^()]*\))*\))?\s*(?:,\s*)?\u2014\s*(.+)/);
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
  grid.innerHTML = TILES.map((t,i) => {
    const hasData = t.users > 0 || t.savings > 0;
    const borderTop = `border-top-color:${t.color}`;
    return `<div class="dash-tile${hasData ? '' : ' dt-zero'}" style="${borderTop}" onclick="clickTile(${i})">
      <div class="dt-label">${escHtml(t.label)}</div>
      <div class="dt-desc">${escHtml(t.desc)}</div>
      <div class="dt-count" style="color:${hasData ? '#ff9f80' : '#6a6a8e'}">${t.users}</div>
      <div class="dt-savings" style="color:${hasData ? '#ff9f80' : '#6a6a8e'}">${t.savings > 0 ? fmtEur(t.savings)+'/yr' : '\u2014'}</div>
      <div class="dt-bar" style="background:${t.color}"></div>
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
      <div class="bd-track"><div class="bd-fill" style="width:${pct}%;background:${t.color}"></div></div>
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
  mb.onclick = null;
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
  const tableRows = matched.map((u, i) =>
    `<tr style="border-bottom:1px solid rgba(255,255,255,.04);cursor:pointer" onclick="showTileUserDetail(${i})" title="Click for full recommendation">
      <td style="padding:10px 12px"><div style="font-weight:500">${escHtml(u.Name||u.UPN)}</div><div style="font-size:11px;color:#6a6a8e">${escHtml(u.UPN||'')}</div></td>
      <td style="padding:10px 12px">${escHtml(u.Dept||'')}</td>
      <td style="padding:10px 12px"><span class="cat-badge">${escHtml(u.Category||'')}</span>${renderTags(u.Tags)}</td>
      <td style="padding:10px 12px;text-align:right">${fmtEur(u.Cost)}</td>
      <td style="padding:10px 12px;text-align:right;font-weight:600;color:#2f9e44">${fmtEur(u.Savings)}</td>
    </tr>`
  ).join('');
  document.getElementById('modal-content').innerHTML = `
    <h2 style="font-size:17px;color:${t.color};margin-bottom:4px">${escHtml(t.label)}</h2>
    <div style="font-size:12px;color:#6a6a8e;margin-bottom:16px">${escHtml(t.desc)} \u2014 ${t.users} finding${t.users!==1?'s':''} (${matched.length} user${matched.length!==1?'s':''} matched)</div>
    <table style="width:100%;border-collapse:collapse;font-size:13px">
      <thead>
        <tr style="background:#181835;border-bottom:1px solid #2a2a55">
          <th style="text-align:left;padding:10px 12px;font-weight:600;color:#9898b8">User</th>
          <th style="text-align:left;padding:10px 12px;font-weight:600;color:#9898b8">Department</th>
          <th style="text-align:left;padding:10px 12px;font-weight:600;color:#9898b8">Category</th>
          <th style="text-align:right;padding:10px 12px;font-weight:600;color:#9898b8">Annual Cost</th>
          <th style="text-align:right;padding:10px 12px;font-weight:600;color:#9898b8">Est. Potential Savings</th>
        </tr>
      </thead>
      <tbody>${tableRows || '<tr><td colspan="5" style="padding:16px;text-align:center;color:#6a6a8e">No matching users found</td></tr>'}</tbody>
    </table>
    ${totalSav > 0 ? `<div style="margin-top:12px;text-align:right;font-size:13px;font-weight:700;color:#2f9e44">Total potential savings: ${fmtEur(totalSav)}/yr</div>` : ''}
    <div style="margin-top:8px;font-size:11px;color:#6a6a8e">Click any row to view the full recommendation.</div>`;
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

(function() {
  const sel = document.getElementById('cat-filter');
  ALL_CATS.forEach(c => {
    const o = document.createElement('option');
    o.value = c;
    o.textContent = c;
    sel.appendChild(o);
  });
})();

function sortTable(col) {
  if (sortCol === col) sortAsc = !sortAsc; else { sortCol = col; sortAsc = (col !== 'Savings' && col !== 'Cost'); }
  document.querySelectorAll('th[data-col]').forEach(th => th.classList.toggle('sorted', th.dataset.col === col));
  renderUserTable();
}

function renderUserTable() {
  const q   = document.getElementById('user-filter').value.toLowerCase();
  const cat = document.getElementById('cat-filter').value;
  const isFiltered = q || cat;
  let data = USERS.filter(u => {
    // Hide rows with zero cost AND zero savings unless explicitly filtered by category
    if (!isFiltered && u.Cost <= 0 && u.Savings <= 0) return false;
    if (q && !(u.Name||'').toLowerCase().includes(q) && !(u.UPN||'').toLowerCase().includes(q) && !(u.Dept||'').toLowerCase().includes(q)) return false;
    if (cat && u.Category !== cat) return false;
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
  const maxSav = data.length ? data[0].Savings : 1;
  document.getElementById('user-tbody').innerHTML = data.map((u,i) => {
    const bg = savingsColor(u.Savings, maxSav);
    const tc = savingsTextColor(u.Savings, maxSav);
    return `<tr class="clickable-row" onclick="showUserModal(${i})">
      <td><div style="font-weight:500">${escHtml(u.Name||u.UPN)}</div><div style="font-size:11px;color:#6a6a8e">${escHtml(u.UPN||'')}</div></td>
      <td>${escHtml(u.Dept||'')}</td>
      <td><span class="savings-cell" style="background:${bg};color:${tc}">${fmtEur(u.Savings)}</span></td>
      <td>${fmtEur(u.Cost)}</td>
      <td><span class="cat-badge">${escHtml(u.Category||'')}</span>${renderTags(u.Tags)}</td>
      <td style="font-size:11px;color:#9898b8;max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${escHtml(u.Licenses||'')}">${escHtml((u.Licenses||'').replace(/;/g,', '))}</td>
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
    <h2 style="font-size:17px;color:#3ddad7;margin-bottom:4px">${escHtml(u.Name||u.UPN)}</h2>
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
    </div>
    <hr class="modal-divider">
    <div class="modal-field">
      <div class="mf-label">Assigned Licenses</div>
      <div class="mf-value" style="margin-top:4px">${licHtml}</div>
    </div>
    <hr class="modal-divider">
    <div class="modal-field">
      <div class="mf-label">Recommendations</div>
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
  'dormant admin':        '#ef6ea7',
  'dormant cloud':        '#e03131',
  'dormant sign-in':      '#fd7e14',
  'dormant':              '#e03131',
  'disabled':             '#e8590c',
  'no activity':          '#f59f00',
  'zero':                 '#f59f00',
  'never signed':         '#f08c00',
  'inactive hold':        '#e8590c',
  'inactive add-on review': '#fd7e14',
  'inactive add-on':      '#fd7e14',
  'inactive mailbox':     '#fd7e14',
  'inactive':             '#fd7e14',
  'shared mailbox':       '#2f9e44',
  'duplicate coverage':   '#5c7cfa',
  'duplicate review':     '#4263eb',
  'duplicate':            '#5c7cfa',
  'standalone':           '#4dabf7',
  'overlapping':          '#748ffc',
  'teams unbundling':     '#1098ad',
  'e5 voice':             '#7048e8',
  'e5 data':              '#1864ab',
  'a la carte':           '#e8590c',
  'frontline':            '#20c997',
  'data gap':             '#6a6a8e',
  'mailbox storage':      '#f08c00',
  'viral license':        '#e64980',
  'windows license':      '#862e9c',
  'unlicensed with data': '#ef6ea7',
  'compliance':           '#e64980',
  'copilot reclaim':      '#1971c2',
  'copilot at risk':      '#0c8599',
  'copilot':              '#1971c2',
  'reclaim':              '#1971c2',
  'at risk':              '#0c8599',
  'add-on':               '#7048e8',
  'visio':                '#7048e8',
  'project':              '#7048e8',
  'pbi':                  '#7048e8',
  'guest':                '#9c36b5',
  'non-human':            '#862e9c',
  'automation':           '#1098ad',
  'admin review':         '#5b89b6',
  'admin':                '#5b89b6',
  'default':              '#6a6a8e'
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
  const tableRows = matched.map((u, i) =>
    `<tr style="border-bottom:1px solid rgba(255,255,255,.04);cursor:pointer" onclick="showTileUserDetail(${i})" title="Click for full recommendation">
      <td style="padding:10px 12px"><div style="font-weight:500">${escHtml(u.Name||u.UPN)}</div><div style="font-size:11px;color:#6a6a8e">${escHtml(u.UPN||'')}</div></td>
      <td style="padding:10px 12px">${escHtml(u.Dept||'')}</td>
      <td style="padding:10px 12px"><span class="cat-badge">${escHtml(u.Category||'')}</span>${renderTags(u.Tags)}</td>
      <td style="padding:10px 12px;text-align:right">${fmtEur(u.Cost)}</td>
      <td style="padding:10px 12px;text-align:right;font-weight:600;color:#2f9e44">${fmtEur(u.Savings)}</td>
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
      </tr></thead>
      <tbody>${tableRows || '<tr><td colspan="5" style="padding:16px;text-align:center;color:#6a6a8e">No matching users found</td></tr>'}</tbody>
    </table>
    ${totalSav > 0 ? `<div style="margin-top:12px;text-align:right;font-size:13px;font-weight:700;color:#2f9e44">Total potential savings: ${fmtEur(totalSav)}/yr</div>` : ''}`;
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
  const tableRows = matched.map((u, i) =>
    `<tr style="border-bottom:1px solid rgba(255,255,255,.04);cursor:pointer" onclick="showTileUserDetail(${i})" title="Click for full recommendation">
      <td style="padding:10px 12px"><div style="font-weight:500">${escHtml(u.Name||u.UPN)}</div><div style="font-size:11px;color:#6a6a8e">${escHtml(u.UPN||'')}</div></td>
      <td style="padding:10px 12px">${escHtml(u.Dept||'')}</td>
      <td style="padding:10px 12px"><span class="cat-badge">${escHtml(u.Category||'')}</span>${renderTags(u.Tags)}</td>
      <td style="padding:10px 12px;text-align:right">${fmtEur(u.Cost)}</td>
      <td style="padding:10px 12px;text-align:right;font-weight:600;color:#2f9e44">${fmtEur(u.Savings)}</td>
    </tr>`
  ).join('');
  const mc = document.getElementById('modal-content');
  mc.innerHTML = `
    <h3 style="margin-bottom:4px">${escHtml(skuName)}</h3>
    <p style="color:#6a6a8e;margin-bottom:16px">${matched.length} user(s) with recommendations \u2022 Potential savings: ${fmtEur(totalSav)}/yr \u2022 Total waste: ${fmtEur(s.waste)}/yr</p>
    <table style="width:100%;border-collapse:collapse">
      <thead><tr style="background:#181835;font-size:12px;color:#9898b8">
        <th style="text-align:left;padding:8px 12px">User</th>
        <th style="text-align:left;padding:8px 12px">Department</th>
        <th style="text-align:left;padding:8px 12px">Category</th>
        <th style="text-align:right;padding:8px 12px">License Cost</th>
        <th style="text-align:right;padding:8px 12px">Potential Savings</th>
      </tr></thead>
      <tbody>${tableRows}</tbody>
    </table>`;
  // Set up back navigation from user detail
  const mb = document.getElementById('modal-box');
  mb.dataset.backSku = skuIdx;
  mb.style.cursor = 'pointer';
  mb.onclick = function(e) {
    if (e.target.closest('tr') || e.target.closest('a')) return;
    const si = parseInt(mb.dataset.backSku);
    if (!isNaN(si)) { showSkuModal(si); }
  };
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
      const tip = `${c.cat}: ${fmtEur(c.val)} — click to view users`;
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
  {k:'ex',kU:'exU',label:'Exchange'},
  {k:'tm',kU:'tmU',label:'Teams'},
  {k:'dt',kU:'dtU',label:'Desktop'},
  {k:'od',kU:'odU',label:'OneDrive'},
  {k:'sp',kU:'spU',label:'SharePoint'},
  {k:'co',kU:'coU',label:'Copilot'}
];

function renderCapMatrix() {
  const tbody = document.getElementById('cap-tbody');
  const capArr = Array.isArray(CAP_USERS) ? CAP_USERS : (CAP_USERS ? [CAP_USERS] : []);
  if (!capArr.length) {
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:20px;color:#6a6a8e">No data</td></tr>';
    return;
  }
  const q = (document.getElementById('cap-filter').value || '').toLowerCase();
  const data = q ? capArr.filter(u => (u.n||'').toLowerCase().includes(q) || (u.upn||'').toLowerCase().includes(q)) : capArr;
  tbody.innerHTML = data.map(u => {
    const cells = CAP_KEYS.map(ck => {
      const c = capCellUser(u[ck.k], u[ck.kU]);
      const tip = ck.label + ': ' + (u[ck.k] ? 'provisioned' : 'not provisioned') + ' / ' + (u[ck.kU] ? 'in use' : 'not in use');
      return `<td title="${escHtml(tip)}"><span class="cap-cell" style="background:${c.bg};color:${c.text}">${c.label}</span></td>`;
    }).join('');
    return `<tr>
      <td class="user-name"><div style="font-weight:500;white-space:nowrap">${escHtml(u.n||u.upn||'')}</div><div style="font-size:10px;color:#6a6a8e;white-space:nowrap">${escHtml(u.upn||'')}</div></td>
      <td style="text-align:left"><span class="cat-badge" style="white-space:nowrap">${escHtml(u.cat||'')}</span></td>
      <td style="text-align:center;font-weight:600;color:#2f9e44;white-space:nowrap">${fmtEur(u.sav||0)}</td>
      ${cells}
    </tr>`;
  }).join('') || '<tr><td colspan="9" style="text-align:center;padding:20px;color:#6a6a8e">No matching users</td></tr>';
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
    +'<td style="text-align:center">'+g.count+'</td>'
    +'<td style="text-align:left">'+skuBadges(g.skus)+'</td></tr>'
  ).join('');
}

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
