# ========================================================
# M365 License Optimization — HTML Heatmap Dashboard
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
      1. Savings by User     — quick-win tiles + sortable user table with recommendation popup
      2. Savings by SKU      — horizontal bar chart of waste per license type
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

# ── Tier-1 categories (full license cost = reclaimable savings) ──────────────
$tier1 = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
@('Dormant','Disabled Account','E5 Data Hoarder','Inactive Hold','No Activity',
  'Shared Mailbox','Never Signed In','Guest Account Waste','Non-Human Account Waste',
  'Admin Review','Automation Account','Dormant Admin Risk','Legacy Service Account') | ForEach-Object { [void]$tier1.Add($_) }

# ── Helper: parse EUR amount ─────────────────────────────────────────────────
# Handles both European (16560,0 / 1.234,56) and standard (1,234.56) formats
function Parse-Decimal([string]$s) {
    $s = ($s -replace '[\u20AC ]','').Trim()
    if (-not $s) { return [decimal]0 }
    $v = [decimal]0
    # European format: ends with comma + 1-3 digits (e.g. "16560,0" or "1.234,56")
    if ($s -match '^[\d.]+,\d{1,3}$') {
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
    # Pattern 3: Annual overlap cost (Overlapping License)
    foreach ($m in [regex]::Matches($recommendation, 'Annual overlap cost: \u20AC([\d.,]+)')) {
        $total += Parse-Decimal $m.Groups[1].Value
    }
    return $total
}

# ── Build per-user heatmap data ───────────────────────────────────────────────
Write-Host "  Processing users..." -ForegroundColor Gray
$userData = foreach ($r in $rows) {
    $cat  = $r.'Recommendation Category'
    $cost = Parse-Decimal $r.'Annual License Cost (EUR)'
    $rec  = $r.'Recommendation'
    if ($cat -eq 'OK' -or $cat -eq '' -or $cat -eq 'Unlicensed') { continue }
    $savings = Get-EstimatedSavings $cat $cost $rec
    [PSCustomObject]@{
        UPN      = $r.'User Principal Name'
        Name     = $r.'Display Name'
        Dept     = if ($r.'Department') { $r.'Department' } else { '(No Department)' }
        Cost     = [math]::Round($cost, 2)
        Savings  = [math]::Round($savings, 2)
        Category = $cat
        Licenses = $r.'License Friendly Names'
        Rec      = $rec
    }
}
$userData = @($userData)
Write-Host "  $($userData.Count) users with savings opportunities" -ForegroundColor Gray

# ── SKU waste rollup ──────────────────────────────────────────────────────────
$skuRollup = @{}
foreach ($u in $userData) {
    $licenses = @($u.Licenses -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
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
    [PSCustomObject]@{ Label='Dormant Accounts';       Desc='No sign-in >30 days';             CatKey='^dormant$';                        RecKey='';                    Color='#e03131' }
    [PSCustomObject]@{ Label='Disabled Accounts';      Desc='Sign-in blocked';                  CatKey='disabled';                         RecKey='';                    Color='#e8590c' }
    [PSCustomObject]@{ Label='Never Signed In';        Desc='No interactive sign-in on record'; CatKey='never.signed';                     RecKey='NEVER SIGNED IN';     Color='#f08c00' }
    [PSCustomObject]@{ Label='Zero M365 Usage';        Desc='No app activity in period';        CatKey='no.activity|zero.*usage';          RecKey='NO ACTIVITY detected'; Color='#f59f00' }
    [PSCustomObject]@{ Label='Admin Review';           Desc='Admin with productivity license';  CatKey='^admin review$';                   RecKey='';                    Color='#868e96' }
    [PSCustomObject]@{ Label='Shared Mailbox';         Desc='No license needed under 50 GB';    CatKey='shared.mailbox';                   RecKey='';                    Color='#2f9e44' }
    [PSCustomObject]@{ Label='Guest w/ Paid Licenses'; Desc='B2B guest holding a paid license'; CatKey='guest';                            RecKey='';                    Color='#9c36b5' }
    [PSCustomObject]@{ Label='Automation Accounts';    Desc='Service/automation account';       CatKey='^automation.account$';             RecKey='';                    Color='#1098ad' }
    [PSCustomObject]@{ Label='Dormant Admin Accounts'; Desc='Admin with no sign-in detected';   CatKey='dormant.admin';                    RecKey='';                    Color='#c92a2a' }

    # ── Tier 2: License optimization (partial savings) ───────────────────────
    [PSCustomObject]@{ Label='Duplicate Coverage';     Desc='Standalone covered by suite';      CatKey='^duplicate.coverage$';             RecKey='';                    Color='#5c7cfa' }
    [PSCustomObject]@{ Label='Duplicate Review';       Desc='Possible duplicate, needs review'; CatKey='^duplicate.review$';               RecKey='';                    Color='#4263eb' }
    [PSCustomObject]@{ Label='Overlapping License';    Desc='Same license via multiple paths';  CatKey='overlapping';                      RecKey='';                    Color='#748ffc' }
    [PSCustomObject]@{ Label='Standalone Licenses';    Desc='Standalone included in suite';     CatKey='standalone';                       RecKey='';                    Color='#4dabf7' }
    [PSCustomObject]@{ Label='Teams Unbundling';       Desc='Suite bundles Teams, no usage';    CatKey='teams.unbundling';                 RecKey='';                    Color='#1098ad' }
    [PSCustomObject]@{ Label='E5 Voice Waste';         Desc='E5 with no calling/conferencing';  CatKey='e5.voice';                         RecKey='';                    Color='#7048e8' }
    [PSCustomObject]@{ Label='A La Carte Waste';       Desc='Standalone apps cheaper as suite'; CatKey='a.la.carte';                       RecKey='';                    Color='#e8590c' }

    # ── Tier 3: Review categories ────────────────────────────────────────────
    [PSCustomObject]@{ Label='Licensing Compliance';   Desc='Policy/entitlement gap detected';  CatKey='licensing.compliance|compliance.gap'; RecKey='';                  Color='#e64980' }
    [PSCustomObject]@{ Label='Frontline Review';       Desc='Check desktop app dependency';     CatKey='frontline';                        RecKey='';                    Color='#20c997' }
    [PSCustomObject]@{ Label='Data Gap';               Desc='Unknown SKU, incomplete analysis'; CatKey='data.gap';                         RecKey='';                    Color='#adb5bd' }
    [PSCustomObject]@{ Label='Mailbox Storage Warning'; Desc='Mailbox near capacity limit';     CatKey='mailbox.storage';                  RecKey='';                    Color='#f08c00' }
    [PSCustomObject]@{ Label='Unlicensed With Data';   Desc='No license but has mailbox data';  CatKey='unlicensed.with.data';             RecKey='';                    Color='#c92a2a' }

    # ── Add-on & Copilot ─────────────────────────────────────────────────────
    [PSCustomObject]@{ Label='Unused Premium Add-Ons'; Desc='Visio / Project / PBI Pro';        CatKey='add.on|visio|project|pbi';        RecKey='';                    Color='#7048e8' }
    [PSCustomObject]@{ Label='Copilot Reclaim';        Desc='Zero usage & zero readiness';      CatKey='reclaim';                         RecKey='';                    Color='#1971c2' }
    [PSCustomObject]@{ Label='Copilot At Risk';        Desc='Zero usage, active in M365';       CatKey='at.risk|copilot.*risk';            RecKey='';                    Color='#0c8599' }

    # ── Cleanup ──────────────────────────────────────────────────────────────
    [PSCustomObject]@{ Label='Viral License Cleanup';  Desc='Self-service trial/free licenses'; CatKey='viral.license';                    RecKey='';                    Color='#e64980' }
    [PSCustomObject]@{ Label='Windows License Waste';  Desc='Windows E3/E5 with no sign-in';    CatKey='windows.license';                  RecKey='';                    Color='#862e9c' }
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
        $_autoPalette = @('#339af0','#51cf66','#fcc419','#ff8787','#b197fc','#63e6be','#ffa94d','#a9e34b','#e599f7','#74c0fc')
        $tileDefs += [PSCustomObject]@{
            Label  = $cat
            Desc   = 'Auto-detected category'
            CatKey = '^' + [regex]::Escape($cat) + '$'
            RecKey = ''
            Color  = $_autoPalette[$tileDefs.Count % $_autoPalette.Count]
        }
        Write-Host "    + Auto-tile: $cat" -ForegroundColor DarkGray
    }
}

$tileData = @($tileDefs | ForEach-Object {
    $def = $_

    # Per-row matching: single source of truth for tile count, savings, AND drill-down
    # This guarantees what the tile shows = what you see when you click it
    $matched = @($rows | Where-Object {
        $cat = $_.'Recommendation Category'
        if ($_skipCats.Contains($cat)) { return $false }
        if ($def.CatKey -and $cat -match $def.CatKey) { return $true }
        if ($def.RecKey -and $_.'Recommendation' -match $def.RecKey) { return $true }
        return $false
    })

    # Savings: sum each matched user's estimated savings (same value shown in drill-down)
    $tileAmount = [decimal]0
    foreach ($m in $matched) {
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
$tileData = @($tileData) + [PSCustomObject]@{
    label   = 'Unassigned Licenses'
    desc    = 'Pool licenses not assigned to any user'
    color   = '#495057'
    key     = ''
    recKey  = ''
    users   = $unassignedSKUs
    savings = [math]::Round($unassignedWaste, 0)
}

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
$kpiWithRec     = @($rows | Where-Object { $_.'Recommendation Category' -ne 'OK' -and $_.'Recommendation Category' -ne '' }).Count

if ($summaryRows) {
    $ovTotalSpend = $summaryRows | Where-Object { $_.'Category' -match 'Total Annual M365 Spend' }
    $ovSavings    = $summaryRows | Where-Object { $_.'Category' -match 'Estimated Optimization Potential' }
    if ($ovTotalSpend) { $kpiTotalSpend = Parse-Decimal ($ovTotalSpend | Select-Object -First 1).'Annual Amount (EUR)' }
    if ($ovSavings)    { $kpiSavingsPot = Parse-Decimal ($ovSavings    | Select-Object -First 1).'Annual Amount (EUR)' }
}
if ($kpiTotalSpend -eq 0) {
    $kpiTotalSpend = ($rows | ForEach-Object { Parse-Decimal $_.'Annual License Cost (EUR)' } | Measure-Object -Sum).Sum
}
if ($kpiSavingsPot -eq 0) {
    $kpiSavingsPot = ($userData | Measure-Object -Property Savings -Sum).Sum
}
$kpiSavingsPct = if ($kpiTotalSpend -gt 0) { [math]::Round($kpiSavingsPot / $kpiTotalSpend * 100, 1) } else { 0 }

# ── JSON helpers ──────────────────────────────────────────────────────────────
function To-JsonString([object]$obj) {
    return (ConvertTo-Json -InputObject $obj -Depth 5 -Compress)
}

# ── Prepare JS data ───────────────────────────────────────────────────────────
$topUsers = @($userData | Sort-Object Savings -Descending |
    Select-Object Name, UPN, Dept, Cost, Savings, Category, Licenses, Rec)

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

$jsTopUsers  = To-JsonString $topUsers
$jsSkuData   = To-JsonString $skuJs
$jsTileData  = To-JsonString $tileData
$jsCapUsers  = To-JsonString $capUsers
$jsPoolSkus  = To-JsonString $poolSkuRows
$jsAllCats   = To-JsonString @($userData | ForEach-Object { $_.Category } | Sort-Object -Unique)

$reportDate = (Get-Item $ReportCsv).LastWriteTime.ToString("dd MMM yyyy HH:mm")
$genDate    = (Get-Date).ToString("dd MMM yyyy HH:mm")

# ── HTML generation ───────────────────────────────────────────────────────────
Write-Host "  Building HTML..." -ForegroundColor Gray

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>M365 License Optimization — Heatmap Dashboard</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#f0f2f5;color:#1a1a2e;font-size:14px}
header{background:linear-gradient(135deg,#0f3460 0%,#16213e 100%);color:#fff;padding:24px 32px 20px}
header h1{font-size:22px;font-weight:600;letter-spacing:.3px}
header p{margin-top:4px;opacity:.7;font-size:13px}
.kpis{display:flex;gap:16px;margin-top:20px;flex-wrap:wrap}
.kpi{background:rgba(255,255,255,.1);border-radius:10px;padding:14px 20px;min-width:160px;flex:1}
.kpi .label{font-size:11px;opacity:.7;text-transform:uppercase;letter-spacing:.5px}
.kpi .value{font-size:26px;font-weight:700;margin-top:4px}
.kpi .sub{font-size:11px;opacity:.6;margin-top:2px}
.kpi.alert .value{color:#ff6b6b}
.kpi.good .value{color:#51cf66}
.tabs{display:flex;gap:0;padding:0 32px;background:#fff;border-bottom:2px solid #e9ecef;position:sticky;top:0;z-index:10;box-shadow:0 2px 8px rgba(0,0,0,.06)}
.tab-btn{padding:14px 24px;cursor:pointer;font-size:13px;font-weight:500;color:#6c757d;border:none;background:none;border-bottom:3px solid transparent;margin-bottom:-2px;transition:all .2s}
.tab-btn:hover{color:#0f3460}
.tab-btn.active{color:#0f3460;border-bottom-color:#0f3460;font-weight:600}
.panel{display:none;padding:24px 32px}
.panel.active{display:block}
h2{font-size:16px;font-weight:600;color:#0f3460;margin-bottom:4px}
.section-desc{font-size:12px;color:#868e96;margin-bottom:20px}
.card{background:#fff;border-radius:12px;padding:20px;margin-bottom:20px;box-shadow:0 1px 4px rgba(0,0,0,.06)}
.card h3{font-size:14px;font-weight:600;margin-bottom:14px;color:#343a40}
/* Dashboard tiles */
.dash-section-title{font-size:15px;font-weight:600;color:#0f3460;margin-bottom:16px;padding-top:4px}
.dash-tiles-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:14px;margin-bottom:24px}
.dash-tile{background:#fff;border-radius:12px;padding:20px 22px;cursor:pointer;transition:all .18s;border-top:4px solid transparent;box-shadow:0 1px 6px rgba(0,0,0,.07);position:relative;overflow:hidden}
.dash-tile:hover{transform:translateY(-3px);box-shadow:0 6px 20px rgba(0,0,0,.13)}
.dash-tile.dt-active{box-shadow:0 6px 24px rgba(0,0,0,.18)}
.dash-tile.dt-zero{opacity:.45}
.dash-tile .dt-label{font-size:13px;font-weight:600;margin-bottom:3px;color:#212529}
.dash-tile .dt-desc{font-size:11px;color:#868e96;margin-bottom:14px;line-height:1.4}
.dash-tile .dt-count{font-size:32px;font-weight:700;line-height:1}
.dash-tile .dt-savings{font-size:13px;font-weight:600;margin-top:5px}
.dash-tile .dt-bar{height:3px;border-radius:2px;margin-top:14px;opacity:.35}
/* Small tile cards (legacy, kept for compat) */
.tiles-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:12px;margin-bottom:24px}
.tile-card{background:#fff;border-radius:10px;padding:16px 18px;cursor:pointer;transition:all .18s;border-left:4px solid transparent;box-shadow:0 1px 4px rgba(0,0,0,.06);position:relative}
.tile-card:hover{transform:translateY(-2px);box-shadow:0 4px 16px rgba(0,0,0,.12)}
.tile-card.t-active{box-shadow:0 4px 20px rgba(0,0,0,.15)}
.tile-card .t-label{font-size:12px;font-weight:600;margin-bottom:2px;color:#343a40}
.tile-card .t-desc{font-size:11px;color:#868e96;margin-bottom:10px}
.tile-card .t-count{font-size:26px;font-weight:700;line-height:1.1}
.tile-card .t-savings{font-size:12px;font-weight:500;margin-top:3px;opacity:.85}
.tile-card .t-zero{opacity:.4}
/* User table */
.tbl-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:13px}
th{background:#f8f9fa;color:#495057;font-weight:600;padding:10px 12px;text-align:left;border-bottom:2px solid #dee2e6;white-space:nowrap;cursor:pointer;user-select:none}
th:hover{background:#e9ecef}
th .sort-icon{font-size:10px;margin-left:4px;opacity:.4}
th.sorted .sort-icon{opacity:1}
td{padding:9px 12px;border-bottom:1px solid #f1f3f5;vertical-align:middle}
tr.clickable-row{cursor:pointer}
tr.clickable-row:hover td{background:#f0f4ff}
.savings-cell{font-weight:600;border-radius:4px;padding:3px 8px;display:inline-block;font-size:12px}
.cat-badge{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:500;background:#e9ecef;color:#495057}
/* SKU bars */
.sku-row{display:flex;align-items:center;gap:12px;margin-bottom:10px}
.sku-name{width:220px;font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex-shrink:0}
.sku-bar-wrap{flex:1;background:#f1f3f5;border-radius:4px;height:22px;overflow:hidden;position:relative}
.sku-bar{height:100%;border-radius:4px;display:flex;align-items:center;padding-left:8px;font-size:11px;font-weight:600;color:#fff;transition:width .6s ease}
.sku-amount{width:90px;font-size:12px;font-weight:600;text-align:right;flex-shrink:0}
.sku-users{width:60px;font-size:11px;color:#868e96;text-align:right;flex-shrink:0}
/* Capability matrix */
.cap-table{width:100%;border-collapse:collapse;font-size:12px}
.cap-table th{background:#f8f9fa;padding:8px 6px;text-align:center;font-size:11px;font-weight:600;color:#495057;border:1px solid #dee2e6}
.cap-table th.user-col{text-align:left;padding-left:12px;min-width:160px}
.cap-table td{padding:4px 4px;border:1px solid #f1f3f5;text-align:center;vertical-align:middle}
.cap-table td.user-name{text-align:left;padding-left:12px}
.cap-cell{border-radius:4px;padding:3px 4px;font-size:10px;font-weight:600;display:inline-block;min-width:42px}
/* Modal */
.modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:1000;align-items:flex-start;justify-content:center;padding-top:60px}
.modal-overlay.open{display:flex}
.modal-box{background:#fff;border-radius:14px;padding:28px 32px;max-width:880px;width:95%;max-height:80vh;overflow-y:auto;position:relative;box-shadow:0 24px 80px rgba(0,0,0,.3)}
.modal-close{position:absolute;top:14px;right:18px;border:none;background:none;font-size:22px;cursor:pointer;color:#adb5bd;line-height:1;padding:2px 6px;border-radius:4px}
.modal-close:hover{background:#f1f3f5;color:#343a40}
.modal-field{margin-bottom:14px}
.modal-field .mf-label{font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#868e96;margin-bottom:3px;font-weight:600}
.modal-field .mf-value{font-size:13px;color:#212529;line-height:1.5}
.modal-rec{background:#f8f9fa;border-radius:8px;padding:14px;font-size:13px;line-height:1.6;color:#343a40;white-space:pre-wrap;word-break:break-word}
.modal-divider{border:none;border-top:1px solid #e9ecef;margin:16px 0}
/* Utilities */
.text-muted{color:#868e96}
.filter-row{display:flex;gap:12px;margin-bottom:16px;align-items:center;flex-wrap:wrap}
.filter-row input,.filter-row select{padding:7px 12px;border:1px solid #dee2e6;border-radius:6px;font-size:13px;outline:none}
.filter-row input:focus,.filter-row select:focus{border-color:#0f3460}
.badge-count{background:#0f3460;color:#fff;border-radius:10px;padding:1px 8px;font-size:11px;margin-left:6px}
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
  <h1>M365 License Optimization — Heatmap Dashboard</h1>
  <p>Report data: $reportDate &nbsp;|&nbsp; Generated: $genDate</p>
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
      <div class="value">€$([string]::Format('{0:N0}', $kpiTotalSpend))</div>
      <div class="sub">licensed users</div>
    </div>
    <div class="kpi good">
      <div class="label">Savings Potential</div>
      <div class="value">€$([string]::Format('{0:N0}', $kpiSavingsPot))</div>
      <div class="sub">$kpiSavingsPct% of annual spend</div>
    </div>
  </div>
</header>

<div class="tabs">
  <button class="tab-btn active" onclick="showTab(0)">&#9733; Overview</button>
  <button class="tab-btn"        onclick="showTab(1)">&#128176; Savings by User</button>
  <button class="tab-btn"        onclick="showTab(2)">&#128230; Savings by SKU</button>
  <button class="tab-btn"        onclick="showTab(3)">&#128309; Over/Under-Licensed</button>
</div>

<!-- TAB 0: DASHBOARD OVERVIEW -->
<div class="panel active" id="panel-0">
  <div class="dash-section-title">Quick Win Categories <span style="font-size:12px;font-weight:400;color:#868e96;margin-left:8px">Click a tile to drill into affected users</span></div>
  <div id="dash-tiles" class="dash-tiles-grid"></div>
  <div class="card" style="margin-top:4px">
    <h3>Savings Breakdown by Category</h3>
    <div id="dash-breakdown" style="margin-top:12px"></div>
  </div>
</div>

<!-- TAB 1: SAVINGS BY USER -->
<div class="panel" id="panel-1">
  <div class="card">
    <h3>All Users by Savings Potential <span class="badge-count" id="user-count"></span></h3>
    <div class="filter-row">
      <input type="text" id="user-filter" placeholder="Filter by name / UPN / department&#8230;" oninput="renderUserTable()" style="flex:1;min-width:200px">
      <select id="cat-filter" onchange="renderUserTable()"><option value="">All categories</option></select>
    </div>
    <p style="font-size:11px;color:#868e96;margin-bottom:12px">Click any row to view the full recommendation.</p>
    <div class="tbl-wrap">
      <table id="user-table">
        <thead>
          <tr>
            <th onclick="sortTable('Name')"     data-col="Name">     Name <span class="sort-icon">&#9660;</span></th>
            <th onclick="sortTable('Dept')"     data-col="Dept">     Department <span class="sort-icon">&#9660;</span></th>
            <th onclick="sortTable('Savings')"  data-col="Savings">  Est. Savings/yr <span class="sort-icon">&#9660;</span></th>
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

<!-- TAB 2: SAVINGS BY SKU -->
<div class="panel" id="panel-2">
  <div class="card">
    <h3>License Waste by SKU (Top 20)</h3>
    <p class="section-desc">Total estimated savings attributed to each license type. Hover a bar for category breakdown.</p>
    <div id="sku-chart"></div>
  </div>
</div>

<!-- TAB 3: OVER/UNDER-LICENSED (per user) -->
<div class="panel" id="panel-3">
  <div class="card">
    <h3>Provisioned vs Used &#8212; Per User</h3>
    <p class="section-desc">
      <span style="display:inline-block;width:12px;height:12px;background:rgba(255,107,107,.7);border-radius:3px;vertical-align:middle"></span> Over-provisioned (licensed, not using) &nbsp;
      <span style="display:inline-block;width:12px;height:12px;background:rgba(116,192,252,.7);border-radius:3px;vertical-align:middle"></span> Under-licensed (using, no license) &nbsp;
      <span style="display:inline-block;width:12px;height:12px;background:rgba(81,207,102,.7);border-radius:3px;vertical-align:middle"></span> Active &nbsp;
      <span style="display:inline-block;width:12px;height:12px;background:#f1f3f5;border-radius:3px;vertical-align:middle;border:1px solid #dee2e6"></span> N/A
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
            <th>Est. Savings</th>
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
    <p style="font-size:11px;color:#adb5bd;margin-top:10px">Showing all non-OK users by savings potential.</p>
  </div>
</div>

<script>
const USERS     = $jsTopUsers;
const SKUS      = $jsSkuData;
const TILES     = $jsTileData;
const CAP_USERS = $jsCapUsers;
const POOL_SKUS = $jsPoolSkus;
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
  if (max === 0) return '#f1f3f5';
  const t = Math.min(val / max, 1);
  if (t < 0.5) {
    const r = Math.round(198 + (255-198)*t*2), g = Math.round(239 - (239-235)*t*2), b = Math.round(206 + (156-206)*t*2);
    return `rgb(${r},${g},${b})`;
  }
  const t2 = (t-0.5)*2;
  const r = 255, g = Math.round(235 - (235-199)*t2), b = Math.round(156 - (156-100)*t2);
  return `rgb(${r},${g},${b})`;
}
function savingsTextColor(val, max) {
  if (max === 0) return '#495057';
  return val / max > 0.6 ? '#fff' : '#212529';
}
function capCellUser(prov, used) {
  if (!prov && !used) return { bg:'#f1f3f5', text:'#adb5bd', label:'—' };
  if (prov && used)   return { bg:'rgba(81,207,102,.65)', text:'#2f9e44', label:'Active' };
  if (prov && !used)  return { bg:'rgba(255,107,107,.65)', text:'#c92a2a', label:'Over' };
  return { bg:'rgba(116,192,252,.65)', text:'#1864ab', label:'Under' };
}
function fmtEur(v) { return '€' + Number(v).toLocaleString('en-GB', {minimumFractionDigits:0,maximumFractionDigits:0}); }
function escHtml(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
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
      <div class="dt-count" style="color:${hasData ? t.color : '#adb5bd'}">${t.users}</div>
      <div class="dt-savings" style="color:${hasData ? t.color : '#adb5bd'}">${t.savings > 0 ? fmtEur(t.savings)+'/yr' : '\u2014'}</div>
      <div class="dt-bar" style="background:${t.color}"></div>
    </div>`;
  }).join('');

  const breakdown = document.getElementById('dash-breakdown');
  const tilesWithSavings = TILES.filter(t => t.savings > 0);
  if (!tilesWithSavings.length) { breakdown.innerHTML = '<p class="text-muted">No savings data available.</p>'; return; }
  const totalSav = tilesWithSavings.reduce((s,t) => s + t.savings, 0);
  const segs = tilesWithSavings.map(t => {
    const pct = (t.savings / totalSav * 100).toFixed(1);
    return `<div title="${escHtml(t.label)}: ${fmtEur(t.savings)} (${pct}%)" style="width:${pct}%;background:${t.color};height:100%;display:inline-block;vertical-align:top;cursor:pointer" onclick="clickTile(${TILES.indexOf(t)})"></div>`;
  }).join('');
  const legend = tilesWithSavings.map(t => {
    const pct = (t.savings / totalSav * 100).toFixed(0);
    return `<span style="display:inline-flex;align-items:center;gap:5px;font-size:11px;color:#495057;white-space:nowrap">
      <span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:${t.color}"></span>${escHtml(t.label)} ${pct}%</span>`;
  }).join('');
  breakdown.innerHTML = `
    <div style="height:28px;border-radius:6px;overflow:hidden;display:flex;margin-bottom:14px">${segs}</div>
    <div style="display:flex;flex-wrap:wrap;gap:10px">${legend}</div>`;
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
  mb.style.cursor = ''; delete mb.dataset.backTile;
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
    `<tr style="border-bottom:1px solid #f1f3f5;cursor:pointer" onclick="showTileUserDetail(${i})" title="Click for full recommendation">
      <td style="padding:10px 12px"><div style="font-weight:500">${escHtml(u.Name||u.UPN)}</div><div style="font-size:11px;color:#868e96">${escHtml(u.UPN||'')}</div></td>
      <td style="padding:10px 12px">${escHtml(u.Dept||'')}</td>
      <td style="padding:10px 12px"><span class="cat-badge">${escHtml(u.Category||'')}</span></td>
      <td style="padding:10px 12px;text-align:right">${fmtEur(u.Cost)}</td>
      <td style="padding:10px 12px;text-align:right;font-weight:600;color:#2f9e44">${fmtEur(u.Savings)}</td>
    </tr>`
  ).join('');
  document.getElementById('modal-content').innerHTML = `
    <h2 style="font-size:17px;color:${t.color};margin-bottom:4px">${escHtml(t.label)}</h2>
    <div style="font-size:12px;color:#868e96;margin-bottom:16px">${escHtml(t.desc)} \u2014 ${t.users} finding${t.users!==1?'s':''} (${matched.length} user${matched.length!==1?'s':''} matched)</div>
    <table style="width:100%;border-collapse:collapse;font-size:13px">
      <thead>
        <tr style="background:#f8f9fa;border-bottom:2px solid #e9ecef">
          <th style="text-align:left;padding:10px 12px;font-weight:600;color:#495057">User</th>
          <th style="text-align:left;padding:10px 12px;font-weight:600;color:#495057">Department</th>
          <th style="text-align:left;padding:10px 12px;font-weight:600;color:#495057">Category</th>
          <th style="text-align:right;padding:10px 12px;font-weight:600;color:#495057">Annual Cost</th>
          <th style="text-align:right;padding:10px 12px;font-weight:600;color:#495057">Est. Savings</th>
        </tr>
      </thead>
      <tbody>${tableRows || '<tr><td colspan="5" style="padding:16px;text-align:center;color:#868e96">No matching users found</td></tr>'}</tbody>
    </table>
    ${totalSav > 0 ? `<div style="margin-top:12px;text-align:right;font-size:13px;font-weight:700;color:#2f9e44">Total savings: ${fmtEur(totalSav)}/yr</div>` : ''}
    <div style="margin-top:8px;font-size:11px;color:#868e96">Click any row to view the full recommendation.</div>`;
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
    `<tr style="border-bottom:1px solid #f1f3f5">
      <td style="padding:10px 12px;font-weight:500">${escHtml(s.sku)}</td>
      <td style="padding:10px 12px;text-align:right;color:#495057">${s.unassigned.toLocaleString()}</td>
      <td style="padding:10px 12px;text-align:right;font-weight:600;color:#c92a2a">${fmtEur(s.waste)}/yr</td>
    </tr>`
  ).join('');
  document.getElementById('modal-content').innerHTML = `
    <h2 style="font-size:17px;color:#495057;margin-bottom:4px">Unassigned Licenses</h2>
    <div style="font-size:12px;color:#868e96;margin-bottom:16px">Paid licenses in the tenant pool with unassigned seats generating waste</div>
    <table style="width:100%;border-collapse:collapse;font-size:13px">
      <thead>
        <tr style="background:#f8f9fa;border-bottom:2px solid #e9ecef">
          <th style="text-align:left;padding:10px 12px;font-weight:600;color:#495057">License SKU</th>
          <th style="text-align:right;padding:10px 12px;font-weight:600;color:#495057">Unassigned Seats</th>
          <th style="text-align:right;padding:10px 12px;font-weight:600;color:#495057">Annual Waste</th>
        </tr>
      </thead>
      <tbody>${tableRows || '<tr><td colspan="3" style="padding:16px;text-align:center;color:#868e96">No unassigned license waste found</td></tr>'}</tbody>
    </table>
    ${rows.length > 1 ? `<div style="margin-top:12px;text-align:right;font-size:13px;font-weight:700;color:#c92a2a">Total: ${fmtEur(total)}/yr</div>` : ''}`;
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
      <td><div style="font-weight:500">${escHtml(u.Name||u.UPN)}</div><div style="font-size:11px;color:#868e96">${escHtml(u.UPN||'')}</div></td>
      <td>${escHtml(u.Dept||'')}</td>
      <td><span class="savings-cell" style="background:${bg};color:${tc}">${fmtEur(u.Savings)}</span></td>
      <td>${fmtEur(u.Cost)}</td>
      <td><span class="cat-badge">${escHtml(u.Category||'')}</span></td>
      <td style="font-size:11px;color:#495057;max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${escHtml(u.Licenses||'')}">${escHtml((u.Licenses||'').replace(/;/g,', '))}</td>
    </tr>`;
  }).join('') || '<tr><td colspan="6" style="text-align:center;padding:20px;color:#868e96">No matching users</td></tr>';
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
    <h2 style="font-size:17px;color:#0f3460;margin-bottom:4px">${escHtml(u.Name||u.UPN)}</h2>
    <div style="font-size:12px;color:#868e96;margin-bottom:16px">${escHtml(u.UPN||'')}</div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:16px">
      <div class="modal-field">
        <div class="mf-label">Department</div>
        <div class="mf-value">${escHtml(u.Dept||'—')}</div>
      </div>
      <div class="modal-field">
        <div class="mf-label">Category</div>
        <div class="mf-value"><span class="cat-badge">${escHtml(u.Category||'')}</span></div>
      </div>
      <div class="modal-field">
        <div class="mf-label">Annual License Cost</div>
        <div class="mf-value" style="font-weight:600">${fmtEur(u.Cost)}</div>
      </div>
      <div class="modal-field">
        <div class="mf-label">Estimated Savings/yr</div>
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
      <div class="mf-label">Recommendation</div>
      <div class="modal-rec" style="margin-top:6px">${escHtml(u.Rec||'No recommendation text available.')}</div>
    </div>`;
  if (activeTileIdx >= 0) {
    setTimeout(function() { mb.style.cursor = 'pointer'; mb.dataset.backTile = activeTileIdx; }, 0);
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
  const bt = this.dataset.backTile;
  if (bt != null) { e.stopPropagation(); showTileModal(parseInt(bt)); }
});

// ── TAB 2: SKU chart ──────────────────────────────────────────────────────────
const CAT_COLORS = {
  'dormant':              '#e03131',
  'disabled':             '#e8590c',
  'no activity':          '#f59f00',
  'zero':                 '#f59f00',
  'never signed':         '#f08c00',
  'inactive':             '#fd7e14',
  'shared mailbox':       '#2f9e44',
  'duplicate coverage':   '#5c7cfa',
  'duplicate review':     '#4263eb',
  'duplicate':            '#5c7cfa',
  'standalone':           '#4dabf7',
  'overlapping':          '#748ffc',
  'teams unbundling':     '#1098ad',
  'e5 voice':             '#7048e8',
  'a la carte':           '#e8590c',
  'frontline':            '#20c997',
  'data gap':             '#adb5bd',
  'mailbox storage':      '#f08c00',
  'viral license':        '#e64980',
  'windows license':      '#862e9c',
  'unlicensed with data': '#c92a2a',
  'compliance':           '#e64980',
  'copilot reclaim':      '#1971c2',
  'reclaim':              '#1971c2',
  'copilot at risk':      '#0c8599',
  'at risk':              '#0c8599',
  'add-on':               '#7048e8',
  'visio':                '#7048e8',
  'project':              '#7048e8',
  'pbi':                  '#7048e8',
  'guest':                '#9c36b5',
  'non-human':            '#862e9c',
  'automation':           '#1098ad',
  'dormant admin':        '#c92a2a',
  'e5 data':              '#1864ab',
  'admin review':         '#868e96',
  'default':              '#adb5bd'
};

function catColor(catName) {
  const lower = (catName||'').toLowerCase();
  for (const key of Object.keys(CAT_COLORS)) {
    if (lower.includes(key)) return CAT_COLORS[key];
  }
  return CAT_COLORS['default'];
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
      return `<span style="display:inline-flex;align-items:center;gap:5px;font-size:11px;color:#495057"><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:${color}"></span>${escHtml(c)}</span>`;
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
      const tip = `${c.cat}: ${fmtEur(c.val)}`;
      return `<div title="${escHtml(tip)}" style="width:${segPct}%;background:${color};height:100%;display:inline-block;vertical-align:top"></div>`;
    }).join('');
    const tipTxt = `${s.lic}: ${fmtEur(s.waste)}\n` + allSegs.map(c => `${c.cat}: ${fmtEur(c.val)}`).join('\n');
    return `<div class="sku-row" title="${escHtml(tipTxt)}">
      <div class="sku-name" title="${escHtml(s.lic)}">${escHtml(s.lic)}</div>
      <div class="sku-bar-wrap" style="position:relative">
        <div style="width:${barPct.toFixed(1)}%;height:100%;display:flex;overflow:hidden;border-radius:4px">${segments}</div>
      </div>
      <div class="sku-amount">${fmtEur(s.waste)}</div>
      <div class="sku-users">${s.users}u</div>
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
    tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:20px;color:#868e96">No data</td></tr>';
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
      <td class="user-name"><div style="font-weight:500;white-space:nowrap">${escHtml(u.n||u.upn||'')}</div><div style="font-size:10px;color:#868e96;white-space:nowrap">${escHtml(u.upn||'')}</div></td>
      <td style="text-align:left"><span class="cat-badge" style="white-space:nowrap">${escHtml(u.cat||'')}</span></td>
      <td style="text-align:center;font-weight:600;color:#2f9e44;white-space:nowrap">${fmtEur(u.sav||0)}</td>
      ${cells}
    </tr>`;
  }).join('') || '<tr><td colspan="9" style="text-align:center;padding:20px;color:#868e96">No matching users</td></tr>';
}

// ── Init ──────────────────────────────────────────────────────────────────────
renderDashboard();
renderUserTable();
renderSkuChart();
renderCapMatrix();
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
