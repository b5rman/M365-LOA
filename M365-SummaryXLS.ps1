<#
.SYNOPSIS
    Post-processes the LOA Excel workbook into a customer-facing assessment report.

.DESCRIPTION
    Creates a copy of M365_LicenseOptimization_*.xlsx with customer-facing language:
    - Renames "Recommendations" tab to "Assessments"
    - Renames column headers (Recommendation → Assessment)
    - Renames category values (Dormant Cloud PC → Cloud PC Review, OK → No Findings)
    - Removes "Copilot Active" and zero-cost rows from the Assessments summary tab
    - Rebuilds category tabs by primary Assessment Category (consistent with HTML)
    - Fixes conditional formatting (OK → No Findings)
    - Removes Executive Summary tab (HTML/Word serve as exec summary)

    Zero changes to source LOA, heatmap, or Word report scripts.

.PARAMETER OutputFolder
    Path containing the LOA output files. The latest XLSX by LastWriteTime is used.

.EXAMPLE
    pwsh -File M365-SummaryXLS.ps1 -OutputFolder ./output
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$OutputFolder
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Locate latest source XLSX ────────────────────────────────────────────────
$sourceFiles = @(Get-ChildItem -Path $OutputFolder -Filter 'M365_LicenseOptimization_*.xlsx' |
    Sort-Object LastWriteTime -Descending)

if ($sourceFiles.Count -eq 0) {
    Write-Error "No M365_LicenseOptimization_*.xlsx found in $OutputFolder"
    return
}
$sourceXlsx = $sourceFiles[0].FullName
Write-Host "[INFO] Source: $($sourceFiles[0].Name)" -ForegroundColor Cyan

# ── Create customer-facing copy ──────────────────────────────────────────────
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$customerXlsx = Join-Path $OutputFolder "M365_LicenseAssessment_$ts.xlsx"
Copy-Item -Path $sourceXlsx -Destination $customerXlsx -Force
Write-Host "[INFO] Created customer copy: M365_LicenseAssessment_$ts.xlsx" -ForegroundColor Cyan

# ── Configuration ────────────────────────────────────────────────────────────
$headerRenames = @{
    'Recommendation'            = 'Assessment'
    'Recommendation Category'   = 'Assessment Category'
    'Recommendation Confidence' = 'Assessment Confidence'
}

$categoryRenames = @{
    'Dormant Cloud PC' = 'Cloud PC Review'
    'OK'               = 'No Findings'
}

$removeCategories = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase)
$removeCategories.Add('Copilot Active') | Out-Null
$removeCategories.Add('No Findings') | Out-Null

$tabCategoryMap = [ordered]@{
    'Admin Review'       = @('Admin Review', 'Dormant Admin Review')
    'Dormant'            = @('Dormant', 'Stale Sign-In')
    'Disabled Account'   = @('Disabled Account', 'Disabled Shared Mailbox')
    'Shared Mailbox'     = @('Shared Mailbox')
    'Automation Account' = @('Automation Account')
    'Never Signed In'    = @('Never Signed In')
}

$tabColors = @{
    'Admin Review'       = [System.Drawing.Color]::FromArgb(180, 198, 231)
    'Dormant'            = [System.Drawing.Color]::FromArgb(255, 199, 206)
    'Disabled Account'   = [System.Drawing.Color]::FromArgb(255, 199, 206)
    'Shared Mailbox'     = [System.Drawing.Color]::FromArgb(198, 239, 206)
    'Automation Account' = [System.Drawing.Color]::FromArgb(217, 217, 217)
    'Never Signed In'    = [System.Drawing.Color]::FromArgb(255, 199, 206)
}

# Category cell colors (matching HTML dashboard CAT_COLORS palette, lightened for Excel)
$catCellColors = @{
    'No Findings'              = [System.Drawing.Color]::FromArgb(220, 245, 220)  # light green
    # teal #3ddad7 — shared mailbox, frontline
    'Shared Mailbox'           = [System.Drawing.Color]::FromArgb(200, 245, 243)
    'Frontline Candidate'      = [System.Drawing.Color]::FromArgb(200, 245, 243)
    'Frontline Rescue'         = [System.Drawing.Color]::FromArgb(200, 245, 243)
    'Frontline Blocked'        = [System.Drawing.Color]::FromArgb(200, 245, 243)
    'Frontline Add-On Stacking' = [System.Drawing.Color]::FromArgb(200, 245, 243)
    # blue #5b8def — inactive products, duplicate, copilot
    'Inactive Add-On'          = [System.Drawing.Color]::FromArgb(214, 224, 247)
    'Duplicate Coverage'       = [System.Drawing.Color]::FromArgb(214, 224, 247)
    'Duplicate Review'         = [System.Drawing.Color]::FromArgb(214, 224, 247)
    'Copilot Watchlist'        = [System.Drawing.Color]::FromArgb(214, 224, 247)
    'AI Overlap Review'        = [System.Drawing.Color]::FromArgb(214, 224, 247)
    'AI Add-On Overlap'        = [System.Drawing.Color]::FromArgb(214, 224, 247)
    'Bundle Consolidation'     = [System.Drawing.Color]::FromArgb(214, 224, 247)
    'Bundle Opportunity'       = [System.Drawing.Color]::FromArgb(214, 224, 247)
    'Web-Only Mailbox'         = [System.Drawing.Color]::FromArgb(214, 224, 247)
    'E5 Upgrade'               = [System.Drawing.Color]::FromArgb(214, 224, 247)
    # peach #ff9f80 — stale sign-in
    'Stale Sign-In'            = [System.Drawing.Color]::FromArgb(255, 230, 220)
    # warm peach #ff8a65 — disabled, inactive hold
    'Disabled Account'         = [System.Drawing.Color]::FromArgb(255, 224, 210)
    'Disabled Shared Mailbox'  = [System.Drawing.Color]::FromArgb(255, 224, 210)
    'Inactive Hold With License' = [System.Drawing.Color]::FromArgb(255, 224, 210)
    'Inactive Hold'            = [System.Drawing.Color]::FromArgb(255, 224, 210)
    'Inactive Mailbox'         = [System.Drawing.Color]::FromArgb(255, 224, 210)
    # soft purple #8b7ed8 — add-on, visio, project, PBI, windows
    'Premium Add-On Review'    = [System.Drawing.Color]::FromArgb(225, 220, 240)
    'Duplicate Assignment'     = [System.Drawing.Color]::FromArgb(225, 220, 240)
    'Free License Overlap'     = [System.Drawing.Color]::FromArgb(225, 220, 240)
    'Suite Inversion'          = [System.Drawing.Color]::FromArgb(225, 220, 240)
    'Windows License Review'   = [System.Drawing.Color]::FromArgb(225, 220, 240)
    'Seeded Visio Overlap'     = [System.Drawing.Color]::FromArgb(225, 220, 240)
    'PBI PPU Overlap'          = [System.Drawing.Color]::FromArgb(225, 220, 240)
    # navy #5b89b6 — admin, copilot reclaim
    'Admin Review'             = [System.Drawing.Color]::FromArgb(210, 225, 240)
    'Dormant Admin Review'     = [System.Drawing.Color]::FromArgb(210, 225, 240)
    'Copilot Reclaim'          = [System.Drawing.Color]::FromArgb(210, 225, 240)
    # deep teal #2ec4b6 — teams unbundling, automation
    'Inactive Teams Entitlement' = [System.Drawing.Color]::FromArgb(200, 240, 235)
    'Automation Account'       = [System.Drawing.Color]::FromArgb(200, 240, 235)
    'Legacy Service Account'   = [System.Drawing.Color]::FromArgb(200, 240, 235)
    'Inactive Audio Conferencing' = [System.Drawing.Color]::FromArgb(200, 240, 235)
    'Forwarding Mailbox Review' = [System.Drawing.Color]::FromArgb(200, 240, 235)
    # pink #ef6ea7 — dormant, no activity, never signed in
    'Dormant'                  = [System.Drawing.Color]::FromArgb(248, 215, 230)
    'Never Signed In'          = [System.Drawing.Color]::FromArgb(248, 215, 230)
    'No Activity'              = [System.Drawing.Color]::FromArgb(248, 215, 230)
    'Unlicensed With Data'     = [System.Drawing.Color]::FromArgb(248, 215, 230)
    # deep pink #d94070 — cloud PC, compliance, trial, storage warnings
    'Cloud PC Review'          = [System.Drawing.Color]::FromArgb(245, 205, 220)
    'Licensing Compliance Gap' = [System.Drawing.Color]::FromArgb(245, 205, 220)
    'Mailbox Storage Warning'  = [System.Drawing.Color]::FromArgb(245, 205, 220)
    'OneDrive Storage Warning' = [System.Drawing.Color]::FromArgb(245, 205, 220)
    'License Error'            = [System.Drawing.Color]::FromArgb(245, 205, 220)
    'Trial License'            = [System.Drawing.Color]::FromArgb(245, 205, 220)
    'Expensive Cold Storage'   = [System.Drawing.Color]::FromArgb(245, 205, 220)
    'Background Sync Only'     = [System.Drawing.Color]::FromArgb(245, 205, 220)
    # muted #6a6a8e — data gap, default
    'Unlicensed'               = [System.Drawing.Color]::FromArgb(230, 230, 235)
    'Room/Equipment'           = [System.Drawing.Color]::FromArgb(230, 230, 235)
    'Guest User'               = [System.Drawing.Color]::FromArgb(230, 230, 235)
    'Non-Human Account Review' = [System.Drawing.Color]::FromArgb(230, 230, 235)
}

$changesApplied = 0

# ══════════════════════════════════════════════════════════════════════════════
# PASS 1: EPPlus in-memory changes (renames, category values, chart, cleanup)
# ══════════════════════════════════════════════════════════════════════════════
$pkg = Open-ExcelPackage -Path $customerXlsx

# ── Rename headers + category values across all tabs ─────────────────────────
foreach ($ws in $pkg.Workbook.Worksheets) {
    if (-not $ws.Dimension) { continue }
    $lastRow = $ws.Dimension.End.Row
    $lastCol = $ws.Dimension.End.Column

    $catCol = -1; $confCol = -1
    for ($c = 1; $c -le $lastCol; $c++) {
        $headerVal = $ws.Cells[1, $c].Text
        if ($headerRenames.ContainsKey($headerVal)) {
            $ws.Cells[1, $c].Value = $headerRenames[$headerVal]
            $changesApplied++
        }
        if ($headerVal -eq 'Recommendation Category' -or $headerVal -eq 'Assessment Category') {
            $catCol = $c
        }
        if ($headerVal -eq 'Recommendation Confidence' -or $headerVal -eq 'Assessment Confidence') {
            $confCol = $c
        }
    }

    if ($catCol -gt 0 -and $lastRow -ge 2) {
        for ($r = 2; $r -le $lastRow; $r++) {
            $cellVal = $ws.Cells[$r, $catCol].Text
            if ($categoryRenames.ContainsKey($cellVal)) {
                $ws.Cells[$r, $catCol].Value = $categoryRenames[$cellVal]
                $changesApplied++
            }
            # Fill empty confidence with "High" (No Findings / Unlicensed are factual)
            if ($confCol -gt 0) {
                $confVal = $ws.Cells[$r, $confCol].Text
                if (-not $confVal -or $confVal -eq '') {
                    $ws.Cells[$r, $confCol].Value = 'High'
                    $changesApplied++
                }
            }

            # Apply category cell color
            $catForColor = $ws.Cells[$r, $catCol].Text
            if ($catCellColors.ContainsKey($catForColor)) {
                $ws.Cells[$r, $catCol].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $ws.Cells[$r, $catCol].Style.Fill.BackgroundColor.SetColor($catCellColors[$catForColor])
            }
        }
    }
}

# ── Collect tab rebuild data from User Report BEFORE deleting tabs ───────────
$tabRebuildData = @{}
$userReport = $pkg.Workbook.Worksheets['User Report']
if ($userReport -and $userReport.Dimension) {
    $urLastRow = $userReport.Dimension.End.Row
    $urLastCol = $userReport.Dimension.End.Column
    $urCatCol = -1
    for ($c = 1; $c -le $urLastCol; $c++) {
        if ($userReport.Cells[1, $c].Text -eq 'Assessment Category') { $urCatCol = $c; break }
    }

    if ($urCatCol -gt 0) {
        foreach ($tabName in $tabCategoryMap.Keys) {
            $existingTab = $pkg.Workbook.Worksheets[$tabName]
            if (-not $existingTab) { continue }

            $matchCats = [System.Collections.Generic.HashSet[string]]::new(
                [string[]]$tabCategoryMap[$tabName], [System.StringComparer]::OrdinalIgnoreCase)

            $matchCount = 0
            for ($r = 2; $r -le $urLastRow; $r++) {
                if ($matchCats.Contains($userReport.Cells[$r, $urCatCol].Text)) { $matchCount++ }
            }

            $oldCount = if ($existingTab.Dimension) { $existingTab.Dimension.End.Row - 1 } else { 0 }
            if ($matchCount -ne $oldCount) {
                $tabRebuildData[$tabName] = @{ OldCount = $oldCount; NewCount = $matchCount }
                # Delete the old tab (removes stale table definitions)
                $pkg.Workbook.Worksheets.Delete($existingTab)
            } else {
                Write-Host "[INFO] Tab '$tabName': $oldCount rows — already matches" -ForegroundColor DarkGray
            }
        }
    }
}

# ── Clean Assessments summary tab (remove zero-cost + excluded categories) ───
$recTab = $pkg.Workbook.Worksheets['Recommendations']
if (-not $recTab) { $recTab = $pkg.Workbook.Worksheets['Assessments'] }

if ($recTab -and $recTab.Dimension) {
    $lastRow = $recTab.Dimension.End.Row
    $lastCol = $recTab.Dimension.End.Column

    $pivotCatCol = 1; $pivotCostCol = -1
    for ($c = 1; $c -le $lastCol; $c++) {
        $h = $recTab.Cells[1, $c].Text
        if ($h -match 'Category') { $pivotCatCol = $c }
        if ($h -match 'Cost') { $pivotCostCol = $c }
    }

    for ($r = $lastRow; $r -ge 2; $r--) {
        $cellVal = $recTab.Cells[$r, $pivotCatCol].Text
        $costVal = if ($pivotCostCol -gt 0) { $recTab.Cells[$r, $pivotCostCol].Value } else { $null }
        $isZeroCost = ($null -ne $costVal -and [decimal]$costVal -eq 0)
        if ($removeCategories.Contains($cellVal) -or $isZeroCost) {
            $reason = if ($removeCategories.Contains($cellVal)) { $cellVal } else { "$cellVal (zero cost)" }
            $recTab.DeleteRow($r)
            $changesApplied++
            Write-Host "[INFO] Removed '$reason' row from summary tab" -ForegroundColor Yellow
        }
    }

    # ── Merge duplicate category rows (caused by category renames, e.g. Dormant Cloud PC → Cloud PC Review) ──
    $mergeLastRow = $recTab.Dimension.End.Row
    # Find all numeric columns (Users count, Annual Amount) for summing
    $pivotUserCol = -1
    for ($c = 1; $c -le $lastCol; $c++) {
        $h = $recTab.Cells[1, $c].Text
        if ($h -match 'Users|Count') { $pivotUserCol = $c }
    }
    # Reverse scan: if category matches a later row, merge into the first occurrence and delete the duplicate
    $seenCats = @{}
    for ($r = 2; $r -le $mergeLastRow; $r++) {
        $cat = $recTab.Cells[$r, $pivotCatCol].Text
        if ($seenCats.ContainsKey($cat)) {
            $keepRow = $seenCats[$cat]
            # Sum numeric columns into the kept row
            if ($pivotUserCol -gt 0) {
                $existingUsers = $recTab.Cells[$keepRow, $pivotUserCol].Value
                $dupeUsers     = $recTab.Cells[$r, $pivotUserCol].Value
                if ($null -ne $existingUsers -and $null -ne $dupeUsers) {
                    $recTab.Cells[$keepRow, $pivotUserCol].Value = [int]$existingUsers + [int]$dupeUsers
                }
            }
            if ($pivotCostCol -gt 0) {
                $existingCost = $recTab.Cells[$keepRow, $pivotCostCol].Value
                $dupeCost     = $recTab.Cells[$r, $pivotCostCol].Value
                if ($null -ne $existingCost -and $null -ne $dupeCost) {
                    $recTab.Cells[$keepRow, $pivotCostCol].Value = [decimal]$existingCost + [decimal]$dupeCost
                }
            }
            Write-Host "[INFO] Merged duplicate '$cat' summary row (row $r into row $keepRow)" -ForegroundColor Yellow
            $recTab.DeleteRow($r)
            $r--; $mergeLastRow--
            $changesApplied++
        } else {
            $seenCats[$cat] = $r
        }
    }

    # Rename tab before chart rebuild (chart refs use current tab name)
    if ($recTab.Name -eq 'Recommendations') {
        $recTab.Name = 'Assessments'
        $changesApplied++
        Write-Host "[INFO] Renamed tab: Recommendations → Assessments" -ForegroundColor Green
    }

    # Delete old chart and rebuild
    $chartNames = @($recTab.Drawings | ForEach-Object { $_.Name })
    foreach ($name in $chartNames) {
        $recTab.Drawings.Remove($name)
        Write-Host "[INFO] Removed old chart: $name" -ForegroundColor Yellow
    }

    [int]$newLastRow = $recTab.Dimension.End.Row
    [int]$chartRows = $newLastRow - 1
    if ($chartRows -gt 0) {
        $chart = $recTab.Drawings.AddChart("AssessmentPieChart",
            [OfficeOpenXml.Drawing.Chart.eChartType]::Pie3D)
        $chart.Title.Text = "Assessment Distribution (by Cost)"
        $chart.SetPosition(1, 0, 4, 0)
        $chart.SetSize(600, 400)
        $series = $chart.Series.Add(
            [OfficeOpenXml.ExcelAddress]::new(2, 3, $newLastRow, 3).Address,
            [OfficeOpenXml.ExcelAddress]::new(2, 1, $newLastRow, 1).Address
        )
        $chart.DataLabel.ShowPercent   = $true
        $chart.DataLabel.ShowCategory  = $true
        $chart.DataLabel.ShowLeaderLines = $true
        $changesApplied++
        Write-Host "[INFO] Rebuilt chart: Assessment Distribution (by Cost) — $chartRows categories" -ForegroundColor Green
    }
}

# ── Fix conditional formatting on User Report ────────────────────────────────
$userReport = $pkg.Workbook.Worksheets['User Report']
if ($userReport) {
    foreach ($cf in $userReport.ConditionalFormatting) {
        if ($cf.Formula -and $cf.Formula -match '<>"OK"') {
            $cf.Formula = $cf.Formula -replace '<>"OK"', '<>"No Findings"'
            $changesApplied++
            Write-Host "[INFO] Fixed conditional formatting: OK → No Findings" -ForegroundColor Green
        }
    }
}

# ── Remove Executive Summary tab ─────────────────────────────────────────────
$execTab = $pkg.Workbook.Worksheets['Executive Summary']
if ($execTab) {
    $pkg.Workbook.Worksheets.Delete($execTab)
    $changesApplied++
    Write-Host "[INFO] Removed Executive Summary tab (use HTML/Word for exec summary)" -ForegroundColor Yellow
}

# ── Save Pass 1 ──────────────────────────────────────────────────────────────
Close-ExcelPackage $pkg

# ══════════════════════════════════════════════════════════════════════════════
# PASS 2: Rebuild category tabs using Export-Excel (consistent table formatting)
# ══════════════════════════════════════════════════════════════════════════════
if ($tabRebuildData.Count -gt 0) {
    # Import User Report data as PSObjects (headers already renamed in Pass 1)
    $allUsers = @(Import-Excel -Path $customerXlsx -WorksheetName 'User Report')

    foreach ($tabName in $tabRebuildData.Keys) {
        $info = $tabRebuildData[$tabName]
        $matchCats = [System.Collections.Generic.HashSet[string]]::new(
            [string[]]$tabCategoryMap[$tabName], [System.StringComparer]::OrdinalIgnoreCase)

        $filtered = @($allUsers | Where-Object { $matchCats.Contains($_.'Assessment Category') })

        if ($filtered.Count -eq 0) { continue }

        # Export-Excel creates proper table with Medium6 style, frozen row, autofilter
        $filtered | Export-Excel -Path $customerXlsx -WorksheetName $tabName `
            -TableName ($tabName -replace '[^A-Za-z0-9]', '') -TableStyle Medium6 `
            -FreezeTopRow -AutoFilter -AutoSize -PassThru | ForEach-Object {
            $ws = $_.Workbook.Worksheets[$tabName]
            if ($tabColors.ContainsKey($tabName)) {
                $ws.TabColor = $tabColors[$tabName]
            }
            # Currency format on cost columns
            [int]$lCol = $ws.Dimension.End.Column
            for ($c = 1; $c -le $lCol; $c++) {
                $hdr = $ws.Cells[1, $c].Text
                if ($hdr -match 'Cost \(EUR\)' -or $hdr -match 'Price \(EUR\)') {
                    $ws.Column($c).Style.Numberformat.Format = '€#,##0.00'
                }
            }
            $_.Save()
            $_.Dispose()
        }

        $changesApplied++
        Write-Host "[INFO] Tab '$tabName': rebuilt $($info.OldCount) → $($filtered.Count) rows (Export-Excel)" -ForegroundColor Green
    }
}

Write-Host "`n[DONE] Customer-facing XLSX created: $customerXlsx" -ForegroundColor Green
Write-Host "[DONE] $changesApplied change(s) applied" -ForegroundColor Green
