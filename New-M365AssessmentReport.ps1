<#
.SYNOPSIS
    Generates a branded Word document from M365 License Assessment summary data.
.DESCRIPTION
    Parses the M365_OptimizationSummary*.txt file produced by the LOA script,
    builds a structured JSON payload, and calls generate-report.js (Node.js + docx)
    to produce a professionally formatted .docx report with EASI branding.
.PARAMETER OutputFolder
    Folder containing the LOA output files (summary .txt, CSV).
.PARAMETER CustomerName
    Customer name for the cover page.
.PARAMETER SummaryFile
    Specific summary file path. Defaults to the latest in OutputFolder.
.EXAMPLE
    .\New-M365AssessmentReport.ps1 -OutputFolder .\output -CustomerName "Xerius"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$OutputFolder,
    [Parameter(Mandatory)][string]$CustomerName,
    [string]$SummaryFile,
    [string]$LogoPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Resolve summary file ─────────────────────────────────────────────────────
if (-not $SummaryFile) {
    $candidates = @(Get-ChildItem -Path $OutputFolder -Filter 'M365_OptimizationSummary_*.txt' |
        Sort-Object LastWriteTime -Descending)
    if ($candidates.Count -eq 0) {
        Write-Error "No M365_OptimizationSummary_*.txt found in $OutputFolder"
        return
    }
    $SummaryFile = $candidates[0].FullName
}

Write-Host "Parsing: $SummaryFile" -ForegroundColor Cyan
$lines = Get-Content -Path $SummaryFile -Encoding UTF8

# ── Helper: extract value after colon ────────────────────────────────────────
function Get-LineValue([string]$line) {
    $idx = $line.IndexOf(':')
    if ($idx -ge 0) { return $line.Substring($idx + 1).Trim() }
    return $line.Trim()
}

# ── Helper: extract euro amount from a line ──────────────────────────────────
function Get-EuroAmount([string]$line) {
    if ($line -match '[^\d](\d{1,3}(?:\.\d{3})*(?:,\d+)?)\s*/yr') {
        return $Matches[1]
    }
    if ($line -match ':\s*[^€]*€([\d.,]+)') {
        return '€' + $Matches[1]
    }
    if ($line -match '€([\d.,]+)') {
        return '€' + $Matches[1]
    }
    return ''
}

# ── Helper: extract count in parentheses ─────────────────────────────────────
function Get-Count([string]$line) {
    # Prefer 'primary' count (matches heatmap tile logic) over 'flagged' count
    if ($line -match '(\d+)\s+primary') {
        return $Matches[1]
    }
    if ($line -match '\((\d+)\s+(?:flagged|users?)') {
        return $Matches[1]
    }
    if ($line -match '\((\d+)\s') {
        return $Matches[1]
    }
    return ''
}

# ── Parse structured sections ────────────────────────────────────────────────
$result = @{
    customerName         = $CustomerName
    reportDate           = ''
    tenantId             = ''
    totalAnnualSpend     = ''
    optimizationPotential = ''
    optimizationPct      = ''
    complianceCost       = ''
    tier1                = @{ items = @(); subtotal = '' }
    tier2                = @{ items = @(); subtotal = '' }
    poolWaste            = @{ items = @(); total = '' }
    accountFlags         = @()
    activityFlags        = @()
    quickWins            = @()
    rightSizing          = @()
    compliance           = @()
    copilot              = @{ totalHolders = 0; activeCount = 0; watchlistCount = 0; reclaimCount = 0; items = @() }
    cloudPc              = @()
    subscriptionAlerts   = @()
    departments          = @()
    recDistribution      = @()
}

$currentSection = ''

for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i]
    $trimmed = $line.Trim()

    # Report date (line 2)
    if ($i -eq 1 -and $trimmed -match '^\d{4}-\d{2}-\d{2}$') {
        $result.reportDate = $trimmed
        continue
    }

    # Tenant ID
    if ($trimmed -match '^Tenant:\s*(.+)') {
        $result.tenantId = $Matches[1]
        continue
    }

    # Section headers (ALL CAPS followed by colon or specific patterns)
    if ($trimmed -match '^EXECUTIVE FINANCIAL SUMMARY') { $currentSection = 'EXEC'; continue }
    if ($trimmed -match '^COST ANALYSIS') { $currentSection = 'COST'; continue }
    if ($trimmed -match '^QUICK WINS:') { $currentSection = 'QUICKWINS'; continue }
    if ($trimmed -match '^ACCOUNT & ROLE FLAGS:') { $currentSection = 'ACCOUNT'; continue }
    if ($trimmed -match '^ACTIVITY-BASED FLAGS:') { $currentSection = 'ACTIVITY'; continue }
    if ($trimmed -match '^RIGHT-SIZING OPPORTUNITIES:') { $currentSection = 'RIGHTSIZE'; continue }
    if ($trimmed -match '^PRODUCT-SPECIFIC FLAGS:') { $currentSection = 'PRODUCT'; continue }
    if ($trimmed -match '^COPILOT RECLAIM PIPELINE:') { $currentSection = 'COPILOT'; continue }
    if ($trimmed -match '^CLOUD PC UTILIZATION:') { $currentSection = 'CLOUDPC'; continue }
    if ($trimmed -match '^LICENSING COMPLIANCE\s*\(') { $currentSection = 'COMPLIANCE'; continue }
    if ($trimmed -match '^DATA QUALITY:') { $currentSection = 'DATAQUALITY'; continue }
    if ($trimmed -match '^SUBSCRIPTION WARNINGS') { $currentSection = 'SUBALERTS'; continue }
    if ($trimmed -match '^RECOMMENDATION DISTRIBUTION:') { $currentSection = 'RECDIST'; continue }
    if ($trimmed -match '^GROUP-BASED LICENSING') { $currentSection = 'GROUPS'; continue }
    if ($trimmed -match '^RUNTIME MAPPING') { $currentSection = 'MAPPING'; continue }
    if ($trimmed -match '^SECURITY & COMPLIANCE UPSELL') { $currentSection = 'UPSELL'; continue }
    if ($trimmed -eq '================================================================') { continue }
    if ($trimmed -match '^─{4,}') { continue }
    if ($trimmed -eq '' -or $trimmed -match '^NOTE:') { continue }

    switch ($currentSection) {
        'EXEC' {
            if ($trimmed -match 'Total Annual M365 Spend\s*:\s*(.+)') {
                $result.totalAnnualSpend = $Matches[1].Trim()
            }
            if ($trimmed -match 'Estimated Optimization Potential\s*:\s*(€[\d.,]+)\s*\((.+?)\)') {
                $result.optimizationPotential = $Matches[1]
                $result.optimizationPct = $Matches[2]
            }
            # Tier 1 items
            if ($trimmed -match 'TIER 1') { $currentSection = 'TIER1'; continue }
            if ($trimmed -match 'TIER 2') { $currentSection = 'TIER2'; continue }
            if ($trimmed -match 'UNASSIGNED LICENSES') { $currentSection = 'UNASSIGNED'; continue }
            if ($trimmed -match 'TENANT POOL') { $currentSection = 'POOL'; continue }
        }
        'TIER1' {
            if ($trimmed -match 'Tier 1 Subtotal\s*:\s*(.+?)\s*\(') {
                $result.tier1.subtotal = $Matches[1]
                $currentSection = 'EXEC'
                continue
            }
            if ($trimmed -match 'TIER 2') { $currentSection = 'TIER2'; continue }
            if ($trimmed -match 'UNASSIGNED') { $currentSection = 'UNASSIGNED'; continue }
            if ($trimmed -match ':\s*€' -and $trimmed -notmatch 'advisory') {
                $parts = $trimmed -split '\s*:\s*', 2
                $cat = $parts[0].Trim()
                $amt = ''
                $cnt = ''
                if ($parts.Count -gt 1) {
                    $amt = if ($parts[1] -match '(€[\d.,]+)') { $Matches[1] } else { '' }
                    $cnt = Get-Count $parts[1]
                }
                if ($amt -ne '€0,00') {
                    $result.tier1.items += @{ category = $cat; amount = $amt; count = $cnt }
                }
            }
        }
        'TIER2' {
            if ($trimmed -match 'Tier 2 Subtotal\s*:\s*(.+?)\s*\(') {
                $result.tier2.subtotal = $Matches[1]
                $currentSection = 'EXEC'
                continue
            }
            if ($trimmed -match 'TENANT POOL') { $currentSection = 'POOL'; continue }
            if ($trimmed -match ':\s*€') {
                $parts = $trimmed -split '\s*:\s*', 2
                $cat = $parts[0].Trim()
                # Expand abbreviations for customer-facing report
                $cat = $cat -replace '\bBiz\b', 'Business' `
                            -replace '\bO365\b', 'Office 365' `
                            -replace '\bEXO\b', 'Exchange Online' `
                            -replace '\bPBI\b', 'Power BI' `
                            -replace '\bPPU\b', 'Premium Per User' `
                            -replace 'F-License', 'Frontline License' `
                            -replace 'No Audio Conf\.', 'No Audio Conferencing'
                $amt = ''
                $cnt = ''
                if ($parts.Count -gt 1) {
                    $amt = if ($parts[1] -match '(€[\d.,]+)') { $Matches[1] } else { '' }
                    $cnt = Get-Count $parts[1]
                }
                if ($amt -ne '€0,00') {
                    $result.tier2.items += @{ category = $cat; amount = $amt; count = $cnt }
                }
            }
        }
        'UNASSIGNED' {
            if ($trimmed -match 'TIER 2') { $currentSection = 'TIER2'; continue }
            if ($trimmed -match 'TENANT POOL') { $currentSection = 'POOL'; continue }
        }
        'POOL' {
            if ($trimmed -match 'Pool waste total\s*:\s*(.+?)\s*\(') {
                $result.poolWaste.total = $Matches[1]
                $currentSection = 'EXEC'
                continue
            }
            if ($trimmed -match '^(.+?):\s*(\d+/\d+)\s+unassigned.*?€([\d.,]+)/yr') {
                $result.poolWaste.items += @{
                    sku   = $Matches[1].Trim()
                    seats = $Matches[2]
                    amount = '€' + $Matches[3]
                }
            }
        }
        'COST' {
            if ($trimmed -match 'Cost by Department') {
                $currentSection = 'DEPARTMENTS'
                continue
            }
        }
        'DEPARTMENTS' {
            if ($trimmed -match '^(.+?):\s*(\d+)\s+users?,\s*€([\d.,]+)/yr') {
                $result.departments += @{
                    name  = $Matches[1].Trim()
                    users = $Matches[2]
                    cost  = '€' + $Matches[3]
                }
            }
        }
        'QUICKWINS' {
            if ($trimmed -match ':\s*\d+') {
                $parts = $trimmed -split '\s*:\s*', 2
                if ($parts.Count -eq 2) {
                    $val = ($parts[1] -split '←')[0].Trim()
                    if ($val -ne '0') {
                        $result.quickWins += @(, @($parts[0].Trim(), $val))
                    }
                }
            }
        }
        'ACCOUNT' {
            if ($trimmed -match ':\s*\d+') {
                $parts = $trimmed -split '\s*:\s*', 2
                if ($parts.Count -eq 2) {
                    $val = ($parts[1] -split '←')[0].Trim()
                    if ($val -ne '0') {
                        $result.accountFlags += @(, @($parts[0].Trim(), $val))
                    }
                }
            }
        }
        'ACTIVITY' {
            if ($trimmed -match ':\s*\d+') {
                $parts = $trimmed -split '\s*:\s*', 2
                if ($parts.Count -eq 2) {
                    $val = ($parts[1] -split '←')[0].Trim()
                    if ($val -ne '0') {
                        $result.activityFlags += @(, @($parts[0].Trim(), $val))
                    }
                }
            }
        }
        'RIGHTSIZE' {
            if ($trimmed -match ':\s*\d+') {
                $parts = $trimmed -split '\s*:\s*', 2
                if ($parts.Count -eq 2) {
                    $val = ($parts[1] -split '←')[0].Trim()
                    if ($val -ne '0') {
                        $result.rightSizing += @(, @($parts[0].Trim(), $val))
                    }
                }
            }
        }
        'COMPLIANCE' {
            if ($trimmed -match ':\s*\d+') {
                $parts = $trimmed -split '\s*:\s*', 2
                if ($parts.Count -eq 2) {
                    $val = ($parts[1] -split '←')[0].Trim()
                    # Skip zero-count items
                    if ($val -ne '0') {
                        $result.compliance += @(, @($parts[0].Trim(), $val))
                    }
                }
            }
        }
        'COPILOT' {
            if ($trimmed -match 'Total Copilot license holders\s*:\s*(\d+)') {
                $result.copilot.totalHolders = [int]$Matches[1]
            }
            if ($trimmed -match '[├└─]+\s*(.+?)\s*:\s*(\d+)') {
                $label = $Matches[1].Trim()
                $count = $Matches[2]
                $note = ''
                if ($trimmed -match '←\s*(.+)') { $note = $Matches[1].Trim() }
                $result.copilot.items += @(, @($label, "$count — $note".Trim(' — ')))
                # Extract numeric counts for metrics
                if ($label -match 'KEEP.*active') { $result.copilot.activeCount = [int]$count }
                if ($label -match 'WATCHLIST') { $result.copilot.watchlistCount = [int]$count }
                if ($label -match 'RECLAIM') { $result.copilot.reclaimCount = [int]$count }
            }
            if ($trimmed -match 'Reclaim savings|Watchlist at-risk') {
                $parts = $trimmed -split '\s*:\s*', 2
                if ($parts.Count -eq 2) {
                    $result.copilot.items += @(, @($parts[0].Trim(), ($parts[1] -split '←')[0].Trim()))
                }
            }
        }
        'CLOUDPC' {
            # Only capture actual Cloud PC items — the summary groups other product flags here
            if ($trimmed -match 'Cloud PC' -and $trimmed -match ':\s*\d+') {
                $parts = $trimmed -split '\s*:\s*', 2
                if ($parts.Count -eq 2) {
                    $val = ($parts[1] -split '←')[0].Trim()
                    if ($val -ne '0') {
                        $result.cloudPc += @(, @($parts[0].Trim(), $val))
                    }
                }
            }
        }
        'SUBALERTS' {
            if ($trimmed -match '^(.+?):\s*(Enabled|Warning|Suspended),\s*(\d+)d\s+remaining') {
                $result.subscriptionAlerts += @{
                    sku    = $Matches[1].Trim()
                    status = $Matches[2]
                    days   = $Matches[3] + 'd'
                }
            }
        }
        'RECDIST' {
            if ($trimmed -match '^(.+?):\s*(\d+)$') {
                $cat = $Matches[1].Trim()
                # Skip non-actionable categories (informational, not findings)
                if ($cat -notin @('No Findings', 'Unlicensed', 'Copilot Active')) {
                    $result.recDistribution += @{
                        category = $cat
                        count    = $Matches[2]
                    }
                }
            }
        }
    }
}

# ── Override financial data from HTML heatmap (single source of truth) ────────
# The HTML dashboard computes savings per-user with primary-category attribution.
# The Word report must show identical numbers. Text summary is kept only for
# qualitative/profile data (account flags, activity flags, detailed breakdowns).
$heatmapHtml = Get-ChildItem -Path $OutputFolder -Filter 'M365_OptimizationHeatmap_*.html' |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($heatmapHtml) {
    try {
        $htmlContent = Get-Content -Path $heatmapHtml.FullName -Raw -Encoding UTF8
        Write-Host "  Heatmap source: $($heatmapHtml.Name)" -ForegroundColor DarkGray
        $deDE = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')

        # ── KPI hero numbers ─────────────────────────────────────────────────
        if ($htmlContent -match '(?s)Total Annual Spend</div>\s*<div[^>]*>(.*?)</div>') {
            $raw = $Matches[1] -replace '&euro;', '€' -replace '<[^>]+>', ''
            if ($raw -match '€?([\d.,]+)') { $result.totalAnnualSpend = '€' + $Matches[1] }
        }
        if ($htmlContent -match '(?s)Potential Annual Savings</div>\s*<div[^>]*>(.*?)</div>') {
            $raw = $Matches[1] -replace '&euro;', '€' -replace '<[^>]+>', ''
            if ($raw -match '€?([\d.,]+)') { $result.optimizationPotential = '€' + $Matches[1] }
        }
        if ($htmlContent -match '(?s)Potential Annual Savings</div>[\s\S]*?<div class="sub">(.*?)</div>') {
            $result.optimizationPct = $Matches[1] -replace '<[^>]+>', ''
        }
        if ($htmlContent -match '(?s)Potential Compliance Costs</div>\s*<div[^>]*>(.*?)</div>') {
            $raw = $Matches[1] -replace '&euro;', '€' -replace '<[^>]+>', ''
            if ($raw -match '€?([\d.,]+)') { $result.complianceCost = '€' + $Matches[1] + '/yr' }
        }
        Write-Host "  KPIs from heatmap — Spend: $($result.totalAnnualSpend), Savings: $($result.optimizationPotential) ($($result.optimizationPct)), Compliance: $($result.complianceCost)" -ForegroundColor DarkGray

        # ── TILES → Tier 1 / Tier 2 items + subtotals ────────────────────────
        if ($htmlContent -match 'const TILES\s*=\s*(\[.*?\]);') {
            $tiles = $Matches[1] | ConvertFrom-Json
            $heatmapTier1 = @()
            $heatmapTier2 = @()
            [decimal]$tier1Sum = 0
            [decimal]$tier2Sum = 0
            foreach ($t in $tiles) {
                if ($t.savings -gt 0 -and $t.label -ne 'Unassigned Licenses') {
                    $item = @{
                        category = $t.label
                        amount   = '€' + ([decimal]$t.savings).ToString('N2', $deDE)
                        count    = [string]$t.users
                    }
                    if ($t.tier -eq 1) {
                        $heatmapTier1 += $item
                        $tier1Sum += [decimal]$t.savings
                    } elseif ($t.tier -eq 2) {
                        $heatmapTier2 += $item
                        $tier2Sum += [decimal]$t.savings
                    }
                }
            }
            if ($heatmapTier1.Count -gt 0) {
                $result.tier1.items = $heatmapTier1
                $result.tier1.subtotal = '€' + $tier1Sum.ToString('N2', $deDE)
            }
            if ($heatmapTier2.Count -gt 0) {
                $result.tier2.items = $heatmapTier2
                $result.tier2.subtotal = '€' + $tier2Sum.ToString('N2', $deDE)
            }
            Write-Host "  Tier 1: $($heatmapTier1.Count) tile(s), €$($tier1Sum.ToString('N2', $deDE))" -ForegroundColor DarkGray
            Write-Host "  Tier 2: $($heatmapTier2.Count) tile(s), €$($tier2Sum.ToString('N2', $deDE))" -ForegroundColor DarkGray
        }

        # ── POOL_SKUS → Pool waste ───────────────────────────────────────────
        if ($htmlContent -match 'const POOL_SKUS\s*=\s*(\[.*?\]);') {
            $poolSkus = $Matches[1] | ConvertFrom-Json
            if ($poolSkus.Count -gt 0) {
                $poolItems = @()
                [decimal]$poolTotal = 0
                foreach ($p in $poolSkus) {
                    $poolItems += @{
                        sku    = $p.sku
                        seats  = [string]$p.unassigned
                        amount = '€' + ([decimal]$p.waste).ToString('N2', $deDE)
                    }
                    $poolTotal += [decimal]$p.waste
                }
                $result.poolWaste.items = $poolItems
                $result.poolWaste.total = '€' + $poolTotal.ToString('N2', $deDE)
                Write-Host "  Pool waste: $($poolItems.Count) SKU(s), €$($poolTotal.ToString('N2', $deDE))" -ForegroundColor DarkGray
            }
        }

        # ── SUB_ALERTS → Subscription alerts ─────────────────────────────────
        if ($htmlContent -match 'const SUB_ALERTS\s*=\s*(\[.*?\]);') {
            $subAlerts = $Matches[1] | ConvertFrom-Json
            if ($subAlerts.Count -gt 0) {
                $alertItems = @()
                foreach ($a in $subAlerts) {
                    $alertItems += @{
                        sku    = $a.sku
                        status = $a.status
                        days   = [string]$a.days + 'd'
                    }
                }
                $result.subscriptionAlerts = $alertItems
                Write-Host "  Sub alerts: $($alertItems.Count) alert(s)" -ForegroundColor DarkGray
            }
        }

        # ── COPILOT_ROI → Copilot metrics ────────────────────────────────────
        if ($htmlContent -match 'const COPILOT_ROI\s*=\s*(\{.*?\});') {
            $cpRoi = $Matches[1] | ConvertFrom-Json
            $result.copilot.totalHolders = [int]$cpRoi.total
            $result.copilot.activeCount = [int]$cpRoi.active
            $result.copilot.watchlistCount = [int]$cpRoi.watchlist
            $result.copilot.reclaimCount = [int]$cpRoi.reclaim
            Write-Host "  Copilot ROI: $($cpRoi.total) holders, $($cpRoi.active) active" -ForegroundColor DarkGray
        }

        # ── Rec distribution from USERS (primary category counts) ────────────
        if ($htmlContent -match '(?s)const USERS\s*=\s*(\[.*?\]);\s*const SKUS') {
            $users = $Matches[1] | ConvertFrom-Json
            $catCounts = @{}
            foreach ($u in $users) {
                $cat = $u.Category
                if ($cat -and $cat -notin @('No Findings', 'Unlicensed', 'Copilot Active')) {
                    if (-not $catCounts.ContainsKey($cat)) { $catCounts[$cat] = 0 }
                    $catCounts[$cat]++
                }
            }
            $result.recDistribution = @($catCounts.GetEnumerator() |
                Sort-Object Value -Descending |
                ForEach-Object { @{ category = $_.Key; count = [string]$_.Value } })
            Write-Host "  Rec distribution: $($result.recDistribution.Count) categories from USERS" -ForegroundColor DarkGray
        }

    } catch {
        Write-Host "[WARN] Could not extract data from HTML heatmap: $_" -ForegroundColor Yellow
    }
} else {
    Write-Host "[WARN] No heatmap HTML found — Word report will use text summary data only" -ForegroundColor Yellow
}

# ── Write JSON and invoke Node.js ────────────────────────────────────────────
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$jsonPath = Join-Path $OutputFolder "report_data_$timestamp.json"
$docxPath = Join-Path $OutputFolder "M365_AssessmentReport_$timestamp.docx"

$result | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonPath -Encoding UTF8

$scriptDir = $PSScriptRoot
$generatorPath = Join-Path $scriptDir 'generate-report.js'

Write-Host "Generating Word document..." -ForegroundColor Cyan
$nodeArgs = @($generatorPath, $jsonPath, $docxPath)
if ($LogoPath) {
    $nodeArgs += '--cover'
    $nodeArgs += $LogoPath
}

& node @nodeArgs

if ($LASTEXITCODE -eq 0) {
    Write-Host "Report saved to: $docxPath" -ForegroundColor Green
    # Clean up temp JSON
    Remove-Item -Path $jsonPath -Force -ErrorAction SilentlyContinue
} else {
    Write-Host "Report generation failed. JSON data preserved at: $jsonPath" -ForegroundColor Red
}
