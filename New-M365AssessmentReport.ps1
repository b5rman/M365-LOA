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
    if ($line -match '\((\d+)\s+(?:flagged|users?|primary)') {
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
                    $cnt = if ($parts[1] -match '\((\d+)\s') { $Matches[1] } else { '' }
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
                    $cnt = if ($parts[1] -match '\((\d+)\s') { $Matches[1] } else { '' }
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
                if ($cat -notin @('OK', 'Unlicensed', 'Copilot Active')) {
                    $result.recDistribution += @{
                        category = $cat
                        count    = $Matches[2]
                    }
                }
            }
        }
    }
}

# ── Extract compliance cost from HTML heatmap ─────────────────────────────────
$heatmapHtml = Get-ChildItem -Path $OutputFolder -Filter 'M365_OptimizationHeatmap_*.html' |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($heatmapHtml -and -not $result.complianceCost) {
    try {
        $htmlContent = Get-Content -Path $heatmapHtml.FullName -Raw -Encoding UTF8
        # Extract compliance cost KPI: the value div after "Potential Compliance Costs" label
        if ($htmlContent -match 'Potential Compliance Costs</div>\s*<div[^>]*>(.*?)</div>') {
            $rawVal = $Matches[1] -replace '&euro;', '€' -replace '<[^>]+>', ''
            if ($rawVal -match '([\d.,]+)(/yr)?') {
                $result.complianceCost = '€' + $Matches[1] + '/yr'
            }
        }
        if ($result.complianceCost) {
            Write-Host "  Compliance cost extracted from heatmap: $($result.complianceCost)" -ForegroundColor DarkGray
        }
    } catch {
        Write-Host "[WARN] Could not extract compliance cost from HTML heatmap: $_" -ForegroundColor Yellow
    }
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
