<#
.SYNOPSIS
    Generates M365SkuData.json from Microsoft's licensing reference CSV.

.DESCRIPTION
    Parses the Microsoft licensing service plan reference CSV and extracts unique
    SKU part numbers with their friendly display names. Outputs M365SkuData.json
    in the same directory as this script.

    The CSV can be downloaded from:
    https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference

.EXAMPLE
    .\_extract_sku_names.ps1
    # Looks for ms_licensing_reference.csv in the script directory

.EXAMPLE
    .\_extract_sku_names.ps1 -CsvPath "C:\Downloads\licensing_reference.csv"
    # Use a specific CSV file
#>
param(
    [string]$CsvPath
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition

if (-not $CsvPath) {
    $CsvPath = Join-Path $scriptRoot 'ms_licensing_reference.csv'
}

if (-not (Test-Path $CsvPath)) {
    Write-Error "CSV not found: $CsvPath`nDownload from: https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference"
    return
}

$csv = Import-Csv $CsvPath
$unique = @{}
foreach ($row in $csv) {
    $sid = $row.String_Id
    $name = $row.Product_Display_Name
    if ($sid -and $name -and -not $unique.ContainsKey($sid)) {
        $unique[$sid] = $name
    }
}
Write-Host "Unique SKU part numbers: $($unique.Count)"

# Build updated JSON — preserve existing licensing matrices (suiteIncludes, planCapabilities, etc.)
$outPath = Join-Path $scriptRoot 'M365SkuData.json'
$jsonObj = $null
if (Test-Path $outPath) {
    try {
        $existingRaw = Get-Content $outPath -Raw -ErrorAction Stop
        $existingData = $existingRaw | ConvertFrom-Json
        # Rebuild as ordered hashtable to preserve section order
        $jsonObj = [ordered]@{
            _meta = "Generated from Microsoft licensing reference CSV (https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference). Last updated: $(Get-Date -Format 'yyyy-MM-dd')."
        }
        # Refresh skuFriendlyNames from CSV
        $jsonObj["skuFriendlyNames"] = [ordered]@{}
        foreach ($entry in ($unique.GetEnumerator() | Sort-Object Name)) {
            $jsonObj["skuFriendlyNames"][$entry.Key] = $entry.Value
        }
        # Carry forward all other sections unchanged
        $preserveSections = @("suiteIncludes","planCapabilities","planCapabilityAliases",
                              "addOnBundles","premiumSuites","skuCoverageAliases","skuMonthlyPricesEUR")
        foreach ($section in $preserveSections) {
            if ($existingData.PSObject.Properties[$section]) {
                $jsonObj[$section] = $existingData.$section
            }
        }
        Write-Host "Preserved $(@($preserveSections | Where-Object { $existingData.PSObject.Properties[$_] }).Count) existing licensing matrix section(s) from M365SkuData.json"
    } catch {
        Write-Warning "Failed to parse existing M365SkuData.json -- rebuilding with skuFriendlyNames only: $($_.Exception.Message)"
        $jsonObj = $null
    }
}

# Fallback: create fresh JSON with only skuFriendlyNames
if (-not $jsonObj) {
    $jsonObj = [ordered]@{
        _meta = "Generated from Microsoft licensing reference CSV (https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference). Last updated: $(Get-Date -Format 'yyyy-MM-dd')."
        skuFriendlyNames = [ordered]@{}
    }
    foreach ($entry in ($unique.GetEnumerator() | Sort-Object Name)) {
        $jsonObj["skuFriendlyNames"][$entry.Key] = $entry.Value
    }
}

$jsonObj | ConvertTo-Json -Depth 10 | Out-File $outPath -Encoding utf8
Write-Host "Written M365SkuData.json with $($unique.Count) SKU friendly names to: $outPath"
