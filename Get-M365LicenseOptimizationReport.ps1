<#
.SYNOPSIS
    M365 License Optimization Report — licenses, service usage intensity, and platform usage per user.

.DESCRIPTION
    Pulls eleven Graph Reports API endpoints + sign-in activity + per-user license detail,
    merges everything by UPN, and produces a consolidated report with computed recommendation
    flags to support license optimization conversations with customers.

    Reports pulled:
      1. getOffice365ActiveUserDetail       — last-activity dates per service + license flags
      2. getM365AppUserDetail               — per-app × per-platform usage (desktop vs web vs mobile)
      3. getOffice365ActivationsUserDetail   — which platforms each product was activated on
      4. getEmailActivityUserDetail          — send/receive/read/meeting counts
      5. getTeamsUserActivityUserDetail      — chat/call/meeting counts
      6. getOneDriveActivityUserDetail       — file view/modify/sync/share counts
      7. getSharePointActivityUserDetail     — file/page activity counts
      8. getMailboxUsageDetail               — mailbox size, item count, quota status
      9. getEmailAppUsageUserDetail          — which email clients each user connects with
     10. getOneDriveUsageAccountDetail       — per-user OneDrive storage consumed & file count
     11. getTeamsDeviceUsageUserDetail        — which platforms users access Teams from

    Additionally:
      - Sign-in activity (beta API)        — last interactive & non-interactive sign-in dates
      - License assignment states (beta)   — direct vs group-based, overlap detection
      - Subscription lifecycle             — expiry dates & status per SKU
      - Get-MgDirectoryRole                — admin role assignments per user
      - Get-EXOMailbox                     — mailbox type (User/Shared/Room/Equipment)
      - Get-MgUser (AssignedLicenses)       — assigned SKUs + disabled plans per user (bulk)
      - Get-MgSubscribedSku                — tenant SKU inventory

    Computed columns in the merged report:
      - UsesDesktopApps          : TRUE if any M365 app used on Windows or Mac
      - UsesWebAppsOnly          : TRUE if apps used on Web but NOT on Windows/Mac
      - UsesMobileOnly           : TRUE if apps used on Mobile but NOT on Windows/Mac/Web
      - DesktopAppsList           : which specific apps were used on desktop (e.g. "Outlook, Word, Excel")
      - WebAppsList               : which specific apps were used on web
      - ExchangeIntensity         : Low / Medium / High based on send+receive counts
      - TeamsIntensity            : Low / Medium / High based on chat+call+meeting counts
      - OneDriveIntensity         : Low / Medium / High
      - SharePointIntensity       : Low / Medium / High
      - MailboxSizeMB             : mailbox storage in MB
      - OneDriveStorageMB         : OneDrive storage in MB
      - EmailClientsUsed          : which email apps the user connects with
      - TeamsDevicePlatforms      : which platforms the user accesses Teams from
      - TeamsNoDesktop            : TRUE if Teams used via web/mobile but not desktop
      - MailboxType               : User / Shared / Room / Equipment
      - AdminRoles               : directory role assignments (e.g. "Global Administrator")
      - UserType                  : Member / Guest
      - IsGuestWithLicense        : TRUE if guest user holds a paid license
      - LicenseAssignment         : Direct / Group / Direct + Group per SKU pattern
      - OverlappingLicenses       : SKUs assigned both directly and via group (waste)
      - LicenseGroups             : group names responsible for license assignments
      - LicenseFriendlyNames      : human-readable SKU names
      - LastSignIn                : last interactive sign-in date
      - DaysSinceLastSignIn       : days since last interactive sign-in
      - LicenseRecommendation     : plain-English recommendation string
      - AccountEnabled            : TRUE/FALSE — disabled accounts with licenses are waste

    Optimization checks performed:
      1. Business 300-seat limit    — warns when Business-family SKUs approach 300 cap
      2. Power BI Pro review         — Pro users when tenant has Premium Capacity (consumers only — creators need Pro)
      3. Duplicate suite coverage   — standalone SKU already included in assigned suite
      4. E3 → E5 upgrade opportunity — E3 + 2+ add-ons may be cheaper as E5
      5. Visio/Project shelfware    — expensive SKU with no detected activity
      6. Teams Phone PSTN review    — Teams Phone but no Microsoft Calling Plan (may use Direct Routing / Operator Connect)
      7. Copilot adoption           — licensed but inactive → reallocate
      8. Frontline right-sizing     — E3/E5 user who only uses web/mobile → F1/F3
      9. EXO Plan 2 downgrade       — mailbox under 50 GB, Plan 1 may suffice
     10. Disabled account licensed   — sign-in blocked but license still assigned

.PARAMETER ReportPeriod
    Usage-report lookback window. Valid: D7, D30, D90, D180. Default: D180.

.PARAMETER OutputFolder
    Folder for output CSVs. Default: current directory.

.PARAMETER IncludeDisabledAccounts
    If set, includes disabled (AccountEnabled = $false) accounts in the full per-user
    report with all recommendation checks. Without this switch, disabled accounts are
    still fetched and counted in summary stats (Disabled Accounts waste metric) but
    are excluded from the main CSV to keep the report focused on active users.

.PARAMETER KeepHashedUPNs
    By default the script unhides user data in Graph usage reports so UPNs are readable.
    Pass this switch to skip that step and leave UPNs hashed. Unhiding requires
    ReportSettings.ReadWrite.All (already included in the app registration). The setting
    persists — the customer must re-enable privacy manually in M365 Admin Center >
    Settings > Org settings > Reports when the audit engagement is complete.

.PARAMETER ExchangeHighThreshold
    Emails (sent+received) above this = High intensity. Default: 500.

.PARAMETER ExchangeLowThreshold
    Emails (sent+received) below this = Low intensity. Default: 50.

.PARAMETER TeamsHighThreshold
    Teams actions (chats+calls+meetings) above this = High. Default: 200.

.PARAMETER TeamsLowThreshold
    Teams actions below this = Low. Default: 20.

.PARAMETER InactiveSignInDays
    Days since last interactive sign-in to flag a user as dormant. Default: 30.

.PARAMETER PricingCsvPath
    Path to a CSV with SkuPartNumber,MonthlyPriceEUR columns. Default: M365SkuPricing.csv alongside the script.

.PARAMETER NoExcel
    Skip Excel workbook generation even if the ImportExcel module is installed.

.PARAMETER AutoInstallModules
    When set, missing required PowerShell modules (Microsoft.Graph.*) are installed automatically
    using Install-Module -Scope CurrentUser. Without this switch, the script stops with a clear
    error listing the exact Install-Module commands needed. This is the safe default for
    locked-down servers, CI pipelines, and multi-admin environments.

.PARAMETER SkipEXO
    Skip Exchange Online connection entirely. Disables mailbox type detection (Shared/Room/Equipment),
    litigation hold detection, and Defender for Office 365 policy coverage evaluation.
    Use this when the ExchangeOnlineManagement module is unavailable or connection is undesired.
    There is NO Microsoft Graph API equivalent for these Exchange Online-specific features.

.PARAMETER PriorReportPath
    Path to a previous run's main CSV (M365_LicenseOptimization_*.csv) for delta analysis.
    When provided, generates a delta report showing user additions/removals, license changes,
    cost trends, recommendation changes, dormancy shifts, and Copilot adoption tracking.
    Works with CSVs from any prior version (missing columns are handled gracefully).

.PARAMETER ClientId
    Application (client) ID from the App Registration created by LOA-App-Registration-Setup.ps1.
    Required for certificate-based authentication. Must be used together with -TenantId and
    either -CertificateThumbprint or -CertificatePath.

.PARAMETER TenantId
    Azure AD tenant ID. Required for certificate-based authentication.

.PARAMETER CertificateThumbprint
    Thumbprint of a certificate already installed in Cert:\CurrentUser\My.
    Use this when the .pfx has been pre-imported into the certificate store.

.PARAMETER CertificatePath
    Path to a .pfx certificate file. The certificate will be imported into the current user
    store automatically. Alternative to -CertificateThumbprint for first-time runs.

.PARAMETER CertificatePassword
    SecureString password for the .pfx file. If omitted when -CertificatePath is used,
    you will be prompted interactively.

.EXAMPLE
    .\Get-M365LicenseOptimizationReport.ps1 -ReportPeriod D90 -OutputFolder "C:\Reports"

.EXAMPLE
    .\Get-M365LicenseOptimizationReport.ps1 -KeepHashedUPNs -ExchangeHighThreshold 1000

.EXAMPLE
    # Certificate-based auth using App Registration from LOA-App-Registration-Setup.ps1
    .\Get-M365LicenseOptimizationReport.ps1 -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -TenantId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
        -CertificateThumbprint "ABCDEF1234567890ABCDEF1234567890ABCDEF12"

.EXAMPLE
    # First-time run with .pfx file (certificate auto-imported)
    .\Get-M365LicenseOptimizationReport.ps1 -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -TenantId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
        -CertificatePath ".\M365-LOA-Audit-Cert.pfx"
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [ValidateSet("D7", "D30", "D90", "D180")]
    [Parameter(HelpMessage = "Usage-report lookback window: D7, D30, D90, D180.")]
    [string]$ReportPeriod = "D90",

    [Parameter(HelpMessage = "Folder for output files. Created if it does not exist.")]
    [string]$OutputFolder = (Join-Path (Get-Location).Path "output"),

    [Parameter(HelpMessage = "Include disabled (AccountEnabled=false) unlicensed users in the report.")]
    [switch]$IncludeDisabledAccounts,

    [Parameter(HelpMessage = "Skip unhiding user data in Graph usage reports (UPNs will remain hashed).")]
    [switch]$KeepHashedUPNs,

    [ValidateRange(1, [int]::MaxValue)]
    [Parameter(HelpMessage = "Emails (sent+received) above this = High intensity.")]
    [int]$ExchangeHighThreshold = 500,

    [ValidateRange(0, [int]::MaxValue)]
    [Parameter(HelpMessage = "Emails (sent+received) below this = Low intensity.")]
    [int]$ExchangeLowThreshold  = 50,

    [ValidateRange(1, [int]::MaxValue)]
    [Parameter(HelpMessage = "Teams actions (chats+calls+meetings) above this = High intensity.")]
    [int]$TeamsHighThreshold    = 200,

    [ValidateRange(0, [int]::MaxValue)]
    [Parameter(HelpMessage = "Teams actions below this = Low intensity.")]
    [int]$TeamsLowThreshold     = 20,

    [ValidateRange(1, [int]::MaxValue)]
    [Parameter(HelpMessage = "OneDrive actions above this = High intensity.")]
    [int]$OneDriveHighThreshold = 100,

    [ValidateRange(0, [int]::MaxValue)]
    [Parameter(HelpMessage = "OneDrive actions below this = Low intensity.")]
    [int]$OneDriveLowThreshold  = 10,

    [ValidateRange(1, [int]::MaxValue)]
    [Parameter(HelpMessage = "SharePoint actions above this = High intensity.")]
    [int]$SharePointHighThreshold = 100,

    [ValidateRange(0, [int]::MaxValue)]
    [Parameter(HelpMessage = "SharePoint actions below this = Low intensity.")]
    [int]$SharePointLowThreshold  = 10,

    [ValidateRange(1, 365)]
    [Parameter(HelpMessage = "Days since last sign-in to flag as dormant (1-365).")]
    [int]$InactiveSignInDays      = 30,

    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [Parameter(HelpMessage = "Path to SKU pricing CSV (default: M365SkuPricing.csv alongside script).")]
    [string]$PricingCsvPath,

    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [Parameter(HelpMessage = "Path to M365SkuData.json with SKU friendly names and pricing.")]
    [string]$SkuDataPath,

    [ValidateRange(1, 365)]
    [Parameter(HelpMessage = "Warn if M365SkuData.json is older than this many days (default: 90).")]
    [int]$SkuStalenessDays = 90,

    [Parameter(HelpMessage = "Abort if external SKU data (M365SkuData.json) is missing or stale.")]
    [switch]$ForceSkuRefresh,

    [Parameter(HelpMessage = "Path to LOA_RulePack_M365.json with audit checklist rules and doc references.")]
    [string]$RulePackPath,

    [Parameter(HelpMessage = "Skip Excel workbook generation (CSVs still produced).")]
    [switch]$NoExcel,

    [Parameter(HelpMessage = "Automatically install missing Microsoft.Graph modules.")]
    [switch]$AutoInstallModules,

    [Parameter(HelpMessage = "Skip Exchange Online; disables mailbox type, litigation hold, and MDO coverage.")]
    [switch]$SkipEXO,

    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [Parameter(HelpMessage = "Path to a prior run CSV for delta analysis.")]
    [string]$PriorReportPath,

    [ValidateRange(1, 11)]
    [Parameter(HelpMessage = "Maximum parallel Graph API report downloads (1-11). Lower values reduce throttling risk.")]
    [int]$MaxParallel = 6,

    [Parameter(HelpMessage = "Application (client) ID from the App Registration. Enables certificate-based auth.")]
    [string]$ClientId,

    [Parameter(HelpMessage = "Tenant ID for certificate-based auth.")]
    [string]$TenantId,

    [Parameter(HelpMessage = "Certificate thumbprint for app-only auth (from LOA-App-Registration-Setup.ps1).")]
    [string]$CertificateThumbprint,

    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [Parameter(HelpMessage = "Path to .pfx certificate file. Alternative to CertificateThumbprint (cert does not need to be pre-installed).")]
    [string]$CertificatePath,

    [Parameter(HelpMessage = "Password for the .pfx certificate file (SecureString).")]
    [securestring]$CertificatePassword
)

#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "Continue"

# ── Validate threshold pairs (High must be > Low) ──
$thresholdPairs = @(
    ,@("ExchangeHighThreshold",   "ExchangeLowThreshold",   $ExchangeHighThreshold,   $ExchangeLowThreshold)
    ,@("TeamsHighThreshold",      "TeamsLowThreshold",      $TeamsHighThreshold,      $TeamsLowThreshold)
    ,@("OneDriveHighThreshold",   "OneDriveLowThreshold",   $OneDriveHighThreshold,   $OneDriveLowThreshold)
    ,@("SharePointHighThreshold", "SharePointLowThreshold", $SharePointHighThreshold, $SharePointLowThreshold)
)
foreach ($pair in $thresholdPairs) {
    if ($pair[2] -le $pair[3]) {
        throw "$($pair[0]) ($($pair[2])) must be greater than $($pair[1]) ($($pair[3]))."
    }
}

# ── Version tracking (LOA v1.0 spec §9) ──
$MappingVersion              = "1.3"    # Increment when $suiteIncludes or $planCapabilities changes
$RecommendationLogicVersion  = "1.2.0"  # Increment when recommendation logic changes

# ── Script-scoped warnings collector — surfaces skipped data in the summary ──
$script:skippedDataWarnings = [System.Collections.Generic.List[string]]::new()

# ── Ensure OutputFolder exists (create if missing) ──
# NOTE: SupportsShouldProcess is declared for the privacy toggle (unhide UPNs by default; skip with -KeepHashedUPNs)
# (the only destructive tenant-modifying action). File output IS the script's purpose,
# so CSV/TXT/XLSX writes are not gated by -WhatIf by design.
if (-not (Test-Path $OutputFolder)) {
    if ($PSCmdlet.ShouldProcess($OutputFolder, "Create output directory")) {
        New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
        Write-Host "Created output folder: $OutputFolder" -ForegroundColor DarkGray
    }
}

# ── Diagnostic log file (always created for post-run analysis) ──
$script:logFile = Join-Path $OutputFolder "M365_OptimizationLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# ═══════════════════════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

function ConvertTo-CsvLine {
    <#
    .SYNOPSIS
        Converts a PSCustomObject to a properly escaped CSV line.
        Handles embedded quotes, commas, and newlines in field values.
    #>
    param (
        [PSCustomObject]$Row,
        [string[]]$Columns
    )
    $parts = [string[]]::new($Columns.Count)
    for ($i = 0; $i -lt $Columns.Count; $i++) {
        $val = $Row.($Columns[$i])
        if ($null -eq $val) {
            $parts[$i] = '""'
        } else {
            [string]$s = $val.ToString()
            if ($s -match '[",\r\n]') {
                $parts[$i] = '"' + ($s -replace '"', '""') + '"'
            } else {
                $parts[$i] = $s
            }
        }
    }
    return $parts -join ','
}

function Write-Log {
    <#
    .SYNOPSIS
        Appends a timestamped, leveled entry to the diagnostic log file.
        For errors, pass -ErrorRecord $_ to capture exception, line number, and stack trace.
    #>
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO',
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )
    $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
    $entry = "[$stamp] [$Level] $Message"
    if ($ErrorRecord) {
        $entry += "`n  Exception : $($ErrorRecord.Exception.GetType().FullName): $($ErrorRecord.Exception.Message)"
        $entry += "`n  Category  : $($ErrorRecord.CategoryInfo.Category)"
        $entry += "`n  Target    : $($ErrorRecord.CategoryInfo.TargetName)"
        $entry += "`n  Line      : $($ErrorRecord.InvocationInfo.ScriptLineNumber)"
        $entry += "`n  Statement : $($ErrorRecord.InvocationInfo.Line.Trim())"
        $entry += "`n  Stack     :`n$($ErrorRecord.ScriptStackTrace)"
        # Include inner exception if present (e.g. Graph API / MSAL chains)
        $inner = $ErrorRecord.Exception.InnerException
        if ($inner) {
            $entry += "`n  Inner     : $($inner.GetType().FullName): $($inner.Message)"
        }
    }
    # Security: redact Bearer tokens / JWT values that may leak via exception messages
    $entry = $entry -replace '(Bearer\s+)[A-Za-z0-9\-_\.]{20,}', '$1[REDACTED]'
    $entry = $entry -replace '(eyJ[A-Za-z0-9\-_]{10,}\.[A-Za-z0-9\-_]+\.[A-Za-z0-9\-_]*)', '[REDACTED-JWT]'
    if ($script:logFile) {
        try { Add-Content -Path $script:logFile -Value $entry -Encoding UTF8 -ErrorAction SilentlyContinue }
        catch { }   # never let logging itself kill the script
    }
}

function Invoke-GraphWithRetry {
    <#
    .SYNOPSIS
        Wraps Invoke-MgGraphRequest with automatic retry on 429 (throttled) and
        transient 5xx errors. Uses exponential backoff with Retry-After header.
    #>
    param (
        [string]$Method = "GET",
        [string]$Uri,
        [string]$Body,
        [string]$ContentType,
        [string]$OutputFilePath,
        [string]$OutputType,
        [hashtable]$Headers,
        [int]$MaxRetries = 5,
        [int]$BaseDelaySeconds = 2
    )
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            $params = @{ Method = $Method; Uri = $Uri }
            if ($Body)           { $params.Body           = $Body }
            if ($ContentType)    { $params.ContentType    = $ContentType }
            if ($OutputFilePath) { $params.OutputFilePath = $OutputFilePath }
            if ($OutputType)     { $params.OutputType     = $OutputType }
            if ($Headers)        { $params.Headers        = $Headers }
            return (Invoke-MgGraphRequest @params)
        } catch {
            $statusCode = $null
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }
            $isRetryable = ($statusCode -eq 429) -or ($statusCode -ge 500 -and $statusCode -lt 600)
            if ($isRetryable -and $attempt -le $MaxRetries) {
                # Respect Retry-After header if present, otherwise exponential backoff
                $retryAfter = $BaseDelaySeconds * [math]::Pow(2, $attempt - 1)
                if ($_.Exception.Response -and $_.Exception.Response.Headers) {
                    try {
                        $raHeader = $_.Exception.Response.Headers | Where-Object { $_.Key -eq 'Retry-After' }
                        if ($raHeader -and $raHeader.Value) {
                            $parsedRA = [int]($raHeader.Value | Select-Object -First 1)
                            if ($parsedRA -gt 0) { $retryAfter = $parsedRA }
                        }
                    } catch { }
                }
                $retryAfter = [math]::Min($retryAfter, 120)  # Cap at 2 minutes
                Write-Log "Graph API throttled/error ($statusCode) on $Uri — retry $attempt/$MaxRetries in ${retryAfter}s" -Level WARN
                Write-Host "    Graph API throttled/error ($statusCode) — retry $attempt/$MaxRetries in ${retryAfter}s ..." -ForegroundColor Yellow
                Start-Sleep -Seconds $retryAfter
            } else {
                throw   # Re-throw non-retryable or exhausted retries
            }
        }
    }
}

function Get-Intensity {
    param([int]$Value, [int]$Low, [int]$High)
    if ($Value -ge $High) { return "High" }
    elseif ($Value -le $Low) { return "Low" }
    else { return "Medium" }
}

function Parse-NumericField {
    param([string]$Value)
    $cleaned = $Value -replace '[^\d]',''
    if ($cleaned) { return [int]$cleaned }
    return 0
}

function Parse-DoubleField {
    param([string]$Value)
    if (-not $Value) { return 0 }
    # Normalise locale-variant decimals: "1.234,56" → "1234.56", "1,234.56" → "1234.56"
    # Detect comma-as-decimal (digit,digit at end with no dot after): 1234,56
    if ($Value -match ',\d{1,2}$' -and $Value -notmatch '\.\d') {
        $cleaned = ($Value -replace '[^\d,]','') -replace ',','.'
    } else {
        # Dot-as-decimal or integer — strip thousands separators (commas) and non-numeric
        $cleaned = $Value -replace '[^\d.]',''
    }
    if ($cleaned) { return [double]$cleaned }
    return 0
}

function Import-CsvStripBom {
    <#
    .SYNOPSIS
        Imports a CSV file, stripping any UTF-8 BOM that Microsoft Graph prepends.
        On PS 5.1, Import-Csv doesn't reliably strip the BOM, corrupting the first
        column header (e.g. "ïUser Principal Name" or invisible U+FEFF prefix).
    #>
    param ([string]$Path)
    $raw = [System.IO.File]::ReadAllText($Path)
    if ($raw.Length -gt 0 -and $raw[0] -eq [char]0xFEFF) {
        $raw = $raw.Substring(1)
    }
    return @($raw | ConvertFrom-Csv)
}

function Download-GraphReport {
    <#
    .SYNOPSIS
        Downloads a Graph usage report CSV to a temp file and returns imported data.
    #>
    param (
        [string]$ReportName,
        [string]$Period
    )
    $tempFile = Join-Path $env:TEMP "$ReportName`_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
    try {
        $uri = "https://graph.microsoft.com/v1.0/reports/$ReportName(period='$Period')"
        Invoke-GraphWithRetry -Method GET -Uri $uri -OutputFilePath $tempFile
        $data = Import-CsvStripBom -Path $tempFile
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        Write-Host "    $ReportName : $($data.Count) rows" -ForegroundColor DarkGreen
        return $data
    } catch {
        Write-Log "Failed to download $ReportName" -Level ERROR -ErrorRecord $_
        Write-Warning "    Failed to download $ReportName : $($_.Exception.Message)"
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        return $null
    }
}

function Download-GraphReportNoPeriod {
    <#
    .SYNOPSIS
        Downloads a Graph usage report that doesn't take a period parameter.
    #>
    param ([string]$ReportName)
    $tempFile = Join-Path $env:TEMP "$ReportName`_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
    try {
        $uri = "https://graph.microsoft.com/v1.0/reports/$ReportName"
        Invoke-GraphWithRetry -Method GET -Uri $uri -OutputFilePath $tempFile
        $data = Import-CsvStripBom -Path $tempFile
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        Write-Host "    $ReportName : $($data.Count) rows" -ForegroundColor DarkGreen
        return $data
    } catch {
        Write-Log "Failed to download $ReportName" -Level ERROR -ErrorRecord $_
        Write-Warning "    Failed to download $ReportName : $($_.Exception.Message)"
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        return $null
    }
}

# ── SKU Part Number → Friendly Name mapping ──
# Minimal fallback — full set loaded from M365SkuData.json below.
$skuFriendlyNames = @{
    "SPE_E3" = "Microsoft 365 E3"; "SPE_E5" = "Microsoft 365 E5"; "SPE_F1" = "Microsoft 365 F3"
    "SPB" = "Microsoft 365 Business Premium"; "ENTERPRISEPACK" = "Office 365 E3"
    "OFFICESUBSCRIPTION" = "Microsoft 365 Apps for Enterprise"
}

# Helper: resolve SKU part number to friendly name
function Resolve-SkuFriendlyName {
    param([string]$SkuPartNumber)
    # Strip zero-width characters (U+200B, U+FEFF) that Microsoft Graph sometimes appends to CPC/W365 SKU IDs
    $clean = $SkuPartNumber -replace '[\u200B\uFEFF]', ''
    if ($skuFriendlyNames.ContainsKey($clean)) { return $skuFriendlyNames[$clean] }
    if ($skuFriendlyNames.ContainsKey($SkuPartNumber)) { return $skuFriendlyNames[$SkuPartNumber] }
    return $clean
}

# Guard $PSScriptRoot — empty when dot-sourced, run from ISE, or invoked via ScriptBlock
$_scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

# ── Monthly EUR list prices per SKU (loaded from M365SkuPricing.csv) ──
$skuMonthlyPrices = @{}
$pricingCsvDefault = Join-Path $_scriptRoot "M365SkuPricing.csv"
$_pricingCsvFile   = if ($PricingCsvPath) { $PricingCsvPath } else { $pricingCsvDefault }
if (Test-Path $_pricingCsvFile) {
    $csvPrices = Import-Csv $_pricingCsvFile
    foreach ($row in $csvPrices) {
        if ($row.SkuPartNumber -and $row.MonthlyPriceEUR) {
            $skuMonthlyPrices[$row.SkuPartNumber] = [decimal]$row.MonthlyPriceEUR
        }
    }
    Write-Host "  Loaded $(@($csvPrices).Count) SKU price(s) from $_pricingCsvFile" -ForegroundColor Green
} else {
    Write-Warning "SKU pricing CSV not found: $_pricingCsvFile — all SKUs will default to EUR 0.00. Place M365SkuPricing.csv alongside the script or use -PricingCsvPath."
}

# ── Override SKU data from external JSON file (M365SkuData.json) ──
$skuJsonPath = if ($SkuDataPath) { $SkuDataPath } else { Join-Path $_scriptRoot "M365SkuData.json" }
$skuDataDate = $null
$skuDataAge  = -1
$skuDataLoaded = $false
if (Test-Path $skuJsonPath) {
    try {
        $jsonRaw  = Get-Content $skuJsonPath -Raw -ErrorAction Stop
        $jsonData = $jsonRaw | ConvertFrom-Json
        # Overwrite friendly names
        if ($jsonData.PSObject.Properties['skuFriendlyNames'] -and $jsonData.skuFriendlyNames) {
            foreach ($prop in $jsonData.skuFriendlyNames.PSObject.Properties) {
                $skuFriendlyNames[$prop.Name] = $prop.Value
            }
        }
        # Overwrite monthly prices (property may not exist in JSON)
        if ($jsonData.PSObject.Properties['skuMonthlyPricesEUR']) {
            foreach ($prop in $jsonData.skuMonthlyPricesEUR.PSObject.Properties) {
                $skuMonthlyPrices[$prop.Name] = [decimal]$prop.Value
            }
        }
        $skuDataLoaded = $true
        # ── Staleness check ──
        $metaProp = $jsonData.PSObject.Properties['_meta']
        if ($metaProp -and $metaProp.Value -match 'Last updated:\s*(?<skuDate>\d{4}-\d{2}-\d{2})') {
            $skuDataDate = [datetime]::ParseExact($Matches['skuDate'], 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
            $skuDataAge  = [int](New-TimeSpan -Start $skuDataDate -End (Get-Date)).TotalDays
            if ($skuDataAge -gt $SkuStalenessDays) {
                Write-Warning "M365SkuData.json is $skuDataAge days old (threshold: ${SkuStalenessDays}d). Re-run _extract_sku_names.ps1 to refresh from Microsoft's licensing reference."
            }
        }
        $skuNameCount  = if ($jsonData.PSObject.Properties['skuFriendlyNames']) { @($jsonData.skuFriendlyNames.PSObject.Properties).Count } else { 0 }
        $skuPriceCount = if ($jsonData.PSObject.Properties['skuMonthlyPricesEUR']) { @($jsonData.skuMonthlyPricesEUR.PSObject.Properties).Count } else { 0 }
        Write-Host "  Loaded SKU naming/pricing from $skuJsonPath ($skuNameCount names, $skuPriceCount prices)" -ForegroundColor Green
    } catch {
        Write-Log "Failed to parse SKU JSON ($skuJsonPath)" -Level ERROR -ErrorRecord $_
        Write-Warning "Failed to parse SKU JSON ($skuJsonPath): $($_.Exception.Message) — using CSV pricing only."
    }
} elseif ($SkuDataPath) {
    Write-Warning "SKU data file not found: $SkuDataPath — using CSV pricing only."
}

# ── ForceSkuRefresh gate: abort if external data is missing or stale ──
if ($ForceSkuRefresh) {
    $abortReason = $null
    if (-not (Test-Path $skuJsonPath)) {
        $abortReason = "M365SkuData.json not found at $skuJsonPath."
    } elseif (-not $skuDataLoaded) {
        $abortReason = "M365SkuData.json failed to load (see warning above)."
    } elseif ($skuDataAge -lt 0) {
        $abortReason = "M365SkuData.json has no 'Last updated' date in _meta — cannot verify freshness."
    } elseif ($skuDataAge -gt $SkuStalenessDays) {
        $abortReason = "SKU data is ${skuDataAge}d old (threshold: ${SkuStalenessDays}d)."
    }
    if ($abortReason) {
        throw "ABORT (-ForceSkuRefresh): $abortReason Re-run _extract_sku_names.ps1 to refresh, or remove -ForceSkuRefresh to use built-in fallbacks."
    }
}

# ── Load LOA Rule Pack (optional) ──
$rulePackData  = $null
$manualRules   = @()
$rulePackDocs  = @{}
$rulePackJsonPath = if ($RulePackPath) { $RulePackPath } else { Join-Path $_scriptRoot "LOA_RulePack_M365.json" }
if (Test-Path $rulePackJsonPath) {
    try {
        $rulePackRaw  = Get-Content $rulePackJsonPath -Raw -ErrorAction Stop
        $rulePackData = $rulePackRaw | ConvertFrom-Json
        $manualRules  = @($rulePackData.rules | Where-Object { $_.automationLevel -eq "MANUAL" })
        if ($rulePackData.docs) {
            foreach ($doc in $rulePackData.docs) { $rulePackDocs[$doc.id] = $doc }
        }
        Write-Host "  LOA Rule Pack loaded: $($rulePackData.rules.Count) rules ($($manualRules.Count) manual audit items)." -ForegroundColor Green
    } catch {
        Write-Warning "  Could not load LOA Rule Pack: $($_.Exception.Message)"
        $rulePackData = $null
        $manualRules  = @()
    }
}

function Get-SkuMonthlyPrice {
    param([string]$SkuPartNumber)
    if ($skuMonthlyPrices.ContainsKey($SkuPartNumber)) { return [decimal]$skuMonthlyPrices[$SkuPartNumber] }
    # Strip zero-width characters (U+200B, U+FEFF) — Microsoft Graph sometimes appends to CPC/W365 SKU IDs
    $clean = $SkuPartNumber -replace '[\u200B\uFEFF]', ''
    if ($clean -ne $SkuPartNumber -and $skuMonthlyPrices.ContainsKey($clean)) { return [decimal]$skuMonthlyPrices[$clean] }
    return [decimal]0.00
}

# Helper: check if a SKU is in our reference data (friendly name or pricing)
function Test-SkuKnown {
    param([string]$SkuPartNumber)
    $clean = $SkuPartNumber -replace '[\u200B\uFEFF]', ''
    return ($skuFriendlyNames.ContainsKey($SkuPartNumber) -or $skuFriendlyNames.ContainsKey($clean) -or
            $skuMonthlyPrices.ContainsKey($SkuPartNumber) -or $skuMonthlyPrices.ContainsKey($clean))
}

# $suiteIncludes: Suite-to-component mapping for duplicate detection -- full set loaded from M365SkuData.json
# Reference: https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference
$suiteIncludes = @{
    "SPE_E3" = @("EXCHANGESTANDARD","EXCHANGEENTERPRISE","EXCHANGE_ARCHIVE","SHAREPOINTSTANDARD","SHAREPOINTENTERPRISE","MCOSTANDARD","OFFICESUBSCRIPTION","INTUNE_A","AAD_PREMIUM","RIGHTSMANAGEMENT","ATP_ENTERPRISE","MDE_LITE","FLOW_FREE","POWERAPPS_VIRAL","STREAM","TEAMS1","TEAMS_EXPLORATORY")
    "SPE_E5" = @("EXCHANGESTANDARD","EXCHANGEENTERPRISE","EXCHANGE_ARCHIVE","SHAREPOINTSTANDARD","SHAREPOINTENTERPRISE","MCOSTANDARD","OFFICESUBSCRIPTION","INTUNE_A","AAD_PREMIUM","AAD_PREMIUM_P2","EMSPREMIUM","RIGHTSMANAGEMENT","ATP_ENTERPRISE","THREAT_INTELLIGENCE","MCOEV","MCOMEETADV","INFORMATION_PROTECTION_COMPLIANCE","POWER_BI_PRO","ATA","ADALLOM_S_STANDALONE","WIN_DEF_ATP","FLOW_FREE","POWERAPPS_VIRAL","STREAM","TEAMS1","TEAMS_EXPLORATORY")
    "SPB" = @("EXCHANGESTANDARD","EXCHANGE_ARCHIVE","SHAREPOINTSTANDARD","MCOSTANDARD","O365_BUSINESS","INTUNE_A","AAD_PREMIUM","ATP_ENTERPRISE","MDE_SMB","RIGHTSMANAGEMENT")
    "O365_BUSINESS_ESSENTIALS" = @("EXCHANGESTANDARD","SHAREPOINTSTANDARD","MCOSTANDARD","TEAMS1")
    "SMB_BUSINESS_ESSENTIALS"  = @("EXCHANGESTANDARD","SHAREPOINTSTANDARD","MCOSTANDARD","TEAMS1")
    "M365_BUSINESS_BASIC"      = @("EXCHANGESTANDARD","SHAREPOINTSTANDARD","MCOSTANDARD","TEAMS1")
    "O365_BUSINESS_PREMIUM"    = @("EXCHANGESTANDARD","SHAREPOINTSTANDARD","MCOSTANDARD","O365_BUSINESS","INTUNE_A","AAD_PREMIUM","TEAMS1")
    "M365_BUSINESS_STANDARD"   = @("EXCHANGESTANDARD","SHAREPOINTSTANDARD","MCOSTANDARD","O365_BUSINESS","TEAMS1")
}

# $addOnBundles: Security/compliance add-on bundles (NOT productivity suites) -- full set loaded from M365SkuData.json
$addOnBundles = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@(“EMS”,”EMSPREMIUM”,
                “IDENTITY_THREAT_PROTECTION”,”IDENTITY_THREAT_PROTECTION_FOR_EMS_E5”,
                “SPE_F5_SEC”,”SPE_F5_SECCOMP”,”M365_SECURITY_COMPLIANCE_FOR_FLW”,
                “DEFENDER_SUITE_FLW”,”PURVIEW_SUITE_FLW”,
                “M365_DEFENDER_SUITE_BUSINESS”,”M365_PURVIEW_SUITE_BUSINESS”,
                “ENTRA_SUITE”,”INTUNE_SUITE”),
    [StringComparer]::OrdinalIgnoreCase)

# $skuCoverageAliases: Canonical SKU resolution for duplicate detection -- full set loaded from M365SkuData.json
$skuCoverageAliases = @{ "ADALLOM_STANDALONE"="ADALLOM_S_STANDALONE"; "DEFENDER_ENDPOINT_P1"="MDE_LITE"; "DEFENDER_ENDPOINT_P2"="WIN_DEF_ATP"; "DEFENDER_BUSINESS"="MDE_SMB"; "MDE_SMB"="WIN_DEF_ATP" }

# ── E3 + add-on → E5 upgrade mapping ──
# If a user has an E3 suite AND multiple of these add-ons, E5 may be cheaper.
# Includes E5-only standalone SKUs: MDO P2, Defender (Endpoint P2/Identity/CloudApps),
# Teams Phone, Audio Conf, PBI Pro, Entra P2, Purview Suite, and Defender Suite bundles.
# NOTE: ATP_ENTERPRISE (MDO P1) and MDE_LITE (MDE P1) are now in E3 — excluded from this list.
$e5AddOns = @("THREAT_INTELLIGENCE","MCOEV","MCOMEETADV",
              "AAD_PREMIUM_P2","INFORMATION_PROTECTION_COMPLIANCE","POWER_BI_PRO",
              # Defender standalone SKUs (E5-only — MDE P1/DEFENDER_ENDPOINT_P1 is now in E3, excluded)
              "WIN_DEF_ATP","DEFENDER_ENDPOINT_P2",
              "ATA","ADALLOM_STANDALONE",
              # Defender Suite (single SKU that wraps multiple Defender components)
              "IDENTITY_THREAT_PROTECTION","IDENTITY_THREAT_PROTECTION_FOR_EMS_E5")
$e3Suites = @("SPE_E3","ENTERPRISEPACK","MICROSOFT365_E3","M365EDU_A3_FACULTY","M365EDU_A3_STUDENT",
              # EEA no-Teams E3 variants
              "Microsoft_365_E3_(no_Teams)","O365_w/o Teams Bundle_M3",
              "O365_w/o_Teams_Bundle_E3","Office_365_E3_(no_Teams)")

# ── Expensive standalone SKUs to flag as shelfware when unused ──
$expensiveStandalone = @{
    "VISIOCLIENT"         = "Visio Plan 2"
    "VISIOONLINE_PLAN1"   = "Visio Plan 1"
    "PROJECTPROFESSIONAL" = "Planner and Project Plan 3"
    "PROJECTPREMIUM"      = "Planner and Project Plan 5"
    "PROJECTESSENTIALS"   = "Project Online Essentials"
    "POWER_BI_PRO"        = "Power BI Pro"
    "POWER_BI_PREMIUM_P"  = "Power BI Premium Per User"
    "Microsoft_Teams_Premium" = "Teams Premium"
}

# ── Teams Phone SKUs and Calling Plan SKUs (HashSets for O(1) lookup in hot loop) ──
$teamsPhoneSkus  = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@("MCOEV","MCOEV_DOD","MCOEV_GOV","PHONESYSTEM_VIRTUALUSER",
                "MCOTEAMS_ESSENTIALS","TEAMS_PHONE_STANDARD"),
    [System.StringComparer]::OrdinalIgnoreCase)
$callingPlanSkus = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@("MCOPSTN1","MCOPSTN2","MCOPSTN5","MCOPSTN_5","MCOPSTNC",
                "MCOPSTN_1_GOV","MCOPSTN_2_GOV"),
    [System.StringComparer]::OrdinalIgnoreCase)

# ── Copilot SKUs ──
# Copilot SKUs — separate productivity Copilot (requires E3/E5/Business Standard/Premium base) from Studio (admin/dev tool)
$copilotProductivitySkus = @("MICROSOFT_COPILOT","Microsoft_365_Copilot","SPE_E3_RPA1")
$copilotBusinessSkus     = @("Microsoft_365_Copilot_Business")
$copilotSecuritySkus     = @("MICROSOFT_SECURITY_COPILOT")
$copilotStudioSkus       = @("COPILOT_STUDIO")
# ── Business-family SKUs (300-seat cap) ──
$businessFamilySkus = @("O365_BUSINESS","O365_BUSINESS_ESSENTIALS","O365_BUSINESS_PREMIUM",
                        "SMB_BUSINESS","SMB_BUSINESS_ESSENTIALS","SMB_BUSINESS_PREMIUM",
                        "SPB","M365_BUSINESS_BASIC","M365_BUSINESS_STANDARD")

# $premiumSuites: E3/E5 suites for frontline right-sizing -- full set loaded from M365SkuData.json
# NOTE: Only enterprise/education suites -- NOT Business SKUs.
$premiumSuites = @("SPE_E3","SPE_E5","ENTERPRISEPACK","ENTERPRISEPREMIUM","MICROSOFT365_E3","MICROSOFT365_E5",
                    "SPE_E5_NOPSTNCONF","ENTERPRISEPREMIUM_NOPSTNCONF",
                    "M365EDU_A3_FACULTY","M365EDU_A3_STUDENT","M365EDU_A5_FACULTY","M365EDU_A5_STUDENT","M365EDU_A5_STUUSEBNFT")

# ── EXO Plan 2 SKUs ──
$exoPlan2Skus = @("EXCHANGEENTERPRISE","EXCHANGE_S_ENTERPRISE")

# ── Security & Compliance upsell mappings ──
# Business licenses WITHOUT any Defender/security features
# Ref: https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference
$businessNoSecurity = @("O365_BUSINESS_ESSENTIALS","SMB_BUSINESS_ESSENTIALS",
                        "O365_BUSINESS_PREMIUM","SMB_BUSINESS_PREMIUM",
                        "O365_BUSINESS","SMB_BUSINESS",
                        "M365_BUSINESS_BASIC","M365_BUSINESS_STANDARD",
                        # no-Teams variants (no Defender)
                        "Microsoft_365_Business_Basic_(no Teams)",
                        "Microsoft_365_Business_Basic_EEA_(no_Teams)",
                        "MICROSOFT_365_BUSINESS_STANDARD_NO_TEAMS",
                        "Microsoft_365_Business_Standard_EEA_(no_Teams)",
                        "Office_365_w/o_Teams_Bundle_Business_Standard")
# Business Premium (has basic Defender — eligible for Defender Suite upsell)
$businessPremiumSkus = @("SPB","Office_365_w/o_Teams_Bundle_Business_Premium",
                        "Microsoft_365_ Business_ Premium_(no Teams)")
# ── Specialist add-on category lists (for upsell suppression + informational flags) ──
# These lists let us suppress false "gap" signals and emit informational LICENSING CHECK notes
# without modeling each feature into planCapabilities (most don't affect workload-level reasoning).
$entraGovSkus   = @("ENTRA_ID_GOVERNANCE","ENTRA_ID_GOVERNANCE_P2",
                     "ENTRA_ID_GOVERNANCE_FLW","ENTRA_ID_GOVERNANCE_P2_FLW")
$entraSuiteSkus = @("ENTRA_SUITE","ENTRA_SUITE_FLW")
$intuneSuiteSkus = @("INTUNE_SUITE","INTUNE_REMOTE_HELP","INTUNE_ADVANCED_ANALYTICS",
                      "INTUNE_EPM","INTUNE_CLOUD_PKI")
$audit10YearSkus = @("COMPLIANCE_AUDIT_10YEAR","M365_COMPLIANCE_AUDIT_10YEAR")
$privaSkus       = @("PRIVA_RISK_MANAGEMENT","PRIVA_SUBJECT_RIGHTS_REQUEST",
                      "PRIVACY_MANAGEMENT","PRIVACY_MANAGEMENT_RISK")
# FLW-specific Defender standalone SKUs — used to suppress "no endpoint protection" on Frontline plans
$flwDefenderSkus = @("DEFENDER_ENDPOINT_P1_FLW","DEFENDER_ENDPOINT_P2_FLW",
                      "MDO_P1_FLW","MDO_P2_FLW","DEFENDER_SUITE_FLW")
# Combined "advanced security/compliance add-on" set — used to suppress generic upsell noise
$advancedSecCompAddons = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]($entraGovSkus + $entraSuiteSkus + $intuneSuiteSkus + $audit10YearSkus + $privaSkus),
    [StringComparer]::OrdinalIgnoreCase)

# ── Plan Capabilities Matrix — full set loaded from M365SkuData.json ──
# Ref: https://go.microsoft.com/fwlink/?linkid=2139145 (Enterprise), /2139553 (SMB)
$planCapabilities = @{}

# ── Plan capability aliases — full set loaded from M365SkuData.json ──
# Minimal fallback: map common aliases to their canonical planCapabilities key.
$planCapabilityAliases = @{
    "MICROSOFT365_E3" = "SPE_E3"; "MICROSOFT365_E5" = "SPE_E5"; "DEVELOPERPACK_E5" = "SPE_E5"
    "SMB_BUSINESS_ESSENTIALS" = "O365_BUSINESS_ESSENTIALS"; "M365_BUSINESS_BASIC" = "O365_BUSINESS_ESSENTIALS"
    "M365_BUSINESS_STANDARD" = "O365_BUSINESS_PREMIUM"; "SMB_BUSINESS_PREMIUM" = "O365_BUSINESS_PREMIUM"
    "SMB_BUSINESS" = "O365_BUSINESS"
}

# ── Phase 2: Load licensing matrices from M365SkuData.json ──
# skuFriendlyNames + skuMonthlyPricesEUR were loaded in Phase 1 above (before these variables existed).
# Now load suiteIncludes, planCapabilities, planCapabilityAliases, addOnBundles, premiumSuites, skuCoverageAliases.
if ($skuDataLoaded -and $jsonData) {
    if ($jsonData.PSObject.Properties['suiteIncludes']) {
        foreach ($prop in $jsonData.suiteIncludes.PSObject.Properties) {
            $suiteIncludes[$prop.Name] = @($prop.Value)
        }
    }
    if ($jsonData.PSObject.Properties['planCapabilities']) {
        foreach ($prop in $jsonData.planCapabilities.PSObject.Properties) {
            $ht = @{}
            foreach ($cap in $prop.Value.PSObject.Properties) {
                $val = $cap.Value
                if ($cap.Name -eq 'MaxUsers' -and $val -eq 0) { $val = [int]::MaxValue }
                $ht[$cap.Name] = $val
            }
            $planCapabilities[$prop.Name] = $ht
        }
    }
    if ($jsonData.PSObject.Properties['planCapabilityAliases']) {
        foreach ($prop in $jsonData.planCapabilityAliases.PSObject.Properties) {
            $planCapabilityAliases[$prop.Name] = $prop.Value
        }
    }
    if ($jsonData.PSObject.Properties['addOnBundles']) {
        foreach ($item in $jsonData.addOnBundles) { [void]$addOnBundles.Add($item) }
    }
    if ($jsonData.PSObject.Properties['premiumSuites']) {
        # NOTE: Full replacement (not merge) — JSON is authoritative for this array.
        $premiumSuites = @($jsonData.premiumSuites)
    }
    if ($jsonData.PSObject.Properties['skuCoverageAliases']) {
        foreach ($prop in $jsonData.skuCoverageAliases.PSObject.Properties) {
            $skuCoverageAliases[$prop.Name] = $prop.Value
        }
    }
    Write-Host "  Loaded licensing matrices: $($suiteIncludes.Count) suites, $($planCapabilities.Count) capability profiles, $($planCapabilityAliases.Count) aliases" -ForegroundColor Green
} elseif (-not $skuDataLoaded) {
    Write-Warning "M365SkuData.json not loaded — using minimal inline fallbacks. Detection accuracy will be reduced."
}

# Helper: resolve SKU to its plan capabilities (returns $null if not a known suite)
function Get-PlanCapabilities {
    param([string]$SkuPartNumber)
    if ($planCapabilities.ContainsKey($SkuPartNumber)) { return $planCapabilities[$SkuPartNumber] }
    if ($planCapabilityAliases.ContainsKey($SkuPartNumber)) {
        return $planCapabilities[$planCapabilityAliases[$SkuPartNumber]]
    }
    return $null
}

# Helper: merge planCapabilities from all user SKUs into a single effective capability hashtable.
# Boolean fields use OR (any SKU providing the capability = true). EntraIdTier uses max (P2 > P1 > Free).
# This enables capability-driven upsell logic — instead of checking SKU names, check what the user
# effectively has across all their assignments combined.
function Merge-UserCapabilities {
    param([string[]]$SkuList)
    $merged = @{
        # Legacy fields (kept for backward compat — driven by the granular ones below)
        DefenderForO365 = $false; DefenderForEndpoint = $false
        DefenderForIdentity = $false; CloudAppSecurity = $false
        DLP = $false; AIPPlan1 = $false; AIPPlan2 = $false
        eDiscoveryStandard = $false; eDiscoveryPremium = $false
        AuditStandard = $false; AuditPremium = $false
        InsiderRisk = $false; IntunePlan1 = $false
        EntraIdTier = "Free"
        # ── Granular threat protection flags ──
        MdoP1 = $false; MdoP2 = $false           # Defender for Office 365 tiers
        MdeP1 = $false; MdeP2OrBusiness = $false  # Defender for Endpoint tiers (P2 or Business)
        Mdi = $false; MdcApps = $false; Xdr = $false  # Identity, Cloud Apps, XDR
        # ── Granular compliance flags ──
        DlpEmailFiles = $false; DlpTeams = $false; EndpointDlp = $false  # DLP scopes
    }
    $entraRank = @{ "Free" = 0; "P1" = 1; "P2" = 2 }
    foreach ($sku in $SkuList) {
        $caps = Get-PlanCapabilities $sku
        if (-not $caps) { continue }
        # OR-merge all boolean security/compliance fields (legacy + granular)
        foreach ($field in @("DefenderForO365","DefenderForEndpoint","DefenderForIdentity",
                             "CloudAppSecurity","DLP","AIPPlan1","AIPPlan2",
                             "eDiscoveryStandard","eDiscoveryPremium",
                             "AuditStandard","AuditPremium","InsiderRisk","IntunePlan1",
                             "MdoP1","MdoP2","MdeP1","MdeP2OrBusiness",
                             "Mdi","MdcApps","Xdr",
                             "DlpEmailFiles","DlpTeams","EndpointDlp")) {
            if ($caps.ContainsKey($field) -and $caps[$field] -eq $true) {
                $merged[$field] = $true
            }
        }
        # Max-merge EntraIdTier
        if ($caps.ContainsKey("EntraIdTier")) {
            $capRank = if ($entraRank.ContainsKey($caps["EntraIdTier"])) { $entraRank[$caps["EntraIdTier"]] } else { 0 }
            $curRank = if ($entraRank.ContainsKey($merged["EntraIdTier"])) { $entraRank[$merged["EntraIdTier"]] } else { 0 }
            if ($capRank -gt $curRank) { $merged["EntraIdTier"] = $caps["EntraIdTier"] }
        }
    }
    return $merged
}

# ── Business Standard SKUs (for downgrade-to-Basic detection) — includes EEA no-Teams variants ──
$businessStandardSkus = @("O365_BUSINESS_PREMIUM","SMB_BUSINESS_PREMIUM",
                          "M365_BUSINESS_STANDARD",
                          "Microsoft_365_Business_Standard_EEA_(no_Teams)",
                          "MICROSOFT_365_BUSINESS_STANDARD_NO_TEAMS",
                          "Office_365_w/o_Teams_Bundle_Business_Standard")

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — Prerequisites & Connection
# ═══════════════════════════════════════════════════════════════════════════════

Write-Log "Script started — PowerShell $($PSVersionTable.PSVersion), OS: $([System.Environment]::OSVersion.VersionString)"
Write-Log "Parameters: ReportPeriod=$ReportPeriod, SkipEXO=$SkipEXO, NoExcel=$NoExcel, OutputFolder=$OutputFolder"

$requiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Reports",
    "Microsoft.Graph.Identity.DirectoryManagement"
)

# Check for missing required modules BEFORE attempting any installs
$missingRequired = @($requiredModules | Where-Object { -not (Get-Module -ListAvailable -Name $_) })
if ($missingRequired.Count -gt 0) {
    if ($AutoInstallModules) {
        foreach ($mod in $missingRequired) {
            Write-Host "  Installing module $mod ..." -ForegroundColor Yellow
            Install-Module -Name $mod -Scope CurrentUser -Force -AllowClobber
        }
    } else {
        Write-Host "`nMissing required modules:" -ForegroundColor Red
        foreach ($mod in $missingRequired) {
            Write-Host "    Install-Module $mod -Scope CurrentUser" -ForegroundColor Yellow
        }
        Write-Host "`nInstall the modules above, or re-run with -AutoInstallModules to install automatically.`n" -ForegroundColor Red
        throw "Missing required modules: $($missingRequired -join ', '). Use -AutoInstallModules to install automatically."
    }
}
foreach ($mod in $requiredModules) {
    Import-Module $mod -ErrorAction Stop
}

# EXO module is optional — used for mailbox type detection (Shared/Room/Equipment)
$exoAvailable = $false
if ($SkipEXO) {
    Write-Host "  -SkipEXO specified — Exchange Online connection will be skipped." -ForegroundColor Yellow
    Write-Host "  Mailbox type, litigation hold, and MDO policy coverage will NOT be available." -ForegroundColor Yellow
} elseif (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
    Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
    $exoAvailable = $true
} else {
    Write-Host "  ExchangeOnlineManagement module not found — mailbox type detection will be skipped." -ForegroundColor Yellow
    Write-Host "  Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser" -ForegroundColor Yellow
}

# ImportExcel module — optional, used for Excel workbook output
$importExcelAvailable = $false
if (-not $NoExcel) {
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel -ErrorAction SilentlyContinue
        $importExcelAvailable = $true
    } else {
        Write-Host "  ImportExcel module not found — Excel output will be skipped (CSVs still produced)." -ForegroundColor Yellow
        Write-Host "  Install with: Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Yellow
    }
} else {
    Write-Host "  Excel output disabled via -NoExcel switch." -ForegroundColor DarkGray
}

# ── Auto-detect LOA-Connection.json if no auth parameters provided ──
if (-not $ClientId -and -not $TenantId -and -not $CertificateThumbprint -and -not $CertificatePath) {
    $configPaths = @(
        (Join-Path $_scriptRoot "LOA-Connection.json"),
        (Join-Path (Get-Location).Path "LOA-Connection.json"),
        (Join-Path $_scriptRoot "M365-LOA-Audit-Package" "LOA-Connection.json"),
        (Join-Path (Get-Location).Path "M365-LOA-Audit-Package" "LOA-Connection.json")
    )
    foreach ($cfgPath in $configPaths) {
        if (Test-Path $cfgPath) {
            try {
                $cfg = Get-Content $cfgPath -Raw | ConvertFrom-Json
                $ClientId              = $cfg.ClientId
                $TenantId              = $cfg.TenantId
                $CertificateThumbprint = $cfg.CertificateThumbprint
                Write-Host "  Auto-detected connection config: $cfgPath" -ForegroundColor Green
                Write-Host "    Tenant: $TenantId  App: $ClientId" -ForegroundColor Gray
                break
            } catch {
                Write-Warning "  LOA-Connection.json found at $cfgPath but failed to parse: $($_.Exception.Message)"
                Write-Warning "  Falling back to interactive login."
            }
        }
    }
}

# ── Validate certificate-auth parameter combinations ──
$useCertAuth = $false
if ($ClientId -or $TenantId -or $CertificateThumbprint -or $CertificatePath) {
    if (-not $ClientId -or -not $TenantId) {
        throw "Certificate-based auth requires both -ClientId and -TenantId."
    }
    if (-not $CertificateThumbprint -and -not $CertificatePath) {
        throw "Certificate-based auth requires either -CertificateThumbprint or -CertificatePath."
    }
    $useCertAuth = $true

    # If a .pfx path was provided, import the certificate to CurrentUser\My
    if ($CertificatePath) {
        if (-not $CertificatePassword) {
            $CertificatePassword = Read-Host -Prompt "Enter certificate password" -AsSecureString
        }
        $importedCert = Import-PfxCertificate -FilePath $CertificatePath `
            -CertStoreLocation Cert:\CurrentUser\My -Password $CertificatePassword -ErrorAction Stop
        $CertificateThumbprint = $importedCert.Thumbprint
        Write-Host "  Certificate imported (thumbprint: $CertificateThumbprint)" -ForegroundColor Green
    }
}

Write-Host "`n[1/12] Connecting to Microsoft Graph ..." -ForegroundColor Cyan
Write-Log "[1/12] Connecting to Microsoft Graph"

if ($useCertAuth) {
    # App-only (certificate) auth — used when the auditor has an App Registration
    Connect-MgGraph -ClientId $ClientId -TenantId $TenantId `
                    -CertificateThumbprint $CertificateThumbprint -NoWelcome
    $ctx = Get-MgContext
    Write-Host "  Connected via certificate auth  App: $ClientId  Tenant: $($ctx.TenantId)" -ForegroundColor Green
} else {
    # Interactive delegated auth — fallback for ad-hoc runs
    # NOTE: CloudLicensing.Read.All may fail consent on some tenants — the script degrades gracefully (try/catch).
    $scopes = @("User.Read.All", "Reports.Read.All", "AuditLog.Read.All", "Policy.Read.All", "RoleManagement.Read.Directory", "Group.Read.All", "DeviceManagementManagedDevices.Read.All", "CloudLicensing.Read.All")
    $scopes += "Organization.Read.All"
    if (-not $KeepHashedUPNs) { $scopes += "ReportSettings.ReadWrite.All" }
    Connect-MgGraph -Scopes $scopes -NoWelcome
    $ctx = Get-MgContext
    Write-Host "  Connected as: $($ctx.Account)  Tenant: $($ctx.TenantId)" -ForegroundColor Green
}

$exoConnected = $false
if ($exoAvailable) {
    Write-Host "  Connecting to Exchange Online ..." -ForegroundColor Cyan
    try {
        if ($useCertAuth) {
            # Certificate-based EXO connection — requires Exchange.ManageAsApp + RBAC roles
            $orgDomain = @((Get-MgOrganization).VerifiedDomains) | Where-Object { $_.IsInitial -eq $true } | Select-Object -First 1 -ExpandProperty Name
            if (-not $orgDomain) { throw "Could not determine initial domain for EXO certificate auth — VerifiedDomains returned no initial domain." }
            Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint `
                                   -Organization $orgDomain -ShowBanner:$false
        } else {
            # Interactive EXO connection (WAM enabled — never disable)
            Connect-ExchangeOnline -ShowBanner:$false
        }
        $exoConnected = $true
        Write-Host "  Exchange Online connected." -ForegroundColor Green
    } catch {
        Write-Log "Exchange Online connection failed" -Level ERROR -ErrorRecord $_
        Write-Warning "  Could not connect to Exchange Online: $($_.Exception.Message)"
        Write-Warning "  Mailbox type detection will be skipped."
        if ($useCertAuth) {
            Write-Warning "  TIP: Ensure Exchange.ManageAsApp is consented and RBAC roles are assigned."
        } else {
            Write-Warning "  TIP: Run 'Update-Module ExchangeOnlineManagement -Force' then retry."
        }
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — Unhide user data in reports (default; skip with -KeepHashedUPNs)
# ═══════════════════════════════════════════════════════════════════════════════

$privacyChanged = $false
if (-not $KeepHashedUPNs) {
    Write-Host "`n[*] Checking report privacy setting ..." -ForegroundColor Yellow
    try {
        $settings = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/admin/reportSettings"
        if ($settings['displayConcealedNames'] -eq $true) {
            if ($PSCmdlet.ShouldProcess("Tenant report privacy setting", "Set displayConcealedNames to false")) {
                Write-Host "  Unhiding user data in reports ..." -ForegroundColor Yellow
                $body = @{ displayConcealedNames = $false } | ConvertTo-Json
                Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/admin/reportSettings" `
                    -Body $body -ContentType "application/json"
                $privacyChanged = $true
                # Allow a few seconds for the setting to propagate
                Start-Sleep -Seconds 5
            } else {
                Write-Host "  Skipped — privacy setting not changed (user declined or -WhatIf)." -ForegroundColor DarkGray
            }
        } else {
            Write-Host "  User data already visible." -ForegroundColor Green
        }
    } catch {
        Write-Log "Could not update privacy setting" -Level ERROR -ErrorRecord $_
        Write-Warning "Could not update privacy setting: $($_.Exception.Message)"
    }
}

# ─── Main execution wrapped in try/finally to guarantee cleanup on crash/Ctrl+C ───
try {

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — Download all usage reports
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host "`n[2/12] Downloading usage reports (period=$ReportPeriod) ..." -ForegroundColor Cyan
Write-Log "[2/12] Downloading usage reports (period=$ReportPeriod)"

# ── Attempt parallel downloads via runspace pool (PS 5.1 + 7+ compatible) ──
$parallelSuccess = $false
$dlStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

try {
    # Extract bearer token from MgGraph session for direct REST calls
    $graphToken = $null
    try {
        $tokenResp  = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -OutputType HttpResponseMessage
        $graphToken = $tokenResp.RequestMessage.Headers.Authorization.Parameter
    } catch { $graphToken = $null }

    if ($graphToken) {
        Write-Host "  Downloading 11 reports ($MaxParallel concurrent) ..." -ForegroundColor DarkGreen

        # Report definitions: [VarName, ReportName, HasPeriod]
        # NOTE: Use comma-prefix ,@(...) to prevent @() from flattening inner arrays
        $reportDefs = @(
            ,@("activeUserDetail",   "getOffice365ActiveUserDetail",   $true)
            ,@("m365AppDetail",      "getM365AppUserDetail",           $true)
            ,@("emailActivity",      "getEmailActivityUserDetail",     $true)
            ,@("teamsActivity",      "getTeamsUserActivityUserDetail",  $true)
            ,@("oneDriveActivity",   "getOneDriveActivityUserDetail",  $true)
            ,@("sharePointActivity", "getSharePointActivityUserDetail", $true)
            ,@("mailboxUsage",       "getMailboxUsageDetail",           $true)
            ,@("emailAppUsage",      "getEmailAppUsageUserDetail",      $true)
            ,@("oneDriveUsage",      "getOneDriveUsageAccountDetail",   $true)
            ,@("teamsDeviceUsage",   "getTeamsDeviceUsageUserDetail",   $true)
            ,@("activations",        "getOffice365ActivationsUserDetail", $false)
        )

        $downloadBlock = {
            param([string]$ReportName, [string]$Period, [bool]$HasPeriod, [string]$Token)
            $uri = if ($HasPeriod) {
                "https://graph.microsoft.com/v1.0/reports/$ReportName(period='$Period')"
            } else {
                "https://graph.microsoft.com/v1.0/reports/$ReportName"
            }
            $headers = @{ Authorization = "Bearer $Token" }
            # Lightweight retry loop: handles 429 (throttled) and transient 5xx inside
            # the runspace so we don't fall back to the slower sequential path.
            $maxAttempts = 3
            $baseDelay   = 2
            for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
                try {
                    # Invoke-RestMethod follows 302 redirects and returns CSV body as string.
                    # Works reliably in runspaces (Invoke-WebRequest requires full HTTP response
                    # infrastructure that may not function in minimal runspace state on PS 7).
                    $csvText = Invoke-RestMethod -Uri $uri -Headers $headers -Method GET -ErrorAction Stop
                    if (-not $csvText -or $csvText.Length -lt 10) { return }
                    # Strip UTF-8 BOM if present — Graph prepends it and PS 5.1 doesn't strip from strings
                    if ($csvText[0] -eq [char]0xFEFF) { $csvText = $csvText.Substring(1) }
                    return ($csvText | ConvertFrom-Csv)
                } catch {
                    $statusCode = 0
                    if ($_.Exception.Response) {
                        try { $statusCode = [int]$_.Exception.Response.StatusCode } catch { }
                    }
                    $isRetryable = ($statusCode -eq 429) -or ($statusCode -ge 500 -and $statusCode -lt 600)
                    if ($isRetryable -and $attempt -lt $maxAttempts) {
                        # Respect Retry-After header if present, otherwise exponential backoff
                        $delay = $baseDelay * [Math]::Pow(2, $attempt - 1)
                        if ($statusCode -eq 429 -and $_.Exception.Response.Headers) {
                            try {
                                $raHeader = $_.Exception.Response.Headers | Where-Object { $_.Key -eq 'Retry-After' }
                                if ($raHeader.Value) {
                                    $parsed = [int]($raHeader.Value | Select-Object -First 1)
                                    if ($parsed -gt 0) { $delay = $parsed }
                                }
                            } catch { }
                        }
                        $delay = [Math]::Min($delay, 60)
                        Start-Sleep -Seconds $delay
                        continue
                    }
                    # Non-retryable or exhausted attempts — signal failure
                    $safeMsg = "$($_.Exception.Message)" -replace '(Bearer\s+)[A-Za-z0-9\-_\.]{20,}','$1[REDACTED]' -replace '(eyJ[A-Za-z0-9\-_]{10,}\.[A-Za-z0-9\-_]+\.[A-Za-z0-9\-_]*)','[REDACTED-JWT]'
                    Write-Error "[$ReportName] $safeMsg (after $attempt attempt(s))"
                    $null
                }
            }
        }

        # Create runspace pool and queue all downloads
        $pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $MaxParallel)
        $pool.Open()
        try {
            $jobs = [System.Collections.Generic.List[hashtable]]::new()
            foreach ($def in $reportDefs) {
                $ps = [System.Management.Automation.PowerShell]::Create().AddScript($downloadBlock)
                $ps.AddParameter("ReportName", $def[1]) | Out-Null
                $ps.AddParameter("Period",     $ReportPeriod) | Out-Null
                $ps.AddParameter("HasPeriod",  $def[2]) | Out-Null
                $ps.AddParameter("Token",      $graphToken) | Out-Null
                $ps.RunspacePool = $pool
                $jobs.Add(@{
                    VarName  = $def[0]
                    Instance = $ps
                    Handle   = $ps.BeginInvoke()
                })
            }

            # Collect results from runspaces.
            # IMPORTANT: In PS 5.1, AddScript() runspaces can wrap pipeline output as a single
            # array element inside the PSDataCollection.  E.g., ConvertFrom-Csv outputs 15 PSObjects,
            # but EndInvoke returns a PSDataCollection with Count=1 containing one Object[15].
            # We must flatten any nested arrays to get the actual CSV rows.
            # A scriptblock returning $null produces an empty PSDataCollection (Count=0).
            $reportResults = @{}
            foreach ($job in $jobs) {
                try {
                    $result = $job.Instance.EndInvoke($job.Handle)
                    # Flatten: enumerate PSDataCollection, unwrap any nested arrays
                    $flat = [System.Collections.Generic.List[object]]::new()
                    foreach ($item in $result) {
                        if ($item -is [System.Array]) {
                            foreach ($sub in $item) { [void]$flat.Add($sub) }
                        } else {
                            [void]$flat.Add($item)
                        }
                    }
                    $reportResults[$job.VarName] = if ($flat.Count -gt 0) { $flat.ToArray() } else { $null }
                    # Log any errors from the runspace's error stream (diagnostic for download failures)
                    $rsErrors = $job.Instance.Streams.Error
                    if ($rsErrors -and $rsErrors.Count -gt 0) {
                        Write-Log "Parallel download error for $($job.VarName): $($rsErrors[0].Exception.Message)" -Level WARN
                    }
                } catch {
                    # EndInvoke can throw if the runspace encountered unrecoverable errors
                    Write-Log "Runspace error for $($job.VarName): $($_.Exception.Message)" -Level WARN
                    $reportResults[$job.VarName] = $null
                }
                $job.Instance.Dispose()
            }
        } finally {
            $pool.Close()
            $pool.Dispose()
        }

        # Retry failed reports sequentially (Invoke-GraphWithRetry has built-in exponential backoff)
        $failedReports = @($reportResults.Keys | Where-Object { $null -eq $reportResults[$_] })
        if ($failedReports.Count -gt 0) {
            Write-Host "    $($failedReports.Count) report(s) failed — retrying sequentially ..." -ForegroundColor Yellow
            Write-Log "$($failedReports.Count) parallel report(s) failed, retrying: $($failedReports -join ', ')" -Level WARN
            $defByVar = @{}
            foreach ($def in $reportDefs) { $defByVar[$def[0]] = $def }

            foreach ($varName in $failedReports) {
                $def = $defByVar[$varName]
                $retryResult = if ($def[2]) {
                    Download-GraphReport -ReportName $def[1] -Period $ReportPeriod
                } else {
                    Download-GraphReportNoPeriod -ReportName $def[1]
                }
                if ($null -ne $retryResult -and $retryResult.Count -gt 0) {
                    $reportResults[$varName] = $retryResult
                    Write-Log "  Retry succeeded for $varName : $($retryResult.Count) rows"
                } else {
                    $reportResults[$varName] = @()   # exhausted retries — treat as empty
                    Write-Log "  Retry also failed for $varName — report data will be missing" -Level WARN
                    [void]$script:skippedDataWarnings.Add("$($def[1]) report — download failed after retry")
                }
            }
        }

        # Assign to script-scope variables
        # CAUTION: @($null) → array with one $null element, NOT empty array!
        # Use ?? pattern: if value is $null, assign @() instead.
        $activeUserDetail   = if ($reportResults["activeUserDetail"])   { $reportResults["activeUserDetail"]   } else { @() }
        $m365AppDetail      = if ($reportResults["m365AppDetail"])      { $reportResults["m365AppDetail"]      } else { @() }
        $emailActivity      = if ($reportResults["emailActivity"])      { $reportResults["emailActivity"]      } else { @() }
        $teamsActivity      = if ($reportResults["teamsActivity"])      { $reportResults["teamsActivity"]      } else { @() }
        $oneDriveActivity   = if ($reportResults["oneDriveActivity"])   { $reportResults["oneDriveActivity"]   } else { @() }
        $sharePointActivity = if ($reportResults["sharePointActivity"]) { $reportResults["sharePointActivity"] } else { @() }
        $mailboxUsage       = if ($reportResults["mailboxUsage"])       { $reportResults["mailboxUsage"]       } else { @() }
        $emailAppUsage      = if ($reportResults["emailAppUsage"])      { $reportResults["emailAppUsage"]      } else { @() }
        $oneDriveUsage      = if ($reportResults["oneDriveUsage"])      { $reportResults["oneDriveUsage"]      } else { @() }
        $teamsDeviceUsage   = if ($reportResults["teamsDeviceUsage"])   { $reportResults["teamsDeviceUsage"]   } else { @() }
        $activations        = if ($reportResults["activations"])        { $reportResults["activations"]        } else { @() }

        # Display row counts — skip retried reports (Download-GraphReport already printed them)
        foreach ($def in $reportDefs) {
            if ($failedReports -contains $def[0]) { continue }
            [int]$count = @($reportResults[$def[0]]).Count
            Write-Host "    $($def[1]) : $count rows" -ForegroundColor DarkGreen
        }

        $parallelSuccess = $true
    }
} catch {
    Write-Log "Parallel download setup failed" -Level WARN -ErrorRecord $_
    Write-Host "  Parallel download setup failed: $($_.Exception.Message)" -ForegroundColor Yellow
}

if (-not $parallelSuccess) {
    # ── Fallback: sequential downloads ──
    Write-Host "  Falling back to sequential downloads ..." -ForegroundColor Yellow
    $activeUserDetail   = Download-GraphReport -ReportName "getOffice365ActiveUserDetail"   -Period $ReportPeriod
    $m365AppDetail      = Download-GraphReport -ReportName "getM365AppUserDetail"           -Period $ReportPeriod
    $emailActivity      = Download-GraphReport -ReportName "getEmailActivityUserDetail"     -Period $ReportPeriod
    $teamsActivity      = Download-GraphReport -ReportName "getTeamsUserActivityUserDetail"  -Period $ReportPeriod
    $oneDriveActivity   = Download-GraphReport -ReportName "getOneDriveActivityUserDetail"  -Period $ReportPeriod
    $sharePointActivity = Download-GraphReport -ReportName "getSharePointActivityUserDetail" -Period $ReportPeriod
    $mailboxUsage       = Download-GraphReport -ReportName "getMailboxUsageDetail"           -Period $ReportPeriod
    $emailAppUsage      = Download-GraphReport -ReportName "getEmailAppUsageUserDetail"      -Period $ReportPeriod
    $oneDriveUsage      = Download-GraphReport -ReportName "getOneDriveUsageAccountDetail"   -Period $ReportPeriod
    $teamsDeviceUsage   = Download-GraphReport -ReportName "getTeamsDeviceUsageUserDetail"   -Period $ReportPeriod
    $activations        = Download-GraphReportNoPeriod -ReportName "getOffice365ActivationsUserDetail"

    # Check for reports that failed even with Invoke-GraphWithRetry backoff
    $seqReportCheck = @{
        'getOffice365ActiveUserDetail'        = $activeUserDetail
        'getM365AppUserDetail'                = $m365AppDetail
        'getEmailActivityUserDetail'          = $emailActivity
        'getTeamsUserActivityUserDetail'      = $teamsActivity
        'getOneDriveActivityUserDetail'       = $oneDriveActivity
        'getSharePointActivityUserDetail'     = $sharePointActivity
        'getMailboxUsageDetail'               = $mailboxUsage
        'getEmailAppUsageUserDetail'          = $emailAppUsage
        'getOneDriveUsageAccountDetail'       = $oneDriveUsage
        'getTeamsDeviceUsageUserDetail'       = $teamsDeviceUsage
        'getOffice365ActivationsUserDetail'   = $activations
    }
    foreach ($rptName in $seqReportCheck.Keys) {
        if ($null -eq $seqReportCheck[$rptName]) {
            [void]$script:skippedDataWarnings.Add("$rptName report — download failed")
        }
    }
}

$dlStopwatch.Stop()
$dlMode = if ($parallelSuccess) { "parallel (max $MaxParallel)" } else { "sequential" }
Write-Host "  Downloaded 11 reports in $([math]::Round($dlStopwatch.Elapsed.TotalSeconds, 1))s ($dlMode)" -ForegroundColor Green
Write-Log "Report downloads completed in $([math]::Round($dlStopwatch.Elapsed.TotalSeconds, 1))s ($dlMode)"
# Log row counts per report for post-run diagnostics
Write-Log "Report row counts: ActiveUsers=$(@($activeUserDetail).Count), M365App=$(@($m365AppDetail).Count), Email=$(@($emailActivity).Count), Teams=$(@($teamsActivity).Count), OneDriveActivity=$(@($oneDriveActivity).Count), SharePoint=$(@($sharePointActivity).Count), MailboxUsage=$(@($mailboxUsage).Count), EmailApp=$(@($emailAppUsage).Count), OneDriveUsage=$(@($oneDriveUsage).Count), TeamsDevice=$(@($teamsDeviceUsage).Count), Activations=$(@($activations).Count)"

# ── Copilot usage report (beta API — separate download, graceful degradation) ──
# Graph beta Copilot CSV sometimes contains duplicate column headers (e.g. "lastActivityDate"
# appears for each Copilot app). ConvertFrom-Csv throws "The member ... is already present".
# Fix: deduplicate headers by appending a numeric suffix to repeats before parsing.
$copilotUsageDetail = @()
$copilotUsageLoaded = $false
$copilotTempFile = $null
try {
    Write-Host "  Downloading Copilot usage report (beta) ..." -ForegroundColor DarkGreen
    $copilotTempFile = Join-Path $env:TEMP "copilotUsageDetail_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
    $copilotUri = "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUsageUserDetail(period='$ReportPeriod')"
    Invoke-GraphWithRetry -Method GET -Uri $copilotUri -OutputFilePath $copilotTempFile
    $copilotRaw = [System.IO.File]::ReadAllText($copilotTempFile)
    if ($copilotRaw.Length -gt 0 -and $copilotRaw[0] -eq [char]0xFEFF) { $copilotRaw = $copilotRaw.Substring(1) }

    # Diagnostic: log raw response size and first 500 chars for debugging parse failures
    Write-Log "Copilot CSV raw size: $($copilotRaw.Length) chars"
    $copilotPreview = if ($copilotRaw.Length -gt 500) { $copilotRaw.Substring(0, 500) + '...' } else { $copilotRaw }
    Write-Log "Copilot CSV preview: $copilotPreview"

    # Detect JSON vs CSV — beta endpoint now returns JSON {"value":[...]} instead of CSV
    if ($copilotRaw.TrimStart()[0] -eq '{') {
        Write-Log "Copilot response is JSON — remapping camelCase fields to expected column names"
        $copilotJson = ConvertFrom-Json $copilotRaw
        $copilotUsageDetail = @($copilotJson.value | ForEach-Object {
            [PSCustomObject]@{
                'User Principal Name'                        = $_.userPrincipalName
                'Display Name'                               = $_.displayName
                'Last Activity Date'                         = $_.lastActivityDate
                'Microsoft Teams Copilot Last Activity Date' = $_.microsoftTeamsCopilotLastActivityDate
                'Word Copilot Last Activity Date'            = $_.wordCopilotLastActivityDate
                'Excel Copilot Last Activity Date'           = $_.excelCopilotLastActivityDate
                'PowerPoint Copilot Last Activity Date'      = $_.powerPointCopilotLastActivityDate
                'Outlook Copilot Last Activity Date'         = $_.outlookCopilotLastActivityDate
                'OneNote Copilot Last Activity Date'         = $_.oneNoteCopilotLastActivityDate
                'Loop Copilot Last Activity Date'            = $_.loopCopilotLastActivityDate
                'Copilot Chat Last Activity Date'            = $_.copilotChatLastActivityDate
            }
        })
    } else {
        # CSV response — deduplicate headers and parse
        # Graph beta Copilot CSV returns headers with embedded date values and quotes
        # (e.g. "lastActivityDate":"2026-03-18") — simple comma-split fails.
        # Use CSV-aware parser that respects quoted fields.
        $copilotLines = $copilotRaw -split "`n", 2
        if ($copilotLines.Count -ge 2) {
            # CSV-aware header split: respect quoted fields containing commas/colons/quotes
            $headerLine = $copilotLines[0].TrimEnd("`r")
            $hdrs = [System.Collections.Generic.List[string]]::new()
            $current = [System.Text.StringBuilder]::new()
            $inQuotes = $false
            foreach ($ch in $headerLine.ToCharArray()) {
                if ($ch -eq '"') { $inQuotes = !$inQuotes; [void]$current.Append($ch) }
                elseif ($ch -eq ',' -and -not $inQuotes) { [void]$hdrs.Add($current.ToString()); [void]$current.Clear() }
                else { [void]$current.Append($ch) }
            }
            if ($current.Length -gt 0) { [void]$hdrs.Add($current.ToString()) }

            $seen = @{}
            for ($hi = 0; $hi -lt $hdrs.Count; $hi++) {
                $bare = $hdrs[$hi].Trim('"')          # strip outer quotes for uniqueness check
                if ($seen.ContainsKey($bare)) {
                    $seen[$bare]++
                    $hdrs[$hi] = "`"${bare}_$($seen[$bare])`""   # re-wrap in quotes
                } else {
                    $seen[$bare] = 1
                }
            }
            $copilotRaw = ($hdrs -join ',') + "`n" + $copilotLines[1]
            Write-Log "Copilot CSV deduped headers ($($hdrs.Count)): $($hdrs -join ' | ')"
        }
        $copilotUsageDetail = @($copilotRaw | ConvertFrom-Csv)
    }
    Remove-Item $copilotTempFile -Force -ErrorAction SilentlyContinue
    $copilotUsageLoaded = $true
    Write-Host "    getMicrosoft365CopilotUsageUserDetail : $(@($copilotUsageDetail).Count) rows" -ForegroundColor DarkGreen
    Write-Log "Copilot usage report loaded: $(@($copilotUsageDetail).Count) row(s)"
} catch {
    if ($copilotTempFile) { Remove-Item $copilotTempFile -Force -ErrorAction SilentlyContinue }
    Write-Log "Copilot usage report download failed (beta API — optional)" -Level WARN -ErrorRecord $_
    Write-Host "    Copilot usage report unavailable (beta API): $($_.Exception.Message)" -ForegroundColor Yellow
    [void]$script:skippedDataWarnings.Add("Copilot usage report (beta) — $($_.Exception.Message)")
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — Build lookup hashtables keyed by UPN
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host "`n[3/12] Indexing report data by UPN ..." -ForegroundColor Cyan
Write-Log "[3/12] Indexing report data by UPN"

# Helper to build a lookup from a report array
function Build-UPNLookup {
    param ([array]$Data, [string]$UPNColumn = 'User Principal Name')
    $ht = @{}
    if (-not $Data -or $Data.Count -eq 0) { return $ht }
    # Validate the UPN column exists in the first row — if not, the report data is
    # corrupted (e.g. BOM-prefixed headers) or the report returned no usable rows.
    $firstRow = $Data[0]
    if ($firstRow -and $UPNColumn -notin $firstRow.PSObject.Properties.Name) {
        $actualCols = ($firstRow.PSObject.Properties.Name | Select-Object -First 3) -join ', '
        Write-Warning "    Report missing expected column '$UPNColumn' (first cols: $actualCols). Skipping $($Data.Count) rows."
        return $ht
    }
    foreach ($row in $Data) {
        if ($null -eq $row) { continue }
        $upn = $row.$UPNColumn
        # Normalize UPN key: case-insensitive + trimmed (Graph/CSV reports can
        # return inconsistent casing across pages, e.g. "User@Domain.com" vs "user@domain.com")
        if ($upn) {
            $upnKey = $upn.ToString().Trim().ToLower()
            if ($ht.ContainsKey($upnKey)) {
                Write-Log "Duplicate UPN '$upnKey' in report (column '$UPNColumn') — last row wins" -Level WARN
            }
            $ht[$upnKey] = $row
        }
    }
    return $ht
}

$lkpActiveUser   = Build-UPNLookup $activeUserDetail
$lkpM365App      = Build-UPNLookup $m365AppDetail
if ($lkpM365App.Count -gt 0) {
    $sampleRow  = ($lkpM365App.GetEnumerator() | Select-Object -First 1).Value
    $sampleCols = @($sampleRow.PSObject.Properties.Name)
    Write-Log "M365 App report: $($lkpM365App.Count) user(s), $($sampleCols.Count) columns"
    $expectedCol = 'Outlook (Windows)'
    if ($expectedCol -in $sampleCols) {
        $sampleVal = $sampleRow.$expectedCol
        Write-Log "  Column '$expectedCol' found — sample value: '$sampleVal' (type: $($sampleVal.GetType().Name))"
    } else {
        $outlookCols = @($sampleCols | Where-Object { $_ -match 'Outlook' })
        Write-Log "  Column '$expectedCol' NOT FOUND — Outlook-related columns: $($outlookCols -join ', ')" -Level WARN
        Write-Log "  All columns: $($sampleCols -join ', ')" -Level WARN
    }
}
$lkpEmail        = Build-UPNLookup $emailActivity
$lkpTeams        = Build-UPNLookup $teamsActivity
$lkpOneDrive     = Build-UPNLookup $oneDriveActivity
$lkpSharePoint   = Build-UPNLookup $sharePointActivity
$lkpMailbox      = Build-UPNLookup $mailboxUsage
$lkpEmailApp     = Build-UPNLookup $emailAppUsage
$lkpODUsage      = Build-UPNLookup $oneDriveUsage -UPNColumn 'Owner Principal Name'
$lkpTeamsDevice  = Build-UPNLookup $teamsDeviceUsage
$lkpCopilotUsage = Build-UPNLookup $copilotUsageDetail

# Log lookup table sizes for post-run diagnostics
Write-Log "Lookup sizes: Email=$($lkpEmail.Count), Teams=$($lkpTeams.Count), OneDrive=$($lkpOneDrive.Count), SharePoint=$($lkpSharePoint.Count), Mailbox=$($lkpMailbox.Count), EmailApp=$($lkpEmailApp.Count), ODUsage=$($lkpODUsage.Count), TeamsDevice=$($lkpTeamsDevice.Count), Copilot=$($lkpCopilotUsage.Count)"

# Activations can have multiple rows per user (one per product type) — aggregate
$lkpActivations = @{}
$_actUPNCol = 'User Principal Name'
if ($activations -and $activations.Count -gt 0 -and $_actUPNCol -notin $activations[0].PSObject.Properties.Name) {
    Write-Warning "    Activations report missing expected column '$_actUPNCol'. Skipping $($activations.Count) rows."
} else {
    foreach ($row in $activations) {
        $upn = $row.$_actUPNCol
        if (-not $upn) { continue }
        $upnKey = $upn.ToString().Trim().ToLower()
        if (-not $lkpActivations.ContainsKey($upnKey)) {
            $lkpActivations[$upnKey] = [System.Collections.Generic.List[PSCustomObject]]::new()
        }
        $lkpActivations[$upnKey].Add($row)
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — SKU inventory & per-user license detail
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host "`n[4/12] Loading tenant SKU inventory & subscription lifecycle ..." -ForegroundColor Cyan
Write-Log "[4/12] Loading tenant SKU inventory & subscription lifecycle"
$_rawSubscribedSkus = Get-MgSubscribedSku -All

# Strip zero-width characters (U+200B, U+FEFF) that Microsoft Graph sometimes appends to CPC/W365 SKU IDs.
# Convert to PSCustomObject array so SkuPartNumber is always writable and the collection is stable.
$subscribedSkus = @(foreach ($sku in $_rawSubscribedSkus) {
    $clone = $sku | Select-Object *
    $clone.SkuPartNumber = $sku.SkuPartNumber -replace '[\u200B\uFEFF]', ''
    $clone
})
$_rawSubscribedSkus = $null   # release original Graph objects

# Build SkuId → SkuPartNumber and SkuId → ServicePlans mappings for bulk license resolution
$skuIdToPartNumber   = @{}
$skuIdToServicePlans = @{}
foreach ($sku in $subscribedSkus) {
    $skuIdToPartNumber[$sku.SkuId]   = $sku.SkuPartNumber
    $skuIdToServicePlans[$sku.SkuId] = $sku.ServicePlans
}

# ── Runtime mapping validation (LOA v1.0 spec §4.2) ──
# Detect mapping drift: suite SKUs in tenant but not in $suiteIncludes, and
# $suiteIncludes items that don't appear in the tenant's known SKUs or plans.
$knownSkuSet  = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$knownPlanSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($sku in $subscribedSkus) {
    [void]$knownSkuSet.Add($sku.SkuPartNumber)
    if ($sku.ServicePlans) {
        foreach ($sp in $sku.ServicePlans) {
            [void]$knownPlanSet.Add($sp.ServicePlanName)
        }
    }
}
# ── Build ServicePlanName → suiteIncludes-String-ID alias map (Flaw #1 fix) ──
# Graph API returns ServicePlanName (e.g. EXCHANGE_S_STANDARD) but $suiteIncludes uses String IDs
# (e.g. EXCHANGESTANDARD). Build a map so disabled-plan gating can match both forms.
$allSuiteIncludeValues = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($vals in $suiteIncludes.Values) {
    foreach ($v in $vals) { [void]$allSuiteIncludeValues.Add($v) }
}
$planNameToStringId = @{}
foreach ($sku in $subscribedSkus) {
    if (-not $sku.ServicePlans) { continue }
    foreach ($sp in $sku.ServicePlans) {
        $spName = $sp.ServicePlanName
        if (-not $spName) { continue }
        # If the ServicePlanName is already a known suiteIncludes value, no alias needed
        if ($allSuiteIncludeValues.Contains($spName)) { continue }
        # Already mapped in a previous iteration
        if ($planNameToStringId.ContainsKey($spName)) { continue }
        # Match by collapsing _S_ infix (Microsoft service-plan naming convention)
        # then stripping remaining underscores and comparing case-insensitively.
        # E.g. EXCHANGE_S_STANDARD → EXCHANGE_STANDARD → EXCHANGESTANDARD ✓
        $spNorm = ($spName -replace '_S_','_' -replace '_','').ToUpperInvariant()
        foreach ($sid in $allSuiteIncludeValues) {
            $sidNorm = ($sid -replace '_S_','_' -replace '_','').ToUpperInvariant()
            if ($spNorm -eq $sidNorm) {
                $planNameToStringId[$spName] = $sid
                break
            }
        }
    }
}
if ($planNameToStringId.Count -gt 0) {
    Write-Log "Service plan alias map built: $($planNameToStringId.Count) mapping(s) — e.g. $( ($planNameToStringId.GetEnumerator() | Select-Object -First 3 | ForEach-Object { "$($_.Key) → $($_.Value)" }) -join ', ' )"
}
# Suites in tenant that have no $suiteIncludes mapping (could be new Microsoft suite)
# Skip free/viral SKUs — duplicate detection has no cost impact for €0 SKUs
$freeSkuSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($priceKey in $skuMonthlyPrices.Keys) {
    if ($skuMonthlyPrices[$priceKey] -eq 0) { [void]$freeSkuSet.Add($priceKey) }
}
$unmappedSuites = [System.Collections.Generic.List[string]]::new()
# Paid SKUs that have ≥3 service plans but are NOT suites for duplicate-detection purposes.
# Cloud PC, Dynamics 365, Visio (has dedicated overlap detection), Copilot add-ons, Viva.
$suiteValidationSkipRx = [regex]'^(CPC_[EB]_|Windows_365_[SE]_|DYN365_|DYNAMICS_365_|D365_|VISIOCLIENT|Microsoft_365_Copilot|Microsoft_Security_Copilot|COPILOT_STUDIO|VIVA|SPE_E3_RPA1|PROJECTPROFESSIONAL|PROJECTPREMIUM|PROJECTESSENTIALS)'
foreach ($tenantSku in $knownSkuSet) {
    if ($freeSkuSet.Contains($tenantSku)) { continue }                   # free SKU — no cost impact
    if ($suiteValidationSkipRx.IsMatch($tenantSku)) { continue }         # standalone product, not a suite
    # A "suite" typically has multiple service plans — heuristic: ≥ 3 plans.
    $spCount = 0
    $skuObj = $subscribedSkus | Where-Object { $_.SkuPartNumber -eq $tenantSku } | Select-Object -First 1
    if ($skuObj -and $skuObj.ServicePlans) { $spCount = $skuObj.ServicePlans.Count }
    if ($spCount -ge 3 -and -not $suiteIncludes.ContainsKey($tenantSku)) {
        $unmappedSuites.Add("$tenantSku ($spCount plans)")
    }
}
# Items in $suiteIncludes values that don't appear in tenant's known SKUs or plans
$unknownMappingItems = [System.Collections.Generic.List[string]]::new()
$checkedItems = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($suite in $suiteIncludes.Keys) {
    foreach ($item in $suiteIncludes[$suite]) {
        if ($checkedItems.Contains($item)) { continue }
        [void]$checkedItems.Add($item)
        if (-not $knownSkuSet.Contains($item) -and -not $knownPlanSet.Contains($item)) {
            $unknownMappingItems.Add($item)
        }
    }
}
if ($unmappedSuites.Count -gt 0) {
    Write-Host "  Mapping validation:" -ForegroundColor DarkYellow
    Write-Host "    $($unmappedSuites.Count) paid suite(s) in tenant but not in `$suiteIncludes: $($unmappedSuites[0..([math]::Min(4,$unmappedSuites.Count-1))] -join ', ')" -ForegroundColor DarkYellow
}
# Log reference-only detail (items in $suiteIncludes for SKUs the tenant doesn't own — expected)
if ($unknownMappingItems.Count -gt 0) {
    Write-Log "$($unknownMappingItems.Count) reference item(s) in `$suiteIncludes not present in this tenant (expected for cross-tenant mapping coverage)"
}

# ── Pricing coverage check — warn about paid SKUs with no pricing data ──
$_knownFreeRx = [regex]'(?i)^(FLOW_FREE|POWER_BI_STANDARD|POWER_BI_PRO_TRIAL|TEAMS_FREE|TEAMS_EXPLORATORY|POWERAPPS_VIRAL|FLOW_P2_VIRAL|WINDOWS_STORE|MICROSOFT_REMOTE_ASSIST|RIGHTSMANAGEMENT_ADHOC|STREAM|CCIBOTS_PRIVPREV_VIRAL|CLIPCHAMP|POWER_AUTOMATE_ATTEND_RPA_PLAN|SPZA_IW|MICROSOFT_BUSINESS_CENTER|PHONESYSTEM_VIRTUALUSER|CDS_FILE_CAPACITY|CDS_DB_CAPACITY|RMSBASIC|POWERAPPS_DEV|Microsoft_Teams_Rooms_Basic|MICROSOFT_LOOP_FREE|FORMS_PRO|CUSTOMER_VOICE_ADDON|COPILOT_.*_VIRAL)'
$_pricingGaps = [System.Collections.Generic.List[string]]::new()
foreach ($sku in $subscribedSkus) {
    $skuClean = ($sku.SkuPartNumber -replace '[\u200B\uFEFF]', '').Trim()
    if (-not $skuClean) { continue }
    if ($_knownFreeRx.IsMatch($skuClean)) { continue }
    if ($sku.ConsumedUnits -eq 0 -and ($sku.PrepaidUnits.Enabled + $sku.PrepaidUnits.Warning) -eq 0) { continue }
    $price = Get-SkuMonthlyPrice $skuClean
    if ($price -eq [decimal]0.00) {
        # Skip SKUs tagged as trial in subscribedSkus (AppliesTo or SkuPartNumber hints)
        $skuTags = @($sku.ServicePlans | ForEach-Object { $_.AppliesTo } | Select-Object -Unique)
        $looksLikeTrial = ($skuClean -match 'TRIAL|_FREE$|_VIRAL$|DEVELOPER|FREEFLOW') -or ($skuTags -contains 'Trial')
        if (-not $looksLikeTrial) {
            [void]$_pricingGaps.Add("$skuClean ($($sku.ConsumedUnits) assigned)")
        }
    }
}
if ($_pricingGaps.Count -gt 0) {
    Write-Host "  Pricing gaps: $($_pricingGaps.Count) paid SKU(s) missing from M365SkuPricing.csv — savings may be understated:" -ForegroundColor Yellow
    foreach ($gap in $_pricingGaps) { Write-Host "    $gap" -ForegroundColor Yellow }
    Write-Log "Pricing coverage gaps: $($_pricingGaps -join '; ')" -Level WARN
    [void]$script:skippedDataWarnings.Add("$($_pricingGaps.Count) paid SKU(s) missing from M365SkuPricing.csv — cost calculations may be incomplete")
}

# ── Cloud PC usage report (beta API — separate fetch, graceful degradation) ──
# Only runs when at least one CPC/W365 SKU exists in the tenant's subscribed SKUs.
# Uses POST endpoint returning JSON (Schema+Values), not CSV.
$cloudPcUsageData    = [System.Collections.Generic.List[PSCustomObject]]::new()
$cloudPcUsageLoaded  = $false
$_hasCpcSku = $false
foreach ($sku in $subscribedSkus) {
    $skuClean = $sku.SkuPartNumber -replace '[\u200B\uFEFF]', ''
    if ($skuClean -match '^(CPC_|Windows_365_)') { $_hasCpcSku = $true; break }
}
if ($_hasCpcSku) {
    Write-Host "  Cloud PC SKUs detected in tenant — fetching usage data ..." -ForegroundColor DarkGray

    # Step 1: Use Get-MgDeviceManagementVirtualEndpointCloudPc (v1.0) to list provisioned Cloud PCs.
    # Having CPC SKUs purchased ≠ Cloud PCs provisioned. If none exist, skip the usage report.
    $provisionedCloudPCs = @()
    try {
        $provisionedCloudPCs = @(Get-MgDeviceManagementVirtualEndpointCloudPc -All `
            -Property 'id','displayName','managedDeviceName','userPrincipalName','servicePlanName','lastModifiedDateTime','provisioningType')
    } catch {
        Write-Host "    Cloud PC service not available — skipping usage report ($($_.Exception.Message))" -ForegroundColor Yellow
        Write-Log "Cloud PC pre-flight (Get-MgDeviceManagementVirtualEndpointCloudPc) failed" -Level WARN -ErrorRecord $_
        [void]$script:skippedDataWarnings.Add("Cloud PC usage — Virtual Endpoint service not available")
    }

    if ($provisionedCloudPCs.Count -eq 0 -and $script:skippedDataWarnings.Count -gt 0 -and $script:skippedDataWarnings[$script:skippedDataWarnings.Count - 1] -like 'Cloud PC*') {
        # Pre-flight error already logged above
    } elseif ($provisionedCloudPCs.Count -eq 0) {
        Write-Host "    No Cloud PCs provisioned (SKUs purchased but no devices) — skipping usage report" -ForegroundColor Yellow
        Write-Log "No Cloud PCs provisioned — CPC SKUs exist but no devices. Skipping usage report."
    } else {
        Write-Host "    $($provisionedCloudPCs.Count) Cloud PC(s) provisioned — fetching aggregated usage report ..." -ForegroundColor DarkGreen

        # Step 2: Fetch aggregated remote connection hours from the beta reports endpoint.
        # This endpoint returns TotalUsageInHour per Cloud PC (90-day window).
        # It requires a direct REST call (Invoke-MgGraphRequest has issues with the octet-stream response).
        try {
            $cpcToken = $null
            try {
                $cpcTokenResp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -OutputType HttpResponseMessage
                $cpcToken = $cpcTokenResp.RequestMessage.Headers.Authorization.Parameter
            } catch { }

            if (-not $cpcToken) {
                throw [System.InvalidOperationException]::new("Could not extract bearer token for Cloud PC API call")
            }

            $cpcHeaders  = @{ Authorization = "Bearer $cpcToken" }
            $cpcUri      = 'https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/reports/getTotalAggregatedRemoteConnectionReports'
            $cpcSkip     = 0
            $cpcTop      = 50
            $cpcAllValues = [System.Collections.Generic.List[object]]::new()
            $cpcSchema    = $null

            do {
                $cpcBodyJson = @{
                    select = @('CloudPcId','ManagedDeviceName','UserPrincipalName','TotalUsageInHour','LastActiveTime','CreatedDate','PcType','DaysSinceLastSignIn','NeverSignedIn','CloudPCStatus')
                    top    = $cpcTop
                    skip   = $cpcSkip
                } | ConvertTo-Json -Depth 3

                $cpcResp = Invoke-RestMethod -Method POST -Uri $cpcUri -Body $cpcBodyJson `
                    -ContentType 'application/json' -Headers $cpcHeaders

                $schemaP = $cpcResp.PSObject.Properties['Schema']
                if (-not $cpcSchema -and $schemaP) { $cpcSchema = $schemaP.Value }
                $valsP   = $cpcResp.PSObject.Properties['Values']
                $cpcVals = if ($valsP) { $valsP.Value } else { $null }
                if ($cpcVals -and $cpcVals.Count -gt 0) {
                    foreach ($v in $cpcVals) { $cpcAllValues.Add($v) }
                    $cpcSkip += $cpcVals.Count
                }
            } while ($cpcVals -and $cpcVals.Count -eq $cpcTop)

            # Convert Schema+Values arrays into PSCustomObject array
            if ($cpcSchema -and $cpcAllValues.Count -gt 0) {
                $cpcColNames = @($cpcSchema | ForEach-Object { $_.Column })
                foreach ($valRow in $cpcAllValues) {
                    $obj = [ordered]@{}
                    for ($i = 0; $i -lt $cpcColNames.Count; $i++) {
                        $obj[$cpcColNames[$i]] = if ($i -lt $valRow.Count) { $valRow[$i] } else { $null }
                    }
                    $cloudPcUsageData.Add([PSCustomObject]$obj)
                }
            }
            $cloudPcUsageLoaded = $true
            Write-Host "    Cloud PC usage report: $($cloudPcUsageData.Count) Cloud PC(s) with usage data" -ForegroundColor DarkGreen
            Write-Log "Cloud PC usage report loaded: $($cloudPcUsageData.Count) Cloud PC(s)"
        } catch {
            # Aggregated report failed — fall back to provisioned list (no usage hours, but we still know who has Cloud PCs)
            Write-Log "Cloud PC aggregated usage report failed (beta) — falling back to provisioned list" -Level WARN -ErrorRecord $_
            # Capture HTTP response body for diagnostics (beta API errors often have inner error codes)
            if ($_.Exception.Response) {
                try {
                    $cpcErrStream = $_.Exception.Response.GetResponseStream()
                    $cpcErrReader = [System.IO.StreamReader]::new($cpcErrStream)
                    $cpcErrBody   = $cpcErrReader.ReadToEnd()
                    $cpcErrReader.Close()
                    Write-Log "Cloud PC API error response: $cpcErrBody" -Level WARN
                } catch { }
            }
            Write-Host "    Cloud PC usage report unavailable — using provisioned list (no usage hours): $($_.Exception.Message)" -ForegroundColor Yellow
            foreach ($cpcDev in $provisionedCloudPCs) {
                $cloudPcUsageData.Add([PSCustomObject]@{
                    CloudPcId            = $cpcDev.Id
                    ManagedDeviceName    = $cpcDev.ManagedDeviceName
                    UserPrincipalName    = $cpcDev.UserPrincipalName
                    TotalUsageInHour     = $null   # unknown — report endpoint unavailable
                    CreatedDate          = $null   # unknown — report endpoint unavailable
                    LastActiveTime       = $cpcDev.LastModifiedDateTime
                    PcType               = $cpcDev.ServicePlanName
                    DaysSinceLastSignIn  = $null   # unknown — report endpoint unavailable
                    NeverSignedIn        = $null   # unknown — report endpoint unavailable
                    CloudPCStatus        = if ($cpcDev.PSObject.Properties['Status']) { $cpcDev.Status } else { $null }
                })
            }
            $cloudPcUsageLoaded = $true
            Write-Host "    Cloud PC fallback: $($cloudPcUsageData.Count) Cloud PC(s) from provisioned list (no usage hours)" -ForegroundColor Yellow
            Write-Log "Cloud PC fallback: $($cloudPcUsageData.Count) Cloud PC(s) from Get-MgDeviceManagementVirtualEndpointCloudPc"
        }
    }
} else {
    Write-Host "  No CPC/W365 SKUs in tenant — skipping Cloud PC usage report." -ForegroundColor DarkGray
    Write-Log "No CPC/W365 SKUs in tenant — skipping Cloud PC usage report"
}

# ── Cloud PC usage lookups (per-UPN aggregates — users can have multiple Cloud PCs) ──
$lkpCloudPcUsageHours       = @{}   # UPN → [double] total connected hours (sum)
$lkpCloudPcLastActive       = @{}   # UPN → [string] most recent LastActiveTime
$lkpCloudPcDaysSinceSignIn  = @{}   # UPN → [int] days since last Cloud PC sign-in (min across devices = most recent)
$lkpCloudPcType             = @{}   # UPN → [string] PcType (comma-joined if multiple)
$lkpCloudPcDeviceName       = @{}   # UPN → [string] ManagedDeviceName (comma-joined if multiple)
$lkpCloudPcNeverSignedIn    = @{}   # UPN → [bool] any Cloud PC never signed in
$lkpCloudPcStatus           = @{}   # UPN → [string] CloudPCStatus (semicolon-joined if multiple)
foreach ($cpc in $cloudPcUsageData) {
    $cpcUpn = $cpc.UserPrincipalName
    if (-not $cpcUpn) { continue }
    $cpcUpnKey = $cpcUpn.ToString().Trim().ToLower()

    # Sum hours — only add to lookup when we have actual usage data (not fallback $null)
    if ($null -ne $cpc.TotalUsageInHour) {
        $hrs = 0.0
        try { $hrs = [double]$cpc.TotalUsageInHour } catch { Write-Log "Could not parse TotalUsageInHour '$($cpc.TotalUsageInHour)' for Cloud PC $($cpc.CloudPcId)" -Level WARN }
        if ($lkpCloudPcUsageHours.ContainsKey($cpcUpnKey)) {
            $lkpCloudPcUsageHours[$cpcUpnKey] += $hrs
        } else {
            $lkpCloudPcUsageHours[$cpcUpnKey] = $hrs
        }
    }

    # Most recent LastActiveTime
    $lat = if ($cpc.LastActiveTime) { $cpc.LastActiveTime.ToString() } else { '' }
    if ($lat) {
        if ($lkpCloudPcLastActive.ContainsKey($cpcUpnKey)) {
            if ($lat -gt $lkpCloudPcLastActive[$cpcUpnKey]) { $lkpCloudPcLastActive[$cpcUpnKey] = $lat }
        } else {
            $lkpCloudPcLastActive[$cpcUpnKey] = $lat
        }
    }

    # DaysSinceLastSignIn — prefer native API value, fall back to manual computation from LastActiveTime
    $dsls = $null
    if ($null -ne $cpc.DaysSinceLastSignIn -and $cpc.DaysSinceLastSignIn -ne '') {
        try { $dsls = [int]$cpc.DaysSinceLastSignIn } catch { }
    }
    if ($null -eq $dsls -and $lat) {
        try {
            $latDate = [datetime]::Parse($lat)
            $dsls = [int][math]::Floor(((Get-Date) - $latDate).TotalDays)
        } catch { }
    }
    if ($null -ne $dsls) {
        if ($dsls -lt 0) { $dsls = 0 }
        if ($lkpCloudPcDaysSinceSignIn.ContainsKey($cpcUpnKey)) {
            if ($dsls -lt $lkpCloudPcDaysSinceSignIn[$cpcUpnKey]) { $lkpCloudPcDaysSinceSignIn[$cpcUpnKey] = $dsls }
        } else {
            $lkpCloudPcDaysSinceSignIn[$cpcUpnKey] = $dsls
        }
    }

    # NeverSignedIn — flag user if any Cloud PC was never signed into (API returns Boolean)
    if ($null -ne $cpc.NeverSignedIn) {
        $nsi = ($cpc.NeverSignedIn -eq $true -or $cpc.NeverSignedIn -eq 1 -or $cpc.NeverSignedIn -eq 'True')
        if ($nsi) {
            $lkpCloudPcNeverSignedIn[$cpcUpnKey] = $true
        } elseif (-not $lkpCloudPcNeverSignedIn.ContainsKey($cpcUpnKey)) {
            $lkpCloudPcNeverSignedIn[$cpcUpnKey] = $false
        }
    }

    # CloudPCStatus (accumulate, semicolon-joined for multi-CPC users)
    $cpcSt = if ($cpc.CloudPCStatus) { $cpc.CloudPCStatus.ToString().Trim() } else { '' }
    if ($cpcSt) {
        if ($lkpCloudPcStatus.ContainsKey($cpcUpnKey)) { $lkpCloudPcStatus[$cpcUpnKey] += "; $cpcSt" } else { $lkpCloudPcStatus[$cpcUpnKey] = $cpcSt }
    }

    # PcType and DeviceName (accumulate in lists, join later)
    $pt = if ($cpc.PcType) { $cpc.PcType.ToString().Trim() } else { '' }
    $dn = if ($cpc.ManagedDeviceName) { $cpc.ManagedDeviceName.ToString().Trim() } else { '' }
    if ($pt) {
        if ($lkpCloudPcType.ContainsKey($cpcUpnKey)) { $lkpCloudPcType[$cpcUpnKey] += "; $pt" } else { $lkpCloudPcType[$cpcUpnKey] = $pt }
    }
    if ($dn) {
        if ($lkpCloudPcDeviceName.ContainsKey($cpcUpnKey)) { $lkpCloudPcDeviceName[$cpcUpnKey] += "; $dn" } else { $lkpCloudPcDeviceName[$cpcUpnKey] = $dn }
    }
}
if ($lkpCloudPcUsageHours.Count -gt 0) {
    Write-Host "  Cloud PC usage: $($lkpCloudPcUsageHours.Count) user(s) with Cloud PC data" -ForegroundColor Green
    $nsiTrue = @($lkpCloudPcNeverSignedIn.GetEnumerator() | Where-Object { $_.Value -eq $true }).Count
    Write-Log "Cloud PC lookups: $($lkpCloudPcUsageHours.Count) usage, $($lkpCloudPcDaysSinceSignIn.Count) daysSinceSignIn, $($lkpCloudPcNeverSignedIn.Count) neverSignedIn ($nsiTrue true), $($lkpCloudPcStatus.Count) status"
}

# Subscription lifecycle — expiry dates and status per SKU (paginated)
$lkpSubscription = @{}
try {
    $subUri = "https://graph.microsoft.com/v1.0/directory/subscriptions"
    while ($subUri) {
        $subResp = Invoke-GraphWithRetry -Method GET -Uri $subUri
        foreach ($sub in $subResp['value']) {
            $sku = $sub['skuPartNumber']
            if ($sku) {
                $subStatus    = $sub['status']
                $subLifecycle = $sub['nextLifecycleDateTime']
                if ($lkpSubscription.ContainsKey($sku)) {
                    $existing = $lkpSubscription[$sku]
                    # De-escalate state: Enabled/Active override Warning/Suspended (healthiest wins)
                    # If ANY subscription for this SKU is active, the SKU is healthy — old/replaced subs may linger as Suspended
                    $_subStateRank = @{ 'Enabled' = 0; 'Active' = 0; 'Warning' = 1; 'Suspended' = 2; 'LockedOut' = 3; 'Expired' = 4 }
                    $_newRank = if ($_subStateRank.ContainsKey($subStatus)) { $_subStateRank[$subStatus] } else { 99 }
                    $_curRank = if ($_subStateRank.ContainsKey($existing.Status)) { $_subStateRank[$existing.Status] } else { 99 }
                    if ($_newRank -lt $_curRank) { $existing.Status = $subStatus }
                    # Preserve latest lifecycle date (active subscription's renewal, not old expiry)
                    if ($subLifecycle) {
                        if (-not $existing.NextLifecycle -or ([datetime]$subLifecycle -gt [datetime]$existing.NextLifecycle)) {
                            $existing.NextLifecycle = $subLifecycle
                        }
                    }
                    $existing.TotalLicenses += [int]($sub['totalLicenses'])
                } else {
                    $lkpSubscription[$sku] = @{
                        Status         = $subStatus                      # Enabled, Warning, Suspended, LockedOut
                        NextLifecycle  = $subLifecycle                    # expiry/renewal date
                        TotalLicenses  = [int]($sub['totalLicenses'])
                        SubscriptionId = $sub['id']
                    }
                }
            }
        }
        $subUri = $subResp['@odata.nextLink']
    }
    $paidSkuCount = @($subscribedSkus | Where-Object { $_.ConsumedUnits -gt 0 }).Count
    $totalAssigned = ($subscribedSkus | Measure-Object -Property ConsumedUnits -Sum).Sum
    Write-Host "  $($subscribedSkus.Count) SKU(s), $($lkpSubscription.Count) subscription(s) with lifecycle data." -ForegroundColor Green
    Write-Log "SKU inventory: $($subscribedSkus.Count) total SKUs ($paidSkuCount with assigned seats), $totalAssigned total assigned licenses"
} catch {
    Write-Log "Subscription lifecycle retrieval failed" -Level ERROR -ErrorRecord $_
    Write-Warning "  Could not retrieve subscription lifecycle: $($_.Exception.Message)"
    [void]$script:skippedDataWarnings.Add("Subscription lifecycle — $($_.Exception.Message)")
    Write-Host "  $($subscribedSkus.Count) SKU(s) loaded (no lifecycle data)." -ForegroundColor Green
}

# ── Cloud Licensing API (beta) — trial detection, assignment errors, waiting members ──
$cloudLicensingData   = @{}   # keyed by skuPartNumber → @{ IsTrial; State; CapacityPct; ... }
$clAssignmentErrors   = @{}   # keyed by target object ID → list of error descriptions
$clWaitingMembers     = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$cloudLicensingLoaded = $false

try {
    Write-Host "  Loading Cloud Licensing API (beta) ..." -ForegroundColor DarkGray
    $allotUri = "https://graph.microsoft.com/beta/admin/cloudLicensing/allotments"
    while ($allotUri) {
        $allotResp = Invoke-GraphWithRetry -Method GET -Uri $allotUri
        foreach ($allot in $allotResp['value']) {
            $skuPart     = $allot['skuPartNumber']
            $allotted    = $allot['allottedUnits']
            $consumed    = $allot['consumedUnits']
            $allotId     = $allot['id']
            $capacityPct = if ($allotted -gt 0) { [math]::Round(($consumed / $allotted) * 100) } else { 0 }

            # Parse subscription metadata (trial tags, state, dates)
            $isTrial          = $false
            $subState         = "unknown"
            $subStartDate     = $null
            $subNextLifecycle = $null
            $subs = $allot['subscriptions']
            if ($subs) {
                foreach ($sub in $subs) {
                    $tags = $sub['tags']
                    if ($tags -and ($tags -match 'trial' -or ($tags -is [System.Collections.IEnumerable] -and $tags -contains 'trial'))) {
                        $isTrial = $true
                    }
                    $st = $sub['state']
                    if ($st) { $subState = $st }
                    if ($sub['startDate'])         { $subStartDate     = $sub['startDate'] }
                    if ($sub['nextLifecycleDate'])  { $subNextLifecycle = $sub['nextLifecycleDate'] }
                }
            }

            # Store / aggregate allotments per SKU (multiple allotments = sum units, worst-case state)
            if (-not $cloudLicensingData.ContainsKey($skuPart)) {
                $cloudLicensingData[$skuPart] = @{
                    IsTrial       = $isTrial
                    State         = $subState
                    StartDate     = $subStartDate
                    NextLifecycle = $subNextLifecycle
                    CapacityPct   = $capacityPct
                    AllottedUnits = $allotted
                    ConsumedUnits = $consumed
                }
            } else {
                $existing = $cloudLicensingData[$skuPart]
                # Merge trial flag (any allotment trial → mark trial)
                if ($isTrial) { $existing.IsTrial = $true }
                # Sum capacity across allotments and recalculate percentage
                $existing.AllottedUnits += $allotted
                $existing.ConsumedUnits += $consumed
                if ($existing.AllottedUnits -gt 0) {
                    $existing.CapacityPct = [math]::Round($existing.ConsumedUnits / $existing.AllottedUnits * 100, 0)
                }
                # Preserve the most urgent lifecycle date (earliest expiration)
                if ($subNextLifecycle) {
                    if (-not $existing.NextLifecycle -or ([datetime]$subNextLifecycle -lt [datetime]$existing.NextLifecycle)) {
                        $existing.NextLifecycle = $subNextLifecycle
                    }
                }
                # Escalate state: Warning/Suspended/LockedOut override Active
                $stateRank = @{ 'Active' = 0; 'Warning' = 1; 'Suspended' = 2; 'LockedOut' = 3; 'Expired' = 4 }
                $newRank = if ($stateRank.ContainsKey($subState)) { $stateRank[$subState] } else { Write-Log "Unknown subscription state '$subState' for SKU $skuPart — treating as highest priority" -Level WARN; 99 }
                $curRank = if ($stateRank.ContainsKey($existing.State)) { $stateRank[$existing.State] } else { 99 }
                if ($newRank -gt $curRank) { $existing.State = $subState }
            }

            # Fetch waiting members for allotments at capacity
            if ($consumed -ge $allotted -and $allotted -gt 0) {
                try {
                    $wmUri = "https://graph.microsoft.com/beta/admin/cloudLicensing/allotments/$allotId/waitingMembers"
                    while ($wmUri) {
                        $wmResp = Invoke-GraphWithRetry -Method GET -Uri $wmUri
                        foreach ($wm in $wmResp['value']) {
                            $wmId = $wm['id']
                            if ($wmId) { [void]$clWaitingMembers.Add($wmId) }
                        }
                        $wmUri = $wmResp['@odata.nextLink']
                    }
                } catch {
                    # Waiting members may not be available — silently skip
                }
            }
        }
        $allotUri = $allotResp['@odata.nextLink']
    }

    # Fetch assignment errors
    try {
        $errUri = "https://graph.microsoft.com/beta/admin/cloudLicensing/assignmentErrors"
        while ($errUri) {
            $errResp = Invoke-GraphWithRetry -Method GET -Uri $errUri
            foreach ($err in $errResp['value']) {
                $targetId = $err['assignedTo']
                if (-not $targetId) {
                    # Try nested relationship format
                    $assignedTo = $err['assignedTo@odata.bind']
                    if ($assignedTo) { $targetId = ($assignedTo -split '/')[-1] -replace "[()']","" }
                }
                if ($targetId) {
                    $errMsg = $err['errorCode']
                    if (-not $errMsg) { $errMsg = $err['message'] }
                    if (-not $errMsg) { $errMsg = "Unknown assignment error" }
                    if (-not $clAssignmentErrors.ContainsKey($targetId)) {
                        $clAssignmentErrors[$targetId] = [System.Collections.Generic.List[string]]::new()
                    }
                    $clAssignmentErrors[$targetId].Add($errMsg)
                }
            }
            $errUri = $errResp['@odata.nextLink']
        }
    } catch {
        # Assignment errors endpoint may not be available yet — silently skip
    }

    $trialCount = @($cloudLicensingData.GetEnumerator() | Where-Object { $_.Value.IsTrial }).Count
    $cloudLicensingLoaded = $true
    Write-Host "  Cloud Licensing: $($cloudLicensingData.Count) allotment(s), $trialCount trial(s), $($clAssignmentErrors.Count) assignment error(s), $($clWaitingMembers.Count) waiting member(s)." -ForegroundColor Green
} catch {
    Write-Log "Cloud Licensing API not available" -Level WARN -ErrorRecord $_
    Write-Warning "  Cloud Licensing API not available: $($_.Exception.Message) — trial/capacity features will be skipped."
    [void]$script:skippedDataWarnings.Add("Cloud Licensing API (beta) — $($_.Exception.Message)")
}

Write-Host "`n[5/12] Retrieving users ..." -ForegroundColor Cyan
Write-Log "[5/12] Retrieving users"

# Pre-initialize license map — unlicensed users are marked during the paginated fetch below
$userLicenseMap      = @{}   # UPN → semicolon-separated SkuPartNumber list (or "[UNLICENSED]")

# Paginated user fetch — process page-by-page to avoid OOM on large tenants (150k+ users).
# Get-MgUser -All loads every user object into a single array in memory before continuing.
# Instead, we call the raw Graph REST API and build lookups from each page as it arrives.
$selectProps = "id,userPrincipalName,displayName,accountEnabled,department,jobTitle,usageLocation,userType,assignedLicenses,companyName,country"
$userGraphUri = "/v1.0/users?`$select=$selectProps&`$top=999"

$lkpUserObj   = @{}
$idToUpn      = @{}
$upnToId      = @{}
$lkpAssignedLicenses = @{}   # UPN → AssignedLicenses array (only licensed users)
$totalUserCount      = 0
$disabledAccountCount = 0

do {
    $response = Invoke-GraphWithRetry -Method GET -Uri $userGraphUri
    foreach ($u in $response['value']) {
        $upn = $u['userPrincipalName']
        if (-not $upn) { continue }
        $upnKey = $upn.ToString().Trim().ToLower()
        $totalUserCount++

        # Build slim user lookup (same shape as before)
        $lkpUserObj[$upnKey] = [PSCustomObject]@{
            DisplayName    = $u['displayName']
            UserType       = $u['userType']
            Department     = $u['department']
            CompanyName    = $u['companyName']
            Country        = $u['country']
            AccountEnabled = $u['accountEnabled']
        }

        # Build Id ↔ UPN lookups (used for PIM & Conditional Access scope mapping)
        if ($u['id']) {
            $idToUpn[$u['id']] = $upnKey
            $upnToId[$upnKey] = $u['id']
        }

        # Store assigned licenses for licensed users; mark unlicensed immediately
        $assigned = $u['assignedLicenses']
        if ($assigned -and $assigned.Count -gt 0) {
            $lkpAssignedLicenses[$upnKey] = $assigned
        } else {
            $userLicenseMap[$upnKey] = "[UNLICENSED]"
        }

        if ($u['accountEnabled'] -eq $false) { $disabledAccountCount++ }
    }

    # Follow pagination link
    $userGraphUri = $response['@odata.nextLink']
    if ($totalUserCount % 5000 -eq 0 -and $totalUserCount -gt 0) {
        Write-Host "    $totalUserCount users fetched so far ..." -ForegroundColor DarkGray
    }
} while ($userGraphUri)

if ($disabledAccountCount -gt 0) {
    Write-Host "  $totalUserCount user(s) retrieved ($disabledAccountCount disabled)." -ForegroundColor Green
} else {
    Write-Host "  $totalUserCount user(s) retrieved." -ForegroundColor Green
}

Write-Host "`n[6/12] Fetching sign-in activity & license assignment states (beta API) ..." -ForegroundColor Cyan
Write-Log "[6/12] Fetching sign-in activity & license assignment states"
$lkpSignIn        = @{}
$lkpLicAssignment = @{}
$groupNameCache   = @{}
$signInDataLoaded = $false
try {
    $betaUri = "https://graph.microsoft.com/beta/users?`$select=userPrincipalName,signInActivity,licenseAssignmentStates&`$top=999"
    $betaPageCount = 0
    while ($betaUri) {
        $betaPageCount++
        $resp = Invoke-GraphWithRetry -Method GET -Uri $betaUri
        foreach ($u in $resp['value']) {
            $uUpn = $u['userPrincipalName']
            if (-not $uUpn) { continue }
            $uUpnKey = $uUpn.ToString().Trim().ToLower()

            $uSia = $u['signInActivity']
            if ($uSia) { $lkpSignIn[$uUpnKey] = $uSia }

            $uLas = $u['licenseAssignmentStates']
            if ($uLas) { $lkpLicAssignment[$uUpnKey] = $uLas }
        }
        $betaUri = $resp['@odata.nextLink']
        if ($betaPageCount % 10 -eq 0) {
            Write-Host "    ... $betaPageCount pages processed" -ForegroundColor DarkGray
        }
    }

    # Collect unique group GUIDs to resolve names
    $groupIds = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($states in $lkpLicAssignment.Values) {
        foreach ($s in $states) {
            $abg = $s['assignedByGroup']
            if ($abg) { [void]$groupIds.Add($abg) }
        }
    }
    foreach ($gid in $groupIds) {
        try {
            $g = Invoke-GraphWithRetry -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$gid?`$select=displayName"
            $groupNameCache[$gid] = $g['displayName']
        } catch {
            $groupNameCache[$gid] = $gid  # Fallback to GUID
        }
    }
    $signInDataLoaded = $true
    Write-Host "  Sign-in: $($lkpSignIn.Count) user(s), Lic-assignment: $($lkpLicAssignment.Count) user(s), $($groupIds.Count) group(s). ($betaPageCount pages)" -ForegroundColor Green
} catch {
    Write-Log "Failed to retrieve beta user data (sign-in activity / license assignment)" -Level ERROR -ErrorRecord $_
    Write-Warning "  Could not retrieve beta user data (requires AuditLog.Read.All): $($_.Exception.Message)"
    [void]$script:skippedDataWarnings.Add("Sign-in activity & license assignment states — $($_.Exception.Message)")
}

# ── Group license inventory (which Entra groups have licenses assigned) ──
$groupLicenseInventory = [System.Collections.Generic.List[PSCustomObject]]::new()
try {
    $licensedGroups = @(Get-MgGroup -All -Property Id, DisplayName, AssignedLicenses, MembershipRule |
        Where-Object { $_.AssignedLicenses -and $_.AssignedLicenses.Count -gt 0 })
    foreach ($lg in $licensedGroups) {
        $memberCount = 0
        try {
            $memberCount = (Invoke-GraphWithRetry -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$($lg.Id)/members/`$count" -Headers @{ 'ConsistencyLevel' = 'eventual' })
        } catch {
            try {
                $memberCount = @(Get-MgGroupMember -GroupId $lg.Id -All).Count
            } catch { $memberCount = -1 }
        }
        $skuNames = @()
        foreach ($alic in $lg.AssignedLicenses) {
            $skuObj = $subscribedSkus | Where-Object { $_.SkuId -eq $alic.SkuId } | Select-Object -First 1
            $skuName = if ($skuObj) { $skuObj.SkuPartNumber } else { $alic.SkuId }
            $disabledCount = if ($alic.DisabledPlans) { $alic.DisabledPlans.Count } else { 0 }
            $skuNames += if ($disabledCount -gt 0) { "$skuName ($disabledCount plans disabled)" } else { $skuName }
        }
        $isDynamic = if ($lg.MembershipRule) { "Dynamic" } else { "Assigned" }
        $groupLicenseInventory.Add([PSCustomObject]@{
            'Group Name'        = $lg.DisplayName
            'Group ID'          = $lg.Id
            'Membership Type'   = $isDynamic
            'Member Count'      = $memberCount
            'Also Direct'       = 0   # backfilled after merge loop with actual overlap count
            'Assigned Licenses' = ($skuNames -join "; ")
            'License Count'     = $lg.AssignedLicenses.Count
        })
    }
    # Backfill groupNameCache from licensed groups (more reliable than individual REST lookups)
    foreach ($lg in $licensedGroups) {
        if ($lg.Id -and $lg.DisplayName -and -not $groupNameCache.ContainsKey($lg.Id)) {
            $groupNameCache[$lg.Id] = $lg.DisplayName
        }
        # Also overwrite GUID-fallback entries (where REST failed and stored GUID as name)
        if ($lg.Id -and $lg.DisplayName -and $groupNameCache[$lg.Id] -eq $lg.Id) {
            $groupNameCache[$lg.Id] = $lg.DisplayName
        }
    }
    if ($groupLicenseInventory.Count -gt 0) {
        Write-Host "  $($groupLicenseInventory.Count) licensing group(s) found." -ForegroundColor Green
    } else {
        Write-Host "  No group-based licensing detected." -ForegroundColor DarkGray
    }
} catch {
    Write-Log "Group licensing inventory failed" -Level ERROR -ErrorRecord $_
    Write-Warning "  Could not retrieve group licensing inventory: $($_.Exception.Message)"
}

Write-Host "`n[7/12] Fetching mailbox types (EXO) ..." -ForegroundColor Cyan
Write-Log "[7/12] Fetching mailbox types (EXO)"
$lkpMailboxType            = @{}
$lkpLitigationHold         = @{}   # UPN → $true if litigation hold is active
$lkpMailboxPrimarySmtp     = @{}   # UPN (lower) → primary SMTP (lower)
$lkpSmtpToUpn              = @{}   # primary SMTP (lower) → UPN (original case)
$lkpArchiveStatus          = @{}   # UPN → ArchiveStatus (Active/None)
$lkpAutoExpandingArchive   = @{}   # UPN → $true/$false
$lkpForwardingTarget       = @{}   # UPN → forwarding target (address string)
$lkpDeliverAndForward      = @{}   # UPN → $true if mail is also delivered to mailbox (not forward-only)

if ($exoConnected) {
    # Inner helper: run Get-EXOMailbox and populate lookup tables.
    # Defined as a scriptblock so the reconnect-retry loop can invoke it without code duplication.
    $script:_FetchMailboxTypes = {
        # Get-EXOMailbox -ResultSize Unlimited loads all mailboxes into a single array.
        # We extract only the lightweight lookup data we need, then release the heavy
        # deserialized objects immediately to avoid OOM on large tenants (150k+).
        $allMailboxes = @(Get-EXOMailbox -ResultSize Unlimited -Properties RecipientTypeDetails, UserPrincipalName, PrimarySmtpAddress, LitigationHoldEnabled, ArchiveStatus, AutoExpandingArchiveEnabled, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward)
        for ($i = 0; $i -lt $allMailboxes.Count; $i++) {
            $mbx = $allMailboxes[$i]
            if (-not $mbx.UserPrincipalName) { continue }

            $mbxUpn      = $mbx.UserPrincipalName
            $mbxUpnLower = $mbxUpn.ToString().Trim().ToLower()

            $lkpMailboxType[$mbxUpnLower] = $mbx.RecipientTypeDetails.ToString()
            if ($mbx.LitigationHoldEnabled -eq $true) { $lkpLitigationHold[$mbxUpnLower] = $true }
            if ($mbx.ArchiveStatus) { $lkpArchiveStatus[$mbxUpnLower] = $mbx.ArchiveStatus.ToString() }
            $lkpAutoExpandingArchive[$mbxUpnLower] = ($mbx.AutoExpandingArchiveEnabled -eq $true)

            # Forwarding configuration (ForwardingAddress = internal DN, ForwardingSmtpAddress = external SMTP)
            $fwdTarget = $null
            if ($mbx.ForwardingSmtpAddress) { $fwdTarget = $mbx.ForwardingSmtpAddress.ToString() -replace '^smtp:',''}
            elseif ($mbx.ForwardingAddress) { $fwdTarget = $mbx.ForwardingAddress.ToString() }
            if ($fwdTarget) {
                $lkpForwardingTarget[$mbxUpnLower]  = $fwdTarget
                $lkpDeliverAndForward[$mbxUpnLower] = ($mbx.DeliverToMailboxAndForward -eq $true)
            }

            $mbxSmtp = $null
            if ($mbx.PrimarySmtpAddress) { $mbxSmtp = $mbx.PrimarySmtpAddress.ToString() }
            if (-not $mbxSmtp) { $mbxSmtp = $mbxUpn }

            $mbxSmtpLower = $mbxSmtp.ToLower()
            $lkpMailboxPrimarySmtp[$mbxUpnLower] = $mbxSmtpLower

            if (-not $lkpSmtpToUpn.ContainsKey($mbxSmtpLower)) {
                $lkpSmtpToUpn[$mbxSmtpLower] = $mbxUpnLower
            }

            # Null out the processed element to allow GC to reclaim the heavy deserialized object
            $allMailboxes[$i] = $null
        }
        # Release the array shell itself
        $allMailboxes = $null
        [System.GC]::Collect()
    }

    $exoMbxRetried = $false
    :exoMbxRetry while ($true) {
        try {
            & $script:_FetchMailboxTypes
            $mbxShared = @($lkpMailboxType.Values | Where-Object { $_ -eq 'SharedMailbox' }).Count
            $mbxRoom   = @($lkpMailboxType.Values | Where-Object { $_ -eq 'RoomMailbox' -or $_ -eq 'EquipmentMailbox' }).Count
            Write-Host "  $($lkpMailboxType.Count) mailbox(es) typed (User/Shared/Room/Equipment)." -ForegroundColor Green
            Write-Log "EXO mailboxes: $($lkpMailboxType.Count) total, $mbxShared shared, $mbxRoom room/equipment, $($lkpLitigationHold.Count) on litigation hold"
            break exoMbxRetry
        } catch {
            # 401 Unauthorized — EXO access token expired between script start and this step.
            # This is common on large tenants where Graph API calls take long enough for the
            # initial token lifetime to elapse. Reconnect once and retry automatically.
            if (-not $exoMbxRetried -and $_.Exception.Message -match '401|HttpStatusCode=401|Unauthorized') {
                $exoMbxRetried = $true
                Write-Host "  EXO session token expired — reconnecting and retrying Get-EXOMailbox ..." -ForegroundColor Yellow
                Write-Log "Get-EXOMailbox returned 401 — attempting EXO reconnect (token expiry)" -Level WARN
                try {
                    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
                    if ($useCertAuth) {
                        if (-not $orgDomain) { throw "Cannot reconnect EXO — orgDomain was not set during initial connection." }
                        Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint `
                                               -Organization $orgDomain -ShowBanner:$false
                    } else {
                        Connect-ExchangeOnline -ShowBanner:$false
                    }
                    continue exoMbxRetry   # retry the fetch
                } catch {
                    Write-Log "EXO reconnect failed" -Level ERROR -ErrorRecord $_
                    # fall through to the error handling below
                }
            }
            Write-Log "Mailbox type retrieval failed (EXO)" -Level ERROR -ErrorRecord $_
            Write-Warning "  Could not retrieve mailbox types: $($_.Exception.Message)"
            [void]$script:skippedDataWarnings.Add("Mailbox types (EXO) — $($_.Exception.Message)")
            break exoMbxRetry
        }
    }
} else {
    if ($SkipEXO) {
        Write-Host "  Skipped — -SkipEXO specified." -ForegroundColor DarkGray
        Write-Warning "  Mailbox type detection (Shared/Room/Equipment) is NOT available without EXO."
        Write-Warning "  Litigation hold detection is NOT available without EXO — no Graph API equivalent exists."
        [void]$script:skippedDataWarnings.Add("Mailbox types & Litigation Hold — skipped (-SkipEXO)")
    } else {
        Write-Host "  Skipped — Exchange Online not connected." -ForegroundColor DarkGray
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# OPTIONAL — Defender for Office 365 policy scope (license exposure signal)
# Light, license-focused check: are mailboxes in scope of Safe Links / Safe
# Attachments rules (including preset policies)?
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host "`n[7b/12] Evaluating Defender for Office 365 policy coverage (Safe Links/Attachments) ..." -ForegroundColor Cyan
Write-Log "[7b/12] Evaluating Defender for Office 365 policy coverage"
$lkpMdoCoverageByUpn = @{}   # UPN → List[string] (sources)
$script:__mdoAllTenantSources = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$mdoPolicyDomains    = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)  # RecipientDomainIs from all policies — must be pre-declared for StrictMode (else block may not run)
$mdoCoverageChecked  = $false  # tracks whether MDO policy evaluation succeeded

if (-not $exoConnected) {
    if ($SkipEXO) {
        Write-Host "  Skipped — -SkipEXO specified (MDO cmdlets require EXO connection)." -ForegroundColor DarkGray
        [void]$script:skippedDataWarnings.Add("MDO policy coverage — skipped (-SkipEXO)")
    } else {
        Write-Host "  Skipped — Exchange Online not connected." -ForegroundColor DarkGray
    }
} else {
    # Always scan MDO policies when EXO is connected — even tenants without an MDO SKU
    # can have Safe Links/Attachments rules configured (policies exist, protection doesn't).
    # Detecting these is critical: users covered by MDO policies without an MDO license
    # are in a false sense of security.  Each cmdlet has its own try-catch for graceful fallback.
    $smtpCoverage = @{}   # SMTP → HashSet of coverage sources
        $script:__mdoAllTenantSources = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

        function Add-MdoCoverage {
            param([string]$Smtp, [string]$Source)
            if (-not $Smtp) { return }
            # [ALL_TENANT] marker from circuit breaker — policy covers entire tenant
            if ($Smtp -eq "[ALL_TENANT]") {
                [void]$script:__mdoAllTenantSources.Add($Source)
                return
            }
            $s = $Smtp.ToLower()
            if (-not $smtpCoverage.ContainsKey($s)) {
                $smtpCoverage[$s] = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            }
            [void]$smtpCoverage[$s].Add($Source)
        }

        $script:__mdoGroupMemberCache = @{}

        function Get-GroupMemberUpns {
            param([string]$GroupId)
            # Use comma operator on all returns to prevent pipeline from unwrapping HashSet
            if (-not $GroupId) { return ,[System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) }
            if ($script:__mdoGroupMemberCache.ContainsKey($GroupId)) { return ,$script:__mdoGroupMemberCache[$GroupId] }

            $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            $success = $false
            try {
                # Circuit breaker: check member count before expanding massive groups (100k+).
                # The /transitiveMembers endpoint heavily throttles on large "All Employees" groups,
                # paginating 100+ times and potentially timing out after 10+ minutes.
                # A single $count call is cheap and returns a scalar integer.
                $memberCount = $null
                try {
                    $countUri = "https://graph.microsoft.com/v1.0/groups/$GroupId/transitiveMembers/`$count"
                    $memberCount = Invoke-MgGraphRequest -Method GET -Uri $countUri -Headers @{ 'ConsistencyLevel' = 'eventual' }
                } catch { }
                if ($null -ne $memberCount -and [int]$memberCount -gt 10000) {
                    Write-Log "Group $GroupId is massive ($memberCount members). Returning [ALL_TENANT] marker to avoid throttling." -Level WARN
                    [void]$set.Add("[ALL_TENANT]")
                    $script:__mdoGroupMemberCache[$GroupId] = $set
                    return ,$set
                }

                $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/transitiveMembers/microsoft.graph.user?`$select=userPrincipalName&`$top=999"
                while ($uri) {
                    $resp = Invoke-GraphWithRetry -Method GET -Uri $uri
                    foreach ($m in $resp['value']) {
                        $mUpn = $m['userPrincipalName']
                        if ($mUpn) { [void]$set.Add($mUpn.ToLower()) }
                    }
                    $uri = $resp['@odata.nextLink']
                }
                $success = $true
            } catch {
                # Group member resolution failure — do NOT cache partial results
            }

            if ($success) { $script:__mdoGroupMemberCache[$GroupId] = $set }
            return ,$set
        }

        function Get-RulePropArray {
            param($Rule, [string]$PropName)
            if ($null -eq $Rule) { return ,([string[]]@()) }
            $p = $null
            try { $p = $Rule.PSObject.Properties[$PropName] } catch { return ,([string[]]@()) }
            if (-not $p) { return ,([string[]]@()) }
            $v = $null
            try { $v = $p.Value } catch { return ,([string[]]@()) }
            if ($null -eq $v) { return ,([string[]]@()) }
            # Deserialized EXO objects throw Int32 conversion errors when any form of
            # enumeration is attempted (foreach, @(), pipe).  Use only index-based access.
            if ($v -is [string]) {
                return ,([string[]]@($v))
            }
            if ($v -is [System.Array]) {
                $list = [System.Collections.Generic.List[string]]::new()
                for ($i = 0; $i -lt $v.Length; $i++) {
                    if ($null -ne $v[$i]) { $list.Add($v[$i].ToString()) }
                }
                return ,($list.ToArray())
            }
            # Deserialized EXO MultiValuedProperty implements ICollection but not Array.
            # Use .Count and index-based access (foreach/pipe throw Int32 conversion errors).
            if ($v -is [System.Collections.ICollection]) {
                $list = [System.Collections.Generic.List[string]]::new()
                for ($i = 0; $i -lt $v.Count; $i++) {
                    if ($null -ne $v[$i]) { $list.Add($v[$i].ToString()) }
                }
                return ,($list.ToArray())
            }
            # Fallback: non-array, non-collection, non-string — treat as single value
            try { return ,([string[]]@($v.ToString())) } catch { return ,([string[]]@()) }
        }

        function Resolve-RecipientObjectId {
            param([string]$Identity)
            try {
                $r = Get-EXORecipient -Identity $Identity -ErrorAction Stop
                if ($r.ExternalDirectoryObjectId) { return $r.ExternalDirectoryObjectId.ToString() }
            } catch { }
            return $null
        }

        function Resolve-IdentityToMailboxSmtps {
            param($Identity)
            $out = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            if ($null -eq $Identity) { return ,$out }
            $id = $Identity.ToString()

            # GUID that matches a user in our local map
            if ($id -match '^[0-9a-fA-F-]{36}$' -and $idToUpn.ContainsKey($id)) {
                $u = $idToUpn[$id]
                $uLower = $u.ToLower()
                $smtp = if ($lkpMailboxPrimarySmtp.ContainsKey($uLower)) { $lkpMailboxPrimarySmtp[$uLower] } else { $uLower }
                if ($smtp) { [void]$out.Add($smtp) }
                return ,$out
            }

            # Looks like an SMTP address
            if ($id -match '@') {
                $smtp = $id.ToLower()
                if ($lkpSmtpToUpn.ContainsKey($smtp)) {
                    $u = $lkpSmtpToUpn[$smtp]
                    $uLower = $u.ToLower()
                    if ($lkpMailboxPrimarySmtp.ContainsKey($uLower)) { $smtp = $lkpMailboxPrimarySmtp[$uLower] }
                    [void]$out.Add($smtp)
                    return ,$out
                }
                # Alias/proxy address not in primary-SMTP map — fall through to
                # EXORecipient resolution below, which can resolve any alias.
            }

            # Try to resolve via EXORecipient
            $objId = Resolve-RecipientObjectId -Identity $id
            if ($objId -and $objId -match '^[0-9a-fA-F-]{36}$') {
                $members = Get-GroupMemberUpns -GroupId $objId
                # Propagate [ALL_TENANT] marker from circuit breaker (massive group)
                if ($members.Contains("[ALL_TENANT]")) {
                    [void]$out.Add("[ALL_TENANT]")
                    return ,$out
                }
                if ($members.Count -gt 0) {
                    foreach ($mUpnLower in $members) {
                        if ($lkpMailboxPrimarySmtp.ContainsKey($mUpnLower)) {
                            [void]$out.Add($lkpMailboxPrimarySmtp[$mUpnLower])
                        }
                    }
                    if ($out.Count -gt 0) { return ,$out }
                }
                if ($idToUpn.ContainsKey($objId)) {
                    $u = $idToUpn[$objId]
                    $uLower = $u.ToLower()
                    $smtp = if ($lkpMailboxPrimarySmtp.ContainsKey($uLower)) { $lkpMailboxPrimarySmtp[$uLower] } else { $uLower }
                    if ($smtp) { [void]$out.Add($smtp) }
                }
            }
            return ,$out
        }

        function Get-ScopedMailboxSmtpsFromRule {
            param($Rule)
            $included = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

            # Use for-loops with index access throughout — foreach/ForEach-Object on
            # deserialized EXO objects triggers Int32 conversion errors.
            $arr = Get-RulePropArray -Rule $Rule -PropName 'SentTo'
            for ($i = 0; $i -lt $arr.Length; $i++) {
                $smtps = Resolve-IdentityToMailboxSmtps -Identity $arr[$i]
                foreach ($smtp in $smtps) { [void]$included.Add($smtp) }  # HashSet[string] is safe
            }
            $arr = Get-RulePropArray -Rule $Rule -PropName 'SentToMemberOf'
            for ($i = 0; $i -lt $arr.Length; $i++) {
                $smtps = Resolve-IdentityToMailboxSmtps -Identity $arr[$i]
                foreach ($smtp in $smtps) { [void]$included.Add($smtp) }
            }

            # Always collect RecipientDomainIs into $mdoPolicyDomains (parent scope) for
            # downstream domain-scope checking — even when SentTo/SentToMemberOf matched.
            $incDomains = Get-RulePropArray -Rule $Rule -PropName 'RecipientDomainIs'
            for ($d = 0; $d -lt $incDomains.Length; $d++) {
                if ($incDomains[$d]) { [void]$mdoPolicyDomains.Add($incDomains[$d].ToLower()) }
            }

            # Domain-scoped policies: when SentTo/SentToMemberOf are both empty but
            # RecipientDomainIs is set, the policy applies to ALL mailboxes on those domains.
            # This is extremely common — preset policies and domain-wide custom rules use this.
            if ($included.Count -eq 0 -and $incDomains.Length -gt 0) {
                foreach ($smtp in $allMailboxSmtps) {
                    for ($d = 0; $d -lt $incDomains.Length; $d++) {
                        if (-not $incDomains[$d]) { continue }
                        if ($smtp -like "*@$($incDomains[$d].ToLower())") {
                            [void]$included.Add($smtp)
                            break
                        }
                    }
                }
            }
            if ($included.Count -eq 0) { return ,$included }

            $excluded = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            $arr = Get-RulePropArray -Rule $Rule -PropName 'ExceptIfSentTo'
            for ($i = 0; $i -lt $arr.Length; $i++) {
                $smtps = Resolve-IdentityToMailboxSmtps -Identity $arr[$i]
                foreach ($smtp in $smtps) { [void]$excluded.Add($smtp) }
            }
            $arr = Get-RulePropArray -Rule $Rule -PropName 'ExceptIfSentToMemberOf'
            for ($i = 0; $i -lt $arr.Length; $i++) {
                $smtps = Resolve-IdentityToMailboxSmtps -Identity $arr[$i]
                foreach ($smtp in $smtps) { [void]$excluded.Add($smtp) }
            }
            $exDomains = Get-RulePropArray -Rule $Rule -PropName 'ExceptIfRecipientDomainIs'

            # [ALL_TENANT] marker from circuit breaker means a massive group was excluded —
            # the entire tenant is effectively excluded from this policy rule.
            if ($excluded.Contains("[ALL_TENANT]")) {
                return ,([System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase))
            }

            $snapshot = [string[]]@($included)  # clone HashSet to string[] for safe iteration
            for ($i = 0; $i -lt $snapshot.Length; $i++) {
                $smtp = $snapshot[$i]
                if ($excluded.Contains($smtp)) { [void]$included.Remove($smtp); continue }
                for ($j = 0; $j -lt $exDomains.Length; $j++) {
                    if (-not $exDomains[$j]) { continue }
                    $dom = $exDomains[$j].ToLower()
                    if ($smtp -like "*@$dom") { [void]$included.Remove($smtp); break }
                }
            }
            return ,$included
        }

        # Build set of all known mailbox SMTPs from the pre-built lookup table.
        # NOTE: $allMailboxes was freed after populating lookups (GC optimization).
        # $lkpMailboxPrimarySmtp holds UPN → primary SMTP (already lowercased).
        $allMailboxSmtps = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($smtpVal in $lkpMailboxPrimarySmtp.Values) {
            if ($smtpVal) { [void]$allMailboxSmtps.Add($smtpVal) }
        }

        # 1) Built-in protection (applies broadly)
        try {
            $bi = @(Get-ATPBuiltInProtectionRule -ErrorAction Stop)
            if ($bi.Count -gt 0) {
                $biRule = $bi[0]
                $covered = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                foreach ($smtp in $allMailboxSmtps) { [void]$covered.Add($smtp) }

                $exSmtps = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                $arr = Get-RulePropArray -Rule $biRule -PropName 'ExceptIfSentTo'
                for ($i = 0; $i -lt $arr.Length; $i++) {
                    $smtps = Resolve-IdentityToMailboxSmtps -Identity $arr[$i]
                    foreach ($smtp in $smtps) { [void]$exSmtps.Add($smtp) }
                }
                $arr = Get-RulePropArray -Rule $biRule -PropName 'ExceptIfSentToMemberOf'
                for ($i = 0; $i -lt $arr.Length; $i++) {
                    $smtps = Resolve-IdentityToMailboxSmtps -Identity $arr[$i]
                    foreach ($smtp in $smtps) { [void]$exSmtps.Add($smtp) }
                }
                $exDomains = Get-RulePropArray -Rule $biRule -PropName 'ExceptIfRecipientDomainIs'

                # [ALL_TENANT] marker from circuit breaker — entire tenant excluded from built-in protection
                if ($exSmtps.Contains("[ALL_TENANT]")) {
                    $covered.Clear()
                }

                $snapshot = [string[]]@($covered)
                for ($i = 0; $i -lt $snapshot.Length; $i++) {
                    $smtp = $snapshot[$i]
                    if ($exSmtps.Contains($smtp)) { [void]$covered.Remove($smtp); continue }
                    for ($j = 0; $j -lt $exDomains.Length; $j++) {
                        if (-not $exDomains[$j]) { continue }
                        $dom = $exDomains[$j].ToLower()
                        if ($smtp -like "*@$dom") { [void]$covered.Remove($smtp); break }
                    }
                }

                foreach ($smtp in $covered) { Add-MdoCoverage -Smtp $smtp -Source "BuiltInProtection" }
            }
        } catch {
            Write-Log "MDO Built-in Protection evaluation failed" -Level ERROR -ErrorRecord $_
            Write-Warning "    MDO Built-in Protection evaluation failed: $($_.Exception.Message)"
            [void]$script:skippedDataWarnings.Add("MDO Built-in Protection — $($_.Exception.Message)")
        }

        # 2) Preset security policy rules (Standard / Strict)
        #    NOTE: EXO implicit remoting returns Deserialized.* objects that break
        #    foreach/ForEach-Object enumeration. Use only for-loops with index access.
        try {
            $presetRules = @(Get-ATPProtectionPolicyRule -ErrorAction Stop)
            for ($ri = 0; $ri -lt $presetRules.Count; $ri++) {
                $r = $presetRules[$ri]
                try {
                    $isEnabled = $true
                    $stP = $r.PSObject.Properties['State']
                    if ($stP) { $isEnabled = ($stP.Value.ToString() -eq 'Enabled') }
                    else {
                        $enP = $r.PSObject.Properties['Enabled']
                        if ($enP) { $isEnabled = ($enP.Value.ToString() -eq 'True') }
                    }
                    if (-not $isEnabled) { continue }

                    $scope = Get-ScopedMailboxSmtpsFromRule -Rule $r
                    if ($scope.Count -eq 0) { continue }

                    $rName = 'Preset'
                    $polP = $r.PSObject.Properties['Policy']
                    $namP = $r.PSObject.Properties['Name']
                    if ($polP -and $polP.Value) { $rName = "Preset:$($polP.Value.ToString())" }
                    elseif ($namP -and $namP.Value) { $rName = "Preset:$($namP.Value.ToString())" }
                    foreach ($smtp in $scope) { Add-MdoCoverage -Smtp $smtp -Source $rName }
                } catch {
                    $ruleName = try { $r.PSObject.Properties['Name'].Value.ToString() } catch { 'unknown' }
                    Write-Log "MDO Preset rule '$ruleName' evaluation failed: $($_.Exception.Message)" -Level WARN
                }
            }
        } catch {
            Write-Log "MDO Preset Security Policy evaluation failed" -Level ERROR -ErrorRecord $_
            Write-Warning "    MDO Preset Security Policy evaluation failed: $($_.Exception.Message)"
            [void]$script:skippedDataWarnings.Add("MDO Preset Security Policies — $($_.Exception.Message)")
        }

        # 3) Custom Safe Links rules
        try {
            $slRules = @(Get-SafeLinksRule -ErrorAction Stop)
            for ($ri = 0; $ri -lt $slRules.Count; $ri++) {
                $r = $slRules[$ri]
                try {
                    $stP = $r.PSObject.Properties['State']
                    if ($stP -and $stP.Value.ToString() -ne 'Enabled') { continue }
                    $scope = Get-ScopedMailboxSmtpsFromRule -Rule $r
                    if ($scope.Count -eq 0) { continue }
                    $rName = 'SafeLinks'
                    $namP = $r.PSObject.Properties['Name']
                    if ($namP -and $namP.Value) { $rName = "SafeLinks:$($namP.Value.ToString())" }
                    foreach ($smtp in $scope) { Add-MdoCoverage -Smtp $smtp -Source $rName }
                } catch {
                    $ruleName = try { $r.PSObject.Properties['Name'].Value.ToString() } catch { 'unknown' }
                    Write-Log "MDO Safe Links rule '$ruleName' evaluation failed: $($_.Exception.Message)" -Level WARN
                }
            }
        } catch {
            Write-Log "MDO Safe Links rule evaluation failed" -Level ERROR -ErrorRecord $_
            Write-Warning "    MDO Safe Links rule evaluation failed: $($_.Exception.Message)"
            [void]$script:skippedDataWarnings.Add("MDO Safe Links rules — $($_.Exception.Message)")
        }

        # 4) Custom Safe Attachments rules
        try {
            $saRules = @(Get-SafeAttachmentRule -ErrorAction Stop)
            for ($ri = 0; $ri -lt $saRules.Count; $ri++) {
                $r = $saRules[$ri]
                try {
                    $stP = $r.PSObject.Properties['State']
                    if ($stP -and $stP.Value.ToString() -ne 'Enabled') { continue }
                    $scope = Get-ScopedMailboxSmtpsFromRule -Rule $r
                    if ($scope.Count -eq 0) { continue }
                    $rName = 'SafeAttach'
                    $namP = $r.PSObject.Properties['Name']
                    if ($namP -and $namP.Value) { $rName = "SafeAttach:$($namP.Value.ToString())" }
                    foreach ($smtp in $scope) { Add-MdoCoverage -Smtp $smtp -Source $rName }
                } catch {
                    $ruleName = try { $r.PSObject.Properties['Name'].Value.ToString() } catch { 'unknown' }
                    Write-Log "MDO Safe Attachments rule '$ruleName' evaluation failed: $($_.Exception.Message)" -Level WARN
                }
            }
        } catch {
            Write-Log "MDO Safe Attachments rule evaluation failed" -Level ERROR -ErrorRecord $_
            Write-Warning "    MDO Safe Attachments rule evaluation failed: $($_.Exception.Message)"
            [void]$script:skippedDataWarnings.Add("MDO Safe Attachments rules — $($_.Exception.Message)")
        }

        # 5) Custom Anti-Phishing rules (advanced impersonation protection requires MDO P1/P2)
        try {
            $apRules = @(Get-AntiPhishRule -ErrorAction Stop)
            for ($ri = 0; $ri -lt $apRules.Count; $ri++) {
                $r = $apRules[$ri]
                try {
                    $stP = $r.PSObject.Properties['State']
                    if ($stP -and $stP.Value.ToString() -ne 'Enabled') { continue }
                    $scope = Get-ScopedMailboxSmtpsFromRule -Rule $r
                    if ($scope.Count -eq 0) { continue }
                    $rName = 'AntiPhish'
                    $namP = $r.PSObject.Properties['Name']
                    if ($namP -and $namP.Value) { $rName = "AntiPhish:$($namP.Value.ToString())" }
                    foreach ($smtp in $scope) { Add-MdoCoverage -Smtp $smtp -Source $rName }
                } catch {
                    $ruleName = try { $r.PSObject.Properties['Name'].Value.ToString() } catch { 'unknown' }
                    Write-Log "MDO Anti-Phishing rule '$ruleName' evaluation failed: $($_.Exception.Message)" -Level WARN
                }
            }
        } catch {
            Write-Log "MDO Anti-Phishing rule evaluation failed" -Level ERROR -ErrorRecord $_
            Write-Warning "    MDO Anti-Phishing rule evaluation failed: $($_.Exception.Message)"
            [void]$script:skippedDataWarnings.Add("MDO Anti-Phishing rules — $($_.Exception.Message)")
        }

        # Convert SMTP coverage to UPN-keyed coverage for the main report
        foreach ($smtp in $smtpCoverage.Keys) {
            $covUpn = $null
            if ($lkpSmtpToUpn.ContainsKey($smtp)) {
                $covUpn = $lkpSmtpToUpn[$smtp]  # already normalized (lowercase)
            } elseif ($smtp -match '@') {
                $covUpn = $smtp.ToLower()
            }
            if (-not $covUpn) { continue }

            if (-not $lkpMdoCoverageByUpn.ContainsKey($covUpn)) {
                $lkpMdoCoverageByUpn[$covUpn] = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            }
            foreach ($src in $smtpCoverage[$smtp]) { [void]$lkpMdoCoverageByUpn[$covUpn].Add($src) }
        }
    $mdoNonBuiltInCount = @($lkpMdoCoverageByUpn.GetEnumerator() | Where-Object { @($_.Value | Where-Object { $_ -ne 'BuiltInProtection' }).Count -gt 0 }).Count
    Write-Host "  MDO coverage mapped for $($lkpMdoCoverageByUpn.Count) mailbox(es)." -ForegroundColor Green
    Write-Log "MDO coverage: $($lkpMdoCoverageByUpn.Count) total, $mdoNonBuiltInCount with explicit policies (non-BuiltIn), $($script:__mdoAllTenantSources.Count) tenant-wide source(s)"
    # Mark as checked because the MDO cmdlets executed (even if no coverage was found).
    # Tying this to result count causes false "MDO not evaluated" warnings when all policies are empty/disabled.
    $mdoCoverageChecked = $true
    if ($mdoPolicyDomains.Count -gt 0) {
        Write-Log "MDO policy domains: $($mdoPolicyDomains.Count) — $($mdoPolicyDomains -join ', ')"
    } else {
        Write-Log "MDO policies use group-based targeting only (no RecipientDomainIs found)"
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION — PIM & Risk-Based Conditional Access (Entra ID P2 usage signals)
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host "`n[8b/12] Detecting Entra feature usage signals (PIM & risk-based Conditional Access) ..." -ForegroundColor Cyan
Write-Log "[8b/12] Detecting Entra feature usage signals (PIM & risk-based CA)"

$lkpPimEligibleRoles = @{}   # UPN → HashSet(RoleNames)
$lkpPimActiveRoles   = @{}   # UPN → HashSet(RoleNames)
$riskPoliciesIncludeAll = [System.Collections.Generic.List[hashtable]]::new()
$riskPoliciesScoped     = [System.Collections.Generic.List[hashtable]]::new()
$caPoliciesIncludeAll   = [System.Collections.Generic.List[hashtable]]::new()   # non-risk CA policies (P1)
$caPoliciesScoped       = [System.Collections.Generic.List[hashtable]]::new()

function Add-ToHashSetLookup {
    param([hashtable]$Lookup, [string]$Key, [string]$Value)
    if (-not $Key -or -not $Value) { return }
    if (-not $Lookup.ContainsKey($Key)) {
        $Lookup[$Key] = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    }
    [void]$Lookup[$Key].Add($Value)
}

# ── Role definitions (used by PIM and admin role display) ──
$roleDefName = @{}
try {
    $uri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?`$select=id,displayName"
    while ($uri) {
        $resp = Invoke-GraphWithRetry -Method GET -Uri $uri
        if (-not $resp) { break }
        foreach ($rd in $resp['value']) {
            $rdId = $rd['id']; $rdName = $rd['displayName']
            if ($rdId -and $rdName) { $roleDefName[$rdId] = $rdName }
        }
        $uri = $resp['@odata.nextLink']
    }
} catch {
    Write-Log "Role definitions retrieval failed" -Level ERROR -ErrorRecord $_
    Write-Warning "  Could not retrieve role definitions: $($_.Exception.Message)"
}

# ── PIM role eligibility / assignment schedule instances (requires Entra ID P2 or Governance) ──
try {
    # Helper: resolve PIM principalId → user UPN(s). Direct assignments map 1:1 via $idToUpn;
    # group-based assignments (memberType = "Group") require expanding group → transitive user members.
    $pimGroupCache = @{}
    function Resolve-PimPrincipal {
        param([string]$PrincipalId, [string]$MemberType)
        # Known user ID — direct assignment
        if ($idToUpn.ContainsKey($PrincipalId)) {
            return @($idToUpn[$PrincipalId])
        }
        # Not a known user — try expanding as group (PIM assigns roles to groups with memberType="Direct",
        # not "Group" as the docs imply; the principalId is the group's object ID, not in $idToUpn)
        if ($pimGroupCache.ContainsKey($PrincipalId)) { return $pimGroupCache[$PrincipalId] }
        $groupUpns = [System.Collections.Generic.List[string]]::new()
        try {
            $gUri = "https://graph.microsoft.com/v1.0/groups/$PrincipalId/transitiveMembers/microsoft.graph.user?`$select=userPrincipalName&`$top=999"
            while ($gUri) {
                $gResp = Invoke-GraphWithRetry -Method GET -Uri $gUri
                if (-not $gResp) { break }
                foreach ($gm in $gResp['value']) {
                    $gmUpn = $gm['userPrincipalName']
                    if ($gmUpn) { $groupUpns.Add($gmUpn.ToString().Trim().ToLower()) }
                }
                $gUri = $gResp['@odata.nextLink']
            }
        } catch {
            Write-Log "PIM principal expansion failed for $PrincipalId — $($_.Exception.Message)" -Level WARN
        }
        $pimGroupCache[$PrincipalId] = $groupUpns.ToArray()
        return $pimGroupCache[$PrincipalId]
    }

    # Eligibility schedule instances (PIM signal)
    $uri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilityScheduleInstances?`$select=principalId,roleDefinitionId,memberType,startDateTime,endDateTime&`$top=999"
    while ($uri) {
        $resp = Invoke-GraphWithRetry -Method GET -Uri $uri
        if (-not $resp) { break }
        foreach ($inst in $resp['value']) {
            $principalId = $inst['principalId']
            if (-not $principalId) { continue }
            $memberType = $inst['memberType']
            $rid = $inst['roleDefinitionId']
            $roleName = if ($rid -and $roleDefName.ContainsKey($rid)) { $roleDefName[$rid] } else { $rid }
            $resolvedUpns = Resolve-PimPrincipal -PrincipalId $principalId -MemberType $memberType
            foreach ($pimUpn in $resolvedUpns) {
                Add-ToHashSetLookup -Lookup $lkpPimEligibleRoles -Key $pimUpn -Value $roleName
            }
        }
        $uri = $resp['@odata.nextLink']
    }

    # Assignment schedule instances — only PIM-activated roles (assignmentType = "Activated")
    # Permanent/direct assignments (assignmentType = "Assigned") do NOT require Entra ID P2
    $uri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleInstances?`$select=principalId,roleDefinitionId,assignmentType,memberType,startDateTime,endDateTime&`$top=999"
    while ($uri) {
        $resp = Invoke-GraphWithRetry -Method GET -Uri $uri
        if (-not $resp) { break }
        foreach ($inst in $resp['value']) {
            if ($inst['assignmentType'] -ne 'Activated') { continue }  # skip permanent/direct assignments
            $principalId = $inst['principalId']
            if (-not $principalId) { continue }
            $memberType = $inst['memberType']
            $rid = $inst['roleDefinitionId']
            $roleName = if ($rid -and $roleDefName.ContainsKey($rid)) { $roleDefName[$rid] } else { $rid }
            $resolvedUpns = Resolve-PimPrincipal -PrincipalId $principalId -MemberType $memberType
            foreach ($pimUpn in $resolvedUpns) {
                Add-ToHashSetLookup -Lookup $lkpPimActiveRoles -Key $pimUpn -Value $roleName
            }
        }
        $uri = $resp['@odata.nextLink']
    }

    $pimGroupCount = $pimGroupCache.Count
    Write-Host "  PIM eligible: $($lkpPimEligibleRoles.Count) user(s); Active privileged roles: $($lkpPimActiveRoles.Count) user(s)$(if ($pimGroupCount -gt 0) { " (expanded $pimGroupCount PIM group(s))" })." -ForegroundColor Green
    Write-Log "PIM: $($lkpPimEligibleRoles.Count) eligible user(s), $($lkpPimActiveRoles.Count) active user(s), $pimGroupCount group(s) expanded"
} catch {
    # PIM schedule endpoints return 400 BadRequest when tenant has no Entra ID P2 / Governance
    Write-Log "PIM not available (requires Entra P2)" -Level WARN -ErrorRecord $_
    Write-Host "  PIM not available (requires Entra ID P2 or Governance license) — skipping." -ForegroundColor DarkGray
    [void]$script:skippedDataWarnings.Add("PIM role schedule instances — tenant may not have Entra ID P2 or Governance")
}

# ── Admin role assignments (must run BEFORE CA parsing so $lkpRoleTemplateMembers is populated for includeRoles/excludeRoles) ──
Write-Host "`n[8/12] Fetching admin role assignments ..." -ForegroundColor Cyan
Write-Log "[8/12] Fetching admin role assignments"
$lkpAdminRoles = @{}
$lkpRoleTemplateMembers = @{}   # RoleTemplateId → HashSet[string] of lowercased UPNs (for CA includeRoles/excludeRoles)
try {
    $directoryRoles = @(Get-MgDirectoryRole -All)
    foreach ($role in $directoryRoles) {
        $members = @(Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All)
        $rtId = $role.RoleTemplateId
        if ($rtId -and -not $lkpRoleTemplateMembers.ContainsKey($rtId)) {
            $lkpRoleTemplateMembers[$rtId] = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        }
        foreach ($member in $members) {
            try {
                # Role members can be users, service principals, or groups — only users have userPrincipalName
                $odataType = $null
                if ($member.AdditionalProperties -and $member.AdditionalProperties.ContainsKey('@odata.type')) {
                    $odataType = $member.AdditionalProperties['@odata.type']
                }
                if ($odataType -and $odataType -ne '#microsoft.graph.user') { continue }
                $memberUpn = $null
                if ($member.AdditionalProperties -and $member.AdditionalProperties.ContainsKey('userPrincipalName')) {
                    $memberUpn = $member.AdditionalProperties['userPrincipalName']
                }
                if ($memberUpn) {
                    $memberUpnKey = $memberUpn.ToString().Trim().ToLower()
                    if (-not $lkpAdminRoles.ContainsKey($memberUpnKey)) {
                        $lkpAdminRoles[$memberUpnKey] = [System.Collections.Generic.List[string]]::new()
                    }
                    $lkpAdminRoles[$memberUpnKey].Add($role.DisplayName)
                    if ($rtId) { [void]$lkpRoleTemplateMembers[$rtId].Add($memberUpnKey) }
                }
            } catch { <# skip individual member parse failures — service principals, groups, etc. #> }
        }
    }
    Write-Host "  $($lkpAdminRoles.Count) user(s) with admin roles." -ForegroundColor Green
    Write-Log "Admin roles: $($lkpAdminRoles.Count) user(s) with direct role assignments"
} catch {
    Write-Log "Admin role assignments retrieval failed" -Level ERROR -ErrorRecord $_
    Write-Warning "  Could not retrieve admin roles (requires RoleManagement.Read.Directory): $($_.Exception.Message)"
    [void]$script:skippedDataWarnings.Add("Admin role assignments — $($_.Exception.Message)")
}

# ── Risk-based Conditional Access policies (User/Sign-in risk) ──
try {
    $script:__caGroupMemberCache = @{}

    function Get-GroupUserUpns {
        param([string]$GroupId)
        if (-not $GroupId) { return ,[System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) }
        if ($script:__caGroupMemberCache.ContainsKey($GroupId)) { return ,$script:__caGroupMemberCache[$GroupId] }

        $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        # Circuit breaker: skip massive groups (>10k members) to avoid throttling — treat as tenant-wide
        $memberCount = $null
        try {
            $countUri = "https://graph.microsoft.com/v1.0/groups/$GroupId/transitiveMembers/`$count"
            $memberCount = Invoke-MgGraphRequest -Method GET -Uri $countUri -Headers @{ 'ConsistencyLevel' = 'eventual' }
        } catch { }
        if ($null -ne $memberCount -and [int]$memberCount -gt 10000) {
            Write-Log "CA group $GroupId is massive ($memberCount members). Returning [ALL_TENANT] marker." -Level WARN
            [void]$set.Add("[ALL_TENANT]")
            $script:__caGroupMemberCache[$GroupId] = $set
            return ,$set
        }
        $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/transitiveMembers/microsoft.graph.user?`$select=userPrincipalName&`$top=999"
        $success = $false
        try {
            while ($uri) {
                $resp = Invoke-GraphWithRetry -Method GET -Uri $uri
                if (-not $resp) { break }
                $members = $resp['value']
                if (-not $members) { break }
                foreach ($m in $members) {
                    $mUpn = $m['userPrincipalName']
                    if ($mUpn) { [void]$set.Add($mUpn.ToLower()) }
                }
                $uri = $resp['@odata.nextLink']
            }
            $success = $true
        } catch {
            # Group member resolution can fail for deleted/inaccessible groups — do NOT cache partial results
        }
        if ($success) { $script:__caGroupMemberCache[$GroupId] = $set }
        return ,$set
    }

    function New-UpnHashSet { return ,[System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) }

    function Add-UpnsFromUserIds {
        param([System.Collections.Generic.HashSet[string]]$Set, $UserIds)
        foreach ($id in @($UserIds)) {
            if (-not $id) { continue }
            $s = $id.ToString()
            if ($s -match '^[0-9a-fA-F-]{36}$' -and $idToUpn.ContainsKey($s)) {
                [void]$Set.Add($idToUpn[$s].ToLower())
            }
        }
    }

    function Add-UpnsFromRoleIds {
        param([System.Collections.Generic.HashSet[string]]$Set, $RoleIds)
        foreach ($rid in @($RoleIds)) {
            if (-not $rid) { continue }
            $s = $rid.ToString()
            if ($s -notmatch '^[0-9a-fA-F-]{36}$') { continue }
            if ($lkpRoleTemplateMembers.ContainsKey($s)) {
                foreach ($u in $lkpRoleTemplateMembers[$s]) { [void]$Set.Add($u) }
            }
        }
    }

    function Add-UpnsFromGroupIds {
        param([System.Collections.Generic.HashSet[string]]$Set, $GroupIds)
        foreach ($gid in @($GroupIds)) {
            if (-not $gid) { continue }
            $s = $gid.ToString()
            if ($s -notmatch '^[0-9a-fA-F-]{36}$') { continue }
            $members = Get-GroupUserUpns -GroupId $s
            foreach ($mUpn in $members) { [void]$Set.Add($mUpn) }
        }
    }

    $uri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies?`$top=200"
    $riskPolicyCount = 0
    $caPolicyCount   = 0

    while ($uri) {
        $resp = Invoke-GraphWithRetry -Method GET -Uri $uri
        if (-not $resp) { break }
        $policies = $resp['value']
        if (-not $policies) { break }
        foreach ($p in $policies) {
            try {
                # Skip disabled/report-only policies — only count enforced ones
                $polState = $p['state']
                if ($polState -ne 'enabled') { continue }

                $conds = $p['conditions']
                if (-not $conds) { continue }

                $polName = if ($p['displayName']) { $p['displayName'] } else { $p['id'] }
                $usersCond = $conds['users']
                if (-not $usersCond) { continue }

                $includeUsers  = @(if ($usersCond['includeUsers'])  { $usersCond['includeUsers'] }  else { @() })
                $excludeUsers  = @(if ($usersCond['excludeUsers'])  { $usersCond['excludeUsers'] }  else { @() })
                $includeGroups = @(if ($usersCond['includeGroups']) { $usersCond['includeGroups'] } else { @() })
                $excludeGroups = @(if ($usersCond['excludeGroups']) { $usersCond['excludeGroups'] } else { @() })
                $includeRoles  = @(if ($usersCond['includeRoles'])  { $usersCond['includeRoles'] }  else { @() })
                $excludeRoles  = @(if ($usersCond['excludeRoles'])  { $usersCond['excludeRoles'] }  else { @() })

                $excluded = New-UpnHashSet
                Add-UpnsFromUserIds  -Set $excluded -UserIds $excludeUsers
                Add-UpnsFromGroupIds -Set $excluded -GroupIds $excludeGroups
                Add-UpnsFromRoleIds  -Set $excluded -RoleIds $excludeRoles

                $targetsAll = ($includeUsers -contains "All")

                # Classify: risk-based (P2) vs general (P1)
                $sir = $conds['signInRiskLevels']
                $usr = $conds['userRiskLevels']
                $hasRisk = ($sir -and ($sir -is [System.Collections.IEnumerable]) -and @($sir).Count -gt 0) -or
                           ($usr -and ($usr -is [System.Collections.IEnumerable]) -and @($usr).Count -gt 0)

                if ($hasRisk) {
                    $riskPolicyCount++
                    if ($targetsAll) {
                        [void]$riskPoliciesIncludeAll.Add(@{ Name = $polName; ExcludedUpns = $excluded })
                    } else {
                        $included = New-UpnHashSet
                        Add-UpnsFromUserIds  -Set $included -UserIds $includeUsers
                        Add-UpnsFromGroupIds -Set $included -GroupIds $includeGroups
                        Add-UpnsFromRoleIds  -Set $included -RoleIds $includeRoles
                        # If [ALL_TENANT] marker present, promote to tenant-wide policy
                        if ($included.Contains("[ALL_TENANT]")) {
                            [void]$included.Remove("[ALL_TENANT]")
                            [void]$riskPoliciesIncludeAll.Add(@{ Name = $polName; ExcludedUpns = $excluded })
                        } else {
                            [void]$riskPoliciesScoped.Add(@{ Name = $polName; IncludedUpns = $included; ExcludedUpns = $excluded })
                        }
                    }
                } else {
                    $caPolicyCount++
                    if ($targetsAll) {
                        [void]$caPoliciesIncludeAll.Add(@{ Name = $polName; ExcludedUpns = $excluded })
                    } else {
                        $included = New-UpnHashSet
                        Add-UpnsFromUserIds  -Set $included -UserIds $includeUsers
                        Add-UpnsFromGroupIds -Set $included -GroupIds $includeGroups
                        Add-UpnsFromRoleIds  -Set $included -RoleIds $includeRoles
                        # If [ALL_TENANT] marker present, promote to tenant-wide policy
                        if ($included.Contains("[ALL_TENANT]")) {
                            [void]$included.Remove("[ALL_TENANT]")
                            [void]$caPoliciesIncludeAll.Add(@{ Name = $polName; ExcludedUpns = $excluded })
                        } else {
                            [void]$caPoliciesScoped.Add(@{ Name = $polName; IncludedUpns = $included; ExcludedUpns = $excluded })
                        }
                    }
                }
            } catch {
                # Skip individual policies that fail to parse (unexpected format)
                $pName = if ($p -and $p['displayName']) { $p['displayName'] } else { "unknown" }
                Write-Warning "  Skipped CA policy '$pName': $($_.Exception.Message)"
            }
        }
        $uri = $resp['@odata.nextLink']
    }

    if ($riskPolicyCount -gt 0 -or $caPolicyCount -gt 0) {
        Write-Host "  Conditional Access policies: $caPolicyCount general (P1), $riskPolicyCount risk-based (P2)." -ForegroundColor Green
        Write-Host "    General CA — All-user: $($caPoliciesIncludeAll.Count), scoped: $($caPoliciesScoped.Count)." -ForegroundColor Green
        Write-Host "    Risk-based — All-user: $($riskPoliciesIncludeAll.Count), scoped: $($riskPoliciesScoped.Count)." -ForegroundColor Green
        Write-Log "CA policies: $caPolicyCount general (all-user: $($caPoliciesIncludeAll.Count), scoped: $($caPoliciesScoped.Count)), $riskPolicyCount risk-based (all-user: $($riskPoliciesIncludeAll.Count), scoped: $($riskPoliciesScoped.Count))"
    } else {
        Write-Host "  No enabled Conditional Access policies detected." -ForegroundColor DarkGray
    }
} catch {
    Write-Log "Conditional Access policies retrieval failed" -Level ERROR -ErrorRecord $_
    Write-Warning "  Could not retrieve Conditional Access policies (requires Policy.Read.All): $($_.Exception.Message)"
    [void]$script:skippedDataWarnings.Add("Conditional Access policies — $($_.Exception.Message)")
}

# ── [8c/12] Fetch Intune managed devices (bulk) for shelfware detection ──
Write-Host "`n[8c/12] Fetching Intune managed device inventory ..." -ForegroundColor Cyan
Write-Log "[8c/12] Fetching Intune managed device inventory"
$lkpManagedDeviceCount = @{}   # UPN (lowercase) → [int] enrolled device count
$managedDevicesLoaded = $false
try {
    $mdUri = "v1.0/deviceManagement/managedDevices?`$select=id,userPrincipalName&`$top=999"
    $mdTotal = 0
    while ($mdUri) {
        $mdPage = Invoke-GraphWithRetry -Method GET -Uri $mdUri
        $mdItems = $mdPage['value']
        if ($mdItems) {
            foreach ($dev in $mdItems) {
                $devUpn = if ($dev -is [hashtable] -or $dev -is [System.Collections.IDictionary]) { $dev['userPrincipalName'] } else { $dev.userPrincipalName }
                if ($devUpn) {
                    $devKey = $devUpn.ToString().Trim().ToLower()
                    if ($lkpManagedDeviceCount.ContainsKey($devKey)) {
                        $lkpManagedDeviceCount[$devKey]++
                    } else {
                        $lkpManagedDeviceCount[$devKey] = 1
                    }
                }
            }
            $mdTotal += $mdItems.Count
        }
        $mdUri = $mdPage['@odata.nextLink']
    }
    $managedDevicesLoaded = $true
    Write-Host "  $mdTotal managed device(s) across $($lkpManagedDeviceCount.Count) user(s)." -ForegroundColor Green
    Write-Log "Intune managed devices loaded: $mdTotal device(s), $($lkpManagedDeviceCount.Count) user(s)"
} catch {
    Write-Log "Intune managed devices retrieval failed" -Level WARN -ErrorRecord $_
    Write-Warning "  Could not retrieve managed devices (requires DeviceManagementManagedDevices.Read.All): $($_.Exception.Message)"
    [void]$script:skippedDataWarnings.Add("Intune managed devices — $($_.Exception.Message)")
}

Write-Host "`n[9/12] Building per-user license maps from bulk data ..." -ForegroundColor Cyan
Write-Log "[9/12] Building per-user license maps"

# ── Prepare timestamp early (needed for streaming CSV writers) ──
$ts = (Get-Date).ToString("yyyyMMdd_HHmmss")

$userDisabledPlansMap = @{}  # UPN → HashSet of disabled ServicePlanName strings
$userHasTeamsClient  = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)  # UPNs with TEAMS1 enabled
$servicePlanRowCount = 0
$consolidatedPlans   = [System.Collections.Generic.Dictionary[string,PSCustomObject]]::new([StringComparer]::OrdinalIgnoreCase)  # "upn|sku" → summary row for Excel

# ── Open StreamWriter for service plan detail CSV ──
$spColumns = @('UserPrincipalName','DisplayName','Department','SkuPartNumber','ServicePlanName','ProvisioningStatus')
$planFile = Join-Path $OutputFolder "M365_ServicePlanDetail_$ts.csv"
$planWriter = [System.IO.StreamWriter]::new($planFile, $false, [System.Text.UTF8Encoding]::new($false))
$planWriter.WriteLine(($spColumns | ForEach-Object { '"' + $_ + '"' }) -join ',')

$counter = 0
$licensedUserCount = $lkpAssignedLicenses.Count

# Iterate only licensed users — unlicensed users were already marked during paginated fetch
foreach ($upn in $lkpAssignedLicenses.Keys) {
    $counter++
    if ($counter % 5000 -eq 0 -or $counter -eq $licensedUserCount) {
        Write-Progress -Activity "License mapping (in-memory)" -Status "$counter / $licensedUserCount" `
            -PercentComplete ([math]::Round(($counter / $licensedUserCount) * 100))
    }

    $assigned = $lkpAssignedLicenses[$upn]
    $userObj  = $lkpUserObj[$upn]

    $skuNames = [System.Collections.Generic.HashSet[string]]::new()
    # Track which service plans are disabled/enabled across all SKUs for this user.
    # A plan is only considered "net disabled" if disabled in EVERY SKU that contains it
    # (e.g., AAD_PREMIUM disabled in E3 but enabled in EMS E5 = enabled).
    $userDisabledPlanNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $userEnabledPlanNames  = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($lic in $assigned) {
        $skuId = $lic.skuId
        $partNumber = if ($skuIdToPartNumber.ContainsKey($skuId)) {
            $skuIdToPartNumber[$skuId]
        } else {
            $skuId   # Fallback to GUID if SKU not in tenant inventory
        }
        [void]$skuNames.Add($partNumber)

        # Build service plan detail rows from tenant SKU definition + user's disabled plans
        if ($skuIdToServicePlans.ContainsKey($skuId)) {
            $disabledPlanIds = [System.Collections.Generic.HashSet[string]]::new()
            if ($lic.disabledPlans) {
                foreach ($dp in $lic.disabledPlans) {
                    [void]$disabledPlanIds.Add($dp.ToString())
                }
            }

            foreach ($plan in $skuIdToServicePlans[$skuId]) {
                $planName = $plan.ServicePlanName
                $isDisabled = $disabledPlanIds.Contains($plan.ServicePlanId.ToString())
                $status = if ($isDisabled) { "Disabled" } else { "Success" }

                # Track disabled vs enabled service plan names for effective entitlement detection
                if ($planName) {
                    if ($isDisabled) {
                        [void]$userDisabledPlanNames.Add($planName)
                    } else {
                        [void]$userEnabledPlanNames.Add($planName)
                    }
                }

                $spRow = [PSCustomObject]@{
                    UserPrincipalName  = $upn
                    DisplayName        = if ($userObj) { $userObj.DisplayName } else { "" }
                    Department         = if ($userObj) { $userObj.Department }  else { "" }
                    SkuPartNumber      = $partNumber
                    ServicePlanName    = $planName
                    ProvisioningStatus = $status
                }
                $planWriter.WriteLine((ConvertTo-CsvLine -Row $spRow -Columns $spColumns))
                $servicePlanRowCount++

                # Accumulate consolidated view for Excel (one row per user per SKU)
                $cKey = "$upn|$partNumber"
                if (-not $consolidatedPlans.ContainsKey($cKey)) {
                    $consolidatedPlans[$cKey] = [PSCustomObject]@{
                        UserPrincipalName = $upn
                        DisplayName       = $userObj.DisplayName
                        Department        = $userObj.Department
                        SkuPartNumber     = $partNumber
                        TotalPlans        = [int]0
                        EnabledCount      = [int]0
                        DisabledCount     = [int]0
                        DisabledPlanNames = [System.Collections.Generic.List[string]]::new()
                    }
                }
                $cEntry = $consolidatedPlans[$cKey]
                $cEntry.TotalPlans++
                if ($isDisabled) {
                    $cEntry.DisabledCount++
                    $cEntry.DisabledPlanNames.Add($planName)
                } else {
                    $cEntry.EnabledCount++
                }
            }
        }
    }
    $userLicenseMap[$upn] = ($skuNames | Sort-Object) -join "; "
    # Expand both sets with suiteIncludes String ID aliases before computing net disabled
    # (Flaw #1 fix: ServicePlanName ≠ String ID — e.g. EXCHANGE_S_STANDARD ≠ EXCHANGESTANDARD)
    if ($planNameToStringId.Count -gt 0) {
        foreach ($dpn in [string[]]$userDisabledPlanNames) {
            if ($planNameToStringId.ContainsKey($dpn)) {
                [void]$userDisabledPlanNames.Add($planNameToStringId[$dpn])
            }
        }
        foreach ($epn in [string[]]$userEnabledPlanNames) {
            if ($planNameToStringId.ContainsKey($epn)) {
                [void]$userEnabledPlanNames.Add($planNameToStringId[$epn])
            }
        }
    }
    # Compute net disabled: only plans that are disabled in ALL containing SKUs
    # If a plan is enabled in ANY SKU, remove it from the disabled set
    $userDisabledPlanNames.ExceptWith($userEnabledPlanNames)
    if ($userDisabledPlanNames.Count -gt 0) {
        $userDisabledPlansMap[$upn] = $userDisabledPlanNames
    }
    # Track Teams client (TEAMS1) availability for EEA no-Teams bundle gating
    if ($userEnabledPlanNames.Contains("TEAMS1")) {
        [void]$userHasTeamsClient.Add($upn)
    }
}

Write-Progress -Activity "License mapping (in-memory)" -Completed
$planWriter.Flush()
$planWriter.Close()
$planWriter.Dispose()
$planWriter = $null
Write-Host "  License data mapped for $totalUserCount user(s) ($licensedUserCount licensed), $servicePlanRowCount service plan rows." -ForegroundColor Green

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — Build the merged optimization report
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host "`n[10/12] Building merged license optimization report ..." -ForegroundColor Cyan
Write-Log "[10/12] Building merged license optimization report"

# Collect all known UPNs across all data sources
$allUPNs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($k in $lkpUserObj.Keys)    { [void]$allUPNs.Add($k) }
foreach ($k in $lkpActiveUser.Keys) { [void]$allUPNs.Add($k) }
foreach ($k in $lkpM365App.Keys)    { [void]$allUPNs.Add($k) }

# Release assigned license data — all needed SKU mappings are now in $userLicenseMap.
# The $lkpAssignedLicenses hashtable can be large on 150k+ tenants; reclaim it now.
$lkpAssignedLicenses = $null
[System.GC]::Collect(2, [System.GCCollectionMode]::Optimized)

# ── Tenant-level pre-computations ──

# Tenant license family detection — Enterprise vs Business vs Mixed
# Used to gate cross-family downgrade recommendations and append advisory notes.
# Enterprise → Business downgrades carry operational implications (300-seat cap, Defender alignment,
# different Office app deployments, no CAL rights). Business → Enterprise upgrades are always safe.
$_hasEnterpriseSku = @($subscribedSkus | Where-Object { $_.SkuPartNumber -match 'SPE_E[35]|SPE_F|M365_F1|ENTERPRISEPACK|ENTERPRISEPREMIUM|ENTERPRISEWITHSCAL|STANDARDPACK|DESKLESSPACK|OFFICESUBSCRIPTION' -and $_.ConsumedUnits -gt 0 }).Count -gt 0
$_hasBusinessSku   = @($subscribedSkus | Where-Object { $_.SkuPartNumber -match '^SPB$|^O365_BUSINESS|^SMB_BUSINESS|^MICROSOFT_BUSINESS' -and $_.ConsumedUnits -gt 0 }).Count -gt 0
$tenantFamily      = if ($_hasEnterpriseSku -and -not $_hasBusinessSku) { 'Enterprise' }
                     elseif ($_hasBusinessSku -and -not $_hasEnterpriseSku) { 'Business' }
                     else { 'Mixed' }
$crossFamilyNote   = "Note: Business and Enterprise licenses can coexist but require coordination — 300-seat Business cap, separate app deployment channels, and endpoint configuration alignment apply. Consult your Microsoft partner."
Write-Host "  Tenant license family: $tenantFamily" -ForegroundColor Cyan
Write-Log "Tenant family: $tenantFamily, Total UPNs: $($allUPNs.Count), Licensed: $licensedUserCount"

# Business 300-seat limit check — aggregate at family level (shared cap per Microsoft docs)
# The 300-seat maximum applies to ALL Business User Subscription Suites collectively, not per-SKU.
$businessLimitWarnings = [System.Collections.Generic.List[string]]::new()
$businessFamilyTotalConsumed = 0
$businessFamilyDetails       = [System.Collections.Generic.List[string]]::new()
foreach ($sku in $subscribedSkus) {
    if ($sku.SkuPartNumber -in $businessFamilySkus) {
        $total    = $sku.PrepaidUnits.Enabled + $sku.PrepaidUnits.Warning
        $consumed = $sku.ConsumedUnits
        $businessFamilyTotalConsumed += $consumed
        if ($consumed -gt 0) {
            $businessFamilyDetails.Add("$(Resolve-SkuFriendlyName $sku.SkuPartNumber): $consumed/$total")
        }
    }
}
if ($businessFamilyTotalConsumed -gt 0) {
    # The 300-seat cap is an absolute limit across the entire Business family
    $familyPct = [math]::Round(($businessFamilyTotalConsumed / 300) * 100)
    if ($businessFamilyTotalConsumed -ge 255) {  # 85% of 300
        $businessLimitWarnings.Add("  Business family aggregate: $businessFamilyTotalConsumed/300 used (${familyPct}%) — across $($businessFamilyDetails.Count) SKU(s)")
        foreach ($detail in $businessFamilyDetails) {
            $businessLimitWarnings.Add("    $detail")
        }
        $businessLimitWarnings.Add("  Action: Plan migration to Enterprise (E3/E5) before hitting the 300-seat cap.")
    }
}

# ── E1 → Business Basic arbitrage pre-computation ──
# STANDARDPACK (O365 E1, ~€8.70) and Business Basic (~€6.00) have identical web/mobile capabilities.
# If migrating all E1 users to Business Basic keeps the business family under 250 seats (safety
# buffer below 300 cap), flag every E1 user as a downgrade candidate.
$standardpackConsumed = 0
foreach ($sku in $subscribedSkus) {
    if ($sku.SkuPartNumber -eq 'STANDARDPACK' -and $sku.CapabilityStatus -eq 'Enabled') {
        $standardpackConsumed += $sku.ConsumedUnits
    }
}
$e1ToBasicEligible = ($standardpackConsumed -gt 0 -and ($businessFamilyTotalConsumed + $standardpackConsumed) -le 250)

# SPE_E3 (M365 E3, ~€34.90) → SPB (M365 Business Premium, ~€22.60): same desktop apps + Intune + better security.
# Business Premium has 50 GB mailbox limit (vs 100 GB on E3) and 300-seat cap.
$speE3Consumed = 0
foreach ($sku in $subscribedSkus) {
    if ($sku.SkuPartNumber -eq 'SPE_E3' -and $sku.CapabilityStatus -eq 'Enabled') {
        $speE3Consumed += $sku.ConsumedUnits
    }
}
$e3ToBpEligible = ($speE3Consumed -gt 0 -and ($businessFamilyTotalConsumed + $speE3Consumed) -le 250)

# Apps for Enterprise (OFFICESUBSCRIPTION) → Apps for Business (O365_BUSINESS): identical desktop apps, cheaper.
# Must account for all migrating seats against the 300-seat Business cap.
$appsEntConsumed = 0
foreach ($sku in $subscribedSkus) {
    if ($sku.SkuPartNumber -eq 'OFFICESUBSCRIPTION' -and $sku.CapabilityStatus -eq 'Enabled') {
        $appsEntConsumed += $sku.ConsumedUnits
    }
}
$appsEntToBizEligible = ($appsEntConsumed -gt 0 -and ($businessFamilyTotalConsumed + $appsEntConsumed) -le 250)

# Power BI Premium Capacity detection (consumers can use Free; creators still need Pro)
$hasPbiPremiumCapacity = $false
foreach ($sku in $subscribedSkus) {
    if ($sku.SkuPartNumber -match '^(PBI_PREMIUM|BI_AZURE_P|POWER_BI_PREMIUM)' -and $sku.CapabilityStatus -eq 'Enabled') {
        $hasPbiPremiumCapacity = $true
        break
    }
}

# ── B2B guest Entra P1/P2 coverage (1:5 ratio) ──
# Microsoft External ID licensing: for every Entra ID P1/P2 license assigned to a member,
# up to 5 B2B guest users are covered for premium features (CA, MFA, Identity Protection).
# This prevents false-positive CA licensing checks on guests who are already covered.
$tenantP1P2ConsumedSeats = 0
foreach ($sku in $subscribedSkus) {
    if ($sku.CapabilityStatus -ne 'Enabled') { continue }
    $hasP1 = $false
    if ($sku.ServicePlans) {
        foreach ($sp in $sku.ServicePlans) {
            if ($sp.ServicePlanName -in @('AAD_PREMIUM','AAD_PREMIUM_P2')) { $hasP1 = $true; break }
        }
    }
    if ($hasP1) { $tenantP1P2ConsumedSeats += $sku.ConsumedUnits }
}
# Deduplicate: if a user has both E5 + standalone P2, they count once. ConsumedUnits per SKU
# may overcount, but for the 1:5 ratio a conservative (higher) member count is safe — it only
# increases the guest capacity, so false positives are still suppressed correctly.
$b2bGuestCapacity = $tenantP1P2ConsumedSeats * 5
# Count actual guest users in the tenant
$tenantGuestCount = 0
foreach ($uKey in $lkpUserObj.Keys) {
    $uObj = $lkpUserObj[$uKey]
    if ($uObj -and $uObj.UserType -eq 'Guest') { $tenantGuestCount++ }
}
$b2bGuestsCovered = ($tenantGuestCount -le $b2bGuestCapacity -and $tenantP1P2ConsumedSeats -gt 0)

# ── Unassigned License Pool Waste detection ──
# Flag ALL paid SKUs with any unassigned seats (price > 0, unassigned > 0)
$unassignedPoolWarnings = [System.Collections.Generic.List[PSCustomObject]]::new()
[decimal]$unassignedPoolTotalAnnual = 0
foreach ($sku in $subscribedSkus) {
    if ($sku.CapabilityStatus -ne 'Enabled') { continue }
    # Skip Company-level SKUs (tenant-wide capacity, not per-user assignments — ConsumedUnits is unreliable)
    if ($sku.AppliesTo -eq 'Company') { continue }
    $total    = $sku.PrepaidUnits.Enabled + $sku.PrepaidUnits.Warning
    $consumed = $sku.ConsumedUnits
    if ($total -le 0) { continue }
    $unassigned = $total - $consumed
    if ($unassigned -le 0) { continue }
    $monthlyPrice = Get-SkuMonthlyPrice $sku.SkuPartNumber
    if ($monthlyPrice -le 0) { continue }
    $unassignedPct   = [math]::Round($unassigned / $total * 100, 1)
    $annualWaste     = [math]::Round($monthlyPrice * 12 * $unassigned, 2)
    $unassignedPoolWarnings.Add([PSCustomObject]@{
        SkuPartNumber  = $sku.SkuPartNumber
        FriendlyName   = Resolve-SkuFriendlyName $sku.SkuPartNumber
        Total          = $total
        Consumed       = $consumed
        Unassigned     = $unassigned
        UnassignedPct  = $unassignedPct
        MonthlyWaste   = [math]::Round($monthlyPrice * $unassigned, 2)
        AnnualWaste    = $annualWaste
    })
    $unassignedPoolTotalAnnual += $annualWaste
}
if ($unassignedPoolWarnings.Count -gt 0) {
    Write-Host "  $($unassignedPoolWarnings.Count) SKU(s) with significant unassigned license pool waste (€$($unassignedPoolTotalAnnual.ToString('N2'))/yr)" -ForegroundColor DarkYellow
}

# ── Unassigned License Inventory (all enabled SKUs with spare seats) ──
# Simple per-SKU unassigned count for the executive summary — no thresholds, no price filter.
$unassignedLicenseInventory = [System.Collections.Generic.List[PSCustomObject]]::new()
$totalUnassignedSeats = 0
foreach ($sku in $subscribedSkus) {
    if ($sku.CapabilityStatus -ne 'Enabled') { continue }
    # Skip Company-level SKUs (tenant-wide capacity — ConsumedUnits not tracked per-user)
    if ($sku.AppliesTo -eq 'Company') { continue }
    $total    = $sku.PrepaidUnits.Enabled + $sku.PrepaidUnits.Warning
    $consumed = $sku.ConsumedUnits
    if ($total -le 0) { continue }
    $unassigned = $total - $consumed
    if ($unassigned -le 0) { continue }
    $unassignedLicenseInventory.Add([PSCustomObject]@{
        SkuPartNumber = $sku.SkuPartNumber
        FriendlyName  = Resolve-SkuFriendlyName $sku.SkuPartNumber
        Total         = $total
        Consumed      = $consumed
        Unassigned    = $unassigned
    })
    $totalUnassignedSeats += $unassigned
}

# ── Teams Rooms Basic vs Pro optimization ──
# Teams Rooms Basic is free (up to 25 per tenant). If total room count ≤ 25 and tenant
# is paying for Rooms Pro, all rooms could use Basic instead (saving ~€37.40/mo each).
$teamsRoomsProSkus   = @('Microsoft_Teams_Rooms_Pro','MEETING_ROOM','MTR_PREM')
$teamsRoomsBasicSkus = @('Microsoft_Teams_Rooms_Basic')
$teamsRoomsProConsumed  = 0
$teamsRoomsBasicConsumed = 0
[decimal]$teamsRoomsProMonthly = 0
foreach ($sku in $subscribedSkus) {
    if ($sku.SkuPartNumber -in $teamsRoomsProSkus -and $sku.CapabilityStatus -eq 'Enabled') {
        $teamsRoomsProConsumed += $sku.ConsumedUnits
        $teamsRoomsProMonthly  += (Get-SkuMonthlyPrice $sku.SkuPartNumber) * $sku.ConsumedUnits
    }
    if ($sku.SkuPartNumber -in $teamsRoomsBasicSkus -and $sku.CapabilityStatus -eq 'Enabled') {
        $teamsRoomsBasicConsumed += $sku.ConsumedUnits
    }
}
$teamsRoomsTotalDevices = $teamsRoomsProConsumed + $teamsRoomsBasicConsumed
$teamsRoomsDowngrade = $null
# Teams Rooms Basic is free up to 25 per tenant. Calculate how many Pro rooms can be
# downgraded to Basic (remaining free slots = 25 - current Basic count, capped at Pro count).
$eligibleDowngrades = [math]::Max(0, [math]::Min($teamsRoomsProConsumed, 25 - $teamsRoomsBasicConsumed))
if ($eligibleDowngrades -gt 0) {
    # Per-unit average Pro price (weighted across possibly multiple Pro SKUs)
    $proPerUnit = if ($teamsRoomsProConsumed -gt 0) { $teamsRoomsProMonthly / $teamsRoomsProConsumed } else { 0 }
    $teamsRoomsAnnualSavings = [math]::Round($proPerUnit * $eligibleDowngrades * 12, 2)
    $teamsRoomsDowngrade = [PSCustomObject]@{
        ProRooms       = $teamsRoomsProConsumed
        EligibleRooms  = $eligibleDowngrades
        BasicRooms     = $teamsRoomsBasicConsumed
        TotalRooms     = $teamsRoomsTotalDevices
        AnnualSavings  = $teamsRoomsAnnualSavings
    }
    Write-Host "  Teams Rooms: $eligibleDowngrades of $teamsRoomsProConsumed Pro license(s) can use free Basic ($teamsRoomsBasicConsumed + $eligibleDowngrades / 25 cap) — €$($teamsRoomsAnnualSavings.ToString('N2'))/yr savings" -ForegroundColor DarkYellow
}

# ── CSV column order for main report (must match PSCustomObject property names) ──
$csvColumns = @(
    'User Principal Name', 'Display Name', 'Assigned Licenses', 'License Friendly Names',
    'License Assignment', 'Overlapping Licenses', 'License Groups', 'License Errors', 'Last License Change',
    'Monthly License Cost (EUR)', 'Annual License Cost (EUR)', 'Department', 'Company', 'Country',
    'User Type', 'Account Enabled', 'Mailbox Type', 'Litigation Hold', 'Admin Roles', 'Admin Privilege Level',
    'PIM Eligible Roles', 'PIM Active Roles', 'Risk-based CA Policies', 'MDO Policy Coverage',
    'Uses Desktop Apps', 'No Desktop Apps', 'Uses Mobile Only',
    'Desktop Apps Used', 'Web Apps Used', 'Mobile Apps Used', 'Activated Platforms',
    'Activated Products', 'Exchange: Sent', 'Exchange: Received', 'Exchange: Read',
    'Exchange Intensity', 'Mailbox Size (MB)', 'Mailbox Item Count', 'Has Archive Mailbox',
    'Email Clients Used', 'No Outlook Desktop', 'Teams: Team Chat', 'Teams: Private Chat',
    'Teams: Calls', 'Teams: Meetings', 'Teams: Meetings Organized', 'Teams Intensity', 'Teams Platforms', 'Teams No Desktop',
    'OneDrive: Files', 'OneDrive: Synced', 'OneDrive: Shared', 'OneDrive Intensity',
    'OneDrive Storage (MB)', 'OneDrive File Count', 'SharePoint: Files', 'SharePoint: Shared',
    'SharePoint: Pages', 'SharePoint Intensity', 'Exchange Last Activity', 'OneDrive Last Activity',
    'SharePoint Last Activity', 'Teams Last Activity', 'Has Exchange License', 'Has Teams License',
    'Has OneDrive License', 'Has SharePoint License', 'Last Sign-In', 'Days Since Sign-In',
    'Last Non-Interactive Sign-In', 'Days Since Non-Interactive Sign-In',
    'Dormant Account', 'Trial License', 'Cloud License Errors', 'Has Unknown SKU',
    'Disabled Plans', 'Missing Data Sources',
    'Copilot Active Apps', 'Copilot Last Activity',
    'Cloud PC Type', 'Cloud PC Total Hours (90d)', 'Cloud PC Days Since Sign-In', 'Cloud PC Last Active', 'Cloud PC Device Name', 'Cloud PC Status',
    'Security Coverage', 'Compliance Coverage',
    'Archive Status', 'Auto-Expanding Archive',
    'Recommendation', 'Recommendation Category', 'Recommendation Confidence'
)

# ── Open StreamWriter for main CSV ──
$mainFile = Join-Path $OutputFolder "M365_LicenseOptimization_$ts.csv"
$mainWriter = [System.IO.StreamWriter]::new($mainFile, $false, [System.Text.UTF8Encoding]::new($false))
$mainWriter.WriteLine(($csvColumns | ForEach-Object { if ($_ -match '[",]') { '"' + $_ + '"' } else { $_ } }) -join ',')

# ── Inline summary accumulators (computed during merge loop — no post-loop pass needed) ──
$totalUsers = 0
$noActivity = 0; $noDesktopCount = 0; $mobileOnly = 0; $unlicensed = 0; $lowExchange = 0
$dormantUsers = 0; $neverSignedIn = 0; $noOutlookDesktopCount = 0; $teamsNoDesktopCount = 0; $sharedMbx = 0; $sharedMbxRemovable = 0; $litigationHold = 0
$roomEquipMbx = 0; $adminUsers = 0; $guestsLicensed = 0; $overlapping = 0; $disabledLicensed = 0; $forwardingWaste = 0; $forwardingReview = 0
$duplicateCov = 0; $e5Upgrade = 0; $shelfware = 0; $phoneNoPlan = 0; $copilotUsers = 0
$pbiProReview = 0; $frontlineCandidate = 0; $exoPlan2Review = 0; $licensingCheck = 0; $licensingCheckCA = 0; $licensingCheckMDO = 0; $licensingCheckPIM = 0; $licensingCheckFrontline = 0
$securityGap = 0; $defenderUpsell = 0; $purviewUpsell = 0; $licenseErrors = 0; $bundleConsolidation = 0
$trialLicenseUsers = 0; $capacityQueueUsers = 0; $businessDowngrade = 0; $e1Downgrade = 0; $o365E3Downgrade = 0; $e3Downgrade = 0; $e5VoiceWaste = 0; $appArbitrage = 0; $ppuArbitrage = 0; $callingPlanWaste = 0; $odPlan2Waste = 0; $entraP2Downgrade = 0; $exoKioskDowngrade = 0; $bizPremInversion = 0; $frontlineRescue = 0; $dataGapUsers = 0
$missingSourceUsers = 0; $frontlineReview = 0; $frontlineBlocked = 0; $businessReview = 0
$mailboxStorageWarning = 0; $copilotPrereq = 0; $copilotStudioUsers = 0; $copilotNonAdopter = 0; $copilotReclaim = 0; $copilotWatchlist = 0; $copilotKeep = 0
$oneDriveStorageWarning = 0; $unlicensedWithData = 0; $disabledFreeSku = 0
$aiOverlapReview = 0; $entraSuiteOverlap = 0; $teamsUnbundling = 0
$guestAccountWaste = 0; $intuneSuiteWaste = 0; $nonHumanWaste = 0
$dormantAdminRisk = 0; $viralCleanup = 0; $windowsLicenseWaste = 0; $overLicensedArchive = 0
$groupDirectOverlap = @{}   # Key: GroupId → count of members who also have a direct assignment for a SKU the group assigns
$standaloneAppsWaste = 0; $f3ToF1Downgrade = 0; $highRiskSharing = 0
$legacyServiceAccount = 0; $automationAccount = 0; $alaCarteWaste = 0; $redundantArchive = 0; $shelfwareReview = 0
$inactiveMailbox = 0; $expensiveColdStorage = 0; $mdmMamWaste = 0; $intuneShelfware = 0
$backgroundSyncOnly = 0; $frontlineAddonBloat = 0
$dormantCloudPc = 0; $cloudPcReview = 0
$bundleInefficiency = 0; $premiumAddonWaste = 0; $teamsPhoneRightSizing = 0; $bizPremSecReview = 0
$suiteInversion = 0; $aiAddonOverlap = 0; $e5DataHoarder = 0; $inactiveHold = 0; $seededVisioOverlap = 0
# Security/Compliance posture counters
$secCoverageNone = 0; $secCoverageBasic = 0; $secCoverageAdvanced = 0; $secCoverageE5 = 0
$compCoverageNone = 0; $compCoverageBasic = 0; $compCoverageAdvanced = 0; $compCoverageE5 = 0

# Cost accumulators ([decimal] to avoid IEEE 754 floating-point drift on large tenants)
[decimal]$totalMonthlySpendAcc = 0; [decimal]$dormantCostAcc = 0; $dormantTier1Count = 0; [decimal]$disabledCostAcc = 0
[decimal]$noActivityCostAcc = 0; [decimal]$shelfwareCostAcc = 0
# NOTE: $copilotNonAdopterCostAcc is accumulated but not consumed in financial outputs — kept for future use
[decimal]$copilotNonAdopterCostAcc = 0; [decimal]$copilotReclaimCostAcc = 0; [decimal]$copilotWatchlistCostAcc = 0; [decimal]$sharedMbxCostAcc = 0; [decimal]$frontlineCostAcc = 0
# Executive Summary accumulators
[decimal]$duplicateCostAcc = 0; [decimal]$frontlineSavingsAcc = 0; [decimal]$businessBasicSavingsAcc = 0; [decimal]$e1DowngradeSavingsAcc = 0; [decimal]$o365E3DowngradeSavingsAcc = 0; [decimal]$e3DowngradeSavingsAcc = 0; [decimal]$e5VoiceSavingsAcc = 0; [decimal]$appArbitrageSavingsAcc = 0; [decimal]$ppuArbitrageSavingsAcc = 0; [decimal]$exoKioskSavingsAcc = 0; [decimal]$bizPremInversionSavingsAcc = 0; [decimal]$frontlineRescueSavingsAcc = 0
[decimal]$exoPlan2SavingsAcc = 0; [decimal]$e5UpgradeSavingsAcc = 0; [decimal]$bundleConsolidationSavingsAcc = 0

# Cost-by-dimension running dictionaries
$deptCostDict    = @{}   # Department → @{ Users = 0; AnnualCost = [decimal]0 }

# Recommendation distribution counter
$recDistribution = @{}   # RecCategory → @{ Count = 0; AnnualCost = [decimal]0 }

# Intensity cross-tab counter
$intensityCrossDict = @{}   # "ExchIntensity,TeamsIntensity" → count

# Pre-computed lookup arrays (avoid per-iteration array concatenation)
# Copilot requires E3/E5, Business Standard, or Business Premium as base license.
# Business Basic and Apps-only SKUs are NOT valid — do NOT include $businessNoSecurity.
$copilotBaseSkus  = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]($premiumSuites + $businessStandardSkus + @("SPB","SMB_BUSINESS_PREMIUM")),
    [StringComparer]::OrdinalIgnoreCase)
$oneDrive1TBSkus  = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]($businessNoSecurity + $businessPremiumSkus + @("STANDARDPACK")),
    [StringComparer]::OrdinalIgnoreCase)
$identityOnlySkus = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@("AAD_PREMIUM","AAD_PREMIUM_P2","IDENTITY_THREAT_PROTECTION","EMSPREMIUM"),
    [StringComparer]::OrdinalIgnoreCase)

# High-privilege admin roles: write access to identity, security, data, or tenant config.
# Roles NOT in this set are treated as low-privilege ($isLowPrivAdmin = $true).
$highPrivRoles = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@(
        # Identity & Access Management
        "Global Administrator",
        "Privileged Role Administrator",
        "Privileged Authentication Administrator",
        "Authentication Administrator",
        "Authentication Extensibility Administrator",
        "Authentication Extensibility Password Administrator",
        "Authentication Policy Administrator",
        "Conditional Access Administrator",
        "Identity Governance Administrator",
        "User Administrator",
        "Password Administrator",
        "Helpdesk Administrator",
        "Groups Administrator",
        "Hybrid Identity Administrator",
        "External Identity Provider Administrator",
        "External ID User Flow Administrator",
        "External ID User Flow Attribute Administrator",
        "B2C IEF Keyset Administrator",
        "B2C IEF Policy Administrator",
        "Permissions Management Administrator",
        "Cloud Device Administrator",
        "Directory Writers",
        "Domain Name Administrator",
        "Directory Synchronization Accounts",
        # Application & Cloud
        "Application Administrator",
        "Cloud Application Administrator",
        "Cloud App Security Administrator",
        "Azure DevOps Administrator",
        "Microsoft Graph Data Connect Administrator",
        # Security & Compliance
        "Security Administrator",
        "Security Operator",
        "Compliance Administrator",
        "Compliance Data Administrator",
        "Azure Information Protection Administrator",
        "Attack Simulation Administrator",
        "Customer Lockbox Access Approver",
        "Global Secure Access Administrator",
        "Global Reader",
        # Workload Admins (write access to Exchange/SharePoint/Teams/Intune)
        "Exchange Administrator",
        "SharePoint Administrator",
        "SharePoint Advanced Management Administrator",
        "Teams Administrator",
        "Teams Communications Administrator",
        "Intune Administrator",
        "Skype for Business Administrator",
        # Finance & Licensing
        "Billing Administrator",
        "License Administrator",
        # Backup & Data
        "Exchange Backup Administrator",
        "SharePoint Backup Administrator",
        "Microsoft 365 Backup Administrator",
        # Attribute & Provisioning (write)
        "Attribute Assignment Administrator",
        "Attribute Definition Administrator",
        "Attribute Log Administrator",
        "Attribute Provisioning Administrator"
    ), [StringComparer]::OrdinalIgnoreCase)

try {
foreach ($upn in $allUPNs) {
    # NOTE: $upn is already lowercase — all lookup keys are normalized with .Trim().ToLower()
    # at insertion time (Build-UPNLookup, user fetch, EXO fetch, admin roles, etc.)

    # ── Active User Detail (last activity dates, license flags) ──
    $au = $lkpActiveUser[$upn]

    # Deleted users have no license impact — Entra ID auto-strips licenses on deletion.
    # Skip entirely regardless of -IncludeDisabledAccounts (that flag is for disabled-but-existing accounts).
    # Check 1: Active User report flags the user as deleted
    $isSoftDeleted = ($au -and $au.PSObject.Properties['Is Deleted'] -and $au.'Is Deleted' -in @('True','Yes'))
    if ($isSoftDeleted) { continue }
    # Check 2: user not returned by Graph /users endpoint — already purged or soft-deleted from Entra.
    # Usage reports can lag behind deletions by up to 48h, so the UPN may still be in $lkpActiveUser
    # without the 'Is Deleted' flag. If Graph doesn't know the user, skip them.
    if (-not $lkpUserObj.ContainsKey($upn)) { continue }

    # ── License data ──
    $assignedSkus = if ($userLicenseMap.ContainsKey($upn)) {
        $userLicenseMap[$upn]
    } else {
        "[NOT IN DIRECTORY]"
    }

    # ── M365 Apps platform detail ──
    $app = $lkpM365App[$upn]

    # Parse platform × app booleans from the M365 App report
    $desktopApps = @()
    $webApps     = @()
    $mobileApps  = @()
    $usesDesktop = $false
    $usesWeb     = $false
    $usesMobile  = $false

    if ($app) {
        # (Flaw #4 fix: Teams removed — $usesDesktop now reflects Office apps only.
        #  Teams desktop is tracked separately via $teamsUsesDesktop from Teams Device report.)
        $appNames = @("Outlook", "Word", "Excel", "PowerPoint", "OneNote")
        foreach ($a in $appNames) {
            $onWin    = $app."$a (Windows)" -in @('True','Yes')
            $onMac    = $app."$a (Mac)"     -in @('True','Yes')
            $onMobile = $app."$a (Mobile)"  -in @('True','Yes')
            $onWeb    = $app."$a (Web)"     -in @('True','Yes')

            if ($onWin -or $onMac) { $desktopApps += $a; $usesDesktop = $true }
            if ($onWeb)            { $webApps += $a;     $usesWeb = $true }
            if ($onMobile)         { $mobileApps += $a;  $usesMobile = $true }
        }
    }

    $noDesktopApps  = ($usesWeb -and -not $usesDesktop)
    $usesMobileOnly = ($usesMobile -and -not $usesDesktop -and -not $usesWeb)

    # ── Activations summary ──
    $activatedPlatforms = ""
    $activatedProducts  = ""
    if ($lkpActivations.ContainsKey($upn)) {
        $actRows = $lkpActivations[$upn]
        $platforms = [System.Collections.Generic.HashSet[string]]::new()
        $products  = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($ar in $actRows) {
            [void]$products.Add($ar.'Product Type')
            if ([int]($ar.'Windows' -as [int]) -gt 0)          { [void]$platforms.Add("Windows") }
            if ([int]($ar.'Mac' -as [int]) -gt 0)              { [void]$platforms.Add("Mac") }
            if ([int]($ar.'Windows 10 Mobile' -as [int]) -gt 0){ [void]$platforms.Add("Win10Mobile") }
            if ([int]($ar.'iOS' -as [int]) -gt 0)              { [void]$platforms.Add("iOS") }
            if ([int]($ar.'Android' -as [int]) -gt 0)          { [void]$platforms.Add("Android") }
        }
        $activatedPlatforms = ($platforms | Sort-Object) -join "; "
        $activatedProducts  = ($products  | Sort-Object) -join "; "
    }

    # ── Email activity ──
    $em = $lkpEmail[$upn]
    $emailSend    = if ($em) { Parse-NumericField $em.'Send Count'    } else { 0 }
    $emailReceive = if ($em) { Parse-NumericField $em.'Receive Count' } else { 0 }
    $emailRead    = if ($em) { Parse-NumericField $em.'Read Count'    } else { 0 }
    $emailTotal   = $emailSend + $emailReceive
    $emailIntensity = Get-Intensity -Value $emailTotal -Low $ExchangeLowThreshold -High $ExchangeHighThreshold

    # ── Teams activity ──
    $tm = $lkpTeams[$upn]
    $teamsChatMsg     = 0; $teamsPrivateMsg = 0; $teamsCalls = 0; $teamsMeetings = 0
    $teamsMeetingsOrganized = 0
    if ($tm) {
        $teamsChatMsg     = Parse-NumericField $tm.'Team Chat Message Count'
        $teamsPrivateMsg  = Parse-NumericField $tm.'Private Chat Message Count'
        $teamsCalls       = Parse-NumericField $tm.'Call Count'
        $teamsMeetings    = Parse-NumericField $tm.'Meeting Count'
        # Meetings Organized Count — distinct from attended; needed for Teams Premium organizer-driven check
        $teamsMeetingsOrganized = if ($tm.PSObject.Properties.Match('Meetings Organized Count').Count) {
            Parse-NumericField $tm.'Meetings Organized Count'
        } else { 0 }
    }
    $teamsTotal     = $teamsChatMsg + $teamsPrivateMsg + $teamsCalls + $teamsMeetings
    $teamsIntensity = Get-Intensity -Value $teamsTotal -Low $TeamsLowThreshold -High $TeamsHighThreshold

    # ── OneDrive activity ──
    $od = $lkpOneDrive[$upn]
    $odViewed = 0; $odSynced = 0; $odShared = 0
    if ($od) {
        $odViewed   = Parse-NumericField $od.'Viewed Or Edited File Count'
        $odSynced   = Parse-NumericField $od.'Synced File Count'
        $odShared   = (Parse-NumericField $od.'Shared Internally File Count') +
                      (Parse-NumericField $od.'Shared Externally File Count')
    }
    $odTotal = $odViewed + $odSynced + $odShared
    $odIntensity = Get-Intensity -Value $odTotal -Low $OneDriveLowThreshold -High $OneDriveHighThreshold

    # ── SharePoint activity ──
    $sp = $lkpSharePoint[$upn]
    $spViewed = 0; $spShared = 0; $spPages = 0
    if ($sp) {
        $spViewed = Parse-NumericField $sp.'Viewed Or Edited File Count'
        $spShared = (Parse-NumericField $sp.'Shared Internally File Count') +
                    (Parse-NumericField $sp.'Shared Externally File Count')
        $spPages  = Parse-NumericField $sp.'Visited Page Count'
    }
    $spTotal = $spViewed + $spShared + $spPages
    $spIntensity = Get-Intensity -Value $spTotal -Low $SharePointLowThreshold -High $SharePointHighThreshold

    # ── Mailbox usage ──
    $mb = $lkpMailbox[$upn]
    $mbSizeMB    = if ($mb) { [math]::Round((Parse-DoubleField $mb.'Storage Used (Byte)') / 1MB, 1) } else { $null }
    $mbItemCount = if ($mb) { Parse-NumericField $mb.'Item Count' } else { 0 }
    $mbHasArchive = if ($mb) { $mb.'Has Archive' } else { "" }

    # ── Email app usage ──
    $ea = $lkpEmailApp[$upn]
    $emailClients = @()
    $usesOutlookDesktop = $false
    $usesOutlookWeb     = $false
    if ($ea) {
        # Graph CSV returns app-name strings ("ProPlus", "Undetermined") when used, empty when not — NOT "True"/"False"
        if ($ea.'Outlook For Windows')       { $emailClients += "Outlook Windows"; $usesOutlookDesktop = $true }
        if ($ea.'Outlook For Mac')           { $emailClients += "Outlook Mac"; $usesOutlookDesktop = $true }
        if ($ea.'Outlook For Web')           { $emailClients += "OWA"; $usesOutlookWeb = $true }
        if ($ea.'Outlook For Mobile')        { $emailClients += "Outlook Mobile" }
        if ($ea.'Other For Mobile')          { $emailClients += "Other Mobile" }
        if ($ea.'POP3 App')                  { $emailClients += "POP3" }
        if ($ea.'IMAP4 App')                 { $emailClients += "IMAP4" }
        if ($ea.'SMTP App')                  { $emailClients += "SMTP" }
    }
    $emailClientsStr = ($emailClients | Sort-Object) -join "; "
    $noOutlookDesktop = ($usesOutlookWeb -and -not $usesOutlookDesktop)

    # ── OneDrive storage usage ──
    $odu = $lkpODUsage[$upn]
    $odStorageMB  = if ($odu) { [math]::Round((Parse-DoubleField $odu.'Storage Used (Byte)') / 1MB, 1) } else { $null }
    $odFileCount  = if ($odu) { Parse-NumericField $odu.'File Count' } else { 0 }

    # ── Teams device usage ──
    $td = $lkpTeamsDevice[$upn]
    $teamsPlats = @()
    $teamsUsesDesktop = $false
    $teamsUsesWeb     = $false
    $teamsUsesMobile  = $false
    if ($td) {
        # CSV column names vary by tenant/locale — use safe property access for strict mode
        $tdProps = $td.PSObject.Properties.Name
        if ('Used Windows'       -in $tdProps -and $td.'Used Windows'       -in @('True','Yes')) { $teamsPlats += "Windows";  $teamsUsesDesktop = $true }
        if ('Used Mac'           -in $tdProps -and $td.'Used Mac'           -in @('True','Yes')) { $teamsPlats += "Mac";      $teamsUsesDesktop = $true }
        if ('Used Web'           -in $tdProps -and $td.'Used Web'           -in @('True','Yes')) { $teamsPlats += "Web";      $teamsUsesWeb     = $true }
        if ('Used iOS'           -in $tdProps -and $td.'Used iOS'           -in @('True','Yes')) { $teamsPlats += "iOS";      $teamsUsesMobile  = $true }
        if ('Used Android Phone' -in $tdProps -and $td.'Used Android Phone' -in @('True','Yes')) { $teamsPlats += "Android";  $teamsUsesMobile  = $true }
        if ('Used Chrome OS'     -in $tdProps -and $td.'Used Chrome OS'     -in @('True','Yes')) { $teamsPlats += "ChromeOS" }
        if ('Used Linux'         -in $tdProps -and $td.'Used Linux'         -in @('True','Yes')) { $teamsPlats += "Linux";    $teamsUsesDesktop = $true }
    }
    $teamsPlatStr    = ($teamsPlats | Sort-Object) -join "; "
    $teamsNoDesktop  = (-not $teamsUsesDesktop -and $teamsPlats.Count -gt 0)

    # ── Per-user missing data sources (LOA v1.0 spec §2.2) ──
    $missingDataSources = [System.Collections.Generic.List[string]]::new()
    if (-not $au)  { $missingDataSources.Add("ActiveUserDetail") }
    if (-not $app) { $missingDataSources.Add("M365AppPlatform") }
    if (-not $em)  { $missingDataSources.Add("EmailActivity") }
    if (-not $tm)  { $missingDataSources.Add("TeamsActivity") }
    if (-not $od)  { $missingDataSources.Add("OneDriveActivity") }
    if (-not $sp)  { $missingDataSources.Add("SharePointActivity") }
    if (-not $mb)  { $missingDataSources.Add("MailboxUsage") }
    if (-not $ea)  { $missingDataSources.Add("EmailAppUsage") }
    if (-not $odu) { $missingDataSources.Add("OneDriveUsage") }
    if (-not $td)  { $missingDataSources.Add("TeamsDeviceUsage") }
    # EXOConfig is missing if EXO connection failed or was skipped
    if (-not $exoConnected) { $missingDataSources.Add("EXOConfig") }
    $missingDataSourcesStr = if ($missingDataSources.Count -gt 0) { $missingDataSources -join "; " } else { "" }

    # ── Mailbox type ──
    $mailboxType = if ($lkpMailboxType.ContainsKey($upn)) { $lkpMailboxType[$upn] } else { "" }
    $isSharedMailbox    = ($mailboxType -eq "SharedMailbox")
    $isRoomOrEquipment  = ($mailboxType -eq "RoomMailbox" -or $mailboxType -eq "EquipmentMailbox")
    $isPhoneResource    = $false  # computed after $userSkuList is built (~line 3887)
    $isLitigationHold   = $lkpLitigationHold.ContainsKey($upn)
    $archiveStatus        = if ($lkpArchiveStatus.ContainsKey($upn))        { $lkpArchiveStatus[$upn] }        else { "" }
    $autoExpandingArchive = if ($lkpAutoExpandingArchive.ContainsKey($upn)) { $lkpAutoExpandingArchive[$upn] } else { $false }
    $forwardingTarget     = if ($lkpForwardingTarget.ContainsKey($upn))  { $lkpForwardingTarget[$upn] }  else { "" }
    $deliverAndForward    = if ($lkpDeliverAndForward.ContainsKey($upn)) { $lkpDeliverAndForward[$upn] } else { $false }

    # ── Admin roles ──
    $adminRolesStr = ""
    $isAdmin       = $false
    if ($lkpAdminRoles.ContainsKey($upn)) {
        $adminRolesStr = ($lkpAdminRoles[$upn] | Sort-Object) -join ", "
        $isAdmin = $true
    }
    # PIM-eligible admins should also be flagged — a dormant PIM-eligible Global Admin is a security risk
    if (-not $isAdmin -and $lkpPimEligibleRoles -and $lkpPimEligibleRoles.ContainsKey($upn)) {
        $isAdmin = $true
        if (-not $adminRolesStr -and $lkpPimEligibleRoles[$upn]) {
            $adminRolesStr = ($lkpPimEligibleRoles[$upn] | Sort-Object) -join ", "
            $adminRolesStr = "$adminRolesStr (PIM-eligible)"
        }
    }
    # Display helper: show roles in parentheses only when available
    $adminRolesDisplay = if ($adminRolesStr) { " ($adminRolesStr)" } else { "" }

    # Low-privilege admin tier: $isLowPrivAdmin = true when ALL roles are read-only/limited scope.
    # High-priv + low-priv mix → treated as high-priv (full $isAdmin protection).
    $isLowPrivAdmin = $false
    if ($isAdmin) {
        $userAllRoles = [System.Collections.Generic.List[string]]::new()
        if ($lkpAdminRoles.ContainsKey($upn)) { foreach ($r in $lkpAdminRoles[$upn]) { $userAllRoles.Add($r) } }
        if ($lkpPimEligibleRoles -and $lkpPimEligibleRoles.ContainsKey($upn)) { foreach ($r in $lkpPimEligibleRoles[$upn]) { $userAllRoles.Add($r) } }
        if ($lkpPimActiveRoles -and $lkpPimActiveRoles.ContainsKey($upn)) { foreach ($r in $lkpPimActiveRoles[$upn]) { $userAllRoles.Add($r) } }
        $hasHighPriv = $false
        foreach ($r in $userAllRoles) {
            # Unresolvable role GUIDs (fallback from $roleDefName) treated as high-priv for safety
            if ($highPrivRoles.Contains($r) -or $r -match '^[0-9a-fA-F-]{36}$') { $hasHighPriv = $true; break }
        }
        if (-not $hasHighPriv -and $userAllRoles.Count -gt 0) { $isLowPrivAdmin = $true }
    }

    # ── Entra feature usage signals (PIM, Conditional Access, risk-based CA) ──
    $upnLower = $upn.ToLower()

    $pimEligibleRoles = ""
    if ($lkpPimEligibleRoles -and $lkpPimEligibleRoles.ContainsKey($upn)) {
        $pimEligibleRoles = (@($lkpPimEligibleRoles[$upn]) | Sort-Object) -join "; "
    }

    $pimActiveRoles = ""
    if ($lkpPimActiveRoles -and $lkpPimActiveRoles.ContainsKey($upn)) {
        $pimActiveRoles = (@($lkpPimActiveRoles[$upn]) | Sort-Object) -join "; "
    }

    # Risk-based CA (P2)
    $riskPolicyNames = [System.Collections.Generic.List[string]]::new()
    $matchedScopedRiskPolicy = $false

    foreach ($rp in @($riskPoliciesIncludeAll)) {
        if (-not $rp.ExcludedUpns.Contains($upnLower)) {
            [void]$riskPolicyNames.Add($rp.Name)
        }
    }
    foreach ($rp in @($riskPoliciesScoped)) {
        if ($rp.IncludedUpns.Contains($upnLower) -and -not $rp.ExcludedUpns.Contains($upnLower)) {
            [void]$riskPolicyNames.Add($rp.Name)
            $matchedScopedRiskPolicy = $true
        }
    }

    $riskBasedCA = if ($riskPolicyNames.Count -gt 0) { (@($riskPolicyNames) | Sort-Object -Unique) -join "; " } else { "" }

    # General CA (P1) — MFA, device compliance, location-based, app restrictions, etc.
    $caPolicyNames = [System.Collections.Generic.List[string]]::new()
    $matchedScopedCaPolicy = $false

    foreach ($cp in @($caPoliciesIncludeAll)) {
        if (-not $cp.ExcludedUpns.Contains($upnLower)) {
            [void]$caPolicyNames.Add($cp.Name)
        }
    }
    foreach ($cp in @($caPoliciesScoped)) {
        if ($cp.IncludedUpns.Contains($upnLower) -and -not $cp.ExcludedUpns.Contains($upnLower)) {
            [void]$caPolicyNames.Add($cp.Name)
            $matchedScopedCaPolicy = $true
        }
    }

    $generalCA = if ($caPolicyNames.Count -gt 0) { (@($caPolicyNames) | Sort-Object -Unique) -join "; " } else { "" }

    # ── Defender for Office 365 policy coverage (Safe Links/Attachments) ──
    $mdoPolicyCoverage     = ""
    $mdoCoverageNonBuiltIn = $false
    # Merge per-user coverage with [ALL_TENANT] coverage from massive group circuit breaker
    $mdoSrcsSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if ($lkpMdoCoverageByUpn -and $lkpMdoCoverageByUpn.ContainsKey($upn)) {
        foreach ($s in $lkpMdoCoverageByUpn[$upn]) { [void]$mdoSrcsSet.Add($s) }
    }
    if ($script:__mdoAllTenantSources -and $script:__mdoAllTenantSources.Count -gt 0) {
        foreach ($s in $script:__mdoAllTenantSources) { [void]$mdoSrcsSet.Add($s) }
    }
    if ($mdoSrcsSet.Count -gt 0) {
        $srcs = @($mdoSrcsSet | Sort-Object)
        $mdoPolicyCoverage = ($srcs -join "; ")
        $mdoCoverageNonBuiltIn = (@($srcs | Where-Object { $_ -ne "BuiltInProtection" }).Count -gt 0)
        # Customer-friendly summary: "Safe Links, Safe Attachments, Anti-Phishing" instead of raw policy names
        $mdoPolicyTypes = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($src in $srcs) {
            if ($src -match '^SafeLinks:')       { [void]$mdoPolicyTypes.Add("Safe Links") }
            elseif ($src -match '^SafeAttach:')   { [void]$mdoPolicyTypes.Add("Safe Attachments") }
            elseif ($src -match '^AntiPhish:')    { [void]$mdoPolicyTypes.Add("Anti-Phishing") }
            elseif ($src -eq 'BuiltInProtection') { [void]$mdoPolicyTypes.Add("Built-in Protection") }
            else                                  { [void]$mdoPolicyTypes.Add($src) }
        }
        $mdoPolicySummary = ($mdoPolicyTypes | Sort-Object) -join ", "
    } else {
        $mdoPolicySummary = ""
    }

    # MDO domain-scope check for shared mailboxes: suppress compliance warnings when the
    # mailbox's primary SMTP domain is not in any MDO policy's RecipientDomainIs list.
    # Fallback: when policies use groups only (no RecipientDomainIs anywhere), suppress for
    # .onmicrosoft.com — the default tenant routing domain, never a production email domain.
    $mdoSharedDomainOk = $true
    $mdoUpnNote = ""
    if ($isSharedMailbox -and $mdoCoverageNonBuiltIn) {
        $userPrimarySmtp = if ($lkpMailboxPrimarySmtp.ContainsKey($upn)) { $lkpMailboxPrimarySmtp[$upn] } else { $upn }
        $userSmtpDomain = ($userPrimarySmtp -split '@')[-1]
        if ($mdoPolicyDomains.Count -gt 0) {
            $mdoSharedDomainOk = $mdoPolicyDomains.Contains($userSmtpDomain)
        } else {
            $mdoSharedDomainOk = -not ($userSmtpDomain -like '*.onmicrosoft.com')
        }
        # Customer-facing note when UPN domain differs from primary SMTP domain
        $upnDomain = ($upn -split '@')[-1]
        $mdoUpnNote = if ($upnDomain -ne $userSmtpDomain) {
            " Note: This mailbox's sign-in name uses the $upnDomain domain, but its primary email address (@$userSmtpDomain) is covered by MDO policies."
        } else { "" }
    }

    # ── Sign-in activity ──
    $lastSignIn     = ""
    $daysSinceSignIn = ""
    $isDormant       = $false
    $lastNonInteractiveSignIn = ""
    $daysSinceNonInteractive  = ""
    $hasRecentNonInteractive  = $false
    if ($lkpSignIn.ContainsKey($upn)) {
        $sia = $lkpSignIn[$upn]
        $rawSignIn = $sia['lastSignInDateTime']
        if ($rawSignIn) {
            $lastSignIn = ([datetime]$rawSignIn).ToString("yyyy-MM-dd")
            $daysSinceSignIn = [int]((Get-Date) - [datetime]$rawSignIn).TotalDays
            $isDormant = ($daysSinceSignIn -ge $InactiveSignInDays)
        }
        # Non-interactive sign-in (Logic Flaw #3: detect automation/service accounts)
        # signInActivity contains lastNonInteractiveSignInDateTime — apps, service principals,
        # scheduled tasks, and scripts sign in non-interactively without human presence.
        $rawNonInteractive = $sia['lastNonInteractiveSignInDateTime']
        if ($rawNonInteractive) {
            $lastNonInteractiveSignIn = ([datetime]$rawNonInteractive).ToString("yyyy-MM-dd")
            $daysSinceNonInteractive  = [int]((Get-Date) - [datetime]$rawNonInteractive).TotalDays
            $hasRecentNonInteractive  = ($daysSinceNonInteractive -lt $InactiveSignInDays)
        }
    }

    # ── User type & Guest flag ──
    $userObj  = $lkpUserObj[$upn]
    $userType = if ($userObj) { $userObj.UserType } else { "" }
    $isGuest  = ($userType -eq "Guest")
    $isGuestWithLicense = ($isGuest -and $assignedSkus -ne "[UNLICENSED]" -and $assignedSkus -ne "[NOT IN DIRECTORY]")

    # ── License assignment path (Direct / Group / Both) ──
    $licAssignmentStr   = ""
    $overlappingSkus    = ""
    $overlappingPartNums = [System.Collections.Generic.List[string]]::new()
    $licenseGroupsStr   = ""
    $licenseErrorsStr   = ""
    $hasOverlap         = $false
    $lastLicenseChange  = ""
    if ($lkpLicAssignment.ContainsKey($upn)) {
        $states = $lkpLicAssignment[$upn]
        # Track the most recent license change date
        $maxLicDate = $null
        # Group states by skuId
        $bySkuId = @{}
        $allGroups = [System.Collections.Generic.HashSet[string]]::new()
        $licErrors = [System.Collections.Generic.List[string]]::new()
        foreach ($s in $states) {
            $ludt = $s['lastUpdatedDateTime']
            if ($ludt) {
                $parsedDate = [datetime]$ludt
                if ($null -eq $maxLicDate -or $parsedDate -gt $maxLicDate) { $maxLicDate = $parsedDate }
            }
            # Check for license assignment errors (e.g., insufficient licenses, conflicting plans)
            $sState = $s['state']
            $sError = $s['error']
            if ($sState -eq 'Error' -or ($sError -and $sError -ne 'None' -and $sError -ne '')) {
                $errSku = $s['skuId']
                $errSkuObj = $subscribedSkus | Where-Object { $_.SkuId -eq $errSku } | Select-Object -First 1
                $errSkuName = if ($errSkuObj) { $errSkuObj.SkuPartNumber } else { $errSku }
                $errSource = $s['assignedByGroup']
                $errSourceName = if ($errSource) {
                    $gn = if ($groupNameCache.ContainsKey($errSource)) { $groupNameCache[$errSource] } else { $errSource }
                    "Group:$gn"
                } else { "Direct" }
                $licErrors.Add("$errSkuName ($errSourceName): $sError")
            }
            $sid = $s['skuId']
            if (-not $bySkuId.ContainsKey($sid)) { $bySkuId[$sid] = @{ Direct = $false; Group = $false; GroupOk = $false; GroupIds = [System.Collections.Generic.List[string]]::new() } }
            $assignedBy = $s['assignedByGroup']
            $isErrorState = ($sState -eq 'Error' -or ($sError -and $sError -ne 'None' -and $sError -ne ''))
            if ($null -eq $assignedBy) {
                $bySkuId[$sid].Direct = $true
            } else {
                $bySkuId[$sid].Group = $true
                if (-not $isErrorState) { $bySkuId[$sid].GroupOk = $true }
                $bySkuId[$sid].GroupIds.Add($assignedBy)
                $gName = if ($groupNameCache.ContainsKey($assignedBy)) { $groupNameCache[$assignedBy] } else { $assignedBy }
                [void]$allGroups.Add($gName)
            }
        }
        # Determine per-user patterns
        $patterns = [System.Collections.Generic.HashSet[string]]::new()
        $overlapList = [System.Collections.Generic.List[string]]::new()
        foreach ($sid in $bySkuId.Keys) {
            $info = $bySkuId[$sid]
            if ($info.Direct -and $info.Group) {
                [void]$patterns.Add("Both")
                # Only flag as overlap if the group assignment is actually working (not in error state).
                # If group assignment has an error (conflicting plans, insufficient seats), the direct
                # assignment is the user's ONLY working source — removing it would delicense the user.
                if ($info.GroupOk) {
                    $hasOverlap = $true
                    # Resolve SKU name from subscribedSkus
                    $skuObj = $subscribedSkus | Where-Object { $_.SkuId -eq $sid } | Select-Object -First 1
                    $skuName = if ($skuObj) { $skuObj.SkuPartNumber } else { $sid }
                    $overlapList.Add((Resolve-SkuFriendlyName $skuName))
                    $overlappingPartNums.Add($skuName)
                    # Track per-group direct overlap count
                    foreach ($gid in $info.GroupIds) {
                        if (-not $groupDirectOverlap.ContainsKey($gid)) { $groupDirectOverlap[$gid] = 0 }
                        $groupDirectOverlap[$gid]++
                    }
                }
            } elseif ($info.Direct) { [void]$patterns.Add("Direct") }
            elseif ($info.Group)  { [void]$patterns.Add("Group") }
        }
        $licAssignmentStr = ($patterns | Sort-Object) -join " + "
        $overlappingSkus  = ($overlapList | Sort-Object) -join "; "
        $licenseGroupsStr = ($allGroups | Sort-Object) -join "; "
        $licenseErrorsStr = ($licErrors) -join "; "
        if ($maxLicDate) { $lastLicenseChange = $maxLicDate.ToString("yyyy-MM-dd") }
    }

    # ── License friendly names ──
    $licenseFriendlyStr = ""
    if ($assignedSkus -ne "[UNLICENSED]" -and $assignedSkus -ne "[NOT IN DIRECTORY]") {
        $friendlyNames = foreach ($sku in ($assignedSkus -split ";\s*")) {
            Resolve-SkuFriendlyName $sku.Trim()
        }
        $licenseFriendlyStr = ($friendlyNames | Sort-Object) -join "; "
    } else {
        $licenseFriendlyStr = $assignedSkus
    }

    # ── Per-user cost computation ──
    $userSkuList = @($assignedSkus -split ";\s*" | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -ne "[UNLICENSED]" -and $_ -ne "[NOT IN DIRECTORY]" })
    $isPhoneResource = ($userSkuList -contains "PHONESYSTEM_VIRTUALUSER" -and @($userSkuList | Where-Object { -not $freeSkuSet.Contains($_) -and $_ -ne "PHONESYSTEM_VIRTUALUSER" }).Count -eq 0)
    $hasCpcEnterprise = @($userSkuList | Where-Object { $_ -match '^(CPC_E_|Windows_365_E_)' }).Count -gt 0
    [decimal]$userMonthlyCost = 0
    foreach ($sku in $userSkuList) { $userMonthlyCost += Get-SkuMonthlyPrice $sku }
    $userAnnualCost = [math]::Round($userMonthlyCost * 12, 2)
    $userMonthlyCost = [math]::Round($userMonthlyCost, 2)
    [decimal]$dupAnnualWaste = 0   # per-user duplicate SKU cost (deducted from Tier 1 accumulators to prevent overlap)

    # ── Detect unknown SKUs (not in reference data) ──
    $unknownSkus  = @($userSkuList | Where-Object { -not (Test-SkuKnown $_) })
    $hasUnknownSku = $unknownSkus.Count -gt 0

    # ── Effective entitlements (expand suite inclusions, respecting disabled service plans) ──
    $effectiveSkuSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $disabledPlans   = if ($userDisabledPlansMap.ContainsKey($upn)) { $userDisabledPlansMap[$upn] } else { $null }
    foreach ($sku in $userSkuList) {
        [void]$effectiveSkuSet.Add($sku)
        if ($suiteIncludes.ContainsKey($sku)) {
            foreach ($inc in $suiteIncludes[$sku]) {
                # Only add if the service plan is not explicitly disabled by the admin
                if (-not $disabledPlans -or -not $disabledPlans.Contains($inc)) {
                    [void]$effectiveSkuSet.Add($inc)
                }
            }
        }
    }

    # ── Entitlement-derived flags (LOA v1.0 spec §4.1 — use entitlements, not report flags) ──
    $hasExchangeEntitlement = ($effectiveSkuSet.Contains("EXCHANGESTANDARD") -or $effectiveSkuSet.Contains("EXCHANGEENTERPRISE") -or
                               $effectiveSkuSet.Contains("EXCHANGEDESKLESS") -or $effectiveSkuSet.Contains("EXCHANGE_S_DESKLESS"))
    $hasTeamsEntitlement    = $effectiveSkuSet.Contains("TEAMS1")
    $hasOneDriveEntitlement = ($effectiveSkuSet.Contains("SHAREPOINTSTANDARD") -or $effectiveSkuSet.Contains("SHAREPOINTENTERPRISE") -or
                               $effectiveSkuSet.Contains("SHAREPOINTDESKLESS") -or $effectiveSkuSet.Contains("ONEDRIVESTANDARD"))

    $hasEntraP1          = ($effectiveSkuSet.Contains("AAD_PREMIUM") -or $effectiveSkuSet.Contains("AAD_PREMIUM_P2"))
    # Entra P2 features (PIM, risk-based CA) are also available via Entra ID Governance and Entra Suite
    $hasEntraP2          = ($effectiveSkuSet.Contains("AAD_PREMIUM_P2") -or
                            $effectiveSkuSet.Contains("ENTRA_ID_GOVERNANCE") -or
                            $effectiveSkuSet.Contains("ENTRA_SUITE"))
    $hasDefenderForO365  = ($effectiveSkuSet.Contains("ATP_ENTERPRISE") -or $effectiveSkuSet.Contains("THREAT_INTELLIGENCE") -or
                            $effectiveSkuSet.Contains("MDO_P1_FLW") -or $effectiveSkuSet.Contains("MDO_P2_FLW"))

    # ── User org properties ──
    $userDept    = if ($userObj -and $userObj.Department)  { $userObj.Department }  else { "" }
    $userCompany = if ($userObj -and $userObj.CompanyName) { $userObj.CompanyName } else { "" }
    $userCountry = if ($userObj -and $userObj.Country)     { $userObj.Country }     else { "" }

    # ── License Recommendation Logic ──
    $recommendations = [System.Collections.Generic.List[string]]::new()
    $hasAnyActivity  = $false
    [decimal]$userCopilotAnnualCost = 0
    [decimal]$userShelfwareCost     = 0
    [decimal]$userNoActRightsizeSave = 0   # net savings when NO ACTIVITY triggers right-sizing (E1 + add-ons) instead of full removal
    $userIsOnTrial   = $false
    $userCloudErrors = ""
    # Defaults for capability flags set inside if($isLicensed) — needed for coverage-level columns
    $hasFullDefenderStack = $false; $hasFullPurviewStack = $false
    $hasAnyDefenderCap = $false; $hasAnyPurviewCap = $false
    $_caConsolidated = $false  # defensive forward-declaration; set properly inside if(-not $isLicensed)
    $isLicensed  = ($assignedSkus -ne "[UNLICENSED]" -and $assignedSkus -ne "[NOT IN DIRECTORY]")
    $isAccountEnabled = if ($userObj) { $userObj.AccountEnabled } else { $true }

    # Without -IncludeDisabledAccounts, skip disabled users that have no license
    # (nothing to flag). Disabled+licensed users are always processed for waste detection.
    if (-not $IncludeDisabledAccounts -and -not $isAccountEnabled -and -not $isLicensed) {
        continue
    }

    if (-not $isLicensed) {
        if ($isSharedMailbox) {
            # Shared mailboxes don't need a license under 50 GB — this is expected.
            # But we must still warn if the mailbox is APPROACHING the 50 GB cap,
            # because $hasExchangeEntitlement is $false for unlicensed mailboxes and
            # the normal Plan 1 warning block (further below) would be skipped entirely.
            if ($null -ne $mbSizeMB -and $mbSizeMB -ge 46080) {
                $pctUsed = [math]::Round($mbSizeMB / 51200 * 100, 0)
                $mbSizeDisplay = if ($mbSizeMB -ge 1024) { "$([math]::Round($mbSizeMB / 1024, 1)) GB" } else { "${mbSizeMB} MB" }
                $recommendations.Add("MAILBOX STORAGE WARNING — unlicensed shared mailbox is $mbSizeDisplay (${pctUsed}% of 50 GB limit). Mail flow stops at 50 GB. Consider assigning an Exchange Online Plan 2 license (100 GB + auto-expanding archive) or migrating data to an archive before the cap is reached.")
            }
            # No informational rec for unlicensed shared mailboxes under 50 GB — the [UNLICENSED] badge
            # and Shared Mailbox category (via $isSharedMailbox fallback in $recCategory) already convey the status.
            # Compliance recs (CA/MDO) are added below if applicable.
        } elseif ($isRoomOrEquipment) {
            $recommendations.Add("$mailboxType — no user license assigned. Room/Equipment mailboxes typically require only a Teams Rooms license for booking or calling features.")
        } else {
            # Orphaned data risk: unlicensed user mailbox/OneDrive is purged after 30 days
            # EXCEPTION: Litigation Hold mailboxes silently convert to free Inactive Mailboxes — no purge.
            $hasOrphanedMail = ($null -ne $mbSizeMB -and $mbSizeMB -gt 0)
            $hasOrphanedOD   = ($null -ne $odStorageMB -and $odStorageMB -gt 0)
            if ($hasOrphanedMail -or $hasOrphanedOD) {
                if ($hasOrphanedMail -and $isLitigationHold) {
                    # Lit Hold → Inactive Mailbox: Microsoft preserves data for eDiscovery at no cost
                    $recommendations.Add("INACTIVE MAILBOX (FREE) — Mailbox (${mbSizeMB} MB) is unlicensed but on Litigation Hold. It is safely archived as an 'Inactive Mailbox' for eDiscovery at no cost. No action required for the mailbox.")
                    # OneDrive is NOT protected by Litigation Hold — warn separately if present
                    if ($hasOrphanedOD) {
                        $recommendations.Add("UNLICENSED WITH DATA — OneDrive: ${odStorageMB} MB. Note: Litigation Hold protects the mailbox but NOT OneDrive. Microsoft purges unlicensed OneDrive data after 30 days. Consider backing up or migrating OneDrive content.")
                    }
                } else {
                    $dataDetails = @()
                    if ($hasOrphanedMail) { $dataDetails += "Mailbox: ${mbSizeMB} MB" }
                    if ($hasOrphanedOD)   { $dataDetails += "OneDrive: ${odStorageMB} MB" }
                    $recommendations.Add("UNLICENSED WITH DATA — $($dataDetails -join '; '). Microsoft purges unlicensed mailbox and OneDrive data after 30 days. Consider backing up, converting to a Shared Mailbox, or re-licensing before data is lost.")
                }
            }
        }

        # MDO gap check for unlicensed mailboxes (shared mailboxes commonly hit by MDO policies without entitlement)
        # Room/Equipment mailboxes: suppress MDO compliance — resource accounts are low-risk
        # Regular unlicensed users: suppress CA/MDO compliance — no license to optimize, account cleanup is outside LOA scope
        # When both CA and MDO gaps exist, consolidate into a single recommendation
        $_caConsolidated = $false  # flag to suppress standalone CA rec at line ~4062 when combined rec fires
        $guestCoveredByRatio = ($isGuest -and $b2bGuestsCovered)
        $pimAlreadyNeedsP2 = (($pimEligibleRoles -or $pimActiveRoles) -and -not $hasEntraP2)
        $_unlicRegularUser = (-not $isSharedMailbox -and -not $isRoomOrEquipment)
        $_unlicCaGap  = ($generalCA -and -not $hasEntraP1 -and -not $guestCoveredByRatio -and -not $pimAlreadyNeedsP2 -and -not $isRoomOrEquipment -and -not $isPhoneResource -and -not $_unlicRegularUser)
        $_unlicMdoGap = ($mdoCoverageNonBuiltIn -and $mdoPolicyCoverage -and -not $hasDefenderForO365 -and -not $isRoomOrEquipment -and -not $isPhoneResource -and -not $_unlicRegularUser -and $mdoSharedDomainOk)
        # When both CA and MDO gaps exist, emit a single combined rec and set $_caConsolidated
        # to suppress the standalone CA rec that would otherwise fire at line ~4064.
        if ($_unlicCaGap -and $_unlicMdoGap) {
            $p1Cost = Get-SkuMonthlyPrice "AAD_PREMIUM"
            $p1Annual = [math]::Round($p1Cost * 12, 2)
            $mdoP1Cost = Get-SkuMonthlyPrice "ATP_ENTERPRISE"
            $mdoP1Annual = [math]::Round($mdoP1Cost * 12, 2)
            $combinedMonthlyCost = [math]::Round($p1Cost + $mdoP1Cost, 2)
            $combinedAnnualCost  = [math]::Round($p1Annual + $mdoP1Annual, 2)
            $_caDescUnlic = if ($matchedScopedCaPolicy) { "targeted by $(@($caPolicyNames).Count) Conditional Access $(if (@($caPolicyNames).Count -eq 1) { 'policy' } else { 'policies' })" } else { "covered by $(@($caPolicyNames).Count) tenant-wide Conditional Access policies" }
            $recommendations.Add("LICENSING CHECK — User is $_caDescUnlic and protected by $($mdoPolicyTypes.Count) Defender for Office 365 $(if ($mdoPolicyTypes.Count -eq 1) { 'policy' } else { 'policies' }) but has no Entra ID P1 or MDO entitlement. To ensure compliance, add Entra P1 (€$($p1Cost.ToString('N2'))/mo) + MDO P1 (€$($mdoP1Cost.ToString('N2'))/mo) = €$($combinedMonthlyCost.ToString('N2'))/mo (€$($combinedAnnualCost.ToString('N2'))/yr). Alternatively, exclude this user from the CA and MDO policies to avoid the compliance cost. Note: M365 E3/E5/Business Premium include both Entra P1 and MDO P1.$mdoUpnNote")
            $_caConsolidated = $true
        } elseif ($_unlicMdoGap) {
            $mdoP1Cost = Get-SkuMonthlyPrice "ATP_ENTERPRISE"
            $mdoP1Annual = [math]::Round($mdoP1Cost * 12, 2)
            $recommendations.Add("LICENSING CHECK — Mailbox is protected by $($mdoPolicyTypes.Count) Defender for Office 365 $(if ($mdoPolicyTypes.Count -eq 1) { 'policy' } else { 'policies' }) but no MDO license entitlement was found. Shared mailboxes covered by MDO policies require an Exchange Online Plan 2 or a standalone Defender for Office 365 P1 add-on (€$($mdoP1Cost.ToString('N2'))/mo, €$($mdoP1Annual.ToString('N2'))/yr) to ensure compliance. Alternatively, exclude this mailbox from the MDO policies to avoid the compliance cost.$mdoUpnNote")
        }
        # Note: standalone CA gap (without MDO) is emitted below at line ~4060 (guarded by $_caConsolidated)
    } elseif ($isGuestWithLicense) {
        if ($userAnnualCost -gt 0) {
            $recommendations.Add("GUEST ACCOUNT REVIEW — External/guest user holding a paid license ($licenseFriendlyStr, €$($userAnnualCost.ToString('N2'))/yr). B2B guests can access shared Teams and SharePoint resources via their home tenant license or Entra ID External Identities. Consider removing the license unless it is required for a dedicated mailbox or specific app.")
        } else {
            $recommendations.Add("GUEST USER with free SKU ($licenseFriendlyStr) — no financial waste but review whether the license assignment is intentional.")
        }
    }

    # ── Security checks that apply regardless of license status ──
    # Dormant admin accounts and PIM role holders are security risks even when unlicensed.
    # The dormant/admin checks inside if($isLicensed) only cover licensed users — these catch the rest.
    # Service/automation account detection by role or UPN pattern (unlicensed path)
    $isServiceAccountByPattern = $false
    if ($adminRolesStr -match 'Directory Synchronization Accounts') {
        $isServiceAccountByPattern = $true
    } elseif ($upn -match '^(sync_|adsync|svc[_\-]|service[_\-]|app[_\-]|bot[_\-]|noreply[_\-@]|azure[_\-]|msol_|aadconnect|sharepoint_|crm[_\-]|robot[_\-]|workflow[_\-]|automation[_\-])') {
        $isServiceAccountByPattern = $true
    }
    if (-not $isLicensed -and -not $isSharedMailbox -and -not $isRoomOrEquipment) {
        # Unlicensed service/automation account — flag for awareness even without license cost
        if ($isServiceAccountByPattern -and $isAccountEnabled) {
            $patternSignal = if ($adminRolesStr -match 'Directory Synchronization Accounts') { "Directory Synchronization Accounts role" } else { "service account UPN pattern" }
            $signInDetail = if ($isDormant) { "no interactive sign-in for $daysSinceSignIn days" } elseif ($lastSignIn -eq "") { "no interactive sign-in on record" } else { "last sign-in $lastSignIn" }
            $recommendations.Add("AUTOMATION ACCOUNT — unlicensed $patternSignal detected ($signInDetail). This is an infrastructure/sync service account. No license cost but review whether Conditional Access covers non-interactive flows. Consider converting to a Workload Identity.")
        }
        # Dormant admin risk — unlicensed admin accounts are still high-value compromise targets
        if ($isDormant -and $isAdmin -and $isAccountEnabled) {
            $recommendations.Add("DORMANT ADMIN REVIEW — unlicensed admin account$adminRolesDisplay has not signed in for $daysSinceSignIn days. As a best practice, inactive admin accounts should be reviewed periodically. Consider removing the admin role. If confirmed unused, consider disabling the account.")
        }
        # PIM licensing gap for unlicensed users: suppressed — no license to optimize.
        # PIM compliance is checked for licensed users in the $isLicensed block below.
    }
    # CA P1 gap — ANY unlicensed user (including shared mailboxes) in scope of CA policies needs P1
    # Shared mailboxes behind CA still require P1 licensing per Microsoft guidance.
    # Exception: B2B guest users are covered by the External ID 1:5 ratio — for every P1/P2 member
    # license, up to 5 guests get premium features (CA, MFA, Identity Protection) at no extra cost.
    # Skip if PIM already recommends P2 (which is a superset of P1) — avoid duplicate recommendations.
    $guestCoveredByRatio = ($isGuest -and $b2bGuestsCovered)
    $pimAlreadyNeedsP2 = (($pimEligibleRoles -or $pimActiveRoles) -and -not $hasEntraP2)
    # Suppressed for regular unlicensed users — no license to optimize, account cleanup is outside LOA scope.
    # Only shared mailboxes retain unlicensed CA compliance recs (they actively receive email behind CA policies).
    if (-not $isLicensed -and $isSharedMailbox -and $generalCA -and -not $hasEntraP1 -and -not $guestCoveredByRatio -and -not $pimAlreadyNeedsP2 -and -not $isRoomOrEquipment -and -not $_caConsolidated) {
        if ($matchedScopedCaPolicy) {
            $p1Cost = Get-SkuMonthlyPrice "AAD_PREMIUM"
            $p1Annual = [math]::Round($p1Cost * 12, 2)
            $recommendations.Add("LICENSING CHECK — User is targeted by $(@($caPolicyNames).Count) Conditional Access $(if (@($caPolicyNames).Count -eq 1) { 'policy' } else { 'policies' }) but has no license at all. Conditional Access requires Entra ID P1 (included in M365 E3/E5, M365 Business Premium, M365 F3, or standalone at €$($p1Cost.ToString('N2'))/mo). Alternatively, exclude this user from the CA $(if (@($caPolicyNames).Count -eq 1) { 'policy' } else { 'policies' }) to avoid the compliance cost. Estimated compliance cost: €$($p1Annual.ToString('N2'))/yr")
        } elseif (@($caPolicyNames).Count -gt 0) {
            $p1Cost = Get-SkuMonthlyPrice "AAD_PREMIUM"
            $p1Annual = [math]::Round($p1Cost * 12, 2)
            $recommendations.Add("LICENSING CHECK — User is covered by $(@($caPolicyNames).Count) tenant-wide Conditional Access policies but has no Entra ID P1 entitlement. CA requires P1 for each covered user (standalone: €$($p1Cost.ToString('N2'))/mo, €$($p1Annual.ToString('N2'))/yr, or included in M365 E3/E5/Business Premium/F3). Alternatively, exclude this user from the CA policies to avoid the compliance cost.")
        }
    }

    if ($isLicensed) {

        # ── Room/Equipment resource accounts — tag early so primary category picks it up ──
        $roomLicenseHandled = $false
        if ($isRoomOrEquipment) {
            $roomSkuRx = '^(MEETING_ROOM|Microsoft_Teams_Rooms_|MTR_PREM|PHONESYSTEM_VIRTUALUSER)'
            $userRoomSkus    = @($userSkuList | Where-Object { $_ -match $roomSkuRx })
            $userNonRoomSkus = @($userSkuList | Where-Object { $_ -notmatch $roomSkuRx -and -not $freeSkuSet.Contains($_) })
            if ($userNonRoomSkus.Count -gt 0 -and $userRoomSkus.Count -gt 0) {
                $roomLicenseHandled = $true
                # Has both Teams Rooms + expensive non-room licenses — the non-room licenses are waste
                $nonRoomFriendly = ($userNonRoomSkus | ForEach-Object { Resolve-SkuFriendlyName $_ }) -join "; "
                [decimal]$nonRoomMonthlyCost = 0; foreach ($nrs in $userNonRoomSkus) { $nonRoomMonthlyCost += Get-SkuMonthlyPrice $nrs }
                $nonRoomAnnualCost = [math]::Round($nonRoomMonthlyCost * 12, 2)
                $recommendations.Add("$mailboxType — resource account already has a Teams Rooms license but also carries $nonRoomFriendly. The non-room license(s) can likely be removed. Annual savings: €$($nonRoomAnnualCost.ToString('N2'))")
            } elseif ($userNonRoomSkus.Count -gt 0) {
                # Has expensive license(s) but no Teams Rooms license — consider switching
                $recommendations.Add("$mailboxType — resource account with $licenseFriendlyStr. Room and equipment mailboxes typically only need a Teams Rooms license. Review whether the current license can be replaced with a Teams Rooms Basic or Pro license.")
            }
            # else: Only Teams Rooms / free licenses — correctly licensed, no action needed (no rec emitted)
        }

        # ── #10b Disabled/blocked account still licensed ──
        # Microsoft Inactive Mailbox: removing a license from a disabled/held mailbox converts it
        # to a free Inactive Mailbox that retains ALL content and holds indefinitely.  You do NOT
        # need a license to maintain a hold.  If EXO not connected, we cannot verify — emit REVIEW.
        $expensiveHoldSkus = @("SPE_E5","SPE_E3","ENTERPRISEPACK","ENTERPRISEPREMIUM","SPE_F1","DESKLESSPACK","M365_F1","SPB")
        $sharedMbxHandledMdo = $false   # set when DISABLED SHARED MAILBOX or SHARED MAILBOX already handles MDO guidance
        $nonHumanReviewFired = $false  # set when NON-HUMAN ACCOUNT REVIEW fires — suppresses TEAMS UNBUNDLING (contradictory)
        if (-not $isAccountEnabled) {
            if ($isLitigationHold) {
                $hasExpensiveHoldSku = @($userSkuList | Where-Object { $_ -in $expensiveHoldSkus }).Count -gt 0
                if ($hasExpensiveHoldSku) {
                    $recommendations.Add("INACTIVE HOLD WITH LICENSE — account is disabled but retains $licenseFriendlyStr (€$($userMonthlyCost.ToString('N2'))/mo) because it is on litigation hold. A license is not required to maintain a hold on a departed user. Consider removing the license — Microsoft will automatically convert this to a free 'Inactive Mailbox' that retains ALL content and holds indefinitely for eDiscovery. Annual savings: €$($userAnnualCost.ToString('N2'))")
                } else {
                    $recommendations.Add("INACTIVE HOLD — account is disabled but retains $licenseFriendlyStr (€$($userMonthlyCost.ToString('N2'))/mo) because it is on litigation hold. A license is not required to maintain a hold. Consider removing the license — Microsoft will automatically convert this to a free 'Inactive Mailbox' for eDiscovery. Annual savings: €$($userAnnualCost.ToString('N2'))")
                }
            } elseif (-not $exoConnected) {
                $recommendations.Add("DISABLED ACCOUNT REVIEW ($licenseFriendlyStr) — sign-in is blocked. Cannot check litigation hold status (EXO not connected). Review whether active holds exist before removing the license to avoid data loss. Annual cost: €$($userAnnualCost.ToString('N2'))")
            } elseif ($isSharedMailbox -and $mdoCoverageNonBuiltIn -and -not $hasDefenderForO365 -and $mdoSharedDomainOk) {
                # Shared mailbox in scope of MDO policies — cannot just remove license
                $mdoP1Cost = Get-SkuMonthlyPrice "ATP_ENTERPRISE"
                $mdoP1Annual = [math]::Round($mdoP1Cost * 12, 2)
                $recommendations.Add("DISABLED SHARED MAILBOX ($licenseFriendlyStr) — account is disabled and converted to a shared mailbox, but is protected by $($mdoPolicyTypes.Count) Defender for Office 365 $(if ($mdoPolicyTypes.Count -eq 1) { 'policy' } else { 'policies' }). A license is needed to maintain this protection. Consider replacing the current license with a standalone Defender for Office 365 P1 add-on (€$($mdoP1Cost.ToString('N2'))/mo, €$($mdoP1Annual.ToString('N2'))/yr) to reduce costs while maintaining MDO coverage. Current annual cost: €$($userAnnualCost.ToString('N2'))$mdoUpnNote")
                $sharedMbxHandledMdo = $true
            } elseif ($isSharedMailbox -and ($userAnnualCost -gt 0 -or $hasUnknownSku)) {
                # Disabled account converted to shared mailbox — no MDO concern (handled above)
                $recommendations.Add("DISABLED SHARED MAILBOX ($licenseFriendlyStr) — account is disabled and converted to a shared mailbox. Shared mailboxes typically do not require a user license (under 50 GB). If the mailbox is still actively delegated, consider removing the paid license to reduce costs. Annual cost: €$($userAnnualCost.ToString('N2'))")
            } elseif ($userAnnualCost -gt 0 -or $hasUnknownSku) {
                $recommendations.Add("DISABLED ACCOUNT still licensed ($licenseFriendlyStr) — account sign-in is blocked, no litigation hold detected. Review whether the license can be removed to reduce costs. Annual cost: €$($userAnnualCost.ToString('N2'))")
            } else {
                $recommendations.Add("DISABLED ACCOUNT with free SKU ($licenseFriendlyStr) — sign-in is blocked. No financial waste but consider removing it for account hygiene.")
            }
        }

        # Shared mailbox licensing
        # Guard: skip when account is disabled — the disabled block (above) already emits
        # DISABLED SHARED MAILBOX with the correct prefix and handles MDO/storage.
        # Running both blocks produces duplicate/contradictory recommendations.
        $sharedMbxRemoveLicense = $false   # track whether we recommended full removal (gates NON-HUMAN block)
        if ($isSharedMailbox -and $isAccountEnabled) {
            $mbDisplay = if ($null -ne $mbSizeMB) { "${mbSizeMB} MB" } else { "unknown" }
            if ($isLitigationHold) {
                $sharedMbxRemoveLicense = $true
                $recommendations.Add("SHARED MAILBOX ($mbDisplay) on LITIGATION HOLD — mailbox is on hold for eDiscovery. A license is not required to maintain the hold. Consider removing the license — Microsoft will automatically convert this to a free Inactive Mailbox that retains all content and holds indefinitely. Annual savings: €$($userAnnualCost.ToString('N2'))")
            } elseif ($null -eq $mbSizeMB) {
                # Mailbox size unknown — cannot safely recommend removal
                $recommendations.Add("SHARED MAILBOX REVIEW — mailbox size unknown (usage report missing). Review whether the mailbox is under 50 GB and not on hold before removing the license. Annual cost: €$($userAnnualCost.ToString('N2'))")
            } elseif ($mbSizeMB -ge 51200) {
                $sharedMbxRemoveLicense = $true   # prevent contradictory NON-HUMAN ACCOUNT REVIEW — license IS needed
                $recommendations.Add("SHARED MAILBOX over 50 GB ($mbDisplay) — requires Exchange Online Plan 2 (100 GB + auto-expanding archive), or an E3/E5 suite. Exchange Plan 1 and Business Basic/Standard cap at 50 GB and will NOT resolve the issue.")
            } elseif ($mdoCoverageNonBuiltIn -and -not $hasDefenderForO365 -and $mdoSharedDomainOk) {
                # Shared mailbox is under 50 GB but in scope of MDO policies — cannot just remove
                $mdoP1Cost = Get-SkuMonthlyPrice "ATP_ENTERPRISE"
                $mdoP1Annual = [math]::Round($mdoP1Cost * 12, 2)
                $recommendations.Add("SHARED MAILBOX ($mbDisplay) — under 50 GB limit but protected by $($mdoPolicyTypes.Count) Defender for Office 365 $(if ($mdoPolicyTypes.Count -eq 1) { 'policy' } else { 'policies' }). A license is needed to maintain this protection. Consider replacing the current license with a standalone Defender for Office 365 P1 add-on (€$($mdoP1Cost.ToString('N2'))/mo, €$($mdoP1Annual.ToString('N2'))/yr) to reduce costs while maintaining MDO coverage. Current annual cost: €$($userAnnualCost.ToString('N2'))$mdoUpnNote")
                $sharedMbxHandledMdo = $true
            } elseif ($archiveStatus -eq 'Active') {
                # In-place archive is enabled — removing the license disables the archive
                $exo2Price = Get-SkuMonthlyPrice "EXCHANGEENTERPRISE"
                $eoaPrice  = Get-SkuMonthlyPrice "EXCHANGE_ARCHIVE"
                $recommendations.Add("SHARED MAILBOX ($mbDisplay) — under 50 GB but has an active in-place archive. Removing the license will disable the archive. Retain an Exchange Online Plan 2 (€$($exo2Price.ToString('N2'))/mo) or Exchange Online Plan 1 + Archiving add-on (€$($eoaPrice.ToString('N2'))/mo) to maintain the archive. Current annual cost: €$($userAnnualCost.ToString('N2'))")
            } else {
                $mdoWarning = if (-not $mdoCoverageChecked) {
                    " WARNING: MDO policy scope was not evaluated (EXO not connected or no MDO SKU detected) — verify this mailbox is not covered by Defender for Office 365 policies before removing license."
                } else { "" }
                $archiveWarning = if (-not $archiveStatus) {
                    " NOTE: Archive status could not be verified (EXO data unavailable) — confirm no in-place archive exists before making license changes."
                } else { "" }
                $sharedMbxRemoveLicense = $true
                $recommendations.Add("SHARED MAILBOX ($mbDisplay) — typically does not require a user license under 50 GB. A paid license is generally not needed for shared mailboxes within the 50 GB limit. Annual cost: €$($userAnnualCost.ToString('N2'))$mdoWarning$archiveWarning")
            }
        }

        # ── Non-human account premium suite waste ──
        # Shared mailboxes, room/equipment accounts assigned expensive suites (E3/E5/Business Premium)
        # when they only need Exchange Online Plan 2 (if > 50 GB) or nothing at all.
        # Skip if shared mailbox was already recommended for full license removal (avoids contradictory advice).
        if (($isSharedMailbox -or $isRoomOrEquipment) -and $userAnnualCost -gt 0 -and -not $sharedMbxRemoveLicense -and -not $roomLicenseHandled -and $isAccountEnabled) {
            $premiumSuitesNH = @("SPE_E3","SPE_E5","MICROSOFT365_E3","Microsoft_365_E3_Extra_Features",
                "ENTERPRISEPACK","ENTERPRISEPREMIUM","ENTERPRISEPREMIUM_NOPSTNCONF",
                "SPB","O365_BUSINESS_PREMIUM")
            $hasPremiumSuiteNH = @($userSkuList | Where-Object { $_ -in $premiumSuitesNH }).Count -gt 0
            if ($hasPremiumSuiteNH) {
                $premiumNamesNH = ($userSkuList | Where-Object { $_ -in $premiumSuitesNH } | ForEach-Object { Resolve-SkuFriendlyName $_ }) -join "; "
                $nhType = if ($isSharedMailbox) { "Shared Mailbox" } else { $mailboxType }
                $exo2PriceNH = Get-SkuMonthlyPrice "EXCHANGEENTERPRISE"
                $nhSavingsAnnual = [math]::Round(($userMonthlyCost - $exo2PriceNH) * 12, 2)
                if ($nhSavingsAnnual -gt 0) {
                    $recommendations.Add("NON-HUMAN ACCOUNT REVIEW — $nhType is holding a premium user suite ($premiumNamesNH). Non-human accounts typically do not require productivity suites. If a license is needed (>50 GB or archive), use Exchange Online Plan 2 (€$($exo2PriceNH.ToString('N2'))/mo) instead. Potential savings: €$($nhSavingsAnnual.ToString('N2'))/yr")
                    $nonHumanReviewFired = $true
                } else {
                    $recommendations.Add("NON-HUMAN ACCOUNT REVIEW — $nhType is holding a premium user suite ($premiumNamesNH). Non-human accounts typically do not require productivity suites. If a license is needed (>50 GB or archive), use Exchange Online Plan 2 (€$($exo2PriceNH.ToString('N2'))/mo) instead. Current annual cost: €$($userAnnualCost.ToString('N2'))/yr")
                    $nonHumanReviewFired = $true
                }
            }
        }

        # Admin account — admin accounts should have a reduced application footprint
        if ($isAdmin) {
            $adminSafeSkus = @(
                # Identity & security SKUs (expected on admin accounts)
                "AAD_PREMIUM","AAD_PREMIUM_P2","INTUNE_A","ATA",
                "IDENTITY_THREAT_PROTECTION","IDENTITY_THREAT_PROTECTION_FOR_EMS_E3",
                "RIGHTSMANAGEMENT","RIGHTSMANAGEMENT_ADHOC","ATP_ENTERPRISE",
                "THREAT_INTELLIGENCE","INFORMATION_PROTECTION_COMPLIANCE",
                "EMSPREMIUM","EMS","ADALLOM_STANDALONE",
                # Free / viral SKUs (no cost, no waste)
                "FLOW_FREE","POWER_BI_STANDARD","TEAMS_FREE","TEAMS_EXPLORATORY",
                "POWERAPPS_VIRAL","FLOW_P2_VIRAL","WINDOWS_STORE","STREAM",
                "MCOPSTNC","RIGHTSMANAGEMENT_ADHOC","POWERAPPS_DEV",
                "PHONESYSTEM_VIRTUALUSER"
            )
            $appBearingSkus = @($userSkuList | Where-Object { $_ -notin $adminSafeSkus })
            if ($appBearingSkus.Count -gt 0) {
                $appBearingFriendly = ($appBearingSkus | ForEach-Object { Resolve-SkuFriendlyName $_ }) -join "; "
                # PIM-eligible admins MUST retain Entra ID P2 or Entra ID Governance
                $pimNote = if ($pimEligibleRoles -or $pimActiveRoles) { " NOTE: PIM role assignments detected — Entra ID P2 (or Entra ID Governance) must be retained for PIM compliance." } else { "" }
                if ($isLowPrivAdmin) {
                    $recommendations.Add("ADMIN$adminRolesDisplay — this account has low-privilege admin roles (read-only or limited scope) that do not require a full productivity suite. Current application-bearing licenses: $appBearingFriendly. Consider whether the admin role can be combined with the user's primary account to avoid maintaining a separate license.$pimNote")
                } else {
                    $recommendations.Add("ADMIN$adminRolesDisplay — admin accounts typically require only Entra ID P2 and security SKUs, not full productivity suites. Current application-bearing licenses: $appBearingFriendly. Consider using a separate daily-driver account for productivity.$pimNote")
                }
            }
        }

        # ── License exposure signals (PIM / risk-based CA / Defender for Office 365 policy scope) ──
        if (($pimEligibleRoles -or $pimActiveRoles) -and -not $hasEntraP2) {
            $pimEligCountL = if ($pimEligibleRoles) { @($pimEligibleRoles -split ';').Count } else { 0 }
            $pimActCountL  = if ($pimActiveRoles)   { @($pimActiveRoles -split ';').Count }   else { 0 }
            $pimRoleDetail = if ($pimEligCountL -gt 0 -and $pimActCountL -gt 0) { "eligible for $pimEligCountL $(if ($pimEligCountL -eq 1) { 'role' } else { 'roles' }), $pimActCountL active" } elseif ($pimEligCountL -gt 0) { "eligible for $pimEligCountL $(if ($pimEligCountL -eq 1) { 'role' } else { 'roles' })" } else { "$pimActCountL active $(if ($pimActCountL -eq 1) { 'role' } else { 'roles' })" }
            $p2CostPimL = Get-SkuMonthlyPrice "AAD_PREMIUM_P2"
            $p2AnnualPimL = [math]::Round($p2CostPimL * 12, 2)
            $recommendations.Add("LICENSING CHECK — PIM role assignments detected ($pimRoleDetail) but no Entra ID P2 / Entra Governance entitlement found in effective SKUs. Entra ID P2 is a superset of P1 and also covers Conditional Access requirements (standalone: €$($p2CostPimL.ToString('N2'))/mo, €$($p2AnnualPimL.ToString('N2'))/yr, or included in M365 E5/Entra Suite). Review whether licensing for PIM usage is in place.")
        }

        # Only flag explicit/scoped targeting per-user — avoid noise from ALL-user risk policies
        if ($matchedScopedRiskPolicy -and $riskBasedCA -and -not $hasEntraP2) {
            $p2Cost = Get-SkuMonthlyPrice "AAD_PREMIUM_P2"
                $p2Annual = [math]::Round($p2Cost * 12, 2)
                $recommendations.Add("LICENSING CHECK — User is targeted by risk-based Conditional Access policy ($riskBasedCA) but no Entra ID P2 entitlement found. Identity Protection features require P2 (standalone: €$($p2Cost.ToString('N2'))/mo, €$($p2Annual.ToString('N2'))/yr, or included in M365 E5/Entra Suite). Alternatively, exclude this user from the risk-based CA policy to avoid the compliance cost.")
        }

        # ── Compliance gap detection: CA P1 + MDO — consolidated when both apply ──
        # Room/Equipment mailboxes suppressed (resource accounts don't need P1/MDO)
        # B2B guests covered by External ID 1:5 ratio are exempt
        # PIM P2 already recommended = skip P1 (P2 is superset)
        $pimAlreadyNeedsP2Licensed = (($pimEligibleRoles -or $pimActiveRoles) -and -not $hasEntraP2)
        $_hasCaGap  = ($generalCA -and -not $hasEntraP1 -and -not $guestCoveredByRatio -and -not $pimAlreadyNeedsP2Licensed -and -not $isRoomOrEquipment -and -not $isPhoneResource)
        $_hasMdoGap = ($mdoCoverageNonBuiltIn -and $mdoPolicyCoverage -and -not $hasDefenderForO365 -and -not $sharedMbxHandledMdo -and -not $isRoomOrEquipment -and -not $isPhoneResource -and (-not $isSharedMailbox -or $mdoSharedDomainOk))

        if ($_hasCaGap -and $_hasMdoGap) {
            # Combined CA + MDO compliance recommendation
            $p1Cost = Get-SkuMonthlyPrice "AAD_PREMIUM"
            $p1Annual = [math]::Round($p1Cost * 12, 2)
            $mdoP1Cost = Get-SkuMonthlyPrice "ATP_ENTERPRISE"
            $mdoP1Annual = [math]::Round($mdoP1Cost * 12, 2)
            $combinedMonthlyCost = [math]::Round($p1Cost + $mdoP1Cost, 2)
            $combinedAnnualCost  = [math]::Round($p1Annual + $mdoP1Annual, 2)
            $_caDesc = if ($matchedScopedCaPolicy) { "targeted by $($caPolicyNames.Count) Conditional Access $(if ($caPolicyNames.Count -eq 1) { 'policy' } else { 'policies' })" } else { "covered by $($caPolicyNames.Count) tenant-wide Conditional Access policies" }
            $recommendations.Add("LICENSING CHECK — User is $_caDesc and protected by $($mdoPolicyTypes.Count) Defender for Office 365 $(if ($mdoPolicyTypes.Count -eq 1) { 'policy' } else { 'policies' }) but has no Entra ID P1 or MDO entitlement. To ensure compliance, add Entra P1 (€$($p1Cost.ToString('N2'))/mo) + MDO P1 (€$($mdoP1Cost.ToString('N2'))/mo) = €$($combinedMonthlyCost.ToString('N2'))/mo (€$($combinedAnnualCost.ToString('N2'))/yr). Alternatively, exclude this user from the CA and MDO policies to avoid the compliance cost. Note: M365 E3/E5/Business Premium include both Entra P1 and MDO P1.$mdoUpnNote")
        } elseif ($_hasCaGap) {
            if ($matchedScopedCaPolicy) {
                $p1Cost = Get-SkuMonthlyPrice "AAD_PREMIUM"
                $p1Annual = [math]::Round($p1Cost * 12, 2)
                $recommendations.Add("LICENSING CHECK — User is targeted by $($caPolicyNames.Count) Conditional Access $(if ($caPolicyNames.Count -eq 1) { 'policy' } else { 'policies' }) but no Entra ID P1 entitlement found in effective SKUs. Conditional Access requires Entra ID P1 (standalone: €$($p1Cost.ToString('N2'))/mo, €$($p1Annual.ToString('N2'))/yr, or included in M365 E3/E5/Business Premium/F3). Alternatively, exclude this user from the CA $(if ($caPolicyNames.Count -eq 1) { 'policy' } else { 'policies' }) to avoid the compliance cost.")
            } elseif ($caPolicyNames.Count -gt 0) {
                $p1Cost = Get-SkuMonthlyPrice "AAD_PREMIUM"
                $p1Annual = [math]::Round($p1Cost * 12, 2)
                $recommendations.Add("LICENSING CHECK — User is covered by $($caPolicyNames.Count) tenant-wide Conditional Access policies but has no Entra ID P1 entitlement. CA requires P1 for each covered user (standalone: €$($p1Cost.ToString('N2'))/mo, €$($p1Annual.ToString('N2'))/yr, or included in M365 E3/E5/Business Premium/F3). Alternatively, exclude this user from the CA policies to avoid the compliance cost.")
            }
        } elseif ($_hasMdoGap) {
            $mdoP1Cost = Get-SkuMonthlyPrice "ATP_ENTERPRISE"
            $mdoP1Annual = [math]::Round($mdoP1Cost * 12, 2)
            $recommendations.Add("LICENSING CHECK — Mailbox is protected by $($mdoPolicyTypes.Count) Defender for Office 365 $(if ($mdoPolicyTypes.Count -eq 1) { 'policy' } else { 'policies' }) but no MDO license entitlement was found. Consider adding a standalone Defender for Office 365 P1 add-on (€$($mdoP1Cost.ToString('N2'))/mo, €$($mdoP1Annual.ToString('N2'))/yr) to ensure compliance. Alternatively, exclude this mailbox from the MDO policies to avoid the compliance cost. Note: M365 E3/E5/Business Premium already include MDO P1.$mdoUpnNote")
        }

        # ── Standalone Entra ID P2 downgrade to P1 ──
        # P2-only user features: PIM and Risk-based CA. If a non-admin user has standalone
        # AAD_PREMIUM_P2, isn't PIM eligible/active, and isn't in scope for risk-based CA,
        # they only need P1 for standard MFA/CA.
        $hasStandaloneP2 = $userSkuList -contains "AAD_PREMIUM_P2"
        if ($hasStandaloneP2 -and -not $isAdmin -and -not $pimEligibleRoles -and -not $pimActiveRoles -and -not $riskBasedCA) {
            $p2Price = Get-SkuMonthlyPrice "AAD_PREMIUM_P2"
            $p1Price = Get-SkuMonthlyPrice "AAD_PREMIUM"
            $p2Savings = [math]::Round($p2Price - $p1Price, 2)
            if ($p2Savings -gt 0) {
                $p2AnnSavings = [math]::Round($p2Savings * 12, 2)
                $recommendations.Add("ENTRA P2 DOWNGRADE — has standalone Entra ID P2 (€$($p2Price.ToString('N2'))/mo) but is not an admin (no PIM required) and is not in scope for Risk-Based Conditional Access. Consider downgrading to Entra ID P1 (€$($p1Price.ToString('N2'))/mo) for standard MFA/CA. Potential savings: €$($p2Savings.ToString('N2'))/mo (€$($p2AnnSavings.ToString('N2'))/yr).")
            }
        }

        # Overlapping license assignments (Direct + Group for same SKU = redundant assignment)
        # Microsoft deduplicates: same SKU via direct + group consumes only 1 seat.
        # Removing the direct assignment is a hygiene task with no cost savings.
        if ($hasOverlap) {
            $recommendations.Add("OVERLAPPING LICENSE — $overlappingSkus assigned both directly and via group ($licenseGroupsStr). No cost impact (Microsoft deduplicates). Consider removing the direct assignment for cleaner administration. Note: verify the group assignment is permanent before removing.")
        }

        # License assignment errors (insufficient seats, conflicting plans, etc.)
        if ($licenseErrorsStr -ne '') {
            $recommendations.Add("LICENSING ERROR — license assignment failed: $licenseErrorsStr. Review group-based licensing in Entra ID for details.")
        }

        # ── #3 Duplicate suite coverage (suite already includes standalone, with alias + suite-covers-suite) ──
        # IMPORTANT: Only count components that are actually ENABLED in the parent suite.
        # If an admin disabled a service plan in the suite, the standalone covering it is NOT redundant.
        $duplicateHits  = [System.Collections.Generic.List[string]]::new()
        $alreadyFlagged = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

        foreach ($sku in $userSkuList) {
            if (-not $suiteIncludes.ContainsKey($sku)) { continue }
            # Build a HashSet of the parent suite's ENABLED components for O(1) lookups
            # Filter out any service plans that the admin has disabled for this user
            $enabledComponents = @($suiteIncludes[$sku] | Where-Object {
                -not $disabledPlans -or -not $disabledPlans.Contains($_)
            })
            if ($enabledComponents.Count -eq 0) { continue }
            $parentComponents = [System.Collections.Generic.HashSet[string]]::new(
                [string[]]$enabledComponents, [StringComparer]::OrdinalIgnoreCase)

            foreach ($otherSku in $userSkuList) {
                if ($otherSku -eq $sku) { continue }
                if ($alreadyFlagged.Contains($otherSku)) { continue }
                # Skip free/viral SKUs — flagging €0 duplicates is noise, not actionable
                if ((Get-SkuMonthlyPrice $otherSku) -le 0) { continue }

                # ── Pass 1: Direct standalone match (with chained alias resolution) ──
                # Resolve up to 2 hops: e.g. DEFENDER_BUSINESS → MDE_SMB → WIN_DEF_ATP
                $canonical = if ($skuCoverageAliases.ContainsKey($otherSku)) { $skuCoverageAliases[$otherSku] } else { $otherSku }
                $canonical2 = if ($canonical -ne $otherSku -and $skuCoverageAliases.ContainsKey($canonical)) { $skuCoverageAliases[$canonical] } else { $null }
                if ($parentComponents.Contains($otherSku) -or $parentComponents.Contains($canonical) -or ($null -ne $canonical2 -and $parentComponents.Contains($canonical2))) {
                    $duplicateHits.Add("$(Resolve-SkuFriendlyName $otherSku) (included in $(Resolve-SkuFriendlyName $sku))")
                    [void]$alreadyFlagged.Add($otherSku)
                    continue
                }

                # ── Pass 2: Suite-covers-suite (all child components are in parent and enabled) ──
                # Apply alias resolution to child components (e.g. MDE_SMB → WIN_DEF_ATP)
                # to match the canonical form used in $parentComponents.
                if ($suiteIncludes.ContainsKey($otherSku)) {
                    $childComponents = $suiteIncludes[$otherSku]
                    $allCovered = $true
                    foreach ($comp in $childComponents) {
                        $compCanonical = if ($skuCoverageAliases.ContainsKey($comp)) { $skuCoverageAliases[$comp] } else { $comp }
                        $compCanonical2 = if ($compCanonical -ne $comp -and $skuCoverageAliases.ContainsKey($compCanonical)) { $skuCoverageAliases[$compCanonical] } else { $null }
                        if (-not $parentComponents.Contains($comp) -and -not $parentComponents.Contains($compCanonical) -and ($null -eq $compCanonical2 -or -not $parentComponents.Contains($compCanonical2))) {
                            $allCovered = $false
                            break
                        }
                    }
                    if ($allCovered) {
                        $duplicateHits.Add("$(Resolve-SkuFriendlyName $otherSku) (fully covered by $(Resolve-SkuFriendlyName $sku))")
                        [void]$alreadyFlagged.Add($otherSku)
                    }
                }
            }
        }

        if ($duplicateHits.Count -gt 0) {
            [decimal]$dupCost = 0
            foreach ($dupSku in $alreadyFlagged) { $dupCost += Get-SkuMonthlyPrice $dupSku }
            $dupAnnualWaste = [math]::Round($dupCost * 12, 2)
            $duplicateCostAcc += $dupAnnualWaste
            # If user has unknown SKUs, downgrade from REMOVE to REVIEW (LOA v1.0 spec §5.4)
            if ($hasUnknownSku) {
                $recommendations.Add("DUPLICATE REVIEW — likely redundant with suite: $($duplicateHits -join '; '). User also has unmapped SKU(s) — review coverage manually before removing. Annual overlap cost: €$($dupAnnualWaste.ToString('N2'))")
            } else {
                $recommendations.Add("DUPLICATE COVERAGE — redundant with suite: $($duplicateHits -join '; '). Consider removing the redundant SKU(s). Annual overlap cost: €$($dupAnnualWaste.ToString('N2'))")
            }
        }

        # ── Exchange Kiosk candidate (Plan 1 → Kiosk) ──
        # Placed AFTER duplicate detection so $alreadyFlagged is populated.
        # Exchange Kiosk (EXCHANGEDESKLESS, €1/mo, 2 GB cap) is sufficient for users who only
        # access email via OWA and have < 2 GB mailbox. Standalone Exchange Plan 1 costs €4/mo.
        $hasStandaloneExoPlan1 = @($userSkuList | Where-Object { $_ -eq "EXCHANGESTANDARD" }).Count -gt 0
        # Skip Kiosk downgrade if EXCHANGESTANDARD is already flagged for removal as duplicate coverage
        if ($hasStandaloneExoPlan1 -and -not $alreadyFlagged.Contains("EXCHANGESTANDARD") -and $isAccountEnabled -and $mailboxType -ne 'SharedMailbox' -and $mailboxType -ne 'RoomMailbox' -and $mailboxType -ne 'EquipmentMailbox') {
            $usesEmailMobile = ($emailClients -contains "Outlook Mobile") -or ($emailClients -contains "Other Mobile")
            if (-not $usesOutlookDesktop -and -not $usesEmailMobile -and $null -ne $mbSizeMB -and $mbSizeMB -lt 2048) {
                $exoP1Price   = Get-SkuMonthlyPrice "EXCHANGESTANDARD"
                $exoKioskPrice = Get-SkuMonthlyPrice "EXCHANGEDESKLESS"
                $exoKioskSave = [math]::Round($exoP1Price - $exoKioskPrice, 2)
                if ($exoKioskSave -gt 0) {
                    $exoKioskAnnSave = [math]::Round($exoKioskSave * 12, 2)
                    $exoKioskSavingsAcc += $exoKioskAnnSave
                    $kioskStorageDisplay = if ($mbSizeMB -ge 1024) { "$([math]::Round($mbSizeMB / 1024, 1)) GB" } elseif ($mbSizeMB -gt 0) { "$([math]::Round($mbSizeMB, 1)) MB" } else { "0 MB" }
                    # Compliance note: warn when user is covered by CA or MDO policies (Exchange standalone SKUs never include P1/MDO)
                    $kioskCompNote = ""
                    $kioskNeedsP1  = ($generalCA -ne '')
                    $kioskNeedsMdo = $mdoCoverageNonBuiltIn
                    if ($kioskNeedsP1 -or $kioskNeedsMdo) {
                        $kioskCaCount  = if ($kioskNeedsP1)  { @($generalCA -split ';').Count } else { 0 }
                        $kioskMdoCount = if ($kioskNeedsMdo -and $mdoPolicyCoverage) { @($mdoPolicyCoverage -split ';').Count } else { 0 }
                        $kioskCompNote = " Note: user is also covered by$(if ($kioskNeedsP1) { " $kioskCaCount Conditional Access" })$(if ($kioskNeedsP1 -and $kioskNeedsMdo) { ' and' })$(if ($kioskNeedsMdo) { " $kioskMdoCount MDO" }) $(if (($kioskCaCount + $kioskMdoCount) -eq 1) { 'policy' } else { 'policies' }) requiring$(if ($kioskNeedsP1) { " Entra ID P1 (€$((Get-SkuMonthlyPrice 'AAD_PREMIUM').ToString('N2'))/mo)" })$(if ($kioskNeedsP1 -and $kioskNeedsMdo) { ' and' })$(if ($kioskNeedsMdo) { " MDO P1 (€$((Get-SkuMonthlyPrice 'ATP_ENTERPRISE').ToString('N2'))/mo)" }) for compliance. Factor in total cost before downgrading."
                    }
                    $recommendations.Add("EXCHANGE KIOSK CANDIDATE — has Exchange Plan 1 (€$($exoP1Price.ToString('N2'))/mo) but only accesses email via OWA and uses $kioskStorageDisplay of storage (< 2 GB). Consider downgrading to Exchange Kiosk (€$($exoKioskPrice.ToString('N2'))/mo).$kioskCompNote Potential savings: €$($exoKioskSave.ToString('N2'))/mo (€$($exoKioskAnnSave.ToString('N2'))/yr).")
                }
            }
        }

        # ── #4 E3 + add-ons → E5 upgrade opportunity ──
        $hasE3 = ($userSkuList | Where-Object { $_ -in $e3Suites }) | Select-Object -First 1
        $suiteInversionFired = $false   # track whether E5 upgrade already recommended (suppresses conflicting bundle consolidation)
        if ($hasE3) {
            $matchedAddons = @($userSkuList | Where-Object { $_ -in $e5AddOns })
            # Exclude $0-priced add-ons (e.g. free Audio Conferencing) from the count trigger —
            # they inflate the add-on count without adding real cost to justify E5 consolidation.
            $paidAddons = @($matchedAddons | Where-Object { (Get-SkuMonthlyPrice $_) -gt 0 })
            if ($paidAddons.Count -ge 1) {
                $addonFriendly = ($matchedAddons | ForEach-Object { Resolve-SkuFriendlyName $_ }) -join "; "
                # Calculate actual cost comparison
                [decimal]$addonCostSum = 0
                foreach ($a in $matchedAddons) { $addonCostSum += Get-SkuMonthlyPrice $a }
                $currentCombined = (Get-SkuMonthlyPrice $hasE3) + $addonCostSum
                $e5Cost  = Get-SkuMonthlyPrice "SPE_E5"
                $delta   = [math]::Round($currentCombined - $e5Cost, 2)
                if ($delta -gt 0) {
                    $annualSave = [math]::Round($delta * 12, 2)
                    $e5UpgradeSavingsAcc += $annualSave
                    $suiteInversionFired = $true
                    $recommendations.Add("SUITE INVERSION — has $(Resolve-SkuFriendlyName $hasE3) + $($matchedAddons.Count) E5-included add-on(s) ($addonFriendly) totalling €$($currentCombined.ToString('N2'))/mo. Full M365 E5 costs €$($e5Cost.ToString('N2'))/mo — upgrade saves €$($delta.ToString('N2'))/mo (€$($annualSave.ToString('N2'))/yr) AND unlocks remaining E5 capabilities (Power BI Pro, Defender for Cloud Apps, risk-based CA, etc.).")
                } else {
                    $recommendations.Add("E5 CONSOLIDATION — has $(Resolve-SkuFriendlyName $hasE3) + $($matchedAddons.Count) add-on(s) ($addonFriendly) at €$($currentCombined.ToString('N2'))/mo. E5 is €$($e5Cost.ToString('N2'))/mo (Δ €$($delta.ToString('N2'))). Not cheaper, but simplifies to 1 SKU with full feature alignment.")
                }
            }
        }

        # ── #4b À la carte bundle consolidation (O365 + EMS + Windows → M365) ──
        # Users paying "à la carte" for the three pillars separately when a single M365 E3/E5 bundle is cheaper.
        # Only fire if user does NOT already have a unified M365 suite (SPE_E3/SPE_E5).
        # Skip if SUITE INVERSION already fired — that recommendation supersedes bundle consolidation
        # (E5 upgrade subsumes pillar consolidation, and the two recommendations would conflict).
        $hasUnifiedM365 = @($userSkuList | Where-Object { $_ -in @("SPE_E3","SPE_E5","MICROSOFT365_E3","Microsoft_365_E3_Extra_Features") }).Count -gt 0
        if (-not $hasUnifiedM365 -and -not $suiteInversionFired) {
            $hasO365E3   = @($userSkuList | Where-Object { $_ -eq "ENTERPRISEPACK" }).Count -gt 0
            $hasEmsE3    = @($userSkuList | Where-Object { $_ -eq "EMS" }).Count -gt 0
            $hasWinE3    = @($userSkuList | Where-Object { $_ -eq "WIN10_PRO_ENT_SUB" }).Count -gt 0
            $hasO365E5   = @($userSkuList | Where-Object { $_ -in @("ENTERPRISEPREMIUM","ENTERPRISEPREMIUM_NOPSTNCONF") }).Count -gt 0
            $hasEmsE5    = @($userSkuList | Where-Object { $_ -eq "EMSPREMIUM" }).Count -gt 0
            $hasWinE5    = @($userSkuList | Where-Object { $_ -eq "WIN10_VDA_E5" }).Count -gt 0

            if ($hasO365E3 -and $hasEmsE3 -and $hasWinE3) {
                # E3 tier: O365 E3 + EMS E3 + Windows E3 → M365 E3
                $alaCarteCost  = (Get-SkuMonthlyPrice "ENTERPRISEPACK") + (Get-SkuMonthlyPrice "EMS") + (Get-SkuMonthlyPrice "WIN10_PRO_ENT_SUB")
                $bundleCost    = Get-SkuMonthlyPrice "SPE_E3"
                $bundleDelta   = [math]::Round($alaCarteCost - $bundleCost, 2)
                if ($bundleDelta -gt 0) {
                    $bundleAnnual = [math]::Round($bundleDelta * 12, 2)
                    $bundleConsolidationSavingsAcc += $bundleAnnual
                    $recommendations.Add("BUNDLE CONSOLIDATION — has Office 365 E3 (€$((Get-SkuMonthlyPrice 'ENTERPRISEPACK').ToString('N2'))/mo) + EMS E3 (€$((Get-SkuMonthlyPrice 'EMS').ToString('N2'))/mo) + Windows E3 (€$((Get-SkuMonthlyPrice 'WIN10_PRO_ENT_SUB').ToString('N2'))/mo) = €$($alaCarteCost.ToString('N2'))/mo. Consider consolidating to Microsoft 365 E3 (€$($bundleCost.ToString('N2'))/mo) and save €$($bundleDelta.ToString('N2'))/mo (€$($bundleAnnual.ToString('N2'))/yr).")
                }
            } elseif ($hasO365E5 -and $hasEmsE5 -and $hasWinE5) {
                # E5 tier: O365 E5 + EMS E5 + Windows E5 → M365 E5
                $o365E5Sku     = if ($hasO365E5) { @($userSkuList | Where-Object { $_ -in @("ENTERPRISEPREMIUM","ENTERPRISEPREMIUM_NOPSTNCONF") }) | Select-Object -First 1 } else { "ENTERPRISEPREMIUM" }
                $alaCarteCost  = (Get-SkuMonthlyPrice $o365E5Sku) + (Get-SkuMonthlyPrice "EMSPREMIUM") + (Get-SkuMonthlyPrice "WIN10_VDA_E5")
                $bundleCost    = Get-SkuMonthlyPrice "SPE_E5"
                $bundleDelta   = [math]::Round($alaCarteCost - $bundleCost, 2)
                if ($bundleDelta -gt 0) {
                    $bundleAnnual = [math]::Round($bundleDelta * 12, 2)
                    $bundleConsolidationSavingsAcc += $bundleAnnual
                    $recommendations.Add("BUNDLE CONSOLIDATION — has Office 365 E5 (€$((Get-SkuMonthlyPrice $o365E5Sku).ToString('N2'))/mo) + EMS E5 (€$((Get-SkuMonthlyPrice 'EMSPREMIUM').ToString('N2'))/mo) + Windows E5 (€$((Get-SkuMonthlyPrice 'WIN10_VDA_E5').ToString('N2'))/mo) = €$($alaCarteCost.ToString('N2'))/mo. Consider consolidating to Microsoft 365 E5 (€$($bundleCost.ToString('N2'))/mo) and save €$($bundleDelta.ToString('N2'))/mo (€$($bundleAnnual.ToString('N2'))/yr).")
                }
            }
        }

        # ── #4c Teams Unbundling (zero Teams activity on bundled suite) ──
        # Microsoft offers "Without Teams" variants of major suites at €2–3/mo less.
        # Flag users on bundled suites with zero Teams activity as candidates for the cheaper SKU.
        $teamsBundledSkus = @("SPE_E3","SPE_E5","MICROSOFT365_E3","Microsoft_365_E3_Extra_Features",
            "ENTERPRISEPACK","ENTERPRISEPREMIUM","ENTERPRISEPREMIUM_NOPSTNCONF",
            "SPB","O365_BUSINESS_PREMIUM","O365_BUSINESS_ESSENTIALS","DESKLESSPACK","M365_F1_COMM")
        $hasBundledTeamsSku = @($userSkuList | Where-Object { $_ -in $teamsBundledSkus }).Count -gt 0
        if ($hasBundledTeamsSku -and $hasTeamsEntitlement -and $teamsTotal -eq 0 -and $au -and -not $nonHumanReviewFired) {
            $recommendations.Add("TEAMS UNBUNDLING — assigned a suite that bundles Teams but shows 0 Teams activity in $ReportPeriod. Consider switching to the equivalent 'Without Teams' SKU to save ~€2–3/user/mo on the bundled Teams component.")
        }

        # ── #4d Windows standalone license waste (Mac/Mobile-only users) ──
        # ideas.md #1: standalone Windows Enterprise E3/E5 on users with no Windows platform usage
        $windowsStandaloneSkus = @("WIN10_PRO_ENT_SUB","WIN10_VDA_E5")
        $hasWinStandalone = @($userSkuList | Where-Object { $_ -in $windowsStandaloneSkus }).Count -gt 0
        if ($hasWinStandalone) {
            $usesWindowsPlatform = $false
            $usesMacOrMobilePlatform = $false
            $hasActivationData = $false
            if ($lkpActivations.ContainsKey($upn)) {
                $hasActivationData = $true
                foreach ($ar in $lkpActivations[$upn]) {
                    if ([int]($ar.'Windows' -as [int]) -gt 0) { $usesWindowsPlatform = $true }
                    if ([int]($ar.'Mac' -as [int]) -gt 0 -or [int]($ar.'iOS' -as [int]) -gt 0 -or [int]($ar.'Android' -as [int]) -gt 0) { $usesMacOrMobilePlatform = $true }
                }
            }
            # Also check Teams Device Usage report for Windows platform
            if (-not $usesWindowsPlatform -and $td) {
                $tdProps2 = $td.PSObject.Properties.Name
                if ('Used Windows' -in $tdProps2 -and $td.'Used Windows' -in @('True','Yes')) { $usesWindowsPlatform = $true }
                if ('Used Mac'     -in $tdProps2 -and $td.'Used Mac'     -in @('True','Yes')) { $usesMacOrMobilePlatform = $true }
                if ('Used iOS'     -in $tdProps2 -and $td.'Used iOS'     -in @('True','Yes')) { $usesMacOrMobilePlatform = $true }
                if ('Used Android Phone' -in $tdProps2 -and $td.'Used Android Phone' -in @('True','Yes')) { $usesMacOrMobilePlatform = $true }
                if (-not $hasActivationData) { $hasActivationData = ($teamsPlats.Count -gt 0) }
            }
            if ($hasActivationData -and -not $usesWindowsPlatform -and $usesMacOrMobilePlatform) {
                $winSku  = ($userSkuList | Where-Object { $_ -in $windowsStandaloneSkus } | Select-Object -First 1)
                $winCost = [math]::Round((Get-SkuMonthlyPrice $winSku) * 12, 2)
                $winName = Resolve-SkuFriendlyName $winSku
                $recommendations.Add("WINDOWS LICENSE REVIEW — standalone $winName (€$((Get-SkuMonthlyPrice $winSku).ToString('N2'))/mo) assigned but no Windows platform activations detected. User only activates on: $activatedPlatforms. Consider removing the Windows subscription. Annual cost: €$($winCost.ToString('N2'))")
            }
        }

        # ── #4e Viral/Exploratory license cleanup ──
        # ideas.md #4: viral SKUs alongside paid suites cause provisioning conflicts in Entra ID
        $viralSkus = @("TEAMS_EXPLORATORY","POWERAPPS_VIRAL","FLOW_P2_VIRAL","TEAMS_FREE")
        $userViralSkus = @($userSkuList | Where-Object { $_ -in $viralSkus })
        if ($userViralSkus.Count -gt 0) {
            $hasPaidProductivitySuite = @($userSkuList | Where-Object {
                $suiteIncludes.ContainsKey($_) -and -not $addOnBundles.Contains($_) -and $_ -notin $viralSkus
            }).Count -gt 0
            if ($hasPaidProductivitySuite) {
                $viralNames = ($userViralSkus | ForEach-Object { Resolve-SkuFriendlyName $_ }) -join "; "
                $recommendations.Add("FREE LICENSE OVERLAP — user holds viral/exploratory license(s) ($viralNames) alongside a paid productivity suite. These free SKUs can cause provisioning conflicts in Entra ID. Consider removing the viral assignment(s) to prevent compliance issues.")
            }
        }

        # ── #4f Visio Plan 1 seeded redundancy ──
        # Microsoft now includes lightweight "Visio in Microsoft 365" web app natively in E1/E3/E5.
        # Users with standalone Visio Plan 1 + qualifying suite likely don't need Plan 1 for
        # viewing/light editing.  Use SharePoint file activity as a proxy for heavy Visio usage.
        $hasVisioPlan1 = @($userSkuList | Where-Object { $_ -eq "VISIOONLINE_PLAN1" }).Count -gt 0
        $visioSeededSuites = @("SPE_E3","SPE_E5","ENTERPRISEPACK","ENTERPRISEPREMIUM","STANDARDPACK")
        $hasVisioSeededSuite = @($userSkuList | Where-Object { $_ -in $visioSeededSuites }).Count -gt 0
        if ($hasVisioPlan1 -and $hasVisioSeededSuite) {
            if ($spViewed -lt 5) {
                $visioPlan1Cost = Get-SkuMonthlyPrice "VISIOONLINE_PLAN1"
                $visioPlan1Annual = [math]::Round($visioPlan1Cost * 12, 2)
                $recommendations.Add("SEEDED VISIO OVERLAP — holds Visio Plan 1 (€$($visioPlan1Cost.ToString('N2'))/mo) alongside $(Resolve-SkuFriendlyName ($userSkuList | Where-Object { $_ -in $visioSeededSuites } | Select-Object -First 1)) which natively includes the 'Visio in Microsoft 365' web app. Based on low SharePoint file activity ($spViewed files viewed/edited in $ReportPeriod), the native app is likely sufficient. Consider removing Visio Plan 1. Annual savings: €$($visioPlan1Annual.ToString('N2'))")
            }
        }

        # ── #5 Visio/Project/Power BI/Teams Premium shelfware ──
        # Product-specific activation map: Visio/Project use activation data, NOT generic app usage.
        # The M365 App usage report only tracks Word/Excel/PPT/OneNote/Outlook/Teams — a user with
        # heavy Word usage would never be flagged even if they haven't opened Visio in years.
        $shelfwareProductMap = @{
            "VISIOCLIENT"         = "Visio"
            # (Flaw #5 fix: VISIOONLINE_PLAN1 removed — web-only product, no desktop activation
            #  record will ever exist. Falls through to generic activity check below.)
            "PROJECTPROFESSIONAL" = "Project"
            "PROJECTPREMIUM"      = "Project"
            # NOTE: PROJECTESSENTIALS (Project Online Essentials) is web-only — no desktop activation
            # record will ever exist, so it uses the generic activity fallback, not this map.
        }
        # Web-only expensive SKUs get REVIEW instead of hard SHELFWARE (no activation telemetry)
        $webOnlyExpensiveSkus = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        [void]$webOnlyExpensiveSkus.Add("VISIOONLINE_PLAN1")
        foreach ($sku in $userSkuList) {
            if ($expensiveStandalone.ContainsKey($sku)) {
                # Teams Premium: features are organizer-driven (webinars, branding, watermarks).
                # Attendees do NOT need a Premium license — only organizers do.
                # Use Meetings Organized Count (not total Meeting Count which includes attended).
                if ($sku -eq "Microsoft_Teams_Premium") {
                    if ($teamsMeetingsOrganized -lt 3) {
                        $shelfCost = [math]::Round((Get-SkuMonthlyPrice $sku) * 12, 2)
                        $mtgNote = if ($teamsMeetingsOrganized -eq 0) { "organized 0 meetings" } else { "organized only $teamsMeetingsOrganized meeting(s)" }
                        $userShelfwareCost += $shelfCost
                        $recommendations.Add("INACTIVE ADD-ON — $($expensiveStandalone[$sku]) license assigned but $mtgNote in $ReportPeriod. Premium features (webinars, branding, watermarks) are organizer-driven; attendees typically do not need this license. Review whether the license is still needed. Annual cost: €$($shelfCost.ToString('N2'))")
                    }
                } elseif ($shelfwareProductMap.ContainsKey($sku)) {
                    # Visio/Project: cross-reference activation report for product-specific usage
                    $targetProduct = $shelfwareProductMap[$sku]
                    $hasProductActivation = $false
                    $hasProductDesktopActivation = $false
                    if ($lkpActivations.ContainsKey($upn)) {
                        foreach ($ar in $lkpActivations[$upn]) {
                            if ($ar.'Product Type' -match $targetProduct) {
                                $hasProductActivation = $true
                                # Check product-specific desktop activation (Windows/Mac)
                                $pWin = [int]($ar.'Windows' -as [int])
                                $pMac = [int]($ar.'Mac' -as [int])
                                if ($pWin -gt 0 -or $pMac -gt 0) { $hasProductDesktopActivation = $true }
                                break
                            }
                        }
                    }
                    if (-not $hasProductActivation) {
                        $shelfCost = [math]::Round((Get-SkuMonthlyPrice $sku) * 12, 2)
                        $userShelfwareCost += $shelfCost
                        $recommendations.Add("INACTIVE ADD-ON — $($expensiveStandalone[$sku]) license assigned but no $targetProduct activation detected. Review whether the license is still needed. Annual cost: €$($shelfCost.ToString('N2'))")
                    } elseif (-not $hasProductDesktopActivation) {
                        # Has activation but NO desktop (Windows/Mac) — user accesses via mobile/web only.
                        # Desktop-tier SKU is overkill; downgrade to web-only Plan 1 equivalent.
                        $desktopToWebMap = @{
                            "VISIOCLIENT"         = "VISIOONLINE_PLAN1"
                            "PROJECTPROFESSIONAL" = "PROJECTESSENTIALS"
                            "PROJECTPREMIUM"      = "PROJECTESSENTIALS"
                        }
                        $webSku = if ($desktopToWebMap.ContainsKey($sku)) { $desktopToWebMap[$sku] } else { $null }
                        if ($webSku) {
                            $desktopPrice = Get-SkuMonthlyPrice $sku
                            $webPrice     = Get-SkuMonthlyPrice $webSku
                            $pawSavings   = [math]::Round($desktopPrice - $webPrice, 2)
                            $pawAnnual    = [math]::Round($pawSavings * 12, 2)
                            $cpcNote = if (@($userSkuList | Where-Object { $_ -match '^(CPC_|Windows_365_)' }).Count -gt 0) {
                                " NOTE: This user has a Cloud PC license assigned. If $targetProduct desktop is used on the Cloud PC, it may appear as a web activation. Verify actual usage before downgrading."
                            } else { "" }
                            $recommendations.Add("PREMIUM ADD-ON REVIEW — $($expensiveStandalone[$sku]) (€$($desktopPrice.ToString('N2'))/mo) assigned but $targetProduct is activated on mobile/web only, no Windows or Mac desktop activation detected. Consider downgrading to $(Resolve-SkuFriendlyName $webSku) (€$($webPrice.ToString('N2'))/mo). Potential savings: €$($pawSavings.ToString('N2'))/mo (€$($pawAnnual.ToString('N2'))/yr).$cpcNote")
                        }
                    }
                } else {
                    # Power BI and other SKUs: use generic app/SharePoint activity
                    $hasAnyAppActivity = $usesDesktop -or $usesWeb -or $usesMobile -or ($spTotal -gt 0)
                    if (-not $hasAnyAppActivity) {
                        $shelfCost = [math]::Round((Get-SkuMonthlyPrice $sku) * 12, 2)
                        if ($webOnlyExpensiveSkus.Contains($sku)) {
                            # (Flaw #5 fix: web-only products lack activation telemetry — downgrade to REVIEW)
                            $recommendations.Add("INACTIVE ADD-ON REVIEW — $($expensiveStandalone[$sku]) license (web-only product) assigned but no app/SharePoint activity detected. Web-only products lack activation telemetry — review actual browser usage before removing. Annual cost: €$($shelfCost.ToString('N2'))")
                        } else {
                            $userShelfwareCost += $shelfCost
                            $recommendations.Add("INACTIVE ADD-ON — $($expensiveStandalone[$sku]) license assigned but no app/SharePoint activity detected. Review whether the license is still needed. Annual cost: €$($shelfCost.ToString('N2'))")
                        }
                    }
                }
            }
        }

        # ── #6 Teams Phone PSTN review (gated on Teams client) ──
        # NOTE: Direct Routing (SBC) and Operator Connect are valid PSTN routes
        # that do NOT produce a Microsoft Calling Plan SKU.  We can only detect
        # Microsoft first-party Calling Plans via licensing; absence of a Calling
        # Plan SKU does NOT mean "no PSTN route."  Emit REVIEW, not REMOVE.
        $hasTeamsPhone  = $false
        $hasCallingPlan = $false
        foreach ($sku in $userSkuList) {
            if ($teamsPhoneSkus.Contains($sku))  { $hasTeamsPhone  = $true }
            if ($callingPlanSkus.Contains($sku)) { $hasCallingPlan = $true }
            if ($hasTeamsPhone -and $hasCallingPlan) { break }
        }
        $hasTeamsClient = $userHasTeamsClient.Contains($upn)
        # Also detect Teams Phone / Audio Conferencing from suite expansion (e.g. EEA no-Teams bundles)
        $hasPhoneEntitlement = ($hasTeamsPhone -or $effectiveSkuSet.Contains("MCOEV"))
        $hasAudioConf        = $effectiveSkuSet.Contains("MCOMEETADV")
        if ($hasPhoneEntitlement -and -not $hasTeamsClient) {
            # User has Teams Phone entitlement but no Teams client (e.g. EEA no-Teams bundle)
            # Determine correct Teams add-on: EEA for EEA/w/o bundles, Enterprise for non-EEA
            $needsEEA = @($userSkuList | Where-Object { $_ -match 'EEA' -or $_ -match 'w/o' -or $_ -match '\(no.?Teams\)' }).Count -gt 0
            $teamsAddonName = if ($needsEEA) { "Microsoft Teams EEA" } else { "Microsoft Teams Enterprise" }
            $recommendations.Add("LICENSING CHECK — Teams Phone$(if ($hasAudioConf) {' and Audio Conferencing'}) entitlement detected but no Teams client (TEAMS1) service plan enabled. Consider adding the $teamsAddonName add-on to activate telephony features, or removing the Teams Phone add-on if not needed.")
        } elseif ($hasTeamsPhone -and $hasTeamsClient -and -not $hasCallingPlan) {
            $recommendations.Add("TEAMS PHONE REVIEW — Teams Phone SKU assigned but no Microsoft Calling Plan detected. If this tenant uses Direct Routing (SBC) or Operator Connect for PSTN, this license is required and valid. If no PSTN route is configured, the Teams Phone license may have no value — review with the Teams administrator before removing.")
        }

        # ── #6b Calling Plan Shelfware (paid PSTN plan with 0 calls) ──
        # Unlike Teams Phone (where Direct Routing is invisible), Microsoft Calling Plans
        # provide Microsoft-managed PSTN. Zero Teams calls over the report period = unused.
        # Exclude MCOPSTNC (Communications Credits = shared pool, not per-user waste).
        $userCallingPlanSkus = @($userSkuList | Where-Object { $_ -in $callingPlanSkus -and $_ -ne "MCOPSTNC" })
        if ($userCallingPlanSkus.Count -gt 0 -and $teamsCalls -eq 0 -and $tm) {
            foreach ($cpSku in $userCallingPlanSkus) {
                $cpPrice  = Get-SkuMonthlyPrice $cpSku
                if ($cpPrice -gt 0) {
                    $cpAnnual = [math]::Round($cpPrice * 12, 2)
                    $cpName   = Resolve-SkuFriendlyName $cpSku
                    $recommendations.Add("CALLING PLAN REVIEW — $cpName (€$($cpPrice.ToString('N2'))/mo) assigned but 0 Teams calls recorded in the $ReportPeriod report period. Consider removing the calling plan and reallocating or cancelling the subscription. Annual cost: €$($cpAnnual.ToString('N2'))")
                }
            }
        }

        # ── #6a Teams Phone Right-Sizing (non-human accounts with full phone license) ──
        # Shared/Room/Equipment mailboxes assigned Teams Phone (€8/mo) instead of
        # the cheaper Teams Shared Devices license (€2.50/mo) designed for common area phones.
        $hasStandaloneMcoev = @($userSkuList | Where-Object { $_ -eq "MCOEV" }).Count -gt 0
        if ($hasStandaloneMcoev -and ($isSharedMailbox -or $isRoomOrEquipment)) {
            $mcoevPrice  = Get-SkuMonthlyPrice "MCOEV"
            $sharedPrice = Get-SkuMonthlyPrice "MCOCAP"
            $phoneSaving = [math]::Round($mcoevPrice - $sharedPrice, 2)
            $phoneAnnual = [math]::Round($phoneSaving * 12, 2)
            $nhTypePhone = if ($isSharedMailbox) { "Shared Mailbox" } else { $mailboxType }
            $recommendations.Add("TEAMS PHONE RIGHT-SIZING — $nhTypePhone has Teams Phone (€$($mcoevPrice.ToString('N2'))/mo) but non-human accounts only need the Teams Shared Devices license (€$($sharedPrice.ToString('N2'))/mo) for common area phones, lobby devices, or conference rooms. Potential savings: €$($phoneSaving.ToString('N2'))/mo (€$($phoneAnnual.ToString('N2'))/yr).")
        }

        # ── #6b Legacy Auth / Service Account Waste ──
        # Users whose ONLY email activity is via POP3/IMAP4/SMTP (no Outlook Desktop, OWA, or Mobile)
        # are almost certainly scan-to-email or script accounts on expensive suites. Legacy protocols
        # also bypass most Conditional Access policies — a dual waste + security gap.
        $legacyOnlyProtocols = @("POP3","IMAP4","SMTP")
        $modernClients = @("Outlook Windows","Outlook Mac","OWA","Outlook Mobile","Other Mobile")
        if ($emailTotal -gt 0 -and $ea -and $emailClients.Count -gt 0) {
            $hasModernClient = $false
            $hasLegacyClient = $false
            foreach ($cl in $emailClients) {
                if ($cl -in $modernClients) { $hasModernClient = $true }
                if ($cl -in $legacyOnlyProtocols) { $hasLegacyClient = $true }
            }
            if ($hasLegacyClient -and -not $hasModernClient -and $userAnnualCost -gt 50) {
                $legacyProtos = ($emailClients | Where-Object { $_ -in $legacyOnlyProtocols }) -join "/"
                $exo1PriceLSA = Get-SkuMonthlyPrice 'EXCHANGESTANDARD'
                $lsaSavingsAnnual = [math]::Round(($userMonthlyCost - $exo1PriceLSA) * 12, 2)
                $lsaSavingsNote = if ($lsaSavingsAnnual -gt 0) { " Potential savings: €$($lsaSavingsAnnual.ToString('N2'))/yr" } else { "" }
                $recommendations.Add("LEGACY SERVICE ACCOUNT — email activity detected but ONLY via legacy protocols ($legacyProtos). No Outlook Desktop, OWA, or Mobile client used. This is likely a scan-to-email or script account on a €$($userMonthlyCost.ToString('N2'))/mo suite. Consider downgrading to Exchange Plan 1 (€$($exo1PriceLSA.ToString('N2'))/mo) or Kiosk.$lsaSavingsNote SECURITY: Legacy protocols bypass most Conditional Access policies — consider migrating to Graph API/SMTP AUTH with Modern Auth.")
            }
        }

        # ── #7 Copilot license flag ──
        # ── Copilot productivity (M365 Copilot — requires base productivity suite) ──
        $hasCopilotProd     = @($userSkuList | Where-Object { $_ -in $copilotProductivitySkus }).Count -gt 0
        $hasCopilotBusiness = @($userSkuList | Where-Object { $_ -in $copilotBusinessSkus }).Count -gt 0
        $hasCopilotSecurity = @($userSkuList | Where-Object { $_ -in $copilotSecuritySkus }).Count -gt 0
        $hasCopilotStudio   = @($userSkuList | Where-Object { $_ -in $copilotStudioSkus }).Count -gt 0
        # Also catch unknown future Copilot SKU variants via name match (but exclude known sub-types)
        $hasCopilotAny      = $hasCopilotProd -or $hasCopilotBusiness -or
                              @($userSkuList | Where-Object { $_ -match 'COPILOT' -and $_ -notin $copilotStudioSkus -and $_ -notin $copilotSecuritySkus -and $_ -notin $copilotBusinessSkus -and $_ -notin $copilotProductivitySkus }).Count -gt 0

        [decimal]$userCopilotAnnualCost = 0   # tracks Copilot-specific cost for deduplication in $noActivityCostAcc
        if ($hasCopilotProd -or $hasCopilotBusiness -or $hasCopilotAny) {
            # Determine Copilot variant for messaging
            $copilotVariant = if ($hasCopilotBusiness) { "M365 Copilot (Business)" }
                              elseif ($hasCopilotProd) { "M365 Copilot" }
                              else { "M365 Copilot" }
            # Check prerequisite: Copilot requires E3/E5, Business Standard/Premium, or Education A3/A5
            $hasValidBase = @($userSkuList | Where-Object { $copilotBaseSkus.Contains($_) }).Count -gt 0
            if (-not $hasValidBase) {
                $recommendations.Add("COPILOT PREREQUISITE MISSING — $copilotVariant assigned but no valid base license (E3/E5, Business Standard/Premium, or Education A3/A5). Copilot will not function. Consider assigning a qualifying base license or reallocating the Copilot license.")
            } else {
                # ── Copilot 3-tier adoption pipeline — prefer real Copilot usage report (beta), fall back to proxy ──
                # Tiers: RECLAIM (no Copilot + no workload readiness), WATCHLIST (no Copilot + active workloads), KEEP (active Copilot)
                $cu = $lkpCopilotUsage[$upn]
                # Workload readiness: is this user actively using M365 base apps? (independent of Copilot)
                $workloadReady = ($teamsTotal -gt 0 -or $emailTotal -gt 0 -or $odTotal -gt 0 -or $spTotal -gt 0 -or $usesDesktop -or $usesWeb -or $usesMobile -or $teamsUsesMobile -or $teamsUsesWeb)
                $copilotSku    = ($userSkuList | Where-Object { $_ -in ($copilotProductivitySkus + $copilotBusinessSkus) } | Select-Object -First 1)
                if (-not $copilotSku) {
                    # Unknown Copilot SKU matched via regex — use the first COPILOT-matching SKU for cost
                    $copilotSku = ($userSkuList | Where-Object { $_ -match 'COPILOT' -and $_ -notin $copilotStudioSkus -and $_ -ne 'Microsoft_Security_Copilot' } | Select-Object -First 1)
                }
                $copilotPrice  = Get-SkuMonthlyPrice $copilotSku
                $copilotAnnual = [math]::Round($copilotPrice * 12, 2)
                if ($copilotUsageLoaded -and $cu) {
                    # Real Copilot activity data available — check all product-specific last-activity columns
                    $copilotActiveApps = @()
                    if ($cu.'Microsoft Teams Copilot Last Activity Date') { $copilotActiveApps += "Teams" }
                    if ($cu.'Word Copilot Last Activity Date')            { $copilotActiveApps += "Word" }
                    if ($cu.'Excel Copilot Last Activity Date')           { $copilotActiveApps += "Excel" }
                    if ($cu.'PowerPoint Copilot Last Activity Date')      { $copilotActiveApps += "PowerPoint" }
                    if ($cu.'Outlook Copilot Last Activity Date')         { $copilotActiveApps += "Outlook" }
                    if ($cu.'OneNote Copilot Last Activity Date')         { $copilotActiveApps += "OneNote" }
                    if ($cu.'Loop Copilot Last Activity Date')            { $copilotActiveApps += "Loop" }
                    if ($cu.'Copilot Chat Last Activity Date')            { $copilotActiveApps += "Copilot Chat" }
                    if ($copilotActiveApps.Count -eq 0) {
                        # Zero Copilot activity — tier by workload readiness
                        $userCopilotAnnualCost = $copilotAnnual
                        if (-not $workloadReady) {
                            $copilotNonAdopterCostAcc += $copilotAnnual
                            $copilotReclaimCostAcc += $copilotAnnual
                            $recommendations.Add("COPILOT RECLAIM — $copilotVariant (€$($copilotPrice.ToString('N2'))/mo) assigned but zero Copilot activity AND zero M365 workload activity in $ReportPeriod. User shows no readiness for AI-assisted productivity. Consider reallocating to an active user. Savings: €$($copilotPrice.ToString('N2'))/mo (€$($copilotAnnual.ToString('N2'))/yr).")
                        } else {
                            $copilotWatchlistCostAcc += $copilotAnnual
                            $recommendations.Add("COPILOT WATCHLIST — $copilotVariant (€$($copilotPrice.ToString('N2'))/mo) assigned with zero Copilot activity, but user IS active in M365 workloads. Candidate for enablement/training before reclaiming. At-risk spend: €$($copilotPrice.ToString('N2'))/mo (€$($copilotAnnual.ToString('N2'))/yr).")
                        }
                    } else {
                        $copilotAppsStr = $copilotActiveApps -join ", "
                        $recommendations.Add("COPILOT ACTIVE — $copilotVariant license assigned, active in: $copilotAppsStr. Monitor adoption depth for ROI.")
                    }
                } elseif ($copilotUsageLoaded) {
                    # Report loaded but user not in it — Copilot license exists but no activity row at all
                    $userCopilotAnnualCost = $copilotAnnual
                    if (-not $workloadReady) {
                        $copilotNonAdopterCostAcc += $copilotAnnual
                        $copilotReclaimCostAcc += $copilotAnnual
                        $recommendations.Add("COPILOT RECLAIM — $copilotVariant (€$($copilotPrice.ToString('N2'))/mo) assigned but user does not appear in the Copilot usage report AND shows zero M365 workload activity. No readiness for AI adoption. Consider reallocating. Savings: €$($copilotPrice.ToString('N2'))/mo (€$($copilotAnnual.ToString('N2'))/yr).")
                    } else {
                        $copilotWatchlistCostAcc += $copilotAnnual
                        $recommendations.Add("COPILOT WATCHLIST — $copilotVariant (€$($copilotPrice.ToString('N2'))/mo) assigned but not in Copilot usage report (zero Copilot activity). User is active in M365 workloads, candidate for enablement/training. At-risk spend: €$($copilotPrice.ToString('N2'))/mo (€$($copilotAnnual.ToString('N2'))/yr).")
                    }
                } else {
                    # Copilot report unavailable — fall back to proxy (generic M365 app activity)
                    if (-not $workloadReady) {
                        $copilotNonAdopterCostAcc += $copilotAnnual
                        $userCopilotAnnualCost = $copilotAnnual
                        $copilotReclaimCostAcc += $copilotAnnual
                        $recommendations.Add("COPILOT RECLAIM — $copilotVariant (€$($copilotPrice.ToString('N2'))/mo) assigned but no M365 workload activity detected. Copilot-specific usage data was not available in the tenant reports. WARNING: Web-based Copilot Chat (copilot.microsoft.com) is NOT captured in standard reports. Review via M365 Admin Center Copilot dashboard before reclaiming. Savings: €$($copilotPrice.ToString('N2'))/mo (€$($copilotAnnual.ToString('N2'))/yr).")
                    } else {
                        $recommendations.Add("COPILOT ACTIVE — $copilotVariant license assigned, user is active in M365 workloads. Copilot-specific usage data was not available in the tenant reports — monitor via M365 Admin Center Copilot dashboard for adoption metrics.")
                    }
                }
            }
        }
        # ── Security Copilot (SOC analyst tool — metered per Security Compute Unit, NOT per-user productivity) ──
        if ($hasCopilotSecurity) {
            $recommendations.Add("COPILOT — Microsoft Security Copilot assigned. Security Copilot is a SOC/security analyst tool billed by Security Compute Units. Review whether this user actively uses Defender/Sentinel investigations.")
        }
        # ── Copilot Studio (admin/developer tool — NOT an end-user productivity assistant) ──
        # Copilot Studio is used to build and deploy custom chatbots in Teams.  Its usage does not
        # appear in standard M365 productivity reports (Word/Excel/Teams meetings).  Only flag if
        # the user has ZERO sign-in activity at all, not just zero M365 app activity.
        if ($hasCopilotStudio -and -not ($hasCopilotProd -or $hasCopilotBusiness)) {
            if ($lastSignIn -eq "") {
                $recommendations.Add("COPILOT STUDIO — Studio license assigned but no sign-in activity detected. Review whether the license is actively used, or consider reallocating.")
            } else {
                $recommendations.Add("COPILOT STUDIO — license assigned. Studio is an admin/developer tool for building chatbots; productivity metrics (Word/Excel) do not apply. Monitor via Teams admin center.")
            }
        }

        # ── #5b Teams Premium + Copilot AI overlap ──
        # Copilot natively includes Teams Intelligent Recap. Teams Premium is redundant
        # unless the user needs advanced webinar branding/registration controls.
        # Two tiers: 0 meetings organized = definitive removal; >0 = review.
        $hasTeamsPremium = @($userSkuList | Where-Object { $_ -eq "Microsoft_Teams_Premium" }).Count -gt 0
        if ($hasTeamsPremium -and ($hasCopilotProd -or $hasCopilotBusiness)) {
            $tpCost = Get-SkuMonthlyPrice "Microsoft_Teams_Premium"
            $tpAnnual = [math]::Round($tpCost * 12, 2)
            if ($teamsMeetingsOrganized -eq 0) {
                $recommendations.Add("AI ADD-ON OVERLAP — has both Teams Premium (€$($tpCost.ToString('N2'))/mo) and Microsoft 365 Copilot. Copilot natively includes Teams Intelligent Recap, and this user organized 0 meetings in $ReportPeriod (meaning they don't use Premium's advanced webinar/branding features). Teams Premium is likely redundant. Consider removing it. Annual savings: €$($tpAnnual.ToString('N2'))")
            } else {
                $recommendations.Add("AI OVERLAP REVIEW — has both Teams Premium (€$($tpCost.ToString('N2'))/mo) and Microsoft 365 Copilot. Copilot natively includes Teams Intelligent Recap (AI meeting notes/tasks). This user organized $teamsMeetingsOrganized meeting(s) — review whether they require Premium's advanced webinar branding or custom meeting templates before removing. Potential savings: €$($tpCost.ToString('N2'))/mo (€$($tpAnnual.ToString('N2'))/yr).")
            }
        }

        # ── #2 Power BI Pro vs Premium Per User vs Free ──
        # NOTE: Premium Capacity only removes the Pro requirement for CONSUMERS (viewers).
        # Creators/publishers who build reports, publish to workspaces, or use dataflows
        # STILL require Power BI Pro (or Premium Per User).  Graph Reports cannot
        # distinguish creator from consumer — this must be a manual review.
        $hasPbiPro = @($userSkuList | Where-Object { $_ -eq 'POWER_BI_PRO' }).Count -gt 0
        $hasPbiPPU = @($userSkuList | Where-Object { $_ -eq 'POWER_BI_PREMIUM_PER_USER' -or $_ -eq 'POWER_BI_PREMIUM_P' }).Count -gt 0
        if ($hasPbiPro -and $hasPbiPremiumCapacity) {
            $pbiProCost = [math]::Round((Get-SkuMonthlyPrice 'POWER_BI_PRO') * 12, 2)
            $recommendations.Add("POWER BI PRO REVIEW — tenant has Power BI Premium Capacity. Report consumers (viewers) can use the Free license, but creators/publishers who build reports, publish to workspaces, or use dataflows still require Pro. Review this user's role before downgrading. Annual cost: €$($pbiProCost.ToString('N2'))")
        }
        # Premium Per User + Pro overlap check: PPU is a superset of Pro — having both is redundant
        if ($hasPbiPPU -and $hasPbiPro) {
            $pbiProCost = [math]::Round((Get-SkuMonthlyPrice 'POWER_BI_PRO') * 12, 2)
            $recommendations.Add("POWER BI PRO REVIEW — user has both Power BI Premium Per User and Power BI Pro. Premium Per User is a superset of Pro — the separate Pro license is not needed. Consider removing it. Annual savings: €$($pbiProCost.ToString('N2'))")
        }

        # ── #9 Exchange Online Plan 2 right-sizing ──
        $hasExoPlan2 = @($userSkuList | Where-Object { $_ -in $exoPlan2Skus }).Count -gt 0
        # Check for actual productivity suites (not just EMS/security add-on bundles which don't include Exchange)
        $noSuiteWithExo = @($userSkuList | Where-Object { $suiteIncludes.ContainsKey($_) -and -not $addOnBundles.Contains($_) }).Count -eq 0
        # Guard: require mailbox and email report rows to exist (missing data ≠ zero)
        $hasMailboxRow = [bool]$mb
        $hasEmailRow   = [bool]$em
        if ($hasExoPlan2 -and $noSuiteWithExo -and $hasMailboxRow -and $hasEmailRow -and $mbSizeMB -lt 51200 -and $emailIntensity -ne "High" -and -not $isLitigationHold) {
            $exo2Price = Get-SkuMonthlyPrice "EXCHANGEENTERPRISE"
            $exo1Price = Get-SkuMonthlyPrice "EXCHANGESTANDARD"
            $exoSavings = [math]::Round($exo2Price - $exo1Price, 2)
            $exoAnnualSavings = [math]::Round($exoSavings * 12, 2)
            $exoPlan2SavingsAcc += $exoAnnualSavings
            $recommendations.Add("EXO PLAN 2 DOWNGRADE — Exchange Online Plan 2 (€$($exo2Price.ToString('N2'))/mo) assigned but mailbox is ${mbSizeMB} MB (under 50 GB) and usage is not high. Plan 1 (€$($exo1Price.ToString('N2'))/mo, 50 GB, no In-Place Hold) may suffice — saves €$($exoSavings.ToString('N2'))/mo (€$($exoAnnualSavings.ToString('N2'))/yr).")
        } elseif ($hasExoPlan2 -and $noSuiteWithExo -and (-not $hasMailboxRow -or -not $hasEmailRow)) {
            $recommendations.Add("EXO PLAN 2 REVIEW — Exchange Plan 2 assigned but mailbox size and/or activity data is missing from reports. Review size and hold requirements before considering downgrade to Plan 1.")
        }

        # ── EXO Plan 1 storage ceiling — warn if approaching 50 GB limit (mail flow stops) ──
        # Applies to Business Basic/Standard, E1, standalone EXO Plan 1, and any suite with Plan 1 Exchange.
        # Plan 2 (100 GB) and E3/E5 (50 GB primary + auto-expanding archive) are not affected.
        $hasExoPlan1Only = ($hasExchangeEntitlement -and
            $effectiveSkuSet.Contains("EXCHANGESTANDARD") -and
            -not $effectiveSkuSet.Contains("EXCHANGEENTERPRISE") -and
            -not $hasExoPlan2)
        if ($hasExoPlan1Only -and $null -ne $mbSizeMB -and $mbSizeMB -ge 46080) {
            $pctUsed = [math]::Round($mbSizeMB / 51200 * 100, 0)
            $mbSizeDisplay = if ($mbSizeMB -ge 1024) { "$([math]::Round($mbSizeMB / 1024, 1)) GB" } else { "${mbSizeMB} MB" }
            $recommendations.Add("MAILBOX STORAGE WARNING — mailbox is $mbSizeDisplay (${pctUsed}% of 50 GB Plan 1 limit). Mail flow stops at 50 GB. Consider adding the standalone Exchange Online Archiving add-on (~€3/mo) to offload data to an auto-expanding archive. If archive is insufficient, consider upgrading to Exchange Plan 2 (100 GB) or a higher suite.")
        }
        # ── EXO Plan 2 / E3/E5 storage ceiling — primary mailbox hard-caps at 100 GB ──
        # Auto-expanding archive only applies to the archive mailbox, NOT the primary mailbox.
        $hasExoPlan2OrHigher = $effectiveSkuSet.Contains("EXCHANGEENTERPRISE")
        if ($hasExoPlan2OrHigher -and $null -ne $mbSizeMB -and $mbSizeMB -ge 92160) {
            $pctUsed = [math]::Round($mbSizeMB / 102400 * 100, 0)
            $mbSizeDisplay = if ($mbSizeMB -ge 1024) { "$([math]::Round($mbSizeMB / 1024, 1)) GB" } else { "${mbSizeMB} MB" }
            $recommendations.Add("MAILBOX STORAGE WARNING — primary mailbox is $mbSizeDisplay (${pctUsed}% of 100 GB limit). Auto-expanding archive only covers the archive mailbox — primary mailbox hard-caps at 100 GB. Consider moving data to the archive or PST to prevent mail flow stoppage.")
        }

        # ── Exchange Kiosk storage ceiling — hard-caps at 2 GB ──
        # Kiosk mailboxes have a brutal 2 GB limit. Warn at 90% (1843 MB).
        $hasKioskExchange = ($effectiveSkuSet.Contains("EXCHANGEDESKLESS") -and
                            -not $effectiveSkuSet.Contains("EXCHANGESTANDARD") -and
                            -not $effectiveSkuSet.Contains("EXCHANGEENTERPRISE"))
        if ($hasKioskExchange -and $null -ne $mbSizeMB -and $mbSizeMB -ge 1843) {
            $pctUsed = [math]::Round($mbSizeMB / 2048 * 100, 0)
            $mbSizeDisplay = if ($mbSizeMB -ge 1024) { "$([math]::Round($mbSizeMB / 1024, 1)) GB" } else { "${mbSizeMB} MB" }
            $recommendations.Add("MAILBOX STORAGE WARNING — Exchange Kiosk mailbox is $mbSizeDisplay (${pctUsed}% of 2 GB Kiosk limit). Mail flow stops at 2 GB. Consider upgrading to Exchange Plan 1 (50 GB) or archiving/deleting data to free space.")
        }

        # ── Over-licensed Archive — standalone EOA with small mailbox and no archive ──
        # User has Exchange Plan 1 (standalone or from suite) + standalone Exchange Online Archiving,
        # but mailbox is under 25 GB and archive hasn't been provisioned. EOA cost is unnecessary.
        $hasEoaStandalone = @($userSkuList | Where-Object { $_ -eq "EXCHANGE_ARCHIVE" }).Count -gt 0
        if ($hasEoaStandalone -and $hasExchangeEntitlement -and -not $effectiveSkuSet.Contains("EXCHANGEENTERPRISE")) {
            # User has Plan 1 + EOA (not Plan 2 which natively includes archiving)
            if ($null -ne $mbSizeMB -and $mbSizeMB -lt 25600 -and $mbHasArchive -notin @('True','Yes')) {
                $eoaCost = [math]::Round((Get-SkuMonthlyPrice "EXCHANGE_ARCHIVE") * 12, 2)
                $recommendations.Add("OVER-LICENSED ARCHIVE — Exchange Online Archiving add-on (€$((Get-SkuMonthlyPrice 'EXCHANGE_ARCHIVE').ToString('N2'))/mo) assigned but mailbox is only ${mbSizeMB} MB (Plan 1 limit: 50 GB) and no archive mailbox exists. Consider removing the archiving add-on until the primary mailbox approaches the 50 GB limit. Annual savings: €$($eoaCost.ToString('N2'))")
            }
        }

        # ── Redundant Archive on Shared Mailbox — Plan 2 natively includes auto-expanding archives ──
        # Junior admins often panic-buy standalone EOA (€3/mo) for shared mailboxes approaching 100 GB,
        # not realizing Exchange Plan 2 already includes auto-expanding archives. The add-on is 100% redundant.
        # NOTE: The duplicate detection engine ($suiteIncludes) also catches this for ALL users, but this
        # check provides shared-mailbox-specific messaging that is more actionable.
        if ($hasEoaStandalone -and $isSharedMailbox -and $hasExoPlan2) {
            $eoaCostRedundant  = Get-SkuMonthlyPrice "EXCHANGE_ARCHIVE"
            $eoaAnnRedundant   = [math]::Round($eoaCostRedundant * 12, 2)
            $recommendations.Add("REDUNDANT ARCHIVE — Shared mailbox is assigned Exchange Plan 2 AND standalone Exchange Online Archiving (€$($eoaCostRedundant.ToString('N2'))/mo). Plan 2 natively includes auto-expanding archives. Consider removing the Archiving add-on. Annual savings: €$($eoaAnnRedundant.ToString('N2'))")
        }

        # ── OneDrive 1 TB storage ceiling — Business and E1 plans hard-cap at 1 TB ──
        # E3/E5 get 1 TB base expandable to 5 TB (or unlimited) via admin request, so not flagged here.
        $hasOneDrive1TBCap = @($userSkuList | Where-Object { $oneDrive1TBSkus.Contains($_) }).Count -gt 0
        if ($hasOneDrive1TBCap -and $null -ne $odStorageMB -and $odStorageMB -ge 972800) {
            $pctUsed = [math]::Round($odStorageMB / 1048576 * 100, 0)
            $odGB = [math]::Round($odStorageMB / 1024, 1)
            $recommendations.Add("ONEDRIVE STORAGE WARNING — OneDrive is ${odGB} GB (${pctUsed}% of 1 TB limit). Business and E1 plans hard-cap at 1 TB. Sync will break if the limit is reached. Consider adding OneDrive Plan 2 for additional storage. If broader productivity features are also needed, an E3/E5 suite provides 5 TB expandable storage.")
        }

        # ── #7b OneDrive Plan 2 → Plan 1 downgrade (standalone only) ──
        # WACONEDRIVEENTERPRISE (Plan 2, unlimited) costs ~double WACONEDRIVESTANDARD (Plan 1, 1 TB).
        # If storage is comfortably under 900 GB, Plan 1 is sufficient.
        $hasOdPlan2Standalone = @($userSkuList | Where-Object { $_ -eq "WACONEDRIVEENTERPRISE" }).Count -gt 0
        if ($hasOdPlan2Standalone -and $null -ne $odStorageMB) {
            $odGB = [math]::Round($odStorageMB / 1024, 1)
            if ($odStorageMB -lt 921600) {  # 900 GB in MB
                $odP2Price = Get-SkuMonthlyPrice "WACONEDRIVEENTERPRISE"
                $odP1Price = Get-SkuMonthlyPrice "WACONEDRIVESTANDARD"
                $odSavings = [math]::Round($odP2Price - $odP1Price, 2)
                if ($odSavings -gt 0) {
                    $odAnnSavings = [math]::Round($odSavings * 12, 2)
                    $recommendations.Add("ONEDRIVE PLAN 2 REVIEW — has OneDrive Plan 2 (€$($odP2Price.ToString('N2'))/mo, unlimited) but only using $odGB GB. Consider downgrading to OneDrive Plan 1 (€$($odP1Price.ToString('N2'))/mo, 1 TB limit). Potential savings: €$($odSavings.ToString('N2'))/mo (€$($odAnnSavings.ToString('N2'))/yr).")
                }
            }
        }

        # ── #8 F1/F3 frontline right-sizing (enhanced with plan capabilities) ──
        # Skip users already flagged for Tier 1 full removal (dormant, deleted, disabled, no activity, shared mailbox)
        # to prevent double-counting savings in Tier 1 waste + Tier 2 right-sizing.
        # Compute $hasAnyActivity here (before $tier1Removal) — it is also re-used in the no-activity block later.
        $hasAnyActivity = ($emailTotal -gt 0) -or ($teamsTotal -gt 0) -or ($odTotal -gt 0) -or
                          ($spTotal -gt 0) -or $usesDesktop -or $usesWeb -or $usesMobile -or
                          $teamsUsesDesktop -or $teamsUsesMobile -or $teamsUsesWeb
        $hasPremiumSuite = @($userSkuList | Where-Object { $_ -in $premiumSuites }).Count -gt 0
        $tier1Removal = ($isDormant -or (-not $isAccountEnabled) -or (-not $hasAnyActivity -and $au) -or $sharedMbxRemoveLicense)
        if ($hasPremiumSuite -and (-not $isAdmin -or $isLowPrivAdmin) -and -not $tier1Removal) {
            # Pre-compute desktop activation count for Multi-PC gate
            # If user has Office activated on 2+ dedicated devices (Windows OR Mac), F3 VDI-only licensing would break them.
            # Bug fix: filter to Office/M365 Apps Product Type only — summing across ALL Product Types
            # (e.g. Visio + Project) would overcount and produce false "blocked" outcomes.
            $winActTotal = 0
            if ($lkpActivations.ContainsKey($upn)) {
                foreach ($ar in $lkpActivations[$upn]) {
                    if ($ar.'Product Type' -match 'Office|Microsoft 365 Apps|M365 Apps') {
                        $w = [int]($ar.'Windows' -as [int])
                        $m = [int]($ar.'Mac' -as [int])
                        $winActTotal += $w + $m
                    }
                }
            }
            # Cloud PC Enterprise prerequisite check — CPC_E requires Windows Enterprise E3 + Intune + Entra P1.
            # F3 lacks Windows Enterprise E3 → downgrade to E3 instead of F3 for CPC users.
            # ($hasCpcEnterprise computed earlier near $isPhoneResource — available for both frontline and NO ACTIVITY paths)
            # Guard: if M365AppPlatform report is entirely missing, log only (not actionable without data)
            if (-not $app) {
                Write-Log "Frontline right-sizing skipped for $upn — M365 app platform usage data missing" -Level INFO
            # Archive mailbox exists → F3 Exchange Kiosk has ZERO archive rights.
            # If a rescue target (Business Basic / E1) is available, recommend that directly
            # instead of emitting a confusing BLOCKED + RESCUE pair.
            } elseif ($mbHasArchive -in @('True','Yes')) {
                $currentSuiteSku = $userSkuList | Where-Object { $_ -in $premiumSuites } | Select-Object -First 1
                $currentSuiteName = Resolve-SkuFriendlyName $currentSuiteSku
                if (-not $usesDesktop -and ($usesMobile -or $usesWeb -or $teamsUsesMobile -or $teamsUsesWeb)) {
                    $mbDisp = if ($null -ne $mbSizeMB) { "${mbSizeMB} MB primary" } else { "unknown primary size" }
                    # Check if a rescue target exists (E1 or Business Basic — both support archives)
                    # Enterprise-only tenants: always use E1 (no cross-family mixing)
                    # Business/Mixed tenants: prefer Business Basic if under 250-seat cap
                    $rescueAvailable = $false
                    if ((-not $isAdmin -or $isLowPrivAdmin) -and ($null -eq $mbSizeMB -or $mbSizeMB -lt 45000)) {
                        $useBusinessBasic = ($tenantFamily -ne 'Enterprise' -and $businessFamilyTotalConsumed -lt 250)
                        $rescueTarget = if ($useBusinessBasic) { "M365 Business Basic" } else { "Office 365 E1" }
                        $rescueSku    = if ($useBusinessBasic) { "O365_BUSINESS_ESSENTIALS" } else { "STANDARDPACK" }
                        $rescuePrice  = Get-SkuMonthlyPrice $rescueSku
                        $currentPrice = Get-SkuMonthlyPrice $currentSuiteSku
                        # Net savings accounts for compliance add-on costs when user is covered by CA/MDO policies
                        # Neither E1 nor Business Basic include Entra P1 or MDO P1
                        $complianceAddonCost = [decimal]0
                        $complianceNote = ""
                        $needsP1  = ($generalCA -ne '')
                        $needsMdo = $mdoCoverageNonBuiltIn
                        if ($needsP1)  { $complianceAddonCost += Get-SkuMonthlyPrice "AAD_PREMIUM" }
                        if ($needsMdo) { $complianceAddonCost += Get-SkuMonthlyPrice "ATP_ENTERPRISE" }
                        if ($complianceAddonCost -gt 0) {
                            $complianceNote = " Note: user is covered by$(if ($needsP1) { ' Conditional Access' })$(if ($needsP1 -and $needsMdo) { ' and' })$(if ($needsMdo) { ' MDO' }) policies — $(if ($needsP1) { "Entra ID P1 (€$(( Get-SkuMonthlyPrice 'AAD_PREMIUM').ToString('N2'))/mo)" })$(if ($needsP1 -and $needsMdo) { ' and ' })$(if ($needsMdo) { "MDO P1 (€$((Get-SkuMonthlyPrice 'ATP_ENTERPRISE').ToString('N2'))/mo)" }) add-ons are required for compliance."
                        }
                        $rescueSave = [math]::Round(($currentPrice - $rescuePrice - $complianceAddonCost) * 12, 2)
                        if ($rescueSave -gt 0) {
                            $rescueAvailable = $true
                            if ($useBusinessBasic) {
                                $businessFamilyTotalConsumed++
                                $e1ToBasicEligible    = ($standardpackConsumed -gt 0 -and ($businessFamilyTotalConsumed + $standardpackConsumed) -le 250)
                                $e3ToBpEligible       = ($speE3Consumed -gt 0 -and ($businessFamilyTotalConsumed + $speE3Consumed) -le 250)
                                $appsEntToBizEligible = ($appsEntConsumed -gt 0 -and ($businessFamilyTotalConsumed + $appsEntConsumed) -le 250)
                            }
                            $frontlineRescueSavingsAcc += $rescueSave
                            $netMonthlySave = [math]::Round($currentPrice - $rescuePrice - $complianceAddonCost, 2)
                            # Direct recommendation: skip BLOCKED, go straight to the actionable target
                            $recommendations.Add("FRONTLINE RESCUE — has $currentSuiteName (€$($currentPrice.ToString('N2'))/mo) but only uses web/mobile apps (no desktop). User has an active archive mailbox ($mbDisp) so F3 is not suitable, but $rescueTarget (€$($rescuePrice.ToString('N2'))/mo) supports 50 GB mailbox + unlimited archive. Consider downgrading to $rescueTarget.$complianceNote Potential savings: €$($netMonthlySave.ToString('N2'))/mo (€$($rescueSave.ToString('N2'))/yr).")
                        }
                    }
                    if (-not $rescueAvailable) {
                        # No rescue path — emit BLOCKED so analyst knows the archive is the obstacle
                        $recommendations.Add("FRONTLINE BLOCKED — has $currentSuiteName and only uses web/mobile apps, but user has an active archive mailbox ($mbDisp). F3 Exchange Kiosk has zero archive rights — downgrade would permanently destroy archive data. Consider migrating or removing the archive before considering F3.")
                    }
                }
            # HARD BLOCKER: Multi-PC gate (Logic Flaw #2) — Office activated on 2+ Windows PCs means
            # dedicated multi-device setup. F3 only provides VDI shared-device rights; downgrade would
            # deactivate Office on all dedicated PCs.
            } elseif ($winActTotal -gt 1 -and -not $usesDesktop -and ($usesMobile -or $usesWeb -or $teamsUsesMobile -or $teamsUsesWeb) -and -not $teamsUsesDesktop) {
                $currentSuiteSku = $userSkuList | Where-Object { $_ -in $premiumSuites } | Select-Object -First 1
                $currentSuiteName = Resolve-SkuFriendlyName $currentSuiteSku
                $recommendations.Add("FRONTLINE BLOCKED — has $currentSuiteName and only uses web/mobile apps, but Office is activated on $winActTotal Windows devices. F3 only provides VDI shared-device rights — downgrading would deactivate Office on all dedicated PCs. Consider consolidating to a single device or retaining the current license.")
            # Cloud PC Enterprise + web/mobile-only → recommend E3 instead of F3.
            # E3 satisfies all CPC Enterprise prerequisites (Intune, Entra P1, Windows Enterprise E3)
            # and has no F3 storage limitations. No storage gate needed — E3 has full 50 GB mailbox + 1 TB OneDrive.
            } elseif ($hasCpcEnterprise -and -not $usesDesktop -and ($usesMobile -or $usesWeb -or $teamsUsesMobile -or $teamsUsesWeb) -and -not $teamsUsesDesktop) {
                $currentSuiteSku = $userSkuList | Where-Object { $_ -in $premiumSuites } | Select-Object -First 1
                $currentSuiteName = Resolve-SkuFriendlyName $currentSuiteSku
                $currentPrice = Get-SkuMonthlyPrice $currentSuiteSku
                $e3Price = Get-SkuMonthlyPrice "SPE_E3"
                $e3MonthlySave = [math]::Round($currentPrice - $e3Price, 2)
                $e3AnnualSave  = [math]::Round($e3MonthlySave * 12, 2)
                if ($e3MonthlySave -gt 0) {
                    $cpcSkuName = ($userSkuList | Where-Object { $_ -match '^(CPC_E_|Windows_365_E_)' } | ForEach-Object { Resolve-SkuFriendlyName $_ }) -join '; '
                    $frontlineSavingsAcc += $e3AnnualSave
                    $recommendations.Add("FRONTLINE CANDIDATE — has $currentSuiteName (€$($currentPrice.ToString('N2'))/mo) but only uses web/mobile apps (no desktop). User also has a Cloud PC Enterprise license ($cpcSkuName) that requires Windows Enterprise E3, Intune, and Entra ID P1 as prerequisites — all included in M365 E3 but NOT in F3. Consider downgrading to Microsoft 365 E3 (€$($e3Price.ToString('N2'))/mo) to maintain Cloud PC compatibility. Potential savings: €$($e3MonthlySave.ToString('N2'))/mo (€$($e3AnnualSave.ToString('N2'))/yr).")
                }
            # User has E3/E5/Education suite but only uses mobile + web (no desktop apps)
            # Storage gate: skip if mailbox >2 GB or OneDrive >2 GB — F3 Kiosk limits make downgrade impractical
            } elseif (-not $usesDesktop -and ($usesMobile -or $usesWeb -or $teamsUsesMobile -or $teamsUsesWeb) -and -not $teamsUsesDesktop -and ($null -eq $mbSizeMB -or $mbSizeMB -le 2048) -and ($null -eq $odStorageMB -or $odStorageMB -le 2048)) {
                $currentSuiteSku = $userSkuList | Where-Object { $_ -in $premiumSuites } | Select-Object -First 1
                $currentSuiteName = Resolve-SkuFriendlyName $currentSuiteSku
                $currentPrice = Get-SkuMonthlyPrice $currentSuiteSku
                # Determine best frontline target based on needs
                $targetName = "M365 F3"
                $targetPrice = Get-SkuMonthlyPrice "SPE_F1"
                # Check if F1 suffices (no mailbox needed, no OneDrive needed)
                # Guard: only recommend F1 if report data confirms zero usage — missing reports ($em/$od = $null)
                # mean "unknown" not "zero", so default to F3 (which has 2 GB mailbox/OneDrive) as safer choice.
                # (Flaw #3 fix: gate on $em/$od — the actual activity sources — not $ea/$odu app/storage reports)
                $isF1Target = $false
                if ($emailTotal -eq 0 -and $odTotal -eq 0 -and $null -ne $mbSizeMB -and $mbSizeMB -lt 5 -and ($em -and $od)) {
                    $targetName = "M365 F1"
                    $targetPrice = Get-SkuMonthlyPrice "M365_F1"
                    $isF1Target = $true
                }
                $monthlySavings = [math]::Round($currentPrice - $targetPrice, 2)
                $annualSavings  = [math]::Round($monthlySavings * 12, 2)
                # Note: $frontlineSavingsAcc accumulation deferred to after confidence tiering
                # (FRONTLINE REVIEW users with desktop activations should not count as confirmed savings)
                # Build compatibility notes
                $notes = [System.Collections.Generic.List[string]]::new()
                if ($isF1Target) {
                    # F1 critical warning: NO mailbox rights — only Exchange Kiosk for calendar in Teams
                    $notes.Add("F1 has NO mailbox rights (Exchange Kiosk only for calendar in Teams) and NO OneDrive storage — suitable only for users who genuinely need zero email/files")
                }
                # F3 storage limits (Exchange Kiosk 2 GB + OneDrive 2 GB) — only relevant for F3 target
                if (-not $isF1Target) {
                    if ($null -eq $mbSizeMB) {
                        $notes.Add("mailbox size unknown — verify under 2 GB Kiosk limit before downgrading")
                    } elseif ($mbSizeMB -gt 2048) {
                        $notes.Add("mailbox ${mbSizeMB} MB exceeds F3 Exchange Kiosk 2 GB limit — migrate or archive first")
                    } elseif ($mbSizeMB -gt 0) {
                        $notes.Add("mailbox ${mbSizeMB} MB (F3 Kiosk 2 GB limit: $(if ($mbSizeMB -le 2048) {'OK'} else {'OVER'}))")
                    }
                    if ($null -eq $odStorageMB) {
                        $notes.Add("OneDrive size unknown — verify under 2 GB limit before downgrading")
                    } elseif ($odStorageMB -gt 2048) {
                        $notes.Add("OneDrive ${odStorageMB} MB exceeds F3 2 GB limit — migrate data first")
                    } elseif ($odStorageMB -gt 0) {
                        $notes.Add("OneDrive ${odStorageMB} MB (F3 2 GB limit: $(if ($odStorageMB -le 2048) {'OK'} else {'OVER'}))")
                    }
                }
                # Check for existing desktop Office activations (user may not have USED desktop this period but still has it installed)
                $hasDesktopActivations = $false
                $desktopActCount = 0
                if ($lkpActivations.ContainsKey($upn)) {
                    foreach ($ar in $lkpActivations[$upn]) {
                        $winAct = [int]($ar.'Windows' -as [int])
                        $macAct = [int]($ar.'Mac' -as [int])
                        $desktopActCount += $winAct + $macAct
                    }
                    if ($desktopActCount -gt 0) { $hasDesktopActivations = $true }
                }
                if ($hasDesktopActivations -and -not $isF1Target) {
                    $notes.Add("WARNING: $desktopActCount desktop Office activation(s) found — F3 will deactivate Office on all PCs/Macs")
                }
                # Windows Enterprise caveat: E3/E5 include Windows Enterprise E3; F-series only provides VDI rights for shared devices
                # Note: CPC Enterprise users are handled in the dedicated E3 branch above — they won't reach this point.
                if (-not $isF1Target -and $currentSuiteSku -in @("SPE_E3","SPE_E5","MICROSOFT365_E3","MICROSOFT365_E5")) {
                    $notes.Add("Windows: E3/E5 includes Windows Enterprise E3 for dedicated devices; F3 only provides VDI shared-device rights")
                }
                # F3-specific: Viva Insights limited to personal insights only (no manager/leader analytics)
                if (-not $isF1Target) {
                    $notes.Add("F3 Viva Insights limited to personal insights only (no manager/leader analytics available in E3/E5)")
                }
                # Mobile screen limit: F-series mobile apps limited to screens < 10.9 inches
                $notes.Add("F-series mobile apps restricted to screens under 10.9 inches (iPads/tablets may lose edit rights)")
                $noteStr = if ($notes.Count -gt 0) { " NOTE: $($notes -join '; ')." } else { "" }
                # Build activation evidence string for recommendation text
                $actEvidence = if ($activatedPlatforms -eq "") {
                    "0 personal device activations"
                } elseif ($desktopActCount -eq 0) {
                    "activated on mobile only ($activatedPlatforms)"
                } else {
                    "$desktopActCount desktop activation(s) ($activatedPlatforms)"
                }
                # Confidence tiering based on activation evidence:
                # - No activations at all → HIGH CONFIDENCE (no hardware footprint)
                # - Desktop activations exist → REVIEW (user has Office installed on PC/Mac,
                #   may use desktop apps sporadically outside the D90 window — needs manual check)
                # - Mobile-only activations → CANDIDATE (no desktop footprint)
                $confidencePrefix = "FRONTLINE CANDIDATE"
                if ($activatedPlatforms -eq "") {
                    $confidencePrefix = "FRONTLINE CANDIDATE (HIGH CONFIDENCE)"
                } elseif ($hasDesktopActivations) {
                    $confidencePrefix = "FRONTLINE REVIEW"
                }
                $sharedDeviceNote = if ($activatedPlatforms -eq "" -and $isF1Target) {
                    " If this user operates on shared/kiosk hardware, consider Teams Shared Devices license instead."
                } else { "" }
                # Accumulate savings only for confirmed candidates, not review items
                if ($confidencePrefix -ne "FRONTLINE REVIEW") {
                    $frontlineSavingsAcc += $annualSavings
                }
                if ($confidencePrefix -eq "FRONTLINE REVIEW") {
                    $recommendations.Add("$confidencePrefix — has $currentSuiteName (€$($currentPrice.ToString('N2'))/mo) with no desktop Office app usage in D90, but $desktopActCount desktop activation(s) found ($activatedPlatforms). The user may use desktop apps sporadically outside the reporting window. Review whether a downgrade to $targetName (€$($targetPrice.ToString('N2'))/mo) is appropriate — F3 would deactivate Office on all PCs/Macs. Potential savings if confirmed: €$($monthlySavings.ToString('N2'))/mo (€$($annualSavings.ToString('N2'))/yr).$noteStr")
                } else {
                    $recommendations.Add("$confidencePrefix — has $currentSuiteName (€$($currentPrice.ToString('N2'))/mo) but only uses web/mobile apps (no desktop). Activation evidence: $actEvidence. Consider downgrading to $targetName (€$($targetPrice.ToString('N2'))/mo). Potential savings: €$($monthlySavings.ToString('N2'))/mo (€$($annualSavings.ToString('N2'))/yr).$sharedDeviceNote$noteStr")
                }
            }
        }

        # ── F3 to F1 Micro-Downgrade (ideas.md #2) ──
        # F3 (SPE_F1 = €8/mo) includes 2 GB mailbox + 2 GB OneDrive.
        # F1 (M365_F1 = €2.25/mo) has no mailbox and no OneDrive — just Teams + SharePoint read.
        # If an F3 user has zero email/OneDrive activity AND storage is empty, they're over-licensed.
        $f3Skus = @("SPE_F1","DESKLESSPACK")   # M365 F3 and O365 F3
        $hasF3 = @($userSkuList | Where-Object { $_ -in $f3Skus }).Count -gt 0
        if ($hasF3 -and $emailTotal -eq 0 -and $odTotal -eq 0 -and $em -and $od) {
            if (($null -ne $mbSizeMB -and $mbSizeMB -lt 5) -and ($null -ne $odStorageMB -and $odStorageMB -lt 5)) {
                $f3Sku   = ($userSkuList | Where-Object { $_ -in $f3Skus } | Select-Object -First 1)
                $f3Price = Get-SkuMonthlyPrice $f3Sku
                $f1Price = Get-SkuMonthlyPrice "M365_F1"
                $f3f1Savings = [math]::Round(($f3Price - $f1Price) * 12, 2)
                if ($f3f1Savings -gt 0) {
                    $f3Name = Resolve-SkuFriendlyName $f3Sku
                    $recommendations.Add("F3 TO F1 DOWNGRADE — has $f3Name (€$($f3Price.ToString('N2'))/mo) but shows 0 email and 0 OneDrive activity with empty storage (mailbox ${mbSizeMB} MB, OneDrive ${odStorageMB} MB). Consider downgrading to M365 F1 (€$($f1Price.ToString('N2'))/mo). Potential savings: €$([math]::Round($f3Price - $f1Price, 2).ToString('N2'))/mo (€$($f3f1Savings.ToString('N2'))/yr).")
                }
            }
        }

        # ── F-series SKU detection (used by compliance check + Frankenstein Frontline below) ──
        $fSeriesSkus = @("SPE_F1","M365_F1","DESKLESSPACK")
        $isOnFrontlineSku = @($userSkuList | Where-Object { $_ -in $fSeriesSkus }).Count -gt 0

        # ── Frontline Compliance Breach — F-series user with desktop app activations ──
        # F1/F3 licenses only entitle web and mobile Office apps.  Desktop activations (Windows/Mac)
        # of Word, Excel, PowerPoint, Outlook, or OneNote on an F-series SKU are a licensing violation
        # that Microsoft flags in audits.  Detection uses the M365 App activation report.
        if ($isOnFrontlineSku -and $lkpActivations.ContainsKey($upn)) {
            $flDesktopProducts = [System.Collections.Generic.List[string]]::new()
            $flDesktopCount = 0
            foreach ($ar in $lkpActivations[$upn]) {
                $pWin = [int]($ar.'Windows' -as [int])
                $pMac = [int]($ar.'Mac' -as [int])
                if (($pWin + $pMac) -gt 0 -and $ar.'Product Type' -match 'Office|Microsoft 365 Apps|M365 Apps') {
                    $flDesktopCount += $pWin + $pMac
                    $ptName = $ar.'Product Type'
                    if ($ptName -and -not $flDesktopProducts.Contains($ptName)) { [void]$flDesktopProducts.Add($ptName) }
                }
            }
            if ($flDesktopCount -gt 0) {
                $fSkuName = Resolve-SkuFriendlyName ($userSkuList | Where-Object { $_ -in $fSeriesSkus } | Select-Object -First 1)
                $prodList = $flDesktopProducts -join ", "
                $recommendations.Add("LICENSING CHECK — $fSkuName does NOT include desktop Office apps, but $flDesktopCount desktop activation(s) detected ($prodList on Windows/Mac). This may be flagged as a compliance gap during audits. To resolve, consider upgrading to E3/Business Premium for desktop license rights, or consider restricting users to web/mobile access only.")
            }
        }

        # ── "Frankenstein Frontline" — F-series base + expensive add-ons exceeding full suite cost ──
        # F3 (€8) + Exchange Plan 2 (€8) + Entra P2 (€9) + PBI Pro (€9.40) = €34.40/mo
        # That exceeds Business Premium (€22.60/mo) and approaches E3 (€36.20/mo).
        # Upgrade to a full suite to remove F3 restrictions (2 GB storage, 10.9" screen cap).
        if ($isOnFrontlineSku) {
            $fBaseSku   = ($userSkuList | Where-Object { $_ -in $fSeriesSkus } | Select-Object -First 1)
            $fBasePrice = Get-SkuMonthlyPrice $fBaseSku
            $addOnCost  = [math]::Round($userMonthlyCost - $fBasePrice, 2)
            if ($addOnCost -gt 0) {
                $bizPremPrice = Get-SkuMonthlyPrice "SPB"
                $e3Price      = Get-SkuMonthlyPrice "SPE_E3"
                if ($userMonthlyCost -gt $bizPremPrice) {
                    $fBaseName = Resolve-SkuFriendlyName $fBaseSku
                    if ($userMonthlyCost -ge $e3Price) {
                        $recommendations.Add("FRONTLINE ADD-ON STACKING — $fBaseName plus €$($addOnCost.ToString('N2'))/mo in add-ons totals €$($userMonthlyCost.ToString('N2'))/mo, which meets or exceeds M365 E3 (€$($e3Price.ToString('N2'))/mo). Consider upgrading to E3 to eliminate F-series restrictions (2 GB storage cap, 10.9-inch mobile screen limit) and consolidate into 1 SKU with full desktop apps and 1 TB OneDrive.")
                    } elseif ($tenantFamily -ne 'Enterprise') {
                        $recommendations.Add("FRONTLINE ADD-ON STACKING — $fBaseName plus €$($addOnCost.ToString('N2'))/mo in add-ons totals €$($userMonthlyCost.ToString('N2'))/mo, which exceeds Business Premium (€$($bizPremPrice.ToString('N2'))/mo). Consider upgrading to Business Premium to unlock full desktop apps and 1 TB storage. Note: Business SKUs are limited to 300-seat tenants. $crossFamilyNote")
                    }
                }
            }
        }

        # ── Security & Compliance upsell (capability-driven, granular flags) ──
        # Merge capabilities from ALL user SKUs (base suite + add-ons) to avoid false positives.
        # E.g., E3 + EMS E5 = "Defender-complete" even without an explicit Defender Suite SKU.
        $userCaps = Merge-UserCapabilities $userSkuList

        $onBusinessNoSec  = @($userSkuList | Where-Object { $_ -in $businessNoSecurity }).Count -gt 0
        $onBusinessPrem   = @($userSkuList | Where-Object { $_ -in $businessPremiumSkus }).Count -gt 0
        $e3Skus = @("SPE_E3","MICROSOFT365_E3","Microsoft_365_E3_(no_Teams)","O365_w/o Teams Bundle_M3","M365EDU_A3_FACULTY","M365EDU_A3_STUDENT")
        $onEntE3 = @($userSkuList | Where-Object { $_ -in $e3Skus }).Count -gt 0
        $e5Skus = @("SPE_E5","MICROSOFT365_E5","SPE_E5_NOPSTNCONF","SPE_E5_CALLINGMINUTES","Microsoft_365_E5_(no_Teams)","ENTERPRISEPREMIUM","ENTERPRISEPREMIUM_NOPSTNCONF","M365EDU_A5_FACULTY","M365EDU_A5_STUDENT","M365EDU_A5_STUUSEBNFT")
        $onEntE5 = @($userSkuList | Where-Object { $_ -in $e5Skus }).Count -gt 0

        # ── Equivalence class booleans (Flaw 2) ──
        # HasFullDefenderStack: true if user has MDO P2 + MDE P2/Business + MDI + Cloud Apps from ANY combination
        $hasFullDefenderStack = ($userCaps.MdoP2 -and $userCaps.MdeP2OrBusiness -and $userCaps.Mdi -and $userCaps.MdcApps)
        # HasFullPurviewStack: true if user has DLP Email/Files + DLP Teams + AIP P2 + eDiscovery Premium + Insider Risk
        $hasFullPurviewStack  = ($userCaps.DlpEmailFiles -and $userCaps.DlpTeams -and $userCaps.AIPPlan2 -and $userCaps.eDiscoveryPremium -and $userCaps.InsiderRisk)

        # Granular "has any" checks for tiered upsell logic
        $hasAnyDefenderCap = ($userCaps.MdoP1 -or $userCaps.MdoP2 -or $userCaps.MdeP1 -or $userCaps.MdeP2OrBusiness)
        $hasAnyPurviewCap  = ($userCaps.DlpEmailFiles -or $userCaps.DlpTeams -or $userCaps.AIPPlan2 -or $userCaps.eDiscoveryPremium -or $userCaps.InsiderRisk)

        # FLW Defender suppression: if user is on a Frontline plan + has FLW Defender SKUs, they DO have protection
        $hasFlwDefender = @($userSkuList | Where-Object { $_ -in $flwDefenderSkus }).Count -gt 0
        # 10-year audit suppression: if user already has 10-year audit, don't mention "advanced audit" in Purview upsells
        $has10YearAudit = @($userSkuList | Where-Object { $_ -in $audit10YearSkus }).Count -gt 0

        # Short-circuit: skip ALL Defender/Purview upsells if already at E5-equivalent
        if (-not $hasFullDefenderStack) {
            # ── SECURITY GAP: no protection at all ──
            # Suppress if user has FLW Defender add-ons (they have protection through FLW-specific SKUs)
            if ($onBusinessNoSec -and -not $hasAnyDefenderCap -and -not $onBusinessPrem -and -not $hasFlwDefender) {
                $bizSku = ($userSkuList | Where-Object { $_ -in $businessNoSecurity } | ForEach-Object { Resolve-SkuFriendlyName $_ }) -join "; "
                $recommendations.Add("SECURITY GAP — $bizSku includes no endpoint or email threat protection. Options: (1) Consider upgrading to Business Premium (adds Defender for Business, MDO P1, Intune, Entra ID P1), (2) add standalone Defender for Business, or (3) add Microsoft Defender Suite for Business (Entra P2, Defender for Identity, Endpoint P2, MDO P2, Cloud Apps).")
            }

            # ── DEFENDER COVERAGE REVIEW: partial coverage, show what's missing (granular) ──
            if ($hasAnyDefenderCap) {
                $missingDef = @()
                if (-not $userCaps.MdoP2)           { $missingDef += "MDO P2" }
                if (-not $userCaps.MdeP2OrBusiness)  { $missingDef += "Endpoint P2" }
                if (-not $userCaps.Mdi)              { $missingDef += "Identity" }
                if (-not $userCaps.MdcApps)           { $missingDef += "Cloud Apps" }

                if ($onBusinessNoSec -or $onBusinessPrem) {
                    $defBizTier = if ($onBusinessPrem) { "Business Premium" } else {
                        ($userSkuList | Where-Object { $_ -in $businessNoSecurity } | ForEach-Object { Resolve-SkuFriendlyName $_ } | Select-Object -First 1)
                    }
                    $recommendations.Add("DEFENDER COVERAGE REVIEW — $defBizTier has partial Defender coverage. Missing: $($missingDef -join ', '). Consider adding Microsoft Defender Suite for Business for comprehensive protection.")
                } elseif ($onEntE3 -and -not $onEntE5) {
                    $recommendations.Add("DEFENDER COVERAGE REVIEW — E3 user has partial Defender coverage (missing: $($missingDef -join ', ')). Consider adding the Microsoft Defender Suite add-on or a full E5 upgrade if multiple add-ons are stacking up.")
                }
            }
        }

        if (-not $hasFullPurviewStack) {
            # Build dynamic Purview feature list (suppress "advanced audit" if user has 10-year audit add-on)
            $purviewFeatures = @("eDiscovery Premium")
            if (-not $has10YearAudit) { $purviewFeatures += "advanced audit" }
            $purviewFeatures += @("insider risk management","records management")
            $purviewFeatureStr = $purviewFeatures -join ", "

            # ── COMPLIANCE COVERAGE REVIEW: only if no advanced compliance ──
            if (($onBusinessNoSec -or $onBusinessPrem) -and -not $hasAnyPurviewCap) {
                if ($hasFullDefenderStack) {
                    $recommendations.Add("COMPLIANCE COVERAGE REVIEW — Defender Suite coverage is complete but no advanced compliance detected. Consider adding Microsoft Purview Suite for Business to complete the stack: $purviewFeatureStr.")
                } else {
                    $recommendations.Add("COMPLIANCE COVERAGE REVIEW — no advanced compliance coverage detected. Consider Microsoft Purview add-ons for $purviewFeatureStr.")
                }
            }
            # Enterprise E3 with Defender-complete but no Purview
            if ($onEntE3 -and -not $onEntE5 -and $hasFullDefenderStack -and -not $hasAnyPurviewCap) {
                $recommendations.Add("COMPLIANCE COVERAGE REVIEW — E3 user has Defender-complete coverage but no advanced compliance. Consider adding the Microsoft Purview Suite add-on for $purviewFeatureStr, or consider full E5 upgrade to consolidate.")
            }
        }

        # ── Behavior-Driven Security Risk: Heavy External Sharer (ideas.md #3) ──
        # If a user actively shares many files externally but lacks DLP/Information Protection,
        # they are a data exfiltration risk. Provides evidence-based Purview upsell.
        $odExtShared = if ($od) { Parse-NumericField $od.'Shared Externally File Count' } else { 0 }
        $spExtShared = if ($sp) { Parse-NumericField $sp.'Shared Externally File Count' } else { 0 }
        $totalExtShared = $odExtShared + $spExtShared
        if ($totalExtShared -gt 25 -and -not $hasAnyPurviewCap) {
            $recommendations.Add("EXTERNAL SHARING REVIEW — user shared $totalExtShared files externally in $ReportPeriod but lacks advanced compliance coverage (DLP/Information Protection). If this volume or data sensitivity requires compliance controls, consider adding Purview add-ons (DLP/Information Protection).")
        }

        # ── Specialist add-on informational flags (Flaw 2: long-tail awareness) ──
        # Emit informational LICENSING CHECK notes for specialist add-ons, and suppress generic
        # upsell noise when the user already has advanced coverage through these add-ons.
        $userAdvancedAddons = @($userSkuList | Where-Object { $advancedSecCompAddons.Contains($_) })
        if ($userAdvancedAddons.Count -gt 0) {
            $hasEntraGov   = @($userSkuList | Where-Object { $_ -in $entraGovSkus }).Count -gt 0
            $hasEntraSuite = @($userSkuList | Where-Object { $_ -in $entraSuiteSkus }).Count -gt 0
            $hasIntuneSuite = @($userSkuList | Where-Object { $_ -in $intuneSuiteSkus }).Count -gt 0
            # Entra Governance + Entra Suite overlap check (strengthened from LICENSING CHECK)
            if ($hasEntraGov -and $hasEntraSuite) {
                [decimal]$entraGovCost = 0
                foreach ($eg in $userSkuList) { if ($eg -in $entraGovSkus) { $entraGovCost += Get-SkuMonthlyPrice $eg } }
                $entraGovAnnual = [math]::Round($entraGovCost * 12, 2)
                $recommendations.Add("ENTRA SUITE OVERLAP — Entra Suite (€$((Get-SkuMonthlyPrice 'ENTRA_SUITE').ToString('N2'))/mo) natively includes Entra ID P2 and Governance. Consider removing the standalone Governance add-on(s) to save €$($entraGovCost.ToString('N2'))/mo (€$($entraGovAnnual.ToString('N2'))/yr).")
            }
            # Reverse consolidation: standalone Entra ID P2 + Governance → Entra Suite bundle
            if (-not $hasEntraSuite -and $hasStandaloneP2 -and $hasEntraGov) {
                $p2Price    = Get-SkuMonthlyPrice "AAD_PREMIUM_P2"
                $govPrice   = Get-SkuMonthlyPrice "ENTRA_ID_GOVERNANCE"
                $alaCarte   = $p2Price + $govPrice
                $suitePrice = Get-SkuMonthlyPrice "ENTRA_SUITE"
                $entraDelta = [math]::Round($alaCarte - $suitePrice, 2)
                if ($entraDelta -gt 0) {
                    $entraAnnSave = [math]::Round($entraDelta * 12, 2)
                    $recommendations.Add("BUNDLE CONSOLIDATION — has Entra ID P2 (€$($p2Price.ToString('N2'))/mo) + Governance (€$($govPrice.ToString('N2'))/mo) = €$($alaCarte.ToString('N2'))/mo. Consider consolidating into Entra Suite (€$($suitePrice.ToString('N2'))/mo) to save €$($entraDelta.ToString('N2'))/mo (€$($entraAnnSave.ToString('N2'))/yr) and also gain Internet/Private Access.")
                }
            }
            # Intune Suite / premium add-on waste (2026 licensing change)
            # Microsoft rolled Intune Remote Help, Advanced Analytics, and Endpoint Privilege Management
            # natively into M365 E3/E5 in late 2025.  Standalone add-ons are now redundant.
            if ($hasIntuneSuite -and ($onEntE3 -or $onEntE5)) {
                $intuneAddonList = @($userSkuList | Where-Object { $_ -in $intuneSuiteSkus })
                [decimal]$intuneAddonCost = 0
                foreach ($ia in $intuneAddonList) { $intuneAddonCost += Get-SkuMonthlyPrice $ia }
                $intuneAnnual = [math]::Round($intuneAddonCost * 12, 2)
                $intuneAddons = ($intuneAddonList | ForEach-Object { Resolve-SkuFriendlyName $_ }) -join "; "
                $entTier = if ($onEntE5) { "M365 E5" } else { "M365 E3" }
                $recommendations.Add("INTUNE SUITE OVERLAP — $intuneAddons (€$($intuneAddonCost.ToString('N2'))/mo) assigned alongside $entTier. Microsoft rolled Intune Remote Help, Advanced Analytics, and Endpoint Privilege Management into M365 E3/E5 in late 2025. Consider removing the standalone add-on(s) to save €$($intuneAddonCost.ToString('N2'))/mo (€$($intuneAnnual.ToString('N2'))/yr).")
            }
            # Entra Suite + E5 consolidation opportunity
            if ($hasEntraSuite -and ($onEntE3 -or $onEntE5)) {
                $recommendations.Add("LICENSING CHECK — Entra Suite present. For E3/E5 users, review whether consolidating into M365 E5 plus a smaller Entra footprint would be cheaper than maintaining separate Entra Suite add-ons.")
            }
        }

        # ── Intune/EMS Shelfware & Web-Only Ghosting ──
        # Only flag when Intune comes from a standalone SKU (e.g. INTUNE_A, EMS, EMSPREMIUM).
        # When Intune is bundled in a suite (E3/E5/Business Premium), the entitlement is not
        # independently removable and flagging it as shelfware is not actionable.
        $hasStandaloneIntune = $false
        $intuneStandaloneSkus = @("INTUNE_A","EMS","EMSPREMIUM","INTUNE_SMB","INTUNE_P1","INTUNE_EDU")
        foreach ($iSku in $userSkuList) {
            if ($iSku -in $intuneStandaloneSkus) { $hasStandaloneIntune = $true; break }
        }
        if ($userCaps['IntunePlan1'] -eq $true -and $hasStandaloneIntune) {
            $userManagedDevices = if ($managedDevicesLoaded -and $lkpManagedDeviceCount.ContainsKey($upn)) { $lkpManagedDeviceCount[$upn] } else { $null }
            if ($managedDevicesLoaded -and ($null -eq $userManagedDevices -or $userManagedDevices -eq 0)) {
                # Tier 1: zero enrolled devices — standalone Intune/EMS is pure shelfware
                $intuneSkuNames = @($userSkuList | Where-Object { $_ -in $intuneStandaloneSkus } | ForEach-Object { Resolve-SkuFriendlyName $_ }) -join ", "
                $recommendations.Add("INTUNE REVIEW — user holds a standalone Intune/EMS license ($intuneSkuNames) but has 0 enrolled devices in Intune. The MDM/MAM capability is entirely unused. Consider removing the standalone Intune/EMS license or enrolling devices.")
            } elseif (-not $usesDesktop -and -not $usesMobile -and -not $teamsUsesDesktop -and -not $teamsUsesMobile -and ($usesWeb -or $teamsUsesWeb)) {
                # Tier 2: web-only access — devices may be enrolled but user never touches desktop/mobile apps
                $recommendations.Add("INTUNE REVIEW — user holds a standalone Intune/EMS license but telemetry shows 100% web-only access (no desktop apps, no mobile apps, no Teams desktop/mobile). Intune device/app management is unutilized. If security is required for web access, Entra ID P1 alone (via Conditional Access) is sufficient.")
            }
        }

        # ── Cloud Licensing API checks (trial, capacity queue, assignment errors) ──
        if ($cloudLicensingLoaded) {
            # Trial subscription detection
            foreach ($sku in $userSkuList) {
                if ($cloudLicensingData.ContainsKey($sku) -and $cloudLicensingData[$sku].IsTrial) {
                    $userIsOnTrial = $true
                    $trialFriendly = Resolve-SkuFriendlyName $sku
                    $trialDaysLeft = ""
                    $nlDate = $cloudLicensingData[$sku].NextLifecycle
                    if ($nlDate) {
                        try {
                            $dLeft = [int]([datetime]$nlDate - (Get-Date)).TotalDays
                            $trialDaysLeft = " ($dLeft days remaining)"
                        } catch { }
                    }
                    $recommendations.Add("TRIAL LICENSE — $trialFriendly is on a trial subscription$trialDaysLeft. Consider converting to a paid license or removing before the trial expires.")
                }
            }

            # Waiting member (license capacity queue)
            $userId = if ($upnToId.ContainsKey($upn.ToLower())) { $upnToId[$upn.ToLower()] } else { $null }
            if ($userId -and $clWaitingMembers.Contains($userId)) {
                $recommendations.Add("LICENSE CAPACITY QUEUE — user is in the waiting room for a license allotment due to insufficient available seats. Consider purchasing additional licenses or removing assignments from inactive users.")
            }

            # Cloud licensing assignment errors
            if ($userId -and $clAssignmentErrors.ContainsKey($userId)) {
                $userCloudErrors = ($clAssignmentErrors[$userId] | Select-Object -Unique) -join "; "
                $recommendations.Add("CLOUD LICENSE SYNC ERROR — license assignment synchronization failed: $userCloudErrors. Check Cloud Licensing allotment assignments in Entra ID.")
            }
        }

        # ── Business Standard → Business Basic downgrade ──
        # NOTE: Both Standard and Basic include full Teams desktop — so $teamsUsesDesktop is irrelevant.
        # The only relevant signal is $usesDesktop (M365 desktop Office apps: Word/Excel/PowerPoint/etc.)
        $hasBizStandard = @($userSkuList | Where-Object { $_ -in $businessStandardSkus }).Count -gt 0
        if ($hasBizStandard -and (-not $isAdmin -or $isLowPrivAdmin)) {
            if (-not $app) {
                # M365AppPlatform report missing — cannot determine desktop usage (LOA v1.0 spec §5.3)
                $recommendations.Add("BUSINESS BASIC REVIEW — has Business Standard but M365 app platform usage data is missing. Review desktop app dependency before downgrading to Business Basic.")
            } elseif (-not $usesDesktop -and ($usesWeb -or $usesMobile)) {
                $stdSku     = ($userSkuList | Where-Object { $_ -in $businessStandardSkus } | Select-Object -First 1)
                $stdPrice   = Get-SkuMonthlyPrice $stdSku
                $basicPrice = Get-SkuMonthlyPrice "O365_BUSINESS_ESSENTIALS"
                $savings    = [math]::Round($stdPrice - $basicPrice, 2)
                $annSavings = [math]::Round($savings * 12, 2)
                $businessBasicSavingsAcc += $annSavings
                # Activation evidence for Business Basic recommendation
                $bbActEvidence = if ($activatedPlatforms -eq "") {
                    " Activation evidence: 0 personal device activations."
                } elseif ($lkpActivations.ContainsKey($upn)) {
                    $bbDesktop = 0
                    foreach ($ar in $lkpActivations[$upn]) {
                        $bbDesktop += $([int]($ar.'Windows' -as [int]))
                        $bbDesktop += $([int]($ar.'Mac' -as [int]))
                    }
                    if ($bbDesktop -eq 0) { " Activation evidence: activated on mobile only ($activatedPlatforms)." }
                    else { " Activation evidence: $bbDesktop desktop activation(s) ($activatedPlatforms) but zero desktop app usage in lookback." }
                } else { "" }
                $recommendations.Add("BUSINESS BASIC CANDIDATE — has Business Standard (€$($stdPrice.ToString('N2'))/mo) but only uses web/mobile apps (no desktop). Consider downgrading to Business Basic (€$($basicPrice.ToString('N2'))/mo). Potential savings: €$($savings.ToString('N2'))/mo (€$($annSavings.ToString('N2'))/yr).$bbActEvidence")
            }
        }

        # ── E1 → Business Basic arbitrage ──
        # STANDARDPACK (O365 E1) and Business Basic have identical web/mobile capabilities
        # but E1 costs ~€2.70/mo more. If the tenant's business family count stays under 250
        # after migrating all E1 users, flag as downgrade candidate.
        # Enterprise-only tenants: skip — do not introduce Business licenses into Enterprise tenants.
        $hasE1 = @($userSkuList | Where-Object { $_ -eq "STANDARDPACK" }).Count -gt 0
        if ($hasE1 -and $e1ToBasicEligible -and (-not $isAdmin -or $isLowPrivAdmin) -and $tenantFamily -ne 'Enterprise') {
            # Skip users with enterprise add-ons that require an enterprise base license
            $enterpriseAddOns = @($userSkuList | Where-Object { $_ -in $e5AddOns })
            if ($enterpriseAddOns.Count -eq 0) {
                $e1Price    = Get-SkuMonthlyPrice "STANDARDPACK"
                $bbPrice    = Get-SkuMonthlyPrice "O365_BUSINESS_ESSENTIALS"
                $e1Savings  = [math]::Round($e1Price - $bbPrice, 2)
                $e1AnnSave  = [math]::Round($e1Savings * 12, 2)
                if ($e1Savings -gt 0) {
                    $e1DowngradeSavingsAcc += $e1AnnSave
                    $recommendations.Add("E1 DOWNGRADE CANDIDATE — has Office 365 E1 (€$($e1Price.ToString('N2'))/mo) but tenant is under the 300-seat Business cap ($($businessFamilyTotalConsumed + $standardpackConsumed)/300). Consider downgrading to M365 Business Basic (€$($bbPrice.ToString('N2'))/mo) for identical web/mobile capabilities. Potential savings: €$($e1Savings.ToString('N2'))/mo (€$($e1AnnSave.ToString('N2'))/yr).")
                }
            }
        }

        # ── Office 365 E3 → E1 Downgrade (desktop-less enterprise) ──
        # ENTERPRISEPACK (O365 E3, ~€23.20) includes desktop apps + 100 GB mailbox.
        # If user only uses web/mobile and mailbox < 50 GB, downgrade to STANDARDPACK (O365 E1, ~€8.70).
        $hasO365E3 = @($userSkuList | Where-Object { $_ -eq "ENTERPRISEPACK" }).Count -gt 0
        if ($hasO365E3 -and -not $usesDesktop -and ($usesWeb -or $usesMobile) -and $app -and (-not $isAdmin -or $isLowPrivAdmin)) {
            if ($null -ne $mbSizeMB -and $mbSizeMB -lt 45000) {
                $o365E3Price = Get-SkuMonthlyPrice "ENTERPRISEPACK"
                $o365E1Price = Get-SkuMonthlyPrice "STANDARDPACK"
                $o365E3Save  = [math]::Round(($o365E3Price - $o365E1Price) * 12, 2)
                if ($o365E3Save -gt 0) {
                    $o365E3DowngradeSavingsAcc += $o365E3Save
                    $recommendations.Add("O365 E3 TO E1 — holds Office 365 E3 (€$($o365E3Price.ToString('N2'))/mo) but uses web/mobile apps only (no desktop activations). Mailbox ($($mbSizeMB) MB) is under E1's 50 GB limit. Consider downgrading to Office 365 E1 (€$($o365E1Price.ToString('N2'))/mo). Potential savings: €$(([math]::Round($o365E3Price - $o365E1Price, 2)).ToString('N2'))/mo (€$($o365E3Save.ToString('N2'))/yr).")
                }
            }
        }

        # ── M365 E3 → Business Premium Arbitrage ──
        # M365 Business Premium (SPB, ~€22.60) includes desktop apps + Intune + Defender for Business.
        # E3 (SPE_E3, ~€34.90) has 100 GB mailbox and some compliance features, but BP is cheaper
        # and actually includes better endpoint security for SMBs.
        # Enterprise-only tenants: skip — do not introduce Business licenses into Enterprise tenants.
        $hasM365E3 = @($userSkuList | Where-Object { $_ -eq "SPE_E3" }).Count -gt 0
        if ($hasM365E3 -and $e3ToBpEligible -and (-not $isAdmin -or $isLowPrivAdmin) -and ($usesDesktop -or $teamsUsesDesktop) -and $tenantFamily -ne 'Enterprise') {
            # Skip users with enterprise add-ons that require an enterprise base license
            $enterpriseAddOnsE3 = @($userSkuList | Where-Object { $_ -in $e5AddOns })
            if ($enterpriseAddOnsE3.Count -eq 0 -and $null -ne $mbSizeMB -and $mbSizeMB -lt 45000) {
                $e3Price   = Get-SkuMonthlyPrice "SPE_E3"
                $bpPrice   = Get-SkuMonthlyPrice "SPB"
                $e3Savings = [math]::Round($e3Price - $bpPrice, 2)
                $e3AnnSave = [math]::Round($e3Savings * 12, 2)
                if ($e3Savings -gt 0) {
                    $e3DowngradeSavingsAcc += $e3AnnSave
                    $recommendations.Add("E3 TO BUSINESS PREMIUM — holds M365 E3 (€$($e3Price.ToString('N2'))/mo). Tenant has spare Business-tier capacity ($($businessFamilyTotalConsumed + $speE3Consumed)/300) and mailbox is under 50 GB. Consider downgrading to M365 Business Premium (€$($bpPrice.ToString('N2'))/mo) for same apps + better endpoint security. Potential savings: €$($e3Savings.ToString('N2'))/mo (€$($e3AnnSave.ToString('N2'))/yr).")
                }
            }
        }

        # ── Business Premium Inversion (Standard + Security/Compliance Add-ons) ──
        # Business Standard (€12.50) + Defender Suite for Business (€6.00) = €18.50/mo.
        # Business Premium (€22.60) includes Intune + Entra P1 + Defender for Business.
        # When add-ons push total ≥ Premium price, upgrading is cheaper AND adds capabilities.
        $bizStdSkus = @("M365_BUSINESS_STANDARD","MICROSOFT_365_BUSINESS_STANDARD_NO_TEAMS",
                        "Microsoft_365_Business_Standard_EEA_(no_Teams)","Office_365_w/o_Teams_Bundle_Business_Standard",
                        "O365_BUSINESS_PREMIUM")
        $hasBizStd = @($userSkuList | Where-Object { $_ -in $bizStdSkus }).Count -gt 0
        if ($hasBizStd) {
            $bizSecAddons = @("M365_DEFENDER_SUITE_BUSINESS","M365_PURVIEW_SUITE_BUSINESS",
                              "MDE_SMB","DEFENDER_BUSINESS","DEFENDER_BUSINESS_PREMIUM",
                              "INTUNE_SMB","INTUNE_A","AAD_PREMIUM")
            $userBizAddons = @($userSkuList | Where-Object { $_ -in $bizSecAddons })
            if ($userBizAddons.Count -gt 0) {
                $stdSku   = ($userSkuList | Where-Object { $_ -in $bizStdSkus } | Select-Object -First 1)
                $stdPrice = Get-SkuMonthlyPrice $stdSku
                $addonTotal = [decimal]0
                foreach ($addon in $userBizAddons) { $addonTotal += Get-SkuMonthlyPrice $addon }
                $totalAlaCartePrice = $stdPrice + $addonTotal
                $bpInvPrice = Get-SkuMonthlyPrice "SPB"
                if ($totalAlaCartePrice -ge $bpInvPrice) {
                    $bpInvSavings = [math]::Round(($totalAlaCartePrice - $bpInvPrice) * 12, 2)
                    $bizPremInversionSavingsAcc += $bpInvSavings
                    $addonNames = ($userBizAddons | ForEach-Object { Resolve-SkuFriendlyName $_ }) -join ' + '
                    $recommendations.Add("BUSINESS PREMIUM INVERSION — Business Standard (€$($stdPrice.ToString('N2'))/mo) + $addonNames = €$($totalAlaCartePrice.ToString('N2'))/mo. Consider upgrading to M365 Business Premium (€$($bpInvPrice.ToString('N2'))/mo) which natively includes Intune, Entra ID P1 and Defender for Business. Potential savings: €$(([math]::Round($totalAlaCartePrice - $bpInvPrice, 2)).ToString('N2'))/mo (€$($bpInvSavings.ToString('N2'))/yr).")
                }
            }
        }

        # ── Business Premium Security Overlap Review ──
        # Business Premium natively includes Defender for Business (MDE_SMB) + MDO P1 (ATP_ENTERPRISE).
        # NOTE: standalone MDE_SMB/DEFENDER_BUSINESS are already caught by the duplicate detection engine.
        # The Defender Suite for Business adds MDI, Defender for Cloud Apps, Entra P2, and MDO P2 — genuine value.
        # Flag as a review: verify the advanced capabilities justify the add-on cost.
        if ($onBusinessPrem) {
            $hasDefSuiteBiz = @($userSkuList | Where-Object { $_ -eq "M365_DEFENDER_SUITE_BUSINESS" }).Count -gt 0
            if ($hasDefSuiteBiz) {
                $defSuitePrice = Get-SkuMonthlyPrice "M365_DEFENDER_SUITE_BUSINESS"
                $defSuiteAnn   = [math]::Round($defSuitePrice * 12, 2)
                # Check if the advanced capabilities in the Defender Suite are also covered by other SKUs
                $hasEntraP2Already = ($effectiveSkuSet.Contains("AAD_PREMIUM_P2"))
                $hasMdiAlready     = ($effectiveSkuSet.Contains("ATA"))
                $hasMdcaAlready    = ($effectiveSkuSet.Contains("ADALLOM_S_STANDALONE"))
                $overlapParts = @()
                if ($hasEntraP2Already) { $overlapParts += "Entra P2" }
                if ($hasMdiAlready)     { $overlapParts += "Defender for Identity" }
                if ($hasMdcaAlready)    { $overlapParts += "Defender for Cloud Apps" }
                $overlapNote = if ($overlapParts.Count -gt 0) { " Additionally, $($overlapParts -join ', ') already present from other SKUs — partial redundancy." } else { "" }
                $recommendations.Add("BUSINESS PREMIUM SECURITY REVIEW — Business Premium already includes Defender for Business (MDE) and MDO P1. The Defender Suite for Business (€$($defSuitePrice.ToString('N2'))/mo) adds MDI, Defender for Cloud Apps, Entra ID P2 and MDO P2. Review whether these advanced capabilities are actively used to justify €$($defSuiteAnn.ToString('N2'))/yr.$overlapNote")
            }
        }

        # ── E5 Voice Shelfware (swap to No-PSTN variant) ──
        # Full E5 SKUs (M365 E5 and Office 365 E5) include Audio Conferencing + Teams Phone.
        # If user organized 0 meetings AND made 0 calls, swap to the No-PSTN variant to drop unused telecom costs.
        # Map each E5 SKU to its correct No-PSTN equivalent:
        #   SPE_E5 / MICROSOFT365_E5              → SPE_E5_NOPSTNCONF (M365 E5 No Audio Conferencing)
        #   ENTERPRISEPREMIUM                     → ENTERPRISEPREMIUM_NOPSTNCONF (O365 E5 No Audio Conferencing)
        #   M365EDU_A5_FACULTY / M365EDU_A5_STUDENT → SPE_E5_NOPSTNCONF (closest Education equivalent)
        # Suppress E5 Voice Review when TEAMS UNBUNDLING already fires — if the user has 0 Teams activity
        # and is being told to unbundle Teams entirely, the "swap to No Audio Conferencing" rec is redundant.
        $teamsUnbundlingFired = ($hasBundledTeamsSku -and $hasTeamsEntitlement -and $teamsTotal -eq 0 -and $au -and -not $nonHumanReviewFired)
        $matchedE5Sku = ($userSkuList | Where-Object { $_ -in @("SPE_E5","MICROSOFT365_E5","ENTERPRISEPREMIUM","M365EDU_A5_FACULTY","M365EDU_A5_STUDENT","M365EDU_A5_STUUSEBNFT") } | Select-Object -First 1)
        if ($matchedE5Sku -and $teamsCalls -eq 0 -and $teamsMeetingsOrganized -eq 0 -and $tm -and -not $teamsUnbundlingFired) {
            $e5Price       = Get-SkuMonthlyPrice $matchedE5Sku
            $e5NoPstnSku   = if ($matchedE5Sku -eq "ENTERPRISEPREMIUM") { "ENTERPRISEPREMIUM_NOPSTNCONF" } else { "SPE_E5_NOPSTNCONF" }
            $e5NoPstnPrice = Get-SkuMonthlyPrice $e5NoPstnSku
            $voiceSavings  = [math]::Round(($e5Price - $e5NoPstnPrice) * 12, 2)
            if ($voiceSavings -gt 0) {
                $e5VoiceSavingsAcc += $voiceSavings
                $e5FriendlyName    = Resolve-SkuFriendlyName $matchedE5Sku
                $noPstnFriendly    = Resolve-SkuFriendlyName $e5NoPstnSku
                $recommendations.Add("E5 VOICE REVIEW — holds $e5FriendlyName (€$($e5Price.ToString('N2'))/mo) but organized 0 meetings and made 0 Teams calls. Consider swapping to $noPstnFriendly (€$($e5NoPstnPrice.ToString('N2'))/mo) to remove unused telecom costs. Potential savings: €$(([math]::Round($e5Price - $e5NoPstnPrice, 2)).ToString('N2'))/mo (€$($voiceSavings.ToString('N2'))/yr).")
            }
        }

        # ── Apps for Enterprise → Apps for Business Arbitrage ──
        # OFFICESUBSCRIPTION (Enterprise, ~€13.20) and O365_BUSINESS (Business, ~€11.50) are identical desktop apps.
        # If tenant is under 300-seat cap, the cheaper Business variant saves €1.70/mo per user.
        $hasAppsEnt = @($userSkuList | Where-Object { $_ -eq "OFFICESUBSCRIPTION" }).Count -gt 0
        if ($hasAppsEnt -and $appsEntToBizEligible) {
            $entAppPrice = Get-SkuMonthlyPrice "OFFICESUBSCRIPTION"
            $bizAppPrice = Get-SkuMonthlyPrice "O365_BUSINESS"
            $appArbSavings = [math]::Round(($entAppPrice - $bizAppPrice) * 12, 2)
            if ($appArbSavings -gt 0) {
                $appArbitrageSavingsAcc += $appArbSavings
                $recommendations.Add("APP ARBITRAGE — holds Apps for Enterprise (€$($entAppPrice.ToString('N2'))/mo). Tenant has spare Business-tier capacity ($($businessFamilyTotalConsumed + $appsEntConsumed)/300). Consider downgrading to Apps for Business (€$($bizAppPrice.ToString('N2'))/mo) for identical desktop applications. Potential savings: €$(([math]::Round($entAppPrice - $bizAppPrice, 2)).ToString('N2'))/mo (€$($appArbSavings.ToString('N2'))/yr).")
            }
        }

        # ── Power BI PPU Add-on Arbitrage ──
        # Standalone PPU (~€22.50) includes Pro + Premium features. If user already gets Pro from a suite
        # (E5, O365 E5), they only need the PPU Add-On (~€8.50) which layers on top of the included Pro.
        $hasPpuStandalone = @($userSkuList | Where-Object { $_ -eq "POWER_BI_PREMIUM_PER_USER" }).Count -gt 0
        $hasProViaSuite = ($effectiveSkuSet.Contains("POWER_BI_PRO") -and -not (@($userSkuList | Where-Object { $_ -eq "POWER_BI_PRO" }).Count -gt 0))
        if ($hasPpuStandalone -and $hasProViaSuite) {
            $ppuPrice      = Get-SkuMonthlyPrice "POWER_BI_PREMIUM_PER_USER"
            $ppuAddonPrice = Get-SkuMonthlyPrice "PBI_PREMIUM_PER_USER_ADDON"
            if ($ppuAddonPrice -le 0) { $ppuAddonPrice = [math]::Round($ppuPrice / 2, 2) }
            $ppuSavings    = [math]::Round($ppuPrice - $ppuAddonPrice, 2)
            if ($ppuSavings -gt 0) {
                $ppuAnnSavings = [math]::Round($ppuSavings * 12, 2)
                $ppuArbitrageSavingsAcc += $ppuAnnSavings
                $recommendations.Add("PBI PPU OVERLAP — has full Power BI Premium Per User (€$($ppuPrice.ToString('N2'))/mo) but already gets Power BI Pro from their base suite. Consider swapping to the PPU Add-On (€$($ppuAddonPrice.ToString('N2'))/mo) which layers Premium features on top of the included Pro. Potential savings: €$($ppuSavings.ToString('N2'))/mo (€$($ppuAnnSavings.ToString('N2'))/yr).")
            }
        }

        # ── Standalone Desktop App Waste (ideas.md #1) ──
        # M365 Apps for Enterprise/Business = desktop-only SKU (no Exchange, no Teams).
        # If user only uses Web/Mobile, the desktop app investment is wasted.
        $standaloneAppSkus = @("OFFICESUBSCRIPTION","O365_BUSINESS","SMB_BUSINESS")
        $hasStandaloneApps = @($userSkuList | Where-Object { $_ -in $standaloneAppSkus }).Count -gt 0
        if ($hasStandaloneApps -and -not $usesDesktop -and ($usesWeb -or $usesMobile) -and $app) {
            $appSku  = ($userSkuList | Where-Object { $_ -in $standaloneAppSkus } | Select-Object -First 1)
            $appCost = Get-SkuMonthlyPrice $appSku
            $appAnnual = [math]::Round($appCost * 12, 2)
            $appName = Resolve-SkuFriendlyName $appSku
            $standaloneTarget = if ($tenantFamily -eq 'Enterprise') { "F3 or O365 E1" } else { "Business Basic or F3" }
            $recommendations.Add("STANDALONE APPS REVIEW — $appName (€$($appCost.ToString('N2'))/mo) assigned but user only uses web/mobile versions (no desktop activations). Consider downgrading to $standaloneTarget if no desktop dependency exists. Annual cost: €$($appAnnual.ToString('N2'))")
        }

        # ── A La Carte Waste — Kiosk + Standalone Desktop Apps Clash ──
        # Exchange Kiosk (€1/mo, 2 GB mailbox) + M365 Apps for Enterprise (€13.90/mo) = €14.90/mo.
        # M365 Business Standard (€12.50/mo) gives 50 GB mailbox + 1 TB OneDrive + Teams desktop.
        # Consolidating saves money AND massively upgrades the user experience.
        $hasKiosk = @($userSkuList | Where-Object { $_ -eq "EXCHANGEDESKLESS" }).Count -gt 0
        if ($hasKiosk -and $hasStandaloneApps -and -not $isSharedMailbox -and -not $isRoomOrEquipment -and $businessFamilyTotalConsumed -lt 250 -and $tenantFamily -ne 'Enterprise') {
            $kioskPrice = Get-SkuMonthlyPrice "EXCHANGEDESKLESS"
            $appSkuALC  = ($userSkuList | Where-Object { $_ -in $standaloneAppSkus } | Select-Object -First 1)
            $appCostALC = Get-SkuMonthlyPrice $appSkuALC
            $combinedCost = $kioskPrice + $appCostALC
            $bizStdPrice  = Get-SkuMonthlyPrice "M365_BUSINESS_STANDARD"
            if ($combinedCost -gt $bizStdPrice) {
                $savings = [math]::Round($combinedCost - $bizStdPrice, 2)
                $annualSavings = [math]::Round($savings * 12, 2)
                $appNameALC = Resolve-SkuFriendlyName $appSkuALC
                $businessFamilyTotalConsumed++
                $e1ToBasicEligible    = ($standardpackConsumed -gt 0 -and ($businessFamilyTotalConsumed + $standardpackConsumed) -le 250)
                $e3ToBpEligible       = ($speE3Consumed -gt 0 -and ($businessFamilyTotalConsumed + $speE3Consumed) -le 250)
                $appsEntToBizEligible = ($appsEntConsumed -gt 0 -and ($businessFamilyTotalConsumed + $appsEntConsumed) -le 250)
                $recommendations.Add("BUNDLE OPPORTUNITY — Exchange Kiosk (€$($kioskPrice.ToString('N2'))/mo) + $appNameALC (€$($appCostALC.ToString('N2'))/mo) = €$($combinedCost.ToString('N2'))/mo. Consider consolidating into M365 Business Standard (€$($bizStdPrice.ToString('N2'))/mo) to save €$($savings.ToString('N2'))/mo (€$($annualSavings.ToString('N2'))/yr) AND upgrade mailbox from 2 GB to 50 GB + add 1 TB OneDrive. Note: Business SKUs limited to 300-seat tenants.")
            }
        }

        # ── A La Carte Waste — Exchange Plan 1 + Standalone Desktop Apps ("Frankenstein Suite") ──
        # Exchange Plan 1 (€4/mo, 50 GB mailbox) + M365 Apps for Business (€11.50/mo) = €15.50/mo.
        # M365 Business Standard (€12.50/mo) includes Exchange + desktop apps + Teams + OneDrive.
        # Consolidating saves money AND adds Teams/OneDrive that the user doesn't currently have.
        $hasExoPlan1Standalone = @($userSkuList | Where-Object { $_ -eq "EXCHANGESTANDARD" }).Count -gt 0
        if ($hasExoPlan1Standalone -and $hasStandaloneApps -and -not $hasKiosk -and -not $isSharedMailbox -and -not $isRoomOrEquipment -and $businessFamilyTotalConsumed -lt 250 -and $tenantFamily -ne 'Enterprise') {
            $exoP1Price    = Get-SkuMonthlyPrice "EXCHANGESTANDARD"
            $appSkuFS      = ($userSkuList | Where-Object { $_ -in $standaloneAppSkus } | Select-Object -First 1)
            $appCostFS     = Get-SkuMonthlyPrice $appSkuFS
            $combinedFS    = $exoP1Price + $appCostFS
            $bizStdPriceFS = Get-SkuMonthlyPrice "M365_BUSINESS_STANDARD"
            if ($combinedFS -gt $bizStdPriceFS) {
                $savingsFS    = [math]::Round($combinedFS - $bizStdPriceFS, 2)
                $annSavingsFS = [math]::Round($savingsFS * 12, 2)
                $appNameFS    = Resolve-SkuFriendlyName $appSkuFS
                $businessFamilyTotalConsumed++
                $e1ToBasicEligible    = ($standardpackConsumed -gt 0 -and ($businessFamilyTotalConsumed + $standardpackConsumed) -le 250)
                $e3ToBpEligible       = ($speE3Consumed -gt 0 -and ($businessFamilyTotalConsumed + $speE3Consumed) -le 250)
                $appsEntToBizEligible = ($appsEntConsumed -gt 0 -and ($businessFamilyTotalConsumed + $appsEntConsumed) -le 250)
                $recommendations.Add("BUNDLE OPPORTUNITY — Exchange Plan 1 (€$($exoP1Price.ToString('N2'))/mo) + $appNameFS (€$($appCostFS.ToString('N2'))/mo) = €$($combinedFS.ToString('N2'))/mo. Consider consolidating into M365 Business Standard (€$($bizStdPriceFS.ToString('N2'))/mo) to save €$($savingsFS.ToString('N2'))/mo (€$($annSavingsFS.ToString('N2'))/yr) AND gain Teams + 1 TB OneDrive included. Note: Business SKUs limited to 300-seat tenants.")
            }
        }

        # ── Bundle Inefficiency — Business Basic + Apps for Business Frankenstein ──
        # Business Basic (email+Teams) + Apps for Business (desktop apps) = two licenses
        # M365 Business Standard natively includes BOTH at a lower combined cost.
        $bizBasicSkus = @("O365_BUSINESS_ESSENTIALS","SMB_BUSINESS_ESSENTIALS","M365_BUSINESS_BASIC")
        $bizAppsSkus  = @("O365_BUSINESS","SMB_BUSINESS")
        $hasBizBasic  = @($userSkuList | Where-Object { $_ -in $bizBasicSkus }).Count -gt 0
        $hasBizApps   = @($userSkuList | Where-Object { $_ -in $bizAppsSkus }).Count -gt 0
        if ($hasBizBasic -and $hasBizApps) {
            $basicSku   = ($userSkuList | Where-Object { $_ -in $bizBasicSkus } | Select-Object -First 1)
            $appsSku    = ($userSkuList | Where-Object { $_ -in $bizAppsSkus }  | Select-Object -First 1)
            $basicPrice = Get-SkuMonthlyPrice $basicSku
            $appsPrice  = Get-SkuMonthlyPrice $appsSku
            $combinedBI = $basicPrice + $appsPrice
            $bizStdPrBI = Get-SkuMonthlyPrice "O365_BUSINESS_PREMIUM"
            if ($combinedBI -gt $bizStdPrBI) {
                $savingsBI    = [math]::Round($combinedBI - $bizStdPrBI, 2)
                $annSavingsBI = [math]::Round($savingsBI * 12, 2)
                $recommendations.Add("BUNDLE OPPORTUNITY — $(Resolve-SkuFriendlyName $basicSku) (€$($basicPrice.ToString('N2'))/mo) + $(Resolve-SkuFriendlyName $appsSku) (€$($appsPrice.ToString('N2'))/mo) = €$($combinedBI.ToString('N2'))/mo. Consider consolidating into M365 Business Standard (€$($bizStdPrBI.ToString('N2'))/mo) which natively includes both. Potential savings: €$($savingsBI.ToString('N2'))/mo (€$($annSavingsBI.ToString('N2'))/yr). Note: Business SKUs limited to 300-seat tenants.")
            }
        }

        # Dormant sign-in — suppress for Room/Equipment mailboxes (resource accounts don't interactively sign in)
        # and for disabled/shared mailbox accounts (activity is from delegates, forwarding, or background sync —
        # STALE SIGN-IN's "do not remove" advice contradicts the primary DISABLED/SHARED MAILBOX rec).
        if ($isDormant -and -not $isRoomOrEquipment -and $isAccountEnabled -and -not $isSharedMailbox) {
            if ($hasAnyActivity) {
                # Sign-in is stale but workload activity detected (cached tokens, mobile apps,
                # background sync).  Do NOT suggest license removal — the user is active.
                $recommendations.Add("STALE SIGN-IN — no interactive sign-in for $daysSinceSignIn days, however M365 workload activity (Exchange, Teams, OneDrive, or SharePoint) was detected in the $ReportPeriod report period. The account is likely active via cached credentials or mobile apps. Review sign-in hygiene but do not remove the license.")
            } else {
                $recommendations.Add("DORMANT — no interactive sign-in for $daysSinceSignIn days (flagged at $InactiveSignInDays+ days of inactivity). Review whether the license can be removed or reassigned. Annual cost: €$($userAnnualCost.ToString('N2'))")
            }
            # Check non-interactive sign-in to distinguish automation accounts from truly abandoned users.
            # If interactive sign-in is dormant but non-interactive is recent, this is likely
            # an automation/service account (scripts, scheduled tasks, app registrations)
            # that logs in programmatically — NOT a truly abandoned account.
            # Suppress when STALE SIGN-IN fired ($hasAnyActivity) — workload activity means
            # the account is genuinely active, not an automation account.
            if ($hasRecentNonInteractive -and -not $hasAnyActivity -and -not $isGuest) {
                if ($isAdmin) {
                    $recommendations.Add("AUTOMATION ACCOUNT — admin account$adminRolesDisplay has no interactive sign-in for $daysSinceSignIn days but has recent non-interactive sign-in ($lastNonInteractiveSignIn, $daysSinceNonInteractive day$(if ([int]$daysSinceNonInteractive -ne 1) {'s'}) ago). This could indicate a service/automation account running scripts or scheduled tasks, or a device with apps refreshing tokens in the background. Review the account's purpose, ensure Conditional Access covers non-interactive flows, and, if this is a service account, consider converting to a dedicated Workload Identity (no user license needed).")
                } else {
                    $recommendations.Add("AUTOMATION ACCOUNT — user has no interactive sign-in for $daysSinceSignIn days but has recent non-interactive sign-in ($lastNonInteractiveSignIn, $daysSinceNonInteractive day$(if ([int]$daysSinceNonInteractive -ne 1) {'s'}) ago). This could indicate a service/automation account, or a device with apps refreshing tokens in the background. Review the account's purpose and, if this is a service account, consider converting to a dedicated Workload Identity (no user license needed). Annual cost: €$($userAnnualCost.ToString('N2'))")
                }
            } elseif ($isServiceAccountByPattern) {
                # Dormant account matching service/sync UPN pattern or Directory Sync role — flag even without non-interactive sign-in
                $patternSignal = if ($adminRolesStr -match 'Directory Synchronization Accounts') { "Directory Synchronization Accounts role" } else { "service account UPN pattern" }
                $recommendations.Add("AUTOMATION ACCOUNT — $patternSignal detected. No interactive sign-in for $daysSinceSignIn days. This is likely an infrastructure/sync service account. Review the account's purpose and consider converting to a dedicated Workload Identity (no user license needed). Annual cost: €$($userAnnualCost.ToString('N2'))")
            } elseif ($isAdmin) {
                $recommendations.Add("DORMANT ADMIN REVIEW — admin account$adminRolesDisplay has not signed in for $daysSinceSignIn days (interactive or non-interactive). This represents both unused license spend (€$($userAnnualCost.ToString('N2'))/yr) and an opportunity to tighten access controls. Consider removing the admin role or reassigning the license. If confirmed unused, consider disabling the account.")
            }
        }

        # Never signed in — licensed user with no sign-in record at all
        # Guard: only emit when sign-in data was actually loaded; otherwise every user looks "never signed in"
        $isNeverSignedIn = $false
        if ($signInDataLoaded -and -not $isDormant -and $lastSignIn -eq "" -and $isAccountEnabled -and -not $isSharedMailbox -and -not $isRoomOrEquipment) {
            $isNeverSignedIn = $true
            if ($isServiceAccountByPattern) {
                # Service/sync account that never signed in interactively — flag as automation, not mystery user
                $patternSignal = if ($adminRolesStr -match 'Directory Synchronization Accounts') { "Directory Synchronization Accounts role" } else { "service account UPN pattern" }
                $recommendations.Add("AUTOMATION ACCOUNT — $patternSignal detected. No interactive sign-in on record. This is an infrastructure/sync service account that operates non-interactively. Review the account's purpose and consider converting to a dedicated Workload Identity (no user license needed). Annual cost: €$($userAnnualCost.ToString('N2'))")
            } elseif ($emailSend -gt 0 -or $hasAnyActivity) {
                # User has sent emails or has workload activity — they likely signed in via OWA/mobile
                # but Entra sign-in logs may have rolled over. Soften the recommendation.
                $recommendations.Add("NEVER SIGNED IN — no interactive sign-in on record, however Exchange or M365 workload activity was detected in $ReportPeriod. The sign-in record may have expired from Entra logs. Review account usage.")
            } else {
                $functionalNote = if ($emailReceive -gt 0 -and $emailSend -eq 0) {
                    " This account receives email but has no interactive sign-in — it may be a functional or shared-purpose account."
                } else { "" }
                if ($userAnnualCost -gt 0) {
                    $recommendations.Add("NEVER SIGNED IN — no interactive sign-in on record. Review whether the license is still needed before next renewal.$functionalNote Annual cost: €$($userAnnualCost.ToString('N2'))")
                } else {
                    $recommendations.Add("NEVER SIGNED IN — no interactive sign-in on record.$functionalNote No financial impact — cleanup candidate.")
                }
            }
        }

        # Forwarding-only mailbox waste — mailbox exists only to forward mail elsewhere
        # A dormant/low-activity user with auto-forwarding configured and no mailbox delivery
        # can be replaced by a free Mail Contact, transport rule, or shared mailbox.
        if ($forwardingTarget -ne "" -and $userAnnualCost -gt 0 -and -not $isSharedMailbox -and -not $isRoomOrEquipment) {
            $fwdMode = if ($deliverAndForward) { "copy" } else { "forward-only" }
            if ($isDormant -or ($lastSignIn -eq "" -and $emailTotal -eq 0)) {
                # Dormant or never-signed-in with no email activity = pure forwarding waste
                $recommendations.Add("FORWARDING MAILBOX REVIEW — mailbox auto-forwards all mail to $forwardingTarget ($fwdMode) with no interactive sign-in $( if ($daysSinceSignIn) { "for $daysSinceSignIn days" } else { 'on record' }). This mailbox appears to exist only to forward email and likely does not require a paid license. Consider converting to a free Mail Contact, shared mailbox, or Exchange transport rule. Annual cost: €$($userAnnualCost.ToString('N2'))")
            } elseif ($emailIntensity -eq 'Low' -and -not $deliverAndForward) {
                # Active user but forward-only (no local delivery) with low exchange = likely unnecessary license
                $recommendations.Add("FORWARDING MAILBOX REVIEW — mailbox is configured to forward all mail to $forwardingTarget (forward-only, no local delivery) with low exchange activity ($emailTotal emails). Consider converting to a free Mail Contact or shared mailbox if user does not actively use this mailbox. Annual cost: €$($userAnnualCost.ToString('N2'))")
            }
        }

        # Desktop vs Web apps
        # Logic Flaw #1: only flag if user holds a SKU that INCLUDES desktop apps — otherwise it's noise
        # (e.g., telling a Business Basic user to "consider web-only license" is redundant)
        $hasDesktopAppEntitlement = $false
        foreach ($sku in $userSkuList) {
            $caps = Get-PlanCapabilities $sku
            if ($caps -and $caps.ContainsKey('DesktopApps') -and $caps['DesktopApps'] -eq $true) {
                $hasDesktopAppEntitlement = $true; break
            }
        }
        # ── Identity-only license detection (moved here so generic notes can use it) ──
        # Admin accounts with only identity SKUs (Entra P1/P2, ITP, EMS E5) don't use M365 workloads —
        # generic "no app activity" notes are irrelevant noise for these accounts.
        $paidNonIdentitySkus = @($userSkuList | Where-Object { -not $identityOnlySkus.Contains($_) -and -not $freeSkuSet.Contains($_) })
        $isIdentityOnlyLicense = ($paidNonIdentitySkus.Count -eq 0 -and ($isAdmin -or $pimEligibleRoles -or $pimActiveRoles))

        # ── Usage observations — only emit when they support a specific actionable downgrade ──
        # For suite licenses (E3/E5/Business Premium), per-workload usage notes are not independently
        # actionable because you cannot remove a single workload from a suite. Only emit these when
        # the user has a standalone workload license where the observation leads to a concrete action
        # (e.g., standalone Exchange Plan 2 → Plan 1, or desktop license → web-only license).
        $isSuiteLicense = @($userSkuList | Where-Object { $suiteIncludes.ContainsKey($_) }).Count -gt 0

        if (-not $isSuiteLicense) {
            if ($noDesktopApps -and $hasDesktopAppEntitlement) {
                $noDesktopTarget = if ($tenantFamily -eq 'Enterprise') { "F3 or O365 E1" } else { "M365 Business Basic or F3" }
                $recommendations.Add("No desktop apps — uses web/mobile only ($($webApps -join ', ')) — consider web-only license (e.g. $noDesktopTarget).")
            }
            if ($usesMobileOnly -and $hasDesktopAppEntitlement) {
                $recommendations.Add("Uses mobile apps only — consider F1/F3 frontline license.")
            }
            if (-not $usesDesktop -and -not $usesWeb -and -not $usesMobile -and $au -and -not $isRoomOrEquipment -and $isAccountEnabled -and $hasAnyActivity -and -not $isIdentityOnlyLicense) {
                $recommendations.Add("No M365 desktop, web, or mobile app activity detected in $ReportPeriod. Review whether the license is still needed.")
            }

            # Email client pattern (using entitlement-derived flag, not report flag)
            if ($noOutlookDesktop -and $au -and $hasExchangeEntitlement) {
                $recommendations.Add("No Outlook desktop client detected, email access is via OWA or mobile only. A desktop-tier Exchange license may not be required.")
            }

            # Teams web-only
            if ($teamsNoDesktop) {
                $recommendations.Add("Teams used without desktop client. This user may be a candidate for a Frontline license which supports web and mobile Teams.")
            }

            # Service-specific
            if ($emailIntensity -eq "Low" -and $em -and $hasExchangeEntitlement -and -not (-not $isAccountEnabled -and $isSharedMailbox)) {
                if ($null -ne $mbSizeMB -and $mbSizeMB -ge 100) {
                    # Significant stored data — mailbox is in use, just low recent activity
                    $recommendations.Add("Low Exchange activity ($emailSend sent, $emailReceive received in $ReportPeriod) but mailbox contains $([math]::Round($mbSizeMB / 1024, 1)) GB of data. The mailbox is actively used for storage. Review whether a lower-tier Exchange plan would be sufficient.")
                } else {
                    $mbNote = if ($null -ne $mbSizeMB -and $mbSizeMB -gt 0) { " (mailbox: ${mbSizeMB} MB)" } else { "" }
                    $recommendations.Add("Low Exchange activity ($emailSend sent, $emailReceive received in $ReportPeriod)$mbNote. Review whether the current Exchange plan is still needed.")
                }
            }
            if ($teamsIntensity -eq "Low" -and $tm -and $hasTeamsEntitlement) {
                $recommendations.Add("Low Teams activity ($teamsTotal actions in $ReportPeriod). Review whether a full Teams license is needed or consider switching to a plan without Teams.")
            }
            if ($odIntensity -eq "Low" -and $od -and $hasOneDriveEntitlement) {
                if ($null -ne $odStorageMB -and $odStorageMB -gt 100) {
                    $recommendations.Add("Low OneDrive activity ($odTotal actions in $ReportPeriod) but $([math]::Round($odStorageMB / 1024, 1)) GB of data stored. NOTE: Migrate or back up data before making any license changes.")
                } else {
                    $recommendations.Add("Low OneDrive activity ($odTotal actions in $ReportPeriod). Review whether OneDrive storage is still needed.")
                }
            }
        }

        # Check for no activity at all ($hasAnyActivity computed earlier, before $tier1Removal)
        # Suppress when DORMANT, DISABLED, or SHARED MAILBOX already flagged — those are higher-priority
        # actionable recommendations that already cover the "remove license" action.
        # Also suppress for identity-only licenses (Entra P1/P2 + free SKUs) with PIM/admin roles —
        # these accounts have no M365 workloads to measure, the license is justified by role, not app usage.
        # Note: $paidNonIdentitySkus and $isIdentityOnlyLicense are computed earlier (before usage observations block).
        $alreadyFlaggedForRemoval = ($isDormant -or (-not $isAccountEnabled) -or $isSharedMailbox -or $isNeverSignedIn)
        if (-not $hasAnyActivity -and $au -and $userAnnualCost -gt 0 -and -not $alreadyFlaggedForRemoval -and -not $isIdentityOnlyLicense) {
            $storageWarning = ""
            if (($null -ne $mbSizeMB -and $mbSizeMB -gt 100) -or ($null -ne $odStorageMB -and $odStorageMB -gt 100)) {
                $storageParts = @()
                if ($null -ne $mbSizeMB -and $mbSizeMB -gt 0) { $storageParts += "mailbox: $([math]::Round($mbSizeMB / 1024, 1)) GB" }
                if ($null -ne $odStorageMB -and $odStorageMB -gt 0) { $storageParts += "OneDrive: $([math]::Round($odStorageMB / 1024, 1)) GB" }
                if ($storageParts.Count -gt 0) {
                    $storageWarning = " Note: user has data ($($storageParts -join ', ')) — verify data retention before removing the license."
                }
            }
            # Logic Flaw #3: Power Platform SKUs run server-side — no M365 app activity telemetry.
            # Append REVIEW caveat so confidence drops from Medium to Review for these users.
            $powerPlatSkus = @("POWERAPPS_PER_USER","POWER_AUTOMATE_PER_USER","POWER_AUTOMATE_PREMIUM","POWERAPPS_PER_APP")
            $userPowerPlatSkus = @($userSkuList | Where-Object { $_ -in $powerPlatSkus })
            $powerPlatNote = ""
            if ($userPowerPlatSkus.Count -gt 0) {
                $ppNames = ($userPowerPlatSkus | ForEach-Object { Resolve-SkuFriendlyName $_ }) -join "; "
                $powerPlatNote = " POWER PLATFORM REVIEW: User holds $ppNames — Power Automate flows and Power Apps run server-side without generating M365 app activity. Review usage via Power Platform admin center before removing."
            }
            # Copilot Chat (copilot.microsoft.com) does not generate standard app telemetry.
            # A Copilot holder with 0 standard activity may be using web Copilot Chat daily.
            $copilotNote = ""
            if ($hasCopilotProd -or $hasCopilotBusiness) {
                $copilotNote = " COPILOT REVIEW: User holds a Copilot license — web-based Copilot Chat activity is NOT captured in standard app usage reports. Verify via the Copilot usage dashboard before removing."
            }
            # Compliance conflict: only warn when removing the license would LOSE an entitlement
            # that currently satisfies a compliance requirement (e.g., E3/E5 includes P1+MDO).
            # If the user already has a LICENSING CHECK (= gap already exists), removing the license
            # doesn't create a new gap — it already exists. Only warn when the license provides coverage.
            $complianceNote = ""
            if (($generalCA -ne '' -and $hasEntraP1) -or ($mdoCoverageNonBuiltIn -and $hasDefenderForO365)) {
                $loseParts = @()
                if ($generalCA -ne '' -and $hasEntraP1) {
                    $caEntLabel = if ($hasEntraP2) { "P2" } else { "P1" }
                    $loseParts += "Conditional Access (current license provides Entra ID $caEntLabel)"
                }
                if ($mdoCoverageNonBuiltIn -and $hasDefenderForO365) { $loseParts += "Defender for Office 365 (current license provides MDO P1)" }
                $complianceNote = " Note: removing this license would remove $($loseParts -join ' and ') coverage. Exclude the user from these policies first, or retain the license."
            }
            # Cloud PC or recent sign-in: user IS active, just not in M365 workloads.
            # Rephrase to right-sizing suggestion instead of "remove license".
            $hasCloudPc       = $lkpCloudPcType.ContainsKey($upn) -and $lkpCloudPcType[$upn] -ne ''
            $hasRecentSignIn  = ($null -ne $daysSinceSignIn -and $daysSinceSignIn -ne '' -and [int]$daysSinceSignIn -lt 30)
            if (($hasCloudPc -or $hasRecentSignIn) -and $userAnnualCost -gt 0) {
                $activityDetail = @()
                if ($hasCloudPc) {
                    $cpcLastNote = if ($lkpCloudPcDaysSinceSignIn.ContainsKey($upn) -and $lkpCloudPcDaysSinceSignIn[$upn] -ne '') { $cpcD = $lkpCloudPcDaysSinceSignIn[$upn]; "last Cloud PC connection $cpcD day$(if ([int]$cpcD -ne 1) {'s'}) ago" } else { "Cloud PC provisioned" }
                    $activityDetail += $cpcLastNote
                }
                if ($hasRecentSignIn) { $activityDetail += "last Entra sign-in $daysSinceSignIn day$(if ([int]$daysSinceSignIn -ne 1) {'s'}) ago" }
                $activityStr = $activityDetail -join '; '
                # CPC Enterprise users: recommend M365 E3 instead of E1 — E1 lacks the three CPC prerequisites
                # (Intune, Entra P1, Windows Enterprise E3). E3 satisfies all three AND includes MDO P1.
                if ($hasCpcEnterprise) {
                    $noActE3Price = Get-SkuMonthlyPrice "SPE_E3"
                    $noActNetSave   = [math]::Round(($userMonthlyCost - $noActE3Price) * 12, 2)
                    $noActMonthlySave = [math]::Round($userMonthlyCost - $noActE3Price, 2)
                    if ($noActNetSave -gt 0) {
                        $userNoActRightsizeSave = $noActNetSave
                        $recommendations.Add("NO ACTIVITY — No M365 workload activity (Exchange, Teams, OneDrive, SharePoint) detected in $ReportPeriod, but user has recent activity ($activityStr). The current license (€$($userMonthlyCost.ToString('N2'))/mo) may be oversized for this usage pattern. Consider replacing with Microsoft 365 E3 (€$($noActE3Price.ToString('N2'))/mo) — E3 is the minimum suite that satisfies Windows 365 Enterprise prerequisites (Intune, Entra P1, Windows Enterprise E3) and includes MDO P1.$storageWarning$powerPlatNote$copilotNote Potential savings: €$($noActMonthlySave.ToString('N2'))/mo (€$($noActNetSave.ToString('N2'))/yr).")
                    } else {
                        $recommendations.Add("NO ACTIVITY — No M365 workload activity detected in $ReportPeriod, but user has recent activity ($activityStr). The current license includes Windows 365 Enterprise prerequisites — review whether the usage pattern justifies the current suite.$storageWarning$powerPlatNote$copilotNote Annual cost: €$($userAnnualCost.ToString('N2'))")
                    }
                } else {
                # Standard path: suggest E1 as right-sized alternative with compliance add-on costs
                $noActE1Price = Get-SkuMonthlyPrice "STANDARDPACK"
                $noActCompAddon = [decimal]0
                $noActCompNote  = ""
                $noActNeedsP1   = ($generalCA -ne '')
                $noActNeedsMdo  = $mdoCoverageNonBuiltIn
                if ($noActNeedsP1)  { $noActCompAddon += Get-SkuMonthlyPrice "AAD_PREMIUM" }
                if ($noActNeedsMdo) { $noActCompAddon += Get-SkuMonthlyPrice "ATP_ENTERPRISE" }
                if ($noActCompAddon -gt 0) {
                    $caCount = if ($noActNeedsP1) { @($generalCA -split ';').Count } else { 0 }
                    $mdoCount = if ($noActNeedsMdo -and $mdoPolicyCoverage) { @($mdoPolicyCoverage -split ';').Count } else { 0 }
                    $noActCompNote = " Note: user is covered by$(if ($noActNeedsP1) { " $caCount Conditional Access" })$(if ($noActNeedsP1 -and $noActNeedsMdo) { ' and' })$(if ($noActNeedsMdo) { " $mdoCount MDO" }) $(if (($caCount + $mdoCount) -eq 1) { 'policy' } else { 'policies' }) — $(if ($noActNeedsP1) { "Entra ID P1 (€$((Get-SkuMonthlyPrice 'AAD_PREMIUM').ToString('N2'))/mo)" })$(if ($noActNeedsP1 -and $noActNeedsMdo) { ' and ' })$(if ($noActNeedsMdo) { "MDO P1 (€$((Get-SkuMonthlyPrice 'ATP_ENTERPRISE').ToString('N2'))/mo)" }) add-ons are required for compliance."
                }
                $noActNetSave   = [math]::Round(($userMonthlyCost - $noActE1Price - $noActCompAddon) * 12, 2)
                $noActMonthlySave = [math]::Round($userMonthlyCost - $noActE1Price - $noActCompAddon, 2)
                if ($noActNetSave -gt 0) {
                    $userNoActRightsizeSave = $noActNetSave
                    $recommendations.Add("NO ACTIVITY — No M365 workload activity (Exchange, Teams, OneDrive, SharePoint) detected in $ReportPeriod, but user has recent activity ($activityStr). The current license (€$($userMonthlyCost.ToString('N2'))/mo) may be oversized for this usage pattern. Consider replacing with Office 365 E1 (€$($noActE1Price.ToString('N2'))/mo) for basic web/mobile access.$noActCompNote$storageWarning$powerPlatNote$copilotNote Potential savings: €$($noActMonthlySave.ToString('N2'))/mo (€$($noActNetSave.ToString('N2'))/yr).")
                } else {
                    $recommendations.Add("NO ACTIVITY — No M365 workload activity detected in $ReportPeriod, but user has recent activity ($activityStr). Review whether the current license is still needed.$storageWarning$powerPlatNote$copilotNote$complianceNote Annual cost: €$($userAnnualCost.ToString('N2'))")
                }
                }
            } else {
                $recommendations.Add("NO ACTIVITY detected in $ReportPeriod — review whether the license can be removed or reassigned.$storageWarning$powerPlatNote$copilotNote$complianceNote Annual cost: €$($userAnnualCost.ToString('N2'))")
            }
            # Cold Storage escalation: 0 activity + significant data = paying premium to store data.
            # Threshold: mailbox > 10 GB or OneDrive > 50 GB — these are high-cost archival candidates.
            if (($null -ne $mbSizeMB -and $mbSizeMB -gt 10240) -or ($null -ne $odStorageMB -and $odStorageMB -gt 51200)) {
                $mbGB = if ($null -ne $mbSizeMB) { [math]::Round($mbSizeMB / 1024, 1) } else { 0 }
                $odGB = if ($null -ne $odStorageMB) { [math]::Round($odStorageMB / 1024, 1) } else { 0 }
                $coldDetails = @()
                if ($null -ne $mbSizeMB -and $mbSizeMB -gt 10240) { $coldDetails += "Mailbox: ${mbGB} GB" }
                if ($null -ne $odStorageMB -and $odStorageMB -gt 51200) { $coldDetails += "OneDrive: ${odGB} GB" }
                $recommendations.Add("EXPENSIVE COLD STORAGE — 0 activity but significant data ($($coldDetails -join '; ')). Current annual cost: €$($userAnnualCost.ToString('N2')). Consider converting the mailbox to Shared or Inactive, migrating OneDrive content to a SharePoint Archive site, and removing the license.")
            }
        }
    }

    # ── "Ghost in the Machine" — Passive OneDrive Sync detection ──
    # OneDrive sync client on an abandoned, powered-on device will silently sync SharePoint library
    # changes in the background. User shows 0 emails, 0 chats, 0 files viewed/shared — but high
    # sync count. This makes $hasAnyActivity = $true, letting ghost users slip through NO ACTIVITY.
    $isPassiveSyncOnly = ($odSynced -gt 0 -and $odViewed -eq 0 -and $odShared -eq 0 -and
                          $emailTotal -eq 0 -and $teamsTotal -eq 0 -and $spViewed -eq 0 -and
                          -not $usesDesktop -and -not $usesWeb -and -not $usesMobile)
    if ($isPassiveSyncOnly -and $userAnnualCost -gt 0 -and $au) {
        $recommendations.Add("BACKGROUND SYNC ONLY — user shows 0 interactive M365 activity (no emails, chats, or files viewed/shared), but OneDrive synced $odSynced file(s) in the background. This usually indicates an abandoned, powered-on device (e.g., laptop in a drawer or stale VDI session). Review device status and whether the account is actively used before considering license removal. Annual cost: €$($userAnnualCost.ToString('N2'))")
    }

    # ── Cloud PC utilization (beta API — only when CPC/W365 SKU assigned to this user) ──
    $userCpcSkus = @($userSkuList | ForEach-Object { $_ -replace '[\u200B\uFEFF]', '' } | Where-Object { $_ -match '^(CPC_|Windows_365_)' })

    # ── Cloud PC Enterprise prerequisite check ──
    # Windows 365 Enterprise requires: (1) Intune, (2) Entra ID P1, (3) Windows 10/11 Enterprise.
    # Qualifying bundles (M365 E3/E5/F3, Business Premium, A3/A5) include all three.
    $userEntCpcSkus = @($userCpcSkus | Where-Object { $_ -match '^(CPC_E|Windows_365_E)' })
    if ($userEntCpcSkus.Count -gt 0) {
        $hasIntuneReq  = $effectiveSkuSet.Contains("INTUNE_A")
        $hasEntraP1Req = ($effectiveSkuSet.Contains("AAD_PREMIUM") -or $effectiveSkuSet.Contains("AAD_PREMIUM_P2"))
        # Windows Enterprise is not tracked in suiteIncludes, so check for qualifying bundle SKUs
        # that implicitly include it, or standalone Windows Enterprise SKUs
        $hasWinEntReq  = ($effectiveSkuSet.Contains("WIN10_PRO_ENT_SUB") -or
                          $effectiveSkuSet.Contains("WIN10_VDA_E5") -or
                          $effectiveSkuSet.Contains("WIN10_VDA_E3") -or
                          @($userSkuList | Where-Object { $_ -match '^(SPE_E[35]|SPE_F1|SPB|M365EDU_A[35]|M365EDU_A5_STUUSEBNFT|Microsoft_365_Business_Premium)' }).Count -gt 0)
        $missingPrereqs = @()
        if (-not $hasIntuneReq)  { $missingPrereqs += "Microsoft Intune" }
        if (-not $hasEntraP1Req) { $missingPrereqs += "Entra ID P1" }
        if (-not $hasWinEntReq)  { $missingPrereqs += "Windows 10/11 Enterprise" }
        if ($missingPrereqs.Count -gt 0) {
            $cpcEntFriendly = ($userEntCpcSkus | ForEach-Object { Resolve-SkuFriendlyName $_ }) -join "; "
            $missingStr = $missingPrereqs -join ", "
            $recommendations.Add("LICENSING CHECK — Windows 365 Enterprise ($cpcEntFriendly) requires $missingStr as prerequisite license(s). Consider assigning a qualifying bundle (M365 E3, E5, F3, or Business Premium) that includes all prerequisites, or add the missing standalone license(s). Alternatively, review whether the Cloud PC assignment is still needed.")
        }
    }

    # Skip Cloud PC recommendations if user connected within the last 14 days (actively using)
    # Primary: Cloud PC connection data. Fallback: Entra sign-in ONLY when CPC-specific data is
    # missing (covers beta API failure). Without the -not ContainsKey guard, any M365-active user
    # would have CPC recs suppressed even if their Cloud PC sits idle.
    # Note: $daysSinceSignIn can be '' (empty string) — must check for both $null and '' to avoid
    # PowerShell coercing '' to 0 in numeric comparison ('' -lt 14 → 0 -lt 14 → $true).
    $cpcRecentConnection = ($lkpCloudPcDaysSinceSignIn.ContainsKey($upn) -and $lkpCloudPcDaysSinceSignIn[$upn] -lt 14) -or
                           (-not $lkpCloudPcDaysSinceSignIn.ContainsKey($upn) -and $signInDataLoaded -and $null -ne $daysSinceSignIn -and $daysSinceSignIn -ne '' -and $daysSinceSignIn -lt 14)

    if ($userCpcSkus.Count -gt 0 -and -not $cloudPcUsageLoaded -and -not $cpcRecentConnection) {
        # Cloud PC API failed — emit data gap so the user isn't silently skipped
        $cpcFriendlyGap = ($userCpcSkus | ForEach-Object { Resolve-SkuFriendlyName $_ }) -join "; "
        $recommendations.Add("DATA GAP — Cloud PC is provisioned ($cpcFriendlyGap) but the usage hours API did not return data for this tenant. Cloud PC utilization cannot be assessed automatically. Review usage in the Intune admin center.")
    }
    if ($userCpcSkus.Count -gt 0 -and $cloudPcUsageLoaded) {
        [decimal]$cpcMonthlyCost = 0
        foreach ($csku in $userCpcSkus) { $cpcMonthlyCost += Get-SkuMonthlyPrice $csku }
        $cpcAnnualCost = [math]::Round($cpcMonthlyCost * 12, 2)
        $cpcFriendly   = ($userCpcSkus | ForEach-Object { Resolve-SkuFriendlyName $_ }) -join "; "

        # Build activity context string: Cloud PC sign-in + Entra sign-in (clearly labelled)
        $cpcActivityParts = @()
        if ($lkpCloudPcDaysSinceSignIn.ContainsKey($upn)) {
            $cpcD2 = $lkpCloudPcDaysSinceSignIn[$upn]; $cpcActivityParts += "last Cloud PC connection $cpcD2 day$(if ([int]$cpcD2 -ne 1) {'s'}) ago"
        }
        if ($lastSignIn -ne '') {
            $cpcActivityParts += "last Entra sign-in $lastSignIn ($daysSinceSignIn day$(if ([int]$daysSinceSignIn -ne 1) {'s'}) ago)"
        } elseif ($signInDataLoaded) {
            $cpcActivityParts += "no Entra sign-in on record"
        }
        $cpcActivityStr = if ($cpcActivityParts.Count -gt 0) { " Activity: $($cpcActivityParts -join '; ')." } else { "" }

        # NeverSignedIn enrichment — stronger signal when Cloud PC was never used since provisioning
        $cpcNeverNote = if ($lkpCloudPcNeverSignedIn.ContainsKey($upn) -and $lkpCloudPcNeverSignedIn[$upn]) {
            " This Cloud PC has never been signed into since provisioning."
        } else { "" }

        if (-not $cpcRecentConnection) {
            if ($lkpCloudPcUsageHours.ContainsKey($upn)) {
                $cpcHours = $lkpCloudPcUsageHours[$upn]
                if ($cpcHours -eq 0) {
                    # Zero hours in usage report
                    if ($isDormant -or $lastSignIn -eq '') {
                        # Sign-in logs confirm inactivity — stronger recommendation
                        $recommendations.Add("DORMANT CLOUD PC — $cpcFriendly (€$($cpcMonthlyCost.ToString('N2'))/mo) has 0 connected hours in the last 90 days and no recent sign-in activity.$cpcNeverNote$cpcActivityStr Consider removing or reassigning the license. Annual cost: €$($cpcAnnualCost.ToString('N2'))")
                    } else {
                        # Zero CPC hours but user has recent sign-ins — might use CPC sporadically or via other means
                        $recommendations.Add("CLOUD PC REVIEW — $cpcFriendly (€$($cpcMonthlyCost.ToString('N2'))/mo) has 0 connected hours in the last 90 days, but user is active in other M365 services.$cpcActivityStr Review whether the Cloud PC is still needed. Annual cost: €$($cpcAnnualCost.ToString('N2'))")
                    }
                } elseif ($cpcHours -lt 10) {
                    # Less than 10 hours in 90 days — ~7 min/day average
                    $cpcHoursRound = [math]::Round($cpcHours, 1)
                    $recommendations.Add("CLOUD PC REVIEW — $cpcFriendly (€$($cpcMonthlyCost.ToString('N2'))/mo) has only $cpcHoursRound connected hours in the last 90 days.$cpcActivityStr Consider downsizing or reclaiming. Annual cost: €$($cpcAnnualCost.ToString('N2'))")
                }
            } else {
                # User has CPC SKU but does NOT appear in the Cloud PC usage report (usage hours API unavailable — only provisioned list)
                if ($isDormant -or $lastSignIn -eq '') {
                    $recommendations.Add("DORMANT CLOUD PC — $cpcFriendly (€$($cpcMonthlyCost.ToString('N2'))/mo) is provisioned but the Cloud PC usage hours API returned no connection data for this user, and there is no recent sign-in activity.$cpcNeverNote$cpcActivityStr Consider removing or reassigning the license. Annual cost: €$($cpcAnnualCost.ToString('N2'))")
                } else {
                    $recommendations.Add("CLOUD PC REVIEW — $cpcFriendly (€$($cpcMonthlyCost.ToString('N2'))/mo) is provisioned but the Cloud PC usage hours API returned no connection data for this user. User is active in other M365 services.$cpcActivityStr Review whether the Cloud PC is still needed. Annual cost: €$($cpcAnnualCost.ToString('N2'))")
                }
            }
        }
    }

    # ── Unknown SKU soft warning (data gap) ──
    if ($hasUnknownSku) {
        $unknownList = ($unknownSkus | Sort-Object) -join "; "
        $recommendations.Add("DATA GAP — SKU(s) not in reference data: $unknownList. Cost and right-sizing recommendations may be incomplete. Update M365SkuData.json to resolve.")
    }

    $recommendationText = if ($recommendations.Count -gt 0) { $recommendations -join " | " } else { "No Findings — active user with matching license profile." }

    # ── Recommendation category (for grouping / pivot tables) ──
    # NOTE: We match against the pipe-joined $recommendationText using (^|\| ) anchoring
    # so that patterns match only recommendation PREFIXES, never body text.
    # This prevents false categorisation (e.g. body text mentioning "shared mailbox"
    # accidentally matching the SHARED MAILBOX category pattern).
    #
    # ORDER-CRITICAL PAIRS — do NOT reorder without understanding these dependencies:
    #   1. INACTIVE HOLD WITH LICENSE  before  INACTIVE HOLD         (prefix substring)
    #   2. SHARED MAILBOX REVIEW       before  SHARED MAILBOX        (prefix substring)
    #   3. INACTIVE ADD-ON REVIEW      before  INACTIVE ADD-ON       (prefix substring)
    #   4. EXO PLAN 2 REVIEW           before  EXO PLAN 2            (prefix substring)
    #   5. COPILOT PREREQUISITE/RECLAIM/WATCHLIST/ACTIVE/STUDIO  before  bare COPILOT (catch-all)
    #   6. NEVER SIGNED IN  before  TEAMS UNBUNDLING / LICENSING CHECK (stronger signal)
    #   7. DORMANT CLOUD PC / STALE SIGN-IN / DORMANT ADMIN REVIEW  before  bare DORMANT (catch-all)
    #   7. FRONTLINE ADD-ON STACKING / BLOCKED / CANDIDATE / REVIEW  are safe relative to each other
    #      but FRONTLINE RESCUE (line 5810) is intentionally placed later — FRONTLINE CANDIDATE wins as primary
    $recCategory = if     ($recommendationText -match "(^|\| )INACTIVE HOLD WITH LICENSE") { "Inactive Hold With License" }
                   elseif ($recommendationText -match "(^|\| )INACTIVE HOLD")        { "Inactive Hold" }
                   elseif ($recommendationText -match "(^|\| )DISABLED SHARED MAILBOX") { "Disabled Account" }
                   elseif ($recommendationText -match "(^|\| )DISABLED ACCOUNT")    { "Disabled Account" }
                   elseif ($recommendationText -match "(^|\| )NEVER SIGNED IN")     { "Never Signed In" }
                   elseif ($recommendationText -match "(^|\| )SHARED MAILBOX REVIEW") { "Shared Mailbox Review" }
                   elseif ($recommendationText -match "(^|\| )SHARED MAILBOX")      { "Shared Mailbox" }
                   elseif ($recommendationText -match "(^|\| )OVERLAPPING LICENSE") { "Overlapping License" }
                   elseif ($recommendationText -match "(^|\| )DUPLICATE REVIEW")     { "Duplicate Review" }
                   elseif ($recommendationText -match "(^|\| )DUPLICATE COVERAGE")  { "Duplicate Coverage" }
                   elseif ($recommendationText -match "(^|\| )SUITE INVERSION")      { "Suite Inversion" }
                   elseif ($recommendationText -match "(^|\| )E5 CONSOLIDATION")    { "E5 Upgrade" }
                   elseif ($recommendationText -match "(^|\| )BUNDLE CONSOLIDATION") { "Bundle Consolidation" }
                   elseif ($recommendationText -match "(^|\| )ENTRA SUITE OVERLAP")  { "Entra Suite Overlap" }
                   elseif ($recommendationText -match "(^|\| )INTUNE SUITE OVERLAP")   { "Intune Suite Overlap" }
                   elseif ($recommendationText -match "(^|\| )INTUNE REVIEW")      { "Intune Review" }
                   elseif ($recommendationText -match "(^|\| )WINDOWS LICENSE REVIEW") { "Windows License Review" }
                   elseif ($recommendationText -match "(^|\| )REDUNDANT ARCHIVE")      { "Redundant Archive" }
                   elseif ($recommendationText -match "(^|\| )OVER-LICENSED ARCHIVE") { "Over-Licensed Archive" }
                   elseif ($recommendationText -match "(^|\| )TEAMS UNBUNDLING")     { "Teams Unbundling" }
                   elseif ($recommendationText -match "(^|\| )F3 TO F1 DOWNGRADE")  { "F3 to F1 Downgrade" }
                   elseif ($recommendationText -match "(^|\| )FRONTLINE ADD-ON STACKING") { "Frontline Add-On Stacking" }
                   elseif ($recommendationText -match "(^|\| )FRONTLINE BLOCKED")   { "Frontline Blocked" }
                   elseif ($recommendationText -match "(^|\| )FRONTLINE CANDIDATE") { "Frontline Candidate" }
                   elseif ($recommendationText -match "(^|\| )FRONTLINE REVIEW")    { "Frontline Review" }
                   elseif ($recommendationText -match "(^|\| )SEEDED VISIO OVERLAP") { "Seeded Visio Overlap" }
                   elseif ($recommendationText -match "(^|\| )PREMIUM ADD-ON REVIEW") { "Premium Add-On Review" }
                   elseif ($recommendationText -match "(^|\| )INACTIVE ADD-ON REVIEW")    { "Inactive Add-On Review" }
                   elseif ($recommendationText -match "(^|\| )INACTIVE ADD-ON")      { "Inactive Add-On" }
                   elseif ($recommendationText -match "(^|\| )CALLING PLAN REVIEW")       { "Calling Plan Review" }
                   elseif ($recommendationText -match "(^|\| )TEAMS PHONE RIGHT-SIZING") { "Teams Phone Right-Sizing" }
                   elseif ($recommendationText -match "(^|\| )TEAMS PHONE REVIEW") { "Teams Phone Review" }
                   elseif ($recommendationText -match "(^|\| )AI ADD-ON OVERLAP")     { "AI Add-On Overlap" }
                   elseif ($recommendationText -match "(^|\| )AI OVERLAP REVIEW")     { "AI Overlap Review" }
                   elseif ($recommendationText -match "(^|\| )COPILOT PREREQUISITE")  { "Copilot Prerequisite" }
                   elseif ($recommendationText -match "(^|\| )COPILOT RECLAIM")     { "Copilot Reclaim" }
                   elseif ($recommendationText -match "(^|\| )COPILOT WATCHLIST")   { "Copilot Watchlist" }
                   elseif ($recommendationText -match "(^|\| )COPILOT ACTIVE")      { "Copilot Active" }
                   elseif ($recommendationText -match "(^|\| )COPILOT STUDIO")      { "Copilot Studio" }
                   elseif ($recommendationText -match "(^|\| )COPILOT")             { "Copilot" }
                   elseif ($recommendationText -match "(^|\| )DORMANT CLOUD PC")   { "Dormant Cloud PC" }
                   elseif ($recommendationText -match "(^|\| )CLOUD PC REVIEW")    { "Cloud PC Review" }
                   elseif ($recommendationText -match "(^|\| )POWER BI PRO REVIEW") { "Power BI Pro Review" }
                   elseif ($recommendationText -match "(^|\| )MAILBOX STORAGE WARNING") { "Mailbox Storage Warning" }
                   elseif ($recommendationText -match "(^|\| )ONEDRIVE PLAN 2 REVIEW")    { "OneDrive Plan 2 Review" }
                   elseif ($recommendationText -match "(^|\| )ONEDRIVE STORAGE WARNING") { "OneDrive Storage Warning" }
                   elseif ($recommendationText -match "(^|\| )EXO PLAN 2 REVIEW")   { "EXO Plan 2 Review" }
                   elseif ($recommendationText -match "(^|\| )EXO PLAN 2")          { "EXO Plan 2 Downgrade" }
                   elseif ($recommendationText -match "(^|\| )RoomMailbox|(^|\| )EquipmentMailbox") { "Room/Equipment" }
                   elseif (-not $isLicensed -and $isSharedMailbox -and $recommendationText -match "(^|\| )LICENSING CHECK") { "Shared Mailbox" }
                   elseif ($recommendationText -match "(^|\| )LICENSING CHECK")     { "Licensing Compliance Gap" }
                   elseif ($recommendationText -match "(^|\| )LICENSING ERROR")    { "License Error" }
                   elseif ($recommendationText -match "(^|\| )TRIAL LICENSE")       { "Trial License" }
                   elseif ($recommendationText -match "(^|\| )LICENSE CAPACITY QUEUE") { "License Capacity" }
                   elseif ($recommendationText -match "(^|\| )CLOUD LICENSE SYNC")  { "Cloud License Error" }
                   elseif ($recommendationText -match "(^|\| )BUNDLE OPPORTUNITY")      { "Bundle Opportunity" }
                   elseif ($recommendationText -match "(^|\| )STANDALONE APPS REVIEW") { "Standalone Apps Review" }
                   elseif ($recommendationText -match "(^|\| )BUSINESS BASIC CANDIDATE") { "Business Downgrade" }
                   elseif ($recommendationText -match "(^|\| )BUSINESS BASIC REVIEW")    { "Business Review" }
                   elseif ($recommendationText -match "(^|\| )E1 DOWNGRADE CANDIDATE")   { "E1 to Business Basic" }
                   elseif ($recommendationText -match "(^|\| )O365 E3 TO E1")             { "O365 E3 to E1" }
                   elseif ($recommendationText -match "(^|\| )FRONTLINE RESCUE")           { "Frontline Rescue" }
                   elseif ($recommendationText -match "(^|\| )E3 TO BUSINESS PREMIUM")    { "E3 to Business Premium" }
                   elseif ($recommendationText -match "(^|\| )BUSINESS PREMIUM INVERSION") { "Business Premium Inversion" }
                   elseif ($recommendationText -match "(^|\| )BUSINESS PREMIUM SECURITY REVIEW") { "Business Premium Security Review" }
                   elseif ($recommendationText -match "(^|\| )E5 VOICE REVIEW")            { "E5 Voice Review" }
                   elseif ($recommendationText -match "(^|\| )APP ARBITRAGE")              { "App Arbitrage" }
                   elseif ($recommendationText -match "(^|\| )PBI PPU OVERLAP")      { "PBI PPU Overlap" }
                   elseif ($recommendationText -match "(^|\| )ENTRA P2 DOWNGRADE")        { "Entra P2 Downgrade" }
                   elseif ($recommendationText -match "(^|\| )EXCHANGE KIOSK CANDIDATE")  { "Exchange Kiosk Downgrade" }
                   elseif ($recommendationText -match "(^|\| )FORWARDING MAILBOX REVIEW") { "Forwarding Mailbox Review" }
                   elseif ($recommendationText -match "(^|\| )GUEST ACCOUNT REVIEW")  { "Guest Account Review" }
                   elseif ($recommendationText -match "(^|\| )GUEST USER")          { "Guest User" }
                   elseif ($recommendationText -match "(^|\| )NON-HUMAN ACCOUNT REVIEW") { "Non-Human Account Review" }
                   elseif ($recommendationText -match "(^|\| )FREE LICENSE OVERLAP") { "Free License Overlap" }
                   elseif ($recommendationText -match "(^|\| )EXTERNAL SHARING REVIEW")   { "External Sharing Review" }
                   elseif ($recommendationText -match "(^|\| )INACTIVE MAILBOX")     { "Inactive Mailbox" }
                   elseif ($recommendationText -match "(^|\| )STALE SIGN-IN")       { "Stale Sign-In" }
                   elseif ($recommendationText -match "(^|\| )DORMANT ADMIN REVIEW")  { "Dormant Admin Review" }
                   elseif ($recommendationText -match "(^|\| )ADMIN\s*\(") { "Admin Review" }
                   elseif ($recommendationText -match "(^|\| )AUTOMATION ACCOUNT")  { "Automation Account" }
                   elseif ($recommendationText -match "(^|\| )LEGACY SERVICE ACCOUNT") { "Legacy Service Account" }
                   elseif ($recommendationText -match "(^|\| )DORMANT")             { "Dormant" }
                   elseif ($recommendationText -match "(^|\| )EXPENSIVE COLD STORAGE") { "Expensive Cold Storage" }
                   elseif ($recommendationText -match "(^|\| )BACKGROUND SYNC ONLY") { "Background Sync Only" }
                   elseif ($recommendationText -match "(^|\| )NO ACTIVITY")         { "No Activity" }
                   elseif ($recommendationText -match "(^|\| )No desktop apps")     { "No Desktop" }
                   elseif ($recommendationText -match "(^|\| )Uses mobile apps only") { "Mobile Only" }
                   elseif ($recommendationText -match "(^|\| )DATA GAP")             { "Data Gap" }
                   elseif ($recommendationText -match "(^|\| )UNLICENSED WITH DATA") { "Unlicensed With Data" }
                   elseif (-not $isLicensed -and $isSharedMailbox)                  { "Shared Mailbox" }
                   elseif (-not $isLicensed -and $assignedSkus -eq "[UNLICENSED]")  { "Unlicensed" }
                   elseif ($recommendationText -match "(^|\| )SECURITY GAP")          { "Security Gap" }
                   elseif ($recommendationText -match "(^|\| )DEFENDER COVERAGE REVIEW") { "Defender Coverage Review" }
                   elseif ($recommendationText -match "(^|\| )COMPLIANCE COVERAGE REVIEW") { "Compliance Coverage Review" }
                   elseif ($recommendationText -eq "No Findings — active user with matching license profile.") { "No Findings" }
                   else { "Partial Optimization" }

    # ── Recommendation confidence (LOA v1.0 spec §6.1) ──
    # High   = deterministic, no missing data  (disabled account, overlapping, duplicate, dormant)
    # Medium = activity-based with complete data  (frontline candidate, business basic candidate, EXO downgrade, shelfware)
    # ── Recommendation Confidence ────────────────────────────────────────────────
    # High   = based on hard facts: account state, license data, policy membership,
    #          zero sign-in, mailbox type, zero product usage, compliance entitlements
    # Medium = based on usage patterns with edge-case potential: 90-day window,
    #          seasonal workers, delegates, platform inference
    # Review = missing data, mapping uncertainty, or explicit REVIEW/DATA GAP tag
    $recConfidence = if     ($recCategory -eq "No Findings")                               { "" }
                     elseif ($recCategory -eq "Unlicensed")                                { "" }
                     elseif ($recommendationText -cmatch "\bREVIEW\b")                     { "Review" }
                     elseif ($recommendationText -cmatch "\bDATA GAP\b")                   { "Review" }
                     elseif ($recCategory -in @(
                                # Account state (factual from Entra/EXO)
                                "Disabled Account","Dormant","Never Signed In","No Activity",
                                "Guest User","Automation Account","Legacy Service Account",
                                "Shared Mailbox","Room/Equipment",
                                # License data (factual from Graph)
                                "Overlapping License","Duplicate Coverage",
                                "License Error","Cloud License Error","License Capacity",
                                "Trial License","Frontline Blocked",
                                "Copilot Prerequisite","Copilot Studio",
                                "Free License Overlap","Unlicensed With Data",
                                "Suite Inversion","AI Add-On Overlap",
                                "Business Premium Inversion","Entra Suite Overlap","Intune Suite Overlap",
                                "Bundle Consolidation","Bundle Opportunity",
                                # Compliance (factual: policy exists + entitlement missing)
                                "Licensing Compliance Gap",
                                # Product usage (zero activity = factual signal)
                                "Inactive Add-On","Teams Unbundling",
                                "Inactive Mailbox","Inactive Hold With License","Inactive Hold",
                                "Background Sync Only","Expensive Cold Storage",
                                # EXO properties (factual from mailbox data)
                                "EXO Plan 2 Downgrade","Non-Human Account Review",
                                "Over-Licensed Archive","Redundant Archive",
                                "Frontline Rescue","Litigation Hold"))                     { "High" }
                     elseif ($recCategory -in @(
                                # Usage-pattern inference (90-day window, platform heuristics)
                                "Frontline Candidate","Business Downgrade",
                                "No Desktop","Mobile Only",
                                "E3 to Business Premium","O365 E3 to E1","E1 to Business Basic",
                                "F3 to F1 Downgrade","Frontline Add-On Stacking",
                                "E5 Upgrade",
                                # Cost comparison / manual validation needed
                                "Teams Phone Right-Sizing","App Arbitrage",
                                "PBI PPU Overlap","Calling Plan Review",
                                "OneDrive Plan 2 Review","Entra P2 Downgrade",
                                "Exchange Kiosk Downgrade","Intune Review",
                                "Seeded Visio Overlap","Windows License Review",
                                # Warnings (threshold-based, not definitive)
                                "Mailbox Storage Warning","OneDrive Storage Warning",
                                "Business Premium Security Review",
                                "Defender Coverage Review","Compliance Coverage Review"))   { "Medium" }
                     else                                                                  { "Medium" }

    # ── Post-hoc confidence upgrade: evidence-driven compliance findings ────────────────
    # LICENSING CHECK recs are based on detected policy membership + missing entitlements —
    # this is factual (policies exist, entitlement missing), not activity-inferred.
    if ($recConfidence -eq "Medium" -and $recommendationText -match "(^|\| )LICENSING CHECK") {
        $recConfidence = "High"
    }

    # ── Post-hoc confidence downgrade: activity-based recs with missing key data sources ──
    # If a recommendation depends on usage/activity signals and the underlying reports are missing,
    # the recommendation is unreliable — force confidence to Review regardless of category mapping.
    # Key activity sources: EmailActivity, TeamsActivity, OneDriveActivity, M365AppPlatform
    if ($recConfidence -in @("High","Medium") -and $missingDataSources.Count -gt 0) {
        $activityBasedCategories = @("No Activity","Frontline Candidate","Business Downgrade",
            "EXO Plan 2 Downgrade","Teams Unbundling","No Desktop","Mobile Only",
            "F3 to F1 Downgrade","Background Sync Only")
        $keyActivitySources = @("EmailActivity","TeamsActivity","OneDriveActivity","M365AppPlatform")
        if ($recCategory -in $activityBasedCategories) {
            $missingKeySources = @($missingDataSources | Where-Object { $_ -in $keyActivitySources })
            if ($missingKeySources.Count -gt 0) {
                $recConfidence = "Review"
            }
        }
    }

    # ── Disabled plans string for CSV output ──
    $disabledPlansStr = if ($disabledPlans -and $disabledPlans.Count -gt 0) {
        ($disabledPlans | Sort-Object) -join "; "
    } else { "" }

    # ── Security & Compliance coverage levels (Flaw 4) ──
    # Compute posture level per user based on merged capability flags.
    # Values: None / Basic / Advanced / E5-equivalent
    $securityCoverageLevel = if (-not $isLicensed) { "None" }
        elseif ($hasFullDefenderStack) { "E5-equivalent" }
        elseif ($userCaps.MdoP2 -and $userCaps.MdeP2OrBusiness) { "Advanced" }
        elseif ($hasAnyDefenderCap) { "Basic" }
        else { "None" }
    $complianceCoverageLevel = if (-not $isLicensed) { "None" }
        elseif ($hasFullPurviewStack) { "E5-equivalent" }
        elseif ($userCaps.DlpTeams -or $userCaps.AIPPlan2 -or $userCaps.eDiscoveryPremium) { "Advanced" }
        elseif ($hasAnyPurviewCap) { "Basic" }
        else { "None" }

    # ── Build merged row + stream to CSV ──
    $row = [PSCustomObject]@{
        'User Principal Name'    = $upn
        'Display Name'           = if ($au -and $au.'Display Name') { $au.'Display Name' } elseif ($userObj) { $userObj.DisplayName } else { "" }
        'Assigned Licenses'      = $assignedSkus
        'License Friendly Names' = $licenseFriendlyStr
        'License Assignment'     = $licAssignmentStr
        'Overlapping Licenses'   = $overlappingSkus
        'License Groups'         = $licenseGroupsStr
        'License Errors'         = $licenseErrorsStr
        'Last License Change'    = $lastLicenseChange
        'Monthly License Cost (EUR)' = $userMonthlyCost
        'Annual License Cost (EUR)'  = $userAnnualCost
        'Department'             = $userDept
        'Company'                = $userCompany
        'Country'                = $userCountry
        'User Type'              = $userType
        'Account Enabled'        = $isAccountEnabled
        'Mailbox Type'           = $mailboxType
        'Litigation Hold'        = $isLitigationHold
        'Admin Roles'            = $adminRolesStr
        'Admin Privilege Level'  = if ($isAdmin -and -not $isLowPrivAdmin) { 'High' } elseif ($isLowPrivAdmin) { 'Low' } else { '' }
        'PIM Eligible Roles'     = $pimEligibleRoles
        'PIM Active Roles'       = $pimActiveRoles
        'Risk-based CA Policies' = $riskBasedCA
        'MDO Policy Coverage'    = $mdoPolicyCoverage
        # Platform usage (M365 Apps)
        'Uses Desktop Apps'      = $usesDesktop
        'No Desktop Apps'        = $noDesktopApps
        'Uses Mobile Only'       = $usesMobileOnly
        'Desktop Apps Used'      = ($desktopApps | Sort-Object) -join ", "
        'Web Apps Used'          = ($webApps     | Sort-Object) -join ", "
        'Mobile Apps Used'       = ($mobileApps  | Sort-Object) -join ", "

        # Activations
        'Activated Platforms'    = $activatedPlatforms
        'Activated Products'     = $activatedProducts

        # Exchange
        'Exchange: Sent'         = $emailSend
        'Exchange: Received'     = $emailReceive
        'Exchange: Read'         = $emailRead
        'Exchange Intensity'     = $emailIntensity
        'Mailbox Size (MB)'      = $mbSizeMB
        'Mailbox Item Count'     = $mbItemCount
        'Has Archive Mailbox'    = $mbHasArchive
        'Archive Status'         = $archiveStatus
        'Auto-Expanding Archive' = $autoExpandingArchive
        'Email Clients Used'     = $emailClientsStr
        'No Outlook Desktop'     = $noOutlookDesktop

        # Teams
        'Teams: Team Chat'       = $teamsChatMsg
        'Teams: Private Chat'    = $teamsPrivateMsg
        'Teams: Calls'           = $teamsCalls
        'Teams: Meetings'        = $teamsMeetings
        'Teams: Meetings Organized' = $teamsMeetingsOrganized
        'Teams Intensity'        = $teamsIntensity
        'Teams Platforms'        = $teamsPlatStr
        'Teams No Desktop'       = $teamsNoDesktop

        # OneDrive
        'OneDrive: Files'        = $odViewed
        'OneDrive: Synced'       = $odSynced
        'OneDrive: Shared'       = $odShared
        'OneDrive Intensity'     = $odIntensity
        'OneDrive Storage (MB)'  = $odStorageMB
        'OneDrive File Count'    = $odFileCount

        # SharePoint
        'SharePoint: Files'      = $spViewed
        'SharePoint: Shared'     = $spShared
        'SharePoint: Pages'      = $spPages
        'SharePoint Intensity'   = $spIntensity

        # Last activity dates
        'Exchange Last Activity'   = if ($au) { $au.'Exchange Last Activity Date' } else { "" }
        'OneDrive Last Activity'   = if ($au) { $au.'OneDrive Last Activity Date' } else { "" }
        'SharePoint Last Activity' = if ($au) { $au.'SharePoint Last Activity Date' } else { "" }
        'Teams Last Activity'      = if ($au) { $au.'Teams Last Activity Date' } else { "" }

        # License flags from Active User report
        'Has Exchange License'     = if ($au) { $au.'Has Exchange License' } else { "" }
        'Has Teams License'        = if ($au) { $au.'Has Teams License' } else { "" }
        'Has OneDrive License'     = if ($au) { $au.'Has OneDrive License' } else { "" }
        'Has SharePoint License'   = if ($au) { $au.'Has SharePoint License' } else { "" }

        # Sign-in activity
        'Last Sign-In'             = $lastSignIn
        'Days Since Sign-In'       = $daysSinceSignIn
        'Last Non-Interactive Sign-In'          = $lastNonInteractiveSignIn
        'Days Since Non-Interactive Sign-In'    = $daysSinceNonInteractive
        'Dormant Account'          = $isDormant

        # Cloud Licensing
        'Trial License'            = $userIsOnTrial
        'Cloud License Errors'     = $userCloudErrors
        'Has Unknown SKU'          = $hasUnknownSku

        # Entitlements & data quality (LOA v1.0 spec §4.1, §2.2, §6.1)
        'Disabled Plans'           = $disabledPlansStr
        'Missing Data Sources'     = $missingDataSourcesStr

        # Copilot activity (from beta usage report, if available)
        'Copilot Active Apps'      = $(
            $cuRow = $lkpCopilotUsage[$upn]
            if ($cuRow) {
                $cpApps = @()
                if ($cuRow.'Microsoft Teams Copilot Last Activity Date') { $cpApps += "Teams" }
                if ($cuRow.'Word Copilot Last Activity Date')            { $cpApps += "Word" }
                if ($cuRow.'Excel Copilot Last Activity Date')           { $cpApps += "Excel" }
                if ($cuRow.'PowerPoint Copilot Last Activity Date')      { $cpApps += "PowerPoint" }
                if ($cuRow.'Outlook Copilot Last Activity Date')         { $cpApps += "Outlook" }
                if ($cuRow.'OneNote Copilot Last Activity Date')         { $cpApps += "OneNote" }
                if ($cuRow.'Loop Copilot Last Activity Date')            { $cpApps += "Loop" }
                if ($cuRow.'Copilot Chat Last Activity Date')            { $cpApps += "Chat" }
                if ($cpApps.Count -gt 0) { $cpApps -join "; " } else { "" }
            } else { "" }
        )
        'Copilot Last Activity'    = $(
            $cuRow2 = $lkpCopilotUsage[$upn]
            if ($cuRow2 -and $cuRow2.'Last Activity Date') { $cuRow2.'Last Activity Date' } else { "" }
        )

        # Cloud PC usage (from beta API, if available)
        'Cloud PC Type'            = if ($lkpCloudPcType.ContainsKey($upn))        { $lkpCloudPcType[$upn] }        else { "" }
        'Cloud PC Total Hours (90d)' = if ($lkpCloudPcUsageHours.ContainsKey($upn)) { [math]::Round($lkpCloudPcUsageHours[$upn], 1) } else { "" }
        'Cloud PC Days Since Sign-In' = if ($lkpCloudPcDaysSinceSignIn.ContainsKey($upn)) { $lkpCloudPcDaysSinceSignIn[$upn] } else { "" }
        'Cloud PC Last Active'     = if ($lkpCloudPcLastActive.ContainsKey($upn))  { $lkpCloudPcLastActive[$upn] }  else { "" }
        'Cloud PC Device Name'     = if ($lkpCloudPcDeviceName.ContainsKey($upn))  { $lkpCloudPcDeviceName[$upn] }  else { "" }
        'Cloud PC Status'          = if ($lkpCloudPcStatus.ContainsKey($upn))      { $lkpCloudPcStatus[$upn] }      else { "" }

        # Security & Compliance posture
        'Security Coverage'        = $securityCoverageLevel
        'Compliance Coverage'      = $complianceCoverageLevel

        # Recommendation
        'Recommendation'           = $recommendationText
        'Recommendation Category'  = $recCategory
        'Recommendation Confidence' = $recConfidence
    }

    $mainWriter.WriteLine((ConvertTo-CsvLine -Row $row -Columns $csvColumns))
    $totalUsers++

    # ── Inline summary statistics ──
    $rec  = $recommendationText
    $cost = $userAnnualCost
    if ($userMonthlyCost) { $totalMonthlySpendAcc += $userMonthlyCost }

    $isLic = ($assignedSkus -ne '[UNLICENSED]' -and $assignedSkus -ne '[NOT IN DIRECTORY]')

    if ($noDesktopApps   -eq $true)  { $noDesktopCount++ }
    if ($usesMobileOnly  -eq $true)  { $mobileOnly++ }
    if ($assignedSkus -eq '[UNLICENSED]') { $unlicensed++ }
    if ($emailIntensity -eq 'Low' -and $au -and $au.'Has Exchange License' -in @('True','Yes')) { $lowExchange++ }
    if ($isDormant       -eq $true)  { $dormantUsers++ }
    if ($rec -match "NEVER SIGNED IN")  { $neverSignedIn++ }
    if ($noOutlookDesktop -eq $true) { $noOutlookDesktopCount++ }
    if ($teamsNoDesktop  -eq $true)  { $teamsNoDesktopCount++ }
    if ($isLitigationHold -eq $true) { $litigationHold++ }
    if ($adminRolesStr   -ne '')     { $adminUsers++ }
    if ($overlappingSkus -ne '')     { $overlapping++ }

    if ($mailboxType -eq 'SharedMailbox')                                   { $sharedMbx++ }
    if ($mailboxType -eq 'RoomMailbox' -or $mailboxType -eq 'EquipmentMailbox') { $roomEquipMbx++ }

    if ($userType -eq 'Guest' -and $isLic) { $guestsLicensed++ }
    if ($isAccountEnabled -eq $false -and $isLic) { $disabledLicensed++ }

    if ($rec -match "(^|\| )NO ACTIVITY")        { $noActivity++;       if ($cost -and $rec -notmatch "(^|\| )DORMANT —|(^|\| )DISABLED ACCOUNT|(^|\| )INACTIVE HOLD|(^|\| )SHARED MAILBOX") { if ($userNoActRightsizeSave -gt 0) { $noActivityCostAcc += $userNoActRightsizeSave } else { $noActivityCostAcc += [math]::Max(0, $cost - $userCopilotAnnualCost - $dupAnnualWaste) } } }
    if ($rec -match "DUPLICATE COVERAGE|DUPLICATE REVIEW") { $duplicateCov++ }
    if ($rec -match "E5 CONSOLIDATION")          { $e5Upgrade++ }
    if ($rec -match "SUITE INVERSION")          { $suiteInversion++ }
    if ($rec -match "BUNDLE CONSOLIDATION")     { $bundleConsolidation++ }
    # Count product shelfware per-recommendation (not on joined $rec) to avoid INTUNE REVIEW masking Visio/Project shelfware
    $hasProductShelfware = @($recommendations | Where-Object { $_ -match "^INACTIVE ADD-ON —" }).Count -gt 0
    if ($hasProductShelfware)  { $shelfware++;        if ($userShelfwareCost -gt 0 -and $rec -notmatch "(^|\| )DORMANT —|(^|\| )DISABLED ACCOUNT|(^|\| )INACTIVE HOLD|(^|\| )NO ACTIVITY|(^|\| )SHARED MAILBOX") { $shelfwareCostAcc += $userShelfwareCost } }
    if ($rec -match "INACTIVE ADD-ON REVIEW")          { $shelfwareReview++ }
    if ($rec -match "PREMIUM ADD-ON REVIEW")     { $premiumAddonWaste++ }
    if ($rec -match "TEAMS PHONE REVIEW")        { $phoneNoPlan++ }
    if ($rec -match "TEAMS PHONE RIGHT-SIZING") { $teamsPhoneRightSizing++ }
    # Use $rec -match (not $recCategory) so this counter stays consistent with sub-counters
    # ($copilotReclaim, $copilotWatchlist, $copilotKeep) which also use $rec -match patterns.
    # $recCategory is primary-only; a Dormant user with secondary Copilot rec must still count here.
    if ($rec -match "(^|\| )COPILOT (RECLAIM|WATCHLIST|ACTIVE)") { $copilotUsers++ }
    if ($rec -match "POWER BI PRO REVIEW") { $pbiProReview++ }
    if ($rec -match "(^|\| )FRONTLINE CANDIDATE") { $frontlineCandidate++; if ($cost) { $frontlineCostAcc += $cost } }
    if ($rec -match "EXO PLAN 2 DOWNGRADE")      { $exoPlan2Review++ }
    if ($recCategory -eq "Licensing Compliance Gap") { $licensingCheck++
        if ($rec -match "Conditional Access")  { $licensingCheckCA++ }
        if ($rec -match "Defender for Office|Safe Links|Safe Attachments|MDO") { $licensingCheckMDO++ }
        if ($rec -match "PIM")                 { $licensingCheckPIM++ }
        if ($rec -match "(^|\| )LICENSING CHECK.*desktop" -or $rec -match "(^|\| )LICENSING CHECK.*Frontline") { $licensingCheckFrontline++ }
    }
    if ($rec -match "SECURITY GAP")             { $securityGap++ }
    if ($rec -match "DEFENDER COVERAGE REVIEW")    { $defenderUpsell++ }
    if ($rec -match "COMPLIANCE COVERAGE REVIEW")           { $purviewUpsell++ }
    if ($rec -match "LICENSING ERROR")          { $licenseErrors++ }
    if ($rec -match "TRIAL LICENSE")            { $trialLicenseUsers++ }
    if ($rec -match "LICENSE CAPACITY QUEUE")   { $capacityQueueUsers++ }
    if ($rec -match "BUSINESS BASIC CANDIDATE") { $businessDowngrade++ }
    if ($rec -match "E1 DOWNGRADE CANDIDATE")   { $e1Downgrade++ }
    if ($rec -match "O365 E3 TO E1")             { $o365E3Downgrade++ }
    if ($rec -match "E3 TO BUSINESS PREMIUM")    { $e3Downgrade++ }
    if ($rec -match "BUSINESS PREMIUM INVERSION") { $bizPremInversion++ }
    if ($rec -match "BUSINESS PREMIUM SECURITY REVIEW") { $bizPremSecReview++ }
    if ($rec -match "E5 VOICE REVIEW")            { $e5VoiceWaste++ }
    if ($rec -match "APP ARBITRAGE")              { $appArbitrage++ }
    if ($rec -match "PBI PPU OVERLAP")      { $ppuArbitrage++ }
    if ($rec -match "CALLING PLAN REVIEW")        { $callingPlanWaste++ }
    if ($rec -match "ONEDRIVE PLAN 2 REVIEW")    { $odPlan2Waste++ }
    if ($rec -match "ENTRA P2 DOWNGRADE")       { $entraP2Downgrade++ }
    if ($rec -match "EXCHANGE KIOSK CANDIDATE")  { $exoKioskDowngrade++ }
    if ($rec -match "DATA GAP")                { $dataGapUsers++ }
    if ($rec -match "FRONTLINE BLOCKED")         { $frontlineBlocked++ }
    if ($rec -match "FRONTLINE RESCUE")          { $frontlineRescue++ }
    if ($rec -match "(^|\| )FRONTLINE REVIEW")   { $frontlineReview++ }
    if ($rec -match "BUSINESS BASIC REVIEW")    { $businessReview++ }
    if ($rec -match "MAILBOX STORAGE WARNING")  { $mailboxStorageWarning++ }
    if ($rec -match "AI ADD-ON OVERLAP")         { $aiAddonOverlap++ }
    if ($rec -match "AI OVERLAP REVIEW")         { $aiOverlapReview++ }
    if ($rec -match "ENTRA SUITE OVERLAP")       { $entraSuiteOverlap++ }
    if ($rec -match "TEAMS UNBUNDLING")          { $teamsUnbundling++ }
    if ($rec -match "GUEST ACCOUNT REVIEW")       { $guestAccountWaste++ }
    if ($rec -match "INTUNE SUITE OVERLAP")        { $intuneSuiteWaste++ }
    if ($rec -match "NON-HUMAN ACCOUNT REVIEW")   { $nonHumanWaste++ }
    if ($rec -match "DORMANT ADMIN REVIEW")       { $dormantAdminRisk++ }
    if ($rec -match "FREE LICENSE OVERLAP")    { $viralCleanup++ }
    if ($rec -match "WINDOWS LICENSE REVIEW")    { $windowsLicenseWaste++ }
    if ($rec -match "OVER-LICENSED ARCHIVE")    { $overLicensedArchive++ }
    if ($rec -match "REDUNDANT ARCHIVE")       { $redundantArchive++ }
    if ($rec -match "STANDALONE APPS REVIEW")    { $standaloneAppsWaste++ }
    if ($rec -match "(^|\| )BUNDLE OPPORTUNITY" -and $rec -match "Exchange (Kiosk|Plan 1)") { $alaCarteWaste++ }
    if ($rec -match "(^|\| )BUNDLE OPPORTUNITY" -and $rec -match "consolidat|includes both") { $bundleInefficiency++ }
    if ($rec -match "F3 TO F1 DOWNGRADE")       { $f3ToF1Downgrade++ }
    if ($rec -match "EXTERNAL SHARING REVIEW")        { $highRiskSharing++ }
    if ($rec -match "LEGACY SERVICE ACCOUNT")  { $legacyServiceAccount++ }
    if ($rec -match "AUTOMATION ACCOUNT")      { $automationAccount++ }
    if ($rec -match "(^|\| )INACTIVE MAILBOX")  { $inactiveMailbox++ }
    if ($rec -match "EXPENSIVE COLD STORAGE")  { $expensiveColdStorage++ }
    if ($rec -match "(^|\| )INTUNE REVIEW" -and $rec -match "0 enrolled|no enrolled") { $intuneShelfware++ }
    if ($rec -match "(^|\| )INTUNE REVIEW" -and $rec -match "web.only access|mobile.only") { $mdmMamWaste++ }
    if ($rec -match "BACKGROUND SYNC ONLY")    { $backgroundSyncOnly++ }
    if ($rec -match "(^|\| )INACTIVE HOLD WITH LICENSE")           { $e5DataHoarder++ }
    if ($rec -match "(^|\| )INACTIVE HOLD" -and $rec -notmatch "(^|\| )INACTIVE HOLD WITH LICENSE") { $inactiveHold++ }
    if ($rec -match "SEEDED VISIO OVERLAP")      { $seededVisioOverlap++ }
    if ($rec -match "FRONTLINE ADD-ON STACKING")  { $frontlineAddonBloat++ }
    if ($rec -match "COPILOT PREREQUISITE")     { $copilotPrereq++ }
    if ($rec -match "(^|\| )COPILOT RECLAIM")    { $copilotReclaim++; $copilotNonAdopter++ }
    if ($rec -match "(^|\| )COPILOT WATCHLIST")  { $copilotWatchlist++; $copilotNonAdopter++ }
    if ($rec -match "COPILOT ACTIVE")           { $copilotKeep++ }
    if ($rec -match "COPILOT STUDIO")           { $copilotStudioUsers++ }
    if ($rec -match "(^|\| )DORMANT CLOUD PC")  { $dormantCloudPc++ }
    if ($rec -match "(^|\| )CLOUD PC REVIEW")  { $cloudPcReview++ }
    if ($rec -match "ONEDRIVE STORAGE WARNING")  { $oneDriveStorageWarning++ }
    if ($rec -match "UNLICENSED WITH DATA")      { $unlicensedWithData++ }
    if ($rec -match "(^|\| )DISABLED ACCOUNT" -and $rec -match "free SKU") { $disabledFreeSku++ }
    # Deduct Copilot-specific cost from cost buckets that feed $totalIdentifiedWaste, because
    # $copilotReclaimCostAcc already captures that share separately — subtracting it here
    # prevents the same Copilot license cost from being counted twice in the tier 1 total.
    if ($missingDataSources.Count -gt 0)        { $missingSourceUsers++ }
    if ($rec -match "(^|\| )DORMANT —" -and $rec -notmatch "(^|\| )STALE SIGN-IN" -and $rec -notmatch "(^|\| )DORMANT CLOUD PC" -and $rec -notmatch "(^|\| )DORMANT ADMIN" -and $rec -notmatch "(^|\| )AUTOMATION ACCOUNT" -and $rec -notmatch "(^|\| )DISABLED ACCOUNT|(^|\| )INACTIVE HOLD") { $dormantTier1Count++; if ($cost) { $dormantCostAcc += [math]::Max(0, $cost - $userCopilotAnnualCost - $dupAnnualWaste) } }
    if ($rec -match "(^|\| )DISABLED ACCOUNT|(^|\| )DISABLED SHARED MAILBOX|(^|\| )INACTIVE HOLD WITH LICENSE|(^|\| )INACTIVE HOLD") { if ($cost) { $disabledCostAcc += [math]::Max(0, $cost - $userCopilotAnnualCost - $dupAnnualWaste) } }
    # Exclude MDO-protected ("A license is needed") and active-archive ("Removing the license will disable") shared mailboxes
    # from the removable count — those recommendations advise retaining or downgrading, not removing.
    if ($rec -match "(^|\| )SHARED MAILBOX \(" -and $rec -notmatch "(^|\| )SHARED MAILBOX REVIEW" -and $rec -notmatch "A license is needed|Removing the license will disable") { $sharedMbxRemovable++; if ($cost) { $sharedMbxCostAcc += [math]::Max(0, $cost - $userCopilotAnnualCost - $dupAnnualWaste) } }
    if ($rec -match "(^|\| )FORWARDING MAILBOX REVIEW" -and $rec -match "no interactive sign-in|no sign-in") { $forwardingWaste++ }
    if ($rec -match "(^|\| )FORWARDING MAILBOX REVIEW" -and $rec -match "low exchange|low email") { $forwardingReview++ }

    # ── Security/Compliance posture counters ──
    if ($isLic) {
        switch ($securityCoverageLevel) {
            "None"           { $secCoverageNone++ }
            "Basic"          { $secCoverageBasic++ }
            "Advanced"       { $secCoverageAdvanced++ }
            "E5-equivalent"  { $secCoverageE5++ }
        }
        switch ($complianceCoverageLevel) {
            "None"           { $compCoverageNone++ }
            "Basic"          { $compCoverageBasic++ }
            "Advanced"       { $compCoverageAdvanced++ }
            "E5-equivalent"  { $compCoverageE5++ }
        }
    }

    # ── Cost-by-dimension accumulation ──
    if ($userDept) {
        if (-not $deptCostDict.ContainsKey($userDept)) { $deptCostDict[$userDept] = @{ Users = 0; AnnualCost = [decimal]0 } }
        $deptCostDict[$userDept].Users++
        $deptCostDict[$userDept].AnnualCost += $userAnnualCost
    }

    # ── Recommendation distribution ──
    if (-not $recDistribution.ContainsKey($recCategory)) { $recDistribution[$recCategory] = @{ Count = 0; AnnualCost = [decimal]0 } }
    $recDistribution[$recCategory].Count++
    $recDistribution[$recCategory].AnnualCost += $userAnnualCost

    # ── Intensity cross-tab ──
    if ($emailIntensity -and $teamsIntensity) {
        $intensityKey = "$emailIntensity,$teamsIntensity"
        if (-not $intensityCrossDict.ContainsKey($intensityKey)) { $intensityCrossDict[$intensityKey] = 0 }
        $intensityCrossDict[$intensityKey]++
    }
}

} finally {
    # Ensure StreamWriter is disposed even on error
    if ($mainWriter) {
        try { $mainWriter.Flush(); $mainWriter.Close(); $mainWriter.Dispose() } catch { }
        $mainWriter = $null
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 7 — Export
# ═══════════════════════════════════════════════════════════════════════════════

Write-Host "`n[11/12] Exporting reports ..." -ForegroundColor Cyan
Write-Log "[11/12] Exporting reports"

# Main CSV and Service Plan CSV were already streamed during the merge loop
Write-Host "  [1] License Optimization : $mainFile ($totalUsers users)" -ForegroundColor Green
Write-Host "  [2] Service Plan Detail  : $planFile ($servicePlanRowCount rows)" -ForegroundColor Green

# 3. SKU inventory (enriched with friendly names and subscription lifecycle)
$skuFile = Join-Path $OutputFolder "M365_SkuInventory_$ts.csv"
$subscribedSkus | Select-Object SkuPartNumber,
    @{N='Friendly Name'; E={ Resolve-SkuFriendlyName $_.SkuPartNumber }},
    SkuId, AppliesTo, CapabilityStatus,
    @{N='Total';     E={$_.PrepaidUnits.Enabled + $_.PrepaidUnits.Warning}},
    @{N='Warning';   E={$_.PrepaidUnits.Warning}},
    @{N='Suspended'; E={$_.PrepaidUnits.Suspended}},
    ConsumedUnits,
    @{N='Available'; E={$_.PrepaidUnits.Enabled + $_.PrepaidUnits.Warning - $_.ConsumedUnits}},
    @{N='Monthly Unit Price (EUR)'; E={ Get-SkuMonthlyPrice $_.SkuPartNumber }},
    @{N='Annual Total Cost (EUR)'; E={ [math]::Round((Get-SkuMonthlyPrice $_.SkuPartNumber) * 12 * $_.ConsumedUnits, 2) }},
    @{N='Subscription Status'; E={
        if ($lkpSubscription.ContainsKey($_.SkuPartNumber)) { $lkpSubscription[$_.SkuPartNumber].Status } else { "" }
    }},
    @{N='Next Lifecycle Date'; E={
        if ($lkpSubscription.ContainsKey($_.SkuPartNumber) -and $lkpSubscription[$_.SkuPartNumber].NextLifecycle) {
            ([datetime]$lkpSubscription[$_.SkuPartNumber].NextLifecycle).ToString("yyyy-MM-dd")
        } else { "" }
    }},
    @{N='Days Until Expiry'; E={
        if ($lkpSubscription.ContainsKey($_.SkuPartNumber) -and $lkpSubscription[$_.SkuPartNumber].NextLifecycle) {
            [int]([datetime]$lkpSubscription[$_.SkuPartNumber].NextLifecycle - (Get-Date)).TotalDays
        } else { "" }
    }},
    @{N='Is Trial'; E={
        if ($cloudLicensingData.ContainsKey($_.SkuPartNumber)) { $cloudLicensingData[$_.SkuPartNumber].IsTrial } else { "" }
    }},
    @{N='Cloud State'; E={
        if ($cloudLicensingData.ContainsKey($_.SkuPartNumber)) { $cloudLicensingData[$_.SkuPartNumber].State } else { "" }
    }},
    @{N='Capacity Utilization %'; E={
        if ($cloudLicensingData.ContainsKey($_.SkuPartNumber)) { $cloudLicensingData[$_.SkuPartNumber].CapacityPct } else { "" }
    }},
    @{N='Pricing Known'; E={ $skuMonthlyPrices.ContainsKey($_.SkuPartNumber) }} |
    Export-Csv -Path $skuFile -NoTypeInformation -Encoding UTF8
Write-Host "  [3] SKU Inventory        : $skuFile" -ForegroundColor Green

# 3b. Group Licensing CSV (for heatmap consumption)
# Backfill "Also Direct" column with per-group overlap counts from the merge loop
foreach ($grp in $groupLicenseInventory) {
    if ($groupDirectOverlap.ContainsKey($grp.'Group ID')) {
        $grp.'Also Direct' = $groupDirectOverlap[$grp.'Group ID']
    }
}
if ($groupLicenseInventory.Count -gt 0) {
    $groupFile = Join-Path $OutputFolder "M365_LicenseGroups_$ts.csv"
    $groupLicenseInventory | Export-Csv -Path $groupFile -NoTypeInformation -Encoding UTF8
    Write-Host "  [3b] License Groups      : $groupFile" -ForegroundColor Green
}

# 4. Summary stats — already computed inline during the merge loop (no second pass needed)

# ── Cost aggregation ──
$totalMonthlySpend = [math]::Round($totalMonthlySpendAcc, 2)
$totalAnnualSpend  = [math]::Round(($totalMonthlySpendAcc * 12) + $unassignedPoolTotalAnnual, 2)

$dormantCost    = [math]::Round($dormantCostAcc, 2)
$disabledCost   = [math]::Round($disabledCostAcc, 2)
$noActivityCost = [math]::Round($noActivityCostAcc, 2)
$shelfwareCost  = [math]::Round($shelfwareCostAcc, 2)
$copilotReclaimCost     = [math]::Round($copilotReclaimCostAcc, 2)
$copilotWatchlistCost   = [math]::Round($copilotWatchlistCostAcc, 2)
$sharedMbxCost  = [math]::Round($sharedMbxCostAcc, 2)
$frontlineCost  = [math]::Round($frontlineCostAcc, 2)
$totalIdentifiedWaste = [math]::Round($dormantCost + $disabledCost + $noActivityCost + $shelfwareCost + $copilotReclaimCost + $sharedMbxCost, 2)

# ── Executive Financial Summary tier variables ──
$duplicateCost        = [math]::Round($duplicateCostAcc, 2)
$frontlineSavings     = [math]::Round($frontlineSavingsAcc, 2)
$businessBasicSavings = [math]::Round($businessBasicSavingsAcc, 2)
$exoPlan2Savings      = [math]::Round($exoPlan2SavingsAcc, 2)
$e5UpgradeSavings     = [math]::Round($e5UpgradeSavingsAcc, 2)
$bundleConsolidationSavings = [math]::Round($bundleConsolidationSavingsAcc, 2)
$e1DowngradeSavings   = [math]::Round($e1DowngradeSavingsAcc, 2)
$o365E3DowngradeSavings = [math]::Round($o365E3DowngradeSavingsAcc, 2)
$e3DowngradeSavings   = [math]::Round($e3DowngradeSavingsAcc, 2)
$e5VoiceSavings       = [math]::Round($e5VoiceSavingsAcc, 2)
$appArbitrageSavings  = [math]::Round($appArbitrageSavingsAcc, 2)
$ppuArbitrageSavings  = [math]::Round($ppuArbitrageSavingsAcc, 2)
$exoKioskSavings      = [math]::Round($exoKioskSavingsAcc, 2)
$bizPremInversionSavings = [math]::Round($bizPremInversionSavingsAcc, 2)
$frontlineRescueSavings  = [math]::Round($frontlineRescueSavingsAcc, 2)
# Tier 1 = quick wins — existing waste + duplicate coverage
$tier1Waste           = [math]::Round($totalIdentifiedWaste + $duplicateCost, 2)
# Tier 2 = right-sizing opportunities (downgrade SKU delta)
$tier2Savings         = [math]::Round($frontlineSavings + $businessBasicSavings + $exoPlan2Savings + $e5UpgradeSavings + $bundleConsolidationSavings + $e1DowngradeSavings + $o365E3DowngradeSavings + $e3DowngradeSavings + $e5VoiceSavings + $appArbitrageSavings + $ppuArbitrageSavings + $exoKioskSavings + $bizPremInversionSavings + $frontlineRescueSavings, 2)
$totalMoneyOnTable    = [math]::Round($tier1Waste + $tier2Savings + $unassignedPoolTotalAnnual, 2)
$wastePercentage      = if ($totalAnnualSpend -gt 0) { [math]::Round($totalMoneyOnTable / $totalAnnualSpend * 100, 1) } else { 0 }
$tier1Percentage      = if ($totalAnnualSpend -gt 0) { [math]::Round($tier1Waste / $totalAnnualSpend * 100, 1) } else { 0 }
$tier2Percentage      = if ($totalAnnualSpend -gt 0) { [math]::Round($tier2Savings / $totalAnnualSpend * 100, 1) } else { 0 }
$poolPercentage       = if ($totalAnnualSpend -gt 0) { [math]::Round($unassignedPoolTotalAnnual / $totalAnnualSpend * 100, 1) } else { 0 }

# Post-run diagnostics: recommendation engine summary
Write-Log "Recommendation summary: $totalUsers users, $($recDistribution.Count) categories"
Write-Log "  Tier 1 (Quick Wins): $($tier1Waste.ToString('N2')) EUR — dormant=$dormantTier1Count, disabled=$(if ($recDistribution.ContainsKey('Disabled Account')) { $recDistribution['Disabled Account'].Count } else { 0 }), neverSignedIn=$neverSignedIn, noActivity=$noActivity, sharedMbx=$sharedMbxRemovable, copilotReclaim=$copilotReclaim"
Write-Log "  Tier 2 (Right-Sizing): $($tier2Savings.ToString('N2')) EUR — frontline=$frontlineCandidate, e1Downgrade=$e1Downgrade, duplicate=$duplicateCov, e5Voice=$e5VoiceWaste"
Write-Log "  Pool waste: $($unassignedPoolTotalAnnual.ToString('N2')) EUR ($($unassignedPoolWarnings.Count) SKU(s))"
Write-Log "  Total annual spend: $($totalAnnualSpend.ToString('N2')) EUR, Total savings potential: $($totalMoneyOnTable.ToString('N2')) EUR ($wastePercentage%)"

# Cost breakdown by Department (from running dictionary — no Group-Object needed)
$costByDepartment = @($deptCostDict.GetEnumerator() | ForEach-Object {
    [PSCustomObject]@{
        Department         = $_.Key
        Users              = $_.Value.Users
        'Annual Cost (EUR)' = [math]::Round($_.Value.AnnualCost, 2)
    }
} | Sort-Object 'Annual Cost (EUR)' -Descending)


# Format cost breakdown strings for summary
$deptCostStr = if ($costByDepartment.Count -gt 0) {
    ($costByDepartment | Select-Object -First 15 | ForEach-Object { "  $($_.Department): $($_.Users) users, €$(($_.'Annual Cost (EUR)').ToString('N2'))/yr" }) -join "`n"
} else { "  (no department data)" }

# Subscription expiry warnings
$expiringSkus = @()
foreach ($sku in $subscribedSkus) {
    if ($lkpSubscription.ContainsKey($sku.SkuPartNumber) -and $lkpSubscription[$sku.SkuPartNumber].NextLifecycle) {
        $daysLeft = [int]([datetime]$lkpSubscription[$sku.SkuPartNumber].NextLifecycle - (Get-Date)).TotalDays
        $status   = $lkpSubscription[$sku.SkuPartNumber].Status
        if ($daysLeft -le 90 -or $status -ne "Enabled") {
            $expiringSkus += "  $(Resolve-SkuFriendlyName $sku.SkuPartNumber): $status, ${daysLeft}d remaining"
        }
    }
}
$expiryWarning = if ($expiringSkus.Count -gt 0) {
    "`nSUBSCRIPTION WARNINGS ($($expiringSkus.Count) SKU(s) expiring within 90d or non-Enabled):`n" + ($expiringSkus -join "`n")
} else { "" }

# Trial subscription warnings (from Cloud Licensing API)
$trialWarnings = @()
if ($cloudLicensingLoaded) {
    foreach ($entry in $cloudLicensingData.GetEnumerator()) {
        if ($entry.Value.IsTrial) {
            $trialFriendly = Resolve-SkuFriendlyName $entry.Key
            $consumed      = $entry.Value.ConsumedUnits
            $trialDays     = ""
            $nlDate        = $entry.Value.NextLifecycle
            if ($nlDate) {
                try {
                    $dLeft = [int]([datetime]$nlDate - (Get-Date)).TotalDays
                    $trialDays = ", ${dLeft}d remaining"
                } catch { }
            }
            $trialWarnings += "  $trialFriendly (TRIAL$trialDays, $consumed assigned)"
        }
    }
}
$trialWarningText = if ($trialWarnings.Count -gt 0) {
    "`nTRIAL SUBSCRIPTION WARNINGS ($($trialWarnings.Count) SKU(s) on trial):`n" + ($trialWarnings -join "`n") +
    "`n  Action: Convert trials to paid subscriptions or remove before expiry to avoid service disruption."
} else { "" }

# Append trial warnings to expiry section
$expiryWarning = $expiryWarning + $trialWarningText

# Business 300-seat limit warning text
$businessLimitText = if ($businessLimitWarnings.Count -gt 0) {
    "`nBUSINESS 300-SEAT LIMIT WARNING:`n" + ($businessLimitWarnings -join "`n")
} else { "" }

$summaryFile = Join-Path $OutputFolder "M365_OptimizationSummary_$ts.txt"
$summary = @"
M365 LICENSE ASSESSMENT SUMMARY
$(Get-Date -Format 'yyyy-MM-dd')
Report Period: $ReportPeriod
Tenant: $($ctx.TenantId)
Mapping Version: $MappingVersion | Recommendation Logic: $RecommendationLogicVersion
SKU Pricing Source: $_pricingCsvFile$(if ($skuDataLoaded) { " + $skuJsonPath" } else { '' })$(if ($skuDataDate) { "`nSKU Data Date: $($skuDataDate.ToString('yyyy-MM-dd')) ($skuDataAge days old$(if ($skuDataAge -gt $SkuStalenessDays) { ' — STALE' } else { '' }))" } else { "" })
================================================================
NOTE: All cost figures are INDICATIVE estimates based on public Microsoft
list prices (EUR). Actual costs may differ due to EA/CSP/volume pricing.
================================================================

EXECUTIVE FINANCIAL SUMMARY
  Total Annual M365 Spend      : €$($totalAnnualSpend.ToString('N2'))
  ╔══════════════════════════════════════════════════════════════╗
  ║  Estimated Optimization Potential   : €$($totalMoneyOnTable.ToString('N2'))  ($wastePercentage% of annual spend)
  ╚══════════════════════════════════════════════════════════════╝

  NOTE: Users may appear in multiple categories below (e.g. a dormant account
  can also have a compliance gap). The "Recommendation Distribution" section
  shows the single primary category assigned to each user.

  TIER 1 — Quick Wins:
    Dormant accounts (no sign-in >$($InactiveSignInDays)d)     : €$($dormantCost.ToString('N2'))  ($dormantTier1Count flagged$(if ($recDistribution.ContainsKey('Dormant')) { ", $($recDistribution['Dormant'].Count) primary" } else { '' }))
    Disabled accounts (sign-in blocked)    : €$($disabledCost.ToString('N2'))  ($($disabledLicensed - $disabledFreeSku) paid, $disabledFreeSku free-SKU-only)
    Zero M365 usage (no app activity)      : €$($noActivityCost.ToString('N2'))  ($noActivity flagged$(if ($recDistribution.ContainsKey('No Activity')) { ", $($recDistribution['No Activity'].Count) primary" } else { '' }))
    Unused premium add-ons                 : €$($shelfwareCost.ToString('N2'))  ($shelfware flagged$(if ($recDistribution.ContainsKey('Inactive Add-On')) { ", $($recDistribution['Inactive Add-On'].Count) primary" } else { '' }))
    Copilot reclaim (zero usage/readiness) : €$($copilotReclaimCost.ToString('N2'))  ($copilotReclaim users)
    Copilot at risk (zero usage, active)   : €$($copilotWatchlistCost.ToString('N2'))  ($copilotWatchlist users)  [advisory — not in subtotal]
    Shared mailbox (no license needed <50 GB) : €$($sharedMbxCost.ToString('N2'))  ($sharedMbxRemovable flagged$(if ($recDistribution.ContainsKey('Shared Mailbox Review')) { ", $($recDistribution['Shared Mailbox Review'].Count) primary" } else { '' }))
    Duplicate licenses (standalone in suite)  : €$($duplicateCost.ToString('N2'))  ($duplicateCov flagged$(if ($recDistribution.ContainsKey('Duplicate Coverage')) { ", $($recDistribution['Duplicate Coverage'].Count) primary" } else { '' }))
    ────────────────────────────────────────
    Tier 1 Subtotal             : €$($tier1Waste.ToString('N2'))/yr  ($tier1Percentage%)
$(if ($unassignedLicenseInventory.Count -gt 0) {
    $uLicLines = ($unassignedLicenseInventory | ForEach-Object { "    $($_.FriendlyName): $($_.Unassigned)/$($_.Total) unassigned" }) -join "`n"
@"

  UNASSIGNED LICENSES ($totalUnassignedSeats seats across $($unassignedLicenseInventory.Count) SKU(s)):
$uLicLines
"@
} else { '' })

  TIER 2 — Right-Sizing Opportunities:
    E3/E5 → Frontline F1/F3 (web/mobile only)       : €$($frontlineSavings.ToString('N2'))  ($frontlineCandidate users)
    Biz Standard → Basic (no desktop apps used)      : €$($businessBasicSavings.ToString('N2'))  ($businessDowngrade users)
    Exchange Plan 2 → Plan 1 (mailbox <50 GB)        : €$($exoPlan2Savings.ToString('N2'))  ($exoPlan2Review users)
    E3 + Add-Ons → E5 Upgrade (cheaper as E5)        : €$($e5UpgradeSavings.ToString('N2'))  ($($e5Upgrade + $suiteInversion) users: $suiteInversion inversion + $e5Upgrade consolidation)
    O365+EMS+Windows → M365 Bundle (cheaper combined): €$($bundleConsolidationSavings.ToString('N2'))  ($bundleConsolidation users)
    O365 E1 → Biz Basic (same features, lower cost)  : €$($e1DowngradeSavings.ToString('N2'))  ($e1Downgrade users)
    O365 E3 → E1 (web/mobile only, mailbox <50 GB)   : €$($o365E3DowngradeSavings.ToString('N2'))  ($o365E3Downgrade users)
    M365 E3 → Biz Premium (<300 seats, cheaper)      : €$($e3DowngradeSavings.ToString('N2'))  ($e3Downgrade users)
    E5 → No Audio Conf. Variant (0 calls)            : €$($e5VoiceSavings.ToString('N2'))  ($e5VoiceWaste flagged$(if ($recDistribution.ContainsKey('E5 Voice Review')) { ", $($recDistribution['E5 Voice Review'].Count) primary" } else { '' }))
    Apps Enterprise → Apps Business (<300 seats)      : €$($appArbitrageSavings.ToString('N2'))  ($appArbitrage users)
    PBI PPU Standalone → Add-On (Pro from suite)     : €$($ppuArbitrageSavings.ToString('N2'))  ($ppuArbitrage users)
    Exchange Plan 1 → Kiosk (web-only, <2 GB)        : €$($exoKioskSavings.ToString('N2'))  ($exoKioskDowngrade users)
    Biz Std + Add-Ons → Premium (cheaper)            : €$($bizPremInversionSavings.ToString('N2'))  ($bizPremInversion users)
    F-License blocked → E1/Basic alternative         : €$($frontlineRescueSavings.ToString('N2'))  ($frontlineRescue users)
    ────────────────────────────────────────
    Tier 2 Subtotal             : €$($tier2Savings.ToString('N2'))/yr  ($tier2Percentage%)
$(if ($unassignedPoolWarnings.Count -gt 0) {
    $poolLines = ($unassignedPoolWarnings | ForEach-Object { "    $($_.FriendlyName): $($_.Unassigned)/$($_.Total) unassigned ($($_.UnassignedPct)%) — €$($_.AnnualWaste.ToString('N2'))/yr" }) -join "`n"
@"

  TENANT POOL — Unassigned License Waste:
$poolLines
    ────────────────────────────────────────
    Pool waste total              : €$($unassignedPoolTotalAnnual.ToString('N2'))/yr  ($poolPercentage%)
    Note: Unassigned licenses are paid but unused. Reduce seat count at next
    renewal or assign to users. Included in Estimated Optimization Potential total.
"@
} else { '' })
================================================================

COST ANALYSIS (EUR):
  Total monthly spend         : €$($totalMonthlySpend.ToString('N2'))
  Total annual spend          : €$($totalAnnualSpend.ToString('N2'))

  Identified waste (annual):
    Dormant accounts (no sign-in >$($InactiveSignInDays)d)     : €$($dormantCost.ToString('N2'))  ($dormantTier1Count flagged$(if ($recDistribution.ContainsKey('Dormant')) { ", $($recDistribution['Dormant'].Count) primary" } else { '' }))
    Disabled accounts (sign-in blocked)    : €$($disabledCost.ToString('N2'))  ($($disabledLicensed - $disabledFreeSku) paid, $disabledFreeSku free-SKU-only)
    Zero M365 usage (no app activity)      : €$($noActivityCost.ToString('N2'))  ($noActivity flagged$(if ($recDistribution.ContainsKey('No Activity')) { ", $($recDistribution['No Activity'].Count) primary" } else { '' }))
    Unused premium add-ons                 : €$($shelfwareCost.ToString('N2'))  ($shelfware flagged$(if ($recDistribution.ContainsKey('Inactive Add-On')) { ", $($recDistribution['Inactive Add-On'].Count) primary" } else { '' }))
    Copilot reclaim (zero usage/readiness) : €$($copilotReclaimCost.ToString('N2'))  ($copilotReclaim users)
    Copilot at risk (zero usage, active)   : €$($copilotWatchlistCost.ToString('N2'))  ($copilotWatchlist users)  [advisory — not in subtotal]
    Shared mailbox (no license needed)     : €$($sharedMbxCost.ToString('N2'))  ($sharedMbxRemovable flagged$(if ($recDistribution.ContainsKey('Shared Mailbox Review')) { ", $($recDistribution['Shared Mailbox Review'].Count) primary" } else { '' }))
    ────────────────────────────────────────
    Total identified waste    : €$($totalIdentifiedWaste.ToString('N2'))/yr (excl. duplicate coverage)

  Right-sizing potential (annual cost of affected users):
    Frontline candidates      : €$($frontlineCost.ToString('N2'))  ($frontlineCandidate users)

  Cost by Department (top 15):
$deptCostStr

================================================================

QUICK WINS:
  Disabled accounts (paid SKU) : $($disabledLicensed - $disabledFreeSku) ← sign-in blocked, license cost is wasted
  Disabled accounts (free SKU) : $disabledFreeSku ← free SKU only, no cost but cleanup candidate
  Overlapping license assign.  : $overlapping  ← same SKU direct + group (redundant)
  Duplicate suite coverage     : $duplicateCov ← standalone already included in suite
  Guest account waste          : $guestAccountWaste ← external/guest users with paid licenses
  Non-human account waste      : $nonHumanWaste ← shared/room mailboxes on premium suites
  Intune Suite waste           : $intuneSuiteWaste ← redundant Intune add-ons (included in E3/E5 since late 2025)
  Windows license waste        : $windowsLicenseWaste ← standalone Windows E3/E5 on Mac/Mobile-only users
  Over-licensed archive        : $overLicensedArchive ← Exchange Online Archiving add-on unused (small mailbox, no archive)
  Intune shelfware (0 devices) : $intuneShelfware ← Intune/EMS entitlement but zero enrolled devices in Intune
  MDM/MAM waste (web-only)     : $mdmMamWaste ← Intune/EMS entitlement on web-only users (no devices to manage)
  Forwarding mailbox waste      : $forwardingWaste ← paid license only used for mail forwarding (replace with Mail Contact/transport rule)
  Forwarding mailbox review     : $forwardingReview ← active user with forward-only config and low exchange activity
$(if ($teamsRoomsDowngrade) {
@"

TENANT-LEVEL OPTIMIZATION:
  Teams Rooms Pro → Basic        : $($teamsRoomsDowngrade.EligibleRooms) of $($teamsRoomsDowngrade.ProRooms) Pro room(s) can use free Basic ($($teamsRoomsDowngrade.BasicRooms)+$($teamsRoomsDowngrade.EligibleRooms)/25 cap)
                                   Potential savings: €$($teamsRoomsDowngrade.AnnualSavings.ToString('N2'))/yr
                                   Note: Basic lacks dual-screen, intelligent camera, AI recap, and
                                   cloud management. Verify room requirements before downgrading.
"@
} else { '' })
ACCOUNT & ROLE FLAGS:
  Total users analyzed         : $totalUsers
  Unlicensed users             : $unlicensed
  Admin accounts               : $adminUsers   ← should only have Entra ID P1/P2
  Guest users with licenses    : $guestsLicensed ← verify if guests need paid licenses
  Shared mailboxes             : $sharedMbx    ← may not need a license (under 50 GB)
  Litigation Hold (active)     : $litigationHold ← license can be removed; Microsoft creates a free Inactive Mailbox
  Inactive mailboxes (free)    : $inactiveMailbox ← unlicensed + litigation hold = free archival for eDiscovery
  Room/Equipment mailboxes     : $roomEquipMbx ← only need a Room license

ACTIVITY-BASED FLAGS:
  Dormant (no sign-in ${InactiveSignInDays}d)  : $dormantUsers  ← strongest signal for license removal
  Dormant admin accounts       : $dormantAdminRisk ← admin + dormant = high security risk
  Automation accounts          : $automationAccount ← dormant interactively but active non-interactive sign-in (scripts/tasks)
  Legacy service accounts      : $legacyServiceAccount ← email via POP3/IMAP4/SMTP only, no modern clients (scan-to-email/scripts)
  Never signed in (licensed)   : $neverSignedIn ← no sign-in on record, verify account usage
  NO activity in period        : $noActivity  ← review for potential license removal
  Expensive cold storage       : $expensiveColdStorage ← 0 activity + large mailbox/OneDrive — paying premium to store data
  Background sync only         : $backgroundSyncOnly ← 0 interactive activity but OneDrive syncing (abandoned device)
  Low Exchange usage (licensed): $lowExchange

RIGHT-SIZING OPPORTUNITIES:
  Frontline candidates (E3/E5) : $frontlineCandidate ← premium suite but only web/mobile usage
  Frontline blocked (archive)  : $frontlineBlocked ← fits frontline profile but archive mailbox prevents downgrade
  Business Basic candidates    : $businessDowngrade ← Business Standard but only web/mobile usage
  E1 → Business Basic arbitrage : $e1Downgrade ← O365 E1 costs more than Business Basic for identical capabilities
  O365 E3 → E1 downgrade        : $o365E3Downgrade ← O365 E3 but web/mobile only, no desktop apps, mailbox < 50 GB
  E3 → Business Premium         : $e3Downgrade ← M365 E3 under 300-seat cap, mailbox < 50 GB, cheaper as Business Premium
  E5 voice → No-PSTN variant    : $e5VoiceWaste ← full E5 but 0 calls and 0 meetings organized, swap to no-audio-conf variant
  Apps Ent → Apps Business       : $appArbitrage ← Apps for Enterprise under 300-seat cap, identical to cheaper Apps for Business
  PBI PPU → PPU Add-On           : $ppuArbitrage ← standalone PPU on users who already get Pro from suite, swap to add-on
  OneDrive Plan 2 → Plan 1      : $odPlan2Waste ← standalone OneDrive Plan 2 (unlimited) but using < 900 GB (Plan 1 1 TB suffices)
  Entra P2 → P1 downgrade       : $entraP2Downgrade ← standalone Entra P2 but no admin roles, no PIM, no risk-based CA
  EXO Plan 1 → Kiosk            : $exoKioskDowngrade ← standalone Exchange Plan 1 but web-only email access and < 2 GB mailbox
  Std → Business Premium         : $bizPremInversion ← Business Standard + security/compliance add-ons exceed Business Premium price
  Biz Premium security review    : $bizPremSecReview ← Business Premium + Defender Suite for Business — verify advanced capabilities justify add-on
  Frontline rescue               : $frontlineRescue ← F3 blocked (archive) but web/mobile only — rescue to E1 or Business Basic
  Standalone apps waste        : $standaloneAppsWaste ← desktop app SKU but only web/mobile usage
  F3 to F1 downgrade           : $f3ToF1Downgrade ← F3 user with empty mailbox/OneDrive
  Frontline add-on bloat       : $frontlineAddonBloat ← F-series base + add-ons exceed Business Premium/E3 cost
  No desktop app users         : $noDesktopCount ← candidates for web-only / F-license
  Mobile-only app users        : $mobileOnly  ← candidates for frontline license
  No Outlook desktop users     : $noOutlookDesktopCount ← may not need Outlook desktop license
  Teams no-desktop users       : $teamsNoDesktopCount ← candidates for F-license
  E5 suite inversions           : $suiteInversion ← E3 + add-ons exceed E5 price — upgrade saves money
  E5 consolidation              : $e5Upgrade   ← E3 + add-ons at break-even — simplifies to 1 SKU
  Bundle consolidation          : $bundleConsolidation ← O365+EMS+Windows separately = cheaper as M365 bundle
  Teams unbundling              : $teamsUnbundling ← bundled suite with 0 Teams activity, switch to "Without Teams" SKU
  EXO Plan 2 downgrade         : $exoPlan2Review ← mailbox under 50 GB, Plan 1 may suffice

PRODUCT-SPECIFIC FLAGS:
  Shelfware (confirmed)       : $shelfware ← expensive license, no desktop/mobile app activity detected
  Shelfware (review)          : $shelfwareReview ← web-only activity detected, verify if standalone license is needed
  Teams Phone PSTN review      : $phoneNoPlan ← Teams Phone SKU, no Microsoft Calling Plan (may use Direct Routing/Operator Connect)
  Calling Plan waste            : $callingPlanWaste ← paid Calling Plan (MCOPSTN) but 0 Teams calls in report period
  AI add-on overlap (definitive): $aiAddonOverlap ← Teams Premium + Copilot, 0 meetings organized — remove Premium
  AI overlap review (soft)       : $aiOverlapReview ← Teams Premium + Copilot, has meetings — check webinar need
  Entra Suite overlap          : $entraSuiteOverlap ← standalone P2/Governance redundant under Entra Suite

COPILOT RECLAIM PIPELINE:
  Total Copilot license holders  : $($copilotUsers + $copilotPrereq)
  ├─ Properly licensed           : $copilotUsers
  ├─ KEEP (active usage)         : $copilotKeep ← Copilot activity detected across M365 apps
  ├─ WATCHLIST (at risk)         : $copilotWatchlist ← zero Copilot activity but active in M365 workloads (enablement candidate)
  ├─ RECLAIM (no readiness)      : $copilotReclaim ← zero Copilot AND zero workload activity (reclaim immediately)
  ├─ Prerequisite missing        : $copilotPrereq ← Copilot assigned without qualifying base license
  └─ Copilot Studio              : $copilotStudioUsers ← admin/developer tool, separate from productivity Copilot
  Reclaim savings                : €$($copilotReclaimCost.ToString('N2'))/yr  (immediate — no usage, no readiness)
  Watchlist at-risk spend        : €$($copilotWatchlistCost.ToString('N2'))/yr  (at-risk — needs enablement or reclaim)
$(if ($_hasCpcSku) {
@"

CLOUD PC UTILIZATION:
  Dormant Cloud PC (0 hrs/90d)  : $dormantCloudPc ← Cloud PC never connected — reclaim license
  Cloud PC Review (< 10 hrs)    : $cloudPcReview ← very low Cloud PC usage — consider downsizing or reclaiming
"@
})
  Power BI Pro (Premium Cap.)  : $pbiProReview ← consumers may use Free; creators/publishers still need Pro
  Mailbox storage warnings     : $mailboxStorageWarning ← Plan 1/Plan 2 mailbox approaching storage limit
  OneDrive storage warnings    : $oneDriveStorageWarning ← Business/E1 OneDrive approaching 1 TB limit
  Unlicensed with data         : $unlicensedWithData ← unlicensed user with mailbox/OneDrive data at risk of 30-day purge
  Seeded Visio overlap          : $seededVisioOverlap ← Visio Plan 1 redundant with built-in Visio web app in E3/E5
  E5 data hoarder               : $e5DataHoarder ← expensive license retained needlessly for litigation hold (free Inactive Mailbox)
  Inactive hold (cheap SKU)    : $inactiveHold ← cheaper license on held mailbox; remove and let Microsoft create free Inactive Mailbox
  Viral/exploratory cleanup    : $viralCleanup ← free trial SKUs alongside paid suites (provisioning conflict risk)
  High risk sharing             : $highRiskSharing ← heavy external file sharing without DLP/Purview coverage

LICENSING COMPLIANCE (evidence-driven):
  Licensing check flags        : $licensingCheck total
    Conditional Access (no P1) : $licensingCheckCA ← users in CA policy scope without Entra ID P1
    MDO (no entitlement)       : $licensingCheckMDO ← users in MDO policy scope without Defender for Office 365
    PIM (no P2)                : $licensingCheckPIM ← users with PIM roles without Entra ID P2
    Frontline desktop breach   : $licensingCheckFrontline ← F1/F3 users with desktop Office activations (audit risk)
  License assignment errors    : $licenseErrors ← group-based licensing failures (insufficient seats, conflicts)
$(if ($cloudLicensingLoaded) {
@"

CLOUD LICENSING HEALTH (beta API):
  Trial subscriptions          : $trialLicenseUsers user(s) on trial licenses
  License capacity queue       : $capacityQueueUsers user(s) waiting for available seats
  Assignment sync errors       : $($clAssignmentErrors.Count) user(s)/group(s) with sync failures
"@
})
DATA QUALITY:
  Users with missing sources  : $missingSourceUsers$(if ($missingSourceUsers -gt 0) { ' ← one or more report sources unavailable' } else { '' })
  Users with unknown SKUs     : $dataGapUsers$(if ($dataGapUsers -gt 0) { ' ← cost/right-size data may be incomplete' } else { '' })
  Frontline REVIEW (data gap) : $frontlineReview$(if ($frontlineReview -gt 0) { ' ← missing M365 app platform data' } else { '' })
  Business REVIEW (data gap)  : $businessReview$(if ($businessReview -gt 0) { ' ← missing M365 app platform data' } else { '' })

RUNTIME MAPPING VALIDATION (v$MappingVersion):
  Unmapped suites in tenant   : $($unmappedSuites.Count)$(if ($unmappedSuites.Count -gt 0) { " ← $($unmappedSuites[0..([math]::Min(2,$unmappedSuites.Count-1))] -join ', ')" } else { '' })
  Reference-only items (OK)   : $($unknownMappingItems.Count)$(if ($unknownMappingItems.Count -gt 0) { " ← mapping entries for SKUs not in this tenant (expected)" } else { '' })

SECURITY & COMPLIANCE UPSELL:
  No Defender protection       : $securityGap  ← Business Basic/Standard without any protection
  Defender Suite upsell        : $defenderUpsell ← Business Premium → add Defender Suite add-on
  Purview upsell               : $purviewUpsell ← no advanced compliance add-on detected

$expiryWarning
$businessLimitText
$(if ($groupLicenseInventory.Count -gt 0) {
"================================================================
GROUP-BASED LICENSING INVENTORY:
$($groupLicenseInventory | ForEach-Object {
    "  $($_.'Group Name') [$($_.'Membership Type'), $($_.'Member Count') members]: $($_.'Assigned Licenses')"
} | Out-String)"
})================================================================
RECOMMENDATION DISTRIBUTION:
$(@($recDistribution.GetEnumerator()) | Sort-Object { $_.Value.Count } -Descending | ForEach-Object { "  $($_.Key): $($_.Value.Count)" } | Out-String)
================================================================
FILES:
  [1] M365_LicenseOptimization_$ts.csv   — main per-user report with recommendations
  [2] M365_ServicePlanDetail_$ts.csv     — granular SKU/service-plan per user
  [3] M365_SkuInventory_$ts.csv          — tenant-level license inventory
  [4] This summary file
  [5] M365_ExecutiveSummary_$ts.csv      — executive financial summary (Estimated Optimization Potential)
  [6] M365_LicenseOptimization_$ts.xlsx  — Excel workbook (if ImportExcel installed)
  [7] M365_LicenseDelta_$ts.csv          — delta comparison (if -PriorReportPath specified)

PRICING DISCLAIMER:
  All cost figures in this report are INDICATIVE estimates based on publicly available
  Microsoft list prices (EUR). Actual costs may differ due to Enterprise Agreement (EA)
  pricing, volume discounts, CSP partner margins, regional variations, promotional rates,
  or contract-specific terms. Treat all financial figures as directional guidance for
  prioritization — not as exact billing amounts. Always verify against your actual
  Microsoft invoice or licensing agreement before making financial commitments.

NOTES:
  - Usage data has ~48h reporting latency.
  - Intensity thresholds are configurable via parameters. Current:
      Exchange: Low < $ExchangeLowThreshold, High > $ExchangeHighThreshold
      Teams:    Low < $TeamsLowThreshold, High > $TeamsHighThreshold
      OneDrive: Low < $OneDriveLowThreshold, High > $OneDriveHighThreshold
      SharePoint: Low < $SharePointLowThreshold, High > $SharePointHighThreshold
  - Dormant threshold: no interactive sign-in for $InactiveSignInDays days.
  - The "No Desktop Apps" flag means the user used M365 apps (Word/Excel/etc.)
    via web and/or mobile but NEVER on a Windows or Mac desktop client in the report period.
  - "No Outlook Desktop" = user accesses email via OWA and/or mobile Outlook but never via Outlook desktop client.
  - "Teams No Desktop" = user accesses Teams via web/mobile but not via desktop client.
  - Mailbox Size and OneDrive Storage help identify data that must be migrated before removal.
  - Sign-in activity uses the beta Graph API — requires AuditLog.Read.All permission.
  - Mailbox type detection requires the ExchangeOnlineManagement module (optional).
    Shared mailboxes under 50 GB do not need a license; over 50 GB they need one for
    auto-expanding archive. Room/Equipment mailboxes only need a Room license.
  - Admin role detection: admin accounts should only have Entra ID P1/P2,
    not full productivity suites. Consider using a separate daily-driver account for productivity.
  - License Assignment Path uses the beta API licenseAssignmentStates property.
    "Overlapping" means the same SKU is assigned both directly and via group — the
    direct assignment is redundant and should be removed.
  - Duplicate Coverage: if a user has a suite (e.g. M365 E3) AND a standalone SKU that
    the suite already includes (e.g. Exchange Online Plan 2), the standalone is redundant.
  - E5 Upgrade: if a user has E3 + E5-only add-ons (MDO P2, Defender for Endpoint P2,
    Defender for Identity, Defender for Cloud Apps, Teams Phone, Audio Conf, PBI Pro, Entra P2,
    Defender Suite/Purview Suite bundles), consolidating to E5 is often cheaper.
    Note: MDO P1 and MDE P1 are now included in E3 — they are flagged as duplicates, not E5 add-ons.
  - Shelfware: Visio, Project, Power BI Pro, and Teams Premium are expensive per-user licenses.
    Visio/Project use product-specific activation data (not generic Office app usage).
    Teams Premium uses Teams meeting activity. Power BI uses general app/SharePoint activity.
    Verify actual usage before renewal.
  - Teams Phone: a Teams Phone SKU without a Microsoft Calling Plan is flagged for review.
    Most enterprises use Direct Routing (SBC) or Operator Connect instead of Microsoft Calling
    Plans — these PSTN routes are not detectable via licensing data. Verify with the Teams
    administrator whether a PSTN route is configured before removing the Teams Phone license.
  - Copilot: flagged for adoption monitoring. Active users get an ROI note; inactive
    users are flagged for reallocation. Copilot requires a qualifying base license
    (E3/E5/Business Standard/Premium) — users without one are flagged as PREREQUISITE MISSING.
    Copilot Studio is tracked separately as an admin/developer tool (not productivity metrics).
  - Power BI Pro + Premium Capacity: when the tenant owns Power BI Premium Capacity,
    report consumers (viewers) can use the Free license. However, creators and publishers
    who build reports, publish to workspaces, or use dataflows STILL require Pro.
    Graph Reports cannot distinguish creator from consumer — manual role verification needed.
  - EXO Plan 2 vs Plan 1: Plan 2 provides 100 GB mailbox + In-Place Hold. If the
    user's mailbox is under 50 GB and usage is not high, Plan 1 (50 GB) may suffice.
  - Frontline (F1/F3): users on E3/E5/Business Premium who only use web and mobile
    apps (no desktop) are candidates for a cheaper frontline license.
  - Mailbox Storage Warning: Plan 1 mailboxes approaching 50 GB (90%+) and Plan 2/E3/E5
    primary mailboxes approaching 100 GB (90%+) risk mail flow stoppage. Auto-expanding
    archive only applies to the archive mailbox, NOT the primary mailbox.
  - OneDrive Storage Warning: Business and E1 plans cap OneDrive at 1 TB. When usage
    exceeds 93%, sync breaks. E3/E5 support up to 5 TB (expandable via admin request).
  - Unlicensed With Data: users without a license but with existing mailbox or OneDrive
    data. Microsoft purges this data after 30 days — back up or convert to shared mailbox.
  - Disabled Accounts: accounts with sign-in blocked but still holding a license are
    candidates for license removal. If the mailbox is on Litigation Hold,
    the license can still be removed — Microsoft creates a free Inactive Mailbox
    that retains all content and holds indefinitely.
    Disabled accounts with only free SKUs (€0 cost) are flagged for hygiene, not waste.
  - Forwarding Mailbox Waste: mailboxes with auto-forwarding configured (ForwardingAddress
    or ForwardingSmtpAddress) where the user is dormant or has no activity. These mailboxes
    exist only to relay mail and can be replaced by a free Mail Contact, shared mailbox,
    or Exchange transport rule — eliminating the license cost entirely.
  - Business 300-Seat Limit: Business-family SKUs (Business Basic/Standard/Premium)
    have a hard cap of 300 seats. When usage exceeds 85%, plan migration to Enterprise.
  - Subscription lifecycle data (expiry, status) is fetched from /v1.0/directory/subscriptions.
    SKUs in Warning/Suspended/LockedOut status are flagged in the summary.
  - Cloud Licensing API (beta) provides trial subscription detection, assignment error
    surfacing, and capacity queue (waiting member) tracking. Requires CloudLicensing.Read
    permission (CloudLicensing.Read.All). Gracefully skipped if not consented.
  - Cloud PC (Windows 365) usage data: fetched via beta API (CloudPC.Read.All permission).
    Shows total connected hours in 90 days, last active time, and device type per Cloud PC.
    Users with CPC/W365 licenses but zero connection time are flagged as DORMANT CLOUD PC.
    Users with very low usage (< 10 hours/90 days) are flagged as CLOUD PC REVIEW.
    Gracefully skipped if no CPC/W365 SKUs exist in the tenant or permission is not granted.
  - Plan capabilities matrix (sourced from Microsoft Modern Work Plan Comparison docs)
    drives enhanced right-sizing: frontline recommendations now show specific cost savings,
    mailbox/OneDrive compatibility notes, and target license recommendations.
  - Business Standard to Basic downgrade: users on Business Standard who only use web/mobile
    apps are flagged as candidates for downgrade to Business Basic with EUR savings.
  - Guest users (UserType=Guest) with paid licenses are flagged for review.
  - SKU friendly names are mapped from M365SkuData.json (if present).
    Unknown SKUs fall back to their raw SkuPartNumber value and are flagged with 'DATA GAP'.
  - SKU pricing is loaded from M365SkuPricing.csv (or -PricingCsvPath), then M365SkuData.json
    overrides on top. Unknown SKUs get EUR 0.00 and 'Pricing Known = False'
    in the SKU inventory. Use -ForceSkuRefresh to abort if external data is missing or stale.
  - If UPNs appear hashed, the tenant privacy setting may not have been updated.
    Re-run without -KeepHashedUPNs (default behavior unhides UPNs automatically).
$(if ($SkipEXO) {
@"

SKIP EXO MODE ACTIVE:
  The -SkipEXO switch was used. The following data is NOT available:
  - Mailbox type detection (Shared/Room/Equipment) — affects shared mailbox licensing recommendations
  - Litigation hold detection — affects license retention recommendations
  - MDO policy coverage (Safe Links/Attachments) — affects licensing compliance checks
  There is NO Microsoft Graph API equivalent for these Exchange Online-specific features.
"@
})
$(if ($script:skippedDataWarnings.Count -gt 0) {
@"

DATA COLLECTION WARNINGS ($($script:skippedDataWarnings.Count) issue(s)):
$( $script:skippedDataWarnings | ForEach-Object { "  [WARN] $_" } | Out-String)  Some columns or recommendations may be incomplete due to the above errors.
  Review the warnings and ensure the required Graph API permissions are granted.
"@
})
$(if ($manualRules -and $manualRules.Count -gt 0) {
    $checklist = "`nMANUAL AUDIT CHECKLIST ($($manualRules.Count) items):"
    $checklist += "`n  The following checks require portal access and cannot be fully automated."
    foreach ($rule in $manualRules) {
        $checklist += "`n"
        $checklist += "`n  [$($rule.id)] $($rule.title)"
        $checklist += "`n    Category: $($rule.category) | Confidence: $($rule.confidence.base)"
        if ($rule.manualCheck -and $rule.manualCheck.Count -gt 0) {
            foreach ($step in $rule.manualCheck) {
                $checklist += "`n    → $step"
            }
        }
        if ($rule.references -and $rule.references.Count -gt 0) {
            foreach ($refId in $rule.references) {
                if ($rulePackDocs.ContainsKey($refId)) {
                    $doc = $rulePackDocs[$refId]
                    $checklist += "`n    Doc: $($doc.title) — $($doc.url)"
                }
            }
        }
    }
    $checklist
})
"@

$summary | Out-File -FilePath $summaryFile -Encoding UTF8
Write-Host "  [4] Summary              : $summaryFile" -ForegroundColor Green

# ── Executive Financial Summary CSV ──
$execSummaryFile = Join-Path $OutputFolder "M365_ExecutiveSummary_$ts.csv"
$execRows = [System.Collections.Generic.List[PSCustomObject]]::new()
$execRows.Add([PSCustomObject]@{ Tier = "Overview"; Category = "Total Annual M365 Spend";      Users = $totalUsers;           'Annual Amount (EUR)' = $totalAnnualSpend;     'Pct of Spend' = "100.0%" })
$execRows.Add([PSCustomObject]@{ Tier = "Overview"; Category = "Estimated Optimization Potential";      Users = "";                    'Annual Amount (EUR)' = $totalMoneyOnTable;    'Pct of Spend' = "$wastePercentage%" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 1";   Category = "Dormant Accounts (no sign-in >$InactiveSignInDays days)"; Users = $dormantTier1Count; 'Annual Amount (EUR)' = $dormantCost;          'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 1";   Category = "Disabled Accounts (sign-in blocked)"; Users = ($disabledLicensed - $disabledFreeSku); 'Annual Amount (EUR)' = $disabledCost;       'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 1";   Category = "Zero M365 Usage (no app activity in period)"; Users = $noActivity; 'Annual Amount (EUR)' = $noActivityCost;    'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 1";   Category = "Unused Premium Add-Ons (Visio/Project/PBI Pro)"; Users = $shelfware; 'Annual Amount (EUR)' = $shelfwareCost;  'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 1";   Category = "Copilot Reclaim (zero usage & zero readiness)"; Users = $copilotReclaim; 'Annual Amount (EUR)' = $copilotReclaimCost; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 1";   Category = "Copilot At Risk (zero usage, active in M365)"; Users = $copilotWatchlist; 'Annual Amount (EUR)' = $copilotWatchlistCost; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 1";   Category = "Shared Mailbox (no license needed under 50 GB)"; Users = $sharedMbxRemovable; 'Annual Amount (EUR)' = $sharedMbxCost; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 1";   Category = "Duplicate Licenses (standalone included in suite)"; Users = $duplicateCov; 'Annual Amount (EUR)' = $duplicateCost; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 1";   Category = "TIER 1 SUBTOTAL";              Users = "";                    'Annual Amount (EUR)' = $tier1Waste;           'Pct of Spend' = "$tier1Percentage%" })
foreach ($uLic in $unassignedLicenseInventory) {
    $execRows.Add([PSCustomObject]@{ Tier = "Tier 1"; Category = "Unassigned Licenses: $($uLic.FriendlyName) ($($uLic.Unassigned)/$($uLic.Total))"; Users = $uLic.Unassigned; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
}
$execRows.Add([PSCustomObject]@{ Tier = "Tier 2";   Category = "E3/E5 to Frontline F1/F3 (web/mobile only users)"; Users = $frontlineCandidate; 'Annual Amount (EUR)' = $frontlineSavings; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 2";   Category = "Business Standard to Basic (no desktop apps used)"; Users = $businessDowngrade; 'Annual Amount (EUR)' = $businessBasicSavings; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 2";   Category = "Exchange Plan 2 to Plan 1 (mailbox under 50 GB)"; Users = $exoPlan2Review; 'Annual Amount (EUR)' = $exoPlan2Savings; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 2";   Category = "E3 + Add-Ons to E5 Upgrade (cheaper as E5)"; Users = ($e5Upgrade + $suiteInversion); 'Annual Amount (EUR)' = $e5UpgradeSavings; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 2";   Category = "O365+EMS+Windows to M365 Bundle (cheaper combined)"; Users = $bundleConsolidation; 'Annual Amount (EUR)' = $bundleConsolidationSavings; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 2";   Category = "O365 E1 to Business Basic (same features, lower cost)"; Users = $e1Downgrade; 'Annual Amount (EUR)' = $e1DowngradeSavings; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 2";   Category = "O365 E3 to E1 (web/mobile only, mailbox <50 GB)"; Users = $o365E3Downgrade; 'Annual Amount (EUR)' = $o365E3DowngradeSavings; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 2";   Category = "M365 E3 to Business Premium (<300 seats, cheaper)"; Users = $e3Downgrade; 'Annual Amount (EUR)' = $e3DowngradeSavings; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 2";   Category = "E5 to No Audio Conferencing Variant (0 calls)"; Users = $e5VoiceWaste; 'Annual Amount (EUR)' = $e5VoiceSavings; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 2";   Category = "Apps Enterprise to Apps Business (<300 seats)"; Users = $appArbitrage; 'Annual Amount (EUR)' = $appArbitrageSavings; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 2";   Category = "PBI PPU Standalone to Add-On (Pro from suite)"; Users = $ppuArbitrage; 'Annual Amount (EUR)' = $ppuArbitrageSavings; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 2";   Category = "Exchange Plan 1 to Kiosk (web-only, <2 GB)"; Users = $exoKioskDowngrade; 'Annual Amount (EUR)' = $exoKioskSavings; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 2";   Category = "Biz Standard + Add-Ons to Premium (cheaper)"; Users = $bizPremInversion; 'Annual Amount (EUR)' = $bizPremInversionSavings; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 2";   Category = "F-License Blocked, E1/Basic Alternative"; Users = $frontlineRescue; 'Annual Amount (EUR)' = $frontlineRescueSavings; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tier 2";   Category = "TIER 2 SUBTOTAL";              Users = "";                    'Annual Amount (EUR)' = $tier2Savings;         'Pct of Spend' = "$tier2Percentage%" })
foreach ($poolWarn in $unassignedPoolWarnings) {
    $poolPctStr = if ($totalAnnualSpend -gt 0) { "$([math]::Round($poolWarn.AnnualWaste / $totalAnnualSpend * 100, 1))%" } else { "" }
    $execRows.Add([PSCustomObject]@{ Tier = "Pool"; Category = "$($poolWarn.FriendlyName) ($($poolWarn.Unassigned)/$($poolWarn.Total) unassigned)"; Users = $poolWarn.Unassigned; 'Annual Amount (EUR)' = $poolWarn.AnnualWaste; 'Pct of Spend' = $poolPctStr })
}
if ($unassignedPoolWarnings.Count -gt 0) {
    $execRows.Add([PSCustomObject]@{ Tier = "Pool"; Category = "POOL WASTE TOTAL"; Users = ""; 'Annual Amount (EUR)' = [math]::Round($unassignedPoolTotalAnnual, 2); 'Pct of Spend' = "" })
}
if ($teamsRoomsDowngrade) {
    $execRows.Add([PSCustomObject]@{ Tier = "Tenant"; Category = "Teams Rooms Pro to Basic ($($teamsRoomsDowngrade.EligibleRooms) of $($teamsRoomsDowngrade.ProRooms) rooms, $($teamsRoomsDowngrade.BasicRooms)+$($teamsRoomsDowngrade.EligibleRooms)/25 cap)"; Users = $teamsRoomsDowngrade.EligibleRooms; 'Annual Amount (EUR)' = $teamsRoomsDowngrade.AnnualSavings; 'Pct of Spend' = "" })
}
# ── Tenant-Level Optimization ──
$execRows.Add([PSCustomObject]@{ Tier = "Tenant"; Category = "Overlapping Assignments (same SKU via direct + group)"; Users = $overlapping; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tenant"; Category = "Intune Entitlement Unused (0 enrolled devices)"; Users = $intuneShelfware; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tenant"; Category = "Intune Suite Add-On Redundant (included in E3/E5)"; Users = $intuneSuiteWaste; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tenant"; Category = "MDM/MAM Unused (web-only users, no devices)"; Users = $mdmMamWaste; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tenant"; Category = "Windows License on Non-Windows Users"; Users = $windowsLicenseWaste; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tenant"; Category = "Archive Add-On Unused (no archive, small mailbox)"; Users = $overLicensedArchive; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tenant"; Category = "Visio Plan 1 Redundant (web Visio in E3/E5)"; Users = $seededVisioOverlap; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tenant"; Category = "Guest Users with Paid Licenses"; Users = $guestAccountWaste; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tenant"; Category = "Shared/Room Mailboxes on Premium Suites"; Users = $nonHumanWaste; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tenant"; Category = "Self-Service & Trial Licenses (cleanup)"; Users = ($viralCleanup + $trialLicenseUsers); 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tenant"; Category = "Teams Unused (switch to Without Teams SKU)"; Users = $teamsUnbundling; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Tenant"; Category = "Archive Add-On Redundant (suite includes archive)"; Users = $redundantArchive; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
# ── Product-Specific Flags ──
$execRows.Add([PSCustomObject]@{ Tier = "Product"; Category = "Teams Phone Without Calling Plan (verify PSTN route)"; Users = $phoneNoPlan; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Product"; Category = "Teams Phone to Resource Account"; Users = $teamsPhoneRightSizing; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Product"; Category = "Calling Plan Unused (0 calls in period)"; Users = $callingPlanWaste; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Product"; Category = "Teams Premium + Copilot Overlap (remove Premium)"; Users = $aiAddonOverlap; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Product"; Category = "Teams Premium + Copilot (review webinar need)"; Users = $aiOverlapReview; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Product"; Category = "Power BI Pro Low Usage (review if needed)"; Users = $pbiProReview; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Product"; Category = "OneDrive Plan 2 to Plan 1 (using <900 GB)"; Users = $odPlan2Waste; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Product"; Category = "Entra P2 to P1 (no admin roles, no PIM, no risk CA)"; Users = $entraP2Downgrade; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Product"; Category = "Desktop App License Unused (web/mobile only)"; Users = $standaloneAppsWaste; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Product"; Category = "Premium Add-On Unused (no activity detected)"; Users = $premiumAddonWaste; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Product"; Category = "Standalone License Replaceable by Cheaper SKU"; Users = $alaCarteWaste; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Product"; Category = "Separate SKUs Cheaper Than Current Bundle"; Users = $bundleInefficiency; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Product"; Category = "Unused Premium Add-Ons Review (web-only activity)"; Users = $shelfwareReview; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Product"; Category = "F3 to F1 (empty mailbox & OneDrive)"; Users = $f3ToF1Downgrade; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Product"; Category = "Frontline + Add-Ons Exceed E3/Premium Price"; Users = $frontlineAddonBloat; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
# ── Copilot Adoption Pipeline ──
$execRows.Add([PSCustomObject]@{ Tier = "Copilot"; Category = "Total Copilot Holders";             Users = ($copilotUsers + $copilotPrereq); 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Copilot"; Category = "Active Users (Copilot usage detected)"; Users = $copilotKeep; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Copilot"; Category = "Missing Base License (needs E3/E5/Biz Std/Prem)"; Users = $copilotPrereq; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Copilot"; Category = "Watchlist (zero Copilot, active M365)"; Users = $copilotWatchlist; 'Annual Amount (EUR)' = $copilotWatchlistCost; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Copilot"; Category = "Reclaim (zero Copilot, zero M365)";     Users = $copilotReclaim;    'Annual Amount (EUR)' = $copilotReclaimCost;   'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Copilot"; Category = "Copilot Studio";                    Users = $copilotStudioUsers;    'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
# ── Cloud PC Utilization ──
if ($_hasCpcSku) {
    $execRows.Add([PSCustomObject]@{ Tier = "Cloud PC"; Category = "Dormant Cloud PC (0 hours in 90 days)"; Users = $dormantCloudPc; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
    $execRows.Add([PSCustomObject]@{ Tier = "Cloud PC"; Category = "Cloud PC Review (< 10 hours in 90 days)"; Users = $cloudPcReview; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
}
# ── Operational Review ──
$execRows.Add([PSCustomObject]@{ Tier = "Risk";   Category = "Dormant Admin Accounts";             Users = $dormantAdminRisk;      'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Risk";   Category = "Automation Accounts";                Users = $automationAccount;     'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Risk";   Category = "Legacy Service Accounts";            Users = $legacyServiceAccount;  'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Risk";   Category = "Premium License as Cold Storage (0 activity + data)"; Users = $expensiveColdStorage; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Risk";   Category = "Unlicensed User With Data (30-day purge risk)"; Users = $unlicensedWithData; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Risk";   Category = "Heavy External Sharing Without DLP/Purview"; Users = $highRiskSharing; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Risk";   Category = "Forwarding-Only Mailbox (replace with Mail Contact)"; Users = $forwardingWaste; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Risk";   Category = "Active User With Mail Forwarding (low usage)"; Users = $forwardingReview; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Risk";   Category = "Mailbox Storage Warning";            Users = $mailboxStorageWarning;  'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Risk";   Category = "OneDrive Storage Warning";           Users = $oneDriveStorageWarning; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
# ── Administrative & Compliance ──
$execRows.Add([PSCustomObject]@{ Tier = "Admin";  Category = "Licensing Compliance Gap (Total)";              Users = $licensingCheck;        'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Admin";  Category = "Conditional Access Without Entra P1 License"; Users = $licensingCheckCA; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Admin";  Category = "Defender for Office Policy Without License"; Users = $licensingCheckMDO; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Admin";  Category = "PIM Role Assignment Without Entra P2 License"; Users = $licensingCheckPIM; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Admin";  Category = "License Assignment Errors";          Users = $licenseErrors;         'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Admin";  Category = "Entra Suite + Standalone P2/Governance (redundant)"; Users = $entraSuiteOverlap; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Admin";  Category = "Missing Base License (needs E3/E5/Biz Std/Prem)"; Users = $copilotPrereq; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
# ── Data Quality ──
$execRows.Add([PSCustomObject]@{ Tier = "Quality"; Category = "License Capacity Queue";              Users = $capacityQueueUsers;    'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Quality"; Category = "Users with Data Gaps";                Users = $dataGapUsers;          'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Quality"; Category = "Users with Missing Sources";          Users = $missingSourceUsers;    'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Quality"; Category = "Frontline Review (Data Gap)";          Users = $frontlineReview;       'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Quality"; Category = "Business Review (Data Gap)";           Users = $businessReview;        'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
# ── Security & Compliance Coverage ──
$execRows.Add([PSCustomObject]@{ Tier = "Security"; Category = "No Defender Protection (Business Basic/Standard)"; Users = $securityGap; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Security"; Category = "Defender Suite Upsell";            Users = $defenderUpsell;        'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Security"; Category = "Purview Upsell";                   Users = $purviewUpsell;         'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Security"; Category = "Business Premium Security Review"; Users = $bizPremSecReview;      'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
# ── Security Posture Distribution ──
$execRows.Add([PSCustomObject]@{ Tier = "Posture"; Category = "Security: None";                    Users = $secCoverageNone;       'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Posture"; Category = "Security: Basic";                   Users = $secCoverageBasic;      'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Posture"; Category = "Security: Advanced";                Users = $secCoverageAdvanced;   'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Posture"; Category = "Security: E5-equivalent";           Users = $secCoverageE5;         'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Posture"; Category = "Compliance: None";                  Users = $compCoverageNone;      'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Posture"; Category = "Compliance: Basic";                 Users = $compCoverageBasic;     'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Posture"; Category = "Compliance: Advanced";              Users = $compCoverageAdvanced;  'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows.Add([PSCustomObject]@{ Tier = "Posture"; Category = "Compliance: E5-equivalent";         Users = $compCoverageE5;        'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
# Subscription lifecycle warnings (for heatmap consumption)
foreach ($sku in $subscribedSkus) {
    if ($sku.AppliesTo -eq 'Company') { continue }   # skip tenant-level capacity SKUs (Dataverse, etc.)
    if ($_knownFreeRx.IsMatch($sku.SkuPartNumber)) { continue }   # skip free/trial SKUs (no cost impact)
    if ($lkpSubscription.ContainsKey($sku.SkuPartNumber) -and $lkpSubscription[$sku.SkuPartNumber].NextLifecycle) {
        $daysLeft = [int]([datetime]$lkpSubscription[$sku.SkuPartNumber].NextLifecycle - (Get-Date)).TotalDays
        $status   = $lkpSubscription[$sku.SkuPartNumber].Status
        if ($daysLeft -le 90 -or $status -ne "Enabled") {
            $subFriendly = Resolve-SkuFriendlyName $sku.SkuPartNumber
            $subSeats    = $sku.PrepaidUnits.Enabled + $sku.PrepaidUnits.Warning
            $subConsumed = $sku.ConsumedUnits
            $execRows.Add([PSCustomObject]@{ Tier = "Sub"; Category = "$subFriendly"; Users = "$subConsumed/$subSeats"; 'Annual Amount (EUR)' = "$status"; 'Pct of Spend' = "${daysLeft}d" })
        }
    }
}
$execRows.Add([PSCustomObject]@{ Tier = "";         Category = "DISCLAIMER: All cost figures are indicative estimates based on public Microsoft list prices (EUR). Actual costs may differ due to EA/CSP/volume pricing Copilot usage and Cloud PC analytics rely on Microsoft Graph BETA APIs — these sections may show limited results until the API becomes generally available."; Users = ""; 'Annual Amount (EUR)' = ""; 'Pct of Spend' = "" })
$execRows | Export-Csv -Path $execSummaryFile -NoTypeInformation -Encoding UTF8
Write-Host "  [5] Executive Summary    : $execSummaryFile" -ForegroundColor Green

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 7b — Delta Report (prior vs current comparison)
# ═══════════════════════════════════════════════════════════════════════════════

$deltaFile = $null
$deltaNewUsers = 0; $deltaRemovedUsers = 0; $deltaLicenseChanged = 0
[decimal]$deltaCostIncrease = 0; [decimal]$deltaCostDecrease = 0
$deltaBecameDormant = 0; $deltaBecameActive = 0; $deltaRecChanged = 0
$deltaNewCopilot = 0; $deltaInactiveCopilot = 0
[decimal]$deltaTotalPriorCost = 0; [decimal]$deltaTotalCurrentCost = 0; [decimal]$deltaWasteAddressed = 0

if ($PriorReportPath) {
    Write-Host "`n[*] Generating delta analysis against prior report ..." -ForegroundColor Cyan

    try {
        $priorRows = @(Import-Csv -Path $PriorReportPath)
        Write-Host "  Prior report: $($priorRows.Count) users loaded from $PriorReportPath" -ForegroundColor Green

        # Build prior lookup by UPN
        $priorByUpn = @{}
        foreach ($pr in $priorRows) {
            $prUpn = $pr.'User Principal Name'
            if ($prUpn) { $priorByUpn[$prUpn] = $pr }
        }

        # Load current report (just written to disk via StreamWriter)
        $currentRows = @(Import-Csv -Path $mainFile)
        $currentByUpn = @{}
        foreach ($cr in $currentRows) {
            $crUpn = $cr.'User Principal Name'
            if ($crUpn) { $currentByUpn[$crUpn] = $cr }
        }

        # Delta CSV columns
        $deltaColumns = @(
            'User Principal Name', 'Display Name', 'Change Type',
            'Prior Licenses', 'Current Licenses', 'License Changed',
            'Prior Annual Cost (EUR)', 'Current Annual Cost (EUR)', 'Cost Delta (EUR)',
            'Prior Recommendation Category', 'Current Recommendation Category', 'Recommendation Changed',
            'Prior Dormant', 'Current Dormant', 'Became Dormant', 'Became Active',
            'Prior Exchange Intensity', 'Current Exchange Intensity',
            'Prior Teams Intensity', 'Current Teams Intensity',
            'Prior OneDrive Intensity', 'Current OneDrive Intensity',
            'Prior SharePoint Intensity', 'Current SharePoint Intensity',
            'Prior Account Enabled', 'Current Account Enabled'
        )

        $deltaFile = Join-Path $OutputFolder "M365_LicenseDelta_$ts.csv"
        $deltaWriter = [System.IO.StreamWriter]::new($deltaFile, $false, [System.Text.UTF8Encoding]::new($false))
        $deltaWriter.WriteLine(($deltaColumns | ForEach-Object { '"' + $_ + '"' }) -join ',')

        # Union all UPNs from both runs
        $allDeltaUpns = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($k in $priorByUpn.Keys)   { [void]$allDeltaUpns.Add($k) }
        foreach ($k in $currentByUpn.Keys) { [void]$allDeltaUpns.Add($k) }

        # Safe column access helper (handles missing columns gracefully across CSV versions)
        $getSafe = {
            param([PSCustomObject]$Row, [string]$Col)
            if ($null -eq $Row) { return "" }
            if ($Col -in $Row.PSObject.Properties.Name) { return $Row.$Col } else { return "" }
        }

        $wasteCategories = @("Disabled Account","Dormant","No Activity","Shelfware",
                              "Shared Mailbox","Overlapping License","Duplicate Coverage")

        foreach ($dupn in $allDeltaUpns) {
            $prior   = if ($priorByUpn.ContainsKey($dupn))   { $priorByUpn[$dupn] }   else { $null }
            $current = if ($currentByUpn.ContainsKey($dupn)) { $currentByUpn[$dupn] } else { $null }

            # Change type
            $changeType = if     (-not $prior)   { "New" }
                          elseif (-not $current)  { "Removed" }
                          else                    { "Existing" }

            if ($changeType -eq "New")     { $deltaNewUsers++ }
            if ($changeType -eq "Removed") { $deltaRemovedUsers++ }

            # License comparison
            $priorLic   = & $getSafe $prior   'Assigned Licenses'
            $currentLic = & $getSafe $current 'Assigned Licenses'
            $licChanged = ($priorLic -ne $currentLic)
            if ($licChanged -and $changeType -eq "Existing") { $deltaLicenseChanged++ }

            # Cost comparison
            $priorCostStr   = & $getSafe $prior   'Annual License Cost (EUR)'
            $currentCostStr = & $getSafe $current 'Annual License Cost (EUR)'
            $priorCost   = if ($priorCostStr)   { try { [decimal]$priorCostStr }   catch { [decimal]0 } } else { [decimal]0 }
            $currentCost = if ($currentCostStr) { try { [decimal]$currentCostStr } catch { [decimal]0 } } else { [decimal]0 }
            $costDelta   = [math]::Round($currentCost - $priorCost, 2)

            $deltaTotalPriorCost   += $priorCost
            $deltaTotalCurrentCost += $currentCost
            if ($costDelta -gt 0) { $deltaCostIncrease += $costDelta }
            if ($costDelta -lt 0) { $deltaCostDecrease += $costDelta }

            # Recommendation comparison
            $priorRecCat   = & $getSafe $prior   'Recommendation Category'
            $currentRecCat = & $getSafe $current 'Recommendation Category'
            $recChanged    = ($priorRecCat -ne $currentRecCat)
            if ($recChanged -and $changeType -eq "Existing") { $deltaRecChanged++ }

            # Waste addressed: prior had a waste category, current does not
            if ($priorRecCat -in $wasteCategories -and $currentRecCat -notin $wasteCategories) {
                $deltaWasteAddressed += $priorCost
            }

            # Dormancy shifts
            $priorDormant   = & $getSafe $prior   'Dormant Account'
            $currentDormant = & $getSafe $current 'Dormant Account'
            $becameDormant  = ($priorDormant -ne 'True' -and $currentDormant -eq 'True')
            $becameActive   = ($priorDormant -eq 'True' -and $currentDormant -ne 'True' -and $changeType -ne "Removed")
            if ($becameDormant) { $deltaBecameDormant++ }
            if ($becameActive)  { $deltaBecameActive++ }

            # Copilot adoption tracking
            $priorHasCopilot   = ($priorLic -match 'COPILOT|Microsoft_365_Copilot')
            $currentHasCopilot = ($currentLic -match 'COPILOT|Microsoft_365_Copilot')
            if ($currentHasCopilot -and -not $priorHasCopilot) { $deltaNewCopilot++ }
            $currentRec = & $getSafe $current 'Recommendation'
            if ($currentHasCopilot -and $currentRec -match 'COPILOT RECLAIM|COPILOT WATCHLIST') {
                $deltaInactiveCopilot++
            }

            # Write delta row
            $deltaRow = [PSCustomObject]@{
                'User Principal Name'            = $dupn
                'Display Name'                   = if ($current) { & $getSafe $current 'Display Name' } else { & $getSafe $prior 'Display Name' }
                'Change Type'                    = $changeType
                'Prior Licenses'                 = $priorLic
                'Current Licenses'               = $currentLic
                'License Changed'                = $licChanged
                'Prior Annual Cost (EUR)'        = $priorCost
                'Current Annual Cost (EUR)'      = $currentCost
                'Cost Delta (EUR)'               = $costDelta
                'Prior Recommendation Category'  = $priorRecCat
                'Current Recommendation Category'= $currentRecCat
                'Recommendation Changed'         = $recChanged
                'Prior Dormant'                  = $priorDormant
                'Current Dormant'                = $currentDormant
                'Became Dormant'                 = $becameDormant
                'Became Active'                  = $becameActive
                'Prior Exchange Intensity'       = & $getSafe $prior   'Exchange Intensity'
                'Current Exchange Intensity'     = & $getSafe $current 'Exchange Intensity'
                'Prior Teams Intensity'          = & $getSafe $prior   'Teams Intensity'
                'Current Teams Intensity'        = & $getSafe $current 'Teams Intensity'
                'Prior OneDrive Intensity'       = & $getSafe $prior   'OneDrive Intensity'
                'Current OneDrive Intensity'     = & $getSafe $current 'OneDrive Intensity'
                'Prior SharePoint Intensity'     = & $getSafe $prior   'SharePoint Intensity'
                'Current SharePoint Intensity'   = & $getSafe $current 'SharePoint Intensity'
                'Prior Account Enabled'          = & $getSafe $prior   'Account Enabled'
                'Current Account Enabled'        = & $getSafe $current 'Account Enabled'
            }
            $deltaWriter.WriteLine((ConvertTo-CsvLine -Row $deltaRow -Columns $deltaColumns))
        }

        try { $deltaWriter.Flush(); $deltaWriter.Close(); $deltaWriter.Dispose() } catch {}
        $deltaWriter = $null

        Write-Host "  [7] Delta Report         : $deltaFile ($($allDeltaUpns.Count) comparison rows)" -ForegroundColor Green

        # Delta summary appended to the main summary TXT
        $deltaSummary = @"

================================================================
DELTA ANALYSIS (vs prior report)
Prior report : $PriorReportPath
Current report: $mainFile
================================================================

USER CHANGES:
  New users (in current, not prior)     : $deltaNewUsers
  Removed users (in prior, not current) : $deltaRemovedUsers
  Existing users with license changes   : $deltaLicenseChanged

COST TREND:
  Prior total annual spend              : EUR $([math]::Round($deltaTotalPriorCost, 2).ToString('N2'))
  Current total annual spend            : EUR $([math]::Round($deltaTotalCurrentCost, 2).ToString('N2'))
  Net cost change                       : EUR $([math]::Round($deltaTotalCurrentCost - $deltaTotalPriorCost, 2).ToString('N2'))
  Cost increases (sum of positive)      : EUR $([math]::Round($deltaCostIncrease, 2).ToString('N2'))
  Cost decreases (sum of negative)      : EUR $([math]::Round([math]::Abs($deltaCostDecrease), 2).ToString('N2'))

RECOMMENDATION CHANGES:
  Users with changed recommendation     : $deltaRecChanged
  Waste addressed (prior waste -> OK)   : EUR $([math]::Round($deltaWasteAddressed, 2).ToString('N2'))/yr

DORMANCY:
  Became dormant since prior            : $deltaBecameDormant
  Became active since prior             : $deltaBecameActive

COPILOT ADOPTION:
  New Copilot license holders           : $deltaNewCopilot
  Inactive Copilot users (current)      : $deltaInactiveCopilot

FILES:
  Delta CSV: M365_LicenseDelta_$ts.csv
"@

        Add-Content -Path $summaryFile -Value $deltaSummary -Encoding UTF8
        Write-Host "  Delta summary appended to: $summaryFile" -ForegroundColor Green

        # Free delta memory (two full CSVs worth of rows)
        $priorRows = $null; $currentRows = $null
        $priorByUpn = $null; $currentByUpn = $null
        [System.GC]::Collect(2, [System.GCCollectionMode]::Optimized)

    } catch {
        Write-Log "Delta analysis failed" -Level ERROR -ErrorRecord $_
        Write-Warning "  Delta analysis failed: $($_.Exception.Message)"
        [void]$script:skippedDataWarnings.Add("Delta analysis — $($_.Exception.Message)")
    } finally {
        if ($deltaWriter) { try { $deltaWriter.Flush(); $deltaWriter.Close(); $deltaWriter.Dispose() } catch {} }
    }
}

# 5. Excel workbook (conditional on ImportExcel)
if ($importExcelAvailable) {
    Write-Host "`n[12/12] Generating Excel workbook ..." -ForegroundColor Cyan
    Write-Log "[12/12] Generating Excel workbook"
    $xlFile = Join-Path $OutputFolder "M365_LicenseOptimization_$ts.xlsx"
    $currentExcelSheet = "init"
  try {

    # ── Sheet 2: User Report (primary data sheet — exported first as pivot source) ──
    $currentExcelSheet = "User Report"
    Import-Csv $mainFile | Export-Excel -Path $xlFile -WorksheetName "User Report" `
        -TableName "UserReport" -TableStyle Medium6 `
        -FreezeTopRow -AutoFilter -AutoSize -PassThru | ForEach-Object {
        $ws = $_.Workbook.Worksheets["User Report"]
        $ws.TabColor = [System.Drawing.Color]::FromArgb(68, 114, 196)  # Dark blue

        # Find column indexes for conditional formatting
        [int]$headerRow = $ws.Dimension.Start.Row
        [int]$lastCol   = $ws.Dimension.End.Column
        $colMap    = @{}
        for ($c = 1; $c -le $lastCol; $c++) {
            $colName = $ws.Cells[$headerRow, $c].Text
            if ($colName) { $colMap[$colName] = $c }
        }
        [int]$lastRow   = $ws.Dimension.End.Row
        [int]$dataStart = $headerRow + 1

        # Guard: skip conditional formatting if no data rows (prevents ExcelAddress crash)
        if ($lastRow -lt $dataStart) {
            Write-Log "XLSX User Report: no data rows — skipping conditional formatting" -Level WARN
        } else {

        # Blue data bars on Annual License Cost column
        if ($colMap.ContainsKey('Annual License Cost (EUR)')) {
            $costCol = $colMap['Annual License Cost (EUR)']
            $costAddr = [OfficeOpenXml.ExcelAddress]::new($dataStart, $costCol, $lastRow, $costCol)
            [void]$ws.ConditionalFormatting.AddDatabar($costAddr, [System.Drawing.Color]::SteelBlue)
        }

        # Red fill on Dormant Account = TRUE
        if ($colMap.ContainsKey('Dormant Account')) {
            $dormantCol  = $colMap['Dormant Account']
            $dormantAddr = [OfficeOpenXml.ExcelAddress]::new($dataStart, $dormantCol, $lastRow, $dormantCol)
            $cfDormant   = $ws.ConditionalFormatting.AddEqual($dormantAddr)
            $cfDormant.Formula = "TRUE"
            $cfDormant.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::FromArgb(255, 199, 206)
        }

        # RAG on intensity columns (Low=red, Medium=yellow, High=green)
        foreach ($intCol in @('Exchange Intensity','Teams Intensity','OneDrive Intensity','SharePoint Intensity')) {
            if ($colMap.ContainsKey($intCol)) {
                $ci   = $colMap[$intCol]
                $addr = [OfficeOpenXml.ExcelAddress]::new($dataStart, $ci, $lastRow, $ci)
                $cfLow = $ws.ConditionalFormatting.AddEqual($addr)
                $cfLow.Formula = '"Low"'
                $cfLow.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::FromArgb(255, 199, 206)
                $cfMed = $ws.ConditionalFormatting.AddEqual($addr)
                $cfMed.Formula = '"Medium"'
                $cfMed.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::FromArgb(255, 235, 156)
                $cfHigh = $ws.ConditionalFormatting.AddEqual($addr)
                $cfHigh.Formula = '"High"'
                $cfHigh.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::FromArgb(198, 239, 206)
            }
        }

        # Yellow highlight on non-OK recommendations
        if ($colMap.ContainsKey('Recommendation Category')) {
            $recCatCol  = $colMap['Recommendation Category']
            $recCatAddr = [OfficeOpenXml.ExcelAddress]::new($dataStart, $recCatCol, $lastRow, $recCatCol)
            $cfNotOK    = $ws.ConditionalFormatting.AddExpression($recCatAddr)
            $colLetter  = [OfficeOpenXml.ExcelCellAddress]::new($dataStart, $recCatCol).Address -replace '\d+',''
            $cfNotOK.Formula = "${colLetter}$($dataStart)<>`"OK`""
            $cfNotOK.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::FromArgb(255, 235, 156)
        }

        } # end if ($lastRow -ge $dataStart) — conditional formatting guard

        # Currency format on cost columns
        foreach ($cName in @('Monthly License Cost (EUR)','Annual License Cost (EUR)')) {
            if ($colMap.ContainsKey($cName)) {
                $ci = $colMap[$cName]
                $ws.Column($ci).Style.Numberformat.Format = '€#,##0.00'
            }
        }

        $_.Save()
        $_.Dispose()
    }

    # ── Category overview tabs (filtered subsets of User Report) ──
    $categoryTabs = @(
        @{ Name = "Admin Review";       Pattern = "ADMIN REVIEW";       Color = [System.Drawing.Color]::FromArgb(180, 198, 231) }
        @{ Name = "Dormant";            Pattern = "DORMANT";            Color = [System.Drawing.Color]::FromArgb(255, 199, 206) }
        @{ Name = "Disabled Account";   Pattern = "DISABLED ACCOUNT";   Color = [System.Drawing.Color]::FromArgb(255, 199, 206) }
        @{ Name = "Guest Account Review";Pattern = "GUEST ACCOUNT";      Color = [System.Drawing.Color]::FromArgb(255, 235, 156) }
        @{ Name = "Shared Mailbox";     Pattern = "SHARED MAILBOX";     Color = [System.Drawing.Color]::FromArgb(198, 239, 206) }
        @{ Name = "Automation Account"; Pattern = "AUTOMATION ACCOUNT"; Color = [System.Drawing.Color]::FromArgb(217, 217, 217) }
        @{ Name = "Never Signed In";    Pattern = "NEVER SIGNED IN";    Color = [System.Drawing.Color]::FromArgb(255, 199, 206) }
    )
    $allCsvRows = @(Import-Csv $mainFile)
    foreach ($catTab in $categoryTabs) {
        $currentExcelSheet = $catTab.Name
        $filtered = @($allCsvRows | Where-Object { $_.'Recommendation' -match $catTab.Pattern })
        if ($filtered.Count -eq 0) { continue }
        try {
            $filtered | Export-Excel -Path $xlFile -WorksheetName $catTab.Name `
                -TableName ($catTab.Name -replace '[^A-Za-z0-9]','') -TableStyle Medium6 `
                -FreezeTopRow -AutoFilter -AutoSize -PassThru | ForEach-Object {
                $ws = $_.Workbook.Worksheets[$catTab.Name]
                $ws.TabColor = $catTab.Color
                # Currency format on cost columns
                [int]$hRow = $ws.Dimension.Start.Row
                [int]$lCol = $ws.Dimension.End.Column
                for ($c = 1; $c -le $lCol; $c++) {
                    $hdr = $ws.Cells[$hRow, $c].Text
                    if ($hdr -match 'Cost \(EUR\)' -or $hdr -match 'Price \(EUR\)') {
                        $ws.Column($c).Style.Numberformat.Format = '€#,##0.00'
                    }
                }
                $_.Save()
                $_.Dispose()
            }
            Write-Host "    [$($catTab.Name)] $($filtered.Count) users" -ForegroundColor Green
        } catch {
            Write-Log "XLSX $($catTab.Name) tab: $($_.Exception.Message)" -Level WARN
        }
    }
    $allCsvRows = $null  # free memory

    # ── Sheet: Service Plans (consolidated: one row per user per SKU) ──
    $currentExcelSheet = "Service Plans"
    $consolidatedPlans.Values |
        Sort-Object UserPrincipalName, SkuPartNumber |
        Select-Object UserPrincipalName, DisplayName, Department, SkuPartNumber,
            @{N='SkuFriendlyName'; E={ Resolve-SkuFriendlyName $_.SkuPartNumber }},
            TotalPlans, EnabledCount, DisabledCount,
            @{N='DisabledPlans'; E={ $_.DisabledPlanNames -join '; ' }} |
        Export-Excel -Path $xlFile -WorksheetName "Service Plans" `
            -TableName "ServicePlans" -TableStyle Medium6 -FreezeTopRow -AutoFilter -AutoSize
    $consolidatedPlans = $null  # free memory

    # ── Sheet 4: SKU Inventory (enriched) ──
    $currentExcelSheet = "SKU Inventory"
    $skuInvData = $subscribedSkus | Select-Object SkuPartNumber,
        @{N='Friendly Name'; E={ Resolve-SkuFriendlyName $_.SkuPartNumber }},
        SkuId, AppliesTo, CapabilityStatus,
        @{N='Total';     E={$_.PrepaidUnits.Enabled + $_.PrepaidUnits.Warning}},
        @{N='Warning';   E={$_.PrepaidUnits.Warning}},
        @{N='Suspended'; E={$_.PrepaidUnits.Suspended}},
        ConsumedUnits,
        @{N='Available'; E={$_.PrepaidUnits.Enabled + $_.PrepaidUnits.Warning - $_.ConsumedUnits}},
        @{N='Monthly Unit Price (EUR)'; E={ Get-SkuMonthlyPrice $_.SkuPartNumber }},
        @{N='Annual Total Cost (EUR)'; E={ [math]::Round((Get-SkuMonthlyPrice $_.SkuPartNumber) * 12 * $_.ConsumedUnits, 2) }},
        @{N='Subscription Status'; E={
            if ($lkpSubscription.ContainsKey($_.SkuPartNumber)) { $lkpSubscription[$_.SkuPartNumber].Status } else { "" }
        }},
        @{N='Next Lifecycle Date'; E={
            if ($lkpSubscription.ContainsKey($_.SkuPartNumber) -and $lkpSubscription[$_.SkuPartNumber].NextLifecycle) {
                ([datetime]$lkpSubscription[$_.SkuPartNumber].NextLifecycle).ToString("yyyy-MM-dd")
            } else { "" }
        }},
        @{N='Days Until Expiry'; E={
            if ($lkpSubscription.ContainsKey($_.SkuPartNumber) -and $lkpSubscription[$_.SkuPartNumber].NextLifecycle) {
                [int]([datetime]$lkpSubscription[$_.SkuPartNumber].NextLifecycle - (Get-Date)).TotalDays
            } else { "" }
        }},
        @{N='Is Trial'; E={
            if ($cloudLicensingData.ContainsKey($_.SkuPartNumber)) { $cloudLicensingData[$_.SkuPartNumber].IsTrial } else { "" }
        }},
        @{N='Cloud State'; E={
            if ($cloudLicensingData.ContainsKey($_.SkuPartNumber)) { $cloudLicensingData[$_.SkuPartNumber].State } else { "" }
        }},
        @{N='Capacity Utilization %'; E={
            if ($cloudLicensingData.ContainsKey($_.SkuPartNumber)) { $cloudLicensingData[$_.SkuPartNumber].CapacityPct } else { "" }
        }},
        @{N='Pricing Known'; E={ $skuMonthlyPrices.ContainsKey($_.SkuPartNumber) }}

    $skuInvData | Export-Excel -Path $xlFile -WorksheetName "SKU Inventory" `
        -TableName "SkuInventory" -TableStyle Medium6 -FreezeTopRow -AutoFilter -AutoSize -PassThru | ForEach-Object {
        $ws = $_.Workbook.Worksheets["SKU Inventory"]
        $ws.TabColor = [System.Drawing.Color]::FromArgb(112, 173, 71)  # Green
        $lastCol = $ws.Dimension.End.Column
        $lastRow = $ws.Dimension.End.Row
        # Find Days Until Expiry column
        for ($c = 1; $c -le $lastCol; $c++) {
            if ($ws.Cells[1, $c].Text -eq 'Days Until Expiry') {
                $addr = [OfficeOpenXml.ExcelAddress]::new(2, $c, $lastRow, $c)
                $cfExpiry = $ws.ConditionalFormatting.AddExpression($addr)
                $colLetter = [OfficeOpenXml.ExcelCellAddress]::new(2, $c).Address -replace '\d+',''
                $cfExpiry.Formula = "AND(ISNUMBER(${colLetter}2),${colLetter}2<30)"
                $cfExpiry.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::FromArgb(255, 199, 206)
                break
            }
        }
        # Currency format on cost columns
        for ($c = 1; $c -le $lastCol; $c++) {
            $hdr = $ws.Cells[1, $c].Text
            if ($hdr -match 'Cost \(EUR\)' -or $hdr -match 'Price \(EUR\)') {
                $ws.Column($c).Style.Numberformat.Format = '€#,##0.00'
            }
        }
        $_.Save()
        $_.Dispose()
    }

    # ── Sheet 5: Group Licensing Inventory ──
    $currentExcelSheet = "Group Licensing"
    if ($groupLicenseInventory.Count -gt 0) {
        @($groupLicenseInventory) | Export-Excel -Path $xlFile -WorksheetName "Group Licensing" `
            -TableName "GroupLicensing" -TableStyle Medium6 -FreezeTopRow -AutoFilter -AutoSize
    }

    # ── Sheet 6: Cost by Department ──
    $currentExcelSheet = "Cost by Department"
    if ($costByDepartment.Count -gt 0) {
        $costByDepartment | Export-Excel -Path $xlFile -WorksheetName "Cost by Department" `
            -TableName "CostByDept" -TableStyle Medium6 -AutoSize -PassThru | ForEach-Object {
            $ws = $_.Workbook.Worksheets["Cost by Department"]
            $ws.TabColor = [System.Drawing.Color]::FromArgb(91, 155, 213)  # Light blue
            # Currency format on Annual Cost column
            $lastCol = $ws.Dimension.End.Column
            for ($c = 1; $c -le $lastCol; $c++) {
                if ($ws.Cells[1, $c].Text -match 'Cost') { $ws.Column($c).Style.Numberformat.Format = '€#,##0.00' }
            }
            # Add bar chart
            [int]$lastRow = $ws.Dimension.End.Row
            [int]$chartRows = [math]::Min($lastRow - 1, 20)
            if ($chartRows -gt 0) {
                [int]$chartEndRow = 1 + $chartRows
                $chart = $ws.Drawings.AddChart("DeptCostChart", [OfficeOpenXml.Drawing.Chart.eChartType]::BarClustered)
                $chart.Title.Text = "Annual License Cost by Department"
                $chart.SetPosition(1, 0, 4, 0)
                $chart.SetSize(700, 400)
                $series = $chart.Series.Add(
                    [OfficeOpenXml.ExcelAddress]::new(2, 3, $chartEndRow, 3).Address,
                    [OfficeOpenXml.ExcelAddress]::new(2, 1, $chartEndRow, 1).Address
                )
                $series.Header = "Annual Cost (EUR)"
            }
            $_.Save()
            $_.Dispose()
        }
    }

    # ── Sheet 7: Recommendations (pivot-style grouping) ──
    $currentExcelSheet = "Recommendations"
    $recPivotData = @($recDistribution.GetEnumerator() | ForEach-Object {
        [PSCustomObject]@{
            'Recommendation Category' = $_.Key
            'User Count'              = $_.Value.Count
            'Annual Cost (EUR)'       = [math]::Round($_.Value.AnnualCost, 2)
        }
    } | Sort-Object 'Annual Cost (EUR)' -Descending)

    $recPivotData | Export-Excel -Path $xlFile -WorksheetName "Recommendations" `
        -TableName "RecSummary" -TableStyle Medium6 -AutoSize -PassThru | ForEach-Object {
        $ws = $_.Workbook.Worksheets["Recommendations"]
        $ws.TabColor = [System.Drawing.Color]::FromArgb(237, 125, 49)  # Orange
        $lastCol = $ws.Dimension.End.Column
        for ($c = 1; $c -le $lastCol; $c++) {
            if ($ws.Cells[1, $c].Text -match 'Cost') { $ws.Column($c).Style.Numberformat.Format = '€#,##0.00' }
        }
        # Pie chart showing recommendation distribution
        [int]$lastRow = $ws.Dimension.End.Row
        [int]$chartRows = $lastRow - 1
        if ($chartRows -gt 0) {
            $chart = $ws.Drawings.AddChart("RecPieChart", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie3D)
            $chart.Title.Text = "Recommendation Distribution (by Cost)"
            $chart.SetPosition(1, 0, 4, 0)
            $chart.SetSize(600, 400)
            $series = $chart.Series.Add(
                [OfficeOpenXml.ExcelAddress]::new(2, 3, $lastRow, 3).Address,
                [OfficeOpenXml.ExcelAddress]::new(2, 1, $lastRow, 1).Address
            )
            $chart.DataLabel.ShowPercent   = $true
            $chart.DataLabel.ShowCategory  = $true
            $chart.DataLabel.ShowLeaderLines = $true
        }
        $_.Save()
        $_.Dispose()
    }

    # ── Sheet 8: Intensity Analysis (cross-tab) ──
    $currentExcelSheet = "Intensity Analysis"
    $intensityCrossTab = @($intensityCrossDict.GetEnumerator() | ForEach-Object {
            $parts = $_.Key -split ','
            [PSCustomObject]@{
                'Exchange Intensity' = $parts[0]
                'Teams Intensity'    = $parts[1]
                'User Count'         = $_.Value
            }
        } | Sort-Object 'Exchange Intensity','Teams Intensity')

    if ($intensityCrossTab.Count -gt 0) {
        $intensityCrossTab | Export-Excel -Path $xlFile -WorksheetName "Intensity Analysis" `
            -TableName "IntensityMatrix" -TableStyle Medium6 -AutoSize
    }

    # ── Executive Summary sheet (inserted as first sheet) ──
    $currentExcelSheet = "Executive Summary"
    $pkg = Open-ExcelPackage -Path $xlFile
    $execWs = $pkg.Workbook.Worksheets.Add("Executive Summary")
    $pkg.Workbook.Worksheets.MoveToStart("Executive Summary")

    # Title
    $execWs.Cells["A1"].Value = "M365 License Assessment — Executive Summary"
    $execWs.Cells["A1"].Style.Font.Size = 16
    $execWs.Cells["A1"].Style.Font.Bold = $true
    $execWs.Cells["A2"].Value = "$(Get-Date -Format 'yyyy-MM-dd')  |  Period: $ReportPeriod  |  Tenant: $($ctx.TenantId)"
    $execWs.Cells["A2"].Style.Font.Color.SetColor([System.Drawing.Color]::Gray)
    $execWs.Cells["A3"].Value = "Pricing disclaimer: All cost figures are indicative estimates based on public Microsoft list prices (EUR). Actual costs may differ due to EA/CSP/volume pricing. Verify against your invoice."
    $execWs.Cells["A3"].Style.Font.Color.SetColor([System.Drawing.Color]::Gray)
    $execWs.Cells["A3"].Style.Font.Italic = $true
    $execWs.Cells["A3"].Style.Font.Size = 9

    # Headline KPIs
    $execWs.Cells["A4"].Value = "Total Annual M365 Spend"
    $execWs.Cells["A4"].Style.Font.Bold = $true
    $execWs.Cells["B4"].Value = $totalAnnualSpend
    $execWs.Cells["B4"].Style.Numberformat.Format = '€#,##0.00'
    $execWs.Cells["B4"].Style.Font.Size = 14

    $execWs.Cells["A5"].Value = "Estimated Optimization Potential"
    $execWs.Cells["A5"].Style.Font.Bold = $true
    $execWs.Cells["A5"].Style.Font.Size = 14
    $execWs.Cells["A5"].Style.Font.Color.SetColor([System.Drawing.Color]::DarkRed)
    $execWs.Cells["B5"].Value = $totalMoneyOnTable
    $execWs.Cells["B5"].Style.Numberformat.Format = '€#,##0.00'
    $execWs.Cells["B5"].Style.Font.Size = 14
    $execWs.Cells["B5"].Style.Font.Bold = $true
    $execWs.Cells["B5"].Style.Font.Color.SetColor([System.Drawing.Color]::DarkRed)
    $execWs.Cells["C5"].Value = "$wastePercentage% of annual spend"
    $execWs.Cells["C5"].Style.Font.Color.SetColor([System.Drawing.Color]::DarkRed)

    # Tier 1 table
    $execWs.Cells["A7"].Value = "TIER 1 — Quick Wins"
    $execWs.Cells["A7"].Style.Font.Bold = $true
    $execWs.Cells["A7"].Style.Font.Size = 12

    $execWs.Cells["A8"].Value = "Category"
    $execWs.Cells["B8"].Value = "Users"
    $execWs.Cells["C8"].Value = "Annual Amount (EUR)"
    $execWs.Cells["A8"].Style.Font.Bold = $true
    $execWs.Cells["B8"].Style.Font.Bold = $true
    $execWs.Cells["C8"].Style.Font.Bold = $true

    $t1Data = @(
        @("Dormant Accounts (no sign-in >$InactiveSignInDays days)", $dormantTier1Count, $dormantCost),
        @("Disabled Accounts (sign-in blocked)", ($disabledLicensed - $disabledFreeSku), $disabledCost),
        @("Zero M365 Usage (no app activity in period)", $noActivity, $noActivityCost),
        @("Unused Premium Add-Ons (Visio/Project/PBI Pro)", $shelfware, $shelfwareCost),
        @("Copilot Reclaim (zero usage & zero readiness)", $copilotReclaim, $copilotReclaimCost),
        @("Copilot At Risk (zero usage, active in M365)", $copilotWatchlist, $copilotWatchlistCost),
        @("Shared Mailbox (no license needed under 50 GB)", $sharedMbxRemovable, $sharedMbxCost),
        @("Duplicate Licenses (standalone included in suite)", $duplicateCov, $duplicateCost)
    )
    [int]$eRow = 9
    foreach ($t in $t1Data) {
        $execWs.Cells[$eRow, 1].Value = $t[0]
        $execWs.Cells[$eRow, 2].Value = $t[1]
        $execWs.Cells[$eRow, 3].Value = $t[2]
        $execWs.Cells[$eRow, 3].Style.Numberformat.Format = '€#,##0.00'
        $eRow++
    }
    # Tier 1 subtotal
    $execWs.Cells[$eRow, 1].Value = "TIER 1 SUBTOTAL"
    $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 3].Value = $tier1Waste
    $execWs.Cells[$eRow, 3].Style.Numberformat.Format = '€#,##0.00'
    $execWs.Cells[$eRow, 3].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 4].Value = "$tier1Percentage%"
    $eRow++

    # Unassigned license inventory (informational — per SKU)
    if ($unassignedLicenseInventory.Count -gt 0) {
        $eRow++
        $execWs.Cells[$eRow, 1].Value = "Unassigned Licenses ($totalUnassignedSeats seats across $($unassignedLicenseInventory.Count) SKU(s))"
        $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
        $execWs.Cells[$eRow, 1].Style.Font.Italic = $true
        $eRow++
        foreach ($uLic in $unassignedLicenseInventory) {
            $execWs.Cells[$eRow, 1].Value = $uLic.FriendlyName
            $execWs.Cells[$eRow, 2].Value = "$($uLic.Unassigned)/$($uLic.Total) unassigned"
            $eRow++
        }
    }
    $eRow++

    # Tier 2 table
    $execWs.Cells[$eRow, 1].Value = "TIER 2 — Right-Sizing Opportunities"
    $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 1].Style.Font.Size = 12
    $eRow++

    $execWs.Cells[$eRow, 1].Value = "Category"
    $execWs.Cells[$eRow, 2].Value = "Users"
    $execWs.Cells[$eRow, 3].Value = "Annual Amount (EUR)"
    $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 2].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 3].Style.Font.Bold = $true
    $eRow++

    $t2Data = @(
        @("E3/E5 to Frontline F1/F3 (web/mobile only users)", $frontlineCandidate, $frontlineSavings),
        @("Business Standard to Basic (no desktop apps used)", $businessDowngrade, $businessBasicSavings),
        @("Exchange Plan 2 to Plan 1 (mailbox under 50 GB)", $exoPlan2Review, $exoPlan2Savings),
        @("E3 + Add-Ons to E5 Upgrade (cheaper as E5)", ($e5Upgrade + $suiteInversion), $e5UpgradeSavings),
        @("O365+EMS+Windows to M365 Bundle (cheaper combined)", $bundleConsolidation, $bundleConsolidationSavings),
        @("O365 E1 to Business Basic (same features, lower cost)", $e1Downgrade, $e1DowngradeSavings),
        @("O365 E3 to E1 (web/mobile only, mailbox <50 GB)", $o365E3Downgrade, $o365E3DowngradeSavings),
        @("M365 E3 to Business Premium (<300 seats, cheaper)", $e3Downgrade, $e3DowngradeSavings),
        @("E5 to No Audio Conferencing Variant (0 calls)", $e5VoiceWaste, $e5VoiceSavings),
        @("Apps Enterprise to Apps Business (<300 seats)", $appArbitrage, $appArbitrageSavings),
        @("PBI PPU Standalone to Add-On (Pro from suite)", $ppuArbitrage, $ppuArbitrageSavings),
        @("Exchange Plan 1 to Kiosk (web-only, <2 GB)", $exoKioskDowngrade, $exoKioskSavings),
        @("Biz Standard + Add-Ons to Premium (cheaper)", $bizPremInversion, $bizPremInversionSavings),
        @("F-License Blocked, E1/Basic Alternative", $frontlineRescue, $frontlineRescueSavings)
    )
    foreach ($t in $t2Data) {
        $execWs.Cells[$eRow, 1].Value = $t[0]
        $execWs.Cells[$eRow, 2].Value = $t[1]
        $execWs.Cells[$eRow, 3].Value = $t[2]
        $execWs.Cells[$eRow, 3].Style.Numberformat.Format = '€#,##0.00'
        $eRow++
    }
    # Tier 2 subtotal
    $execWs.Cells[$eRow, 1].Value = "TIER 2 SUBTOTAL"
    $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 3].Value = $tier2Savings
    $execWs.Cells[$eRow, 3].Style.Numberformat.Format = '€#,##0.00'
    $execWs.Cells[$eRow, 3].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 4].Value = "$tier2Percentage%"
    $eRow += 2

    # ── Tenant-Level Optimization table ──
    $execWs.Cells[$eRow, 1].Value = "TENANT-LEVEL OPTIMIZATION"
    $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 1].Style.Font.Size = 12
    $eRow++
    $execWs.Cells[$eRow, 1].Value = "Category"
    $execWs.Cells[$eRow, 2].Value = "Users"
    $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 2].Style.Font.Bold = $true
    $eRow++
    $tenantData = @(
        @("Overlapping Assignments (same SKU via direct + group)", $overlapping),
        @("Intune Entitlement Unused (0 enrolled devices)", $intuneShelfware),
        @("Intune Suite Add-On Redundant (included in E3/E5)", $intuneSuiteWaste),
        @("MDM/MAM Unused (web-only users, no devices)", $mdmMamWaste),
        @("Windows License on Non-Windows Users", $windowsLicenseWaste),
        @("Archive Add-On Unused (no archive, small mailbox)", $overLicensedArchive),
        @("Visio Plan 1 Redundant (web Visio in E3/E5)", $seededVisioOverlap),
        @("Guest Users with Paid Licenses", $guestAccountWaste),
        @("Shared/Room Mailboxes on Premium Suites", $nonHumanWaste),
        @("Self-Service & Trial Licenses (cleanup)", ($viralCleanup + $trialLicenseUsers)),
        @("Teams Unused (switch to Without Teams SKU)", $teamsUnbundling),
        @("Archive Add-On Redundant (suite includes archive)", $redundantArchive)
    )
    foreach ($t in $tenantData) {
        $execWs.Cells[$eRow, 1].Value = $t[0]
        $execWs.Cells[$eRow, 2].Value = $t[1]
        $eRow++
    }
    $eRow++

    # ── Product-Specific Flags table ──
    $execWs.Cells[$eRow, 1].Value = "PRODUCT-SPECIFIC FLAGS"
    $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 1].Style.Font.Size = 12
    $eRow++
    $execWs.Cells[$eRow, 1].Value = "Category"
    $execWs.Cells[$eRow, 2].Value = "Users"
    $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 2].Style.Font.Bold = $true
    $eRow++
    $productData = @(
        @("Teams Phone Without Calling Plan (verify PSTN route)", $phoneNoPlan),
        @("Calling Plan Unused (0 calls in period)", $callingPlanWaste),
        @("Teams Premium + Copilot Overlap (remove Premium)", $aiAddonOverlap),
        @("Teams Premium + Copilot (review webinar need)", $aiOverlapReview),
        @("Power BI Pro Low Usage (review if needed)", $pbiProReview),
        @("OneDrive Plan 2 to Plan 1 (using <900 GB)", $odPlan2Waste),
        @("Entra P2 to P1 (no admin roles, no PIM, no risk CA)", $entraP2Downgrade),
        @("Desktop App License Unused (web/mobile only)", $standaloneAppsWaste),
        @("F3 to F1 (empty mailbox & OneDrive)", $f3ToF1Downgrade),
        @("Frontline + Add-Ons Exceed E3/Premium Price", $frontlineAddonBloat),
        @("Teams Phone to Resource Account", $teamsPhoneRightSizing),
        @("Premium Add-On Unused (no activity detected)", $premiumAddonWaste),
        @("Standalone License Replaceable by Cheaper SKU", $alaCarteWaste),
        @("Separate SKUs Cheaper Than Current Bundle", $bundleInefficiency)
    )
    foreach ($t in $productData) {
        $execWs.Cells[$eRow, 1].Value = $t[0]
        $execWs.Cells[$eRow, 2].Value = $t[1]
        $eRow++
    }
    $eRow++

    # ── Copilot Adoption Pipeline table ──
    $execWs.Cells[$eRow, 1].Value = "COPILOT ADOPTION PIPELINE"
    $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 1].Style.Font.Size = 12
    $eRow++
    $execWs.Cells[$eRow, 1].Value = "Category"
    $execWs.Cells[$eRow, 2].Value = "Users"
    $execWs.Cells[$eRow, 3].Value = "Annual Amount (EUR)"
    $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 2].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 3].Style.Font.Bold = $true
    $eRow++
    $copilotData = @(
        @("Total Copilot Holders",       ($copilotUsers + $copilotPrereq), ""),
        @("Active Users (Copilot usage detected)", $copilotKeep, ""),
        @("At Risk (zero Copilot usage, active in M365)", $copilotWatchlist, $copilotWatchlistCost),
        @("Reclaim (zero Copilot & zero M365 activity)", $copilotReclaim, $copilotReclaimCost),
        @("Missing Base License (needs E3/E5/Biz Std/Prem)", $copilotPrereq, ""),
        @("Copilot Studio",              $copilotStudioUsers,  "")
    )
    foreach ($t in $copilotData) {
        $execWs.Cells[$eRow, 1].Value = $t[0]
        $execWs.Cells[$eRow, 2].Value = $t[1]
        if ($t[2] -ne "") { $execWs.Cells[$eRow, 3].Value = $t[2]; $execWs.Cells[$eRow, 3].Style.Numberformat.Format = '€#,##0.00' }
        $eRow++
    }
    $eRow++

    # ── Cloud PC Utilization table (only if CPC/W365 SKUs exist) ──
    if ($_hasCpcSku) {
        $execWs.Cells[$eRow, 1].Value = "CLOUD PC UTILIZATION"
        $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
        $execWs.Cells[$eRow, 1].Style.Font.Size = 12
        $eRow++
        $execWs.Cells[$eRow, 1].Value = "Category"
        $execWs.Cells[$eRow, 2].Value = "Users"
        $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
        $execWs.Cells[$eRow, 2].Style.Font.Bold = $true
        $eRow++
        $cpcData = @(
            @("Dormant Cloud PC (0 hours in 90 days)", $dormantCloudPc),
            @("Cloud PC Review (< 10 hours in 90 days)", $cloudPcReview)
        )
        foreach ($t in $cpcData) {
            $execWs.Cells[$eRow, 1].Value = $t[0]
            $execWs.Cells[$eRow, 2].Value = $t[1]
            $eRow++
        }
        $eRow++
    }

    # ── Licensing Compliance table ──
    $execWs.Cells[$eRow, 1].Value = "LICENSING COMPLIANCE"
    $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 1].Style.Font.Size = 12
    $eRow++
    $execWs.Cells[$eRow, 1].Value = "Category"
    $execWs.Cells[$eRow, 2].Value = "Users"
    $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 2].Style.Font.Bold = $true
    $eRow++
    $licCompData = @(
        @("Licensing Compliance Gap (Total)",     $licensingCheck),
        @("Conditional Access Without Entra P1 License", $licensingCheckCA),
        @("Defender for Office Policy Without License", $licensingCheckMDO),
        @("PIM Role Assignment Without Entra P2 License", $licensingCheckPIM),
        @("License Assignment Errors",   $licenseErrors),
        @("Entra Suite + Standalone P2/Governance (redundant)", $entraSuiteOverlap)
    )
    foreach ($t in $licCompData) {
        $execWs.Cells[$eRow, 1].Value = $t[0]
        $execWs.Cells[$eRow, 2].Value = $t[1]
        $eRow++
    }
    $eRow++

    # ── Operational Review table ──
    $execWs.Cells[$eRow, 1].Value = "OPERATIONAL REVIEW"
    $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 1].Style.Font.Size = 12
    $eRow++
    $execWs.Cells[$eRow, 1].Value = "Category"
    $execWs.Cells[$eRow, 2].Value = "Users"
    $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 2].Style.Font.Bold = $true
    $eRow++
    $riskData = @(
        @("Dormant Admin Accounts",      $dormantAdminRisk),
        @("Automation Accounts",          $automationAccount),
        @("Legacy Service Accounts",      $legacyServiceAccount),
        @("Premium License as Cold Storage (0 activity + data)", $expensiveColdStorage),
        @("Unlicensed User With Data (30-day purge risk)", $unlicensedWithData),
        @("Heavy External Sharing Without DLP/Purview", $highRiskSharing),
        @("Forwarding-Only Mailbox (replace with Mail Contact)", $forwardingWaste),
        @("Active User With Mail Forwarding (low usage)", $forwardingReview),
        @("Mailbox Storage Warning",      $mailboxStorageWarning),
        @("OneDrive Storage Warning",     $oneDriveStorageWarning)
    )
    foreach ($t in $riskData) {
        $execWs.Cells[$eRow, 1].Value = $t[0]
        $execWs.Cells[$eRow, 2].Value = $t[1]
        $eRow++
    }
    $eRow++

    # ── Security & Compliance Posture table ──
    $execWs.Cells[$eRow, 1].Value = "SECURITY & COMPLIANCE POSTURE"
    $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 1].Style.Font.Size = 12
    $eRow++
    # Upsell opportunities
    $execWs.Cells[$eRow, 1].Value = "Category"
    $execWs.Cells[$eRow, 2].Value = "Users"
    $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 2].Style.Font.Bold = $true
    $eRow++
    $upsellData = @(
        @("No Defender Protection (Business Basic/Standard)", $securityGap),
        @("Defender Suite Upsell",        $defenderUpsell),
        @("Purview Upsell",              $purviewUpsell),
        @("Business Premium Security Review", $bizPremSecReview)
    )
    foreach ($t in $upsellData) {
        $execWs.Cells[$eRow, 1].Value = $t[0]
        $execWs.Cells[$eRow, 2].Value = $t[1]
        $eRow++
    }
    $eRow++

    # Combined posture matrix (Security vs Compliance side by side)
    $execWs.Cells[$eRow, 1].Value = "Coverage Level"
    $execWs.Cells[$eRow, 2].Value = "Security"
    $execWs.Cells[$eRow, 3].Value = "Compliance"
    $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 2].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 3].Style.Font.Bold = $true
    $eRow++
    $postureMatrix = @(
        @("None",          $secCoverageNone,     $compCoverageNone),
        @("Basic",         $secCoverageBasic,    $compCoverageBasic),
        @("Advanced",      $secCoverageAdvanced, $compCoverageAdvanced),
        @("E5-equivalent", $secCoverageE5,       $compCoverageE5)
    )
    foreach ($t in $postureMatrix) {
        $execWs.Cells[$eRow, 1].Value = $t[0]
        $execWs.Cells[$eRow, 2].Value = $t[1]
        $execWs.Cells[$eRow, 3].Value = $t[2]
        $eRow++
    }
    $eRow += 2

    # Stacked bar chart: Tier 1 vs Tier 2 vs Remaining Spend
    $remainingSpend = [math]::Round($totalAnnualSpend - $totalMoneyOnTable, 2)
    $chartDataStart = $eRow
    $execWs.Cells[$eRow, 1].Value = "Category"
    $execWs.Cells[$eRow, 2].Value = "Amount (EUR)"
    $eRow++
    $execWs.Cells[$eRow, 1].Value = "Tier 1 — Quick Wins"
    $execWs.Cells[$eRow, 2].Value = $tier1Waste
    $execWs.Cells[$eRow, 2].Style.Numberformat.Format = '€#,##0.00'
    $eRow++
    $execWs.Cells[$eRow, 1].Value = "Tier 2 — Right-Sizing Opportunities"
    $execWs.Cells[$eRow, 2].Value = $tier2Savings
    $execWs.Cells[$eRow, 2].Style.Numberformat.Format = '€#,##0.00'
    $eRow++
    $execWs.Cells[$eRow, 1].Value = "Remaining Productive Spend"
    $execWs.Cells[$eRow, 2].Value = $remainingSpend
    $execWs.Cells[$eRow, 2].Style.Numberformat.Format = '€#,##0.00'
    $chartDataEnd = $eRow

    if ($totalAnnualSpend -gt 0) {
        [int]$chartDataBodyStart = $chartDataStart + 1   # first data row (after header)
        $barChart = $execWs.Drawings.AddChart("ExecSpendBreakdown", [OfficeOpenXml.Drawing.Chart.eChartType]::BarStacked)
        $barChart.Title.Text = "Annual Spend Breakdown"
        $barChart.SetPosition(3, 0, 4, 0)
        $barChart.SetSize(550, 350)
        $series = $barChart.Series.Add(
            [OfficeOpenXml.ExcelAddress]::new($chartDataBodyStart, 2, $chartDataEnd, 2).Address,
            [OfficeOpenXml.ExcelAddress]::new($chartDataBodyStart, 1, $chartDataEnd, 1).Address
        )
        $barChart.DataLabel.ShowPercent = $true
    }

    # ── Recommendation Distribution table + Pie chart ──
    $eRow += 2
    $execWs.Cells[$eRow, 1].Value = "RECOMMENDATION DISTRIBUTION"
    $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 1].Style.Font.Size = 12
    $eRow++
    $execWs.Cells[$eRow, 1].Value = "Category"
    $execWs.Cells[$eRow, 2].Value = "Users"
    $execWs.Cells[$eRow, 3].Value = "Annual Cost (EUR)"
    $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 2].Style.Font.Bold = $true
    $execWs.Cells[$eRow, 3].Style.Font.Bold = $true
    [int]$recTableStart = $eRow
    $eRow++
    foreach ($rec in $recPivotData) {
        $execWs.Cells[$eRow, 1].Value = $rec.'Recommendation Category'
        $execWs.Cells[$eRow, 2].Value = $rec.'User Count'
        $execWs.Cells[$eRow, 3].Value = $rec.'Annual Cost (EUR)'
        $execWs.Cells[$eRow, 3].Style.Numberformat.Format = '€#,##0.00'
        $eRow++
    }
    [int]$recTableEnd = $eRow - 1

    if ($recPivotData.Count -gt 0) {
        [int]$recDataStart  = $recTableStart + 1
        [int]$recChartAnchor = $recTableStart
        $pieChart = $execWs.Drawings.AddChart("ExecRecPie", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie3D)
        $pieChart.Title.Text = "Recommendation Distribution"
        $pieChart.SetPosition($recChartAnchor, 0, 4, 0)
        $pieChart.SetSize(500, 350)
        $series = $pieChart.Series.Add(
            [OfficeOpenXml.ExcelAddress]::new($recDataStart, 3, $recTableEnd, 3).Address,
            [OfficeOpenXml.ExcelAddress]::new($recDataStart, 1, $recTableEnd, 1).Address
        )
        $pieChart.DataLabel.ShowPercent  = $true
        $pieChart.DataLabel.ShowCategory = $true
    }

    # ── Unassigned License Pool Waste (conditional) ──
    if ($unassignedPoolWarnings.Count -gt 0) {
        $eRow += 2
        $execWs.Cells[$eRow, 1].Value = "UNASSIGNED LICENSE POOL WASTE"
        $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
        $execWs.Cells[$eRow, 1].Style.Font.Size = 12
        $eRow++
        foreach ($poolWarn in $unassignedPoolWarnings) {
            $execWs.Cells[$eRow, 1].Value = $poolWarn.FriendlyName
            $execWs.Cells[$eRow, 2].Value = "$($poolWarn.Unassigned)/$($poolWarn.Total) unassigned ($($poolWarn.UnassignedPct)%)"
            $execWs.Cells[$eRow, 3].Value = $poolWarn.AnnualWaste
            $execWs.Cells[$eRow, 3].Style.Numberformat.Format = '€#,##0.00'
            $eRow++
        }
        $execWs.Cells[$eRow, 1].Value = "Pool Waste Total"
        $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
        $execWs.Cells[$eRow, 3].Value = [math]::Round($unassignedPoolTotalAnnual, 2)
        $execWs.Cells[$eRow, 3].Style.Numberformat.Format = '€#,##0.00'
        $execWs.Cells[$eRow, 3].Style.Font.Bold = $true
        $eRow++
    }

    # ── Teams Rooms Optimization (conditional) ──
    if ($teamsRoomsDowngrade) {
        $eRow += 2
        $execWs.Cells[$eRow, 1].Value = "TEAMS ROOMS OPTIMIZATION"
        $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
        $execWs.Cells[$eRow, 1].Style.Font.Size = 12
        $eRow++
        $execWs.Cells[$eRow, 1].Value = "Teams Rooms Pro to Basic"
        $execWs.Cells[$eRow, 2].Value = "$($teamsRoomsDowngrade.EligibleRooms) of $($teamsRoomsDowngrade.ProRooms) Pro rooms, $($teamsRoomsDowngrade.BasicRooms)+$($teamsRoomsDowngrade.EligibleRooms)/25 cap"
        $execWs.Cells[$eRow, 3].Value = $teamsRoomsDowngrade.AnnualSavings
        $execWs.Cells[$eRow, 3].Style.Numberformat.Format = '€#,##0.00'
        $eRow++
    }

    # ── Delta vs Prior Run (conditional) ──
    if ($PriorReportPath -and $deltaFile) {
        $eRow += 2
        $execWs.Cells[$eRow, 1].Value = "DELTA vs PRIOR RUN"
        $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
        $execWs.Cells[$eRow, 1].Style.Font.Size = 12
        $eRow++
        $deltaItems = @(
            @("Net Cost Change (Annual)", "EUR $([math]::Round($deltaTotalCurrentCost - $deltaTotalPriorCost, 2).ToString('N2'))"),
            @("Users Added", $deltaNewUsers),
            @("Users Removed", $deltaRemovedUsers),
            @("License Changes", $deltaLicenseChanged),
            @("Waste Addressed", "EUR $([math]::Round($deltaWasteAddressed, 2).ToString('N2'))/yr"),
            @("New Copilot Users", $deltaNewCopilot),
            @("Became Dormant", $deltaBecameDormant),
            @("Became Active", $deltaBecameActive)
        )
        foreach ($d in $deltaItems) {
            $execWs.Cells[$eRow, 1].Value = $d[0]
            $execWs.Cells[$eRow, 1].Style.Font.Bold = $true
            $execWs.Cells[$eRow, 2].Value = $d[1]
            $eRow++
        }
    }

    # Auto-fit columns
    $execWs.Cells[$execWs.Dimension.Address].AutoFitColumns()
    $execWs.Column(1).Width = 50
    $execWs.Column(3).Width = 22

    # Tab colors for Executive Summary and sheets created without -PassThru
    $execWs.TabColor = [System.Drawing.Color]::FromArgb(255, 192, 0)  # Gold
    $tabColorMap = @{
        "Service Plans"     = [System.Drawing.Color]::FromArgb(146, 208, 80)   # Light green
        "Group Licensing"   = [System.Drawing.Color]::FromArgb(169, 208, 142)  # Sage green
        "Intensity Analysis"= [System.Drawing.Color]::FromArgb(155, 187, 227)  # Soft blue
    }
    foreach ($entry in $tabColorMap.GetEnumerator()) {
        $tabWs = $pkg.Workbook.Worksheets[$entry.Key]
        if ($tabWs) { $tabWs.TabColor = $entry.Value }
    }

    Close-ExcelPackage $pkg
    Write-Host "  [6] Excel Workbook       : $xlFile" -ForegroundColor Green

  } catch {
    Write-Log "Excel generation failed on sheet '$currentExcelSheet'" -Level ERROR -ErrorRecord $_
    Write-Warning "  Excel generation failed on sheet '$currentExcelSheet': $($_.Exception.Message)"
    Write-Warning "  Line: $($_.InvocationInfo.ScriptLineNumber)  |  $($_.InvocationInfo.Line.Trim())"
    [void]$script:skippedDataWarnings.Add("Excel workbook — $currentExcelSheet — $($_.Exception.Message)")
    # Attempt to close any open package
    try { if ($pkg) { Close-ExcelPackage $pkg } } catch { Write-Log "Failed to close Excel package: $($_.Exception.Message)" -Level WARN }
  }
} else {
    Write-Host "`n  Excel output skipped (ImportExcel module not available)." -ForegroundColor DarkGray
}

# ── Log file path in output listing ──
if ($script:logFile -and (Test-Path $script:logFile)) {
    Write-Host "  [LOG] Diagnostic log     : $($script:logFile)" -ForegroundColor DarkGray
}

# ── Final stats to log ──
Write-Log "Completed — $totalUsers users analyzed, $licensedUserCount licensed"
Write-Log "Warnings: $($script:skippedDataWarnings.Count) data warnings"
Write-Log "Output folder: $OutputFolder"
if ($script:logFile) { Write-Log "Log file: $($script:logFile)" }

# ─── End of main try block ───
} finally {
    # ═══════════════════════════════════════════════════════════════════════════
    # Cleanup — always runs, even on error or Ctrl+C
    # ═══════════════════════════════════════════════════════════════════════════

    if ($privacyChanged) {
        Write-Host "`n[!] Report privacy was changed to show real UPNs." -ForegroundColor Yellow
        Write-Host "    Re-enable when done: M365 Admin Center > Settings > Org settings > Reports" -ForegroundColor Yellow
        Write-Host "    > 'Display concealed user, group, and site names in all reports'" -ForegroundColor Yellow
    }

    if ($exoConnected) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}

Write-Host "`nDone. Summary saved to: $summaryFile" -ForegroundColor Cyan
if ($script:logFile -and (Test-Path $script:logFile)) {
    Write-Host "Diagnostic log: $($script:logFile)" -ForegroundColor DarkGray
}
