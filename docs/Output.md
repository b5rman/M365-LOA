by looking closely at the generated CSVs and how the script formatted your data, I have identified one critical internationalization bug that will corrupt financial/storage data on European systems, and three logical "noise" flaws where the script is over-reporting.

Here is the analysis of the output and the code fixes required.

1. The Critical Culture/Locale Bug (The "1,1" MB issue)
The Evidence: In your M365_LicenseOptimization_...csv output, look at the Mailbox Size (MB) column. The values are output as "1,1", "62,1", etc. Notice the comma instead of a dot. Your system is using a European locale (e.g., Belgium nl-BE or fr-BE).
The Flaw: Graph API always returns data in US format (e.g., "1234.56" bytes). Look at your Parse-DoubleField function:

PowerShell
function Parse-DoubleField {
    param([string]$Value)
    $cleaned = $Value -replace '[^\d.]',''
    if ($cleaned) { return [double]$cleaned } # <--- DANGER
    return 0
}
When PowerShell casts [double]"1234.56" on a European machine, it expects a comma as the decimal separator. Depending on the PS version, it will either throw an error, drop the decimal entirely (reading it as 123456), or treat the . as a thousands separator.
The Fix: You must force InvariantCulture when parsing Graph data. Replace your Parse-DoubleField function at the top of the script with this:

PowerShell
function Parse-DoubleField {
    param([string]$Value)
    $cleaned = $Value -replace '[^\d.]',''
    if ($cleaned) { 
        return [double]::Parse($cleaned, [System.Globalization.CultureInfo]::InvariantCulture) 
    }
    return 0
}
2. Warning Fatigue (The "Dead Account" Cascade)
The Evidence: Look at the recommendation generated for Emma.Wilson@alost...:

"DORMANT — no interactive sign-in for 378 days... | No M365 desktop/web/mobile app activity in D180... | Low Exchange usage (0 emails)... | Low OneDrive usage (0 actions)."

The Flaw: The script is suffering from "Warning Fatigue". If an account is Dormant, Disabled, or has 0 activity across the board, it is a dead account. There is no reason to also nag the admin that this dead account "has low Exchange usage" or "should be unbundled from Teams."
The Fix: Implement a "Suppression Hierarchy." We need to skip micro-optimizations (Right-Sizing, Low Usage, Desktop vs Web) if the account is fundamentally dead.
Wrap your service-specific checks (around line 1916) in a gate:

PowerShell
# Only run micro-optimization/usage warnings if the account is actually alive
$isDeadAccount = ($isDormant -or -not $isAccountEnabled -or -not $hasAnyActivity)

if (-not $isDeadAccount) {
    # Desktop vs Web apps
    if ($noDesktopApps -and $hasDesktopAppEntitlement) {
        $recommendations.Add("No desktop apps — uses web/mobile only...")
    }
    
    # Service-specific low usage
    if ($emailIntensity -eq "Low" -and $em -and $hasExchangeEntitlement) {
        $mbNote = if ($null -ne $mbSizeMB -and $mbSizeMB -gt 0) { " (mailbox: ${mbSizeMB} MB)" } else { "" }
        $recommendations.Add("Low Exchange usage ($emailTotal emails)$mbNote — consider if mailbox is needed or downgrade.")
    }
    if ($teamsIntensity -eq "Low" -and $tm -and $hasTeamsEntitlement) {
        $recommendations.Add("Low Teams usage ($teamsTotal actions) — may not need full Teams license.")
    }
    
    # Teams Unbundling (move this inside the gate too)
    if ($hasBundledTeamsSku -and $hasTeamsEntitlement -and $teamsTotal -eq 0 -and $au) {
        $recommendations.Add("TEAMS UNBUNDLING — assigned a suite that bundles Teams but shows 0 Teams activity...")
    }
}
3. False-Positive "Missing Data Sources"
The Evidence: In your summary file, it states:

Users with missing sources : 11

If we look at easi.audit@alost..., the script flags them for missing 9 different data sources. But easi.audit is [UNLICENSED]. Unlicensed users do not generate M365 Usage Reports.
The Flaw: Flagging unlicensed (or Guest) users for missing M365 report data pollutes your Data Quality metrics. You only care if a Licensed user is missing telemetry.
The Fix: Update the Missing Data Sources block (around line 1404) to only append if the user is licensed:

PowerShell
$missingDataSources = [System.Collections.Generic.List[string]]::new()
if ($isLicensed) {
    if (-not $au)  { $missingDataSources.Add("ActiveUserDetail") }
    if (-not $app) { $missingDataSources.Add("M365AppPlatform") }
    if (-not $em)  { $missingDataSources.Add("EmailActivity") }
    # ... keep the rest ...
}
if (-not $exoConnected) { $missingDataSources.Add("EXOConfig") }
$missingDataSourcesStr = if ($missingDataSources.Count -gt 0) { $missingDataSources -join "; " } else { "" }
4. Group-Assigned Waste Instructions
The Evidence: For test@alost..., the script output says:

DISABLED ACCOUNT with free SKU... consider cleanup for hygiene.

However, looking at the License Assignment column for that user, it says Group, and the group is SG_LIC_M365_E5.
The Flaw: If a junior admin reads "Remove license," they will go to Entra ID, click the user, and try to remove the license. The button will be greyed out because it's inherited from a group.
The Fix: Enhance the $hasOverlap logic and the general removal strings to explicitly mention group removals.
Add this helper variable near where you define your recommendations:

PowerShell
$removalInstruction = if ($licAssignmentStr -match "Group" -and $licAssignmentStr -notmatch "Both") {
    "Note: License is assigned via group. Remove user from group '$licenseGroupsStr' to reclaim license."
} else {
    "Remove license assignment."
}
Then, on your major waste flags (like Dormant or Disabled), append it:

PowerShell
$recommendations.Add("DORMANT — no interactive sign-in for $daysSinceSignIn days. $re

Drop in the [double]::Parse fix immediately, as that will break your actual storage warnings when you scan mailboxes larger than 1GB!