$tokens = $null
$errors = $null
$null = [System.Management.Automation.Language.Parser]::ParseFile(
    'D:\Claude License Optimisation\Get-M365LicenseOptimizationReport.ps1',
    [ref]$tokens,
    [ref]$errors
)
if ($errors.Count -gt 0) {
    foreach ($e in $errors) {
        Write-Host "ERROR: $($e.Message) at line $($e.Extent.StartLineNumber)"
    }
} else {
    Write-Host "SYNTAX OK - no parse errors."
}
