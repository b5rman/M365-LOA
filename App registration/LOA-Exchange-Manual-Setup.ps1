# ========================================================
# MANUAL EXCHANGE RBAC CONFIGURATION
# Run this script if automatic setup failed
# ========================================================

# Ensure module is installed and imported
Write-Host "Checking ExchangeOnlineManagement module..." -ForegroundColor Yellow
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Installing ExchangeOnlineManagement module..." -ForegroundColor Yellow
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
}

Import-Module ExchangeOnlineManagement
Write-Host "+ Module loaded" -ForegroundColor Green

# Connect to Exchange Online
Write-Host "
Connecting to Exchange Online..." -ForegroundColor Yellow
Connect-ExchangeOnline

# Your App Details
$AppId = "fb7486e8-5ede-4d5b-9422-fc37d18fa3f5"
$ServicePrincipalObjectId = "2cdbb519-928a-4126-a49e-48612fb3242b"

Write-Host "
Checking if Service Principal exists in Exchange..." -ForegroundColor Yellow
$sp = Get-ServicePrincipal -ErrorAction SilentlyContinue | Where-Object { $_.AppId -eq $AppId }

if (-not $sp) {
    Write-Host "Creating Service Principal..." -ForegroundColor Yellow
    try { New-ServicePrincipal -AppId $AppId -ServiceId $ServicePrincipalObjectId }
    catch [System.Management.Automation.ParameterBindingException] { New-ServicePrincipal -AppId $AppId -ObjectId $ServicePrincipalObjectId }

    Write-Host "Waiting 30 seconds for replication..." -ForegroundColor Yellow
    Start-Sleep -Seconds 30

    $sp = Get-ServicePrincipal | Where-Object { $_.AppId -eq $AppId }
}

if ($sp) {
    Write-Host "Service Principal found:" -ForegroundColor Green
    $sp | Format-List AppId, DisplayName, Identity

    $exoAppId = $sp.AppId

    # Assign required roles for License Optimization Audit
    Write-Host "
Assigning 'View-Only Configuration' role (MDO/ATP policies)..." -ForegroundColor Yellow
    try {
        New-ManagementRoleAssignment -Role 'View-Only Configuration' -App $exoAppId -ErrorAction Stop
        Write-Host "+ View-Only Configuration assigned" -ForegroundColor Green
    } catch {
        Write-Host "! View-Only Configuration failed or already exists: $_" -ForegroundColor Yellow
    }

    Write-Host "
Assigning 'View-Only Recipients' role (mailbox properties)..." -ForegroundColor Yellow
    try {
        New-ManagementRoleAssignment -Role 'View-Only Recipients' -App $exoAppId -ErrorAction Stop
        Write-Host "+ View-Only Recipients assigned" -ForegroundColor Green
    } catch {
        Write-Host "! View-Only Recipients failed or already exists: $_" -ForegroundColor Yellow
    }

    Write-Host "
Verifying assignments..." -ForegroundColor Yellow
    Get-ManagementRoleAssignment | Where-Object { $_.RoleAssignee -eq $sp.Identity } | Format-Table Role, RoleAssignee, Name

    Write-Host "
+ Exchange role assignments complete!" -ForegroundColor Green
} else {
    Write-Host "X Service Principal still not found. Please wait 5-10 minutes and try again." -ForegroundColor Red
}
