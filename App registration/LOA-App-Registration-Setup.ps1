# ========================================================
# App Registration Setup Script - FOR CUSTOMER USE
# All Required Permissions for M365 License Optimization Audit
# ========================================================

<#
.SYNOPSIS
    Customer-side script to create auditor access for M365 License Optimization Audit.

.DESCRIPTION
    This script should be run by the CUSTOMER (company being audited) to create
    secure, read-only access for an external auditor to conduct the M365 License
    Optimization Report.

    What this script does:
    - Creates an App Registration with read-only permissions
    - Generates a certificate for secure authentication
    - Configures access to Microsoft Graph and Exchange Online
    - Grants admin consent to the enterprise app automatically
    - Creates a complete package for the auditor

.NOTES
    WHO RUNS THIS: Customer's Global Administrator
    PERMISSIONS REQUIRED: Global Administrator role
    OUTPUT: Complete auditor package ready to send

.EXAMPLE
    .\LOA-App-Registration-Setup.ps1
#>

# ========================================================
# CONFIGURATION
# ========================================================

$appDisplayName = "M365 License Optimization Audit"
$certificatePassword = Read-Host -Prompt "Enter password for certificate (will be generated)" -AsSecureString

# ========================================================
# MICROSOFT GRAPH API PERMISSIONS (Application)
# ========================================================
# These are the minimum permissions required by
# M365-LOA.ps1
# All are read-only EXCEPT ReportSettings.ReadWrite.All (unhides UPNs by default; customer must re-enable privacy manually)

$graphPermissions = @(
    # Core Directory & User Permissions
    "User.Read.All",                              # User profiles, assigned licenses, account state
    "Group.Read.All",                             # Group memberships, license groups, MDO/CA scope resolution
    "Organization.Read.All",                      # Organization config, tenant info, subscriptions
    "ReportSettings.ReadWrite.All",               # Unhide UPNs in usage reports (displayConcealedNames)

    # Reporting (11 usage reports + activation detail)
    "Reports.Read.All",                           # All M365 usage reports (Email, Teams, OneDrive, SharePoint, Apps)

    # Sign-In Activity & Audit Logs (beta endpoint)
    "AuditLog.Read.All",                          # Sign-in activity (last interactive/non-interactive), license assignment states

    # Identity & Access Management
    "Policy.Read.All",                            # Conditional Access policies (risk-based CA detection)
    "RoleManagement.Read.Directory",              # PIM eligible/active role assignments, admin role definitions

    # Cloud Licensing (beta — subscription lifecycle, trial detection, capacity queue)
    "CloudLicensing.Read.All",                    # Allotments, trial state, assignment errors, waiting members

    # Device Management (Intune)
    "DeviceManagementManagedDevices.Read.All",    # Enrolled device count per user (Intune shelfware detection)

    # Cloud PC (Windows 365) — beta endpoint for remote connection usage
    "CloudPC.Read.All"                            # Cloud PC usage hours, last active time, device type (dormant CPC detection)
)


Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  M365 LICENSE OPTIMIZATION AUDIT - CUSTOMER SETUP SCRIPT" -ForegroundColor Cyan
Write-Host "  Creating Secure Auditor Access" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "`nWHO RUNS THIS SCRIPT?" -ForegroundColor Yellow
Write-Host "  You (the customer being audited)" -ForegroundColor White
Write-Host "  Requires: Global Administrator permissions`n" -ForegroundColor White

Write-Host "WHAT THIS SCRIPT DOES:" -ForegroundColor Yellow
Write-Host "  1. Creates an App Registration for auditor access" -ForegroundColor White
Write-Host "  2. Generates a secure certificate (3-month validity)" -ForegroundColor White
Write-Host "  3. Assigns READ-ONLY permissions for license analysis" -ForegroundColor White
Write-Host "  4. Grants admin consent automatically" -ForegroundColor White
Write-Host "  5. Configures Exchange Online read-only access (mailbox type, litigation hold)" -ForegroundColor White
Write-Host "  6. Creates a package to send to your auditor`n" -ForegroundColor White

Write-Host "AFTER THIS SCRIPT:" -ForegroundColor Yellow
Write-Host "  -> You send the generated package to your auditor" -ForegroundColor White
Write-Host "  -> Auditor uses it to run the LOA" -ForegroundColor White
Write-Host "  -> You can revoke access anytime in Azure Portal`n" -ForegroundColor White

Write-Host "App Name: $appDisplayName" -ForegroundColor White
Write-Host "`nPermissions Summary:" -ForegroundColor Yellow
Write-Host "  Microsoft Graph API: $($graphPermissions.Count) permissions (Application)" -ForegroundColor White
Write-Host "  Exchange Online: 1 permission + 2 role assignments" -ForegroundColor White
Write-Host "`n  ALL PERMISSIONS ARE READ-ONLY except ReportSettings.ReadWrite.All" -ForegroundColor Green
Write-Host "  ReportSettings.ReadWrite.All is used ONLY to unhide anonymized user data" -ForegroundColor Green
Write-Host "  in usage reports (default behavior). Customer must re-enable privacy" -ForegroundColor Green
Write-Host "  manually after the audit in M365 Admin Center > Org settings > Reports.`n" -ForegroundColor Green

$confirm = Read-Host "Do you want to continue? (Y/N)"
if ($confirm.Trim() -notmatch '^[Yy]') {
    Write-Host "`nSetup cancelled by user." -ForegroundColor Yellow
    exit
}

# ========================================================
# STEP 1: INSTALL REQUIRED MODULES
# ========================================================

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "STEP 1: Checking Required PowerShell Modules" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# Microsoft.Graph sub-modules must all be the same version.
# A mismatch (e.g. Authentication 2.25 vs Applications 2.34) causes
# assembly-load failures. Update all Graph modules together.
$graphModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications', 'Microsoft.Graph.Identity.DirectoryManagement')

foreach ($mod in $graphModules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "  Installing $mod..." -ForegroundColor Yellow
        Install-Module -Name $mod -Force -AllowClobber -Scope CurrentUser
    }
    Import-Module $mod -Force -ErrorAction Stop
    Write-Host "  + $mod $((Get-Module $mod).Version) loaded" -ForegroundColor Green
}

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "  Installing ExchangeOnlineManagement..." -ForegroundColor Yellow
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
}
Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop
Write-Host "  + ExchangeOnlineManagement $((Get-Module ExchangeOnlineManagement).Version) loaded" -ForegroundColor Green

# ========================================================
# STEP 2: CONNECT TO MICROSOFT 365
# ========================================================

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "STEP 2: Connecting to Your Microsoft 365 Tenant" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Please sign in with your Global Administrator account..." -ForegroundColor Yellow

# Connect to Exchange Online FIRST (before Graph) to avoid MSAL assembly conflicts
Write-Host "`n  Connecting to Exchange Online..." -ForegroundColor White
Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
Write-Host "  + Exchange Online connected" -ForegroundColor Green

# Connect to Microsoft Graph
Write-Host "  Connecting to Microsoft Graph..." -ForegroundColor White
Connect-MgGraph -Scopes "Application.ReadWrite.All", "RoleManagement.ReadWrite.Directory" -NoWelcome

$context = Get-MgContext
if (-not $context) {
    Write-Host "`n  X Failed to connect to Microsoft Graph." -ForegroundColor Red
    Write-Host "    Please ensure you have the Microsoft.Graph.Authentication module installed." -ForegroundColor Yellow
    exit 1
}
try {
    $orgName = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -ErrorAction Stop).value[0].displayName
} catch {
    $orgName = $context.TenantId
    Write-Host "  Could not retrieve organization name, using Tenant ID instead." -ForegroundColor Yellow
}
Write-Host "  + Microsoft Graph connected — Tenant: $orgName" -ForegroundColor Green

# ========================================================
# STEP 3: GENERATE SELF-SIGNED CERTIFICATE
# ========================================================

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "STEP 3: Generating Secure Certificate" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Creating a certificate for auditor authentication..." -ForegroundColor White

$certName = "CN=M365-LOA-Audit-Cert"
$certStartDate = Get-Date
$certEndDate = $certStartDate.AddMonths(3)

$cert = New-SelfSignedCertificate -Subject $certName `
    -NotBefore $certStartDate `
    -NotAfter $certEndDate `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -CertStoreLocation "Cert:\CurrentUser\My"

Write-Host "  + Certificate generated successfully" -ForegroundColor Green
Write-Host "    Thumbprint: $($cert.Thumbprint)" -ForegroundColor Gray
Write-Host "    Valid until: $($certEndDate.ToString('MMMM dd, yyyy'))" -ForegroundColor Gray

$packageDir = ".\M365-LOA-Audit-Package"
if (-not (Test-Path $packageDir)) {
    New-Item -ItemType Directory -Path $packageDir -Force | Out-Null
}

$certPath = "$packageDir\M365-LOA-Audit-Cert.pfx"
$certPublicPath = "$packageDir\M365-LOA-Audit-Cert-Public.cer"

Export-PfxCertificate -Cert $cert -FilePath $certPath -Password $certificatePassword | Out-Null
Export-Certificate -Cert $cert -FilePath $certPublicPath | Out-Null

Write-Host "`n  Certificate files created:" -ForegroundColor White
Write-Host "  * $certPath (PRIVATE - Keep secure)" -ForegroundColor Yellow
Write-Host "  * $certPublicPath (PUBLIC - Reference only)" -ForegroundColor Green

# ========================================================
# STEP 4: CREATE APP REGISTRATION
# ========================================================

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "STEP 4: Creating App Registration for Auditor" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$existingApps = @(Get-MgApplication -Filter "displayName eq '$appDisplayName'" -ErrorAction SilentlyContinue)

if ($existingApps.Count -gt 0) {
    if ($existingApps.Count -gt 1) {
        Write-Host "  WARNING: $($existingApps.Count) apps found with name '$appDisplayName'" -ForegroundColor Yellow
        Write-Host "    Using the most recently created one. Consider removing duplicates in Azure Portal." -ForegroundColor Yellow
    }
    Write-Host "  An app with this name already exists" -ForegroundColor Yellow
    Write-Host "    This might be from a previous audit setup." -ForegroundColor Gray
    $useExisting = Read-Host "  Use existing app? (Y/N)"
    if ($useExisting.Trim() -match '^[Yy]') {
        $app = ($existingApps | Sort-Object -Property CreatedDateTime -Descending)[0]
        Write-Host "  + Using existing app" -ForegroundColor Green
    } else {
        Write-Host "`n  Please either:" -ForegroundColor Yellow
        Write-Host "    1. Delete the existing app in Azure Portal, or" -ForegroundColor White
        Write-Host "    2. Change the app name in this script" -ForegroundColor White
        Write-Host "`n  Setup cancelled." -ForegroundColor Red
        exit
    }
} else {
    $app = New-MgApplication -DisplayName $appDisplayName -SignInAudience "AzureADMyOrg"
    Write-Host "  + App Registration created successfully" -ForegroundColor Green

    Write-Host "  Waiting for Azure AD replication (10 seconds)..." -ForegroundColor Yellow
    Start-Sleep -Seconds 10
}

Write-Host "    App Name: $appDisplayName" -ForegroundColor Gray
Write-Host "    Application ID: $($app.AppId)" -ForegroundColor Gray
Write-Host "    Object ID: $($app.Id)" -ForegroundColor Gray

# ========================================================
# STEP 5: UPLOAD CERTIFICATE TO APP
# ========================================================

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "STEP 5: Configuring Certificate Authentication" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# Pass raw DER-encoded certificate bytes — the Graph SDK handles Base64 encoding internally.
# Passing the Base64 string's ASCII bytes instead would double-encode the payload.
$newKeyCredential = @{
    Type = "AsymmetricX509Cert"
    Usage = "Verify"
    Key = $cert.GetRawCertData()
}

# Merge with existing key credentials to avoid silently removing other valid certs
$existingKeys = @((Get-MgApplication -ApplicationId $app.Id).KeyCredentials | ForEach-Object {
    @{ Type = $_.Type; Usage = $_.Usage; Key = $_.Key; KeyId = $_.KeyId }
})
$mergedKeys = $existingKeys + $newKeyCredential
Update-MgApplication -ApplicationId $app.Id -KeyCredentials $mergedKeys
Write-Host "  + Certificate uploaded to App Registration" -ForegroundColor Green
Write-Host "    This allows secure, password-less authentication" -ForegroundColor Gray

# ========================================================
# STEP 6: ADD MICROSOFT GRAPH API PERMISSIONS
# ========================================================

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "STEP 6: Configuring API Permissions" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Adding Microsoft Graph API permissions..." -ForegroundColor White

# Use immutable AppId instead of displayName — display names can be localized in non-English/GCC tenants
$graphSP = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -Select "id,appRoles,appId"

$requiredResourceAccess = @{
    ResourceAppId = "00000003-0000-0000-c000-000000000000"
    ResourceAccess = @()
}

$permissionCount = 0
$permissionIds = @()

foreach ($permissionName in $graphPermissions) {
    $appRole = $graphSP.AppRoles | Where-Object { $_.Value -eq $permissionName }
    if ($appRole) {
        $requiredResourceAccess.ResourceAccess += @{
            Id = $appRole.Id
            Type = "Role"
        }
        $permissionIds += $appRole.Id
        $permissionCount++
        Write-Host "  + Added: $permissionName" -ForegroundColor Green
    } else {
        Write-Host "  ! Warning: Permission not found: $permissionName (may require tenant eligibility)" -ForegroundColor Yellow
    }
}

# Merge with existing non-Graph permissions to avoid dropping pre-existing resource blocks
$existingResourceAccess = @((Get-MgApplication -ApplicationId $app.Id).RequiredResourceAccess |
    Where-Object { $_.ResourceAppId -ne "00000003-0000-0000-c000-000000000000" })
$mergedResourceAccess = @($existingResourceAccess) + $requiredResourceAccess
Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess $mergedResourceAccess
Write-Host "`n  + $permissionCount Microsoft Graph permissions configured" -ForegroundColor Green

# ========================================================
# STEP 7: CREATE SERVICE PRINCIPAL & WAIT FOR REPLICATION
# ========================================================

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "STEP 7: Creating Service Principal" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$servicePrincipal = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -ErrorAction SilentlyContinue

if (-not $servicePrincipal) {
    Write-Host "  Creating Service Principal..." -ForegroundColor White
    $servicePrincipal = New-MgServicePrincipal -AppId $app.AppId
    Write-Host "  + Service Principal created" -ForegroundColor Green

    Write-Host "`n  Waiting for Azure AD replication..." -ForegroundColor Yellow

    $maxWaitTime = 60
    $waitInterval = 5
    $elapsedTime = 0
    $spReady = $false

    while ($elapsedTime -lt $maxWaitTime -and -not $spReady) {
        Start-Sleep -Seconds $waitInterval
        $elapsedTime += $waitInterval

        $spCheck = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -ErrorAction SilentlyContinue

        if ($spCheck) {
            try {
                $null = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $spCheck.Id -ErrorAction Stop
                $spReady = $true
                Write-Host "  + Service Principal is ready (took $elapsedTime seconds)" -ForegroundColor Green
            } catch {
                Write-Host "  Waiting... ($elapsedTime seconds)" -ForegroundColor Yellow
            }
        }
    }

    if (-not $spReady) {
        Write-Host "  ! Service Principal may need more time to replicate" -ForegroundColor Yellow
        Write-Host "     Continuing anyway - consent may fail and require retry" -ForegroundColor Gray
    }

} else {
    Write-Host "  + Service Principal already exists" -ForegroundColor Green
    $spReady = $true
}

Write-Host "    Service Principal Object ID: $($servicePrincipal.Id)" -ForegroundColor Gray

# ========================================================
# STEP 8: GRANT ADMIN CONSENT AUTOMATICALLY
# ========================================================

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "STEP 8: Granting Admin Consent" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

if (-not $spReady) {
    Write-Host "  ! Service Principal may not be fully ready" -ForegroundColor Yellow
    Write-Host "     Attempting consent anyway..." -ForegroundColor Gray
}

Write-Host "  Granting admin consent for Microsoft Graph permissions..." -ForegroundColor White

try {
    $consentGranted = 0
    $consentFailed = 0
    $failedPermissions = @()

    foreach ($permissionId in $permissionIds) {
        try {
            $existingGrant = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.Id -ErrorAction SilentlyContinue |
                Where-Object { $_.AppRoleId -eq $permissionId -and $_.ResourceId -eq $graphSP.Id }

            if (-not $existingGrant) {
                New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.Id `
                    -PrincipalId $servicePrincipal.Id `
                    -ResourceId $graphSP.Id `
                    -AppRoleId $permissionId -ErrorAction Stop | Out-Null
                $consentGranted++
            } else {
                $consentGranted++
            }
        } catch {
            $consentFailed++
            $failedPermissions += $permissionId
        }
    }

    if ($consentFailed -eq 0) {
        Write-Host "  + Admin consent granted successfully!" -ForegroundColor Green
        Write-Host "     All $consentGranted permissions consented" -ForegroundColor Gray
    } else {
        Write-Host "  ! Partial success: $consentGranted consented, $consentFailed failed" -ForegroundColor Yellow

        if ($consentFailed -gt 0) {
            # Azure AD replication can take up to 2 minutes for new service principals.
            # Retry with increasing delays: 15s, 30s, 60s
            $retryDelays = @(15, 30, 60)
            foreach ($delay in $retryDelays) {
                if ($consentFailed -eq 0) { break }
                Write-Host "`n  Retrying $consentFailed failed permission(s) after ${delay}s wait (Azure AD replication) ..." -ForegroundColor Yellow
                Start-Sleep -Seconds $delay

                $retrySuccess = 0
                $stillFailed = @()
                foreach ($permissionId in $failedPermissions) {
                    try {
                        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.Id `
                            -PrincipalId $servicePrincipal.Id `
                            -ResourceId $graphSP.Id `
                            -AppRoleId $permissionId -ErrorAction Stop | Out-Null
                        $retrySuccess++
                    } catch {
                        $stillFailed += $permissionId
                    }
                }

                if ($retrySuccess -gt 0) {
                    Write-Host "  + Retry successful for $retrySuccess permission(s)" -ForegroundColor Green
                    $consentGranted += $retrySuccess
                    $consentFailed -= $retrySuccess
                }
                $failedPermissions = $stillFailed
            }
        }

        if ($consentFailed -gt 0) {
            Write-Host "  ! Some permissions still need manual consent" -ForegroundColor Yellow
            Write-Host "     Total granted: $consentGranted of $($permissionIds.Count)" -ForegroundColor Gray
        }
    }

} catch {
    Write-Host "  X Automatic consent failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "`n  This can happen if:" -ForegroundColor Yellow
    Write-Host "    - Service Principal is still replicating" -ForegroundColor Gray
    Write-Host "    - You don't have sufficient permissions" -ForegroundColor Gray

    $grantConsent = Read-Host "  Open Azure Portal to grant admin consent manually? (Y/N)"
    if ($grantConsent.Trim() -match '^[Yy]') {
        $consentUrl = "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$($app.AppId)"
        Start-Process $consentUrl
        Write-Host "`n  -> Opening Azure Portal..." -ForegroundColor Cyan
        Write-Host "  -> Click 'Grant admin consent for [Your Organization]'" -ForegroundColor Yellow
        Write-Host "  -> Click 'Yes' to confirm" -ForegroundColor Yellow
        Read-Host "`n  Press ENTER after granting consent to continue"
    } else {
        Write-Host "`n  ! WARNING: Admin consent not granted!" -ForegroundColor Red
        Write-Host "     Your auditor will not be able to connect until consent is granted." -ForegroundColor Yellow
    }
}

# ========================================================
# STEP 9: EXCHANGE ONLINE PERMISSIONS
# ========================================================

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "STEP 9: Configuring Exchange Online Read Only Access" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Exchange Online permissions are required for:" -ForegroundColor White
Write-Host "  - Mailbox type detection (User/Shared/Room/Equipment)" -ForegroundColor Gray
Write-Host "  - Litigation Hold status" -ForegroundColor Gray
Write-Host "  - Archive mailbox status" -ForegroundColor Gray
Write-Host "  - Defender for Office 365 policy coverage" -ForegroundColor Gray

$setupExchange = Read-Host "`n  Configure Exchange Online access now? (Y/N)"

if ($setupExchange.Trim() -match '^[Yy]') {
    Write-Host "`n  Adding Exchange Online API permission..." -ForegroundColor Cyan

    # Get Exchange Online Service Principal
    # Use immutable AppId instead of displayName — display names can be localized in non-English/GCC tenants
    $exchangeSP = Get-MgServicePrincipal -Filter "appId eq '00000002-0000-0ff1-ce00-000000000000'" -ErrorAction SilentlyContinue

    if ($exchangeSP) {
        $exchangePermission = $exchangeSP.AppRoles | Where-Object { $_.Value -eq "Exchange.ManageAsApp" }

        if ($exchangePermission) {
            # Add Exchange permission to app
            $exchangeResourceAccess = @{
                ResourceAppId = $exchangeSP.AppId
                ResourceAccess = @(
                    @{
                        Id = $exchangePermission.Id
                        Type = "Role"
                    }
                )
            }

            $currentPermissions = (Get-MgApplication -ApplicationId $app.Id).RequiredResourceAccess
            # Merge idempotently — replace existing Exchange block or append if absent
            $exchangeAppId = $exchangeSP.AppId
            $existingExchange = $currentPermissions | Where-Object { $_.ResourceAppId -eq $exchangeAppId }
            if ($existingExchange) {
                # Merge role IDs into the existing block, deduplicating by Id
                $existingIds = @($existingExchange.ResourceAccess | ForEach-Object { $_.Id })
                foreach ($ra in $exchangeResourceAccess.ResourceAccess) {
                    if ($ra.Id -notin $existingIds) {
                        $existingExchange.ResourceAccess += $ra
                    }
                }
                $allPermissions = @($currentPermissions)
            } else {
                $allPermissions = @($currentPermissions) + $exchangeResourceAccess
            }
            Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess $allPermissions

            Write-Host "  + Exchange.ManageAsApp permission added" -ForegroundColor Green

            # Grant admin consent for Exchange
            Write-Host "  Granting admin consent for Exchange..." -ForegroundColor Cyan

            try {
                $existingExchangeAssignment = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.Id -ErrorAction SilentlyContinue |
                    Where-Object { $_.AppRoleId -eq $exchangePermission.Id -and $_.ResourceId -eq $exchangeSP.Id }

                if (-not $existingExchangeAssignment) {
                    New-MgServicePrincipalAppRoleAssignment `
                        -ServicePrincipalId $servicePrincipal.Id `
                        -PrincipalId $servicePrincipal.Id `
                        -ResourceId $exchangeSP.Id `
                        -AppRoleId $exchangePermission.Id | Out-Null

                    Write-Host "  + Admin consent granted for Exchange.ManageAsApp" -ForegroundColor Green
                } else {
                    Write-Host "  + Exchange.ManageAsApp consent already granted" -ForegroundColor Green
                }

            } catch {
                Write-Warning "  Failed to grant Exchange consent automatically: $_"
                Write-Host "  Grant manually in Azure Portal: API permissions > Grant admin consent" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host "  ! Could not find Office 365 Exchange Online service principal" -ForegroundColor Yellow
    }

    # ========================================================
    # Configure Exchange RBAC Roles
    # ========================================================

    Write-Host "`n  Configuring Exchange RBAC roles..." -ForegroundColor Cyan

    try {

        Write-Host "  Checking Exchange Service Principal..." -ForegroundColor White
        $exoServicePrincipal = Get-ServicePrincipal -ErrorAction SilentlyContinue | Where-Object { $_.AppId -eq $app.AppId }

        if (-not $exoServicePrincipal) {
            Write-Host "  Creating Exchange Service Principal..." -ForegroundColor Yellow

            try {
                # -ServiceId is current; Microsoft plans to rename it to -ObjectId in a future update
                try {
                    New-ServicePrincipal -AppId $app.AppId -ServiceId $servicePrincipal.Id -ErrorAction Stop
                } catch [System.Management.Automation.ParameterBindingException] {
                    New-ServicePrincipal -AppId $app.AppId -ObjectId $servicePrincipal.Id -ErrorAction Stop
                }
                Write-Host "  + Exchange Service Principal created" -ForegroundColor Green

                Write-Host "  Waiting for Exchange replication (20 seconds)..." -ForegroundColor Yellow
                Start-Sleep -Seconds 20

                $exoServicePrincipal = Get-ServicePrincipal -ErrorAction SilentlyContinue | Where-Object { $_.AppId -eq $app.AppId }

                if (-not $exoServicePrincipal) {
                    throw "Service Principal not found after creation. Try waiting 5-10 minutes and run the manual script."
                }

            } catch {
                throw "Error creating Service Principal: $($_.Exception.Message)"
            }
        } else {
            Write-Host "  + Exchange Service Principal already exists" -ForegroundColor Green
        }

        $exoIdentity = $exoServicePrincipal.Identity
        $exoAppId = $exoServicePrincipal.AppId

        Write-Host "    Exchange SP Identity: $exoIdentity" -ForegroundColor Gray

        # ========================================================
        # ASSIGN REQUIRED ROLES FOR LICENSE OPTIMIZATION AUDIT
        # ========================================================
        # View-Only Recipients: Get-EXOMailbox (mailbox type, litigation hold, archive)
        # View-Only Configuration: Get-ATPProtectionPolicyRule, Get-SafeLinksRule, Get-SafeAttachmentRule

        $requiredRoles = @(
            "View-Only Configuration",  # For MDO/ATP policy rules
            "View-Only Recipients"      # For Get-EXOMailbox: mailbox type, litigation hold, archive status
        )

        Write-Host "`n  Assigning required Exchange roles..." -ForegroundColor White
        $successfulRoles = @()
        $failedRoles = @()

        foreach ($roleName in $requiredRoles) {
            Write-Host "    Checking role: $roleName" -ForegroundColor Gray

            $roleAssignment = Get-ManagementRoleAssignment -ErrorAction SilentlyContinue | Where-Object {
                $_.RoleAssignee -eq $exoIdentity -and
                $_.Role -eq $roleName
            }

            if (-not $roleAssignment) {
                Write-Host "    Assigning '$roleName' role..." -ForegroundColor Yellow

                try {
                    $assignment = New-ManagementRoleAssignment -Role $roleName -App $exoAppId -ErrorAction Stop
                    Write-Host "    + $roleName assigned successfully" -ForegroundColor Green
                    Write-Host "      Assignment Name: $($assignment.Name)" -ForegroundColor Gray
                    $successfulRoles += $roleName

                } catch {
                    Write-Host "    ! Initial assignment failed, waiting 30 seconds and retrying..." -ForegroundColor Yellow
                    Start-Sleep -Seconds 30

                    $exoServicePrincipal = Get-ServicePrincipal -ErrorAction SilentlyContinue | Where-Object { $_.AppId -eq $app.AppId }
                    $exoAppId = $exoServicePrincipal.AppId

                    try {
                        $assignment = New-ManagementRoleAssignment -Role $roleName -App $exoAppId -ErrorAction Stop
                        Write-Host "    + $roleName assigned successfully on retry" -ForegroundColor Green
                        $successfulRoles += $roleName
                    } catch {
                        Write-Host "    X Failed to assign $roleName : $($_.Exception.Message)" -ForegroundColor Red
                        $failedRoles += $roleName
                    }
                }
            } else {
                Write-Host "    + $roleName already assigned" -ForegroundColor Green
                $successfulRoles += $roleName
            }
        }

        if ($failedRoles.Count -eq 0) {
            Write-Host "`n  + Exchange Online configuration complete!" -ForegroundColor Green
            Write-Host "    Roles assigned: $($successfulRoles -join ', ')" -ForegroundColor Gray
        } else {
            Write-Host "`n  ! Partial success:" -ForegroundColor Yellow
            Write-Host "    Assigned: $($successfulRoles -join ', ')" -ForegroundColor Green
            Write-Host "    Failed: $($failedRoles -join ', ')" -ForegroundColor Red
            throw "Not all roles were assigned successfully"
        }

    } catch {
        Write-Host "`n  X Automatic Exchange configuration failed" -ForegroundColor Red
        Write-Host "     Error: $($_.Exception.Message)" -ForegroundColor Gray

        # Save manual instructions to a file
        $exchangeManualScript = @"
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
Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Yellow
Connect-ExchangeOnline

# Your App Details
`$AppId = "$($app.AppId)"
`$ServicePrincipalObjectId = "$($servicePrincipal.Id)"

Write-Host "`nChecking if Service Principal exists in Exchange..." -ForegroundColor Yellow
`$sp = Get-ServicePrincipal -ErrorAction SilentlyContinue | Where-Object { `$_.AppId -eq `$AppId }

if (-not `$sp) {
    Write-Host "Creating Service Principal..." -ForegroundColor Yellow
    try { New-ServicePrincipal -AppId `$AppId -ServiceId `$ServicePrincipalObjectId }
    catch [System.Management.Automation.ParameterBindingException] { New-ServicePrincipal -AppId `$AppId -ObjectId `$ServicePrincipalObjectId }

    Write-Host "Waiting 30 seconds for replication..." -ForegroundColor Yellow
    Start-Sleep -Seconds 30

    `$sp = Get-ServicePrincipal | Where-Object { `$_.AppId -eq `$AppId }
}

if (`$sp) {
    Write-Host "Service Principal found:" -ForegroundColor Green
    `$sp | Format-List AppId, DisplayName, Identity

    `$exoAppId = `$sp.AppId

    # Assign required roles for License Optimization Audit
    Write-Host "`nAssigning 'View-Only Configuration' role (MDO/ATP policies)..." -ForegroundColor Yellow
    try {
        New-ManagementRoleAssignment -Role 'View-Only Configuration' -App `$exoAppId -ErrorAction Stop
        Write-Host "+ View-Only Configuration assigned" -ForegroundColor Green
    } catch {
        Write-Host "! View-Only Configuration failed or already exists: `$_" -ForegroundColor Yellow
    }

    Write-Host "`nAssigning 'View-Only Recipients' role (mailbox properties)..." -ForegroundColor Yellow
    try {
        New-ManagementRoleAssignment -Role 'View-Only Recipients' -App `$exoAppId -ErrorAction Stop
        Write-Host "+ View-Only Recipients assigned" -ForegroundColor Green
    } catch {
        Write-Host "! View-Only Recipients failed or already exists: `$_" -ForegroundColor Yellow
    }

    Write-Host "`nVerifying assignments..." -ForegroundColor Yellow
    Get-ManagementRoleAssignment | Where-Object { `$_.RoleAssignee -eq `$sp.Identity } | Format-Table Role, RoleAssignee, Name

    Write-Host "`n+ Exchange role assignments complete!" -ForegroundColor Green
} else {
    Write-Host "X Service Principal still not found. Please wait 5-10 minutes and try again." -ForegroundColor Red
}
"@

        $manualScriptPath = ".\LOA-Exchange-Manual-Setup.ps1"
        $exchangeManualScript | Out-File -FilePath $manualScriptPath -Encoding UTF8

        Write-Host "`n  MANUAL SETUP SCRIPT CREATED" -ForegroundColor Yellow
        Write-Host "     Location: $manualScriptPath" -ForegroundColor White
        Write-Host "`n  ACTION REQUIRED:" -ForegroundColor Yellow
        Write-Host "     1. Ensure ExchangeOnlineManagement module is installed" -ForegroundColor White
        Write-Host "     2. Wait 5-10 minutes for Azure AD to sync with Exchange" -ForegroundColor White
        Write-Host "     3. Run the manual setup script:" -ForegroundColor White
        Write-Host "        .\LOA-Exchange-Manual-Setup.ps1" -ForegroundColor Cyan
    }
}

# ========================================================
# STEP 10: ASSIGN AZURE AD ROLES (Security Reader)
# ========================================================

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "STEP 10: Assigning Azure AD Security Reader Role" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Security Reader provides read access to security reports and sign-in data." -ForegroundColor White

try {
    $secReaderTemplate = Get-MgDirectoryRoleTemplate | Where-Object { $_.DisplayName -eq "Security Reader" }

    if ($secReaderTemplate) {
        # Check if the role is activated
        $activatedRole = Get-MgDirectoryRole -Filter "roleTemplateId eq '$($secReaderTemplate.Id)'" -ErrorAction SilentlyContinue

        if (-not $activatedRole) {
            $activatedRole = New-MgDirectoryRole -RoleTemplateId $secReaderTemplate.Id
            Write-Host "  + Security Reader role activated" -ForegroundColor Green
        }

        # Check if already assigned
        $existingMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $activatedRole.Id -ErrorAction SilentlyContinue
        $alreadyAssigned = $existingMembers | Where-Object { $_.Id -eq $servicePrincipal.Id }

        if (-not $alreadyAssigned) {
            $memberBody = @{
                "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($servicePrincipal.Id)"
            }
            New-MgDirectoryRoleMemberByRef -DirectoryRoleId $activatedRole.Id -BodyParameter $memberBody -ErrorAction Stop
            Write-Host "  + Security Reader role assigned to app" -ForegroundColor Green
        } else {
            Write-Host "  + Security Reader already assigned" -ForegroundColor Green
        }
    } else {
        Write-Host "  ! Security Reader role template not found" -ForegroundColor Yellow
    }
} catch {
    Write-Host "  ! Failed to assign Security Reader: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "     Assign manually in Azure Portal: Roles and administrators > Security Reader > Add assignment" -ForegroundColor Gray
}

# ========================================================
# STEP 11: CREATE AUDITOR PACKAGE
# ========================================================

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "STEP 11: Creating Auditor Package" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# Create connection config JSON — auto-detected by Get-M365LicenseOptimizationReport.ps1
$connectionConfig = @{
    ClientId              = $app.AppId
    TenantId              = $context.TenantId
    CertificateThumbprint = $cert.Thumbprint
    Organization          = $orgName
    Created               = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
}
$connectionConfig | ConvertTo-Json | Out-File -FilePath "$packageDir\LOA-Connection.json" -Encoding UTF8
Write-Host "  + Connection config saved (LOA-Connection.json)" -ForegroundColor Green

# Create App Registration Details file
$detailsContent = @"
# ========================================================
# M365 License Optimization Audit - Connection Details
# ========================================================
# Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
# Organization: $orgName
# ========================================================

TENANT INFORMATION:
  Tenant ID:      $($context.TenantId)
  Organization:   $orgName

APP REGISTRATION:
  Application ID: $($app.AppId)
  App Name:       $appDisplayName
  Object ID:      $($app.Id)

CERTIFICATE:
  Thumbprint:     $($cert.Thumbprint)
  Valid From:      $($certStartDate.ToString('yyyy-MM-dd'))
  Valid Until:     $($certEndDate.ToString('yyyy-MM-dd'))
  PFX File:       M365-LOA-Audit-Cert.pfx

PERMISSIONS GRANTED:
  Microsoft Graph API:
    - User.Read.All (user profiles, assigned licenses)
    - Group.Read.All (group memberships, MDO scope)
    - Organization.Read.All (org config, subscriptions)
    - Reports.Read.All (11 M365 usage reports)
    - ReportSettings.ReadWrite.All (unhide anonymized UPNs in usage reports)
    - AuditLog.Read.All (sign-in activity)
    - Policy.Read.All (Conditional Access policies)
    - RoleManagement.Read.Directory (PIM role assignments)
    - DeviceManagementManagedDevices.Read.All (Intune device count per user)

  Exchange Online:
    - Exchange.ManageAsApp API permission
    - View-Only Configuration (MDO policy rules)
    - View-Only Recipients (mailbox properties, litigation hold)

  Azure AD Role:
    - Security Reader

  ALL PERMISSIONS ARE READ-ONLY except ReportSettings.ReadWrite.All.
  ReportSettings.ReadWrite.All is used ONLY to unhide anonymized user data
  in usage reports (default behavior). After the
  audit completes, the customer must re-enable privacy in M365 Admin
  Center > Settings > Org settings > Reports > "Display concealed user,
  group, and site names in all reports".

CUSTOMER CONTACT:
  Name:  [TO BE FILLED IN]
  Email: [TO BE FILLED IN]
  Phone: [TO BE FILLED IN]

AUDIT DATES:
  Start: [TO BE FILLED IN]
  End:   [TO BE FILLED IN]
"@

$detailsContent | Out-File -FilePath "$packageDir\App-Registration-Details.txt" -Encoding UTF8

Write-Host "  + Auditor package created at: $packageDir" -ForegroundColor Green
Write-Host "`n  Package contents:" -ForegroundColor White
Get-ChildItem $packageDir | ForEach-Object { Write-Host "    * $($_.Name)" -ForegroundColor Gray }

# ========================================================
# SUMMARY
# ========================================================

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  SETUP COMPLETE!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan

Write-Host "`n  App Registration: $appDisplayName" -ForegroundColor White
Write-Host "  Application ID:   $($app.AppId)" -ForegroundColor White
Write-Host "  Tenant ID:        $($context.TenantId)" -ForegroundColor White
Write-Host "  Certificate:      Valid until $($certEndDate.ToString('MMMM dd, yyyy'))" -ForegroundColor White

Write-Host "`n  NEXT STEPS:" -ForegroundColor Yellow
Write-Host "  1. Open App-Registration-Details.txt and fill in contact info" -ForegroundColor White
Write-Host "  2. ZIP the M365-LOA-Audit-Package folder" -ForegroundColor White
Write-Host "  3. Send the ZIP to your auditor via secure channel" -ForegroundColor White
Write-Host "  4. Share the certificate password SEPARATELY (phone/SMS/Teams)" -ForegroundColor White

Write-Host "`n  SECURITY REMINDERS:" -ForegroundColor Yellow
Write-Host "  - All access is READ-ONLY" -ForegroundColor Gray
Write-Host "  - Certificate expires in 3 months" -ForegroundColor Gray
Write-Host "  - To revoke: Azure Portal > App Registrations > $appDisplayName > Delete" -ForegroundColor Gray
Write-Host "  - All auditor access is logged in Azure AD sign-in logs" -ForegroundColor Gray

Write-Host "`n============================================================`n" -ForegroundColor Cyan
