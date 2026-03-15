# LOA App Registration Setup

Customer-side PowerShell script that creates secure, read-only auditor access for the
M365 License Optimization Report. Run by the **customer's Global Administrator** — not
the auditor.

## What It Does

1. Creates an App Registration (`M365 License Optimization Audit`)
2. Generates a self-signed certificate (3-month validity)
3. Assigns read-only Microsoft Graph API permissions
4. Grants admin consent automatically
5. Configures Exchange Online RBAC roles (View-Only)
6. Assigns the Security Reader Azure AD role
7. Produces a ready-to-send auditor package

## Prerequisites

| Requirement | Details |
|-------------|---------|
| **Role** | Global Administrator on the target tenant |
| **OS** | Windows 10/11 or Windows Server 2016+ |
| **PowerShell** | 5.1+ (Desktop) or 7+ (Core) |
| **Modules** | Installed automatically: `Microsoft.Graph.Authentication`, `Microsoft.Graph.Applications`, `ExchangeOnlineManagement` |

## Usage

```powershell
.\LOA-App-Registration-Setup.ps1
```

The script is fully interactive — it prompts for a certificate password, confirms before
proceeding, and asks whether to configure Exchange Online access.

## Permissions Granted

### Microsoft Graph API (Application)

| Permission | Purpose |
|------------|---------|
| `User.Read.All` | User profiles, assigned licenses, account state |
| `Group.Read.All` | Group memberships, license groups, MDO/CA scope |
| `Organization.ReadWrite.All` | Org config + temporarily unhide anonymized usage report data |
| `Reports.Read.All` | 11 M365 usage reports (Email, Teams, OneDrive, SharePoint, Apps) |
| `AuditLog.Read.All` | Sign-in activity (last interactive/non-interactive) |
| `Policy.Read.All` | Conditional Access policies |
| `RoleManagement.Read.Directory` | PIM eligible/active role assignments |
| `DeviceManagementManagedDevices.Read.All` | Enrolled device count per user (Intune shelfware detection) |

> All permissions are read-only except `Organization.ReadWrite.All`, which is used solely
> to temporarily unhide anonymized user data in usage reports.
>
> **Note:** `CloudLicensing.Read.All` (subscription lifecycle / trial detection) is not
> included because the app role is not registered in all tenants. The script degrades
> gracefully without it. Add manually in Azure Portal if your tenant supports it.

### Exchange Online

| Permission / Role | Purpose |
|-------------------|---------|
| `Exchange.ManageAsApp` | API permission for certificate-based EXO connection |
| `View-Only Configuration` | MDO/ATP policy rules (Safe Links, Safe Attachments, Preset) |
| `View-Only Recipients` | Mailbox type, Litigation Hold, Archive status, Forwarding rules |

### Azure AD Role

| Role | Purpose |
|------|---------|
| `Security Reader` | Read access to security reports and sign-in data |

## Output

The script creates a `M365-LOA-Audit-Package/` folder containing:

| File | Description |
|------|-------------|
| `M365-LOA-Audit-Cert.pfx` | Private certificate (password-protected) |
| `M365-LOA-Audit-Cert-Public.cer` | Public certificate (reference only) |
| `LOA-Connection.json` | Connection config (ClientId, TenantId, Thumbprint) |
| `App-Registration-Details.txt` | Full permission summary + contact fields |
| `Quick-Start-Guide.txt` | Step-by-step instructions for the auditor |

The audit script auto-detects `LOA-Connection.json` from the script root,
current directory, or the `M365-LOA-Audit-Package/` subfolder in either location.

## Sending the Package to Your Auditor

1. Fill in the contact and audit date fields in `App-Registration-Details.txt`
2. ZIP the `M365-LOA-Audit-Package/` folder
3. Send the ZIP via a secure channel (e.g., encrypted email, SharePoint link)
4. Share the certificate password **separately** (phone, SMS, or Teams call)

## Revoking Access

Access can be revoked at any time:

- **Azure Portal** > App Registrations > `M365 License Optimization Audit` > Delete
- The certificate expires automatically after 3 months
- All auditor access is logged in Azure AD sign-in logs

## Troubleshooting

### Exchange RBAC setup fails

If the automatic Exchange configuration fails (common when Azure AD hasn't finished
replicating the Service Principal to Exchange), the script generates a
`LOA-Exchange-Manual-Setup.ps1` fallback script. Wait 5-10 minutes, then run it.

### Admin consent partially fails

If some Graph permissions fail to get consent (replication lag), the script retries
automatically after 10 seconds. If it still fails, grant consent manually:

**Azure Portal** > App Registrations > `M365 License Optimization Audit` > API
Permissions > Grant admin consent

### Module version mismatch

All `Microsoft.Graph.*` sub-modules must be the same version. If you see assembly-load
errors, update all Graph modules:

```powershell
Update-Module Microsoft.Graph.Authentication, Microsoft.Graph.Applications -Force
```
