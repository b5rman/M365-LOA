# M365 License Optimization Audit — Permissions Guide

## What is this?

This document explains the permissions required by the **M365 License Optimization Audit (LOA)** app registration. The setup script (`LOA-App-Registration-Setup.ps1`) creates a read-only app registration in your Azure AD tenant so the auditor can analyse your Microsoft 365 license usage and provide optimization recommendations.

---

## Who runs the setup script?

**Your Global Administrator.** The script must be executed by someone with the Global Administrator role in your tenant. It will prompt for interactive sign-in — no credentials are stored or shared.

---

## How does authentication work?

The script generates a **self-signed certificate** (valid for 3 months) that the auditor uses to authenticate. No passwords or user credentials are involved.

| Item | Detail |
|---|---|
| Authentication method | Certificate-based (no passwords) |
| Certificate validity | 3 months from creation |
| Certificate key length | 2048-bit RSA, SHA-256 |
| Stored in | Your local certificate store (`Cert:\CurrentUser\My`) |

After the audit is complete, you can delete the app registration and certificate at any time to revoke all access immediately.

---

## Microsoft Graph API Permissions
morgen 
All Graph permissions are **Application** type (no user context) and **read-only** unless noted.

| Permission | Type | What it reads | Why the audit needs it |
|---|---|---|---|
| `User.Read.All` | Read | User profiles, assigned licenses, sign-in state, account status | Identify which users have which licenses, detect disabled/dormant accounts |
| `Group.Read.All` | Read | Group memberships, license groups | Determine group-based license assignments, resolve MDO/Conditional Access policy scope |
| `Organization.Read.All` | Read | Tenant configuration, subscribed SKUs | Inventory all purchased licenses and available units |
| `Reports.Read.All` | Read | M365 usage reports (Exchange, Teams, OneDrive, SharePoint, M365 Apps) | Measure actual usage per user — the core of license optimization |
| `ReportSettings.ReadWrite.All` | **Read/Write** | Report privacy setting only | Temporarily unhides anonymized user names in usage reports so the audit can match usage to users (*see note below*) |
| `AuditLog.Read.All` | Read | Sign-in activity logs | Detect last sign-in date per user (interactive and non-interactive) to identify dormant accounts |
| `Policy.Read.All` | Read | Conditional Access policies | Detect risk-based Conditional Access policies that may affect license recommendations |
| `RoleManagement.Read.Directory` | Read | Admin role assignments (PIM eligible/active) | Identify admin accounts for dormant admin detection and license compliance |
| `DeviceManagementManagedDevices.Read.All` | Read | Intune enrolled device count per user | Detect Intune license usage (shelfware detection) |

### Note on `ReportSettings.ReadWrite.All`

This is the only non-read permission. It is used to **unhide user principal names** in Microsoft 365 usage reports. By default, Microsoft anonymizes user names in reports for privacy. The audit temporarily disables this setting so it can match usage data to specific users.

**After the audit**, you should re-enable the privacy setting:
> **M365 Admin Center** > **Settings** > **Org settings** > **Reports** > Enable *"Display concealed user, group, and site names in all reports"*

---

## Exchange Online Permissions

Exchange Online access is configured separately from Graph and uses its own permission model.

| Permission / Role | Type | What it reads | Why the audit needs it |
|---|---|---|---|
| `Exchange.ManageAsApp` | API permission | Allows the app to connect to Exchange Online | Required for any Exchange cmdlet access via certificate auth |
| `View-Only Recipients` | Management role | Mailbox properties: type (User/Shared/Room), litigation hold status, archive status | Detect shared mailboxes with paid licenses, litigation hold dependencies before recommending license removal |
| `View-Only Configuration` | Management role | Exchange transport rules, MDO/ATP policies (Safe Links, Safe Attachments) | Identify which users are covered by Microsoft Defender for Office 365 policies to avoid removing needed licenses |

---

## What the audit does NOT have access to

- Email content, attachments, or message bodies
- File contents in OneDrive or SharePoint
- Teams chat messages or call recordings
- User passwords or authentication credentials
- Write access to any user, group, or mailbox data
- Ability to modify licenses, policies, or configurations
- Calendar contents or contact details

The audit reads **metadata and usage statistics only** — never the content of communications or files.

---

## Security controls

| Control | Detail |
|---|---|
| **Read-only access** | All permissions except `ReportSettings.ReadWrite.All` are strictly read-only |
| **Certificate authentication** | No passwords involved; certificate expires automatically after 3 months |
| **Time-limited** | Certificate validity is 3 months — access expires automatically |
| **Full audit trail** | All access is logged in Azure AD sign-in logs under the app registration name |
| **Instant revocation** | Delete the app registration in Azure Portal to revoke all access immediately |
| **No data export** | The app cannot export, modify, or delete any tenant data |

---

## How to revoke access

You can revoke auditor access at any time using either method:

### Option 1: Delete the app registration (recommended)
1. Go to **Azure Portal** > **Azure Active Directory** > **App registrations**
2. Search for **"M365 License Optimization Audit"**
3. Click **Delete**

### Option 2: Delete the certificate
1. Go to **Azure Portal** > **Azure Active Directory** > **App registrations**
2. Open **"M365 License Optimization Audit"** > **Certificates & secrets**
3. Remove the certificate

Either option immediately prevents any further access.

---

## Questions?

If you have any questions about the permissions or the audit process, please contact your auditor before running the setup script.
