# Logic Flaws — Tracking

> Active flaws/bugs to fix. Move to "Implemented" section once done.

## Open

*(none — all caught up)*

## Implemented

1. **"No Desktop Apps" False Positive** — Flagged users on web-only SKUs (e.g., Business Basic) to "consider web-only license" — redundant. Fixed: gated on `$hasDesktopAppEntitlement` via `Get-PlanCapabilities`.

2. **Audio Conferencing Price Inflation** — MCOMEETADV priced at €2.50 but became free in 2023. Fixed: set to €0.00.

3. **Teams Premium Shelfware Threshold** — Was `-eq 0`, too tight for €10/mo SKU. Fixed: changed to `-lt 3`.

4. **EOA Overlap Missing from `$suiteIncludes`** — Exchange Online Archiving wasn't listed as a component of E3/E5 suites, so duplicate detection engine missed standalone EOA overlap. Fixed: added `EXCHANGE_ARCHIVE` to all relevant `$suiteIncludes` entries.

5. **Power Platform False NO ACTIVITY** — Power Automate/Power Apps run server-side without M365 app telemetry. Fixed: appended `POWER PLATFORM REVIEW` caveat to NO ACTIVITY (drops confidence to Review).

6. **Legacy Auth / Service Account Waste** — POP3/IMAP4/SMTP-only users on premium suites. Fixed: added `LEGACY SERVICE ACCOUNT` detection.

7. **Frontline "Multiple PC" Blocker** — Frontline downgrade recommended for users with 2+ Windows PC activations (F3 VDI-only would break them). Fixed: added `$winActTotal > 1` gate → `FRONTLINE BLOCKED`.

8. **False Dormancy of Admin Accounts** — Admin flagged as dormant when non-interactive sign-in is recent (automation/service account). Fixed: reads `lastNonInteractiveSignInDateTime`, emits `AUTOMATION ACCOUNT` instead of `DORMANT ADMIN RISK`.

9. **E5 Step-Up vs. Shelfware Collision** — E5 upgrade recommended based on add-on cost sum, but add-ons might be shelfware. Fixed: appended shelfware validation caveat to E5 UPGRADE OPPORTUNITY message.

10. **Copilot "False Dormancy" Trap** — Copilot Chat (copilot.microsoft.com) doesn't generate standard app telemetry. Fixed: updated Copilot inactive message + added `COPILOT REVIEW` caveat to NO ACTIVITY block.

11. **"Ghost in the Machine" (Passive OneDrive Sync)** — Abandoned laptop syncing in background counts as activity, hiding ghost users. Fixed: added `BACKGROUND SYNC ONLY` detection for sync-only users with zero interactive activity.

12. **Teams Premium "Organizer" Illusion** — Shelfware check used total `Meeting Count` (organized + attended). Premium features are organizer-driven. Fixed: now uses `Meetings Organized Count`.

13. **Unlicensed Mailbox Time Bomb** — Script warned "data purged in 30 days" for ALL unlicensed mailboxes. Litigation Hold mailboxes silently become free Inactive Mailboxes. Fixed: split into `INACTIVE MAILBOX (FREE)` vs `UNLICENSED WITH DATA`.
