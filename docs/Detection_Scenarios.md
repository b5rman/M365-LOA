# M365 License Optimization — Detection Scenarios

> Every detection in the script, explained with a real-world scenario and a name you'll actually remember.

---

## TIER 1 — Pure Waste (Remove License Immediately)

### The Ghost Ship
**Tag:** `DORMANT` | **Confidence:** High

A user hasn't interactively signed in for 90+ days (configurable). Their license is drifting in the ocean, burning money. The strongest signal for license removal — no human has touched this account in months.

**Scenario:** Sarah from Marketing left on maternity leave 4 months ago. Her E3 license at €36/mo is still active. That's €432/yr floating into the void.

---

### The Walking Dead
**Tag:** `DELETED USER` | **Confidence:** High

The account is in the Azure AD recycle bin (soft-deleted), but the license is still assigned. Microsoft keeps soft-deleted users for 30 days. The license is 100% wasted.

**Scenario:** James was terminated last Tuesday. IT deleted his account but forgot to reclaim the E5 license. That's €57/mo burning until someone notices.

---

### The Locked Door
**Tag:** `DISABLED ACCOUNT` | **Confidence:** High

Sign-in is blocked, but a paid license is still attached. The user literally cannot use the service. Pure waste.

**Scenario:** A contractor's account was disabled after their project ended, but their M365 Business Premium license was never removed. €22.60/mo for a door nobody can open.

---

### The Empty Office
**Tag:** `NO ACTIVITY` | **Confidence:** High

The user can sign in, but hasn't touched Exchange, Teams, OneDrive, SharePoint, or any M365 app in the entire reporting period. Strong candidate for removal.

**Scenario:** A legacy service account created for a project that ended 2 years ago. Still has an E3 license. Zero emails, zero Teams chats, zero files. €432/yr for an empty chair.

---

### The Double Agent
**Tag:** `OVERLAPPING LICENSE` | **Confidence:** High

The same SKU is assigned both directly AND via a group. One of them is pure waste — the user only needs one path.

**Scenario:** IT assigned E3 directly to Bob. Then Bob's department got added to an "E3 Users" group. Now Bob has two E3 assignments. One is invisible waste.

---

### The Russian Doll
**Tag:** `DUPLICATE COVERAGE` | **Confidence:** High

A user has a suite (e.g., M365 E3) AND a standalone SKU that the suite already includes (e.g., Exchange Online Plan 2). The standalone is completely redundant.

**Scenario:** Maria has M365 E5 (which includes everything) plus a standalone Exchange Online Plan 2 (€8/mo). The E5 already covers Exchange Enterprise. That €96/yr is a nesting doll of waste.

---

### The Gate Crasher
**Tag:** `GUEST ACCOUNT WASTE` | **Confidence:** High

An external/guest user has a paid license assigned. Guest users typically shouldn't have productivity licenses — they access shared resources via their home tenant.

**Scenario:** A vendor's guest account was accidentally included in a group-based licensing policy. Now they have an E3 license they'll never use. €432/yr for a gate crasher with a VIP badge.

---

### The Impersonator
**Tag:** `NON-HUMAN ACCOUNT WASTE` | **Confidence:** High

A shared mailbox, room mailbox, or equipment mailbox is sitting on a premium suite license. These accounts don't need E3/E5 — shared mailboxes under 50 GB need no license at all.

**Scenario:** The "Conference Room A" room mailbox somehow got an M365 E3 license. A room doesn't need Teams, OneDrive, or desktop Office apps. That's €432/yr for a room that can't read email.

---

### The Zombie Add-On
**Tag:** `INTUNE SUITE WASTE` | **Confidence:** High

Standalone Intune add-ons (Remote Help, Advanced Analytics, EPM) assigned alongside M365 E3/E5. Microsoft rolled these into E3/E5 in late 2025. The add-ons are now the walking dead.

**Scenario:** The IT team bought Intune Remote Help add-ons in 2024. Microsoft included them in E3/E5 bundles in late 2025. Nobody removed the now-redundant add-ons. Ghost money.

---

### The Wrong Platform
**Tag:** `WINDOWS LICENSE WASTE` | **Confidence:** Medium

A standalone Windows Enterprise E3/E5 license is assigned to a user who only uses Mac or mobile devices. Windows licensing for someone without a Windows device.

**Scenario:** A designer uses a MacBook Pro exclusively. They were assigned Windows E3 (€7/mo) as part of a blanket IT policy. They'll never boot Windows. €84/yr for a platform they don't own.

---

### The Empty Vault
**Tag:** `OVER-LICENSED ARCHIVE` | **Confidence:** High

Exchange Online Archiving add-on assigned to a user with a small mailbox (<25 GB) and no active archive. They're paying for vault space they never opened.

**Scenario:** A junior employee got an EOA license (€3/mo) during a department-wide rollout. Their mailbox is 800 MB and they've never enabled an archive. €36/yr for an empty vault.

---

### The Double Archive
**Tag:** `REDUNDANT ARCHIVE` | **Confidence:** High

A shared mailbox has both Exchange Plan 2 and standalone Exchange Online Archiving. Plan 2 natively includes auto-expanding archives — the EOA add-on is 100% redundant.

**Scenario:** A shared mailbox hits 95 GB. A junior admin panics and buys standalone EOA (€3/mo) to prevent data loss. But Exchange Plan 2 (€8/mo) already includes auto-expanding archives. That €3/mo add-on is pure waste — €36/yr per shared mailbox. In a tenant with 50 shared mailboxes, that's €1,800/yr down the drain.

---

### The Free Lunch
**Tag:** `INACTIVE MAILBOX (FREE)` | **Confidence:** High

An unlicensed mailbox on Litigation Hold. Microsoft silently converts it to a free "Inactive Mailbox" for eDiscovery. No license needed — the data is preserved at zero cost. Don't let a panicked admin re-license it.

**Scenario:** A departed employee's mailbox is on legal hold for an ongoing investigation. Their license was removed. A junior admin sees "unlicensed with data" and panics, re-assigning an E3. That €432/yr is unnecessary — the mailbox is safely archived for free.

---

## TIER 2 — Right-Sizing (Downgrade SKU)

### The Frontline Candidate
**Tag:** `FRONTLINE CANDIDATE` | **Confidence:** Medium (High when zero activations)

A user on an E3/E5 premium suite who only uses web/mobile apps — no desktop Office. They don't need the desktop app entitlement. Downgrade to F1/F3.

**Scenario:** A warehouse worker has E3 (€36/mo) but only checks Teams on their phone and reads email via OWA. F3 (€8/mo) gives them everything they need. Savings: €336/yr per user.

---

### The Device-less Wonder
**Tag:** `FRONTLINE CANDIDATE (HIGH CONFIDENCE)` | **Confidence:** High

The strongest frontline downgrade signal possible. The user has a premium suite, uses only web/mobile, AND has zero device activations ever — they've literally never installed Office on any PC or Mac.

**Scenario:** Same warehouse worker as above, but we can also prove they've never activated Office on any hardware device. Zero ambiguity. This is the slam-dunk downgrade.

---

### The Immovable Object
**Tag:** `FRONTLINE BLOCKED` | **Confidence:** High

The user fits the frontline profile (web/mobile only), but something blocks the downgrade: an archive mailbox (F3 has zero archive rights — data would be destroyed) or multiple Windows PC activations (F3 is VDI-only).

**Scenario:** A field worker uses only OWA on her phone, but she has a 15 GB archive mailbox from years of email. Downgrading to F3 would permanently destroy her archive. The system flags it but blocks the recommendation.

---

### The Frankenstein Frontline
**Tag:** `FRONTLINE ADD-ON BLOAT` | **Confidence:** Medium

An F3 user (€8/mo) has been bolted together with expensive add-ons: Exchange Plan 2, Entra P2, Power BI Pro. The total exceeds Business Premium or E3. The monster costs more than the full suite.

**Scenario:** A shift manager started on F3. Local IT added Exchange Plan 2 (€8), Entra P2 (€9), and PBI Pro (€9.40). Total: €34.40/mo. An E3 at €36.20/mo would remove F3's 2 GB storage cap and 10.9" screen limit — for roughly the same price.

---

### The Overachiever
**Tag:** `BUSINESS BASIC CANDIDATE` | **Confidence:** Medium

A user on Business Standard who only uses web/mobile apps. They're paying for desktop Office they never launch. Business Basic covers web/mobile at a lower price.

**Scenario:** A part-time consultant has Business Standard (€11.70/mo) but only ever uses Word and Excel in the browser. Business Basic (€5.60/mo) gives them everything they need. Savings: €73/yr.

---

### The Desktop Mirage
**Tag:** `STANDALONE APPS WASTE` | **Confidence:** Medium

A user has a standalone desktop apps license (M365 Apps for Enterprise/Business) but only uses web/mobile versions. The desktop app investment is a mirage.

**Scenario:** A temp worker was given M365 Apps (€12.90/mo) for desktop Office, but they only use the web versions from a shared kiosk. Downgrade to Business Basic or F3.

---

### The Kiosk Frankenstein
**Tag:** `A LA CARTE WASTE` | **Confidence:** High

Exchange Kiosk (€1/mo, 2 GB mailbox) + standalone M365 Apps (€13.90/mo) costs more than M365 Business Standard (€12.50/mo) — which gives 50 GB mailbox, 1 TB OneDrive, and Teams desktop. A poorly combined a-la-carte stack that's both more expensive AND worse.

**Scenario:** A sales rep gets Exchange Kiosk for cheap email, then their manager adds M365 Apps for desktop Word/Excel. Combined: €14.90/mo. Business Standard at €12.50/mo would save money AND give them a real mailbox instead of a 2 GB toy. That's €28.80/yr savings per user, plus a massive UX upgrade.

---

### The Frankenstein Stack
**Tag:** `BUNDLE INEFFICIENCY` | **Confidence:** High

Business Basic (€6/mo) + Apps for Business (€11.50/mo) bolted together by different admins at different times. Business Standard (€12.50/mo) natively includes both — for less money.

**Scenario:** A new hire gets Business Basic for Teams and email. Six months later, they request desktop Word and Excel, so someone adds Apps for Business. Now they cost €17.50/mo. Business Standard at €12.50/mo would save €60/yr per user and simplify the license assignment to a single SKU.

---

### The Phantom Desktop
**Tag:** `PREMIUM ADD-ON WASTE` | **Confidence:** Medium

A user has Visio Plan 2 or Project Plan 3/5 (expensive desktop-tier SKUs) but their product-specific activation data shows zero Windows/Mac desktop activations — they only access via mobile or web. The web-only Plan 1 equivalent costs 60–75% less.

**Scenario:** An architect has Visio Plan 2 (€14.60/mo) but only views Visio diagrams in the browser (no Visio desktop installed). Visio Plan 1 (€4.70/mo) gives the same web access. That's €118.80/yr per user saved. For Project Plan 3 → Essentials, the savings are €267.60/yr.

---

### The Lobby Phone Tax
**Tag:** `TEAMS PHONE RIGHT-SIZING` | **Confidence:** High

A shared mailbox, room, or equipment account has Teams Phone Standard (€8/mo) — designed for human users. Non-human endpoints (lobby phones, conference room devices) only need Teams Shared Devices (€2.50/mo).

**Scenario:** An office has 20 conference rooms, each with a Room Mailbox and Teams Phone Standard for the conference phone. That's 20 × €8 = €160/mo. Switching to Teams Shared Devices: 20 × €2.50 = €50/mo. Annual savings: €1,320.

---

### The Micro-Saver
**Tag:** `F3 TO F1 DOWNGRADE` | **Confidence:** Medium

An F3 user (€8/mo) with zero email activity, zero OneDrive activity, and empty storage. F1 (€2.25/mo) gives them Teams + SharePoint read access — which is all they're using.

**Scenario:** A factory floor worker has F3 but only uses Teams for shift announcements. Their 2 GB mailbox is empty, OneDrive is untouched. F1 saves €69/yr per user. At 500 frontline workers, that's €34,500/yr.

---

### The Silent Quitter
**Tag:** `TEAMS UNBUNDLING` | **Confidence:** Medium

A user on a bundled suite (M365 E3) with zero Teams activity. Microsoft now offers "without Teams" SKUs at a lower price in the EU.

**Scenario:** A data analyst has E3 and uses Excel and SharePoint all day but has literally never opened Teams (0 chats, 0 calls, 0 meetings). The "M365 E3 without Teams" SKU saves €2/mo.

---

### The EXO Downgrade
**Tag:** `EXO PLAN 2 DOWNGRADE` | **Confidence:** Medium

A user on Exchange Online Plan 2 (100 GB, €8/mo) whose mailbox is under 50 GB and has no archive. Plan 1 (€3.70/mo) would suffice.

**Scenario:** A user has EXO Plan 2 for the 100 GB mailbox, but their mailbox is only 12 GB with no archive enabled. Plan 1's 50 GB limit is more than enough. Savings: €51.60/yr.

---

### The False Economy
**Tag:** `E5 UPGRADE OPPORTUNITY` | **Confidence:** Medium

A user has E3 plus multiple E5-included add-ons (Defender, Purview, etc.) that combined cost more than E5. BUT — the recommendation now warns to verify the add-ons are actually used. If they're shelfware, removing them from E3 is cheaper than upgrading.

**Scenario:** A user has E3 + Defender P2 + Purview + PBI Pro = €78/mo. E5 is €57/mo — looks like a €252/yr save. But if they have zero Power BI reports and zero eDiscovery cases, those add-ons are shelfware. Remove them: E3 at €36/mo. The "upgrade" would have locked in €684/yr permanently.

---

### The Transformer
**Tag:** `BUNDLE CONSOLIDATION` | **Confidence:** High

A user has O365 + EMS + Windows Enterprise as separate SKUs. Combining them into a single M365 E3/E5 bundle is cheaper and simpler.

**Scenario:** A user has O365 E3 (€22/mo) + EMS E3 (€8.80/mo) + Windows E3 (€7/mo) = €37.80/mo. M365 E3 is €36.20/mo. Same features, one SKU, €19/yr savings.

---

## ACTIVITY & BEHAVIORAL DETECTIONS

### The Ghost in the Machine
**Tag:** `BACKGROUND SYNC ONLY` | **Confidence:** High

Zero interactive activity (no emails, chats, meetings, files viewed), but OneDrive shows hundreds of synced files. A powered-on laptop in a drawer is silently syncing SharePoint libraries in the background. The user is gone — the machine isn't.

**Scenario:** An employee left 3 months ago. Their laptop is still on a shelf in the office, connected to Wi-Fi. OneDrive synced 847 files this quarter. The script saw "activity" and thought they were alive. This detection exposes the ghost.

---

### The Digital Hoarder
**Tag:** `EXPENSIVE COLD STORAGE` | **Confidence:** High

Zero activity, but the mailbox is >10 GB or OneDrive is >50 GB. The organization is paying E3/E5 prices just to store data nobody accesses. Convert to Shared Mailbox / SharePoint Archive and strip the license.

**Scenario:** A departed director's account has 80 GB of email and 400 GB of OneDrive files. Zero activity for 6 months. Nobody wants to delete it "just in case." That's €432/yr to store data that could live in a free Shared Mailbox + SharePoint.

---

### The Dinosaur
**Tag:** `LEGACY SERVICE ACCOUNT` | **Confidence:** High

Email activity detected, but ONLY via POP3, IMAP4, or SMTP — no Outlook Desktop, OWA, or Mobile. This is almost certainly a scan-to-email device or a legacy script, not a human. They don't need an E3.

**Scenario:** A multifunction printer uses a service account with E3 (€36/mo) to send scanned documents via SMTP. It doesn't need Teams, OneDrive, or desktop Office. Exchange Plan 1 (€3.70/mo) covers it. Savings: €387/yr. Security bonus: legacy protocols bypass most Conditional Access.

---

### The Robot in Disguise
**Tag:** `AUTOMATION ACCOUNT` | **Confidence:** High

An admin account appears dormant (no interactive sign-in for 90+ days), but has recent non-interactive sign-ins. This isn't an abandoned admin — it's a service/automation account running scripts or scheduled tasks behind the scenes. Don't panic and disable it; investigate its purpose.

**Scenario:** A Global Admin account hasn't logged in interactively for 120 days. Old logic would scream "DORMANT ADMIN RISK!" But non-interactive sign-in was yesterday — it's running a nightly Azure Automation script. The fix isn't "disable immediately" — it's "convert to a Workload Identity."

---

### The Ticking Time Bomb
**Tag:** `DORMANT ADMIN RISK` | **Confidence:** High

An admin account with no interactive AND no non-interactive sign-in for 90+ days. This is a genuine security risk — dormant admin accounts are prime targets for credential stuffing and compromise.

**Scenario:** A former IT contractor's Global Admin account hasn't been used in 6 months — interactively or programmatically. It's sitting there with full tenant access, waiting for an attacker to find it. Disable immediately, reclaim license, audit for unauthorized activity.

---

### The Invisible Pilot
**Tag:** `COPILOT` (inactive variant) | **Confidence:** Review

A Copilot license holder shows zero standard M365 activity. BUT — web-based Copilot Chat (copilot.microsoft.com) doesn't generate telemetry in Exchange, Teams, or M365 App reports. The executive might be using Copilot Chat every single day.

**Scenario:** The CFO has a €30/mo Copilot license and zero Teams/Email/Office activity. Standard logic says "remove it." But she uses Copilot Chat for 2 hours daily via the browser. Check the Copilot usage dashboard before pulling the trigger.

---

### The Phantom Caller
**Tag:** `TEAMS PHONE REVIEW` | **Confidence:** Review

A user has a Teams Phone System SKU but no Microsoft Calling Plan. They might use Direct Routing or Operator Connect (which wouldn't show as a Calling Plan SKU). Needs manual verification.

**Scenario:** A user has Teams Phone Standard but no PSTN Calling Plan visible in licensing. They might have Direct Routing through the company SBC. Or the Phone System is genuinely unused. Only the admin knows.

---

## SHELFWARE & ADD-ON WASTE

### The Shelf Ornament
**Tag:** `SHELFWARE` | **Confidence:** Medium

An expensive standalone license (Visio, Project, Power BI Pro, Teams Premium) with zero detected usage activity. For Teams Premium, this specifically checks **meetings organized** — not attended — because Premium features are organizer-driven.

**Scenario:** 50 users have Visio Plan 2 (€14.60/mo) but none have opened Visio in 90 days. That's €8,760/yr of shelf ornaments. For Teams Premium: a user attended 30 meetings but organized zero. The €10/mo license only benefits organizers.

---

### The AI Collision
**Tag:** `AI OVERLAP REVIEW` | **Confidence:** Review

A user has both Teams Premium AND Microsoft 365 Copilot. Copilot natively includes Teams Intelligent Recap (AI meeting notes/tasks), making the Teams Premium recap feature redundant.

**Scenario:** A manager has Teams Premium (€10/mo) and Copilot (€30/mo). Copilot already provides AI meeting summaries. Unless the manager specifically needs advanced webinar branding or custom meeting templates, Teams Premium is the redundant one. Savings: €120/yr.

---

### The Matryoshka
**Tag:** `ENTRA SUITE OVERLAP` | **Confidence:** High

Standalone Entra P2 or Governance add-ons assigned alongside a full Entra Suite license. The suite already includes everything — the standalones are nesting dolls.

**Scenario:** A user has the Entra Suite (€12/mo) plus standalone Entra ID P2 (€9/mo). The Suite already includes P2. That €108/yr is a matryoshka of licensing redundancy.

---

### The Virus
**Tag:** `VIRAL LICENSE CLEANUP` | **Confidence:** Medium

Free trial/viral SKUs (Teams Exploratory, Power Apps Viral, Flow P2 Viral) lurking alongside paid productivity suites. They create service plan conflicts during provisioning and should be cleaned up.

**Scenario:** A user has M365 E3 plus a "Teams Exploratory" viral license that self-provisioned when they first clicked Teams. The E3 already includes Teams. The viral SKU is just clutter that can cause group-based licensing conflicts.

---

## SECURITY & COMPLIANCE FLAGS

### The Leaky Bucket
**Tag:** `HIGH RISK SHARING` | **Confidence:** Medium

A user shared 25+ files externally in the reporting period without any Purview/DLP coverage. Data leakage risk without compliance controls.

**Scenario:** A sales rep shared 47 files externally via OneDrive — client proposals, pricing sheets, contracts. No DLP policies, no Information Protection labels, no Purview coverage. One accidental share of the customer database and it's front-page news.

---

### The Phantom Bodyguard
**Tag:** `MDM/MAM WASTE` | **Confidence:** Medium

A user has an Intune/EMS entitlement but telemetry shows 100% web-only access (no desktop apps, no mobile apps, no Teams desktop/mobile). Intune manages devices and apps — there's nothing to manage for a browser-only user.

**Scenario:** A contractor accesses M365 exclusively through Chrome on a personal laptop. They have EMS E3 (€8.80/mo) for "security." But Intune can't manage their personal browser. Entra ID P1 with Conditional Access is all they need.

---

### The Security Gap
**Tag:** `SECURITY GAP` / `DEFENDER SUITE UPSELL` / `PURVIEW UPSELL` | **Confidence:** Medium

Users on Business or E3 plans missing advanced security or compliance capabilities that their usage profile suggests they need.

**Scenario:** A finance team on E3 handles sensitive financial data daily with no Purview DLP, no Information Protection labels, and no eDiscovery Premium. The compliance risk far exceeds the add-on cost.

---

## STORAGE & INFRASTRUCTURE

### The Overflowing Inbox
**Tag:** `MAILBOX STORAGE WARNING` | **Confidence:** Medium

A mailbox on a plan with a storage ceiling (Exchange Plan 1 = 50 GB, Plan 2 = 100 GB) is approaching that limit. Mail flow stops when the cap is hit.

**Scenario:** A user on Exchange Plan 1 has a 48 GB mailbox (96% of the 50 GB limit). When it hits 50 GB, they can't send or receive. Upgrade to Plan 2 or enable archiving before the inbox explodes.

---

### The OneDrive Cliff
**Tag:** `ONEDRIVE STORAGE WARNING` | **Confidence:** Medium

OneDrive on a Business/E1 plan is approaching the hard 1 TB limit. Sync will break catastrophically when the limit is reached.

**Scenario:** A content creator on Business Standard has 950 GB of video files in OneDrive. Business plans hard-cap at 1 TB. At their growth rate, sync will break in 2 weeks.

---

### The Unlicensed Time Bomb
**Tag:** `UNLICENSED WITH DATA` | **Confidence:** High

An unlicensed user account with mailbox or OneDrive data still present. Microsoft purges unlicensed data after 30 days. This is a countdown timer.

**Scenario:** An employee was offboarded and their license removed, but their 25 GB mailbox and 100 GB OneDrive are still there. In 30 days, it's gone. Back up, convert to Shared Mailbox, or re-license.

---

## ADMINISTRATIVE & COMPLIANCE

### The Trial Balloon
**Tag:** `TRIAL LICENSE` | **Confidence:** High

A user is on a trial subscription that will expire. Plan conversion to paid or removal before expiry.

**Scenario:** IT started a Power BI Pro trial for 25 users to evaluate. The trial expires in 12 days. If nobody converts or removes it, 25 users lose access overnight.

---

### The Waiting Room
**Tag:** `LICENSE CAPACITY QUEUE` | **Confidence:** High

A user is in the license allotment waiting room because there aren't enough seats. They need a license but can't get one.

**Scenario:** The company bought 500 E3 licenses but has 503 users in the licensing group. Three users are in the queue, unable to use their assigned services. Buy 3 more seats or remove 3 inactive assignments.

---

### The Broken Pipe
**Tag:** `CLOUD LICENSE SYNC ERROR` | **Confidence:** High

Group-based licensing assignment failed. The user's license state is inconsistent. Could be insufficient seats, service plan conflicts, or dependency issues.

**Scenario:** A user was added to both the "E3 Users" and "E1 Users" groups. The conflicting assignments create an error state. Neither license is fully applied. The user experiences random service outages.

---

### The Litigation Shield
**Tag:** `LITIGATION HOLD` | **Confidence:** High

An active Litigation Hold is on the mailbox. The license MUST be retained while the hold is active — removing it could constitute spoliation of evidence.

**Scenario:** The Legal department placed a hold on 15 mailboxes for an ongoing lawsuit. Even if those users have left the company, their licenses cannot be removed until Legal lifts the hold.

---

## DATA QUALITY FLAGS

### The Fog of War
**Tag:** `DATA GAP` | **Confidence:** Review

One or more SKUs on this user aren't in the reference data. Cost calculations and right-sizing recommendations may be incomplete.

**Scenario:** A user has a brand-new "Microsoft 365 Copilot for Sales" SKU that isn't in the script's pricing table yet. Their total cost shows €0 for that SKU and no recommendations fire for it. Update the reference data.

---

### The Never-Was
**Tag:** `NEVER SIGNED IN` | **Confidence:** High

A licensed user with no interactive sign-in record at all. They were provisioned but may have never actually used the service.

**Scenario:** A new hire was assigned an E3 license on their first day 2 months ago, but they're still in training and haven't set up their laptop. Or the account was created for a project that never started.

---

### The Price Inversion
**Tag:** `SUITE INVERSION` | **Confidence:** High

User has E3 plus one or more E5-included add-ons that together cost MORE than full E5. This is a mathematical certainty — upgrading saves money AND unlocks remaining E5 capabilities.

**Scenario:** A company buys M365 E3 (€36.20/mo) and later the security team mandates E5 Security add-on (€12/mo). Total: €48.20/mo. Full M365 E5 costs €53.70/mo — wait, that's not inverted yet. But add Teams Phone Standard (€8/mo) and you're at €56.20/mo. Now E5 is €2.50/mo cheaper AND gives you Power BI Pro, Defender for Cloud Apps, and risk-based CA for free.

---

### The Abandoned Vault
**Tag:** `E5 DATA HOARDER` | **Confidence:** High

A disabled account on Litigation Hold still has an expensive suite (E3/E5) attached because IT feared removing the license would break the hold. This is a common misconception — Microsoft automatically converts it to a free Inactive Mailbox that retains ALL content and holds indefinitely.

**Scenario:** An executive left 6 months ago. Legal placed their mailbox on Litigation Hold. IT disabled the account but left the €53.70/mo E5 license "just in case." That's €322/yr burned on a misconception. Remove the license — Microsoft will safely create a free Inactive Mailbox.

---

### The Quiet Archive
**Tag:** `INACTIVE HOLD` | **Confidence:** High

Same as E5 Data Hoarder but for cheaper licenses. A disabled account on Litigation Hold with a less expensive license still doesn't need it. Safe to remove — Microsoft will convert to a free Inactive Mailbox.

**Scenario:** A contractor left and their Exchange Online Plan 2 (€8.80/mo) is kept because of a legal hold. Remove the license — the hold persists on the free Inactive Mailbox. Saves €105.60/yr.

---

### The AI Double-Dip
**Tag:** `AI ADD-ON OVERLAP` | **Confidence:** High

User has both Copilot for M365 and Teams Premium, and organizes zero meetings. Copilot natively includes Teams Intelligent Recap. Since the user doesn't organize meetings, they don't use Premium's advanced webinar branding or custom meeting template features either. Teams Premium is 100% redundant.

**Scenario:** An analyst has Copilot (€30/mo) and Teams Premium (€10/mo). They attend meetings but never organize any — they just need the AI recap, which Copilot already provides. Remove Teams Premium. Savings: €120/yr.

---

### The Hidden Freebie
**Tag:** `SEEDED VISIO OVERLAP` | **Confidence:** Medium

Microsoft now includes a lightweight "Visio in Microsoft 365" web app natively in E1/E3/E5 suites. Users with standalone Visio Plan 1 (€4.70/mo) alongside a qualifying suite likely don't need Plan 1 for viewing or light editing.

**Scenario:** 200 network engineers have E3 + Visio Plan 1 just to view network diagrams. Since E3 now includes the native Visio web app, those 200 × €4.70/mo = €940/mo (€11,280/yr) in Visio Plan 1 licenses can likely be removed.

---

*Last updated: 2026-02-23 | Detections: 50+ unique scenarios*
