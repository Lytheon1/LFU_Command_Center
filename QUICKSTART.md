# LFU Command Center — Quick Start Guide

**V1 · Google Apps Script · No Node · No External Hosting**

Install time: ~5 minutes.

---

## Step 1 — Make a copy of the template Google Sheet

1. Open the template link your seller provided.
2. Click **File → Make a copy**.
3. Name it whatever you like and save it to your Google Drive.
4. Open your copy. You are now the owner.

> **Why a copy?** The Apps Script project is "container-bound" — it lives inside the Sheet. You must own the copy to authorize and run it.

---

## Step 2 — Authorize and run Setup Wizard

1. In your Sheet, click the **Autopilot** menu. If you don't see it, wait a moment and refresh the page once.
2. Click **Autopilot → Setup Wizard**.
3. Complete the Google authorization prompts (click through — this is expected and safe).
4. Fill in your settings:
   - **Stage Pack** — choose the pipeline that fits your business
   - **Sender name** — your name (used in email signatures)
   - **Send mode** — recommend: REVIEW (drafts first, you approve before sending)
5. Click **Save and Install**.

The wizard creates all required sheets (Leads, Settings, Templates, Activities, etc.) and sets up your pipeline.

---

## Step 3 — Open the Command Center

**Option A — Sidebar (no setup required, works immediately):**
Click **Autopilot → Open CRM (Sidebar)**.

**Option B — Full-screen Web App (recommended for daily use):**

1. In your Sheet, click **Extensions → Apps Script**.
2. Click **Deploy → New deployment**.
3. Click the gear ⚙ next to "Select type" → choose **Web app**.
4. Set:
   - **Execute as:** `Me`
   - **Who has access:** `Only myself` (solo) or `Anyone with Google account` (team)
5. Click **Deploy** → copy the `/exec` URL (it ends in `/exec`).
6. Go back to your Sheet → click **Autopilot → Deployment Helper**.
7. Paste the URL into the URL field → click **Save URL**.
8. Click **Autopilot → Open Command Center (Web App)** to open it.

> You only need to deploy once. Bookmark the URL for daily use.

---

## Step 4 — Add your leads

There are three ways to add leads:

- **Quick Add:** Click **+ Add Lead** in the Command Center (fastest for 1–5 leads).
- **Direct entry:** Type directly into rows in the Leads sheet.
- **Import:** Paste a CSV into the Import sheet → run **Autopilot → Run Import**.

> **Important:** Leads need at least a Name (or Company) and a phone number or email address, plus a **Follow Up Due** date, to appear in the Today list.

---

## Step 5 — Run your first autopilot cycle

1. Click **Autopilot → Run Autopilot (REVIEW)** — generates email drafts for all due leads.
2. Review the drafts in the Leads sheet (Draft column).
3. Set **Approved to Send = TRUE** for drafts you want to send.
4. Click **Autopilot → Send Approved Drafts** to send.

Or use the **▷ Run Review** and **✓ Send Approved** buttons directly inside the Command Center.

---

## Daily Workflow

1. Open Command Center → **⚡ Today** tab.
2. Work due leads sorted by Priority Score (highest urgency first).
3. For each lead: **Call**, **Text**, **Email**, click **✅ Called**, or **Snooze**.
4. Approve good email drafts with **✓ Approve**.
5. Click **✓ Send Approved** when done.

**Keyboard shortcuts (web app):**
- `1` / `2` / `3` → switch tabs (Today / All Leads / Activity)
- `A` → Quick Add Lead
- `R` → Refresh
- `Esc` → close modals

---

## Style Your Spreadsheet (Optional)

Click **Autopilot → Style Spreadsheet** to apply:
- Frozen header row and filter views
- Data validation dropdowns (Status, Approved to Send)
- Conditional formatting (Due, Errors, Approved, Sent, Closed)
- System column protection (warn-only — no data is locked)
- Hidden auxiliary sheets (keeps it clean)

Safe to re-run anytime. **Never deletes data.**

---

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| Autopilot menu missing | Wait a moment and refresh the page |
| "Leads sheet not found" error | Run **Autopilot → Setup Wizard** |
| Today tab is empty | Switch to **All Leads** to confirm leads exist. Today only shows leads with a due date ≤ today AND a phone or email |
| Leads not showing as due | Check: does the lead have a Phone or Email set? Is Follow Up Due ≤ today? Is Status not CLOSED/LOST/BOOKED? |
| Authorization error on web app | Share the Sheet with the user (Viewer or higher) before they open the URL |
| Drafts not generating | Check: lead has Email + Follow Up Due ≤ today + non-closed Status |
| Web app probe blocked / 401 | Expected if login required — open the URL in your browser to confirm it loads |
| Names not showing in Command Center | Run Setup Wizard once — it ensures all required header columns are present |
| "Unknown server function" error | Confirm `WebAppCode.gs` is present in Apps Script (not the old `WebApp.gs`) |

---

## Security Defaults

- Web app defaults to **Only myself**. Change to "Anyone with Google account" for team use.
- All email sending uses the authorized Google account's Gmail.
- See `SECURITY.md` for a full security overview.

---

## File Reference

| File | Purpose |
|------|---------|
| `Code.gs` | Core backend — autopilot engine, rules, email, sheet operations |
| `WebAppCode.gs` | Web app layer — `doGet()`, safe wrappers, repair, dopamine system |
| `CommandCenter.html` | Full-screen Command Center web app (served by `doGet()`) |
| `DeploymentHelper.html` | Deployment Helper sidebar (install URL, probe, repair) |
| `appsscript.json` | Apps Script manifest (V8 runtime, timezone) |
| `QUICKSTART.md` | This file |
| `SECURITY.md` | Security and privacy guide |
| `QA_REPORT.md` | Acceptance test checklist |
| `SELLER_SETUP.md` | Seller/reseller one-time setup guide |
| `HANDOFF_CHECKLIST.md` | Pre-delivery checklist |

---

*LFU Command Center V1*
