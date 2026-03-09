# Seller Setup Guide — Creating the Template Sheet

**This document is for you (the seller), not for the buyer.**
You do this once. After this, buyers only need to "Make a copy" and run Setup Wizard.

---

## What You Are Creating

A "master template" Google Sheet with the script pre-installed, pre-configured with sensible defaults, and ready for buyers to copy. No buyer should ever need to paste code.

---

## One-Time Setup Steps

### Step 1 — Create a clean Google Sheet

1. Go to [sheets.new](https://sheets.new) to create a new Sheet.
2. Rename it: `LFU Command Center Template` (or your brand name).

---

### Step 2 — Open Apps Script

1. Click **Extensions → Apps Script**.
2. You will see a default `Code.gs` file with an empty function. You will replace this.

---

### Step 3 — Install the script files

You need to add **4 files** in the exact order below.

> ⚠️ **File naming matters.** Apps Script identifies files by name. The names below must match exactly — including capitalization. The old `WebApp.gs` / `WebApp.html` names caused a naming collision and have been permanently replaced in V1.

---

**File 1: Replace Code.gs**
1. Click the existing `Code.gs` in the left sidebar
2. Select all content → delete it
3. Paste the entire contents of `Code.gs` from this package
4. Save (Ctrl+S or Cmd+S)

---

**File 2: Add WebAppCode.gs**
1. Click **+** next to "Files" in the left sidebar → select **Script**
2. Name it exactly: `WebAppCode` (the `.gs` extension is added automatically)
3. Delete any placeholder content → paste the entire contents of `WebAppCode.gs`
4. Save

> This file contains the `doGet()` web app entry point, safe wrappers, and the Deployment Helper. It must be named `WebAppCode`.

---

**File 3: Add CommandCenter.html**
1. Click **+** → select **HTML**
2. Name it exactly: `CommandCenter` (the `.html` extension is added automatically)
3. Delete any placeholder content → paste the entire contents of `CommandCenter.html`
4. Save

> This is the full-screen Command Center web app UI. The `doGet()` function in WebAppCode.gs references this file by name (`"CommandCenter"`). The name must match exactly.

---

**File 4: Add DeploymentHelper.html**
1. Click **+** → select **HTML**
2. Name it exactly: `DeploymentHelper`
3. Delete any placeholder content → paste the entire contents of `DeploymentHelper.html`
4. Save

---

**Update appsscript.json:**
1. In the left sidebar, click the **⚙ Project Settings** gear icon
2. Scroll down to **Show "appsscript.json" manifest file in editor** → enable it
3. Click `appsscript.json` in the file list
4. Replace all contents with the contents of `appsscript.json` from this package
5. Save

Your file list in Apps Script should now show:
- `Code.gs`
- `WebAppCode.gs`
- `CommandCenter.html`
- `DeploymentHelper.html`
- `appsscript.json` (manifest)

---

### Step 4 — Run Setup Wizard to initialize sheets

1. Go back to your Google Sheet (the spreadsheet tab)
2. Refresh the page
3. Click **Autopilot → Setup Wizard** in the menu bar
4. Authorize when prompted (click through Google's permission screens — this is expected)
5. Fill in your defaults:
   - **Sender name** — your name or brand name
   - **Send mode** — recommend: REVIEW
   - **Follow-up days** — recommend: 3
   - **Stage Pack** — choose the pack that fits your niche
6. Click **Save and Install**

The wizard will create all required sheets: Leads, Settings, Templates, Activities, System Ops, and more.

---

### Step 5 — Deploy the Web App (recommended)

Deploying the web app gives buyers a permanent URL for the full-screen Command Center.

1. In the Apps Script editor: **Deploy → New deployment**
2. Click the gear icon next to "Select type" → choose **Web app**
3. Configure:
   - **Description:** `LFU Command Center V1`
   - **Execute as:** `Me` (your Google account)
   - **Who has access:** `Anyone with Google account` (for team use) or `Only myself` (for solo buyers)
4. Click **Deploy** → copy the `/exec` URL
5. Back in the Google Sheet: **Autopilot → Deployment Helper**
6. Paste the URL → click **Save URL**

Now when a buyer copies the Sheet, the URL is pre-saved and they can use it immediately.

> **Note on "Execute as: Me"** — Because each buyer copies the Sheet and deploys their own web app from their own Google account, "Me" refers to the buyer's account, not yours. Each copy is fully independent.

---

### Step 6 — (Optional) Add sample leads

Add 2–3 clearly fake sample leads to the Leads sheet so buyers can see how the system looks:
- Name: `Jane Example`
- Email: `example@example.com`
- Company: `Sample Corp`
- Status: `NEW`
- Follow Up Due: today's date
- Deal Value: `1000`

---

### Step 7 — Share the template

1. In the Google Sheet: **File → Share → Share with others**
2. Change the access to: **Anyone with the link → Viewer**
3. Copy the share link

Give buyers this share link. When they open it and choose **File → Make a copy**, they get their own independent copy with all scripts installed.

---

## Important Notes

- **Each buyer's copy is fully independent.** Their data, triggers, and Script Properties are completely separate.
- **Triggers do not copy.** Buyers must run Setup Wizard to install their own daily trigger.
- **Authorization does not copy.** Each buyer authorizes the script on their first use.
- **Script is version-locked at copy time.** To deliver updates, notify buyers to re-copy or provide patched files.

---

## Pre-Handoff Checklist

- [ ] Apps Script has exactly these 4 files: `Code.gs`, `WebAppCode.gs`, `CommandCenter.html`, `DeploymentHelper.html`
- [ ] `appsscript.json` has been replaced with the package version
- [ ] Setup Wizard has been run at least once (all required sheets exist in the spreadsheet)
- [ ] Settings sheet has sensible defaults for your niche
- [ ] At least 1–2 sample leads exist (clearly fake data)
- [ ] Web app has been deployed (optional but recommended) and URL saved
- [ ] Sheet is shared as "Anyone with the link → Viewer"
- [ ] You have tested "File → Make a copy" yourself and verified the buyer experience

---

*LFU Command Center V1*
