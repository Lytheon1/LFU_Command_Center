# Handoff Checklist — One-Time Sale

**For the seller.** Complete every item before handing over to the buyer.

---

## What to Hand Over

| Item | How to hand it over | Notes |
|------|---------------------|-------|
| Template Google Sheet URL | Share as "Make a copy" link | Do NOT share your master copy directly |
| This ZIP package | Attach to email or shared Drive folder | Buyer receives all files listed below |
| Buyer's email address | For sharing the Sheet if needed | Buyer must have at least Viewer access to open the web app |

---

## Pre-Handoff Checklist (Seller Does This)

- [ ] Apps Script project contains exactly **4 files** with these exact names:
  - `Code.gs`
  - `WebAppCode.gs`
  - `CommandCenter.html`
  - `DeploymentHelper.html`
- [ ] `appsscript.json` manifest has been replaced with the V1 package version
- [ ] Setup Wizard has been run at least once (all required sheets exist: Leads, Settings, Templates, Activities, System Ops, Rules, etc.)
- [ ] Settings sheet has sensible defaults for your niche (Sender name, Stage Pack, Send Mode)
- [ ] At least 1–2 rows of clearly fake sample lead data exist in the Leads sheet
- [ ] Web app deployed (optional but recommended) — URL saved via Deployment Helper
- [ ] Sheet is set to: **Anyone with the link → Viewer**
- [ ] You have done a "File → Make a copy" test yourself and verified the buyer experience end-to-end

---

## ZIP Package Contents

The ZIP file delivered to buyers should contain:

- `Code.gs`
- `WebAppCode.gs`
- `CommandCenter.html`
- `DeploymentHelper.html`
- `appsscript.json`
- `QUICKSTART.md`
- `SELLER_SETUP.md`
- `HANDOFF_CHECKLIST.md`
- `SECURITY.md`
- `QA_REPORT.md`
- `CHANGELOG.md`

---

## What the Buyer Owns (After Install)

Everything below belongs to the buyer forever. No recurring cost, no dependency on you after handoff.

| Asset | Location | Owner after handoff |
|-------|----------|---------------------|
| Google Sheet | Buyer's Google Drive | Buyer |
| Apps Script project (bound to Sheet) | Embedded in Sheet | Buyer |
| All lead data | Leads sheet | Buyer |
| Settings | Settings sheet | Buyer |
| Activity log | Activities sheet | Buyer |
| Templates | Templates sheet | Buyer |
| Web app URL (if deployed) | Script Properties | Buyer (can redeploy for new URL) |
| Daily trigger | Apps Script Triggers | Buyer (can reinstall via Setup Wizard) |

---

## What the Buyer Can Change Safely

- All data in the Leads sheet: add, edit, delete rows
- Settings sheet: change any value in column B
- Templates sheet: add or edit templates
- Stage names: change `pipeline_stages` value in Settings
- Web app access: re-deploy with different "Who has access" setting
- Time zone: change `timeZone` in `appsscript.json`, then re-save

---

## What the Buyer Should NOT Do Without Care

- **Do not rename sheets.** The script finds sheets by exact name ("Leads", "Settings", etc.). Renaming will break things.
- **Do not delete the header row** in any sheet. Row 1 is always the header.
- **Do not add a second `doGet` function** if they add other scripts — merge the logic into the existing one in WebAppCode.gs.
- **Do not change the web app URL** without updating it in the Deployment Helper — existing bookmarks will break.

---

## How to Recover If Something Breaks

| Problem | Fix |
|---------|-----|
| "Leads sheet not found" in Command Center | Autopilot → Setup Wizard (non-destructive — safe to re-run) |
| Web app blank page or error after redeploy | Deployment Helper → Probe URL. If 401, open URL in browser to confirm login |
| Daily emails stopped | Apps Script → Triggers — if trigger is missing, re-run Setup Wizard |
| Missing columns in Leads sheet | Deployment Helper → Repair (adds missing columns without touching data) |
| Buyer lost web app URL | Apps Script → Deploy → Manage deployments → copy URL again |
| Script authorization expired | Extensions → Apps Script → Run any function → approve permissions again |
| Leads not appearing in Today tab | Check: lead needs Phone or Email + Follow Up Due ≤ today + non-closed Status |
| Tab blue line not switching | Clear browser cache or try a hard reload (Ctrl+Shift+R) |

---

## Support Included vs. Not Included

**Typically included (one-time sale):**
- Initial setup walkthrough (one session, up to 30 minutes)
- Fix for bugs in the delivered code
- Clarification on QUICKSTART steps

**Typically not included:**
- Ongoing feature development
- Data migration from other CRMs
- Custom integrations (Slack, Notion, HubSpot, etc.)
- Google Workspace admin issues (buyer resolves with their IT)
- Issues caused by buyer modifying the script after handoff

---

*LFU Command Center V1*
