# QA Report — LFU Command Center V1

**Status: PASS** | All critical acceptance tests passed.

---

## Test Environment

| Item | Value |
|------|-------|
| Version | V1 (Code.gs v8.1.0) |
| Runtime | Google Apps Script V8 |
| Tested browsers | Chrome, Safari, Firefox |
| Sheet type | Google Sheets (container-bound) |

---

## Critical Path Tests

### P0 — Must Pass Before Release

| Test | Expected | Status |
|------|----------|--------|
| Setup Wizard runs and creates all sheets | All 13 required sheets created, Settings populated | ✅ PASS |
| Leads sheet headers created correctly | All 33 required columns present | ✅ PASS |
| `uiGetStateReadOnly()` returns leads | `{ ok:true, data:{ kpis, leads, pipeline, features } }` | ✅ PASS |
| Today tab shows due leads | Leads with Follow Up Due ≤ today + phone or email + non-closed status appear | ✅ PASS |
| Today tab empty state (has leads, none due) | Shows "All caught up!" message + "View All Leads →" button | ✅ PASS |
| Today tab empty state (no leads at all) | Shows "No leads yet" + "Quick Add Lead" button | ✅ PASS |
| All Leads tab shows all leads regardless of due status | All qualified rows appear | ✅ PASS |
| Tab switching blue line moves correctly | Active tab border-bottom updates on click | ✅ PASS |
| Quick Add Lead — email only | Lead added, appears in Today (if due date is today) | ✅ PASS |
| Quick Add Lead — phone only | Lead added, appears in Today (V1 fix: phone now valid contact) | ✅ PASS |
| Quick Add Lead — no email or phone | Error message shown inline, lead not added | ✅ PASS |
| `uiApproveToggle()` approves a lead | Approved to Send = TRUE written to sheet | ✅ PASS |
| `uiMarkCalledSafe()` marks called | Last Contacted = now, Follow Up Due advances, Due Flag = FALSE | ✅ PASS |
| `uiSnoozeSafe()` snoozes lead | Follow Up Due advances by N days, Due Flag = FALSE | ✅ PASS |
| `uiMoveLeadSafe()` changes status | Status updated, Won/Lost date set where applicable | ✅ PASS |
| `uiRunReviewSafe()` generates drafts | Draft column populated, Draft Status = DRAFTED | ✅ PASS |
| `uiSendApprovedSafe()` sends emails | Emails sent for rows with Approved to Send = TRUE | ✅ PASS |
| `doGet()` serves CommandCenter.html | Web app URL returns correct HTML | ✅ PASS |
| `doGet(?action=health)` returns JSON | Health check returns `{ ok:true, data:{ version, sheets, triggers } }` | ✅ PASS |
| Deployment Helper opens as sidebar | Sidebar appears via Autopilot → Deployment Helper | ✅ PASS |
| Repair function non-destructive | Creates missing sheets only, never deletes or overwrites data | ✅ PASS |
| Error banner shows on load failure | Error message surfaced without stack trace | ✅ PASS |

---

## P1 — Should Pass

| Test | Expected | Status |
|------|----------|--------|
| Auto-refresh every 25s | `load(true)` fires on schedule, refreshes lead list | ✅ PASS |
| Window focus triggers refresh | Navigating back to tab refreshes silently | ✅ PASS |
| Activity tab loads last 100 actions | Acts list rendered with timestamps | ✅ PASS |
| Lead detail drawer opens on "Details" | Drawer slides in with all fields populated | ✅ PASS |
| Lead drawer — activity history loads | Per-lead activity shown async | ✅ PASS |
| Priority Score displayed on lead cards | P0–P10 badge shown, fire/hot class on high scores | ✅ PASS |
| Stage dropdown moves lead | `uiMoveLeadSafe` fires, lead moves, list refreshes | ✅ PASS |
| Schedule modal disabled when calendar off | Button hidden when `enable_calendar = FALSE` | ✅ PASS |
| Dopamine counter increments on actions | Count increases on Called, Approve, Snooze, Move, Quick Add | ✅ PASS |
| Toast system shows correct type | success/error/info/warn types display correct color + icon | ✅ PASS |
| Keyboard shortcut 1/2/3 switches tabs | Tab switches correctly from keyboard | ✅ PASS |
| Keyboard shortcut A opens Quick Add | Modal opens | ✅ PASS |
| Keyboard shortcut R refreshes | `load(true)` fires | ✅ PASS |
| Esc closes all modals + drawer | All overlay elements close | ✅ PASS |
| Search filters leads client-side | Results filter in real-time with no server call | ✅ PASS |
| Stage filter works | Only leads in selected stage shown | ✅ PASS |
| Owner filter works | Only leads with selected owner shown | ✅ PASS |
| Due filter (due only / not due) | Correct subset shown | ✅ PASS |
| Filter count shows X / Y | Correct filtered count vs. total | ✅ PASS |

---

## P2 — Nice to Have

| Test | Expected | Status |
|------|----------|--------|
| Style Spreadsheet applies formatting | CF rules, validation, header style applied to Leads | ✅ PASS |
| Rules Engine fires on stage change | RULE-001 through RULE-008 trigger on matching stages | ✅ PASS |
| Daily Briefing sends email | Email with pipeline snapshot sent to authorized user | ✅ PASS |
| Import Wizard processes CSV | Leads created/updated per mapping in row 9 | ✅ PASS |
| Dual token `{token}` and `{{token}}` in templates | Both formats replaced correctly | ✅ PASS |
| `getFirstName_()` handles "Last, First" format | First name extracted correctly | ✅ PASS |

---

## Known Limitations

| Item | Notes |
|------|-------|
| Scheduling requires `enable_calendar = TRUE` | Off by default. Buyer must enable in Settings. |
| Daily trigger must be installed by buyer | Setup Wizard installs it. Does not auto-install. |
| Web app URL is deployment-specific | Re-deploying creates a new URL. Old bookmarks break. |
| Max 500 emails per run | `max_emails_per_run` setting. Google Apps Script MailApp limits apply. |
| Sheet writes not concurrent-safe | Single-user system by design. Race conditions possible with multiple simultaneous users. |
| Intake/SOP/Import features not in Autopilot menu | Advanced features are in the codebase but not exposed in Track A menu. |

---

*LFU Command Center V1*
