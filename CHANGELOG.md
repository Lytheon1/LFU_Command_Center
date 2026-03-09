# CHANGELOG — LFU Command Center

---

## V1.0 (current release)

### Bug Fixes

**BUG FIX: Leads Not Appearing in Today Tab**
- `addComputedDueFlag_()` previously required an email address to mark a lead as due. Leads with only a phone number never appeared in the Today list.
- Fixed: leads are now flagged as due if they have **either a phone number OR an email address** (plus a valid Follow Up Due date ≤ today in a non-closed status).
- Impact: Phone-only leads now correctly appear in Today.

**BUG FIX: Tab Blue Line Not Moving on Click**
- `switchTab()` now uses an explicit ID-based approach (`$('tab-today')`, `$('tab-all')`, `$('tab-activity')`) with `classList.add('active')` / `classList.remove('active')` instead of the previous `data-tab` selector + `classList.toggle(force)` pattern.
- This resolves a browser compatibility edge case where the CSS active indicator did not reliably update in Google Apps Script's HtmlService rendering environment.

**BUG FIX: Empty Today Tab Not Differentiating States**
- When Today had no due leads but leads existed in the system, the empty state showed a generic message. Users interpreted this as "leads not loaded."
- Fixed: Today now shows three distinct empty states:
  1. **Has leads, none due today** → "All caught up! 🎉 Your N leads are visible in All Leads." + "View All Leads →" button
  2. **No leads at all** → "No leads yet 🚀" + "Quick Add Lead" button  
  3. **Has leads, filtered out** → "No leads match your filters"

**BUG FIX: Incorrect File Names in Docs**
- `SELLER_SETUP.md` referenced `WebApp.gs` (Step 2) and `WebApp.html` (Step 3) — both of which were renamed in v2.3 to resolve a naming collision. These references now correctly say `WebAppCode.gs` and `CommandCenter.html`.
- `HANDOFF_CHECKLIST.md` referenced old file names in the ZIP contents list — updated.

**BUG FIX: Deployment Instructions Inconsistency**
- SELLER_SETUP.md and QUICKSTART.md had different "Execute as" recommendations.
- Both now consistently recommend **Execute as: Me**, with an explanation that "Me" refers to the buyer's own account (not the seller's), since each buyer deploys their own copy.

### Improvements

**IMPROVED: CommandCenter.html Rewritten**
- Tab switching uses explicit ID-based selectors — more robust and browser-compatible.
- Tab IDs: `tab-today`, `tab-all`, `tab-activity` (plain IDs, no `data-tab` dependency for JS logic).
- Better empty states with context-aware messaging and action buttons.
- Improved lead cards: phone pill always shown when present; due date shown in pill format.
- Improved drawer: cleaner section layout, appointment section hidden when empty.
- Keyboard shortcuts added: `1/2/3` switch tabs, `A` opens Quick Add, `R` refreshes, `Esc` closes modals.
- Auto-refresh every 25 seconds + on window focus (unchanged behavior, more reliable).
- Toast system and error banner unchanged — kept for consistency.

**IMPROVED: Quick Add Validation**
- Quick Add now accepts a lead with Name OR Company (not requiring both).
- Requires at least an Email OR Phone (not email-only).
- Aligns with the fixed due-flag logic: a phone-only lead will now appear in Today.

**IMPROVED: Version Bump**
- `Code.gs` internal version: `8.1.0`
- Distribution version: `V1`

### Architecture (unchanged from v2.3)

File naming (established in v2.3, confirmed in V1):
- `Code.gs` — canonical backend
- `WebAppCode.gs` — `doGet()` + safe wrappers + repair + dopamine
- `CommandCenter.html` — full-screen web app UI (served by `doGet()`)
- `DeploymentHelper.html` — deployment helper sidebar

---

## v2.3 (previous release — reference only)

- Critical: Fixed naming collision between WebApp.gs / WebApp.html (renamed to WebAppCode.gs / CommandCenter.html)
- Fixed: Lead names rendering correctly; `isQualifiedLeadRow_()` accepts any non-empty identity field
- Fixed: Quick Add button opens a proper modal with inline validation
- New: Style Spreadsheet function
- New: `safe_()` envelope on all WebApp endpoints — no stack traces to browser
- New: `uiGetStateReadOnly()` — read-only state getter (no sheet writes on refresh)
- New: `safeLogActivity_()` — non-throwing activity log wrapper
