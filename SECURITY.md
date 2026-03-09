# SECURITY.md — LFU Command Center V1

Plain-English security guide. Read this before sharing your setup with others.

---

## Who can access the Command Center?

**Sidebar (default):** Only you. The sidebar runs inside your Google Sheet and requires Sheet access.

**Web App:** Controlled by your deployment settings:
- **Only myself** — only you can open the URL. Recommended for solo use.
- **Anyone with Google account** — anyone with the URL who has a Google account can open it. Use for teams where everyone has the Sheet shared with them.
- **Anyone** — public. Not recommended. Do not use unless you understand the risk.

**Recommendation:** Start with "Only myself". Expand to "Anyone with Google account" when onboarding team members, and ensure they have at least Viewer access to the Sheet.

---

## What data does the web app expose?

The web app reads from and writes to your Google Sheet. It does not:
- Store data outside your Google account
- Send data to any third party
- Use any external database or API

All processing happens server-side in your Google Apps Script. The client (browser) only sees what the server returns.

---

## Email sending

Emails are sent from **your** Gmail account via `MailApp.sendEmail()`. Recipients see your email address as the sender. Monitor your Gmail Sent folder to verify.

Daily send quota: Google enforces ~100 emails/day for personal accounts, ~1,500/day for Workspace. The `max_emails_per_run` setting prevents you from accidentally burning your quota in one run.

---

## Error handling

- Stack traces never reach the browser. Errors are logged server-side only (Apps Script Executions log).
- Each error gets a `debugId` (UUID) that appears in the UI. You can search the Apps Script execution log by this ID for technical details.
- Activity log records only action codes and debugIds — not raw technical error text.

---

## OAuth scopes

The script requests these Google OAuth scopes (defined in appsscript.json if explicitly set, otherwise determined by usage):
- `spreadsheets` — read/write your Google Sheet
- `gmail` — send emails on your behalf
- `script.external_request` — UrlFetchApp for web app health probe
- `userinfo.email` — Session.getActiveUser().getEmail()
- `calendar` — only if enable_calendar = TRUE in Settings

Calendar scope is disabled by default. It only activates when you set `enable_calendar = TRUE` in your Settings sheet and re-authorize.

---

## Template Sheet security

If you distribute a template Sheet to buyers:
1. Publish it as "view only" to prevent editing the template.
2. Direct buyers to **File → Make a copy** — they own their copy.
3. Never put live lead data in a template Sheet.
4. Never set web app access to "Anyone" in a template.

---

## Responsible disclosure

If you discover a security issue, do not post it publicly. Contact the seller directly.
