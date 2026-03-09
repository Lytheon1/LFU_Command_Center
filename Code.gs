/**
 * Lead Follow-Up Autopilot Command Center (Google)
 * v8.0.0
 *
 * Single-file install: paste entire file into Apps Script → Code.gs
 * No Deploy required. Run "Setup Wizard" from the Autopilot menu.
 *
 * NEW in v8.0.0 vs v7.3.3:
 *   ✅ Rules Engine Lite   — table-driven automations (stage changes trigger actions)
 *   ✅ Activity Timeline   — per-lead history in Command Center
 *   ✅ Quick Add Lead      — add leads without leaving Command Center
 *   ✅ Schedule Modal      — proper modal (no more browser prompts)
 *   ✅ Auto-refresh UI     — live updates, configurable interval
 *   ✅ Lead Priority Score — 0-10 score badge on every card
 *   ✅ Import Wizard       — CSV paste → column map → dedupe import
 *   ✅ Intake Processor    — form submissions → leads (Typeform/Jotform/Google Forms)
 *   ✅ Scheduler Lite      — appointments + Google Calendar events
 *   ✅ Appointment Reminders — auto email reminders before appointments
 *   ✅ Daily Briefing      — morning email summary of pipeline state
 *   ✅ SOP Pack Impact 15  — 60 pre-built SOPs across 15 business categories
 *   ✅ Dual token support  — {token} AND {{token}} both work in templates
 *   ✅ appendLeadRow_      — missing helper now defined
 *   ✅ Constant unification — SHEET_SOP_TMPL consistent throughout
 *   ✅ Kanban XSS guard    — name/company escaped before injecting into HTML
 *   ✅ Error resilience    — all sheet ops wrapped with graceful fallbacks
 *
 * Sheets auto-created on first run:
 *   Leads | Settings | Templates | Activities | System Ops
 *   SOP Library | Process Improvements | SOP Builder | SOP Templates
 *   Rules | Import | Intake | Appointments
 */

// ═══════════════════════════════════════════════════════════════
//  CONSTANTS
// ═══════════════════════════════════════════════════════════════

const LFU = {
  VERSION: "8.1.0",
  SHEET_LEADS:        "Leads",
  SHEET_SETTINGS:     "Settings",
  SHEET_TEMPLATES:    "Templates",
  SHEET_ACTIVITIES:   "Activities",
  SHEET_SYSOPS:       "System Ops",
  SHEET_SOPS:         "SOP Library",
  SHEET_IMPROVE:      "Process Improvements",
  SHEET_SOP_BUILDER:  "SOP Builder",
  SHEET_SOP_TMPL:     "SOP Templates",
  SHEET_RULES:        "Rules",           // v8: Rules Engine
  SHEET_IMPORT:       "Import",          // v8: Import Wizard
  SHEET_INTAKE:       "Intake",          // v8: Intake Processor
  SHEET_APPOINTMENTS: "Appointments",    // v8: Scheduler
  CLOSED_STATUSES:  ["BOOKED","CLOSED"],
  WON_STATUSES:     ["BOOKED","CLOSED","SCHEDULED"],
  LOST_STATUSES:    ["LOST","NO_SHOW"],
  HEADERS_REQUIRED: [
    "Lead ID","Name","Preferred Name","Email","Phone","Company",
    "Template Pack","Status","Owner","Deal Value","Probability %",
    "Expected Value","Close Date","Won Date","Lost Date",
    "Next Action","Last Contacted","Follow Up Due",
    "Cadence Override Days","Notes","Draft","Draft Status",
    "Approved to Send","Last Error","Last Drafted","Last Sent",
    "Due Flag","Priority Score",         // v8: Priority Score column
    "Appointment Start","Appointment End",
    "Calendar Event ID","Calendar Event Link",
    "Intake Source","Intake Raw"
  ],
  SETTINGS_DEFAULTS: {
    send_mode:                      "REVIEW",
    max_emails_per_run:             "10",
    business_days_only:             "TRUE",
    follow_up_days:                 "3",
    autopilot_hour_local:           "9",
    default_template_pack:          "general",
    stage_pack:                     "local_service",
    pipeline_stages:                "NEW,CONTACTED,NURTURE,BOOKED,CLOSED",
    stage_colors:                   "NEW=#7aa2ff;CONTACTED=#33d17a;NURTURE=#ffcc00;BOOKED=#b197fc;CLOSED=#8b949e",
    review_send_source:             "DRAFT",
    text_body_template:             "Hi {first_name}, quick follow up on {company}. Text me when you can. {sender_name}",
    email_subject_template:         "Quick follow up, {first_name}",
    email_body_template:            "Hi {first_name},\n\nJust following up on {company}. If it makes sense, I can share options and next steps.\n\nBest,\n{sender_name}",
    // Appointments
    appointment_default_duration_min:   "30",
    appointment_followup_days_after:    "1",
    appointment_reminder_enabled:       "TRUE",
    appointment_reminder_hours_before:  "24",
    appointment_reminder_subject:       "Reminder: your appointment on {date}",
    appointment_reminder_body:          "Hi {first_name},\n\nQuick reminder of your appointment on {date} at {time}.\n\nIf you need to reschedule, reply to this email.\n\nThanks,\n{sender_name}",
    // Rules Engine
    rules_enabled:                  "TRUE",
    // Calendar / Scheduling (disabled by default — enable to unlock schedule features)
    enable_calendar:                "FALSE"
  }
};

// ═══════════════════════════════════════════════════════════════
//  MENU
// ═══════════════════════════════════════════════════════════════

function onOpen() {
  // DISTRIBUTION MENU — Track A scope only.
  // Advanced features (SOP, Rules, Import, Intake, Scheduler, Briefings)
  // remain in the codebase but are not exposed to buyers in this release.
  SpreadsheetApp.getUi()
    .createMenu("Autopilot")
    .addItem("Setup Wizard",                         "showSetupWizard")
    .addItem("Open CRM (Sidebar)",                   "showCrmSidebar")
    .addSeparator()
    .addItem("Run Autopilot (REVIEW)",               "autopilotRunReviewNow")
    .addItem("Send Approved Drafts",                 "sendApprovedDrafts")
    .addSeparator()
    .addItem("Open Command Center (Web App)",        "openCommandCenterWeb_")
    .addItem("Deployment Helper",                    "showDeploymentHelper")
    .addSeparator()
    .addItem("Style Spreadsheet",                    "styleSpreadsheet")
    .addSeparator()
    .addItem("Help",                                 "openHelp_")
    .addToUi();
}
function showSetupWizard() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutput(getSetupWizardHtml_()).setTitle("Autopilot Setup Wizard")
  );
}

function showCrmSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutput(getCrmSidebarHtml_()).setTitle("Command Center")
  );
}

function openCommandCenterFull_() {
  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutput(getCommandCenterAppHtml_())
      .setTitle("Command Center").setWidth(1260).setHeight(840),
    "Command Center"
  );
}

function openHelp_() {
  SpreadsheetApp.getUi().alert("Autopilot Help",
    "v" + LFU.VERSION + " — Lead Follow-Up Autopilot Command Center\n\n" +
    "CORE ACTIONS:\n" +
    "• Run Autopilot (REVIEW) → creates drafts only\n" +
    "• Send Approved Drafts → sends rows with Approved to Send = TRUE\n" +
    "• Open Command Center (Full Screen) → full pipeline + actions\n\n" +
    "RULES ENGINE:\n" +
    "• Open Rules Engine → view/edit automation rules\n" +
    "• Run Rules Now → manually trigger rules against all leads\n" +
    "• Rules auto-run on stage changes and autopilot runs\n\n" +
    "INTAKE + SCHEDULING:\n" +
    "• Connect Typeform/Jotform/Google Forms → write to Intake tab\n" +
    "• Install Intake Trigger → processes new form submissions every 5 min\n" +
    "• Schedule Selected Lead → create Google Calendar event from a lead row\n\n" +
    "DAILY BRIEFING:\n" +
    "• Install Daily Briefing → morning email at 8am with pipeline snapshot\n\n" +
    "DOCS: see your downloaded package docs/ folder for full setup guide.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ═══════════════════════════════════════════════════════════════
//  WIZARD DATA
// ═══════════════════════════════════════════════════════════════

function getWizardData() {
  ensureSheets_();
  return { version: LFU.VERSION, timezone: Session.getScriptTimeZone(), settings: getAllSettings_(), sheetNames: getSheetNames_() };
}

function saveWizardData(data) {
  ensureSheets_();
  if (!data) throw new Error("Missing wizard payload.");
  const mode = String(data.send_mode || "").toUpperCase().trim();
  if (mode !== "REVIEW" && mode !== "AUTO") throw new Error("send_mode must be REVIEW or AUTO.");
  const max = parseInt(data.max_emails_per_run, 10);
  if (isNaN(max) || max < 0 || max > 500) throw new Error("max_emails_per_run must be 0-500.");
  setAllSettings_(data);
  if (String(data.enable_daily_trigger || "FALSE").toUpperCase() === "TRUE") {
    const hour = parseInt(data.autopilot_hour_local, 10) || 9;
    installDailyTrigger_(hour);
  } else {
    uninstallTriggers_();
  }
  return { ok: true };
}

// ═══════════════════════════════════════════════════════════════
//  AUTOPILOT RUNNERS
// ═══════════════════════════════════════════════════════════════

function autopilotRunNow() {
  const cfg = getAllSettings_();
  const result = autopilotRun_(cfg.send_mode);
  logActivity_({ leadId:"", name:"", action:"Autopilot Run",
    notes:"Mode="+(result.mode||"")+" Due="+(result.due_count||0)+" Drafted="+(result.drafted||0)+" Sent="+(result.sent||0)+" Errors="+(result.errors||0) });
  safeUiAlert_(formatRunSummary_(result));
  return result;
}

function autopilotRunReviewNow() {
  const result = autopilotRun_("REVIEW");
  logActivity_({ leadId:"", name:"", action:"Autopilot Run",
    notes:"Mode=REVIEW Due="+(result.due_count||0)+" Drafted="+(result.drafted||0) });
  safeUiAlert_(formatRunSummary_(result));
  return result;
}

function sendApprovedDrafts() {
  const cfg   = getAllSettings_();
  const ss    = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(LFU.SHEET_LEADS);
  if (!sheet) throw new Error("Missing Leads sheet.");
  ensureLeadColumns_(sheet);
  const { headerIndex, rows } = readLeads_(sheet);
  const approved = rows.filter(r =>
    toBool_(r["Approved to Send"]) &&
    String(r["Draft"] || "").trim() &&
    ["DRAFTED","APPROVED"].indexOf(String(r["Draft Status"] || "").trim().toUpperCase()) !== -1
  );
  const maxParsed = parseInt(cfg.max_emails_per_run, 10);
  const max = isNaN(maxParsed) ? 10 : Math.max(0, maxParsed);
  const batch = approved.slice(0, max);
  let sent = 0, errors = 0;
  batch.forEach(lead => {
    try {
      let t = null, subject = "", body = "";
      const sendSource = String(cfg.review_send_source || "DRAFT").trim().toUpperCase();
      if (sendSource === "DRAFT") {
        const parsed = parseDraft_(lead["Draft"]);
        if (parsed && parsed.body) {
          subject = (parsed.subject && parsed.subject !== "(no subject)")
            ? parsed.subject : buildSubject_(lead, cfg, null);
          body = parsed.body;
        } else {
          t = pickTemplateForLead_(lead, cfg);
          subject = buildSubject_(lead, cfg, t);
          body = buildBody_(lead, cfg, t);
        }
      } else {
        t = pickTemplateForLead_(lead, cfg);
        subject = buildSubject_(lead, cfg, t);
        body = buildBody_(lead, cfg, t);
      }
      const to = String(lead["Email"] || "").trim();
      if (!to) throw new Error("Missing Email.");
      MailApp.sendEmail({ to, subject, body });
      updateLeadAfterSend_(sheet, headerIndex, lead, cfg, "SENT", t);
      logActivity_({ leadId:lead["Lead ID"], name:lead["Name"], action:"Email Sent", notes:"Subject: "+subject });
      sent++;
    } catch (e) {
      updateLeadError_(sheet, headerIndex, lead, e);
      errors++;
    }
  });
  logActivity_({ leadId:"", name:"", action:"Review Send Batch", notes:"Sent="+sent+" Errors="+errors });
  const msg = sent > 0
    ? "✅ Done. Sent: " + sent + (errors ? " | Errors: " + errors : "")
    : "No approved leads found.\n\nSet Approved to Send = TRUE on the rows you want to send.";
  safeUiAlert_(msg);
  return { ok:true, sent, errors, mode:"REVIEW_SEND" };
}

function autopilotRun_(mode) {
  ensureSheets_();
  const cfg = getAllSettings_();
  const effectiveMode = String(mode || cfg.send_mode || "REVIEW").toUpperCase();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(LFU.SHEET_LEADS);
  if (!sheet) throw new Error("Missing Leads sheet.");
  ensureLeadColumns_(sheet);
  const { headerIndex, rows } = readLeads_(sheet);
  const today = startOfDay_(new Date());

  // Recalculate due flags + priority scores
  rows.forEach(lead => {
    addComputedDueFlag_(lead, today);
    addPriorityScore_(lead);
  });
  writeColumnFromLeads_(sheet, headerIndex, rows, "Due Flag",
    lead => toBool_(lead["Due Flag"]) ? "TRUE" : "FALSE");
  writeColumnFromLeads_(sheet, headerIndex, rows, "Priority Score",
    lead => lead["Priority Score"] || 0);

  const maxParsed = parseInt(cfg.max_emails_per_run, 10);
  const max = isNaN(maxParsed) ? 10 : Math.max(0, maxParsed);
  const dueLeads = rows.filter(lead => toBool_(lead["Due Flag"]) === true);
  const batch = dueLeads.slice(0, max);

  let drafted = 0, sent = 0, errors = 0;
  batch.forEach(lead => {
    try {
      const t = pickTemplateForLead_(lead, cfg);
      const subject = buildSubject_(lead, cfg, t);
      const body = buildBody_(lead, cfg, t);
      updateLeadDraft_(sheet, headerIndex, lead, "Subject: " + subject + "\n\n" + body, effectiveMode);
      drafted++;
      if (effectiveMode === "AUTO") {
        const to = String(lead["Email"] || "").trim();
        if (!to) throw new Error("Missing Email.");
        MailApp.sendEmail({ to, subject, body });
        updateLeadAfterSend_(sheet, headerIndex, lead, cfg, "SENT", t);
        logActivity_({ leadId:lead["Lead ID"], name:lead["Name"], action:"Email Sent (AUTO)", notes:"Subject: "+subject });
        sent++;
      }
    } catch (e) {
      updateLeadError_(sheet, headerIndex, lead, e);
      errors++;
    }
  });

  // Run rules after autopilot if enabled
  if (String(cfg.rules_enabled || "TRUE").toUpperCase() === "TRUE") {
    try { runRulesEngine_("EVERY_RUN", null); } catch(e) { /* non-fatal */ }
  }

  return { ok:true, mode:effectiveMode, due_count:dueLeads.length, drafted, sent, errors };
}

// ═══════════════════════════════════════════════════════════════
//  TEMPLATE ENGINE  (FIX: supports {token} AND {{token}})
// ═══════════════════════════════════════════════════════════════

function pickTemplateForLead_(lead, cfg) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(LFU.SHEET_TEMPLATES);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) return null;
  const headers = data[0].map(h => String(h || "").trim());
  const idx = indexMap_(headers);
  const req = ["Pack","Status","Subject","Body","Delay Days","Active"];
  for (let i = 0; i < req.length; i++) {
    if (idx[req[i]] === undefined) return null;
  }
  const leadStatus = String(lead["Status"] || "").trim().toUpperCase();
  const leadPack = String(lead["Template Pack"] || "").trim()
    || String(cfg.default_template_pack || "general").trim();
  const trows = data.slice(1).map(r => ({
    pack:      String(r[idx["Pack"]]       || "").trim(),
    status:    String(r[idx["Status"]]     || "").trim().toUpperCase(),
    subject:   String(r[idx["Subject"]]    || "").trim(),
    body:      String(r[idx["Body"]]       || "").trim(),
    delayDays: String(r[idx["Delay Days"]] || "").trim(),
    active:    toBool_(r[idx["Active"]])
  })).filter(t => t.active);
  if (!trows.length) return null;
  const score = t => {
    const pm = (t.pack   === leadPack)   ? 2 : (t.pack   === "*" ? 1 : 0);
    const sm = (t.status === leadStatus) ? 2 : (t.status === "*" ? 1 : 0);
    return pm * 10 + sm;
  };
  let best = null, bestScore = -1;
  trows.forEach(t => { const s = score(t); if (s > bestScore) { best = t; bestScore = s; } });
  return (!best || bestScore <= 0) ? null : best;
}

function buildSubject_(lead, cfg, tpl) {
  const firstName = getFirstName_(lead["Preferred Name"] || lead["Name"]);
  const tmpl = (tpl && tpl.subject) ? tpl.subject
    : String(cfg.email_subject_template || LFU.SETTINGS_DEFAULTS.email_subject_template);
  return renderTemplate_(tmpl, lead, cfg, firstName);
}

function buildBody_(lead, cfg, tpl) {
  const firstName = getFirstName_(lead["Preferred Name"] || lead["Name"]);
  const tmpl = (tpl && tpl.body) ? tpl.body
    : String(cfg.email_body_template || LFU.SETTINGS_DEFAULTS.email_body_template);
  return renderTemplate_(tmpl, lead, cfg, firstName);
}

/**
 * Renders a template string, replacing tokens in both {token} and {{token}} formats.
 * Also supports optional spaces: {{ token }}.
 */
function renderTemplate_(tmpl, lead, cfg, firstName) {
  const company    = String(lead["Company"]    || "").trim();
  const senderName = String(cfg.sender_name || Session.getActiveUser().getEmail() || "").trim();
  const map = {
    first_name:  firstName   || "there",
    company:     company     || "your team",
    sender_name: senderName  || "",
    // appointment tokens (used in reminder emails)
    date:        String(lead["__date"]  || "").trim(),
    time:        String(lead["__time"]  || "").trim()
  };
  let out = String(tmpl || "");
  Object.keys(map).forEach(k => {
    const v = map[k];
    // Match {{token}}, {{ token }}, {token}, { token }
    const re = new RegExp("\\{\\{\\s*" + k + "\\s*\\}\\}|\\{\\s*" + k + "\\s*\\}", "g");
    out = out.replace(re, v);
  });
  return out;
}

// ═══════════════════════════════════════════════════════════════
//  LEAD PRIORITY SCORE (v8 NEW)
//  0-10: helps sort Today list by urgency, not just deal value
// ═══════════════════════════════════════════════════════════════

function addPriorityScore_(lead) {
  let score = 0;
  if (toBool_(lead["Due Flag"])) score += 3;
  const dv = parseFloat(String(lead["Deal Value"] || "0").replace(/[^0-9.]/g,"")) || 0;
  if (dv > 0) score += 1;
  if (dv >= 1000) score += 1;
  if (dv >= 5000) score += 1;
  const prob = parseFloat(String(lead["Probability %"] || "0")) || 0;
  if (prob >= 50) score += 1;
  if (prob >= 75) score += 1;
  if (String(lead["Phone"] || "").trim()) score += 1;
  const lc = asDate_(lead["Last Contacted"]);
  if (lc) {
    const daysSince = (new Date() - lc) / 86400000;
    if (daysSince <= 7) score += 1;
  }
  lead["Priority Score"] = Math.min(10, score);
  return lead;
}

// ═══════════════════════════════════════════════════════════════
//  RULES ENGINE LITE (v8 NEW)
//  Sheet: Rules
//  Columns: Rule ID | Name | Trigger Type | Trigger Value |
//           Action Type | Action Value | Priority | Active
//
//  Trigger Types:  STAGE_BECOMES | EVERY_RUN | STATUS_IS
//  Action Types:   SET_STATUS | SET_TEMPLATE_PACK | SNOOZE_DAYS |
//                  CLEAR_DUE | LOG_NOTE | CREATE_DRAFT | SET_PROBABILITY
// ═══════════════════════════════════════════════════════════════

function openRulesEngine_() {
  ensureSheets_();
  const ss = SpreadsheetApp.getActive();
  ss.setActiveSheet(ss.getSheetByName(LFU.SHEET_RULES));
  ss.toast("Rules Engine: add/edit rules here. Rules run on stage changes and autopilot runs.", "Rules Engine", 5);
}

function runRulesNow_() {
  ensureSheets_();
  const result = runRulesEngine_("EVERY_RUN", null);
  safeUiAlert_("Rules run complete.\nMatched: " + (result.matched||0) + " | Actions: " + (result.actions||0));
}

/**
 * Core rules engine. Called after stage changes and during autopilot.
 * @param {string} triggerType - "STAGE_BECOMES" | "EVERY_RUN" | "STATUS_IS"
 * @param {string|null} triggerValue - the stage/status value, or null for EVERY_RUN
 * @param {number|null} specificLeadRow - if set, only apply to this row
 */
function runRulesEngine_(triggerType, triggerValue, specificLeadRow) {
  const cfg = getAllSettings_();
  if (String(cfg.rules_enabled || "TRUE").toUpperCase() !== "TRUE") return { matched:0, actions:0 };

  const ss = SpreadsheetApp.getActive();
  const rulesSheet = ss.getSheetByName(LFU.SHEET_RULES);
  const leadsSheet = ss.getSheetByName(LFU.SHEET_LEADS);
  if (!rulesSheet || !leadsSheet) return { matched:0, actions:0 };

  ensureRulesSheet_(rulesSheet);
  ensureLeadColumns_(leadsSheet);

  // Load rules
  const rdata = rulesSheet.getDataRange().getValues();
  if (rdata.length < 2) return { matched:0, actions:0 };
  const rh = rdata[0].map(h => String(h||"").trim());
  const ri = indexMap_(rh);

  const rules = rdata.slice(1)
    .filter(r => toBool_(r[ri["Active"]]))
    .map(r => ({
      id:           String(r[ri["Rule ID"]]       || "").trim(),
      name:         String(r[ri["Name"]]          || "").trim(),
      triggerType:  String(r[ri["Trigger Type"]]  || "").trim().toUpperCase(),
      triggerValue: String(r[ri["Trigger Value"]] || "").trim().toUpperCase(),
      actionType:   String(r[ri["Action Type"]]   || "").trim().toUpperCase(),
      actionValue:  String(r[ri["Action Value"]]  || "").trim(),
      priority:     parseInt(String(r[ri["Priority"]] || "0"),10) || 0
    }))
    .filter(r => r.triggerType && r.actionType)
    .sort((a,b) => a.priority - b.priority);

  if (!rules.length) return { matched:0, actions:0 };

  // Load leads
  const { headerIndex, rows } = readLeads_(leadsSheet);
  const today = startOfDay_(new Date());
  let matched = 0, actions = 0;

  const applyRule = (rule, lead) => {
    const row = lead.__row;
    const status = String(lead["Status"] || "").trim().toUpperCase();
    const businessOnly = toBool_(cfg.business_days_only);

    switch (rule.actionType) {
      case "SET_STATUS":
        if (status !== rule.actionValue.toUpperCase()) {
          setCell_(leadsSheet, row, headerIndex, "Status", rule.actionValue);
          logActivity_({ leadId:lead["Lead ID"]||"", name:lead["Name"]||"",
            action:"Rule: "+rule.name, notes:"SET_STATUS="+rule.actionValue });
          actions++;
        }
        break;
      case "SET_TEMPLATE_PACK":
        setCell_(leadsSheet, row, headerIndex, "Template Pack", rule.actionValue);
        actions++;
        break;
      case "SNOOZE_DAYS": {
        const d = Math.max(1, parseInt(rule.actionValue,10)||1);
        const next = nextFollowUpDate_(today, d, businessOnly);
        setCell_(leadsSheet, row, headerIndex, "Follow Up Due", next);
        setCell_(leadsSheet, row, headerIndex, "Due Flag", "FALSE");
        setCell_(leadsSheet, row, headerIndex, "Approved to Send", "");
        logActivity_({ leadId:lead["Lead ID"]||"", name:lead["Name"]||"",
          action:"Rule: "+rule.name, notes:"SNOOZE="+d+"d" });
        actions++;
        break;
      }
      case "CLEAR_DUE":
        setCell_(leadsSheet, row, headerIndex, "Due Flag", "FALSE");
        setCell_(leadsSheet, row, headerIndex, "Approved to Send", "");
        actions++;
        break;
      case "LOG_NOTE": {
        const existing = String(lead["Notes"] || "").trim();
        const note = rule.actionValue + " [" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(),"yyyy-MM-dd") + "]";
        setCell_(leadsSheet, row, headerIndex, "Notes", existing ? existing + "\n" + note : note);
        logActivity_({ leadId:lead["Lead ID"]||"", name:lead["Name"]||"",
          action:"Rule: "+rule.name, notes:note });
        actions++;
        break;
      }
      case "CREATE_DRAFT": {
        const t = pickTemplateForLead_(lead, cfg);
        const subject = buildSubject_(lead, cfg, t);
        const body = buildBody_(lead, cfg, t);
        setCell_(leadsSheet, row, headerIndex, "Draft", "Subject: "+subject+"\n\n"+body);
        setCell_(leadsSheet, row, headerIndex, "Draft Status", "DRAFTED");
        setCell_(leadsSheet, row, headerIndex, "Last Drafted", new Date());
        logActivity_({ leadId:lead["Lead ID"]||"", name:lead["Name"]||"",
          action:"Rule: "+rule.name, notes:"Draft created" });
        actions++;
        break;
      }
      case "SET_PROBABILITY": {
        const prob = Math.min(100, Math.max(0, parseInt(rule.actionValue,10)||0));
        setCell_(leadsSheet, row, headerIndex, "Probability %", prob);
        actions++;
        break;
      }
    }
  };

  rows.forEach(lead => {
    if (specificLeadRow !== null && specificLeadRow !== undefined && lead.__row !== specificLeadRow) return;
    const status = String(lead["Status"] || "").trim().toUpperCase();
    const closedOrLost = [...LFU.CLOSED_STATUSES, ...LFU.LOST_STATUSES].map(s=>s.toUpperCase());

    rules.forEach(rule => {
      let matches = false;
      switch (rule.triggerType) {
        case "STAGE_BECOMES":
          // For STAGE_BECOMES, only matches when called with that specific trigger
          if (triggerType === "STAGE_BECOMES" && triggerValue &&
              rule.triggerValue === triggerValue.toUpperCase() &&
              (specificLeadRow === undefined || specificLeadRow === null || specificLeadRow === lead.__row)) {
            matches = true;
          }
          break;
        case "STATUS_IS":
          if (status === rule.triggerValue) matches = true;
          break;
        case "EVERY_RUN":
          matches = true;
          break;
      }
      if (matches) {
        matched++;
        try { applyRule(rule, lead); } catch(e) { /* non-fatal */ }
      }
    });
  });

  return { matched, actions };
}

function ensureRulesSheet_(sheet) {
  if (!sheet) return;
  const lastCol = Math.max(sheet.getLastColumn(),1);
  const existing = sheet.getRange(1,1,1,lastCol).getValues()[0].map(h=>String(h||"").trim()).filter(Boolean);
  const H = ["Rule ID","Name","Trigger Type","Trigger Value","Action Type","Action Value","Priority","Active"];
  if (existing.length === 0) {
    sheet.clear();
    sheet.getRange(1,1,1,H.length).setValues([H]);
    sheet.setFrozenRows(1);
  }
  if (sheet.getLastRow() < 2) {
    const seeds = [
      ["RULE-001","No-Show Snooze",         "STAGE_BECOMES","NO_SHOW",   "SNOOZE_DAYS",       "1",           "1","TRUE"],
      ["RULE-002","No-Show Template Pack",  "STAGE_BECOMES","NO_SHOW",   "SET_TEMPLATE_PACK", "no_show",     "2","TRUE"],
      ["RULE-003","No-Show Draft",          "STAGE_BECOMES","NO_SHOW",   "CREATE_DRAFT",      "",            "3","TRUE"],
      ["RULE-004","Booked Clear Due",       "STAGE_BECOMES","BOOKED",    "CLEAR_DUE",         "",            "1","TRUE"],
      ["RULE-005","Closed Clear Due",       "STAGE_BECOMES","CLOSED",    "CLEAR_DUE",         "",            "1","TRUE"],
      ["RULE-006","Lost Clear Due",         "STAGE_BECOMES","LOST",      "CLEAR_DUE",         "",            "1","TRUE"],
      ["RULE-007","Lost Note",              "STAGE_BECOMES","LOST",      "LOG_NOTE",          "Lead marked lost", "2","TRUE"],
      ["RULE-008","Scheduled Clear Due",    "STAGE_BECOMES","SCHEDULED", "CLEAR_DUE",         "",            "1","TRUE"],
      ["RULE-009","High Prob Log",          "STATUS_IS",    "CONTACTED", "SET_PROBABILITY",   "50",          "5","FALSE"]
    ];
    sheet.getRange(2,1,seeds.length,H.length).setValues(seeds);
    sheet.autoResizeColumns(1,H.length);
    // Add note row
    sheet.getRange("A1").setNote(
      "Trigger Types: STAGE_BECOMES | STATUS_IS | EVERY_RUN\n" +
      "Action Types: SET_STATUS | SET_TEMPLATE_PACK | SNOOZE_DAYS | CLEAR_DUE | LOG_NOTE | CREATE_DRAFT | SET_PROBABILITY\n" +
      "Priority: lower number runs first\n" +
      "Active: TRUE/FALSE"
    );
  }
}

// ═══════════════════════════════════════════════════════════════
//  SHEET SETUP
// ═══════════════════════════════════════════════════════════════

function ensureSheets_() {
  const ss = SpreadsheetApp.getActive();
  const needed = [
    LFU.SHEET_LEADS, LFU.SHEET_SETTINGS, LFU.SHEET_TEMPLATES,
    LFU.SHEET_ACTIVITIES, LFU.SHEET_SYSOPS, LFU.SHEET_SOPS,
    LFU.SHEET_IMPROVE, LFU.SHEET_SOP_BUILDER, LFU.SHEET_SOP_TMPL,
    LFU.SHEET_RULES, LFU.SHEET_IMPORT, LFU.SHEET_INTAKE, LFU.SHEET_APPOINTMENTS
  ];
  needed.forEach(name => { if (!ss.getSheetByName(name)) ss.insertSheet(name); });
  ensureLeadColumns_         (ss.getSheetByName(LFU.SHEET_LEADS));
  ensureSettingsSheet_       (ss.getSheetByName(LFU.SHEET_SETTINGS));
  ensureTemplatesSheet_      (ss.getSheetByName(LFU.SHEET_TEMPLATES));
  ensureActivitiesSheet_     (ss.getSheetByName(LFU.SHEET_ACTIVITIES));
  ensureSysOpsSheet_         (ss.getSheetByName(LFU.SHEET_SYSOPS));
  ensureSopSheet_            (ss.getSheetByName(LFU.SHEET_SOPS));
  ensureProcessImprovementsSheet_(ss.getSheetByName(LFU.SHEET_IMPROVE));
  ensureSopBuilderSheet_     (ss.getSheetByName(LFU.SHEET_SOP_BUILDER));
  ensureSopTemplatesSheet_   (ss.getSheetByName(LFU.SHEET_SOP_TMPL));
  ensureRulesSheet_          (ss.getSheetByName(LFU.SHEET_RULES));
  ensureImportSheet_         (ss.getSheetByName(LFU.SHEET_IMPORT));
  ensureIntakeSheet_         (ss.getSheetByName(LFU.SHEET_INTAKE));
  ensureAppointmentsSheet_   (ss.getSheetByName(LFU.SHEET_APPOINTMENTS));
}

function ensureLeadColumns_(sheet) {
  if (!sheet) return;
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1,1,1,lastCol).getValues()[0].map(h => String(h||"").trim());
  const existing = new Set(headers.filter(Boolean));
  if (headers.filter(Boolean).length === 0) {
    sheet.getRange(1,1,1,LFU.HEADERS_REQUIRED.length).setValues([LFU.HEADERS_REQUIRED]);
    sheet.setFrozenRows(1);
    return;
  }
  const missing = LFU.HEADERS_REQUIRED.filter(h => !existing.has(h));
  if (missing.length > 0) {
    sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  }
}

function ensureTemplatesSheet_(sheet) {
  if (!sheet) return;
  const lastCol = Math.max(sheet.getLastColumn(),1);
  const existH = sheet.getRange(1,1,1,lastCol).getValues()[0].map(h=>String(h||"").trim()).filter(Boolean);
  const headers = ["Pack","Status","Subject","Body","Delay Days","Active"];
  if (existH.length === 0) {
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  if (sheet.getLastRow() < 2) {
    const seeds = [
      ["general",        "NEW",       "Quick follow up, {first_name}",        "Hi {first_name},\n\nJust following up on {company}.\n\nBest,\n{sender_name}",                              "3","TRUE"],
      ["general",        "*",         "Following up, {first_name}",            "Hi {first_name},\n\nCircling back on {company}.\n\nBest,\n{sender_name}",                                "3","TRUE"],
      ["real_estate",    "NEW",       "Quick question, {first_name}",          "Hi {first_name},\n\nAre you still looking in the next 30-60 days? Happy to send a few options.\n\nBest,\n{sender_name}",     "2","TRUE"],
      ["agency",         "CONTACTED", "Next steps for {company}",              "Hi {first_name},\n\nI can outline a simple 2-step plan for {company} and ballpark timelines.\n\nBest,\n{sender_name}",      "2","TRUE"],
      ["local_services", "NEW",       "Scheduling help, {first_name}",         "Hi {first_name},\n\nWant me to help you pick a time this week for {company}?\n\nBest,\n{sender_name}",                       "1","TRUE"],
      ["no_show",        "*",         "Missed our appointment, {first_name}",  "Hi {first_name},\n\nWe missed each other. Happy to reschedule at a time that works better.\n\nBest,\n{sender_name}",         "1","TRUE"],
      ["*",              "*",         "Follow up, {first_name}",               "Hi {first_name},\n\nJust following up.\n\nBest,\n{sender_name}",                                          "3","TRUE"]
    ];
    sheet.getRange(2,1,seeds.length,headers.length).setValues(seeds);
    sheet.autoResizeColumns(1,headers.length);
  }
}

function ensureActivitiesSheet_(sheet) {
  if (!sheet) return;
  const lastCol = Math.max(sheet.getLastColumn(),1);
  const existing = sheet.getRange(1,1,1,lastCol).getValues()[0].map(h=>String(h||"").trim()).filter(Boolean);
  if (existing.length === 0) {
    const headers = ["Timestamp","Lead ID","Name","Action Type","Notes","Performed By"];
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1,160); sheet.setColumnWidth(4,160); sheet.setColumnWidth(5,320);
  }
}

// ═══════════════════════════════════════════════════════════════
//  ACTIVITY LOG
// ═══════════════════════════════════════════════════════════════


// ═══════════════════════════════════════════════════════════════
//  TRACK A ADDITIONS — required by WebApp.gs
// ═══════════════════════════════════════════════════════════════

/**
 * isQualifiedLeadRow_ — filters out blank/header rows from readLeads_ output.
 * A row is qualified if it has a non-empty Name field.
 * Used by uiGetStateReadOnly in WebApp.gs.
 */
/**
 * isQualifiedLeadRow_ — returns true if any of the primary identity fields
 * (Lead ID, Name, Email, Phone, Company) are non-empty.
 * This prevents blank sheet rows from appearing as leads while still
 * accepting email-only, phone-only, or company-only entries.
 */
function isQualifiedLeadRow_(r) {
  if (!r) return false;
  return (
    String(r["Lead ID"]  || "").trim().length > 0 ||
    String(r["Name"]     || "").trim().length > 0 ||
    String(r["Email"]    || "").trim().length > 0 ||
    String(r["Phone"]    || "").trim().length > 0 ||
    String(r["Company"]  || "").trim().length > 0
  );
}

/**
 * safeLogActivity_ — non-throwing wrapper around logActivity_.
 * PRIVACY RULE: never log raw technical error text in the notes field.
 * For errors: log only debugId + action code. Technical stays in console.error.
 */
function safeLogActivity_(opts) {
  try {
    logActivity_(opts);
  } catch (e) {
    console.warn("[safeLogActivity_] logging failed: " + (e && e.message ? e.message : String(e)));
  }
}

/**
 * isCalendarEnabled_() — returns true only when enable_calendar is explicitly "TRUE"
 * in the Settings sheet. Defaults to false so new installs never prompt for Calendar
 * OAuth until the buyer deliberately turns this on.
 */
function isCalendarEnabled_() {
  try {
    var cfg = getAllSettings_();
    return String(cfg.enable_calendar || "").trim().toUpperCase() === "TRUE";
  } catch (e) {
    return false; // fail safe: if settings unreadable, disable calendar
  }
}

function logActivity_(opts) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(LFU.SHEET_ACTIVITIES);
    if (!sheet) return;
    sheet.appendRow([
      new Date(),
      opts && opts.leadId ? opts.leadId : "",
      opts && opts.name   ? opts.name   : "",
      opts && opts.action ? opts.action : "",
      opts && opts.notes  ? opts.notes  : "",
      Session.getActiveUser().getEmail() || ""
    ]);
  } catch(e) { /* non-fatal */ }
}

// ═══════════════════════════════════════════════════════════════
//  SHEET READ / WRITE
// ═══════════════════════════════════════════════════════════════

function readLeads_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1,1,1,lastCol).getValues()[0].map(h=>String(h||"").trim());
  const headerIndex = {};
  headers.forEach((h,i) => { if (h) headerIndex[h] = i; });
  if (lastRow < 2) return { headerIndex, rows:[] };
  const values = sheet.getRange(2,1,lastRow-1,lastCol).getValues();
  const rows = values.map((row,idx) => {
    const obj = { __row: idx+2 };
    headers.forEach((h,ci) => { if (h) obj[h] = row[ci]; });
    return obj;
  });
  return { headerIndex, rows };
}

function writeColumnFromLeads_(sheet, headerIndex, leads, colName, getterFn) {
  const colIdx = headerIndex[colName];
  if (colIdx === undefined || !leads.length) return;
  sheet.getRange(2, colIdx+1, leads.length, 1).setValues(leads.map(l => [getterFn(l)]));
}

function setCell_(sheet, row, headerIndex, headerName, value) {
  const idx = headerIndex[headerName];
  if (idx === undefined) return;
  sheet.getRange(row, idx+1).setValue(value);
}

/**
 * Appends a new lead row to the Leads sheet.
 * Auto-generates Lead ID if not provided.
 * Returns the new sheet row number.
 */
function appendLeadRow_(sheet, headerIndex, upd) {
  if (!upd["Lead ID"] || String(upd["Lead ID"]).trim() === "") {
    upd["Lead ID"] = "L-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
  }
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1,1,1,lastCol).getValues()[0].map(h=>String(h||"").trim());
  const rowValues = headers.map(h => (upd[h] !== undefined ? upd[h] : ""));
  sheet.appendRow(rowValues);
  return sheet.getLastRow();
}

function updateLeadDraft_(sheet, headerIndex, lead, draft, mode) {
  const row = lead.__row, now = new Date();
  setCell_(sheet, row, headerIndex, "Draft",            draft);
  setCell_(sheet, row, headerIndex, "Draft Status",     mode === "AUTO" ? "SENT_PENDING" : "DRAFTED");
  setCell_(sheet, row, headerIndex, "Approved to Send", "");
  setCell_(sheet, row, headerIndex, "Last Error",       "");
  setCell_(sheet, row, headerIndex, "Last Drafted",     now);
}

function updateLeadAfterSend_(sheet, headerIndex, lead, cfg, statusLabel, tpl) {
  const row = lead.__row, now = new Date();
  setCell_(sheet, row, headerIndex, "Draft Status",     statusLabel);
  setCell_(sheet, row, headerIndex, "Last Sent",        now);
  setCell_(sheet, row, headerIndex, "Last Contacted",   now);
  setCell_(sheet, row, headerIndex, "Last Error",       "");
  const businessOnly = toBool_(cfg.business_days_only);
  const override  = parseInt(String(lead["Cadence Override Days"] || "").trim(),10);
  const tplDelay  = tpl ? parseInt(String(tpl.delayDays || "").trim(),10) : NaN;
  const cfgDelay  = parseInt(cfg.follow_up_days,10);
  const days = Math.max(1, !isNaN(override) ? override : !isNaN(tplDelay) ? tplDelay : isNaN(cfgDelay) ? 3 : cfgDelay);
  setCell_(sheet, row, headerIndex, "Follow Up Due",    nextFollowUpDate_(startOfDay_(now), days, businessOnly));
  setCell_(sheet, row, headerIndex, "Due Flag",         "FALSE");
  setCell_(sheet, row, headerIndex, "Approved to Send", "");
}

function updateLeadError_(sheet, headerIndex, lead, err) {
  setCell_(sheet, lead.__row, headerIndex, "Draft Status", "ERROR");
  setCell_(sheet, lead.__row, headerIndex, "Last Error",   (err && err.message) ? err.message : String(err));
}

// ═══════════════════════════════════════════════════════════════
//  DUE LOGIC
// ═══════════════════════════════════════════════════════════════

function addComputedDueFlag_(lead, today) {
  const status  = String(lead["Status"] || "").trim().toUpperCase();
  const email   = String(lead["Email"]  || "").trim();
  const dueDate = asDate_(lead["Follow Up Due"]);
  const closedStatuses = [...LFU.CLOSED_STATUSES, ...LFU.LOST_STATUSES];
  const phone   = String(lead["Phone"]  || "").trim();
  const hasContact = email || phone;
  const due = hasContact && dueDate &&
    startOfDay_(dueDate).getTime() <= today.getTime() &&
    closedStatuses.indexOf(status) === -1;
  lead["Due Flag"] = due ? "TRUE" : "FALSE";
  return lead;
}

// ═══════════════════════════════════════════════════════════════
//  SETTINGS
// ═══════════════════════════════════════════════════════════════

function ensureSettingsSheet_(sheet) {
  if (!sheet) return;
  const vals = sheet.getDataRange().getValues();
  const hasData = vals.length >= 1 && vals.some(r => String(r[0]||"").trim());
  if (!hasData) {
    const keys = Object.keys(LFU.SETTINGS_DEFAULTS);
    sheet.getRange(1,1,keys.length,2).setValues(keys.map(k=>[k,LFU.SETTINGS_DEFAULTS[k]]));
    sheet.setFrozenRows(1);
  }
}

function getAllSettings_() {
  const merged = Object.assign({}, LFU.SETTINGS_DEFAULTS,
    PropertiesService.getScriptProperties().getProperties());
  const settingsSheet = SpreadsheetApp.getActive().getSheetByName(LFU.SHEET_SETTINGS);
  if (settingsSheet) Object.assign(merged, readSettingsSheet_(settingsSheet));
  const isDefault = merged.pipeline_stages === LFU.SETTINGS_DEFAULTS.pipeline_stages || !merged.pipeline_stages;
  return isDefault ? applyStagePack_(merged, null) : merged;
}

function setAllSettings_(data) {
  const props = PropertiesService.getScriptProperties();
  const clean = {};
  Object.keys(LFU.SETTINGS_DEFAULTS).forEach(k => {
    if (data[k] !== undefined && data[k] !== null && String(data[k]).trim() !== "")
      clean[k] = String(data[k]).trim();
  });
  if (data.enable_daily_trigger !== undefined) clean.enable_daily_trigger = String(data.enable_daily_trigger).trim();
  if (data.sender_name && String(data.sender_name).trim()) clean.sender_name = String(data.sender_name).trim();
  props.setProperties(clean, true);
  const settingsSheet = SpreadsheetApp.getActive().getSheetByName(LFU.SHEET_SETTINGS);
  if (settingsSheet) writeSettingsSheet_(settingsSheet, Object.assign({}, LFU.SETTINGS_DEFAULTS, clean));
}

function readSettingsSheet_(sheet) {
  const out = {};
  sheet.getDataRange().getValues().forEach(r => {
    const k = String(r[0]||"").trim();
    if (!k) return;
    out[k] = (r[1] === undefined || r[1] === null) ? "" : String(r[1]).trim();
  });
  return out;
}

function writeSettingsSheet_(sheet, obj) {
  const keys = Object.keys(obj);
  sheet.clearContents();
  sheet.getRange(1,1,keys.length,2).setValues(keys.map(k=>[k,obj[k]]));
}

// ═══════════════════════════════════════════════════════════════
//  TRIGGERS
// ═══════════════════════════════════════════════════════════════

function installDailyTrigger_(hourLocal) {
  uninstallTriggers_();
  ScriptApp.newTrigger("autopilotRunNow").timeBased().everyDays(1).atHour(hourLocal || 9).create();
}

function uninstallTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "autopilotRunNow") ScriptApp.deleteTrigger(t);
  });
}

// ═══════════════════════════════════════════════════════════════
//  CRM SIDEBAR / COMMAND CENTER SERVER ENDPOINTS
// ═══════════════════════════════════════════════════════════════

function uiGetState() {
  ensureSheets_();
  const cfg      = getAllSettings_();
  const pipeline = uiGetPipelineConfig();
  const ss       = SpreadsheetApp.getActive();
  const sheet    = ss.getSheetByName(LFU.SHEET_LEADS);
  ensureLeadColumns_(sheet);
  const { headerIndex, rows } = readLeads_(sheet);
  const today = startOfDay_(new Date());

  rows.forEach(lead => { addComputedDueFlag_(lead, today); addPriorityScore_(lead); });
  writeColumnFromLeads_(sheet, headerIndex, rows, "Due Flag",
    lead => toBool_(lead["Due Flag"]) ? "TRUE" : "FALSE");
  writeColumnFromLeads_(sheet, headerIndex, rows, "Priority Score",
    lead => lead["Priority Score"] || 0);

  const { rows: rows2 } = readLeads_(sheet);

  const leads = rows2.map(r => {
    const name          = String(r["Name"]           || "").trim();
    const preferredName = String(r["Preferred Name"] || "").trim();
    const firstName     = getFirstName_(preferredName || name);
    const status        = (String(r["Status"] || "NEW").trim().toUpperCase()) || "NEW";
    const dueFlag       = toBool_(r["Due Flag"]);
    const draftStatus   = String(r["Draft Status"]   || "").trim().toUpperCase();
    const approved      = toBool_(r["Approved to Send"]);
    const phone         = String(r["Phone"]          || "").trim();
    const email         = String(r["Email"]          || "").trim();
    const company       = String(r["Company"]        || "").trim();
    const owner         = String(r["Owner"]          || "").trim();
    const dealValue     = String(r["Deal Value"]     || "").trim();
    const expectedValue = String(r["Expected Value"] || "").trim();
    const probability   = String(r["Probability %"]  || "").trim();
    const templatePack  = String(r["Template Pack"]  || "").trim();
    const lastError     = String(r["Last Error"]     || "").trim();
    const leadId        = String(r["Lead ID"]        || "").trim();
    const notes         = String(r["Notes"]          || "").trim();
    const nextAction    = String(r["Next Action"]    || "").trim();
    const priorityScore = parseInt(String(r["Priority Score"] || "0"),10) || 0;
    const draft         = String(r["Draft"]          || "").trim();
    const apptStart     = r["Appointment Start"] ? Utilities.formatDate(
      asDate_(r["Appointment Start"]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : "";
    const calLink       = String(r["Calendar Event Link"] || "").trim();
    const followUpDue   = r["Follow Up Due"]
      ? Utilities.formatDate(asDate_(r["Follow Up Due"]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
    const lastContacted = r["Last Contacted"]
      ? Utilities.formatDate(asDate_(r["Last Contacted"]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
    return {
      row: r.__row, leadId, name, preferredName, firstName, status,
      dueFlag, draftStatus, approved, phone, email, company, owner,
      dealValue, expectedValue, probability, templatePack, lastError,
      notes, nextAction, priorityScore, draft, apptStart, calLink,
      followUpDue, lastContacted,
      textPreview: renderTemplate_(
        String(pipeline.textTemplate || ""),
        { Name:name, Company:company, "Preferred Name":preferredName },
        cfg, firstName
      )
    };
  });

  const toNum = s => parseFloat(String(s||"0").replace(/[^0-9.]/g,"")) || 0;
  const openLeads = leads.filter(l => [...LFU.CLOSED_STATUSES,...LFU.LOST_STATUSES].indexOf(l.status) === -1);
  const pipelineValue    = openLeads.reduce((sum,l) => sum + toNum(l.dealValue), 0);
  const weightedPipeline = openLeads.reduce((sum,l) => {
    const dv = toNum(l.dealValue), ev = toNum(l.expectedValue);
    const prob = ev && dv ? ev/dv : toNum(l.probability)/100;
    return sum + dv * prob;
  }, 0);
  const overdueValue = leads.filter(l=>l.dueFlag).reduce((sum,l) => sum+toNum(l.dealValue),0);

  const kpis = {
    total:            leads.length,
    due:              leads.filter(l=>l.dueFlag).length,
    drafted:          leads.filter(l=>l.draftStatus==="DRAFTED").length,
    approved:         leads.filter(l=>l.approved).length,
    errors:           leads.filter(l=>l.draftStatus==="ERROR"||l.lastError.length>0).length,
    pipelineValue:    Math.round(pipelineValue),
    weightedPipeline: Math.round(weightedPipeline),
    overdueValue:     Math.round(overdueValue)
  };

  return { version:LFU.VERSION, kpis, leads, pipeline };
}

/** Returns last N activities, optionally filtered by leadId */
function uiGetActivities(leadId, limit) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(LFU.SHEET_ACTIVITIES);
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const lim = Math.min(parseInt(limit,10)||50, 200);
    const startRow = Math.max(2, lastRow - 200);
    const data = sheet.getRange(startRow,1,lastRow-startRow+1,6).getValues().reverse();
    const results = [];
    for (let i = 0; i < data.length && results.length < lim; i++) {
      const r = data[i];
      const rLeadId = String(r[1]||"").trim();
      if (leadId && rLeadId !== leadId) continue;
      const ts = r[0] ? Utilities.formatDate(asDate_(r[0]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : "";
      results.push({ ts, leadId:rLeadId, name:String(r[2]||""), action:String(r[3]||""), notes:String(r[4]||"") });
    }
    return results;
  } catch(e) { return []; }
}

function uiUpdateLead(row, updates) {
  if (!row || !updates) throw new Error("Missing row or updates.");
  const sheet = SpreadsheetApp.getActive().getSheetByName(LFU.SHEET_LEADS);
  ensureLeadColumns_(sheet);
  const { headerIndex } = readLeads_(sheet);
  Object.keys(updates).forEach(k => setCell_(sheet, row, headerIndex, k, updates[k]));
  return { ok:true };
}

/** Quick Add Lead from Command Center */
function uiQuickAddLead(fields) {
  ensureSheets_();
  if (!fields) throw new Error("Missing lead fields.");
  const ss    = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(LFU.SHEET_LEADS);
  ensureLeadColumns_(sheet);
  const { headerIndex } = readLeads_(sheet);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const upd = {
    "Name":          String(fields.name    || "").trim(),
    "Email":         String(fields.email   || "").trim(),
    "Phone":         String(fields.phone   || "").trim(),
    "Company":       String(fields.company || "").trim(),
    "Deal Value":    String(fields.dealValue || "").trim(),
    "Status":        String(fields.status  || "NEW").trim().toUpperCase(),
    "Follow Up Due": today,
    "Owner":         String(fields.owner   || "").trim(),
    "Notes":         String(fields.notes   || "").trim()
  };
  if (!upd["Name"] && !upd["Email"]) throw new Error("Name or Email is required.");
  const newRow = appendLeadRow_(sheet, headerIndex, upd);
  logActivity_({ leadId:"", name:upd["Name"], action:"Quick Add Lead", notes:"Via Command Center" });
  return { ok:true, row:newRow };
}

function uiSnooze(row, days) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(LFU.SHEET_LEADS);
  ensureLeadColumns_(sheet);
  const { headerIndex, rows } = readLeads_(sheet);
  const d = Math.max(1, parseInt(days,10) || 1);
  const businessOnly = toBool_(getAllSettings_().business_days_only);
  const next = nextFollowUpDate_(startOfDay_(new Date()), d, businessOnly);
  setCell_(sheet, row, headerIndex, "Follow Up Due",    next);
  setCell_(sheet, row, headerIndex, "Approved to Send", "");
  setCell_(sheet, row, headerIndex, "Due Flag",         "FALSE");
  const lead = rows.find(r => r.__row === row);
  logActivity_({ leadId:lead?lead["Lead ID"]:"", name:lead?lead["Name"]:"",
    action:"Snoozed", notes:"Snoozed "+d+"d. Next: "+Utilities.formatDate(next,Session.getScriptTimeZone(),"yyyy-MM-dd") });
  return { ok:true };
}

function uiMarkCalled(row) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(LFU.SHEET_LEADS);
  ensureLeadColumns_(sheet);
  const cfg = getAllSettings_();
  const { headerIndex, rows } = readLeads_(sheet);
  const now = new Date();
  setCell_(sheet, row, headerIndex, "Last Contacted",   now);
  setCell_(sheet, row, headerIndex, "Last Error",       "");
  setCell_(sheet, row, headerIndex, "Approved to Send", "");
  const lead     = rows.find(r => r.__row === row);
  const override = lead ? parseInt(String(lead["Cadence Override Days"]||"").trim(),10) : NaN;
  const cfgDelay = parseInt(cfg.follow_up_days,10);
  const days     = Math.max(1, !isNaN(override) ? override : isNaN(cfgDelay) ? 3 : cfgDelay);
  const next     = nextFollowUpDate_(startOfDay_(now), days, toBool_(cfg.business_days_only));
  setCell_(sheet, row, headerIndex, "Follow Up Due", next);
  setCell_(sheet, row, headerIndex, "Due Flag",      "FALSE");
  logActivity_({ leadId:lead?lead["Lead ID"]:"", name:lead?lead["Name"]:"",
    action:"Called", notes:"Marked called. Next: "+Utilities.formatDate(next,Session.getScriptTimeZone(),"yyyy-MM-dd") });
  return { ok:true };
}

function uiRunReview()    { return autopilotRun_("REVIEW"); }
function uiSendApproved() { return sendApprovedDrafts(); }

function uiMoveLead(row, newStatus) {
  if (!row) throw new Error("Missing row.");
  const status = String(newStatus||"").trim().toUpperCase();
  if (!status) throw new Error("Missing status.");
  const sheet = SpreadsheetApp.getActive().getSheetByName(LFU.SHEET_LEADS);
  ensureLeadColumns_(sheet);
  const { headerIndex, rows } = readLeads_(sheet);
  const now  = new Date();
  const lead = rows.find(r => r.__row === row);
  const prev = lead ? String(lead["Status"]||"").trim().toUpperCase() : "";
  setCell_(sheet, row, headerIndex, "Status", status);
  if (LFU.WON_STATUSES.indexOf(status) !== -1) {
    setCell_(sheet, row, headerIndex, "Won Date",         now);
    setCell_(sheet, row, headerIndex, "Approved to Send", "");
    setCell_(sheet, row, headerIndex, "Due Flag",         "FALSE");
  } else if (LFU.LOST_STATUSES.indexOf(status) !== -1) {
    setCell_(sheet, row, headerIndex, "Lost Date",        now);
    setCell_(sheet, row, headerIndex, "Approved to Send", "");
    setCell_(sheet, row, headerIndex, "Due Flag",         "FALSE");
  } else if (LFU.CLOSED_STATUSES.indexOf(status) !== -1) {
    setCell_(sheet, row, headerIndex, "Approved to Send", "");
    setCell_(sheet, row, headerIndex, "Due Flag",         "FALSE");
  }
  logActivity_({ leadId:lead?lead["Lead ID"]:"", name:lead?lead["Name"]:"",
    action:"Stage Move", notes:(prev||"?")+" → "+status });
  // Run rules for STAGE_BECOMES trigger
  try { runRulesEngine_("STAGE_BECOMES", status, row); } catch(e) { /* non-fatal */ }
  return { ok:true };
}

/**
 * uiScheduleLead — public wrapper callable from google.script.run.
 * Gated by enable_calendar setting. Returns FEATURE_DISABLED if off.
 */
function uiScheduleLead(row, dateStr, timeStr, durationMin, location) {
  if (!isCalendarEnabled_()) {
    return {
      ok: false,
      error: {
        code:        "FEATURE_DISABLED",
        userMessage: "Scheduling is disabled. Enable it in Settings (enable_calendar = TRUE) to use this feature.",
        debugId:     Utilities.getUuid()
      }
    };
  }
  return uiScheduleLead_(row, dateStr, timeStr, durationMin, location);
}

function uiGetPipelineConfig() {
  const cfg = getAllSettings_();
  const stages = String(cfg.pipeline_stages||"NEW,CONTACTED,NURTURE,BOOKED,CLOSED")
    .split(",").map(s=>String(s||"").trim().toUpperCase()).filter(Boolean);
  const colors = {};
  String(cfg.stage_colors||"").split(";").forEach(pair => {
    const pts = String(pair||"").trim().split("=");
    const k = String(pts[0]||"").trim().toUpperCase();
    const v = String(pts[1]||"").trim();
    if (k && v) colors[k] = v;
  });
  return { stages, colors, textTemplate: String(cfg.text_body_template||"") };
}

// ═══════════════════════════════════════════════════════════════
//  STAGE PACK CONFIG
// ═══════════════════════════════════════════════════════════════

function getStagePackConfig_(pack) {
  const p = String(pack||"local_service").trim().toLowerCase();
  const packs = {
    local_service: { stages:["NEW","ATTEMPTED","SCHEDULED","NO_SHOW","CLOSED"],    colors:{NEW:"#7aa2ff",ATTEMPTED:"#ffcc00",SCHEDULED:"#33d17a",NO_SHOW:"#ff5c5c",CLOSED:"#8b949e"}, defaultPack:"local_services" },
    agency:        { stages:["NEW","CONTACTED","NURTURE","PROPOSAL","CLOSED"],      colors:{NEW:"#7aa2ff",CONTACTED:"#33d17a",NURTURE:"#ffcc00",PROPOSAL:"#b197fc",CLOSED:"#8b949e"}, defaultPack:"agency" },
    real_estate:   { stages:["NEW","CONTACTED","SHOWINGS","OFFER","CLOSED"],        colors:{NEW:"#7aa2ff",CONTACTED:"#33d17a",SHOWINGS:"#ffcc00",OFFER:"#b197fc",CLOSED:"#8b949e"},   defaultPack:"real_estate" },
    recruiting:    { stages:["NEW","SCREEN","INTERVIEW","OFFER","CLOSED"],          colors:{NEW:"#7aa2ff",SCREEN:"#33d17a",INTERVIEW:"#ffcc00",OFFER:"#b197fc",CLOSED:"#8b949e"},     defaultPack:"recruiting" },
    universal:     { stages:["NEW","CONTACTED","NURTURE","BOOKED","CLOSED"],        colors:{NEW:"#7aa2ff",CONTACTED:"#33d17a",NURTURE:"#ffcc00",BOOKED:"#b197fc",CLOSED:"#8b949e"},   defaultPack:"universal" }
  };
  return packs[p] || packs.local_service;
}

function applyStagePack_(cfg, packName) {
  const pack = getStagePackConfig_(packName || cfg.stage_pack);
  cfg.stage_pack            = String(packName || cfg.stage_pack || "local_service");
  cfg.pipeline_stages       = pack.stages.join(",");
  cfg.stage_colors          = Object.keys(pack.colors).map(k=>k+"="+pack.colors[k]).join(";");
  cfg.default_template_pack = pack.defaultPack;
  return cfg;
}

// ═══════════════════════════════════════════════════════════════
//  UTILS
// ═══════════════════════════════════════════════════════════════

function safeUiAlert_(msg) {
  try { SpreadsheetApp.getUi().alert(String(msg||"")); } catch(e) { /* trigger context */ }
}

function formatRunSummary_(result) {
  if (!result) return "Done.";
  return ["✅ Done.","Mode:"+(result.mode||""),"Due:"+(result.due_count||0),
    "Drafted:"+(result.drafted||0),"Sent:"+(result.sent||0),"Errors:"+(result.errors||0)].join("\n");
}

function getSheetNames_() {
  return SpreadsheetApp.getActive().getSheets().map(s=>s.getName());
}

function toBool_(v) {
  const s = String(v||"").trim().toUpperCase();
  return s==="TRUE"||s==="YES"||s==="1";
}

function asDate_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v)==="[object Date]" && !isNaN(v.getTime())) return v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function startOfDay_(d) {
  const x = new Date(d); x.setHours(0,0,0,0); return x;
}

function nextFollowUpDate_(startDay, days, businessOnly) {
  let d = new Date(startDay), remaining = days;
  while (remaining > 0) {
    d.setDate(d.getDate()+1);
    if (businessOnly && (d.getDay()===0||d.getDay()===6)) continue;
    remaining--;
  }
  return d;
}

function getFirstName_(nameVal) {
  const raw = String(nameVal||"").trim();
  if (!raw) return "";
  if (raw.indexOf(",")!==-1) {
    const after = String(raw.split(",")[1]||"").trim();
    if (after) return after.split(/\s+/)[0];
  }
  return raw.split(/\s+/)[0];
}

function indexMap_(headers) {
  const m = {};
  headers.forEach((h,i) => { if (h) m[h]=i; });
  return m;
}

function parseDraft_(draftText) {
  const raw = String(draftText||"").trim();
  if (!raw) return null;
  const lines = raw.split(/\r?\n/);
  if (!lines.length) return null;
  const first = String(lines[0]||"").trim();
  if (/^subject\s*:/i.test(first)) {
    const subject = first.replace(/^subject\s*:/i,"").trim();
    let bodyLines = lines.slice(1);
    if (bodyLines.length && String(bodyLines[0]).trim()==="") bodyLines = bodyLines.slice(1);
    return { subject: subject||"(no subject)", body: bodyLines.join("\n").trim() };
  }
  return { subject:"(no subject)", body:raw };
}

// ═══════════════════════════════════════════════════════════════
//  OPS SYSTEM — Sheet openers
// ═══════════════════════════════════════════════════════════════

function openSystemOps_()          { _activateSheet(LFU.SHEET_SYSOPS,       "System Ops ready. KPIs are live."); }
function openSopLibrary_()         { _activateSheet(LFU.SHEET_SOPS,         "SOP Library. Add or edit SOPs here."); }
function openSopBuilder_()         { _activateSheet(LFU.SHEET_SOP_BUILDER,  "Fill Column B, then Autopilot → Generate SOP from Builder."); }
function openSopTemplates_()       { _activateSheet(LFU.SHEET_SOP_TMPL,     "Copy a template into SOP Builder to start fast."); }
function openProcessImprovements_(){ _activateSheet(LFU.SHEET_IMPROVE,      "Track pain points and planned improvements here."); }
function openImportWizard_()       { _activateSheet(LFU.SHEET_IMPORT,       "Paste CSV (with headers) into row 10. Map row 9. Then Autopilot → Run Import."); }
function openIntake_()             { _activateSheet(LFU.SHEET_INTAKE,       "Connect your form tool to write here. Then run Process Intake."); }
function openAppointments_()       { _activateSheet(LFU.SHEET_APPOINTMENTS, "Appointments tab is the scheduler audit trail."); }

function _activateSheet(name, toastMsg) {
  ensureSheets_();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name);
  if (sh) { ss.setActiveSheet(sh); ss.toast(toastMsg, "Autopilot", 4); }
}

// ═══════════════════════════════════════════════════════════════
//  OPS SYSTEM — Sheet initializers
// ═══════════════════════════════════════════════════════════════

function ensureSysOpsSheet_(sheet) {
  if (!sheet) return;
  const a1 = String(sheet.getRange(1,1).getValue()||"");
  if (a1.toLowerCase().includes("system ops")) return;
  sheet.clear();
  sheet.getRange("A1").setValue("System Ops").setFontWeight("bold").setFontSize(14);
  sheet.getRange("A2").setValue("Operational dashboard. KPIs are live formula-driven. Check Apps Script → Executions for errors.");
  const rows = [
    ["Version",            LFU.VERSION],
    ["Last Autopilot Run", '=IFERROR(TEXT(MAX(FILTER(Activities!A:A,Activities!D:D="Autopilot Run")),"yyyy-mm-dd hh:mm"),"—")'],
    ["Last Review Send",   '=IFERROR(TEXT(MAX(FILTER(Activities!A:A,Activities!D:D="Review Send Batch")),"yyyy-mm-dd hh:mm"),"—")'],
    ["",""],
    ["— Live KPIs —",""],
    ["Due Now",           '=IFERROR(COUNTIF(INDEX(Leads!A:ZZ,,MATCH("Due Flag",Leads!1:1,0)),"TRUE"),0)'],
    ["Drafted",           '=IFERROR(COUNTIF(INDEX(Leads!A:ZZ,,MATCH("Draft Status",Leads!1:1,0)),"DRAFTED"),0)'],
    ["Approved to Send",  '=IFERROR(COUNTIF(INDEX(Leads!A:ZZ,,MATCH("Approved to Send",Leads!1:1,0)),"TRUE"),0)'],
    ["Errors",            '=IFERROR(COUNTIF(INDEX(Leads!A:ZZ,,MATCH("Draft Status",Leads!1:1,0)),"ERROR"),0)'],
    ["Total Leads",       '=IFERROR(COUNTA(INDEX(Leads!A:ZZ,,MATCH("Email",Leads!1:1,0)))-1,0)'],
    ["",""],
    ["— Revenue View —",""],
    ["Pipeline Value",       '=IFERROR(SUM(IFERROR(FILTER(INDEX(Leads!A:ZZ,,MATCH("Deal Value",Leads!1:1,0)),INDEX(Leads!A:ZZ,,MATCH("Status",Leads!1:1,0))<>"CLOSED"),0)),0)'],
    ["Expected Value (WP)",  '=IFERROR(SUM(IFERROR(FILTER(INDEX(Leads!A:ZZ,,MATCH("Expected Value",Leads!1:1,0)),INDEX(Leads!A:ZZ,,MATCH("Status",Leads!1:1,0))<>"CLOSED"),0)),0)'],
    ["Won Deals (Count)",    '=IFERROR(COUNTA(FILTER(INDEX(Leads!A:ZZ,,MATCH("Won Date",Leads!1:1,0)),INDEX(Leads!A:ZZ,,MATCH("Won Date",Leads!1:1,0))<>"")),0)'],
    ["Lost Deals (Count)",   '=IFERROR(COUNTA(FILTER(INDEX(Leads!A:ZZ,,MATCH("Lost Date",Leads!1:1,0)),INDEX(Leads!A:ZZ,,MATCH("Lost Date",Leads!1:1,0))<>"")),0)'],
    ["Avg Priority Score",   '=IFERROR(AVERAGE(FILTER(INDEX(Leads!A:ZZ,,MATCH("Priority Score",Leads!1:1,0)),INDEX(Leads!A:ZZ,,MATCH("Email",Leads!1:1,0))<>"")),0)']
  ];
  sheet.getRange(4,1,rows.length,2).setValues(rows);
  sheet.getRange("A4:A"+(4+rows.length)).setFontWeight("bold");
  sheet.autoResizeColumns(1,2);
}

function ensureSopSheet_(sheet) {
  if (!sheet) return;
  const lastCol = Math.max(sheet.getLastColumn(),1);
  const headers = sheet.getRange(1,1,1,lastCol).getValues()[0].map(h=>String(h||"").trim()).filter(Boolean);
  const H = ["SOP ID","Title","Category","When / Trigger","Steps / Body","Owner","Automation Hook","Last Updated","Active"];
  if (headers.length === 0) {
    sheet.clear();
    sheet.getRange(1,1,1,H.length).setValues([H]);
    sheet.setFrozenRows(1);
  }
  if (sheet.getLastRow() < 2) {
    const seeds = [
      ["SOP-001","Daily Lead Review","Sales Ops","Daily @ 9am",
        "1) Open Command Center\n2) Work Today list (due leads sorted by Priority Score)\n3) Call/Text each due lead\n4) Mark Called or Snooze\n5) Approve good drafts\n6) Send Approved",
        "Owner/Team","Today workflow","","TRUE"],
      ["SOP-002","No-Show Follow-Up","Sales Ops","When Status = NO_SHOW",
        "1) Rules Engine auto-snoozes 1 day + drafts follow-up\n2) Review draft in Command Center\n3) Approve + Send",
        "Owner/Team","Rules Engine: RULE-001 + RULE-003","","TRUE"],
      ["SOP-003","Weekly Pipeline Review","Management","Every Friday",
        "1) Open System Ops tab\n2) Review Pipeline Value + Expected Value\n3) Sort by Priority Score in Full Screen\n4) Fix stage hygiene\n5) Update Probability % for open deals",
        "Manager","System Ops KPIs","","TRUE"],
      ["SOP-004","New Lead Entry","Sales","When new lead arrives",
        "1) Use Quick Add in Command Center (or add row to Leads)\n2) Set Deal Value + Probability\n3) Set Follow Up Due = today\n4) Run Autopilot (REVIEW)\n5) Approve + Send",
        "Sales Rep","Quick Add → REVIEW","","TRUE"]
    ];
    sheet.getRange(2,1,seeds.length,H.length).setValues(seeds);
    sheet.autoResizeColumns(1,H.length);
  }
}

function ensureProcessImprovementsSheet_(sheet) {
  if (!sheet) return;
  const lastCol = Math.max(sheet.getLastColumn(),1);
  const headers = sheet.getRange(1,1,1,lastCol).getValues()[0].map(h=>String(h||"").trim()).filter(Boolean);
  const H = ["ID","Area","Pain Point","Proposed Fix","Impact","Effort","Priority","Status","Owner","Notes","Created","Updated"];
  if (headers.length === 0) {
    sheet.clear();
    sheet.getRange(1,1,1,H.length).setValues([H]);
    sheet.setFrozenRows(1);
  }
  if (sheet.getLastRow() < 2) {
    const now = new Date();
    const seeds = [
      ["IMP-001","Install","Users struggle with Apps Script permissions","Clearer permission guide + screenshot","High","Low","P0","Planned","Owner","Reduces refunds",now,now],
      ["IMP-002","Import","Existing lead list import is painful","Import Wizard (v8 shipped)","High","Medium","P0","Shipped","Owner","Direct activation ROI",now,now],
      ["IMP-003","Automation","No rules engine","Rules Engine Lite (v8 shipped)","High","High","P0","Shipped","Owner","Game changer",now,now],
      ["IMP-004","UI","No quick add from Command Center","Quick Add modal (v8 shipped)","Medium","Low","P1","Shipped","Owner","UX win",now,now],
      ["IMP-005","UI","No activity timeline","Activity Timeline tab (v8 shipped)","Medium","Medium","P1","Shipped","Owner","Trust builder",now,now]
    ];
    sheet.getRange(2,1,seeds.length,H.length).setValues(seeds);
    sheet.autoResizeColumns(1,H.length);
  }
}

function ensureSopBuilderSheet_(sheet) {
  if (!sheet) return;
  const a1 = String(sheet.getRange(1,1).getValue()||"");
  if (a1.toLowerCase().includes("sop builder")) return;
  sheet.clear();
  sheet.getRange("A1").setValue("SOP Builder").setFontWeight("bold").setFontSize(14);
  sheet.getRange("A2").setValue("Fill out Column B. Then run: Autopilot → Generate SOP from Builder");
  sheet.getRange("A3").setValue("Copy a starter from SOP Templates to paste into Column B rows.");
  const rows = [
    ["SOP Title",""],["Category","Sales Ops"],["Trigger / When to Use",""],
    ["Goal / Outcome",""],["Owner (person or role)",""],["Tools / Systems","Google Sheets, Phone"],
    ["Inputs Needed",""],["Steps (one per line)","1)\n2)\n3)\n4)\n5)"],
    ["Quality Checks",""],["Escalation / Edge Cases",""],["KPIs / What good looks like",""],
    ["Automation Hook (optional)","e.g. When stage=NO_SHOW → Rules Engine fires RULE-001"],
    ["Notes",""]
  ];
  sheet.getRange(5,1,rows.length,2).setValues(rows);
  sheet.setFrozenRows(4);
  sheet.setColumnWidth(1,240); sheet.setColumnWidth(2,760);
  sheet.getRange("A5:A17").setFontWeight("bold");
  sheet.getRange("B12:B17").setWrap(true);
}

function ensureSopTemplatesSheet_(sheet) {
  if (!sheet) return;
  const lastCol = Math.max(sheet.getLastColumn(),1);
  const headers = sheet.getRange(1,1,1,lastCol).getValues()[0].map(h=>String(h||"").trim()).filter(Boolean);
  const H = ["Template Name","Category","Trigger","Steps (outline)","Automation Hook","Notes"];
  if (headers.length === 0) {
    sheet.clear();
    sheet.getRange(1,1,1,H.length).setValues([H]);
    sheet.setFrozenRows(1);
  }
  if (sheet.getLastRow() < 2) {
    const seeds = [
      ["Daily Lead Review","Sales Ops","Daily @ 9am",
        "1) Open Command Center\n2) Work Today list\n3) Call/Text due leads\n4) Mark Called or Snooze\n5) Approve drafts\n6) Send Approved",
        "Today workflow","Core daily SOP."],
      ["No-Show Recovery","Sales Ops","Status = NO_SHOW",
        "1) Rules Engine snoozes + drafts automatically\n2) Review in Command Center\n3) Approve + Send",
        "Rules Engine RULE-001+003","High-value SOP."],
      ["New Lead Entry","Sales","New lead comes in",
        "1) Quick Add from Command Center\n2) Set Deal Value + Probability\n3) Run REVIEW\n4) Approve + Send",
        "Quick Add → REVIEW","Activation SOP."],
      ["Win / Closed Deal","Sales","Moving to BOOKED or CLOSED",
        "1) Drag card to BOOKED/CLOSED\n2) System stamps Won Date + clears due\n3) Send thank-you\n4) Note referral opportunity",
        "Won Date stamp","Celebrate + referral."]
    ];
    sheet.getRange(2,1,seeds.length,H.length).setValues(seeds);
    sheet.autoResizeColumns(1,H.length);
  }
}

function ensureImportSheet_(sheet) {
  if (!sheet) return;
  const a1 = String(sheet.getRange("A1").getValue()||"");
  if (a1.toLowerCase().includes("import wizard")) return;
  sheet.clear();
  sheet.getRange("A1").setValue("Import Wizard").setFontWeight("bold").setFontSize(14);
  sheet.getRange("A2").setValue("Step 1: Paste your CSV INCLUDING HEADERS starting at cell A10.");
  sheet.getRange("A3").setValue("Step 2: In row 9 (MAP TO row), select which Leads column each CSV column maps to.");
  sheet.getRange("A4").setValue("Step 3: Run Autopilot → Run Import.");
  sheet.getRange("A6").setValue("Import Mode (SKIP or UPSERT)");
  sheet.getRange("B6").setValue("UPSERT");
  sheet.getRange("B6").setNote("UPSERT = update existing leads (matched by Email). SKIP = skip if email exists.");
  sheet.getRange("A7").setValue("Set Follow Up Due = today if empty? (TRUE/FALSE)");
  sheet.getRange("B7").setValue("TRUE");
  sheet.getRange("A9").setValue("MAP TO ↓ (choose Leads column name or leave blank to skip)");
  sheet.getRange("A10").setValue("PASTE CSV HEADERS HERE →");
  sheet.getRange("A10").setNote("Paste your CSV starting here (row 10). Row 9 above is the mapping row.");
  sheet.setFrozenRows(9);
  sheet.setColumnWidth(1,280);
}

function ensureIntakeSheet_(sheet) {
  if (!sheet) return;
  const lastCol = Math.max(sheet.getLastColumn(),1);
  const existing = sheet.getRange(1,1,1,lastCol).getValues()[0].map(h=>String(h||"").trim()).filter(Boolean);
  if (existing.length === 0) {
    const headers = ["Timestamp","Full Name","Preferred Name","Email","Phone","Company","Message",
      "Source","Preferred Contact","Desired Date","Processed","Processed At","Lead Row","Error"];
    sheet.clear();
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1,headers.length);
    sheet.getRange("A1").setNote(
      "Connect Typeform/Jotform/Google Forms to write responses here.\n" +
      "Then run: Autopilot → Process Intake Now\n" +
      "Or install the 5-minute trigger: Autopilot → Install Intake Trigger"
    );
  }
}

function ensureAppointmentsSheet_(sheet) {
  if (!sheet) return;
  const lastCol = Math.max(sheet.getLastColumn(),1);
  const existing = sheet.getRange(1,1,1,lastCol).getValues()[0].map(h=>String(h||"").trim()).filter(Boolean);
  const H = ["Appointment ID","Created At","Lead Row","Lead ID","Name","Email","Phone","Company",
    "Status","Start","End","Duration (min)","Calendar Event ID","Calendar Link","Location","Notes",
    "Reminder Sent","Reminder Sent At"];
  const existingSet = new Set(existing);
  if (existing.length === 0) {
    sheet.clear();
    sheet.getRange(1,1,1,H.length).setValues([H]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1,H.length);
    return;
  }
  const missing = H.filter(h=>!existingSet.has(h));
  if (missing.length) {
    sheet.getRange(1, lastCol+1, 1, missing.length).setValues([missing]);
    sheet.autoResizeColumns(lastCol+1, missing.length);
  }
}

// ═══════════════════════════════════════════════════════════════
//  SOP GENERATOR
// ═══════════════════════════════════════════════════════════════

function generateSopFromBuilder_() {
  ensureSheets_();
  const ss      = SpreadsheetApp.getActive();
  const builder = ss.getSheetByName(LFU.SHEET_SOP_BUILDER);
  const sopSheet = ss.getSheetByName(LFU.SHEET_SOPS);
  if (!builder || !sopSheet) throw new Error("Missing SOP Builder or SOP Library sheet.");
  const data = builder.getRange(5,1,13,2).getValues();
  const get = label => {
    const row = data.find(r => String(r[0]||"").trim()===label);
    return row ? String(row[1]||"").trim() : "";
  };
  const title = get("SOP Title");
  if (!title) { safeUiAlert_("SOP Title is required. Fill in Column B of SOP Builder first."); return; }
  const category = get("Category") || "General";
  const trigger  = get("Trigger / When to Use");
  const goal     = get("Goal / Outcome");
  const owner    = get("Owner (person or role)");
  const tools    = get("Tools / Systems");
  const inputs   = get("Inputs Needed");
  const steps    = get("Steps (one per line)");
  const qc       = get("Quality Checks");
  const edges    = get("Escalation / Edge Cases");
  const kpis     = get("KPIs / What good looks like");
  const hook     = get("Automation Hook (optional)");
  const notes    = get("Notes");
  const body = "GOAL:\n"+(goal||"—")+"\n\nTRIGGER:\n"+(trigger||"—")+
    "\n\nOWNER:\n"+(owner||"—")+"\n\nTOOLS:\n"+(tools||"—")+
    "\n\nINPUTS:\n"+(inputs||"—")+"\n\nSTEPS:\n"+(steps||"—")+
    "\n\nQUALITY CHECKS:\n"+(qc||"—")+"\n\nEDGE CASES:\n"+(edges||"—")+
    "\n\nKPIs:\n"+(kpis||"—")+"\n\nNOTES:\n"+(notes||"—");
  const id = "SOP-"+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyyMMdd-HHmm");
  sopSheet.appendRow([id,title,category,trigger,body,owner,hook,new Date(),"TRUE"]);
  logActivity_({ leadId:"", name:"", action:"SOP Generated", notes:id+": "+title });
  ss.setActiveSheet(sopSheet);
  ss.toast("✅ SOP generated: "+id+" — "+title, "SOP Builder", 6);
}

// ═══════════════════════════════════════════════════════════════
//  SOP PACK — IMPACT 15 (60 SOPs across 15 categories)
// ═══════════════════════════════════════════════════════════════

function loadSopPackImpact15_() {
  ensureSheets_();
  const ss = SpreadsheetApp.getActive();
  const sopSheet = ss.getSheetByName(LFU.SHEET_SOPS);
  if (!sopSheet) throw new Error("Missing SOP Library sheet.");
  const existing = new Set();
  const lastRow = sopSheet.getLastRow();
  if (lastRow >= 2) {
    sopSheet.getRange(2,1,lastRow-1,1).getValues().flat()
      .forEach(v=>existing.add(String(v||"").trim()));
  }
  const now = new Date();
  const S = (id,title,cat,trigger,steps,owner,hook) =>
    [id,title,cat,trigger,steps,owner,hook,now,"TRUE"];
  const pack = [
    S("SOP-S1-001","5-Min Rule (Inbound)","Speed-to-Lead","New lead created","1) Call within 5 min\n2) If no answer: voicemail+email+text\n3) Set Follow Up Due=today\n4) Log outcome","Owner/Team","Priority Score bump"),
    S("SOP-S1-002","Inbound Routing","Speed-to-Lead","New lead arrives","1) Assign Owner\n2) Set Deal Value+Probability\n3) Set Template Pack\n4) Follow Up Due=today","Manager","Assignment rule"),
    S("SOP-S1-003","First Call Script","Speed-to-Lead","First call attempt","1) Identify+reason\n2) Ask 2 qualifier questions\n3) Confirm best contact\n4) Close with next step+time","Sales","Call script"),
    S("SOP-S1-004","No Response Ladder","Speed-to-Lead","No reply after first touch","Day 0: Call+email. Day 1: Call+text. Day 3: Email value+call. Day 7: Close-the-loop. Day 14: Nurture or close.","Sales","Cadence rule"),
    S("SOP-S2-001","Daily Power Hour","Daily Execution","Daily 9am","1) Open Command Center → Today\n2) Work due leads high-to-low by Priority Score\n3) Call/text/email\n4) Mark Called or Snooze\n5) Approve drafts\n6) Send Approved","Owner/Team","Today list + Priority Score"),
    S("SOP-S2-002","Next Action Standard","Daily Execution","Any lead touch","Every lead gets a dated Next Action. Follow Up Due matches. If unclear: nurture cadence.","Owner/Team","Rules-lite"),
    S("SOP-S2-003","Overdue Value Triage","Daily Execution","Overdue exists","1) Sort due leads by Deal Value\n2) Work top 10 first\n3) Escalate stuck high-value\n4) Convert or close","Owner","Overdue value KPI"),
    S("SOP-S2-004","End-of-Day Hygiene","Daily Execution","Daily close","1) Clear approvals\n2) Fix wrong stages\n3) Ensure tomorrow has a call list\n4) Log blockers in Process Improvements","Owner/Team","System Ops"),
    S("SOP-S3-001","Discovery Scorecard","Qualification","Discovery call complete","Score Need/Timeline/Budget/Authority 0-3 each.\n<6: nurture. 6-9: proposal. 10+: schedule next step within 48h.","Sales","Probability update"),
    S("SOP-S3-002","Disqualify Cleanly","Qualification","Lead unqualified","1) Move to LOST\n2) Log reason\n3) If timing: NURTURE+future due\n4) Send polite closeout","Sales","Stage move + Rules Engine"),
    S("SOP-S3-003","Required Field Hygiene","Qualification","Weekly","Audit missing email/phone/value/probability. Fix top gaps.","Admin","System Ops check"),
    S("SOP-S3-004","Qualification Call Script","Qualification","Discovery call","1) Problem\n2) Desired outcome\n3) Constraints\n4) Decision process\n5) Confirm next step time","Sales","Script"),
    S("SOP-S4-001","24-Hour Proposal SLA","Proposal/Close","After qualified discovery","1) Send proposal within 24h\n2) Include options\n3) Expiration\n4) Set due +2 days","Sales","Draft template"),
    S("SOP-S4-002","Proposal Follow-up Cadence","Proposal/Close","Proposal sent","Day 1: questions email. Day 3: call. Day 7: hold-your-spot email. Day 14: nurture or lost.","Sales","Cadence"),
    S("SOP-S4-003","Objection Handling","Proposal/Close","Objection raised","1) Restate\n2) Clarify\n3) Match script\n4) Offer two paths\n5) Book next step or close","Sales","Script"),
    S("SOP-S4-004","Close + Deposit","Proposal/Close","Customer says yes","1) Collect deposit\n2) Schedule via Scheduler\n3) Move to BOOKED\n4) Set follow-up for review/referral","Sales/Ops","Scheduler Lite + Rules Engine"),
    S("SOP-S5-001","Booking Confirmation","Scheduling","Appointment scheduled","1) Confirm time/location\n2) Prep instructions\n3) Cancellation policy\n4) System sends 24h reminder","Ops","Scheduler + Reminders"),
    S("SOP-S5-002","24h Reminder + Confirm","Scheduling","24h before appointment","System sends automatic reminder via Appointment Reminders trigger. Install via Autopilot menu.","Ops","Reminder trigger"),
    S("SOP-S5-003","No-Show Recovery","Scheduling","Status = NO_SHOW","Rules Engine RULE-001: snooze 1d. RULE-002: set no_show template. RULE-003: draft follow-up. Review + Approve + Send.","Sales","Rules Engine"),
    S("SOP-S5-004","Reschedule Policy","Scheduling","Customer requests reschedule","1) Offer 2 options\n2) Update appointment via Scheduler\n3) Confirm policy\n4) Log reason","Ops","Appointment update"),
    S("SOP-S6-001","Definition of Done","Delivery Quality","Before marking complete","1) Checklist complete\n2) Photo proof\n3) Customer walkthrough\n4) Payment collected\n5) Notes logged","Ops","QC checklist"),
    S("SOP-S6-002","Warranty/Callback Workflow","Delivery Quality","Issue reported","1) Acknowledge in 2h\n2) Triage\n3) Schedule fix\n4) Root cause tag\n5) Update SOP if recurring","Ops","Issue log"),
    S("SOP-S6-003","Before/After Documentation","Delivery Quality","On completion","Capture required photos/notes. Store in Notes. Use for reviews/referrals.","Ops","Evidence required"),
    S("SOP-S6-004","Service Standard Work","Delivery Quality","Every job","Run step checklist. Confirm safety. Confirm clean finish. Confirm customer sign-off.","Ops","Checklist"),
    S("SOP-S7-001","Complaint Intake","Customer Success","Complaint received","1) Acknowledge same day\n2) Capture facts\n3) Offer options\n4) Escalate by severity\n5) Close loop in writing","Manager","Escalation ladder"),
    S("SOP-S7-002","Refund / Chargeback","Customer Success","Refund requested","1) Check eligibility\n2) Approve/deny with script\n3) Process\n4) Log reason\n5) Prevent recurrence","Manager","Policy"),
    S("SOP-S7-003","Service Recovery","Customer Success","Service failure confirmed","1) Apologize\n2) Fix plan\n3) Recovery offer\n4) Confirm satisfaction\n5) Later: review ask","Manager","Recovery"),
    S("SOP-S7-004","Communication Scripts","Customer Success","Ongoing","Maintain approved scripts for: delays, complaints, scheduling, pricing. Train team. Update quarterly.","Manager","Scripts"),
    S("SOP-S8-001","Review Request","Reviews/Referrals","After success","Day 0: ask. Day 3: reminder. If negative: escalate immediately.","Ops","Template"),
    S("SOP-S8-002","Referral Ask Script","Reviews/Referrals","After positive outcome","Ask for 1 referral by name. Provide copy/paste intro. Follow up referred lead in 5 minutes.","Owner/Team","Script"),
    S("SOP-S8-003","Dormant Winback","Reviews/Referrals","Dormant 60-90 days","Send winback. Text handoff high value. Offer quick check-in. Move back to NEW when engaged.","Owner/Team","Winback"),
    S("SOP-S8-004","Reputation Monitoring","Reviews/Referrals","Weekly","Check new reviews. Respond within 48h. Escalate negative. Log trends.","Manager","Ops cadence"),
    S("SOP-S9-001","Weekly Cash Review","Cash Flow","Every Monday","1) Cash on hand\n2) AR due\n3) AP+payroll+taxes\n4) 14-day runway\n5) Decide actions","Owner","System Ops"),
    S("SOP-S9-002","14-Day Runway Rules","Cash Flow","After cash review","<14: freeze discretionary + collections daily. 14-30: tighten approvals. 30+: invest in pipeline.","Owner","Decision rules"),
    S("SOP-S9-003","Monthly Budget vs Actual","Cash Flow","Monthly","Review revenue vs target. Review expense variances. Decide corrections. Log decisions.","Owner","Cadence"),
    S("SOP-S9-004","Tax/Compliance Calendar","Cash Flow","Monthly/Quarterly","Maintain deadlines. Confirm filings. Store receipts. Audit quarterly.","Owner/Admin","Calendar"),
    S("SOP-S10-001","Same-Day Invoicing","Collections","After delivery","Invoice same day. Terms visible. Reminder schedule. Log payment status.","Admin","Policy"),
    S("SOP-S10-002","Collections Ladder","Collections","Invoice overdue","Day 0 reminder. Day 7 call. Day 14 final notice. Day 30 stop-work/collections. Log every touch.","Admin","Cadence"),
    S("SOP-S10-003","Payment Plan Policy","Collections","Customer requests plan","Min down. Schedule payments. Auto reminders. Default handling. Log agreement.","Owner/Admin","Policy"),
    S("SOP-S10-004","Deposit Enforcement","Collections","Booking confirmed","Deposit required. Cancellation window. No-show fee. Exceptions logged.","Owner","Policy"),
    S("SOP-S11-001","Quarterly Pricing Review","Pricing/Margins","Quarterly","Review margins+win rate. Compare competitors. Adjust pricing. Update scripts/templates.","Owner","Cadence"),
    S("SOP-S11-002","Discount Approval Rules","Pricing/Margins","Discount requested","Approval above threshold. Trade for something (prepay/upsell). Log reason.","Owner","Policy"),
    S("SOP-S11-003","Change Order Guardrail","Pricing/Margins","Scope expands","Stop+document. Price change. Written approval. Update invoice. Log.","Owner/Ops","Policy"),
    S("SOP-S11-004","Vendor Cost Review","Pricing/Margins","Quarterly","Review top costs. Renegotiate. Swap vendors. Track savings.","Owner","Procurement"),
    S("SOP-S12-001","Role Scorecard + Job Post","Hiring","Before posting","Define outcomes. Define must-have skills. Rubric. Comp range. Post.","Owner","Hiring pipeline"),
    S("SOP-S12-002","Interview Loop","Hiring","Candidate evaluation","Screen. Skills test. Values interview. References. Decision meeting.","Owner/Manager","Rubrics"),
    S("SOP-S12-003","Reference Check Script","Hiring","Before offer","Verify dates. Strengths. Weaknesses. Rehire question. Document.","Manager","Script"),
    S("SOP-S12-004","Offer + Start Date","Hiring","Candidate selected","Send offer. Confirm start date. Prepare onboarding. Assign mentor. Schedule training.","Manager","Onboarding trigger"),
    S("SOP-S13-001","Day 1 Onboarding","Onboarding/Training","New hire start","Tools access. SOP tour. Shadow plan. Week 1 goals. Schedule 1:1.","Manager","SOP system"),
    S("SOP-S13-002","30/60/90 Ramp Plan","Onboarding/Training","Week 1","Define outcomes. Weekly check-ins. Skills checklist. KPIs. Sign-off checkpoints.","Manager","Cadence"),
    S("SOP-S13-003","Training Certification","Onboarding/Training","Role training required","Modules. Checklist. Assessment. Sign-off. Refresh cadence.","Manager","Certification"),
    S("SOP-S13-004","Offboarding SOP","Onboarding/Training","Employee exits","Access removal. Equipment return. Knowledge capture. Customer handoffs. Exit notes.","Manager","Security"),
    S("SOP-S14-001","Weekly 1:1 Template","Management","Weekly","Wins. Blockers. Metrics. Priorities. Feedback. Commitments.","Manager","Cadence"),
    S("SOP-S14-002","KPI Scoreboard Review","Management","Weekly","Pipeline KPIs. Cash KPIs. Ops KPIs. Top 3 priorities. Owners.","Owner","Cadence"),
    S("SOP-S14-003","Meeting Hygiene","Management","All recurring meetings","Agenda required. Owner required. Notes+actions. Cancel if no agenda. Track actions.","Owner","Ops discipline"),
    S("SOP-S14-004","Performance Issue Process","Management","Issue identified","Document. Set expectations. Improvement plan. Check-ins. Decision timeline.","Owner","People ops"),
    S("SOP-S15-001","Process Improvement Capture","System Ops","Any mistake/rework","Log issue. Root cause. Fix SOP. Retrain. Confirm fix. Log in Process Improvements tab.","Owner/Team","Process Improvements sheet"),
    S("SOP-S15-002","SOP Review Cadence","System Ops","Monthly/Quarterly","Review SOPs on schedule. Update steps. Announce changes. Confirm adoption. Track completion.","Owner","SOP governance"),
    S("SOP-S15-003","Incident / Near-Miss Log","System Ops","Incident occurs","Stop work. Document. Escalate. Corrective action. Prevent recurrence. Log.","Owner","Risk log"),
    S("SOP-S15-004","Quarterly Business Review","System Ops","Quarterly","Review KPIs. What worked/failed. Set goals. Update SOPs. Assign owners.","Owner","Cadence")
  ];
  const toAdd = pack.filter(r=>!existing.has(r[0]));
  if (!toAdd.length) { ss.toast("SOP Pack already loaded.","SOP Library",3); return; }
  sopSheet.getRange(sopSheet.getLastRow()+1,1,toAdd.length,9).setValues(toAdd);
  ss.toast("Loaded Impact 15 SOP Pack: "+toAdd.length+" SOPs","SOP Library",4);
}

// ═══════════════════════════════════════════════════════════════
//  IMPORT WIZARD
// ═══════════════════════════════════════════════════════════════

function runImportWizard_() {
  ensureSheets_();
  const ss = SpreadsheetApp.getActive();
  const importSheet = ss.getSheetByName(LFU.SHEET_IMPORT);
  const leadsSheet  = ss.getSheetByName(LFU.SHEET_LEADS);
  if (!importSheet || !leadsSheet) throw new Error("Missing Import or Leads sheet.");
  ensureLeadColumns_(leadsSheet);

  const { headerIndex, rows: existingRows } = readLeads_(leadsSheet);
  const maxCols = Math.max(importSheet.getLastColumn(),1);
  const csvHeaders = importSheet.getRange(10,1,1,maxCols).getValues()[0].map(h=>String(h||"").trim());
  let lastCol = 0;
  for (let i = csvHeaders.length-1; i >= 0; i--) { if (csvHeaders[i]) { lastCol = i+1; break; } }
  if (lastCol === 0) throw new Error("No CSV headers found in row 10. Paste CSV starting at A10.");

  const mapTo = importSheet.getRange(9,1,1,lastCol).getValues()[0].map(v=>String(v||"").trim());
  const dataStartRow = 11;
  const lastRow = importSheet.getLastRow();
  if (lastRow < dataStartRow) throw new Error("No CSV data rows found under the headers.");

  const values = importSheet.getRange(dataStartRow,1,lastRow-dataStartRow+1,lastCol).getValues();
  const mode = String(importSheet.getRange("B6").getValue()||"UPSERT").trim().toUpperCase();
  const fillDue = String(importSheet.getRange("B7").getValue()||"TRUE").trim().toUpperCase()==="TRUE";
  const today = Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyy-MM-dd");

  const existingByEmail = new Map();
  existingRows.forEach(r => {
    const e = String(r["Email"]||"").trim().toLowerCase();
    if (e) existingByEmail.set(e, r.__row);
  });

  let created = 0, updated = 0, skipped = 0;

  values.forEach(row => {
    const hasAny = row.some(cell=>String(cell||"").trim()!=="");
    if (!hasAny) return;
    const upd = {};
    for (let c = 0; c < lastCol; c++) {
      const target = mapTo[c];
      if (!target) continue;
      upd[target] = row[c];
    }
    if (!upd["Status"]) upd["Status"] = "NEW";
    if (fillDue && (!upd["Follow Up Due"]||String(upd["Follow Up Due"]).trim()==="")) upd["Follow Up Due"] = today;

    const email = String(upd["Email"]||"").trim().toLowerCase();
    if (!email) { appendLeadRow_(leadsSheet, headerIndex, upd); created++; return; }

    const existingRowNum = existingByEmail.get(email);
    if (existingRowNum && mode==="SKIP") { skipped++; return; }
    if (existingRowNum && mode==="UPSERT") {
      Object.keys(upd).forEach(k => setCell_(leadsSheet, existingRowNum, headerIndex, k, upd[k]));
      updated++;
      return;
    }
    const newRow = appendLeadRow_(leadsSheet, headerIndex, upd);
    existingByEmail.set(email, newRow);
    created++;
  });

  logActivity_({ leadId:"", name:"", action:"Import", notes:"created="+created+" updated="+updated+" skipped="+skipped });
  SpreadsheetApp.getUi().alert("Import complete",
    "Created: "+created+"\nUpdated: "+updated+"\nSkipped: "+skipped+
    "\n\nTip: Filter Due Flag = TRUE in Leads for your call list.",
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// ═══════════════════════════════════════════════════════════════
//  INTAKE PROCESSOR  (form submissions → leads)
// ═══════════════════════════════════════════════════════════════

function installIntakeTrigger_() {
  removeIntakeTrigger_();
  ScriptApp.newTrigger("processIntakeNow_").timeBased().everyMinutes(5).create();
  SpreadsheetApp.getActive().toast("Intake trigger installed (every 5 minutes).","Intake",4);
}

function removeIntakeTrigger_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction()==="processIntakeNow_") ScriptApp.deleteTrigger(t);
  });
  SpreadsheetApp.getActive().toast("Intake trigger removed.","Intake",3);
}

function processIntakeNow_() {
  ensureSheets_();
  const ss = SpreadsheetApp.getActive();
  const intake = ss.getSheetByName(LFU.SHEET_INTAKE);
  const leads  = ss.getSheetByName(LFU.SHEET_LEADS);
  if (!intake || !leads) throw new Error("Missing Intake or Leads sheet.");
  ensureLeadColumns_(leads);
  ensureIntakeSheet_(intake);

  const intakeRange = intake.getDataRange().getValues();
  if (intakeRange.length < 2) return;
  const h = intakeRange[0].map(x=>String(x||"").trim());
  const idx = indexMap_(h);

  const { headerIndex, rows: leadRows } = readLeads_(leads);
  const leadByEmail = new Map();
  leadRows.forEach(r => {
    const e = String(r["Email"]||"").trim().toLowerCase();
    if (e) leadByEmail.set(e, r.__row);
  });

  let created = 0, updated = 0, errored = 0;
  const today = Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyy-MM-dd");

  for (let r = 1; r < intakeRange.length; r++) {
    const row = intakeRange[r];
    if (idx["Processed"]===undefined) break;
    if (toBool_(row[idx["Processed"]])) continue;
    const email   = String(row[idx["Email"]]     || "").trim();
    const name    = String(row[idx["Full Name"]] || "").trim();
    const phone   = String(row[idx["Phone"]]     || "").trim();
    const company = String(row[idx["Company"]]   || "").trim();
    const msg     = String(row[idx["Message"]]   || "").trim();
    const source  = String(row[idx["Source"]]    || "").trim();
    if (!email && !phone && !name) {
      if (idx["Processed"]!==undefined) intake.getRange(r+1,idx["Processed"]+1).setValue("TRUE");
      if (idx["Error"]!==undefined) intake.getRange(r+1,idx["Error"]+1).setValue("Missing contact info.");
      errored++;
      continue;
    }
    try {
      const upd = {
        "Name":         name,
        "Email":        email,
        "Phone":        phone,
        "Company":      company,
        "Status":       "NEW",
        "Follow Up Due": today,
        "Intake Source": source,
        "Intake Raw":    msg
      };
      const existingRow = email ? leadByEmail.get(email.toLowerCase()) : null;
      let leadRowNum;
      if (existingRow) {
        Object.keys(upd).forEach(k => { if (upd[k]) setCell_(leads,existingRow,headerIndex,k,upd[k]); });
        leadRowNum = existingRow;
        updated++;
      } else {
        leadRowNum = appendLeadRow_(leads, headerIndex, upd);
        if (email) leadByEmail.set(email.toLowerCase(), leadRowNum);
        created++;
      }
      if (idx["Processed"]!==undefined)   intake.getRange(r+1,idx["Processed"]+1).setValue("TRUE");
      if (idx["Processed At"]!==undefined) intake.getRange(r+1,idx["Processed At"]+1).setValue(new Date());
      if (idx["Lead Row"]!==undefined)     intake.getRange(r+1,idx["Lead Row"]+1).setValue(leadRowNum);
      if (idx["Error"]!==undefined)        intake.getRange(r+1,idx["Error"]+1).setValue("");
      logActivity_({ leadId:"", name, action:"Intake Processed",
        notes:"LeadRow="+leadRowNum+" Source="+source });
    } catch (e) {
      if (idx["Error"]!==undefined)
        intake.getRange(r+1,idx["Error"]+1).setValue(e && e.message ? e.message : String(e));
      errored++;
    }
  }
  SpreadsheetApp.getActive().toast(
    "Intake processed. Created="+created+" Updated="+updated+" Errors="+errored,"Intake",5);
}

// ═══════════════════════════════════════════════════════════════
//  SCHEDULER LITE  (leads → Google Calendar → Appointments)
// ═══════════════════════════════════════════════════════════════

function scheduleSelectedLead_() {
  ensureSheets_();
  const ss    = SpreadsheetApp.getActive();
  const leads = ss.getSheetByName(LFU.SHEET_LEADS);
  if (!leads) throw new Error("Missing Leads sheet.");
  const row = leads.getActiveRange() ? leads.getActiveRange().getRow() : 0;
  if (row < 2) { SpreadsheetApp.getUi().alert("Select a lead row first (not the header)."); return; }
  const ui = SpreadsheetApp.getUi();
  const dateResp = ui.prompt("Schedule","Date (YYYY-MM-DD):",ui.ButtonSet.OK_CANCEL);
  if (dateResp.getSelectedButton()!==ui.Button.OK) return;
  const timeResp = ui.prompt("Schedule","Time (HH:MM, 24-hour):",ui.ButtonSet.OK_CANCEL);
  if (timeResp.getSelectedButton()!==ui.Button.OK) return;
  const durResp  = ui.prompt("Schedule","Duration minutes (default 30):",ui.ButtonSet.OK_CANCEL);
  if (durResp.getSelectedButton()!==ui.Button.OK) return;
  const locResp  = ui.prompt("Schedule","Location (optional):",ui.ButtonSet.OK_CANCEL);
  if (locResp.getSelectedButton()!==ui.Button.OK) return;
  uiScheduleLead_(row,
    dateResp.getResponseText().trim(),
    timeResp.getResponseText().trim(),
    parseInt(durResp.getResponseText()||"30",10)||30,
    locResp.getResponseText().trim()
  );
}

function uiScheduleLead_(leadRow, dateStr, timeStr, durationMin, location) {
  ensureSheets_();
  const ss    = SpreadsheetApp.getActive();
  const leads = ss.getSheetByName(LFU.SHEET_LEADS);
  const appts = ss.getSheetByName(LFU.SHEET_APPOINTMENTS);
  if (!leads || !appts) throw new Error("Missing Leads or Appointments sheet.");
  ensureLeadColumns_(leads);
  ensureAppointmentsSheet_(appts);

  const { headerIndex, rows } = readLeads_(leads);
  const lead = rows.find(r=>r.__row===leadRow);
  if (!lead) throw new Error("Lead row not found.");

  const start = parseLocalDateTime_(dateStr, timeStr);
  if (!start) throw new Error("Invalid date/time. Use YYYY-MM-DD and HH:MM (24-hour).");
  const dur = Math.max(5, durationMin||30);
  const end = new Date(start.getTime() + dur*60000);

  const name    = String(lead["Name"]    ||"").trim() || "Lead";
  const company = String(lead["Company"] ||"").trim();
  const email   = String(lead["Email"]   ||"").trim();
  const phone   = String(lead["Phone"]   ||"").trim();
  const title   = company ? name+" - "+company : name;

  const cal = CalendarApp.getDefaultCalendar();
  const ev  = cal.createEvent(title, start, end, {
    location: location||"",
    description:"Command Center Appointment\n\nName: "+name+"\nCompany: "+company+
      "\nEmail: "+email+"\nPhone: "+phone+"\n\nNotes: "+String(lead["Notes"]||"").trim()
  });
  if (email) { try { ev.addGuest(email); } catch(e) { /* ignore */ } }

  const eventId   = ev.getId();
  const eventLink = ev.getHtmlLink ? ev.getHtmlLink() : "";

  setCell_(leads, leadRow, headerIndex, "Appointment Start",    start);
  setCell_(leads, leadRow, headerIndex, "Appointment End",      end);
  setCell_(leads, leadRow, headerIndex, "Calendar Event ID",    eventId);
  setCell_(leads, leadRow, headerIndex, "Calendar Event Link",  eventLink);

  // Move to SCHEDULED if it's a valid stage
  const pipeline = uiGetPipelineConfig();
  if (pipeline.stages.indexOf("SCHEDULED")!==-1) {
    setCell_(leads, leadRow, headerIndex, "Status", "SCHEDULED");
    try { runRulesEngine_("STAGE_BECOMES","SCHEDULED",leadRow); } catch(e){}
  }

  // Follow-up due: next business day after appointment
  const cfg = getAllSettings_();
  const businessOnly = toBool_(cfg.business_days_only);
  const followUpDays = parseInt(String(cfg.appointment_followup_days_after||"1"),10)||1;
  const followUp = nextFollowUpDate_(startOfDay_(end), followUpDays, businessOnly);
  setCell_(leads, leadRow, headerIndex, "Follow Up Due", followUp);
  setCell_(leads, leadRow, headerIndex, "Due Flag",      "FALSE");

  const apptId = "APT-"+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyyMMdd-HHmmss");
  // Build appointment row matching ensureAppointmentsSheet_ columns
  const apptHeaders = ["Appointment ID","Created At","Lead Row","Lead ID","Name","Email","Phone","Company",
    "Status","Start","End","Duration (min)","Calendar Event ID","Calendar Link","Location","Notes",
    "Reminder Sent","Reminder Sent At"];
  const apptRow = [apptId,new Date(),leadRow,String(lead["Lead ID"]||""),name,email,phone,company,
    String(lead["Status"]||""),start,end,dur,eventId,eventLink,location||"","","",""];
  appts.appendRow(apptRow.slice(0,apptHeaders.length));

  logActivity_({ leadId:String(lead["Lead ID"]||""), name,
    action:"Appointment Scheduled", notes:dateStr+" "+timeStr+" dur="+dur });
  ss.toast("✅ Scheduled and added to calendar.","Scheduler",4);
  return { ok:true, apptId, eventId, eventLink };
}

function parseLocalDateTime_(dateStr, timeStr) {
  const m1 = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dateStr||"");
  const m2 = /^(\d{1,2}):(\d{2})$/.exec(timeStr||"");
  if (!m1||!m2) return null;
  const y=parseInt(m1[1],10), mo=parseInt(m1[2],10)-1, d=parseInt(m1[3],10);
  const h=parseInt(m2[1],10), mi=parseInt(m2[2],10);
  if (h<0||h>23||mi<0||mi>59) return null;
  return new Date(y,mo,d,h,mi,0,0);
}

// ═══════════════════════════════════════════════════════════════
//  APPOINTMENT REMINDERS
// ═══════════════════════════════════════════════════════════════

function installAppointmentReminderTrigger_() {
  removeAppointmentReminderTrigger_();
  ScriptApp.newTrigger("runAppointmentRemindersNow_").timeBased().everyMinutes(15).create();
  SpreadsheetApp.getActive().toast("Reminder trigger installed (every 15 min).","Scheduler",4);
}

function removeAppointmentReminderTrigger_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction()==="runAppointmentRemindersNow_") ScriptApp.deleteTrigger(t);
  });
  SpreadsheetApp.getActive().toast("Reminder trigger removed.","Scheduler",3);
}

function runAppointmentRemindersNow_() {
  const res = runAppointmentReminders_();
  SpreadsheetApp.getActive().toast(
    "Reminders: sent="+res.sent+" skipped="+res.skipped+" errors="+res.errors,"Scheduler",5);
}

function runAppointmentReminders_() {
  ensureSheets_();
  const cfg = getAllSettings_();
  if (String(cfg.appointment_reminder_enabled||"TRUE").toUpperCase()!=="TRUE")
    return { sent:0, skipped:0, errors:0 };

  const hoursBefore = parseInt(String(cfg.appointment_reminder_hours_before||"24"),10)||24;
  const now = new Date();
  const windowEnd = now.getTime()+(hoursBefore*3600000);

  const ss    = SpreadsheetApp.getActive();
  const appts = ss.getSheetByName(LFU.SHEET_APPOINTMENTS);
  if (!appts) return { sent:0, skipped:0, errors:0 };
  ensureAppointmentsSheet_(appts);

  const data = appts.getDataRange().getValues();
  if (data.length < 2) return { sent:0, skipped:0, errors:0 };
  const h = data[0].map(x=>String(x||"").trim());
  const idx = indexMap_(h);

  let sent = 0, skipped = 0, errors = 0;
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (toBool_(row[idx["Reminder Sent"]])) { skipped++; continue; }
    const start = asDate_(row[idx["Start"]]);
    if (!start) { skipped++; continue; }
    const ts = start.getTime();
    if (ts <= now.getTime() || ts > windowEnd) { skipped++; continue; }
    const to = String(row[idx["Email"]]||"").trim();
    if (!to) { skipped++; continue; }
    try {
      const fullName = String(row[idx["Name"]]||"").trim();
      const firstName = getFirstName_(fullName);
      const dateStr = Utilities.formatDate(start,Session.getScriptTimeZone(),"yyyy-MM-dd");
      const timeStr = Utilities.formatDate(start,Session.getScriptTimeZone(),"h:mm a");
      const senderName = String(cfg.sender_name||Session.getActiveUser().getEmail()||"").trim();
      // Reuse renderTemplate_ with date/time tokens
      const fakeLead = { Company:"", __date:dateStr, __time:timeStr };
      const fakeCfg = { sender_name:senderName };
      const subject = renderTemplate_(String(cfg.appointment_reminder_subject||"Reminder: your appointment on {date}"),
        fakeLead, fakeCfg, firstName);
      const body = renderTemplate_(String(cfg.appointment_reminder_body||"Hi {first_name}, reminder: {date} {time}. — {sender_name}"),
        fakeLead, fakeCfg, firstName);
      MailApp.sendEmail({ to, subject, body });
      if (idx["Reminder Sent"]!==undefined)    appts.getRange(r+1,idx["Reminder Sent"]+1).setValue("TRUE");
      if (idx["Reminder Sent At"]!==undefined) appts.getRange(r+1,idx["Reminder Sent At"]+1).setValue(new Date());
      logActivity_({ leadId:"", name:fullName, action:"Appointment Reminder Sent",
        notes:dateStr+" "+timeStr+" to "+to });
      sent++;
    } catch(e) { errors++; }
  }
  return { sent, skipped, errors };
}

// ═══════════════════════════════════════════════════════════════
//  DAILY BRIEFING
// ═══════════════════════════════════════════════════════════════

function installDailyBriefingTrigger_() {
  removeDailyBriefingTrigger_();
  ScriptApp.newTrigger("sendDailyBriefingScheduled_").timeBased().everyDays(1).atHour(8).create();
  SpreadsheetApp.getActive().toast("Daily briefing trigger installed (8am).","Command Center",4);
}

function removeDailyBriefingTrigger_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction()==="sendDailyBriefingScheduled_") ScriptApp.deleteTrigger(t);
  });
  SpreadsheetApp.getActive().toast("Daily briefing trigger removed.","Command Center",3);
}

function sendDailyBriefingNow_() {
  const cfg = getAllSettings_();
  const to  = Session.getActiveUser().getEmail();
  sendDailyBriefing_(to, cfg);
  SpreadsheetApp.getActive().toast("Daily briefing sent to "+to,"Command Center",4);
}

function sendDailyBriefingScheduled_() {
  const cfg = getAllSettings_();
  const to  = Session.getActiveUser().getEmail();
  sendDailyBriefing_(to, cfg);
}

function sendDailyBriefing_(toEmail, cfg) {
  ensureSheets_();
  const state = uiGetState();
  const k = state.kpis || {};
  const topDue = (state.leads || [])
    .filter(l => l.dueFlag)
    .sort((a,b) => (b.priorityScore||0)-(a.priorityScore||0))
    .slice(0,10);

  const lines = [
    "Command Center Daily Briefing — v"+LFU.VERSION,
    "Date: "+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyy-MM-dd EEE"),
    "",
    "TODAY'S PRIORITIES",
    "  Due leads:        "+k.due,
    "  Overdue value:   $"+k.overdueValue,
    "  Drafted (queue): "+k.drafted+"  |  Approved: "+k.approved+"  |  Errors: "+k.errors,
    "",
    "PIPELINE SNAPSHOT",
    "  Pipeline value:  $"+k.pipelineValue,
    "  Weighted pipeline: $"+k.weightedPipeline,
    "  Total leads:      "+k.total,
    "",
    "TOP DUE LEADS (by priority score)"
  ];
  if (!topDue.length) {
    lines.push("  None due today 🎉");
  } else {
    topDue.forEach(l => {
      lines.push("  [P"+l.priorityScore+"] "+
        (l.name||"(no name)")+" | "+(l.company||"—")+" | $"+(l.dealValue||"0")+
        " | "+(l.phone||"—")+" | "+(l.email||"—"));
    });
  }
  lines.push("","SYSTEM ACTIONS","  Rules Engine: auto-ran on all stage changes.");
  lines.push("  Intake Trigger: processes form submissions every 5 min (if installed).");
  lines.push("  Reminder Trigger: sends appointment reminders every 15 min (if installed).");
  lines.push("","NEXT STEP: Open Command Center → Today, clear due leads, approve drafts, send approved.");

  MailApp.sendEmail({
    to: toEmail,
    subject: "📋 Daily Briefing: "+k.due+" due | $"+k.overdueValue+" overdue | P-score leads ready",
    body: lines.join("\n")
  });
  logActivity_({ leadId:"", name:"", action:"Daily Briefing",
    notes:"Due="+k.due+" Overdue=$"+k.overdueValue });
}

// ═══════════════════════════════════════════════════════════════
//  SETUP WIZARD HTML
// ═══════════════════════════════════════════════════════════════

function getSetupWizardHtml_() {
  return `<!DOCTYPE html><html><head><base target="_top">
<style>
  body{font-family:Arial,sans-serif;padding:12px;margin:0;background:#0b0f14;color:#e6edf3;}
  h2{margin:0 0 6px;font-size:16px;color:#7aa2ff;}
  .row{margin:8px 0;}
  label{display:block;font-weight:600;margin-bottom:3px;font-size:13px;color:#adbac7;}
  input,select,textarea{width:100%;padding:7px 8px;box-sizing:border-box;border:1px solid rgba(255,255,255,.15);border-radius:6px;font-size:13px;background:#111823;color:#e6edf3;}
  .inline{display:flex;gap:10px;}
  .inline>div{flex:1;}
  .btn{padding:9px 14px;cursor:pointer;background:rgba(122,162,255,.2);color:#7aa2ff;border:1px solid rgba(122,162,255,.4);border-radius:6px;font-size:13px;font-weight:600;}
  .ok{color:#33d17a;font-size:12px;} .err{color:#ff5c5c;white-space:pre-wrap;font-size:12px;}
  .small{font-size:11px;opacity:.65;margin-top:2px;}
  hr{border:none;border-top:1px solid rgba(255,255,255,.1);margin:10px 0;}
  .badge{background:rgba(122,162,255,.15);border:1px solid rgba(122,162,255,.3);color:#7aa2ff;padding:2px 8px;border-radius:999px;font-size:10px;font-weight:700;margin-left:6px;}
</style></head><body>
<h2>Autopilot Setup Wizard <span class="badge">v${LFU.VERSION}</span></h2>
<div class="small" id="meta" style="margin-bottom:8px;color:#adbac7"></div>

<div class="row"><label>Stage Pack</label>
  <select id="stage_pack">
    <option value="local_service">Local Service (NEW→ATTEMPTED→SCHEDULED→CLOSED)</option>
    <option value="agency">Agency (NEW→CONTACTED→NURTURE→PROPOSAL→CLOSED)</option>
    <option value="real_estate">Real Estate (NEW→CONTACTED→SHOWINGS→OFFER→CLOSED)</option>
    <option value="recruiting">Recruiting (NEW→SCREEN→INTERVIEW→OFFER→CLOSED)</option>
    <option value="universal">Universal (NEW→CONTACTED→NURTURE→BOOKED→CLOSED)</option>
  </select>
  <div class="small">Sets pipeline stages, colors, and default template pack.</div></div>

<hr>
<div class="row"><label>Send Mode</label>
  <select id="send_mode">
    <option value="REVIEW">REVIEW — draft first, you approve then send</option>
    <option value="AUTO">AUTO — sends immediately for all due leads</option>
  </select></div>

<div class="row"><label>REVIEW send source</label>
  <select id="review_send_source">
    <option value="DRAFT">DRAFT — sends what's in the Draft column (WYSIWYG)</option>
    <option value="TEMPLATE">TEMPLATE — rebuilds from Templates at send time</option>
  </select>
  <div class="small">Recommended: DRAFT. You see exactly what gets sent.</div></div>

<hr>
<div class="inline">
  <div class="row"><label>Max emails per run</label><input id="max_emails_per_run" type="number" min="0" max="500"/></div>
  <div class="row"><label>Default follow-up days</label><input id="follow_up_days" type="number" min="1" max="90"/></div>
</div>
<div class="row"><label>Business day scheduling</label>
  <select id="business_days_only">
    <option value="TRUE">ON — skip weekends</option>
    <option value="FALSE">OFF — calendar days</option>
  </select></div>

<hr>
<div class="inline">
  <div class="row"><label>Default template pack</label>
    <input id="default_template_pack" type="text" placeholder="general"/>
    <div class="small">Used when a lead has no Template Pack set.</div></div>
  <div class="row"><label>Enable daily trigger</label>
    <select id="enable_daily_trigger">
      <option value="FALSE">No</option>
      <option value="TRUE">Yes</option>
    </select></div>
</div>
<div class="inline">
  <div class="row"><label>Daily run hour (0-23)</label><input id="autopilot_hour_local" type="number" min="0" max="23"/></div>
  <div class="row"><label>Sender name</label><input id="sender_name" type="text" placeholder="Your name"/></div>
</div>

<hr>
<div class="row"><label>Rules Engine</label>
  <select id="rules_enabled">
    <option value="TRUE">Enabled — rules fire on stage changes + autopilot runs</option>
    <option value="FALSE">Disabled</option>
  </select>
  <div class="small">Recommended: Enabled. Edit rules in the Rules sheet.</div></div>

<hr>
<div class="row"><label>Fallback subject template</label>
  <input id="email_subject_template" type="text"/>
  <div class="small">Tokens: {first_name} {company} {sender_name} — both {token} and {{token}} work</div></div>
<div class="row"><label>Fallback body template</label>
  <textarea id="email_body_template" rows="5"></textarea></div>

<div class="row"><button class="btn" onclick="save()">Save and Install</button></div>
<div id="status" class="row small"></div>
<div id="error" class="row err"></div>

<script>
  function byId(id){return document.getElementById(id);}
  function setStatus(msg,ok){byId('error').textContent='';byId('status').textContent=msg||'';byId('status').className='row small'+(ok?' ok':'');}
  function setError(msg){byId('status').textContent='';byId('error').textContent=msg||'';}
  function load(){
    setStatus('Loading…');
    google.script.run.withSuccessHandler(function(data){
      byId('meta').textContent='v'+data.version+' | Timezone: '+data.timezone;
      var s=data.settings||{};
      byId('stage_pack').value=(s.stage_pack||'local_service').toLowerCase();
      byId('send_mode').value=(s.send_mode||'REVIEW').toUpperCase();
      byId('review_send_source').value=(s.review_send_source||'DRAFT').toUpperCase();
      byId('max_emails_per_run').value=s.max_emails_per_run||10;
      byId('business_days_only').value=String(s.business_days_only||'TRUE').toUpperCase()==='FALSE'?'FALSE':'TRUE';
      byId('follow_up_days').value=s.follow_up_days||3;
      byId('autopilot_hour_local').value=s.autopilot_hour_local||9;
      byId('enable_daily_trigger').value=String(s.enable_daily_trigger||'FALSE').toUpperCase()==='TRUE'?'TRUE':'FALSE';
      byId('default_template_pack').value=s.default_template_pack||'general';
      byId('sender_name').value=s.sender_name||'';
      byId('email_subject_template').value=s.email_subject_template||'';
      byId('email_body_template').value=s.email_body_template||'';
      byId('rules_enabled').value=String(s.rules_enabled||'TRUE').toUpperCase()==='FALSE'?'FALSE':'TRUE';
      setStatus('Loaded.',true);
    }).withFailureHandler(function(err){setError(err&&err.message?err.message:String(err));}).getWizardData();
  }
  function save(){
    var payload={
      stage_pack:byId('stage_pack').value,
      send_mode:byId('send_mode').value,
      review_send_source:byId('review_send_source').value,
      max_emails_per_run:byId('max_emails_per_run').value,
      business_days_only:byId('business_days_only').value,
      follow_up_days:byId('follow_up_days').value,
      autopilot_hour_local:byId('autopilot_hour_local').value,
      enable_daily_trigger:byId('enable_daily_trigger').value,
      default_template_pack:byId('default_template_pack').value,
      sender_name:byId('sender_name').value,
      email_subject_template:byId('email_subject_template').value,
      email_body_template:byId('email_body_template').value,
      rules_enabled:byId('rules_enabled').value
    };
    setStatus('Saving…');
    google.script.run
      .withSuccessHandler(function(){setStatus('✅ Saved. Reload Sheet to apply stage changes.',true);})
      .withFailureHandler(function(err){setError(err&&err.message?err.message:String(err));})
      .saveWizardData(payload);
  }
  load();
</script></body></html>`;
}

// ═══════════════════════════════════════════════════════════════
//  CRM SIDEBAR HTML
// ═══════════════════════════════════════════════════════════════

function getCrmSidebarHtml_() {
  return `<!DOCTYPE html><html><head><base target="_top">
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<style>
  :root{--bg:#0b0f14;--panel:#111823;--panel2:#0f1620;--text:#e6edf3;--muted:rgba(230,237,243,.70);--muted2:rgba(230,237,243,.45);--border:rgba(230,237,243,.10);--accent:#7aa2ff;--good:#33d17a;--warn:#ffcc00;--bad:#ff5c5c;}
  *{box-sizing:border-box;}
  body{margin:0;font-family:ui-sans-serif,system-ui,-apple-system,Arial,sans-serif;background:var(--bg);color:var(--text);font-size:12px;}
  .wrap{padding:10px;}
  .topbar{display:flex;gap:6px;align-items:center;margin-bottom:8px;flex-wrap:wrap;}
  .title{font-weight:800;font-size:13px;margin-right:auto;}
  .btn{background:var(--panel);border:1px solid var(--border);color:var(--text);border-radius:9px;padding:6px 9px;font-weight:600;cursor:pointer;font-size:11px;}
  .btn.primary{border-color:rgba(122,162,255,.45);background:rgba(122,162,255,.12);}
  .btn:disabled{opacity:.5;cursor:not-allowed;}
  .fin-bar{display:grid;grid-template-columns:repeat(3,1fr);gap:5px;margin-bottom:8px;}
  .fin{background:var(--panel2);border:1px solid var(--border);border-radius:10px;padding:7px;}
  .fin .label{color:var(--muted2);font-size:10px;} .fin .val{font-size:13px;font-weight:800;margin-top:2px;}
  .kpis{display:grid;grid-template-columns:repeat(5,1fr);gap:5px;margin-bottom:8px;}
  .kpi{background:var(--panel2);border:1px solid var(--border);border-radius:10px;padding:7px;}
  .kpi .label{color:var(--muted2);font-size:10px;} .kpi .val{font-size:14px;font-weight:800;margin-top:2px;}
  .tabs{display:flex;gap:5px;margin-bottom:8px;}
  .tab{flex:1;text-align:center;padding:7px;border-radius:9px;background:var(--panel2);border:1px solid var(--border);cursor:pointer;font-weight:700;font-size:11px;color:var(--muted);}
  .tab.active{background:rgba(122,162,255,.14);border-color:rgba(122,162,255,.45);color:var(--accent);}
  .filters{display:grid;grid-template-columns:1fr auto;gap:5px;margin-bottom:8px;}
  .input,.select{background:var(--panel);border:1px solid var(--border);color:var(--text);border-radius:9px;padding:7px 9px;font-size:11px;outline:none;width:100%;}
  .today-list{display:grid;gap:6px;}
  .tcard{background:var(--panel2);border:1px solid rgba(255,204,0,.30);border-radius:12px;padding:10px;}
  .tcard-top{display:flex;justify-content:space-between;align-items:baseline;}
  .tcard-name{font-weight:900;font-size:13px;} .tcard-val{font-weight:700;color:var(--warn);font-size:12px;}
  .tcard-co{color:var(--muted);font-size:11px;margin-top:2px;}
  .tcard-score{font-size:10px;background:rgba(122,162,255,.15);border:1px solid rgba(122,162,255,.3);color:var(--accent);padding:2px 6px;border-radius:999px;margin-top:4px;display:inline-block;}
  .tcard-actions{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:5px;margin-top:7px;}
  .board{display:grid;grid-auto-flow:column;grid-auto-columns:minmax(180px,1fr);gap:6px;align-items:start;}
  .col{background:var(--panel2);border:1px solid var(--border);border-radius:12px;overflow:hidden;}
  .colhead{padding:8px 10px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;}
  .colhead .name{font-weight:800;font-size:11px;} .colhead .count{color:var(--muted);font-size:10px;}
  .cards{padding:6px;display:grid;gap:6px;}
  .card{background:rgba(255,255,255,.03);border:1px solid rgba(255,255,255,.08);border-radius:12px;padding:8px;}
  .card.due{border-color:rgba(255,204,0,.45);background:rgba(255,204,0,.06);}
  .card.error{border-color:rgba(255,92,92,.45);background:rgba(255,92,92,.06);}
  .cardtop{display:flex;gap:6px;align-items:baseline;}
  .cardtop .who{font-weight:900;font-size:11px;} .cardtop .co{color:var(--muted);font-size:10px;margin-left:auto;}
  .meta{margin-top:5px;display:grid;gap:3px;} .row{display:flex;gap:4px;align-items:center;flex-wrap:wrap;}
  .pill{padding:3px 7px;border-radius:999px;border:1px solid var(--border);font-size:10px;color:var(--muted);}
  .pill.good{color:rgba(51,209,122,.95);border-color:rgba(51,209,122,.35);}
  .pill.warn{color:rgba(255,204,0,.95);border-color:rgba(255,204,0,.35);}
  .pill.bad{color:rgba(255,92,92,.95);border-color:rgba(255,92,92,.35);}
  .actions{margin-top:6px;display:grid;grid-template-columns:1fr 1fr;gap:5px;}
  .smallbtn{background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.10);color:var(--text);border-radius:9px;padding:6px 7px;font-weight:800;cursor:pointer;font-size:10px;text-align:center;user-select:none;}
  .smallbtn.primary{border-color:rgba(122,162,255,.45);background:rgba(122,162,255,.12);}
  .smallbtn.good{border-color:rgba(51,209,122,.35);background:rgba(51,209,122,.10);}
  .smallbtn.warn{border-color:rgba(255,204,0,.35);background:rgba(255,204,0,.10);}
  .smallbtn.bad{border-color:rgba(255,92,92,.35);background:rgba(255,92,92,.10);}
  .smallbtn:active{transform:translateY(1px);}
  .empty{color:var(--muted2);text-align:center;padding:24px;font-size:11px;}
  .toast{position:fixed;left:10px;right:10px;bottom:10px;background:rgba(17,24,35,.98);border:1px solid rgba(230,237,243,.14);border-radius:12px;padding:9px 12px;font-size:11px;display:none;z-index:999;}
  .toast.show{display:block;} .toast .sub{color:var(--muted);margin-top:2px;font-size:10px;}
</style></head>
<body><div class="wrap">
  <div class="topbar">
    <div class="title">Command Center</div>
    <button class="btn" id="refreshBtn" onclick="load()">↺</button>
    <button class="btn primary" onclick="runReview()">Run REVIEW</button>
    <button class="btn primary" onclick="sendApproved()">Send ✓</button>
  </div>
  <div class="fin-bar">
    <div class="fin"><div class="label">Pipeline</div><div class="val" id="f_pipeline">—</div></div>
    <div class="fin"><div class="label">Weighted</div><div class="val" id="f_weighted">—</div></div>
    <div class="fin"><div class="label">Due Value</div><div class="val" id="f_overdue">—</div></div>
  </div>
  <div class="kpis">
    <div class="kpi"><div class="label">Due</div><div class="val" id="k_due">0</div></div>
    <div class="kpi"><div class="label">Drafted</div><div class="val" id="k_drafted">0</div></div>
    <div class="kpi"><div class="label">Approved</div><div class="val" id="k_approved">0</div></div>
    <div class="kpi"><div class="label">Errors</div><div class="val" id="k_errors">0</div></div>
    <div class="kpi"><div class="label">Total</div><div class="val" id="k_total">0</div></div>
  </div>
  <div class="tabs">
    <div class="tab active" id="tabToday" onclick="switchTab('today')">⚡ Today</div>
    <div class="tab" id="tabKanban" onclick="switchTab('kanban')">📋 Pipeline</div>
  </div>
  <div class="filters">
    <input class="input" id="search" placeholder="Search…" oninput="render()"/>
    <select class="select" id="statusFilter" onchange="render()" style="width:auto;min-width:80px"></select>
  </div>
  <div id="viewToday" class="today-list"></div>
  <div id="viewKanban" class="board" style="display:none"></div>
</div>
<div class="toast" id="toast">
  <div id="toastTitle"></div>
  <div class="sub" id="toastSub"></div>
</div>
<script>
  var STATE={leads:[],kpis:{},pipeline:{}};
  var viewMode='today';

  function toast(title,sub){
    var t=document.getElementById('toast');
    document.getElementById('toastTitle').textContent=title||'';
    document.getElementById('toastSub').textContent=sub||'';
    t.classList.add('show');
    setTimeout(function(){t.classList.remove('show');},2600);
  }
  function esc(s){return String(s||'').replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];});}
  function fmt$(n){n=parseInt(n,10)||0;if(!n)return'$0';if(n>=1000000)return'$'+(n/1e6).toFixed(1)+'M';if(n>=1000)return'$'+(n/1000).toFixed(1)+'K';return'$'+n.toLocaleString();}

  function switchTab(mode){
    viewMode=mode;
    document.getElementById('tabToday').className='tab'+(mode==='today'?' active':'');
    document.getElementById('tabKanban').className='tab'+(mode==='kanban'?' active':'');
    document.getElementById('viewToday').style.display=mode==='today'?'':'none';
    document.getElementById('viewKanban').style.display=mode==='kanban'?'':'none';
    render();
  }

  function load(){
    document.getElementById('refreshBtn').disabled=true;
    google.script.run
      .withSuccessHandler(function(data){STATE=data||{leads:[],kpis:{},pipeline:{}};updateKpis();populateStatusFilter();render();document.getElementById('refreshBtn').disabled=false;toast('Loaded','Pipeline updated');})
      .withFailureHandler(function(err){document.getElementById('refreshBtn').disabled=false;toast('Error',err&&err.message?err.message:String(err));})
      .uiGetState();
  }

  function updateKpis(){
    var k=STATE.kpis||{};
    document.getElementById('k_due').textContent=k.due||0;
    document.getElementById('k_drafted').textContent=k.drafted||0;
    document.getElementById('k_approved').textContent=k.approved||0;
    document.getElementById('k_errors').textContent=k.errors||0;
    document.getElementById('k_total').textContent=k.total||0;
    document.getElementById('f_pipeline').textContent=fmt$(k.pipelineValue||0);
    document.getElementById('f_weighted').textContent=fmt$(k.weightedPipeline||0);
    document.getElementById('f_overdue').textContent=fmt$(k.overdueValue||0);
  }

  function populateStatusFilter(){
    var stages=(STATE.pipeline&&STATE.pipeline.stages)||['NEW','CONTACTED','NURTURE','BOOKED','CLOSED'];
    var sel=document.getElementById('statusFilter');
    sel.innerHTML='<option value="all">All</option>'+stages.map(function(s){return'<option value="'+esc(s)+'">'+esc(s)+'</option>';}).join('');
  }

  function matchesFilters(lead){
    var q=document.getElementById('search').value.toLowerCase().trim();
    var status=document.getElementById('statusFilter').value;
    if(q&&![lead.name,lead.company,lead.email,lead.phone].join(' ').toLowerCase().includes(q))return false;
    if(status!=='all'&&lead.status!==status)return false;
    return true;
  }

  function scoreEmoji(s){if(s>=8)return'🔥';if(s>=5)return'⬆';if(s>=3)return'◆';return'·';}

  function render(){if(viewMode==='today')renderToday();else renderKanban();}

  function renderToday(){
    var container=document.getElementById('viewToday');
    var leads=(STATE.leads||[]).filter(function(l){return l.dueFlag;}).filter(matchesFilters)
      .sort(function(a,b){return(b.priorityScore||0)-(a.priorityScore||0)||(parseFloat(String(b.dealValue||'0').replace(/[^0-9.]/g,''))||0)-(parseFloat(String(a.dealValue||'0').replace(/[^0-9.]/g,''))||0);});
    if(!leads.length){container.innerHTML='<div class="empty">No due leads today 🎉<br>Relax or add new leads.</div>';return;}
    container.innerHTML=leads.map(function(l){
      return '<div class="tcard">'+
        '<div class="tcard-top"><div class="tcard-name">'+esc(l.preferredName||l.firstName||l.name||'Lead')+'</div>'+
        '<div class="tcard-val">'+(l.dealValue?'$'+esc(l.dealValue):'')+'</div></div>'+
        '<div class="tcard-co">'+esc(l.company||'')+'&nbsp;·&nbsp;'+esc(l.status||'')+'</div>'+
        '<div class="tcard-score">'+scoreEmoji(l.priorityScore||0)+' P-Score: '+(l.priorityScore||0)+'/10</div>'+
        '<div class="tcard-actions">'+
        (l.phone?'<div class="smallbtn" onclick="callLead(\''+esc(l.phone)+'\')">📞 Call</div>':'<div class="smallbtn" style="opacity:.3">Call</div>')+
        (l.phone?'<div class="smallbtn" onclick="openText(\''+esc(l.phone)+'\',\''+esc(l.textPreview||'')+'\')">💬 Text</div>':'<div class="smallbtn" style="opacity:.3">Text</div>')+
        '<div class="smallbtn warn" onclick="markCalled('+l.row+')">Called ✓</div>'+
        '<div class="smallbtn" onclick="snooze('+l.row+',1)">Snooze 1d</div>'+
        '</div></div>';
    }).join('');
  }

  function renderKanban(){
    var pipeline=(STATE.pipeline&&STATE.pipeline.stages)||['NEW','CONTACTED','NURTURE','BOOKED','CLOSED'];
    var colors=(STATE.pipeline&&STATE.pipeline.colors)||{};
    var board=document.getElementById('viewKanban');
    board.innerHTML='';
    var leads=(STATE.leads||[]).filter(matchesFilters);
    pipeline.forEach(function(stage){
      var colLeads=leads.filter(function(l){return l.status===stage;});
      var col=document.createElement('div');col.className='col';col.dataset.stage=stage;
      var head=document.createElement('div');head.className='colhead';
      head.innerHTML='<div class="name">'+esc(stage)+'</div><div class="count">'+colLeads.length+'</div>';
      var c=colors[stage]||'';
      if(c)head.style.boxShadow='inset 0 -3px 0 '+c;
      col.appendChild(head);
      var cards=document.createElement('div');cards.className='cards';cards.dataset.stage=stage;
      cards.addEventListener('dragover',function(ev){ev.preventDefault();ev.dataTransfer.dropEffect='move';});
      cards.addEventListener('drop',function(ev){
        ev.preventDefault();
        try{var d=JSON.parse(ev.dataTransfer.getData('text/plain'));if(d&&d.row&&d.stage!==stage)moveLead(d.row,stage);}catch(e){}
      });
      if(!colLeads.length)cards.innerHTML='<div class="empty">—</div>';
      colLeads.forEach(function(l){
        var isError=l.draftStatus==='ERROR'||(l.lastError||'').length>0;
        var card=document.createElement('div');
        card.className='card'+(l.dueFlag?' due':'')+(isError?' error':'');
        card.draggable=true;
        card.addEventListener('dragstart',function(ev){ev.dataTransfer.setData('text/plain',JSON.stringify({row:l.row,stage:stage}));ev.dataTransfer.effectAllowed='move';});
        var duePill=l.dueFlag?'<span class="pill warn">DUE</span>':'<span class="pill">Next: '+esc(l.followUpDue||'-')+'</span>';
        var draftPill=l.draftStatus==='DRAFTED'?'<span class="pill good">Drafted</span>':l.draftStatus==='ERROR'?'<span class="pill bad">Error</span>':l.draftStatus?'<span class="pill">'+esc(l.draftStatus)+'</span>':'';
        var valuePill=l.dealValue?'<span class="pill">$'+esc(l.dealValue)+'</span>':'';
        var scorePill='<span class="pill">'+scoreEmoji(l.priorityScore||0)+''+( l.priorityScore||0)+'</span>';
        card.innerHTML='<div class="cardtop"><div class="who">'+esc(l.preferredName||l.firstName||l.name||'Lead')+'</div><div class="co">'+esc(l.company||'')+'</div></div>'+
          '<div class="meta"><div class="row">'+duePill+draftPill+valuePill+scorePill+'</div>'+
          (isError?'<div class="row"><span class="pill bad">'+esc(l.lastError||'Error')+'</span></div>':'')+
          '</div><div class="actions">'+
          (l.phone?'<div class="smallbtn" onclick="callLead(\''+esc(l.phone)+'\')">📞 Call</div>':'')+
          '<div class="smallbtn '+(l.approved?'bad':'good')+'" onclick="setApproved('+l.row+','+(!l.approved)+')">'+(l.approved?'Unapprove':'Approve')+'</div>'+
          '<div class="smallbtn warn" onclick="markCalled('+l.row+')">Called ✓</div>'+
          '<div class="smallbtn" onclick="snooze('+l.row+',1)">Snooze 1d</div>'+
          '</div>';
        cards.appendChild(card);
      });
      col.appendChild(cards);board.appendChild(col);
    });
  }

  function runReview(){toast('Running','Generating drafts…');google.script.run.withSuccessHandler(function(res){toast('Done','Drafted:'+(res.drafted||0));load();}).withFailureHandler(function(err){toast('Error',err&&err.message?err.message:String(err));}).uiRunReview();}
  function sendApproved(){toast('Sending','Sending approved drafts…');google.script.run.withSuccessHandler(function(res){toast('Done','Sent:'+(res.sent||0)+(res.errors?' Errors:'+res.errors:''));load();}).withFailureHandler(function(err){toast('Error',err&&err.message?err.message:String(err));}).uiSendApproved();}
  function markCalled(row){toast('Updating','Marking called');google.script.run.withSuccessHandler(function(){load();}).withFailureHandler(function(err){toast('Error',err&&err.message?err.message:String(err));}).uiMarkCalled(row);}
  function snooze(row,days){toast('Snoozing','Adding '+days+'d');google.script.run.withSuccessHandler(function(){load();}).withFailureHandler(function(err){toast('Error',err&&err.message?err.message:String(err));}).uiSnooze(row,days);}
  function setApproved(row,val){google.script.run.withSuccessHandler(function(){toast('Updated',val?'Approved':'Unapproved');load();}).withFailureHandler(function(err){toast('Error',err&&err.message?err.message:String(err));}).uiUpdateLead(row,{'Approved to Send':val?'TRUE':''});}
  function moveLead(row,stage){toast('Moving','Updating to '+stage);google.script.run.withSuccessHandler(function(){load();}).withFailureHandler(function(err){toast('Error',err&&err.message?err.message:String(err));}).uiMoveLead(row,stage);}
  function callLead(phone){if(!phone){toast('No phone','');return;}window.open('tel:'+encodeURIComponent(phone),'_blank');}
  function openText(phone,msg){if(!phone){toast('No phone','');return;}window.open('sms:'+encodeURIComponent(phone)+'?body='+encodeURIComponent(msg||''),'_blank');if(msg)navigator.clipboard.writeText(msg).catch(function(){});toast('Text','Opened Messages. Message pre-copied.');}
  load();
</script></body></html>`;
}

// ═══════════════════════════════════════════════════════════════
//  FULL SCREEN COMMAND CENTER HTML
//  v8: Schedule modal, Quick Add modal, Activity Timeline tab,
//      Auto-refresh, Priority Score badge, XSS-safe rendering
// ═══════════════════════════════════════════════════════════════

function getCommandCenterAppHtml_() {
  return `<!DOCTYPE html><html><head><base target="_top">
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<style>
  *{box-sizing:border-box;} html,body{margin:0;padding:0;background:#0b0f14;color:#e6edf3;font-family:ui-sans-serif,system-ui,-apple-system,Arial,sans-serif;}
  .wrap{max-width:1240px;margin:0 auto;padding:14px;}
  .topbar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:12px;}
  .title{font-weight:800;font-size:16px;margin-right:auto;}
  .hint{font-size:11px;color:rgba(230,237,243,.45);margin-top:2px;}
  button{background:#111827;color:#e6edf3;border:1px solid rgba(255,255,255,.12);padding:8px 12px;border-radius:10px;cursor:pointer;font-weight:600;font-size:13px;}
  button.primary{background:rgba(122,162,255,.18);border-color:rgba(122,162,255,.45);}
  button.success{background:rgba(51,209,122,.15);border-color:rgba(51,209,122,.4);}
  button:disabled{opacity:.5;cursor:not-allowed;}
  .refresh-bar{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin:0 0 12px;padding:8px 12px;background:#0f172a;border:1px solid rgba(255,255,255,.08);border-radius:12px;}
  .refresh-bar label{display:flex;gap:6px;align-items:center;font-size:12px;color:rgba(230,237,243,.65);}
  .refresh-bar input[type=checkbox]{width:14px;height:14px;}
  .refresh-bar input[type=number]{width:70px;padding:4px 7px;border-radius:8px;border:1px solid rgba(255,255,255,.12);background:#0b1220;color:#e6edf3;font-size:12px;}
  .fin-bar{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:12px;}
  .fin{background:#0f172a;border:1px solid rgba(255,255,255,.10);border-radius:14px;padding:12px;}
  .fin .label{font-size:12px;color:rgba(230,237,243,.55);} .fin .val{font-size:22px;font-weight:800;margin-top:4px;}
  .kpis{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:12px;}
  .kpi{background:#0f172a;border:1px solid rgba(255,255,255,.10);border-radius:14px;padding:12px;}
  .kpi .label{font-size:12px;color:rgba(230,237,243,.55);} .kpi .val{font-size:20px;font-weight:800;margin-top:4px;}
  .tabs{display:flex;gap:8px;margin-bottom:12px;}
  .tab{padding:8px 14px;border-radius:10px;background:#0f172a;border:1px solid rgba(255,255,255,.10);cursor:pointer;font-weight:700;font-size:13px;color:rgba(230,237,243,.55);}
  .tab.active{background:rgba(122,162,255,.14);border-color:rgba(122,162,255,.45);color:#7aa2ff;}
  .filters{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px;}
  input[type=text],select{background:#0b1220;color:#e6edf3;border:1px solid rgba(255,255,255,.12);padding:8px 10px;border-radius:10px;font-size:13px;}
  input[type=text]{flex:1;min-width:200px;}
  select{min-width:150px;}
  .list{display:grid;gap:10px;}
  .card{background:#0f172a;border:1px solid rgba(255,255,255,.10);border-radius:14px;padding:14px;}
  .card.due{border-color:rgba(255,204,0,.40);background:rgba(255,204,0,.05);}
  .card.error{border-color:rgba(255,92,92,.40);background:rgba(255,92,92,.05);}
  .cardTop{display:flex;justify-content:space-between;gap:10px;align-items:flex-start;}
  .name{font-weight:900;font-size:15px;} .sub{font-size:12px;color:rgba(230,237,243,.55);margin-top:3px;}
  .pills{display:flex;gap:6px;flex-wrap:wrap;margin-top:8px;}
  .pill{font-size:12px;padding:4px 9px;border-radius:999px;border:1px solid rgba(255,255,255,.12);color:rgba(230,237,243,.70);}
  .pill.due{color:rgba(255,204,0,.95);border-color:rgba(255,204,0,.35);}
  .pill.green{color:rgba(51,209,122,.95);border-color:rgba(51,209,122,.35);}
  .pill.red{color:rgba(255,92,92,.95);border-color:rgba(255,92,92,.35);}
  .pill.blue{color:rgba(122,162,255,.95);border-color:rgba(122,162,255,.35);}
  .actions{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px;}
  .actions button{padding:7px 11px;font-size:12px;}
  .empty{color:rgba(230,237,243,.35);text-align:center;padding:60px;font-size:14px;}
  .err{display:none;background:#2b0b0b;border:1px solid rgba(255,80,80,.35);padding:10px;border-radius:12px;margin-bottom:12px;font-size:13px;}
  /* Activity Tab */
  .act-list{display:grid;gap:8px;}
  .act-item{background:#0f172a;border:1px solid rgba(255,255,255,.08);border-radius:12px;padding:10px 14px;display:flex;gap:12px;align-items:flex-start;}
  .act-ts{font-size:11px;color:rgba(230,237,243,.40);min-width:110px;}
  .act-body{flex:1;}
  .act-action{font-weight:700;font-size:12px;color:#7aa2ff;}
  .act-notes{font-size:12px;color:rgba(230,237,243,.65);margin-top:2px;}
  .act-who{font-size:11px;color:rgba(230,237,243,.45);}
  /* MODAL */
  .modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:100;align-items:center;justify-content:center;}
  .modal-overlay.open{display:flex;}
  .modal{background:#111823;border:1px solid rgba(255,255,255,.14);border-radius:16px;padding:24px;width:500px;max-width:95vw;max-height:90vh;overflow-y:auto;}
  .modal h3{margin:0 0 16px;font-size:16px;color:#7aa2ff;}
  .modal .field{margin-bottom:12px;}
  .modal label{display:block;font-size:12px;color:rgba(230,237,243,.65);margin-bottom:4px;font-weight:600;}
  .modal input,.modal select,.modal textarea{width:100%;background:#0b1220;color:#e6edf3;border:1px solid rgba(255,255,255,.15);padding:9px 11px;border-radius:10px;font-size:13px;}
  .modal textarea{resize:vertical;font-family:inherit;}
  .modal .row2{display:grid;grid-template-columns:1fr 1fr;gap:10px;}
  .modal .btn-row{display:flex;gap:8px;margin-top:18px;justify-content:flex-end;}
  .modal .btn-row button{padding:9px 16px;font-size:13px;}
  .modal .moderr{color:#ff5c5c;font-size:12px;margin-top:8px;display:none;}
  @media(max-width:900px){.fin-bar{grid-template-columns:1fr 1fr;}.kpis{grid-template-columns:repeat(2,1fr);}  }
</style></head>
<body>
<div class="wrap">
  <div class="topbar">
    <div>
      <div class="title">Command Center <span style="font-size:12px;opacity:.5;font-weight:400">v${LFU.VERSION}</span></div>
      <div class="hint">REVIEW mode: drafts first, you approve, then Send Approved.</div>
    </div>
    <button id="btnAddLead" class="success" onclick="openQuickAdd()">+ Quick Add Lead</button>
    <button id="btnRefresh">↺ Refresh</button>
    <button id="btnRun" class="primary">Run REVIEW</button>
    <button id="btnSend" class="primary">Send Approved</button>
  </div>

  <!-- Auto-refresh controls -->
  <div class="refresh-bar">
    <label><input type="checkbox" id="autoRefreshToggle" checked/> Auto-refresh</label>
    <label>Interval (sec): <input type="number" id="autoRefreshSec" min="5" max="300" value="30"/></label>
    <span style="font-size:11px;color:rgba(230,237,243,.35);">Min 5s. Refresh only reads data, never sends.</span>
    <span style="font-size:11px;color:rgba(51,209,122,.6);margin-left:auto;" id="lastRefreshLabel"></span>
  </div>

  <div id="err" class="err"></div>

  <div class="fin-bar">
    <div class="fin"><div class="label">Pipeline Value</div><div class="val" id="fPipeline">—</div></div>
    <div class="fin"><div class="label">Weighted Pipeline</div><div class="val" id="fWeighted">—</div></div>
    <div class="fin"><div class="label">Due Today Value</div><div class="val" id="fOverdue">—</div></div>
  </div>
  <div class="kpis">
    <div class="kpi"><div class="label">Due Now</div><div class="val" id="kDue">—</div></div>
    <div class="kpi"><div class="label">Drafted</div><div class="val" id="kDrafted">—</div></div>
    <div class="kpi"><div class="label">Approved</div><div class="val" id="kApproved">—</div></div>
    <div class="kpi"><div class="label">Errors</div><div class="val" id="kErrors">—</div></div>
    <div class="kpi"><div class="label">Total</div><div class="val" id="kTotal">—</div></div>
  </div>

  <div class="tabs">
    <div class="tab active" id="tabToday"    onclick="switchTab('today')">⚡ Today</div>
    <div class="tab"        id="tabAll"      onclick="switchTab('all')">📋 All Leads</div>
    <div class="tab"        id="tabActivity" onclick="switchTab('activity')">📜 Activity Feed</div>
  </div>
  <div class="filters" id="filterRow">
    <input type="text" id="q" placeholder="Search name, company, phone, email…"/>
    <select id="statusFilter"><option value="">All statuses</option></select>
  </div>
  <div id="list" class="list"></div>
  <div id="actList" class="act-list" style="display:none"></div>
</div>

<!-- ── SCHEDULE MODAL ── -->
<div class="modal-overlay" id="scheduleModal">
  <div class="modal">
    <h3>📅 Schedule Appointment</h3>
    <div id="scheduleLeadName" style="margin-bottom:14px;font-weight:700;font-size:14px;color:#e6edf3;"></div>
    <div class="row2">
      <div class="field"><label>Date (YYYY-MM-DD)</label><input type="text" id="schedDate" placeholder="2025-01-15"/></div>
      <div class="field"><label>Time (HH:MM, 24-hour)</label><input type="text" id="schedTime" placeholder="14:00"/></div>
    </div>
    <div class="row2">
      <div class="field"><label>Duration (minutes)</label><input type="number" id="schedDur" value="30" min="5" max="480"/></div>
      <div class="field"><label>Location (optional)</label><input type="text" id="schedLoc" placeholder="Zoom / 123 Main St"/></div>
    </div>
    <div class="moderr" id="schedErr"></div>
    <div class="btn-row">
      <button onclick="closeModal('scheduleModal')">Cancel</button>
      <button class="primary" onclick="submitSchedule()">Create Appointment</button>
    </div>
  </div>
</div>

<!-- ── QUICK ADD LEAD MODAL ── -->
<div class="modal-overlay" id="addLeadModal">
  <div class="modal">
    <h3>+ Quick Add Lead</h3>
    <div class="row2">
      <div class="field"><label>Name *</label><input type="text" id="addName" placeholder="Jane Smith"/></div>
      <div class="field"><label>Email</label><input type="text" id="addEmail" placeholder="jane@example.com"/></div>
    </div>
    <div class="row2">
      <div class="field"><label>Phone</label><input type="text" id="addPhone" placeholder="+1 555 000 0000"/></div>
      <div class="field"><label>Company</label><input type="text" id="addCompany" placeholder="Acme Co."/></div>
    </div>
    <div class="row2">
      <div class="field"><label>Deal Value ($)</label><input type="text" id="addDealValue" placeholder="2500"/></div>
      <div class="field"><label>Status</label>
        <select id="addStatus"></select>
      </div>
    </div>
    <div class="field"><label>Owner</label><input type="text" id="addOwner" placeholder="Your name"/></div>
    <div class="field"><label>Notes</label><textarea id="addNotes" rows="2" placeholder="Source, context…"></textarea></div>
    <div class="moderr" id="addErr"></div>
    <div class="btn-row">
      <button onclick="closeModal('addLeadModal')">Cancel</button>
      <button class="success" onclick="submitQuickAdd()">Add Lead</button>
    </div>
  </div>
</div>

<script>
  var STATE = null;
  var viewMode = 'today';
  var busy = false;
  var autoTimer = null;
  var pendingScheduleRow = null;

  var el = function(id){return document.getElementById(id);};
  var esc = function(s){return String(s||'').replace(/[&<>"']/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];});};
  var fmt$ = function(n){n=parseInt(n,10)||0;if(!n)return'$0';if(n>=1e6)return'$'+(n/1e6).toFixed(1)+'M';if(n>=1000)return'$'+(n/1000).toFixed(1)+'K';return'$'+n.toLocaleString();};
  var scoreEmoji = function(s){if(s>=8)return'🔥';if(s>=5)return'⬆';if(s>=3)return'◆';return'·';};

  function setBusy(b){busy=b;['btnRefresh','btnRun','btnSend','btnAddLead'].forEach(function(id){var e=el(id);if(e)e.disabled=b;});}
  function showErr(m){var e=el('err');e.style.display='block';e.textContent=m;}
  function clearErr(){el('err').style.display='none';}

  function switchTab(mode){
    viewMode=mode;
    ['tabToday','tabAll','tabActivity'].forEach(function(id){var t=el(id);if(t)t.className='tab';});
    el('tab'+mode.charAt(0).toUpperCase()+mode.slice(1)).className='tab active';
    el('filterRow').style.display=mode==='activity'?'none':'flex';
    el('list').style.display=mode==='activity'?'none':'grid';
    el('actList').style.display=mode==='activity'?'grid':'none';
    if(mode==='activity'){loadActivities();}else if(STATE){render();}
  }

  function load(){
    if(busy)return;
    setBusy(true);clearErr();
    google.script.run
      .withSuccessHandler(function(data){
        setBusy(false);STATE=data;
        renderKpis();buildStatusFilter();populateAddStatus();render();
        var now=new Date();el('lastRefreshLabel').textContent='Refreshed '+now.getHours()+':'+String(now.getMinutes()).padStart(2,'0')+':'+String(now.getSeconds()).padStart(2,'0');
      })
      .withFailureHandler(function(e){setBusy(false);showErr(e&&e.message?e.message:String(e));})
      .uiGetState();
  }

  function renderKpis(){
    if(!STATE)return;
    var k=STATE.kpis||{};
    el('kDue').textContent=k.due||0;el('kDrafted').textContent=k.drafted||0;
    el('kApproved').textContent=k.approved||0;el('kErrors').textContent=k.errors||0;el('kTotal').textContent=k.total||0;
    el('fPipeline').textContent=fmt$(k.pipelineValue||0);el('fWeighted').textContent=fmt$(k.weightedPipeline||0);el('fOverdue').textContent=fmt$(k.overdueValue||0);
  }

  function buildStatusFilter(){
    if(!STATE)return;
    var stages=(STATE.pipeline&&STATE.pipeline.stages)||[];
    var statuses=stages.length?stages:[...new Set((STATE.leads||[]).map(function(l){return l.status;}).filter(Boolean))].sort();
    el('statusFilter').innerHTML='<option value="">All statuses</option>'+statuses.map(function(s){return'<option value="'+esc(s)+'">'+esc(s)+'</option>';}).join('');
  }

  function populateAddStatus(){
    var sel=el('addStatus');
    if(!sel||!STATE)return;
    var stages=(STATE.pipeline&&STATE.pipeline.stages)||['NEW','CONTACTED','NURTURE','BOOKED','CLOSED'];
    sel.innerHTML=stages.map(function(s){return'<option value="'+esc(s)+'">'+esc(s)+'</option>';}).join('');
  }

  function render(){
    if(!STATE)return;
    var q=el('q').value.trim().toLowerCase();
    var status=el('statusFilter').value;
    var leads=(STATE.leads||[]).slice();
    if(viewMode==='today')leads=leads.filter(function(l){return l.dueFlag===true;});
    if(q)leads=leads.filter(function(l){return[l.name,l.company,l.phone,l.email].join(' ').toLowerCase().includes(q);});
    if(status)leads=leads.filter(function(l){return l.status===status;});
    leads.sort(function(a,b){
      if(b.dueFlag!==a.dueFlag)return b.dueFlag?1:-1;
      return(b.priorityScore||0)-(a.priorityScore||0)||(parseFloat(String(b.dealValue||'0').replace(/[^0-9.]/g,''))||0)-(parseFloat(String(a.dealValue||'0').replace(/[^0-9.]/g,''))||0);
    });
    if(!leads.length){
      el('list').innerHTML=viewMode==='today'?'<div class="empty">No due leads today 🎉<br>Add leads with Quick Add or run Import Wizard.</div>':'<div class="empty">No leads match your filters.</div>';
      return;
    }
    el('list').innerHTML=leads.map(function(l){
      var isDue=l.dueFlag===true;
      var row=l.row;
      var duePill=isDue?'<span class="pill due">DUE '+esc(l.followUpDue||'')+'</span>':'<span class="pill">Next: '+esc(l.followUpDue||'-')+'</span>';
      var scorePill='<span class="pill blue">'+scoreEmoji(l.priorityScore||0)+' P'+esc(String(l.priorityScore||0))+'/10</span>';
      var draftPill=l.draftStatus?'<span class="pill '+(l.draftStatus==='DRAFTED'?'green':l.draftStatus==='ERROR'?'red':'')+'">'+(l.draftStatus==='DRAFTED'?'✎ Drafted':l.draftStatus==='ERROR'?'⚠ Error':esc(l.draftStatus))+'</span>':'';
      var valuePill=l.dealValue?'<span class="pill">$'+esc(l.dealValue)+'</span>':'';
      var approvedPill=l.approved?'<span class="pill green">✓ APPROVED</span>':'';
      var phonePill=l.phone?'<span class="pill">📞 '+esc(l.phone)+'</span>':'';
      var errPill=l.lastError?'<span class="pill red">⚠ '+esc(l.lastError)+'</span>':'';
      var apptPill=l.apptStart?'<span class="pill">📅 '+esc(l.apptStart)+'</span>':'';
      var calBtn=l.calLink?'<button onclick="window.open(\''+esc(l.calLink)+'\',\'_blank\')">📆 Cal Event</button>':'';
      var ownerPill=l.owner?'<span class="pill">👤 '+esc(l.owner)+'</span>':'';
      return'<div class="card'+(isDue?' due':'')+(l.draftStatus==='ERROR'?' error':'')+'">'+
        '<div class="cardTop"><div><div class="name">'+esc(l.name||'(no name)')+'</div>'+
        '<div class="sub">'+esc(l.company||'—')+' · '+esc(l.status||'—')+'</div></div>'+ownerPill+'</div>'+
        '<div class="pills">'+duePill+scorePill+draftPill+approvedPill+valuePill+phonePill+
        (l.email?'<span class="pill">'+esc(l.email)+'</span>':'')+errPill+apptPill+'</div>'+
        '<div class="actions">'+
        (l.phone?'<button onclick="window.open(\'tel:'+encodeURIComponent(l.phone||'')+'\',\'_blank\')">📞 Call</button>':'')+
        (l.phone?'<button onclick="window.open(\'sms:'+encodeURIComponent(l.phone||'')+'\',\'_blank\')">💬 Text</button>':'')+
        (l.email?'<button onclick="window.open(\'mailto:'+encodeURIComponent(l.email||'')+'\',\'_blank\')">✉ Email</button>':'')+
        '<button onclick="markCalled('+row+')">✅ Mark Called</button>'+
        '<button onclick="snoozeRow('+row+',1)">Snooze 1d</button>'+
        '<button onclick="snoozeRow('+row+',3)">Snooze 3d</button>'+
        '<button onclick="setApproved('+row+','+(!l.approved)+')">'+(l.approved?'⬛ Unapprove':'✓ Approve')+'</button>'+
        '<button class="primary" onclick="openScheduleModal('+row+',\''+esc(l.name||'Lead')+'\',\''+esc(l.company||'')+'\')">📅 Schedule</button>'+
        calBtn+
        '</div></div>';
    }).join('');
  }

  function loadActivities(){
    el('actList').innerHTML='<div class="empty">Loading…</div>';
    google.script.run
      .withSuccessHandler(function(acts){
        if(!acts||!acts.length){el('actList').innerHTML='<div class="empty">No activity recorded yet.</div>';return;}
        el('actList').innerHTML=acts.map(function(a){
          return'<div class="act-item">'+
            '<div class="act-ts">'+esc(a.ts||'')+'</div>'+
            '<div class="act-body">'+
            '<div class="act-action">'+esc(a.action||'')+'</div>'+
            (a.name?'<div class="act-who">'+esc(a.name||'')+'</div>':'')+
            (a.notes?'<div class="act-notes">'+esc(a.notes||'')+'</div>':'')+
            '</div></div>';
        }).join('');
      })
      .withFailureHandler(function(e){el('actList').innerHTML='<div class="empty">Error loading activities.</div>';})
      .uiGetActivities(null, 50);
  }

  // ── SCHEDULE MODAL ──
  function openScheduleModal(row, name, company){
    pendingScheduleRow=row;
    el('scheduleLeadName').textContent=(name||'Lead')+(company?' — '+company:'');
    el('schedDate').value='';el('schedTime').value='';el('schedDur').value='30';el('schedLoc').value='';
    el('schedErr').style.display='none';el('schedErr').textContent='';
    // Pre-fill today's date
    var today=new Date();
    el('schedDate').value=today.getFullYear()+'-'+String(today.getMonth()+1).padStart(2,'0')+'-'+String(today.getDate()).padStart(2,'0');
    el('scheduleModal').classList.add('open');
    setTimeout(function(){el('schedTime').focus();},100);
  }

  function submitSchedule(){
    var dateStr=el('schedDate').value.trim();
    var timeStr=el('schedTime').value.trim();
    var dur=parseInt(el('schedDur').value,10)||30;
    var loc=el('schedLoc').value.trim();
    if(!dateStr||!timeStr){el('schedErr').style.display='block';el('schedErr').textContent='Date and time are required.';return;}
    if(!pendingScheduleRow){el('schedErr').style.display='block';el('schedErr').textContent='No lead selected.';return;}
    el('schedErr').style.display='none';
    closeModal('scheduleModal');
    setBusy(true);
    google.script.run
      .withSuccessHandler(function(){setBusy(false);load();})
      .withFailureHandler(function(e){setBusy(false);showErr(e&&e.message?e.message:String(e));})
      .uiScheduleLead(pendingScheduleRow, dateStr, timeStr, dur, loc);
  }

  // ── QUICK ADD MODAL ──
  function openQuickAdd(){
    el('addName').value='';el('addEmail').value='';el('addPhone').value='';
    el('addCompany').value='';el('addDealValue').value='';el('addOwner').value='';el('addNotes').value='';
    el('addErr').style.display='none';el('addErr').textContent='';
    populateAddStatus();
    el('addLeadModal').classList.add('open');
    setTimeout(function(){el('addName').focus();},100);
  }

  function submitQuickAdd(){
    var name=el('addName').value.trim();
    var email=el('addEmail').value.trim();
    if(!name&&!email){el('addErr').style.display='block';el('addErr').textContent='Name or Email is required.';return;}
    el('addErr').style.display='none';
    var fields={name:name,email:email,phone:el('addPhone').value.trim(),
      company:el('addCompany').value.trim(),dealValue:el('addDealValue').value.trim(),
      status:el('addStatus').value,owner:el('addOwner').value.trim(),notes:el('addNotes').value.trim()};
    closeModal('addLeadModal');
    setBusy(true);
    google.script.run
      .withSuccessHandler(function(res){setBusy(false);load();})
      .withFailureHandler(function(e){setBusy(false);showErr(e&&e.message?e.message:String(e));})
      .uiQuickAddLead(fields);
  }

  function closeModal(id){el(id).classList.remove('open');}

  // Close modal on overlay click
  document.addEventListener('click',function(ev){
    if(ev.target.classList.contains('modal-overlay'))ev.target.classList.remove('open');
  });

  // ── SERVER ACTIONS ──
  function runReview(){if(busy)return;setBusy(true);clearErr();google.script.run.withSuccessHandler(function(){setBusy(false);load();}).withFailureHandler(function(e){setBusy(false);showErr(e&&e.message?e.message:String(e));}).uiRunReview();}
  function sendApproved(){if(busy)return;setBusy(true);clearErr();google.script.run.withSuccessHandler(function(){setBusy(false);load();}).withFailureHandler(function(e){setBusy(false);showErr(e&&e.message?e.message:String(e));}).uiSendApproved();}
  function markCalled(row){if(busy)return;setBusy(true);google.script.run.withSuccessHandler(function(){setBusy(false);load();}).withFailureHandler(function(e){setBusy(false);showErr(e&&e.message?e.message:String(e));}).uiMarkCalled(row);}
  function snoozeRow(row,days){if(busy)return;setBusy(true);google.script.run.withSuccessHandler(function(){setBusy(false);load();}).withFailureHandler(function(e){setBusy(false);showErr(e&&e.message?e.message:String(e));}).uiSnooze(row,days);}
  function setApproved(row,val){if(busy)return;setBusy(true);google.script.run.withSuccessHandler(function(){setBusy(false);load();}).withFailureHandler(function(e){setBusy(false);showErr(e&&e.message?e.message:String(e));}).uiUpdateLead(row,{'Approved to Send':val?'TRUE':''});}

  // ── AUTO-REFRESH ──
  function startAutoRefresh(){
    stopAutoRefresh();
    var toggle=el('autoRefreshToggle');var secInput=el('autoRefreshSec');
    if(!toggle||!secInput||!toggle.checked)return;
    var sec=parseInt(secInput.value,10);
    if(isNaN(sec)||sec<5)sec=5;if(sec>300)sec=300;
    secInput.value=sec;
    autoTimer=setInterval(function(){if(!busy&&viewMode!=='activity')load();},sec*1000);
  }
  function stopAutoRefresh(){if(autoTimer){clearInterval(autoTimer);autoTimer=null;}}

  el('autoRefreshToggle').addEventListener('change',startAutoRefresh);
  el('autoRefreshSec').addEventListener('change',startAutoRefresh);
  el('autoRefreshSec').addEventListener('keyup',function(e){if(e.key==='Enter')startAutoRefresh();});
  el('q').addEventListener('input',function(){if(STATE)render();});
  el('statusFilter').addEventListener('change',function(){if(STATE)render();});
  el('btnRefresh').addEventListener('click',load);
  el('btnRun').addEventListener('click',runReview);
  el('btnSend').addEventListener('click',sendApproved);

  // Keyboard shortcuts
  document.addEventListener('keydown',function(e){
    if(e.key==='Escape'){document.querySelectorAll('.modal-overlay.open').forEach(function(m){m.classList.remove('open');});}
  });

  load();
  startAutoRefresh();
</script></body></html>`;
}

// ═══════════════════════════════════════════════════════════════
//  SPREADSHEET POLISH  (Phase 5 — safe to re-run)
//  Autopilot → Style Spreadsheet
//  ‑ Freezes header row
//  ‑ Sets filter on Leads sheet
//  ‑ Data validation: Status, Approved to Send, Owner
//  ‑ Conditional formatting: Due, Overdue, Approved, Closed
//  ‑ Protect system columns (warn-only)
//  ‑ Hides system / auxiliary sheets
//  NEVER deletes data. Safe to call multiple times.
// ═══════════════════════════════════════════════════════════════

function styleSpreadsheet() {
  ensureSheets_();
  var ss     = SpreadsheetApp.getActive();
  var leads  = ss.getSheetByName(LFU.SHEET_LEADS);
  if (!leads) { SpreadsheetApp.getUi().alert("Leads sheet not found. Run Setup Wizard first."); return; }

  var lastCol = Math.max(leads.getLastColumn(), LFU.HEADERS_REQUIRED.length);
  var lastRow = Math.max(leads.getLastRow(), 2);
  var headers = leads.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h||"").trim(); });
  var hi = {};
  headers.forEach(function(h, i){ if(h) hi[h] = i+1; }); // 1-based column index

  // ── 1. Freeze header row
  leads.setFrozenRows(1);

  // ── 2. Filter view (set basic filter on entire data range)
  try {
    var existingFilter = leads.getFilter();
    if (!existingFilter) {
      leads.getRange(1, 1, lastRow, lastCol).createFilter();
    }
  } catch(e) { /* non-fatal */ }

  // ── 3. Header row styling
  var headerRange = leads.getRange(1, 1, 1, lastCol);
  headerRange
    .setBackground("#0d1117")
    .setFontColor("#7aa2ff")
    .setFontWeight("bold")
    .setFontSize(10);

  // ── 4. Data validation — Status column
  var cfg    = getAllSettings_();
  var stages = String(cfg.pipeline_stages || "NEW,CONTACTED,NURTURE,BOOKED,CLOSED")
    .split(",").map(function(s){ return s.trim().toUpperCase(); }).filter(Boolean);
  if (hi["Status"] && lastRow > 1) {
    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(stages, true)
      .setAllowInvalid(true)
      .build();
    leads.getRange(2, hi["Status"], lastRow-1, 1).setDataValidation(statusRule);
  }

  // ── 5. Data validation — Approved to Send column
  if (hi["Approved to Send"] && lastRow > 1) {
    var approvedRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["TRUE", "FALSE", ""], true)
      .setAllowInvalid(true)
      .build();
    leads.getRange(2, hi["Approved to Send"], lastRow-1, 1).setDataValidation(approvedRule);
  }

  // ── 6. Conditional formatting
  var rules = leads.getConditionalFormatRules();
  // Remove only our named rules (by comment is impossible, so we replace all CF)
  // Safe: CF is cosmetic only, no data lost
  var newRules = [];
  var dataRange = leads.getRange(2, 1, Math.max(lastRow-1, 1), lastCol);

  // Due flag = TRUE → amber tint
  if (hi["Due Flag"]) {
    var dueCol = leads.getRange(2, hi["Due Flag"], Math.max(lastRow-1,1), 1);
    newRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("TRUE")
      .setBackground("#2e2100")
      .setFontColor("#ffcc00")
      .setRanges([dueCol])
      .build());
  }

  // Approved to Send = TRUE → green tint
  if (hi["Approved to Send"]) {
    var apprCol = leads.getRange(2, hi["Approved to Send"], Math.max(lastRow-1,1), 1);
    newRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("TRUE")
      .setBackground("#0d2b14")
      .setFontColor("#33d17a")
      .setRanges([apprCol])
      .build());
  }

  // Draft Status = ERROR → red tint
  if (hi["Draft Status"]) {
    var dsCol = leads.getRange(2, hi["Draft Status"], Math.max(lastRow-1,1), 1);
    newRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("ERROR")
      .setBackground("#2b0b0b")
      .setFontColor("#ff5c5c")
      .setRanges([dsCol])
      .build());
    newRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("SENT")
      .setBackground("#0d2b14")
      .setFontColor("#33d17a")
      .setRanges([dsCol])
      .build());
    newRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("DRAFTED")
      .setBackground("#1a2040")
      .setFontColor("#7aa2ff")
      .setRanges([dsCol])
      .build());
  }

  // Status in CLOSED_STATUSES → dim the row (use a per-column approach on Status)
  if (hi["Status"]) {
    var stCol = leads.getRange(2, hi["Status"], Math.max(lastRow-1,1), 1);
    LFU.CLOSED_STATUSES.concat(LFU.LOST_STATUSES).forEach(function(s) {
      newRules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(s)
        .setBackground("#111823")
        .setFontColor("#4a5568")
        .setRanges([stCol])
        .build());
    });
  }

  leads.setConditionalFormatRules(newRules);

  // ── 7. Protect system columns (warn-only: no editors locked, just a note)
  var SYSTEM_COLS = ["Lead ID", "Due Flag", "Last Error", "Last Drafted", "Last Sent",
                     "Draft Status", "Priority Score", "Calendar Event ID", "Calendar Event Link"];
  SYSTEM_COLS.forEach(function(colName) {
    if (!hi[colName]) return;
    try {
      var colRange = leads.getRange(1, hi[colName], lastRow, 1);
      var prot = colRange.protect().setDescription("System column — " + colName);
      prot.setWarningOnly(true);
    } catch(e) { /* non-fatal if protection fails */ }
  });

  // ── 8. Auto-resize key columns
  var resizeCols = ["Name","Email","Phone","Company","Status","Follow Up Due","Deal Value","Notes"];
  resizeCols.forEach(function(n) {
    if (hi[n]) {
      try { leads.autoResizeColumn(hi[n]); } catch(e) {}
    }
  });

  // ── 9. Hide system/auxiliary sheets
  var HIDE_SHEETS = [
    LFU.SHEET_SYSOPS, LFU.SHEET_SOPS, LFU.SHEET_IMPROVE,
    LFU.SHEET_SOP_BUILDER, LFU.SHEET_SOP_TMPL, LFU.SHEET_RULES,
    LFU.SHEET_IMPORT, LFU.SHEET_INTAKE, LFU.SHEET_APPOINTMENTS
  ];
  HIDE_SHEETS.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (sh && !sh.isSheetHidden()) {
      try { sh.hideSheet(); } catch(e) {}
    }
  });

  // Make sure the Leads sheet is visible and active
  try { leads.showSheet(); ss.setActiveSheet(leads); } catch(e) {}

  logActivity_({ leadId:"", name:"", action:"Style Spreadsheet", notes:"Applied formatting and validation." });
  ss.toast("✅ Spreadsheet styled. Filters, validation, conditional formats, and system column protection applied.", "Autopilot", 6);
}
