/**
 * WebAppCode.gs — Apps Script Web App Layer (Track A)
 * ================================================
 * LFU Command Center v8.0.1 — Track A final
 *
 * ADD THIS FILE to your Apps Script project alongside Code.gs.
 * Do NOT paste this inside Code.gs — keep it as a separate file.
 *
 * Deploy → New deployment → Web app
 *   Execute as:   User accessing the web app  (recommended)
 *                 — OR — Me (if you want one shared session)
 *   Who has access: Only myself  (solo)
 *                   Anyone with Google account  (team)
 *
 * NOTE: All team members must have Viewer access to the Sheet or
 *       they will receive an authorization error when opening the
 *       web app. This is a Google security requirement, not a bug.
 */

// ═══════════════════════════════════════════════════════════════
//  ENTRYPOINT
// ═══════════════════════════════════════════════════════════════

/**
 * doGet — serves the Command Center web app.
 * Supports ?action=health for server-side probe diagnostics.
 *
 * IMPORTANT: If Code.gs already has a doGet function, you must
 * MERGE this logic into it. Apps Script only runs one doGet per
 * project — a second one silently wins or loses depending on load
 * order, which is undefined. The safest fix: rename the old doGet
 * in Code.gs to something else (e.g. doGet_legacy_) and keep only
 * this one.
 */
function doGet(e) {
  var p      = (e && e.parameter) ? e.parameter : {};
  var action = String(p.action || "").trim().toLowerCase();

  if (action) {
    return webAppRouteGet_(action, p);
  }

  return HtmlService
    .createHtmlOutputFromFile("CommandCenter")
    .setTitle("LFU Command Center")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function webAppRouteGet_(action, p) {
  try {
    if (action === "health") {
      return jsonRes_({ ok: true, data: webAppHealth_() });
    }
    return jsonRes_({ ok: false, error: { message: "Unknown action: " + action, code: "BAD_ACTION" } });
  } catch (err) {
    return jsonRes_({
      ok: false,
      error: { message: "Server error", code: "SERVER_ERROR" }
    });
  }
}

function jsonRes_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════════
//  SAFE WRAPPER
//  Wraps every google.script.run-callable function.
//  Returns { ok:true, data, debugId } on success.
//  Returns { ok:false, error:{ code, userMessage, debugId } } on failure.
//  UI shows userMessage only. Stack traces NEVER reach the client.
// ═══════════════════════════════════════════════════════════════

/**
 * safeUserMessage_(err) — derives a safe, sanitized message from a caught error.
 * Rules:
 *   - Use err.userMessage if set (explicit caller-set message)
 *   - Else use err.message if it looks like a user-readable string
 *   - Else fall back to generic
 *   - Strip newlines, cap to 180 chars
 *   - Never return a stack trace
 */
function safeUserMessage_(err) {
  var raw = (err && (err.userMessage || err.message)) || "Something went wrong. Please try again.";
  // Strip stack trace indicators
  var sanitized = String(raw)
    .replace(/\n[\s\S]*/m, "")   // cut at first newline (removes stack lines)
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 180);
  return sanitized || "Something went wrong. Please try again.";
}

/**
 * safe_(fn) — standard wrapper for all google.script.run-callable functions.
 *
 * Returns on success:  { ok:true, data, debugId }
 * Returns on failure:  { ok:false, error:{ code, userMessage, debugId } }
 *
 * The "technical" field is intentionally ABSENT from the return value.
 * Technical details are logged server-side only via console.error.
 * The client NEVER receives a stack trace or raw exception message.
 */
function safe_(fn) {
  var debugId = Utilities.getUuid();
  try {
    var result = fn();
    if (result && result.ok === false) return result;
    return { ok: true, data: result, debugId: debugId };
  } catch (err) {
    var technical = err && err.message ? err.message : String(err);
    // Server-side logging only — never sent to client
    console.error("[LFU " + debugId + "] " + technical);
    // PRIVACY: Activity sheet gets only debugId + action code.
    // Full technical detail stays in console.error (server-side only).
    safeLogActivity_({
      leadId: "", name: "", action: "ERROR",
      notes:  "debugId:" + debugId
    });
    return {
      ok: false,
      error: {
        code:        "SERVER_ERROR",
        userMessage: safeUserMessage_(err),
        debugId:     debugId
        // NOTE: no "technical" field — technical details stay server-side
      }
    };
  }
}

// ═══════════════════════════════════════════════════════════════
//  READ-ONLY STATE GETTER
//  Called by the 25-second refresh loop and focus events.
//  NEVER writes to the sheet — safe for repeated polling.
//
//  P0.4 FIX: ensureSheets_() and ensureLeadColumns_() removed from
//  this path. They write to the sheet and must only run via
//  Setup Wizard or Repair, not on every refresh tick.
//
//  P0.1 FIX: isQualifiedLeadRow_ is now defined in Code.gs.
// ═══════════════════════════════════════════════════════════════

function uiGetStateReadOnly() {
  return safe_(function () {
    var cfg      = getAllSettings_();
    var pipeline = uiGetPipelineConfig();
    var ss       = SpreadsheetApp.getActive();
    var sheet    = ss.getSheetByName(LFU.SHEET_LEADS);
    if (!sheet) {
      throw new Error("Leads sheet not found. Open the Autopilot menu → Setup Wizard to initialize the spreadsheet.");
    }

    var ref   = readLeads_(sheet);
    var today = startOfDay_(new Date());

    // Compute flags IN MEMORY ONLY — zero sheet writes
    ref.rows.forEach(function (lead) {
      addComputedDueFlag_(lead, today);
      addPriorityScore_(lead);
    });

    var tz    = Session.getScriptTimeZone();
    var fmtD  = function (v) {
      if (!v) return "";
      var d = asDate_(v);
      return d ? Utilities.formatDate(d, tz, "yyyy-MM-dd") : "";
    };
    var fmtDT = function (v) {
      if (!v) return "";
      var d = asDate_(v);
      return d ? Utilities.formatDate(d, tz, "yyyy-MM-dd HH:mm") : "";
    };

    var leads = ref.rows.filter(isQualifiedLeadRow_).map(function (r) {
      var name          = String(r["Name"]           || "").trim();
      var preferredName = String(r["Preferred Name"] || "").trim();
      var firstName     = getFirstName_(preferredName || name);
      var status        = String(r["Status"] || "NEW").trim().toUpperCase() || "NEW";
      var dueFlag       = toBool_(r["Due Flag"]);
      var draftStatus   = String(r["Draft Status"]   || "").trim().toUpperCase();
      var approved      = toBool_(r["Approved to Send"]);
      var lastError     = String(r["Last Error"]     || "").trim();
      var dealValue     = String(r["Deal Value"]     || "").trim();
      var expectedValue = String(r["Expected Value"] || "").trim();
      var probability   = String(r["Probability %"]  || "").trim();
      var company       = String(r["Company"]        || "").trim();
      return {
        row:          r.__row,
        leadId:       String(r["Lead ID"]        || "").trim(),
        name:         name,
        preferredName:preferredName,
        firstName:    firstName,
        status:       status,
        dueFlag:      dueFlag,
        draftStatus:  draftStatus,
        approved:     approved,
        phone:        String(r["Phone"]          || "").trim(),
        email:        String(r["Email"]          || "").trim(),
        company:      company,
        owner:        String(r["Owner"]          || "").trim(),
        dealValue:    dealValue,
        expectedValue:expectedValue,
        probability:  probability,
        templatePack: String(r["Template Pack"]  || "").trim(),
        lastError:    lastError,
        notes:        String(r["Notes"]          || "").trim(),
        source:       String(r["Intake Source"]  || "").trim(),
        nextAction:   String(r["Next Action"]    || "").trim(),
        priorityScore:parseInt(String(r["Priority Score"] || "0"), 10) || 0,
        draft:        String(r["Draft"]          || "").trim(),
        apptStart:    fmtDT(r["Appointment Start"]),
        calLink:      String(r["Calendar Event Link"] || "").trim(),
        followUpDue:  fmtD(r["Follow Up Due"]),
        lastContacted:fmtD(r["Last Contacted"]),
        textPreview:  renderTemplate_(
          String(pipeline.textTemplate || ""),
          { Name: name, Company: company, "Preferred Name": preferredName },
          cfg,
          firstName
        )
      };
    });

    var toNum  = function (s) { return parseFloat(String(s || "0").replace(/[^0-9.]/g, "")) || 0; };
    var closed = LFU.CLOSED_STATUSES.concat(LFU.LOST_STATUSES);
    var open   = leads.filter(function (l) { return closed.indexOf(l.status) === -1; });

    var pipelineValue    = open.reduce(function (sum, l) { return sum + toNum(l.dealValue); }, 0);
    var weightedPipeline = open.reduce(function (sum, l) {
      var dv = toNum(l.dealValue), ev = toNum(l.expectedValue);
      var prob = (ev && dv) ? ev / dv : toNum(l.probability) / 100;
      return sum + dv * prob;
    }, 0);
    var overdueValue = leads.filter(function (l) { return l.dueFlag; })
      .reduce(function (sum, l) { return sum + toNum(l.dealValue); }, 0);

    var kpis = {
      total:            leads.length,
      due:              leads.filter(function (l) { return l.dueFlag; }).length,
      drafted:          leads.filter(function (l) { return l.draftStatus === "DRAFTED"; }).length,
      approved:         leads.filter(function (l) { return l.approved; }).length,
      errors:           leads.filter(function (l) { return l.draftStatus === "ERROR" || l.lastError.length > 0; }).length,
      pipelineValue:    Math.round(pipelineValue),
      weightedPipeline: Math.round(weightedPipeline),
      overdueValue:     Math.round(overdueValue)
    };

    // Feature flags — read once per state fetch, used by UI to show/hide controls
    var features = {
      calendarEnabled: isCalendarEnabled_()
    };

    return { version: LFU.VERSION, kpis: kpis, leads: leads, pipeline: pipeline, features: features };
  });
}

// ═══════════════════════════════════════════════════════════════
//  HEALTH CHECK  (served at ?action=health)
//  Must not throw even if optional sheets are missing.
// ═══════════════════════════════════════════════════════════════

function webAppHealth_() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var req = [LFU.SHEET_LEADS, LFU.SHEET_SETTINGS, LFU.SHEET_TEMPLATES, LFU.SHEET_ACTIVITIES];
  var triggerList = [];
  try {
    triggerList = ScriptApp.getProjectTriggers().map(function (t) {
      return { handler: t.getHandlerFunction(), type: String(t.getEventType()) };
    });
  } catch (e) { /* non-fatal */ }

  return {
    version:         (typeof LFU !== "undefined" && LFU.VERSION) ? LFU.VERSION : "unknown",
    spreadsheetName: ss.getName(),
    sheets: {
      all:      ss.getSheets().map(function (s) { return s.getName(); }),
      required: req.map(function (n) { return { name: n, present: !!ss.getSheetByName(n) }; })
    },
    triggers:  triggerList,
    checkedAt: new Date().toISOString()
  };
}

// ═══════════════════════════════════════════════════════════════
//  APPROVE TOGGLE
//  P1 FIX: Cannot approve leads in CLOSED or BOOKED status.
// ═══════════════════════════════════════════════════════════════

function uiApproveToggle(row, approved) {
  return safe_(function () {
    if (!row) throw new Error("Missing row.");
    var sheet = SpreadsheetApp.getActive().getSheetByName(LFU.SHEET_LEADS);
    if (!sheet) throw new Error("Leads sheet not found.");
    var ref         = readLeads_(sheet);
    var headerIndex = ref.headerIndex;
    var rows        = ref.rows;
    var lead        = rows.filter(function (r) { return r.__row === row; })[0];

    // P1: Block approval for closed stages
    if (approved && lead) {
      var status = String(lead["Status"] || "").trim().toUpperCase();
      if (LFU.CLOSED_STATUSES.indexOf(status) !== -1) {
        throw new Error("Cannot approve a lead in status " + status + ". Move to an active stage first.");
      }
    }

    setCell_(sheet, row, headerIndex, "Approved to Send", approved ? "TRUE" : "");
    safeLogActivity_({
      leadId: lead ? lead["Lead ID"] : "",
      name:   lead ? lead["Name"]   : "",
      action: approved ? "Approved Draft" : "Unapproved Draft",
      notes:  "Row " + row
    });
    return { ok: true, row: row, approved: approved };
  });
}

// ═══════════════════════════════════════════════════════════════
//  DOPAMINE STATE  (stored per-user in UserProperties by date)
// ═══════════════════════════════════════════════════════════════

function uiGetDopamineState() {
  return safe_(function () {
    var props    = PropertiesService.getUserProperties();
    var tz       = Session.getScriptTimeZone();
    var today    = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    var lastDate = props.getProperty("LFU_DOPC_DATE")   || "";
    var streak   = parseInt(props.getProperty("LFU_DOPC_STREAK") || "0", 10) || 0;
    var count    = 0;

    if (lastDate === today) {
      count = parseInt(props.getProperty("LFU_DOPC_COUNT") || "0", 10) || 0;
    } else if (lastDate) {
      var yest = new Date();
      yest.setDate(yest.getDate() - 1);
      var yStr = Utilities.formatDate(yest, tz, "yyyy-MM-dd");
      if (lastDate !== yStr) streak = 0;
    }

    return { count: count, streak: streak, today: today, goal: 10 };
  });
}

function uiLogDopamineAction() {
  return safe_(function () {
    var props    = PropertiesService.getUserProperties();
    var tz       = Session.getScriptTimeZone();
    var today    = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    var lastDate = props.getProperty("LFU_DOPC_DATE")   || "";
    var streak   = parseInt(props.getProperty("LFU_DOPC_STREAK") || "0", 10) || 0;
    var count    = 0;

    if (lastDate === today) {
      count = parseInt(props.getProperty("LFU_DOPC_COUNT") || "0", 10) || 0;
    } else {
      var yest = new Date();
      yest.setDate(yest.getDate() - 1);
      var yStr = Utilities.formatDate(yest, tz, "yyyy-MM-dd");
      streak = (lastDate === yStr) ? streak + 1 : 1;
      count  = 0;
    }

    count++;
    props.setProperty("LFU_DOPC_DATE",   today);
    props.setProperty("LFU_DOPC_COUNT",  String(count));
    props.setProperty("LFU_DOPC_STREAK", String(streak));

    return { count: count, streak: streak, today: today, goal: 10 };
  });
}

// ═══════════════════════════════════════════════════════════════
//  REPAIR / RESET  (non-destructive only)
//  - NEVER deletes sheets, clears rows, or overwrites data.
//  - Only: creates missing required sheets + calls ensureSheets_.
//  - Returns a report of what changed.
// ═══════════════════════════════════════════════════════════════

function uiRepair() {
  return safe_(function () {
    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var required = [
      LFU.SHEET_LEADS, LFU.SHEET_SETTINGS, LFU.SHEET_TEMPLATES,
      LFU.SHEET_ACTIVITIES, LFU.SHEET_SYSOPS
    ];
    var created = [];
    required.forEach(function (name) {
      if (!ss.getSheetByName(name)) {
        ss.insertSheet(name);
        created.push(name);
      }
    });

    // ensureSheets_ sets up headers on newly created sheets — this is safe here
    // because Repair is an explicit user action, not a polling call.
    try { ensureSheets_(); } catch (e) { /* non-fatal */ }

    var triggers = [];
    try {
      triggers = ScriptApp.getProjectTriggers().map(function (t) {
        return t.getHandlerFunction();
      });
    } catch (e) { /* non-fatal */ }

    safeLogActivity_({ leadId: "", name: "", action: "Repair", notes: "Created: " + (created.join(", ") || "none") });
    return {
      ok:            true,
      createdSheets: created,
      triggersFound: triggers,
      message:       created.length
        ? "Created " + created.length + " missing sheet(s). All headers verified."
        : "All required sheets present. Headers verified. No changes made."
    };
  });
}

// ═══════════════════════════════════════════════════════════════
//  PROBE WEB APP URL  (server-side via UrlFetchApp)
// ═══════════════════════════════════════════════════════════════

function uiProbeWebApp(url) {
  return safe_(function () {
    var u = String(url || "").trim();
    if (!u) throw new Error("No URL saved. Paste your web app URL first.");
    if (u.indexOf("https://") !== 0) throw new Error("URL must start with https://");

    var probeUrl = u + (u.indexOf("?") !== -1 ? "&" : "?") + "action=health";
    var resp, code;
    try {
      resp = UrlFetchApp.fetch(probeUrl, { muteHttpExceptions: true, followRedirects: true });
      code = resp.getResponseCode();
    } catch (e) {
      return {
        status: "unreachable", code: 0,
        message: "Could not reach the URL. Check that it is correct and that the web app is deployed."
      };
    }

    if (code === 200) {
      var text = resp.getContentText();
      var json = null;
      try { json = JSON.parse(text); } catch (e) { /* ignore */ }
      return {
        status:  "ok",
        code:    code,
        message: "Web app is responding correctly.",
        version: (json && json.data && json.data.version) ? json.data.version : "?"
      };
    } else if (code === 401 || code === 403) {
      return {
        status:  "blocked",
        code:    code,
        message: "Probe blocked by access settings — that's expected if the web app requires Google sign-in. " +
                 "Open the URL directly in your browser to confirm it loads correctly."
      };
    } else {
      return {
        status:  "error",
        code:    code,
        message: "Unexpected HTTP " + code + ". Try re-deploying the web app (Deploy → Manage deployments → New version)."
      };
    }
  });
}

// ═══════════════════════════════════════════════════════════════
//  DEPLOYMENT HELPER DATA FUNCTIONS
// ═══════════════════════════════════════════════════════════════

function uiSaveWebAppUrl(url) {
  return safe_(function () {
    var u = String(url || "").trim();
    if (!u) throw new Error("URL is required.");
    PropertiesService.getScriptProperties().setProperty("LFU_WEBAPP_URL", u);
    return { ok: true };
  });
}

function uiGetWebAppUrl() {
  return safe_(function () {
    var url = PropertiesService.getScriptProperties().getProperty("LFU_WEBAPP_URL") || "";
    return { url: url, health: webAppHealth_() };
  });
}

/** Opens the Deployment Helper as a sidebar. Called from the Autopilot menu. */
function showDeploymentHelper() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService
      .createHtmlOutputFromFile("DeploymentHelper")
      .setTitle("LFU Deployment Helper")
  );
}

// ═══════════════════════════════════════════════════════════════
//  OPEN COMMAND CENTER  (from menu / as fallback)
// ═══════════════════════════════════════════════════════════════

/**
 * Opens the saved web app URL in a tiny dialog that immediately
 * redirects the browser to the external URL in a new tab.
 * Called from Autopilot → Open Command Center (Web App).
 */
function openCommandCenterWeb_() {
  var url = PropertiesService.getScriptProperties().getProperty("LFU_WEBAPP_URL") || "";
  if (!url) {
    SpreadsheetApp.getUi().alert(
      "No Web App URL Saved",
      "Use Autopilot → Deployment Helper to deploy the web app and save its URL.\n\n" +
      "Steps:\n1. Deployment Helper → Deploy → copy the /exec URL\n2. Paste it in the URL field → Save URL",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  var html = HtmlService.createHtmlOutput(
    '<script>window.open("' + url + '","_blank"); google.script.host.close();</script>'
  ).setWidth(1).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, "Opening…");
}

// ═══════════════════════════════════════════════════════════════
//  SAFE ACTION WRAPPERS
//  All google.script.run calls from CommandCenter.html route through
//  these. They guarantee the { ok, error:{ userMessage, debugId } }
//  contract even for functions in Code.gs that don't use safe_().
// ═══════════════════════════════════════════════════════════════

function uiRunReviewSafe() {
  var debugId = Utilities.getUuid();
  try {
    var r = uiRunReview();
    return { ok: true, data: r, debugId: debugId };
  } catch (err) {
    console.error("[" + debugId + "] uiRunReview: " + (err && err.message ? err.message : err));
    return {
      ok: false,
      error: {
        code:        "REVIEW_ERROR",
        userMessage: safeUserMessage_(err),
        debugId:     debugId
      }
    };
  }
}

function uiSendApprovedSafe() {
  var debugId = Utilities.getUuid();
  try {
    var r = uiSendApproved();
    return { ok: true, data: r, debugId: debugId };
  } catch (err) {
    console.error("[" + debugId + "] uiSendApproved: " + (err && err.message ? err.message : err));
    return {
      ok: false,
      error: {
        code:        "SEND_ERROR",
        userMessage: safeUserMessage_(err),
        debugId:     debugId
      }
    };
  }
}

function uiQuickAddLeadSafe(fields) {
  var debugId = Utilities.getUuid();
  try {
    var r = uiQuickAddLead(fields);
    return { ok: true, data: r, debugId: debugId };
  } catch (err) {
    console.error("[" + debugId + "] uiQuickAddLead: " + (err && err.message ? err.message : err));
    return {
      ok: false,
      error: {
        code:        "ADD_LEAD_ERROR",
        userMessage: safeUserMessage_(err),
        debugId:     debugId
      }
    };
  }
}

function uiSnoozeSafe(row, days) {
  var debugId = Utilities.getUuid();
  try {
    var r = uiSnooze(row, days);
    return { ok: true, data: r, debugId: debugId };
  } catch (err) {
    console.error("[" + debugId + "] uiSnooze: " + (err && err.message ? err.message : err));
    return {
      ok: false,
      error: {
        code:        "SNOOZE_ERROR",
        userMessage: safeUserMessage_(err),
        debugId:     debugId
      }
    };
  }
}

function uiMarkCalledSafe(row) {
  var debugId = Utilities.getUuid();
  try {
    var r = uiMarkCalled(row);
    return { ok: true, data: r, debugId: debugId };
  } catch (err) {
    console.error("[" + debugId + "] uiMarkCalled: " + (err && err.message ? err.message : err));
    return {
      ok: false,
      error: {
        code:        "MARK_CALLED_ERROR",
        userMessage: safeUserMessage_(err),
        debugId:     debugId
      }
    };
  }
}

function uiMoveLeadSafe(row, status) {
  var debugId = Utilities.getUuid();
  try {
    var r = uiMoveLead(row, status);
    return { ok: true, data: r, debugId: debugId };
  } catch (err) {
    console.error("[" + debugId + "] uiMoveLead: " + (err && err.message ? err.message : err));
    return {
      ok: false,
      error: {
        code:        "MOVE_LEAD_ERROR",
        userMessage: safeUserMessage_(err),
        debugId:     debugId
      }
    };
  }
}

// ═══════════════════════════════════════════════════════════════
//  SCHEDULE LEAD SAFE WRAPPER
//  Gated by isCalendarEnabled_(). Returns FEATURE_DISABLED if off.
//  If calendar IS enabled, delegates to uiScheduleLead in Code.gs.
//  CommandCenter.html calls this via google.script.run.
// ═══════════════════════════════════════════════════════════════

function uiScheduleLeadSafe(row, dateStr, timeStr, durationMin, location) {
  var debugId = Utilities.getUuid();
  try {
    if (!isCalendarEnabled_()) {
      return {
        ok: false,
        error: {
          code:        "FEATURE_DISABLED",
          userMessage: "Scheduling is disabled. Enable it in Settings (enable_calendar = TRUE) to use this feature.",
          debugId:     debugId
        }
      };
    }
    var r = uiScheduleLead(row, dateStr, timeStr, durationMin, location);
    return { ok: true, data: r, debugId: debugId };
  } catch (err) {
    console.error("[" + debugId + "] uiScheduleLeadSafe: " + (err && err.message ? err.message : err));
    return {
      ok: false,
      error: {
        code:        "SCHEDULE_ERROR",
        userMessage: safeUserMessage_(err),
        debugId:     debugId
      }
    };
  }
}
