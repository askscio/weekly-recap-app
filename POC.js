// =======================================================================
// POC COMMAND CENTER — Backend Functions
// =======================================================================

// -----------------------------------------------------------------------
// POC Sheet helpers
// -----------------------------------------------------------------------
function getOrCreatePocAccountsSheet() {
  var ss = SpreadsheetApp.openById("1tK7hslM--NY0fU6z7zvDwjw3K63DOTayAyQ0KE20J14");
  var sheet = ss.getSheetByName(POC_ACCOUNTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(POC_ACCOUNTS_SHEET, ss.getSheets().length);
    sheet.appendRow(POC_ACCOUNTS_HEADERS);
    sheet.getRange(1, 1, 1, POC_ACCOUNTS_HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreatePocActionsSheet() {
  var ss = SpreadsheetApp.openById("1tK7hslM--NY0fU6z7zvDwjw3K63DOTayAyQ0KE20J14");
  var sheet = ss.getSheetByName(POC_ACTIONS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(POC_ACTIONS_SHEET, ss.getSheets().length);
    sheet.appendRow(POC_ACTIONS_HEADERS);
    sheet.getRange(1, 1, 1, POC_ACTIONS_HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// -----------------------------------------------------------------------
// POC CRUD
// -----------------------------------------------------------------------
function getPocData() {
  try {
    var accountsSheet = getOrCreatePocAccountsSheet();
    var actionsSheet = getOrCreatePocActionsSheet();

    var accountsData = accountsSheet.getDataRange().getValues();
    var accounts = [];
    for (var i = 1; i < accountsData.length; i++) {
      var r = accountsData[i];
      accounts.push({
        account: r[0], pilotType: r[1], status: r[2], stage: r[3],
        owner: r[4], se: r[5], champion: r[6],
        pilotStart: r[7] ? new Date(r[7]).toISOString() : "",
        pilotEnd: r[8] ? new Date(r[8]).toISOString() : "",
        competitor: r[9], sfdcUrl: r[10]
      });
    }

    var actionsData = actionsSheet.getDataRange().getValues();
    var actions = [];
    for (var j = 1; j < actionsData.length; j++) {
      var a = actionsData[j];
      actions.push({
        account: a[0], actionItem: a[1], rag: a[2], signalSource: a[3],
        notes: a[4],
        lastUpdated: a[5] ? new Date(a[5]).toISOString() : "",
        updatedBy: a[6]
      });
    }

    return { success: true, accounts: accounts, actions: actions };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function updateRagStatus(account, actionItem, newRag) {
  try {
    var sheet = getOrCreatePocActionsSheet();
    var data = sheet.getDataRange().getValues();
    var email = Session.getActiveUser().getEmail() || "unknown";
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === account && data[i][1] === actionItem) {
        sheet.getRange(i + 1, 3).setValue(newRag);
        sheet.getRange(i + 1, 6).setValue(new Date());
        sheet.getRange(i + 1, 7).setValue(email);
        return { success: true };
      }
    }
    return { success: false, error: "Action item not found" };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function updateActionNotes(account, actionItem, notes) {
  try {
    var sheet = getOrCreatePocActionsSheet();
    var data = sheet.getDataRange().getValues();
    var email = Session.getActiveUser().getEmail() || "unknown";
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === account && data[i][1] === actionItem) {
        sheet.getRange(i + 1, 5).setValue(notes);
        sheet.getRange(i + 1, 6).setValue(new Date());
        sheet.getRange(i + 1, 7).setValue(email);
        return { success: true };
      }
    }
    return { success: false, error: "Action item not found" };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function addActionItem(data) {
  try {
    var sheet = getOrCreatePocActionsSheet();
    var email = Session.getActiveUser().getEmail() || "unknown";
    sheet.appendRow([
      data.account, data.actionItem, data.rag || "Red",
      data.signalSource || "", data.notes || "",
      new Date(), email
    ]);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function deleteActionItem(account, actionItem) {
  try {
    var sheet = getOrCreatePocActionsSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === account && data[i][1] === actionItem) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: "Action item not found" };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getPocAccountNames_() {
  var sheet = getOrCreatePocAccountsSheet();
  var rows = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    var name = String(rows[i][0] || "").trim();
    if (name) out.push(name);
  }
  return out;
}

function normalizeGleanRag_(v) {
  var s = String(v || "").toLowerCase().trim();
  if (s === "green") return "Green";
  if (s === "yellow" || s === "amber") return "Yellow";
  if (s === "red") return "Red";
  return "Yellow";
}

function getGleanPocAgentConfig_() {
  var props = PropertiesService.getScriptProperties();
  return {
    endpoint: String(props.getProperty("GLEAN_POC_AGENT_ENDPOINT") || "").trim(),
    token: String(props.getProperty("GLEAN_POC_AGENT_TOKEN") || "").trim(),
    method: String(props.getProperty("GLEAN_POC_AGENT_METHOD") || "post").toLowerCase(),
    daysBack: parseInt(props.getProperty("GLEAN_POC_DAYS_BACK") || "3", 10) || 3,
    base: String(props.getProperty("GLEAN_BASE") || "").trim(),
    agentId: String(props.getProperty("GLEAN_AGENT_ID") || "").trim(),
    apiToken: String(props.getProperty("GLEAN_API_TOKEN") || "").trim(),
    actAs: String(props.getProperty("GLEAN_ACT_AS") || "").trim()
  };
}

function validateGleanPocAgentConfig_() {
  var cfg = getGleanPocAgentConfig_();
  function isPlaceholder_(v) {
    var s = String(v || "").toLowerCase().trim();
    return !s ||
      s.indexOf("your glean") !== -1 ||
      s.indexOf("api url") !== -1 ||
      s.indexOf("bearer token") !== -1 ||
      s.indexOf("<paste") !== -1;
  }

  // Mode A: direct Apps Script -> Glean runs/wait
  var hasDirect = !!cfg.base || !!cfg.agentId || !!cfg.apiToken;
  if (hasDirect) {
    if (isPlaceholder_(cfg.base)) throw new Error("Missing Script Property: GLEAN_BASE");
    if (isPlaceholder_(cfg.agentId)) throw new Error("Missing Script Property: GLEAN_AGENT_ID");
    if (isPlaceholder_(cfg.apiToken)) throw new Error("Missing Script Property: GLEAN_API_TOKEN");
    return cfg;
  }

  // Mode B: custom endpoint proxy
  if (isPlaceholder_(cfg.endpoint)) throw new Error("Missing Script Property: GLEAN_POC_AGENT_ENDPOINT");
  if (isPlaceholder_(cfg.token)) throw new Error("Missing Script Property: GLEAN_POC_AGENT_TOKEN");
  return cfg;
}

function buildGleanPocAgentPayload_() {
  return {
    mode: "poc_daily_refresh",
    accounts: getPocAccountNames_(),
    daysBack: getGleanPocAgentConfig_().daysBack,
    requestedAt: new Date().toISOString(),
    requestedBy: Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || ""
  };
}

function callGleanPocAgent_() {
  var cfg = validateGleanPocAgentConfig_();
  var payload = buildGleanPocAgentPayload_();
  var url = cfg.endpoint;
  var headers = {};
  if (cfg.base && cfg.agentId && cfg.apiToken) {
    url = String(cfg.base).replace(/\/+$/, "") + "/rest/api/v1/agents/runs/wait";
    payload = { agentId: cfg.agentId };
    headers.Authorization = "Bearer " + cfg.apiToken;
    if (cfg.actAs) headers["X-Glean-ActAs"] = cfg.actAs;
  } else {
    headers.Authorization = "Bearer " + cfg.token;
  }

  var options = {
    method: cfg.method || "post",
    muteHttpExceptions: true,
    contentType: "application/json",
    headers: headers,
    payload: JSON.stringify(payload)
  };
  var res = UrlFetchApp.fetch(url, options);
  var code = res.getResponseCode();
  var body = String(res.getContentText() || "");
  if (code < 200 || code >= 300) {
    throw new Error("Glean agent HTTP " + code + ": " + body.slice(0, 400));
  }
  var parsed;
  try {
    parsed = JSON.parse(body);
  } catch (e) {
    throw new Error("Glean agent response is not JSON.");
  }
  return parsed;
}

function parseGleanPocActions_(responseObj) {
  var src = responseObj || {};
  var arr = src.actions || src.results || [];
  if (!Array.isArray(arr)) throw new Error("Glean response must include an array at `actions`.");
  var out = [];
  for (var i = 0; i < arr.length; i++) {
    var a = arr[i] || {};
    var account = String(a.account || "").trim();
    var actionItem = String(a.actionItem || a.title || "").trim();
    if (!account || !actionItem) continue;
    out.push({
      account: account,
      actionItem: actionItem,
      rag: normalizeGleanRag_(a.rag),
      signalSource: String(a.signalSource || "GLEAN").trim() || "GLEAN",
      notes: String(a.notes || a.summary || "").trim(),
      updatedBy: String(a.updatedBy || "Glean Agent").trim() || "Glean Agent"
    });
  }
  return out;
}

function upsertPocActions_(actions) {
  var sheet = getOrCreatePocActionsSheet();
  var values = sheet.getDataRange().getValues();
  var idx = {}; // account|action
  for (var i = 1; i < values.length; i++) {
    var key = String(values[i][0] || "").toLowerCase().trim() + "|" + String(values[i][1] || "").toLowerCase().trim();
    if (key) idx[key] = i + 1; // sheet row
  }
  var now = new Date();
  var updated = 0;
  var inserted = 0;
  for (var j = 0; j < actions.length; j++) {
    var a = actions[j];
    var key2 = String(a.account || "").toLowerCase().trim() + "|" + String(a.actionItem || "").toLowerCase().trim();
    if (!key2) continue;
    if (idx[key2]) {
      var row = idx[key2];
      sheet.getRange(row, 3).setValue(a.rag);
      sheet.getRange(row, 4).setValue(a.signalSource || "GLEAN");
      sheet.getRange(row, 5).setValue(a.notes || "");
      sheet.getRange(row, 6).setValue(now);
      sheet.getRange(row, 7).setValue(a.updatedBy || "Glean Agent");
      updated++;
    } else {
      sheet.appendRow([
        a.account, a.actionItem, a.rag || "Yellow", a.signalSource || "GLEAN",
        a.notes || "", now, a.updatedBy || "Glean Agent"
      ]);
      inserted++;
    }
  }
  return { updated: updated, inserted: inserted };
}

function savePocGleanRefreshStatus_(statusObj) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty("POC_GLEAN_REFRESH_STATUS", JSON.stringify(statusObj || {}));
}

function getPocGleanRefreshStatus() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty("POC_GLEAN_REFRESH_STATUS") || "";
    if (!raw) return { success: true, status: null };
    var parsed = JSON.parse(raw);
    return { success: true, status: parsed };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function runPocGleanRefreshNow() {
  var startedAt = new Date();
  try {
    var caller = getCallerEmailSafe_();
    // Enforce admin check only when a caller email is available.
    // Time-based triggers may not expose active user context.
    if (caller && caller !== String(ADMIN_EMAIL || "").toLowerCase()) {
      var denied = {
        ok: false,
        ranAt: startedAt.toISOString(),
        error: "Admin access required."
      };
      savePocGleanRefreshStatus_(denied);
      return { success: false, error: "Admin access required." };
    }
    var responseObj = callGleanPocAgent_();
    var actions = [];
    var result = { inserted: 0, updated: 0 };
    var mode = "glean_writeback";
    // If caller returns structured actions, we can upsert here.
    if (responseObj && (Array.isArray(responseObj.actions) || Array.isArray(responseObj.results))) {
      actions = parseGleanPocActions_(responseObj);
      result = upsertPocActions_(actions);
      mode = "apps_script_upsert";
    }

    var actionsSheet = getOrCreatePocActionsSheet();
    var rowCount = Math.max(0, actionsSheet.getLastRow() - 1);
    var okStatus = {
      ok: true,
      ranAt: startedAt.toISOString(),
      mode: mode,
      fetchedActions: actions.length,
      inserted: result.inserted,
      updated: result.updated,
      rowsInPocActions: rowCount
    };
    savePocGleanRefreshStatus_(okStatus);
    return {
      success: true,
      mode: mode,
      fetchedActions: actions.length,
      inserted: result.inserted,
      updated: result.updated,
      rowsInPocActions: rowCount
    };
  } catch (err) {
    savePocGleanRefreshStatus_({
      ok: false,
      ranAt: startedAt.toISOString(),
      error: err.message
    });
    return { success: false, error: err.message };
  }
}

function installDailyPocGleanRefreshTrigger() {
  var caller = getCallerEmailSafe_();
  if (caller && caller !== String(ADMIN_EMAIL || "").toLowerCase()) {
    throw new Error("Admin access required.");
  }
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "runPocGleanRefreshNow") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("runPocGleanRefreshNow")
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
  return "Installed daily trigger for runPocGleanRefreshNow at ~6 AM script timezone.";
}

function getCallerEmailSafe_() {
  var email = "";
  try { email = String(Session.getActiveUser().getEmail() || ""); } catch (e1) {}
  if (!email) {
    try { email = String(Session.getEffectiveUser().getEmail() || ""); } catch (e2) {}
  }
  return String(email || "").trim().toLowerCase();
}

// -----------------------------------------------------------------------
// Seed POC Data — run once from the spreadsheet menu
// -----------------------------------------------------------------------
function seedPocData() {
  try {
    var ss = SpreadsheetApp.openById("1tK7hslM--NY0fU6z7zvDwjw3K63DOTayAyQ0KE20J14");

    var oldAccts = ss.getSheetByName(POC_ACCOUNTS_SHEET);
    if (oldAccts) {
      var hdr = oldAccts.getRange(1, 1, 1, oldAccts.getLastColumn()).getValues()[0];
      if (hdr[0] !== POC_ACCOUNTS_HEADERS[0]) ss.deleteSheet(oldAccts);
    }
    var oldActions = ss.getSheetByName(POC_ACTIONS_SHEET);
    if (oldActions) {
      var hdrA = oldActions.getRange(1, 1, 1, oldActions.getLastColumn()).getValues()[0];
      if (hdrA[0] !== POC_ACTIONS_HEADERS[0]) ss.deleteSheet(oldActions);
    }

    var accountsSheet = getOrCreatePocAccountsSheet();
    var existing = accountsSheet.getDataRange().getValues();
    if (existing.length > 1) {
      SpreadsheetApp.getUi().alert(
        "POC data already exists. Delete the rows in the \"POC Accounts\" and \"POC Actions\" sheet tabs first if you want to re-seed."
      );
      return;
    }

    var accounts = [
      ["A.T. Kearney", "Paid", "Active", "P6 Crawling/Indexing", "Taylor Lundberg", "Mike Koscak", "", "2/23/2026", "4/3/2026", "Microsoft Copilot", "https://glean.lightning.force.com/lightning/r/Opportunity/006PZ00000Q0hFhYAJ/view"],
      ["Bread Financial", "Free", "Upcoming", "P3 EB Alignment", "Johnny Hatch", "Ryan Krueger", "Joe Petrelli", "10/30/2025", "12/31/2025", "", "https://glean.lightning.force.com/lightning/r/Opportunity/006PZ00000MyI5SYAV/view"],
      ["Hudl", "Free", "Active", "P4 Deploy", "Johnny Hatch", "Dylan Kuntz", "Fabio Girolami", "2/20/2026", "4/10/2026", "Guru", "https://glean.lightning.force.com/lightning/r/Opportunity/006PZ00000QBiGTYA1/view"],
      ["Chamberlain", "Free", "Upcoming", "P1 Security/Arch", "Niko Argaez", "Ryan Krueger", "", "11/17/2025", "1/23/2026", "", "https://glean.lightning.force.com/lightning/r/Opportunity/006PZ00000N9fgDYAR/view"],
      ["Morningstar", "Free", "Upcoming", "P1 Security/Arch", "Taylor Lundberg", "Mike Koscak", "Samuel Castano", "3/16/2026", "6/14/2026", "", "https://glean.lightning.force.com/lightning/r/Opportunity/006PZ00000QFwW9YAL/view"]
    ];
    accountsSheet.getRange(2, 1, accounts.length, accounts[0].length).setValues(accounts);

    var actionsSheet = getOrCreatePocActionsSheet();
    var now = new Date();
    var user = "System Seed";

    var actions = [
      ["A.T. Kearney", "Instance Deployed & Live", "Green", "SLACK", "Mike Koscak confirmed instance initialized via magic links", now, user],
      ["A.T. Kearney", "User Onboarding Kickoff", "Green", "SFDC", "Onboarding scheduled for 3/4", now, user],
      ["A.T. Kearney", "Security Questionnaires Cleared", "Green", "SLACK", "Dylan: 4 left but status accepted, good to go", now, user],
      ["A.T. Kearney", "TOS Signed", "Green", "SLACK", "Taylor confirmed TOS signed in #help-legal", now, user],
      ["A.T. Kearney", "DITL + Exec Validation", "Green", "SFDC", "Both marked complete", now, user],
      ["A.T. Kearney", "PPT Creation Gap \u2014 Beta Request", "Yellow", "SLACK", "ROAD-898 requested, waiting Zane approval", now, user],
      ["A.T. Kearney", "Competitive Bake-Off (3 vendors)", "Yellow", "GONG", "Evaluating Glean + Copilot + LLM provider. CTO calls week of 3/13", now, user],
      ["A.T. Kearney", "Taylor Mat Leave Handoff (3/15)", "Yellow", "MGR", "Niko backing up. 11 days to ensure continuity mid-pilot", now, user],
      ["Bread Financial", "Customer Engagement (Gong)", "Red", "GONG", "ZERO Gong calls in 14 days. Total radio silence.", now, user],
      ["Bread Financial", "Internal Activity (Slack)", "Red", "SLACK", "No Slack activity found in past 2 weeks", now, user],
      ["Bread Financial", "Pilot Dates", "Red", "SFDC", "Original dates Oct 30 \u2013 Dec 31, 2025. 60+ days expired. No new dates.", now, user],
      ["Bread Financial", "Connectors & Corpus Defined", "Red", "SFDC", "No connectors, no corpus size in SFDC", now, user],
      ["Bread Financial", "EB Alignment on Scope", "Yellow", "SFDC", "Stage stuck at P3 since Oct 30. No progression in 125 days.", now, user],
      ["Bread Financial", "Security / Architecture Sign-Off", "Yellow", "SFDC", "Not started. Architecture Review = false.", now, user],
      ["Hudl", "Pilot Kicked Off", "Green", "GONG", "Kickoff call Feb 20 recorded. $0 Order Form approved.", now, user],
      ["Hudl", "Active Customer Engagement", "Green", "GONG", "Two calls in 7 days \u2014 Kickoff (2/20) + Checkpoint (2/27). Weekly cadence.", now, user],
      ["Hudl", "Connector Setup In Progress", "Green", "GONG", "Jira/Slack/Salesforce discussed. Daniel driving technical.", now, user],
      ["Hudl", "User Adoption (20 seats)", "Yellow", "SFDC", "Seat Utilization = 0%. Need first adoption data within 1 week.", now, user],
      ["Hudl", "Competitive Positioning vs Guru", "Yellow", "SFDC", "Primary competitor identified. Need differentiation narrative.", now, user],
      ["Hudl", "DITL Session", "Yellow", "SFDC", "Not yet scheduled. Good rapport \u2014 right time to propose.", now, user],
      ["Chamberlain", "Legal/TOS \u2014 NTT Miscommunication", "Red", "GONG", "Jeff: \"I was told you guys would be sending us an agreement months ago.\" NTT never sent TOS. Niko sending directly.", now, user],
      ["Chamberlain", "DPA Review Re-engagement", "Red", "GONG", "DPA review from Nov/Oct paused. Niko to pull prior work and resume.", now, user],
      ["Chamberlain", "Pilot Dates Expired 40+ Days", "Red", "SFDC", "Niko proposed: stood up by end of next week, 4-5 weeks testing.", now, user],
      ["Chamberlain", "Dept Lead 1 Buy-In", "Red", "MGR", "Not started \u2014 legal must clear first", now, user],
      ["Chamberlain", "Dept Lead 2 Buy-In", "Red", "MGR", "Not started \u2014 legal must clear first", now, user],
      ["Chamberlain", "Dept Lead 3 Buy-In", "Red", "MGR", "Not started \u2014 legal must clear first", now, user],
      ["Chamberlain", "Dept Lead 4 Buy-In", "Red", "MGR", "Not started \u2014 legal must clear first", now, user],
      ["Chamberlain", "Dept Lead 5 Buy-In", "Red", "MGR", "Not started \u2014 legal must clear first", now, user],
      ["Chamberlain", "Renewal Risk ($69.4K by 3/22)", "Red", "SFDC", "Renewal date Mar 22. Current value $69.4K.", now, user],
      ["Chamberlain", "NTT Ground Lead Alignment", "Yellow", "GONG", "Amitav Dash confirmed as NTT executor. Weekly calls proposed.", now, user],
      ["Chamberlain", "Expansion Path Strategy", "Red", "GONG", "Glean wants full Chamberlain (~3,000 seats) not just 50 NTT licenses.", now, user],
      ["Morningstar", "Legal Review", "Yellow", "GONG", "Samuel driving legal internally. Suggested removing word \"POC\" to help legal.", now, user],
      ["Morningstar", "ROI Measurement Strategy", "Yellow", "GONG", "CEO Kunal is data driven. Need clear ROI metrics before kickoff.", now, user],
      ["Morningstar", "AHEAD Partner Engagement", "Green", "GONG", "Josh (AHEAD) investing resources. POC proposal $39K (down from $54K).", now, user],
      ["Morningstar", "Kickoff Prep (3/16)", "Yellow", "SFDC", "Taylor starts mat leave 3/15 \u2014 one day before. Johnny must lead.", now, user],
      ["Morningstar", "Connectors & Technical Scope", "Red", "SFDC", "No connectors, corpus, SSO, or people data source defined. Must lock by 3/16.", now, user],
      ["Morningstar", "AI Enablement Team Formed", "Yellow", "GONG", "Todd Dernity (Sr. Manager) now helping drive adoption. Positive signal.", now, user],
      ["Morningstar", "Morningstar Pace Risk", "Yellow", "GONG", "Josh: \"Morningstar prizes poorly at first sign of dollar figure.\" Slow pace.", now, user]
    ];
    actionsSheet.getRange(2, 1, actions.length, actions[0].length).setValues(actions);

    SpreadsheetApp.getUi().alert(
      "Done! POC data seeded. You should see 'POC Accounts' and 'POC Actions' tabs at the bottom of your spreadsheet. Open your dashboard and click the POC Command Center tab."
    );
  } catch (err) {
    SpreadsheetApp.getUi().alert("Error seeding POC data: " + err.message);
  }
}
