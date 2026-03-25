// -----------------------------------------------------------------------
// ROUTING
// -----------------------------------------------------------------------
function doGet(e) {
  try {
  var APP_VERSION = "v215";
  function safeJsonString_(obj) {
    return JSON.stringify(obj === undefined ? null : obj)
      .replace(/<\//g, "<\\/")
      .replace(/\u2028/g, "\\u2028")
      .replace(/\u2029/g, "\\u2029");
  }
  function setBootstrapJson_(t) {
    var boot = {
      sessionEmail: t.sessionEmail || "",
      scriptUrl: t.scriptUrl || "",
      benchmarks: t.benchmarks || {},
      rainmakerImage: t.rainmakerImage || "",
      lastData: t.lastData || null,
      userStats: t.userStats || {},
      sfData: t.sfData || {},
      qbrCall: t.qbrCall || {},
      onBehalf: t.onBehalf || "",
      isAdmin: !!t.isAdmin
    };
    t.bootstrapJson = safeJsonString_(boot);
  }
  function blankBenchmarks_() {
    return {
      disco_5: 0, disco_20: 0,
      nbm_5: 0, nbm_20: 0,
      pipe_5: 0, pipe_20: 0,
      stg4_5: 0, stg4_20: 0
    };
  }
  // cleanEmail_ is defined globally in Config.gs
  var sessionEmail = "";
  try {
    sessionEmail = cleanEmail_(Session.getEffectiveUser().getEmail());
  } catch (eSessionEffective) {
    Logger.log("Session.getEffectiveUser failed in doGet: " + eSessionEffective.message);
  }
  if (!sessionEmail) {
    try {
      sessionEmail = cleanEmail_(Session.getActiveUser().getEmail());
    } catch (eSessionActive) {
      Logger.log("Session.getActiveUser failed in doGet: " + eSessionActive.message);
    }
  }
  var isAdmin = sessionEmail === String(ADMIN_EMAIL || "").toLowerCase();
  var isRep = REP_EMAILS.indexOf(String(sessionEmail || "").toLowerCase()) !== -1;
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : "";

  // Base URL routing:
  // - reps land on recap
  // - admin and everyone else land on summary
  if (page === "") {
    if (isRep && !((e && e.parameter && e.parameter.desktop) || "")) {
      var routerTemplate = HtmlService.createTemplateFromFile("rep_router");
      routerTemplate.scriptUrl = ScriptApp.getService().getUrl();
      return routerTemplate.evaluate()
        .setTitle("Opening Weekly Recap")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag("viewport", "width=device-width, initial-scale=1");
    }
    page = isRep ? "recap" : "summary";
  }

  if (page === "diag") {
    var activeUser = "";
    var effectiveUser = "";
    try { effectiveUser = cleanEmail_(Session.getEffectiveUser().getEmail()); } catch (_) {}
    if (!effectiveUser) {
      try { activeUser = cleanEmail_(Session.getActiveUser().getEmail()); } catch (_) {}
    }
    var params = "";
    try { params = JSON.stringify((e && e.parameter) ? e.parameter : {}); } catch (_) { params = "{}"; }
    var htmlDiag = ""
      + "<!doctype html><html><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'>"
      + "<title>Weekly Recap Diag</title>"
      + "<style>body{font-family:Arial,sans-serif;background:#0f172a;color:#e2e8f0;padding:20px}code{background:#1e293b;padding:2px 6px;border-radius:6px}pre{white-space:pre-wrap;background:#111827;border:1px solid #334155;padding:10px;border-radius:8px}</style>"
      + "</head><body><h2>Weekly Recap Diagnostics</h2>"
      + "<p><b>App version:</b> <code>" + APP_VERSION + "</code></p>"
      + "<p><b>Admin email:</b> <code>" + String(ADMIN_EMAIL || "") + "</code></p>"
      + "<p><b>Session active user:</b> <code>" + activeUser + "</code></p>"
      + "<p><b>Session effective user:</b> <code>" + effectiveUser + "</code></p>"
      + "<p><b>Resolved sessionEmail:</b> <code>" + sessionEmail + "</code></p>"
      + "<p><b>Is admin:</b> <code>" + String(isAdmin) + "</code></p>"
      + "<p><b>Request params:</b></p><pre>" + params + "</pre>"
      + "</body></html>";
    return HtmlService.createHtmlOutput(htmlDiag)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (!isRep && page === "summary") {
    var summaryTemplate = HtmlService.createTemplateFromFile("summary");
    summaryTemplate.summaryData = JSON.stringify(getExecutiveSummaryData());
    summaryTemplate.scriptUrl = ScriptApp.getService().getUrl();
    summaryTemplate.sessionEmail = sessionEmail;
    summaryTemplate.isAdmin = isAdmin;
    return summaryTemplate.evaluate()
      .setTitle("NorthCentral Enterprise Weekly Summary")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  if (page === "poc-legacy" || (isAdmin && page === "poc-legacy")) {
    var pocTemplate = HtmlService.createTemplateFromFile("dashboard");
    pocTemplate.dashboardData = JSON.stringify([]);
    pocTemplate.benchmarks = "{}";
    pocTemplate.quotaMetrics = "{}";
    pocTemplate.pocDataJson = safeJsonString_(getPocData());
    pocTemplate.pocRefreshStatusJson = "null";
    pocTemplate.qaRepOptionsJson = safeJsonString_(getQaRepOptions_());
    pocTemplate.scriptUrl = ScriptApp.getService().getUrl();
    pocTemplate.sessionEmail = sessionEmail;
    pocTemplate.isAdmin = isAdmin;
    pocTemplate.initialPage = "poc";
    return pocTemplate.evaluate()
      .setTitle("POC Command Center")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  if (page === "poc" || (isAdmin && page === "poc")) {
    var pocPageTemplate = HtmlService.createTemplateFromFile("poc");
    pocPageTemplate.scriptUrl = ScriptApp.getService().getUrl();
    pocPageTemplate.sessionEmail = sessionEmail;
    pocPageTemplate.isAdmin = isAdmin;
    pocPageTemplate.pocDataJson = safeJsonString_(getPocDashboardSnapshot_());
    return pocPageTemplate.evaluate()
      .setTitle("POC Command Center")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  if (page === "recap-mobile") {
    var recapMobileTemplate = HtmlService.createTemplateFromFile("recap_mobile");
    var mobileTargetEmail = cleanEmail_(sessionEmail);
    var mobileOnBehalf = "";
    if (isAdmin && e && e.parameter && e.parameter.rep) {
      mobileTargetEmail = cleanEmail_(e.parameter.rep);
      mobileOnBehalf = mobileTargetEmail;
    } else if (isAdmin) {
      var mobileProps = PropertiesService.getScriptProperties();
      var mobilePending = mobileProps.getProperty("adminEdit_" + sessionEmail);
      if (mobilePending) {
        mobileTargetEmail = cleanEmail_(mobilePending);
        mobileOnBehalf = mobileTargetEmail;
        mobileProps.deleteProperty("adminEdit_" + sessionEmail);
      }
    }
    recapMobileTemplate.sessionEmail = sessionEmail;
    recapMobileTemplate.scriptUrl = ScriptApp.getService().getUrl();
    try {
      recapMobileTemplate.benchmarks = getBenchmarks();
    } catch (eBenchRecapMobile) {
      Logger.log("getBenchmarks failed for recap-mobile " + sessionEmail + ": " + eBenchRecapMobile.message);
      recapMobileTemplate.benchmarks = blankBenchmarks_();
    }
    recapMobileTemplate.rainmakerImage = "";
    try {
      recapMobileTemplate.lastData = getLastDataForUser(mobileTargetEmail);
    } catch (eLastRecapMobile) {
      Logger.log("getLastDataForUser failed for recap-mobile " + mobileTargetEmail + ": " + eLastRecapMobile.message);
      recapMobileTemplate.lastData = null;
    }
    recapMobileTemplate.lastDataJson = safeJsonString_(recapMobileTemplate.lastData);
    try { recapMobileTemplate.userStats = getUserStats(mobileTargetEmail); } catch (eStatsRecapMobile) { Logger.log("getUserStats failed for recap-mobile " + mobileTargetEmail + ": " + eStatsRecapMobile.message); recapMobileTemplate.userStats = {}; }
    try { recapMobileTemplate.sfData = getSFDataForUser(mobileTargetEmail); } catch (eSfRecapMobile) { Logger.log("getSFDataForUser failed for recap-mobile " + mobileTargetEmail + ": " + eSfRecapMobile.message); recapMobileTemplate.sfData = {}; }
    try { recapMobileTemplate.qbrCall = getQBRCall(mobileTargetEmail); } catch (eQbrRecapMobile) { Logger.log("getQBRCall failed for recap-mobile " + mobileTargetEmail + ": " + eQbrRecapMobile.message); recapMobileTemplate.qbrCall = {}; }
    recapMobileTemplate.onBehalf = mobileOnBehalf;
    recapMobileTemplate.isAdmin = isAdmin;
    setBootstrapJson_(recapMobileTemplate);
    return recapMobileTemplate.evaluate()
      .setTitle(mobileOnBehalf ? "Editing: " + mobileOnBehalf : "Q4 Weekly Recap")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  if (page === "recap") {
    var recapTemplate = HtmlService.createTemplateFromFile("index");
    var targetEmail = cleanEmail_(sessionEmail);
    var onBehalf = "";
    if (isAdmin && e && e.parameter && e.parameter.rep) {
      targetEmail = cleanEmail_(e.parameter.rep);
      onBehalf = targetEmail;
    } else if (isAdmin) {
      var props = PropertiesService.getScriptProperties();
      var pending = props.getProperty("adminEdit_" + sessionEmail);
      if (pending) {
        targetEmail = cleanEmail_(pending);
        onBehalf = targetEmail;
        props.deleteProperty("adminEdit_" + sessionEmail);
      }
    }
    recapTemplate.sessionEmail = sessionEmail;
    recapTemplate.scriptUrl = ScriptApp.getService().getUrl();
    try {
      recapTemplate.benchmarks = getBenchmarks();
    } catch (eBenchRecap) {
      Logger.log("getBenchmarks failed for recap " + sessionEmail + ": " + eBenchRecap.message);
      recapTemplate.benchmarks = blankBenchmarks_();
    }
    recapTemplate.rainmakerImage = "";  // legacy field, kept for template compat
    try {
      recapTemplate.lastData = getLastDataForUser(targetEmail);
    } catch (eLastRecap) {
      Logger.log("getLastDataForUser failed for recap " + targetEmail + ": " + eLastRecap.message);
      recapTemplate.lastData = null;
    }
    recapTemplate.lastDataJson = safeJsonString_(recapTemplate.lastData);
    try { recapTemplate.userStats = getUserStats(targetEmail); } catch (eStatsRecap) { Logger.log("getUserStats failed for recap " + targetEmail + ": " + eStatsRecap.message); recapTemplate.userStats = {}; }
    try { recapTemplate.sfData = getSFDataForUser(targetEmail); } catch (eSfRecap) { Logger.log("getSFDataForUser failed for recap " + targetEmail + ": " + eSfRecap.message); recapTemplate.sfData = {}; }
    try { recapTemplate.qbrCall = getQBRCall(targetEmail); } catch (eQbrRecap) { Logger.log("getQBRCall failed for recap " + targetEmail + ": " + eQbrRecap.message); recapTemplate.qbrCall = {}; }
    recapTemplate.onBehalf = onBehalf;
    recapTemplate.isAdmin = isAdmin;
    setBootstrapJson_(recapTemplate);
    return recapTemplate.evaluate()
      .setTitle(onBehalf ? "Editing: " + onBehalf : "Q4 Weekly Recap")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  if (page === "dashboard" && !isAdmin) {
    return HtmlService.createHtmlOutput("<p style='font-family:Arial,sans-serif;padding:20px;max-width:700px;margin:auto;'>Dashboard is admin-only. Please go to <a href='?page=summary'>Summary</a>.</p>")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  if (isAdmin && page === "dashboard") {
    var dashTemplate = HtmlService.createTemplateFromFile("dashboard");
    dashTemplate.dashboardData = JSON.stringify(getTeamDashboardData());
    dashTemplate.benchmarks = "{}";
    dashTemplate.quotaMetrics = JSON.stringify(getQuotaMetricsForDashboard_());
    dashTemplate.qaRepOptionsJson = safeJsonString_(getQaRepOptions_());
    dashTemplate.scriptUrl = ScriptApp.getService().getUrl();
    dashTemplate.sessionEmail = sessionEmail;
    dashTemplate.isAdmin = isAdmin;
    dashTemplate.initialPage = "dashboard";
    return dashTemplate.evaluate()
      .setTitle("Team Weekly Dashboard")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  if (isAdmin && page === "rep") {
    var repEmail = (e && e.parameter && e.parameter.email) ? cleanEmail_(e.parameter.email) : "";
    if (!repEmail) {
      return HtmlService.createHtmlOutput("<p>Missing rep email. Go back and select a rep.</p>")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    var repTemplate = HtmlService.createTemplateFromFile("index");
    repTemplate.sessionEmail = sessionEmail;
    repTemplate.scriptUrl = ScriptApp.getService().getUrl();
    try {
      repTemplate.benchmarks = getBenchmarks();
    } catch (eBenchRep) {
      Logger.log("getBenchmarks failed for rep view " + repEmail + ": " + eBenchRep.message);
      repTemplate.benchmarks = blankBenchmarks_();
    }
    repTemplate.rainmakerImage = "";
    try {
      repTemplate.lastData = getLastDataForUser(repEmail);
    } catch (eLastRep) {
      Logger.log("getLastDataForUser failed for rep view " + repEmail + ": " + eLastRep.message);
      repTemplate.lastData = null;
    }
    repTemplate.lastDataJson = safeJsonString_(repTemplate.lastData);
    try { repTemplate.userStats = getUserStats(repEmail); } catch (eStatsRep) { Logger.log("getUserStats failed for rep view " + repEmail + ": " + eStatsRep.message); repTemplate.userStats = {}; }
    try { repTemplate.sfData = getSFDataForUser(repEmail); } catch (eSfRep) { Logger.log("getSFDataForUser failed for rep view " + repEmail + ": " + eSfRep.message); repTemplate.sfData = {}; }
    try { repTemplate.qbrCall = getQBRCall(repEmail); } catch (eQbrRep) { Logger.log("getQBRCall failed for rep view " + repEmail + ": " + eQbrRep.message); repTemplate.qbrCall = {}; }
    repTemplate.onBehalf = repEmail;
    repTemplate.isAdmin = true;
    setBootstrapJson_(repTemplate);
    return repTemplate.evaluate()
      .setTitle("Edit Recap: " + repEmail)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  var template = HtmlService.createTemplateFromFile("index");
  template.sessionEmail = sessionEmail;
  template.scriptUrl = ScriptApp.getService().getUrl();
  try {
    template.benchmarks = getBenchmarks();
  } catch (eBenchDefault) {
    Logger.log("getBenchmarks failed for rep " + sessionEmail + ": " + eBenchDefault.message);
    template.benchmarks = blankBenchmarks_();
  }
  template.rainmakerImage = "";
  try {
    template.lastData = getLastDataForUser(sessionEmail);
  } catch (eLastDefault) {
    Logger.log("getLastDataForUser failed for default " + sessionEmail + ": " + eLastDefault.message);
    template.lastData = null;
  }
  template.lastDataJson = safeJsonString_(template.lastData);
  try { template.userStats = getUserStats(sessionEmail); } catch (eStatsDefault) { Logger.log("getUserStats failed for default " + sessionEmail + ": " + eStatsDefault.message); template.userStats = {}; }
  try { template.sfData = getSFDataForUser(sessionEmail); } catch (eSfDefault) { Logger.log("getSFDataForUser failed for default " + sessionEmail + ": " + eSfDefault.message); template.sfData = {}; }
  try { template.qbrCall = getQBRCall(sessionEmail); } catch (eQbrDefault) { Logger.log("getQBRCall failed for default " + sessionEmail + ": " + eQbrDefault.message); template.qbrCall = {}; }
  template.onBehalf = "";
  template.isAdmin = isAdmin;
  setBootstrapJson_(template);
  return template.evaluate()
    .setTitle("Q4 Weekly Recap")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
  } catch (err) {
    var msg = String((err && err.message) || err || "Unknown doGet error");
    var stack = String((err && err.stack) || "");
    var req = "";
    try {
      req = JSON.stringify((e && e.parameter) ? e.parameter : {});
    } catch (jsonErr) {
      req = "{}";
    }
    Logger.log("FATAL doGet: " + msg + "\n" + stack + "\nparams=" + req);
    function esc_(s) {
      return String(s || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
    }
    var html = ''
      + '<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
      + '<title>Weekly Recap Error</title>'
      + '<style>body{font-family:Arial,sans-serif;background:#111827;color:#f9fafb;padding:20px}h1{font-size:20px}pre{white-space:pre-wrap;background:#1f2937;padding:12px;border-radius:8px;border:1px solid #374151}</style>'
      + '</head><body><h1>Weekly Recap failed to load</h1>'
      + '<p>Please screenshot this page and send to Billy/Codex.</p>'
      + '<p><strong>Error:</strong> ' + esc_(msg) + '</p>'
      + '<p><strong>Params:</strong> ' + esc_(req) + '</p>'
      + '<pre>' + esc_(stack) + '</pre>'
      + '</body></html>';
    return HtmlService.createHtmlOutput(html)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function setAdminEditTarget(email) {
  var caller = "";
  try { caller = String(Session.getEffectiveUser().getEmail() || "").toLowerCase().trim(); } catch (_) {}
  if (!caller) {
    try { caller = String(Session.getActiveUser().getEmail() || "").toLowerCase().trim(); } catch (_) {}
  }
  if (caller !== String(ADMIN_EMAIL || "").toLowerCase().trim()) return;
  PropertiesService.getScriptProperties().setProperty(
    "adminEdit_" + caller, String(email || "").toLowerCase().trim()
  );
}

function getResolvedSessionEmail() {
  // cleanEmail_ is defined globally in Config.gs
  var email = "";
  try { email = cleanEmail_(Session.getEffectiveUser().getEmail()); } catch (_) {}
  if (!email) {
    try { email = cleanEmail_(Session.getActiveUser().getEmail()); } catch (_) {}
  }
  return email;
}

function testStatsDebug() {
  var ss = SpreadsheetApp.openById(TRACKER_SHEET_ID);
  var results = {};
  var sheetNames = [DISCO_SHEET_NAME, NBM_SHEET_NAME, PIPELINE_SHEET_NAME, STAGE4_SHEET_NAME];
  sheetNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) { results[name] = "SHEET NOT FOUND"; return; }
    var data = sheet.getRange(1, 1, Math.min(5, sheet.getLastRow()), 3).getValues();
    results[name] = data;
  });
  Logger.log(JSON.stringify(results, null, 2));
}

function testDashboard() {
  var start = new Date();
  var result = getTeamDashboardData();
  var elapsed = (new Date() - start) / 1000;
  Logger.log("Time: " + elapsed + "s, Reps: " + result.length);
}

function testDoGet2() {
  var start = new Date();
  var bd = getTeamDashboardData();
  var t1 = (new Date() - start) / 1000;
  Logger.log("dashboard data: " + t1 + "s");
  var start2 = new Date();
  var bm = getBenchmarks();
  var t2 = (new Date() - start2) / 1000;
  Logger.log("benchmarks: " + t2 + "s");
}
