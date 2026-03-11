// -----------------------------------------------------------------------
// CONFIG — global constants and shared helpers
// -----------------------------------------------------------------------
var ADMIN_EMAIL         = "billy.schuett@glean.com";
var MAIN_SHEET_NAME    = "Main Sheet";
var DEALS_SHEET_NAME    = "Billy WR - Total Pipe SF-data";
var PIPELINE_SHEET_NAME = "Billy WR - Total Ent Stage 2 FY";
var STAGE4_SHEET_NAME   = "Billy WR - Ent Stage 4+ CurrentQ/NQ";
var DISCO_SHEET_NAME    = "Billy - WR Disco Meetings CQ";
var NBM_SHEET_NAME      = "Billy WR - Enterprise - NBM CQ";
var POC_ACCOUNTS_SHEET  = "POC Accounts";
var POC_ACTIONS_SHEET   = "POC Actions";

var ACCT_BASES = [9, 15, 21, 27, 33, 39];

var DEALS_COL_OWNER    = 11;
var DEALS_COL_AMOUNT   = 4;
var DEALS_COL_CLOSE    = 7;
var DEALS_COL_FORECAST = 10;
var DEALS_COL_ACCOUNT  = 1;
var NBM_COL_KEY        = 1;
var NBM_MIN_AMOUNT     = 100000;

var NAME_MAP = {
  "niko.argaez@glean.com":      "Niko Argaez",
  "taylor.lundberg@glean.com":  "Taylor Lundberg",
  "james.demory@glean.com":     "James DeMory",
  "johnny.hatch@glean.com":     "Johnny Hatch",
  "melissa.richards@glean.com": "Melissa Richards",
  "daniel.broderick@glean.com": "Daniel Broderick",
  "billy.schuett@glean.com":    "Billy Schuett"
};

var REP_EMAILS = [
  "niko.argaez@glean.com",
  "taylor.lundberg@glean.com",
  "james.demory@glean.com",
  "johnny.hatch@glean.com",
  "melissa.richards@glean.com",
  "daniel.broderick@glean.com"
];

var _repStatsCache_ = null;

var POC_ACCOUNTS_HEADERS = ["Account","Pilot Type","Status","Stage","Owner","SE","Champion","Pilot Start","Pilot End","Competitor","SFDC URL"];
var POC_ACTIONS_HEADERS  = ["Account","Action Item","RAG","Signal Source","Notes","Last Updated","Updated By"];

var RECAP_HEADERS = [
  "Timestamp","Email","Week Of",
  "Pulse Score","Pulse Reason",
  "RM Disco","RM NBM","RM Opp ARR","RM Target Accts",
  "Acct1 Name","Acct1 Stage","Acct1 ARR","Acct1 Engage","Acct1 Move","Acct1 Next",
  "Acct2 Name","Acct2 Stage","Acct2 ARR","Acct2 Engage","Acct2 Move","Acct2 Next",
  "Acct3 Name","Acct3 Stage","Acct3 ARR","Acct3 Engage","Acct3 Move","Acct3 Next",
  "Acct4 Name","Acct4 Stage","Acct4 ARR","Acct4 Engage","Acct4 Move","Acct4 Next",
  "Acct5 Name","Acct5 Stage","Acct5 ARR","Acct5 Engage","Acct5 Move","Acct5 Next",
  "Acct6 Name","Acct6 Stage","Acct6 ARR","Acct6 Engage","Acct6 Move","Acct6 Next",
  "Forecast Note","Commit","Most Likely","Best Case","NQ Commit",
  "Goal","Risk","Ask",
  "Engagement Changes","Forecast Changes","SF Accurate"
];

// -----------------------------------------------------------------------
// HELPERS
// -----------------------------------------------------------------------
function _propOrDefault_(key, fallback) {
  var val = PropertiesService.getScriptProperties().getProperty(key);
  return val !== null ? val : fallback;
}

function getRecapSheet_() {
  var ss = SpreadsheetApp.openById("1tK7hslM--NY0fU6z7zvDwjw3K63DOTayAyQ0KE20J14");
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[0];
  return sheet;
}

function getUserNameFromEmail(email) {
  return NAME_MAP[email.toLowerCase().trim()] || null;
}

function normalizeEngagement(v) {
  if (!v) return "";
  var s = v.toString().trim().toLowerCase();
  if (s === "green")  return "Green";
  if (s === "yellow") return "Yellow";
  if (s === "red")    return "Red";
  return v.toString().trim();
}

function parseMoneyish_(val) {
  if (val === null || val === undefined || val === "") return "";
  if (typeof val === "number") return val;
  var s = String(val).toLowerCase().trim();
  if (!s) return "";
  var n = parseFloat(s.replace(/[^0-9.]/g, "")) || 0;
  if (s.indexOf("m") !== -1) n = n * 1000000;
  else if (s.indexOf("k") !== -1) n = n * 1000;
  return n ? n : "";
}

function normalizeRainmakerOpp_(rawVal) {
  var raw = (rawVal === null || rawVal === undefined) ? "" : String(rawVal).toLowerCase().trim();
  var n = parseMoneyish_(rawVal);
  if (!(n > 0)) return "";
  if (!/[mk]/.test(raw) && n < 1000) n = n * 1000;
  return Math.min(Math.round(n), 9999999);
}

function normalizeMoneyThousandsCapped_(rawVal) {
  var raw = (rawVal === null || rawVal === undefined) ? "" : String(rawVal).toLowerCase().trim();
  var n = parseMoneyish_(rawVal);
  if (!(n > 0)) return "";
  if (!/[mk]/.test(raw) && n < 1000) n = n * 1000;
  return Math.min(Math.round(n), 9999999);
}

function normalizeForecastFields_(formObject) {
  ["commit","likely","upside","nq_commit"].forEach(function(f) {
    var n = normalizeMoneyThousandsCapped_(formObject[f]);
    if (n === "") return;
    formObject[f] = n;
  });

  var rmOpp = normalizeRainmakerOpp_(formObject.rm_opp);
  formObject.rm_opp = (rmOpp === "") ? "" : rmOpp;

  for (var i = 1; i <= 6; i++) {
    var key = "acct" + i + "_arr";
    var arr = normalizeMoneyThousandsCapped_(formObject[key]);
    formObject[key] = (arr === "") ? "" : arr;
  }

  return formObject;
}

function getBlankStats() {
  return { disco_rank:"-", disco_amt:0, nbm_rank:"-", nbm_amt:0, pipe_rank:"-", pipe_amt:0, stg4_rank:"-", stg4_amt:0 };
}

function getRainmakerImageBase64() {
  return "";
}
