// -----------------------------------------------------------------------
// CONFIG — global constants and shared helpers
// -----------------------------------------------------------------------
var ADMIN_EMAIL         = "billy.schuett@glean.com";
var TRACKER_SHEET_ID    = "1tK7hslM--NY0fU6z7zvDwjw3K63DOTayAyQ0KE20J14";
var MAIN_SHEET_NAME    = "Main Sheet";
var SUBMISSIONS_SHEET_NAME = MAIN_SHEET_NAME;
var DEALS_SHEET_NAME    = "Billy WR - Total Pipe SF-data";
var PIPELINE_SHEET_NAME = "Billy WR - Total Ent Stage 2 FY";
var STAGE4_SHEET_NAME   = "Billy WR - Ent Stage 4+ CurrentQ/NQ";
var DISCO_SHEET_NAME    = "Billy - WR Disco Meetings CQ";
var NBM_SHEET_NAME      = "Billy WR - Enterprise - NBM CQ";
var POC_ACCOUNTS_SHEET  = "POC Accounts";
var POC_ACTIONS_SHEET   = "POC Actions";
var SUMMARY_ADMIN_SHEET = "Summary_Admin";
var QUOTA_CONFIG_SHEET  = "Quota_Config";
var CLOSED_WON_REPORT_SHEET_NAME = "SF_Closed_Won_CQ";
var SUMMARY_NEXT_QUARTER_OVERRIDE = 8437800;
var SUMMARY_TEAM_NAME   = "NorthCentral Enterprise";
var SUMMARY_AI_SHEET    = "Summary_AI";
var SUMMARY_AI_MODEL    = "gpt-5.1";

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

var NAME_ALIASES_MAP = {
  "niko.argaez@glean.com":      ["Niko Argaez", "Nicolas Argaez"],
  "taylor.lundberg@glean.com":  ["Taylor Lundberg"],
  "james.demory@glean.com":     ["James DeMory", "James Demory"],
  "johnny.hatch@glean.com":     ["Johnny Hatch", "John Hatch"],
  "melissa.richards@glean.com": ["Melissa Richards"],
  "daniel.broderick@glean.com": ["Daniel Broderick", "Dan Broderick"],
  "billy.schuett@glean.com":    ["Billy Schuett", "William Schuett"]
};

var REP_EMAILS = [
  "niko.argaez@glean.com",
  "taylor.lundberg@glean.com",
  "james.demory@glean.com",
  "johnny.hatch@glean.com",
  "melissa.richards@glean.com",
  "daniel.broderick@glean.com"
];

var POC_ACCOUNTS_HEADERS = ["Account","Pilot Type","Status","Stage","Owner","SE","Champion","Pilot Start","Pilot End","Competitor","SFDC URL"];
var POC_ACTIONS_HEADERS  = ["Account","Action Item","RAG","Signal Source","Notes","Last Updated","Updated By"];
var SUMMARY_ADMIN_HEADERS = ["week_of","leader_note","shout_outs","weekly_ask","theme","footer_note"];
var QUOTA_CONFIG_HEADERS  = ["quarter_key","quarter_label","team_name","team_quota","closed_to_date","manager_name","notes"];

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
  "Engagement Changes","Forecast Changes","SF Accurate","NBM Scheduled This Week","NBM Scheduled Count"
];

// -----------------------------------------------------------------------
// HELPERS
// -----------------------------------------------------------------------
function _propOrDefault_(key, fallback) {
  var val = PropertiesService.getScriptProperties().getProperty(key);
  return val !== null ? val : fallback;
}

function getRecapSheet_() {
  return getSubmissionsSheet_();
}

function getOrCreateSummaryAdminSheet() {
  var ss = SpreadsheetApp.openById(TRACKER_SHEET_ID);
  var sheet = ss.getSheetByName(SUMMARY_ADMIN_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SUMMARY_ADMIN_SHEET, ss.getSheets().length);
    sheet.appendRow(SUMMARY_ADMIN_HEADERS);
    sheet.getRange(1, 1, 1, SUMMARY_ADMIN_HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateQuotaConfigSheet() {
  var ss = SpreadsheetApp.openById(TRACKER_SHEET_ID);
  var sheet = ss.getSheetByName(QUOTA_CONFIG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(QUOTA_CONFIG_SHEET, ss.getSheets().length);
    sheet.appendRow(QUOTA_CONFIG_HEADERS);
    sheet.getRange(1, 1, 1, QUOTA_CONFIG_HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// -----------------------------------------------------------------------
// OPERATING STANDARDS — US Enterprise KPI targets
// -----------------------------------------------------------------------
var OPERATING_STANDARDS = {
  stage2_plus_coverage_x: 4,       // 4x quota in Stage 2+ pipeline
  nbms_per_quarter: 5,             // 5 NBMs per quarter
  nnms_per_quarter: 15,            // 15 NNMs per quarter (source: pending)
  meetings_per_week: 8,            // 8 meetings per week (definition needs leadership clarity)
  ebr_per_tier1_per_quarter: 1     // 1 EBR per Tier 1 customer per quarter (source: pending)
};

// -----------------------------------------------------------------------
// RAINMAKER — Glean Agents API integration
// -----------------------------------------------------------------------
var GLEAN_API_BASE = 'https://scio-prod-be.glean.com/rest/api/v1';
var GLEAN_AGENT_WAIT_ENDPOINT = GLEAN_API_BASE + '/agents/runs/wait';
var RAINMAKER_AGENT_ID = '';  // Filled in after agent is built in Glean Agent Builder
var RAINMAKER_SHEET_NAME = 'Rainmaker';
var RAINMAKER_LOG_SHEET_NAME = 'Rainmaker_Log';

var RAINMAKER_CATEGORIES = [
  { key: 'nbm',          label: 'NBMs',                  weight: 1 },
  { key: 'pipe_adds',    label: 'Pipe Adds',             weight: 1 },
  { key: 'pipe_dollars', label: 'Pipe $$',               weight: 1 },
  { key: 'c_level',      label: 'C-Level',               weight: 1 },
  { key: 'stage4_plus',  label: 'Stage 4+',              weight: 2 },
  { key: 'pocs',         label: 'POCs',                  weight: 1 },
  { key: 'partner_reg',  label: 'Partner Registrations', weight: 1 },
  { key: 'closed_won',   label: 'Closed Won Business',   weight: 3 }
];

var RAINMAKER_POINT_BUCKETS = [
  { minPercentile: 90, points: 6 },
  { minPercentile: 75, points: 5 },
  { minPercentile: 50, points: 4 },
  { minPercentile: 25, points: 3 },
  { minPercentile: 10, points: 2 },
  { minPercentile: 0,  points: 1 }
];

var RAINMAKER_EXCLUDED_EMAILS = ['billy.schuett@glean.com'];

// -----------------------------------------------------------------------
// HELPERS
// -----------------------------------------------------------------------
function getSubmissionsSheet_() {
  var ss = SpreadsheetApp.openById(TRACKER_SHEET_ID);

  // 1) Exact configured sheet name first.
  var byName = ss.getSheetByName(SUBMISSIONS_SHEET_NAME);
  if (byName) return byName;

  // 2) Fallback: resolve by expected header signature.
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var lastCol = sh.getLastColumn();
    if (lastCol < 3) continue;

    var header = sh.getRange(1, 1, 1, Math.min(lastCol, RECAP_HEADERS.length)).getValues()[0];
    var match = 0;
    var checkLen = Math.min(header.length, RECAP_HEADERS.length);
    for (var c = 0; c < checkLen; c++) {
      if (String(header[c] || "").trim() === String(RECAP_HEADERS[c] || "").trim()) match++;
    }

    // Require a strong signature match to avoid accidental misrouting.
    if (match >= 12) return sh;
  }

  throw new Error(
    "Submissions sheet not found. Expected sheet name '" + SUBMISSIONS_SHEET_NAME +
    "' or a matching recap header row."
  );
}

function getUserNameFromEmail(email) {
  var cleaned = String(email || "").toLowerCase().trim();
  if (!cleaned) return null;
  if (NAME_MAP[cleaned]) return NAME_MAP[cleaned];
  var local = cleaned.split("@")[0] || "";
  if (!local) return null;
  var parts = local.split(".").filter(Boolean).map(function(part) {
    return part.charAt(0).toUpperCase() + part.slice(1);
  });
  return parts.length ? parts.join(" ") : null;
}

function getUserNameCandidatesFromEmail(email) {
  var cleaned = String(email || "").toLowerCase().trim();
  if (!cleaned) return [];
  var seen = {};
  var out = [];

  function push_(name) {
    var s = String(name || "").trim();
    if (!s) return;
    var key = s.toLowerCase();
    if (seen[key]) return;
    seen[key] = true;
    out.push(s);
  }

  push_(NAME_MAP[cleaned]);
  var aliases = NAME_ALIASES_MAP[cleaned] || [];
  for (var i = 0; i < aliases.length; i++) push_(aliases[i]);
  push_(getUserNameFromEmail(cleaned));

  return out;
}

function normalizeEngagement(v) {
  if (!v) return "";
  var s = v.toString().trim().toLowerCase();
  if (s === "green")  return "Green";
  if (s === "yellow") return "Yellow";
  if (s === "red")    return "Red";
  return v.toString().trim();
}
// Canonical strict currency/number parser
// Rules:
// - Accepts plain numbers, optional leading '$' and comma separators, optional one-letter suffixes 'k' or 'm' (case-insensitive).
// - Does NOT silently scale plain small numbers. '1000' -> 1000. '1' -> 1. '1k' -> 1000 only when suffix present.
// - Returns null for ambiguous / invalid inputs.
function parseCurrencyCanonical_(val) {
  if (val === null || val === undefined || val === "") return null;
  if (typeof val === 'number' && !isNaN(val)) return Number(val);
  var s = String(val).trim();
  if (!s) return null;

  // Remove leading currency symbol(s) and surrounding whitespace
  var cleaned = s.replace(/^\$+/, '').trim();

  // Accept formats like: 1234, 1,234, 1,234.56, 1234.56, 1k, 1.5M, -2,345
  var m = cleaned.match(/^(-)?(?:([0-9]{1,3}(?:,[0-9]{3})*)|([0-9]+))(\.[0-9]+)?\s*([kKmM])?$/);
  if (!m) return null;

  var sign = m[1] ? -1 : 1;
  var intPart = m[2] || m[3] || '0';
  intPart = intPart.replace(/,/g, '');
  var frac = m[4] || '';
  var suffix = (m[5] || '').toLowerCase();
  var num = parseFloat((sign === -1 ? '-' : '') + intPart + frac);
  if (isNaN(num)) return null;
  if (suffix === 'k') num = num * 1000;
  if (suffix === 'm') num = num * 1000000;
  return num;
}

function parseMoneyish_(val) {
  var n = parseCurrencyCanonical_(val);
  // Keep legacy return types: empty-string for missing/invalid, numeric for valid values
  return n === null ? "" : n;
}

function normalizeRainmakerOpp_(rawVal) {
  var n = parseCurrencyCanonical_(rawVal);
  if (!(n > 0)) return "";
  return Math.min(Math.round(n), 9999999);
}

function normalizeMoneyThousandsCapped_(rawVal) {
  var n = parseCurrencyCanonical_(rawVal);
  if (!(n > 0)) return "";
  return Math.min(Math.round(n), 9999999);
}

function normalizeRainmakerCount_(rawVal) {
  if (rawVal === null || rawVal === undefined || rawVal === "") return 0;
  var n = Number(rawVal);
  if (isNaN(n) || n < 0) return 0;
  return Math.min(Math.round(n), 9999);
}

function normalizeForecastFields_(formObject) {
  ["commit","likely","upside","nq_commit"].forEach(function(f) {
    var raw = formObject[f];
    if (raw === null || raw === undefined || raw === "") {
      formObject[f] = "";
      return;
    }
    var n = parseCurrencyCanonical_(raw);
    if (n === null) {
      // Keep explicit empty string to indicate invalid/missing to caller; server validation will enforce if necessary
      formObject[f] = "";
      return;
    }
    // Accept zero as explicit input
    formObject[f] = Math.min(Math.round(n), 9999999);
  });

  var rmOpp = normalizeRainmakerOpp_(formObject.rm_opp);
  formObject.rm_opp = (rmOpp === "") ? "" : rmOpp;

  // Normalize rainmaker count fields (non-negative integers)
  formObject.rm_disco = normalizeRainmakerCount_(formObject.rm_disco);
  formObject.rm_nbm = normalizeRainmakerCount_(formObject.rm_nbm);
  formObject.rm_accts = normalizeRainmakerCount_(formObject.rm_accts);

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

// -----------------------------------------------------------------------
// SHARED HELPER — used by Code.js and other server-side files
// -----------------------------------------------------------------------
function cleanEmail_(v) {
  return String(v || "").replace(/^["']+|["']+$/g, "").trim().toLowerCase();
}
