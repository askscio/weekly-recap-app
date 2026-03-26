// -----------------------------------------------------------------------
// LAST DATA
// -----------------------------------------------------------------------
function getLastDataForUser(email) {
  email = String(email || "").replace(/^["']+|["']+$/g, "").trim().toLowerCase();
  var sheet = getRecapSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var fallback = null;
  for (var i = data.length - 1; i >= 0; i--) {
    var rowEmail = String(data[i][1] || "").trim().toLowerCase();
    if (rowEmail !== email) continue;
    var obj = parseRecapRow_(data[i]);
    if (!fallback) fallback = obj;
    // meaningful check
    var meaningful = obj.weekOf || obj.pulseReason ||
      Number(obj.rm_disco) > 0 || Number(obj.rm_nbm) > 0 ||
      obj.rm_opp || obj.commit || obj.likely || obj.upside ||
      obj.goal || obj.risk || obj.ask;
    if (!meaningful) {
      for (var ai = 1; ai <= 6; ai++) {
        if (obj["acct"+ai+"_name"] || obj["acct"+ai+"_move"] ||
            obj["acct"+ai+"_next"] || obj["acct"+ai+"_engage"]) {
          meaningful = true; break;
        }
      }
    }
    if (meaningful) {
      fillMissingFieldsFromHistory_(data, email, i, obj);
      return obj;
    }
  }
  if (fallback) return fallback;
  return null;
}

function parseRecapRow_(r) {
  var obj = {
    weekOf: r[2] ? (r[2] instanceof Date ? r[2].toISOString() : String(r[2])) : "",
    pulseScore: r[3], pulseReason: r[4],
    rm_disco: r[5], rm_nbm: r[6], rm_opp: r[7], rm_accts: r[8],
    forecast_note: r[45],
    commit: r[46], likely: r[47], upside: r[48], nq_commit: r[49],
    goal: r[50], risk: r[51], ask: r[52],
    sf_accurate: r[55],
    nbm_scheduled_week: r[56],
    nbm_scheduled_count: r[57]
  };
  ACCT_BASES.forEach(function(b, idx) {
    var n = idx + 1;
    obj["acct" + n + "_name"]   = r[b];
    obj["acct" + n + "_stage"]  = r[b + 1];
    obj["acct" + n + "_arr"]    = r[b + 2];
    obj["acct" + n + "_engage"] = r[b + 3];
    obj["acct" + n + "_move"]   = r[b + 4];
    obj["acct" + n + "_next"]   = r[b + 5];
  });
  return obj;
}

function isBlankVal_(v) {
  return v === null || v === undefined || (typeof v === "string" && v.trim() === "");
}

function fillMissingFieldsFromHistory_(rows, email, startIdx, target) {
  var keys = [
    "weekOf", "pulseReason",
    "forecast_note", "commit", "likely", "upside", "nq_commit",
    "goal", "risk", "ask", "sf_accurate"
  ];
  for (var a = 1; a <= 6; a++) {
    keys.push("acct" + a + "_name");
    keys.push("acct" + a + "_stage");
    keys.push("acct" + a + "_arr");
    keys.push("acct" + a + "_engage");
    keys.push("acct" + a + "_move");
    keys.push("acct" + a + "_next");
  }

  for (var i = startIdx - 1; i >= 0; i--) {
    var rowEmail = String(rows[i][1] || "").trim().toLowerCase();
    if (rowEmail !== email) continue;
    var older = parseRecapRow_(rows[i]);
    var pending = false;
    for (var k = 0; k < keys.length; k++) {
      var key = keys[k];
      if (isBlankVal_(target[key]) && !isBlankVal_(older[key])) target[key] = older[key];
      if (isBlankVal_(target[key])) pending = true;
    }
    if (!pending) break;
  }
}

function parseMoneyishServer_(val) {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === "number") return val;
  var s = String(val).toLowerCase().trim();
  if (!s) return 0;
  var n = parseFloat(s.replace(/[^0-9.]/g, "")) || 0;
  if (s.indexOf("m") !== -1) n = n * 1000000;
  else if (s.indexOf("k") !== -1) n = n * 1000;
  return n;
}

function getJsonCache_(key, fallback) {
  try {
    var raw = CacheService.getScriptCache().get(key);
    if (!raw) return fallback;
    return JSON.parse(raw);
  } catch (err) {
    Logger.log("getJsonCache_ failed for " + key + ": " + err.message);
    return fallback;
  }
}

function putJsonCache_(key, value, ttlSec) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(value), ttlSec || 30);
  } catch (err) {
    Logger.log("putJsonCache_ failed for " + key + ": " + err.message);
  }
}

function getQaRepOptions_() {
  var out = [];
  for (var i = 0; i < REP_EMAILS.length; i++) {
    var email = String(REP_EMAILS[i] || "").toLowerCase().trim();
    if (!email) continue;
    out.push({
      email: email,
      name: getUserNameFromEmail(email) || email
    });
  }
  return out;
}

function findHeaderIndex_(headers, candidates) {
  var normalized = headers.map(function(h) {
    return String(h || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
  });
  for (var i = 0; i < candidates.length; i++) {
    var c = candidates[i].toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
    for (var j = 0; j < normalized.length; j++) {
      if (normalized[j] === c) return j;
    }
  }
  return -1;
}

function getLatestSelfForecastForUser(email) {
  email = String(email || "").replace(/^["']+|["']+$/g, "").trim().toLowerCase();
  var out = { commit: 0, likely: 0, upside: 0, nq_commit: 0 };
  var sheet = getRecapSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return out;
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var colEmail = findHeaderIndex_(headers, ["Email"]);
  var colCommit = findHeaderIndex_(headers, ["Commit"]);
  var colLikely = findHeaderIndex_(headers, ["Most Likely", "Likely"]);
  var colUpside = findHeaderIndex_(headers, ["Best Case", "Upside"]);
  var colNQ = findHeaderIndex_(headers, ["NQ Commit", "Next Quarter Commit"]);

  if (colEmail < 0) colEmail = 1; // fallback
  if (colCommit < 0) colCommit = 46;
  if (colLikely < 0) colLikely = 47;
  if (colUpside < 0) colUpside = 48;
  if (colNQ < 0) colNQ = 49;

  var rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var target = String(email || "").toLowerCase().trim();
  for (var i = rows.length - 1; i >= 0; i--) {
    var rowEmail = String(rows[i][colEmail] || "").toLowerCase().trim();
    if (rowEmail !== target) continue;
    if (!out.commit) out.commit = parseMoneyishServer_(rows[i][colCommit]);
    if (!out.likely) out.likely = parseMoneyishServer_(rows[i][colLikely]);
    if (!out.upside) out.upside = parseMoneyishServer_(rows[i][colUpside]);
    if (!out.nq_commit) out.nq_commit = parseMoneyishServer_(rows[i][colNQ]);
    if (out.commit && out.likely && out.upside && out.nq_commit) break;
  }
  return out;
}

// -----------------------------------------------------------------------
// CHANGE SUMMARIES
// -----------------------------------------------------------------------
function buildEngagementSummary(lastData, formObject) {
  if (!lastData) return "";
  var moves = {};
  for (var i = 1; i <= 6; i++) {
    var prev = normalizeEngagement(lastData["acct" + i + "_engage"]);
    var curr = normalizeEngagement(formObject["acct" + i + "_engage"]);
    if (prev && curr && prev !== curr) {
      var key = prev + " to " + curr;
      moves[key] = (moves[key] || 0) + 1;
    }
  }
  return Object.keys(moves).map(function(k) {
    return moves[k] + " account" + (moves[k] > 1 ? "s" : "") + " moved " + k;
  }).join(", ");
}

function buildForecastSummary(lastData, formObject) {
  if (!lastData) return "";
  return [
    { label: "Commit", prev: lastData.commit, curr: formObject.commit },
    { label: "Most Likely", prev: lastData.likely, curr: formObject.likely },
    { label: "Best Case", prev: lastData.upside, curr: formObject.upside }
  ].filter(function(f) {
    var p = (f.prev || "").toString().replace(/[^0-9]/g, "");
    var c = (f.curr || "").toString().replace(/[^0-9]/g, "");
    return p && c && p !== c;
  }).map(function(f) {
    var p2 = parseInt((f.prev || "").toString().replace(/[^0-9]/g, ""), 10);
    var c2 = parseInt((f.curr || "").toString().replace(/[^0-9]/g, ""), 10);
    return f.label + ": $" + p2.toLocaleString("en-US") + " to $" + c2.toLocaleString("en-US");
  }).join(" | ");
}

// -----------------------------------------------------------------------
// PROCESS FORM (SAVE)
// -----------------------------------------------------------------------
function processForm(formObject) {
  var ss = SpreadsheetApp.openById(TRACKER_SHEET_ID);
  var sheet = getRecapSheet_();
  function resolveEmail_() {
    var e = "";
    try { e = Session.getEffectiveUser().getEmail(); } catch (_) {}
    if (!e) {
      try { e = Session.getActiveUser().getEmail(); } catch (_) {}
    }
    return String(e || "").trim().toLowerCase();
  }
  var submitEmail = (formObject.previewEmail && formObject.previewEmail !== "")
    ? String(formObject.previewEmail).trim().toLowerCase()
    : resolveEmail_();
  formObject = normalizeForecastFields_(formObject);
  var lastData = getLastDataForUser(submitEmail);
  var engagementSummary = buildEngagementSummary(lastData, formObject);
  var forecastSummary = buildForecastSummary(lastData, formObject);
  if (sheet.getLastRow() < 1) {
    sheet.appendRow(RECAP_HEADERS);
    sheet.getRange(1, 1, 1, RECAP_HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  } else if (sheet.getLastColumn() < RECAP_HEADERS.length) {
    sheet.getRange(1, 1, 1, RECAP_HEADERS.length).setValues([RECAP_HEADERS]);
    sheet.getRange(1, 1, 1, RECAP_HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    new Date(), submitEmail, formObject.weekOf,
    formObject.pulseScore, formObject.pulseReason,
    formObject.rm_disco, formObject.rm_nbm, formObject.rm_opp, formObject.rm_accts,
    formObject.acct1_name, formObject.acct1_stage, formObject.acct1_arr, formObject.acct1_engage, formObject.acct1_move, formObject.acct1_next,
    formObject.acct2_name, formObject.acct2_stage, formObject.acct2_arr, formObject.acct2_engage, formObject.acct2_move, formObject.acct2_next,
    formObject.acct3_name, formObject.acct3_stage, formObject.acct3_arr, formObject.acct3_engage, formObject.acct3_move, formObject.acct3_next,
    formObject.acct4_name, formObject.acct4_stage, formObject.acct4_arr, formObject.acct4_engage, formObject.acct4_move, formObject.acct4_next,
    formObject.acct5_name, formObject.acct5_stage, formObject.acct5_arr, formObject.acct5_engage, formObject.acct5_move, formObject.acct5_next,
    formObject.acct6_name, formObject.acct6_stage, formObject.acct6_arr, formObject.acct6_engage, formObject.acct6_move, formObject.acct6_next,
    formObject.forecast_note,
    formObject.commit, formObject.likely, formObject.upside, formObject.nq_commit,
    formObject.goal, formObject.risk, formObject.ask,
    engagementSummary, forecastSummary, formObject.sf_accurate, formObject.nbm_scheduled_week, formObject.nbm_scheduled_count
  ]);
  try {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: "\u2705 Weekly Recap: " + submitEmail,
      htmlBody: "<b>" + submitEmail + "</b> just submitted.<br>Pulse: " + formObject.pulseScore +
        "<br><a href='" + ss.getUrl() + "'>View Sheet</a>"
    });
  } catch (e) {
    Logger.log("Email failed: " + e.message);
  }
  return {
    message: "\u2705 Recap saved!",
    engagementSummary: engagementSummary,
    forecastSummary: forecastSummary
  };
}

function csvEscape_(v) {
  var s = (v === null || v === undefined) ? "" : String(v);
  return '"' + s.replace(/"/g, '""') + '"';
}

function getResolvedCallerEmail_() {
  var caller = "";
  try { caller = String(Session.getEffectiveUser().getEmail() || ""); } catch (_) {}
  if (!caller) {
    try { caller = String(Session.getActiveUser().getEmail() || ""); } catch (_) {}
  }
  return caller.trim().toLowerCase();
}

function getMainSheetSnapshotForCodex() {
  var caller = getResolvedCallerEmail_();
  if (caller !== ADMIN_EMAIL.toLowerCase()) {
    throw new Error("Admin access required.");
  }

  var sheet = getRecapSheet_();
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
    return { filename: "main-sheet-snapshot-empty.csv", csv: "", rowCount: 0, generatedAt: new Date().toISOString() };
  }

  var values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getDisplayValues();
  var csv = values.map(function(row) {
    return row.map(csvEscape_).join(",");
  }).join("\r\n");

  var ts = Utilities.formatDate(new Date(), "America/Chicago", "yyyy-MM-dd_HH-mm-ss");
  return {
    filename: "main-sheet-snapshot-" + ts + ".csv",
    csv: csv,
    rowCount: Math.max(0, values.length - 1),
    generatedAt: new Date().toISOString()
  };
}

function ensureAdminCaller_() {
  var caller = getResolvedCallerEmail_();
  var admin = String(ADMIN_EMAIL || "").toLowerCase();
  if (caller && caller !== admin) throw new Error("Admin access required.");
}

function normHeaderKey_(h) {
  return String(h || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function parseMoneyishForFix_(val) {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === "number") return val > 0 ? val : 0;
  var s = String(val).toLowerCase().trim();
  if (!s) return 0;
  var n = parseFloat(s.replace(/[^0-9.]/g, "")) || 0;
  if (!n) return 0;
  if (s.indexOf("m") !== -1) n = n * 1000000;
  else if (s.indexOf("k") !== -1) n = n * 1000;
  return n > 0 ? n : 0;
}

function hasUnsupportedMoneyText_(val) {
  if (val === null || val === undefined) return false;
  var s = String(val).trim().toLowerCase();
  if (!s) return false;
  // Allow digits, spaces, $, comma, dot, dash, and k/m suffixes.
  return /[^0-9\s$.,km-]/i.test(s);
}

function getMainSheetDataHygieneReport() {
  ensureAdminCaller_();
  var sheet = getRecapSheet_();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { ok: true, rowCount: 0, issues: [] };

  var values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headers = values[0] || [];
  var rows = values.slice(1);
  var headerMap = {};
  for (var c = 0; c < headers.length; c++) headerMap[normHeaderKey_(headers[c])] = c;

  var issues = [];
  var blankHeaderCount = headers.filter(function(h) { return String(h || "").trim() === ""; }).length;
  if (blankHeaderCount > 0) issues.push(blankHeaderCount + " blank header columns detected.");

  var emailIdx = headerMap.email;
  var weekIdx = headerMap.weekof;
  var pulseReasonIdx = headerMap.pulsereason;
  var nqIdx = headerMap.nqcommit;
  var newOppIdx = headerMap.newopparr;
  if (newOppIdx === undefined) newOppIdx = headerMap.rmopparr;
  var commitIdx = headerMap.commit;
  var likelyIdx = headerMap.mostlikely;
  var bestIdx = headerMap.bestcase;

  var blankWeek = 0, blankPulseReason = 0, nqZero = 0;
  var dupKeys = {};
  var weirdMoney = [];
  var legacySmall = [];

  for (var r = 0; r < rows.length; r++) {
    var rowNum = r + 2;
    var email = emailIdx === undefined ? "" : String(rows[r][emailIdx] || "").trim().toLowerCase();
    var week = weekIdx === undefined ? "" : String(rows[r][weekIdx] || "").trim();
    var pulseReason = pulseReasonIdx === undefined ? "" : String(rows[r][pulseReasonIdx] || "").trim();

    if (!week) blankWeek++;
    if (!pulseReason) blankPulseReason++;
    if (email && week) {
      var key = email + "|" + week;
      dupKeys[key] = (dupKeys[key] || 0) + 1;
    }

    if (nqIdx !== undefined) {
      var nqv = parseMoneyishForFix_(rows[r][nqIdx]);
      if (!(nqv > 0)) nqZero++;
    }

    var fields = [];
    if (newOppIdx !== undefined) fields.push({ name: "New Opp ARR", idx: newOppIdx });
    if (commitIdx !== undefined) fields.push({ name: "Commit", idx: commitIdx });
    if (likelyIdx !== undefined) fields.push({ name: "Most Likely", idx: likelyIdx });
    if (bestIdx !== undefined) fields.push({ name: "Best Case", idx: bestIdx });
    if (nqIdx !== undefined) fields.push({ name: "NQ Commit", idx: nqIdx });
    for (var i = 1; i <= 6; i++) {
      var k = headerMap["acct" + i + "arr"];
      if (k !== undefined) fields.push({ name: "Acct" + i + " ARR", idx: k });
    }

    for (var f = 0; f < fields.length; f++) {
      var raw = rows[r][fields[f].idx];
      var n = parseMoneyishForFix_(raw);
      var s = String(raw || "").trim();
      if (hasUnsupportedMoneyText_(raw) && weirdMoney.length < 20) {
        weirdMoney.push({ row: rowNum, email: email, field: fields[f].name, value: s });
      }
      if (n > 0 && n < 1000 && legacySmall.length < 20) {
        legacySmall.push({ row: rowNum, email: email, field: fields[f].name, value: s });
      }
    }
  }

  var dupList = [];
  for (var k2 in dupKeys) if (dupKeys[k2] > 1) dupList.push({ key: k2, count: dupKeys[k2] });
  dupList.sort(function(a, b) { return b.count - a.count; });

  return {
    ok: true,
    generatedAt: new Date().toISOString(),
    rowCount: rows.length,
    blankHeaderCount: blankHeaderCount,
    blankWeekCount: blankWeek,
    blankPulseReasonCount: blankPulseReason,
    nqZeroOrBlankCount: nqZero,
    duplicateEmailWeekCount: dupList.length,
    duplicateEmailWeekTop: dupList.slice(0, 10),
    weirdMoneyCount: weirdMoney.length,
    weirdMoneySamples: weirdMoney,
    legacySmallMoneyCount: legacySmall.length,
    legacySmallMoneySamples: legacySmall,
    issues: issues
  };
}

function runMainSheetMoneyBackfill(apply) {
  ensureAdminCaller_();
  var doApply = !!apply;
  var sheet = getRecapSheet_();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { ok: true, apply: doApply, changedCells: 0, rowsTouched: 0, samples: [] };

  var range = sheet.getRange(1, 1, lastRow, lastCol);
  var values = range.getValues();
  var headers = values[0] || [];
  var headerMap = {};
  for (var c = 0; c < headers.length; c++) headerMap[normHeaderKey_(headers[c])] = c;

  var targetCols = [];
  function addCol(name) {
    var idx = headerMap[name];
    if (idx !== undefined && targetCols.indexOf(idx) === -1) targetCols.push(idx);
  }
  addCol("newopparr");
  addCol("rmopparr");
  addCol("commit");
  addCol("mostlikely");
  addCol("bestcase");
  addCol("nqcommit");
  for (var i = 1; i <= 6; i++) addCol("acct" + i + "arr");

  var changed = 0;
  var rowsTouchedMap = {};
  var samples = [];

  for (var r = 1; r < values.length; r++) {
    for (var t = 0; t < targetCols.length; t++) {
      var col = targetCols[t];
      var raw = values[r][col];
      if (raw === null || raw === undefined || raw === "") continue;
      if (hasUnsupportedMoneyText_(raw)) continue;

      var rawStr = String(raw).toLowerCase().trim();
      var parsed = parseMoneyishForFix_(raw);
      if (!(parsed > 0)) continue;

      var normalized = parsed;
      if (!/[mk]/.test(rawStr) && parsed < 1000) normalized = parsed * 1000;
      normalized = Math.min(Math.round(normalized), 9999999);

      var currentNumeric = parseMoneyishForFix_(raw);
      var rawIsNumber = typeof raw === "number";
      var sameNumeric = Math.round(currentNumeric) === normalized;
      if (rawIsNumber && sameNumeric) continue;

      if (doApply) values[r][col] = normalized;
      changed++;
      rowsTouchedMap[r + 1] = true;
      if (samples.length < 25) {
        samples.push({
          row: r + 1,
          col: col + 1,
          header: headers[col],
          from: String(raw),
          to: normalized
        });
      }
    }
  }

  if (doApply && changed > 0) {
    range.setValues(values);
  }

  return {
    ok: true,
    apply: doApply,
    changedCells: changed,
    rowsTouched: Object.keys(rowsTouchedMap).length,
    samples: samples
  };
}

function toMillis_(v) {
  if (v instanceof Date) return v.getTime();
  var d = new Date(v);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function runMainSheetCleanup(apply) {
  ensureAdminCaller_();
  var doApply = !!apply;
  var sheet = getRecapSheet_();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return { ok: true, apply: doApply, changedCells: 0, rowsTouched: 0, duplicateRowsRemoved: 0, duplicateKeys: 0, samples: [], removedRows: [] };
  }

  var range = sheet.getRange(1, 1, lastRow, lastCol);
  var values = range.getValues();
  var headers = values[0] || [];
  var headerMap = {};
  for (var c = 0; c < headers.length; c++) headerMap[normHeaderKey_(headers[c])] = c;

  var timestampIdx = headerMap.timestamp;
  var emailIdx = headerMap.email;
  var weekIdx = headerMap.weekof;

  var targetCols = [];
  function addCol(name) {
    var idx = headerMap[name];
    if (idx !== undefined && targetCols.indexOf(idx) === -1) targetCols.push(idx);
  }
  addCol("newopparr");
  addCol("rmopparr");
  addCol("commit");
  addCol("mostlikely");
  addCol("bestcase");
  addCol("nqcommit");
  for (var i = 1; i <= 6; i++) addCol("acct" + i + "arr");

  var changed = 0;
  var rowsTouchedMap = {};
  var samples = [];

  // Pass 1: normalize money values.
  for (var r = 1; r < values.length; r++) {
    for (var t = 0; t < targetCols.length; t++) {
      var col = targetCols[t];
      var raw = values[r][col];
      if (raw === null || raw === undefined || raw === "") continue;
      if (hasUnsupportedMoneyText_(raw)) continue;

      var rawStr = String(raw).toLowerCase().trim();
      var parsed = parseMoneyishForFix_(raw);
      if (!(parsed > 0)) continue;

      var normalized = parsed;
      if (!/[mk]/.test(rawStr) && parsed < 1000) normalized = parsed * 1000;
      normalized = Math.min(Math.round(normalized), 9999999);

      var currentNumeric = parseMoneyishForFix_(raw);
      var rawIsNumber = typeof raw === "number";
      var sameNumeric = Math.round(currentNumeric) === normalized;
      if (rawIsNumber && sameNumeric) continue;

      if (doApply) values[r][col] = normalized;
      changed++;
      rowsTouchedMap[r + 1] = true;
      if (samples.length < 25) {
        samples.push({
          type: "money",
          row: r + 1,
          col: col + 1,
          header: headers[col],
          from: String(raw),
          to: normalized
        });
      }
    }
  }

  // Pass 2: detect duplicates by email+week and keep latest timestamp.
  var latestByKey = {};
  var rowsToDelete = {};
  var duplicateKeySet = {};
  var removedRows = [];

  if (emailIdx !== undefined && weekIdx !== undefined) {
    for (var r2 = 1; r2 < values.length; r2++) {
      var email = String(values[r2][emailIdx] || "").trim().toLowerCase();
      var week = String(values[r2][weekIdx] || "").trim();
      if (!email || !week) continue;

      var key = email + "|" + week;
      var ts = (timestampIdx !== undefined) ? toMillis_(values[r2][timestampIdx]) : 0;
      if (!latestByKey[key]) {
        latestByKey[key] = { rowIdx: r2, ts: ts };
      } else {
        duplicateKeySet[key] = true;
        var prev = latestByKey[key];
        var keepCurrent = ts >= prev.ts;
        var dropIdx = keepCurrent ? prev.rowIdx : r2;
        if (keepCurrent) latestByKey[key] = { rowIdx: r2, ts: ts };
        rowsToDelete[dropIdx] = true;
      }
    }
  }

  var deleteIdxs = Object.keys(rowsToDelete).map(function(k) { return parseInt(k, 10); });
  deleteIdxs.sort(function(a, b) { return b - a; });

  if (doApply && changed > 0) {
    range.setValues(values);
  }
  if (doApply && deleteIdxs.length > 0) {
    for (var di = 0; di < deleteIdxs.length; di++) {
      var rowIdx = deleteIdxs[di];
      var rowNum = rowIdx + 1; // sheet row (values includes header)
      if (removedRows.length < 25) {
        var rv = values[rowIdx] || [];
        removedRows.push({
          row: rowNum,
          email: (emailIdx !== undefined) ? String(rv[emailIdx] || "") : "",
          weekOf: (weekIdx !== undefined) ? String(rv[weekIdx] || "") : ""
        });
      }
      sheet.deleteRow(rowNum);
    }
  }

  return {
    ok: true,
    apply: doApply,
    changedCells: changed,
    rowsTouched: Object.keys(rowsTouchedMap).length,
    duplicateRowsRemoved: deleteIdxs.length,
    duplicateKeys: Object.keys(duplicateKeySet).length,
    samples: samples,
    removedRows: removedRows
  };
}

// -----------------------------------------------------------------------
// SALESFORCE DATA
// -----------------------------------------------------------------------
function normNameToken_(v) {
  return String(v || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function ownerMatchesRep_(ownerName, repName) {
  var owner = normNameToken_(ownerName);
  var rep = normNameToken_(repName);
  if (!owner || !rep) return false;
  if (owner === rep) return true;
  if (owner.indexOf(rep) !== -1 || rep.indexOf(owner) !== -1) return true;

  var ownerParts = owner.split(" ").filter(Boolean);
  var repParts = rep.split(" ").filter(Boolean);
  if (!ownerParts.length || !repParts.length) return false;

  var ownerFirst = ownerParts[0];
  var ownerLast = ownerParts[ownerParts.length - 1];
  var repFirst = repParts[0];
  var repLast = repParts[repParts.length - 1];

  // Require last-name consistency, then allow first-name variants/initials.
  if (ownerLast !== repLast) return false;
  if (ownerFirst === repFirst) return true;
  if (ownerFirst.charAt(0) && repFirst.charAt(0) && ownerFirst.charAt(0) === repFirst.charAt(0)) return true;
  if (ownerFirst.indexOf(repFirst) === 0 || repFirst.indexOf(ownerFirst) === 0) return true;
  return false;
}

function ownerMatchesAnyRepName_(ownerName, repNames) {
  if (!repNames || !repNames.length) return false;
  for (var i = 0; i < repNames.length; i++) {
    if (ownerMatchesRep_(ownerName, repNames[i])) return true;
  }
  return false;
}

function getTeamForecastTotalsFromDealsSheet_(quarterKey, nextQuarterKey, fallbackTotals) {
  var out = {
    closed: Number((fallbackTotals && fallbackTotals.closed) || 0),
    commit: Number((fallbackTotals && fallbackTotals.commit) || 0),
    likely: Number((fallbackTotals && fallbackTotals.likely) || 0),
    bestCase: Number((fallbackTotals && fallbackTotals.bestCase) || 0),
    nextQuarter: Number((fallbackTotals && fallbackTotals.nextQuarter) || 0),
    sources: {
      closed: { source: 'fallback', label: 'Fallback', usedFallback: true, warning: '' },
      commit: { source: 'rep_submission_fallback', label: 'Rep submission fallback', usedFallback: true, warning: '' },
      likely: { source: 'rep_submission_fallback', label: 'Rep submission fallback', usedFallback: true, warning: '' },
      bestCase: { source: 'rep_submission_fallback', label: 'Rep submission fallback', usedFallback: true, warning: '' },
      nextQuarter: { source: 'rep_submission_fallback', label: 'Rep submission fallback', usedFallback: true, warning: '' }
    }
  };

  try {
    var ss = SpreadsheetApp.openById(TRACKER_SHEET_ID);
    var dealsSheet = ss.getSheetByName(DEALS_SHEET_NAME);
    if (!dealsSheet || dealsSheet.getLastRow() < 3) {
      var warn = DEALS_SHEET_NAME + ' sheet missing or empty.';
      out.sources.closed.warning = warn;
      out.sources.commit.warning = warn;
      out.sources.likely.warning = warn;
      out.sources.bestCase.warning = warn;
      out.sources.nextQuarter.warning = warn;
      return out;
    }

    var numRows = dealsSheet.getLastRow() - 2;
    var numCols = Math.max(DEALS_COL_OWNER, DEALS_COL_FORECAST, DEALS_COL_AMOUNT, DEALS_COL_CLOSE, DEALS_COL_ACCOUNT);
    var rows = dealsSheet.getRange(3, 1, numRows, numCols).getValues();
    var totals = { closed: 0, commit: 0, likely: 0, bestCase: 0, nextQuarter: 0 };

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var forecast = String(row[DEALS_COL_FORECAST - 1] || '').trim().toLowerCase();
      var amountRaw = row[DEALS_COL_AMOUNT - 1];
      var closeRaw = row[DEALS_COL_CLOSE - 1];
      var closeDate = closeRaw instanceof Date ? closeRaw : new Date(closeRaw);
      var amount = typeof amountRaw === 'number' ? amountRaw : parseFloat(String(amountRaw || '0').replace(/[^0-9.]/g, '')) || 0;
      if (!amount) continue;
      if (!(closeDate instanceof Date) || isNaN(closeDate.getTime())) continue;

      var rowQuarter = quarterKeyFromDate_(closeDate);
      var isCommit = forecast.indexOf('commit') !== -1;
      var isMostLikely = forecast.indexOf('most likely') !== -1 || forecast.indexOf('most_likely') !== -1;
      var isBestCase = forecast.indexOf('best case') !== -1 || forecast.indexOf('best_case') !== -1;
      var isClosed = forecast.indexOf('closed') !== -1;

      // The Total Pipe SF-data tab is already the scoped forecast universe.
      // Roll the full stack cumulatively from all rows in that tab.
      if (isClosed) {
        totals.closed += amount;
        totals.commit += amount;
        totals.likely += amount;
        totals.bestCase += amount;
      } else if (isCommit) {
        totals.commit += amount;
        totals.likely += amount;
        totals.bestCase += amount;
      } else if (isMostLikely) {
        totals.likely += amount;
        totals.bestCase += amount;
      } else if (isBestCase) {
        totals.bestCase += amount;
      }

      // Keep next-quarter commit date-based if those rows exist in the sheet.
      if (rowQuarter === nextQuarterKey && (isCommit || isClosed)) {
        totals.nextQuarter += amount;
      }
    }

    out.closed = totals.closed;
    out.commit = totals.commit;
    out.likely = totals.likely;
    out.bestCase = totals.bestCase;
    out.nextQuarter = totals.nextQuarter;
    out.sources.closed = { source: 'salesforce_total_pipe_sheet', label: DEALS_SHEET_NAME, usedFallback: false, warning: '' };
    out.sources.commit = { source: 'salesforce_total_pipe_sheet', label: DEALS_SHEET_NAME, usedFallback: false, warning: '' };
    out.sources.likely = { source: 'salesforce_total_pipe_sheet', label: DEALS_SHEET_NAME, usedFallback: false, warning: '' };
    out.sources.bestCase = { source: 'salesforce_total_pipe_sheet', label: DEALS_SHEET_NAME, usedFallback: false, warning: '' };
    out.sources.nextQuarter = { source: 'salesforce_total_pipe_sheet', label: DEALS_SHEET_NAME, usedFallback: false, warning: '' };
    return out;
  } catch (err) {
    var msg = DEALS_SHEET_NAME + ' source failed: ' + err.message;
    out.sources.closed.warning = msg;
    out.sources.commit.warning = msg;
    out.sources.likely.warning = msg;
    out.sources.bestCase.warning = msg;
    out.sources.nextQuarter.warning = msg;
    return out;
  }
}

function getSFDataForUser(email) {
  var blank = { commit: 0, mostLikely: 0, bestCase: 0, deals: [], nbmBlocks: [], nbmWarnings: [] };
  var nameCandidates = getUserNameCandidatesFromEmail(email);
  if (!nameCandidates.length) return blank;
  var ss = SpreadsheetApp.openById(TRACKER_SHEET_ID);
  var dealsSheet = ss.getSheetByName(DEALS_SHEET_NAME);
  if (!dealsSheet || dealsSheet.getLastRow() < 2) {
    Logger.log("getSFDataForUser: sheet missing or empty: " + DEALS_SHEET_NAME);
    return blank;
  }

  var lastRow = dealsSheet.getLastRow();
  var lastCol = dealsSheet.getLastColumn();
  var scanRows = Math.min(lastRow, 5);
  var scanValues = dealsSheet.getRange(1, 1, scanRows, lastCol).getValues();
  var headerRowIdx = -1;
  var colOwner = -1, colForecast = -1, colAmount = -1, colClose = -1, colAccount = -1;

  for (var hr = 0; hr < scanValues.length; hr++) {
    var header = scanValues[hr];
    var ownerIdx = findHeaderIndex_(header, ["Opportunity Owner", "Owner"]);
    var forecastIdx = findHeaderIndex_(header, ["Forecast Category", "Forecast"]);
    var amountIdx = findHeaderIndex_(header, ["Amount"]);
    var closeIdx = findHeaderIndex_(header, ["Close Date"]);
    var accountIdx = findHeaderIndex_(header, ["Account Name", "Account"]);
    if (ownerIdx !== -1 && forecastIdx !== -1 && amountIdx !== -1 && closeIdx !== -1 && accountIdx !== -1) {
      headerRowIdx = hr;
      colOwner = ownerIdx;
      colForecast = forecastIdx;
      colAmount = amountIdx;
      colClose = closeIdx;
      colAccount = accountIdx;
      break;
    }
  }

  if (headerRowIdx === -1) {
    Logger.log("getSFDataForUser: could not find expected headers in " + DEALS_SHEET_NAME);
    return blank;
  }

  var startRow = headerRowIdx + 2;
  var numRows = lastRow - (startRow - 1);
  if (numRows <= 0) return blank;
  var dealsValues = dealsSheet.getRange(startRow, 1, numRows, lastCol).getValues();

  var nbmSheet = ss.getSheetByName(NBM_SHEET_NAME);
  var nbmKeysSet = {};
  if (nbmSheet && nbmSheet.getLastRow() > 2) {
    var nbmRows = nbmSheet.getLastRow() - 2;
    var nbmValues = nbmSheet.getRange(3, 1, nbmRows, NBM_COL_KEY).getValues();
    for (var i = 0; i < nbmValues.length; i++) {
      var key = nbmValues[i][NBM_COL_KEY - 1];
      if (key && key.toString().trim() !== "") nbmKeysSet[key.toString().trim().toLowerCase()] = true;
    }
  }
  var commitTotal = 0, mlTotal = 0, bcTotal = 0;
  var deals = [], nbmBlocks = [], nbmWarnings = [];
  var primaryName = String(nameCandidates[0] || "");
  for (var r = 0; r < dealsValues.length; r++) {
    var row = dealsValues[r];
    var owner    = (row[colOwner] || "").toString().trim();
    var account  = (row[colAccount] || "").toString().trim();
    var forecast = (row[colForecast] || "").toString().trim();
    var amount   = row[colAmount];
    var closeDate = row[colClose];
    if (!owner || !account || !forecast || !amount) continue;
    if (!ownerMatchesAnyRepName_(owner, nameCandidates)) continue;
    var amt = typeof amount === "number"
      ? amount
      : parseFloat((amount || "0").toString().replace(/[^0-9.]/g, "")) || 0;
    if (!amt) continue;
    var forecastStr  = forecast.toLowerCase();
    var isClosed     = forecastStr.indexOf("closed") !== -1;
    var isCommit     = forecastStr.indexOf("commit") !== -1;
    var isMostLikely = forecastStr.indexOf("most likely") !== -1 || forecastStr.indexOf("most_likely") !== -1;
    var isBestCase   = forecastStr.indexOf("best case") !== -1 || forecastStr.indexOf("best_case") !== -1;
    if (!isClosed && !isCommit && !isMostLikely && !isBestCase) continue;
    var hasNBM = !!nbmKeysSet[account.toLowerCase()];
    deals.push({ category: forecast, account: account, amount: amt, closeDate: closeDate, hasNBM: hasNBM, owner: owner });
    if (isClosed)          { commitTotal += amt; mlTotal += amt; bcTotal += amt; }
    else if (isCommit)     { commitTotal += amt; mlTotal += amt; bcTotal += amt; }
    else if (isMostLikely) { mlTotal += amt; bcTotal += amt; }
    else if (isBestCase)   { bcTotal += amt; }
    if (amt >= NBM_MIN_AMOUNT && !hasNBM) {
      var tag = account + " \u2014 $" + Math.round(amt / 1000) + "k";
      if (isCommit)          nbmBlocks.push(tag + " \u2014 Commit \u2014 No NBM logged");
      else if (isMostLikely) nbmWarnings.push(tag + " \u2014 Most Likely \u2014 No NBM scheduled");
    }
  }
  Logger.log("getSFDataForUser [" + normNameToken_(primaryName) + "]: deals=" + deals.length + " commit=$" + commitTotal);
  return { commit: commitTotal, mostLikely: mlTotal, bestCase: bcTotal, deals: deals, nbmBlocks: nbmBlocks, nbmWarnings: nbmWarnings };
}

// -----------------------------------------------------------------------
// QBR CALL
// -----------------------------------------------------------------------
function getQBRCall(email) {
  var blank = { commit: 0, mostLikely: 0, bestCase: 0 };
  var nameCandidates = getUserNameCandidatesFromEmail(email);
  if (!nameCandidates.length) return blank;
  var sheet = SpreadsheetApp.openById(TRACKER_SHEET_ID).getSheetByName("QBR_Call");
  if (!sheet || sheet.getLastRow() < 2) return blank;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  function parseAmt(v) {
    if (!v) return 0;
    if (typeof v === "number") return v;
    return parseFloat(v.toString().replace(/[^0-9.]/g, "")) || 0;
  }
  for (var i = 0; i < data.length; i++) {
    if (ownerMatchesAnyRepName_(data[i][0], nameCandidates)) {
      return { commit: parseAmt(data[i][1]), mostLikely: parseAmt(data[i][2]), bestCase: parseAmt(data[i][3]) };
    }
  }
  return blank;
}

// -----------------------------------------------------------------------
// TEAM DASHBOARD DATA — lean, no SF/stats calls (those load async)
// -----------------------------------------------------------------------
function getTeamDashboardData() {
  var cached = getJsonCache_("team_dashboard_data_v1", null);
  if (cached) return cached;

  var sheet = getRecapSheet_();
  Logger.log("getTeamDashboardData: using sheet '" + sheet.getName() + "' with " + sheet.getLastRow() + " rows");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var allRows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var tz = Session.getScriptTimeZone();
  var teamData = [];
  REP_EMAILS.forEach(function(email) {
    var emailLower = email.toLowerCase().trim();
    var _bestRow = null;
    var _forecastFallback = { commit: 0, likely: 0, upside: 0, nq_commit: 0 };
    for (var i = allRows.length - 1; i >= 0; i--) {
      var rowEmail = String(allRows[i][1] || "").trim().toLowerCase();
      if (rowEmail !== emailLower) continue;
      if (!_bestRow) _bestRow = allRows[i]; // fallback: newest
      var _r = allRows[i];
      // Capture latest non-empty forecast fields as fallback values.
      if (!_forecastFallback.commit && _r[46] !== "" && _r[46] !== null && _r[46] !== undefined) _forecastFallback.commit = _r[46];
      if (!_forecastFallback.likely && _r[47] !== "" && _r[47] !== null && _r[47] !== undefined) _forecastFallback.likely = _r[47];
      if (!_forecastFallback.upside && _r[48] !== "" && _r[48] !== null && _r[48] !== undefined) _forecastFallback.upside = _r[48];
      if (!_forecastFallback.nq_commit && _r[49] !== "" && _r[49] !== null && _r[49] !== undefined) _forecastFallback.nq_commit = _r[49];
      var _meaningful = _r[2] || _r[4] || Number(_r[5]) > 0 || Number(_r[6]) > 0 || _r[7] || _r[46] || _r[47] || _r[48] || _r[50] || _r[51] || _r[52] || _r[56] || _r[57];
      if (!_meaningful) {
        // check accounts
        var _bases = [9,15,21,27,33,39];
        for (var _bi = 0; _bi < _bases.length && !_meaningful; _bi++) { if (_r[_bases[_bi]]) _meaningful = true; }
      }
      if (_meaningful) { _bestRow = _r; break; }
    }
    if (!_bestRow) { return; }
    var row = _bestRow;
    // dummy for-loop to allow break pattern below
    for (var _once = 0; _once < 1; _once++) {
      var engage = ACCT_BASES.map(function(b) { return normalizeEngagement(row[b + 3]); });
      var weekDate = "\u2014";
      try { weekDate = row[2] ? Utilities.formatDate(new Date(row[2]), tz, "MMM d, yyyy") : "\u2014"; } catch(ex) {}
      teamData.push({
        email: email,
        name: getUserNameFromEmail(email),
        isManager: false,
        weekOf: weekDate,
        pulseScore: row[3] || 0,
        pulseReason: row[4] || "",
        rm_disco: row[5] || 0, rm_nbm: row[6] || 0, rm_opp: row[7] || 0, rm_accts: row[8] || "",
        accounts: ACCT_BASES.map(function(b, idx) {
          return { name: row[b], stage: row[b+1], arr: row[b+2], engage: engage[idx], move: row[b+4], next: row[b+5] };
        }),
        forecast_note: row[45] || "",
        commit: (row[46] !== "" && row[46] !== null && row[46] !== undefined) ? row[46] : _forecastFallback.commit,
        likely: (row[47] !== "" && row[47] !== null && row[47] !== undefined) ? row[47] : _forecastFallback.likely,
        upside: (row[48] !== "" && row[48] !== null && row[48] !== undefined) ? row[48] : _forecastFallback.upside,
        nq_commit: (row[49] !== "" && row[49] !== null && row[49] !== undefined) ? row[49] : _forecastFallback.nq_commit,
        goal: row[50] || "", risk: row[51] || "", ask: row[52] || "",
        engagementChanges: row[53] || "", forecastChanges: row[54] || "", sfAccurate: row[55] || "",
        nbm_scheduled_week: row[56] || "",
        nbm_scheduled_count: row[57] || "",
        sfCommit: 0, sfMostLikely: 0, sfBestCase: 0,
        qbrCommit: 0, qbrMostLikely: 0, qbrBestCase: 0,
        nbmBlocks: [],
        disco_rank: "-", disco_amt: 0, nbm_rank: "-", nbm_amt: 0,
        pipe_rank: "-", pipe_amt: 0, stg4_rank: "-", stg4_amt: 0,
        redAccounts:    engage.filter(function(e) { return e === "Red"; }).length,
        yellowAccounts: engage.filter(function(e) { return e === "Yellow"; }).length,
        greenAccounts:  engage.filter(function(e) { return e === "Green"; }).length
      });
    } // end _once
  });
  Logger.log("getTeamDashboardData: returning " + teamData.length + " reps");
  putJsonCache_("team_dashboard_data_v1", teamData, 45);
  return teamData;
}

function countKeywordHits_(text, keywords) {
  var s = String(text || "").toLowerCase();
  if (!s) return 0;
  var hits = 0;
  for (var i = 0; i < keywords.length; i++) {
    if (s.indexOf(String(keywords[i]).toLowerCase()) !== -1) hits++;
  }
  return hits;
}

function trimSentenceWords_(text, maxWords) {
  var words = String(text || "").trim().split(/\s+/).filter(Boolean);
  if (words.length <= maxWords) return String(text || "").trim();
  return words.slice(0, maxWords).join(" ").replace(/[,:;.-]?$/, "") + "…";
}

function cleanSnippet_(text, maxLen) {
  var s = String(text || "").replace(/\s+/g, " ").trim();
  if (!s) return "";
  if (s.length <= maxLen) return s;
  return s.slice(0, maxLen).replace(/[,:;.-]?$/, "") + "…";
}


function editorialStopPhrases_() {
  return [
    'n/a', 'none yet', 'no changes', 'nothing new', 'all good', 'same as last week',
    'not currently', 'no current notes', 'tbd', 'na', 'none', 'unchanged'
  ];
}

function looksWeakSnippet_(text) {
  var s = String(text || '').replace(/\s+/g, ' ').trim().toLowerCase();
  if (!s) return true;
  var stops = editorialStopPhrases_();
  for (var i = 0; i < stops.length; i++) {
    if (s === stops[i] || s.indexOf(stops[i] + '.') === 0) return true;
  }
  return s.length < 18;
}

function scoreSnippet_(text, keywords, baseScore) {
  var s = String(text || '').toLowerCase();
  if (!s) return 0;
  var score = Number(baseScore || 0);
  for (var i = 0; i < keywords.length; i++) {
    if (s.indexOf(String(keywords[i]).toLowerCase()) !== -1) score += 3;
  }
  if (/\$[0-9]|\b\d{1,2}\/\d{1,2}\b|\bq[1-4]\b/.test(s)) score += 2;
  if (s.length > 70) score += 1;
  return score;
}

function formatRepPrefix_(repName) {
  var first = String(repName || '').trim().split(/\s+/)[0] || String(repName || '').trim();
  return first ? first + ': ' : '';
}

function formatAccountMention_(acctName) {
  var s = String(acctName || '').trim();
  return s ? s : 'the account';
}

function buildNarrativeLine_(prefix, text, maxLen) {
  var clean = cleanSnippet_(text, maxLen || 140);
  if (!clean) return '';
  return prefix + clean.charAt(0).toUpperCase() + clean.slice(1);
}

function pushCandidate_(arr, section, repName, text, score, accountName) {
  var clean = cleanSnippet_(text, 180);
  if (!clean || looksWeakSnippet_(clean)) return;
  arr.push({
    section: section,
    repName: repName,
    repFirst: String(repName || '').trim().split(/\s+/)[0] || repName,
    accountName: String(accountName || '').trim(),
    text: clean,
    score: Number(score || 0)
  });
}

function dedupeCandidates_(items) {
  var seen = {};
  var out = [];
  for (var i = 0; i < items.length; i++) {
    var key = [items[i].section, String(items[i].accountName || '').toLowerCase(), String(items[i].text || '').toLowerCase()].join('|');
    if (seen[key]) continue;
    seen[key] = true;
    out.push(items[i]);
  }
  return out;
}

function selectEditorialItems_(candidates, maxItems) {
  var list = dedupeCandidates_(candidates || []).sort(function(a, b) { return (b.score || 0) - (a.score || 0); });
  var out = [];
  var usedRep = {};
  for (var pass = 0; pass < 2 && out.length < maxItems; pass++) {
    for (var i = 0; i < list.length && out.length < maxItems; i++) {
      var item = list[i];
      if (item._used) continue;
      var repKey = String(item.repName || '').toLowerCase();
      if (pass === 0 && usedRep[repKey]) continue;
      item._used = true;
      usedRep[repKey] = true;
      out.push(item);
    }
  }
  return out;
}

function renderEditorialSection_(section, selected) {
  var lines = [];
  for (var i = 0; i < selected.length; i++) {
    var item = selected[i];
    var prefix = formatRepPrefix_(item.repName);
    if (section === 'deal_progression') {
      if (item.accountName) lines.push(prefix + formatAccountMention_(item.accountName) + ': ' + item.text);
      else lines.push(prefix + item.text);
    } else {
      lines.push(prefix + item.text);
    }
  }
  return lines;
}

function summarizeTeamRecapsEditorial_(latestRecaps) {
  var pulseSum = 0, pulseCount = 0;
  var themes = { momentum: 0, risk: 0, execution: 0, competition: 0, expansion: 0 };
  var bigDealAdds = [];
  var dealProgress = [];
  var forecastSignals = [];
  var priorities = [];
  var risks = [];
  var asks = [];
  var forecastNotes = [];

  for (var r = 0; r < latestRecaps.length; r++) {
    var row = latestRecaps[r];
    var parsed = parseRecapRow_(row);
    var repEmail = String(row[1] || '').trim().toLowerCase();
    var repName = getUserNameFromEmail(repEmail) || repEmail;
    var pulse = Number(parsed.pulseScore) || 0;
    if (pulse > 0) { pulseSum += pulse; pulseCount++; }

    var themeBundle = [parsed.pulseReason, parsed.goal, parsed.risk, parsed.ask, parsed.forecast_note].join(' ').toLowerCase();
    themes.momentum += countKeywordHits_(themeBundle, ['momentum', 'progress', 'strong', 'win', 'closed', 'expansion']);
    themes.risk += countKeywordHits_(themeBundle, ['risk', 'blocked', 'delay', 'legal', 'security', 'procurement', 'slip']);
    themes.execution += countKeywordHits_(themeBundle, ['execute', 'finalize', 'schedule', 'meeting', 'follow up', 'kickoff']);
    themes.competition += countKeywordHits_(themeBundle, ['copilot', 'competitor', 'competitive', 'claude', 'guru']);
    themes.expansion += countKeywordHits_(themeBundle, ['expansion', 'pipeline', 'land', 'rollout', 'department', 'enterprise']);

    var newOpp = parseMoneyishServer_(parsed.rm_opp);
    if (newOpp > 0) {
      pushCandidate_(bigDealAdds, 'big_deal_adds', repName,
        'Added ' + Math.round(newOpp).toLocaleString('en-US') + ' in new opp ARR.',
        8 + Math.min(8, Math.round(newOpp / 100000)), '');
    }

    var bestForecast = [
      { label: 'commit', value: parseMoneyishServer_(parsed.commit) },
      { label: 'most likely', value: parseMoneyishServer_(parsed.likely) },
      { label: 'best case', value: parseMoneyishServer_(parsed.upside) },
      { label: 'next quarter', value: parseMoneyishServer_(parsed.nq_commit) }
    ].sort(function(a, b) { return b.value - a.value; })[0];
    if (bestForecast && bestForecast.value >= 150000) {
      var forecastText = 'Carrying ' + bestForecast.label + ' at $' + Math.round(bestForecast.value).toLocaleString('en-US') + (parsed.forecast_note && !looksWeakSnippet_(parsed.forecast_note) ? ' — ' + cleanSnippet_(parsed.forecast_note, 90) : '');
      pushCandidate_(forecastSignals, 'forecast_signals', repName, forecastText, 7 + Math.min(8, Math.round(bestForecast.value / 150000)), '');
    }

    if (parsed.goal && !looksWeakSnippet_(parsed.goal)) {
      pushCandidate_(priorities, 'rep_priorities', repName, parsed.goal, scoreSnippet_(parsed.goal, ['close', 'finalize', 'schedule', 'land', 'commit', 'expand', 'kickoff'], 4), '');
    }
    if (parsed.risk && !looksWeakSnippet_(parsed.risk)) {
      pushCandidate_(risks, 'top_risks', repName, parsed.risk, scoreSnippet_(parsed.risk, ['legal', 'security', 'procurement', 'delay', 'blocked', 'stalled', 'slip', 'risk'], 5), '');
    }
    if (parsed.ask && !looksWeakSnippet_(parsed.ask)) {
      pushCandidate_(asks, 'manager_asks', repName, parsed.ask, scoreSnippet_(parsed.ask, ['support', 'feedback', 'join', 'escalat', 'help', 'intro', 'coach'], 4), '');
    }
    if (parsed.forecast_note && !looksWeakSnippet_(parsed.forecast_note)) {
      pushCandidate_(forecastNotes, 'forecast_notes', repName, parsed.forecast_note, scoreSnippet_(parsed.forecast_note, ['pull in', 'slip', 'close', 'signed', 'commit', 'likely', 'upside'], 4), '');
    }

    for (var a = 1; a <= 6; a++) {
      var acctName = String(parsed['acct' + a + '_name'] || '').trim();
      var move = String(parsed['acct' + a + '_move'] || '').trim();
      var next = String(parsed['acct' + a + '_next'] || '').trim();
      var stage = String(parsed['acct' + a + '_stage'] || '').trim();
      var arr = parseMoneyishServer_(parsed['acct' + a + '_arr']);
      var combo = [move, next, parsed.forecast_note, parsed.pulseReason].join(' ').toLowerCase();
      var score = 0;
      score += scoreSnippet_(move + ' ' + next, ['closed', 'signed', 'procurement', 'legal', 'approval', 'security', 'kickoff', 'deployment', 'commercial'], 3);
      if (arr >= 100000) score += Math.min(6, Math.round(arr / 200000));
      if (stage && /closed won|s4|s3|stage 4|procurement|security|pilot/i.test(stage)) score += 3;
      if (score <= 0) continue;
      var narrative = '';
      if (/closed won|contract signed|signed/i.test(combo) || /closed won/i.test(stage)) {
        narrative = 'Signed or commercial outcome on ' + formatAccountMention_(acctName) + ' is now real and worth leaning into.';
      } else if (/legal|procurement|approval|security|paperwork|msa|dpa/i.test(combo)) {
        narrative = formatAccountMention_(acctName) + ' is moving through a real commercial gate: ' + cleanSnippet_(move || next, 100);
      } else if (/kickoff|deploy|rollout|pilot|launch|implementation/i.test(combo)) {
        narrative = formatAccountMention_(acctName) + ' is moving operationally: ' + cleanSnippet_(move || next, 100);
      } else {
        narrative = formatAccountMention_(acctName) + ': ' + cleanSnippet_(move || next, 100);
      }
      pushCandidate_(dealProgress, 'deal_progression', repName, narrative, score, acctName);
    }
  }

  var avgPulse = pulseCount ? (pulseSum / pulseCount).toFixed(1) : '—';
  var themeArray = [
    { label: 'execution momentum', score: themes.momentum },
    { label: 'risk / blockers', score: themes.risk },
    { label: 'inspection and follow-through', score: themes.execution },
    { label: 'competitive pressure', score: themes.competition },
    { label: 'expansion potential', score: themes.expansion }
  ].sort(function(a, b) { return b.score - a.score; });
  var themeLabels = themeArray.filter(function(t) { return t.score > 0; }).slice(0, 3).map(function(t) { return t.label; });

  var selectedBigDealAdds = selectEditorialItems_(bigDealAdds, 2);
  var selectedDealProgress = selectEditorialItems_(dealProgress, 3);
  var selectedForecastSignals = selectEditorialItems_(forecastSignals, 3);
  var selectedPriorities = selectEditorialItems_(priorities, 3);
  var selectedRisks = selectEditorialItems_(risks, 3);
  var selectedAsks = selectEditorialItems_(asks, 3);
  var selectedNotes = selectEditorialItems_(forecastNotes, 3);

  var summaryParts = [];
  summaryParts.push('Across ' + latestRecaps.length + ' rep recaps updated in the last 7 days, the average pulse is ' + avgPulse + '/10.');
  if (themeLabels.length) summaryParts.push('This week reads as ' + themeLabels.join(', ') + '.');
  if (selectedBigDealAdds.length) summaryParts.push('The clearest creation signal is ' + renderEditorialSection_('big_deal_adds', selectedBigDealAdds).join(' '));
  if (selectedDealProgress.length) summaryParts.push('The most important commercial movement is ' + renderEditorialSection_('deal_progression', selectedDealProgress).slice(0,2).join(' '));
  if (selectedRisks.length) summaryParts.push('The highest-risk issues are ' + renderEditorialSection_('top_risks', selectedRisks).slice(0,2).join(' '));
  var note = trimSentenceWords_(summaryParts.join(' '), 120);

  return {
    success: true,
    note: note,
    generatedAt: new Date().toISOString(),
    recapsUsed: latestRecaps.length,
    avgPulse: avgPulse,
    themeItems: themeLabels,
    newDealItems: renderEditorialSection_('big_deal_adds', selectedBigDealAdds),
    dealProgressItems: renderEditorialSection_('deal_progression', selectedDealProgress),
    // Forecast signals intentionally omitted in heuristic fallback path to prevent category hallucination.
    // The deterministic forecast signals require context.metrics and context.forecast_deal_context which are not available here.
    forecastItems: [],
    priorityItems: renderEditorialSection_('rep_priorities', selectedPriorities),
    riskItems: renderEditorialSection_('top_risks', selectedRisks),
    askItems: renderEditorialSection_('manager_asks', selectedAsks),
    noteItems: renderEditorialSection_('forecast_notes', selectedNotes)
  };
}

function getTeamRecapNote() {
  ensureAdminCaller_();

  var cached = getJsonCache_("team_recap_note_v2", null);
  if (cached) return cached;

  var now = new Date();
  var daysBack = 7;
  var cutoff = new Date(now.getTime() - daysBack * 86400000);
  var sheet = getRecapSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    var emptyResult = {
      success: true,
      note: 'No team recap submissions were found in the sheet yet.',
      generatedAt: now.toISOString(),
      daysBack: daysBack,
      recapsUsed: 0,
      avgPulse: '—',
      themeItems: [],
      newDealItems: [],
      dealProgressItems: [],
      forecastItems: [],
      priorityItems: [],
      riskItems: [],
      askItems: [],
      noteItems: []
    };
    putJsonCache_("team_recap_note_v2", emptyResult, 45);
    return emptyResult;
  }

  var rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var recentRows = [];
  for (var i = 0; i < rows.length; i++) {
    var ts = rows[i][0];
    var tsDate = ts instanceof Date ? ts : new Date(ts);
    if (!(tsDate instanceof Date) || isNaN(tsDate.getTime())) continue;
    if (tsDate < cutoff) continue;
    recentRows.push(rows[i]);
  }

  if (!recentRows.length) {
    var noRecent = {
      success: true,
      note: 'No team recap submissions were updated in the last 7 days, so there is no current team summary to report.',
      generatedAt: now.toISOString(),
      daysBack: daysBack,
      recapsUsed: 0,
      avgPulse: '—',
      themeItems: [],
      newDealItems: [],
      dealProgressItems: [],
      forecastItems: [],
      priorityItems: [],
      riskItems: [],
      askItems: [],
      noteItems: []
    };
    putJsonCache_("team_recap_note_v2", noRecent, 45);
    return noRecent;
  }

  var byRep = {};
  for (var j = recentRows.length - 1; j >= 0; j--) {
    var email = String(recentRows[j][1] || '').trim().toLowerCase();
    if (!email || byRep[email]) continue;
    byRep[email] = recentRows[j];
  }

  var latestRecaps = [];
  for (var emailKey in byRep) latestRecaps.push(byRep[emailKey]);

  var result = summarizeTeamRecapsEditorial_(latestRecaps);
  result.daysBack = daysBack;
  putJsonCache_("team_recap_note_v2", result, 45);
  return result;
}
function toIsoDateOnly_(value) {
  if (!value) return "";
  var d = value instanceof Date ? value : new Date(value);
  if (!(d instanceof Date) || isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, "America/Chicago", "yyyy-MM-dd");
}

function quarterKeyFromDate_(dateObj) {
  var d = dateObj instanceof Date ? dateObj : new Date(dateObj);
  if (!(d instanceof Date) || isNaN(d.getTime())) d = new Date();
  var year = d.getFullYear();
  var quarter = Math.floor(d.getMonth() / 3) + 1;
  return year + "-Q" + quarter;
}

function quarterLabelFromKey_(quarterKey) {
  var m = String(quarterKey || "").match(/^(\d{4})-Q([1-4])$/);
  if (!m) return String(quarterKey || "");
  return "Q" + m[2] + " " + m[1];
}

function isMeaningfulSummaryAdminText_(value) {
  var s = String(value || "").trim();
  if (!s) return false;
  var normalized = s.toLowerCase();
  var placeholders = [
    "use this section for northcentral rep shout-outs. replace with real names from my team.",
    "keep building top-of-funnel while closing what is in front of us.",
    "weekly summary",
    "confidential - team only"
  ];
  return placeholders.indexOf(normalized) === -1;
}

function findLatestSummaryAdminRow_() {
  var sheet = getOrCreateSummaryAdminSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {
      week_of: toIsoDateOnly_(new Date()),
      leader_note: "",
      shout_outs: "",
      weekly_ask: "",
      theme: "Weekly Summary",
      footer_note: ""
    };
  }
  var rows = sheet.getRange(2, 1, lastRow - 1, SUMMARY_ADMIN_HEADERS.length).getValues();
  var best = null;
  var bestDate = null;
  for (var i = 0; i < rows.length; i++) {
    var d = rows[i][0] instanceof Date ? rows[i][0] : new Date(rows[i][0]);
    if (!(d instanceof Date) || isNaN(d.getTime())) continue;
    if (!bestDate || d > bestDate) {
      bestDate = d;
      best = rows[i];
    }
  }
  if (!best) best = rows[rows.length - 1];

  var latest = {
    week_of: toIsoDateOnly_(best[0]) || toIsoDateOnly_(new Date()),
    leader_note: String(best[1] || "").trim(),
    shout_outs: String(best[2] || "").trim(),
    weekly_ask: String(best[3] || "").trim(),
    theme: String(best[4] || "").trim(),
    footer_note: String(best[5] || "").trim()
  };

  var latestRecapWeek = getLatestRecapWeekOf_();
  if (latestRecapWeek) {
    var adminWeekDate = new Date(String(latest.week_of) + "T12:00:00");
    var recapWeekDate = new Date(String(latestRecapWeek) + "T12:00:00");
    if (!isNaN(recapWeekDate.getTime()) && (isNaN(adminWeekDate.getTime()) || recapWeekDate > adminWeekDate)) {
      latest.week_of = latestRecapWeek;
    }
  }

  var leader = latest.leader_note;
  var shouts = latest.shout_outs;
  var ask = latest.weekly_ask;
  var theme = latest.theme;
  var footer = latest.footer_note;

  for (var j = rows.length - 1; j >= 0; j--) {
    if (!isMeaningfulSummaryAdminText_(leader) && isMeaningfulSummaryAdminText_(rows[j][1])) leader = String(rows[j][1] || "").trim();
    if (!isMeaningfulSummaryAdminText_(shouts) && isMeaningfulSummaryAdminText_(rows[j][2])) shouts = String(rows[j][2] || "").trim();
    if (!isMeaningfulSummaryAdminText_(ask) && isMeaningfulSummaryAdminText_(rows[j][3])) ask = String(rows[j][3] || "").trim();
    if (!isMeaningfulSummaryAdminText_(theme) && isMeaningfulSummaryAdminText_(rows[j][4])) theme = String(rows[j][4] || "").trim();
    if (!isMeaningfulSummaryAdminText_(footer) && isMeaningfulSummaryAdminText_(rows[j][5])) footer = String(rows[j][5] || "").trim();
  }

  latest.leader_note = isMeaningfulSummaryAdminText_(leader) ? leader : "";
  latest.shout_outs = isMeaningfulSummaryAdminText_(shouts) ? shouts : "";
  latest.weekly_ask = isMeaningfulSummaryAdminText_(ask) ? ask : "";
  latest.theme = isMeaningfulSummaryAdminText_(theme) ? theme : (latest.theme || "Weekly Summary");
  latest.footer_note = isMeaningfulSummaryAdminText_(footer) ? footer : "";

  return latest;
}

function getLatestRecapWeekOf_() {
  try {
    var sheet = getRecapSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return "";
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var colWeek = findHeaderIndex_(headers, ["Week Of"]);
    var colTs = findHeaderIndex_(headers, ["Timestamp"]);
    if (colWeek < 0) colWeek = 2;
    if (colTs < 0) colTs = 0;
    var rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var bestWeek = "";
    var bestTs = null;
    for (var i = 0; i < rows.length; i++) {
      var rawTs = rows[i][colTs];
      var ts = rawTs instanceof Date ? rawTs : new Date(rawTs);
      if (!(ts instanceof Date) || isNaN(ts.getTime())) continue;
      var rawWeek = rows[i][colWeek];
      var week = rawWeek instanceof Date ? toIsoDateOnly_(rawWeek) : toIsoDateOnly_(new Date(rawWeek));
      if (!bestTs || ts > bestTs) {
        bestTs = ts;
        bestWeek = toIsoDateOnly_(ts) || week;
      }
    }
    return bestWeek;
  } catch (err) {
    Logger.log("getLatestRecapWeekOf_ failed: " + err.message);
    return "";
  }
}

function getEffectiveReportingDate_(adminWeekOf) {
  var now = new Date();
  var latestRecapWeek = getLatestRecapWeekOf_();
  var adminDate = adminWeekOf ? new Date(String(adminWeekOf) + "T12:00:00") : null;
  var recapDate = latestRecapWeek ? new Date(String(latestRecapWeek) + "T12:00:00") : null;

  if (recapDate instanceof Date && !isNaN(recapDate.getTime())) {
    if (!(adminDate instanceof Date) || isNaN(adminDate.getTime()) || recapDate > adminDate) return recapDate;
  }
  if (adminDate instanceof Date && !isNaN(adminDate.getTime())) return adminDate;
  return now;
}

function findQuotaConfigForQuarter_(quarterKey) {
  var sheet = getOrCreateQuotaConfigSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {
      quarter_key: quarterKey,
      quarter_label: quarterLabelFromKey_(quarterKey),
      team_name: SUMMARY_TEAM_NAME,
      team_quota: 0,
      closed_to_date: 0,
      manager_name: "Billy Schuett",
      notes: ""
    };
  }
  var rows = sheet.getRange(2, 1, lastRow - 1, QUOTA_CONFIG_HEADERS.length).getValues();
  var fallback = null;
  for (var i = 0; i < rows.length; i++) {
    var quarter = String(rows[i][0] || "").trim();
    var teamName = String(rows[i][2] || "").trim();
    var rec = {
      quarter_key: quarter,
      quarter_label: String(rows[i][1] || "").trim() || quarterLabelFromKey_(quarter),
      team_name: teamName || SUMMARY_TEAM_NAME,
      team_quota: Number(rows[i][3]) || 0,
      closed_to_date: Number(rows[i][4]) || 0,
      manager_name: String(rows[i][5] || "").trim() || "Billy Schuett",
      notes: String(rows[i][6] || "").trim()
    };
    if (quarter === quarterKey && (!teamName || teamName === SUMMARY_TEAM_NAME)) return rec;
    if (quarter === quarterKey && !fallback) fallback = rec;
  }
  if (fallback) return fallback;
  return {
    quarter_key: quarterKey,
    quarter_label: quarterLabelFromKey_(quarterKey),
    team_name: SUMMARY_TEAM_NAME,
    team_quota: 0,
    closed_to_date: 0,
    manager_name: "Billy Schuett",
    notes: ""
  };
}

function findClosedWonReportSheet_() {
  return findReportSheetByNameOrHeader_(CLOSED_WON_REPORT_SHEET_NAME);
}

function findReportSheetByNameOrHeader_(preferredName) {
  var ss = SpreadsheetApp.openById(TRACKER_SHEET_ID);
  var byName = ss.getSheetByName(preferredName);
  if (byName) return byName;

  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    if (sh.getLastRow() < 2 || sh.getLastColumn() < 6) continue;
    var header = sh.getRange(2, 1, 1, Math.min(sh.getLastColumn(), 6)).getValues()[0];
    var h0 = String(header[0] || '').trim().toLowerCase();
    var h1 = String(header[1] || '').trim().toLowerCase();
    var h2 = String(header[2] || '').trim().toLowerCase();
    var h3 = String(header[3] || '').trim().toLowerCase();
    var h4 = String(header[4] || '').trim().toLowerCase();
    var h5 = String(header[5] || '').trim().toLowerCase();
    if (
      h0 === 'opportunity name' &&
      h1 === 'account name' &&
      h2 === 'amount' &&
      h3 === 'close date' &&
      h4 === 'stage' &&
      h5 === 'opportunity owner'
    ) {
      return sh;
    }
  }
  return null;
}

function getClosedWonSourceForQuarter_(quarterKey, fallbackValue) {
  try {
    var sheet = findClosedWonReportSheet_();
    if (!sheet || sheet.getLastRow() < 3) {
      return {
        amount: Number(fallbackValue) || 0,
        source: "quota_config_fallback",
        label: "Quota_Config fallback",
        usedFallback: true,
        warning: "Closed Won Salesforce report sheet not found or empty."
      };
    }

    var data = sheet.getRange(3, 1, sheet.getLastRow() - 2, Math.min(sheet.getLastColumn(), 6)).getValues();
    var total = 0;

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var amount = Number(row[2]) || 0;
      var closeRaw = row[3];
      var stage = String(row[4] || "").trim().toLowerCase();
      var closeDate = closeRaw instanceof Date ? closeRaw : new Date(closeRaw);

      if (!amount) continue;
      if (stage.indexOf("closed won") !== 0) continue;
      if (!(closeDate instanceof Date) || isNaN(closeDate.getTime())) continue;
      if (quarterKeyFromDate_(closeDate) !== quarterKey) continue;

      total += amount;
    }

    return {
      amount: total,
      source: "salesforce_closed_won_report",
      label: sheet.getName(),
      usedFallback: false,
      warning: ""
    };
  } catch (err) {
    Logger.log("getClosedWonSourceForQuarter_ failed: " + err.message);
    return {
      amount: Number(fallbackValue) || 0,
      source: "quota_config_fallback",
      label: "Quota_Config fallback",
      usedFallback: true,
      warning: "Closed Won source failed: " + err.message
    };
  }
}

// Lightweight quota + closed won for Dashboard/POC pages
function getQuotaMetricsForDashboard_() {
  try {
    var quarterKey = quarterKeyFromDate_(new Date());
    var quotaCfg = findQuotaConfigForQuarter_(quarterKey);
    var closedMeta = getClosedWonSourceForQuarter_(quarterKey, quotaCfg.closed_to_date);
    var closed = Number(closedMeta.amount) || 0;
    var quota = Number(quotaCfg.team_quota) || 0;
    return {
      teamQuota: quota,
      closed: closed,
      attainmentPct: pctOfQuota_(closed, quota),
      quarterLabel: quotaCfg.quarter_label || quarterLabelFromKey_(quarterKey),
      closedSource: closedMeta
    };
  } catch (err) {
    Logger.log("getQuotaMetricsForDashboard_ failed: " + err.message);
    return {
      teamQuota: 0,
      closed: 0,
      attainmentPct: 0,
      quarterLabel: "",
      closedSource: {
        amount: 0,
        source: "error",
        label: "Unavailable",
        usedFallback: true,
        warning: err.message
      }
    };
  }
}

function pctOfQuota_(value, quota) {
  if (!(quota > 0)) return 0;
  return (Number(value) || 0) / quota;
}

function buildExecutiveTeamRollup_(teamData, teamRecap) {
  var avgPulse = 0;
  var pulseCount = 0;
  var green = 0, yellow = 0, red = 0;
  var risks = [];
  var asks = [];
  var notes = [];
  for (var i = 0; i < teamData.length; i++) {
    var rep = teamData[i];
    var p = Number(rep.pulseScore) || 0;
    if (p > 0) {
      avgPulse += p;
      pulseCount++;
    }
    green += Number(rep.greenAccounts) || 0;
    yellow += Number(rep.yellowAccounts) || 0;
    red += Number(rep.redAccounts) || 0;
    if (rep.risk) risks.push((rep.name || rep.email) + ": " + cleanSnippet_(rep.risk, 90));
    if (rep.ask) asks.push((rep.name || rep.email) + ": " + cleanSnippet_(rep.ask, 90));
    if (rep.forecast_note) notes.push((rep.name || rep.email) + ": " + cleanSnippet_(rep.forecast_note, 90));
  }
  var recapCount = Number(teamRecap && teamRecap.recapsUsed) || 0;
  var recapAvgPulse = (teamRecap && teamRecap.avgPulse) ? String(teamRecap.avgPulse) : "";
  return {
    repsSubmitted: recapCount > 0 ? recapCount : teamData.length,
    avgPulse: recapAvgPulse || (pulseCount ? (avgPulse / pulseCount).toFixed(1) : "—"),
    greenAccounts: green,
    yellowAccounts: yellow,
    redAccounts: red,
    themes: (teamRecap && teamRecap.themeItems) ? teamRecap.themeItems : [],
    risks: (teamRecap && teamRecap.riskItems && teamRecap.riskItems.length) ? teamRecap.riskItems : risks.slice(0, 4),
    asks: (teamRecap && teamRecap.askItems && teamRecap.askItems.length) ? teamRecap.askItems : asks.slice(0, 4),
    notes: (teamRecap && teamRecap.noteItems && teamRecap.noteItems.length) ? teamRecap.noteItems : notes.slice(0, 4),
    priorities: (teamRecap && teamRecap.priorityItems) ? teamRecap.priorityItems : []
  };
}

function getExecutiveSummaryDataBase_(forcedQuarterKey) {
  ensureAdminCaller_();

  var admin = findLatestSummaryAdminRow_();
  var reportingDate = getEffectiveReportingDate_(admin.week_of);
  var quarterKey = forcedQuarterKey || quarterKeyFromDate_(reportingDate);
  var quotaCfg = findQuotaConfigForQuarter_(quarterKey);
  var teamData = getTeamDashboardData();
  var teamRecap = getTeamRecapNote();
  var closedMeta = getClosedWonSourceForQuarter_(quarterKey, quotaCfg.closed_to_date);
  var closedAmount = Number(closedMeta.amount) || 0;

  var repFallbackTotals = {
    closed: closedAmount,
    commit: 0,
    likely: 0,
    bestCase: 0,
    nextQuarter: 0
  };

  for (var i = 0; i < teamData.length; i++) {
    repFallbackTotals.commit += parseMoneyishServer_(teamData[i].commit);
    repFallbackTotals.likely += parseMoneyishServer_(teamData[i].likely);
    repFallbackTotals.bestCase += parseMoneyishServer_(teamData[i].upside);
    repFallbackTotals.nextQuarter += parseMoneyishServer_(teamData[i].nq_commit);
  }

  var nextQuarterKey = quarterKeyFromDate_(new Date(reportingDate.getFullYear(), reportingDate.getMonth() + 3, reportingDate.getDate()));
  var nextQuarterQuotaCfg = findQuotaConfigForQuarter_(nextQuarterKey);
  var nextQuarterQuota = Number(nextQuarterQuotaCfg.team_quota) || 0;
  var forecastTotals = getTeamForecastTotalsFromDealsSheet_(quarterKey, nextQuarterKey, repFallbackTotals);
  var commitMeta = forecastTotals.sources.commit;
  var likelyMeta = forecastTotals.sources.likely;
  var bestCaseMeta = forecastTotals.sources.bestCase;
  var nextQuarterMeta = forecastTotals.sources.nextQuarter;
  var nextQuarterOverride = Number(SUMMARY_NEXT_QUARTER_OVERRIDE || 0);
  if (nextQuarterOverride > 0) {
    nextQuarterMeta = {
      source: 'manual_override',
      label: 'Temporary manual next-quarter override',
      usedFallback: false,
      warning: ''
    };
  }

  var totals = {
    closed: Number(forecastTotals.closed) || closedAmount,
    commit: Number(forecastTotals.commit) || 0,
    likely: Number(forecastTotals.likely) || 0,
    bestCase: Number(forecastTotals.bestCase) || 0,
    nextQuarter: nextQuarterOverride > 0 ? nextQuarterOverride : (Number(forecastTotals.nextQuarter) || 0)
  };

  var quota = Number(quotaCfg.team_quota) || 0;
  var rollup = buildExecutiveTeamRollup_(teamData, teamRecap);
  var nbmSummary = getSummaryNBMData_(reportingDate);
  var rainmakerSummary = getSummaryRainmakerData_(teamData, reportingDate, nbmSummary);

  return {
    generatedAt: new Date().toISOString(),
    teamName: quotaCfg.team_name || SUMMARY_TEAM_NAME,
    reportingWeek: toIsoDateOnly_(reportingDate),
    quarterKey: quarterKey,
    quarterLabel: quotaCfg.quarter_label || quarterLabelFromKey_(quarterKey),
    managerName: quotaCfg.manager_name || "Billy Schuett",
    theme: admin.theme || "Weekly Summary",
    footerNote: admin.footer_note || "",
    leaderNote: admin.leader_note || "",
    shoutOuts: admin.shout_outs || "",
    weeklyAsk: admin.weekly_ask || "",
    quotaNotes: quotaCfg.notes || "",
    adminNotes: admin.footer_note || "",
    diagnostics: {
      closedSource: forecastTotals.sources.closed && !forecastTotals.sources.closed.usedFallback ? forecastTotals.sources.closed : closedMeta,
      commitSource: commitMeta,
      mostLikelySource: likelyMeta,
      bestCaseSource: bestCaseMeta,
      nextQuarterSource: nextQuarterMeta,
      quotaConfigQuarter: quarterKey,
      reportingWeek: toIsoDateOnly_(reportingDate),
      recapRowsUsed: Number(rollup && rollup.repsSubmitted || 0),
      nbmDateSource: nbmSummary.dateSource || ""
    },
    metrics: {
      teamQuota: quota,
      closed: totals.closed,
      commit: totals.commit,
      mostLikely: totals.likely,
      bestCase: totals.bestCase,
      nextQuarter: totals.nextQuarter,
      nextQuarterQuota: nextQuarterQuota,
      nextQuarterCoverageX: nextQuarterQuota > 0 ? ((Number(totals.nextQuarter) || 0) / nextQuarterQuota) : 0,
      attainmentPct: pctOfQuota_(totals.closed, quota),
      commitPct: pctOfQuota_(totals.commit, quota),
      likelyPct: pctOfQuota_(totals.likely, quota),
      bestCasePct: pctOfQuota_(totals.bestCase, quota),
      closedGap: quota - totals.closed,
      commitGap: quota - totals.commit,
      likelyGap: quota - totals.likely,
      bestCaseGap: quota - totals.bestCase
    },
    nbmSummary: nbmSummary,
    rainmakerSummary: rainmakerSummary,
    teamRollup: rollup,
    teamRecap: teamRecap
  };
}

function getSummaryNBMData_(reportingDate) {
  var out = {
    completedQtd: [],
    scheduledThisWeek: [],
    weekLabel: "",
    dateSource: ""
  };

  try {
    var ss = SpreadsheetApp.openById(TRACKER_SHEET_ID);
    var sheet = ss.getSheetByName(NBM_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return out;

    var weekStart = reportingDate instanceof Date && !isNaN(reportingDate.getTime())
      ? new Date(reportingDate.getFullYear(), reportingDate.getMonth(), reportingDate.getDate())
      : new Date();
    var weekEnd = new Date(weekStart.getTime() + 6 * 86400000);
    out.weekLabel = Utilities.formatDate(weekStart, Session.getScriptTimeZone(), "MMM d")
      + "–"
      + Utilities.formatDate(weekEnd, Session.getScriptTimeZone(), "MMM d");

    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return out;

    var recapScheduled = getLatestRecapNBMScheduled_();
    if (recapScheduled.length) {
      out.scheduledThisWeek = recapScheduled;
      out.dateSource = "latest_recap_field";
    }

    var headerRowIdx = -1;
    var ownerIdx = 0;
    var dateIdx = -1;
    var accountIdx = -1;

    for (var hr = 0; hr < Math.min(5, data.length); hr++) {
      var header = data[hr];
      var ownerCandidate = findHeaderIndex_(header, ["Owner", "Opportunity Owner", "AE", "Rep", "Meeting Owner"]);
      if (ownerCandidate === -1) continue;
      headerRowIdx = hr;
      ownerIdx = ownerCandidate;
      dateIdx = findHeaderIndex_(header, ["NBM Date", "Meeting Date", "Date", "Scheduled Date", "Start Date"]);
      accountIdx = findHeaderIndex_(header, ["Account", "Account Name", "Company", "Account / Opportunity"]);
      break;
    }

    var startRow = headerRowIdx === -1 ? 2 : (headerRowIdx + 2);
    var counts = {};
    var scheduled = [];

    function addCount_(repName) {
      counts[repName] = (counts[repName] || 0) + 1;
    }

    for (var r = startRow - 1; r < data.length; r++) {
      var row = data[r];
      var owner = String(row[ownerIdx] || "").trim();
      if (!owner) continue;

      var matchedRep = null;
      for (var i = 0; i < REP_EMAILS.length; i++) {
        var candidates = getUserNameCandidatesFromEmail(REP_EMAILS[i]);
        if (ownerMatchesAnyRepName_(owner, candidates)) {
          matchedRep = getUserNameFromEmail(REP_EMAILS[i]);
          break;
        }
      }
      if (!matchedRep) continue;

      addCount_(matchedRep);

      if (dateIdx !== -1) {
        var rawDate = row[dateIdx];
        var d = rawDate instanceof Date ? rawDate : new Date(rawDate);
        if (d instanceof Date && !isNaN(d.getTime())) {
          var day = new Date(d.getFullYear(), d.getMonth(), d.getDate());
          if (day >= weekStart && day <= weekEnd) {
            var acct = accountIdx !== -1 ? String(row[accountIdx] || "").trim() : "";
            scheduled.push({
              rep: matchedRep,
              date: day,
              account: acct
            });
          }
        }
      }
    }

    for (var j = 0; j < REP_EMAILS.length; j++) {
      var repName = getUserNameFromEmail(REP_EMAILS[j]);
      out.completedQtd.push({
        rep: repName,
        count: Number(counts[repName] || 0)
      });
    }

    if (!out.scheduledThisWeek.length) {
      out.scheduledThisWeek = scheduled.sort(function(a, b) {
        return a.date.getTime() - b.date.getTime();
      }).map(function(item) {
        var dateLabel = Utilities.formatDate(item.date, Session.getScriptTimeZone(), "EEE M/d");
        return dateLabel + " — " + item.rep + (item.account ? " — " + item.account : "");
      });
      out.dateSource = dateIdx !== -1 ? "nbm_sheet_date_column" : "none";
    }
    return out;
  } catch (err) {
    Logger.log("getSummaryNBMData_ failed: " + err.message);
    return out;
  }
}

function getLatestRecapNBMScheduled_() {
  var sheet = getRecapSheet_();
  if (!sheet || sheet.getLastRow() < 2) return [];

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var latestByEmail = {};

  for (var i = rows.length - 1; i >= 0; i--) {
    var email = String(rows[i][1] || "").trim().toLowerCase();
    if (!email || latestByEmail[email]) continue;
    latestByEmail[email] = parseRecapRow_(rows[i]);
  }

  var out = [];
  for (var r = 0; r < REP_EMAILS.length; r++) {
    var repEmail = REP_EMAILS[r].toLowerCase().trim();
    var parsed = latestByEmail[repEmail];
    if (!parsed) continue;
    var note = String(parsed.nbm_scheduled_week || "").trim();
    var scheduledCount = Number(parsed.nbm_scheduled_count) || 0;
    if (note) {
      if (scheduledCount > 0) out.push(getUserNameFromEmail(repEmail) + ": " + scheduledCount + " — " + note);
      else out.push(getUserNameFromEmail(repEmail) + ": " + note);
      continue;
    }
    if (scheduledCount > 0) {
      out.push(getUserNameFromEmail(repEmail) + ": " + scheduledCount + " NBM" + (scheduledCount === 1 ? "" : "s") + " scheduled this week");
      continue;
    }
    var rmNbm = Number(parsed.rm_nbm) || 0;
    if (rmNbm > 0) {
      out.push(getUserNameFromEmail(repEmail) + ": " + rmNbm + " NBM" + (rmNbm === 1 ? "" : "s") + " logged this week");
    }
  }
  return out;
}

function getSummaryRainmakerData_(teamData, reportingDate, nbmSummary) {
  var out = {
    weekLabel: "",
    discovery: [],
    nbms: [],
    channelMeetings: [],
    totals: { discovery: 0, nbms: 0, channelMeetings: 0 }
  };

  var weekStart = reportingDate instanceof Date && !isNaN(reportingDate.getTime())
    ? new Date(reportingDate.getFullYear(), reportingDate.getMonth(), reportingDate.getDate())
    : new Date();
  var weekEnd = new Date(weekStart.getTime() + 6 * 86400000);
  out.weekLabel = Utilities.formatDate(weekStart, Session.getScriptTimeZone(), "MMM d")
    + "–"
    + Utilities.formatDate(weekEnd, Session.getScriptTimeZone(), "MMM d");

  var currentNbmByRep = {};
  var completedQtd = (nbmSummary && Array.isArray(nbmSummary.completedQtd)) ? nbmSummary.completedQtd : [];
  for (var c = 0; c < completedQtd.length; c++) {
    var completedItem = completedQtd[c] || {};
    var completedRep = String(completedItem.rep || "").trim();
    if (!completedRep) continue;
    currentNbmByRep[completedRep] = Number(completedItem.count) || 0;
  }

  var reps = Array.isArray(teamData) ? teamData : [];
  for (var i = 0; i < reps.length; i++) {
    var rep = reps[i] || {};
    var repName = String(rep.name || rep.email || "").trim();
    if (!repName) continue;

    var disco = Number(rep.rm_disco) || 0;
    var nbm = Number(rep.rm_nbm) || 0;
    var channel = Number(rep.rm_accts) || 0;
    var nbmCustomer = String(rep.nbm_scheduled_week || "").trim();
    var currentNbm = Number(currentNbmByRep[repName] || 0);

    out.totals.discovery += disco;
    out.totals.nbms += nbm;
    out.totals.channelMeetings += channel;

    out.discovery.push({ rep: repName, count: disco });
    out.nbms.push({ rep: repName, count: nbm, note: nbmCustomer, currentCount: currentNbm });
    out.channelMeetings.push({ rep: repName, count: channel });
  }

  return out;
}

function getExecutiveSummaryData() {
  var base = getExecutiveSummaryDataBase_();
  function mergeSectionItems_(primary, fallbackA, fallbackB, maxItems) {
    var out = [];
    var seen = {};
    function addAll_(list) {
      list = Array.isArray(list) ? list : [];
      for (var i = 0; i < list.length && out.length < maxItems; i++) {
        var item = String(list[i] || '').trim();
        var key = item.toLowerCase();
        if (!item || seen[key]) continue;
        seen[key] = true;
        out.push(item);
      }
    }
    addAll_(primary);
    addAll_(fallbackA);
    addAll_(fallbackB);
    return out.slice(0, maxItems);
  }
  var ai = getJsonCache_("summary_ai_" + String(base.quarterKey || ""), null) || getLatestSummaryAISnapshot_(base.quarterKey);
  var fallbackRecap = base.teamRecap || {};
  var fallbackRollup = base.teamRollup || {};
  if (ai) {
    var normalizedAi = ai;
    try {
      if (typeof buildSummaryAiInputContext_ === 'function' && typeof validateSummaryAIPayload_ === 'function') {
        normalizedAi = validateSummaryAIPayload_(ai, buildSummaryAiInputContext_(base.quarterKey));
      }
    } catch (err) {
      Logger.log("AI payload normalization failed, using raw payload: " + err.message);
      normalizedAi = ai;
    }
    base.leaderNote = normalizedAi.leader_note || base.leaderNote || "";
    if (normalizedAi.themes && normalizedAi.themes.length) base.teamRollup.themes = normalizedAi.themes;
    var combinedRAN = normalizedAi.risks_asks_notes || [].concat(normalizedAi.top_risks || [], normalizedAi.manager_asks || [], normalizedAi.forecast_notes || []).slice(0, 5);
    base.teamRollup.priorities = mergeSectionItems_(normalizedAi.rep_priorities || [], fallbackRollup.priorities || [], fallbackRecap.priorityItems || [], 4);
    base.teamRollup.risks = normalizedAi.top_risks || [];
    base.teamRollup.asks = normalizedAi.manager_asks || [];
    base.teamRollup.notes = normalizedAi.forecast_notes || [];
    base.teamRollup.risksAsksNotes = combinedRAN;
    base.teamRecap = {
      success: true,
      note: "",
      generatedAt: normalizedAi.generated_at || "",
      recapsUsed: base.teamRollup.repsSubmitted || 0,
      avgPulse: base.teamRollup.avgPulse || "—",
      themeItems: normalizedAi.themes || [],
      newDealItems: normalizedAi.big_deal_adds || [],
      dealProgressItems: mergeSectionItems_(normalizedAi.deal_progression || [], fallbackRecap.dealProgressItems || [], normalizedAi.big_deal_adds || [], 4),
      forecastItems: mergeSectionItems_(normalizedAi.forecast_signals || [], fallbackRecap.forecastItems || [], [], 4),
      priorityItems: mergeSectionItems_(normalizedAi.rep_priorities || [], fallbackRollup.priorities || [], fallbackRecap.priorityItems || [], 4),
      riskItems: normalizedAi.top_risks || [],
      askItems: normalizedAi.manager_asks || [],
      noteItems: normalizedAi.forecast_notes || [],
      risksAsksNotesItems: combinedRAN,
      source: "openai"
    };
    base.diagnostics.summarySource = "openai";
    base.diagnostics.summaryGeneratedAt = normalizedAi.generated_at || "";
  } else {
    base.diagnostics.summarySource = "heuristic_fallback";
  }
  return base;
}

function getRepDataForAdmin(email) {
  var caller = getResolvedCallerEmail_();
  Logger.log("getRepDataForAdmin caller=" + caller);
  // Allow admin or empty string (web app context)
  if (caller && caller !== String(ADMIN_EMAIL || "").trim().toLowerCase()) return null;
  return JSON.stringify({
    lastData: getLastDataForUser(email),
    latestForecast: getLatestSelfForecastForUser(email),
    userStats: getUserStats(email),
    sfData: getSFDataForUser(email),
    qbrCall: getQBRCall(email)
  });
}

function getRepDataForRecap(email) {
  var target = String(email || "").trim().toLowerCase();
  if (!target) return null;

  var caller = getResolvedCallerEmail_();
  var admin = String(ADMIN_EMAIL || "").trim().toLowerCase();

  // Allow: admin, same-user, or empty caller from web app context.
  if (caller && caller !== admin && caller !== target) return null;

  return JSON.stringify({
    lastData: getLastDataForUser(target),
    userStats: getUserStats(target),
    sfData: getSFDataForUser(target),
    qbrCall: getQBRCall(target),
    latestForecast: getLatestSelfForecastForUser(target)
  });
}

function testNikoData() {
  var result = getLastDataForUser("niko.argaez@glean.com");
  Logger.log(JSON.stringify(result));
  return result;
}

function getSfQbrSanityReportForAllReps() {
  ensureAdminCaller_();
  var report = [];
  for (var i = 0; i < REP_EMAILS.length; i++) {
    var email = String(REP_EMAILS[i] || "").toLowerCase().trim();
    if (!email) continue;
    var name = getUserNameFromEmail(email) || email;
    var sf = getSFDataForUser(email) || {};
    var qbr = getQBRCall(email) || {};
    var sfTotal = Number(sf.commit || 0) + Number(sf.mostLikely || 0) + Number(sf.bestCase || 0);
    var qbrTotal = Number(qbr.commit || 0) + Number(qbr.mostLikely || 0) + Number(qbr.bestCase || 0);
    report.push({
      email: email,
      name: name,
      sfCommit: Number(sf.commit || 0),
      sfMostLikely: Number(sf.mostLikely || 0),
      sfBestCase: Number(sf.bestCase || 0),
      qbrCommit: Number(qbr.commit || 0),
      qbrMostLikely: Number(qbr.mostLikely || 0),
      qbrBestCase: Number(qbr.bestCase || 0),
      sfHasData: sfTotal > 0,
      qbrHasData: qbrTotal > 0
    });
  }
  Logger.log("SF/QBR sanity report: " + JSON.stringify(report));
  return report;
}
