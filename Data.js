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
    sf_accurate: r[55]
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
    "weekOf", "pulseReason", "rm_disco", "rm_nbm", "rm_opp", "rm_accts",
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
  var ss = SpreadsheetApp.openById("1tK7hslM--NY0fU6z7zvDwjw3K63DOTayAyQ0KE20J14");
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
    engagementSummary, forecastSummary, formObject.sf_accurate
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

function getMainSheetSnapshotForCodex() {
  var caller = (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || "").toLowerCase();
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
  var caller = (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || "").toLowerCase();
  if (caller !== String(ADMIN_EMAIL || "").toLowerCase()) throw new Error("Admin access required.");
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

function getSFDataForUser(email) {
  var blank = { commit: 0, mostLikely: 0, bestCase: 0, deals: [], nbmBlocks: [], nbmWarnings: [] };
  var name = getUserNameFromEmail(email);
  if (!name) return blank;
  var ss = SpreadsheetApp.openById("1tK7hslM--NY0fU6z7zvDwjw3K63DOTayAyQ0KE20J14");
  var dealsSheet = ss.getSheetByName(DEALS_SHEET_NAME);
  if (!dealsSheet || dealsSheet.getLastRow() < 3) {
    Logger.log("getSFDataForUser: sheet missing or empty: " + DEALS_SHEET_NAME);
    return blank;
  }
  var startRow = 3;
  var numRows = dealsSheet.getLastRow() - 2;
  var numCols = Math.max(DEALS_COL_OWNER, DEALS_COL_FORECAST, DEALS_COL_AMOUNT, DEALS_COL_CLOSE, DEALS_COL_ACCOUNT);
  var dealsValues = dealsSheet.getRange(startRow, 1, numRows, numCols).getValues();
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
  var nameStr = String(name || "");
  for (var r = 0; r < dealsValues.length; r++) {
    var row = dealsValues[r];
    var owner    = (row[DEALS_COL_OWNER - 1] || "").toString().trim();
    var account  = (row[DEALS_COL_ACCOUNT - 1] || "").toString().trim();
    var forecast = (row[DEALS_COL_FORECAST - 1] || "").toString().trim();
    var amount   = row[DEALS_COL_AMOUNT - 1];
    var closeDate = row[DEALS_COL_CLOSE - 1];
    if (!owner || !account || !forecast || !amount) continue;
    if (!ownerMatchesRep_(owner, nameStr)) continue;
    var amt = typeof amount === "number"
      ? amount
      : parseFloat((amount || "0").toString().replace(/[^0-9.]/g, "")) || 0;
    if (!amt) continue;
    var forecastStr  = forecast.toLowerCase();
    var isCommit     = forecastStr.indexOf("commit") !== -1;
    var isMostLikely = forecastStr.indexOf("most likely") !== -1 || forecastStr.indexOf("most_likely") !== -1;
    var isBestCase   = forecastStr.indexOf("best case") !== -1 || forecastStr.indexOf("best_case") !== -1;
    if (!isCommit && !isMostLikely && !isBestCase) continue;
    var hasNBM = !!nbmKeysSet[account.toLowerCase()];
    deals.push({ category: forecast, account: account, amount: amt, closeDate: closeDate, hasNBM: hasNBM, owner: owner });
    if (isCommit)          { commitTotal += amt; mlTotal += amt; bcTotal += amt; }
    else if (isMostLikely) { mlTotal += amt; bcTotal += amt; }
    else if (isBestCase)   { bcTotal += amt; }
    if (amt >= NBM_MIN_AMOUNT && !hasNBM) {
      var tag = account + " \u2014 $" + Math.round(amt / 1000) + "k";
      if (isCommit)          nbmBlocks.push(tag + " \u2014 Commit \u2014 No NBM logged");
      else if (isMostLikely) nbmWarnings.push(tag + " \u2014 Most Likely \u2014 No NBM scheduled");
    }
  }
  Logger.log("getSFDataForUser [" + normNameToken_(nameStr) + "]: deals=" + deals.length + " commit=$" + commitTotal);
  return { commit: commitTotal, mostLikely: mlTotal, bestCase: bcTotal, deals: deals, nbmBlocks: nbmBlocks, nbmWarnings: nbmWarnings };
}

// -----------------------------------------------------------------------
// QBR CALL
// -----------------------------------------------------------------------
function getQBRCall(email) {
  var blank = { commit: 0, mostLikely: 0, bestCase: 0 };
  var name = getUserNameFromEmail(email);
  if (!name) return blank;
  var sheet = SpreadsheetApp.openById("1tK7hslM--NY0fU6z7zvDwjw3K63DOTayAyQ0KE20J14").getSheetByName("QBR_Call");
  if (!sheet || sheet.getLastRow() < 2) return blank;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  function parseAmt(v) {
    if (!v) return 0;
    if (typeof v === "number") return v;
    return parseFloat(v.toString().replace(/[^0-9.]/g, "")) || 0;
  }
  for (var i = 0; i < data.length; i++) {
    if (ownerMatchesRep_(data[i][0], name)) {
      return { commit: parseAmt(data[i][1]), mostLikely: parseAmt(data[i][2]), bestCase: parseAmt(data[i][3]) };
    }
  }
  return blank;
}

// -----------------------------------------------------------------------
// TEAM DASHBOARD DATA — lean, no SF/stats calls (those load async)
// -----------------------------------------------------------------------
function getTeamDashboardData() {
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
      var _meaningful = _r[2] || _r[4] || Number(_r[5]) > 0 || Number(_r[6]) > 0 || _r[7] || _r[46] || _r[47] || _r[48] || _r[50] || _r[51] || _r[52];
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
  return teamData;
}

function getRepDataForAdmin(email) {
  var caller = "";
  try { caller = Session.getEffectiveUser().getEmail(); } catch (_) {}
  if (!caller) {
    try { caller = Session.getActiveUser().getEmail(); } catch (_) {}
  }
  caller = String(caller || "").trim().toLowerCase();
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

  var callerRaw = "";
  try { callerRaw = Session.getEffectiveUser().getEmail(); } catch (_) {}
  if (!callerRaw) {
    try { callerRaw = Session.getActiveUser().getEmail(); } catch (_) {}
  }
  var caller = String(callerRaw || "")
    .trim()
    .toLowerCase();
  var admin = String(ADMIN_EMAIL || "").trim().toLowerCase();

  // Allow: admin, same-user, or empty caller from web app context.
  if (caller && caller !== admin && caller !== target) return null;

  return JSON.stringify({
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
