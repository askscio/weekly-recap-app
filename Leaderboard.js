// -----------------------------------------------------------------------
// LEADERBOARD HELPERS
// -----------------------------------------------------------------------
var _repStatsCache_ = null;

function aggregateByRep_(sheetName, ownerCol, amtCol, mode) {
  try {
    var ss = SpreadsheetApp.openById("1tK7hslM--NY0fU6z7zvDwjw3K63DOTayAyQ0KE20J14");
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) {
      Logger.log("aggregateByRep_: sheet missing or empty: " + sheetName);
      return [];
    }
    var data = sheet.getDataRange().getValues();
    var totals = {};

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var owner = (row[ownerCol] || "").toString().trim();
      if (!owner) continue;
      if (mode === "count") {
        totals[owner] = (totals[owner] || 0) + 1;
      } else {
        var amt = typeof row[amtCol] === "number"
          ? row[amtCol]
          : parseFloat((row[amtCol] || "0").toString().replace(/[^0-9.]/g, "")) || 0;
        totals[owner] = (totals[owner] || 0) + amt;
      }
    }

    var result = [];
    for (var name in totals) result.push({ name: name, value: totals[name] });
    result.sort(function(a, b) { return b.value - a.value; });
    return result;
  } catch (e) {
    Logger.log("aggregateByRep_ failed for " + sheetName + ": " + e.message);
    return [];
  }
}

function getRankAndValue_(sortedList, repNameOrNames) {
  var repNames = Array.isArray(repNameOrNames) ? repNameOrNames : [repNameOrNames];
  var exactNames = repNames.map(function(n) { return String(n || "").toLowerCase().trim(); }).filter(Boolean);
  for (var i = 0; i < sortedList.length; i++) {
    var rowName = String(sortedList[i].name || "").toLowerCase().trim();
    for (var j = 0; j < exactNames.length; j++) {
      if (rowName === exactNames[j]) return { rank: i + 1, amt: sortedList[i].value };
    }
  }
  for (var i = 0; i < sortedList.length; i++) {
    if (typeof ownerMatchesAnyRepName_ === "function" && ownerMatchesAnyRepName_(sortedList[i].name, repNames)) {
      return { rank: i + 1, amt: sortedList[i].value };
    }
  }
  var firstNames = exactNames.map(function(n) { return n.split(" ")[0]; }).filter(Boolean);
  for (var i = 0; i < sortedList.length; i++) {
    var sn = String(sortedList[i].name || "").toLowerCase().trim();
    var snFirst = sn.split(" ")[0];
    for (var k = 0; k < firstNames.length; k++) {
      if (sn === firstNames[k] || snFirst === firstNames[k]) return { rank: i + 1, amt: sortedList[i].value };
    }
  }
  return { rank: "-", amt: 0 };
}

function getTopValue_(sortedList, position) {
  if (!sortedList || sortedList.length === 0) return 0;
  var idx = Math.min(position - 1, sortedList.length - 1);
  return sortedList[idx].value;
}

function _getRepStats_() {
  if (_repStatsCache_) return _repStatsCache_;
  _repStatsCache_ = {
    disco: aggregateByRep_(DISCO_SHEET_NAME, 0, null, "count"),
    nbm: aggregateByRep_(NBM_SHEET_NAME, 0, null, "count"),
    pipe: aggregateByRep_(PIPELINE_SHEET_NAME, 0, 1, "sum"),
    stg4: aggregateByRep_(STAGE4_SHEET_NAME, 0, 1, "sum")
  };
  Logger.log("_getRepStats_: disco=" + _repStatsCache_.disco.length
    + " nbm=" + _repStatsCache_.nbm.length
    + " pipe=" + _repStatsCache_.pipe.length
    + " stg4=" + _repStatsCache_.stg4.length);
  return _repStatsCache_;
}

// -----------------------------------------------------------------------
// BENCHMARKS (from Coefficient tabs)
// -----------------------------------------------------------------------
function getBenchmarks() {
  var stats = _getRepStats_();
  return {
    disco_5: getTopValue_(stats.disco, 5),
    disco_20: getTopValue_(stats.disco, 20),
    nbm_5: getTopValue_(stats.nbm, 5),
    nbm_20: getTopValue_(stats.nbm, 20),
    pipe_5: getTopValue_(stats.pipe, 5),
    pipe_20: getTopValue_(stats.pipe, 20),
    stg4_5: getTopValue_(stats.stg4, 5),
    stg4_20: getTopValue_(stats.stg4, 20)
  };
}

// -----------------------------------------------------------------------
// USER STATS (from Coefficient tabs)
// -----------------------------------------------------------------------
function getUserStats(email) {
  var names = (typeof getUserNameCandidatesFromEmail === "function") ? getUserNameCandidatesFromEmail(email) : [];
  if (!names.length) {
    var fallbackName = getUserNameFromEmail(email);
    names = fallbackName ? [fallbackName] : [];
  }
  if (!names.length) return getBlankStats();

  var stats = _getRepStats_();

  var disco = getRankAndValue_(stats.disco, names);
  var nbm = getRankAndValue_(stats.nbm, names);
  var pipe = getRankAndValue_(stats.pipe, names);
  var stg4 = getRankAndValue_(stats.stg4, names);

  return {
    disco_rank: disco.rank,
    disco_amt: disco.amt,
    nbm_rank: nbm.rank === "-" ? "dead_last" : nbm.rank,
    nbm_amt: nbm.amt,
    pipe_rank: pipe.rank,
    pipe_amt: pipe.amt,
    stg4_rank: stg4.rank,
    stg4_amt: stg4.amt
  };
}
