// -----------------------------------------------------------------------
// RainmakerScoring.js
// Pure scoring/ranking layer for Rainmaker dashboard
// -----------------------------------------------------------------------

// -----------------------------------------------------------------------
// MAIN ENTRY POINT
// -----------------------------------------------------------------------
function computeRainmakerScoreboard() {
  return computeRainmakerScoreboard_();
}

function computeRainmakerScoreboard_() {
  var ss = SpreadsheetApp.openById(TRACKER_SHEET_ID);

  // Read source data
  var reps = readRainmakerSheet_(ss);
  var entBenchmarks = readEntBenchmarkSheet_(ss);
  var lastRefresh = readLastRefreshTimestamp_(ss);

  if (!reps || reps.length === 0) {
    throw new Error('RainmakerScoring: no Rainmaker sheet data — run runRainmakerRefresh first');
  }

  // Get fiscal quarter from first rep
  var fiscalQuarter = reps[0].fiscal_quarter || '';

  // Category metadata
  var categories = [
    { key: 'nbm',          weight: 1, type: 'count',  enterpriseBenchmark: true },
    { key: 'pipe_adds',    weight: 1, type: 'count',  enterpriseBenchmark: true },
    { key: 'pipe_dollars', weight: 1, type: 'amount', enterpriseBenchmark: true },
    { key: 'c_level',      weight: 1, type: 'count',  enterpriseBenchmark: false },
    { key: 'stage4_plus',  weight: 2, type: 'amount', enterpriseBenchmark: true },
    { key: 'closed_won',   weight: 3, type: 'amount', enterpriseBenchmark: true },
    { key: 'pocs',         weight: 1, type: 'count',  enterpriseBenchmark: false },
    { key: 'partner_reg',  weight: 1, type: 'count',  enterpriseBenchmark: true }
  ];

  // Score each category and build category details for each rep
  var repScores = [];

  for (var i = 0; i < reps.length; i++) {
    var rep = reps[i];
    var repScore = {
      rep_email: rep.rep_email,
      rep_name: rep.rep_name,
      total_points: 0,
      categories: {},
      errors: rep.errors || []
    };
    repScores.push(repScore);
  }

  // For each category, rank all reps and assign points
  for (var j = 0; j < categories.length; j++) {
    var cat = categories[j];

    // Rank reps in this category
    var rankings = rankRepsInCategory_(reps, cat.key, cat.weight);

    // For Enterprise-benchmarked categories, get ent rank and color
    var entRows = cat.enterpriseBenchmark ? (entBenchmarks[cat.key] || []) : null;

    // Populate each rep's category details
    for (var k = 0; k < repScores.length; k++) {
      var rs = repScores[k];
      var ranking = rankings[rs.rep_email];

      if (!ranking) continue;

      var catDetail = {
        value: ranking.value,
        display: cat.type === 'amount' ? formatCurrency_(ranking.value) : formatCount_(ranking.value),
        place_points: ranking.place_points,
        weighted_points: ranking.weighted_points,
        team_rank: ranking.rank
      };

      if (cat.enterpriseBenchmark) {
        var entInfo = getEnterpriseRankAndColor_(ranking.value, entRows);
        catDetail.ent_rank = entInfo.rank;
        catDetail.ent_total = entInfo.totalRows;
        catDetail.color = entInfo.color;
      } else {
        // Team-internal color
        catDetail.color = getTeamNorthInternalColor_(ranking.rank);
      }

      rs.categories[cat.key] = catDetail;
      rs.total_points += ranking.weighted_points;
    }
  }

  // Rank reps by total_points (with tiebreakers: closed_won, stage4_plus, name)
  repScores.sort(function(a, b) {
    if (a.total_points !== b.total_points) {
      return b.total_points - a.total_points;
    }
    // Tiebreak by closed_won value
    var aCW = a.categories.closed_won ? a.categories.closed_won.value : 0;
    var bCW = b.categories.closed_won ? b.categories.closed_won.value : 0;
    if (aCW !== bCW) return bCW - aCW;

    // Tiebreak by stage4_plus value
    var aS4 = a.categories.stage4_plus ? a.categories.stage4_plus.value : 0;
    var bS4 = b.categories.stage4_plus ? b.categories.stage4_plus.value : 0;
    if (aS4 !== bS4) return bS4 - aS4;

    // Final tiebreak: alphabetical by name
    return (a.rep_name || '').localeCompare(b.rep_name || '');
  });

  // Assign team_rank
  for (var m = 0; m < repScores.length; m++) {
    repScores[m].team_rank = m + 1;
  }

  // Build ent_leaderboards for the 6 Enterprise-benchmarked categories
  var entLeaderboards = {};
  var entCategories = ['nbm', 'pipe_adds', 'pipe_dollars', 'stage4_plus', 'closed_won', 'partner_reg'];

  for (var n = 0; n < entCategories.length; n++) {
    var entCatKey = entCategories[n];
    var entCatRows = entBenchmarks[entCatKey] || [];

    if (entCatRows.length === 0) {
      entLeaderboards[entCatKey] = {
        median: 0,
        top_10_percent_threshold: 0,
        top_10: []
      };
      continue;
    }

    // Extract values (already sorted descending)
    var values = entCatRows.map(function(r) { return r.value; });

    // Median
    var median = computeMedian_(values);

    // Top 10% threshold
    var top10Pos = Math.ceil(values.length * 0.1);
    var top10Threshold = values[top10Pos - 1] || 0;

    // Top 10 rows
    var top10Rows = entCatRows.slice(0, Math.min(10, entCatRows.length));

    // Find category metadata
    var catMeta = null;
    for (var p = 0; p < categories.length; p++) {
      if (categories[p].key === entCatKey) {
        catMeta = categories[p];
        break;
      }
    }

    var top10Display = top10Rows.map(function(row) {
      var isTeamNorth = REP_EMAILS.indexOf(row.owner_email) !== -1;
      return {
        owner_name: row.owner_name,
        owner_email: row.owner_email,
        value: row.value,
        display: catMeta && catMeta.type === 'amount' ? formatCurrency_(row.value) : formatCount_(row.value),
        is_team_north: isTeamNorth
      };
    });

    entLeaderboards[entCatKey] = {
      median: median,
      top_10_percent_threshold: top10Threshold,
      top_10: top10Display
    };
  }

  return {
    fiscal_quarter: fiscalQuarter,
    last_refresh: lastRefresh,
    team: repScores,
    ent_leaderboards: entLeaderboards
  };
}

// -----------------------------------------------------------------------
// SHEET READERS
// -----------------------------------------------------------------------
function readRainmakerSheet_(ss) {
  var sheet = ss.getSheetByName(RAINMAKER_SHEET_NAME);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var reps = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rep = {
      rep_email: row[0] || '',
      rep_name: row[1] || '',
      fiscal_quarter: row[2] || '',
      nbm: row[3] || 0,
      pipe_adds: row[4] || 0,
      pipe_dollars: row[5] || 0,
      c_level: row[6] || 0,
      stage4_plus: row[7] || 0,
      closed_won: row[8] || 0,
      pocs: row[9] || 0,
      partner_reg: row[10] || 0,
      errors: parseErrorsColumn_(row[11])
    };
    reps.push(rep);
  }

  return reps;
}

function parseErrorsColumn_(val) {
  if (!val || val === 'none') return [];
  return String(val).split(';').map(function(s) { return s.trim(); }).filter(Boolean);
}

function readEntBenchmarkSheet_(ss) {
  var sheet = ss.getSheetByName(RAINMAKER_SHEET_NAME + '_Ent');
  if (!sheet) return {};

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return {};

  var byCategory = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var cat = row[0] || '';
    var ownerName = row[1] || '';
    var ownerEmail = row[2] || '';
    var value = row[3] || 0;

    if (!cat) continue;

    if (!byCategory[cat]) byCategory[cat] = [];
    byCategory[cat].push({
      owner_name: ownerName,
      owner_email: ownerEmail,
      value: value
    });
  }

  // Sort each category descending by value
  for (var cat in byCategory) {
    byCategory[cat].sort(function(a, b) {
      return b.value - a.value;
    });
  }

  return byCategory;
}

function readLastRefreshTimestamp_(ss) {
  var sheet = ss.getSheetByName(RAINMAKER_LOG_SHEET_NAME);
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return null;

  // Last row (most recent)
  var lastRow = data[data.length - 1];
  var timestamp = lastRow[0];

  if (!timestamp || !(timestamp instanceof Date)) return null;

  return timestamp.toISOString();
}

// -----------------------------------------------------------------------
// RANKING & SCORING
// -----------------------------------------------------------------------
function rankRepsInCategory_(reps, categoryKey, weight) {
  // Returns { email: { rank, value, place_points, weighted_points } }

  // Extract values
  var repValues = [];
  for (var i = 0; i < reps.length; i++) {
    var rep = reps[i];
    var value = rep[categoryKey] || 0;
    repValues.push({
      email: rep.rep_email,
      value: value
    });
  }

  // Check if all values are 0/null
  var allZero = true;
  for (var j = 0; j < repValues.length; j++) {
    if (repValues[j].value > 0) {
      allZero = false;
      break;
    }
  }

  if (allZero) {
    // Everyone gets 0 points
    var result = {};
    for (var k = 0; k < repValues.length; k++) {
      result[repValues[k].email] = {
        rank: 1,
        value: 0,
        place_points: 0,
        weighted_points: 0
      };
    }
    return result;
  }

  // Sort descending by value
  repValues.sort(function(a, b) {
    if (a.value !== b.value) return b.value - a.value;
    // Tie-break by email alphabetically for stability
    return a.email.localeCompare(b.email);
  });

  // Assign competition ranks and place points
  var rankMap = {};
  var currentRank = 1;

  for (var m = 0; m < repValues.length; m++) {
    var rv = repValues[m];

    // If this value differs from previous, update rank
    if (m > 0 && repValues[m - 1].value !== rv.value) {
      currentRank = m + 1;
    }

    // Place points based on rank: 1st=5, 2nd=4, 3rd=3, 4th=2, 5th=1
    var placePoints = 0;
    if (currentRank === 1) placePoints = 5;
    else if (currentRank === 2) placePoints = 4;
    else if (currentRank === 3) placePoints = 3;
    else if (currentRank === 4) placePoints = 2;
    else if (currentRank === 5) placePoints = 1;

    rankMap[rv.email] = {
      rank: currentRank,
      value: rv.value,
      place_points: placePoints,
      weighted_points: placePoints * weight
    };
  }

  return rankMap;
}

function getEnterpriseRankAndColor_(value, entRows) {
  // entRows is sorted descending by value
  // Returns { rank, totalRows, color }

  if (!entRows || entRows.length === 0) {
    return { rank: null, totalRows: 0, color: 'yellow' };
  }

  var totalRows = entRows.length;

  // Find rank (1-indexed)
  var rank = null;
  for (var i = 0; i < entRows.length; i++) {
    if (entRows[i].value <= value) {
      rank = i + 1;
      break;
    }
  }

  // If value is lower than all rows, rank = totalRows + 1
  if (rank === null) {
    rank = totalRows + 1;
  }

  // Determine color
  var color = 'red';
  var top10Cutoff = Math.ceil(totalRows * 0.1);
  var top50Cutoff = Math.ceil(totalRows * 0.5);

  if (rank <= top10Cutoff) {
    color = 'green';
  } else if (rank <= top50Cutoff) {
    color = 'yellow';
  }

  return { rank: rank, totalRows: totalRows, color: color };
}

function getTeamNorthInternalColor_(rank) {
  // rank is 1-5 within Team North
  if (rank <= 2) return 'green';
  if (rank === 3) return 'yellow';
  return 'red';
}

// -----------------------------------------------------------------------
// FORMATTING
// -----------------------------------------------------------------------
function formatCurrency_(amount) {
  if (amount >= 1000000) {
    return '$' + (amount / 1e6).toFixed(2) + 'M';
  }
  if (amount >= 100000) {
    return '$' + Math.round(amount / 1e3) + 'K';
  }
  if (amount >= 1000) {
    return '$' + (amount / 1e3).toFixed(1) + 'K';
  }
  return '$' + amount;
}

function formatCount_(n) {
  if (n >= 10000) {
    return n.toLocaleString();
  }
  return String(n);
}

// -----------------------------------------------------------------------
// HELPERS
// -----------------------------------------------------------------------
function computeMedian_(values) {
  // values is array of numbers, already sorted descending
  if (values.length === 0) return 0;

  var sorted = values.slice().sort(function(a, b) { return a - b; });
  var mid = Math.floor(sorted.length / 2);

  if (sorted.length % 2 === 0) {
    return (sorted[mid - 1] + sorted[mid]) / 2;
  } else {
    return sorted[mid];
  }
}

// -----------------------------------------------------------------------
// SMOKE TEST
// -----------------------------------------------------------------------
function testRainmakerScoreboard() {
  var result = computeRainmakerScoreboard();
  Logger.log('=== Rainmaker Scoreboard ===');
  Logger.log('Fiscal Quarter: ' + result.fiscal_quarter);
  Logger.log('Last Refresh: ' + result.last_refresh);
  Logger.log('Reps: ' + result.team.length);
  for (var i = 0; i < result.team.length; i++) {
    var r = result.team[i];
    Logger.log('  #' + r.team_rank + ' ' + r.rep_name + ': ' + r.total_points + ' pts');
  }
  Logger.log('Ent leaderboards: ' + Object.keys(result.ent_leaderboards).join(', '));
  for (var cat in result.ent_leaderboards) {
    var lb = result.ent_leaderboards[cat];
    Logger.log('  ' + cat + ': median=' + lb.median + ', top10%=' + lb.top_10_percent_threshold + ', top10 has ' + lb.top_10.length + ' entries');
  }
  return result;
}

function dumpScoreboardDetails() {
  var result = computeRainmakerScoreboard();
  Logger.log('=== PER-REP PER-CATEGORY BREAKDOWN ===');
  var cats = ['nbm','pipe_adds','pipe_dollars','c_level','stage4_plus','closed_won','pocs','partner_reg'];
  for (var i = 0; i < result.team.length; i++) {
    var r = result.team[i];
    Logger.log('--- ' + r.rep_name + ' (Total: ' + r.total_points + ' pts, Team Rank: #' + r.team_rank + ') ---');
    for (var c = 0; c < cats.length; c++) {
      var cat = cats[c];
      var cd = r.categories[cat] || {};
      Logger.log('  ' + cat + ': value=' + cd.value + ', display="' + cd.display + '", team_rank=' + cd.team_rank + ', place_pts=' + cd.place_points + ', weighted_pts=' + cd.weighted_points + ', color=' + cd.color);
    }
  }
  return result;
}
