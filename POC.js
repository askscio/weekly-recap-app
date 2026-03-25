// =======================================================================
// POC COMMAND CENTER — Backend Functions
// =======================================================================

// -----------------------------------------------------------------------
// POC Sheet helpers
// -----------------------------------------------------------------------
function getOrCreatePocAccountsSheet() {
  var ss = SpreadsheetApp.openById(TRACKER_SHEET_ID);
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
  var ss = SpreadsheetApp.openById(TRACKER_SHEET_ID);
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
    var cached = getJsonCache_("poc_data_v4", null);
    if (cached) return cached;

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
    var rawActions = [];
    for (var j = 1; j < actionsData.length; j++) {
      var a = actionsData[j];
      var dt = a[5] ? new Date(a[5]) : null;
      var ts = (dt instanceof Date && !isNaN(dt.getTime())) ? dt.getTime() : 0;
      rawActions.push({
        account: a[0], actionItem: a[1], rag: a[2], signalSource: a[3],
        notes: a[4],
        lastUpdated: ts ? new Date(ts).toISOString() : "",
        updatedBy: a[6],
        _sortTs: ts,
        _seq: j
      });
    }
    rawActions = collapsePocActionHistory_(rawActions);

    var maxPer = getPocActionsMaxPerAccount_();
    var grouped = {};
    for (var k = 0; k < rawActions.length; k++) {
      var act = rawActions[k];
      var acctKey = String(act.account || "").trim();
      if (!acctKey) continue;
      if (!grouped[acctKey]) grouped[acctKey] = [];
      grouped[acctKey].push(act);
    }

    var latestActivityByAccount = {};
    var actions = [];
    for (var acct in grouped) {
      grouped[acct].sort(function(x, y) {
        var diff = Number(y._sortTs || 0) - Number(x._sortTs || 0);
        if (diff !== 0) return diff;
        return Number(y._seq || 0) - Number(x._seq || 0);
      });
      if (grouped[acct].length) latestActivityByAccount[acct] = Number(grouped[acct][0]._sortTs || 0);
      for (var n = 0; n < grouped[acct].length && n < maxPer; n++) actions.push(grouped[acct][n]);
    }

    accounts.sort(function(a, b) {
      var at = Number(latestActivityByAccount[a.account] || 0);
      var bt = Number(latestActivityByAccount[b.account] || 0);
      if (bt !== at) return bt - at;
      return String(a.account || "").localeCompare(String(b.account || ""));
    });

    var result = {
      success: true,
      accounts: accounts,
      actions: actions,
      rowsInPocActions: Math.max(0, actionsSheet.getLastRow() - 1),
      refreshConfig: getPocRefreshCapabilities_(),
      maxActionsPerAccount: maxPer
    };
    putJsonCache_("poc_data_v4", result, 45);
    return result;
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getPocDashboardSnapshot_() {
  return {
    generated_at: "2026-03-23T15:00:00Z",
    accounts: [
      {
        name: "Morningstar",
        status: "Yellow",
        health_summary: [
          "Morningstar chose to move forward with a full deployment POC in partnership with AHEAD, and the POC quote was approved on 3/4.",
          "Legal/security remains the primary blocker: on 3/18, Morningstar pushed a meeting to 3/21 due to security-related questions on MSA terms, with signature expected the following week.",
          "Taylor Lundberg is on maternity leave; Johnny Hatch and Billy Schuett are covering, introducing execution risk.",
          "The pilot is scheduled to start 3/16 with an end date of 6/14, but the SFDC pilot record still shows P1 – Security/Arch sign-off, meaning the POC has not technically launched yet."
        ],
        key_signals: [
          "POC quote approved 3/4; Morningstar opted for full deployment POC with AHEAD services.",
          "MSA terms under review; Morningstar pushed 3/18 meeting to 3/21 due to security questions — signature expected next week.",
          "SE Mike Koscak identified a contact (Hannah) from a prior Salesforce CPQ project who could help expand the land scope.",
          "Multiple senior leaders (marketing ops, sales ops, finance tech, product) have expressed interest in participating in the POC.",
          "Taylor on maternity leave — coverage transition introduces execution risk.",
          "Competitor landscape includes Microsoft Copilot, OpenAI Enterprise, Agentforce, and DIY solutions."
        ],
        recommended_actions: [
          "Follow up on the 3/21 meeting outcome — confirm whether MSA was signed and unblock the POC start immediately.",
          "Define exact use cases and connected systems with Samuel Castano and Ping Lu to ensure the POC demonstrates meaningful business value.",
          "Begin quantifying ROI with Glean's business value team; align with AHEAD on the Morningstar-specific ROI framework.",
          "Leverage AHEAD partnership to accelerate setup and services — ensure the 6/14 pilot end date is achievable.",
          "Engage Mike Koscak's contact Hannah to expand the land scope beyond the initial target."
        ]
      },
      {
        name: "A.T. Kearney",
        status: "Yellow",
        health_summary: [
          "The Kearney POC is active and in P6 – Crawling/Indexing stage with ~40 pilot users across multiple geographies; the paid pilot ($10K) kicked off on 3/9.",
          "A significant risk emerged on 3/12: pilot users are confused about how Glean differs from Microsoft Copilot due to the small pilot scope, putting the $2.94M land opportunity at risk.",
          "On the 3/16 sync call, Kearney user Nadim referenced a 'Thursday debacle' where multiple users questioned Copilot vs. Glean and suggested a live differentiation session.",
          "SharePoint permissions issues affected at least one pilot user and more may be silently impacted."
        ],
        key_signals: [
          "Pilot is live with users actively testing; POC officially kicked off 3/9.",
          "Copilot differentiation confusion is the #1 risk — multiple users questioned Glean vs. Copilot in the pilot group chat.",
          "SharePoint permissions gap affected pilot user Pranav Sharma; more users may be silently affected.",
          "Kearney was asked to send sample SharePoint docs/folder links for agent seeding and RFP workflows — not yet received.",
          "The $2.94M Intelligence Platform opportunity is at Stage 4 with a Q2-2027 close date of 5/29/2026.",
          "Billy committed to follow up with AHEAD partner Josh Weisfuse to press at every angle on differentiation."
        ],
        recommended_actions: [
          "Schedule the Copilot differentiation live session ASAP — this was requested by Kearney users and has strong interest.",
          "Proactively reach out to all pilot users to check for SharePoint permissions gaps and prevent silent churn.",
          "Collect sample SharePoint docs and folder links from Kearney to tailor starter agents and RFP workflows.",
          "Coordinate with AHEAD partner Josh Weisfuse to reinforce Glean's value proposition at every touchpoint.",
          "Prepare stronger use-case demos that clearly differentiate Glean's enterprise search and knowledge graph from Copilot."
        ]
      },
      {
        name: "Chamberlain Group",
        status: "Green",
        health_summary: [
          "The 100-user land deal ($69.4K TCV) closed won on 2/24/2026 — Chamberlain is now a Glean customer, positioned as phase 1 of a larger ~3,000-user expansion.",
          "On 3/19, Chamberlain confirmed that their contracts (ToS/DPA) are complete, clearing the legal blocker for the paid pilot and the $650K expansion opportunity.",
          "DEM Connor Waldo and AIOM Fernando Ponce have been assigned for post-close kickoff readiness.",
          "The $650K expansion opportunity is at Stage 2 – Blueprint with a Q2-2027 close date of 5/29/2026."
        ],
        key_signals: [
          "Land deal closed won at $69.4K for 100 users on 2/24/2026.",
          "Contracts (ToS/DPA) confirmed complete by Chamberlain on 3/19 — legal blocker cleared.",
          "$650K expansion opportunity created at Stage 2 with 5/29/2026 close date.",
          "Jeff Wyzywany (Strategic Sourcing) sent redlines on ToS/DPA on 3/19 — Niko looped in Glean legal for review.",
          "Chamberlain prefers to use their own established channel partners for future expansion rather than NTT.",
          "CIO David Stahl and Director Manoj Chinnasamy are actively engaged."
        ],
        recommended_actions: [
          "Schedule kickoff/rollout call immediately now that contracts are complete — begin platform deployment with the initial 100 users.",
          "Resolve any remaining ToS/DPA redlines from Jeff Wyzywany quickly to avoid delays in the expansion motion.",
          "Ensure a strong initial deployment experience to build momentum for the $650K expansion.",
          "Coordinate with DEM Connor Waldo and AIOM Fernando Ponce on kickoff logistics.",
          "Begin mapping expansion use cases across engineering, company-wide search, and additional departments to justify the 3,000-user target."
        ]
      },
      {
        name: "HUDL",
        status: "Yellow",
        health_summary: [
          "The HUDL POC is active and in P10 – User Piloting stage, running from ~2/20 to 4/10/2026, with the $250K new business opportunity at Stage 4.",
          "A critical Salesforce connector authentication failure (RETRIES_EXHAUSTED) was escalated on 3/13 — root cause is a systemic scio-apps issue. Jira EE-23939 created.",
          "Daniel Shaefer (Head of Business Technology) confirmed positive POC feedback and discussed expanding from 350 support users to 2,800–3,000 enterprise-wide, targeting a mid-April executive write-up.",
          "Hudl's recent enterprise Claude agreement creates competitive noise, though internal champions view Glean as the stronger platform play."
        ],
        key_signals: [
          "POC is active with positive feedback; Fabio Girolami says Glean has won vs. Gemini and Claude in his personal testing.",
          "Salesforce connector auth failure is a critical blocker — directly impacts the active pilot experience ending 4/10.",
          "Daniel Shaefer confirmed potential enterprise rollout to 2,800–3,000 users with mid-April executive write-up target.",
          "Hudl signed an enterprise Claude agreement — CTO-driven, primarily for engineering/code, but creates competitive noise.",
          "support.hudl is not indexed, causing search to return internal Guru content instead of public support articles.",
          "Johnny working to grow deal from Support-only to org-wide; business case meeting with AI lead scheduled for 3/24."
        ],
        recommended_actions: [
          "Resolve the Salesforce connector authentication failure urgently — this is the #1 technical blocker with the pilot ending 4/10.",
          "Add support.hudl as an indexed data source before the pilot ends to demonstrate public-facing support answer quality.",
          "Execute the 3/23 working session with Daniel Shaefer to document the business case and competitive differentiation vs. Claude/Slack.",
          "Arrange the executive intro between Stephanie (Glean) and Courtney Rogers (Hudl SVP Business Operations).",
          "Connect Gong and Google Drive as additional data sources to expand POC value beyond support use cases.",
          "Confirm Salesforce embedded widget visibility for all pilot users to ensure consistent experience."
        ]
      },
      {
        name: "Bread Financial",
        status: "Red",
        health_summary: [
          "The Bread Financial pilot has never started — originally scheduled for Oct–Dec 2025, it remains stuck at P3 – EB Alignment on Scope with no meaningful activity in over a month.",
          "Johnny noted on 3/11 that Bread is 'mostly radio silent' except for a short note about concerns with Glean contract language; legal redlines are on hold pending a POC test plan revamp.",
          "CTO Allegra Driscoll is the Exec Sponsor but has never replied to any outreach. Champion Joe Petrelli's contact status is 'Marketing Nurture – Last Activity 30+ Days.'",
          "The $500K new business opportunity has a Q3-2027 close date of 9/30/2026 but is at serious risk of deprioritization given zero engagement momentum."
        ],
        key_signals: [
          "Pilot has never started — original dates (Oct–Dec 2025) are long past; opportunity is 416 days old.",
          "Contract language concerns raised internally; legal redlines on hold pending POC test plan revamp.",
          "CTO Allegra Driscoll (Exec Sponsor) has never responded to any outreach.",
          "Champion Joe Petrelli's contact status is 'Marketing Nurture – Last Activity 30+ Days.'",
          "Bread Financial has high intent signals (surging for Generative AI, Agentic AI, AI strategy) but zero engagement momentum.",
          "Last substantive update from John Downs on 3/9: internal discussions happened 3–4 weeks ago with no follow-up since."
        ],
        recommended_actions: [
          "Escalate internally — make a leadership-level decision on whether to continue investing or deprioritize this account.",
          "Re-engage champion Joe Petrelli with a simplified POC scope and revised timeline to restart momentum.",
          "Attempt a different approach to reach CTO Allegra Driscoll — consider leveraging a partner, event invite, or executive-to-executive outreach.",
          "Propose a simplified, lower-friction POC plan that addresses contract language concerns and reduces internal review burden.",
          "If no movement within 2 weeks, formally deprioritize and reallocate team resources to higher-probability accounts."
        ]
      }
    ]
  };
}

function normalizePocText_(v) {
  return String(v || "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function getPocContentFingerprint_(action) {
  var raw = String((action && action.actionItem) || "") + " " + String((action && action.notes) || "");
  return normalizePocText_(
    raw
      .replace(/https?:\/\/\S+/g, " ")
      .replace(/\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\s+\d{1,2}(?:,\s*\d{4})?/gi, " ")
      .replace(/\b\d{1,2}\/\d{1,2}(?:\/\d{2,4})?\b/g, " ")
      .replace(/\b\d+\b/g, " ")
      .replace(/[^a-z0-9]+/gi, " ")
  );
}

function getPocKeywordStopwords_() {
  return {
    a: 1, an: 1, and: 1, are: 1, as: 1, at: 1, be: 1, been: 1, being: 1, but: 1,
    by: 1, call: 1, can: 1, could: 1, for: 1, from: 1, had: 1, has: 1, have: 1,
    into: 1, its: 1, just: 1, may: 1, more: 1, need: 1, needs: 1, not: 1, now: 1,
    off: 1, one: 1, onto: 1, our: 1, out: 1, over: 1, same: 1, should: 1, show: 1,
    that: 1, the: 1, their: 1, them: 1, then: 1, there: 1, they: 1, this: 1,
    today: 1, was: 1, were: 1, what: 1, when: 1, where: 1, while: 1, with: 1,
    would: 1, yesterday: 1, user: 1, users: 1, team: 1, account: 1, source: 1,
    updated: 1, refresh: 1, glean: 1, auto: 1, pilot: 1
  };
}

function simplifyPocToken_(token) {
  var t = String(token || "");
  if (t.length > 6 && /ing$/.test(t)) t = t.replace(/ing$/, "");
  else if (t.length > 5 && /ed$/.test(t)) t = t.replace(/ed$/, "");
  else if (t.length > 4 && /es$/.test(t)) t = t.replace(/es$/, "");
  else if (t.length > 4 && /s$/.test(t)) t = t.replace(/s$/, "");
  return t;
}

function getPocConceptKey_(action) {
  var stop = getPocKeywordStopwords_();
  var text = getPocContentFingerprint_(action);
  var toks = text.split(/\s+/);
  var seen = {};
  var keep = [];
  for (var i = 0; i < toks.length; i++) {
    var tok = simplifyPocToken_(toks[i]);
    if (!tok || tok.length < 4 || stop[tok] || seen[tok]) continue;
    seen[tok] = true;
    keep.push(tok);
  }
  keep.sort();
  if (!keep.length) return normalizePocText_(action && action.actionItem);
  return keep.slice(0, 8).join("|");
}

function getPocActionHistoryKey_(action) {
  return normalizePocText_(action && action.account) + "|" + getPocConceptKey_(action);
}

function getPocActionStateSignature_(action) {
  return [
    getPocConceptKey_(action),
    normalizePocText_(action && action.rag)
  ].join("|");
}

function collapsePocActionHistory_(actions) {
  var grouped = {};
  var order = [];

  for (var i = 0; i < actions.length; i++) {
    var action = actions[i];
    var key = getPocActionHistoryKey_(action);
    if (!key || key === "|") continue;
    if (!grouped[key]) {
      grouped[key] = [];
      order.push(key);
    }
    grouped[key].push(action);
  }

  var out = [];
  for (var g = 0; g < order.length; g++) {
    var history = grouped[order[g]];
    history.sort(function(a, b) {
      var at = Number(a && a._sortTs || 0);
      var bt = Number(b && b._sortTs || 0);
      if (at !== bt) return at - bt;
      return Number(a && a._seq || 0) - Number(b && b._seq || 0);
    });

    var current = null;
    for (var h = 0; h < history.length; h++) {
      var row = history[h];
      var sig = getPocActionStateSignature_(row);
      if (!current || current._stateSig !== sig) {
        if (current) out.push(current);
        current = {
          account: row.account,
          actionItem: row.actionItem,
          rag: row.rag,
          signalSource: row.signalSource,
          notes: row.notes,
          updatedBy: row.updatedBy,
          lastUpdated: row.lastUpdated,
          _sortTs: row._sortTs,
          _seq: row._seq,
          _stateSig: sig
        };
      } else {
        current.updatedBy = row.updatedBy || current.updatedBy;
      }
    }
    if (current) out.push(current);
  }

  return out;
}

function normalizePocAccountKey_(v) {
  return String(v || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function daysAgoLabel_(dateObj) {
  if (!(dateObj instanceof Date) || isNaN(dateObj.getTime())) return "";
  var diffMs = new Date().getTime() - dateObj.getTime();
  var diffDays = Math.floor(diffMs / 86400000);
  if (diffDays <= 0) return "today";
  if (diffDays === 1) return "1 day ago";
  return diffDays + " days ago";
}

function buildPocRecapSentence_(accountName, actionInfo, recapInfo) {
  var bits = [];
  bits.push(accountName + ":");

  if (actionInfo && actionInfo.total > 0) {
    var ragBits = [];
    if (actionInfo.red) ragBits.push(actionInfo.red + " red");
    if (actionInfo.yellow) ragBits.push(actionInfo.yellow + " yellow");
    if (actionInfo.green) ragBits.push(actionInfo.green + " green");
    if (ragBits.length) bits.push(ragBits.join(", ") + " signals.");
    if (actionInfo.topAction) bits.push("Top item: " + actionInfo.topAction + ".");
  } else {
    bits.push("No current POC actions logged.");
  }

  if (recapInfo) {
    var recapBits = [];
    recapBits.push("Recent rep update from " + recapInfo.repName + " (" + recapInfo.whenLabel + ").");
    if (recapInfo.move) recapBits.push("Move: " + recapInfo.move + ".");
    if (recapInfo.next) recapBits.push("Next: " + recapInfo.next + ".");
    if (recapInfo.risk) recapBits.push("Risk: " + recapInfo.risk + ".");
    else if (recapInfo.goal) recapBits.push("Goal: " + recapInfo.goal + ".");
    else if (recapInfo.forecastNote) recapBits.push("Forecast note: " + recapInfo.forecastNote + ".");
    bits.push(recapBits.join(" "));
  }

  return bits.join(" ").replace(/\s+/g, " ").trim();
}

function trimToWordCount_(text, maxWords) {
  var words = String(text || "").trim().split(/\s+/).filter(Boolean);
  if (words.length <= maxWords) return String(text || "").trim();
  return words.slice(0, maxWords).join(" ").replace(/[,:;.-]?$/, "") + "…";
}

function getPocRecapNote() {
  ensureAdminCaller_();

  var daysBack = 7;
  var now = new Date();
  var cutoff = new Date(now.getTime() - daysBack * 86400000);
  var accountsSheet = getOrCreatePocAccountsSheet();
  var actionsSheet = getOrCreatePocActionsSheet();

  var accountRows = accountsSheet.getDataRange().getValues();
  var pocAccounts = [];
  var pocAccountKeyToName = {};
  for (var i = 1; i < accountRows.length; i++) {
    var acctName = String(accountRows[i][0] || "").trim();
    if (!acctName) continue;
    pocAccounts.push(acctName);
    pocAccountKeyToName[normalizePocAccountKey_(acctName)] = acctName;
  }

  var actionRows = actionsSheet.getDataRange().getValues();
  var actionStats = {};
  for (var j = 1; j < actionRows.length; j++) {
    var actionAcct = String(actionRows[j][0] || "").trim();
    if (!actionAcct) continue;
    var acctKey = normalizePocAccountKey_(actionAcct);
    if (!pocAccountKeyToName[acctKey]) continue;
    var rag = String(actionRows[j][2] || "").toLowerCase().trim();
    var actionItem = String(actionRows[j][1] || "").trim();
    if (!actionStats[acctKey]) actionStats[acctKey] = { total: 0, red: 0, yellow: 0, green: 0, topAction: "" };
    actionStats[acctKey].total++;
    if (rag === "red") actionStats[acctKey].red++;
    else if (rag === "yellow") actionStats[acctKey].yellow++;
    else if (rag === "green") actionStats[acctKey].green++;
    if (!actionStats[acctKey].topAction && actionItem) actionStats[acctKey].topAction = actionItem;
  }

  var recapByAccount = {};
  var recapSheet = getRecapSheet_();
  if (recapSheet.getLastRow() >= 2) {
    var recapRows = recapSheet.getRange(2, 1, recapSheet.getLastRow() - 1, recapSheet.getLastColumn()).getValues();
    for (var r = recapRows.length - 1; r >= 0; r--) {
      var row = recapRows[r];
      var ts = row[0];
      var tsDate = ts instanceof Date ? ts : new Date(ts);
      if (!(tsDate instanceof Date) || isNaN(tsDate.getTime())) continue;
      if (tsDate < cutoff) continue;
      var email = String(row[1] || "").trim().toLowerCase();
      var parsed = parseRecapRow_(row);
      var repName = getUserNameFromEmail(email) || email;
      for (var a = 1; a <= 6; a++) {
        var recapAcct = String(parsed["acct" + a + "_name"] || "").trim();
        if (!recapAcct) continue;
        var recapKey = normalizePocAccountKey_(recapAcct);
        if (!pocAccountKeyToName[recapKey]) continue;
        if (recapByAccount[recapKey]) continue; // newest only
        recapByAccount[recapKey] = {
          repName: repName,
          whenLabel: daysAgoLabel_(tsDate),
          move: String(parsed["acct" + a + "_move"] || "").trim(),
          next: String(parsed["acct" + a + "_next"] || "").trim(),
          risk: String(parsed.risk || "").trim(),
          goal: String(parsed.goal || "").trim(),
          forecastNote: String(parsed.forecast_note || "").trim()
        };
      }
    }
  }

  var ranked = [];
  for (var p = 0; p < pocAccounts.length; p++) {
    var name = pocAccounts[p];
    var key = normalizePocAccountKey_(name);
    var stats = actionStats[key] || { total: 0, red: 0, yellow: 0, green: 0, topAction: "" };
    var recap = recapByAccount[key] || null;
    var score = (stats.red * 100) + (stats.yellow * 20) + (recap ? 5 : 0) + stats.total;
    ranked.push({ name: name, key: key, stats: stats, recap: recap, score: score });
  }
  ranked.sort(function(a, b) { return b.score - a.score; });

  var intro = "Current POC picture: ";
  var sentences = [];
  var recapsUsed = 0;
  for (var q = 0; q < ranked.length; q++) {
    var item = ranked[q];
    if (item.recap) recapsUsed++;
    sentences.push(buildPocRecapSentence_(item.name, item.stats, item.recap));
  }
  if (!recapsUsed) {
    sentences.unshift("No rep recap updates were submitted in the last 7 days, so this note is based on current POC Actions only.");
  }

  var note = trimToWordCount_(intro + sentences.join(" "), 200);
  return {
    success: true,
    note: note,
    generatedAt: now.toISOString(),
    daysBack: daysBack,
    accountsConsidered: pocAccounts.length,
    recentRepUpdatesUsed: recapsUsed
  };
}

function updateRagStatus(account, actionItem, newRag) {
  try {
    var sheet = getOrCreatePocActionsSheet();
    var data = sheet.getDataRange().getValues();
    var email = getCallerEmailSafe_() || "unknown";
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
    var email = getCallerEmailSafe_() || "unknown";
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
    var email = getCallerEmailSafe_() || "unknown";
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

function isPlaceholderValue_(v) {
  var s = String(v || "").toLowerCase().trim();
  return !s ||
    s.indexOf("your glean") !== -1 ||
    s.indexOf("api url") !== -1 ||
    s.indexOf("bearer token") !== -1 ||
    s.indexOf("<paste") !== -1;
}

function getPocRefreshCapabilities_() {
  var cfg = getGleanPocAgentConfig_();
  var hasDirectAny = !!cfg.base || !!cfg.agentId || !!cfg.apiToken;
  var hasDirectAll = !isPlaceholderValue_(cfg.base) &&
                     !isPlaceholderValue_(cfg.agentId) &&
                     !isPlaceholderValue_(cfg.apiToken);
  if (hasDirectAny && hasDirectAll) {
    return {
      canTrigger: true,
      mode: "apps_script_direct",
      source: "Apps Script -> Glean /agents/runs/wait"
    };
  }

  var hasProxyAny = !!cfg.endpoint || !!cfg.token;
  var hasProxyAll = !isPlaceholderValue_(cfg.endpoint) &&
                    !isPlaceholderValue_(cfg.token);
  if (hasProxyAny && hasProxyAll) {
    return {
      canTrigger: true,
      mode: "apps_script_proxy",
      source: "Apps Script -> custom proxy endpoint"
    };
  }

  return {
    canTrigger: false,
    mode: "glean_native_writeback",
    source: "Glean-native writeback to sheet",
    reason: "No valid Apps Script Glean credentials configured. Using passive status mode."
  };
}

function validateGleanPocAgentConfig_() {
  var cfg = getGleanPocAgentConfig_();

  // Mode A: direct Apps Script -> Glean runs/wait
  var hasDirect = !!cfg.base || !!cfg.agentId || !!cfg.apiToken;
  if (hasDirect) {
    if (isPlaceholderValue_(cfg.base)) throw new Error("Missing Script Property: GLEAN_BASE");
    if (isPlaceholderValue_(cfg.agentId)) throw new Error("Missing Script Property: GLEAN_AGENT_ID");
    if (isPlaceholderValue_(cfg.apiToken)) throw new Error("Missing Script Property: GLEAN_API_TOKEN");
    return cfg;
  }

  // Mode B: custom endpoint proxy
  if (isPlaceholderValue_(cfg.endpoint)) throw new Error("Missing Script Property: GLEAN_POC_AGENT_ENDPOINT");
  if (isPlaceholderValue_(cfg.token)) throw new Error("Missing Script Property: GLEAN_POC_AGENT_TOKEN");
  return cfg;
}

function buildGleanPocAgentPayload_() {
  return {
    mode: "poc_daily_refresh",
    accounts: getPocAccountNames_(),
    daysBack: getGleanPocAgentConfig_().daysBack,
    requestedAt: new Date().toISOString(),
    requestedBy: getCallerEmailSafe_() || ""
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
    var ts = a.updatedAt || a.lastUpdated || a.timestamp || a.date || "";
    var sortTs = 0;
    try { sortTs = ts ? new Date(ts).getTime() : 0; } catch (e) { sortTs = 0; }
    out.push({
      account: account,
      actionItem: actionItem,
      rag: normalizeGleanRag_(a.rag),
      signalSource: String(a.signalSource || "GLEAN").trim() || "GLEAN",
      notes: String(a.notes || a.summary || "").trim(),
      updatedBy: String(a.updatedBy || "Glean Agent").trim() || "Glean Agent",
      _sortTs: sortTs,
      _seq: i
    });
  }
  return out;
}

function getPocActionsMaxPerAccount_() {
  var props = PropertiesService.getScriptProperties();
  var raw = parseInt(String(props.getProperty("POC_MAX_ACTIONS_PER_ACCOUNT") || "8"), 10);
  if (!raw || raw < 1) raw = 8;
  if (raw > 8) raw = 8; // hard cap safety
  return raw;
}

function capLatestActionsPerAccount_(actions) {
  var maxPer = getPocActionsMaxPerAccount_();
  var grouped = {};
  var order = [];

  for (var i = 0; i < actions.length; i++) {
    var a = actions[i];
    var acctKey = String(a.account || "").toLowerCase().trim();
    var uniqKey = acctKey + "|" + String(a.actionItem || "").toLowerCase().trim();
    if (!acctKey || !a.actionItem) continue;
    if (!grouped[acctKey]) {
      grouped[acctKey] = { items: [], seen: {} };
      order.push(acctKey);
    }
    if (grouped[acctKey].seen[uniqKey]) continue;
    grouped[acctKey].seen[uniqKey] = true;
    grouped[acctKey].items.push(a);
  }

  var out = [];
  for (var j = 0; j < order.length; j++) {
    var k = order[j];
    var items = grouped[k].items;
    items.sort(function(x, y) {
      var xt = Number(x._sortTs || 0);
      var yt = Number(y._sortTs || 0);
      if (yt !== xt) return yt - xt; // newest first if timestamps available
      return Number(x._seq || 0) - Number(y._seq || 0); // preserve original order
    });
    for (var n = 0; n < items.length && n < maxPer; n++) out.push(items[n]);
  }
  return out;
}

function replacePocActionsSnapshot_(actions) {
  var sheet = getOrCreatePocActionsSheet();
  var existingRows = Math.max(0, sheet.getLastRow() - 1);

  // Always clear old rows and keep only newest snapshot.
  if (existingRows > 0) {
    sheet.getRange(2, 1, existingRows, sheet.getLastColumn()).clearContent();
  }

  var capped = capLatestActionsPerAccount_(actions);
  var now = new Date();

  if (capped.length > 0) {
    var rows = [];
    for (var i = 0; i < capped.length; i++) {
      var a = capped[i];
      rows.push([
        a.account, a.actionItem, a.rag || "Yellow", a.signalSource || "GLEAN",
        a.notes || "", now, a.updatedBy || "Glean Agent"
      ]);
    }
    sheet.getRange(2, 1, rows.length, POC_ACTIONS_HEADERS.length).setValues(rows);
  }

  return {
    replaced: capped.length,
    removed: existingRows,
    maxPerAccount: getPocActionsMaxPerAccount_()
  };
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

    var capabilities = getPocRefreshCapabilities_();
    var actionsSheet = getOrCreatePocActionsSheet();
    var currentRows = Math.max(0, actionsSheet.getLastRow() - 1);

    // Default/primary model: Glean-native writeback. If no Apps Script creds, do passive status check.
    if (!capabilities.canTrigger) {
      var passiveStatus = getPocGleanRefreshStatus();
      var passive = {
        success: true,
        passive: true,
        mode: "glean_native_writeback",
        message: capabilities.reason,
        rowsInPocActions: currentRows,
        lastStatus: (passiveStatus && passiveStatus.success) ? (passiveStatus.status || null) : null
      };
      return passive;
    }

    var responseObj = callGleanPocAgent_();
    var actions = [];
    var result = { replaced: 0, removed: 0, maxPerAccount: getPocActionsMaxPerAccount_() };
    var mode = "apps_script_snapshot_replace";
    // Snapshot mode: require structured actions so we can clear + replace deterministically.
    if (!(responseObj && (Array.isArray(responseObj.actions) || Array.isArray(responseObj.results)))) {
      throw new Error("Glean response must include structured actions/results for snapshot refresh.");
    }
    actions = parseGleanPocActions_(responseObj);
    result = replacePocActionsSnapshot_(actions);

    var rowCount = Math.max(0, actionsSheet.getLastRow() - 1);
    var okStatus = {
      ok: true,
      ranAt: startedAt.toISOString(),
      mode: mode,
      fetchedActions: actions.length,
      replaced: result.replaced,
      removed: result.removed,
      maxPerAccount: result.maxPerAccount,
      rowsInPocActions: rowCount
    };
    savePocGleanRefreshStatus_(okStatus);
    return {
      success: true,
      mode: mode,
      fetchedActions: actions.length,
      replaced: result.replaced,
      removed: result.removed,
      maxPerAccount: result.maxPerAccount,
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
  try { email = String(Session.getEffectiveUser().getEmail() || ""); } catch (e1) {}
  if (!email) {
    try { email = String(Session.getActiveUser().getEmail() || ""); } catch (e2) {}
  }
  return String(email || "").trim().toLowerCase();
}

// -----------------------------------------------------------------------
// Seed POC Data — run once from the spreadsheet menu
// -----------------------------------------------------------------------
function seedPocData() {
  try {
    var ss = SpreadsheetApp.openById(TRACKER_SHEET_ID);

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
