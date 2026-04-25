// -----------------------------------------------------------------------
// AI SUMMARY GENERATION
// -----------------------------------------------------------------------
var SUMMARY_AI_SHEET = typeof SUMMARY_AI_SHEET !== 'undefined' ? SUMMARY_AI_SHEET : 'Summary_AI';
var SUMMARY_AI_MODEL = typeof SUMMARY_AI_MODEL !== 'undefined' ? SUMMARY_AI_MODEL : 'gpt-5.2';

function getOrCreateSummaryAISheet() {
  var ss = SpreadsheetApp.openById(TRACKER_SHEET_ID);
  var sheet = ss.getSheetByName(SUMMARY_AI_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SUMMARY_AI_SHEET, ss.getSheets().length);
    sheet.appendRow(['generated_at', 'quarter_key', 'model', 'payload_json', 'source', 'status']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getLatestSummaryAISnapshot_(quarterKey) {
  try {
    var sheet = getOrCreateSummaryAISheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    var rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    for (var i = rows.length - 1; i >= 0; i--) {
      var rowQuarter = String(rows[i][1] || '').trim();
      var payloadRaw = String(rows[i][3] || '').trim();
      var status = String(rows[i][5] || '').trim().toLowerCase();
      if (!payloadRaw || (quarterKey && rowQuarter && rowQuarter !== quarterKey) || (status && status !== 'ok')) continue;
      try {
        var payload = JSON.parse(payloadRaw);
        if (payload && payload.generated_at) return payload;
      } catch (_) {}
    }
  } catch (err) {
    Logger.log('getLatestSummaryAISnapshot_ failed: ' + err.message);
  }
  return null;
}

function writeSummaryAISnapshot_(quarterKey, model, payload, source, status) {
  var sheet = getOrCreateSummaryAISheet();
  sheet.appendRow([
    new Date(),
    quarterKey || '',
    model || '',
    JSON.stringify(payload || {}),
    source || 'glean_agent',
    status || 'ok'
  ]);
}

function getOpenAIApiKey_() {
  var key = String(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY') || '').trim();
  if (!key) throw new Error('Missing OPENAI_API_KEY in Script Properties.');
  return key;
}

function buildTeamForecastDealContext_(repEmails) {
  var out = [];
  var seen = {};
  var emails = Array.isArray(repEmails) ? repEmails : [];
  for (var i = 0; i < emails.length; i++) {
    var email = String(emails[i] || '').trim().toLowerCase();
    if (!email) continue;
    var repName = getUserNameFromEmail(email) || email;
    var sf = getSFDataForUser(email);
    var deals = (sf && Array.isArray(sf.deals)) ? sf.deals : [];
    for (var j = 0; j < deals.length; j++) {
      var d = deals[j] || {};
      var account = String(d.account || '').trim();
      var rawCategory = String(d.category || '').trim();
      if (!account || !rawCategory) continue;
      var category = rawCategory.toLowerCase();
      if (category.indexOf('closed') !== -1) category = 'Closed';
      else if (category.indexOf('commit') !== -1) category = 'Commit';
      else if (category.indexOf('most likely') !== -1 || category.indexOf('most_likely') !== -1) category = 'Most Likely';
      else if (category.indexOf('best case') !== -1 || category.indexOf('best_case') !== -1) category = 'Best Case';
      else continue;
      var key = [repName.toLowerCase(), account.toLowerCase(), category].join('|');
      if (seen[key]) continue;
      seen[key] = true;
      out.push({
        rep_name: repName,
        account: account,
        category: category,
        amount: Number(d.amount) || 0,
        close_date: d.closeDate instanceof Date && !isNaN(d.closeDate.getTime()) ? toIsoDateOnly_(d.closeDate) : String(d.closeDate || '')
      });
    }
  }
  out.sort(function(a, b) {
    return (Number(b.amount) || 0) - (Number(a.amount) || 0);
  });
  return out.slice(0, 80);
}

function fmtSummaryMoneyCompact_(n) {
  var v = Number(n) || 0;
  if (Math.abs(v) >= 1000000) return '$' + (v / 1000000).toFixed(v % 1000000 === 0 ? 0 : 1) + 'M';
  if (Math.abs(v) >= 1000) return '$' + Math.round(v / 1000) + 'k';
  return '$' + Math.round(v).toLocaleString('en-US');
}

function topForecastDealsByCategory_(context, category, maxItems) {
  var list = Array.isArray(context && context.forecast_deal_context) ? context.forecast_deal_context : [];
  return list.filter(function(item) {
    return String(item.category || '') === String(category || '');
  }).sort(function(a, b) {
    return (Number(b.amount) || 0) - (Number(a.amount) || 0);
  }).slice(0, maxItems || 3);
}

function joinForecastDealTags_(deals) {
  return (deals || []).map(function(d) {
    return String(d.account || '').trim() + ' (' + fmtSummaryMoneyCompact_(d.amount) + ')';
  }).filter(Boolean).join(', ');
}

function buildDeterministicForecastSignals_(context) {
  var metrics = context && context.metrics ? context.metrics : {};
  var lines = [];
  var quota = Number(metrics.teamQuota) || 0;
  var closed = Number(metrics.closed) || 0;
  var commit = Number(metrics.commit) || 0;
  var ml = Number(metrics.mostLikely) || 0;
  var best = Number(metrics.bestCase) || 0;
  var nextQ = Number(metrics.nextQuarter) || 0;
  var nextQQuota = Number((context && context.admin_context && context.admin_context.next_quarter_quota) || 0);

  lines.push(
    'Team is at ' + fmtSummaryMoneyCompact_(closed) + ' closed, ' + fmtSummaryMoneyCompact_(commit) +
    ' commit, and ' + fmtSummaryMoneyCompact_(ml) + ' most likely against a ' + fmtSummaryMoneyCompact_(quota) + ' quota.'
  );

  var commitDeals = topForecastDealsByCategory_(context, 'Commit', 3);
  if (commitDeals.length) {
    lines.push(
      'Commit is concentrated in ' + joinForecastDealTags_(commitDeals) +
      '; any slippage would pressure in-quarter attainment.'
    );
  } else {
    lines.push(
      'Commit outside closed business is thin, so the path still depends on converting likely-stage deals.'
    );
  }

  var likelyDeals = topForecastDealsByCategory_(context, 'Most Likely', 3);
  if (likelyDeals.length) {
    lines.push(
      'Most-likely upside is led by ' + joinForecastDealTags_(likelyDeals) + '.'
    );
  } else {
    lines.push('Most-likely upside remains thin outside the current commit set.');
  }

  var bestDeals = topForecastDealsByCategory_(context, 'Best Case', 3);
  if (bestDeals.length) {
    lines.push(
      'Best-case upside is led by ' + joinForecastDealTags_(bestDeals) + '.'
    );
  } else if (nextQ > 0 && nextQQuota > 0) {
    lines.push(
      'Next-quarter coverage is ' + fmtSummaryMoneyCompact_(nextQ) + ' against a ' + fmtSummaryMoneyCompact_(nextQQuota) +
      ' quota and still needs disciplined qualification.'
    );
  } else {
    lines.push('Best-case upside is present but still needs cleaner qualification.');
  }

  while (lines.length < 4) {
    lines.push('Forecast quality still depends on disciplined deal inspection and category hygiene.');
  }

  return lines.slice(0, 4);
}

function buildSummaryAiInputContext_(quarterKey) {
  function short_(value, maxLen) {
    var s = String(value || '').trim();
    if (!s) return '';
    return s.length > maxLen ? s.slice(0, maxLen).trim() : s;
  }
  var admin = findLatestSummaryAdminRow_();
  var reportingDate = getEffectiveReportingDate_(admin.week_of);
  var quarter = quarterKey || quarterKeyFromDate_(reportingDate);
  var quotaCfg = findQuotaConfigForQuarter_(quarter);
  var nextQuarterKey = quarterKeyFromDate_(new Date(reportingDate.getFullYear(), reportingDate.getMonth() + 3, reportingDate.getDate()));
  var nextQuarterQuotaCfg = findQuotaConfigForQuarter_(nextQuarterKey);
  var teamData = getTeamDashboardData();
  var recapSheet = getRecapSheet_();
  var lastRow = recapSheet.getLastRow();
  var rows = lastRow > 1 ? recapSheet.getRange(2, 1, lastRow - 1, recapSheet.getLastColumn()).getValues() : [];
  var cutoff = new Date(Date.now() - 7 * 86400000);
  var byRep = {};
  for (var i = rows.length - 1; i >= 0; i--) {
    var ts = rows[i][0] instanceof Date ? rows[i][0] : new Date(rows[i][0]);
    if (!(ts instanceof Date) || isNaN(ts.getTime()) || ts < cutoff) continue;
    var email = String(rows[i][1] || '').trim().toLowerCase();
    if (!email || byRep[email]) continue;
    byRep[email] = rows[i];
  }

  var repRecaps = [];
  var repEmails = [];
  for (var emailKey in byRep) {
    var parsed = parseRecapRow_(byRep[emailKey]);
    var repName = getUserNameFromEmail(emailKey) || emailKey;
    repEmails.push(emailKey);
    var accounts = [];
    for (var a = 1; a <= 6; a++) {
      var acct = {
        name: short_(parsed['acct' + a + '_name'], 60),
        stage: short_(parsed['acct' + a + '_stage'], 24),
        arr: parseMoneyishServer_(parsed['acct' + a + '_arr']) || 0,
        engagement: short_(parsed['acct' + a + '_engage'], 12),
        critical_move: short_(parsed['acct' + a + '_move'], 140),
        next_step: short_(parsed['acct' + a + '_next'], 140)
      };
      if (acct.name || acct.critical_move || acct.next_step || acct.arr) accounts.push(acct);
    }
    accounts = accounts.slice(0, 4);
    repRecaps.push({
      rep_name: repName,
      pulse_score: Number(parsed.pulseScore) || 0,
      pulse_reason: short_(parsed.pulseReason, 220),
      rainmaker: {
        disco: Number(parsed.rm_disco) || 0,
        nbm: Number(parsed.rm_nbm) || 0,
        opp_arr: parseMoneyishServer_(parsed.rm_opp) || 0,
        target_accounts: Number(parsed.rm_accts) || 0
      },
      forecast: {
        commit: parseMoneyishServer_(parsed.commit) || 0,
        most_likely: parseMoneyishServer_(parsed.likely) || 0,
        best_case: parseMoneyishServer_(parsed.upside) || 0,
        next_quarter_commit: parseMoneyishServer_(parsed.nq_commit) || 0,
        note: short_(parsed.forecast_note, 180)
      },
      weekly_priority: short_(parsed.goal, 180),
      top_risk: short_(parsed.risk, 180),
      manager_ask: short_(parsed.ask, 180),
      accounts: accounts
    });
  }

  var executiveData = getExecutiveSummaryDataBase_ ? getExecutiveSummaryDataBase_(quarter) : null;
  return {
    team_name: quotaCfg.team_name || SUMMARY_TEAM_NAME,
    manager_name: quotaCfg.manager_name || 'Billy Schuett',
    reporting_week: toIsoDateOnly_(reportingDate),
    quarter_key: quarter,
    quarter_label: quotaCfg.quarter_label || quarterLabelFromKey_(quarter),
    metrics: executiveData ? executiveData.metrics : null,
    team_rollup: executiveData ? executiveData.teamRollup : null,
    admin_context: {
      theme: admin.theme || 'Weekly Summary',
      existing_leader_note: String(admin.leader_note || '').trim(),
      quota_notes: String(quotaCfg.notes || '').trim(),
      next_quarter_quota: Number(nextQuarterQuotaCfg.team_quota) || 0,
      next_quarter_label: nextQuarterQuotaCfg.quarter_label || quarterLabelFromKey_(nextQuarterKey)
    },
    rep_recaps: repRecaps
    ,
    forecast_deal_context: buildTeamForecastDealContext_(repEmails)
  };
}

function getSummaryAISchema_() {
  return {
    type: 'object',
    additionalProperties: false,
    required: ['generated_at', 'headline_verdict', 'leader_note', 'themes', 'big_deal_adds', 'deal_progression', 'rep_priorities', 'risks_asks_notes', 'top_risks', 'manager_asks', 'forecast_notes'],
    properties: {
      generated_at: { type: 'string' },
      headline_verdict: { type: 'string' },
      leader_note: { type: 'string' },
      themes: {
        type: 'array',
        maxItems: 3,
        items: { type: 'string' }
      },
      big_deal_adds: {
        type: 'array',
        maxItems: 1,
        items: { type: 'string' }
      },
      deal_progression: {
        type: 'array',
        maxItems: 4,
        items: { type: 'string' }
      },
      rep_priorities: {
        type: 'array',
        maxItems: 4,
        items: { type: 'string' }
      },
      risks_asks_notes: {
        type: 'array',
        minItems: 5,
        maxItems: 5,
        items: { type: 'string' }
      },
      top_risks: {
        type: 'array',
        maxItems: 2,
        items: { type: 'string' }
      },
      manager_asks: {
        type: 'array',
        maxItems: 2,
        items: { type: 'string' }
      },
      forecast_notes: {
        type: 'array',
        maxItems: 2,
        items: { type: 'string' }
      }
    }
  };
}

function getSummaryAIPrompt_() {
  return [
    'You are a world-class CRO + CEO chief of staff writing a one-page executive weekly summary for NorthCentral Enterprise, a Glean enterprise sales team run by Billy Schuett.',
    'This summary will be read at 7 AM on Monday by a senior executive with two minutes. Every word must earn its place.',
    'Your voice is sharp, current, commercial, and consequential. Think Bloomberg Businessweek editorial judgment meets CRO-grade forecast discipline.',
    '',
    'Rules for sections:',
    '',
    '- headline_verdict: ONE complete sentence, 10 to 18 words, that captures the single most important truth about this week. It is the newspaper headline a CRO would write if they could only say one thing. It should read as a verdict — what is actually happening this week — not a summary of the dashboard. Examples of the right voice: "Pipeline creation holds steady, but closed won is still pacing 43% behind quota with five weeks left." OR "Chamberlain legal drag and SPS positioning against Anthropic now define the quarter." Avoid empty phrases like "The team is focused on..." or "This week saw..." or "Overall performance is...". Never start with the team name. Never start with a date. Lead with the verdict itself.',
    '- leader_note: 2-3 sentences total, max 85 words, complete sentences only. This is the narrative context behind the headline_verdict.',
    '- big_deal_adds: at most 1 line, only if truly meaningful.',
    '- themes: maximum 3 items, each 4-7 words, crisp executive labels.',
    '- deal_progression and rep_priorities: each should contain 3 listed items whenever enough signal exists; drop to 2 only when fewer than 3 meaningful items exist.',
    '- each deal_progression and rep_priority item should be 10-16 words, specific enough to stand alone, one complete business-ready sentence, no fragments or ellipses.',
    '- deal_progression focuses on real commercial movement, not generic activity.',
    '- rep_priorities captures the highest-value execution priorities, not task noise.',
    '- risks_asks_notes contains exactly 5 items, 8-20 words each. These combine the most important risks, leadership asks, and forecast caveats.',
    '- In risks_asks_notes, if an item is a leadership ask, begin it with "Ask - Owner Name:" so the owner can be bolded in the UI.',
    '- top_risks and manager_asks and forecast_notes stay tight, 8-14 words per line.',
    '',
    'Hard rules:',
    '- Do not invent or import context from Glean, Salesforce, or outside systems unless it is present in the provided input context.',
    '- Do not infer quarter boundaries from calendar month names. The business quarter may not align to calendar quarters.',
    '- If a note says something may slip into April or another month, do not say it leaves the current quarter unless the source recap explicitly says it moves out of quarter, misses quarter, or impacts quarter attainment.',
    '- Prefer naming the account when the account is the real story. Name the rep only when that adds clarity.',
    '- Spread attention across the team intelligently. Do not repeat the same rep or account in multiple sections unless they genuinely own the week.',
    '- Default to the shorter phrasing when choosing between two options. One crisp sentence per item. No second clause unless it materially changes the decision.',
    '- No filler. No "N/A" or "no change" entries unless materially important.',
    '- Return valid JSON matching the schema exactly.'
  ].join('\n');
}
function extractResponseText_(json) {
  if (!json) return '';
  if (typeof json.output_text === 'string' && json.output_text.trim()) return json.output_text;
  if (json.output && json.output.length) {
    for (var i = 0; i < json.output.length; i++) {
      var item = json.output[i];
      if (item && item.content && item.content.length) {
        for (var j = 0; j < item.content.length; j++) {
          var c = item.content[j];
          if (c && typeof c.text === 'string' && c.text.trim()) return c.text;
        }
      }
    }
  }
  return '';
}

function extractJsonObjectString_(text) {
  var raw = String(text || '').trim();
  if (!raw) return '';
  raw = raw.replace(/^```(?:json)?/i, '').replace(/```$/i, '').trim();
  if (raw.charAt(0) === '{' && raw.charAt(raw.length - 1) === '}') return raw;
  var start = raw.indexOf('{');
  if (start === -1) return '';
  var depth = 0;
  var inString = false;
  var escaped = false;
  for (var i = start; i < raw.length; i++) {
    var ch = raw.charAt(i);
    if (inString) {
      if (escaped) {
        escaped = false;
      } else if (ch === '\\') {
        escaped = true;
      } else if (ch === '"') {
        inString = false;
      }
      continue;
    }
    if (ch === '"') {
      inString = true;
      continue;
    }
    if (ch === '{') depth++;
    if (ch === '}') {
      depth--;
      if (depth === 0) return raw.slice(start, i + 1);
    }
  }
  return '';
}

function validateSummaryAIPayload_(payload, context) {
  var out = payload || {};
  function arr_(key, maxItems, maxWords) {
    var arr = Array.isArray(out[key]) ? out[key] : [];
    return arr.map(function(v) { return trimSentenceWords_(String(v || '').trim(), maxWords || 22); }).filter(Boolean).slice(0, maxItems);
  }
  return {
    generated_at: String(out.generated_at || new Date().toISOString()),
    headline_verdict: trimSentenceWords_(String(out.headline_verdict || '').trim(), 22),
    leader_note: trimSentenceWords_(String(out.leader_note || '').trim(), 85),
    themes: arr_('themes', 3, 8),
    big_deal_adds: arr_('big_deal_adds', 1, 40),
    deal_progression: arr_('deal_progression', 4, 36),
    forecast_signals: buildDeterministicForecastSignals_(context).map(function(v) {
      return trimSentenceWords_(String(v || '').trim(), 32);
    }).filter(Boolean).slice(0, 4),
    rep_priorities: arr_('rep_priorities', 4, 36),
    risks_asks_notes: arr_('risks_asks_notes', 5, 28),
    top_risks: arr_('top_risks', 2, 28),
    manager_asks: arr_('manager_asks', 2, 28),
    forecast_notes: arr_('forecast_notes', 2, 28)
  };
}

function callGleanAgentSummary_(contextJsonString) {
  var GLEAN_BASE = 'https://scio-prod-be.glean.com';
  var AGENT_ID = 'ae756d1d02e54829803d49d540631a90';
  var token = PropertiesService.getScriptProperties().getProperty('GLEAN_API_TOKEN');
  if (!token) throw new Error('GLEAN_API_TOKEN not set in Script Properties.');

  var url = GLEAN_BASE + '/rest/api/v1/agents/runs/wait';
  var payload = {
    agent_id: AGENT_ID,
    input: { text: contextJsonString }
  };

  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var body = response.getContentText();
  if (code !== 200) {
    throw new Error('Glean agent run failed (' + code + '): ' + body);
  }

  var envelope;
  try { envelope = JSON.parse(body); }
  catch (e) { throw new Error('Glean response was not JSON: ' + body.slice(0, 500)); }

  var inner;
  try {
    inner = envelope.messages[0].content[0].text;
  } catch (e) {
    throw new Error('Glean response missing messages[0].content[0].text: ' + body.slice(0, 500));
  }

  // Strip possible markdown fences just in case.
  var cleaned = inner.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '').trim();

  var parsed;
  try { parsed = JSON.parse(cleaned); }
  catch (e) { throw new Error('Inner agent JSON failed to parse. First 500 chars: ' + cleaned.slice(0, 500)); }

  return parsed;
}

function callOpenAISummary_(context) {
  var apiKey = getOpenAIApiKey_();
  var inputJson = JSON.stringify(context);
  Logger.log('callOpenAISummary_: input bytes=' + inputJson.length);
  var payload = {
    model: SUMMARY_AI_MODEL,
    reasoning: { effort: 'low' },
    instructions: getSummaryAIPrompt_(),
    input: inputJson,
    max_output_tokens: 2600,
    text: {
      format: {
        type: 'json_schema',
        name: 'weekly_summary_payload',
        strict: true,
        schema: getSummaryAISchema_()
      }
    }
  };

  var resp = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  var body = resp.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('OpenAI request failed (' + code + '): ' + body.slice(0, 600));
  }

  var json = JSON.parse(body);
  var text = extractResponseText_(json);
  if (!text) throw new Error('OpenAI response did not include output text.');
  var jsonText = extractJsonObjectString_(text);
  if (!jsonText) {
    throw new Error('OpenAI response did not contain a complete JSON object. Raw text: ' + text.slice(0, 400));
  }
  return validateSummaryAIPayload_(JSON.parse(jsonText), context);
}

function generateExecutiveSummaryAI() {
  ensureAdminCaller_();
  var quarterKey = quarterKeyFromDate_(new Date());
  var context = buildSummaryAiInputContext_(quarterKey);

  try {
    var ai = callGleanAgentSummary_(JSON.stringify(context));

    // Validate AI output before publishing
    if (!ai || !ai.generated_at || !ai.leader_note) {
      throw new Error('AI output validation failed: missing required fields');
    }

    // Only write snapshot and cache if validation passed
    writeSummaryAISnapshot_(quarterKey, SUMMARY_AI_MODEL, ai, 'glean_agent', 'ok');
    putJsonCache_('summary_ai_' + quarterKey, ai, 60);

    return {
      success: true,
      quarterKey: quarterKey,
      generatedAt: ai.generated_at,
      model: SUMMARY_AI_MODEL,
      source: 'glean_agent_validated'
    };
  } catch (err) {
    // Log failure but don't overwrite good cache/snapshot with bad data
    var errorMsg = err.message || String(err);
    Logger.log('generateExecutiveSummaryAI failed: ' + errorMsg);
    writeSummaryAISnapshot_(quarterKey, SUMMARY_AI_MODEL, { error: errorMsg, timestamp: new Date().toISOString() }, 'glean_agent', 'failed');

    // Check if we have a previous good snapshot to fall back to
    var existingSnapshot = getLatestSummaryAISnapshot_(quarterKey);
    if (existingSnapshot) {
      Logger.log('AI generation failed, but existing snapshot available. Keeping existing snapshot.');
      return {
        success: false,
        quarterKey: quarterKey,
        error: errorMsg,
        fallbackAvailable: true,
        source: 'cached_fallback'
      };
    }

    return {
      success: false,
      quarterKey: quarterKey,
      error: errorMsg,
      fallbackAvailable: false,
      source: 'generation_failed'
    };
  }
}

function setupExecutiveSummaryAI() {
  ensureAdminCaller_();
  getOrCreateSummaryAISheet();
  var triggers = ScriptApp.getProjectTriggers();
  var hasTrigger = false;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction && triggers[i].getHandlerFunction() === 'generateExecutiveSummaryAI') {
      hasTrigger = true;
      break;
    }
  }
  if (!hasTrigger) {
    ScriptApp.newTrigger('generateExecutiveSummaryAI').timeBased().everyDays(1).atHour(6).create();
  }
  return generateExecutiveSummaryAI();
}
