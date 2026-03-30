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
    source || 'openai_responses',
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
    required: ['generated_at', 'leader_note', 'themes', 'big_deal_adds', 'deal_progression', 'rep_priorities', 'risks_asks_notes', 'top_risks', 'manager_asks', 'forecast_notes'],
    properties: {
      generated_at: { type: 'string' },
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
    'You are the best CRO, CEO chief of staff, and business communications director in enterprise software.',
    'You are writing a one-page weekly executive summary for NorthCentral Enterprise.',
    'Your output must sound sharp, current, commercial, and consequential.',
    'Prioritize what matters most to the business this week: deal movement, forecast quality, execution blockers, manager intervention, and real next steps.',
    'Do not output filler. Do not output generic snippets. Do not output weak updates like no change, all good, or N/A unless they are materially important.',
    'Do not repeat the same account or rep over and over if others contributed meaningful signal. Spread attention intelligently across the team.',
    'Write with crisp executive judgment and communications polish.',
    'Keep everything concise enough to fit a one-page summary.',
    'Do not infer quarter boundaries from calendar month names. The business quarter may not align to calendar quarters.',
    'If a note says something may slip into April or another month, do not say it leaves the current quarter unless the source recap explicitly says it moves out of quarter, misses quarter, or impacts quarter attainment.',
    'Do not invent or import context from Glean, Salesforce, or outside systems unless it is present in the provided input context.',
    'Rules for sections:',
    '- leader_note: 2-3 sentences total, maximum about 85 words, complete sentences only.',
    '- big_deal_adds: at most 1 line, only if truly meaningful.',
    '- Deal Progression and Rep Priorities should each contain 3 listed items whenever enough signal exists; only drop to 2 if there are genuinely fewer than 3 meaningful items.',
    '- those sections should read like sharp bullet-point lists, not stitched paragraphs.',
    '- avoid ellipses and sentence fragments in those sections.',
    '- each Deal Progression and Rep Priorities item should usually be 10 to 16 words, with enough specificity to stand alone on the page.',
    '- Forecast Signals should be extremely tight and executive, usually 8 to 14 words per line.',
    '- Risks, Asks, and Notes stays compressed in its combined section and should be materially shorter than a normal update.',
    '- Other sections should stay concise.',
    '- each line should be a complete business-ready sentence, not a fragment.',
    '- each line should name the rep only if that adds clarity. Prefer naming the account when the account is the real story.',
    '- deal_progression should focus on real commercial movement, not generic activity.',
    '- rep_priorities should capture the highest-value execution priorities, not task noise.',
    '- risks_asks_notes should contain exactly 5 numbered-worthy items total, combining the most important risks, leadership asks, and forecast caveats for the week.',
    '- Across the right side of the page, reduce total word count by roughly 40% versus a normal operating summary. Prefer omission over over-explaining. Default to the shorter version when deciding between two phrasings.',
    '- Use one crisp sentence per item. No second clause unless it materially changes the decision.',
    '- In risks_asks_notes, if an item is a leadership ask, begin it with Ask - Owner Name: so the owner can be bolded in the UI.',
    '- risks_asks_notes should favor the most consequential notes across the whole team, not one bucket at a time.',
    '- top_risks should capture the sharpest blockers or downside risks.',
    '- manager_asks should capture only asks that actually require leadership involvement.',
    '- forecast_notes should elevate only meaningful notes that affect how the forecast should be read.',
    'Return valid JSON matching the schema exactly.'
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
    var ai = callOpenAISummary_(context);

    // Validate AI output before publishing
    if (!ai || !ai.generated_at || !ai.leader_note) {
      throw new Error('AI output validation failed: missing required fields');
    }

    // Only write snapshot and cache if validation passed
    writeSummaryAISnapshot_(quarterKey, SUMMARY_AI_MODEL, ai, 'openai_responses', 'ok');
    putJsonCache_('summary_ai_' + quarterKey, ai, 60);

    return {
      success: true,
      quarterKey: quarterKey,
      generatedAt: ai.generated_at,
      model: SUMMARY_AI_MODEL,
      source: 'openai_validated'
    };
  } catch (err) {
    // Log failure but don't overwrite good cache/snapshot with bad data
    var errorMsg = err.message || String(err);
    Logger.log('generateExecutiveSummaryAI failed: ' + errorMsg);
    writeSummaryAISnapshot_(quarterKey, SUMMARY_AI_MODEL, { error: errorMsg, timestamp: new Date().toISOString() }, 'openai_responses', 'failed');

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
