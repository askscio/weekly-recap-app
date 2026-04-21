// =======================================================================
// RAINMAKER — Glean Agents API Integration Layer
// =======================================================================
// This file provides server-side infrastructure for calling Glean's Agents
// API from Apps Script. The Glean API token must be stored in Script
// Properties as GLEAN_API_TOKEN (never in source code).
//
// Part 1 of N in Rainmaker workstream — scaffolding only, no UI yet.
// =======================================================================

// -----------------------------------------------------------------------
// TOKEN MANAGEMENT
// -----------------------------------------------------------------------
function getGleanApiToken_() {
  var token = PropertiesService.getScriptProperties().getProperty('GLEAN_API_TOKEN');
  if (!token || token.trim() === '') {
    throw new Error(
      'Missing GLEAN_API_TOKEN in Script Properties. ' +
      'Set it via: PropertiesService.getScriptProperties().setProperty("GLEAN_API_TOKEN", "your-token-here")'
    );
  }
  return token.trim();
}

// -----------------------------------------------------------------------
// API CONNECTION TESTING
// -----------------------------------------------------------------------
function testGleanApiConnection_() {
  try {
    var token = getGleanApiToken_();

    // Cheap health check: call the agents endpoint with an empty agent_id.
    // Expected behavior:
    // - 401/403: token is invalid
    // - 400/422: token is valid (endpoint rejected empty agent_id as expected)
    // - 200: unexpected but means endpoint is reachable
    var payload = {
      agent_id: '',
      input: {
        messages: [{ role: 'user', content: 'test' }]
      }
    };

    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + token
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(GLEAN_AGENT_WAIT_ENDPOINT, options);
    var statusCode = response.getResponseCode();
    var responseText = response.getContentText();

    Logger.log('testGleanApiConnection_: status=' + statusCode + ', body=' + responseText.slice(0, 200));

    // Interpret results
    if (statusCode === 401 || statusCode === 403) {
      return {
        ok: false,
        status: statusCode,
        message: 'Token is invalid or expired (received ' + statusCode + '). Check GLEAN_API_TOKEN in Script Properties.'
      };
    }

    if (statusCode === 400 || statusCode === 422) {
      // Endpoint returned expected error for empty/invalid agent_id — token is valid
      return {
        ok: true,
        status: statusCode,
        message: 'Connection successful. Token is valid (endpoint returned expected ' + statusCode + ' for test payload).'
      };
    }

    if (statusCode === 200) {
      return {
        ok: true,
        status: statusCode,
        message: 'Connection successful. Token is valid (endpoint returned 200).'
      };
    }

    // Any other status code
    return {
      ok: false,
      status: statusCode,
      message: 'Unexpected status code ' + statusCode + '. Response: ' + responseText.slice(0, 500)
    };

  } catch (err) {
    Logger.log('testGleanApiConnection_ error: ' + err.message);
    return {
      ok: false,
      status: null,
      message: 'Connection test failed: ' + err.message
    };
  }
}

// -----------------------------------------------------------------------
// CORE API WRAPPER
// -----------------------------------------------------------------------
function callGleanAgent_(agentId, prompt, opts) {
  var options = opts || {};
  var retries = options.retries !== undefined ? options.retries : 3;
  var backoffMs = options.backoffMs !== undefined ? options.backoffMs : 2000;

  if (!agentId || agentId.trim() === '') {
    throw new Error('callGleanAgent_: agentId is required');
  }

  if (!prompt || prompt.trim() === '') {
    throw new Error('callGleanAgent_: prompt is required');
  }

  var token = getGleanApiToken_();
  var payload = {
    agent_id: agentId.trim(),
    input: {
      messages: [{ role: 'user', content: prompt.trim() }]
    }
  };

  var requestOptions = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + token
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var attempt = 0;
  var lastError = null;

  while (attempt <= retries) {
    attempt++;

    try {
      Logger.log('callGleanAgent_: attempt ' + attempt + '/' + (retries + 1) + ', agentId=' + agentId);

      var response = UrlFetchApp.fetch(GLEAN_AGENT_WAIT_ENDPOINT, requestOptions);
      var statusCode = response.getResponseCode();
      var responseText = response.getContentText();

      // Success
      if (statusCode === 200) {
        Logger.log('callGleanAgent_: success on attempt ' + attempt);
        try {
          return JSON.parse(responseText);
        } catch (parseErr) {
          throw new Error('callGleanAgent_: response was 200 but body is not valid JSON. Body: ' + responseText.slice(0, 500));
        }
      }

      // Retry on 429 (rate limit) and 5xx (server errors)
      if (statusCode === 429 || (statusCode >= 500 && statusCode < 600)) {
        lastError = new Error('callGleanAgent_: received ' + statusCode + ' on attempt ' + attempt + '. Response: ' + responseText.slice(0, 500));
        Logger.log(lastError.message);

        if (attempt <= retries) {
          var waitMs = backoffMs * Math.pow(2, attempt - 1);
          Logger.log('callGleanAgent_: waiting ' + waitMs + 'ms before retry...');
          Utilities.sleep(waitMs);
          continue;
        }
      }

      // Non-retryable 4xx errors (except 429)
      if (statusCode >= 400 && statusCode < 500 && statusCode !== 429) {
        throw new Error('callGleanAgent_: received non-retryable ' + statusCode + '. Response: ' + responseText.slice(0, 500));
      }

      // Any other unexpected status
      throw new Error('callGleanAgent_: unexpected status ' + statusCode + '. Response: ' + responseText.slice(0, 500));

    } catch (err) {
      lastError = err;
      Logger.log('callGleanAgent_: error on attempt ' + attempt + ': ' + err.message);

      // If it's a known non-retryable error, throw immediately
      if (err.message.indexOf('non-retryable') !== -1) {
        throw err;
      }

      // Otherwise retry if attempts remain
      if (attempt <= retries) {
        var waitMs = backoffMs * Math.pow(2, attempt - 1);
        Logger.log('callGleanAgent_: waiting ' + waitMs + 'ms before retry...');
        Utilities.sleep(waitMs);
        continue;
      }
    }
  }

  // All retries exhausted
  throw new Error('callGleanAgent_: all ' + (retries + 1) + ' attempts failed. Last error: ' + (lastError ? lastError.message : 'unknown'));
}

// -----------------------------------------------------------------------
// RESPONSE PARSING
// -----------------------------------------------------------------------
function extractAgentResponseText_(gleanResponse) {
  if (!gleanResponse || !gleanResponse.output || !Array.isArray(gleanResponse.output.messages)) {
    throw new Error('extractAgentResponseText_: invalid response structure. Expected output.messages array.');
  }

  var messages = gleanResponse.output.messages;

  // Walk from the end to find the last assistant message
  for (var i = messages.length - 1; i >= 0; i--) {
    var msg = messages[i];
    if (msg && msg.role === 'assistant' && msg.content) {
      return String(msg.content).trim();
    }
  }

  throw new Error('extractAgentResponseText_: no assistant message found in response. Messages: ' + JSON.stringify(messages).slice(0, 500));
}

function parseAgentJsonResponse_(text) {
  if (!text || text.trim() === '') {
    throw new Error('parseAgentJsonResponse_: text is empty');
  }

  var cleaned = text.trim();

  // Strip markdown JSON code fences if present
  if (cleaned.indexOf('```json') !== -1 || cleaned.indexOf('```') !== -1) {
    // Try to extract content between ```json and ``` or just ``` and ```
    var jsonMatch = cleaned.match(/```(?:json)?\s*\n?([\s\S]*?)\n?```/);
    if (jsonMatch && jsonMatch[1]) {
      cleaned = jsonMatch[1].trim();
    }
  }

  // Try direct parse first
  try {
    return JSON.parse(cleaned);
  } catch (directErr) {
    // Direct parse failed, try to extract first {...} or [...] block
    Logger.log('parseAgentJsonResponse_: direct parse failed, trying regex extraction');

    // Try to find first complete JSON object
    var objMatch = cleaned.match(/\{[\s\S]*\}/);
    if (objMatch) {
      try {
        return JSON.parse(objMatch[0]);
      } catch (objErr) {
        // Continue to array attempt
      }
    }

    // Try to find first complete JSON array
    var arrMatch = cleaned.match(/\[[\s\S]*\]/);
    if (arrMatch) {
      try {
        return JSON.parse(arrMatch[0]);
      } catch (arrErr) {
        // Continue to final error
      }
    }

    // Nothing worked, throw with raw text preview
    throw new Error('parseAgentJsonResponse_: could not parse JSON. Raw text (first 500 chars): ' + text.slice(0, 500));
  }
}

// -----------------------------------------------------------------------
// PUBLIC ADMIN ENTRY POINTS
// -----------------------------------------------------------------------
function smokeTestGleanConnection() {
  ensureAdminCaller_();

  Logger.log('smokeTestGleanConnection: starting connection test...');
  var result = testGleanApiConnection_();

  Logger.log('smokeTestGleanConnection: test complete');
  Logger.log('  ok: ' + result.ok);
  Logger.log('  status: ' + result.status);
  Logger.log('  message: ' + result.message);

  return result;
}

// -----------------------------------------------------------------------
// RAINMAKER DATA REFRESH — End-to-End Pipeline
// -----------------------------------------------------------------------
function runRainmakerRefresh() {
  var startTime = new Date();
  var summary = {
    success: false,
    reps_scored: 0,
    reps_with_errors: 0,
    ent_rows_total: 0,
    runtime_seconds: 0,
    error: null
  };

  try {
    Logger.log('runRainmakerRefresh: starting Rainmaker data refresh...');

    // Collect per-rep data from Agent 1
    var repsData = [];
    var fiscalQuarter = null;

    for (var i = 0; i < REP_EMAILS.length; i++) {
      var repEmail = REP_EMAILS[i];
      try {
        Logger.log('runRainmakerRefresh: calling Agent 1 for ' + repEmail);

        var agentResponse = callGleanAgent_(
          RAINMAKER_AGENT_ID_REP_SCORECARD,
          JSON.stringify({ rep_email: repEmail })
        );

        var responseText = extractAgentResponseText_(agentResponse);
        var repData = parseAgentJsonResponse_(responseText);

        // Skip if rep_excluded
        if (repData.rep_excluded) {
          Logger.log('runRainmakerRefresh: ' + repEmail + ' is excluded, skipping');
          continue;
        }

        // Store fiscal quarter from first successful rep
        if (!fiscalQuarter && repData.fiscal_quarter) {
          fiscalQuarter = repData.fiscal_quarter;
        }

        repsData.push(repData);
        summary.reps_scored++;

        if (repData.errors && Array.isArray(repData.errors) && repData.errors.length > 0) {
          summary.reps_with_errors++;
          Logger.log('runRainmakerRefresh: ' + repEmail + ' has errors: ' + repData.errors.join('; '));
        }

        Logger.log('runRainmakerRefresh: successfully processed ' + repEmail);

      } catch (repErr) {
        Logger.log('runRainmakerRefresh: FAILED for ' + repEmail + ': ' + repErr.message);
        summary.reps_with_errors++;
        // Continue to next rep
      }
    }

    Logger.log('runRainmakerRefresh: Agent 1 calls complete. Scored ' + summary.reps_scored + ' reps.');

    // Call Agent 2 for Enterprise benchmarks
    var entBenchmarks = null;
    try {
      Logger.log('runRainmakerRefresh: calling Agent 2 for Enterprise benchmarks');

      var agent2Response = callGleanAgent_(
        RAINMAKER_AGENT_ID_ENT_BENCHMARKS,
        '{}'
      );

      var ent2Text = extractAgentResponseText_(agent2Response);
      entBenchmarks = parseAgentJsonResponse_(ent2Text);

      Logger.log('runRainmakerRefresh: successfully retrieved Enterprise benchmarks');

    } catch (entErr) {
      Logger.log('runRainmakerRefresh: Agent 2 FAILED: ' + entErr.message);
      Logger.log('runRainmakerRefresh: continuing without Enterprise data');
      // Continue without Enterprise data
    }

    // Get/create sheets
    var ss = SpreadsheetApp.openById(TRACKER_SHEET_ID);
    var rainmakerSheet = getOrCreateSheet_(ss, RAINMAKER_SHEET_NAME);
    var entSheet = getOrCreateSheet_(ss, RAINMAKER_SHEET_NAME + '_Ent');

    Logger.log('runRainmakerRefresh: clearing existing data...');

    // Clear both sheets completely
    if (rainmakerSheet.getLastRow() > 0) {
      rainmakerSheet.getRange(1, 1, rainmakerSheet.getMaxRows(), rainmakerSheet.getMaxColumns()).clearContent();
    }
    if (entSheet.getLastRow() > 0) {
      entSheet.getRange(1, 1, entSheet.getMaxRows(), entSheet.getMaxColumns()).clearContent();
    }

    Logger.log('runRainmakerRefresh: writing per-rep data to Rainmaker sheet...');

    // Write Rainmaker sheet header
    var rainmakerHeaders = [
      'rep_email', 'rep_name', 'fiscal_quarter', 'nbm', 'pipe_adds', 'pipe_dollars',
      'c_level', 'stage4_plus', 'closed_won', 'pocs', 'partner_reg', 'errors', 'generated_at'
    ];
    rainmakerSheet.appendRow(rainmakerHeaders);
    rainmakerSheet.getRange(1, 1, 1, rainmakerHeaders.length).setFontWeight('bold');
    rainmakerSheet.setFrozenRows(1);

    // Write rep data rows
    var now = new Date();
    for (var j = 0; j < repsData.length; j++) {
      var rd = repsData[j];
      var errorsText = (rd.errors && Array.isArray(rd.errors) && rd.errors.length > 0)
        ? rd.errors.join('; ')
        : 'none';

      var row = [
        rd.rep_email || '',
        rd.rep_name || '',
        rd.fiscal_quarter || fiscalQuarter || '',
        rd.nbm || 0,
        rd.pipe_adds || 0,
        rd.pipe_dollars || 0,
        rd.c_level || 0,
        rd.stage4_plus || 0,
        rd.closed_won || 0,
        rd.pocs || 0,
        rd.partner_reg || 0,
        errorsText,
        now
      ];
      rainmakerSheet.appendRow(row);
    }

    Logger.log('runRainmakerRefresh: wrote ' + repsData.length + ' rep rows');
    Logger.log('runRainmakerRefresh: writing Enterprise benchmark data to Rainmaker_Ent sheet...');

    // Write Enterprise benchmark sheet header
    var entHeaders = ['category', 'owner_name', 'owner_email', 'value'];
    entSheet.appendRow(entHeaders);
    entSheet.getRange(1, 1, 1, entHeaders.length).setFontWeight('bold');
    entSheet.setFrozenRows(1);

    // Write Enterprise benchmark rows
    var entRowsTotal = 0;
    if (entBenchmarks && entBenchmarks.rows_by_category) {
      var categories = ['nbm', 'pipe_adds', 'pipe_dollars', 'stage4_plus', 'closed_won', 'partner_reg'];

      for (var k = 0; k < categories.length; k++) {
        var cat = categories[k];
        var rows = entBenchmarks.rows_by_category[cat];

        if (rows && Array.isArray(rows)) {
          for (var m = 0; m < rows.length; m++) {
            var entRow = rows[m];
            entSheet.appendRow([
              cat,
              entRow.owner_name || '',
              entRow.owner_email || '',
              entRow.value || 0
            ]);
            entRowsTotal++;
          }
          Logger.log('runRainmakerRefresh: wrote ' + rows.length + ' rows for category: ' + cat);
        }
      }
    } else {
      Logger.log('runRainmakerRefresh: no Enterprise benchmark data available');
    }

    summary.ent_rows_total = entRowsTotal;

    // Calculate runtime
    var endTime = new Date();
    summary.runtime_seconds = Math.round((endTime.getTime() - startTime.getTime()) / 1000 * 10) / 10;

    // Log summary
    Logger.log('runRainmakerRefresh: COMPLETE ✓');
    Logger.log('  Reps scored: ' + summary.reps_scored);
    Logger.log('  Reps with errors: ' + summary.reps_with_errors);
    Logger.log('  Enterprise rows written: ' + summary.ent_rows_total);
    Logger.log('  Runtime: ' + summary.runtime_seconds + 's');

    // Append to log sheet
    var logSheet = getOrCreateRainmakerLogSheet_(ss);
    logSheet.appendRow([
      now,
      fiscalQuarter || '',
      summary.reps_scored,
      summary.reps_with_errors,
      summary.ent_rows_total,
      summary.runtime_seconds,
      'Success'
    ]);

    summary.success = true;
    return summary;

  } catch (err) {
    Logger.log('runRainmakerRefresh: FATAL ERROR: ' + err.message);
    Logger.log('runRainmakerRefresh: Stack trace: ' + err.stack);

    summary.error = err.message;
    summary.runtime_seconds = Math.round((new Date().getTime() - startTime.getTime()) / 1000 * 10) / 10;

    // Try to log the failure
    try {
      var ss = SpreadsheetApp.openById(TRACKER_SHEET_ID);
      var logSheet = getOrCreateRainmakerLogSheet_(ss);
      logSheet.appendRow([
        new Date(),
        '',
        summary.reps_scored,
        summary.reps_with_errors,
        summary.ent_rows_total,
        summary.runtime_seconds,
        'ERROR: ' + err.message.slice(0, 200)
      ]);
    } catch (logErr) {
      Logger.log('runRainmakerRefresh: could not write error to log: ' + logErr.message);
    }

    return summary;
  }
}

// -----------------------------------------------------------------------
// SHEET HELPERS
// -----------------------------------------------------------------------
function getOrCreateSheet_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('getOrCreateSheet_: creating new sheet: ' + sheetName);
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function getOrCreateRainmakerLogSheet_(ss) {
  var sheet = ss.getSheetByName(RAINMAKER_LOG_SHEET_NAME);
  if (!sheet) {
    Logger.log('getOrCreateRainmakerLogSheet_: creating new log sheet');
    sheet = ss.insertSheet(RAINMAKER_LOG_SHEET_NAME);
    sheet.appendRow([
      'timestamp',
      'fiscal_quarter',
      'reps_scored',
      'reps_with_errors',
      'ent_rows_total',
      'runtime_seconds',
      'notes'
    ]);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}
