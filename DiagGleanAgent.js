/**
 * DIAGNOSTIC — probes the Glean agent wait endpoint to discover:
 *   (a) whether the endpoint accepts our auth token
 *   (b) the correct request shape to pass a chat message
 *   (c) the raw response wrapper structure
 *   (d) how the agent's JSON output is embedded in the wrapper
 *
 * This file is TEMPORARY. It will be deleted after we understand
 * the response shape and write the production call site.
 *
 * Run from the Apps Script editor: diagGleanAgentEndpoint
 */

var DIAG_GLEAN_BASE = 'https://scio-prod-be.glean.com';
var DIAG_AGENT_ID = 'ae756d1d02e54829803d49d540631a90';

function diagGleanAgentEndpoint() {
  var token = PropertiesService.getScriptProperties().getProperty('GLEAN_API_TOKEN');
  if (!token) {
    Logger.log('ERROR: GLEAN_API_TOKEN not set in Script Properties');
    return;
  }
  Logger.log('Token present, length=' + token.length);

  // A minimal test payload. This is a small hand-written JSON
  // that roughly matches our weekly recap schema. We're testing
  // whether the agent accepts input AND what the response shape
  // looks like. Real payload will be much larger.
  var testPayload = {
    reporting_week: '2026-04-20',
    recaps: [
      {
        rep_name: 'Niko Argaez',
        submitted: true,
        pulse_score: 8,
        pulse_reason: 'Chamberlain cleared legal this week — biggest unlock of the quarter.',
        forecast_note: 'Moving Chamberlain from Most Likely to Commit.',
        goal: 'Close Chamberlain procurement review.',
        risk: 'Zurich still pushing meetings.',
        ask: 'Billy — join the Chamberlain procurement call Tuesday.'
      }
    ],
    current_2x4_snapshot: [
      {
        rep_name: 'Niko Argaez',
        accounts: [
          { name: 'Chamberlain', stage: 6, arr: 1200000, engagement: 'Green', critical_moment: 'Procurement Tuesday', next_step: 'Billy joins call' }
        ]
      }
    ],
    movements: [
      { rep_name: 'Niko Argaez', account: 'Chamberlain', type: 'stage_jump', detail: 'Stage 3 -> 6, ARR $1.2M, legal cleared' }
    ],
    reps_missing_this_week: []
  };

  var payloadJsonString = JSON.stringify(testPayload);
  Logger.log('Test payload JSON length: ' + payloadJsonString.length + ' bytes');

  // TRY A: input as structured object with 'text' key (guess #1 — common chat pattern)
  tryShape_('A: input as { text: <json string> }', token, {
    agent_id: DIAG_AGENT_ID,
    input: { text: payloadJsonString }
  });

  // TRY B: input as structured object with 'message' key (guess #2)
  tryShape_('B: input as { message: <json string> }', token, {
    agent_id: DIAG_AGENT_ID,
    input: { message: payloadJsonString }
  });

  // TRY C: messages array with role/content (OpenAI-like pattern)
  tryShape_('C: messages array with role/content', token, {
    agent_id: DIAG_AGENT_ID,
    messages: [{ role: 'user', content: payloadJsonString }]
  });

  // TRY D: just query field (Glean search-like)
  tryShape_('D: { query: <json string> } at top level', token, {
    agent_id: DIAG_AGENT_ID,
    query: payloadJsonString
  });

  // TRY E: empty input, see what error message says about required fields
  tryShape_('E: empty input {}', token, {
    agent_id: DIAG_AGENT_ID,
    input: {}
  });

  // Also fetch the agent's input schema to see what Glean expects
  try {
    var schemaUrl = DIAG_GLEAN_BASE + '/rest/api/v1/agents/' + DIAG_AGENT_ID + '/schemas';
    var schemaResp = UrlFetchApp.fetch(schemaUrl, {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    Logger.log('--- AGENT SCHEMAS ---');
    Logger.log('Schema fetch HTTP: ' + schemaResp.getResponseCode());
    Logger.log('Schema body (first 3000 chars): ' + String(schemaResp.getContentText()).substring(0, 3000));
  } catch (e) {
    Logger.log('Schema fetch threw: ' + e.message);
  }
}

function tryShape_(label, token, bodyObj) {
  Logger.log('');
  Logger.log('========== TRY ' + label + ' ==========');
  try {
    var resp = UrlFetchApp.fetch(DIAG_GLEAN_BASE + '/rest/api/v1/agents/runs/wait', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify(bodyObj),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    var body = resp.getContentText();
    Logger.log('HTTP: ' + code);
    Logger.log('Body length: ' + body.length);
    Logger.log('Body (first 4000 chars): ' + body.substring(0, 4000));
  } catch (e) {
    Logger.log('Fetch threw: ' + e.message);
  }
}
