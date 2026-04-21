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
