/**
 * Revitin Support Chats (Suntek) — Resumable Analyzer (Hardened, AI-only tags)
 * - Per-channel Sentiment • Tags (incl. health_safety + query) • Short Summary (+ solved?)
 * - Writes rows as it goes; safe to stop/resume; auto-resumes via time trigger
 * - AI decides tags; we only filter to the approved list (unknowns → 'other')
 * - Builds Tag Summary and Query Subcategories from output
 */

/* =================== Config =================== */

// Name of the tab with your raw data (must contain channel + body headers)
const SOURCE_SHEET_NAME = 'RAW_DATA';   // <-- CHANGE to your input sheet name

const OPENAI_URL = "https://api.openai.com/v1/chat/completions";
const MODEL = "gpt-4o-mini";
const TEMPERATURE = 0.0;

const CHUNK_SIZE     = 40;    // channels per execution slice (smaller = safer)
const TRANSCRIPT_CAP = 8000;  // max characters sent per channel
const SLEEP_MS       = 200;   // pacing between calls
const MAX_RETRIES    = 5;     // OpenAI retries with backoff

// Approved tag taxonomy (+ health_safety + query) with 'other' fallback
const APPROVED_TAGS = [
  'refund_issue',
  'billing_issue',
  'subscription_change',
  'order_status',
  'shipping_issue',
  'product_quality',
  'packaging_issue',
  'pricing_value',
  'promotion_issue',     // ONLY coupons/discounts/codes
  'technical_issue',
  'account_issue',
  'marketing_spam',
  'complaint',
  'positive_feedback',
  'negative_feedback',
  'health_safety',       // ingredients / safety / toxicity / side-effects
  'query',               // general informational questions
  'other'
];

// Subcategory basis for "query" rollup
const QUERY_TOPICAL_TAGS = [
  'order_status',
  'shipping_issue',
  'billing_issue',
  'subscription_change',
  'product_quality',
  'health_safety',
  'packaging_issue',
  'promotion_issue',
  'pricing_value',
  'technical_issue',
  'account_issue'
];

// Flexible header aliases (lowercased)
const CHANNEL_HEADER_ALIASES = [
  'channel_identifier','channel id','channel_id','channel','conversation_id','thread_id','ticket_id'
].map(function(s){ return s.toLowerCase(); });

const TEXT_HEADER_ALIASES = [
  'body','message','messages','chat','text','content','transcript','chat_body'
].map(function(s){ return s.toLowerCase(); });

// State key (document-scoped)
const STATE_KEY = 'REVITIN_ANALYSIS_STATE_V3'; // {offset,total,channelsProcessed,startedAtISO}

/* =================== Menu =================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AI Support Chat Analysis')
    .addItem('Start Full Run (Auto-Resume)', 'startFullRun')
    .addItem('Resume Now', 'resumeFullRun')
    .addSeparator()
    .addItem('Analyze Selected Range', 'analyzeSelectedRange')
    .addItem('Build Tag Summary (from output)', 'buildTagSummaryFromOutput')
    .addItem('Build Query Subcategories (from output)', 'buildQuerySubcategoriesFromOutput')
    .addSeparator()
    .addItem('Diagnose Selection', 'diagnoseSelection')
    .addItem('Debug Triggers & State', 'debugTriggersAndState')
    .addToUi();
}

/* =================== Entrypoints =================== */

function startFullRun() {
  _prepareOutputSheet(); // ensure header; DO NOT clear existing rows
  _saveState({ offset: 0, total: -1, channelsProcessed: 0, startedAtISO: new Date().toISOString() });
  _processNextChunk(true);   // run first chunk immediately; next chunks will self-schedule
}

function resumeFullRun() {
  _processNextChunk(true);   // processes one chunk and schedules next if more remains
}

function analyzeSelectedRange() {
  var range = SpreadsheetApp.getActiveRange();
  if (!range) { SpreadsheetApp.getUi().alert('Select some rows first.'); return; }
  _analyzeRangeOnce(range);
}

/* =================== Core Resumable Engine =================== */

function _processNextChunk(manual) {
  return _withDocLock('_processNextChunk', function () {
    var key = _getApiKey();
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Use configured input sheet (do NOT depend on active tab)
    var inSheet = ss.getSheetByName(SOURCE_SHEET_NAME) || SpreadsheetApp.getActiveSheet();
    if (!inSheet) { _clearState(); SpreadsheetApp.getUi().alert('Input sheet not found: ' + SOURCE_SHEET_NAME); return; }

    var values = inSheet.getDataRange().getValues();
    if (values.length < 2) { _clearState(); SpreadsheetApp.getUi().alert('No data in input sheet: ' + SOURCE_SHEET_NAME); return; }

    var headers = values[0].map(function(h){ return String(h || '').trim().toLowerCase(); });
    var chIdx = _findHeaderIdx(headers, CHANNEL_HEADER_ALIASES);
    var txIdx = _findHeaderIdx(headers, TEXT_HEADER_ALIASES);
    if (chIdx === -1 || txIdx === -1) {
      _clearState();
      SpreadsheetApp.getUi().alert(
        'Missing headers. Found: [' + headers.join(', ') + ']\n' +
        'Need one of: ' + CHANNEL_HEADER_ALIASES.join(', ') + '\n' +
        'and one of: ' + TEXT_HEADER_ALIASES.join(', ')
      );
      return;
    }

    // Group rows by channel
    var byChannel = {};
    for (var r = 1; r < values.length; r++) {
      var channel = String(values[r][chIdx] || '').trim();
      var text    = String(values[r][txIdx] || '').trim();
      if (!channel || !text) continue;
      if (!byChannel[channel]) byChannel[channel] = [];
      byChannel[channel].push(text);
    }
    var allChannels = Object.keys(byChannel);
    if (!allChannels.length) { _clearState(); SpreadsheetApp.getUi().alert('No usable rows.'); return; }

    // Load state & determine remaining (skip outputs that already exist)
    var st = _loadState();
    if (!st) st = { offset: 0, total: -1, channelsProcessed: 0, startedAtISO: new Date().toISOString() };
    if (st.total < 0) st.total = allChannels.length;

    var out = _prepareOutputSheet();
    var existing = _readAnalyzedChannels(out);
    var remaining = allChannels.filter(function(ch){ return !existing.has(ch); });

    var start = st.offset;
    if (start >= remaining.length) {
      buildTagSummaryFromOutput();
      _clearState();
      if (manual) SpreadsheetApp.getUi().alert('All done.');
      return;
    }

    var slice = remaining.slice(start, Math.min(start + CHUNK_SIZE, remaining.length));
    SpreadsheetApp.getActive().toast(
      'Processing channels ' + (start+1) + '-' + (start+slice.length) + ' of ' + remaining.length + '…',
      'AI Support Chat Analysis',
      5
    );

    var writeRow = _nextEmptyRow(out);

    for (var c = 0; c < slice.length; c++) {
      var channel = slice[c];
      var full = (byChannel[channel] || []).join('\n---\n');
      var transcript = full.substring(0, TRANSCRIPT_CAP);

      var payload = {
        model: MODEL,
        temperature: TEMPERATURE,
        response_format: { type: "json_object" },
        messages: [
          { role: "system", content: _buildSystemPrompt() },
          { role: "user",   content: _buildUserMessage(transcript) }
        ]
      };

      var parsed = null, errMsg = null;
      for (var attempt = 1; attempt <= MAX_RETRIES; attempt++) {
        try {
          var resp = _httpPostJson(OPENAI_URL, payload, key);
          var content = '{}';
          if (resp && resp.choices && resp.choices[0] && resp.choices[0].message &&
              typeof resp.choices[0].message.content === 'string') {
            content = resp.choices[0].message.content;
          }
          parsed = _safeParseJson(content) || {};
          break;
        } catch (e) {
          errMsg = String(e);
          _setLastError(errMsg);
          var backoff = Math.min(2000 * attempt, 8000); // 2s,4s,6s,8s,8s
          Utilities.sleep(backoff);
        }
      }

      // Sentiment
      var sentiment = _oneOf(parsed && parsed.sentiment, ['positive','neutral','negative']) || 'neutral';

      // ==== SIMPLE TAG NORMALIZATION (no keyword hints, no synonyms) ====
      var tagsArr = (parsed && parsed.tags && parsed.tags.map)
        ? parsed.tags.map(function(t){ return String(t || '').trim().toLowerCase().replace(/\s+/g,'_'); })
        : [];

      var seen = {};
      var finalTags = [];
      for (var j = 0; j < tagsArr.length; j++) {
        var tNorm = tagsArr[j];
        if (APPROVED_TAGS.indexOf(tNorm) >= 0 && !seen[tNorm]) {
          seen[tNorm] = true;
          finalTags.push(tNorm);
        }
        if (finalTags.length >= 6) break; // cap
      }
      if (!finalTags.length) finalTags = ['other'];
      // ==== end normalization ====

      // Summary + solved
      var solved = _oneOf(String((parsed && parsed.solved) || '').toLowerCase(), ['yes','no','unclear']) || 'unclear';
      var baseSummary = String((parsed && parsed.summary) || (errMsg ? 'API error: ' + errMsg : '')).trim();
      var summary = baseSummary ? (baseSummary + ' (Solved: ' + solved + ')') : ('(Solved: ' + solved + ')');

      // Write row immediately
      out.getRange(writeRow, 1, 1, 4).setValues([[channel, _cap(sentiment), finalTags.join(', '), summary]]);
      writeRow++;

      Utilities.sleep(SLEEP_MS);
    }

    SpreadsheetApp.flush();

    // Update state
    st.offset = start + slice.length;
    st.channelsProcessed = (st.channelsProcessed || 0) + slice.length;
    _saveState(st);

    // ALWAYS schedule the next chunk when more remains
    var moreRemaining = st.offset < remaining.length;
    if (moreRemaining) {
      _scheduleNext(1); // ~1 minute
      if (manual) SpreadsheetApp.getUi().alert('Processed ' + slice.length + ' channels. Next chunk is auto-scheduled.');
    } else {
      buildTagSummaryFromOutput();
      _clearState();
      if (manual) SpreadsheetApp.getUi().alert('All done.');
    }
  });
}

/* =================== One-off (selection) =================== */

function _analyzeRangeOnce(range) {
  var key = _getApiKey();
  var values = range.getValues();
  if (values.length < 2) { SpreadsheetApp.getUi().alert('Need a header row + data.'); return; }

  var headers = values[0].map(function(h){ return String(h || '').trim().toLowerCase(); });
  var chIdx = _findHeaderIdx(headers, CHANNEL_HEADER_ALIASES);
  var txIdx = _findHeaderIdx(headers, TEXT_HEADER_ALIASES);
  if (chIdx === -1 || txIdx === -1) {
    SpreadsheetApp.getUi().alert('Missing headers in the selection.'); return;
  }

  var byChannel = {};
  for (var r = 1; r < values.length; r++) {
    var channel = String(values[r][chIdx] || '').trim();
    var text    = String(values[r][txIdx] || '').trim();
    if (!channel || !text) continue;
    if (!byChannel[channel]) byChannel[channel] = [];
    byChannel[channel].push(text);
  }
  var channels = Object.keys(byChannel);
  if (!channels.length) { SpreadsheetApp.getUi().alert('No usable rows in selection.'); return; }

  var out = _prepareOutputSheet();
  var existing = _readAnalyzedChannels(out);
  var writeRow = _nextEmptyRow(out);

  for (var c = 0; c < channels.length; c++) {
    var channel = channels[c];
    if (existing.has(channel)) continue;

    var full = byChannel[channel].join('\n---\n');
    var transcript = full.substring(0, TRANSCRIPT_CAP);

    var payload = {
      model: MODEL,
      temperature: TEMPERATURE,
      response_format: { type: "json_object" },
      messages: [
        { role: "system", content: _buildSystemPrompt() },
        { role: "user",   content: _buildUserMessage(transcript) }
      ]
    };

    var parsed = null, errMsg = null;
    for (var attempt = 1; attempt <= MAX_RETRIES; attempt++) {
      try {
        var resp = _httpPostJson(OPENAI_URL, payload, _getApiKey());
        var content = '{}';
        if (resp && resp.choices && resp.choices[0] && resp.choices[0].message &&
            typeof resp.choices[0].message.content === 'string') {
          content = resp.choices[0].message.content;
        }
        parsed = _safeParseJson(content) || {};
        break;
      } catch (e) {
        errMsg = String(e);
        _setLastError(errMsg);
        var backoff = Math.min(2000 * attempt, 8000);
        Utilities.sleep(backoff);
      }
    }

    var sentiment = _oneOf(parsed && parsed.sentiment, ['positive','neutral','negative']) || 'neutral';

    var tagsArr = (parsed && parsed.tags && parsed.tags.map)
      ? parsed.tags.map(function(t){ return String(t || '').trim().toLowerCase().replace(/\s+/g,'_'); })
      : [];

    var seen = {};
    var finalTags = [];
    for (var j = 0; j < tagsArr.length; j++) {
      var tNorm = tagsArr[j];
      if (APPROVED_TAGS.indexOf(tNorm) >= 0 && !seen[tNorm]) {
        seen[tNorm] = true;
        finalTags.push(tNorm);
      }
      if (finalTags.length >= 6) break;
    }
    if (!finalTags.length) finalTags = ['other'];

    var solved = _oneOf(String((parsed && parsed.solved) || '').toLowerCase(), ['yes','no','unclear']) || 'unclear';
    var baseSummary = String((parsed && parsed.summary) || (errMsg ? 'API error: ' + errMsg : '')).trim();
    var summary = baseSummary ? (baseSummary + ' (Solved: ' + solved + ')') : ('(Solved: ' + solved + ')');

    out.getRange(writeRow, 1, 1, 4).setValues([[channel, _cap(sentiment), finalTags.join(', '), summary]]);
    writeRow++;

    Utilities.sleep(SLEEP_MS);
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Selection analyzed: ' + channels.length + ' channel(s).');
}

/* =================== Prompt =================== */

function _buildSystemPrompt() {
  return [
    "You are a data analyst at Suntek analyzing a single support chat transcript for Revitin.",
    "Return STRICT JSON with keys: sentiment, tags, summary, solved.",
    "- sentiment: one of [positive, neutral, negative].",
    "- tags: choose 2–6 ONLY from this approved list; use 'other' only if none reasonably fit:",
    "  " + APPROVED_TAGS.join(', '),
    "- summary: 1–2 sentences summarizing what the customer wanted and why.",
    "- solved: one of [yes, no, unclear] — did the issue appear resolved?",
    "Tagging rules (AI decides from content only—no keyword heuristics):",
    "• Use 'health_safety' for ingredient/safety/toxicity/side-effects topics (e.g., fluoride, nano hydroxyapatite, nanoparticles, SLS, parabens).",
    "• Use 'promotion_issue' ONLY for coupons/discounts/promo codes not applying.",
    "• For general information requests (e.g., availability, formula questions), include 'query' plus any topical tag if applicable.",
    "• Use 'positive_feedback' only when the user expresses praise; 'negative_feedback' only for dissatisfaction.",
    "• Prefer accurate topical tags over sentiment-only tags when the text is neutral.",
    "• Use only tags from the list verbatim. Do not invent new tags."
  ].join('\n');
}

function _buildUserMessage(transcript) {
  return JSON.stringify({ transcript: transcript });
}

/* =================== Summaries from Output =================== */

function buildTagSummaryFromOutput() {
  var out = _getOrCreateSheet('AI_Chat_Analysis');
  var lastRow = out.getLastRow();
  var rows = out.getRange(2,1, Math.max(0,lastRow-1), 4).getValues(); // channel, sentiment, tags, summary

  var tally = {};
  for (var i = 0; i < APPROVED_TAGS.length; i++) {
    tally[APPROVED_TAGS[i]] = { channels: new Set(), positive:0, neutral:0, negative:0 };
  }

  for (var r = 0; r < rows.length; r++) {
    var ch = rows[r][0], sentiment = String(rows[r][1] || '').toLowerCase();
    var tags = String(rows[r][2] || '').split(',').map(function(s){ return s.trim().toLowerCase(); });
    var unique = {};
    for (var t = 0; t < tags.length; t++) if (tags[t]) unique[tags[t]] = true;

    for (var tag in unique) {
      if (!tally[tag]) tally[tag] = { channels: new Set(), positive:0, neutral:0, negative:0 };
      tally[tag].channels.add(ch);
      if (sentiment === 'positive' || sentiment === 'neutral' || sentiment === 'negative') {
        tally[tag][sentiment]++;
      }
    }
  }

  var sh = _getOrCreateSheet('AI_Tag_Summary');
  sh.clear();
  var header = ['tag','channels','positive','neutral','negative'];
  var outRows = [header];
  for (var i2 = 0; i2 < APPROVED_TAGS.length; i2++) {
    var tagName = APPROVED_TAGS[i2];
    var rec = tally[tagName] || { channels: new Set(), positive:0, neutral:0, negative:0 };
    outRows.push([
      tagName,
      rec.channels.size || 0,
      rec.positive || 0,
      rec.neutral || 0,
      rec.negative || 0
    ]);
  }
  sh.getRange(1,1,outRows.length,outRows[0].length).setValues(outRows);
  sh.getRange(1,1,1,outRows[0].length).setFontWeight('bold');
  sh.autoResizeColumns(1, outRows[0].length);

  // Also refresh query subcategories
  buildQuerySubcategoriesFromOutput();

  SpreadsheetApp.getUi().alert('Tag summary built from output.');
}

// Build "AI_Query_Subcategories" by splitting 'query' rows by their topical co-tags
function buildQuerySubcategoriesFromOutput() {
  var out = _getOrCreateSheet('AI_Chat_Analysis');
  var lastRow = out.getLastRow();
  var rows = out.getRange(2,1, Math.max(0,lastRow-1), 4).getValues(); // [channel, sentiment, tags, summary]

  var tally = {}; // subcat -> { channels:Set, positive, neutral, negative }

  for (var r = 0; r < rows.length; r++) {
    var ch = String(rows[r][0] || '').trim();
    if (!ch) continue;
    var sentiment = String(rows[r][1] || '').toLowerCase();
    var tagStr = String(rows[r][2] || '');
    if (!tagStr) continue;

    var tags = tagStr.split(',').map(function(s){ return s.trim().toLowerCase(); });
    var hasQuery = false;
    var tagSet = {};
    for (var t = 0; t < tags.length; t++) { tagSet[tags[t]] = true; if (tags[t] === 'query') hasQuery = true; }
    if (!hasQuery) continue;

    var coTopicals = [];
    for (var i = 0; i < QUERY_TOPICAL_TAGS.length; i++) {
      var ct = QUERY_TOPICAL_TAGS[i];
      if (tagSet[ct]) coTopicals.push(ct);
    }
    if (coTopicals.length === 0) coTopicals.push('general');

    for (var j = 0; j < coTopicals.length; j++) {
      var subcat = 'query:' + coTopicals[j];
      if (!tally[subcat]) tally[subcat] = { channels: new Set(), positive:0, neutral:0, negative:0 };
      tally[subcat].channels.add(ch);
      if (sentiment === 'positive' || sentiment === 'neutral' || sentiment === 'negative') {
        tally[subcat][sentiment]++;
      }
    }
  }

  var sh = _getOrCreateSheet('AI_Query_Subcategories');
  sh.clear();
  var header = ['subcategory','channels','positive','neutral','negative'];
  var outRows = [header];

  var keys = Object.keys(tally).sort(function(a,b){
    if (a === 'query:general') return 1;
    if (b === 'query:general') return -1;
    return a < b ? -1 : (a > b ? 1 : 0);
  });

  for (var k = 0; k < keys.length; k++) {
    var name = keys[k];
    var rec = tally[name];
    outRows.push([ name, rec.channels.size || 0, rec.positive || 0, rec.neutral || 0, rec.negative || 0 ]);
  }

  sh.getRange(1,1,outRows.length,outRows[0].length).setValues(outRows);
  sh.getRange(1,1,1,outRows[0].length).setFontWeight('bold');
  sh.autoResizeColumns(1, outRows[0].length);

  SpreadsheetApp.getUi().alert('Query subcategories built from output.');
}

/* =================== Utilities & Hardening =================== */

function _httpPostJson(url, obj, apiKey) {
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(obj),
    muteHttpExceptions: true,
    followRedirects: true
  };
  var res = UrlFetchApp.fetch(url, options);
  var code = res.getResponseCode();
  var body = res.getContentText();
  if (code >= 300) {
    throw new Error('OpenAI HTTP ' + code + ': ' + body.slice(0, 500));
  }
  var json = _safeParseJson(body);
  if (!json) throw new Error('Bad JSON from OpenAI');
  return json;
}

function _safeParseJson(text) { try { return JSON.parse(text); } catch (e) { return null; } }
function _sleepMs(ms) { Utilities.sleep(ms); }

function _setLastError(msg) {
  PropertiesService.getDocumentProperties().setProperty('LAST_ERROR', new Date().toISOString() + ' :: ' + msg);
}

function _getApiKey() {
  var key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!key) throw new Error('Set OPENAI_API_KEY in Script Properties first.');
  return key;
}

function _getOrCreateSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function _prepareOutputSheet() {
  var out = _getOrCreateSheet('AI_Chat_Analysis');
  if (out.getLastRow() === 0) {
    out.getRange(1,1,1,4).setValues([['channel_identifier','sentiment','tags','summary']]);
  } else {
    var h = out.getRange(1,1,1,4).getValues()[0];
    if (String(h[0] || '') !== 'channel_identifier') {
      out.clear();
      out.getRange(1,1,1,4).setValues([['channel_identifier','sentiment','tags','summary']]);
    }
  }
  return out;
}

function _readAnalyzedChannels(outSheet) {
  var lastRow = outSheet.getLastRow();
  var set = new Set();
  if (lastRow <= 1) return set;
  var col = outSheet.getRange(2,1,lastRow-1,1).getValues();
  for (var i = 0; i < col.length; i++) {
    var ch = String(col[i][0] || '').trim();
    if (ch) set.add(ch);
  }
  return set;
}

function _nextEmptyRow(sh) {
  var lr = sh.getLastRow();
  return lr > 0 ? lr + 1 : 2;
}

function _findHeaderIdx(headers, aliases) {
  // Exact
  for (var i = 0; i < headers.length; i++) if (aliases.indexOf(headers[i]) >= 0) return i;
  // Relaxed
  var norm = function(h){ return h.replace(/[^a-z0-9]+/g,' ').trim().replace(/\s+/g,' '); };
  var aliasNorm = aliases.map(norm);
  for (var j = 0; j < headers.length; j++) if (aliasNorm.indexOf(norm(headers[j])) >= 0) return j;
  return -1;
}

function _oneOf(val, arr) {
  val = String(val || '').toLowerCase().trim();
  return arr.indexOf(val) >= 0 ? val : null;
}

function _cap(s) {
  s = String(s || '').toLowerCase();
  return s ? s.charAt(0).toUpperCase() + s.slice(1) : s;
}

/* =================== Locking & Triggers =================== */

function _withDocLock(fnName, workFn) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    Logger.log(fnName + ': could not obtain lock; exiting');
    return false;
  }
  try {
    workFn();
    return true;
  } finally {
    lock.releaseLock();
  }
}

function _loadState() {
  var json = PropertiesService.getDocumentProperties().getProperty(STATE_KEY);
  if (!json) return null;
  try { return JSON.parse(json); } catch (e) { return null; }
}

function _saveState(obj) {
  PropertiesService.getDocumentProperties().setProperty(STATE_KEY, JSON.stringify(obj));
}

function _clearState() {
  PropertiesService.getDocumentProperties().deleteProperty(STATE_KEY);
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === '_autoResumeTrigger') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function _scheduleNext(minutesFromNow) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === '_autoResumeTrigger') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('_autoResumeTrigger')
    .timeBased()
    .after(Math.max(1, minutesFromNow) * 60 * 1000)
    .create();
}

function _autoResumeTrigger() {
  try {
    _processNextChunk(false);
  } catch (e) {
    Logger.log('autoResumeTrigger failed: ' + e);
    _setLastError(String(e));
  }
}

/* =================== Diagnostics =================== */

function diagnoseSelection() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inSheet = ss.getSheetByName(SOURCE_SHEET_NAME) || SpreadsheetApp.getActiveSheet();
  var range = inSheet.getDataRange();
  var values = range.getValues();
  if (values.length < 2) { SpreadsheetApp.getUi().alert('Need header row + at least 1 data row.'); return; }

  var headers = values[0].map(function(h){ return String(h||'').trim().toLowerCase(); });
  var chIdx = _findHeaderIdx(headers, CHANNEL_HEADER_ALIASES);
  var txIdx = _findHeaderIdx(headers, TEXT_HEADER_ALIASES);

  var withBoth = 0, uniqueChannels = {};
  for (var r = 1; r < values.length; r++) {
    var ch = String(values[r][chIdx] || '').trim();
    var tx = String(values[r][txIdx] || '').trim();
    if (ch && tx) { withBoth++; uniqueChannels[ch] = true; }
  }

  var uniqueCount = Object.keys(uniqueChannels).length;
  SpreadsheetApp.getUi().alert(
    'Input sheet: ' + inSheet.getName() +
    '\nHeaders: [' + headers.join(', ') + ']' +
    '\nchannel idx: ' + chIdx + ', body idx: ' + txIdx +
    '\nRows with both fields present: ' + withBoth +
    '\nUnique channels found: ' + uniqueCount
  );
}

function debugTriggersAndState() {
  var dp = PropertiesService.getDocumentProperties();
  var st = dp.getProperty(STATE_KEY);
  var lastErr = dp.getProperty('LAST_ERROR');
  var triggers = ScriptApp.getProjectTriggers().map(function(t){
    return t.getHandlerFunction();
  });
  SpreadsheetApp.getUi().alert(
    'STATE: ' + (st || '(none)') +
    '\nLAST_ERROR: ' + (lastErr || '(none)') +
    '\nTRIGGERS:\n- ' + (triggers.join('\n- ') || 'none')
  );
}
