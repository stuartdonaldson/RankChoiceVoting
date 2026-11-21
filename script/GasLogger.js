/**
 * Drive-mapped structured logging for server-side test validation, with an
 * optional Axiom sink. Adapted from F3Go30's GasLogger.js, trimmed to this
 * project's needs.
 *
 * Creates one log file per execution run in a Drive subfolder (RankChoiceVoting/).
 * All GasLogger.log() entries within a run share the same execId and go
 * to the same file — one file per execution run, not one file per flush().
 *
 * Setup (via the admin JSON API — see WebApp.js's setScriptProperties action,
 * or run once from the GAS editor):
 *   setScriptProperty('GAS_LOGGER_PARENT_FOLDER_ID', '<Drive folder ID>');
 *
 * Axiom is optional: set AXIOM_TOKEN + AXIOM_DATASET script properties to
 * enable it. Once both are set, flush() POSTs to Axiom EXCLUSIVELY — it does
 * not also write the Drive file, even if the POST fails. A broken Axiom pipe
 * is meant to surface as a visible gap (Logger.log only), not be silently
 * absorbed by falling back to Drive. Unset either property to revert to
 * Drive-only behavior with zero code changes.
 *
 * Usage:
 *   GasLogger.run('triggerName', function() { ... });  // preferred — see below
 *
 *   // or manually:
 *   GasLogger.init('triggerName');           // call at start of each trigger
 *   GasLogger.log('tag', { key: value });    // accumulate + Logger.log()
 *   GasLogger.flush();                       // write to Drive or Axiom at end of execution
 *
 * GasLogger.run(triggerName, fn) wraps an entry point (simple trigger, web app
 * request, menu item) so init/flush happen automatically: it calls init(), runs
 * fn(), and flushes in a finally block so accumulated entries are written even if
 * fn() throws or returns early. Apps Script has no execution-end hook, so every
 * entry point that wants guaranteed flushing must go through run() (or call
 * flush() itself before every return path).
 *
 * Every entry is stamped by log() itself with `version` (APP_VERSION) and `target`
 * (APP_DEPLOY_TARGET, e.g. SIT/PROD) from version.js, so SIT and PROD runs stay
 * distinguishable in one shared Axiom dataset without depending on every call site
 * to add it.
 *
 * PII rule (rcballot-1j4): voter names and caller emails are logged as-is — the Axiom dataset
 * keeps 30 days, and an unmasked name is what lets a usage report say who voted. Secrets are
 * still redacted by scrubForLog_().
 */

/**
 * Maps GasLogger entries to Axiom ingest rows. Pure — no GAS globals — so it's
 * unit testable in Node. execId/runId are this project's correlation fields
 * (set by GasLogger.init()); included only when present on the entry. version/target
 * are stamped by GasLogger.log() onto every entry so SIT vs PROD runs are always
 * distinguishable in a shared dataset, without depending on every call site to add it;
 * fallbackVersion only covers entries built before that existed.
 * @param {Array<Object>} entries - Entries as built by GasLogger.log() (ts, tag, data, execId, runId?, version?, target?).
 * @param {string=} fallbackVersion - Used only when an entry has no version of its own.
 * @returns {Array<Object>} Axiom rows: { _time, name, side, version, target, ...data, execId?, runId? }.
 */
function buildAxiomRows_(entries, fallbackVersion) {
  return (entries || []).map(function(e) {
    var row = Object.assign({
      _time: e.ts,
      name: e.tag,
      side: 'gas',
      version: e.version || fallbackVersion,
      target: e.target || 'unknown'
    }, e.data || {});
    if (e.execId) row.execId = e.execId;
    if (e.actor) row.actor = e.actor;
    if (e.email) row.email = e.email;
    if (e.runId) row.runId = e.runId;
    return row;
  });
}

/**
 * Makes a request payload safe to log: secret-like keys redacted, `properties` reduced to its
 * key names, long strings/arrays truncated, nesting capped. Pure.
 * @param {*} value
 * @param {number=} depth
 * @returns {*}
 */
function scrubForLog_(value, depth) {
  depth = depth || 0;
  if (value === null || value === undefined) return value;
  if (typeof value === 'string') return value.length > 200 ? value.slice(0, 200) + '...(+' + (value.length - 200) + ')' : value;
  if (typeof value !== 'object') return value;
  if (depth > 3) return Array.isArray(value) ? '[array]' : '[object]';
  if (Array.isArray(value)) return value.slice(0, 30).map(function(v) { return scrubForLog_(v, depth + 1); });
  var out = {};
  Object.keys(value).forEach(function(k) {
    // assertion/credential/profile: the signed-in envelope (rcballot-dx3.2) — never logged.
    if (/secret|token|password|assertion|credential|profile/i.test(k)) out[k] = '[redacted]';
    else if (k === 'properties' && value[k] && typeof value[k] === 'object') out[k] = Object.keys(value[k]);
    else out[k] = scrubForLog_(value[k], depth + 1);
  });
  return out;
}

var GasLogger = {
  _folder: null,
  _entries: [],
  _enabled: true,
  _execId: null,
  _runId: null,
  _fileId: null,
  _axiomConfig: null,
  _actor: null,
  _email: null,

  /**
   * Stamps the caller's email (rcballot-8un: owner policy — every log carries the user's email
   * when known) on every later entry of this execution AND back-fills entries already buffered,
   * so the request line logged before identity verification is attributed too. This is the
   * verified Google email (or the Apps Script session user).
   * @param {string} email
   */
  setEmail: function(email) {
    var e = email ? String(email).trim().toLowerCase() : '';
    this._email = e || null;
    if (!e) return;
    this._entries.forEach(function(entry) { if (!entry.email) entry.email = e; });
  },

  /**
   * Stamps `actor` (an identity, e.g. the entered voter name or 'admin') on every
   * later entry of this execution, so one person's calls can be tied together.
   * @param {string} actor
   */
  setActor: function(actor) {
    this._actor = actor ? String(actor) : null;
  },
  /**
   * Call at the start of every trigger function.
   * Generates a fresh execId and reads the optional test runId from Script Properties.
   * @param {string} triggerName - Caller name for the init log line.
   * @returns {string} The generated execId.
   */
  init: function(triggerName) {
    this._execId = Utilities.getUuid();
    this._runId = PropertiesService.getScriptProperties().getProperty('RCV_TEST_RUN_ID') || null;
    this._entries = [];
    this._fileId = null;
    this._actor = null;
    this._email = null;
    // Fallback identity: the Apps Script session user (sheet menus/triggers, the deployer).
    // Anonymous web-app callers give ''. A Dispatch-verified email replaces it (Identity.js).
    try {
      var sessionEmail = typeof Session !== 'undefined' ? Session.getActiveUser().getEmail() : '';
      if (sessionEmail) this._email = String(sessionEmail).trim().toLowerCase();
    } catch (err) { /* no session user available */ }
    Logger.log('[GasLogger] init — trigger=' + triggerName +
      ' execId=' + this._execId + (this._runId ? ' runId=' + this._runId : ''));
    return this._execId;
  },

  _getFolder: function() {
    if (this._folder) return this._folder;
    var parentId = PropertiesService.getScriptProperties().getProperty('GAS_LOGGER_PARENT_FOLDER_ID');
    if (!parentId) {
      Logger.log('[GasLogger] GAS_LOGGER_PARENT_FOLDER_ID not set — Drive writes disabled');
      return null;
    }
    try {
      var parent = DriveApp.getFolderById(parentId);
      var iter = parent.getFoldersByName('RankChoiceVoting');
      this._folder = iter.hasNext() ? iter.next() : parent.createFolder('RankChoiceVoting');
      return this._folder;
    } catch (e) {
      Logger.log('[GasLogger] _getFolder failed: ' + e);
      return null;
    }
  },

  _getAxiomConfig: function() {
    if (this._axiomConfig) return this._axiomConfig;
    var props = PropertiesService.getScriptProperties();
    this._axiomConfig = {
      token: props.getProperty('AXIOM_TOKEN'),
      dataset: props.getProperty('AXIOM_DATASET')
    };
    return this._axiomConfig;
  },

  _postToAxiom: function(entries) {
    var config = this._getAxiomConfig();
    var rows = buildAxiomRows_(entries);
    try {
      var resp = UrlFetchApp.fetch(
        'https://api.axiom.co/v1/datasets/' + config.dataset + '/ingest',
        {
          method: 'post',
          contentType: 'application/json',
          headers: { Authorization: 'Bearer ' + config.token },
          payload: JSON.stringify(rows),
          muteHttpExceptions: true
        }
      );
      if (resp.getResponseCode() >= 300) {
        // Visible in clasp logs (Stackdriver) only — never recurse through GasLogger.log()
        // itself, and intentionally not written to Drive either (see file header).
        Logger.log('[GasLogger] Axiom ingest non-2xx ' + resp.getResponseCode() + ': ' + resp.getContentText());
      }
    } catch (e) {
      Logger.log('[GasLogger] Axiom POST threw: ' + e);
    }
  },

  /**
   * Accumulate a structured log entry. Also writes to Logger.log().
   * @param {string}  tag    - Entry type (e.g. 'webapp.ballot', 'handleAdminPost_.error').
   * @param {Object}  data   - Payload. Must not contain voter names or secrets. The caller's
   *                           email is added at entry level via setEmail()/init(), not in data.
   * @param {boolean} flush  - If true, flush accumulated entries to Drive/Axiom immediately.
   * @param {boolean} newLog - If true, reset the file reference after flushing so the
   *                           next flush() creates a new file for subsequent entries.
   */
  log: function(tag, data, flush, newLog) {
    if (!this._execId) this.init('auto');
    // version/target stamped here (not just at Axiom-export time) so every entry — Drive
    // or Axiom — always carries which build and which deployment target produced it,
    // without depending on every call site to remember to add it.
    var version = (typeof APP_VERSION !== 'undefined' && APP_VERSION) || 'unknown';
    var target = (typeof APP_DEPLOY_TARGET !== 'undefined' && APP_DEPLOY_TARGET) || 'unknown';
    var entry = { ts: new Date().toISOString(), tag: tag, data: data, execId: this._execId, version: version, target: target };
    if (this._runId) entry.runId = this._runId;
    if (this._actor) entry.actor = this._actor;
    if (this._email) entry.email = this._email;
    Logger.log('[GasLogger] ' + JSON.stringify(entry));
    if (this._enabled) this._entries.push(entry);
    if (flush) this.flush();
    if (newLog) this._fileId = null;
  },

  /**
   * Write accumulated entries to Drive or Axiom. Call once at the end of each trigger
   * function (or rely on run()'s finally block).
   */
  flush: function() {
    if (!this._enabled || this._entries.length === 0) return;

    var axiomConfig = this._getAxiomConfig();
    if (axiomConfig.token && axiomConfig.dataset) {
      this._postToAxiom(this._entries);
      this._entries = [];
      return;
    }

    var folder = this._getFolder();
    if (!folder) { this._entries = []; return; }
    var content = this._entries.map(function(e) { return JSON.stringify(e); }).join('\n');
    try {
      if (this._fileId) {
        // Append to the execution-run file (read + overwrite — Drive has no native append)
        var existing = DriveApp.getFileById(this._fileId).getBlob().getDataAsString();
        var newContent = existing + (existing ? '\n' : '') + content;
        UrlFetchApp.fetch(
          'https://www.googleapis.com/upload/drive/v3/files/' + this._fileId + '?uploadType=media',
          {
            method: 'PATCH',
            contentType: 'text/plain; charset=UTF-8',
            payload: newContent,
            headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
            muteHttpExceptions: true
          }
        );
      } else {
        // First flush for this execution run — create the file
        var filename = new Date().getTime() + '-' + (this._execId || Utilities.getUuid()) + '.log';
        var file = folder.createFile(filename, content, MimeType.PLAIN_TEXT);
        this._fileId = file.getId();
      }
    } catch (e) {
      Logger.log('[GasLogger] flush failed: ' + e);
    }
    this._entries = [];
  },

  enable: function() { this._enabled = true; },
  disable: function() { this._enabled = false; },

  /**
   * Standardizes a caught-exception log entry so every hand-rolled try/catch logs the
   * same shape as run()'s own catch: message + stack, plus whatever call-site context is
   * useful (e.g. { action: payload.action }).
   * @param {string} tag   - Entry tag (e.g. 'handleAdminPost_.error').
   * @param {Error}  err   - The caught exception.
   * @param {Object=} extra - Extra fields merged in (e.g. { action: payload.action }).
   */
  logError: function(tag, err, extra) {
    var data = Object.assign({ message: err && err.message, stack: err && err.stack }, extra || {});
    this.log(tag, data);
  },

  /**
   * Wraps an entry-point function with init/flush so callers don't have to manage the
   * lifecycle by hand. On error, logs the error entry, flushes, then rethrows.
   * @param {string}   triggerName - Passed to init(); identifies this execution in logs.
   * @param {Function} fn          - The entry-point body. No arguments — close over them.
   * @returns {*} fn()'s return value.
   */
  run: function(triggerName, fn) {
    this.init(triggerName);
    try {
      return fn();
    } catch (e) {
      this.log('error', { message: e && e.message, stack: e && e.stack });
      throw e;
    } finally {
      this.flush();
    }
  }
};

/** Run from GAS editor to configure without opening Script Properties UI manually. */
function setScriptProperty(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
}

function getScriptProperty(key) {
  Logger.log(PropertiesService.getScriptProperties().getProperty(key));
}

/**
 * Run this ONCE from the Apps Script editor (select it in the function dropdown, click Run)
 * after adding/changing oauthScopes in appsscript.json. Apps Script grants scopes
 * incrementally — only a call that actually exercises a scope (e.g. UrlFetchApp.fetch for
 * https://www.googleapis.com/auth/script.external_request) will trigger that scope's consent
 * prompt; running an unrelated function (e.g. getScriptProperty, which only touches
 * PropertiesService) will silently skip it, leaving the deployed web app's UrlFetchApp calls
 * (GasLogger's Axiom ingest, Drive PATCH) throwing "You do not have permission" with no
 * interactive prompt possible (there's no user session to prompt during an anonymous web
 * app request). Grant consent here first; the deployed web app then inherits it.
 */
function authorizeExternalRequestScope() {
  var resp = UrlFetchApp.fetch('https://api.axiom.co/v1/datasets', { muteHttpExceptions: true });
  Logger.log('authorizeExternalRequestScope: HTTP ' + resp.getResponseCode());
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    GasLogger: GasLogger,
    buildAxiomRows_: buildAxiomRows_,
    scrubForLog_: scrubForLog_
  };
}
