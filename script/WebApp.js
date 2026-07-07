/**
 * WebApp.js
 *
 * doGet/doPost router for the RankChoiceVoting web app. Modeled on F3Go30's
 * WebApp.js, trimmed to this project's needs.
 *
 *   (no cmd)               — home page: list all ballots, with links to view
 *                            results/edit each, plus a create-new-ballot form
 *                            (same content as ?cmd=admin — see webAdmin.js)
 *   ?cmd=ballot&id=<name>  — ballot page + RPCs (see webBallot.js)
 *   ?cmd=admin             — ballot list/create/edit/analyze (see webAdmin.js)
 *
 * doPost ?cmd=admin dispatches a minimal, best-practice set of administrative/
 * diagnostic JSON actions (see handleAdminPost_ below), gated by
 * ADMIN_SHARED_SECRET once bootstrapped. tools/callWebapp.js is the CLI client
 * for this API; it is what the deploy pipeline and integration tests use to
 * stamp WEBAPP_URL, configure script properties (e.g. Axiom credentials), and
 * download/upload sheet data to verify state.
 *
 * Deployed as a Web app (Deploy > New deployment > Web app). The single active
 * named deployment per script project is updated by tools/manage-deployments.js.
 */

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Never includes postData.contents — request bodies (cmd=admin) may carry
 * secrets/voter data, and GasLogger.log() data must never contain either.
 * Only type/length are safe to log.
 */
function buildWebAppRequestLog_(e) {
  return {
    cmd: (e && e.parameter && e.parameter.cmd) || '',
    queryString: (e && e.queryString) || null,
    postData: e && e.postData ? { type: e.postData.type, length: e.postData.length } : null
  };
}

function doGet(e) {
  return GasLogger.run('doGet', function () {
    GasLogger.log('doGet', buildWebAppRequestLog_(e));
    var cmd = (e && e.parameter && e.parameter.cmd) || '';

    if (cmd === 'ballot') {
      return _handleBallot(e);
    }
    if (cmd === 'admin' || !cmd) {
      return _handleAdmin(e);
    }

    return HtmlService.createHtmlOutput(
      '<p>Unknown cmd "' + cmd + '". Append <code>?cmd=ballot&amp;id=&lt;name&gt;</code> to open a ballot, ' +
      'or <code>?cmd=admin</code> (or no cmd at all) for the ballot list.</p>'
    );
  });
}

function doPost(e) {
  return GasLogger.run('doPost', function () {
    GasLogger.log('doPost', buildWebAppRequestLog_(e));
    var cmd = e && e.parameter && e.parameter.cmd;

    if (cmd === 'admin') {
      return handleAdminPost_(e);
    }

    return jsonOutput_({ ok: false, error: 'unknown_cmd' });
  });
}

/**
 * Sets ADMIN_SHARED_SECRET the first time only — whoever calls this first owns the
 * secret going forward. Never settable again via the web app; clearing it requires
 * the Apps Script editor's Script Properties UI by hand.
 */
function bootstrapAdminSecret_(secret) {
  if (!secret || String(secret).length < 16) {
    return { ok: false, error: 'secret must be at least 16 characters' };
  }
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('ADMIN_SHARED_SECRET')) {
    return { ok: false, error: 'already_bootstrapped' };
  }
  props.setProperty('ADMIN_SHARED_SECRET', String(secret));
  GasLogger.log('bootstrapAdminSecret_.bootstrapped', {});
  return { ok: true };
}

/**
 * Dispatches a cmd=admin doPost JSON body to administrative/diagnostic actions, gated by
 * ADMIN_SHARED_SECRET (set once via bootstrapSecret — never typed in by hand). Every other
 * action must echo the secret back in the POST body (never the query string, so it never
 * lands in access logs / curl history).
 *
 * This is deliberately a minimal set — config, URL stamping, sheet diagnostics/data I/O, one
 * domain action (createBallot), and one cleanup action (deleteSheet) — rather than growing an
 * action per feature. Add to this list only when a genuinely new capability is needed.
 */
function handleAdminPost_(e) {
  var payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput_({ ok: false, error: 'invalid_json' });
  }

  if (payload.action === 'bootstrapSecret') {
    return jsonOutput_(bootstrapAdminSecret_(payload.secret));
  }

  // setWebappUrl is handled BEFORE the secret gate: it only stores the running deployment's
  // own exec URL (non-sensitive, idempotent) and is called by tools/manage-deployments.js on
  // every PROD deploy — on a fresh project the admin secret isn't bootstrapped yet, so gating
  // it would deadlock the very first deploy's URL stamp.
  if (payload.action === 'setWebappUrl') {
    var stampUrl = ScriptApp.getService().getUrl();
    PropertiesService.getScriptProperties().setProperty('WEBAPP_URL', stampUrl);
    GasLogger.log('handleAdminPost_.setWebappUrl', { webappUrl: stampUrl });
    return jsonOutput_({ ok: true, webappUrl: stampUrl });
  }

  var storedSecret = PropertiesService.getScriptProperties().getProperty('ADMIN_SHARED_SECRET');
  if (!storedSecret || payload.adminSecret !== storedSecret) {
    GasLogger.log('handleAdminPost_.forbidden', { action: payload.action });
    return jsonOutput_({ ok: false, error: 'forbidden' });
  }

  try {
    switch (payload.action) {
      case 'setScriptProperties': {
        var keys = Object.keys(payload.properties || {});
        PropertiesService.getScriptProperties().setProperties(payload.properties || {});
        GasLogger.log('handleAdminPost_.setScriptProperties', { keys: keys });
        return jsonOutput_({ ok: true, keysSet: keys });
      }

      case 'getDiagnostics': {
        // Read-only visibility into logging/config wiring — never returns secret values,
        // only presence/length, so it's safe to leave gated behind the admin secret.
        var diagProps = PropertiesService.getScriptProperties();
        var axiomToken = diagProps.getProperty('AXIOM_TOKEN');
        var axiomDataset = diagProps.getProperty('AXIOM_DATASET');
        var diagResult = {
          ok: true,
          deployTarget: (typeof APP_DEPLOY_TARGET !== 'undefined' ? APP_DEPLOY_TARGET : 'unknown'),
          appVersion: (typeof APP_VERSION !== 'undefined' ? APP_VERSION : 'unknown'),
          webappUrl: diagProps.getProperty('WEBAPP_URL') || null,
          axiomTokenSet: !!axiomToken,
          axiomTokenLength: axiomToken ? axiomToken.length : 0,
          axiomDataset: axiomDataset || null,
          gasLoggerFolderIdSet: !!diagProps.getProperty('GAS_LOGGER_PARENT_FOLDER_ID')
        };
        // Optional: fire a real test ingest and report the raw HTTP result (or thrown
        // exception message) so a broken Axiom pipe is diagnosable without Stackdriver access.
        if (payload.testAxiomIngest && axiomToken && axiomDataset) {
          try {
            var testResp = UrlFetchApp.fetch(
              'https://api.axiom.co/v1/datasets/' + axiomDataset + '/ingest',
              {
                method: 'post',
                contentType: 'application/json',
                headers: { Authorization: 'Bearer ' + axiomToken },
                payload: JSON.stringify([{ _time: new Date().toISOString(), name: 'getDiagnostics.testAxiomIngest', side: 'gas' }]),
                muteHttpExceptions: true
              }
            );
            diagResult.axiomTestIngest = { httpCode: testResp.getResponseCode(), body: testResp.getContentText() };
          } catch (axiomErr) {
            diagResult.axiomTestIngest = { threw: String(axiomErr && axiomErr.message || axiomErr) };
          }
        }
        return jsonOutput_(diagResult);
      }

      case 'listSheets': {
        var allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
        return jsonOutput_({
          ok: true,
          sheets: allSheets.map(function (s) {
            return { name: s.getName(), hidden: s.isSheetHidden(), index: s.getIndex() };
          })
        });
      }

      case 'getSheet': {
        // Downloads a sheet's full data range as CSV — used by integration tests to
        // verify state written by the app (e.g. that createBallot actually produced a sheet).
        if (!payload.sheetName) {
          return jsonOutput_({ ok: false, error: 'sheetName is required' });
        }
        var getSheetSs = SpreadsheetApp.getActiveSpreadsheet();
        var getSheetObj = getSheetSs.getSheetByName(payload.sheetName);
        if (!getSheetObj) {
          return jsonOutput_({ ok: false, error: 'sheet_not_found' });
        }
        var rows = getSheetObj.getLastRow() > 0
          ? getSheetObj.getDataRange().getValues()
          : [];
        var csv = rows.map(function (row) {
          return row.map(function (cell) {
            var s = String(cell == null ? '' : cell);
            return /[\t"\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
          }).join('\t');
        }).join('\n');
        return jsonOutput_({ ok: true, csv: csv, rowCount: rows.length });
      }

      case 'setSheet': {
        // Writes a 2D array of values starting at (row, col) (default 1,1) into a sheet —
        // creating the sheet if it doesn't exist. Used by integration tests to seed/overwrite
        // state (e.g. a ballot's Responses rows) without going through the UI.
        if (!payload.sheetName || !payload.rows) {
          return jsonOutput_({ ok: false, error: 'sheetName and rows are required' });
        }
        var setSheetSs = SpreadsheetApp.getActiveSpreadsheet();
        var setSheetObj = setSheetSs.getSheetByName(payload.sheetName) || setSheetSs.insertSheet(payload.sheetName);
        var startRow = payload.row || 1;
        var startCol = payload.col || 1;
        var setRows = payload.rows;
        if (setRows.length > 0) {
          var maxCols = setRows.reduce(function (m, r) { return Math.max(m, r.length); }, 1);
          var padded = setRows.map(function (r) {
            var copy = r.slice();
            while (copy.length < maxCols) copy.push('');
            return copy;
          });
          setSheetObj.getRange(startRow, startCol, padded.length, maxCols).setValues(padded);
        }
        GasLogger.log('handleAdminPost_.setSheet', { sheetName: payload.sheetName, rowCount: setRows.length });
        return jsonOutput_({ ok: true, rowCount: setRows.length });
      }

      case 'deleteSheet': {
        // Cleanup for integration tests (e.g. remove a ballot created by createBallot).
        if (!payload.sheetName) {
          return jsonOutput_({ ok: false, error: 'sheetName is required' });
        }
        var deleteSs = SpreadsheetApp.getActiveSpreadsheet();
        var deleteSheetObj = deleteSs.getSheetByName(payload.sheetName);
        if (!deleteSheetObj) {
          return jsonOutput_({ ok: false, error: 'sheet_not_found' });
        }
        deleteSs.deleteSheet(deleteSheetObj);
        GasLogger.log('handleAdminPost_.deleteSheet', { sheetName: payload.sheetName });
        return jsonOutput_({ ok: true });
      }

      case 'createBallot': {
        // Headless equivalent of the admin page's "Create New Ballot" form — lets
        // integration tests exercise/verify ballot creation without a browser.
        if (!payload.id) {
          return jsonOutput_({ ok: false, error: 'id is required' });
        }
        var createSs = SpreadsheetApp.getActiveSpreadsheet();
        var createdSheet = createNewBallot_(createSs, payload.id);
        GasLogger.log('handleAdminPost_.createBallot', { id: payload.id });
        return jsonOutput_({ ok: true, sheetName: createdSheet.getName() });
      }

      default:
        return jsonOutput_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    GasLogger.logError('handleAdminPost_.error', err, { action: payload.action });
    return jsonOutput_({ ok: false, error: 'server_error', detail: err.message });
  }
}
