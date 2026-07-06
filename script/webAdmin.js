/**
 * webAdmin.js
 *
 * doGet ?cmd=admin — lists every Survey-<name> sheet in the spreadsheet and
 * lets you run RCV + Condorcet analysis against any of them, viewing the
 * results in the browser. Each run also writes a summary back into that
 * survey's own Results section (see SurveyModel.js for the sheet layout).
 *
 * Reachable both from the web app URL directly and from the spreadsheet's
 * "Voting and Ballot Tools > Open Survey Admin Page" menu (onOpen.js).
 *
 * Entry point: _handleAdmin(e), wired from WebApp.js doGet (cmd === 'admin').
 */

/**
 * ?cmd=admin                          — list all surveys, with a create-new-survey form
 * ?cmd=admin&action=create&id=<id>    — create a new Survey-<id> sheet, then open its edit form
 * ?cmd=admin&action=edit&id=<id>      — edit a survey's Title/Description/Footer/Contact/Accept-New/Items
 * ?cmd=admin&action=save&id=<id>&...  — save the edit form
 * ?cmd=admin&action=saveItems&id=...  — rename/re-describe items and/or add a new one
 * ?cmd=admin&action=analyze&id=<id>   — run analysis on one survey and show results
 *
 * @param {Object} e doGet event.
 * @return {HtmlOutput}
 */
function _handleAdmin(e) {
  var params = (e && e.parameter) || {};
  var action = params.action || 'list';

  GasLogger.log('webapp.admin', { action: action, id: params.id || '' });
  GasLogger.flush();

  var body;
  switch (action) {
    case 'analyze':
      body = params.id ? _renderAdminAnalysis_(params.id) : _renderAdminList_();
      break;
    case 'create':
      body = _handleAdminCreate_(params);
      break;
    case 'edit':
      body = _renderAdminEdit_(params.id, null);
      break;
    case 'save':
      body = _handleAdminSave_(params);
      break;
    case 'saveItems':
      body = _handleAdminSaveItems_(params);
      break;
    default:
      body = _renderAdminList_();
  }

  return HtmlService.createHtmlOutput(_ADMIN_STYLE_ + body + _versionFooterHtml_())
    .setTitle('Survey Admin')
    // GAS wraps webapp output in an outer frame that only reliably honors a mobile
    // viewport when set via addMetaTag — a plain <meta viewport> tag inside the page's
    // own HTML is not enough (same reason webSurvey.js's _handleSurvey sets it this way).
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

// Mobile-first: no table layout (tables force horizontal scrolling on narrow screens).
// Each survey is a "card" — a flex column that stacks its own action buttons and, on a
// narrow viewport, stacks the buttons themselves too (flex-wrap). Buttons/inputs use a
// minimum touch-target height (44px, per WCAG 2.5.5) since this page has no hover state
// to fall back on for a touch device.
var _ADMIN_STYLE_ =
  '<style>' +
  '*{box-sizing:border-box;}' +
  'body{font-family:Arial,sans-serif;max-width:640px;margin:24px auto;padding:0 16px;color:#202124;}' +
  'h1{font-size:1.4rem;}' +
  'a{color:#1a73e8;}' +
  'form{margin-bottom:24px;}' +
  'label{display:block;font-weight:bold;margin:10px 0 4px;}' +
  'input[type="text"],textarea{width:100%;font-size:1rem;padding:8px;border:1px solid #dadce0;border-radius:4px;font-family:inherit;}' +
  'textarea{resize:vertical;}' +
  'button,a.btn{font-size:1rem;padding:10px 16px;min-height:44px;border:none;border-radius:4px;' +
    'background:#1a73e8;color:#fff;cursor:pointer;text-decoration:none;text-align:center;' +
    'display:inline-flex;align-items:center;justify-content:center;white-space:nowrap;}' +
  'button{width:100%;margin-top:12px;}' +
  '.survey-list{display:flex;flex-direction:column;gap:12px;margin-bottom:16px;}' +
  '.survey-card{border:1px solid #dadce0;border-radius:8px;padding:12px 16px;}' +
  '.survey-card h3{margin:0 0 2px;font-size:1.05rem;}' +
  '.survey-card .survey-id{color:#5f6368;font-size:0.85em;margin-bottom:6px;}' +
  '.survey-card .survey-meta{color:#5f6368;font-size:0.9em;margin-bottom:10px;}' +
  '.survey-actions{display:flex;flex-wrap:wrap;gap:8px;}' +
  '.survey-actions a.btn{flex:1 1 auto;min-width:100px;}' +
  '.survey-info{margin-top:10px;}' +
  '.survey-info summary{cursor:pointer;color:#1a73e8;font-size:0.9em;padding:6px 0;min-height:44px;display:flex;align-items:center;}' +
  '.survey-info p{margin:4px 0 0;color:#3c4043;font-size:0.9em;white-space:pre-wrap;}' +
  '.item-row{border:1px solid #dadce0;border-radius:6px;padding:10px 12px;margin-bottom:10px;}' +
  '.item-row label{margin-top:6px;}' +
  '.item-row label:first-child{margin-top:0;}' +
  '.app-version-footer{color:#9aa0a6;font-size:0.75em;margin-top:24px;}' +
  '@media (max-width:420px){button,a.btn{width:100%;}}' +
  '</style>';

/**
 * Small "vX.X.X (TARGET)" footer appended to every admin page render, so anyone
 * looking at the page can confirm which deployed build they're actually viewing
 * (see version.js — stamped by tools/manage-deployments.js on every deploy).
 * @return {string}
 */
function _versionFooterHtml_() {
  var version = (typeof APP_VERSION !== 'undefined' && APP_VERSION) || 'unknown';
  var target = (typeof APP_DEPLOY_TARGET !== 'undefined' && APP_DEPLOY_TARGET) || 'unknown';
  var date = (typeof APP_VERSION_DATE !== 'undefined' && APP_VERSION_DATE) || '';
  return '<footer class="app-version-footer">v' + _escapeHtml_(version) + ' (' + _escapeHtml_(target) + ')' +
    (date ? ' — deployed ' + _escapeHtml_(date) : '') + '</footer>';
}

/**
 * @param {string=} createError shown above the create form if the last
 *   creation attempt failed (e.g. duplicate/invalid id).
 * @param {string=} prefillId re-populates the id field after a failed create.
 * @return {string} HTML listing every Survey-<name> sheet.
 */
function _renderAdminList_(createError, prefillId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ids = listSurveyIds_(ss);
  var baseUrl = _getWebAppUrl_();

  var html = '<h1>Survey Admin</h1>';
  html += _renderCreateForm_(createError, prefillId);

  if (!ids.length) {
    return html + '<p>No survey sheets found yet — create one above.</p>';
  }

  html += '<div class="survey-list">';
  ids.forEach(function (id) {
    var sheet = findSurveySheet_(ss, id);
    var config = readSurveyConfig_(sheet);
    var responseCount = readSurveyResponseRows_(sheet).length;
    var surveyUrl = baseUrl + '?cmd=survey&id=' + encodeURIComponent(id);
    var info = String(config.Info || '').trim();

    html += '<div class="survey-card">' +
      '<h3>' + _escapeHtml_(config.Title || id) + '</h3>' +
      '<div class="survey-id">' + _escapeHtml_(id) + '</div>' +
      '<div class="survey-meta">' + responseCount + ' response' + (responseCount === 1 ? '' : 's') + '</div>' +
      '<div class="survey-actions">' +
      '<a class="btn" target="_blank" href="' + _escapeHtml_(surveyUrl) + '">View Survey</a>' +
      '<a class="btn" target="_top" href="' + _escapeHtml_(baseUrl) + '?cmd=admin&action=edit&id=' + encodeURIComponent(id) + '">Edit</a>' +
      '<a class="btn" target="_top" href="' + _escapeHtml_(baseUrl) + '?cmd=admin&action=analyze&id=' + encodeURIComponent(id) + '">Run Analysis</a>' +
      '</div>';

    // A <details>/<summary> disclosure is used instead of a title="" tooltip because
    // tooltips have no reliable trigger on touch devices — <details> opens on tap or
    // click alike, needs no JS, and is screen-reader friendly out of the box. Omitted
    // entirely when there's no Info text, so the list stays uncluttered by default.
    if (info) {
      html += '<details class="survey-info"><summary>ℹ️ More info</summary><p>' + _escapeHtml_(info) + '</p></details>';
    }

    html += '</div>';
  });
  html += '</div>';
  return html;
}

/**
 * @param {string=} error
 * @param {string=} prefillId
 * @return {string}
 */
function _renderCreateForm_(error, prefillId) {
  var html = '<h2>Create New Survey</h2>';
  if (error) html += '<p style="color:#d93025;">' + _escapeHtml_(error) + '</p>';
  // Forms/links that navigate to a NEW doGet-rendered page must use both target="_top"
  // (to escape HtmlService's sandboxed iframe) AND an ABSOLUTE action/href built from the
  // real deployed webapp URL — a relative URL (e.g. just "?cmd=admin...") resolves against
  // the sandbox iframe's own script.googleusercontent.com origin, not the real
  // script.google.com/macros/.../exec URL, so target="_top" alone still lands on a bare,
  // never-meant-to-be-viewed-standalone sandbox URL and renders blank.
  html += '<form method="get" target="_top" action="' + _escapeHtml_(_getWebAppUrl_()) + '">' +
    '<input type="hidden" name="cmd" value="admin">' +
    '<input type="hidden" name="action" value="create">' +
    '<label for="newSurveyId">Survey ID (used in the survey link and sheet name — letters, numbers, "-", "_" only)</label>' +
    '<input type="text" id="newSurveyId" name="id" value="' + _escapeHtml_(prefillId || '') + '" required>' +
    '<button type="submit">Create Survey</button>' +
    '</form>';
  return html;
}

/**
 * @param {Object} params
 * @return {string} HTML for the resulting page (edit form on success, list w/ error on failure).
 */
function _handleAdminCreate_(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    createNewSurvey_(ss, params.id);
  } catch (err) {
    return _renderAdminList_(err.message, params.id);
  }
  return _renderAdminEdit_(params.id, 'Survey created — fill in the details below and Save. Add items using the Items section below, or let respondents add them if Accept-New is checked.');
}

/**
 * @param {string} id
 * @param {?string} message success/info banner text
 * @return {string}
 */
function _renderAdminEdit_(id, message, errorMessage) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSurveySheet_(ss, id);
  var baseUrl = _getWebAppUrl_();
  var back = '<p><a target="_top" href="' + _escapeHtml_(baseUrl) + '?cmd=admin">&larr; Back to survey list</a></p>';
  if (!sheet) return back + '<p>No survey found for id "' + _escapeHtml_(id) + '".</p>';

  // Self-heals older sheets: adds the section-marker highlighting and the Items
  // section (if missing) the first time an admin opens this survey to edit it.
  _highlightSectionMarkers_(sheet);

  var config = readSurveyConfig_(sheet);
  var acceptNew = _surveyAcceptsNew_(config);
  var surveyUrl = baseUrl + '?cmd=survey&id=' + encodeURIComponent(id);

  var html = back + '<h1>Edit Survey: ' + _escapeHtml_(id) + '</h1>';
  if (message) html += '<p style="color:#188038;">' + _escapeHtml_(message) + '</p>';
  if (errorMessage) html += '<p style="color:#d93025;">' + _escapeHtml_(errorMessage) + '</p>';
  html += '<p>Survey link: <a href="' + _escapeHtml_(surveyUrl) + '" target="_blank">' + _escapeHtml_(surveyUrl) + '</a></p>';

  html += '<form method="get" target="_top" action="' + _escapeHtml_(baseUrl) + '">' +
    '<input type="hidden" name="cmd" value="admin">' +
    '<input type="hidden" name="action" value="save">' +
    '<input type="hidden" name="id" value="' + _escapeHtml_(id) + '">' +
    '<label for="cfgTitle">Title</label>' +
    '<input type="text" id="cfgTitle" name="Title" value="' + _escapeHtml_(config.Title || '') + '">' +
    '<label for="cfgDescription">Description</label>' +
    '<textarea id="cfgDescription" name="Description" rows="3">' + _escapeHtml_(config.Description || '') + '</textarea>' +
    '<label for="cfgFooter">Footer</label>' +
    '<input type="text" id="cfgFooter" name="Footer" value="' + _escapeHtml_(config.Footer || '') + '">' +
    '<label for="cfgContact">Contact</label>' +
    '<input type="text" id="cfgContact" name="Contact" value="' + _escapeHtml_(config.Contact || '') + '">' +
    '<label><input type="checkbox" name="AcceptNew" value="TRUE"' + (acceptNew ? ' checked' : '') + '> Accept new candidates from respondents</label>' +
    '<label for="cfgInfo">Info (admin-only notes — shown as a "More info" disclosure in the survey list, never shown to respondents)</label>' +
    '<textarea id="cfgInfo" name="Info" rows="3">' + _escapeHtml_(config.Info || '') + '</textarea>' +
    '<button type="submit">Save</button>' +
    '</form>';

  html += _renderItemsForm_(id, sheet, baseUrl);

  return html;
}

/**
 * Renders the Items management form: one Name/Details field pair per existing
 * candidate (rename + describe in place, position-aligned — see SurveyModel.js's
 * saveSurveyItems_) plus a section to add one new item with its own details. Both
 * submit together as a single form so a multi-item edit round-trips as one save.
 *
 * @param {string} id
 * @param {Sheet} sheet
 * @param {string} baseUrl
 * @return {string}
 */
function _renderItemsForm_(id, sheet, baseUrl) {
  var candidates = readSurveyCandidates_(sheet);
  var items = readSurveyItemDetails_(sheet);
  // Backfill any candidate that predates the Items table (or predates having its own
  // row yet) so every candidate always has a field pair to edit.
  while (items.length < candidates.length) {
    items.push({ name: candidates[items.length], details: '' });
  }

  var html = '<h2>Items</h2>';
  html += '<form method="get" target="_top" action="' + _escapeHtml_(baseUrl) + '">' +
    '<input type="hidden" name="cmd" value="admin">' +
    '<input type="hidden" name="action" value="saveItems">' +
    '<input type="hidden" name="id" value="' + _escapeHtml_(id) + '">' +
    '<input type="hidden" name="itemCount" value="' + candidates.length + '">';

  if (!candidates.length) {
    html += '<p><em>No items yet — add the first one below.</em></p>';
  }

  candidates.forEach(function (name, i) {
    html += '<div class="item-row">' +
      '<label for="itemName' + i + '">Name</label>' +
      '<input type="text" id="itemName' + i + '" name="item_name_' + i + '" value="' + _escapeHtml_(items[i].name || name) + '">' +
      '<label for="itemDetails' + i + '">Details</label>' +
      '<textarea id="itemDetails' + i + '" name="item_details_' + i + '" rows="2">' + _escapeHtml_(items[i].details || '') + '</textarea>' +
      '</div>';
  });

  html += '<h3>Add a new item</h3>' +
    '<div class="item-row">' +
    '<label for="newItemName">Name</label>' +
    '<input type="text" id="newItemName" name="newItemName">' +
    '<label for="newItemDetails">Details</label>' +
    '<textarea id="newItemDetails" name="newItemDetails" rows="2"></textarea>' +
    '</div>' +
    '<button type="submit">Save Items</button>' +
    '</form>';

  return html;
}

/**
 * @param {Object} params
 * @return {string}
 */
function _handleAdminSaveItems_(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSurveySheet_(ss, params.id);
  if (!sheet) return _renderAdminList_('No survey found for id "' + params.id + '".');

  try {
    var count = parseInt(params.itemCount, 10) || 0;
    var items = [];
    for (var i = 0; i < count; i++) {
      items.push({ name: params['item_name_' + i] || '', details: params['item_details_' + i] || '' });
    }
    if (items.length) {
      saveSurveyItemsForId_(params.id, items);
    }

    var newName = String(params.newItemName || '').trim();
    if (newName) {
      addSurveyItemForAdmin_(params.id, newName, params.newItemDetails || '');
    }
  } catch (err) {
    return _renderAdminEdit_(params.id, null, err.message);
  }

  return _renderAdminEdit_(params.id, 'Items saved.');
}

/**
 * @param {Object} params
 * @return {string}
 */
function _handleAdminSave_(params) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSurveySheet_(ss, params.id);
  if (!sheet) return _renderAdminList_('No survey found for id "' + params.id + '".');

  writeSurveyConfig_(sheet, {
    Title: params.Title || '',
    Description: params.Description || '',
    Footer: params.Footer || '',
    Contact: params.Contact || '',
    'Accept-New': params.AcceptNew === 'TRUE' ? 'TRUE' : 'FALSE',
    Info: params.Info || ''
  });

  return _renderAdminEdit_(params.id, 'Saved.');
}

/**
 * @param {string} id survey id
 * @return {string} HTML with RCV + Condorcet results for one survey.
 */
function _renderAdminAnalysis_(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSurveySheet_(ss, id);
  var back = '<p><a target="_top" href="' + _escapeHtml_(_getWebAppUrl_()) + '?cmd=admin">&larr; Back to survey list</a></p>';

  if (!sheet) {
    return back + '<p>No survey found for id "' + _escapeHtml_(id) + '".</p>';
  }

  var results = runSurveyAnalysis_(sheet);
  var html = back + '<h1>' + _escapeHtml_(id) + ' — Analysis Results</h1>';

  if (results.error) {
    return html + '<p><em>' + _escapeHtml_(results.error) + '</em></p>';
  }

  html += '<h2>Ranked Choice Voting</h2>';
  if (results.rcv.winner) {
    html += '<p><b>Winner: ' + _escapeHtml_(results.rcv.winner) + '</b></p>';
  } else if (results.rcv.tie) {
    html += '<p><b>Tie after all eliminations:</b> ' + results.rcv.tie.map(_escapeHtml_).join(', ') + '</p>';
  }
  html += formatCandidateSummaryHtml(results.rcv.summary);

  html += generateCondorcetResultsHtml(results.condorcet);

  html += '<p><em>Results have been written to the Results section of the ' +
    _escapeHtml_(getSurveySheetName_(id)) + ' sheet.</em></p>';

  return html;
}

/**
 * Runs RCV + all four Condorcet methods against a survey's current
 * candidates/responses, writes a summary into the sheet's Results section,
 * and returns the raw results for HTML rendering.
 *
 * @param {Sheet} sheet a Survey-<name> sheet
 * @return {{rcv:Object, condorcet:Object}|{error:string}}
 */
function runSurveyAnalysis_(sheet) {
  var candidateNames = readSurveyCandidates_(sheet);
  var responseRows = readSurveyResponseRows_(sheet);

  if (candidateNames.length < 2) {
    return { error: 'Survey needs at least two candidates before it can be analyzed.' };
  }
  if (responseRows.length === 0) {
    return { error: 'Survey has no responses yet.' };
  }

  var ballots = responseRows.map(function (r) {
    return { voterName: r.name, ranks: r.ranks, weight: r.weight };
  });

  var rcv = runRankedChoiceVoting(candidateNames, ballots, null);

  var candidates = candidateNames.map(function (name) { return { name: name }; });
  var condorcet = {
    condorcet: findCondorcetWinner(ballots, candidateNames),
    schulze: findSchulzeWinner(ballots, candidateNames),
    rankedPairs: findRankedPairsWinner(ballots, candidateNames),
    minimax: findMinimaxWinner(ballots, candidateNames)
  };

  writeSurveyResults_(sheet, _buildResultsRows_(rcv, condorcet));

  return { rcv: rcv, condorcet: condorcet };
}

/**
 * @param {Object} rcv result of runRankedChoiceVoting
 * @param {Object} condorcet {condorcet, schulze, rankedPairs, minimax}
 * @return {Array<Array>} rows to write into the Results section.
 */
function _buildResultsRows_(rcv, condorcet) {
  var rows = [];
  rows.push(['Analysis run: ' + new Date().toISOString()]);
  rows.push(['']);
  rows.push(['Ranked Choice Voting (RCV)']);
  rows.push([rcv.winner ? 'Winner: ' + rcv.winner : 'Tie: ' + (rcv.tie || []).join(', ')]);
  rcv.summary.forEach(function (row) { rows.push(row); });
  rows.push(['']);
  rows.push(['Condorcet Analysis', 'Winner']);
  rows.push(['Basic Condorcet', condorcet.condorcet.winner || 'None (cycle)']);
  rows.push(['Schulze', condorcet.schulze.winner || 'None (cycle)']);
  rows.push(['Ranked Pairs', condorcet.rankedPairs.winner || 'None (cycle)']);
  rows.push(['Minimax', condorcet.minimax.winner || 'None (tie/cycle)']);

  ['condorcet', 'schulze', 'rankedPairs', 'minimax'].forEach(function (key) {
    rows.push(['']);
    rows.push([key + ' — ranked candidates']);
    rows.push(['Rank', 'Candidate', 'Score']);
    condorcet[key].rankedCandidates.forEach(function (rc, i) {
      rows.push([i + 1, rc.candidate, rc.score]);
    });
  });

  return rows;
}

/**
 * @param {*} value
 * @return {string}
 */
function _escapeHtml_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    _buildResultsRows_: _buildResultsRows_,
    _escapeHtml_: _escapeHtml_,
    _versionFooterHtml_: _versionFooterHtml_
  };
}
