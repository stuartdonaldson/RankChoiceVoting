/**
 * webAdmin.js
 *
 * doGet ?cmd=admin — lists every Ballot-<name> sheet in the spreadsheet and
 * lets you run RCV + Condorcet analysis against any of them, viewing the
 * results in the browser. Each run also writes a summary back into that
 * ballot's own Results section (see BallotModel.js for the sheet layout).
 *
 * Reachable both from the web app URL directly and from the spreadsheet's
 * "Voting and Ballot Tools > Open Ballot Admin Page" menu (onOpen.js).
 *
 * Entry point: _handleAdmin(e), wired from WebApp.js doGet (cmd === 'admin').
 */

/**
 * ?cmd=admin                          — list all ballots, with a create-new-ballot form
 * ?cmd=admin&action=create&id=<id>    — create a new Ballot-<id> sheet, then open its edit page
 * ?cmd=admin&action=edit&id=<id>      — edit a ballot: a live ballot preview (webAdminEditPage.html)
 *                                        with a pencil icon per editable field. Each field saves
 *                                        itself immediately via google.script.run (see the
 *                                        adminSave... / adminAddCandidate RPCs below) — there
 *                                        is no page-wide Save, only Back.
 * ?cmd=admin&action=analyze&id=<id>   — run analysis on one ballot and show results
 *
 * @param {Object} e doGet event.
 * @return {HtmlOutput}
 */
function _handleAdmin(e) {
  var params = (e && e.parameter) || {};
  var action = params.action || 'list';

  GasLogger.log('webapp.admin', { action: action, id: params.id || '' });
  GasLogger.flush();

  if (action === 'create') {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    try {
      createNewBallot_(ss, params.id);
    } catch (err) {
      return HtmlService.createHtmlOutput(_ADMIN_STYLE_ + _renderAdminList_(err.message, params.id) + _versionFooterHtml_())
        .setTitle('Ballot Admin')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
    }
    return _renderAdminEditPage_();
  }

  if (action === 'edit') {
    return _renderAdminEditPage_();
  }

  var body;
  switch (action) {
    case 'analyze':
      body = params.id ? _renderAdminAnalysis_(params.id) : _renderAdminList_();
      break;
    default:
      body = _renderAdminList_();
  }

  return HtmlService.createHtmlOutput(_ADMIN_STYLE_ + body + _versionFooterHtml_())
    .setTitle('Ballot Admin')
    // GAS wraps webapp output in an outer frame that only reliably honors a mobile
    // viewport when set via addMetaTag — a plain <meta viewport> tag inside the page's
    // own HTML is not enough (same reason webBallot.js's _handleBallot sets it this way).
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

/**
 * @return {HtmlOutput} the webAdminEditPage.html template — it reads id (and, right after
 *   a create, action=create for a one-time banner) from google.script.url.getLocation()
 *   client-side, same pattern as webBallotPage.html, then fetches data via getAdminEditData.
 */
function _renderAdminEditPage_() {
  return HtmlService.createHtmlOutputFromFile('webAdminEditPage')
    .setTitle('Edit Ballot')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

// Mobile-first: no table layout (tables force horizontal scrolling on narrow screens).
// Each ballot is a "card" — a flex column that stacks its own action buttons and, on a
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
  '.ballot-list{display:flex;flex-direction:column;gap:12px;margin-bottom:16px;}' +
  '.ballot-card{border:1px solid #dadce0;border-radius:8px;padding:12px 16px;}' +
  '.ballot-card h3{margin:0 0 2px;font-size:1.05rem;}' +
  '.ballot-card .ballot-id{color:#5f6368;font-size:0.85em;margin-bottom:6px;}' +
  '.ballot-card .ballot-meta{color:#5f6368;font-size:0.9em;margin-bottom:10px;}' +
  '.ballot-actions{display:flex;flex-wrap:wrap;gap:8px;}' +
  '.ballot-actions a.btn{flex:1 1 auto;min-width:100px;}' +
  '.ballot-info{margin-top:10px;}' +
  '.ballot-info summary{cursor:pointer;color:#1a73e8;font-size:0.9em;padding:6px 0;min-height:44px;display:flex;align-items:center;}' +
  '.ballot-info p{margin:4px 0 0;color:#3c4043;font-size:0.9em;white-space:pre-wrap;}' +
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
 * @return {string} HTML listing every Ballot-<name> sheet.
 */
function _renderAdminList_(createError, prefillId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ids = listBallotIds_(ss);
  var baseUrl = _getWebAppUrl_();

  var html = '<h1>Ballot Admin</h1>';
  html += _renderCreateForm_(createError, prefillId);

  if (!ids.length) {
    return html + '<p>No ballot sheets found yet — create one above.</p>';
  }

  html += '<div class="ballot-list">';
  ids.forEach(function (id) {
    var sheet = findBallotSheet_(ss, id);
    var config = readBallotConfig_(sheet);
    var candidateCount = readBallotCandidates_(sheet).length;
    var responseCount = readBallotResponseRows_(sheet).length;
    var uniqueRespondentCount = countUniqueBallotRespondents_(sheet);
    var staleRespondents = findRespondentsWithNewCandidates_(sheet);
    var ballotUrl = baseUrl + '?cmd=ballot&id=' + encodeURIComponent(id);
    var adminNotes = String(config['Admin-Only-Notes'] || '').trim();

    html += '<div class="ballot-card">' +
      '<h3>' + _escapeHtml_(config.Title || id) + '</h3>' +
      '<div class="ballot-id">' + _escapeHtml_(id) + '</div>' +
      '<div class="ballot-meta">' + candidateCount + ' candidate' + (candidateCount === 1 ? '' : 's') +
      ' &middot; ' + responseCount + ' response' + (responseCount === 1 ? '' : 's') +
      ' (' + uniqueRespondentCount + ' unique respondent' + (uniqueRespondentCount === 1 ? '' : 's') + ')' + '</div>' +
      '<div class="ballot-actions">' +
      '<a class="btn" target="_blank" href="' + _escapeHtml_(ballotUrl) + '">View Ballot</a>' +
      '<a class="btn" target="_top" href="' + _escapeHtml_(baseUrl) + '?cmd=admin&action=edit&id=' + encodeURIComponent(id) + '">Edit</a>' +
      '<a class="btn" target="_top" href="' + _escapeHtml_(baseUrl) + '?cmd=admin&action=analyze&id=' + encodeURIComponent(id) + '">Run Analysis</a>' +
      '</div>';

    // A <details>/<summary> disclosure is used instead of a title="" tooltip because
    // tooltips have no reliable trigger on touch devices — <details> opens on tap or
    // click alike, needs no JS, and is screen-reader friendly out of the box. Omitted
    // entirely when there's no Admin-Only-Notes text, so the list stays uncluttered
    // by default.
    if (adminNotes) {
      html += '<details class="ballot-info"><summary>ℹ️ More info</summary><p>' + _escapeHtml_(adminNotes) + '</p></details>';
    }

    // Flags respondents whose last submission predates a since-added candidate — see
    // findRespondentsWithNewCandidates_ for how "stale" is detected. Omitted when
    // empty for the same reason as the Admin-Only-Notes disclosure above.
    if (staleRespondents.length) {
      html += '<details class="ballot-info"><summary>⚠️ ' + staleRespondents.length +
        ' respondent' + (staleRespondents.length === 1 ? '' : 's') +
        ' may want to review (candidates added since they responded)</summary><p>' +
        _escapeHtml_(staleRespondents.join(', ')) + '</p></details>';
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
  var html = '<h2>Create New Ballot</h2>';
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
    '<label for="newBallotId">Ballot ID (used in the ballot link and sheet name — letters, numbers, "-", "_" only)</label>' +
    '<input type="text" id="newBallotId" name="id" value="' + _escapeHtml_(prefillId || '') + '" required>' +
    '<button type="submit">Create Ballot</button>' +
    '</form>';
  return html;
}

/**
 * RPC: called once by webAdminEditPage.html on load (after reading the ballot id from
 * google.script.url.getLocation) to fetch everything the edit-mode ballot preview needs
 * to render. Field names match what the client displays/edits, not the sheet's config
 * key spelling — see BallotModel.js for the underlying "Title"/"Accept-New"/etc. keys.
 *
 * @param {string} id
 * @return {Object} or {error:'no_ballot'} if the id doesn't match a ballot sheet.
 */
function getAdminEditData(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findBallotSheet_(ss, id);
  if (!sheet) return { error: 'no_ballot' };

  // Self-heals older sheets: adds the section-marker highlighting and the Candidates
  // section (if missing) the first time an admin opens this ballot to edit it.
  _highlightSectionMarkers_(sheet);

  var config = readBallotConfig_(sheet);
  var candidates = readBallotCandidates_(sheet);
  var candidateRows = readBallotCandidateDetails_(sheet);
  // Backfill any candidate that predates the Candidates table (or predates having its
  // own row yet) so every candidate always has a name/details pair to edit.
  while (candidateRows.length < candidates.length) {
    candidateRows.push({ name: candidates[candidateRows.length], details: '' });
  }

  var baseUrl = _getWebAppUrl_();
  return {
    id: id,
    title: config.Title || '',
    description: config.Description || '',
    instructions: config.Instructions || '',
    footer: config.Footer || '',
    contact: config.Contact || '',
    acceptNew: _ballotAcceptsNew_(config),
    addInstructions: config['Add-Instructions'] || '',
    adminOnlyNotes: config['Admin-Only-Notes'] || '',
    candidates: candidateRows,
    ballotUrl: baseUrl + '?cmd=ballot&id=' + encodeURIComponent(id),
    adminListUrl: baseUrl + '?cmd=admin',
    appVersion: (typeof APP_VERSION !== 'undefined' && APP_VERSION) || '',
    appDeployTarget: (typeof APP_DEPLOY_TARGET !== 'undefined' && APP_DEPLOY_TARGET) || ''
  };
}

/**
 * @param {string} id
 * @param {Object} updates passed straight through to writeBallotConfig_
 * @return {{ok:boolean}}
 */
function _adminWriteConfig_(id, updates) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findBallotSheet_(ss, id);
  if (!sheet) throw new Error('No ballot found for id "' + id + '".');
  writeBallotConfig_(sheet, updates);
  return { ok: true };
}

/** RPC: saves the Title field's pencil-icon editor. @return {{ok:boolean}} */
function adminSaveTitle(id, value) { return _adminWriteConfig_(id, { Title: String(value || '') }); }

/** RPC: saves the landing-page Description field's pencil-icon editor. @return {{ok:boolean}} */
function adminSaveDescription(id, value) { return _adminWriteConfig_(id, { Description: String(value || '') }); }

/** RPC: saves the ballot-page Instructions field's pencil-icon editor. @return {{ok:boolean}} */
function adminSaveInstructions(id, value) { return _adminWriteConfig_(id, { Instructions: String(value || '') }); }

/** RPC: saves the Footer field's pencil-icon editor. @return {{ok:boolean}} */
function adminSaveFooter(id, value) { return _adminWriteConfig_(id, { Footer: String(value || '') }); }

/** RPC: saves the Contact field's pencil-icon editor. @return {{ok:boolean}} */
function adminSaveContact(id, value) { return _adminWriteConfig_(id, { Contact: String(value || '') }); }

/** RPC: saves the Admin-Only Notes field's pencil-icon editor. @return {{ok:boolean}} */
function adminSaveAdminOnlyNotes(id, value) { return _adminWriteConfig_(id, { 'Admin-Only-Notes': String(value || '') }); }

/**
 * RPC: saves the combined "Adding New Candidates" editor (Accept-New checkbox +
 * Add-Instructions text) in one call, since they're edited together as one field group.
 *
 * @param {string} id
 * @param {boolean} acceptNew
 * @param {string} addInstructions
 * @return {{ok:boolean}}
 */
function adminSaveAddSettings(id, acceptNew, addInstructions) {
  return _adminWriteConfig_(id, {
    'Accept-New': acceptNew ? 'TRUE' : 'FALSE',
    'Add-Instructions': String(addInstructions || '')
  });
}

/**
 * RPC: saves one existing candidate's name/details from its pencil-icon editor.
 * saveBallotCandidatesForId_ requires the full candidate list (it throws if the count
 * changed since the page loaded), so this re-reads the current rows and replaces just
 * the edited one, position-aligned by index — matching how the old combined form saved
 * candidate edits.
 *
 * @param {string} id
 * @param {number} index
 * @param {string} name
 * @param {string} details
 * @return {{ok:boolean, name:string, details:string}}
 */
function adminSaveCandidate(id, index, name, details) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findBallotSheet_(ss, id);
  if (!sheet) throw new Error('No ballot found for id "' + id + '".');

  var candidates = readBallotCandidates_(sheet);
  var rows = readBallotCandidateDetails_(sheet);
  while (rows.length < candidates.length) {
    rows.push({ name: candidates[rows.length], details: '' });
  }
  if (index < 0 || index >= rows.length) throw new Error('Invalid candidate index.');

  rows[index] = { name: String(name || '').trim(), details: String(details || '') };
  saveBallotCandidatesForId_(id, rows);
  return { ok: true, name: rows[index].name, details: rows[index].details };
}

/**
 * RPC: adds a new candidate from the "+ Add Candidate" panel.
 *
 * @param {string} id
 * @param {string} name
 * @param {string=} details
 * @return {{candidate:string}}
 */
function adminAddCandidate(id, name, details) {
  return addBallotCandidateForAdmin_(id, name, details);
}

/**
 * @param {string} id ballot id
 * @return {string} HTML with RCV + Condorcet results for one ballot.
 */
function _renderAdminAnalysis_(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findBallotSheet_(ss, id);
  var back = '<p><a target="_top" href="' + _escapeHtml_(_getWebAppUrl_()) + '?cmd=admin">&larr; Back to ballot list</a></p>';

  if (!sheet) {
    return back + '<p>No ballot found for id "' + _escapeHtml_(id) + '".</p>';
  }

  var results = runBallotAnalysis_(sheet);
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
    _escapeHtml_(getBallotSheetName_(id)) + ' sheet.</em></p>';

  return html;
}

/**
 * Runs RCV + all four Condorcet methods against a ballot's current
 * candidates/responses, writes a summary into the sheet's Results section,
 * and returns the raw results for HTML rendering.
 *
 * @param {Sheet} sheet a Ballot-<name> sheet
 * @return {{rcv:Object, condorcet:Object}|{error:string}}
 */
function runBallotAnalysis_(sheet) {
  var candidateNames = readBallotCandidates_(sheet);
  // Deduped to one ballot per respondent (case-insensitive name, latest wins) — a
  // respondent who re-voted must only count once, using their most recent ranking.
  var responseRows = _latestResponseByRespondent_(readBallotResponseRows_(sheet));

  if (candidateNames.length < 2) {
    return { error: 'Ballot needs at least two candidates before it can be analyzed.' };
  }
  if (responseRows.length === 0) {
    return { error: 'Ballot has no responses yet.' };
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

  writeBallotResults_(sheet, _buildResultsRows_(rcv, condorcet));

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
