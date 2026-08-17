function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Voting and Ballot Tools')
    .addItem('Open Ballot Admin Page', 'showAdminPage')
    .addSeparator()
    .addItem('About', 'showAbout')
    .addToUi();
}

/**
 * Displays an About dialog with version info, the deployed web app URL, the
 * admin page link, and a link to every current Ballot-<name> sheet — so
 * anyone opening the spreadsheet can find ballot/admin URLs without digging
 * through Script Properties.
 */
function showAbout() {
  var url = _getWebAppUrl_();
  // The admin/ballot UI lives on the static front end now (static-pages/src/index.html,
  // published by tools/publish-static-pages.js) — user-facing links point there directly
  // rather than at the GAS exec URL, which now only serves as a one-tap redirect to it (see
  // ApiBridge.js's _renderStaticRedirect_). `url` above is still shown separately below as the
  // backend URL, which is genuinely useful diagnostic info, just not where you want to click.
  var staticBase = (typeof _staticPagesBaseUrl_ === 'function') ? _staticPagesBaseUrl_() : '';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ballotIds = (typeof listBallotIds_ === 'function') ? listBallotIds_(ss) : [];

  var ballotLinksHtml;
  if (!staticBase) {
    ballotLinksHtml = '<p><em>Static hosting is not configured for this deployment ' +
      '(APP_DEPLOY_TARGET="' + APP_DEPLOY_TARGET + '") — deploy this project ' +
      '(tools/manage-deployments.js publishes the static pages automatically).</em></p>';
  } else if (!ballotIds.length) {
    ballotLinksHtml = '<p><em>No ballots yet — use the admin page to create one.</em></p>';
  } else {
    ballotLinksHtml = '<ul style="padding-left:18px;">' + ballotIds.map(function (id) {
      var ballotUrl = staticBase + '?cmd=ballot&id=' + encodeURIComponent(id);
      var editUrl = staticBase + '?cmd=admin&action=edit&id=' + encodeURIComponent(id);
      return '<li><b>' + id + '</b> — ' +
        '<a href="' + ballotUrl + '" target="_blank">view</a> | ' +
        '<a href="' + editUrl + '" target="_blank">edit</a></li>';
    }).join('') + '</ul>';
  }

  var html = HtmlService.createHtmlOutput(
    '<style>' +
    '  body { font-family: Arial, sans-serif; padding: 16px; font-size: 13px; color: #333; }' +
    '  h2 { margin-top: 0; }' +
    '  p { margin: 6px 0; }' +
    '  .label { font-weight: bold; }' +
    '  .code { font-family: monospace; font-size: 11px; word-break: break-all; background: #f5f5f5; padding: 4px; border-radius: 3px; }' +
    '  hr { border: none; border-top: 1px solid #ddd; margin: 12px 0; }' +
    '  a { color: #1a73e8; }' +
    '</style>' +
    '<h2>RankChoiceVoting</h2>' +
    '<p>Multi-ballot ranked-choice/Condorcet voting web app bound to this spreadsheet.</p>' +
    '<hr>' +
    '<p><span class="label">Version:</span> ' + APP_VERSION + ' (' + APP_VERSION_DATE + ', ' + APP_DEPLOY_TARGET + ')</p>' +
    '<p><span class="label">Author:</span> ' + APP_AUTHOR + '</p>' +
    '<p><span class="label">Contact:</span> <a href="mailto:' + APP_CONTACT + '">' + APP_CONTACT + '</a></p>' +
    '<hr>' +
    '<p><span class="label">Web app URL (backend):</span></p>' +
    '<p class="code">' + (url || 'unknown') + '</p>' +
    (staticBase ? '<p><span class="label">Admin page:</span> <a href="' + staticBase + '?cmd=admin" target="_blank">' + staticBase + '?cmd=admin</a></p>' : '') +
    '<hr>' +
    '<p><span class="label">Ballots:</span></p>' +
    ballotLinksHtml
  ).setWidth(480).setHeight(420);

  SpreadsheetApp.getUi().showModalDialog(html, 'About RankChoiceVoting');
}

/**
 * Returns the deployed web app's exec URL — WEBAPP_URL is set authoritatively
 * by tools/manage-deployments.js after each PROD deploy (via doPost
 * setWebappUrl in WebApp.js). Falls back to ScriptApp.getService().getUrl(),
 * which only resolves correctly when called from inside a running web app
 * request, not from the spreadsheet-bound editor context — so on a project
 * that has never been deployed/stamped, this falls back to an empty string.
 */
function _getWebAppUrl_() {
  var stored = PropertiesService.getScriptProperties().getProperty('WEBAPP_URL');
  if (stored) return stored;
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (err) {
    return '';
  }
}

/**
 * Show a dialog with a link to the ballot admin page — the static front end
 * (static-pages/src/index.html) directly, not the GAS exec URL (which now only redirects
 * there — see ApiBridge.js's _renderStaticRedirect_). Linking straight to the static page
 * skips that redirect hop for the one link people actually click from this menu.
 */
function showAdminPage() {
  var staticBase = (typeof _staticPagesBaseUrl_ === 'function') ? _staticPagesBaseUrl_() : '';
  var html;
  if (!staticBase) {
    html = HtmlService.createHtmlOutput(
      '<p>Static hosting is not configured for this deployment (APP_DEPLOY_TARGET="' +
      ((typeof APP_DEPLOY_TARGET !== 'undefined' && APP_DEPLOY_TARGET) || '') + '"). ' +
      'Deploy this project (see tools/manage-deployments.js), which publishes the static admin ' +
      'page automatically.</p>'
    ).setWidth(420).setHeight(160);
  } else {
    var adminUrl = staticBase + '?cmd=admin';
    html = HtmlService.createHtmlOutput(
      '<p>Ballot admin page: <a href="' + adminUrl + '" target="_blank">' + adminUrl + '</a></p>'
    ).setWidth(420).setHeight(150);
  }
  SpreadsheetApp.getUi().showModalDialog(html, 'Ballot Admin');
}

// Renders a finish-order array (as returned by computeCondorcetFinishOrder /
// computeRCVFinishOrder) as HTML, e.g. "1. C  2. B  3. A", with ties for a
// place shown as "(tie) B / A".
function _formatFinishOrderHtml_(order) {
  if (!order || !order.length) return '';
  return order.map(function (entry) {
    var label = entry.names.length > 1 ? '(tie) ' + entry.names.join(' / ') : entry.names[0];
    return entry.place + '. ' + label;
  }).join('&nbsp;&nbsp;');
}

// generate html output for the results of runBallotAnalysis_() (webAdmin.js)
// finishOrders is the {condorcet, schulze, rankedPairs, minimax} map of
// computeCondorcetFinishOrder() results - the true finish order, computed by
// removing each method's winner and re-running on who's left. It is not the
// same as rankedCandidates below, which is each method's own display
// heuristic scored against the full field (see rcballot-14e).
function generateCondorcetResultsHtml(results, finishOrders) {
  finishOrders = finishOrders || {};

  // Basic Condorcet
  var html = '<h3>Basic Condorcet</h3>';
  if (results.condorcet.winner) {
    html += '<p>Condorcet winner: <strong>' + results.condorcet.winner + '</strong></p>';
  } else {
    html += '<p>No Condorcet winner (cycle detected).</p>';
  }
  if (finishOrders.condorcet) {
    html += '<p><strong>Finish order:</strong> ' + _formatFinishOrderHtml_(finishOrders.condorcet) + '</p>';
  }
  html += '<p><strong>Heuristic display score (not the finish order):</strong></p>';
  html += '<p style="font-size: 0.9em; margin-top: 5px; margin-bottom: 5px;"><em>Score = Number of head-to-head victories</em></p>';
  html += formatRankedCandidates(results.condorcet.rankedCandidates);
 // html += '<pre>' + JSON.stringify(results.condorcet.matrix, null, 2) + '</pre>';

  // Schulze Method
  html += '<h3>Schulze Method</h3>';
  if (results.schulze.winner) {
    html += '<p>Schulze winner: <strong>' + results.schulze.winner + '</strong></p>';
  } else {
    html += '<p>No Schulze winner (cycle detected).</p>';
  }
  if (finishOrders.schulze) {
    html += '<p><strong>Finish order:</strong> ' + _formatFinishOrderHtml_(finishOrders.schulze) + '</p>';
  }
  html += '<p><strong>Heuristic display score (not the finish order):</strong></p>';
  html += '<p style="font-size: 0.9em; margin-top: 5px; margin-bottom: 5px;"><em>Score = Sum of strongest path strengths over all opponents</em></p>';
  html += formatRankedCandidates(results.schulze.rankedCandidates);
//  html += '<pre>' + JSON.stringify(results.schulze.matrix, null, 2) + '</pre>';
//  html += '<pre>' + JSON.stringify(results.schulze.paths, null, 2) + '</pre>';

  // Ranked Pairs
  html += '<h3>Ranked Pairs (Tideman)</h3>';
  if (results.rankedPairs.winner) {
    html += '<p>Ranked Pairs winner: <strong>' + results.rankedPairs.winner + '</strong></p>';
  } else if (results.rankedPairs.tie && results.rankedPairs.tie.length) {
    html += '<p>No Ranked Pairs winner (tie between: <strong>' + results.rankedPairs.tie.join(', ') + '</strong>).</p>';
  } else {
    html += '<p>No Ranked Pairs winner (cycle detected).</p>';
  }
  if (finishOrders.rankedPairs) {
    html += '<p><strong>Finish order:</strong> ' + _formatFinishOrderHtml_(finishOrders.rankedPairs) + '</p>';
  }
  html += '<p><strong>Heuristic display score (not the finish order):</strong></p>';
  html += '<p style="font-size: 0.9em; margin-top: 5px; margin-bottom: 5px;"><em>Score = Net locked edges (outgoing minus incoming)</em></p>';
  html += formatRankedCandidates(results.rankedPairs.rankedCandidates);
//  html += '<pre>' + JSON.stringify(results.rankedPairs.matrix, null, 2) + '</pre>';
//  html += '<pre>' + JSON.stringify(results.rankedPairs.locked, null, 2) + '</pre>';

  // Minimax
  html += '<h3>Minimax (Simpson)</h3>';
  if (results.minimax.winner) {
    html += '<p>Minimax winner: <strong>' + results.minimax.winner + '</strong></p>';
  } else {
    html += '<p>No Minimax winner (tie or cycle detected).</p>';
  }
  if (finishOrders.minimax) {
    html += '<p><strong>Finish order:</strong> ' + _formatFinishOrderHtml_(finishOrders.minimax) + '</p>';
  }
  html += '<p><strong>Heuristic display score (not the finish order):</strong></p>';
  html += '<p style="font-size: 0.9em; margin-top: 5px; margin-bottom: 5px;"><em>Score = Worst pairwise defeat (lower is better)</em></p>';
  html += formatRankedCandidates(results.minimax.rankedCandidates);
//  html += '<pre>' + JSON.stringify(results.minimax.matrix, null, 2) + '</pre>';
  return html;
}

/**
 * Formats an RCV candidate summary table (as returned by
 * runRankedChoiceVoting(), first row = header) as an HTML table. Used by
 * the web admin page (webAdmin.js).
 *
 * @param {Array<Array>} summary
 * @return {string}
 */
function formatCandidateSummaryHtml(summary) {
  var table = '<table border="1" style="border-collapse: collapse;">';
  table += '<tr>';
  for (var i = 0; i < summary[0].length; i++) {
    table += '<th>' + summary[0][i] + '</th>';
  }
  table += '</tr>';
  for (var i = 1; i < summary.length; i++) {
    table += '<tr>';
    for (var j = 0; j < summary[i].length; j++) {
      if (j === 0) {
        table += '<td style="text-align: left;">' + summary[i][j] + "</td>";
      } else {
        table += '<td style="text-align: center;">' + summary[i][j] + "</td>";
      }
    }
    table += '</tr>';
  }
  table += '</table>';
  return table;
}

// Helper function to format ranked candidates as an HTML table
function formatRankedCandidates(rankedCandidates) {
  var html = '<table border="1" style="border-collapse: collapse; margin-bottom: 15px;">';
  html += '<tr><th style="padding: 5px;">Rank</th><th style="padding: 5px;">Candidate</th><th style="padding: 5px;">Score</th></tr>';
  for (var i = 0; i < rankedCandidates.length; i++) {
    html += '<tr>';
    html += '<td style="text-align: center; padding: 5px;">' + (i + 1) + '</td>';
    html += '<td style="text-align: left; padding: 5px;">' + rankedCandidates[i].candidate + '</td>';
    html += '<td style="text-align: center; padding: 5px;">' + rankedCandidates[i].score + '</td>';
    html += '</tr>';
  }
  html += '</table>';
  return html;
}