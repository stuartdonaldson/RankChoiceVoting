function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Voting and Ballot Tools')
    .addItem('Create Ballot Form', 'showCreateRCVFormDialog')
    .addItem('Run RCV Analysis', 'showProcessRCVResultsDialog')
    .addItem('Run Condorcet Analysis', 'showCondorcetAnalysisDialog')
    .addSeparator()
    .addItem('Open Survey Admin Page', 'showAdminPage')
    .addSeparator()
    .addItem('About', 'showAbout')
    .addToUi();
}

/**
 * Displays an About dialog with version info, the deployed web app URL, the
 * admin page link, and a link to every current Survey-<name> sheet — so
 * anyone opening the spreadsheet can find survey/admin URLs without digging
 * through Script Properties.
 */
function showAbout() {
  var url = _getWebAppUrl_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var surveyIds = (typeof listSurveyIds_ === 'function') ? listSurveyIds_(ss) : [];

  var surveyLinksHtml;
  if (!url) {
    surveyLinksHtml = '<p><em>WEBAPP_URL is not set yet — deploy this project as a web app ' +
      '(tools/manage-deployments.js sets it automatically after a PROD deploy).</em></p>';
  } else if (!surveyIds.length) {
    surveyLinksHtml = '<p><em>No surveys yet — use the admin page to create one.</em></p>';
  } else {
    surveyLinksHtml = '<ul style="padding-left:18px;">' + surveyIds.map(function (id) {
      var surveyUrl = url + '?cmd=survey&id=' + encodeURIComponent(id);
      var editUrl = url + '?cmd=admin&action=edit&id=' + encodeURIComponent(id);
      return '<li><b>' + id + '</b> — ' +
        '<a href="' + surveyUrl + '" target="_blank">view</a> | ' +
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
    '<p>Multi-survey ranked-choice/Condorcet voting web app bound to this spreadsheet.</p>' +
    '<hr>' +
    '<p><span class="label">Version:</span> ' + APP_VERSION + ' (' + APP_VERSION_DATE + ', ' + APP_DEPLOY_TARGET + ')</p>' +
    '<p><span class="label">Author:</span> ' + APP_AUTHOR + '</p>' +
    '<p><span class="label">Contact:</span> <a href="mailto:' + APP_CONTACT + '">' + APP_CONTACT + '</a></p>' +
    '<hr>' +
    '<p><span class="label">Web app URL:</span></p>' +
    '<p class="code">' + (url || 'unknown') + '</p>' +
    (url ? '<p><span class="label">Admin page:</span> <a href="' + url + '?cmd=admin" target="_blank">' + url + '?cmd=admin</a></p>' : '') +
    '<hr>' +
    '<p><span class="label">Surveys:</span></p>' +
    surveyLinksHtml
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

// Show a dialog with a link to the survey admin web page.
function showAdminPage() {
  var url = _getWebAppUrl_();
  var html;
  if (!url) {
    html = HtmlService.createHtmlOutput(
      '<p>WEBAPP_URL is not set yet. Deploy this project as a web app (see tools/manage-deployments.js), ' +
      'which sets the WEBAPP_URL script property automatically after a PROD deploy.</p>'
    ).setWidth(420).setHeight(160);
  } else {
    var adminUrl = url + '?cmd=admin';
    html = HtmlService.createHtmlOutput(
      '<p>Survey admin page: <a href="' + adminUrl + '" target="_blank">' + adminUrl + '</a></p>'
    ).setWidth(420).setHeight(150);
  }
  SpreadsheetApp.getUi().showModalDialog(html, 'Survey Admin');
}

// Show a dialog with the link to the created RCV form
function showCreateRCVFormDialog() {
  var formUrl = createRankedChoiceVotingForm(); // You must implement createRCVForm() to return the form URL
  var html = HtmlService.createHtmlOutput(
    '<p>RCV Form created: <a href="' + formUrl + '" target="_blank">' + formUrl + '</a>.  You may send this link out to the respective voters.  When voting has concluded, or to look at the current standing, select <b>Process RCV Data</b> from the <b>Ranked Choice Voting</b> menu. </p>'
  ).setWidth(400).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'RCV Form Created');
}

// Show a dialog with the results of the RCV processing
function showProcessRCVResultsDialog() {
  var {winner, tie, summary} = processRankedChoiceVotes(); // returns array of winners., 1 element means there is a winner more than one is a tie even after elimination rounds.
  var results;
  if (winner) {
    results = "<b>Winner: " + winner + "</b>";
  } else if (tie) {
    results = "<b>Tie after all eliminations:</b><br>" + summary.join(",<br>");    
  }
  results += "<br><br>" + formatCandidateSummaryHtml(summary);
        
  var instructions = '<p><b>Instructions:</b> To see the step-by-step elimination of candidates by round, view the "RCV Processing" sheet in this spreadsheet.</p>';
  var html = HtmlService.createHtmlOutput(
    instructions + results
  ).setWidth(500).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'RCV Results');
}

function showCondorcetAnalysisDialog() {
  var results = processCondorcetVoting(); // You must implement processCondorcetVotes() to return the results
  var html = HtmlService.createHtmlOutput(
    // add style for font to be a sans-serif font
    '<style>body { font-family: Arial, sans-serif; }</style>' +
    generateCondorcetResultsHtml(results)
  ).setWidth(600).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Condorcet Analysis Results');
}

// generate html output for the results of processcondorcet()
function generateCondorcetResultsHtml(results) {
  // Basic Condorcet
  var html = '<h3>Basic Condorcet</h3>';
  if (results.condorcet.winner) {
    html += '<p>Condorcet winner: <strong>' + results.condorcet.winner + '</strong></p>';
  } else {
    html += '<p>No Condorcet winner (cycle detected).</p>';
  }
  html += '<p><strong>Ranked Results:</strong></p>';
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
  html += '<p><strong>Ranked Results:</strong></p>';
  html += '<p style="font-size: 0.9em; margin-top: 5px; margin-bottom: 5px;"><em>Score = Sum of strongest path strengths over all opponents</em></p>';
  html += formatRankedCandidates(results.schulze.rankedCandidates);
//  html += '<pre>' + JSON.stringify(results.schulze.matrix, null, 2) + '</pre>';
//  html += '<pre>' + JSON.stringify(results.schulze.paths, null, 2) + '</pre>';

  // Ranked Pairs
  html += '<h3>Ranked Pairs (Tideman)</h3>';
  if (results.rankedPairs.winner) {
    html += '<p>Ranked Pairs winner: <strong>' + results.rankedPairs.winner + '</strong></p>';
  } else {
    html += '<p>No Ranked Pairs winner (cycle detected).</p>';
  }
  html += '<p><strong>Ranked Results:</strong></p>';
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
  html += '<p><strong>Ranked Results:</strong></p>';
  html += '<p style="font-size: 0.9em; margin-top: 5px; margin-bottom: 5px;"><em>Score = Worst pairwise defeat (lower is better)</em></p>';
  html += formatRankedCandidates(results.minimax.rankedCandidates);
//  html += '<pre>' + JSON.stringify(results.minimax.matrix, null, 2) + '</pre>';
  return html;
}

/**
 * Formats an RCV candidate summary table (as returned by
 * processRankedChoiceVotes()/runRankedChoiceVoting(), first row = header)
 * as an HTML table. Shared by the "Run RCV Analysis" menu dialog and the
 * web admin page (webAdmin.js).
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