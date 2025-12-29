function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Voting and Ballot Tools')
    .addItem('Create Ballot Form', 'showCreateRCVFormDialog')
    .addItem('Run RCV Analysis', 'showProcessRCVResultsDialog')
    .addItem('Run Condorcet Analysis', 'showCondorcetAnalysisDialog')
    .addToUi();
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
  // summary is an a candidate summary table with the first row as a header row.
  // format this as an html table.
  var table = '<table border="1" style="border-collapse: collapse;">';
  // the first row of summary is the header row.
  table += '<tr>';
  for (var i = 0; i < summary[0].length; i++) {
    table += '<th>' + summary[0][i] + '</th>';
  }
  table += '</tr>';
  // the rest of the rows are the data rows.
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
  results += "<br><br>" + table;
        
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