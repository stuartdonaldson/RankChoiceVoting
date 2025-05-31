function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ranked Choice Voting')
    .addItem('Create RCV Form', 'showCreateRCVFormDialog')
    .addItem('Process RCV Data', 'showProcessRCVResultsDialog')
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
  var results = processRankedChoiceVotes(); // returns array of winners., 1 element means there is a winner more than one is a tie even after elimination rounds.
  if (!results || results.length === 0) {
    results = "No votes found or no candidates available.";
  } else if (results.length === 1) {
    results = "Winner: " + results[0];
  } else {
    results = "Tie after all eliminations: " + results.join(", ");
  }
        
  var instructions = '<p><b>Instructions:</b> To see the step-by-step elimination of candidates by round, view the "RCV Processing" sheet in this spreadsheet.</p>';
  var html = HtmlService.createHtmlOutput(
    instructions + '<pre>' + results + '</pre>'
  ).setWidth(500).setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, 'RCV Results');
}