/**
 * https://chatgpt.com/c/670fe32e-143c-8002-80c3-1f46a48f72f9
 */
/**
 * Requirement:
 * This script creates a Google Form for ranked choice voting (RCV).
 * - The form prompts for the voter's name.
 * - It uses a single grid question for voters to rank all candidates with numerical ranks.
 * - Instructions clarify that voters must assign a unique ranking to each candidate (1 for 1st choice, 2 for 2nd choice, etc.).
 * - The form responses are automatically linked to a Google Sheet with a sheet named "Candidate Responses".
 * - The script pulls candidate names from a Google Sheet named 'Candidates'.
 * 
 * Form Interface:
 * - The first question asks for the voter's name (Short answer).
 * - The second question is a grid with all candidates listed as rows and numeric ranks (1, 2, 3, etc.) as columns.
 * - Example for each candidate row: "Rank Candidate A" with columns for 1, 2, 3, etc.
 */

function createRankedChoiceVotingForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Candidates"); // Sheet with candidate list

  // Get candidate names and descriptions from columns 1 and 2
  var candidateData = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 2)
    .getValues(); // [[name, description], ...]

  var candidates = candidateData.map(function(row) { return row[0]; });
  var candidateListText = candidateData.map(function(row) {
    return '• ' + row[1]; // row[1] is the description (including book title)
  }).join('\n');

  // Get the name of the spreadsheet and append 'Form'
  var spreadsheetName = ss.getName();
  var formName = spreadsheetName + " Form";

  // Create the Google Form
  var form = FormApp.create(formName);

  // Add a question to capture the voter's name
  form.addTextItem().setTitle("What is your name?").setRequired(true);

  // Add form description with ranking instructions and candidate descriptions
  form.setDescription(
    "Please rank each of the candidates below. Assign a unique rank to each candidate, " +
    "starting with 1 for your first choice. Do not assign the same rank to more than one candidate.\n\n" +
    "Learn more about each candidate (book) below:\n" +
    candidateListText
  );

  // Create a grid question for ranking the candidates
  var grid = form.addGridItem();
  grid
    .setTitle("Rank the candidates below")
    .setRows(candidates) // Each row is a candidate
    .setColumns(candidates.map((_, index) => (index + 1).toString())) // 1, 2, 3, ... for ranking
    .setRequired(true); // Ensure that each candidate is ranked

  // delete all previous forms linked to this spreadsheet, as well as all sheets named "Form Responses..."
  cleanupFormResponses();

  // Link the form responses to the Google Sheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // Wait for the form to create the responses sheet and rename it
  var found = false;
  for (var attempt = 0; attempt < 20; attempt++) {
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().startsWith("Form Responses")) {
        sheets[i].setName("Candidate Responses");
        found = true;
        break;
      }
    }
    if (found) break;
    Utilities.sleep(500); // Wait 500 ms before trying again
  }
  if (!found) {
    Logger.log("Error: Could not find a sheet starting with 'Form Responses' to rename to 'Candidate Responses'.");
  }

  var formURL = form.getPublishedUrl();

  Logger.log("Form created: " + formURL);
  return formURL;
}

/** loadCandidates
 * Loads candidate names and descriptions from the "Candidates" sheet,
 * and returns an array of candidate objects.
 * Each candidate object has a name and description.
 */
function loadCandidates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var candidateSheet = ss.getSheetByName("Candidates");

  if (!candidateSheet) {
    Logger.log("Error: Candidates sheet not found.");
    return [];
  }

  // Get candidate names and descriptions from columns 1 and 2
  var candidateData = candidateSheet
    .getRange(2, 1, candidateSheet.getLastRow() - 1, 2)
    .getValues(); // [[name, description], ...]

  if (candidateData.length === 0) {
    Logger.log("Error: No candidates found in the Candidates sheet.");
    return [];
  }

  // Transform to { name, description }
  var candidates = candidateData.map(function(row) {
    return {
      name: row[0],
      description: row[1] || "", // Handle empty descriptions
    };
  });

  return candidates;
} 
/** loadBallots
 * Loads ballots from the "Candidate Responses" sheet,
 * deduplicates by voter name, and returns an array of ballots.
 */
function loadBallots() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var responseSheet = ss.getSheetByName("Candidate Responses");

  // Use loadCandidates to get candidate info and count
  var candidates = loadCandidates();
  if (!responseSheet || candidates.length === 0) {
    Logger.log("Error: Candidate Responses sheet not found or no candidates available.");
    return [];
  }

  var candidateNames = candidates.map(function(c) { return c.name; });

  if (responseSheet.getLastRow() < 2) {
    Logger.log("Error: No responses found in the Candidate Responses sheet.");
    return [];
  }

  // Read all responses, skipping the header row and timestamp column
  var rawResponses = responseSheet
    .getRange(2, 2, responseSheet.getLastRow() - 1, candidateNames.length + 1)
    .getValues();

  // Deduplicate by voter name, keeping only the latest response
  var uniqueResponses = {};
  rawResponses.forEach((response, index) => {
    var voterName = response[0]; // Voter name is in the first column after timestamp
    if (
      !uniqueResponses[voterName] ||
      uniqueResponses[voterName].index < index
    ) {
      uniqueResponses[voterName] = { response, index };
    }
  });

  // Transform to { voterName, ranks }
  var allBallots = Object.values(uniqueResponses).map((entry) => {
    var response = entry.response;
    return {
      voterName: response[0],
      ranks: response.slice(1),
    };
  });

  return allBallots;
}
