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

  // delete all previous forms linked to this spreadsheet
  cleanupFormResponses();

  // Link the form responses to the Google Sheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // Wait for the form to create the responses sheet
  Utilities.sleep(2000); // Wait 2 seconds (may need to adjust)

  // Find the new responses sheet and rename it
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().startsWith("Form Responses")) {
      sheets[i].setName("Candidate Responses");
      break;
    }
  }

  var formURL = form.getPublishedUrl();

  Logger.log("Form created: " + formURL);
  return formURL;
}
