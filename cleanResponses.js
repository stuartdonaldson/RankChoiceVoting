/**
 * Cleanup form responses - removing all forms that link to this sheet, and then removing all sheets named "Form Responses..."
 * which would also remove any duplicates created during testing or leftover by interrupted processes.
 */

function cleanupFormResponses() {
  deleteFormsLinkedToCurrentSpreadsheet();
  deleteAllFormResponsesSheets();
}

/**
 * Deletes all sheets in the active spreadsheet whose names start with "Form Responses".
 */
function deleteAllFormResponsesSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  sheets.forEach(function (sheet) {
    if (sheet.getName().startsWith("Form Responses") || sheet.getName() === "Candidate Responses") {
      Logger.log("Deleting sheet: " + sheet.getName());
      ss.deleteSheet(sheet);
    }
  });
}

/**
 * Deletes all forms that are linked to the current spreadsheet.
 * Use with caution: this will permanently delete the forms from Drive.
 */
function deleteFormsLinkedToCurrentSpreadsheet() {
  var forms = getFormsLinkedToCurrentSpreadsheet();
  if (forms.length === 0) {
    Logger.log("No forms are linked to this spreadsheet.");
    return;
  }
  forms.forEach(function (file) {
    try {
      var form = FormApp.openById(file.getId());
      form.removeDestination(); // Unlink the form from the spreadsheet
      Logger.log(
        "Unlinked form: " + file.getName() + " (" + file.getId() + ")"
      );
    } catch (e) {
      Logger.log(
        "Could not unlink form: " + file.getName() + " (" + file.getId() + ")"
      );
    }
    Logger.log("Deleting form: " + file.getName() + " (" + file.getId() + ")");
    file.setTrashed(true); // Moves to trash
  });
  Logger.log(forms.length + " form(s) unlinked and moved to trash.");
}
/**
 * Returns an array of form file objects (from DriveApp) that are linked to the current spreadsheet.
 */
function getFormsLinkedToCurrentSpreadsheet() {
  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var linkedForms = [];
  var files = DriveApp.getFilesByType(MimeType.GOOGLE_FORMS);

  while (files.hasNext()) {
    var file = files.next();
    try {
      var form = FormApp.openById(file.getId());
      var destId = form.getDestinationId();
      if (destId && destId === ssId) {
        linkedForms.push(file);
      }
    } catch (e) {
      // Skip forms you can't open
      continue;
    }
  }
  return linkedForms;
}
