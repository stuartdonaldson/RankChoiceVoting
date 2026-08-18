/**
 * Triggers.js
 *
 * onEdit(e) — a simple trigger (Apps Script auto-runs any project-level function named
 * exactly `onEdit`, no installable-trigger setup needed) that keeps BallotCache.js's cache
 * in sync with manual edits made directly in Sheets — typing over a config cell, renaming
 * a candidate in the Candidates table, etc. Anything made instead through the admin/ballot
 * RPCs (webAdmin.js's adminSave-prefixed setters and adminAddCandidate, webBallot.js's
 * addBallotTopic) already
 * refreshes the cache itself right after writing — see refreshBallotCache_'s other callers.
 *
 * Simple triggers run with restricted authorization (see
 * https://developers.google.com/apps-script/guides/triggers#restrictions) — they cannot
 * call services that need extra authorization (UrlFetchApp, GmailApp, etc.). Everything this
 * needs — PropertiesService, reading the sheet that was just edited — is allowed for a
 * simple trigger, so no installable trigger is required.
 *
 * Deliberately coarse: any edit anywhere on a Ballot-<id> sheet rebuilds that whole ballot's
 * cache entry (config + Candidates table), rather than working out whether the edited cell
 * actually falls inside the config/Candidates range. Rebuilding is cheap (a handful of
 * getRange().getValues() calls) and this only runs once per user edit, so the simplicity is
 * worth it over a second, edit-range-aware code path duplicating BallotModel.js's own
 * marker-scanning logic.
 *
 * @param {Object} e onEdit event object (Range/Spreadsheet — see
 *   https://developers.google.com/apps-script/guides/triggers/events#edit).
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    var name = sheet.getName();
    if (name.indexOf(BALLOT_SHEET_PREFIX) !== 0) return;
    var id = name.substring(BALLOT_SHEET_PREFIX.length);
    refreshBallotCache_(id, sheet);
  } catch (err) {
    // Never let a cache-refresh failure surface as a visible error to someone who's just
    // hand-editing the sheet — worst case the cache is left stale until the next
    // successful write/edit, and every read path already falls back to the live sheet on
    // a cache miss, so correctness is unaffected either way.
    Logger.log('onEdit cache refresh failed: ' + err);
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { onEdit: onEdit };
}
