/**
 * BallotModel.js
 *
 * Sheet-layout model for "Ballot-<name>" sheets. Each ballot lives entirely in
 * one sheet, top to bottom:
 *
 *   Row 1..8   Config (key in col A, value in col B): Title, Description,
 *              Instructions, Footer, Contact, Accept-New, Add-Instructions, Admin-Only-Notes
 *              — Description is the landing-page intro (shown before a respondent
 *              enters their name); Instructions is shown above the ranking list on
 *              the ballot page instead; Add-Instructions is shown above the "+ Add
 *              New" button, only when Accept-New is on.
 *   Row 9      blank
 *   Row 10     "[Results]" marker (col A)
 *   Row 11+    Analysis output — cleared and rewritten by runBallotAnalysis_
 *   Row M      "[Candidates]" marker (col A)
 *   Row M+1    Candidates header: (col A blank) Name (col B) | Details (col C)
 *   Row M+2+   One row per candidate — name/details start at col B, position-aligned
 *              with the Responses header's candidate columns (candidate i's row here
 *              == candidate column i)
 *   Row N      "[Responses]" marker (col A)
 *   Row N+1    Responses header: Date | Name | Weight | Comment | <candidate1> | <candidate2> | ...
 *   Row N+2+   One row per respondent
 *
 * Column A is reserved EXCLUSIVELY for section markers/config keys — never for
 * user-entered or app-generated data — and markers are bracket-decorated ("[Results]",
 * not "Results") as a second, independent guard. Both exist because _findMarkerRow_
 * does an unscoped top-to-bottom scan of column A for exact marker text: if a
 * candidate's name (user-controlled, arbitrary) ever landed in column A, naming a
 * candidate e.g. "Responses" would shadow the real marker and corrupt every section
 * boundary computed from it (this was a real bug in an earlier version of this file,
 * where the Candidates table's Name column started at col A). Reserving col A for
 * markers only closes the hole entirely; the bracket decoration is defense in depth.
 * Hence: the Results output and the Candidates table both start at col B; marker
 * matching is exact bracketed text only (plain "Responses" etc. is never matched or
 * rewritten); and candidate names equal to any marker word are rejected.
 *
 * The one documented exception: the Responses section keeps column A for its Date
 * column (header "Date" plus app-written timestamps in the response rows) — it sits
 * below the [Responses] marker, where no marker scan needs to find anything after it.
 *
 * "Admin-Only-Notes" (config key; formerly "Info" — _findMarkerRow_-style legacy
 * migration in readBallotConfig_ upgrades old sheets on first read) is free-text
 * notes about the ballot (e.g. purpose, audience, scheduling context) — shown in
 * the admin list as a per-row disclosure, never exposed to respondents (see
 * getBallotForRespondent_, which omits it).
 *
 * The Candidates table is the PRIMARY identification of candidates — it is the
 * source of truth for who is a candidate, including ones pre-populated there ahead
 * of any votes. The Responses header's candidate columns (E onward) are a derived,
 * position-aligned mirror that response rows key their ranks off of; readBallotCandidates_
 * reconciles the two on every read (via _reconcileCandidatesWithResponses_), appending
 * a Responses column for any Candidates-table entry that doesn't have one yet, so a
 * candidate typed only into the Candidates table still shows up everywhere (respondent
 * page, admin edit form, analysis) even before its first vote. Existing Responses
 * columns are never reordered or removed by this reconciliation, so previously recorded
 * ranks stay aligned to their column. The table also carries a per-candidate "Details"
 * note editable from the admin page; it is kept in sync by saveBallotCandidates_/
 * addBallotCandidate_, not by hand-editing. A ballot created before the Candidates
 * table existed has no Candidates marker — _ensureCandidatesSection_ inserts it (empty)
 * the first time anything here needs it, so older sheets self-heal rather than throwing.
 *
 * webBallot.js (respondent UI) and webAdmin.js (analysis) both read/write
 * through this module so the sheet layout only needs to be understood once.
 */

var BALLOT_SHEET_PREFIX = 'Ballot-';
var _BALLOT_RESULTS_MARKER = '[Results]';
var _BALLOT_CANDIDATES_MARKER = '[Candidates]';
var _BALLOT_CANDIDATES_HEADER = ['Name', 'Details', 'Link Text', 'Link URL'];
var _BALLOT_CANDIDATES_FIRST_COL = 2; // col B — col A stays blank/reserved throughout the Candidates section
var _BALLOT_RESPONSES_MARKER = '[Responses]';
// Bare words (lowercase, brackets stripped) a candidate may not be named: the marker
// words plus the former "Items" section name. See the file header's column-A rule.
var _BALLOT_RESERVED_NAMES_ = ['results', 'candidates', 'responses', 'items'];
var _BALLOT_CONFIG_ROWS = ['Title', 'Description', 'Instructions', 'Footer', 'Contact', 'Accept-New', 'Add-Instructions', 'Admin-Only-Notes'];
// Legacy config key recognized by readBallotConfig_ purely to migrate older sheets in
// place (the cell is rewritten to the current key name the first time it is read).
var _BALLOT_LEGACY_CONFIG_KEYS_ = { 'Admin-Only-Notes': ['Info'] };
// Fixed Responses columns before the candidate columns begin (col 5 / E).
var _BALLOT_RESPONSE_FIXED_COLS = ['Date', 'Name', 'Weight', 'Comment'];
var _BALLOT_FIRST_CANDIDATE_COL = _BALLOT_RESPONSE_FIXED_COLS.length + 1; // 5
var _BALLOT_LOCK_TIMEOUT_MS = 30000;
var _ballotLockDepth_ = 0; // nesting depth of _withBallotLock_ in this execution (GAS runs one request per isolate)
// Bold + light-blue background applied to every section marker/header row so the
// sheet's structure (Results / Candidates / Responses) is easy to scan visually.
var _BALLOT_SECTION_HIGHLIGHT_BG = '#e8f0fe';

/**
 * @param {string} id
 * @return {string} sheet name, e.g. "Ballot-BoardElection2026"
 */
function getBallotSheetName_(id) {
  return BALLOT_SHEET_PREFIX + String(id || '').trim();
}

/**
 * Lists ballot ids (sheet name with the "Ballot-" prefix stripped) for every
 * matching sheet in the spreadsheet, in sheet order.
 *
 * @param {Spreadsheet} ss
 * @return {Array<string>}
 */
function listBallotIds_(ss) {
  return ss.getSheets()
    .map(function (s) { return s.getName(); })
    .filter(function (name) { return name.indexOf(BALLOT_SHEET_PREFIX) === 0; })
    .map(function (name) { return name.substring(BALLOT_SHEET_PREFIX.length); });
}

/**
 * Returns the sheet for a ballot id, or null if it does not exist.
 *
 * @param {Spreadsheet} ss
 * @param {string} id
 * @return {Sheet|null}
 */
function findBallotSheet_(ss, id) {
  return ss.getSheetByName(getBallotSheetName_(id));
}

/**
 * Creates a new Ballot-<id> sheet with placeholder configuration and an
 * empty Results/Responses skeleton. Placeholder values start with a "[TODO: ...]"
 * instruction (easy to spot and grep for) followed by example text to edit.
 *
 * @param {Spreadsheet} ss
 * @param {string} id
 * @return {Sheet}
 */
function createBallotSheet_(ss, id) {
  var sheet = ss.insertSheet(getBallotSheetName_(id));
  sheet.getRange(1, 1, 8, 2).setValues([
    // Each placeholder is "[TODO: what to enter — delete this bracket] <example>": the
    // bracketed lead-in says what the field is for, and the text after it is a working
    // example, so the admin deletes the bracket and edits the example rather than writing
    // from scratch. Description/Instructions also fold in the "how to use this" guidance the
    // static ballot page (static-pages/src/index.html) used to render as fixed copy — every
    // visible string on the page is admin-editable, none of it hardcoded. Existing ballots
    // keep whatever they already have; this only affects new ballots' starting text.
    ['Title', '[TODO: ballot title shown to respondents — delete this bracket and edit the example] F3 Active Book Club — Next Book Selection'],
    ['Description', '[TODO: describe what this ballot is about — delete this bracket and edit the example] Help choose our next ABC read! Rank the books you’d be most interested in reading and discussing with the PAX. You don’t need to have read any of them to vote. Enter your name below to start or resume ranking — if you’ve already been here, your previous answers will be loaded so you can review or change them.'],
    ['Instructions', '[TODO: any extra guidance for ranking these particular candidates — delete this bracket and edit the example] Drag your favorites to the top, ranking as many or as few as you’d like. Your first choice should be the book you’re most excited to discuss.'],
    ['Footer', '[TODO: footer text, e.g. deadline or sponsoring group — delete this bracket and edit the example] Voting closes [date and time]. The book with the strongest ranked-choice support will be our next ABC selection.'],
    ['Contact', '[TODO: contact name/email for questions — delete this bracket and edit the example] Questions or suggestions? Post in #ao-active-book-club.'],
    ['Accept-New', 'TRUE'],
    ['Add-Instructions', '[TODO: instructions shown above the "+ Add New" button, if enabled — delete this bracket and edit the example] Have another book you’d like ABC to consider? Add the title and author, and preferably a link to the book.'],
    ['Admin-Only-Notes', '']
  ]);
  sheet.getRange(10, 1).setValue(_BALLOT_RESULTS_MARKER);
  sheet.getRange(11, 1).setValue(_BALLOT_CANDIDATES_MARKER);
  sheet.getRange(12, _BALLOT_CANDIDATES_FIRST_COL, 1, _BALLOT_CANDIDATES_HEADER.length).setValues([_BALLOT_CANDIDATES_HEADER]);
  sheet.getRange(13, 1).setValue(_BALLOT_RESPONSES_MARKER);
  sheet.getRange(14, 1, 1, _BALLOT_RESPONSE_FIXED_COLS.length).setValues([_BALLOT_RESPONSE_FIXED_COLS]);
  _highlightSectionMarkers_(sheet);
  return sheet;
}

/**
 * Bolds + colors every section marker cell (Results/Candidates/Responses) and the
 * Candidates/Responses header rows, so the sheet's structure is easy to pick out
 * visually. Idempotent and safe to call repeatedly (e.g. every time the admin edit
 * page loads), which is how older sheets created before this existed pick up the
 * highlighting.
 *
 * @param {Sheet} sheet
 */
function _highlightSectionMarkers_(sheet) {
  function highlightRow(row, numCols) {
    if (row === -1) return;
    sheet.getRange(row, 1, 1, numCols || 1).setBackground(_BALLOT_SECTION_HIGHLIGHT_BG).setFontWeight('bold');
  }

  var resultsRow = _findMarkerRow_(sheet, _BALLOT_RESULTS_MARKER);
  var candidatesRow = _findMarkerRow_(sheet, _BALLOT_CANDIDATES_MARKER);
  var responsesRow = _findMarkerRow_(sheet, _BALLOT_RESPONSES_MARKER);

  highlightRow(resultsRow, 1);
  highlightRow(candidatesRow, 1);
  // Highlight col A through the end of the Name/Details header as one continuous band,
  // even though col A itself is blank in the Candidates section (reserved — see file header).
  if (candidatesRow !== -1) highlightRow(candidatesRow + 1, _BALLOT_CANDIDATES_FIRST_COL - 1 + _BALLOT_CANDIDATES_HEADER.length);
  highlightRow(responsesRow, 1);
  if (responsesRow !== -1) {
    var lastCol = Math.max(sheet.getLastColumn(), _BALLOT_RESPONSE_FIXED_COLS.length);
    highlightRow(responsesRow + 1, lastCol);
    _unfreezeStaleHeaderRows_(sheet, responsesRow + 1);
  }
}

/**
 * Undoes the old fixed 14-row freeze (formerly set at sheet creation, where row 14 was
 * the Responses header) once the Responses header has moved off row 14 — the freeze then
 * pinned analysis output instead. Only exactly 14 is treated as that legacy value; any
 * other frozen-row count is a user choice and is left alone.
 *
 * @param {Sheet} sheet
 * @param {number} responsesHeaderRow
 */
function _unfreezeStaleHeaderRows_(sheet, responsesHeaderRow) {
  if (sheet.getFrozenRows() !== 14 || responsesHeaderRow === 14) return;
  _withBallotLock_(function () {
    if (sheet.getFrozenRows() !== 14) return; // re-check under the lock
    sheet.setFrozenRows(0);
    _logBallotEvent_('ballot.repair.frozenRows', { sheet: sheet.getName(), responsesHeaderRow: responsesHeaderRow });
  });
}

/**
 * Inserts an empty Candidates marker + header row directly above the Responses marker
 * if this sheet doesn't have one yet (a ballot created before the Candidates section
 * existed). No-op if the sheet already has a Candidates marker, or has no Responses
 * marker to anchor the insert to (shouldn't happen for any sheet created via
 * createBallotSheet_).
 *
 * @param {Sheet} sheet
 */
function _ensureCandidatesSection_(sheet) {
  if (_findMarkerRow_(sheet, _BALLOT_CANDIDATES_MARKER) !== -1) return;
  if (_findMarkerRow_(sheet, _BALLOT_RESPONSES_MARKER) === -1) return;
  _withBallotLock_(function () {
    // Re-check under the lock: another request may have repaired (or shifted rows) meanwhile.
    if (_findMarkerRow_(sheet, _BALLOT_CANDIDATES_MARKER) !== -1) return;
    var responsesRow = _findMarkerRow_(sheet, _BALLOT_RESPONSES_MARKER);
    if (responsesRow === -1) return;
    sheet.insertRowsBefore(responsesRow, 2);
    sheet.getRange(responsesRow, 1).setValue(_BALLOT_CANDIDATES_MARKER);
    sheet.getRange(responsesRow + 1, _BALLOT_CANDIDATES_FIRST_COL, 1, _BALLOT_CANDIDATES_HEADER.length).setValues([_BALLOT_CANDIDATES_HEADER]);
  });
}

/**
 * Runs fn while holding the script lock, so every sheet read-modify-write on a ballot
 * is serialized against every other. Re-entrant: only the outermost call (depth 0)
 * acquires the lock and releases it (in finally, so a throw can't leak it); nested calls
 * — e.g. adminSaveCandidate holding the lock around saveBallotCandidatesForId_, or a
 * self-repair inside a locked write — just run fn. The script lock itself is not
 * re-entrant, so a second waitLock in the same execution would block until it timed out.
 *
 * @param {function(): *} fn
 * @return {*} fn's return value.
 */
function _withBallotLock_(fn) {
  if (_ballotLockDepth_ > 0) {
    _ballotLockDepth_++;
    try {
      return fn();
    } finally {
      _ballotLockDepth_--;
    }
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(_BALLOT_LOCK_TIMEOUT_MS);
  _ballotLockDepth_ = 1;
  try {
    return fn();
  } finally {
    _ballotLockDepth_ = 0;
    lock.releaseLock();
  }
}

/**
 * Logs a sheet-parsing/repair event. Goes through GasLogger when it's loaded (web app
 * requests, where entries also reach Drive/Axiom); otherwise — e.g. the onEdit simple
 * trigger's restricted context, or Node tests — straight to Logger, which still shows in
 * the Apps Script Executions log. Callers pass sheet names, row numbers and counts only.
 *
 * @param {string} tag
 * @param {Object} data
 */
function _logBallotEvent_(tag, data) {
  if (typeof GasLogger !== 'undefined') {
    GasLogger.log(tag, data);
  } else {
    Logger.log(tag + ' ' + JSON.stringify(data));
  }
}

/**
 * Re-creates a lost [Responses] marker + header directly after the sheet's last row —
 * the typical cause being a hand-paste of candidates over rows that held them. Without
 * it every section boundary below [Candidates] is undefined and readBallotCandidateDetails_
 * finds nothing. Only repairs when column A below the Candidates header is entirely blank:
 * column A is marker-only, so any text there may be orphaned response rows (Date in col A)
 * and appending a header beneath them would misfile them — that case is logged, not touched.
 * No-op when the sheet has a [Responses] marker or no [Candidates] marker.
 *
 * @param {Sheet} sheet
 */
function _ensureResponsesSection_(sheet) {
  var candidatesRow = _findMarkerRow_(sheet, _BALLOT_CANDIDATES_MARKER);
  if (candidatesRow === -1 || _findMarkerRow_(sheet, _BALLOT_RESPONSES_MARKER) !== -1) return;
  _withBallotLock_(function () {
    // Re-check under the lock before touching anything.
    var candidatesRow = _findMarkerRow_(sheet, _BALLOT_CANDIDATES_MARKER);
    if (candidatesRow === -1 || _findMarkerRow_(sheet, _BALLOT_RESPONSES_MARKER) !== -1) return;
    _repairResponsesSectionLocked_(sheet, candidatesRow);
  });
}

/**
 * Body of _ensureResponsesSection_; caller holds the ballot lock and has verified the
 * [Candidates] marker exists at candidatesRow and [Responses] does not.
 *
 * @param {Sheet} sheet
 * @param {number} candidatesRow
 */
function _repairResponsesSectionLocked_(sheet, candidatesRow) {
  var lastRow = sheet.getLastRow();
  var firstDataRow = candidatesRow + 2;
  var colARows = [];
  if (lastRow >= firstDataRow) {
    sheet.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, 1).getValues().forEach(function (r, i) {
      if (String(r[0] || '').trim()) colARows.push(firstDataRow + i);
    });
  }
  if (colARows.length) {
    _logBallotEvent_('ballot.parse.responsesMarkerMissing', { sheet: sheet.getName(), candidatesRow: candidatesRow, colARows: colARows });
    return;
  }
  var markerRow = Math.max(lastRow, candidatesRow + 1) + 1;
  sheet.getRange(markerRow, 1).setValue(_BALLOT_RESPONSES_MARKER);
  sheet.getRange(markerRow + 1, 1, 1, _BALLOT_RESPONSE_FIXED_COLS.length).setValues([_BALLOT_RESPONSE_FIXED_COLS]);
  _logBallotEvent_('ballot.repair.responsesSection', { sheet: sheet.getName(), row: markerRow });
}

/**
 * Migrates a Candidates section still using the OLD column layout (Name/Details in
 * columns A/B, from before column A was reserved for markers only) to the current
 * layout (columns B/C, col A left blank). Detected by the Candidates header's column A
 * cell still holding the literal "Name" label — the current layout never writes
 * anything to column A there. A no-op for a sheet already on the new layout (or
 * with no Candidates section at all — call _ensureCandidatesSection_ first).
 *
 * Without this, readBallotCandidateDetails_ silently reads garbage (the OLD Details
 * column mistaken for the new Name column) rather than throwing, because both
 * layouts are structurally valid 2-column reads — this was found live on a ballot
 * created before the column-A-reservation fix shipped, where every candidate's Details
 * came back empty even though the sheet clearly had them.
 *
 * @param {Sheet} sheet
 */
function _migrateCandidatesColumnsIfNeeded_(sheet) {
  var candidatesRow = _findMarkerRow_(sheet, _BALLOT_CANDIDATES_MARKER);
  if (candidatesRow === -1) return;
  var candidatesHeaderRow = candidatesRow + 1;
  var headerColA = String(sheet.getRange(candidatesHeaderRow, 1).getValue() || '').trim();
  if (headerColA !== _BALLOT_CANDIDATES_HEADER[0]) return; // already migrated (col A blank) or unrecognized

  _withBallotLock_(function () {
    // Re-check under the lock: a concurrent request may already have migrated.
    var candidatesRow = _findMarkerRow_(sheet, _BALLOT_CANDIDATES_MARKER);
    if (candidatesRow === -1) return;
    var candidatesHeaderRow = candidatesRow + 1;
    var headerColA = String(sheet.getRange(candidatesHeaderRow, 1).getValue() || '').trim();
    if (headerColA !== _BALLOT_CANDIDATES_HEADER[0]) return;

    var responsesRow = _findMarkerRow_(sheet, _BALLOT_RESPONSES_MARKER);
    if (responsesRow === -1) return;
    var count = responsesRow - candidatesHeaderRow - 1;
    var oldValues = count > 0 ? sheet.getRange(candidatesHeaderRow + 1, 1, count, 2).getValues() : [];

    sheet.getRange(candidatesHeaderRow, 1, count + 1, 1).clearContent(); // vacate col A (header + data rows)
    sheet.getRange(candidatesHeaderRow, _BALLOT_CANDIDATES_FIRST_COL, 1, _BALLOT_CANDIDATES_HEADER.length).setValues([_BALLOT_CANDIDATES_HEADER]);
    if (count > 0) {
      sheet.getRange(candidatesHeaderRow + 1, _BALLOT_CANDIDATES_FIRST_COL, count, 2).setValues(oldValues);
    }
  });
}

/**
 * Backfills the Candidates header row with any header columns added to
 * _BALLOT_CANDIDATES_HEADER after a sheet was created (e.g. "Link Text"/"Link URL"
 * added after Name/Details) — a sheet already on the current header is left untouched.
 * Only ever appends missing trailing header cells; never rewrites the existing ones,
 * so it can't clobber a sheet mid-migration from the old column-A layout.
 *
 * @param {Sheet} sheet
 */
function _ensureCandidatesHeaderColumns_(sheet) {
  var candidatesRow = _findMarkerRow_(sheet, _BALLOT_CANDIDATES_MARKER);
  if (candidatesRow === -1) return;
  var headerRow = candidatesRow + 1;
  var current = sheet.getRange(headerRow, _BALLOT_CANDIDATES_FIRST_COL, 1, _BALLOT_CANDIDATES_HEADER.length).getValues()[0];
  var missing = _BALLOT_CANDIDATES_HEADER.filter(function (label, i) { return String(current[i] || '').trim() === ''; });
  if (!missing.length) return;
  _withBallotLock_(function () {
    // Re-check under the lock before appending.
    var candidatesRow = _findMarkerRow_(sheet, _BALLOT_CANDIDATES_MARKER);
    if (candidatesRow === -1) return;
    var headerRow = candidatesRow + 1;
    var current = sheet.getRange(headerRow, _BALLOT_CANDIDATES_FIRST_COL, 1, _BALLOT_CANDIDATES_HEADER.length).getValues()[0];
    var missing = _BALLOT_CANDIDATES_HEADER.filter(function (label, i) { return String(current[i] || '').trim() === ''; });
    if (!missing.length) return;
    var firstMissingCol = _BALLOT_CANDIDATES_FIRST_COL + (_BALLOT_CANDIDATES_HEADER.length - missing.length);
    sheet.getRange(headerRow, firstMissingCol, 1, missing.length).setValues([missing]);
  });
}

/**
 * Inserts/deletes rows so exactly rows.length rows are available directly below
 * `afterRow`, ending just before `beforeMarkerRow`, then writes rows into that space.
 * Shared resize primitive behind both the Results section (writeBallotResults_) and
 * the Candidates section (writeBallotCandidateDetails_) — everything at/after
 * beforeMarkerRow shifts as one block, so content further down the sheet is never
 * disturbed.
 *
 * @param {Sheet} sheet
 * @param {number} afterRow
 * @param {number} beforeMarkerRow
 * @param {Array<Array>} rows
 * @param {number=} startCol defaults to col A (1) — Candidates-section and Results callers
 *   pass _BALLOT_CANDIDATES_FIRST_COL so their data never lands in the marker-reserved column A.
 */
function _resizeSectionRows_(sheet, afterRow, beforeMarkerRow, rows, startCol) {
  startCol = startCol || 1;
  var available = beforeMarkerRow - afterRow - 1;
  var needed = rows.length;
  if (needed > available) {
    sheet.insertRowsAfter(afterRow, needed - available);
  } else if (needed < available) {
    sheet.deleteRows(afterRow + 1, available - needed);
  }

  if (needed > 0) {
    var maxCols = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 1);
    var padded = rows.map(function (r) {
      var copy = r.slice();
      while (copy.length < maxCols) copy.push('');
      return copy;
    });
    sheet.getRange(afterRow + 1, startCol, needed, maxCols).setValues(padded);
  }
}

/**
 * Returns the sheet for a ballot id, creating a placeholder skeleton if it
 * does not already exist.
 *
 * @param {Spreadsheet} ss
 * @param {string} id
 * @return {Sheet}
 */
function getOrCreateBallotSheet_(ss, id) {
  return findBallotSheet_(ss, id) || createBallotSheet_(ss, id);
}

/**
 * Validates a candidate ballot id and creates a brand-new Ballot-<id> sheet.
 * Used by the admin "Create New Ballot" form (webAdmin.js). Ids are
 * restricted to characters that are safe in both a sheet name and a URL
 * query parameter without encoding.
 *
 * @param {Spreadsheet} ss
 * @param {string} id
 * @return {Sheet}
 */
function createNewBallot_(ss, id) {
  id = String(id || '').trim();
  if (!id) throw new Error('Ballot ID is required.');
  if (!/^[A-Za-z0-9][A-Za-z0-9_-]*$/.test(id)) {
    throw new Error('Ballot ID may only contain letters, numbers, "-" and "_", and must start with a letter or number.');
  }
  if (findBallotSheet_(ss, id)) {
    throw new Error('A ballot with id "' + id + '" already exists.');
  }
  return createBallotSheet_(ss, id);
}

/**
 * Overwrites the config key/value rows for a ballot. Keys already present are
 * matched by scanning column A rather than assuming fixed row numbers, so it
 * tolerates a sheet whose config rows have been manually reordered. Any update
 * key with no existing row (e.g. a config field — like "Admin-Only-Notes" — added
 * to the model after this particular sheet was created) is appended as a new row,
 * inserted just above the Results marker, so older sheets self-heal on first
 * save instead of silently dropping the value.
 *
 * @param {Sheet} sheet
 * @param {Object} updates e.g. {Title, Description, Instructions, Footer, Contact, 'Accept-New', 'Add-Instructions', 'Admin-Only-Notes'}
 */
function writeBallotConfig_(sheet, updates) {
  var resultsRow = _findMarkerRow_(sheet, _BALLOT_RESULTS_MARKER);
  var lastConfigRow = resultsRow === -1 ? _BALLOT_CONFIG_ROWS.length : resultsRow - 1;
  if (lastConfigRow < 1) return;
  var keys = sheet.getRange(1, 1, lastConfigRow, 1).getValues();
  var seenKeys = {};
  var blankRows = [];
  for (var i = 0; i < keys.length; i++) {
    var key = String(keys[i][0] || '').trim();
    if (key) {
      seenKeys[key] = true;
      if (updates.hasOwnProperty(key)) {
        sheet.getRange(i + 1, 2).setValue(updates[key]);
      }
    } else {
      blankRows.push(i + 1);
    }
  }

  var missingKeys = Object.keys(updates).filter(function (k) { return !seenKeys[k]; });
  if (!missingKeys.length) return;

  // Reuse existing blank config rows (e.g. the spacer row above the Results marker)
  // before inserting brand-new ones, so an older sheet doesn't accumulate a stray
  // blank line every time a new config field is introduced.
  missingKeys.forEach(function (key) {
    if (blankRows.length) {
      sheet.getRange(blankRows.shift(), 1, 1, 2).setValues([[key, updates[key]]]);
    } else if (resultsRow !== -1) {
      sheet.insertRowsBefore(resultsRow, 1);
      sheet.getRange(lastConfigRow + 1, 1, 1, 2).setValues([[key, updates[key]]]);
      resultsRow++;
      lastConfigRow++;
    }
    // else: no marker and no spare blank row — nothing safe to insert relative to,
    // so this key is left unwritten (matches the historical no-op behavior).
  });
}

/**
 * Scans column A for a row whose value exactly matches the (bracketed) marker. Plain
 * look-alike text such as "Responses" is never matched.
 *
 * @param {Sheet} sheet
 * @param {string} marker
 * @return {number} 1-based row, or -1 if not found.
 */
function _findMarkerRow_(sheet, marker) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return -1;
  var values = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === marker) return i + 1;
  }
  return -1;
}

/**
 * Throws if name (trimmed, case-insensitive, with or without surrounding brackets) is
 * one of the section-marker words — such a candidate would put marker-looking text into
 * the sheet (e.g. in the Responses header) and invite confusion with section boundaries.
 *
 * @param {string} name
 */
function _assertNotMarkerName_(name) {
  var bare = String(name || '').trim().toLowerCase();
  if (bare.charAt(0) === '[' && bare.charAt(bare.length - 1) === ']') bare = bare.slice(1, -1).trim();
  if (_BALLOT_RESERVED_NAMES_.indexOf(bare) !== -1) {
    throw new Error('Candidate name cannot be a section marker.');
  }
}

/**
 * Reads the config key/value rows (rows 1 through the row before the
 * Results marker).
 *
 * @param {Sheet} sheet
 * @return {Object} e.g. {Title, Description, Instructions, Footer, Contact, 'Accept-New', 'Add-Instructions'}
 */
function readBallotConfig_(sheet) {
  var resultsRow = _findMarkerRow_(sheet, _BALLOT_RESULTS_MARKER);
  var lastConfigRow = resultsRow === -1 ? _BALLOT_CONFIG_ROWS.length : resultsRow - 1;
  var config = {};
  if (lastConfigRow < 1) return config;
  var values = sheet.getRange(1, 1, lastConfigRow, 2).getValues();
  values.forEach(function (row, i) {
    var key = String(row[0] || '').trim();
    if (!key) return;
    Object.keys(_BALLOT_LEGACY_CONFIG_KEYS_).forEach(function (currentKey) {
      if (_BALLOT_LEGACY_CONFIG_KEYS_[currentKey].indexOf(key) !== -1) {
        var legacyKey = key;
        _withBallotLock_(function () {
          if (String(sheet.getRange(i + 1, 1).getValue() || '').trim() === legacyKey) {
            sheet.getRange(i + 1, 1).setValue(currentKey); // one-time upgrade to the current key name
          }
        });
        key = currentKey;
      }
    });
    config[key] = row[1];
  });
  return config;
}

/**
 * @param {Object} config as returned by readBallotConfig_
 * @return {boolean}
 */
function _ballotAcceptsNew_(config) {
  return String(config['Accept-New'] || '').trim().toUpperCase() === 'TRUE';
}

/**
 * @param {Sheet} sheet
 * @return {number} 1-based row of the Responses header (Date/Name/Weight/Comment/...).
 */
function _getResponsesHeaderRow_(sheet) {
  var marker = _findMarkerRow_(sheet, _BALLOT_RESPONSES_MARKER);
  return marker === -1 ? -1 : marker + 1;
}

/**
 * Reads the candidate names from the Responses header row (columns E+), first
 * reconciling that header against the Candidates table (the primary source of
 * candidate identity) so a candidate typed only into the Candidates table — e.g.
 * pre-populated ahead of any votes — is always included.
 *
 * @param {Sheet} sheet
 * @return {Array<string>}
 */
function readBallotCandidates_(sheet) {
  _reconcileCandidatesWithResponses_(sheet);
  var headerRow = _getResponsesHeaderRow_(sheet);
  if (headerRow === -1) return [];
  var lastCol = sheet.getLastColumn();
  if (lastCol < _BALLOT_FIRST_CANDIDATE_COL) return [];
  var values = sheet.getRange(headerRow, _BALLOT_FIRST_CANDIDATE_COL, 1, lastCol - _BALLOT_FIRST_CANDIDATE_COL + 1).getValues()[0];
  return values.map(function (v) { return String(v || '').trim(); }).filter(function (v) { return v; });
}

/**
 * Ensures every candidate listed in the Candidates table (the primary source of
 * candidate identity — see file header) has a matching column in the Responses
 * header, appending one for each that's missing. Never reorders or removes existing
 * Responses columns, so ranks already recorded against them stay aligned to their
 * column; a Candidates-table entry with no column yet is always appended at the end,
 * which also matches how addBallotCandidate_ grows the two sections in lockstep, so
 * position alignment between the two sections is preserved either way.
 *
 * @param {Sheet} sheet
 */
function _reconcileCandidatesWithResponses_(sheet) {
  var plan = _planCandidateReconcile_(sheet);
  if (!plan) return;
  _withBallotLock_(function () {
    // Re-plan under the lock: another request may have appended the same columns meanwhile.
    var current = _planCandidateReconcile_(sheet);
    if (current) _applyCandidateReconcile_(sheet, current);
  });
}

/**
 * Read-only half of _reconcileCandidatesWithResponses_: works out what (if anything) the
 * Responses header needs.
 *
 * @param {Sheet} sheet
 * @return {{headerRow:number, existingNames:Array<string>, compacted:Array<string>, missingNames:Array<string>}|null}
 *   null when the header already matches the Candidates table.
 */
function _planCandidateReconcile_(sheet) {
  var candidateRows = readBallotCandidateDetails_(sheet);
  if (!candidateRows.length) return null;
  var headerRow = _getResponsesHeaderRow_(sheet);
  if (headerRow === -1) return null;

  var lastCol = sheet.getLastColumn();
  var existingNames = lastCol < _BALLOT_FIRST_CANDIDATE_COL ? [] :
    sheet.getRange(headerRow, _BALLOT_FIRST_CANDIDATE_COL, 1, lastCol - _BALLOT_FIRST_CANDIDATE_COL + 1)
      .getValues()[0].map(function (v) { return String(v || '').trim(); });

  // Sanity check: a blank cell before a non-blank one means an earlier write left a
  // gap in the Responses header — e.g. by computing the append column from
  // sheet.getLastColumn(), which can be pushed right of the header's own last
  // candidate column by a wider Candidates table (its Link URL column included).
  // Compact left to realign header names with the vote columns beneath them, which
  // are always written starting at _BALLOT_FIRST_CANDIDATE_COL regardless of the
  // header (see submitBallotResponse_) — so the header, not the data, is what drifted.
  var compacted = existingNames.filter(function (n) { return n; });

  var existingSet = {};
  compacted.forEach(function (n) { existingSet[n] = true; });

  var missingNames = candidateRows
    .map(function (it) { return it.name; })
    .filter(function (n) { return n && !existingSet[n]; });
  if (compacted.length === existingNames.length && !missingNames.length) return null;
  return { headerRow: headerRow, existingNames: existingNames, compacted: compacted, missingNames: missingNames };
}

/**
 * Write half of _reconcileCandidatesWithResponses_; caller holds the ballot lock.
 *
 * @param {Sheet} sheet
 * @param {{headerRow:number, existingNames:Array<string>, compacted:Array<string>, missingNames:Array<string>}} plan
 */
function _applyCandidateReconcile_(sheet, plan) {
  var headerRow = plan.headerRow;
  var existingNames = plan.existingNames;
  var compacted = plan.compacted;
  var missingNames = plan.missingNames;
  if (compacted.length !== existingNames.length) {
    Logger.log('Ballot "' + sheet.getName() + '": correcting ' + (existingNames.length - compacted.length) +
      ' blank gap column(s) in the Responses header.');
    sheet.getRange(headerRow, _BALLOT_FIRST_CANDIDATE_COL, 1, existingNames.length).clearContent();
    if (compacted.length) {
      sheet.getRange(headerRow, _BALLOT_FIRST_CANDIDATE_COL, 1, compacted.length).setValues([compacted]);
    }
    existingNames = compacted;
  }
  if (!missingNames.length) return;

  // Computed from the header's own (now-compacted) candidate count, not
  // sheet.getLastColumn() — see the compaction comment above for why that's unsafe.
  var nextCol = _BALLOT_FIRST_CANDIDATE_COL + existingNames.length;
  sheet.getRange(headerRow, nextCol, 1, missingNames.length).setValues([missingNames]);
}

/**
 * Reads all response rows below the Responses header.
 *
 * @param {Sheet} sheet
 * @return {Array<{row:number, date:*, name:string, weight:number, comment:string, ranks:Array}>}
 */
function readBallotResponseRows_(sheet) {
  var headerRow = _getResponsesHeaderRow_(sheet);
  var lastRow = sheet.getLastRow();
  if (headerRow === -1 || lastRow < headerRow + 1) return [];
  var candidates = readBallotCandidates_(sheet);
  var numCols = _BALLOT_RESPONSE_FIXED_COLS.length + candidates.length;
  var values = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, numCols).getValues();
  var out = [];
  values.forEach(function (row, i) {
    var name = String(row[1] || '').trim();
    if (!name) return;
    out.push({
      row: headerRow + 1 + i,
      date: row[0],
      name: name,
      weight: Number(row[2]) || 1,
      comment: String(row[3] || ''),
      ranks: row.slice(_BALLOT_RESPONSE_FIXED_COLS.length)
    });
  });
  return out;
}

/**
 * Collapses response rows down to one per respondent (case-insensitive name
 * match), keeping only the last-written row for each — sheet rows are in
 * top-to-bottom submission order, and a re-submission is meant to supersede
 * that respondent's earlier ranking entirely (see submitBallotResponse_, which
 * normally overwrites the existing row in place via _findBallotResponseRowForName_;
 * this dedup is what makes counts/analysis correct even if duplicate rows for
 * the same name exist anyway, e.g. from a hand-edited sheet).
 *
 * @param {Array<{name:string}>} rows as returned by readBallotResponseRows_
 * @return {Array} one entry per distinct respondent, in first-seen order
 */
function _latestResponseByRespondent_(rows) {
  var latestByLowerName = {};
  var order = [];
  rows.forEach(function (r) {
    var key = r.name.toLowerCase();
    if (!latestByLowerName.hasOwnProperty(key)) order.push(key);
    latestByLowerName[key] = r;
  });
  return order.map(function (key) { return latestByLowerName[key]; });
}

/**
 * @param {Sheet} sheet
 * @return {number} count of distinct respondents (case-insensitive name), i.e.
 *   how many people voted — a respondent who re-voted still counts once.
 */
function countUniqueBallotRespondents_(sheet) {
  return _latestResponseByRespondent_(readBallotResponseRows_(sheet)).length;
}

/**
 * Finds respondents whose latest submitted ranking predates one or more of the
 * ballot's current candidates, so they may want to come back and rank the new
 * addition(s). Detected via a blank rank cell: _reconcileCandidatesWithResponses_
 * only ever appends new candidate columns and never backfills existing rows, so
 * a response row written before a candidate existed has a genuinely empty cell
 * under that candidate's column — there's no other way a rank cell goes blank,
 * since submitBallotResponse_ always writes every candidate known at submit time.
 *
 * @param {Sheet} sheet
 * @return {Array<string>} respondent names (as last submitted), one per stale respondent
 */
function findRespondentsWithNewCandidates_(sheet) {
  var rows = _latestResponseByRespondent_(readBallotResponseRows_(sheet));
  return rows
    .filter(function (r) { return r.ranks.some(function (v) { return v === '' || v === null || v === undefined; }); })
    .map(function (r) { return r.name; });
}

/**
 * Finds the last response row for a respondent (case-insensitive), so a
 * re-submission overwrites their existing row instead of adding a duplicate.
 *
 * @param {Sheet} sheet
 * @param {string} name
 * @return {number} 1-based row, or -1 if not found.
 */
function _findBallotResponseRowForName_(sheet, name) {
  var target = String(name).trim().toLowerCase();
  var rows = readBallotResponseRows_(sheet);
  for (var i = rows.length - 1; i >= 0; i--) {
    if (rows[i].name.toLowerCase() === target) return rows[i].row;
  }
  return -1;
}

/**
 * Appends a new candidate column to the Responses header row, and a matching row
 * to the Candidates table (position-aligned — this candidate's Candidates-table row
 * ends up at the same index as its new column). Caller must check Accept-New first.
 *
 * @param {Sheet} sheet
 * @param {string} phrase
 * @param {string=} details admin-only note for this candidate (respondents never pass this).
 * @param {string=} linkText admin-only link label shown alongside details (respondents never pass this).
 * @param {string=} linkUrl admin-only link target, opened in a new tab (respondents never pass this).
 */
function addBallotCandidate_(sheet, phrase, details, linkText, linkUrl) {
  _assertNotMarkerName_(phrase);
  var headerRow = _getResponsesHeaderRow_(sheet);
  if (headerRow === -1) throw new Error('Ballot has no Responses section.');
  var existingCandidates = readBallotCandidates_(sheet); // BEFORE the new column exists
  // Computed from the Responses header's own candidate-column count (not
  // sheet.getLastColumn()) — the Candidates table sits in earlier columns and can be
  // wider than the Responses section (e.g. its Link URL column), which would otherwise
  // throw off where the next candidate column belongs.
  var nextCol = _BALLOT_FIRST_CANDIDATE_COL + existingCandidates.length;
  sheet.getRange(headerRow, nextCol).setValue(phrase);

  var candidateRows = readBallotCandidateDetails_(sheet);
  // Backfill any candidates that predate the Candidates table (e.g. a ballot created, or
  // a candidate added, before this feature existed) with blank details, so positions
  // never drift between the Candidates table and the Responses header's candidate columns.
  while (candidateRows.length < existingCandidates.length) {
    candidateRows.push({ name: existingCandidates[candidateRows.length], details: '', linkText: '', linkUrl: '' });
  }
  candidateRows.push({ name: phrase, details: details || '', linkText: linkText || '', linkUrl: linkUrl || '' });
  writeBallotCandidateDetails_(sheet, candidateRows);
}

/**
 * RPC-facing: returns everything the ballot page needs to render for a
 * given respondent (or generic branding/ordering if name is blank/unknown).
 *
 * @param {string} id
 * @param {string} name
 * @return {Object|null} null if the ballot id does not exist.
 */
function getBallotForRespondent_(id, name) {
  name = String(name || '').trim();

  if (!name) {
    // No respondent name means nothing below depends on the Responses section (no
    // ranking to look up), so a cache hit lets this return without ever opening the
    // spreadsheet — see BallotCache.js. This is the common landing-page-load path.
    var cached = getCachedBallotData_(id);
    if (!cached) return null;
    return _buildRespondentBranding_(id, cached.config, cached.candidates);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findBallotSheet_(ss, id);
  if (!sheet) return null;

  // A name means this respondent's saved ranking (Responses section) is needed, which
  // always requires the live sheet — read config/candidates fresh here too, rather than
  // from cache, since readBallotCandidates_'s reconciliation against the Responses
  // header must run against the real sheet for the rank-column alignment below to be
  // trustworthy. refreshBallotCache_ also backfills any candidate that predates the
  // Candidates table, and re-warms the cache from this same read for the next
  // (nameless) load. `candidates` is name-only, in the same Responses-header-aligned
  // order, for the rank lookup further down.
  var freshCache = refreshBallotCache_(id, sheet);
  var candidates = freshCache.candidates.map(function (it) { return it.name; });

  var result = _buildRespondentBranding_(id, freshCache.config, freshCache.candidates);

  var rows = readBallotResponseRows_(sheet);
  var target = name.toLowerCase();
  var existing = null;
  for (var i = rows.length - 1; i >= 0; i--) {
    if (rows[i].name.toLowerCase() === target) { existing = rows[i]; break; }
  }
  if (!existing) return result;

  var withRank = candidates.map(function (c, idx) {
    var rank = Number(existing.ranks[idx]);
    return { candidate: c, rank: isNaN(rank) ? Infinity : rank };
  });
  withRank.sort(function (a, b) { return a.rank - b.rank; });

  result.candidates = withRank.map(function (r) { return r.candidate; });
  result.comment = existing.comment;
  return result;
}

/**
 * Builds the branding/candidate portion of getBallotForRespondent_'s result (everything
 * except a specific respondent's saved ranking/comment) from already-loaded config +
 * candidate-detail rows, whether those came from the cache or a fresh sheet read.
 *
 * @param {string} id
 * @param {Object} config as returned by readBallotConfig_
 * @param {Array<{name:string, details:string, linkText:string, linkUrl:string}>} candidateRows
 * @return {Object}
 */
function _buildRespondentBranding_(id, config, candidateRows) {
  var candidates = candidateRows.map(function (it) { return it.name; });

  // Keyed by name (not position) so it stays correct however `candidates` above ends up
  // reordered per-respondent to match their saved ranking — a candidate with no
  // Candidates-table entry (e.g. added by a respondent, who can't set details) simply
  // has no key here, which the client treats as "no details to show," not an error.
  // NOTE: the RPC field is still named `itemDetails` — it's an established part of the
  // client contract (webBallotPage.html reads config.itemDetails), unrelated to the
  // Items->Candidates section rename. Each entry is {details, linkText, linkUrl} rather
  // than a bare string so the client can render the optional link alongside it.
  var itemDetails = {};
  candidateRows.forEach(function (it) {
    if (it.details || it.linkUrl) {
      itemDetails[it.name] = { details: it.details || '', linkText: it.linkText || '', linkUrl: it.linkUrl || '' };
    }
  });

  return {
    id: id,
    title: config.Title || '',
    description: config.Description || '',
    instructions: config.Instructions || '',
    footer: config.Footer || '',
    contact: config.Contact || '',
    acceptNew: _ballotAcceptsNew_(config),
    addInstructions: config['Add-Instructions'] || '',
    candidates: candidates,
    itemDetails: itemDetails,
    comment: '',
    appVersion: (typeof APP_VERSION !== 'undefined' && APP_VERSION) || '',
    appDeployTarget: (typeof APP_DEPLOY_TARGET !== 'undefined' && APP_DEPLOY_TARGET) || ''
  };
}

/**
 * RPC-facing: appends a new candidate to a ballot (if Accept-New allows it).
 * Respondents can set a Details note too (previously admin-only) — it's the same
 * underlying field, just now writable from the "Add New" panel on the ballot page.
 *
 * @param {string} id
 * @param {string} phrase
 * @param {string=} details
 * @return {{candidate:string}}
 */
function addBallotCandidateForId_(id, phrase, details) {
  phrase = String(phrase || '').trim();
  if (!phrase) throw new Error('Item cannot be empty.');
  _assertNotMarkerName_(phrase);

  return _withBallotLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findBallotSheet_(ss, id);
    if (!sheet) throw new Error('No ballot found for id "' + id + '".');
    var config = readBallotConfig_(sheet);
    if (!_ballotAcceptsNew_(config)) throw new Error('This ballot is not accepting new items.');
    addBallotCandidate_(sheet, phrase, details);
    refreshBallotCache_(id, sheet); // keep the cache (BallotCache.js) in sync with this write
    return { candidate: phrase };
  });
}

/**
 * RPC-facing: saves a respondent's ranking (and optional comment). Reuses
 * their existing row and weight if one exists, else appends a new row with
 * weight 1.
 *
 * @param {string} id
 * @param {string} name
 * @param {Array<string>} orderedCandidates preference order, most preferred first.
 * @param {string=} comment
 * @return {{ok:boolean}}
 */
function submitBallotResponse_(id, name, orderedCandidates, comment) {
  name = String(name || '').trim();
  if (!name) throw new Error('Name is required.');
  if (!orderedCandidates || !orderedCandidates.length) throw new Error('Ranking cannot be empty.');

  return _withBallotLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findBallotSheet_(ss, id);
    if (!sheet) throw new Error('No ballot found for id "' + id + '".');

    var candidates = readBallotCandidates_(sheet);
    var rankByCandidate = {};
    orderedCandidates.forEach(function (c, idx) {
      rankByCandidate[String(c).trim()] = idx + 1;
    });

    var row = _findBallotResponseRowForName_(sheet, name);
    var weight = 1;
    if (row === -1) {
      row = sheet.getLastRow() + 1;
    } else {
      weight = Number(sheet.getRange(row, 3).getValue()) || 1;
    }

    var ranks = candidates.map(function (c) {
      return rankByCandidate.hasOwnProperty(c) ? rankByCandidate[c] : '';
    });
    sheet.getRange(row, 1, 1, _BALLOT_RESPONSE_FIXED_COLS.length + candidates.length).setValues([
      [new Date(), name, weight, String(comment || '').trim()].concat(ranks)
    ]);

    return { ok: true };
  });
}

/**
 * Inserts or deletes rows so the Results section has exactly rows.length
 * rows available, then writes rows into it. Leaves the Candidates/Responses sections
 * (which sit immediately below) untouched — the insert/delete happens strictly
 * between the Results and Candidates markers, so everything below shifts as one block.
 * The region is cleared in full first and output starts at column B (column A is
 * markers-only), so older column-A output self-corrects on the next run.
 *
 * @param {Sheet} sheet
 * @param {Array<Array>} rows
 */
function writeBallotResults_(sheet, rows) {
  _ensureCandidatesSection_(sheet);
  var resultsRow = _findMarkerRow_(sheet, _BALLOT_RESULTS_MARKER);
  var candidatesRow = _findMarkerRow_(sheet, _BALLOT_CANDIDATES_MARKER);
  if (resultsRow === -1 || candidatesRow === -1) {
    throw new Error('Ballot sheet is missing its Results/Candidates markers.');
  }
  // Clear the whole region (every column, col A included) first: column A is
  // markers-only, so output from the older layout that wrote names there must not
  // survive the rewrite; the resize then inserts/deletes blank rows only.
  var regionRows = candidatesRow - resultsRow - 1;
  if (regionRows > 0) {
    sheet.getRange(resultsRow + 1, 1, regionRows, Math.max(sheet.getLastColumn(), 1)).clearContent();
  }
  _resizeSectionRows_(sheet, resultsRow, candidatesRow, rows, _BALLOT_CANDIDATES_FIRST_COL);
}

/**
 * Reads the Candidates table (Name/Details per row), position-aligned with the
 * Responses header's candidate columns — candidate i corresponds to candidate column i.
 * This table is the primary source of candidate identity (see file header); it may
 * list candidates that don't have a Responses column yet.
 *
 * @param {Sheet} sheet
 * @return {Array<{name:string, details:string, linkText:string, linkUrl:string}>}
 */
function readBallotCandidateDetails_(sheet) {
  _ensureCandidatesSection_(sheet);
  _migrateCandidatesColumnsIfNeeded_(sheet);
  _ensureCandidatesHeaderColumns_(sheet);
  _ensureResponsesSection_(sheet);
  var candidatesRow = _findMarkerRow_(sheet, _BALLOT_CANDIDATES_MARKER);
  var responsesRow = _findMarkerRow_(sheet, _BALLOT_RESPONSES_MARKER);
  if (candidatesRow === -1 || responsesRow === -1) {
    _logBallotEvent_('ballot.parse.markerMissing', { sheet: sheet.getName(), candidatesRow: candidatesRow, responsesRow: responsesRow });
    return [];
  }
  var candidatesHeaderRow = candidatesRow + 1;
  var count = responsesRow - candidatesHeaderRow - 1;
  if (count <= 0) return [];
  // Column A is marker-only (see file header) and never read as a candidate — but a name
  // typed there is an easy mistake, so say so rather than let it vanish silently.
  var colARows = [];
  sheet.getRange(candidatesHeaderRow + 1, 1, count, 1).getValues().forEach(function (r, i) {
    if (String(r[0] || '').trim()) colARows.push(candidatesHeaderRow + 1 + i);
  });
  if (colARows.length) {
    _logBallotEvent_('ballot.parse.candidateInColumnA', { sheet: sheet.getName(), rows: colARows });
  }
  var values = sheet.getRange(candidatesHeaderRow + 1, _BALLOT_CANDIDATES_FIRST_COL, count, _BALLOT_CANDIDATES_HEADER.length).getValues();
  // A row with no name (e.g. a stray blank row someone left in the sheet) isn't a
  // candidate — skip it rather than returning a phantom entry that would throw off
  // the position alignment with the Responses header's candidate columns.
  return values
    .map(function (r) {
      return { name: String(r[0] || '').trim(), details: String(r[1] || ''), linkText: String(r[2] || ''), linkUrl: String(r[3] || '') };
    })
    .filter(function (it) { return it.name; });
}

/**
 * Resizes + rewrites the Candidates table to exactly match `candidateRows` (array of
 * {name, details, linkText, linkUrl}, one per candidate column, in the same order).
 * Leaves the Responses section untouched — the insert/delete happens strictly between
 * the Candidates header and the Responses marker.
 *
 * @param {Sheet} sheet
 * @param {Array<{name:string, details:string, linkText:string, linkUrl:string}>} candidateRows
 */
function writeBallotCandidateDetails_(sheet, candidateRows) {
  _ensureCandidatesSection_(sheet);
  _migrateCandidatesColumnsIfNeeded_(sheet);
  _ensureCandidatesHeaderColumns_(sheet);
  var candidatesRow = _findMarkerRow_(sheet, _BALLOT_CANDIDATES_MARKER);
  var responsesRow = _findMarkerRow_(sheet, _BALLOT_RESPONSES_MARKER);
  if (candidatesRow === -1 || responsesRow === -1) {
    throw new Error('Ballot sheet is missing its Candidates/Responses markers.');
  }
  var rows = candidateRows.map(function (it) { return [it.name || '', it.details || '', it.linkText || '', it.linkUrl || '']; });
  _resizeSectionRows_(sheet, candidatesRow + 1, responsesRow, rows, _BALLOT_CANDIDATES_FIRST_COL);
}

/**
 * Admin-facing: overwrites every candidate's name + details in one pass, given a
 * full ordered list matching the ballot's current candidate positions (position,
 * not name, is the join key — renaming a candidate never touches response data,
 * since ranks are keyed by column index). Throws if the count doesn't match the
 * ballot's current candidate count (stale form — caller should reload and retry).
 *
 * @param {Sheet} sheet
 * @param {Array<{name:string, details:string, linkText:string, linkUrl:string}>} candidateRows
 */
function saveBallotCandidates_(sheet, candidateRows) {
  var candidates = readBallotCandidates_(sheet);
  if (candidateRows.length !== candidates.length) {
    throw new Error('Candidate count changed since the page was loaded — reload and try again.');
  }
  var names = candidateRows.map(function (it) {
    var name = String(it.name || '').trim();
    if (!name) throw new Error('Candidate name cannot be empty.');
    _assertNotMarkerName_(name);
    return name;
  });
  var headerRow = _getResponsesHeaderRow_(sheet);
  if (names.length) {
    sheet.getRange(headerRow, _BALLOT_FIRST_CANDIDATE_COL, 1, names.length).setValues([names]);
  }
  writeBallotCandidateDetails_(sheet, candidateRows.map(function (it, i) {
    return { name: names[i], details: String(it.details || ''), linkText: String(it.linkText || ''), linkUrl: String(it.linkUrl || '') };
  }));
  _highlightSectionMarkers_(sheet);
}

/**
 * RPC-facing: locates the ballot by id and applies saveBallotCandidates_ under the
 * script lock (mirrors submitBallotResponse_/addBallotCandidateForId_'s locking —
 * admin edits and respondent submissions/candidate-adds must not interleave).
 *
 * @param {string} id
 * @param {Array<{name:string, details:string, linkText:string, linkUrl:string}>} candidateRows
 */
function saveBallotCandidatesForId_(id, candidateRows) {
  return _withBallotLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findBallotSheet_(ss, id);
    if (!sheet) throw new Error('No ballot found for id "' + id + '".');
    saveBallotCandidates_(sheet, candidateRows);
    refreshBallotCache_(id, sheet); // keep the cache (BallotCache.js) in sync with this write
  });
}

/**
 * Admin-facing: appends a new candidate WITH a details note in one call (the
 * respondent-facing addBallotCandidateForId_ has no details param — Details is
 * admin-only). Locked the same way as addBallotCandidateForId_.
 *
 * @param {string} id
 * @param {string} name
 * @param {string=} details
 * @param {string=} linkText
 * @param {string=} linkUrl
 * @return {{candidate:string}}
 */
function addBallotCandidateForAdmin_(id, name, details, linkText, linkUrl) {
  name = String(name || '').trim();
  if (!name) throw new Error('Candidate name cannot be empty.');

  return _withBallotLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findBallotSheet_(ss, id);
    if (!sheet) throw new Error('No ballot found for id "' + id + '".');
    addBallotCandidate_(sheet, name, details, linkText, linkUrl);
    _highlightSectionMarkers_(sheet);
    refreshBallotCache_(id, sheet); // keep the cache (BallotCache.js) in sync with this write
    return { candidate: name };
  });
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    getBallotSheetName_: getBallotSheetName_,
    listBallotIds_: listBallotIds_,
    findBallotSheet_: findBallotSheet_,
    createBallotSheet_: createBallotSheet_,
    getOrCreateBallotSheet_: getOrCreateBallotSheet_,
    createNewBallot_: createNewBallot_,
    writeBallotConfig_: writeBallotConfig_,
    readBallotConfig_: readBallotConfig_,
    readBallotCandidates_: readBallotCandidates_,
    readBallotResponseRows_: readBallotResponseRows_,
    countUniqueBallotRespondents_: countUniqueBallotRespondents_,
    _latestResponseByRespondent_: _latestResponseByRespondent_,
    findRespondentsWithNewCandidates_: findRespondentsWithNewCandidates_,
    addBallotCandidate_: addBallotCandidate_,
    getBallotForRespondent_: getBallotForRespondent_,
    addBallotCandidateForId_: addBallotCandidateForId_,
    submitBallotResponse_: submitBallotResponse_,
    writeBallotResults_: writeBallotResults_,
    readBallotCandidateDetails_: readBallotCandidateDetails_,
    writeBallotCandidateDetails_: writeBallotCandidateDetails_,
    saveBallotCandidates_: saveBallotCandidates_,
    saveBallotCandidatesForId_: saveBallotCandidatesForId_,
    addBallotCandidateForAdmin_: addBallotCandidateForAdmin_,
    _highlightSectionMarkers_: _highlightSectionMarkers_,
    _ensureCandidatesSection_: _ensureCandidatesSection_,
    _ensureResponsesSection_: _ensureResponsesSection_,
    _migrateCandidatesColumnsIfNeeded_: _migrateCandidatesColumnsIfNeeded_,
    _ensureCandidatesHeaderColumns_: _ensureCandidatesHeaderColumns_,
    _withBallotLock_: _withBallotLock_,
    _findMarkerRow_: _findMarkerRow_
  };
}
