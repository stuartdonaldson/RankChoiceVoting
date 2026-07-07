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
 * user-entered data — and markers are bracket-decorated ("[Results]", not "Results")
 * as a second, independent guard. Both exist because _findMarkerRow_ does an
 * unscoped top-to-bottom scan of column A for exact marker text: if a candidate's
 * name (user-controlled, arbitrary) ever landed in column A, naming a candidate e.g.
 * "Responses" would shadow the real marker and corrupt every section boundary
 * computed from it (this was a real bug in an earlier version of this file, where
 * the Candidates table's Name column started at col A). Reserving col A for markers
 * only closes the hole entirely; the bracket decoration is defense in depth against
 * any future mistake (or a sheet hand-edited outside the app) reintroducing it.
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
// Legacy marker text recognized by _findMarkerRow_ purely to migrate older sheets in
// place — the moment one is found, the cell is rewritten to the decorated/current form,
// so every later scan matches directly. This covers both the pre-bracket-decoration
// bare text ("Results", "Responses") and the section's former name ("[Items]"/"Items",
// before it was renamed to "[Candidates]" to reflect that it's the primary candidate list).
var _BALLOT_LEGACY_MARKER_TEXT_ = {
  '[Results]': ['Results'],
  '[Candidates]': ['Candidates', '[Items]', 'Items'],
  '[Responses]': ['Responses']
};
var _BALLOT_CONFIG_ROWS = ['Title', 'Description', 'Instructions', 'Footer', 'Contact', 'Accept-New', 'Add-Instructions', 'Admin-Only-Notes'];
// Legacy config key recognized by readBallotConfig_ purely to migrate older sheets in
// place, same pattern as _BALLOT_LEGACY_MARKER_TEXT_ above.
var _BALLOT_LEGACY_CONFIG_KEYS_ = { 'Admin-Only-Notes': ['Info'] };
// Fixed Responses columns before the candidate columns begin (col 5 / E).
var _BALLOT_RESPONSE_FIXED_COLS = ['Date', 'Name', 'Weight', 'Comment'];
var _BALLOT_FIRST_CANDIDATE_COL = _BALLOT_RESPONSE_FIXED_COLS.length + 1; // 5
var _BALLOT_LOCK_TIMEOUT_MS = 30000;
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
 * empty Results/Responses skeleton. Placeholder values start with "[TODO"
 * so they are easy to spot and grep for.
 *
 * @param {Spreadsheet} ss
 * @param {string} id
 * @return {Sheet}
 */
function createBallotSheet_(ss, id) {
  var sheet = ss.insertSheet(getBallotSheetName_(id));
  sheet.getRange(1, 1, 8, 2).setValues([
    ['Title', '[TODO: ballot title shown to respondents]'],
    ['Description', '[TODO: intro text shown before respondents enter their name]'],
    ['Instructions', '[TODO: instructions shown above the ranking list on the ballot page]'],
    ['Footer', '[TODO: footer text, e.g. deadline or sponsoring group]'],
    ['Contact', '[TODO: contact name/email for questions]'],
    ['Accept-New', 'TRUE'],
    ['Add-Instructions', '[TODO: instructions shown above the "+ Add New" button, if enabled]'],
    ['Admin-Only-Notes', '']
  ]);
  sheet.getRange(10, 1).setValue(_BALLOT_RESULTS_MARKER);
  sheet.getRange(11, 1).setValue(_BALLOT_CANDIDATES_MARKER);
  sheet.getRange(12, _BALLOT_CANDIDATES_FIRST_COL, 1, _BALLOT_CANDIDATES_HEADER.length).setValues([_BALLOT_CANDIDATES_HEADER]);
  sheet.getRange(13, 1).setValue(_BALLOT_RESPONSES_MARKER);
  sheet.getRange(14, 1, 1, _BALLOT_RESPONSE_FIXED_COLS.length).setValues([_BALLOT_RESPONSE_FIXED_COLS]);
  sheet.setFrozenRows(14);
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
  }
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
  var responsesRow = _findMarkerRow_(sheet, _BALLOT_RESPONSES_MARKER);
  if (responsesRow === -1) return;
  sheet.insertRowsBefore(responsesRow, 2);
  sheet.getRange(responsesRow, 1).setValue(_BALLOT_CANDIDATES_MARKER);
  sheet.getRange(responsesRow + 1, _BALLOT_CANDIDATES_FIRST_COL, 1, _BALLOT_CANDIDATES_HEADER.length).setValues([_BALLOT_CANDIDATES_HEADER]);
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

  var responsesRow = _findMarkerRow_(sheet, _BALLOT_RESPONSES_MARKER);
  if (responsesRow === -1) return;
  var count = responsesRow - candidatesHeaderRow - 1;
  var oldValues = count > 0 ? sheet.getRange(candidatesHeaderRow + 1, 1, count, 2).getValues() : [];

  sheet.getRange(candidatesHeaderRow, 1, count + 1, 1).clearContent(); // vacate col A (header + data rows)
  sheet.getRange(candidatesHeaderRow, _BALLOT_CANDIDATES_FIRST_COL, 1, _BALLOT_CANDIDATES_HEADER.length).setValues([_BALLOT_CANDIDATES_HEADER]);
  if (count > 0) {
    sheet.getRange(candidatesHeaderRow + 1, _BALLOT_CANDIDATES_FIRST_COL, count, 2).setValues(oldValues);
  }
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
  var firstMissingCol = _BALLOT_CANDIDATES_FIRST_COL + (_BALLOT_CANDIDATES_HEADER.length - missing.length);
  sheet.getRange(headerRow, firstMissingCol, 1, missing.length).setValues([missing]);
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
 * @param {number=} startCol defaults to col A (1) — Candidates-section callers pass
 *   _BALLOT_CANDIDATES_FIRST_COL so candidate data never lands in the marker-reserved column A.
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
 * Scans column A for a row whose value exactly matches marker.
 *
 * @param {Sheet} sheet
 * @param {string} marker
 * @return {number} 1-based row, or -1 if not found.
 */
function _findMarkerRow_(sheet, marker) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return -1;
  var values = sheet.getRange(1, 1, lastRow, 1).getValues();
  var legacyTexts = _BALLOT_LEGACY_MARKER_TEXT_[marker] || [];
  for (var i = 0; i < values.length; i++) {
    var cell = String(values[i][0] || '').trim();
    if (cell === marker) return i + 1;
    if (legacyTexts.indexOf(cell) !== -1) {
      sheet.getRange(i + 1, 1).setValue(marker); // one-time upgrade to the current decorated form
      return i + 1;
    }
  }
  return -1;
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
        sheet.getRange(i + 1, 1).setValue(currentKey); // one-time upgrade to the current key name
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
  var candidateRows = readBallotCandidateDetails_(sheet);
  if (!candidateRows.length) return;
  var headerRow = _getResponsesHeaderRow_(sheet);
  if (headerRow === -1) return;

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
  if (compacted.length !== existingNames.length) {
    Logger.log('Ballot "' + sheet.getName() + '": correcting ' + (existingNames.length - compacted.length) +
      ' blank gap column(s) in the Responses header.');
    sheet.getRange(headerRow, _BALLOT_FIRST_CANDIDATE_COL, 1, existingNames.length).clearContent();
    if (compacted.length) {
      sheet.getRange(headerRow, _BALLOT_FIRST_CANDIDATE_COL, 1, compacted.length).setValues([compacted]);
    }
    existingNames = compacted;
  }

  var existingSet = {};
  existingNames.forEach(function (n) { if (n) existingSet[n] = true; });

  var missingNames = candidateRows
    .map(function (it) { return it.name; })
    .filter(function (n) { return n && !existingSet[n]; });
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findBallotSheet_(ss, id);
  if (!sheet) return null;

  var config = readBallotConfig_(sheet);
  var candidates = readBallotCandidates_(sheet);
  name = String(name || '').trim();

  // Keyed by name (not position) so it stays correct even though `candidates` below
  // gets reordered per-respondent to match their saved ranking — a candidate with no
  // Candidates-table entry (e.g. added by a respondent, who can't set details) simply
  // has no key here, which the client treats as "no details to show," not an error.
  // NOTE: the RPC field is still named `itemDetails` — it's an established part of
  // the client contract (webBallotPage.html reads config.itemDetails), unrelated to
  // the Items->Candidates section rename. Each entry is {details, linkText, linkUrl}
  // rather than a bare string so the client can render the optional link alongside it.
  var itemDetails = {};
  readBallotCandidateDetails_(sheet).forEach(function (it) {
    if (it.details || it.linkUrl) {
      itemDetails[it.name] = { details: it.details || '', linkText: it.linkText || '', linkUrl: it.linkUrl || '' };
    }
  });

  var result = {
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
  if (!name) return result;

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

  var lock = LockService.getScriptLock();
  lock.waitLock(_BALLOT_LOCK_TIMEOUT_MS);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findBallotSheet_(ss, id);
    if (!sheet) throw new Error('No ballot found for id "' + id + '".');
    var config = readBallotConfig_(sheet);
    if (!_ballotAcceptsNew_(config)) throw new Error('This ballot is not accepting new items.');
    addBallotCandidate_(sheet, phrase, details);
    return { candidate: phrase };
  } finally {
    lock.releaseLock();
  }
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

  var lock = LockService.getScriptLock();
  lock.waitLock(_BALLOT_LOCK_TIMEOUT_MS);
  try {
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
  } finally {
    lock.releaseLock();
  }
}

/**
 * Inserts or deletes rows so the Results section has exactly rows.length
 * rows available, then writes rows into it. Leaves the Candidates/Responses sections
 * (which sit immediately below) untouched — the insert/delete happens strictly
 * between the Results and Candidates markers, so everything below shifts as one block.
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
  _resizeSectionRows_(sheet, resultsRow, candidatesRow, rows);
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
  var candidatesRow = _findMarkerRow_(sheet, _BALLOT_CANDIDATES_MARKER);
  var responsesRow = _findMarkerRow_(sheet, _BALLOT_RESPONSES_MARKER);
  if (candidatesRow === -1 || responsesRow === -1) return [];
  var candidatesHeaderRow = candidatesRow + 1;
  var count = responsesRow - candidatesHeaderRow - 1;
  if (count <= 0) return [];
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
  var lock = LockService.getScriptLock();
  lock.waitLock(_BALLOT_LOCK_TIMEOUT_MS);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findBallotSheet_(ss, id);
    if (!sheet) throw new Error('No ballot found for id "' + id + '".');
    saveBallotCandidates_(sheet, candidateRows);
  } finally {
    lock.releaseLock();
  }
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

  var lock = LockService.getScriptLock();
  lock.waitLock(_BALLOT_LOCK_TIMEOUT_MS);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findBallotSheet_(ss, id);
    if (!sheet) throw new Error('No ballot found for id "' + id + '".');
    addBallotCandidate_(sheet, name, details, linkText, linkUrl);
    _highlightSectionMarkers_(sheet);
    return { candidate: name };
  } finally {
    lock.releaseLock();
  }
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
    _migrateCandidatesColumnsIfNeeded_: _migrateCandidatesColumnsIfNeeded_,
    _ensureCandidatesHeaderColumns_: _ensureCandidatesHeaderColumns_
  };
}
