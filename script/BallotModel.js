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
 *   Row M+2+   One row per candidate — name/details start at col B
 *   Row N      "[Responses]" marker (col A)
 *   Row N+1    Responses header: Date | Name | Weight | Comment | <candidate name> | ... — one rank
 *              column per candidate, matched to the Candidates table BY NAME (any order)
 *   Row N+2+   One row per respondent
 *
 * That is the order createBallotSheet_ writes, but nothing depends on it: the three
 * sections may appear in any order below the config rows. _readBallotLayout_ finds every
 * marker in one column-A read, and each section runs from its marker to the row before the
 * next marker of any kind (or to the last row). Config is everything above the first
 * marker. Readers skip blank rows inside a section rather than stopping at them, and every
 * write resizes its own section in place, so the sections around it only shift.
 *
 * Column A is reserved EXCLUSIVELY for section markers/config keys — never for
 * user-entered or app-generated data — and markers are bracket-decorated ("[Results]",
 * not "Results") as a second, independent guard. Both exist because _readBallotLayout_
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
 * "Admin-Only-Notes" (config key; formerly "Info" — read under the new name, and renamed
 * on the sheet by the repair path) is free-text
 * notes about the ballot (e.g. purpose, audience, scheduling context) — shown in
 * the admin list as a per-row disclosure, never exposed to respondents (see
 * getBallotForRespondent_, which omits it).
 *
 * The Candidates table is the PRIMARY identification of candidates — it is the
 * source of truth for who is a candidate, including ones pre-populated there ahead
 * of any votes. Each candidate's ranks live in the Responses column whose header is its
 * name (trim, case-insensitive); columns are found by label, never by position, so a
 * hand-sorted or hand-trimmed Candidates table cannot move a rank to another candidate.
 * A Responses rank column matching no candidate is an ORPHAN: kept, shown on the admin
 * edit page, never counted.
 *
 * READS NEVER WRITE (rcballot-8i5.3). Every reader is one getDataRange read parsed by
 * _parseBallot_ (see the Snapshot section). Repairs — lost markers, the legacy "Info" key,
 * the old column-A Candidates layout, a missing/incomplete Candidates header, labels left
 * on the [Responses] marker row, an empty Candidates table on a pre-Candidates ballot,
 * candidates without a rank column — run only through _repairBallotSheet_, called on admin
 * edit-page load (getAdminEditData) and by the write paths, under the ballot lock.
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

  var layout = _readBallotLayout_(sheet);
  var resultsRow = layout.markers.results;
  var candidatesRow = layout.markers.candidates;
  var responsesRow = layout.markers.responses;

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
 * Self-repair for a ballot sheet missing any of its three section markers, run at the start
 * of every section read/write entry. Without a marker, that section's rows count as part of
 * the section above — and if that is Results, the next analysis run deletes them. Repairs,
 * in this order, each logged as ballot.repair.<section>Section {sheet,row}:
 *   - [Candidates] missing: an empty marker + header is inserted directly above
 *     [Responses], or appended at the end when [Responses] is missing too (a ballot
 *     created before the Candidates section existed);
 *   - [Responses] missing: see _repairResponsesSectionLocked_ (appended only when that
 *     is provably safe, else logged and left alone);
 *   - [Results] missing: a marker is inserted directly above the first present marker,
 *     which keeps the canonical order.
 * Existing rows are never reassigned to another section: every repair inserts new
 * marker/header rows (or appends below everything). A healthy sheet costs one layout read
 * and takes no lock; otherwise the lock is taken and the layout re-checked under it.
 *
 * @param {Sheet} sheet
 * @param {Object=} layout from _readBallotLayout_, if the caller already has a current one.
 * @return {Object} the current layout (re-read after any repair). A marker can still be
 *   missing when its repair was refused.
 */
function _ensureBallotSections_(sheet, layout) {
  layout = layout || _readBallotLayout_(sheet);
  if (_ballotMarkersComplete_(layout)) return layout;
  return _withBallotLock_(function () {
    var current = _readBallotLayout_(sheet);
    if (current.markers.candidates === -1) {
      var row;
      if (current.markers.responses !== -1) {
        row = current.markers.responses;
        sheet.insertRowsBefore(row, 2);
      } else {
        row = current.lastRow + 1;
      }
      sheet.getRange(row, 1).setValue(_BALLOT_CANDIDATES_MARKER);
      sheet.getRange(row + 1, _BALLOT_CANDIDATES_FIRST_COL, 1, _BALLOT_CANDIDATES_HEADER.length).setValues([_BALLOT_CANDIDATES_HEADER]);
      _logBallotEvent_('ballot.repair.candidatesSection', { sheet: sheet.getName(), row: row });
      current = _readBallotLayout_(sheet);
    }
    if (current.markers.responses === -1 && _repairResponsesSectionLocked_(sheet, current)) {
      current = _readBallotLayout_(sheet);
    }
    if (current.markers.results === -1 && current.firstMarkerRow !== -1) {
      var resultsRow = current.firstMarkerRow;
      sheet.insertRowsBefore(resultsRow, 1);
      sheet.getRange(resultsRow, 1).setValue(_BALLOT_RESULTS_MARKER);
      _logBallotEvent_('ballot.repair.resultsSection', { sheet: sheet.getName(), row: resultsRow });
      current = _readBallotLayout_(sheet);
    }
    return current;
  });
}

/**
 * @param {Object} layout from _readBallotLayout_
 * @return {boolean} true when all three section markers are present.
 */
function _ballotMarkersComplete_(layout) {
  return _BALLOT_SECTION_KEYS_.every(function (k) { return layout.markers[k] !== -1; });
}

/**
 * Former single-purpose repair, kept as a name for callers/tests: runs the generic
 * _ensureBallotSections_.
 *
 * @param {Sheet} sheet
 * @param {Object=} layout
 * @return {Object} the current layout.
 */
function _ensureCandidatesSection_(sheet, layout) {
  return _ensureBallotSections_(sheet, layout);
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
 * Former single-purpose repair, kept as a name for callers/tests: runs the generic
 * _ensureBallotSections_ (see _repairResponsesSectionLocked_ for the Responses rule).
 *
 * @param {Sheet} sheet
 * @param {Object=} layout
 * @return {Object} the current layout.
 */
function _ensureResponsesSection_(sheet, layout) {
  return _ensureBallotSections_(sheet, layout);
}

/**
 * Re-creates a lost [Responses] marker + header directly after the sheet's last row —
 * the typical cause being a hand-paste of candidates over rows that held them (rcballot-4wm).
 * Only repairs when column A below the last present marker's header (below the marker
 * itself for [Results], which has no header) is entirely blank: column A is marker-only, so
 * any text there may be orphaned response rows (Date in col A) and appending a header
 * beneath them would misfile them — that case is logged as ballot.parse.responsesMarkerMissing
 * and nothing is written. Caller holds the ballot lock and has a current layout with no
 * [Responses] marker.
 *
 * @param {Sheet} sheet
 * @param {Object} layout from _readBallotLayout_
 * @return {boolean} true if the section was appended.
 */
function _repairResponsesSectionLocked_(sheet, layout) {
  var last = null;
  _BALLOT_SECTION_KEYS_.forEach(function (k) {
    var section = layout.sections[k];
    if (section && (!last || section.markerRow > last.markerRow)) last = section;
  });
  if (!last) return false; // no marker to anchor to: not a ballot layout we can reason about
  var lastRow = layout.lastRow;
  var firstScanRow = (last.headerRow || last.markerRow) + 1;
  var colARows = [];
  if (lastRow >= firstScanRow) {
    sheet.getRange(firstScanRow, 1, lastRow - firstScanRow + 1, 1).getValues().forEach(function (r, i) {
      if (String(r[0] || '').trim()) colARows.push(firstScanRow + i);
    });
  }
  if (colARows.length) {
    _logBallotEvent_('ballot.parse.responsesMarkerMissing', { sheet: sheet.getName(), lastMarkerRow: last.markerRow, colARows: colARows });
    return false;
  }
  var markerRow = Math.max(lastRow, firstScanRow - 1) + 1;
  sheet.getRange(markerRow, 1).setValue(_BALLOT_RESPONSES_MARKER);
  sheet.getRange(markerRow + 1, 1, 1, _BALLOT_RESPONSE_FIXED_COLS.length).setValues([_BALLOT_RESPONSE_FIXED_COLS]);
  _logBallotEvent_('ballot.repair.responsesSection', { sheet: sheet.getName(), row: markerRow });
  return true;
}

/**
 * Migrates a Candidates section still using the OLD column layout (Name/Details in
 * columns A/B, from before column A was reserved for markers only) to the current
 * layout (columns B/C, col A left blank). Detected by the Candidates header's column A
 * cell still holding the literal "Name" label — the current layout never writes
 * anything to column A there. A no-op for a sheet already on the new layout (or
 * with no Candidates section at all — call _ensureBallotSections_ first).
 *
 * Without this, readBallotCandidateDetails_ silently reads garbage (the OLD Details
 * column mistaken for the new Name column) rather than throwing, because both
 * layouts are structurally valid 2-column reads — this was found live on a ballot
 * created before the column-A-reservation fix shipped, where every candidate's Details
 * came back empty even though the sheet clearly had them.
 *
 * @param {Sheet} sheet
 * @param {Object=} layout from _readBallotLayout_, if the caller already has a current one.
 * @return {Object} the current layout (migration rewrites cells only, never rows).
 */
function _migrateCandidatesColumnsIfNeeded_(sheet, layout) {
  layout = layout || _readBallotLayout_(sheet);
  var candidates = layout.sections.candidates;
  if (!candidates) return layout;
  var headerColA = String(sheet.getRange(candidates.headerRow, 1).getValue() || '').trim();
  if (headerColA !== _BALLOT_CANDIDATES_HEADER[0]) return layout; // already migrated (col A blank) or unrecognized

  return _withBallotLock_(function () {
    // Re-check under the lock: a concurrent request may already have migrated.
    var current = _readBallotLayout_(sheet);
    var candidates = current.sections.candidates;
    if (!candidates || !current.sections.responses) return current;
    var candidatesHeaderRow = candidates.headerRow;
    var headerColA = String(sheet.getRange(candidatesHeaderRow, 1).getValue() || '').trim();
    if (headerColA !== _BALLOT_CANDIDATES_HEADER[0]) return current;

    var count = candidates.endRow - candidatesHeaderRow;
    var oldValues = count > 0 ? sheet.getRange(candidatesHeaderRow + 1, 1, count, 2).getValues() : [];

    sheet.getRange(candidatesHeaderRow, 1, count + 1, 1).clearContent(); // vacate col A (header + data rows)
    sheet.getRange(candidatesHeaderRow, _BALLOT_CANDIDATES_FIRST_COL, 1, _BALLOT_CANDIDATES_HEADER.length).setValues([_BALLOT_CANDIDATES_HEADER]);
    if (count > 0) {
      sheet.getRange(candidatesHeaderRow + 1, _BALLOT_CANDIDATES_FIRST_COL, count, 2).setValues(oldValues);
    }
    return current;
  });
}

/**
 * Backfills the Candidates header row with any header columns added to
 * _BALLOT_CANDIDATES_HEADER after a sheet was created (e.g. "Link Text"/"Link URL"
 * added after Name/Details) — a sheet already on the current header is left untouched.
 * Only ever appends missing trailing header cells; never rewrites the existing ones,
 * so it can't clobber a sheet mid-migration from the old column-A layout. Skipped when
 * the row below [Candidates] is already the next marker (no header row to backfill).
 *
 * @param {Sheet} sheet
 * @param {Object=} layout from _readBallotLayout_, if the caller already has a current one.
 * @return {Object} the current layout (header cells only, never rows).
 */
function _ensureCandidatesHeaderColumns_(sheet, layout) {
  layout = layout || _readBallotLayout_(sheet);
  function missingLabels(l) {
    var candidates = l.sections.candidates;
    if (!candidates || candidates.headerRow > candidates.endRow) return [];
    var current = sheet.getRange(candidates.headerRow, _BALLOT_CANDIDATES_FIRST_COL, 1, _BALLOT_CANDIDATES_HEADER.length).getValues()[0];
    return _BALLOT_CANDIDATES_HEADER.filter(function (label, i) { return String(current[i] || '').trim() === ''; });
  }
  if (!missingLabels(layout).length) return layout;
  return _withBallotLock_(function () {
    // Re-check under the lock before appending.
    var current = _readBallotLayout_(sheet);
    var missing = missingLabels(current);
    if (!missing.length) return current;
    var firstMissingCol = _BALLOT_CANDIDATES_FIRST_COL + (_BALLOT_CANDIDATES_HEADER.length - missing.length);
    sheet.getRange(current.sections.candidates.headerRow, firstMissingCol, 1, missing.length).setValues([missing]);
    return current;
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
 * inserted just above the first section marker, so older sheets self-heal on first
 * save instead of silently dropping the value.
 *
 * @param {Sheet} sheet
 * @param {Object} updates e.g. {Title, Description, Instructions, Footer, Contact, 'Accept-New', 'Add-Instructions', 'Admin-Only-Notes'}
 */
function writeBallotConfig_(sheet, updates) {
  var layout = _readBallotLayout_(sheet);
  var firstMarkerRow = layout.firstMarkerRow;
  var lastConfigRow = layout.configEndRow;
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

  // Reuse existing blank config rows (e.g. the spacer row above the first marker)
  // before inserting brand-new ones, so an older sheet doesn't accumulate a stray
  // blank line every time a new config field is introduced.
  missingKeys.forEach(function (key) {
    if (blankRows.length) {
      sheet.getRange(blankRows.shift(), 1, 1, 2).setValues([[key, updates[key]]]);
    } else if (firstMarkerRow !== -1) {
      sheet.insertRowsBefore(firstMarkerRow, 1);
      sheet.getRange(lastConfigRow + 1, 1, 1, 2).setValues([[key, updates[key]]]);
      firstMarkerRow++;
      lastConfigRow++;
    }
    // else: no marker and no spare blank row — nothing safe to insert relative to,
    // so this key is left unwritten (matches the historical no-op behavior).
  });
}

/**
 * Section keys in canonical order, and the column-A marker text for each.
 */
var _BALLOT_SECTION_KEYS_ = ['results', 'candidates', 'responses'];
var _BALLOT_SECTION_MARKER_TEXT_ = {
  results: _BALLOT_RESULTS_MARKER,
  candidates: _BALLOT_CANDIDATES_MARKER,
  responses: _BALLOT_RESPONSES_MARKER
};
// "<sheet>|<marker>|<rows>" keys already logged as ballot.parse.duplicateMarker in this
// execution — the layout is read several times per request and one line per finding is enough.
var _ballotLoggedDuplicates_ = {};

/**
 * Reads the whole section layout of a ballot sheet from ONE column-A read. Sections may sit
 * in any order: each runs from its marker to the row before the next marker of any kind
 * (duplicate occurrences included — a stray copy of a marker is never read as data), or to
 * the last row when it is the bottom section. A duplicated marker resolves to its first
 * occurrence and is logged. Recompute after any row insert/delete: every row number here
 * is a snapshot.
 *
 * @param {Sheet} sheet
 * @return {{lastRow:number,
 *   markers:{results:number, candidates:number, responses:number},
 *   duplicates:Array<{marker:string, rows:Array<number>}>,
 *   sections:Object<string, {markerRow:number, headerRow:?number, firstDataRow:number, endRow:number, nextMarkerRow:number}>,
 *   firstMarkerRow:number, configEndRow:number}}
 *   markers/nextMarkerRow/firstMarkerRow are -1 when absent; sections has no key for a
 *   missing marker; headerRow is null for Results (it has no header row).
 */
function _readBallotLayout_(sheet) {
  var lastRow = sheet.getLastRow();
  var colA = lastRow > 0 ? sheet.getRange(1, 1, lastRow, 1).getValues() : [];
  var layout = _scanBallotLayout_(colA.map(function (r) { return r[0]; }), lastRow);
  layout.duplicates.forEach(function (d) {
    var logKey = sheet.getName() + '|' + d.marker + '|' + d.rows.join(',');
    if (_ballotLoggedDuplicates_[logKey]) return;
    _ballotLoggedDuplicates_[logKey] = true;
    _logBallotEvent_('ballot.parse.duplicateMarker', { sheet: sheet.getName(), marker: d.marker, rows: d.rows });
  });
  return layout;
}

/**
 * Pure half of _readBallotLayout_ (and of _parseBallot_): the section rules applied to
 * column-A values already in hand. No sheet access, no logging.
 *
 * @param {Array<*>} colA column-A cell values, index 0 = row 1.
 * @param {number} lastRow last row with content anywhere on the sheet.
 * @return {Object} the layout shape documented on _readBallotLayout_.
 */
function _scanBallotLayout_(colA, lastRow) {
  var keyByText = {};
  _BALLOT_SECTION_KEYS_.forEach(function (k) { keyByText[_BALLOT_SECTION_MARKER_TEXT_[k]] = k; });

  var occurrences = { results: [], candidates: [], responses: [] };
  var allMarkerRows = []; // ascending
  for (var i = 0; i < colA.length; i++) {
    var key = keyByText[String(colA[i] || '').trim()];
    if (!key) continue;
    occurrences[key].push(i + 1);
    allMarkerRows.push(i + 1);
  }

  var markers = {};
  var duplicates = [];
  var sections = {};
  _BALLOT_SECTION_KEYS_.forEach(function (k) {
    var rows = occurrences[k];
    markers[k] = rows.length ? rows[0] : -1;
    if (rows.length > 1) duplicates.push({ marker: _BALLOT_SECTION_MARKER_TEXT_[k], rows: rows });
    if (!rows.length) return;
    var markerRow = rows[0];
    var nextMarkerRow = -1;
    for (var j = 0; j < allMarkerRows.length; j++) {
      if (allMarkerRows[j] > markerRow) { nextMarkerRow = allMarkerRows[j]; break; }
    }
    var headerRow = k === 'results' ? null : markerRow + 1;
    sections[k] = {
      markerRow: markerRow,
      headerRow: headerRow,
      firstDataRow: (headerRow || markerRow) + 1,
      // The bottom section reaches at least its own header row, so an empty last section
      // still has a well-formed (zero-row) data region.
      endRow: nextMarkerRow !== -1 ? nextMarkerRow - 1 : Math.max(lastRow, headerRow || markerRow),
      nextMarkerRow: nextMarkerRow
    };
  });

  var firstMarkerRow = allMarkerRows.length ? allMarkerRows[0] : -1;
  return {
    lastRow: lastRow,
    markers: markers,
    duplicates: duplicates,
    sections: sections,
    firstMarkerRow: firstMarkerRow,
    configEndRow: firstMarkerRow !== -1 ? firstMarkerRow - 1 : _BALLOT_CONFIG_ROWS.length
  };
}

/**
 * Scans column A for a row whose value exactly matches the (bracketed) marker. Plain
 * look-alike text such as "Responses" is never matched. Diagnostic/test helper only —
 * model code reads every marker at once through _readBallotLayout_.
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
 * Reads the config key/value rows (rows 1 through the row before the first section
 * marker, whichever section that is). Snapshot-backed (one read, no write): a legacy key
 * ('Info') reads as its current name and is renamed on the sheet only by the repair path.
 *
 * @param {Sheet} sheet
 * @return {Object} e.g. {Title, Description, Instructions, Footer, Contact, 'Accept-New', 'Add-Instructions'}
 */
function readBallotConfig_(sheet) {
  return _readBallotView_(sheet).config;
}

/**
 * @param {Object} config as returned by readBallotConfig_
 * @return {boolean}
 */
function _ballotAcceptsNew_(config) {
  return String(config['Accept-New'] || '').trim().toUpperCase() === 'TRUE';
}

/**
 * Unreferenced since rcballot-8i5.3 (delete in rcballot-8i5.5).
 *
 * @param {Sheet} sheet
 * @return {number} 1-based row of the Responses header (Date/Name/Weight/Comment/...), or -1.
 */
function _getResponsesHeaderRow_(sheet) {
  var responses = _readBallotLayout_(sheet).sections.responses;
  return responses ? responses.headerRow : -1;
}

/**
 * Candidate names, in Candidates-table order (the table is the source of candidate
 * identity). Snapshot-backed: one read, no write. A candidate with no Responses rank
 * column is still listed (its ranks read ''); the repair path adds the column.
 *
 * @param {Sheet} sheet
 * @return {Array<string>}
 */
function readBallotCandidates_(sheet) {
  return _readBallotView_(sheet).names;
}

/**
 * Superseded by the snapshot readers (rcballot-8i5.3); kept until rcballot-8i5.5. Repairs
 * as it reads — no reader calls it any more.
 *
 * @param {Sheet} sheet
 * @return {{names:Array<string>, layout:Object}}
 */
function _readBallotCandidatesAndLayout_(sheet) {
  return _reconcileCandidatesWithResponses_(sheet);
}

/**
 * SUPERSEDED (rcballot-8i5.3; delete in rcballot-8i5.5): unreferenced. Its positional
 * gap-compaction would now misattribute votes — the repair path's rankColumns step
 * replaces it.
 *
 * Ensures every candidate listed in the Candidates table (the primary source of
 * candidate identity — see file header) has a matching column in the Responses
 * header, appending one for each that's missing. Never reorders or removes existing
 * Responses columns, so ranks already recorded against them stay aligned to their
 * column; a Candidates-table entry with no column yet is always appended at the end,
 * which also matches how addBallotCandidate_ grows the two sections in lockstep, so
 * position alignment between the two sections is preserved either way.
 *
 * @param {Sheet} sheet
 * @return {{names:Array<string>, layout:Object}} the reconciled Responses-header candidate
 *   names (blank cells dropped) and the current layout — reconciling writes header cells
 *   only, never rows, so the layout stays valid.
 */
function _reconcileCandidatesWithResponses_(sheet) {
  var plan = _planCandidateReconcile_(sheet);
  if (plan.needed) {
    plan = _withBallotLock_(function () {
      // Re-plan under the lock: another request may have appended the same columns meanwhile.
      var current = _planCandidateReconcile_(sheet);
      if (current.needed) _applyCandidateReconcile_(sheet, current);
      return current;
    });
  }
  return { names: plan.compacted.concat(plan.missingNames), layout: plan.layout };
}

/**
 * Read-only half of _reconcileCandidatesWithResponses_: works out what (if anything) the
 * Responses header needs.
 *
 * @param {Sheet} sheet
 * @return {{layout:Object, needed:boolean, headerRow:number, existingNames:Array<string>, compacted:Array<string>, missingNames:Array<string>}}
 *   needed is false when the header already matches the Candidates table (or there is no
 *   Candidates table to match against); compacted is then the header's names as they stand.
 */
function _planCandidateReconcile_(sheet) {
  var read = _readCandidateDetailsAndLayout_(sheet);
  var candidateRows = read.rows;
  var responses = read.layout.sections.responses;
  var plan = { layout: read.layout, needed: false, headerRow: responses ? responses.headerRow : -1, existingNames: [], compacted: [], missingNames: [] };
  if (!responses) return plan;
  var headerRow = responses.headerRow;

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
  plan.existingNames = existingNames;
  plan.compacted = compacted;
  if (!candidateRows.length) return plan; // nothing to reconcile against: report the header as-is

  var existingSet = {};
  compacted.forEach(function (n) { existingSet[n] = true; });

  plan.missingNames = candidateRows
    .map(function (it) { return it.name; })
    .filter(function (n) { return n && !existingSet[n]; });
  plan.needed = compacted.length !== existingNames.length || plan.missingNames.length > 0;
  return plan;
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
 * Reads all response rows in the Responses section, skipping rows with no Name.
 * Snapshot-backed: one read, no write. ranks[i] is the rank for readBallotCandidates_()[i],
 * found by that candidate's NAME in the Responses header (never by position); rank columns
 * matching no candidate (orphans) are left out.
 *
 * @param {Sheet} sheet
 * @return {Array<{row:number, date:*, name:string, weight:number, comment:string, ranks:Array, ranksByName:Object}>}
 */
function readBallotResponseRows_(sheet) {
  return _readBallotView_(sheet).rows;
}

/**
 * Superseded by the snapshot readers (rcballot-8i5.3); kept until rcballot-8i5.5. Positional
 * and repairs as it reads — no reader or writer calls it any more.
 *
 * @param {Sheet} sheet
 * @return {{rows:Array<{row:number, date:*, name:string, weight:number, comment:string, ranks:Array}>, names:Array<string>, layout:Object}}
 */
function _readResponseRowsAndLayout_(sheet) {
  var read = _readBallotCandidatesAndLayout_(sheet);
  var result = { rows: [], names: read.names, layout: read.layout };
  var responses = read.layout.sections.responses;
  if (!responses) return result;
  var count = responses.endRow - responses.firstDataRow + 1;
  if (count <= 0) return result;
  var numCols = _BALLOT_RESPONSE_FIXED_COLS.length + read.names.length;
  var values = sheet.getRange(responses.firstDataRow, 1, count, numCols).getValues();
  values.forEach(function (row, i) {
    var name = String(row[1] || '').trim();
    if (!name) return;
    result.rows.push({
      row: responses.firstDataRow + i,
      date: row[0],
      name: name,
      weight: Number(row[2]) || 1,
      comment: String(row[3] || ''),
      ranks: row.slice(_BALLOT_RESPONSE_FIXED_COLS.length)
    });
  });
  return result;
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
 * addition(s). Detected via a blank rank: a candidate added after a response was written
 * has a blank cell (or no column at all) for that response, since submitBallotResponse_
 * writes every candidate known at submit time.
 *
 * @param {Sheet} sheet
 * @return {Array<string>} respondent names (as last submitted), one per stale respondent
 */
function findRespondentsWithNewCandidates_(sheet) {
  return _staleRespondents_(_latestResponseByRespondent_(readBallotResponseRows_(sheet)));
}

/**
 * @param {Array<{name:string, ranks:Array}>} latestRows one row per respondent
 * @return {Array<string>} names of respondents with any blank rank.
 */
function _staleRespondents_(latestRows) {
  return latestRows
    .filter(function (r) { return r.ranks.some(function (v) { return v === '' || v === null || v === undefined; }); })
    .map(function (r) { return r.name; });
}

/**
 * Finds the last response row for a respondent (case-insensitive), so a
 * re-submission overwrites their existing row instead of adding a duplicate.
 *
 * @param {Array<{row:number, name:string}>} rows as returned by readBallotResponseRows_
 * @param {string} name
 * @return {number} 1-based row, or -1 if not found.
 */
function _findBallotResponseRowForName_(rows, name) {
  var target = String(name).trim().toLowerCase();
  for (var i = rows.length - 1; i >= 0; i--) {
    if (rows[i].name.toLowerCase() === target) return rows[i].row;
  }
  return -1;
}

/**
 * Appends a new candidate: a row at the end of the Candidates table, and a rank column in
 * the Responses header — reusing an existing column already labelled with this name
 * (e.g. an orphan left by a hand-deleted candidate), else appended after the last used
 * Responses column. Runs the repair path first. Caller must check Accept-New first.
 *
 * @param {Sheet} sheet
 * @param {string} phrase
 * @param {string=} details admin-only note for this candidate (respondents never pass this).
 * @param {string=} linkText admin-only link label shown alongside details (respondents never pass this).
 * @param {string=} linkUrl admin-only link target, opened in a new tab (respondents never pass this).
 */
function addBallotCandidate_(sheet, phrase, details, linkText, linkUrl) {
  _assertNotMarkerName_(phrase);
  var snap = _repairBallotSheet_(sheet).snap;
  if (!snap.tables.responses) throw new Error('Ballot has no Responses section.');
  var candidateRows = _snapCandidates_(snap).map(function (c) {
    return { name: c.name, details: c.details, linkText: c.linkText, linkUrl: c.linkUrl };
  });
  candidateRows.push({ name: phrase, details: details || '', linkText: linkText || '', linkUrl: linkUrl || '' });
  writeBallotCandidateDetails_(sheet, candidateRows);

  // Re-read: the Candidates resize may have shifted the Responses section.
  var cur = _readBallotGridAndSnap_(sheet);
  var responses = cur.snap.tables.responses;
  if (responses.headerRow === null) throw new Error('Ballot sheet is missing its Responses header row.');
  if (_snapRankColumns_(cur.snap)[_labelKey_(phrase)]) return;
  sheet.getRange(responses.headerRow, _nextResponsesCol_(cur.grid, cur.snap)).setValue(phrase);
}

/**
 * @param {Array<Array<*>>} grid
 * @param {Object} snap from _parseBallot_(grid), with a [Responses] section
 * @return {number} the first column right of every labelled Responses header cell and of
 *   any content in the Responses section's rows — where a new rank column may go without
 *   adopting existing (unlabelled) data.
 */
function _nextResponsesCol_(grid, snap) {
  var section = snap.layout.sections.responses;
  var t = snap.tables.responses;
  var last = _BALLOT_RESPONSE_FIXED_COLS.length;
  Object.keys(t.col).forEach(function (l) { if (t.col[l] > last) last = t.col[l]; });
  for (var row = section.headerRow; row <= section.endRow; row++) {
    var cells = grid[row - 1] || [];
    for (var c = cells.length; c > last; c--) {
      if (_cellHasContent_(cells[c - 1])) { last = c; break; }
    }
  }
  return last + 1;
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
  // always requires the live sheet: one snapshot read serves config, candidates and
  // responses, and re-warms the cache for the next (nameless) load. No sheet write.
  var view = _readBallotView_(sheet);
  var freshCache = refreshBallotCache_(id, sheet, view);
  var result = _buildRespondentBranding_(id, freshCache.config, freshCache.candidates);

  if (_responsesMissingName_(view.snap)) {
    // No vote can be keyed by name on this sheet: the respondent's saved ranking would
    // silently not load. Repaired on the next admin edit-page load or write (rcballot-8i5.7).
    _logBallotEvent_('ballot.parse.responsesMissingName', { sheet: sheet.getName() });
  }
  var rows = view.rows;
  var target = name.toLowerCase();
  var existing = null;
  for (var i = rows.length - 1; i >= 0; i--) {
    if (rows[i].name.toLowerCase() === target) { existing = rows[i]; break; }
  }
  if (!existing) return result;

  // Ranks by candidate NAME (rcballot-8i5.3), so a hand-sorted or hand-trimmed Candidates
  // table cannot hand this respondent's rank for one candidate to another.
  var withRank = view.names.map(function (c) {
    var rank = Number(existing.ranksByName[c]);
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
    _repairBallotSheet_(sheet);
    var config = readBallotConfig_(sheet);
    if (!_ballotAcceptsNew_(config)) throw new Error('This ballot is not accepting new items.');
    addBallotCandidate_(sheet, phrase, details);
    refreshBallotCache_(id, sheet); // keep the cache (BallotCache.js) in sync with this write
    return { candidate: phrase };
  });
}

/**
 * Picks (and if needed makes room for) the row a new respondent's vote goes in: directly
 * below the last non-empty row of the Responses section — named vote or not, so a stray
 * unnamed row with content (a hand-typed note) is never overwritten — or below its header
 * when it has none. When that row is the next section's marker, a blank row is inserted in front of
 * the marker first, so the following section shifts down intact instead of being
 * overwritten. When Responses is the bottom section the row may simply be past the last
 * row. Never sheet.getLastRow()+1 — that is only right while Responses happens to be last.
 * Caller holds the ballot lock.
 *
 * @param {Sheet} sheet
 * @param {{headerRow:number, endRow:number, nextMarkerRow:number}} responses layout section
 * @param {Array<{row:number}>} rows the section's named response rows, ascending
 * @return {number} 1-based row to write the new vote into.
 */
function _newResponseRow_(sheet, responses, rows) {
  var lastUsed = rows.length ? rows[rows.length - 1].row : responses.headerRow;
  if (responses.endRow > lastUsed) {
    var tailStart = lastUsed + 1;
    var tail = sheet.getRange(tailStart, 1, responses.endRow - lastUsed, Math.max(sheet.getLastColumn(), 1)).getValues();
    tail.forEach(function (r, i) {
      var used = r.some(function (v) { return String(v === null || v === undefined ? '' : v).trim() !== ''; });
      if (used) lastUsed = tailStart + i;
    });
  }
  var row = lastUsed + 1;
  if (row > responses.endRow && responses.nextMarkerRow !== -1) {
    row = responses.nextMarkerRow;
    sheet.insertRowsBefore(row, 1);
  }
  return row;
}

/**
 * RPC-facing: saves a respondent's ranking (and optional comment). Reuses
 * their existing row and weight if one exists (overwritten in place), else adds a new
 * row with weight 1 at the end of the Responses section (see _newResponseRow_).
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
    var rep = _repairBallotSheet_(sheet);
    var snap = rep.snap;
    var t = snap.tables.responses;
    if (!t) throw new Error('Ballot has no Responses section.');
    if (t.headerRow === null || _responsesMissingName_(snap)) {
      _logBallotEvent_('ballot.write.refused', { sheet: sheet.getName(), reason: 'responsesMissingName' });
      throw new Error('Ballot sheet layout is damaged — the Responses header has no Name column.');
    }
    var view = _ballotViewFromSnapshot_(snap);
    var rankByCandidate = {};
    orderedCandidates.forEach(function (c, idx) {
      rankByCandidate[String(c).trim()] = idx + 1;
    });

    var row = _findBallotResponseRowForName_(view.rows, name);
    var base = [];
    var weight = 1;
    if (row === -1) {
      row = _newResponseRow_(sheet, snap.layout.sections.responses, view.rows);
    } else {
      base = rep.grid[row - 1] || [];
      if (t.col.Weight) weight = Number(_gridCell_(rep.grid, row, t.col.Weight)) || 1;
    }

    // Every cell is placed by its header label: the fixed columns, and each candidate's
    // rank under the column bearing its name (rcballot-8i5.3). Cells in between (orphan
    // rank columns, unlabelled columns) keep what the row already had.
    var put = {};
    var fixed = { Date: new Date(), Name: name, Weight: weight, Comment: String(comment || '').trim() };
    Object.keys(fixed).forEach(function (label) { if (t.col[label]) put[t.col[label]] = fixed[label]; });
    _snapCandidates_(snap).forEach(function (c) {
      if (c.rankCol !== null) put[c.rankCol] = rankByCandidate.hasOwnProperty(c.name) ? rankByCandidate[c.name] : '';
    });
    var cols = Object.keys(put).map(Number);
    var minCol = Math.min.apply(null, cols);
    var maxCol = Math.max.apply(null, cols);
    var values = [];
    for (var c = minCol; c <= maxCol; c++) {
      if (put.hasOwnProperty(c)) values.push(put[c]);
      else values.push(base[c - 1] === undefined || base[c - 1] === null ? '' : base[c - 1]);
    }
    sheet.getRange(row, minCol, 1, values.length).setValues([values]);

    return { ok: true };
  });
}

/**
 * Decides whether writeBallotResults_ must refuse to rewrite the Results region. The rewrite
 * clears and resizes everything between [Results] and the next marker, so it is only safe
 * when that region provably holds analysis output: refuses when a marker is still missing
 * after repair, when any marker is duplicated, or when the region holds what looks like
 * another section's header row absorbed into it (Candidates: col B "Name" + col C
 * "Details"; Responses: col A "Date" + col B "Name") — the signature of a deleted marker.
 *
 * @param {Sheet} sheet
 * @param {Object} layout from _readBallotLayout_
 * @return {?{reason:string, markers:(Array<string>|undefined), row:(number|undefined)}} null when safe.
 */
function _resultsRewriteRefusal_(sheet, layout) {
  var missing = _BALLOT_SECTION_KEYS_.filter(function (k) { return layout.markers[k] === -1; });
  if (missing.length) {
    return { reason: 'markerMissing', markers: missing.map(function (k) { return _BALLOT_SECTION_MARKER_TEXT_[k]; }) };
  }
  if (layout.duplicates.length) {
    return { reason: 'duplicateMarker', markers: layout.duplicates.map(function (d) { return d.marker; }) };
  }
  var results = layout.sections.results;
  var count = results.endRow - results.firstDataRow + 1;
  if (count <= 0) return null;
  var values = sheet.getRange(results.firstDataRow, 1, count, 3).getValues();
  for (var i = 0; i < values.length; i++) {
    var a = String(values[i][0] || '').trim();
    var b = String(values[i][1] || '').trim();
    var c = String(values[i][2] || '').trim();
    if (b === _BALLOT_CANDIDATES_HEADER[0] && c === _BALLOT_CANDIDATES_HEADER[1]) {
      return { reason: 'absorbedCandidatesHeader', row: results.firstDataRow + i };
    }
    if (a === _BALLOT_RESPONSE_FIXED_COLS[0] && b === _BALLOT_RESPONSE_FIXED_COLS[1]) {
      return { reason: 'absorbedResponsesHeader', row: results.firstDataRow + i };
    }
  }
  return null;
}

/**
 * Inserts or deletes rows so the Results section has exactly rows.length
 * rows available, then writes rows into it. Leaves every other section untouched — the
 * insert/delete happens strictly between the [Results] marker and the next marker below
 * it (whichever section that is), so everything below shifts as one block.
 * The region is cleared in full first and output starts at column B (column A is
 * markers-only), so older column-A output self-corrects on the next run.
 *
 * @param {Sheet} sheet
 * @param {Array<Array>} rows
 */
function writeBallotResults_(sheet, rows) {
  var layout = _ensureBallotSections_(sheet);
  var refusal = _resultsRewriteRefusal_(sheet, layout);
  if (refusal) {
    refusal.sheet = sheet.getName();
    _logBallotEvent_('ballot.write.refused', refusal);
    throw new Error('Ballot sheet layout is damaged — fix the markers before running analysis.');
  }
  var results = layout.sections.results;
  // Clear the whole region (every column, col A included) first: column A is
  // markers-only, so output from the older layout that wrote names there must not
  // survive the rewrite; the resize then inserts/deletes blank rows only.
  var regionRows = results.endRow - results.firstDataRow + 1;
  if (regionRows > 0) {
    sheet.getRange(results.firstDataRow, 1, regionRows, Math.max(sheet.getLastColumn(), 1)).clearContent();
  }
  _resizeSectionRows_(sheet, results.markerRow, results.endRow + 1, rows, _BALLOT_CANDIDATES_FIRST_COL);
}

/**
 * Runs the section self-repair plus the Candidates-side migrations and returns the layout
 * they leave behind — the shared preamble of readBallotCandidateDetails_ and
 * writeBallotCandidateDetails_, and through them of every other section read/write entry.
 *
 * @param {Sheet} sheet
 * @return {Object} layout from _readBallotLayout_.
 */
function _prepareCandidatesSection_(sheet) {
  var layout = _ensureBallotSections_(sheet);
  layout = _migrateCandidatesColumnsIfNeeded_(sheet, layout);
  return _ensureCandidatesHeaderColumns_(sheet, layout);
}

/**
 * Reads the Candidates table (Name/Details/Link Text/Link URL per row, columns found by
 * header label), in sheet order — the primary source of candidate identity (see file
 * header). Snapshot-backed: one read, no write. Rows with no Name are skipped.
 *
 * @param {Sheet} sheet
 * @return {Array<{name:string, details:string, linkText:string, linkUrl:string}>}
 */
function readBallotCandidateDetails_(sheet) {
  return _readBallotView_(sheet).candidates;
}

/**
 * Superseded by the snapshot readers (rcballot-8i5.3); kept until rcballot-8i5.5. Repairs
 * as it reads — no reader calls it any more.
 *
 * @param {Sheet} sheet
 * @return {{rows:Array<{name:string, details:string, linkText:string, linkUrl:string}>, layout:Object}}
 */
function _readCandidateDetailsAndLayout_(sheet) {
  var layout = _prepareCandidatesSection_(sheet);
  var candidates = layout.sections.candidates;
  // Without a [Responses] marker the rows below [Candidates] may be orphaned votes (see
  // _repairResponsesSectionLocked_), so nothing is read as a candidate until it is repaired.
  if (!candidates || !layout.sections.responses) {
    _logBallotEvent_('ballot.parse.markerMissing', { sheet: sheet.getName(), candidatesRow: layout.markers.candidates, responsesRow: layout.markers.responses });
    return { rows: [], layout: layout };
  }
  var firstRow = candidates.firstDataRow;
  var count = candidates.endRow - firstRow + 1;
  if (count <= 0) return { rows: [], layout: layout };
  // One read: column A (for the warning below) plus the table's own columns.
  var values = sheet.getRange(firstRow, 1, count, _BALLOT_CANDIDATES_FIRST_COL - 1 + _BALLOT_CANDIDATES_HEADER.length).getValues();
  // Column A is marker-only (see file header) and never read as a candidate — but a name
  // typed there is an easy mistake, so say so rather than let it vanish silently.
  var colARows = [];
  values.forEach(function (r, i) {
    if (String(r[0] || '').trim()) colARows.push(firstRow + i);
  });
  if (colARows.length) {
    _logBallotEvent_('ballot.parse.candidateInColumnA', { sheet: sheet.getName(), rows: colARows });
  }
  // A row with no name (e.g. a stray blank row someone left in the sheet) isn't a
  // candidate — skip it (and keep reading below it) rather than returning a phantom entry
  // that would throw off the position alignment with the Responses header's candidate columns.
  var rows = values
    .map(function (r) {
      var c = _BALLOT_CANDIDATES_FIRST_COL - 1;
      return { name: String(r[c] || '').trim(), details: String(r[c + 1] || ''), linkText: String(r[c + 2] || ''), linkUrl: String(r[c + 3] || '') };
    })
    .filter(function (it) { return it.name; });
  return { rows: rows, layout: layout };
}

/**
 * Resizes + rewrites the Candidates table to exactly match `candidateRows` (array of
 * {name, details, linkText, linkUrl}, one per candidate column, in the same order).
 * Leaves every other section untouched — the insert/delete happens strictly between
 * the Candidates header and the next marker below it (or the end of the sheet).
 *
 * @param {Sheet} sheet
 * @param {Array<{name:string, details:string, linkText:string, linkUrl:string}>} candidateRows
 */
function writeBallotCandidateDetails_(sheet, candidateRows) {
  var layout = _prepareCandidatesSection_(sheet);
  var candidates = layout.sections.candidates;
  if (!candidates || !layout.sections.responses) {
    throw new Error('Ballot sheet is missing its Candidates/Responses markers.');
  }
  if (candidates.headerRow > candidates.endRow) {
    // The row below [Candidates] is already the next marker: there is no header row to
    // write beneath, and resizing from it would land inside the next section.
    throw new Error('Ballot sheet is missing its Candidates header row.');
  }
  var rows = candidateRows.map(function (it) { return [it.name || '', it.details || '', it.linkText || '', it.linkUrl || '']; });
  _resizeSectionRows_(sheet, candidates.headerRow, candidates.endRow + 1, rows, _BALLOT_CANDIDATES_FIRST_COL);
}

/**
 * Admin-facing: overwrites every candidate's name + details in one pass, given a full
 * list in the ballot's current Candidates-table order (as getAdminEditData returns it).
 * A rename relabels that candidate's own Responses rank column (found by its old name), so
 * recorded ranks follow the candidate. Throws if the count doesn't match the ballot's
 * current candidate count (stale form — caller should reload and retry). Runs the repair
 * path first.
 *
 * @param {Sheet} sheet
 * @param {Array<{name:string, details:string, linkText:string, linkUrl:string}>} candidateRows
 */
function saveBallotCandidates_(sheet, candidateRows) {
  var snap = _repairBallotSheet_(sheet).snap;
  var candidates = _snapCandidates_(snap);
  if (candidateRows.length !== candidates.length) {
    throw new Error('Candidate count changed since the page was loaded — reload and try again.');
  }
  var names = candidateRows.map(function (it) {
    var name = String(it.name || '').trim();
    if (!name) throw new Error('Candidate name cannot be empty.');
    _assertNotMarkerName_(name);
    return name;
  });
  var responses = snap.tables.responses;
  candidates.forEach(function (c, i) {
    if (c.rankCol !== null && c.rankLabel !== names[i]) {
      sheet.getRange(responses.headerRow, c.rankCol).setValue(names[i]);
    }
  });
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
    _repairBallotSheet_(sheet);
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
    _repairBallotSheet_(sheet);
    addBallotCandidate_(sheet, name, details, linkText, linkUrl);
    _highlightSectionMarkers_(sheet);
    refreshBallotCache_(id, sheet); // keep the cache (BallotCache.js) in sync with this write
    return { candidate: name };
  });
}

// ---------------------------------------------------------------------------
// Snapshot reader (rcballot-8i5.1). One getDataRange read parsed into a plain Snapshot;
// columns resolved by header label, ranks keyed by candidate name. Lives alongside the old
// per-section readers until rcballot-8i5.3 switches callers over. Never writes.
// ---------------------------------------------------------------------------

/**
 * Section registry. config is the key/value block above the first marker; results is
 * free-form output from col B; candidates/responses are header-labelled tables keyed by
 * their 'Name' column. responses' extra (non-fixed) columns are candidate rank columns.
 */
var _BALLOT_SECTIONS_ = {
  config: { marker: null, kind: 'kv' },
  results: { marker: _BALLOT_RESULTS_MARKER, kind: 'output', firstCol: 2 },
  candidates: { marker: _BALLOT_CANDIDATES_MARKER, kind: 'table', firstCol: _BALLOT_CANDIDATES_FIRST_COL, key: 'Name', columns: _BALLOT_CANDIDATES_HEADER },
  responses: { marker: _BALLOT_RESPONSES_MARKER, kind: 'table', firstCol: 1, key: 'Name', columns: _BALLOT_RESPONSE_FIXED_COLS, extraColumns: 'candidates' }
};

/** @return {*} grid cell at 1-based (row, col), '' when absent. */
function _gridCell_(grid, row, col) {
  var r = grid[row - 1];
  if (!r) return '';
  var v = r[col - 1];
  return v === undefined || v === null ? '' : v;
}

/** @return {boolean} true when v is a non-blank cell value (0 and false count as content). */
function _cellHasContent_(v) {
  return v !== '' && v !== null && v !== undefined && String(v).trim() !== '';
}

/** @return {boolean} true when any cell in grid row `row` within [fromCol, toCol] has content. */
function _gridRowHasContent_(grid, row, fromCol, toCol) {
  for (var c = fromCol; c <= toCol; c++) {
    if (_cellHasContent_(_gridCell_(grid, row, c))) return true;
  }
  return false;
}

/** Normalised form used for every label/name comparison: trimmed, lower-case. */
function _labelKey_(v) {
  return String(v === null || v === undefined ? '' : v).trim().toLowerCase();
}

/**
 * Parses a whole ballot sheet's values into a Snapshot. Pure: plain 2D array in, plain
 * object out; touches no GAS global and logs nothing — findings are returned in `problems`.
 *
 * @param {Array<Array<*>>} grid sheet values, grid[0] = row 1 (ragged rows allowed).
 * @return {{
 *   layout: Object,   // _scanBallotLayout_ shape; lastRow = last row with content
 *   config: {byKey: Object<string, {row:number, value:*, legacyKey:(string|undefined)}>, rowsUsed:number},
 *   tables: {candidates: ?Table, responses: ?Table},
 *   resultsRows: ?Array<{row:number, values:Array<*>}>,
 *   problems: Array<{code:string}>
 * }}
 *   Table = {headerRow:?number, col:Object<string,number>, labels:Array<string>,
 *            missing:Array<string>, rows:Array<{row:number, key:string, cells:Object<string,*>}>,
 *            rowsUsed:number}
 *   Fixed labels are matched trim/case-insensitively and keyed by their canonical spelling
 *   (e.g. 'Link URL'); other labels keep their trimmed header spelling. A table or
 *   resultsRows is null when its marker is missing. Problem codes: duplicateMarker
 *   {marker, rows}, missingMarker {section}, missingLabel {section, labels},
 *   duplicateLabel {section, label, col}, headerGap {section, cols}, candidateInColumnA {rows}.
 */
function _parseBallot_(grid) {
  grid = grid || [];
  var lastCol = 0;
  var lastRow = 0;
  for (var r = 0; r < grid.length; r++) {
    var cells = grid[r] || [];
    for (var c = cells.length - 1; c >= 0; c--) {
      if (_cellHasContent_(cells[c])) {
        lastRow = r + 1;
        if (c + 1 > lastCol) lastCol = c + 1;
        break;
      }
    }
  }
  var colA = [];
  for (var i = 1; i <= lastRow; i++) colA.push(_gridCell_(grid, i, 1));
  var layout = _scanBallotLayout_(colA, lastRow);

  var problems = [];
  layout.duplicates.forEach(function (d) { problems.push({ code: 'duplicateMarker', marker: d.marker, rows: d.rows }); });
  _BALLOT_SECTION_KEYS_.forEach(function (k) {
    if (!layout.sections[k]) problems.push({ code: 'missingMarker', section: k });
  });

  var tables = {};
  ['candidates', 'responses'].forEach(function (k) {
    var section = layout.sections[k];
    tables[k] = section ? _parseBallotTable_(grid, section, k, lastCol, problems) : null;
  });
  var candidatesSection = layout.sections.candidates;
  if (candidatesSection) {
    var colARows = [];
    for (var cr = candidatesSection.firstDataRow; cr <= candidatesSection.endRow; cr++) {
      if (String(_gridCell_(grid, cr, 1) || '').trim()) colARows.push(cr);
    }
    if (colARows.length) problems.push({ code: 'candidateInColumnA', rows: colARows });
  }

  return {
    layout: layout,
    config: _parseBallotConfig_(grid, layout),
    tables: tables,
    resultsRows: _parseBallotResults_(grid, layout.sections.results, lastCol),
    problems: problems
  };
}

/**
 * Config rows 1..layout.configEndRow: key col A (trimmed; legacy names mapped to their
 * current key, recorded as legacyKey), value col B. Blank keys skipped; a repeated key's
 * last row wins (as readBallotConfig_). rowsUsed = last row with content in cols A-B, or 0.
 */
function _parseBallotConfig_(grid, layout) {
  var legacyToCurrent = {};
  Object.keys(_BALLOT_LEGACY_CONFIG_KEYS_).forEach(function (current) {
    _BALLOT_LEGACY_CONFIG_KEYS_[current].forEach(function (legacy) { legacyToCurrent[legacy] = current; });
  });
  var config = { byKey: {}, rowsUsed: 0 };
  for (var row = 1; row <= layout.configEndRow; row++) {
    if (_gridRowHasContent_(grid, row, 1, 2)) config.rowsUsed = row;
    var key = String(_gridCell_(grid, row, 1) || '').trim();
    if (!key) continue;
    var entry = { row: row, value: _gridCell_(grid, row, 2) };
    if (legacyToCurrent.hasOwnProperty(key)) {
      entry.legacyKey = key;
      key = legacyToCurrent[key];
    }
    config.byKey[key] = entry;
  }
  return config;
}

/**
 * One header-labelled table section. The header row is the row below the marker; a
 * section with no room for one (next marker directly below) gets headerRow null.
 * Rows with a blank key cell are skipped and the scan continues; rowsUsed is the last row
 * of the section with content in any column (blank-key rows included), or the header
 * row (marker row when there is no header).
 */
function _parseBallotTable_(grid, section, sectionKey, lastCol, problems) {
  var def = _BALLOT_SECTIONS_[sectionKey];
  var table = { headerRow: null, col: {}, labels: [], missing: [], rows: [], rowsUsed: section.markerRow };
  var fixedByKey = {};
  def.columns.forEach(function (label) { fixedByKey[_labelKey_(label)] = label; });

  if (section.headerRow <= section.endRow) {
    table.headerRow = section.headerRow;
    table.rowsUsed = section.headerRow;
    var seen = {};
    var blankCols = [];
    var gapCols = [];
    for (var c = def.firstCol; c <= lastCol; c++) {
      var raw = _gridCell_(grid, section.headerRow, c);
      var lk = _labelKey_(raw);
      if (!lk) { blankCols.push(c); continue; }
      gapCols = gapCols.concat(blankCols);
      blankCols = [];
      var label = fixedByKey.hasOwnProperty(lk) ? fixedByKey[lk] : String(raw).trim();
      if (seen[lk]) {
        problems.push({ code: 'duplicateLabel', section: sectionKey, label: label, col: c });
        continue;
      }
      seen[lk] = true;
      table.col[label] = c;
      table.labels.push(label);
    }
    if (gapCols.length) problems.push({ code: 'headerGap', section: sectionKey, cols: gapCols });
  }

  def.columns.forEach(function (label) {
    if (!table.col.hasOwnProperty(label)) table.missing.push(label);
  });
  if (table.missing.length) problems.push({ code: 'missingLabel', section: sectionKey, labels: table.missing.slice() });
  if (table.headerRow === null) return table;

  var keyCol = table.col[def.key];
  for (var row = table.headerRow + 1; row <= section.endRow; row++) {
    if (_gridRowHasContent_(grid, row, 1, lastCol)) table.rowsUsed = row;
    if (!keyCol) continue;
    var key = String(_gridCell_(grid, row, keyCol) || '').trim();
    if (!key) continue;
    var cells = {};
    table.labels.forEach(function (l) { cells[l] = _gridCell_(grid, row, table.col[l]); });
    table.rows.push({ row: row, key: key, cells: cells });
  }
  return table;
}

/**
 * Results output rows with content from col B on (trailing blanks trimmed); null when
 * there is no [Results] marker.
 */
function _parseBallotResults_(grid, section, lastCol) {
  if (!section) return null;
  var firstCol = _BALLOT_SECTIONS_.results.firstCol;
  var out = [];
  for (var row = section.firstDataRow; row <= section.endRow; row++) {
    var values = [];
    for (var c = firstCol; c <= lastCol; c++) values.push(_gridCell_(grid, row, c));
    while (values.length && !_cellHasContent_(values[values.length - 1])) values.pop();
    if (values.length) out.push({ row: row, values: values });
  }
  return out;
}

/**
 * Reads a ballot sheet as a Snapshot: exactly one getDataRange().getValues() call, no
 * writes, no lock.
 *
 * @param {Sheet} sheet
 * @return {Object} see _parseBallot_.
 */
function _readBallot_(sheet) {
  return _parseBallot_(sheet.getDataRange().getValues());
}

/**
 * Responses extra columns (neither a fixed column nor blank), by normalised label.
 * @return {Object<string, {label:string, col:number}>}
 */
function _snapRankColumns_(snap) {
  var t = snap.tables.responses;
  var out = {};
  if (!t) return out;
  var fixed = {};
  _BALLOT_SECTIONS_.responses.columns.forEach(function (l) { fixed[l] = true; });
  t.labels.forEach(function (l) {
    if (!fixed[l]) out[_labelKey_(l)] = { label: l, col: t.col[l] };
  });
  return out;
}

/**
 * Candidates-table rows in sheet order, each with its Responses rank column found by
 * name (trim, case-insensitive) or null. Empty when either the Candidates or the
 * Responses marker is missing (as readBallotCandidateDetails_: rows below [Candidates]
 * may then be orphaned votes).
 *
 * @param {Object} snap from _parseBallot_
 * @return {Array<{row:number, name:string, details:string, linkText:string, linkUrl:string,
 *   rankLabel:?string, rankCol:?number}>}
 */
function _snapCandidates_(snap) {
  var t = snap.tables.candidates;
  if (!t || !snap.layout.sections.responses) return [];
  var rankCols = _snapRankColumns_(snap);
  return t.rows.map(function (r) {
    var rank = rankCols[_labelKey_(r.key)] || null;
    return {
      row: r.row,
      name: r.key,
      details: String(r.cells.Details || ''),
      linkText: String(r.cells['Link Text'] || ''),
      linkUrl: String(r.cells['Link URL'] || ''),
      rankLabel: rank ? rank.label : null,
      rankCol: rank ? rank.col : null
    };
  });
}

/**
 * Responses header labels that are neither fixed columns nor any candidate's name — rank
 * data with no candidate (e.g. its Candidates row was deleted or renamed by hand).
 *
 * @param {Object} snap
 * @return {Array<{label:string, col:number}>} in column order.
 */
function _snapOrphanColumns_(snap) {
  var rankCols = _snapRankColumns_(snap);
  var claimed = {};
  _snapCandidates_(snap).forEach(function (c) { claimed[_labelKey_(c.name)] = true; });
  return Object.keys(rankCols)
    .filter(function (k) { return !claimed[k]; })
    .map(function (k) { return rankCols[k]; })
    .sort(function (a, b) { return a.col - b.col; });
}

/**
 * @param {Object} snap
 * @return {{candidates:?Array<string>, responses:?Array<string>}} fixed header labels
 *   missing from each table (null when the section's marker is missing).
 */
function _snapMissingLabels_(snap) {
  return {
    candidates: snap.tables.candidates ? snap.tables.candidates.missing.slice() : null,
    responses: snap.tables.responses ? snap.tables.responses.missing.slice() : null
  };
}

/**
 * Response rows (non-blank Name) in sheet order, ranks keyed by candidate name. A
 * candidate with no rank column ranks ''. Values under orphan columns are kept apart in
 * orphanRanks (keyed by header label). Missing fixed columns read as '' (weight 1).
 *
 * @param {Object} snap
 * @return {Array<{row:number, date:*, name:string, weight:number, comment:string,
 *   ranks:Object<string,*>, orphanRanks:Object<string,*>}>}
 */
function _snapResponses_(snap) {
  var t = snap.tables.responses;
  if (!t) return [];
  var candidates = _snapCandidates_(snap);
  var orphans = _snapOrphanColumns_(snap);
  return t.rows.map(function (r) {
    var ranks = {};
    candidates.forEach(function (c) { ranks[c.name] = c.rankLabel !== null ? r.cells[c.rankLabel] : ''; });
    var orphanRanks = {};
    orphans.forEach(function (o) { orphanRanks[o.label] = r.cells[o.label]; });
    return {
      row: r.row,
      date: r.cells.hasOwnProperty('Date') ? r.cells.Date : '',
      name: r.key,
      weight: Number(r.cells.Weight) || 1,
      comment: String(r.cells.Comment || ''),
      ranks: ranks,
      orphanRanks: orphanRanks
    };
  });
}

/**
 * Everything the public readers return, derived from one snapshot (rcballot-8i5.3).
 * Candidates are in Candidates-table order; each response's ranks array is aligned to
 * that list and filled by candidate NAME. Orphan rank columns are listed apart and never
 * reach `rows`, so analysis cannot count them.
 *
 * @param {Object} snap from _parseBallot_
 * @return {{snap:Object, config:Object, names:Array<string>,
 *   candidates:Array<{name:string, details:string, linkText:string, linkUrl:string}>,
 *   rows:Array<{row:number, date:*, name:string, weight:number, comment:string, ranks:Array, ranksByName:Object}>,
 *   orphans:Array<{label:string, col:number}>}}
 */
function _ballotViewFromSnapshot_(snap) {
  var config = {};
  Object.keys(snap.config.byKey).forEach(function (k) { config[k] = snap.config.byKey[k].value; });
  var cands = _snapCandidates_(snap);
  var names = cands.map(function (c) { return c.name; });
  return {
    snap: snap,
    config: config,
    names: names,
    candidates: cands.map(function (c) { return { name: c.name, details: c.details, linkText: c.linkText, linkUrl: c.linkUrl }; }),
    rows: _snapResponses_(snap).map(function (r) {
      return {
        row: r.row, date: r.date, name: r.name, weight: r.weight, comment: r.comment,
        ranks: names.map(function (n) { return r.ranks[n]; }),
        ranksByName: r.ranks
      };
    }),
    orphans: _snapOrphanColumns_(snap)
  };
}

// "<sheet>|<code>|<detail>" keys already logged by _logSnapshotProblems_ in this execution.
var _ballotLoggedProblems_ = {};

/**
 * Logs a snapshot's problems, each once per execution (readers run several times per
 * request). duplicateMarker keeps _readBallotLayout_'s tag, payload and de-dup key; every
 * other code logs as ballot.parse.<code>. Logging only — never a sheet write.
 */
function _logSnapshotProblems_(sheet, snap) {
  var name = sheet.getName();
  snap.problems.forEach(function (p) {
    if (p.code === 'duplicateMarker') {
      var dupKey = name + '|' + p.marker + '|' + p.rows.join(',');
      if (_ballotLoggedDuplicates_[dupKey]) return;
      _ballotLoggedDuplicates_[dupKey] = true;
      _logBallotEvent_('ballot.parse.duplicateMarker', { sheet: name, marker: p.marker, rows: p.rows });
      return;
    }
    var data = { sheet: name };
    Object.keys(p).forEach(function (k) { if (k !== 'code') data[k] = p[k]; });
    var key = name + '|' + p.code + '|' + JSON.stringify(data);
    if (_ballotLoggedProblems_[key]) return;
    _ballotLoggedProblems_[key] = true;
    _logBallotEvent_('ballot.parse.' + p.code, data);
  });
}

/**
 * One snapshot read of a ballot sheet, problems logged, as the readers' view. No write, no lock.
 *
 * @param {Sheet} sheet
 * @return {Object} see _ballotViewFromSnapshot_.
 */
function _readBallotView_(sheet) {
  var snap = _readBallot_(sheet);
  _logSnapshotProblems_(sheet, snap);
  return _ballotViewFromSnapshot_(snap);
}

// ---------------------------------------------------------------------------
// Repair path (rcballot-8i5.7). One entry, _repairBallotSheet_, run on admin edit-page load
// and at the start of the locked write entries. A healthy sheet costs one getDataRange read
// and takes no lock; otherwise each step re-plans from a fresh snapshot under the lock.
// ---------------------------------------------------------------------------

/**
 * Plans the "labels on the marker row" repair: a legacy Responses layout where the
 * [Responses] marker row carries Name|Weight|Comment in cols B-D and the header row below
 * holds Date + candidate names with B-D blank (live: Ballot-ABC-2026-7 on SIT and PROD).
 * Each marker-row cell (col B on) whose text is a fixed Responses label missing from the
 * header, sitting above a blank header cell, is planned to move down into that cell. Pure.
 *
 * @param {Array<Array<*>>} grid
 * @param {Object} snap from _parseBallot_(grid)
 * @return {?{markerRow:number, headerRow:number, cells:Array<{col:number, label:string}>}} null when nothing to do.
 */
function _planHeaderOnMarkerRow_(grid, snap) {
  var section = snap.layout.sections.responses;
  var table = snap.tables.responses;
  if (!section || !table || table.headerRow === null || !table.missing.length) return null;
  var missing = {};
  table.missing.forEach(function (l) { missing[_labelKey_(l)] = l; });
  var marker = grid[section.markerRow - 1] || [];
  var cells = [];
  for (var c = 2; c <= marker.length; c++) {
    var lk = _labelKey_(_gridCell_(grid, section.markerRow, c));
    if (!lk || !missing.hasOwnProperty(lk)) continue;
    if (_cellHasContent_(_gridCell_(grid, table.headerRow, c))) continue;
    cells.push({ col: c, label: missing[lk] });
    delete missing[lk];
  }
  return cells.length ? { markerRow: section.markerRow, headerRow: table.headerRow, cells: cells } : null;
}

/**
 * Writes a _planHeaderOnMarkerRow_ plan: header cells only, one setValues per contiguous
 * run of columns. The marker row is left as it is (its cols B+ are never read). Caller
 * holds the ballot lock.
 */
function _applyHeaderOnMarkerRow_(sheet, plan) {
  var cells = plan.cells.slice().sort(function (a, b) { return a.col - b.col; });
  var i = 0;
  while (i < cells.length) {
    var j = i;
    while (j + 1 < cells.length && cells[j + 1].col === cells[j].col + 1) j++;
    var run = cells.slice(i, j + 1).map(function (x) { return x.label; });
    sheet.getRange(plan.headerRow, cells[i].col, 1, run.length).setValues([run]);
    i = j + 1;
  }
  _logBallotEvent_('ballot.repair.headerOnMarkerRow', {
    sheet: sheet.getName(), row: plan.headerRow,
    cols: cells.map(function (x) { return x.col; }), labels: cells.map(function (x) { return x.label; })
  });
}

/** Legacy config keys (e.g. 'Info') to rename to their current key. Pure. */
function _planLegacyConfigKeys_(grid, snap) {
  var out = [];
  Object.keys(snap.config.byKey).forEach(function (key) {
    var e = snap.config.byKey[key];
    if (e.legacyKey) out.push({ row: e.row, from: e.legacyKey, to: key });
  });
  return out.length ? out : null;
}

function _applyLegacyConfigKeys_(sheet, plan) {
  plan.forEach(function (p) { sheet.getRange(p.row, 1).setValue(p.to); });
  _logBallotEvent_('ballot.repair.legacyConfigKey', { sheet: sheet.getName(), keys: plan });
}

/** Candidates table still on the pre-column-A-reservation layout (header 'Name' in col A). Pure. */
function _planCandidatesOldColumns_(grid, snap) {
  var t = snap.tables.candidates;
  if (!t || t.headerRow === null || !snap.layout.sections.responses) return null;
  return _labelKey_(_gridCell_(grid, t.headerRow, 1)) === _labelKey_(_BALLOT_CANDIDATES_HEADER[0]) ? { headerRow: t.headerRow } : null;
}

function _applyCandidatesOldColumns_(sheet, plan) {
  _migrateCandidatesColumnsIfNeeded_(sheet);
  _logBallotEvent_('ballot.repair.candidatesOldColumns', { sheet: sheet.getName(), row: plan.headerRow });
}

/**
 * Candidates header row missing or incomplete. Only while [Responses] exists (without it,
 * rows below [Candidates] may be orphaned votes). Pure. Kinds:
 *   insert — no room for a header (next marker directly below), or the row below the marker
 *            holds content but none of the fixed labels: it is a candidate, so a header row
 *            is inserted above it (live: SIT Ballot-ABC-2026-08);
 *   append — some fixed labels present but others missing: the missing ones are written
 *            right of the last labelled column (blank row: written from col B).
 */
function _planCandidatesHeader_(grid, snap) {
  var t = snap.tables.candidates;
  var section = snap.layout.sections.candidates;
  if (!t || !snap.layout.sections.responses || !t.missing.length) return null;
  if (t.headerRow === null) return { kind: 'insert', markerRow: section.markerRow };
  var fixedPresent = t.labels.filter(function (l) { return _BALLOT_CANDIDATES_HEADER.indexOf(l) !== -1; }).length;
  var rowHasContent = _gridRowHasContent_(grid, t.headerRow, _BALLOT_CANDIDATES_FIRST_COL, Math.max((grid[t.headerRow - 1] || []).length, _BALLOT_CANDIDATES_FIRST_COL));
  if (!fixedPresent && rowHasContent) return { kind: 'insert', markerRow: section.markerRow };
  if (!fixedPresent && t.labels.length) return null; // unrecognised labels only: leave to a human
  var last = _BALLOT_CANDIDATES_FIRST_COL - 1;
  Object.keys(t.col).forEach(function (l) { if (t.col[l] > last) last = t.col[l]; });
  return { kind: 'append', headerRow: t.headerRow, col: last + 1, labels: t.missing.slice() };
}

function _applyCandidatesHeader_(sheet, plan) {
  if (plan.kind === 'insert') {
    sheet.insertRowsAfter(plan.markerRow, 1);
    sheet.getRange(plan.markerRow + 1, _BALLOT_CANDIDATES_FIRST_COL, 1, _BALLOT_CANDIDATES_HEADER.length).setValues([_BALLOT_CANDIDATES_HEADER]);
  } else {
    sheet.getRange(plan.headerRow, plan.col, 1, plan.labels.length).setValues([plan.labels]);
  }
  _logBallotEvent_('ballot.repair.candidatesHeader', { sheet: sheet.getName(), kind: plan.kind, row: plan.kind === 'insert' ? plan.markerRow + 1 : plan.headerRow, labels: plan.labels || _BALLOT_CANDIDATES_HEADER });
}

/**
 * A Candidates table with no content below its header while the Responses header has rank
 * columns: a ballot from before the Candidates table existed. Each rank-column label
 * becomes a Candidates row (blank details). Not applied once the table has any row, so a
 * hand-deleted candidate stays deleted (its column is an orphan). Pure.
 */
function _planCandidatesFromResponses_(grid, snap) {
  var t = snap.tables.candidates;
  if (!t || t.headerRow === null || !t.col[_BALLOT_SECTIONS_.candidates.key] || t.rowsUsed !== t.headerRow) return null;
  if (!snap.tables.responses || _responsesMissingName_(snap)) return null;
  var rankCols = _snapRankColumns_(snap);
  var labels = Object.keys(rankCols).map(function (k) { return rankCols[k]; })
    .sort(function (a, b) { return a.col - b.col; }).map(function (r) { return r.label; });
  if (!labels.length) return null;
  var section = snap.layout.sections.candidates;
  return { headerRow: t.headerRow, nameCol: t.col[_BALLOT_SECTIONS_.candidates.key], available: section.endRow - t.headerRow, names: labels };
}

function _applyCandidatesFromResponses_(sheet, plan) {
  var n = plan.names.length;
  if (n > plan.available) sheet.insertRowsAfter(plan.headerRow + plan.available, n - plan.available);
  sheet.getRange(plan.headerRow + 1, plan.nameCol, n, 1).setValues(plan.names.map(function (x) { return [x]; }));
  _logBallotEvent_('ballot.repair.candidatesFromResponses', { sheet: sheet.getName(), row: plan.headerRow + 1, count: n });
}

/** Candidates with no Responses rank column: one is appended for each. Pure. */
function _planRankColumns_(grid, snap) {
  var t = snap.tables.responses;
  if (!t || t.headerRow === null || _responsesMissingName_(snap)) return null;
  var seen = {};
  var labels = [];
  _snapCandidates_(snap).forEach(function (c) {
    var k = _labelKey_(c.name);
    if (c.rankCol !== null || seen[k]) return;
    seen[k] = true;
    labels.push(c.name);
  });
  if (!labels.length) return null;
  return { headerRow: t.headerRow, col: _nextResponsesCol_(grid, snap), labels: labels };
}

function _applyRankColumns_(sheet, plan) {
  sheet.getRange(plan.headerRow, plan.col, 1, plan.labels.length).setValues([plan.labels]);
  _logBallotEvent_('ballot.repair.rankColumns', { sheet: sheet.getName(), row: plan.headerRow, col: plan.col, labels: plan.labels });
}

/**
 * Repair steps in the order they run. Each has a pure `plan(grid, snap)` (null = nothing to
 * do) and an `apply(sheet, plan)` that writes under the ballot lock. The repair path is
 * the ONLY place these run (rcballot-8i5.3): never from a reader.
 */
var _BALLOT_REPAIR_STEPS_ = [
  {
    name: 'sections',
    plan: function (grid, snap) { return _ballotMarkersComplete_(snap.layout) ? null : true; },
    apply: function (sheet) { _ensureBallotSections_(sheet); }
  },
  { name: 'legacyConfigKey', plan: _planLegacyConfigKeys_, apply: _applyLegacyConfigKeys_ },
  { name: 'candidatesOldColumns', plan: _planCandidatesOldColumns_, apply: _applyCandidatesOldColumns_ },
  { name: 'candidatesHeader', plan: _planCandidatesHeader_, apply: _applyCandidatesHeader_ },
  { name: 'headerOnMarkerRow', plan: _planHeaderOnMarkerRow_, apply: _applyHeaderOnMarkerRow_ },
  { name: 'candidatesFromResponses', plan: _planCandidatesFromResponses_, apply: _applyCandidatesFromResponses_ },
  { name: 'rankColumns', plan: _planRankColumns_, apply: _applyRankColumns_ }
];

/**
 * The locked repair path: runs every repair step a ballot sheet needs. Healthy sheet: one
 * getDataRange read, no lock, no write. Otherwise, under the lock, each step re-plans
 * from a fresh read and applies; each write is logged as ballot.repair.<kind>. Called only
 * from admin edit-page load (getAdminEditData) and the write paths.
 *
 * @param {Sheet} sheet
 * @return {{repaired:boolean, snap:Object, grid:Array<Array<*>>}} reflecting the sheet after any repair.
 */
function _repairBallotSheet_(sheet) {
  var first = _readBallotGridAndSnap_(sheet);
  var needed = _BALLOT_REPAIR_STEPS_.some(function (step) { return step.plan(first.grid, first.snap) !== null; });
  if (!needed) return { repaired: false, snap: first.snap, grid: first.grid };
  return _withBallotLock_(function () {
    var wrote = false;
    _BALLOT_REPAIR_STEPS_.forEach(function (step) {
      var cur = _readBallotGridAndSnap_(sheet);
      var plan = step.plan(cur.grid, cur.snap);
      if (plan === null) return;
      step.apply(sheet, plan);
      wrote = true;
    });
    var after = _readBallotGridAndSnap_(sheet);
    return { repaired: wrote, snap: after.snap, grid: after.grid };
  });
}

/** @return {{grid:Array<Array<*>>, snap:Object}} one getDataRange read and its parse. */
function _readBallotGridAndSnap_(sheet) {
  var grid = sheet.getDataRange().getValues();
  return { grid: grid, snap: _parseBallot_(grid) };
}

/**
 * @param {Object} snap
 * @return {boolean} true when a [Responses] section exists but its header has no 'Name'
 *   label — no vote can be keyed, so reading it would silently report 0 votes.
 */
function _responsesMissingName_(snap) {
  var t = snap.tables.responses;
  return !!t && t.missing.indexOf(_BALLOT_SECTIONS_.responses.key) !== -1;
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
    _ensureBallotSections_: _ensureBallotSections_,
    _ensureResponsesSection_: _ensureResponsesSection_,
    _migrateCandidatesColumnsIfNeeded_: _migrateCandidatesColumnsIfNeeded_,
    _ensureCandidatesHeaderColumns_: _ensureCandidatesHeaderColumns_,
    _withBallotLock_: _withBallotLock_,
    _findMarkerRow_: _findMarkerRow_,
    _readBallotLayout_: _readBallotLayout_,
    _scanBallotLayout_: _scanBallotLayout_,
    _BALLOT_SECTIONS_: _BALLOT_SECTIONS_,
    _parseBallot_: _parseBallot_,
    _readBallot_: _readBallot_,
    _snapCandidates_: _snapCandidates_,
    _snapResponses_: _snapResponses_,
    _snapOrphanColumns_: _snapOrphanColumns_,
    _snapMissingLabels_: _snapMissingLabels_,
    _planHeaderOnMarkerRow_: _planHeaderOnMarkerRow_,
    _repairBallotSheet_: _repairBallotSheet_,
    _responsesMissingName_: _responsesMissingName_
  };
}
