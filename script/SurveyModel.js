/**
 * SurveyModel.js
 *
 * Sheet-layout model for "Survey-<name>" sheets. Each survey lives entirely in
 * one sheet, top to bottom:
 *
 *   Row 1..6   Config (key in col A, value in col B): Title, Description,
 *              Footer, Contact, Accept-New, Info
 *   Row 7      blank
 *   Row 8      "[Results]" marker (col A)
 *   Row 9+     Analysis output — cleared and rewritten by runSurveyAnalysis_
 *   Row M      "[Items]" marker (col A)
 *   Row M+1    Items header: (col A blank) Name (col B) | Details (col C)
 *   Row M+2+   One row per item — name/details start at col B, position-aligned
 *              with the Responses header's candidate columns (item i's row here
 *              == candidate column i)
 *   Row N      "[Responses]" marker (col A)
 *   Row N+1    Responses header: Date | Name | Weight | Comment | <candidate1> | <candidate2> | ...
 *   Row N+2+   One row per respondent
 *
 * Column A is reserved EXCLUSIVELY for section markers/config keys — never for
 * user-entered data — and markers are bracket-decorated ("[Results]", not "Results")
 * as a second, independent guard. Both exist because _findMarkerRow_ does an
 * unscoped top-to-bottom scan of column A for exact marker text: if an item's name
 * (user-controlled, arbitrary) ever landed in column A, naming an item e.g.
 * "Responses" would shadow the real marker and corrupt every section boundary
 * computed from it (this was a real bug in an earlier version of this file, where
 * the Items table's Name column started at col A). Reserving col A for markers only
 * closes the hole entirely; the bracket decoration is defense in depth against any
 * future mistake (or a sheet hand-edited outside the app) reintroducing it.
 *
 * "Info" is admin-only free-text notes about the survey (e.g. purpose, audience,
 * scheduling context) — shown in the admin list as a per-row disclosure, never
 * exposed to respondents (see getSurveyForRespondent_, which omits it).
 *
 * Candidates ("items") are the Responses header columns from E onward — that header
 * row remains the authoritative source for RCV/Condorcet analysis (column position is
 * what response rows key ranks off of). The Items table is a second, position-aligned
 * view of the same list that adds a per-item "Details" note editable from the admin
 * page; it is kept in sync by saveSurveyItems_/addSurveyCandidate_, not by hand-editing.
 * A survey created before Items existed has no Items marker — _ensureItemsSection_
 * inserts it (empty) the first time anything here needs it, so older sheets self-heal
 * rather than throwing.
 *
 * webSurvey.js (respondent UI) and webAdmin.js (analysis) both read/write
 * through this module so the sheet layout only needs to be understood once.
 */

var SURVEY_SHEET_PREFIX = 'Survey-';
var _SURVEY_RESULTS_MARKER = '[Results]';
var _SURVEY_ITEMS_MARKER = '[Items]';
var _SURVEY_ITEMS_HEADER = ['Name', 'Details'];
var _SURVEY_ITEMS_FIRST_COL = 2; // col B — col A stays blank/reserved throughout the Items section
var _SURVEY_RESPONSES_MARKER = '[Responses]';
// Bare (undecorated) marker text used before markers were bracket-decorated. Recognized
// by _findMarkerRow_ purely to migrate older sheets in place — the moment one is found,
// the cell is rewritten to the decorated form, so every later scan matches directly.
var _SURVEY_LEGACY_MARKER_TEXT_ = { '[Results]': 'Results', '[Items]': 'Items', '[Responses]': 'Responses' };
var _SURVEY_CONFIG_ROWS = ['Title', 'Description', 'Footer', 'Contact', 'Accept-New', 'Info'];
// Fixed Responses columns before the candidate columns begin (col 5 / E).
var _SURVEY_RESPONSE_FIXED_COLS = ['Date', 'Name', 'Weight', 'Comment'];
var _SURVEY_FIRST_CANDIDATE_COL = _SURVEY_RESPONSE_FIXED_COLS.length + 1; // 5
var _SURVEY_LOCK_TIMEOUT_MS = 30000;
// Bold + light-blue background applied to every section marker/header row so the
// sheet's structure (Results / Items / Responses) is easy to scan visually.
var _SURVEY_SECTION_HIGHLIGHT_BG = '#e8f0fe';

/**
 * @param {string} id
 * @return {string} sheet name, e.g. "Survey-BoardElection2026"
 */
function getSurveySheetName_(id) {
  return SURVEY_SHEET_PREFIX + String(id || '').trim();
}

/**
 * Lists survey ids (sheet name with the "Survey-" prefix stripped) for every
 * matching sheet in the spreadsheet, in sheet order.
 *
 * @param {Spreadsheet} ss
 * @return {Array<string>}
 */
function listSurveyIds_(ss) {
  return ss.getSheets()
    .map(function (s) { return s.getName(); })
    .filter(function (name) { return name.indexOf(SURVEY_SHEET_PREFIX) === 0; })
    .map(function (name) { return name.substring(SURVEY_SHEET_PREFIX.length); });
}

/**
 * Returns the sheet for a survey id, or null if it does not exist.
 *
 * @param {Spreadsheet} ss
 * @param {string} id
 * @return {Sheet|null}
 */
function findSurveySheet_(ss, id) {
  return ss.getSheetByName(getSurveySheetName_(id));
}

/**
 * Creates a new Survey-<id> sheet with placeholder configuration and an
 * empty Results/Responses skeleton. Placeholder values start with "[TODO"
 * so they are easy to spot and grep for.
 *
 * @param {Spreadsheet} ss
 * @param {string} id
 * @return {Sheet}
 */
function createSurveySheet_(ss, id) {
  var sheet = ss.insertSheet(getSurveySheetName_(id));
  sheet.getRange(1, 1, 6, 2).setValues([
    ['Title', '[TODO: survey title shown to respondents]'],
    ['Description', '[TODO: description/instructions shown above the ranking list]'],
    ['Footer', '[TODO: footer text, e.g. deadline or sponsoring group]'],
    ['Contact', '[TODO: contact name/email for questions]'],
    ['Accept-New', 'TRUE'],
    ['Info', '']
  ]);
  sheet.getRange(8, 1).setValue(_SURVEY_RESULTS_MARKER);
  sheet.getRange(9, 1).setValue(_SURVEY_ITEMS_MARKER);
  sheet.getRange(10, _SURVEY_ITEMS_FIRST_COL, 1, _SURVEY_ITEMS_HEADER.length).setValues([_SURVEY_ITEMS_HEADER]);
  sheet.getRange(11, 1).setValue(_SURVEY_RESPONSES_MARKER);
  sheet.getRange(12, 1, 1, _SURVEY_RESPONSE_FIXED_COLS.length).setValues([_SURVEY_RESPONSE_FIXED_COLS]);
  sheet.setFrozenRows(12);
  _highlightSectionMarkers_(sheet);
  return sheet;
}

/**
 * Bolds + colors every section marker cell (Results/Items/Responses) and the Items/
 * Responses header rows, so the sheet's structure is easy to pick out visually.
 * Idempotent and safe to call repeatedly (e.g. every time the admin edit page loads),
 * which is how older sheets created before this existed pick up the highlighting.
 *
 * @param {Sheet} sheet
 */
function _highlightSectionMarkers_(sheet) {
  function highlightRow(row, numCols) {
    if (row === -1) return;
    sheet.getRange(row, 1, 1, numCols || 1).setBackground(_SURVEY_SECTION_HIGHLIGHT_BG).setFontWeight('bold');
  }

  var resultsRow = _findMarkerRow_(sheet, _SURVEY_RESULTS_MARKER);
  var itemsRow = _findMarkerRow_(sheet, _SURVEY_ITEMS_MARKER);
  var responsesRow = _findMarkerRow_(sheet, _SURVEY_RESPONSES_MARKER);

  highlightRow(resultsRow, 1);
  highlightRow(itemsRow, 1);
  // Highlight col A through the end of the Name/Details header as one continuous band,
  // even though col A itself is blank in the Items section (reserved — see file header).
  if (itemsRow !== -1) highlightRow(itemsRow + 1, _SURVEY_ITEMS_FIRST_COL - 1 + _SURVEY_ITEMS_HEADER.length);
  highlightRow(responsesRow, 1);
  if (responsesRow !== -1) {
    var lastCol = Math.max(sheet.getLastColumn(), _SURVEY_RESPONSE_FIXED_COLS.length);
    highlightRow(responsesRow + 1, lastCol);
  }
}

/**
 * Inserts an empty Items marker + header row directly above the Responses marker if
 * this sheet doesn't have one yet (a survey created before the Items section existed).
 * No-op if the sheet already has an Items marker, or has no Responses marker to anchor
 * the insert to (shouldn't happen for any sheet created via createSurveySheet_).
 *
 * @param {Sheet} sheet
 */
function _ensureItemsSection_(sheet) {
  if (_findMarkerRow_(sheet, _SURVEY_ITEMS_MARKER) !== -1) return;
  var responsesRow = _findMarkerRow_(sheet, _SURVEY_RESPONSES_MARKER);
  if (responsesRow === -1) return;
  sheet.insertRowsBefore(responsesRow, 2);
  sheet.getRange(responsesRow, 1).setValue(_SURVEY_ITEMS_MARKER);
  sheet.getRange(responsesRow + 1, _SURVEY_ITEMS_FIRST_COL, 1, _SURVEY_ITEMS_HEADER.length).setValues([_SURVEY_ITEMS_HEADER]);
}

/**
 * Migrates an Items section still using the OLD column layout (Name/Details in
 * columns A/B, from before column A was reserved for markers only) to the current
 * layout (columns B/C, col A left blank). Detected by the Items header's column A
 * cell still holding the literal "Name" label — the current layout never writes
 * anything to column A there. A no-op for a sheet already on the new layout (or
 * with no Items section at all — call _ensureItemsSection_ first).
 *
 * Without this, readSurveyItemDetails_ silently reads garbage (the OLD Details
 * column mistaken for the new Name column) rather than throwing, because both
 * layouts are structurally valid 2-column reads — this was found live on a survey
 * created before the column-A-reservation fix shipped, where every item's Details
 * came back empty even though the sheet clearly had them.
 *
 * @param {Sheet} sheet
 */
function _migrateItemsColumnsIfNeeded_(sheet) {
  var itemsRow = _findMarkerRow_(sheet, _SURVEY_ITEMS_MARKER);
  if (itemsRow === -1) return;
  var itemsHeaderRow = itemsRow + 1;
  var headerColA = String(sheet.getRange(itemsHeaderRow, 1).getValue() || '').trim();
  if (headerColA !== _SURVEY_ITEMS_HEADER[0]) return; // already migrated (col A blank) or unrecognized

  var responsesRow = _findMarkerRow_(sheet, _SURVEY_RESPONSES_MARKER);
  if (responsesRow === -1) return;
  var count = responsesRow - itemsHeaderRow - 1;
  var oldValues = count > 0 ? sheet.getRange(itemsHeaderRow + 1, 1, count, 2).getValues() : [];

  sheet.getRange(itemsHeaderRow, 1, count + 1, 1).clearContent(); // vacate col A (header + data rows)
  sheet.getRange(itemsHeaderRow, _SURVEY_ITEMS_FIRST_COL, 1, _SURVEY_ITEMS_HEADER.length).setValues([_SURVEY_ITEMS_HEADER]);
  if (count > 0) {
    sheet.getRange(itemsHeaderRow + 1, _SURVEY_ITEMS_FIRST_COL, count, 2).setValues(oldValues);
  }
}

/**
 * Inserts/deletes rows so exactly rows.length rows are available directly below
 * `afterRow`, ending just before `beforeMarkerRow`, then writes rows into that space.
 * Shared resize primitive behind both the Results section (writeSurveyResults_) and
 * the Items section (writeSurveyItemDetails_) — everything at/after beforeMarkerRow
 * shifts as one block, so content further down the sheet is never disturbed.
 *
 * @param {Sheet} sheet
 * @param {number} afterRow
 * @param {number} beforeMarkerRow
 * @param {Array<Array>} rows
 * @param {number=} startCol defaults to col A (1) — Items-section callers pass
 *   _SURVEY_ITEMS_FIRST_COL so item data never lands in the marker-reserved column A.
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
 * Returns the sheet for a survey id, creating a placeholder skeleton if it
 * does not already exist.
 *
 * @param {Spreadsheet} ss
 * @param {string} id
 * @return {Sheet}
 */
function getOrCreateSurveySheet_(ss, id) {
  return findSurveySheet_(ss, id) || createSurveySheet_(ss, id);
}

/**
 * Validates a candidate survey id and creates a brand-new Survey-<id> sheet.
 * Used by the admin "Create New Survey" form (webAdmin.js). Ids are
 * restricted to characters that are safe in both a sheet name and a URL
 * query parameter without encoding.
 *
 * @param {Spreadsheet} ss
 * @param {string} id
 * @return {Sheet}
 */
function createNewSurvey_(ss, id) {
  id = String(id || '').trim();
  if (!id) throw new Error('Survey ID is required.');
  if (!/^[A-Za-z0-9][A-Za-z0-9_-]*$/.test(id)) {
    throw new Error('Survey ID may only contain letters, numbers, "-" and "_", and must start with a letter or number.');
  }
  if (findSurveySheet_(ss, id)) {
    throw new Error('A survey with id "' + id + '" already exists.');
  }
  return createSurveySheet_(ss, id);
}

/**
 * Overwrites the config key/value rows for a survey. Keys already present are
 * matched by scanning column A rather than assuming fixed row numbers, so it
 * tolerates a sheet whose config rows have been manually reordered. Any update
 * key with no existing row (e.g. a config field — like "Info" — added to the
 * model after this particular sheet was created) is appended as a new row,
 * inserted just above the Results marker, so older sheets self-heal on first
 * save instead of silently dropping the value.
 *
 * @param {Sheet} sheet
 * @param {Object} updates e.g. {Title, Description, Footer, Contact, 'Accept-New', Info}
 */
function writeSurveyConfig_(sheet, updates) {
  var resultsRow = _findMarkerRow_(sheet, _SURVEY_RESULTS_MARKER);
  var lastConfigRow = resultsRow === -1 ? _SURVEY_CONFIG_ROWS.length : resultsRow - 1;
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
  var legacyText = _SURVEY_LEGACY_MARKER_TEXT_[marker];
  for (var i = 0; i < values.length; i++) {
    var cell = String(values[i][0] || '').trim();
    if (cell === marker) return i + 1;
    if (legacyText && cell === legacyText) {
      sheet.getRange(i + 1, 1).setValue(marker); // one-time upgrade to the decorated form
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
 * @return {Object} e.g. {Title, Description, Footer, Contact, 'Accept-New'}
 */
function readSurveyConfig_(sheet) {
  var resultsRow = _findMarkerRow_(sheet, _SURVEY_RESULTS_MARKER);
  var lastConfigRow = resultsRow === -1 ? _SURVEY_CONFIG_ROWS.length : resultsRow - 1;
  var config = {};
  if (lastConfigRow < 1) return config;
  var values = sheet.getRange(1, 1, lastConfigRow, 2).getValues();
  values.forEach(function (row) {
    var key = String(row[0] || '').trim();
    if (key) config[key] = row[1];
  });
  return config;
}

/**
 * @param {Object} config as returned by readSurveyConfig_
 * @return {boolean}
 */
function _surveyAcceptsNew_(config) {
  return String(config['Accept-New'] || '').trim().toUpperCase() === 'TRUE';
}

/**
 * @param {Sheet} sheet
 * @return {number} 1-based row of the Responses header (Date/Name/Weight/Comment/...).
 */
function _getResponsesHeaderRow_(sheet) {
  var marker = _findMarkerRow_(sheet, _SURVEY_RESPONSES_MARKER);
  return marker === -1 ? -1 : marker + 1;
}

/**
 * Reads the candidate names from the Responses header row (columns E+).
 *
 * @param {Sheet} sheet
 * @return {Array<string>}
 */
function readSurveyCandidates_(sheet) {
  var headerRow = _getResponsesHeaderRow_(sheet);
  if (headerRow === -1) return [];
  var lastCol = sheet.getLastColumn();
  if (lastCol < _SURVEY_FIRST_CANDIDATE_COL) return [];
  var values = sheet.getRange(headerRow, _SURVEY_FIRST_CANDIDATE_COL, 1, lastCol - _SURVEY_FIRST_CANDIDATE_COL + 1).getValues()[0];
  return values.map(function (v) { return String(v || '').trim(); }).filter(function (v) { return v; });
}

/**
 * Reads all response rows below the Responses header.
 *
 * @param {Sheet} sheet
 * @return {Array<{row:number, date:*, name:string, weight:number, comment:string, ranks:Array}>}
 */
function readSurveyResponseRows_(sheet) {
  var headerRow = _getResponsesHeaderRow_(sheet);
  var lastRow = sheet.getLastRow();
  if (headerRow === -1 || lastRow < headerRow + 1) return [];
  var candidates = readSurveyCandidates_(sheet);
  var numCols = _SURVEY_RESPONSE_FIXED_COLS.length + candidates.length;
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
      ranks: row.slice(_SURVEY_RESPONSE_FIXED_COLS.length)
    });
  });
  return out;
}

/**
 * Finds the last response row for a respondent (case-insensitive), so a
 * re-submission overwrites their existing row instead of adding a duplicate.
 *
 * @param {Sheet} sheet
 * @param {string} name
 * @return {number} 1-based row, or -1 if not found.
 */
function _findSurveyResponseRowForName_(sheet, name) {
  var target = String(name).trim().toLowerCase();
  var rows = readSurveyResponseRows_(sheet);
  for (var i = rows.length - 1; i >= 0; i--) {
    if (rows[i].name.toLowerCase() === target) return rows[i].row;
  }
  return -1;
}

/**
 * Appends a new candidate column to the Responses header row, and a matching row
 * to the Items table (position-aligned — this candidate's Items row ends up at the
 * same index as its new column). Caller must check Accept-New first.
 *
 * @param {Sheet} sheet
 * @param {string} phrase
 * @param {string=} details admin-only note for this item (respondents never pass this).
 */
function addSurveyCandidate_(sheet, phrase, details) {
  var headerRow = _getResponsesHeaderRow_(sheet);
  if (headerRow === -1) throw new Error('Survey has no Responses section.');
  var existingCandidates = readSurveyCandidates_(sheet); // BEFORE the new column exists
  var nextCol = sheet.getLastColumn() + 1;
  if (nextCol < _SURVEY_FIRST_CANDIDATE_COL) nextCol = _SURVEY_FIRST_CANDIDATE_COL;
  sheet.getRange(headerRow, nextCol).setValue(phrase);

  var items = readSurveyItemDetails_(sheet);
  // Backfill any candidates that predate the Items table (e.g. a survey created, or a
  // candidate added, before this feature existed) with blank details, so positions
  // never drift between the Items table and the Responses header's candidate columns.
  while (items.length < existingCandidates.length) {
    items.push({ name: existingCandidates[items.length], details: '' });
  }
  items.push({ name: phrase, details: details || '' });
  writeSurveyItemDetails_(sheet, items);
}

/**
 * RPC-facing: returns everything the survey page needs to render for a
 * given respondent (or generic branding/ordering if name is blank/unknown).
 *
 * @param {string} id
 * @param {string} name
 * @return {Object|null} null if the survey id does not exist.
 */
function getSurveyForRespondent_(id, name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSurveySheet_(ss, id);
  if (!sheet) return null;

  var config = readSurveyConfig_(sheet);
  var candidates = readSurveyCandidates_(sheet);
  name = String(name || '').trim();

  // Keyed by name (not position) so it stays correct even though `candidates` below
  // gets reordered per-respondent to match their saved ranking — a candidate with no
  // Items-table entry (e.g. added by a respondent, who can't set details) simply has
  // no key here, which the client treats as "no details to show," not an error.
  var itemDetails = {};
  readSurveyItemDetails_(sheet).forEach(function (it) {
    if (it.details) itemDetails[it.name] = it.details;
  });

  var result = {
    id: id,
    title: config.Title || '',
    description: config.Description || '',
    footer: config.Footer || '',
    contact: config.Contact || '',
    acceptNew: _surveyAcceptsNew_(config),
    candidates: candidates,
    itemDetails: itemDetails,
    comment: '',
    appVersion: (typeof APP_VERSION !== 'undefined' && APP_VERSION) || '',
    appDeployTarget: (typeof APP_DEPLOY_TARGET !== 'undefined' && APP_DEPLOY_TARGET) || ''
  };
  if (!name) return result;

  var rows = readSurveyResponseRows_(sheet);
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
 * RPC-facing: appends a new candidate to a survey (if Accept-New allows it).
 * Respondents can set a Details note too (previously admin-only) — it's the same
 * underlying field, just now writable from the "Add New" panel on the survey page.
 *
 * @param {string} id
 * @param {string} phrase
 * @param {string=} details
 * @return {{candidate:string}}
 */
function addSurveyCandidateForId_(id, phrase, details) {
  phrase = String(phrase || '').trim();
  if (!phrase) throw new Error('Item cannot be empty.');

  var lock = LockService.getScriptLock();
  lock.waitLock(_SURVEY_LOCK_TIMEOUT_MS);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findSurveySheet_(ss, id);
    if (!sheet) throw new Error('No survey found for id "' + id + '".');
    var config = readSurveyConfig_(sheet);
    if (!_surveyAcceptsNew_(config)) throw new Error('This survey is not accepting new items.');
    addSurveyCandidate_(sheet, phrase, details);
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
function submitSurveyResponse_(id, name, orderedCandidates, comment) {
  name = String(name || '').trim();
  if (!name) throw new Error('Name is required.');
  if (!orderedCandidates || !orderedCandidates.length) throw new Error('Ranking cannot be empty.');

  var lock = LockService.getScriptLock();
  lock.waitLock(_SURVEY_LOCK_TIMEOUT_MS);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findSurveySheet_(ss, id);
    if (!sheet) throw new Error('No survey found for id "' + id + '".');

    var candidates = readSurveyCandidates_(sheet);
    var rankByCandidate = {};
    orderedCandidates.forEach(function (c, idx) {
      rankByCandidate[String(c).trim()] = idx + 1;
    });

    var row = _findSurveyResponseRowForName_(sheet, name);
    var weight = 1;
    if (row === -1) {
      row = sheet.getLastRow() + 1;
    } else {
      weight = Number(sheet.getRange(row, 3).getValue()) || 1;
    }

    var ranks = candidates.map(function (c) {
      return rankByCandidate.hasOwnProperty(c) ? rankByCandidate[c] : '';
    });
    sheet.getRange(row, 1, 1, _SURVEY_RESPONSE_FIXED_COLS.length + candidates.length).setValues([
      [new Date(), name, weight, String(comment || '').trim()].concat(ranks)
    ]);

    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Inserts or deletes rows so the Results section has exactly rows.length
 * rows available, then writes rows into it. Leaves the Items/Responses sections
 * (which sit immediately below) untouched — the insert/delete happens strictly
 * between the Results and Items markers, so everything below shifts as one block.
 *
 * @param {Sheet} sheet
 * @param {Array<Array>} rows
 */
function writeSurveyResults_(sheet, rows) {
  _ensureItemsSection_(sheet);
  var resultsRow = _findMarkerRow_(sheet, _SURVEY_RESULTS_MARKER);
  var itemsRow = _findMarkerRow_(sheet, _SURVEY_ITEMS_MARKER);
  if (resultsRow === -1 || itemsRow === -1) {
    throw new Error('Survey sheet is missing its Results/Items markers.');
  }
  _resizeSectionRows_(sheet, resultsRow, itemsRow, rows);
}

/**
 * Reads the Items table (Name/Details per row), position-aligned with the Responses
 * header's candidate columns — item i corresponds to candidate column i.
 *
 * @param {Sheet} sheet
 * @return {Array<{name:string, details:string}>}
 */
function readSurveyItemDetails_(sheet) {
  _ensureItemsSection_(sheet);
  _migrateItemsColumnsIfNeeded_(sheet);
  var itemsRow = _findMarkerRow_(sheet, _SURVEY_ITEMS_MARKER);
  var responsesRow = _findMarkerRow_(sheet, _SURVEY_RESPONSES_MARKER);
  if (itemsRow === -1 || responsesRow === -1) return [];
  var itemsHeaderRow = itemsRow + 1;
  var count = responsesRow - itemsHeaderRow - 1;
  if (count <= 0) return [];
  var values = sheet.getRange(itemsHeaderRow + 1, _SURVEY_ITEMS_FIRST_COL, count, 2).getValues();
  return values.map(function (r) {
    return { name: String(r[0] || '').trim(), details: String(r[1] || '') };
  });
}

/**
 * Resizes + rewrites the Items table to exactly match `items` (array of
 * {name, details}, one per candidate column, in the same order). Leaves the
 * Responses section untouched — the insert/delete happens strictly between the
 * Items header and the Responses marker.
 *
 * @param {Sheet} sheet
 * @param {Array<{name:string, details:string}>} items
 */
function writeSurveyItemDetails_(sheet, items) {
  _ensureItemsSection_(sheet);
  _migrateItemsColumnsIfNeeded_(sheet);
  var itemsRow = _findMarkerRow_(sheet, _SURVEY_ITEMS_MARKER);
  var responsesRow = _findMarkerRow_(sheet, _SURVEY_RESPONSES_MARKER);
  if (itemsRow === -1 || responsesRow === -1) {
    throw new Error('Survey sheet is missing its Items/Responses markers.');
  }
  var rows = items.map(function (it) { return [it.name || '', it.details || '']; });
  _resizeSectionRows_(sheet, itemsRow + 1, responsesRow, rows, _SURVEY_ITEMS_FIRST_COL);
}

/**
 * Admin-facing: overwrites every candidate's name + details in one pass, given a
 * full ordered list matching the survey's current candidate positions (position,
 * not name, is the join key — renaming a candidate never touches response data,
 * since ranks are keyed by column index). Throws if the count doesn't match the
 * survey's current candidate count (stale form — caller should reload and retry).
 *
 * @param {Sheet} sheet
 * @param {Array<{name:string, details:string}>} items
 */
function saveSurveyItems_(sheet, items) {
  var candidates = readSurveyCandidates_(sheet);
  if (items.length !== candidates.length) {
    throw new Error('Item count changed since the page was loaded — reload and try again.');
  }
  var names = items.map(function (it) {
    var name = String(it.name || '').trim();
    if (!name) throw new Error('Item name cannot be empty.');
    return name;
  });
  var headerRow = _getResponsesHeaderRow_(sheet);
  if (names.length) {
    sheet.getRange(headerRow, _SURVEY_FIRST_CANDIDATE_COL, 1, names.length).setValues([names]);
  }
  writeSurveyItemDetails_(sheet, items.map(function (it, i) {
    return { name: names[i], details: String(it.details || '') };
  }));
  _highlightSectionMarkers_(sheet);
}

/**
 * RPC-facing: locates the survey by id and applies saveSurveyItems_ under the
 * script lock (mirrors submitSurveyResponse_/addSurveyCandidateForId_'s locking —
 * admin edits and respondent submissions/candidate-adds must not interleave).
 *
 * @param {string} id
 * @param {Array<{name:string, details:string}>} items
 */
function saveSurveyItemsForId_(id, items) {
  var lock = LockService.getScriptLock();
  lock.waitLock(_SURVEY_LOCK_TIMEOUT_MS);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findSurveySheet_(ss, id);
    if (!sheet) throw new Error('No survey found for id "' + id + '".');
    saveSurveyItems_(sheet, items);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Admin-facing: appends a new candidate WITH a details note in one call (the
 * respondent-facing addSurveyCandidateForId_ has no details param — Details is
 * admin-only). Locked the same way as addSurveyCandidateForId_.
 *
 * @param {string} id
 * @param {string} name
 * @param {string=} details
 * @return {{candidate:string}}
 */
function addSurveyItemForAdmin_(id, name, details) {
  name = String(name || '').trim();
  if (!name) throw new Error('Item name cannot be empty.');

  var lock = LockService.getScriptLock();
  lock.waitLock(_SURVEY_LOCK_TIMEOUT_MS);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findSurveySheet_(ss, id);
    if (!sheet) throw new Error('No survey found for id "' + id + '".');
    addSurveyCandidate_(sheet, name, details);
    _highlightSectionMarkers_(sheet);
    return { candidate: name };
  } finally {
    lock.releaseLock();
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    getSurveySheetName_: getSurveySheetName_,
    listSurveyIds_: listSurveyIds_,
    findSurveySheet_: findSurveySheet_,
    createSurveySheet_: createSurveySheet_,
    getOrCreateSurveySheet_: getOrCreateSurveySheet_,
    createNewSurvey_: createNewSurvey_,
    writeSurveyConfig_: writeSurveyConfig_,
    readSurveyConfig_: readSurveyConfig_,
    readSurveyCandidates_: readSurveyCandidates_,
    readSurveyResponseRows_: readSurveyResponseRows_,
    addSurveyCandidate_: addSurveyCandidate_,
    getSurveyForRespondent_: getSurveyForRespondent_,
    addSurveyCandidateForId_: addSurveyCandidateForId_,
    submitSurveyResponse_: submitSurveyResponse_,
    writeSurveyResults_: writeSurveyResults_,
    readSurveyItemDetails_: readSurveyItemDetails_,
    writeSurveyItemDetails_: writeSurveyItemDetails_,
    saveSurveyItems_: saveSurveyItems_,
    saveSurveyItemsForId_: saveSurveyItemsForId_,
    addSurveyItemForAdmin_: addSurveyItemForAdmin_,
    _highlightSectionMarkers_: _highlightSectionMarkers_,
    _ensureItemsSection_: _ensureItemsSection_,
    _migrateItemsColumnsIfNeeded_: _migrateItemsColumnsIfNeeded_
  };
}
