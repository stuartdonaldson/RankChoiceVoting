'use strict';

const assert = require('assert');
const vm = require('vm');
const fs = require('fs');
const path = require('path');

const { createFakeSpreadsheet, createFakeGasGlobals } = require('./fakeGas');

/** Loads script/SurveyModel.js into a fresh vm context wired to a fresh FakeSpreadsheet. */
function loadSurveyModel() {
  const ss = createFakeSpreadsheet();
  const sandbox = Object.assign({ module: { exports: {} } }, createFakeGasGlobals(ss));
  vm.createContext(sandbox);
  const src = fs.readFileSync(path.join(__dirname, '..', 'script', 'SurveyModel.js'), 'utf8');
  vm.runInContext(src, sandbox, { filename: 'SurveyModel.js' });
  return { SM: sandbox.module.exports, ss };
}

function testCreateNewSurveyCreatesSkeleton() {
  const { SM, ss } = loadSurveyModel();
  const sheet = SM.createNewSurvey_(ss, 'Test123');

  assert.equal(sheet.getName(), 'Survey-Test123');

  const config = SM.readSurveyConfig_(sheet);
  assert.equal(config.Title, '[TODO: survey title shown to respondents]');
  assert.equal(config.Description, '[TODO: intro text shown before respondents enter their name]');
  assert.equal(config.Instructions, '[TODO: instructions shown above the ranking list on the ballot page]');
  assert.equal(config.Footer, '[TODO: footer text, e.g. deadline or sponsoring group]');
  assert.equal(config.Contact, '[TODO: contact name/email for questions]');
  assert.equal(config['Accept-New'], 'TRUE');
  assert.equal(config['Add-Instructions'], '[TODO: instructions shown above the "+ Add New" button, if enabled]');
  assert.equal(config['Admin-Only-Notes'], '');

  // Results marker at row 10, Candidates marker/header at 11/12, Responses marker at row
  // 13, Responses header at row 14. Markers are bracket-decorated ("[Results]", not
  // "Results") so a survey candidate literally named "Results"/"Candidates"/"Responses"
  // can never be mistaken for the real marker; the Candidates header (row 12) starts
  // at col B (col A stays blank/reserved) so candidate data is never in the
  // marker-scanned column.
  assert.equal(sheet.getRange(10, 1).getValue(), '[Results]');
  assert.equal(sheet.getRange(11, 1).getValue(), '[Candidates]');
  assert.equal(sheet.getRange(12, 1).getValue(), ''); // col A reserved, always blank here
  assert.deepEqual(sheet.getRange(12, 2, 1, 2).getValues()[0], ['Name', 'Details']);
  assert.equal(sheet.getRange(13, 1).getValue(), '[Responses]');
  assert.deepEqual(sheet.getRange(14, 1, 1, 4).getValues()[0], ['Date', 'Name', 'Weight', 'Comment']);
}

function testListSurveyIdsAndFindSurveySheet() {
  const { SM, ss } = loadSurveyModel();
  SM.createNewSurvey_(ss, 'Alpha');
  SM.createNewSurvey_(ss, 'Beta');

  assert.deepEqual(SM.listSurveyIds_(ss), ['Alpha', 'Beta']);
  assert.ok(SM.findSurveySheet_(ss, 'Alpha'));
  assert.equal(SM.findSurveySheet_(ss, 'Nope'), null);
}

function testWriteThenReadConfigRoundTrips() {
  const { SM, ss } = loadSurveyModel();
  const sheet = SM.createNewSurvey_(ss, 'RoundTrip');

  SM.writeSurveyConfig_(sheet, {
    Title: 'Board Election 2026',
    Description: 'Rank your top three.',
    Footer: 'Closes Friday',
    Contact: 'admin@example.com',
    'Accept-New': 'FALSE',
  });

  const config = SM.readSurveyConfig_(sheet);
  assert.equal(config.Title, 'Board Election 2026');
  assert.equal(config.Description, 'Rank your top three.');
  assert.equal(config.Footer, 'Closes Friday');
  assert.equal(config.Contact, 'admin@example.com');
  assert.equal(config['Accept-New'], 'FALSE');
}

/**
 * Simulates a sheet created by a PRE-"Admin-Only-Notes" version of createSurveySheet_
 * (5 config rows, a blank spacer at row 6, Results at row 7, Responses at row 9) to
 * confirm writeSurveyConfig_ reuses that blank spacer row for "Admin-Only-Notes" on
 * first save rather than silently dropping the value (or needlessly inserting a new
 * row and shifting Results/Responses).
 */
function testWriteSurveyConfigSelfHealsMissingAdminOnlyNotesRowOnOldSheet() {
  const { SM, ss } = loadSurveyModel();
  const sheet = ss.insertSheet('Survey-OldFormat');
  sheet.getRange(1, 1, 5, 2).setValues([
    ['Title', 'Old Title'],
    ['Description', 'Old Description'],
    ['Footer', 'Old Footer'],
    ['Contact', 'old@example.com'],
    ['Accept-New', 'TRUE'],
  ]);
  sheet.getRange(7, 1).setValue('Results');
  sheet.getRange(9, 1).setValue('Responses');
  sheet.getRange(10, 1, 1, 4).setValues([['Date', 'Name', 'Weight', 'Comment']]);

  SM.writeSurveyConfig_(sheet, {
    Title: 'Old Title',
    Description: 'Old Description',
    Footer: 'Old Footer',
    Contact: 'old@example.com',
    'Accept-New': 'TRUE',
    'Admin-Only-Notes': 'newly added notes',
  });

  const config = SM.readSurveyConfig_(sheet);
  assert.equal(config['Admin-Only-Notes'], 'newly added notes');
  assert.equal(config.Title, 'Old Title'); // pre-existing keys untouched

  // Admin-Only-Notes reused the existing blank spacer row (6) — Results/Responses/
  // header stay put. The bare "Results" text this test seeded gets migrated to
  // "[Results]" in place the moment writeSurveyConfig_ looks it up (via
  // _findMarkerRow_); "Responses" is never looked up in this test, so it's left
  // as-is — legacy text is only ever migrated on actual use, not proactively.
  assert.equal(sheet.getRange(6, 1).getValue(), 'Admin-Only-Notes');
  assert.equal(sheet.getRange(7, 1).getValue(), '[Results]');
  assert.equal(sheet.getRange(9, 1).getValue(), 'Responses');
  assert.deepEqual(sheet.getRange(10, 1, 1, 4).getValues()[0], ['Date', 'Name', 'Weight', 'Comment']);
}

/**
 * Same self-heal scenario, but with NO spare blank row above Results (config rows
 * butt directly against the marker) — exercises the insert-a-new-row fallback path.
 */
function testWriteSurveyConfigInsertsRowWhenNoBlankSpacerExists() {
  const { SM, ss } = loadSurveyModel();
  const sheet = ss.insertSheet('Survey-NoSpacer');
  sheet.getRange(1, 1, 4, 2).setValues([
    ['Title', 'T'],
    ['Description', 'D'],
    ['Footer', 'F'],
    ['Contact', 'C'],
  ]);
  sheet.getRange(5, 1).setValue('Results'); // no blank row between config and marker
  sheet.getRange(7, 1).setValue('Responses');
  sheet.getRange(8, 1, 1, 4).setValues([['Date', 'Name', 'Weight', 'Comment']]);

  SM.writeSurveyConfig_(sheet, { Title: 'T', 'Admin-Only-Notes': 'notes' });

  const config = SM.readSurveyConfig_(sheet);
  assert.equal(config['Admin-Only-Notes'], 'notes');

  // A new row was inserted at row 5 (just above Results), shifting Results/Responses/
  // header down by one row each. Results migrates to "[Results]" (looked up by
  // writeSurveyConfig_); Responses is never looked up here, so stays bare.
  assert.equal(sheet.getRange(5, 1).getValue(), 'Admin-Only-Notes');
  assert.equal(sheet.getRange(6, 1).getValue(), '[Results]');
  assert.equal(sheet.getRange(8, 1).getValue(), 'Responses');
  assert.deepEqual(sheet.getRange(9, 1, 1, 4).getValues()[0], ['Date', 'Name', 'Weight', 'Comment']);
}

/**
 * Confirms the old "Info" config key is recognized as a legacy alias for
 * "Admin-Only-Notes" and migrated to the new key name in place the first time
 * readSurveyConfig_ scans it — covers a sheet that hasn't yet been touched since
 * the Info->Admin-Only-Notes rename.
 */
function testLegacyInfoKeyMigratesToAdminOnlyNotes() {
  const { SM, ss } = loadSurveyModel();
  const sheet = ss.insertSheet('Survey-LegacyInfo');
  sheet.getRange(1, 1, 6, 2).setValues([
    ['Title', 'T'], ['Description', 'D'], ['Footer', 'F'],
    ['Contact', 'C'], ['Accept-New', 'TRUE'], ['Info', 'old notes'],
  ]);
  sheet.getRange(8, 1).setValue('[Results]');

  const config = SM.readSurveyConfig_(sheet);
  assert.equal(config['Admin-Only-Notes'], 'old notes');
  assert.equal(config.Info, undefined);
  assert.equal(sheet.getRange(6, 1).getValue(), 'Admin-Only-Notes'); // migrated in place
}

function testCandidatesAndSubmitResponseRoundTrip() {
  const { SM, ss } = loadSurveyModel();
  const sheet = SM.createNewSurvey_(ss, 'Candidates');

  SM.addSurveyCandidate_(sheet, 'Alice');
  SM.addSurveyCandidate_(sheet, 'Bob');
  SM.addSurveyCandidate_(sheet, 'Carol');

  assert.deepEqual(SM.readSurveyCandidates_(sheet), ['Alice', 'Bob', 'Carol']);

  // submitSurveyResponse_ resolves the survey via SpreadsheetApp.getActiveSpreadsheet(),
  // which fakeGas wires to the same `ss` instance passed to createNewSurvey_ above.
  const result = SM.submitSurveyResponse_('Candidates', 'Voter One', ['Bob', 'Carol', 'Alice'], 'my comment');
  assert.deepEqual(result, { ok: true });

  const rows = SM.readSurveyResponseRows_(sheet);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].name, 'Voter One');
  assert.equal(rows[0].weight, 1);
  assert.equal(rows[0].comment, 'my comment');
  // ranks are positional per candidate column: Alice=3, Bob=1, Carol=2
  assert.deepEqual(rows[0].ranks, [3, 1, 2]);

  // Re-submitting for the same voter (case-insensitive) overwrites the row rather than
  // appending a new one, and preserves their existing weight.
  SM.submitSurveyResponse_('Candidates', 'voter one', ['Alice', 'Bob', 'Carol'], 'updated');
  const rows2 = SM.readSurveyResponseRows_(sheet);
  assert.equal(rows2.length, 1);
  assert.equal(rows2[0].comment, 'updated');
  assert.deepEqual(rows2[0].ranks, [1, 2, 3]);
}

/**
 * countUniqueSurveyRespondents_ counts distinct respondents (case-insensitive name),
 * collapsing duplicate rows for the same person down to one. Normal re-submission
 * never creates a duplicate row (submitSurveyResponse_ overwrites in place), so this
 * simulates the one way a duplicate can still appear: a hand-edited sheet or legacy
 * data imported outside the app.
 */
function testCountUniqueSurveyRespondents() {
  const { SM, ss } = loadSurveyModel();
  const sheet = SM.createNewSurvey_(ss, 'DupeCheck');

  SM.addSurveyCandidate_(sheet, 'Alice');
  SM.addSurveyCandidate_(sheet, 'Bob');
  SM.submitSurveyResponse_('DupeCheck', 'Voter One', ['Alice', 'Bob'], '');
  SM.submitSurveyResponse_('DupeCheck', 'Voter Two', ['Bob', 'Alice'], '');
  assert.equal(SM.readSurveyResponseRows_(sheet).length, 2);
  assert.equal(SM.countUniqueSurveyRespondents_(sheet), 2);

  // Manually append a second row for "voter one" (different case), simulating a
  // hand-edited duplicate that submitSurveyResponse_'s own overwrite-in-place logic
  // would never produce.
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, 6).setValues([[new Date(), 'voter one', 1, 'dup', 2, 1]]);

  assert.equal(SM.readSurveyResponseRows_(sheet).length, 3);
  assert.equal(SM.countUniqueSurveyRespondents_(sheet), 2);
}

/**
 * findRespondentsWithNewCandidates_ flags a respondent whose last submission
 * predates a candidate added afterward — detected via the blank rank cell left
 * under that candidate's newly-appended Responses column (_reconcileCandidatesWithResponses_
 * never backfills old rows). A respondent who re-submits after the new candidate
 * exists ranks it too, so they drop off the stale list.
 */
function testFindRespondentsWithNewCandidates() {
  const { SM, ss } = loadSurveyModel();
  const sheet = SM.createNewSurvey_(ss, 'StaleCheck');

  SM.addSurveyCandidate_(sheet, 'Alice');
  SM.addSurveyCandidate_(sheet, 'Bob');
  SM.submitSurveyResponse_('StaleCheck', 'Voter One', ['Alice', 'Bob'], '');
  SM.submitSurveyResponse_('StaleCheck', 'Voter Two', ['Bob', 'Alice'], '');

  assert.deepEqual(SM.findRespondentsWithNewCandidates_(sheet), []);

  SM.addSurveyCandidate_(sheet, 'Carol'); // added after both respondents' submissions
  assert.deepEqual(SM.findRespondentsWithNewCandidates_(sheet), ['Voter One', 'Voter Two']);

  // Voter One re-ranks including Carol — no longer stale. Voter Two still is.
  SM.submitSurveyResponse_('StaleCheck', 'Voter One', ['Carol', 'Alice', 'Bob'], '');
  assert.deepEqual(SM.findRespondentsWithNewCandidates_(sheet), ['Voter Two']);
}

/**
 * getSurveyForRespondent_ is the RPC-facing payload the survey page's client JS
 * actually renders from — confirms itemDetails is included (keyed by name, so it
 * survives the per-respondent candidate reordering below it), only for items that
 * have a non-empty Details note, and is never accidentally omitted.
 */
function testGetSurveyForRespondentIncludesItemDetails() {
  const { SM, ss } = loadSurveyModel();
  const sheet = SM.createNewSurvey_(ss, 'RespondentDetails');
  SM.addSurveyCandidate_(sheet, 'Alice', 'Loves cats');
  SM.addSurveyCandidate_(sheet, 'Bob', '');

  const result = SM.getSurveyForRespondent_('RespondentDetails', '');
  assert.deepEqual(result.itemDetails, { Alice: 'Loves cats' }); // Bob has no entry (empty details)
  assert.deepEqual(result.candidates, ['Alice', 'Bob']);
}

function testAddSurveyCandidateTracksDetailsPositionAligned() {
  const { SM, ss } = loadSurveyModel();
  const sheet = SM.createNewSurvey_(ss, 'CandidatesTest');

  SM.addSurveyCandidate_(sheet, 'Alice', 'Loves cats');
  SM.addSurveyCandidate_(sheet, 'Bob', '');
  SM.addSurveyCandidate_(sheet, 'Carol', 'Prefers mornings');

  assert.deepEqual(SM.readSurveyCandidates_(sheet), ['Alice', 'Bob', 'Carol']);
  assert.deepEqual(SM.readSurveyCandidateDetails_(sheet), [
    { name: 'Alice', details: 'Loves cats' },
    { name: 'Bob', details: '' },
    { name: 'Carol', details: 'Prefers mornings' },
  ]);
}

function testSaveSurveyCandidatesRenamesAndUpdatesDetailsWithoutDisturbingResponses() {
  const { SM, ss } = loadSurveyModel();
  const sheet = SM.createNewSurvey_(ss, 'RenameTest');

  SM.addSurveyCandidate_(sheet, 'Alice', '');
  SM.addSurveyCandidate_(sheet, 'Bob', '');
  SM.submitSurveyResponse_('RenameTest', 'Voter One', ['Bob', 'Alice'], '');

  SM.saveSurveyCandidates_(sheet, [
    { name: 'Alicia', details: 'renamed from Alice' },
    { name: 'Robert', details: 'renamed from Bob' },
  ]);

  assert.deepEqual(SM.readSurveyCandidates_(sheet), ['Alicia', 'Robert']);
  assert.deepEqual(SM.readSurveyCandidateDetails_(sheet), [
    { name: 'Alicia', details: 'renamed from Alice' },
    { name: 'Robert', details: 'renamed from Bob' },
  ]);

  // Rename is purely positional (column index unchanged) — the voter's ranks,
  // recorded against column position, must survive untouched.
  const rows = SM.readSurveyResponseRows_(sheet);
  assert.equal(rows.length, 1);
  assert.deepEqual(rows[0].ranks, [2, 1]); // Alice(now Alicia)=2, Bob(now Robert)=1
}

function testSaveSurveyCandidatesRejectsCountMismatch() {
  const { SM, ss } = loadSurveyModel();
  const sheet = SM.createNewSurvey_(ss, 'MismatchTest');
  SM.addSurveyCandidate_(sheet, 'Alice', '');

  assert.throws(
    () => SM.saveSurveyCandidates_(sheet, [{ name: 'A', details: '' }, { name: 'B', details: '' }]),
    /count changed/
  );
}

function testSaveSurveyCandidatesRejectsEmptyName() {
  const { SM, ss } = loadSurveyModel();
  const sheet = SM.createNewSurvey_(ss, 'EmptyNameTest');
  SM.addSurveyCandidate_(sheet, 'Alice', '');

  assert.throws(() => SM.saveSurveyCandidates_(sheet, [{ name: '  ', details: '' }]), /cannot be empty/);
}

function testAddSurveyCandidateForAdminAndSaveSurveyCandidatesForId() {
  const { SM, ss } = loadSurveyModel();
  SM.createNewSurvey_(ss, 'IdWrappers');

  const result = SM.addSurveyCandidateForAdmin_('IdWrappers', 'Dana', 'added via admin');
  assert.deepEqual(result, { candidate: 'Dana' });

  SM.saveSurveyCandidatesForId_('IdWrappers', [{ name: 'Dana Renamed', details: 'updated' }]);

  const sheet = SM.findSurveySheet_(ss, 'IdWrappers');
  assert.deepEqual(SM.readSurveyCandidateDetails_(sheet), [{ name: 'Dana Renamed', details: 'updated' }]);
}

/**
 * Respondents can now set a Details note too when adding their own item (previously
 * admin-only) — via the survey page's "Add New" panel, RPC'd through
 * addSurveyCandidateForId_'s new optional `details` param.
 */
function testAddSurveyCandidateForIdAcceptsRespondentDetails() {
  const { SM, ss } = loadSurveyModel();
  SM.createNewSurvey_(ss, 'RespondentAdd');

  const result = SM.addSurveyCandidateForId_('RespondentAdd', 'Elm Street', 'suggested by a voter');
  assert.deepEqual(result, { candidate: 'Elm Street' });

  const sheet = SM.findSurveySheet_(ss, 'RespondentAdd');
  assert.deepEqual(SM.readSurveyCandidateDetails_(sheet), [{ name: 'Elm Street', details: 'suggested by a voter' }]);
}

/**
 * Simulates a sheet created before the Candidates section existed (Results/Responses
 * markers only, candidates already typed directly into the Responses header) to
 * confirm _ensureCandidatesSection_ inserts an empty Candidates section on first
 * touch, and that adding a further candidate correctly backfills the pre-existing
 * ones instead of misaligning the new candidate's details against the wrong column.
 */
function testEnsureCandidatesSectionSelfHealsOldSheetAndBackfillsOnAdd() {
  const { SM, ss } = loadSurveyModel();
  const sheet = ss.insertSheet('Survey-PreCandidates');
  sheet.getRange(1, 1, 6, 2).setValues([
    ['Title', 'T'], ['Description', 'D'], ['Footer', 'F'],
    ['Contact', 'C'], ['Accept-New', 'TRUE'], ['Info', ''],
  ]);
  sheet.getRange(8, 1).setValue('Results');
  sheet.getRange(10, 1).setValue('Responses');
  sheet.getRange(11, 1, 1, 6).setValues([['Date', 'Name', 'Weight', 'Comment', 'Alice', 'Bob']]);

  assert.deepEqual(SM.readSurveyCandidateDetails_(sheet), []); // no Candidates rows yet, but no throw
  // _ensureCandidatesSection_ inserts the Candidates marker+header directly above the
  // (now shifted) Responses marker: old Responses(10) -> 12, Candidates marker lands
  // at 10. Looking up the Responses marker along the way migrates its bare legacy text
  // to "[Responses]"; the new Candidates marker is written decorated from the start.
  assert.equal(sheet.getRange(10, 1).getValue(), '[Candidates]');
  assert.equal(sheet.getRange(11, 1).getValue(), ''); // col A reserved, blank
  assert.deepEqual(sheet.getRange(11, 2, 1, 2).getValues()[0], ['Name', 'Details']);
  assert.equal(sheet.getRange(12, 1).getValue(), '[Responses]');

  SM.addSurveyCandidate_(sheet, 'Carol', 'newest candidate');

  assert.deepEqual(SM.readSurveyCandidates_(sheet), ['Alice', 'Bob', 'Carol']);
  assert.deepEqual(SM.readSurveyCandidateDetails_(sheet), [
    { name: 'Alice', details: '' },
    { name: 'Bob', details: '' },
    { name: 'Carol', details: 'newest candidate' },
  ]);
}

/**
 * Reproduces the live bug found on Survey-ABC-2026-7: a Candidates section that DOES
 * exist (created before the column-A-reservation fix shipped), with Name/Details
 * still sitting in columns A/B instead of B/C. Before _migrateCandidatesColumnsIfNeeded_
 * existed, readSurveyCandidateDetails_ would silently misread column B (the old
 * Details value) as the new Name column and an always-empty column C as Details —
 * returning every candidate with blank details instead of throwing, which is why this
 * went unnoticed until a user reported "I don't see any candidate details."
 */
function testMigratesExistingOldColumnCandidatesSectionOnRead() {
  const { SM, ss } = loadSurveyModel();
  const sheet = ss.insertSheet('Survey-OldCandidatesCols');
  sheet.getRange(1, 1, 6, 2).setValues([
    ['Title', 'T'], ['Description', 'D'], ['Footer', 'F'],
    ['Contact', 'C'], ['Accept-New', 'TRUE'], ['Info', ''],
  ]);
  sheet.getRange(8, 1).setValue('[Results]');
  sheet.getRange(9, 1).setValue('[Candidates]');
  // Old layout: Name/Details in columns A/B (not B/C).
  sheet.getRange(10, 1, 1, 2).setValues([['Name', 'Details']]);
  sheet.getRange(11, 1, 3, 2).setValues([
    ['book 1', 'details for b1'],
    ['book 2', 'details for b2'],
    ['book 3', 'details for b3'],
  ]);
  sheet.getRange(14, 1).setValue('[Responses]');
  sheet.getRange(15, 1, 1, 7).setValues([['Date', 'Name', 'Weight', 'Comment', 'book 1', 'book 2', 'book 3']]);

  const candidateRows = SM.readSurveyCandidateDetails_(sheet);
  assert.deepEqual(candidateRows, [
    { name: 'book 1', details: 'details for b1' },
    { name: 'book 2', details: 'details for b2' },
    { name: 'book 3', details: 'details for b3' },
  ]);

  // Migrated in place: header/data now live at columns B/C, column A blank.
  assert.equal(sheet.getRange(10, 1).getValue(), '');
  assert.deepEqual(sheet.getRange(10, 2, 1, 2).getValues()[0], ['Name', 'Details']);
  assert.equal(sheet.getRange(11, 1).getValue(), '');
  assert.deepEqual(sheet.getRange(11, 2, 1, 2).getValues()[0], ['book 1', 'details for b1']);

  // Candidates/Responses section untouched by the migration.
  assert.deepEqual(SM.readSurveyCandidates_(sheet), ['book 1', 'book 2', 'book 3']);
}

/**
 * Confirms the old "[Items]" marker (and its bare pre-decoration form "Items") is
 * recognized as a legacy alias for "[Candidates]" and migrated to the decorated
 * "[Candidates]" text in place the first time it's looked up — covers a sheet that
 * hasn't yet been manually renamed since the Items->Candidates rename.
 */
function testLegacyItemsMarkerMigratesToCandidatesMarker() {
  const { SM, ss } = loadSurveyModel();
  const sheet = ss.insertSheet('Survey-LegacyItems');
  sheet.getRange(1, 1, 6, 2).setValues([
    ['Title', 'T'], ['Description', 'D'], ['Footer', 'F'],
    ['Contact', 'C'], ['Accept-New', 'TRUE'], ['Info', ''],
  ]);
  sheet.getRange(8, 1).setValue('[Results]');
  sheet.getRange(9, 1).setValue('[Items]');
  sheet.getRange(10, 2, 1, 2).setValues([['Name', 'Details']]);
  sheet.getRange(11, 2, 1, 2).setValues([['Alice', 'Loves cats']]);
  sheet.getRange(12, 1).setValue('[Responses]');
  sheet.getRange(13, 1, 1, 5).setValues([['Date', 'Name', 'Weight', 'Comment', 'Alice']]);

  assert.deepEqual(SM.readSurveyCandidateDetails_(sheet), [{ name: 'Alice', details: 'Loves cats' }]);
  assert.equal(sheet.getRange(9, 1).getValue(), '[Candidates]'); // migrated in place
}

/**
 * The Candidates table is the PRIMARY identification of candidates: a candidate
 * pre-populated there directly (e.g. by hand-editing the sheet, ahead of any votes)
 * but not yet present in the Responses header must still show up everywhere that
 * reads candidates — readSurveyCandidates_ reconciles the Responses header to match
 * on every read, appending the missing column(s) without disturbing existing ones.
 */
function testPrePopulatedCandidateNotInResponsesIsReconciledIn() {
  const { SM, ss } = loadSurveyModel();
  const sheet = SM.createNewSurvey_(ss, 'PrePopulated');

  SM.addSurveyCandidate_(sheet, 'Alice', '');
  SM.submitSurveyResponse_('PrePopulated', 'Voter One', ['Alice'], '');

  // Hand-populate a second candidate directly into the Candidates table only —
  // simulates pre-populating candidates in the spreadsheet without going through
  // addSurveyCandidate_, so the Responses header doesn't have a column for it yet.
  SM.writeSurveyCandidateDetails_(sheet, [
    { name: 'Alice', details: '' },
    { name: 'Bob', details: '' },
  ]);

  // Responses header still only has Alice's column at this point.
  assert.equal(sheet.getLastColumn(), 5);

  // readSurveyCandidates_ reconciles: Bob's column gets appended, Alice's existing
  // column/position (and Voter One's recorded rank) are left untouched.
  assert.deepEqual(SM.readSurveyCandidates_(sheet), ['Alice', 'Bob']);
  const rows = SM.readSurveyResponseRows_(sheet);
  assert.equal(rows.length, 1);
  assert.deepEqual(rows[0].ranks, [1, '']); // Alice=1 (unchanged), Bob unranked

  // getSurveyForRespondent_ (the RPC payload the survey page renders from) also
  // sees the reconciled, complete candidate list.
  const result = SM.getSurveyForRespondent_('PrePopulated', '');
  assert.deepEqual(result.candidates, ['Alice', 'Bob']);
}

function testDuplicateIdThrows() {
  const { SM, ss } = loadSurveyModel();
  SM.createNewSurvey_(ss, 'Dup');
  assert.throws(() => SM.createNewSurvey_(ss, 'Dup'), /already exists/);
}

function testInvalidIdThrows() {
  const { SM, ss } = loadSurveyModel();
  assert.throws(() => SM.createNewSurvey_(ss, ''), /required/);
  assert.throws(() => SM.createNewSurvey_(ss, 'bad id with spaces'), /may only contain/);
  assert.throws(() => SM.createNewSurvey_(ss, '_leadingUnderscore'), /may only contain/);
}

function run() {
  testCreateNewSurveyCreatesSkeleton();
  testListSurveyIdsAndFindSurveySheet();
  testWriteThenReadConfigRoundTrips();
  testWriteSurveyConfigSelfHealsMissingAdminOnlyNotesRowOnOldSheet();
  testWriteSurveyConfigInsertsRowWhenNoBlankSpacerExists();
  testLegacyInfoKeyMigratesToAdminOnlyNotes();
  testCandidatesAndSubmitResponseRoundTrip();
  testCountUniqueSurveyRespondents();
  testFindRespondentsWithNewCandidates();
  testGetSurveyForRespondentIncludesItemDetails();
  testAddSurveyCandidateTracksDetailsPositionAligned();
  testSaveSurveyCandidatesRenamesAndUpdatesDetailsWithoutDisturbingResponses();
  testSaveSurveyCandidatesRejectsCountMismatch();
  testSaveSurveyCandidatesRejectsEmptyName();
  testAddSurveyCandidateForAdminAndSaveSurveyCandidatesForId();
  testAddSurveyCandidateForIdAcceptsRespondentDetails();
  testEnsureCandidatesSectionSelfHealsOldSheetAndBackfillsOnAdd();
  testMigratesExistingOldColumnCandidatesSectionOnRead();
  testLegacyItemsMarkerMigratesToCandidatesMarker();
  testPrePopulatedCandidateNotInResponsesIsReconciledIn();
  testDuplicateIdThrows();
  testInvalidIdThrows();
  console.log('test_survey_model: all tests passed');
}

run();
