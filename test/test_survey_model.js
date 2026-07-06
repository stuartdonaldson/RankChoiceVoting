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
  assert.equal(config.Description, '[TODO: description/instructions shown above the ranking list]');
  assert.equal(config.Footer, '[TODO: footer text, e.g. deadline or sponsoring group]');
  assert.equal(config.Contact, '[TODO: contact name/email for questions]');
  assert.equal(config['Accept-New'], 'TRUE');
  assert.equal(config.Info, '');

  // Results marker at row 8, Items marker/header at 9/10, Responses marker at row 11,
  // Responses header at row 12. Markers are bracket-decorated ("[Results]", not
  // "Results") so a survey item literally named "Results"/"Items"/"Responses" can
  // never be mistaken for the real marker; the Items header (row 10) starts at col B
  // (col A stays blank/reserved) so item data is never in the marker-scanned column.
  assert.equal(sheet.getRange(8, 1).getValue(), '[Results]');
  assert.equal(sheet.getRange(9, 1).getValue(), '[Items]');
  assert.equal(sheet.getRange(10, 1).getValue(), ''); // col A reserved, always blank here
  assert.deepEqual(sheet.getRange(10, 2, 1, 2).getValues()[0], ['Name', 'Details']);
  assert.equal(sheet.getRange(11, 1).getValue(), '[Responses]');
  assert.deepEqual(sheet.getRange(12, 1, 1, 4).getValues()[0], ['Date', 'Name', 'Weight', 'Comment']);
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
 * Simulates a sheet created by a PRE-"Info" version of createSurveySheet_ (5 config
 * rows, a blank spacer at row 6, Results at row 7, Responses at row 9) to confirm
 * writeSurveyConfig_ reuses that blank spacer row for "Info" on first save rather
 * than silently dropping the value (or needlessly inserting a new row and shifting
 * Results/Responses).
 */
function testWriteSurveyConfigSelfHealsMissingInfoRowOnOldSheet() {
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
    Info: 'newly added notes',
  });

  const config = SM.readSurveyConfig_(sheet);
  assert.equal(config.Info, 'newly added notes');
  assert.equal(config.Title, 'Old Title'); // pre-existing keys untouched

  // Info reused the existing blank spacer row (6) — Results/Responses/header stay put.
  // The bare "Results" text this test seeded gets migrated to "[Results]" in place the
  // moment writeSurveyConfig_ looks it up (via _findMarkerRow_); "Responses" is never
  // looked up in this test, so it's left as-is — legacy text is only ever migrated on
  // actual use, not proactively.
  assert.equal(sheet.getRange(6, 1).getValue(), 'Info');
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

  SM.writeSurveyConfig_(sheet, { Title: 'T', Info: 'notes' });

  const config = SM.readSurveyConfig_(sheet);
  assert.equal(config.Info, 'notes');

  // A new row was inserted at row 5 (just above Results), shifting Results/Responses/
  // header down by one row each. Results migrates to "[Results]" (looked up by
  // writeSurveyConfig_); Responses is never looked up here, so stays bare.
  assert.equal(sheet.getRange(5, 1).getValue(), 'Info');
  assert.equal(sheet.getRange(6, 1).getValue(), '[Results]');
  assert.equal(sheet.getRange(8, 1).getValue(), 'Responses');
  assert.deepEqual(sheet.getRange(9, 1, 1, 4).getValues()[0], ['Date', 'Name', 'Weight', 'Comment']);
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
  const sheet = SM.createNewSurvey_(ss, 'ItemsTest');

  SM.addSurveyCandidate_(sheet, 'Alice', 'Loves cats');
  SM.addSurveyCandidate_(sheet, 'Bob', '');
  SM.addSurveyCandidate_(sheet, 'Carol', 'Prefers mornings');

  assert.deepEqual(SM.readSurveyCandidates_(sheet), ['Alice', 'Bob', 'Carol']);
  assert.deepEqual(SM.readSurveyItemDetails_(sheet), [
    { name: 'Alice', details: 'Loves cats' },
    { name: 'Bob', details: '' },
    { name: 'Carol', details: 'Prefers mornings' },
  ]);
}

function testSaveSurveyItemsRenamesAndUpdatesDetailsWithoutDisturbingResponses() {
  const { SM, ss } = loadSurveyModel();
  const sheet = SM.createNewSurvey_(ss, 'RenameTest');

  SM.addSurveyCandidate_(sheet, 'Alice', '');
  SM.addSurveyCandidate_(sheet, 'Bob', '');
  SM.submitSurveyResponse_('RenameTest', 'Voter One', ['Bob', 'Alice'], '');

  SM.saveSurveyItems_(sheet, [
    { name: 'Alicia', details: 'renamed from Alice' },
    { name: 'Robert', details: 'renamed from Bob' },
  ]);

  assert.deepEqual(SM.readSurveyCandidates_(sheet), ['Alicia', 'Robert']);
  assert.deepEqual(SM.readSurveyItemDetails_(sheet), [
    { name: 'Alicia', details: 'renamed from Alice' },
    { name: 'Robert', details: 'renamed from Bob' },
  ]);

  // Rename is purely positional (column index unchanged) — the voter's ranks,
  // recorded against column position, must survive untouched.
  const rows = SM.readSurveyResponseRows_(sheet);
  assert.equal(rows.length, 1);
  assert.deepEqual(rows[0].ranks, [2, 1]); // Alice(now Alicia)=2, Bob(now Robert)=1
}

function testSaveSurveyItemsRejectsCountMismatch() {
  const { SM, ss } = loadSurveyModel();
  const sheet = SM.createNewSurvey_(ss, 'MismatchTest');
  SM.addSurveyCandidate_(sheet, 'Alice', '');

  assert.throws(
    () => SM.saveSurveyItems_(sheet, [{ name: 'A', details: '' }, { name: 'B', details: '' }]),
    /count changed/
  );
}

function testSaveSurveyItemsRejectsEmptyName() {
  const { SM, ss } = loadSurveyModel();
  const sheet = SM.createNewSurvey_(ss, 'EmptyNameTest');
  SM.addSurveyCandidate_(sheet, 'Alice', '');

  assert.throws(() => SM.saveSurveyItems_(sheet, [{ name: '  ', details: '' }]), /cannot be empty/);
}

function testAddSurveyItemForAdminAndSaveSurveyItemsForId() {
  const { SM, ss } = loadSurveyModel();
  SM.createNewSurvey_(ss, 'IdWrappers');

  const result = SM.addSurveyItemForAdmin_('IdWrappers', 'Dana', 'added via admin');
  assert.deepEqual(result, { candidate: 'Dana' });

  SM.saveSurveyItemsForId_('IdWrappers', [{ name: 'Dana Renamed', details: 'updated' }]);

  const sheet = SM.findSurveySheet_(ss, 'IdWrappers');
  assert.deepEqual(SM.readSurveyItemDetails_(sheet), [{ name: 'Dana Renamed', details: 'updated' }]);
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
  assert.deepEqual(SM.readSurveyItemDetails_(sheet), [{ name: 'Elm Street', details: 'suggested by a voter' }]);
}

/**
 * Simulates a sheet created before the Items section existed (Results/Responses
 * markers only, candidates already typed directly into the Responses header) to
 * confirm _ensureItemsSection_ inserts an empty Items section on first touch, and
 * that adding a further candidate correctly backfills the pre-existing ones instead
 * of misaligning the new item's details against the wrong column.
 */
function testEnsureItemsSectionSelfHealsOldSheetAndBackfillsOnAdd() {
  const { SM, ss } = loadSurveyModel();
  const sheet = ss.insertSheet('Survey-PreItems');
  sheet.getRange(1, 1, 6, 2).setValues([
    ['Title', 'T'], ['Description', 'D'], ['Footer', 'F'],
    ['Contact', 'C'], ['Accept-New', 'TRUE'], ['Info', ''],
  ]);
  sheet.getRange(8, 1).setValue('Results');
  sheet.getRange(10, 1).setValue('Responses');
  sheet.getRange(11, 1, 1, 6).setValues([['Date', 'Name', 'Weight', 'Comment', 'Alice', 'Bob']]);

  assert.deepEqual(SM.readSurveyItemDetails_(sheet), []); // no Items rows yet, but no throw
  // _ensureItemsSection_ inserts the Items marker+header directly above the (now
  // shifted) Responses marker: old Responses(10) -> 12, Items marker lands at 10.
  // Looking up the Responses marker along the way migrates its bare legacy text to
  // "[Responses]"; the new Items marker is written decorated from the start.
  assert.equal(sheet.getRange(10, 1).getValue(), '[Items]');
  assert.equal(sheet.getRange(11, 1).getValue(), ''); // col A reserved, blank
  assert.deepEqual(sheet.getRange(11, 2, 1, 2).getValues()[0], ['Name', 'Details']);
  assert.equal(sheet.getRange(12, 1).getValue(), '[Responses]');

  SM.addSurveyCandidate_(sheet, 'Carol', 'newest item');

  assert.deepEqual(SM.readSurveyCandidates_(sheet), ['Alice', 'Bob', 'Carol']);
  assert.deepEqual(SM.readSurveyItemDetails_(sheet), [
    { name: 'Alice', details: '' },
    { name: 'Bob', details: '' },
    { name: 'Carol', details: 'newest item' },
  ]);
}

/**
 * Reproduces the live bug found on Survey-ABC-2026-7: an Items section that DOES
 * exist (created before the column-A-reservation fix shipped), with Name/Details
 * still sitting in columns A/B instead of B/C. Before _migrateItemsColumnsIfNeeded_
 * existed, readSurveyItemDetails_ would silently misread column B (the old Details
 * value) as the new Name column and an always-empty column C as Details — returning
 * every item with blank details instead of throwing, which is why this went
 * unnoticed until a user reported "I don't see any item details."
 */
function testMigratesExistingOldColumnItemsSectionOnRead() {
  const { SM, ss } = loadSurveyModel();
  const sheet = ss.insertSheet('Survey-OldItemsCols');
  sheet.getRange(1, 1, 6, 2).setValues([
    ['Title', 'T'], ['Description', 'D'], ['Footer', 'F'],
    ['Contact', 'C'], ['Accept-New', 'TRUE'], ['Info', ''],
  ]);
  sheet.getRange(8, 1).setValue('[Results]');
  sheet.getRange(9, 1).setValue('[Items]');
  // Old layout: Name/Details in columns A/B (not B/C).
  sheet.getRange(10, 1, 1, 2).setValues([['Name', 'Details']]);
  sheet.getRange(11, 1, 3, 2).setValues([
    ['book 1', 'details for b1'],
    ['book 2', 'details for b2'],
    ['book 3', 'details for b3'],
  ]);
  sheet.getRange(14, 1).setValue('[Responses]');
  sheet.getRange(15, 1, 1, 7).setValues([['Date', 'Name', 'Weight', 'Comment', 'book 1', 'book 2', 'book 3']]);

  const items = SM.readSurveyItemDetails_(sheet);
  assert.deepEqual(items, [
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
  testWriteSurveyConfigSelfHealsMissingInfoRowOnOldSheet();
  testWriteSurveyConfigInsertsRowWhenNoBlankSpacerExists();
  testCandidatesAndSubmitResponseRoundTrip();
  testGetSurveyForRespondentIncludesItemDetails();
  testAddSurveyCandidateTracksDetailsPositionAligned();
  testSaveSurveyItemsRenamesAndUpdatesDetailsWithoutDisturbingResponses();
  testSaveSurveyItemsRejectsCountMismatch();
  testSaveSurveyItemsRejectsEmptyName();
  testAddSurveyItemForAdminAndSaveSurveyItemsForId();
  testAddSurveyCandidateForIdAcceptsRespondentDetails();
  testEnsureItemsSectionSelfHealsOldSheetAndBackfillsOnAdd();
  testMigratesExistingOldColumnItemsSectionOnRead();
  testDuplicateIdThrows();
  testInvalidIdThrows();
  console.log('test_survey_model: all tests passed');
}

run();
