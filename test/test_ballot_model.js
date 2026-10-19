'use strict';

const assert = require('assert');
const vm = require('vm');
const fs = require('fs');
const path = require('path');

const { createFakeSpreadsheet, createFakeGasGlobals } = require('./fakeGas');
const { prodAbc2026_7Rows, sitAbc2026_08Rows, writeGrid } = require('./ballotFixtures');

/**
 * Loads script/BallotCache.js and script/BallotModel.js into one fresh vm context wired to a
 * fresh FakeSpreadsheet — BallotModel.js's getBallotForRespondent_ calls BallotCache.js's cache
 * helpers directly (same as they share one global namespace in the real GAS runtime), so both
 * must be evaluated into the same sandbox for it to resolve. Each file's `module.exports` is
 * captured and merged separately (rather than loaded via Node's require, which BallotModel.js
 * itself guards against with its own `typeof module !== 'undefined'` check) since a second
 * `module.exports = {...}` at file scope would otherwise clobber the first file's exports.
 */
function loadBallotModel() {
  const ss = createFakeSpreadsheet();
  const sandbox = Object.assign({ module: { exports: {} } }, createFakeGasGlobals(ss));
  vm.createContext(sandbox);

  const cacheSrc = fs.readFileSync(path.join(__dirname, '..', 'script', 'BallotCache.js'), 'utf8');
  vm.runInContext(cacheSrc, sandbox, { filename: 'BallotCache.js' });
  const cacheExports = sandbox.module.exports;

  sandbox.module = { exports: {} };
  const src = fs.readFileSync(path.join(__dirname, '..', 'script', 'BallotModel.js'), 'utf8');
  vm.runInContext(src, sandbox, { filename: 'BallotModel.js' });
  const modelExports = sandbox.module.exports;

  return { SM: Object.assign({}, cacheExports, modelExports), ss, sandbox };
}

function testCreateNewBallotCreatesSkeleton() {
  const { SM, ss } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'Test123');

  assert.equal(sheet.getName(), 'Ballot-Test123');

  const config = SM.readBallotConfig_(sheet);
  assert.equal(config.Title, '[TODO: ballot title shown to respondents — delete this bracket and edit the example] F3 Active Book Club — Next Book Selection');
  assert.equal(config.Description, '[TODO: describe what this ballot is about — delete this bracket and edit the example] Help choose our next ABC read! Rank the books you’d be most interested in reading and discussing with the PAX. You don’t need to have read any of them to vote. Enter your name below to start or resume ranking — if you’ve already been here, your previous answers will be loaded so you can review or change them.');
  assert.equal(config.Instructions, '[TODO: any extra guidance for ranking these particular candidates — delete this bracket and edit the example] Drag your favorites to the top, ranking as many or as few as you’d like. Your first choice should be the book you’re most excited to discuss.');
  assert.equal(config.Footer, '[TODO: footer text, e.g. deadline or sponsoring group — delete this bracket and edit the example] Voting closes [date and time]. The book with the strongest ranked-choice support will be our next ABC selection.');
  assert.equal(config.Contact, '[TODO: contact name/email for questions — delete this bracket and edit the example] Questions or suggestions? Post in #ao-active-book-club.');
  assert.equal(config['Accept-New'], 'TRUE');
  assert.equal(config['Add-Instructions'], '[TODO: instructions shown above the "+ Add New" button, if enabled — delete this bracket and edit the example] Have another book you’d like ABC to consider? Add the title and author, and preferably a link to the book.');
  assert.equal(config['Admin-Only-Notes'], '');

  // Results marker at row 10, Candidates marker/header at 11/12, Responses marker at row
  // 13, Responses header at row 14. Markers are bracket-decorated ("[Results]", not
  // "Results") so a ballot candidate literally named "Results"/"Candidates"/"Responses"
  // can never be mistaken for the real marker; the Candidates header (row 12) starts
  // at col B (col A stays blank/reserved) so candidate data is never in the
  // marker-scanned column.
  assert.equal(sheet.getRange(10, 1).getValue(), '[Results]');
  assert.equal(sheet.getRange(11, 1).getValue(), '[Candidates]');
  assert.equal(sheet.getRange(12, 1).getValue(), ''); // col A reserved, always blank here
  assert.deepEqual(sheet.getRange(12, 2, 1, 2).getValues()[0], ['Name', 'Details']);
  assert.equal(sheet.getRange(13, 1).getValue(), '[Responses]');
  assert.deepEqual(sheet.getRange(14, 1, 1, 4).getValues()[0], ['Date', 'Name', 'Weight', 'Comment']);
  assertSnapshotMatchesOldReaders(SM, sheet, 'testCreateNewBallotCreatesSkeleton');
}

function testListBallotIdsAndFindBallotSheet() {
  const { SM, ss } = loadBallotModel();
  SM.createNewBallot_(ss, 'Alpha');
  SM.createNewBallot_(ss, 'Beta');

  assert.deepEqual(SM.listBallotIds_(ss), ['Alpha', 'Beta']);
  assert.ok(SM.findBallotSheet_(ss, 'Alpha'));
  assert.equal(SM.findBallotSheet_(ss, 'Nope'), null);
}

function testWriteThenReadConfigRoundTrips() {
  const { SM, ss } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'RoundTrip');

  SM.writeBallotConfig_(sheet, {
    Title: 'Board Election 2026',
    Description: 'Rank your top three.',
    Footer: 'Closes Friday',
    Contact: 'admin@example.com',
    'Accept-New': 'FALSE',
  });

  const config = SM.readBallotConfig_(sheet);
  assert.equal(config.Title, 'Board Election 2026');
  assert.equal(config.Description, 'Rank your top three.');
  assert.equal(config.Footer, 'Closes Friday');
  assert.equal(config.Contact, 'admin@example.com');
  assert.equal(config['Accept-New'], 'FALSE');
  assertSnapshotMatchesOldReaders(SM, sheet, 'testWriteThenReadConfigRoundTrips');
}

/**
 * Simulates a sheet created by a PRE-"Admin-Only-Notes" version of createBallotSheet_
 * (5 config rows, a blank spacer at row 6, Results at row 7, Responses at row 9) to
 * confirm writeBallotConfig_ reuses that blank spacer row for "Admin-Only-Notes" on
 * first save rather than silently dropping the value (or needlessly inserting a new
 * row and shifting Results/Responses).
 */
function testWriteBallotConfigSelfHealsMissingAdminOnlyNotesRowOnOldSheet() {
  const { SM, ss } = loadBallotModel();
  const sheet = ss.insertSheet('Ballot-OldFormat');
  sheet.getRange(1, 1, 5, 2).setValues([
    ['Title', 'Old Title'],
    ['Description', 'Old Description'],
    ['Footer', 'Old Footer'],
    ['Contact', 'old@example.com'],
    ['Accept-New', 'TRUE'],
  ]);
  sheet.getRange(7, 1).setValue('[Results]');
  sheet.getRange(9, 1).setValue('[Responses]');
  sheet.getRange(10, 1, 1, 4).setValues([['Date', 'Name', 'Weight', 'Comment']]);

  SM.writeBallotConfig_(sheet, {
    Title: 'Old Title',
    Description: 'Old Description',
    Footer: 'Old Footer',
    Contact: 'old@example.com',
    'Accept-New': 'TRUE',
    'Admin-Only-Notes': 'newly added notes',
  });

  const config = SM.readBallotConfig_(sheet);
  assert.equal(config['Admin-Only-Notes'], 'newly added notes');
  assert.equal(config.Title, 'Old Title'); // pre-existing keys untouched

  // Admin-Only-Notes reused the existing blank spacer row (6) — Results stays put. The
  // write runs the repair path first (rcballot-8i5.4), which gives this pre-Candidates
  // ballot an empty [Candidates] section directly above [Responses] (rows 9-10).
  assert.equal(sheet.getRange(6, 1).getValue(), 'Admin-Only-Notes');
  assert.equal(sheet.getRange(7, 1).getValue(), '[Results]');
  assert.equal(sheet.getRange(9, 1).getValue(), '[Candidates]');
  assert.equal(sheet.getRange(11, 1).getValue(), '[Responses]');
  assert.deepEqual(sheet.getRange(12, 1, 1, 4).getValues()[0], ['Date', 'Name', 'Weight', 'Comment']);
  assertSnapshotMatchesOldReaders(SM, sheet, 'testWriteBallotConfigSelfHealsMissingAdminOnlyNotesRowOnOldSheet');
}

/**
 * Same self-heal scenario, but with NO spare blank row above Results (config rows
 * butt directly against the marker) — exercises the insert-a-new-row fallback path.
 */
function testWriteBallotConfigInsertsRowWhenNoBlankSpacerExists() {
  const { SM, ss } = loadBallotModel();
  const sheet = ss.insertSheet('Ballot-NoSpacer');
  sheet.getRange(1, 1, 4, 2).setValues([
    ['Title', 'T'],
    ['Description', 'D'],
    ['Footer', 'F'],
    ['Contact', 'C'],
  ]);
  sheet.getRange(5, 1).setValue('[Results]'); // no blank row between config and marker
  sheet.getRange(7, 1).setValue('[Responses]');
  sheet.getRange(8, 1, 1, 4).setValues([['Date', 'Name', 'Weight', 'Comment']]);

  SM.writeBallotConfig_(sheet, { Title: 'T', 'Admin-Only-Notes': 'notes' });

  const config = SM.readBallotConfig_(sheet);
  assert.equal(config['Admin-Only-Notes'], 'notes');

  // A new row was inserted at row 5 (just above Results), shifting everything below down
  // one row; the repair path had first put an empty [Candidates] section above [Responses].
  assert.equal(sheet.getRange(5, 1).getValue(), 'Admin-Only-Notes');
  assert.equal(sheet.getRange(6, 1).getValue(), '[Results]');
  assert.equal(sheet.getRange(8, 1).getValue(), '[Candidates]');
  assert.equal(sheet.getRange(10, 1).getValue(), '[Responses]');
  assert.deepEqual(sheet.getRange(11, 1, 1, 4).getValues()[0], ['Date', 'Name', 'Weight', 'Comment']);
  assertSnapshotMatchesOldReaders(SM, sheet, 'testWriteBallotConfigInsertsRowWhenNoBlankSpacerExists');
}

/**
 * Confirms the old "Info" config key is recognized as a legacy alias for
 * "Admin-Only-Notes": readBallotConfig_ reads it under the new name without writing
 * (rcballot-8i5.3), and the repair path (admin load / writes) renames the cell in place.
 */
function testLegacyInfoKeyMigratesToAdminOnlyNotes() {
  const { SM, ss, sandbox } = loadBallotModel();
  const sheet = ss.insertSheet('Ballot-LegacyInfo');
  sheet.getRange(1, 1, 6, 2).setValues([
    ['Title', 'T'], ['Description', 'D'], ['Footer', 'F'],
    ['Contact', 'C'], ['Accept-New', 'TRUE'], ['Info', 'old notes'],
  ]);
  sheet.getRange(8, 1).setValue('[Results]');

  const config = SM.readBallotConfig_(sheet);
  assert.equal(config['Admin-Only-Notes'], 'old notes');
  assert.equal(config.Info, undefined);
  assert.equal(sheet.getRange(6, 1).getValue(), 'Info'); // the read wrote nothing

  SM._repairBallotSheet_(sheet);
  assert.equal(sheet.getRange(6, 1).getValue(), 'Admin-Only-Notes'); // migrated in place by the repair path
  assert.equal(logLinesWith(sandbox, 'ballot.repair.legacyConfigKey').length, 1);
  assertSnapshotMatchesOldReaders(SM, sheet, 'testLegacyInfoKeyMigratesToAdminOnlyNotes');
}

function testCandidatesAndSubmitResponseRoundTrip() {
  const { SM, ss } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'Candidates');

  SM.addBallotCandidate_(sheet, 'Alice');
  SM.addBallotCandidate_(sheet, 'Bob');
  SM.addBallotCandidate_(sheet, 'Carol');

  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Alice', 'Bob', 'Carol']);

  // submitBallotResponse_ resolves the ballot via SpreadsheetApp.getActiveSpreadsheet(),
  // which fakeGas wires to the same `ss` instance passed to createNewBallot_ above.
  const result = SM.submitBallotResponse_('Candidates', 'Voter One', ['Bob', 'Carol', 'Alice'], 'my comment');
  assert.deepEqual(result, { ok: true });

  const rows = SM.readBallotResponseRows_(sheet);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].name, 'Voter One');
  assert.equal(rows[0].weight, 1);
  assert.equal(rows[0].comment, 'my comment');
  // ranks align with readBallotCandidates_ (filled by name): Alice=3, Bob=1, Carol=2
  assert.deepEqual(rows[0].ranks, [3, 1, 2]);

  // Re-submitting for the same voter (case-insensitive) overwrites the row rather than
  // appending a new one, and preserves their existing weight.
  SM.submitBallotResponse_('Candidates', 'voter one', ['Alice', 'Bob', 'Carol'], 'updated');
  const rows2 = SM.readBallotResponseRows_(sheet);
  assert.equal(rows2.length, 1);
  assert.equal(rows2[0].comment, 'updated');
  assert.deepEqual(rows2[0].ranks, [1, 2, 3]);
  assertSnapshotMatchesOldReaders(SM, sheet, 'testCandidatesAndSubmitResponseRoundTrip');
}

/**
 * countUniqueBallotRespondents_ counts distinct respondents (case-insensitive name),
 * collapsing duplicate rows for the same person down to one. Normal re-submission
 * never creates a duplicate row (submitBallotResponse_ overwrites in place), so this
 * simulates the one way a duplicate can still appear: a hand-edited sheet or legacy
 * data imported outside the app.
 */
function testCountUniqueBallotRespondents() {
  const { SM, ss } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'DupeCheck');

  SM.addBallotCandidate_(sheet, 'Alice');
  SM.addBallotCandidate_(sheet, 'Bob');
  SM.submitBallotResponse_('DupeCheck', 'Voter One', ['Alice', 'Bob'], '');
  SM.submitBallotResponse_('DupeCheck', 'Voter Two', ['Bob', 'Alice'], '');
  assert.equal(SM.readBallotResponseRows_(sheet).length, 2);
  assert.equal(SM.countUniqueBallotRespondents_(sheet), 2);

  // Manually append a second row for "voter one" (different case), simulating a
  // hand-edited duplicate that submitBallotResponse_'s own overwrite-in-place logic
  // would never produce.
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, 1, 6).setValues([[new Date(), 'voter one', 1, 'dup', 2, 1]]);

  assert.equal(SM.readBallotResponseRows_(sheet).length, 3);
  assert.equal(SM.countUniqueBallotRespondents_(sheet), 2);
  assertSnapshotMatchesOldReaders(SM, sheet, 'testCountUniqueBallotRespondents');
}

/**
 * findRespondentsWithNewCandidates_ flags a respondent whose last submission
 * predates a candidate added afterward — detected via the blank rank cell left
 * under that candidate's newly-appended Responses column (old rows are never backfilled). A respondent who re-submits after the new candidate
 * exists ranks it too, so they drop off the stale list.
 */
function testFindRespondentsWithNewCandidates() {
  const { SM, ss } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'StaleCheck');

  SM.addBallotCandidate_(sheet, 'Alice');
  SM.addBallotCandidate_(sheet, 'Bob');
  SM.submitBallotResponse_('StaleCheck', 'Voter One', ['Alice', 'Bob'], '');
  SM.submitBallotResponse_('StaleCheck', 'Voter Two', ['Bob', 'Alice'], '');

  assert.deepEqual(SM.findRespondentsWithNewCandidates_(sheet), []);

  SM.addBallotCandidate_(sheet, 'Carol'); // added after both respondents' submissions
  assert.deepEqual(SM.findRespondentsWithNewCandidates_(sheet), ['Voter One', 'Voter Two']);

  // Voter One re-ranks including Carol — no longer stale. Voter Two still is.
  SM.submitBallotResponse_('StaleCheck', 'Voter One', ['Carol', 'Alice', 'Bob'], '');
  assert.deepEqual(SM.findRespondentsWithNewCandidates_(sheet), ['Voter Two']);
  assertSnapshotMatchesOldReaders(SM, sheet, 'testFindRespondentsWithNewCandidates');
}

/**
 * getBallotForRespondent_ is the RPC-facing payload the ballot page's client JS
 * actually renders from — confirms itemDetails is included (keyed by name, so it
 * survives the per-respondent candidate reordering below it), only for items that
 * have a non-empty Details note, and is never accidentally omitted.
 */
function testGetBallotForRespondentIncludesItemDetails() {
  const { SM, ss } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'RespondentDetails');
  SM.addBallotCandidate_(sheet, 'Alice', 'Loves cats');
  SM.addBallotCandidate_(sheet, 'Bob', '');

  const result = SM.getBallotForRespondent_('RespondentDetails', '');
  assert.deepEqual(result.itemDetails, { Alice: { details: 'Loves cats', linkText: '', linkUrl: '' } }); // Bob has no entry (empty details)
  assert.deepEqual(result.candidates, ['Alice', 'Bob']);
  assertSnapshotMatchesOldReaders(SM, sheet, 'testGetBallotForRespondentIncludesItemDetails');
}

function testAddBallotCandidateTracksDetailsPositionAligned() {
  const { SM, ss } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'CandidatesTest');

  SM.addBallotCandidate_(sheet, 'Alice', 'Loves cats');
  SM.addBallotCandidate_(sheet, 'Bob', '');
  SM.addBallotCandidate_(sheet, 'Carol', 'Prefers mornings');

  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Alice', 'Bob', 'Carol']);
  assert.deepEqual(SM.readBallotCandidateDetails_(sheet), [
    { name: 'Alice', details: 'Loves cats', linkText: '', linkUrl: '' },
    { name: 'Bob', details: '', linkText: '', linkUrl: '' },
    { name: 'Carol', details: 'Prefers mornings', linkText: '', linkUrl: '' },
  ]);
  assertSnapshotMatchesOldReaders(SM, sheet, 'testAddBallotCandidateTracksDetailsPositionAligned');
}

function testSaveBallotCandidatesRenamesAndUpdatesDetailsWithoutDisturbingResponses() {
  const { SM, ss } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'RenameTest');

  SM.addBallotCandidate_(sheet, 'Alice', '');
  SM.addBallotCandidate_(sheet, 'Bob', '');
  SM.submitBallotResponse_('RenameTest', 'Voter One', ['Bob', 'Alice'], '');

  SM.saveBallotCandidates_(sheet, [
    { name: 'Alicia', details: 'renamed from Alice' },
    { name: 'Robert', details: 'renamed from Bob' },
  ]);

  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Alicia', 'Robert']);
  assert.deepEqual(SM.readBallotCandidateDetails_(sheet), [
    { name: 'Alicia', details: 'renamed from Alice', linkText: '', linkUrl: '' },
    { name: 'Robert', details: 'renamed from Bob', linkText: '', linkUrl: '' },
  ]);

  // A rename relabels each candidate's own rank column (found by its old name) — the
  // voter's ranks follow the candidate.
  const rows = SM.readBallotResponseRows_(sheet);
  assert.equal(rows.length, 1);
  assert.deepEqual(rows[0].ranks, [2, 1]); // Alice(now Alicia)=2, Bob(now Robert)=1
  assertSnapshotMatchesOldReaders(SM, sheet, 'testSaveBallotCandidatesRenamesAndUpdatesDetailsWithoutDisturbingResponses');
}

function testSaveBallotCandidatesRejectsCountMismatch() {
  const { SM, ss } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'MismatchTest');
  SM.addBallotCandidate_(sheet, 'Alice', '');

  assert.throws(
    () => SM.saveBallotCandidates_(sheet, [{ name: 'A', details: '' }, { name: 'B', details: '' }]),
    /count changed/
  );
}

function testSaveBallotCandidatesRejectsEmptyName() {
  const { SM, ss } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'EmptyNameTest');
  SM.addBallotCandidate_(sheet, 'Alice', '');

  assert.throws(() => SM.saveBallotCandidates_(sheet, [{ name: '  ', details: '' }]), /cannot be empty/);
}

function testAddBallotCandidateForAdminAndSaveBallotCandidatesForId() {
  const { SM, ss } = loadBallotModel();
  SM.createNewBallot_(ss, 'IdWrappers');

  const result = SM.addBallotCandidateForAdmin_('IdWrappers', 'Dana', 'added via admin');
  assert.deepEqual(result, { candidate: 'Dana' });

  SM.saveBallotCandidatesForId_('IdWrappers', [{ name: 'Dana Renamed', details: 'updated' }]);

  const sheet = SM.findBallotSheet_(ss, 'IdWrappers');
  assert.deepEqual(SM.readBallotCandidateDetails_(sheet), [{ name: 'Dana Renamed', details: 'updated', linkText: '', linkUrl: '' }]);
  assertSnapshotMatchesOldReaders(SM, sheet, 'testAddBallotCandidateForAdminAndSaveBallotCandidatesForId');
}

/**
 * Respondents can now set a Details note too when adding their own item (previously
 * admin-only) — via the ballot page's "Add New" panel, RPC'd through
 * addBallotCandidateForId_'s new optional `details` param.
 */
function testAddBallotCandidateForIdAcceptsRespondentDetails() {
  const { SM, ss } = loadBallotModel();
  SM.createNewBallot_(ss, 'RespondentAdd');

  const result = SM.addBallotCandidateForId_('RespondentAdd', 'Elm Street', 'suggested by a voter');
  assert.deepEqual(result, { candidate: 'Elm Street' });

  const sheet = SM.findBallotSheet_(ss, 'RespondentAdd');
  assert.deepEqual(SM.readBallotCandidateDetails_(sheet), [{ name: 'Elm Street', details: 'suggested by a voter', linkText: '', linkUrl: '' }]);
}

/**
 * Simulates a sheet created before the Candidates section existed (Results/Responses
 * markers only, candidates already typed directly into the Responses header). Reading it
 * writes nothing and lists no candidates (the rank columns are orphans); the first write
 * (here addBallotCandidate_) runs the repair path: inserts the Candidates section, backfills
 * the pre-existing candidates from the Responses header, then appends the new one.
 */
function testEnsureCandidatesSectionSelfHealsOldSheetAndBackfillsOnAdd() {
  const { SM, ss } = loadBallotModel();
  const sheet = ss.insertSheet('Ballot-PreCandidates');
  sheet.getRange(1, 1, 6, 2).setValues([
    ['Title', 'T'], ['Description', 'D'], ['Footer', 'F'],
    ['Contact', 'C'], ['Accept-New', 'TRUE'], ['Info', ''],
  ]);
  sheet.getRange(8, 1).setValue('[Results]');
  sheet.getRange(10, 1).setValue('[Responses]');
  sheet.getRange(11, 1, 1, 6).setValues([['Date', 'Name', 'Weight', 'Comment', 'Alice', 'Bob']]);

  sheet._writeCalls = 0;
  assert.deepEqual(SM.readBallotCandidateDetails_(sheet), []); // no Candidates rows yet, but no throw
  assert.equal(sheet._writeCalls, 0); // and no repair on read
  assert.equal(sheet.getRange(10, 1).getValue(), '[Responses]');

  SM.addBallotCandidate_(sheet, 'Carol', 'newest candidate');
  // The repair inserted the Candidates marker+header directly above the old Responses
  // marker (row 10), written decorated; col A of the header is reserved/blank.
  assert.equal(sheet.getRange(10, 1).getValue(), '[Candidates]');
  assert.equal(sheet.getRange(11, 1).getValue(), '');
  assert.deepEqual(sheet.getRange(11, 2, 1, 2).getValues()[0], ['Name', 'Details']);
  assert.equal(sheet.getRange(15, 1).getValue(), '[Responses]'); // + Alice, Bob, Carol rows

  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Alice', 'Bob', 'Carol']);
  assert.deepEqual(SM.readBallotCandidateDetails_(sheet), [
    { name: 'Alice', details: '', linkText: '', linkUrl: '' },
    { name: 'Bob', details: '', linkText: '', linkUrl: '' },
    { name: 'Carol', details: 'newest candidate', linkText: '', linkUrl: '' },
  ]);
  assertSnapshotMatchesOldReaders(SM, sheet, 'testEnsureCandidatesSectionSelfHealsOldSheetAndBackfillsOnAdd');
}

/**
 * Reproduces the live bug found on Ballot-ABC-2026-7: a Candidates section that DOES
 * exist (created before the column-A-reservation fix shipped), with Name/Details
 * still sitting in columns A/B instead of B/C. Migrated by the repair path (admin load /
 * writes), never by a read (rcballot-8i5.3). Before _migrateCandidatesColumnsIfNeeded_
 * existed, readBallotCandidateDetails_ would silently misread column B (the old
 * Details value) as the new Name column and an always-empty column C as Details —
 * returning every candidate with blank details instead of throwing, which is why this
 * went unnoticed until a user reported "I don't see any candidate details."
 */
function testMigratesExistingOldColumnCandidatesSectionOnRepair() {
  const { SM, ss } = loadBallotModel();
  const sheet = ss.insertSheet('Ballot-OldCandidatesCols');
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

  // Before repair the snapshot finds no 'Name' label at col B+: no candidates, no write.
  sheet._writeCalls = 0;
  assert.deepEqual(SM.readBallotCandidateDetails_(sheet), []);
  assert.equal(sheet._writeCalls, 0);

  SM._repairBallotSheet_(sheet); // admin load / any write
  const candidateRows = SM.readBallotCandidateDetails_(sheet);
  assert.deepEqual(candidateRows, [
    { name: 'book 1', details: 'details for b1', linkText: '', linkUrl: '' },
    { name: 'book 2', details: 'details for b2', linkText: '', linkUrl: '' },
    { name: 'book 3', details: 'details for b3', linkText: '', linkUrl: '' },
  ]);

  // Migrated in place: header/data now live at columns B/C, column A blank.
  assert.equal(sheet.getRange(10, 1).getValue(), '');
  assert.deepEqual(sheet.getRange(10, 2, 1, 2).getValues()[0], ['Name', 'Details']);
  assert.equal(sheet.getRange(11, 1).getValue(), '');
  assert.deepEqual(sheet.getRange(11, 2, 1, 2).getValues()[0], ['book 1', 'details for b1']);

  // Candidates/Responses section untouched by the migration.
  assert.deepEqual(SM.readBallotCandidates_(sheet), ['book 1', 'book 2', 'book 3']);
  assertSnapshotMatchesOldReaders(SM, sheet, 'testMigratesExistingOldColumnCandidatesSectionOnRead');
}

/**
 * Column A is markers-only: marker matching is exact bracketed text. Plain "Responses" /
 * "Candidates" / "Results" / "Items" cells (and the former "[Items]" marker) are neither
 * matched nor rewritten, so a candidate or result line with such a name can never
 * shadow or become a section boundary.
 */
function testPlainTextMarkersAreNotMatchedOrRewritten() {
  const { SM, ss } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'PlainText');
  // Plain-text look-alikes sitting in column A of the Results region (rows 11..) —
  // above the real [Candidates]/[Responses] markers, where a top-down scan hits them first.
  sheet.getRange(10, 1).setValue('[Results]');
  SM.writeBallotResults_(sheet, [['x'], ['x'], ['x'], ['x']]);
  ['Responses', 'Candidates', 'Items', '[Items]', 'Results'].forEach((text, i) => {
    sheet.getRange(11 + i, 1).setValue(text);
  });
  const before = sheet.getRange(11, 1, 5, 1).getValues().map((r) => r[0]);

  SM.writeBallotCandidateDetails_(sheet, [{ name: 'Alice', details: 'a' }]);
  assert.deepEqual(SM.readBallotCandidateDetails_(sheet).map((it) => it.name), ['Alice']);
  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Alice']);
  assert.equal(SM._findMarkerRow_(sheet, '[Results]'), 10);
  assert.equal(SM._findMarkerRow_(sheet, '[Candidates]') > 15, true);
  assert.equal(SM._findMarkerRow_(sheet, '[Responses]') > SM._findMarkerRow_(sheet, '[Candidates]'), true);
  assert.deepEqual(sheet.getRange(11, 1, 5, 1).getValues().map((r) => r[0]), before); // not rewritten
  assert.equal(sheet.getRange(11, 1).getValue(), 'Responses');
  assert.equal(sheet.getRange(12, 1).getValue(), 'Candidates');

  // A sheet whose ONLY marker-like cells are plain text has no markers.
  const bare = ss.insertSheet('Ballot-BareOnly');
  bare.getRange(1, 1, 3, 1).setValues([['Results'], ['Items'], ['Responses']]);
  assert.equal(SM._findMarkerRow_(bare, '[Results]'), -1);
  assert.equal(SM._findMarkerRow_(bare, '[Candidates]'), -1);
  assert.equal(SM._findMarkerRow_(bare, '[Responses]'), -1);
  assert.deepEqual(bare.getRange(1, 1, 3, 1).getValues().map((r) => r[0]), ['Results', 'Items', 'Responses']);
}

/** Results output starts at column B and leaves column A blank; stale column-A text is cleared. */
function testWriteBallotResultsUsesColumnBAndClearsColumnA() {
  const { SM, ss } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'ResultsColB');
  SM.addBallotCandidate_(sheet, 'Alice', '');
  SM.submitBallotResponse_('ResultsColB', 'V', ['Alice'], '');

  // Legacy-style output previously written from column A (candidate names included).
  sheet.getRange(10, 1).setValue('[Results]');
  SM.writeBallotResults_(sheet, [['placeholder']]);
  sheet.getRange(11, 1, 1, 2).setValues([['Responses', 'old']]);

  SM.writeBallotResults_(sheet, [['Header', 'Winner'], ['Alice', 3], ['Bob', 1]]);

  const candidatesRow = SM._findMarkerRow_(sheet, '[Candidates]');
  assert.equal(candidatesRow, 14); // 3 rows between [Results] (10) and [Candidates]
  for (let r = 11; r < candidatesRow; r++) {
    assert.equal(sheet.getRange(r, 1).getValue(), '', 'col A row ' + r + ' must be blank');
  }
  assert.deepEqual(sheet.getRange(11, 2, 3, 2).getValues(), [['Header', 'Winner'], ['Alice', 3], ['Bob', 1]]);

  // Shrinking the output leaves no stale text in any column of the region.
  SM.writeBallotResults_(sheet, [['only']]);
  const newCandidatesRow = SM._findMarkerRow_(sheet, '[Candidates]');
  assert.equal(newCandidatesRow, 12);
  assert.deepEqual(sheet.getRange(11, 1, 1, 3).getValues()[0], ['', 'only', '']);
}

/** A region already holding column-A text (old output) is fully cleared on rewrite. */
function testWriteBallotResultsClearsPreexistingColumnAText() {
  const { SM, ss } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'ResultsClear');
  SM.writeBallotResults_(sheet, [['a'], ['b'], ['c']]);
  sheet.getRange(11, 1, 3, 3).setValues([['Alice', 'x', 'y'], ['Bob', 'x', 'y'], ['Carol', 'x', 'y']]);

  SM.writeBallotResults_(sheet, [['n1'], ['n2'], ['n3']]);
  assert.deepEqual(sheet.getRange(11, 1, 3, 3).getValues(), [['', 'n1', ''], ['', 'n2', ''], ['', 'n3', '']]);
}

/** Candidate names equal to any marker text (bracketed or not, any case/padding) are rejected. */
function testMarkerNamedCandidatesAreRejected() {
  const { SM, ss } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'MarkerNames');
  SM.addBallotCandidate_(sheet, 'Alice', '');
  const bad = ['Responses', '[Results]', ' items ', 'CANDIDATES', '[Candidates]', 'Results', '[responses]', '[Items]'];

  bad.forEach((name) => {
    assert.throws(() => SM.addBallotCandidate_(sheet, name, ''), /Candidate name cannot be a section marker\./, 'addBallotCandidate_ ' + name);
    assert.throws(() => SM.addBallotCandidateForId_('MarkerNames', name, ''), /Candidate name cannot be a section marker\./, 'addBallotCandidateForId_ ' + name);
    assert.throws(() => SM.addBallotCandidateForAdmin_('MarkerNames', name, '', '', ''), /Candidate name cannot be a section marker\./, 'addBallotCandidateForAdmin_ ' + name);
    assert.throws(() => SM.saveBallotCandidates_(sheet, [{ name: name, details: '' }]), /Candidate name cannot be a section marker\./, 'saveBallotCandidates_ ' + name);
    assert.throws(() => SM.saveBallotCandidatesForId_('MarkerNames', [{ name: name, details: '' }]), /Candidate name cannot be a section marker\./, 'saveBallotCandidatesForId_ ' + name);
  });
  // Nothing was written by the rejected calls.
  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Alice']);
  // Names merely containing a marker word are fine.
  SM.addBallotCandidate_(sheet, 'Responses to Intervention', '');
  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Alice', 'Responses to Intervention']);
}

/**
 * The Candidates table is the PRIMARY identification of candidates: a candidate
 * pre-populated there directly (e.g. by hand-editing the sheet, ahead of any votes)
 * but not yet present in the Responses header must still show up everywhere that
 * reads candidates. Reads list it (ranks '') without writing; the repair path (admin
 * load / writes) appends its Responses column without disturbing existing ones.
 */
function testPrePopulatedCandidateNotInResponsesIsReconciledIn() {
  const { SM, ss } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'PrePopulated');

  SM.addBallotCandidate_(sheet, 'Alice', '');
  SM.submitBallotResponse_('PrePopulated', 'Voter One', ['Alice'], '');

  // Hand-populate a second candidate directly into the Candidates table only —
  // simulates pre-populating candidates in the spreadsheet without going through
  // addBallotCandidate_, so the Responses header doesn't have a column for it yet.
  SM.writeBallotCandidateDetails_(sheet, [
    { name: 'Alice', details: '' },
    { name: 'Bob', details: '' },
  ]);

  // Responses header still only has Alice's column at this point.
  assert.equal(sheet.getLastColumn(), 5);

  // Reads already list Bob (unranked) — without writing anything.
  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Alice', 'Bob']);
  let rows = SM.readBallotResponseRows_(sheet);
  assert.equal(rows.length, 1);
  assert.deepEqual(rows[0].ranks, [1, '']); // Alice=1 (unchanged), Bob unranked
  assert.equal(sheet.getLastColumn(), 5);

  // The repair path appends Bob's column; Alice's column and Voter One's rank stay put.
  SM._repairBallotSheet_(sheet);
  assert.equal(sheet.getRange(SM._readBallotLayout_(sheet).sections.responses.headerRow, 6).getValue(), 'Bob');
  rows = SM.readBallotResponseRows_(sheet);
  assert.deepEqual(rows[0].ranks, [1, '']);

  // getBallotForRespondent_ (the RPC payload the ballot page renders from) also
  // sees the reconciled, complete candidate list.
  const result = SM.getBallotForRespondent_('PrePopulated', '');
  assert.deepEqual(result.candidates, ['Alice', 'Bob']);
  assertSnapshotMatchesOldReaders(SM, sheet, 'testPrePopulatedCandidateNotInResponsesIsReconciledIn');
}

/**
 * A stray blank row in the Candidates table (e.g. hand-edited into the sheet, or left
 * behind by some other bug) must not be treated as a candidate — readBallotCandidateDetails_
 * is position-aligned with the Responses header's candidate columns, so a phantom blank
 * entry would throw off every candidate after it.
 */
function testReadBallotCandidateDetailsIgnoresBlankRow() {
  const { SM, ss } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'BlankRow');

  SM.writeBallotCandidateDetails_(sheet, [
    { name: 'Alice', details: '' },
    { name: '', details: '' }, // stray blank row
    { name: 'Bob', details: '' },
  ]);

  assert.deepEqual(SM.readBallotCandidateDetails_(sheet), [
    { name: 'Alice', details: '', linkText: '', linkUrl: '' },
    { name: 'Bob', details: '', linkText: '', linkUrl: '' },
  ]);
  assertSnapshotMatchesOldReaders(SM, sheet, 'testReadBallotCandidateDetailsIgnoresBlankRow');
}

/**
 * Regression test for a real bug seen on the SIT sheet: pre-populating the Candidates
 * table directly (skipping addBallotCandidate_) before any Responses columns exist meant
 * sheet.getLastColumn() reflected the Candidates table's own fixed-width header (through
 * its Link URL column), so the old reconcile appended candidate names one column right of
 * where ranks were written. The repair path's rankColumns step (run here by the submit)
 * places new columns right after the Responses section's own last used column, and votes
 * are written under each candidate's labelled column.
 */
function testReconcileCompactsGapCausedByWiderCandidatesTable() {
  const { SM, ss } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'GapTest');

  SM.writeBallotCandidateDetails_(sheet, [
    { name: 'Alice', details: '' },
    { name: 'Bob', details: '' },
  ]);

  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Alice', 'Bob']);

  SM.submitBallotResponse_('GapTest', 'Voter One', ['Bob', 'Alice'], '');
  const rows = SM.readBallotResponseRows_(sheet);
  assert.equal(rows.length, 1);
  // Alice=2, Bob=1 — proves the header lines up with the vote columns rather than
  // being shifted right by a gap column.
  assert.deepEqual(rows[0].ranks, [2, 1]);
  assertSnapshotMatchesOldReaders(SM, sheet, 'testReconcileCompactsGapCausedByWiderCandidatesTable');
}

function testDuplicateIdThrows() {
  const { SM, ss } = loadBallotModel();
  SM.createNewBallot_(ss, 'Dup');
  assert.throws(() => SM.createNewBallot_(ss, 'Dup'), /already exists/);
}

function testInvalidIdThrows() {
  const { SM, ss } = loadBallotModel();
  assert.throws(() => SM.createNewBallot_(ss, ''), /required/);
  assert.throws(() => SM.createNewBallot_(ss, 'bad id with spaces'), /may only contain/);
  assert.throws(() => SM.createNewBallot_(ss, '_leadingUnderscore'), /may only contain/);
}

/** Logger lines from a loadBallotModel() sandbox containing `tag`. */
function logLinesWith(sandbox, tag) {
  return sandbox.Logger._lines.filter((l) => String(l).indexOf(tag) !== -1);
}

/**
 * Builds a ballot whose [Responses] marker + header were lost (e.g. overwritten while
 * hand-pasting candidates) — the sheet simply ends after the Candidates header.
 */
function createBallotWithoutResponsesSection(SM, ss, id) {
  const sheet = SM.createNewBallot_(ss, id);
  sheet.deleteRows(13, 2); // rows 13/14: [Responses] marker + header
  assert.equal(sheet.getLastRow(), 12);
  return sheet;
}

/**
 * A missing [Responses] section with only candidate data below the Candidates header is
 * re-created after the last row by the repair path (admin load / writes), so hand-pasted
 * candidates are found instead of silently dropped. A read alone finds none and writes nothing.
 */
function testMissingResponsesSectionSelfHealsAndReturnsCandidates() {
  const { SM, ss, sandbox } = loadBallotModel();
  const sheet = createBallotWithoutResponsesSection(SM, ss, 'NoResponses');
  sheet.getRange(13, 2, 2, 2).setValues([['Alice', 'first'], ['Bob', 'second']]);

  sheet._writeCalls = 0;
  assert.deepEqual(SM.readBallotCandidateDetails_(sheet), []);
  assert.equal(sheet._writeCalls, 0);
  SM._repairBallotSheet_(sheet);
  assert.deepEqual(SM.readBallotCandidateDetails_(sheet).map((it) => it.name), ['Alice', 'Bob']);
  assert.equal(sheet.getRange(15, 1).getValue(), '[Responses]');
  assert.deepEqual(sheet.getRange(16, 1, 1, 4).getValues()[0], ['Date', 'Name', 'Weight', 'Comment']);
  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Alice', 'Bob']);
  assert.equal(logLinesWith(sandbox, 'ballot.repair.responsesSection').length, 1);
  assertSnapshotMatchesOldReaders(SM, sheet, 'testMissingResponsesSectionSelfHealsAndReturnsCandidates');
}

/**
 * Column-A text below the Candidates header with no [Responses] marker could be orphaned
 * response rows — never guess: write nothing, log why no candidates were found.
 */
function testMissingResponsesSectionWithColumnAContentIsNotRepaired() {
  const { SM, ss, sandbox } = loadBallotModel();
  const sheet = createBallotWithoutResponsesSection(SM, ss, 'NoResponsesColA');
  sheet.getRange(13, 1).setValue('Alice');

  assert.deepEqual(SM.readBallotCandidateDetails_(sheet), []);
  SM._repairBallotSheet_(sheet); // the repair path refuses to guess
  assert.equal(sheet.getLastRow(), 13);
  assert.equal(logLinesWith(sandbox, 'ballot.parse.responsesMarkerMissing').length > 0, true);
  assertSnapshotMatchesOldReaders(SM, sheet, 'testMissingResponsesSectionWithColumnAContentIsNotRepaired');
}

/** A candidate typed into column A is ignored (col A is marker-only) but logged. */
function testCandidateInColumnAIsLoggedAndIgnored() {
  const { SM, ss, sandbox } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'ColA');
  SM.writeBallotCandidateDetails_(sheet, [{ name: 'Alice', details: '' }, { name: 'Bob', details: '' }]);
  sheet.getRange(14, 1).setValue('Carol'); // beside Bob, in the reserved column

  assert.deepEqual(SM.readBallotCandidateDetails_(sheet).map((it) => it.name), ['Alice', 'Bob']);
  const lines = logLinesWith(sandbox, 'ballot.parse.candidateInColumnA');
  assert.equal(lines.length > 0, true);
  assert.equal(lines[0].indexOf('14') !== -1, true);
  assertSnapshotMatchesOldReaders(SM, sheet, 'testCandidateInColumnAIsLoggedAndIgnored');
}

/** Every cache rebuild logs what discovery found. */
function testCacheRefreshLogsDiscoverySummary() {
  const { SM, ss, sandbox } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'Summary');
  SM.writeBallotCandidateDetails_(sheet, [{ name: 'Alice', details: '' }]);
  SM.refreshBallotCache_('Summary', sheet);

  const lines = logLinesWith(sandbox, 'ballot.cache.refresh');
  assert.equal(lines.length, 1);
  assert.equal(lines[0].indexOf('"candidateCount":1') !== -1, true);
}

/**
 * Instruments a sheet so every mutating call (setValue/setValues/clearContent via
 * getRange, plus insert/delete rows) records whether the script lock was held. Returns
 * { writes, unlocked } — both arrays of op names.
 */
function spyWrites(sheet, lockState) {
  const log = { writes: [], unlocked: [] };
  function note(op) {
    log.writes.push(op);
    if (!lockState.held) log.unlocked.push(op);
  }
  const origGetRange = sheet.getRange.bind(sheet);
  sheet.getRange = function () {
    const range = origGetRange.apply(null, arguments);
    ['setValue', 'setValues', 'clearContent'].forEach(function (m) {
      const orig = range[m].bind(range);
      range[m] = function () { note(m); return orig.apply(null, arguments); };
    });
    return range;
  };
  ['insertRowsBefore', 'insertRowsAfter', 'deleteRows'].forEach(function (m) {
    const orig = sheet[m].bind(sheet);
    sheet[m] = function () { note(m); return orig.apply(null, arguments); };
  });
  return log;
}

function assertAllWritesLocked(log, label) {
  assert.ok(log.writes.length > 0, label + ': expected the scenario to write something');
  assert.deepEqual(log.unlocked, [], label + ': writes outside the script lock');
}

/** _withBallotLock_ is re-entrant (one waitLock/releaseLock pair) and releases on throw. */
function testWithBallotLockIsReentrantAndReleasesOnThrow() {
  const { SM, sandbox } = loadBallotModel();
  const state = sandbox.LockService._state;

  const result = SM._withBallotLock_(function () {
    assert.equal(state.held, true);
    return SM._withBallotLock_(function () {
      return SM._withBallotLock_(function () { return 42; });
    });
  });
  assert.equal(result, 42);
  assert.equal(state.waitCount, 1);
  assert.equal(state.releaseCount, 1);
  assert.equal(state.held, false);

  assert.throws(function () {
    SM._withBallotLock_(function () {
      SM._withBallotLock_(function () { throw new Error('boom'); });
    });
  }, /boom/);
  assert.equal(state.waitCount, 2);
  assert.equal(state.releaseCount, 2);
  assert.equal(state.held, false);

  // Depth was reset by the throw: a later call acquires again.
  SM._withBallotLock_(function () {});
  assert.equal(state.waitCount, 3);
}

/** Every model-level mutating entry point does its sheet writes inside the lock. */
function testModelEntryPointsWriteUnderLock() {
  const { SM, ss, sandbox } = loadBallotModel();
  const state = sandbox.LockService._state;
  const sheet = SM.createNewBallot_(ss, 'LockCover');
  const log = spyWrites(sheet, state);

  SM.addBallotCandidateForId_('LockCover', 'Alice', 'a');
  assertAllWritesLocked(log, 'addBallotCandidateForId_');
  log.writes.length = 0;

  SM.addBallotCandidateForAdmin_('LockCover', 'Bob', 'b', '', '');
  assertAllWritesLocked(log, 'addBallotCandidateForAdmin_');
  log.writes.length = 0;

  SM.submitBallotResponse_('LockCover', 'Voter', ['Bob', 'Alice'], '');
  assertAllWritesLocked(log, 'submitBallotResponse_');
  log.writes.length = 0;

  SM.saveBallotCandidatesForId_('LockCover', [
    { name: 'Alice', details: 'x', linkText: '', linkUrl: '' },
    { name: 'Bob', details: 'y', linkText: '', linkUrl: '' },
  ]);
  assertAllWritesLocked(log, 'saveBallotCandidatesForId_');
  assert.equal(state.held, false);
}

/** Every repair write runs inside the lock (repairs run only from admin load / write paths). */
function testSelfRepairWritesRunUnderLock() {
  // Missing [Candidates] section (old sheet) + legacy bare marker text.
  let t = loadBallotModel();
  let sheet = t.ss.insertSheet('Ballot-RepairA');
  sheet.getRange(1, 1, 6, 2).setValues([
    ['Title', 'T'], ['Description', 'D'], ['Footer', 'F'],
    ['Contact', 'C'], ['Accept-New', 'TRUE'], ['Admin-Only-Notes', ''],
  ]);
  sheet.getRange(8, 1).setValue('[Results]');
  sheet.getRange(10, 1).setValue('[Responses]');
  sheet.getRange(11, 1, 1, 6).setValues([['Date', 'Name', 'Weight', 'Comment', 'Alice', 'Bob']]);
  let log = spyWrites(sheet, t.sandbox.LockService._state);
  t.SM._ensureCandidatesSection_(sheet);
  assertAllWritesLocked(log, '_ensureCandidatesSection_');

  // Lost [Responses] section.
  t = loadBallotModel();
  sheet = t.SM.createNewBallot_(t.ss, 'RepairB');
  sheet.deleteRows(13, 2);
  log = spyWrites(sheet, t.sandbox.LockService._state);
  t.SM._ensureResponsesSection_(sheet);
  assertAllWritesLocked(log, '_ensureResponsesSection_');

  // Old column-A Candidates layout.
  t = loadBallotModel();
  sheet = t.ss.insertSheet('Ballot-RepairC');
  sheet.getRange(8, 1).setValue('[Results]');
  sheet.getRange(9, 1).setValue('[Candidates]');
  sheet.getRange(10, 1, 1, 2).setValues([['Name', 'Details']]);
  sheet.getRange(11, 1, 1, 2).setValues([['book 1', 'd1']]);
  sheet.getRange(12, 1).setValue('[Responses]');
  sheet.getRange(13, 1, 1, 5).setValues([['Date', 'Name', 'Weight', 'Comment', 'book 1']]);
  log = spyWrites(sheet, t.sandbox.LockService._state);
  t.SM._migrateCandidatesColumnsIfNeeded_(sheet);
  assertAllWritesLocked(log, '_migrateCandidatesColumnsIfNeeded_');

  // Missing trailing Candidates header columns.
  t = loadBallotModel();
  sheet = t.SM.createNewBallot_(t.ss, 'RepairD');
  sheet.getRange(12, 4, 1, 2).clearContent();
  log = spyWrites(sheet, t.sandbox.LockService._state);
  t.SM._ensureCandidatesHeaderColumns_(sheet);
  assertAllWritesLocked(log, '_ensureCandidatesHeaderColumns_');

  // Legacy 'Info' config key rename.
  t = loadBallotModel();
  sheet = t.SM.createNewBallot_(t.ss, 'RepairE');
  sheet.getRange(8, 1).setValue('Info');
  log = spyWrites(sheet, t.sandbox.LockService._state);
  t.SM._repairBallotSheet_(sheet);
  assertAllWritesLocked(log, '_repairBallotSheet_ legacy key');

  // Candidate typed into the Candidates table only: Responses header reconcile.
  t = loadBallotModel();
  sheet = t.SM.createNewBallot_(t.ss, 'RepairF');
  t.SM.writeBallotCandidateDetails_(sheet, [{ name: 'Alice', details: '' }]);
  log = spyWrites(sheet, t.sandbox.LockService._state);
  t.SM._repairBallotSheet_(sheet);
  assertAllWritesLocked(log, '_repairBallotSheet_ rank columns');

  // Every damaged fixture through the one repair entry.
  [['header on marker row', (ss) => writeGrid(ss.insertSheet('Ballot-RepairG'), prodAbc2026_7Rows())],
    ['headerless Candidates', (ss) => writeGrid(ss.insertSheet('Ballot-RepairH'), sitAbc2026_08Rows())]].forEach(([label, make]) => {
    t = loadBallotModel();
    sheet = make(t.ss);
    log = spyWrites(sheet, t.sandbox.LockService._state);
    t.SM._repairBallotSheet_(sheet);
    assertAllWritesLocked(log, '_repairBallotSheet_ ' + label);
  });

  // A healthy sheet read takes no lock at all.
  t = loadBallotModel();
  sheet = t.SM.createNewBallot_(t.ss, 'Healthy');
  t.SM.readBallotCandidates_(sheet);
  t.SM.readBallotConfig_(sheet);
  assert.equal(t.sandbox.LockService._state.waitCount, 0);
}

/** New ballots are not frozen at all (the old fixed 14-row freeze is gone). */
function testNewBallotHasNoFrozenRows() {
  const { SM, ss } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'NoFreeze');
  assert.equal(sheet.getFrozenRows(), 0);
}

/** The legacy freeze-at-14 is undone by _highlightSectionMarkers_ only when the Responses header isn't row 14. */
function testHighlightSectionMarkersUnfreezesStaleFourteenRowFreeze() {
  const { SM, ss, sandbox } = loadBallotModel();
  const sheet = SM.createNewBallot_(ss, 'StaleFreeze');
  SM.writeBallotResults_(sheet, [['a'], ['b'], ['c']]); // pushes the Responses header below row 14
  sheet.setFrozenRows(14);

  SM._highlightSectionMarkers_(sheet);
  assert.equal(sheet.getFrozenRows(), 0);
  const lines = logLinesWith(sandbox, 'ballot.repair.frozenRows');
  assert.equal(lines.length, 1);
  assert.equal(lines[0].indexOf('StaleFreeze') !== -1, true);

  // Header genuinely at row 14: a deliberate freeze, left alone.
  const same = SM.createNewBallot_(ss, 'SameFreeze');
  same.setFrozenRows(14);
  SM._highlightSectionMarkers_(same);
  assert.equal(same.getFrozenRows(), 14);

  // Any other value is a user choice.
  const other = SM.createNewBallot_(ss, 'OtherFreeze');
  SM.writeBallotResults_(other, [['a'], ['b'], ['c']]);
  other.setFrozenRows(3);
  SM._highlightSectionMarkers_(other);
  assert.equal(other.getFrozenRows(), 3);
  assert.equal(logLinesWith(sandbox, 'ballot.repair.frozenRows').length, 1);
}

// ---------------------------------------------------------------------------
// Order-independent layout (rcballot-kix): hand-built sheets in any marker order.
// ---------------------------------------------------------------------------

const SECTION_MARKERS = { results: '[Results]', candidates: '[Candidates]', responses: '[Responses]' };

/** All 6 orderings of the three section keys. */
const SECTION_ORDERS = (function () {
  const keys = ['results', 'candidates', 'responses'];
  const out = [];
  keys.forEach((a) => keys.forEach((b) => keys.forEach((c) => {
    if (a !== b && b !== c && a !== c) out.push([a, b, c]);
  })));
  return out;
}());

/**
 * Rows (from column A) for each section of a hand-built ballot. Candidates and Responses
 * agree on candidate names/order, so no reconcile write fires on read. `blankRows` puts a
 * blank row in the middle of every section's data (never at the very end of the sheet,
 * where it would be invisible to getLastRow anyway).
 */
function sectionRows(key, blankRows) {
  const blank = blankRows ? [[]] : [];
  if (key === 'results') {
    return [['[Results]'], ['', 'Analysis run: earlier']].concat(blank, [['', 'Winner: Alice', 3]]);
  }
  if (key === 'candidates') {
    return [['[Candidates]'], ['', 'Name', 'Details', 'Link Text', 'Link URL'], ['', 'Alice', 'a note']]
      .concat(blank, [['', 'Bob', 'b note', 'site', 'https://example.com/b']]);
  }
  return [['[Responses]'], ['Date', 'Name', 'Weight', 'Comment', 'Alice', 'Bob'], ['2026-10-01', 'V1', 1, '', 1, 2]]
    .concat(blank, [['2026-10-02', 'V2', 2, 'c2', 2, 1]]);
}

/**
 * Builds Ballot-<id> by hand: 8 config rows, a blank spacer, then the three sections in
 * `order`. Only non-empty cells are written, so getLastRow matches real Sheets.
 */
function buildOrderedBallot(ss, id, order, opts) {
  const sheet = ss.insertSheet('Ballot-' + id);
  const rows = orderedBallotRows(id, order, opts);
  rows.forEach((r, i) => r.forEach((v, c) => { if (v !== '' && v != null) sheet.getRange(i + 1, c + 1).setValue(v); }));
  return sheet;
}

/** The plain 2D grid buildOrderedBallot writes (rows may be ragged/empty). */
function orderedBallotRows(id, order, opts) {
  const rows = [
    ['Title', 'T-' + id], ['Description', 'D'], ['Instructions', 'I'], ['Footer', 'F'],
    ['Contact', 'C'], ['Accept-New', 'TRUE'], ['Add-Instructions', 'AI'], ['Admin-Only-Notes', 'N'], [],
  ];
  order.forEach((key) => { sectionRows(key, opts && opts.blankRows).forEach((r) => rows.push(r)); });
  return rows;
}

/** Every row whose column A is exactly one of the three markers, as {key: [rows]}. */
function scanMarkers(sheet) {
  const out = { results: [], candidates: [], responses: [] };
  const last = sheet.getLastRow();
  const col = last ? sheet.getRange(1, 1, last, 1).getValues() : [];
  col.forEach((r, i) => {
    Object.keys(SECTION_MARKERS).forEach((k) => { if (String(r[0]).trim() === SECTION_MARKERS[k]) out[k].push(i + 1); });
  });
  return out;
}

const SNAPSHOT_COLS = 12;

/** Trims trailing blank cells/rows so a snapshot ignores sheet extent, not content. */
function trimRows(rows) {
  const trimmed = rows.map((r) => {
    const copy = r.slice();
    while (copy.length && (copy[copy.length - 1] === '' || copy[copy.length - 1] == null)) copy.pop();
    return copy;
  });
  while (trimmed.length && !trimmed[trimmed.length - 1].length) trimmed.pop();
  return trimmed;
}

/**
 * Values of one section, marker row through the row before the next marker (any kind) or
 * the last row — computed independently of the model's own layout code.
 */
function sectionSnapshot(sheet, key) {
  const markers = scanMarkers(sheet);
  const start = markers[key][0];
  assert.ok(start, 'marker ' + key + ' not found');
  const all = [].concat(markers.results, markers.candidates, markers.responses).sort((a, b) => a - b);
  const next = all.find((r) => r > start);
  const end = next ? next - 1 : sheet.getLastRow();
  return trimRows(sheet.getRange(start, 1, end - start + 1, SNAPSHOT_COLS).getValues());
}

function configSnapshot(sheet) {
  return trimRows(sheet.getRange(1, 1, 9, 2).getValues());
}

function assertMarkersExactlyOnce(sheet, label) {
  const m = scanMarkers(sheet);
  Object.keys(m).forEach((k) => assert.equal(m[k].length, 1, label + ': ' + k + ' marker rows ' + JSON.stringify(m[k])));
}

/** Everything the four section readers return, with sheet row numbers dropped. */
function readAllSections(SM, sheet) {
  return {
    config: SM.readBallotConfig_(sheet),
    details: SM.readBallotCandidateDetails_(sheet),
    candidates: SM.readBallotCandidates_(sheet),
    responses: SM.readBallotResponseRows_(sheet).map((r) => ({ date: r.date, name: r.name, weight: r.weight, comment: r.comment, ranks: r.ranks })),
  };
}

/** kix.4 AC1/AC2/AC4: every marker order (with and without blank rows) reads the same data. */
function testAllMarkerOrdersReadTheSameData() {
  const canonical = (function () {
    const { SM, ss } = loadBallotModel();
    return readAllSections(SM, buildOrderedBallot(ss, 'Canon', ['results', 'candidates', 'responses']));
  }());
  assert.deepEqual(canonical.candidates, ['Alice', 'Bob']);
  assert.deepEqual(canonical.responses.map((r) => r.name), ['V1', 'V2']);
  assert.equal(canonical.config.Title, 'T-Canon');

  SECTION_ORDERS.forEach((order) => {
    [false, true].forEach((blankRows) => {
      const { SM, ss, sandbox } = loadBallotModel();
      const sheet = buildOrderedBallot(ss, 'Canon', order, { blankRows: blankRows });
      const label = order.join('/') + (blankRows ? ' +blank rows' : '');
      const lastRowBefore = sheet.getLastRow();
      assert.deepEqual(readAllSections(SM, sheet), canonical, label);
      // A healthy sheet in any order needs no repair: nothing written, no lock taken.
      assert.equal(sheet.getLastRow(), lastRowBefore, label + ': rows changed');
      assert.equal(sandbox.LockService._state.waitCount, 0, label + ': lock taken');
    });
  });
}

/** kix.4 AC3: a duplicated marker uses the first occurrence, bounds the section above it, and is logged once. */
function testDuplicateMarkerUsesFirstOccurrenceAndIsLogged() {
  const { SM, ss, sandbox } = loadBallotModel();
  const sheet = buildOrderedBallot(ss, 'Dup', ['results', 'responses', 'candidates']);
  // A stray copy of [Responses] (plus a would-be vote under it) at the bottom of Candidates.
  const dupRow = sheet.getLastRow() + 1;
  sheet.getRange(dupRow, 1).setValue('[Responses]');
  sheet.getRange(dupRow + 1, 1, 1, 6).setValues([['2026-10-03', 'Stray', 1, '', 1, 2]]);

  const layout = SM._readBallotLayout_(sheet);
  assert.equal(layout.markers.responses, 13); // first occurrence: config 1..9, Results 10..12, Responses 13
  assert.deepEqual(layout.duplicates.map((d) => d.marker), ['[Responses]']);
  assert.equal(layout.sections.candidates.endRow, dupRow - 1); // the duplicate bounds Candidates

  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Alice', 'Bob']);
  assert.deepEqual(SM.readBallotResponseRows_(sheet).map((r) => r.name), ['V1', 'V2']); // 'Stray' never read
  const lines = logLinesWith(sandbox, 'ballot.parse.duplicateMarker');
  assert.equal(lines.length, 1); // many layout reads, one log line
  assert.equal(lines[0].indexOf('"rows":[13,' + dupRow + ']') !== -1, true, lines[0]);
}

/** kix.4 AC4: with Responses not last, the next section's rows are not read as votes. */
function testResponsesSectionEndsAtNextMarker() {
  const { SM, ss } = loadBallotModel();
  const sheet = buildOrderedBallot(ss, 'RespMid', ['candidates', 'responses', 'results']);
  // Results output in column B of the following section would read as voter names if
  // Responses ran to the end of the sheet.
  const layout = SM._readBallotLayout_(sheet);
  assert.equal(layout.sections.responses.endRow, layout.markers.results - 1);
  assert.equal(layout.sections.results.endRow, sheet.getLastRow());
  assert.equal(layout.configEndRow, 9);
  assert.deepEqual(SM.readBallotResponseRows_(sheet).map((r) => r.name), ['V1', 'V2']);
}

/** kix.4 AC5, tightened by 8i5.3: each public reader is exactly one getValues (the snapshot read), no column-A rescans. */
function testSectionReadersReadTheSheetOnce() {
  const { SM, ss } = loadBallotModel();
  const sheet = buildOrderedBallot(ss, 'OneScan', ['responses', 'results', 'candidates']);
  ['readBallotConfig_', 'readBallotCandidateDetails_', 'readBallotCandidates_', 'readBallotResponseRows_',
    'countUniqueBallotRespondents_', 'findRespondentsWithNewCandidates_'].forEach((fn) => {
    sheet._getValuesCalls = 0;
    SM[fn](sheet);
    assert.equal(sheet._getValuesCalls, 1, fn + ' getValues calls');
  });
}

/**
 * kix.5 AC1/AC2: with Responses not the last section, a new vote lands inside Responses
 * (the following section shifts down one row intact) and a re-vote overwrites in place.
 */
function testNewVoteLandsInsideResponsesWhenNotLast() {
  ['results', 'candidates'].forEach((after) => {
    const { SM, ss } = loadBallotModel();
    const order = after === 'results' ? ['candidates', 'responses', 'results'] : ['results', 'responses', 'candidates'];
    const sheet = buildOrderedBallot(ss, 'Mid' + after, order);
    const before = SM._readBallotLayout_(sheet);
    const nextRow = before.markers[after];
    const nextBefore = sheet.getRange(nextRow, 1, before.lastRow - nextRow + 1, 6).getValues();

    SM.submitBallotResponse_('Mid' + after, 'V3', ['Bob', 'Alice'], '');
    const layout = SM._readBallotLayout_(sheet);
    assert.equal(layout.markers[after], nextRow + 1, after + ' marker shifted by one');
    assert.deepEqual(sheet.getRange(nextRow + 1, 1, before.lastRow - nextRow + 1, 6).getValues(), nextBefore, after + ' contents intact');
    assert.equal(sheet.getRange(nextRow, 2).getValue(), 'V3'); // the old marker row is now the new vote
    assert.deepEqual(SM.readBallotResponseRows_(sheet).map((r) => r.name), ['V1', 'V2', 'V3']);

    // Re-vote: same row, no row inserted.
    const rowsBefore = sheet.getLastRow();
    SM.submitBallotResponse_('Mid' + after, 'v2', ['Alice', 'Bob'], 'again');
    assert.equal(sheet.getLastRow(), rowsBefore);
    const v2 = SM.readBallotResponseRows_(sheet).filter((r) => r.name.toLowerCase() === 'v2');
    assert.equal(v2.length, 1);
    assert.equal(v2[0].row, nextRow - 1);
    assert.equal(v2[0].comment, 'again');
    assert.equal(v2[0].weight, 2); // existing weight kept
  });
}

/** A new vote goes right below the last named row, filling a blank row before the next marker. */
function testNewVoteFillsBlankRowBeforeNextMarker() {
  const { SM, ss } = loadBallotModel();
  const sheet = buildOrderedBallot(ss, 'Gap', ['responses', 'results', 'candidates']);
  const resultsRow = SM._readBallotLayout_(sheet).markers.results;
  sheet.insertRowsBefore(resultsRow, 1); // a spare blank row at the end of Responses
  SM.submitBallotResponse_('Gap', 'V3', ['Alice', 'Bob'], '');
  assert.equal(sheet.getRange(resultsRow, 2).getValue(), 'V3');
  assert.equal(SM._readBallotLayout_(sheet).markers.results, resultsRow + 1); // no extra insert
}

/**
 * A trailing Responses row with content but no name (e.g. a stray hand-typed note) is
 * not a vote, but it isn't empty either — a new vote goes below it, never over it.
 */
function testNewVoteNeverOverwritesUnnamedRowWithContent() {
  const { SM, ss } = loadBallotModel();
  const sheet = buildOrderedBallot(ss, 'Stray', ['responses', 'results', 'candidates']);
  const resultsRow = SM._readBallotLayout_(sheet).markers.results;
  sheet.insertRowsBefore(resultsRow, 1);
  sheet.getRange(resultsRow, 4).setValue('stray note'); // Comment column, no name
  SM.submitBallotResponse_('Stray', 'V3', ['Alice', 'Bob'], '');
  assert.equal(sheet.getRange(resultsRow, 4).getValue(), 'stray note');
  assert.equal(sheet.getRange(resultsRow + 1, 2).getValue(), 'V3');
  assert.equal(SM._readBallotLayout_(sheet).markers.results, resultsRow + 2);
}

/**
 * addBallotCandidate_ takes the Responses header row only after the read that may repair
 * the sheet — on an old sheet with no [Candidates] section the repair inserts two rows
 * above [Responses], and a row number taken earlier would put the name in the wrong row.
 */
function testAddBallotCandidateAfterRepairWritesResponsesHeader() {
  const { SM, ss } = loadBallotModel();
  const sheet = ss.insertSheet('Ballot-OldAdd');
  sheet.getRange(1, 1, 2, 2).setValues([['Title', 'T'], ['Accept-New', 'TRUE']]);
  sheet.getRange(8, 1).setValue('[Results]');
  sheet.getRange(10, 1).setValue('[Responses]');
  sheet.getRange(11, 1, 1, 5).setValues([['Date', 'Name', 'Weight', 'Comment', 'Alice']]);

  SM.addBallotCandidate_(sheet, 'Bob', '');
  const headerRow = SM._readBallotLayout_(sheet).sections.responses.headerRow;
  assert.equal(headerRow, 15); // +2 Candidates marker/header, +2 Candidates rows (Alice backfilled, Bob)
  assert.deepEqual(sheet.getRange(headerRow, 1, 1, 6).getValues()[0], ['Date', 'Name', 'Weight', 'Comment', 'Alice', 'Bob']);
  assert.deepEqual(sheet.getRange(11, 2, 1, 6).getValues()[0], ['Name', 'Details', 'Link Text', 'Link URL', '', '']); // Candidates header untouched
  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Alice', 'Bob']);
  assert.deepEqual(SM.readBallotCandidateDetails_(sheet).map((it) => it.name), ['Alice', 'Bob']);
}

/** kix.5 AC3: no script code appends at getLastRow()+1 (only section-repair code may append). */
function testNoGetLastRowPlusOneWrites() {
  const dir = path.join(__dirname, '..', 'script');
  const hits = [];
  fs.readdirSync(dir).filter((f) => f.endsWith('.js')).forEach((f) => {
    fs.readFileSync(path.join(dir, f), 'utf8').split('\n').forEach((line, i) => {
      const code = line.replace(/\/\/.*$/, '');
      if (/^\s*\*/.test(code)) return; // JSDoc/comment line
      if (/getLastRow\(\)\s*\+\s*1/.test(code)) hits.push(f + ':' + (i + 1));
    });
  });
  assert.deepEqual(hits, []);
}

/** Row of the first column-A cell exactly equal to text, or -1. */
function rowOf(sheet, text) {
  const col = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  for (let i = 0; i < col.length; i++) if (col[i][0] === text) return i + 1;
  return -1;
}

/** kix.6 AC1: each marker, removed alone, is restored per its rule and logged — by the repair path (admin load / writes), not by reads. */
function testEachMissingMarkerIsRestoredAndLogged() {
  // [Candidates] marker deleted: an empty section goes in directly above [Responses].
  let t = loadBallotModel();
  let sheet = buildOrderedBallot(t.ss, 'NoCand', ['results', 'candidates', 'responses']);
  const responsesRow = rowOf(sheet, '[Responses]');
  sheet.getRange(rowOf(sheet, '[Candidates]'), 1).setValue('');
  sheet._writeCalls = 0;
  t.SM.readBallotCandidates_(sheet);
  assert.equal(sheet._writeCalls, 0, 'read repaired');
  t.SM._repairBallotSheet_(sheet);
  assert.equal(sheet.getRange(responsesRow, 1).getValue(), '[Candidates]');
  assert.deepEqual(sheet.getRange(responsesRow + 1, 2, 1, 2).getValues()[0], ['Name', 'Details']);
  // The new, empty table is then backfilled from the Responses header's rank columns.
  assert.deepEqual(sheet.getRange(responsesRow + 2, 2, 2, 1).getValues(), [['Alice'], ['Bob']]);
  assert.equal(sheet.getRange(responsesRow + 4, 1).getValue(), '[Responses]');
  assert.equal(logLinesWith(t.sandbox, 'ballot.repair.candidatesFromResponses').length, 1);
  let lines = logLinesWith(t.sandbox, 'ballot.repair.candidatesSection');
  assert.equal(lines.length, 1);
  assert.equal(lines[0].indexOf('"row":' + responsesRow) !== -1, true, lines[0]);

  // [Candidates] and [Responses] both absent: Candidates is appended at the end, then
  // Responses below it (column A under the new Candidates header is blank).
  t = loadBallotModel();
  sheet = buildOrderedBallot(t.ss, 'OnlyResults', ['results']);
  const last = sheet.getLastRow();
  t.SM._repairBallotSheet_(sheet);
  assert.equal(sheet.getRange(last + 1, 1).getValue(), '[Candidates]');
  assert.equal(sheet.getRange(last + 3, 1).getValue(), '[Responses]');
  assert.equal(logLinesWith(t.sandbox, 'ballot.repair.candidatesSection').length, 1);
  assert.equal(logLinesWith(t.sandbox, 'ballot.repair.responsesSection').length, 1);

  // [Responses] section gone, Results is the bottom section with column A blank below its
  // marker: Responses is appended after the last row.
  t = loadBallotModel();
  sheet = buildOrderedBallot(t.ss, 'NoResp', ['candidates', 'results']);
  const lastRow = sheet.getLastRow();
  t.SM._repairBallotSheet_(sheet);
  assert.deepEqual(t.SM.readBallotCandidateDetails_(sheet).map((it) => it.name), ['Alice', 'Bob']);
  assert.equal(sheet.getRange(lastRow + 1, 1).getValue(), '[Responses]');
  assert.deepEqual(sheet.getRange(lastRow + 2, 1, 1, 4).getValues()[0], ['Date', 'Name', 'Weight', 'Comment']);
  lines = logLinesWith(t.sandbox, 'ballot.repair.responsesSection');
  assert.equal(lines.length, 1);
  assert.equal(lines[0].indexOf('"row":' + (lastRow + 1)) !== -1, true, lines[0]);

  // [Results] marker deleted: re-inserted directly above the first present marker, in
  // canonical and non-canonical orders alike; no existing row changes section. (When
  // Results sat BELOW another section, deleting its marker already handed its rows to that
  // section; the repair cannot know where the marker was and does not move them back.)
  [['results', 'candidates', 'responses'], ['results', 'responses', 'candidates']].forEach((order) => {
    t = loadBallotModel();
    sheet = buildOrderedBallot(t.ss, 'NoRes', order);
    sheet.getRange(rowOf(sheet, '[Results]'), 1).setValue('');
    const firstMarker = rowOf(sheet, SECTION_MARKERS[order[1]]);
    t.SM._repairBallotSheet_(sheet);
    const before = t.SM.readBallotResponseRows_(sheet).map((r) => r.name);
    assert.equal(sheet.getRange(firstMarker, 1).getValue(), '[Results]', order.join('/'));
    assert.equal(sheet.getRange(firstMarker + 1, 1).getValue(), SECTION_MARKERS[order[1]]);
    assert.deepEqual(before, ['V1', 'V2']);
    assert.deepEqual(t.SM.readBallotCandidates_(sheet), ['Alice', 'Bob']);
    lines = logLinesWith(t.sandbox, 'ballot.repair.resultsSection');
    assert.equal(lines.length, 1, order.join('/'));
    assert.equal(lines[0].indexOf('"row":' + firstMarker) !== -1, true, lines[0]);
  });
}

/** kix.6 AC2: [Responses] marker lost but its rows remain (Date in column A) — no repair, logged. */
function testMissingResponsesMarkerWithOrphanedRowsIsNotRepaired() {
  ['results', 'candidates'].forEach((lastBefore) => {
    const { SM, ss, sandbox } = loadBallotModel();
    const order = lastBefore === 'results' ? ['results', 'candidates', 'responses'] : ['candidates', 'results', 'responses'];
    const sheet = buildOrderedBallot(ss, 'Orphan', order);
    sheet.getRange(rowOf(sheet, '[Responses]'), 1).setValue('');
    const lastRow = sheet.getLastRow();
    assert.deepEqual(SM.readBallotCandidateDetails_(sheet), []);
    SM._repairBallotSheet_(sheet);
    assert.equal(sheet.getLastRow(), lastRow);
    assert.equal(rowOf(sheet, '[Responses]'), -1);
    assert.equal(logLinesWith(sandbox, 'ballot.parse.responsesMarkerMissing').length > 0, true);
    assert.equal(logLinesWith(sandbox, 'ballot.repair.responsesSection').length, 0);
  });
}

function wholeSheet(sheet) {
  return { lastRow: sheet.getLastRow(), values: sheet.getRange(1, 1, sheet.getLastRow(), Math.max(sheet.getLastColumn(), 1)).getValues() };
}

/** kix.6 AC3: writeBallotResults_ refuses (throws + logs, writes nothing) on a damaged layout. */
function testWriteBallotResultsRefusesOnDamagedLayout() {
  const damage = {
    duplicateMarker: (sheet) => sheet.getRange(sheet.getLastRow() + 1, 1).setValue('[Candidates]'),
    absorbedCandidatesHeader: (sheet) => sheet.getRange(rowOf(sheet, '[Results]') + 1, 2, 1, 2).setValues([['Name', 'Details']]),
    absorbedResponsesHeader: (sheet) => sheet.getRange(rowOf(sheet, '[Results]') + 2, 1, 1, 2).setValues([['Date', 'Name']]),
  };
  Object.keys(damage).forEach((reason) => {
    const { SM, ss, sandbox } = loadBallotModel();
    const sheet = buildOrderedBallot(ss, 'Refuse', ['results', 'candidates', 'responses']);
    damage[reason](sheet);
    const before = wholeSheet(sheet);
    const log = spyWrites(sheet, sandbox.LockService._state);
    assert.throws(() => SM.writeBallotResults_(sheet, [['new output']]), /^Error: Ballot sheet layout is damaged — fix the markers before running analysis\.$/, reason);
    assert.deepEqual(log.writes, [], reason + ': refusal wrote');
    assert.deepEqual(wholeSheet(sheet), before, reason);
    const lines = logLinesWith(sandbox, 'ballot.write.refused');
    assert.equal(lines.length, 1, reason);
    assert.equal(lines[0].indexOf('"reason":"' + reason + '"') !== -1, true, lines[0]);
    assert.equal(lines[0].indexOf('"sheet":"Ballot-Refuse"') !== -1, true, lines[0]);
  });

  // A deleted [Candidates] marker: the repair inserts a fresh section above [Responses],
  // which leaves the old Candidates header inside the Results region — refused, and the
  // Results region (old candidate rows included) is left exactly as the repair left it.
  // writeBallotResults_ runs the full repair path (rcballot-8i5.4), which also refills the
  // empty new Candidates table from the Responses header — so `before` is taken after it.
  let t = loadBallotModel();
  let sheet = buildOrderedBallot(t.ss, 'LostCand', ['results', 'candidates', 'responses']);
  sheet.getRange(rowOf(sheet, '[Candidates]'), 1).setValue('');
  t.SM._repairBallotSheet_(sheet);
  let before = wholeSheet(sheet);
  assert.throws(() => t.SM.writeBallotResults_(sheet, [['x']]), /layout is damaged/);
  assert.deepEqual(wholeSheet(sheet), before);
  assert.equal(logLinesWith(t.sandbox, '"reason":"absorbedCandidatesHeader"').length, 1);

  // A [Responses] marker that could not be repaired: still missing after repair → refused.
  t = loadBallotModel();
  sheet = buildOrderedBallot(t.ss, 'LostResp', ['results', 'candidates', 'responses']);
  sheet.getRange(rowOf(sheet, '[Responses]'), 1).setValue('');
  before = wholeSheet(sheet);
  assert.throws(() => t.SM.writeBallotResults_(sheet, [['x']]), /layout is damaged/);
  assert.deepEqual(wholeSheet(sheet), before);
  assert.equal(logLinesWith(t.sandbox, '"reason":"markerMissing"').length, 1);

  // Healthy sheet in a non-canonical order: not refused.
  t = loadBallotModel();
  sheet = buildOrderedBallot(t.ss, 'Fine', ['responses', 'results', 'candidates']);
  t.SM.writeBallotResults_(sheet, [['ok']]);
  assert.equal(logLinesWith(t.sandbox, 'ballot.write.refused').length, 0);
}

/**
 * Invariant (rcballot-kix.7): for every marker order, growing/shrinking/rewriting one
 * section through its real write path leaves the config rows and the other sections'
 * values (relative to their markers) unchanged, and every marker is found exactly once.
 */
function testSectionWritesLeaveOtherSectionsUnchangedInAllOrders() {
  const cases = [
    { section: 'results', label: 'results grow', act: (SM, sheet) => SM.writeBallotResults_(sheet, [['r1'], ['r2'], ['r3'], ['r4'], ['r5']]),
      check: (SM, sheet) => assert.deepEqual(sectionSnapshot(sheet, 'results'), [['[Results]'], ['', 'r1'], ['', 'r2'], ['', 'r3'], ['', 'r4'], ['', 'r5']]) },
    { section: 'results', label: 'results shrink', act: (SM, sheet) => SM.writeBallotResults_(sheet, [['only']]),
      check: (SM, sheet) => assert.deepEqual(sectionSnapshot(sheet, 'results'), [['[Results]'], ['', 'only']]) },
    { section: 'candidates', label: 'candidates grow', act: (SM, sheet) => SM.writeBallotCandidateDetails_(sheet, [
      { name: 'Alice', details: 'a note' }, { name: 'Bob', details: 'b note', linkText: 'site', linkUrl: 'https://example.com/b' },
      { name: 'Carol', details: 'c' }, { name: 'Dan', details: 'd' }]),
      check: (SM, sheet) => assert.deepEqual(SM.readBallotCandidateDetails_(sheet).map((it) => it.name), ['Alice', 'Bob', 'Carol', 'Dan']) },
    { section: 'candidates', label: 'candidates shrink', act: (SM, sheet) => SM.writeBallotCandidateDetails_(sheet, [{ name: 'Alice', details: 'a note' }]),
      check: (SM, sheet) => assert.deepEqual(sectionSnapshot(sheet, 'candidates').slice(2), [['', 'Alice', 'a note']]) },
    { section: 'responses', label: 'responses grow (new vote)', act: (SM, sheet, id) => SM.submitBallotResponse_(id, 'V3', ['Bob', 'Alice'], 'new'),
      check: (SM, sheet) => assert.deepEqual(SM.readBallotResponseRows_(sheet).map((r) => r.name + ':' + r.ranks.join(',')), ['V1:1,2', 'V2:2,1', 'V3:2,1']) },
    { section: 'responses', label: 'responses re-vote in place', act: (SM, sheet, id) => SM.submitBallotResponse_(id, 'v1', ['Bob', 'Alice'], 'changed'),
      check: (SM, sheet) => {
        const rows = SM.readBallotResponseRows_(sheet);
        assert.deepEqual(rows.map((r) => r.name + ':' + r.ranks.join(',')), ['v1:2,1', 'V2:2,1']);
      } },
    // Grows Candidates by a row and Responses by a header column: only Results is a bystander.
    { section: 'candidates', also: 'responses', label: 'addBallotCandidate_', act: (SM, sheet) => SM.addBallotCandidate_(sheet, 'Carol', 'c'),
      check: (SM, sheet) => assert.deepEqual(SM.readBallotCandidates_(sheet), ['Alice', 'Bob', 'Carol']) },
  ];

  SECTION_ORDERS.forEach((order) => {
    // Each section's grow-then-shrink runs as a sequence on one sheet; other cases get a fresh one.
    const sequences = [
      cases.filter((c) => c.section === 'results'),
      cases.filter((c) => c.section === 'candidates' && !c.also),
      cases.slice(4, 5), cases.slice(5, 6),
      cases.filter((c) => c.also),
    ];
    sequences.forEach((seq) => {
      const { SM, ss } = loadBallotModel();
      const id = 'Inv' + order.map((k) => k[0]).join('');
      const sheet = buildOrderedBallot(ss, id, order);
      seq.forEach((c) => {
        const label = order.join('/') + ' ' + c.label;
        const bystanders = Object.keys(SECTION_MARKERS).filter((k) => k !== c.section && k !== c.also);
        const before = bystanders.map((k) => sectionSnapshot(sheet, k));
        const configBefore = configSnapshot(sheet);
        c.act(SM, sheet, id);
        assertMarkersExactlyOnce(sheet, label);
        assert.deepEqual(bystanders.map((k) => sectionSnapshot(sheet, k)), before, label + ': other sections changed');
        assert.deepEqual(configSnapshot(sheet), configBefore, label + ': config rows changed');
        c.check(SM, sheet);
        assertMarkersExactlyOnce(sheet, label + ' (after read)');
      });
    });
  });
}

/** writeBallotConfig_ inserts a missing key above the first marker, whichever section it is. */
function testWriteBallotConfigInsertsAboveFirstMarkerInAnyOrder() {
  SECTION_ORDERS.forEach((order) => {
    const { SM, ss } = loadBallotModel();
    const sheet = buildOrderedBallot(ss, 'Cfg', order);
    sheet.getRange(9, 1, 1, 2).setValues([['Spare', 'x']]); // no blank spacer row left to reuse
    const others = Object.keys(SECTION_MARKERS).map((k) => sectionSnapshot(sheet, k));
    SM.writeBallotConfig_(sheet, { Title: 'New', 'Extra-Key': 'v' });
    assert.equal(sheet.getRange(10, 1).getValue(), 'Extra-Key', order.join('/'));
    assert.equal(sheet.getRange(11, 1).getValue(), SECTION_MARKERS[order[0]], order.join('/'));
    assert.deepEqual(Object.keys(SECTION_MARKERS).map((k) => sectionSnapshot(sheet, k)), others, order.join('/'));
    const config = SM.readBallotConfig_(sheet);
    assert.equal(config.Title, 'New');
    assert.equal(config['Extra-Key'], 'v');
  });
}

// ---------------------------------------------------------------------------
// Snapshot reader (rcballot-8i5.1): _parseBallot_ / _readBallot_ and derivations.
// ---------------------------------------------------------------------------

/** The old readers' output, ranks re-keyed by Responses-header name (old ranks are positional over it). */
function oldReadersView(SM, sheet) {
  const names = SM.readBallotCandidates_(sheet);
  return {
    config: SM.readBallotConfig_(sheet),
    details: SM.readBallotCandidateDetails_(sheet),
    candidates: names,
    responses: SM.readBallotResponseRows_(sheet).map((r) => {
      const ranks = {};
      names.forEach((n, i) => { ranks[n] = r.ranks[i]; });
      return { row: r.row, date: r.date, name: r.name, weight: r.weight, comment: r.comment, ranks: ranks };
    }),
  };
}

/** The same view derived from a snapshot. */
function snapshotView(SM, snap) {
  const config = {};
  Object.keys(snap.config.byKey).forEach((k) => { config[k] = snap.config.byKey[k].value; });
  const cands = SM._snapCandidates_(snap);
  return {
    config: config,
    details: cands.map((c) => ({ name: c.name, details: c.details, linkText: c.linkText, linkUrl: c.linkUrl })),
    candidates: cands.map((c) => c.name),
    responses: SM._snapResponses_(snap).map((r) => ({ row: r.row, date: r.date, name: r.name, weight: r.weight, comment: r.comment, ranks: r.ranks })),
  };
}

/**
 * 8i5.1 AC2: snapshot-derived config/candidates/details/ranks-by-name equal the old readers'.
 * The old readers run first (they may self-repair an old-format sheet); the snapshot then
 * reads the sheet they leave behind.
 */
function assertSnapshotMatchesOldReaders(SM, sheet, label) {
  const old = oldReadersView(SM, sheet);
  const snap = SM._readBallot_(sheet);
  const view = snapshotView(SM, snap);
  assert.deepEqual(view, old, label);
  // Strict too (types matter: 1 vs '1'); JSON drops the vm-realm prototypes strict mode would trip on.
  assert.deepStrictEqual(JSON.parse(JSON.stringify(view)), JSON.parse(JSON.stringify(old)), label + ' (strict)');
  assert.deepEqual(SM._snapOrphanColumns_(snap), [], label + ': orphans');
  return snap;
}

/** Pads a ragged grid to a rectangle, as getDataRange().getValues() would return it. */
function rectGrid(rows) {
  const width = Math.max.apply(null, rows.map((r) => r.length).concat([1]));
  return rows.map((r) => { const out = r.slice(); while (out.length < width) out.push(''); return out; });
}

function rowIndexOf(rows, col, text) {
  const i = rows.findIndex((r) => r[col] === text);
  assert.ok(i !== -1, text + ' not found');
  return i;
}

/** Loads only BallotModel.js into a context with NO GAS globals. */
function loadBallotModelWithoutGas() {
  const sandbox = { module: { exports: {} } };
  vm.createContext(sandbox);
  const src = fs.readFileSync(path.join(__dirname, '..', 'script', 'BallotModel.js'), 'utf8');
  vm.runInContext(src, sandbox, { filename: 'BallotModel.js' });
  ['SpreadsheetApp', 'Logger', 'LockService', 'PropertiesService', 'Utilities', 'CacheService'].forEach((g) => {
    assert.equal(g in sandbox, false, g + ' leaked into the sandbox');
  });
  return sandbox.module.exports;
}

/** AC1: parse + every derivation runs on a plain array with no GAS globals at all. */
function testParseBallotNeedsNoGasGlobals() {
  const SM = loadBallotModelWithoutGas();
  const snap = SM._parseBallot_(orderedBallotRows('Pure', ['results', 'candidates', 'responses']));
  assert.deepEqual(snap.problems, []);
  assert.equal(snap.config.byKey.Title.value, 'T-Pure');
  assert.equal(snap.config.byKey.Title.row, 1);
  assert.equal(snap.config.rowsUsed, 8);
  assert.deepEqual(snap.resultsRows, [{ row: 11, values: ['Analysis run: earlier'] }, { row: 12, values: ['Winner: Alice', 3] }]);
  const c = snap.tables.candidates;
  assert.equal(c.headerRow, 14);
  assert.deepEqual(c.labels, ['Name', 'Details', 'Link Text', 'Link URL']);
  assert.deepEqual(c.col, { Name: 2, Details: 3, 'Link Text': 4, 'Link URL': 5 });
  assert.deepEqual(c.rows.map((r) => [r.row, r.key]), [[15, 'Alice'], [16, 'Bob']]);
  assert.equal(c.rowsUsed, 16);
  const r = snap.tables.responses;
  assert.deepEqual(r.col, { Date: 1, Name: 2, Weight: 3, Comment: 4, Alice: 5, Bob: 6 });
  assert.deepEqual(r.rows[1].cells, { Date: '2026-10-02', Name: 'V2', Weight: 2, Comment: 'c2', Alice: 2, Bob: 1 });
  assert.deepEqual(SM._snapCandidates_(snap).map((x) => [x.name, x.rankCol]), [['Alice', 5], ['Bob', 6]]);
  assert.deepEqual(SM._snapResponses_(snap).map((x) => [x.name, x.weight, x.ranks]), [['V1', 1, { Alice: 1, Bob: 2 }], ['V2', 2, { Alice: 2, Bob: 1 }]]);
  assert.deepEqual(SM._snapMissingLabels_(snap), { candidates: [], responses: [] });
  // Degenerate input: empty sheet / nothing at all.
  const empty = SM._parseBallot_([['']]);
  assert.equal(empty.layout.lastRow, 0);
  assert.equal(empty.tables.candidates, null);
  assert.equal(empty.resultsRows, null);
  assert.deepEqual(SM._snapResponses_(empty), []);
  assert.deepEqual(empty.problems.map((p) => p.section), ['results', 'candidates', 'responses']);
}

/** AC2: all 6 marker orders, with and without blank rows: snapshot == old readers, layout == _readBallotLayout_. */
function testSnapshotMatchesOldReadersInAllMarkerOrders() {
  SECTION_ORDERS.forEach((order) => {
    [false, true].forEach((blankRows) => {
      const { SM, ss } = loadBallotModel();
      const sheet = buildOrderedBallot(ss, 'Snap', order, { blankRows: blankRows });
      const label = order.join('/') + (blankRows ? ' +blank rows' : '');
      const snap = assertSnapshotMatchesOldReaders(SM, sheet, label);
      assert.deepEqual(snap.layout, SM._readBallotLayout_(sheet), label + ': layout');
      assert.deepEqual(snap.problems, [], label + ': problems');
      assert.deepEqual(SM._parseBallot_(rectGrid(orderedBallotRows('Snap', order, { blankRows: blankRows }))), snap, label + ': grid vs sheet');
    });
  });
}

/** AC3: a hand-sorted Candidates table; labels matched trim/case-insensitively. Ranks follow names. */
function testSnapshotHandSortedCandidatesKeepRanksByName() {
  const { SM } = loadBallotModel();
  const rows = orderedBallotRows('Sorted', ['results', 'candidates', 'responses']);
  const a = rowIndexOf(rows, 1, 'Alice');
  const b = rowIndexOf(rows, 1, 'Bob');
  const tmp = rows[a]; rows[a] = rows[b]; rows[b] = tmp;
  rows[rowIndexOf(rows, 1, 'Name')][2] = 'details'; // Candidates header, lower-case
  const respHeader = rows[rowIndexOf(rows, 0, 'Date')];
  respHeader[4] = ' alice '; // Responses header, spacing + case

  const snap = SM._parseBallot_(rows);
  assert.deepEqual(snap.problems, []);
  assert.deepEqual(SM._snapCandidates_(snap).map((c) => [c.name, c.details, c.rankCol]), [['Bob', 'b note', 6], ['Alice', 'a note', 5]]);
  assert.deepEqual(SM._snapResponses_(snap).map((r) => r.ranks), [{ Bob: 2, Alice: 1 }, { Bob: 1, Alice: 2 }]);
  assert.deepEqual(SM._snapOrphanColumns_(snap), []);
}

/** AC3: a hand-deleted Candidates row: remaining ranks still follow names; the deleted one's column is an orphan. */
function testSnapshotHandDeletedCandidateLeavesOrphanColumn() {
  const { SM } = loadBallotModel();
  const rows = orderedBallotRows('Deleted', ['results', 'candidates', 'responses']);
  rows.splice(rowIndexOf(rows, 1, 'Alice'), 1);

  const snap = SM._parseBallot_(rows);
  assert.deepEqual(SM._snapCandidates_(snap).map((c) => [c.name, c.rankCol]), [['Bob', 6]]);
  assert.deepEqual(SM._snapOrphanColumns_(snap), [{ label: 'Alice', col: 5 }]);
  const responses = SM._snapResponses_(snap);
  assert.deepEqual(responses.map((r) => r.ranks), [{ Bob: 2 }, { Bob: 1 }]);
  assert.deepEqual(responses.map((r) => r.orphanRanks), [{ Alice: 1 }, { Alice: 2 }]);
}

/** AC3: an extra Responses column matching no candidate is listed as an orphan; a candidate with no column ranks ''. */
function testSnapshotOrphanRankColumnAndCandidateWithoutColumn() {
  const { SM } = loadBallotModel();
  const rows = orderedBallotRows('Orphan', ['candidates', 'responses', 'results']);
  const h = rowIndexOf(rows, 0, 'Date');
  rows[h][6] = 'Zed';
  rows[h + 1][6] = 3;
  rows.splice(rowIndexOf(rows, 1, 'Bob') + 1, 0, ['', 'Carol']); // in Candidates only

  const snap = SM._parseBallot_(rows);
  assert.deepEqual(SM._snapOrphanColumns_(snap), [{ label: 'Zed', col: 7 }]);
  assert.deepEqual(SM._snapCandidates_(snap).map((c) => [c.name, c.rankCol]), [['Alice', 5], ['Bob', 6], ['Carol', null]]);
  const responses = SM._snapResponses_(snap);
  assert.deepEqual(responses[0].ranks, { Alice: 1, Bob: 2, Carol: '' });
  assert.deepEqual(responses.map((r) => r.orphanRanks), [{ Zed: 3 }, { Zed: '' }]);
}

/** AC3: a missing fixed header label is reported (table.missing, _snapMissingLabels_, problems); reads default. */
function testSnapshotReportsMissingFixedLabels() {
  const { SM } = loadBallotModel();
  const rows = orderedBallotRows('Missing', ['results', 'candidates', 'responses']);
  rows[rowIndexOf(rows, 1, 'Name')][4] = ''; // Candidates 'Link URL'
  rows[rowIndexOf(rows, 0, 'Date')][2] = ''; // Responses 'Weight'

  const snap = SM._parseBallot_(rows);
  assert.deepEqual(SM._snapMissingLabels_(snap), { candidates: ['Link URL'], responses: ['Weight'] });
  assert.deepEqual(snap.problems, [
    { code: 'missingLabel', section: 'candidates', labels: ['Link URL'] },
    { code: 'headerGap', section: 'responses', cols: [3] },
    { code: 'missingLabel', section: 'responses', labels: ['Weight'] },
  ]);
  const bob = SM._snapCandidates_(snap).find((c) => c.name === 'Bob');
  assert.equal(bob.linkText, 'site');
  assert.equal(bob.linkUrl, ''); // unlabelled column is not read
  assert.deepEqual(SM._snapResponses_(snap).map((r) => [r.name, r.weight, r.ranks]), [['V1', 1, { Alice: 1, Bob: 2 }], ['V2', 1, { Alice: 2, Bob: 1 }]]);

  // Missing key column: no rows can be keyed.
  rows[rowIndexOf(rows, 0, 'Date')][1] = '';
  const noName = SM._parseBallot_(rows);
  assert.deepEqual(noName.tables.responses.rows, []);
  assert.deepEqual(SM._snapMissingLabels_(noName).responses, ['Name', 'Weight']);
}

/** AC3: blank-key rows mid-section are skipped, the scan continues, and rowsUsed still counts unnamed content. */
function testSnapshotSkipsBlankKeyRowsMidSection() {
  const { SM } = loadBallotModel();
  const rows = orderedBallotRows('Blank', ['candidates', 'responses', 'results'], { blankRows: true });
  rows.splice(rowIndexOf(rows, 1, 'Bob'), 0, ['', '', 'details with no name']);
  const v2 = rowIndexOf(rows, 1, 'V2');
  rows.splice(v2, 0, ['2026-10-05', '', 1, '', 2, 1]); // unnamed Responses row with content
  rows.push(['', '', 'dangling candidate-less row']); // trailing content in Results, ignored by tables

  const snap = SM._parseBallot_(rows);
  const c = snap.tables.candidates;
  assert.deepEqual(c.rows.map((r) => r.key), ['Alice', 'Bob']);
  assert.equal(c.rowsUsed, c.rows[1].row);
  const r = snap.tables.responses;
  assert.deepEqual(r.rows.map((x) => x.key), ['V1', 'V2']);
  assert.equal(r.rows[1].row, v2 + 2); // 1-based, below the inserted unnamed row
  assert.equal(r.rowsUsed, r.rows[1].row);
  // An unnamed row with content as the LAST row of the section still counts toward rowsUsed.
  rows.splice(v2 + 2, 0, ['', '', '', 'note only']);
  assert.equal(SM._parseBallot_(rows).tables.responses.rowsUsed, v2 + 3);
}

/** Legacy config key 'Info' maps to 'Admin-Only-Notes' without writing anything. */
function testSnapshotMapsLegacyConfigKey() {
  const { SM } = loadBallotModel();
  const snap = SM._parseBallot_([['Title', 'T'], ['Info', 'old notes'], [], ['[Results]']]);
  assert.deepEqual(snap.config.byKey['Admin-Only-Notes'], { row: 2, value: 'old notes', legacyKey: 'Info' });
  assert.equal(snap.config.byKey.Info, undefined);
  assert.equal(snap.config.rowsUsed, 2);
}

/** AC4: _readBallot_ is exactly one getValues call, no writes, no lock. */
function testReadBallotIsOneGetValuesAndWritesNothing() {
  const { SM, ss, sandbox } = loadBallotModel();
  const sheet = buildOrderedBallot(ss, 'OneRead', ['responses', 'results', 'candidates'], { blankRows: true });
  const before = JSON.stringify(Array.from(sheet._cells.entries()));
  sheet._getValuesCalls = 0;
  const snap = SM._readBallot_(sheet);
  assert.equal(sheet._getValuesCalls, 1);
  assert.equal(JSON.stringify(Array.from(sheet._cells.entries())), before);
  assert.equal(sandbox.LockService._state.waitCount, 0);
  assert.deepEqual(SM._snapCandidates_(snap).map((c) => c.name), ['Alice', 'Bob']);
}

// ---------------------------------------------------------------------------
// Legacy Responses layout: labels on the marker row (rcballot-8i5.7).
// ---------------------------------------------------------------------------

/** 8i5.7 AC1/AC2: PROD Ballot-ABC-2026-7 — the repair moves Name|Weight|Comment down; the snapshot then reads all 5 votes. */
function testHeaderOnMarkerRowRepairRestoresLegacyVotes() {
  const { SM, ss, sandbox } = loadBallotModel();
  const sheet = writeGrid(ss.insertSheet('Ballot-ABC-2026-7'), prodAbc2026_7Rows());

  const before = SM._readBallot_(sheet);
  assert.equal(SM._responsesMissingName_(before), true);
  assert.equal(SM._snapResponses_(before).length, 0); // the bug: votes invisible to the snapshot
  assert.deepEqual(SM._planHeaderOnMarkerRow_(rectGrid(prodAbc2026_7Rows()), before),
    { markerRow: 66, headerRow: 67, cells: [{ col: 2, label: 'Name' }, { col: 3, label: 'Weight' }, { col: 4, label: 'Comment' }] });

  const log = spyWrites(sheet, sandbox.LockService._state);
  const result = SM._repairBallotSheet_(sheet);
  assert.equal(result.repaired, true);
  // AC2: exactly one write — the three blank header cells B-D of row 67 — under the lock.
  assert.deepEqual(log.writes, ['setValues']);
  assert.deepEqual(log.unlocked, []);
  assert.deepEqual(sheet.getRange(67, 1, 1, 5).getValues()[0], ['Date', 'Name', 'Weight', 'Comment', 'The Death of Ivan Ilyich']);
  assert.deepEqual(sheet.getRange(66, 1, 1, 4).getValues()[0], ['[Responses]', 'Name', 'Weight', 'Comment']); // marker row untouched
  const lines = logLinesWith(sandbox, 'ballot.repair.headerOnMarkerRow');
  assert.equal(lines.length, 1);
  assert.equal(lines[0].indexOf('"row":67') !== -1 && lines[0].indexOf('"cols":[2,3,4]') !== -1, true, lines[0]);

  // AC1: 5 responses, ranks by name equal to what the old (positional) readers return.
  const snap = assertSnapshotMatchesOldReaders(SM, sheet, 'PROD ABC-2026-7 after repair');
  const votes = SM._snapResponses_(snap);
  assert.deepEqual(votes.map((v) => v.name), ['Little John', 'Crazy Ivan', 'Paul Cooley', 'Glory', 'Pogo']);
  assert.deepEqual(votes[0].ranks, { 'The Death of Ivan Ilyich': 2, 'Man’s Search for Meaning': 1, 'The Foreign Correspondent': 3,
    'Admiral McRaven Commencement Speech': 4, 'Blood in the Machine': 5 });
  assert.deepEqual(snap.problems, []);

  // Idempotent: a second run finds nothing to do and takes no lock.
  const waits = sandbox.LockService._state.waitCount;
  log.writes.length = 0;
  assert.equal(SM._repairBallotSheet_(sheet).repaired, false);
  assert.deepEqual(log.writes, []);
  assert.equal(sandbox.LockService._state.waitCount, waits);
}

/** 8i5.7: a write entry (here a respondent's vote) runs the same repair first, so the vote is keyed correctly. */
function testWriteEntryRepairsHeaderOnMarkerRow() {
  const { SM, ss, sandbox } = loadBallotModel();
  const sheet = writeGrid(ss.insertSheet('Ballot-ABC-2026-7'), prodAbc2026_7Rows());
  SM.submitBallotResponse_('ABC-2026-7', 'Glory', ['Blood in the Machine'], 'revote');
  assert.equal(logLinesWith(sandbox, 'ballot.repair.headerOnMarkerRow').length, 1);
  const votes = SM._snapResponses_(SM._readBallot_(sheet));
  assert.equal(votes.length, 5);
  const glory = votes.find((v) => v.name === 'Glory');
  assert.equal(glory.row, 71); // re-vote overwrote in place
  assert.equal(glory.ranks['Blood in the Machine'], 1);
}

// ---------------------------------------------------------------------------
// Reads never write; ranks by name (rcballot-8i5.3).
// ---------------------------------------------------------------------------

/** Damaged-sheet builders: each returns a sheet some repair step would change. */
const DAMAGED_SHEETS = {
  missingResults: (ss) => { const sh = buildOrderedBallot(ss, 'D1', ['results', 'candidates', 'responses']); sh.getRange(rowOf(sh, '[Results]'), 1).setValue(''); return sh; },
  missingCandidates: (ss) => { const sh = buildOrderedBallot(ss, 'D2', ['results', 'candidates', 'responses']); sh.getRange(rowOf(sh, '[Candidates]'), 1).setValue(''); return sh; },
  missingResponses: (ss) => { const sh = buildOrderedBallot(ss, 'D3', ['responses', 'results', 'candidates']); sh.getRange(rowOf(sh, '[Responses]'), 1).setValue(''); return sh; },
  duplicateMarker: (ss) => { const sh = buildOrderedBallot(ss, 'D4', ['results', 'candidates', 'responses']); sh.getRange(sh.getLastRow() + 1, 1).setValue('[Candidates]'); return sh; },
  legacyInfoKey: (ss) => { const sh = buildOrderedBallot(ss, 'D5', ['results', 'candidates', 'responses']); sh.getRange(8, 1).setValue('Info'); return sh; },
  oldColumnACandidates: (ss) => {
    const sh = ss.insertSheet('Ballot-D6');
    writeGrid(sh, [['Title', 'T'], [], ['[Results]'], ['[Candidates]'], ['Name', 'Details'], ['book 1', 'd1'], ['[Responses]'], ['Date', 'Name', 'Weight', 'Comment', 'book 1'], ['x', 'V', 1, '', 1]]);
    return sh;
  },
  candidatesHeaderMissingLabels: (ss) => { const sh = buildOrderedBallot(ss, 'D7', ['results', 'candidates', 'responses']); sh.getRange(rowOf(sh, '[Candidates]') + 1, 4, 1, 2).clearContent(); return sh; },
  preCandidatesTable: (ss) => {
    const sh = ss.insertSheet('Ballot-D8');
    writeGrid(sh, [['Title', 'T'], [], ['[Results]'], ['[Candidates]'], ['', 'Name', 'Details', 'Link Text', 'Link URL'], ['[Responses]'], ['Date', 'Name', 'Weight', 'Comment', 'Alice', 'Bob'], ['x', 'V', 1, '', 1, 2]]);
    return sh;
  },
  candidateWithoutRankColumn: (ss) => { const sh = buildOrderedBallot(ss, 'D9', ['candidates', 'responses', 'results']); sh.insertRowsAfter(rowOf(sh, '[Responses]') - 1, 1); sh.getRange(rowOf(sh, '[Responses]') - 1, 2).setValue('Carol'); return sh; },
  headerOnMarkerRow: (ss) => writeGrid(ss.insertSheet('Ballot-ABC-2026-7'), prodAbc2026_7Rows()),
  headerlessCandidates: (ss) => writeGrid(ss.insertSheet('Ballot-ABC-2026-08'), sitAbc2026_08Rows()),
};

/** 8i5.3 AC1: no read entry point writes or locks, on healthy sheets (all orders) and on every damaged kind. */
function testReadEntryPointsNeverWrite() {
  const sheets = [];
  SECTION_ORDERS.forEach((order) => sheets.push(['healthy ' + order.join('/'), (ss) => buildOrderedBallot(ss, 'H' + order.map((k) => k[0]).join(''), order, { blankRows: true })]));
  Object.keys(DAMAGED_SHEETS).forEach((k) => sheets.push([k, DAMAGED_SHEETS[k]]));
  sheets.forEach(([label, make]) => {
    const { SM, ss, sandbox } = loadBallotModel();
    const sheet = make(ss);
    const id = sheet.getName().substring('Ballot-'.length);
    // duplicateMarker is reported, never auto-repaired (which copy is real is a human call).
    const repairNeeded = label.indexOf('healthy') === -1 && label !== 'duplicateMarker';
    const before = JSON.stringify(Array.from(sheet._cells.entries()));
    sheet._writeCalls = 0;
    SM.readBallotConfig_(sheet);
    SM.readBallotCandidates_(sheet);
    SM.readBallotCandidateDetails_(sheet);
    SM.readBallotResponseRows_(sheet);
    SM.countUniqueBallotRespondents_(sheet);
    SM.findRespondentsWithNewCandidates_(sheet);
    SM.getBallotForRespondent_(id, '');
    SM.getBallotForRespondent_(id, 'V1');
    SM.getBallotForRespondent_(id, 'Pogo');
    SM.refreshBallotCache_(id, sheet);
    SM.getCachedBallotData_(id, sheet);
    assert.equal(sheet._writeCalls, 0, label + ': reads wrote');
    assert.equal(JSON.stringify(Array.from(sheet._cells.entries())), before, label + ': cells changed');
    assert.equal(sandbox.LockService._state.waitCount, 0, label + ': lock taken');
    if (repairNeeded) {
      // ...while the repair path does have work on every damaged kind (so the fixtures are real).
      assert.equal(SM._repairBallotSheet_(sheet).repaired, true, label + ': nothing to repair');
    }
  });
}

/** SIT Ballot-ABC-2026-08: headerless Candidates. Reads show no candidates (7 orphans); the repair restores all 8 and votes by name. */
function testHeaderlessCandidatesRepairedOnRepairPath() {
  const { SM, ss, sandbox } = loadBallotModel();
  const sheet = writeGrid(ss.insertSheet('Ballot-ABC-2026-08'), sitAbc2026_08Rows());
  const ranked = ['Endurance', 'The Righteous Mind', 'Hansel and Gretel + other original fairy tales', 'Joseph',
    'The Feather Thief', 'The Good Life', 'Who Moved My Cheese?'];

  assert.deepEqual(SM.readBallotCandidates_(sheet), []);
  assert.deepEqual(SM._snapOrphanColumns_(SM._readBallot_(sheet)).map((o) => o.label), ranked);

  SM._repairBallotSheet_(sheet);
  assert.deepEqual(logLinesWith(sandbox, 'ballot.repair.candidatesHeader').length, 1);
  assert.deepEqual(logLinesWith(sandbox, 'ballot.repair.rankColumns').length, 1);
  assert.equal(sheet.getRange(77, 1).getValue(), '[Candidates]');
  assert.deepEqual(sheet.getRange(78, 1, 1, 5).getValues()[0], ['', 'Name', 'Details', 'Link Text', 'Link URL']); // inserted
  assert.deepEqual(SM.readBallotCandidates_(sheet), ['The Boys in the Boat'].concat(ranked));
  assert.equal(SM.readBallotCandidateDetails_(sheet)[1].details, 'Endurance - details');
  const snap = SM._readBallot_(sheet);
  assert.deepEqual(SM._snapOrphanColumns_(snap), []);
  assert.deepEqual(snap.problems, []);
  // Votes by name, exactly as the live sheet's positional header said.
  const lj = SM.readBallotResponseRows_(sheet).find((r) => r.name === 'Little John');
  assert.deepEqual(lj.ranks, ['', 1, 2, 6, 7, 5, 4, 3]); // The Boys in the Boat: new column, unranked
  assert.equal(sheet.getRange(88, 12).getValue(), 'The Boys in the Boat'); // header row 87 -> 88 after the insert
  assert.equal(SM.countUniqueBallotRespondents_(sheet), 5);
  assert.equal(SM._repairBallotSheet_(sheet).repaired, false); // idempotent
}

/** AC5 (respondent page) + writers: a hand-sorted Candidates table never moves a rank to another candidate. */
function testHandSortedCandidatesKeepRanksByNameThroughReadsAndWrites() {
  const { SM, ss } = loadBallotModel();
  const sheet = buildOrderedBallot(ss, 'Sorted', ['results', 'candidates', 'responses']);
  // Swap Alice and Bob rows by hand (Responses columns stay Alice=E, Bob=F).
  const a = rowOf(sheet, '[Candidates]') + 2;
  const rowA = sheet.getRange(a, 1, 1, 5).getValues()[0];
  const rowB = sheet.getRange(a + 1, 1, 1, 5).getValues()[0];
  sheet.getRange(a, 1, 1, 5).setValues([rowB]);
  sheet.getRange(a + 1, 1, 1, 5).setValues([rowA]);

  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Bob', 'Alice']);
  // V1 ranked Alice=1, Bob=2.
  assert.deepEqual(SM.readBallotResponseRows_(sheet)[0].ranks, [2, 1]);
  assert.deepEqual(SM.getBallotForRespondent_('Sorted', 'V1').candidates, ['Alice', 'Bob']);

  // A new vote lands under each candidate's own column.
  SM.submitBallotResponse_('Sorted', 'V3', ['Bob', 'Alice'], '');
  const header = SM._readBallotLayout_(sheet).sections.responses.headerRow;
  const v3 = SM.readBallotResponseRows_(sheet).find((r) => r.name === 'V3');
  assert.deepEqual(sheet.getRange(v3.row, 5, 1, 2).getValues()[0], [2, 1]); // E=Alice 2, F=Bob 1
  assert.deepEqual(sheet.getRange(header, 5, 1, 2).getValues()[0], ['Alice', 'Bob']);

  // Renaming the first table row (Bob) relabels Bob's column, not Alice's.
  SM.saveBallotCandidates_(sheet, [{ name: 'Robert', details: '' }, { name: 'Alice', details: 'a note' }]);
  assert.deepEqual(sheet.getRange(header, 5, 1, 2).getValues()[0], ['Alice', 'Robert']);
  assert.deepEqual(SM.readBallotResponseRows_(sheet)[0].ranks, [2, 1]); // V1: Robert(ex-Bob)=2, Alice=1
}

/** AC5: a hand-deleted Candidates row: remaining ranks follow names; a new candidate never adopts the orphan column. */
function testHandDeletedCandidateOrphanNeverAdopted() {
  const { SM, ss } = loadBallotModel();
  const sheet = buildOrderedBallot(ss, 'Deleted', ['results', 'candidates', 'responses']);
  sheet.deleteRows(rowOf(sheet, '[Candidates]') + 2, 1); // Alice's row; her column E stays

  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Bob']);
  assert.deepEqual(SM.readBallotResponseRows_(sheet).map((r) => r.ranks), [[2], [1]]);
  assert.deepEqual(SM.findRespondentsWithNewCandidates_(sheet), []);

  SM.addBallotCandidate_(sheet, 'Carol', '');
  const header = SM._readBallotLayout_(sheet).sections.responses.headerRow;
  assert.deepEqual(sheet.getRange(header, 5, 1, 3).getValues()[0], ['Alice', 'Bob', 'Carol']); // appended after the orphan
  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Bob', 'Carol']);
  assert.deepEqual(SM.readBallotResponseRows_(sheet).map((r) => r.ranks), [[2, ''], [1, '']]);

  // Re-adding the deleted name reuses its column, so its recorded ranks come back.
  SM.addBallotCandidate_(sheet, 'Alice', '');
  assert.deepEqual(SM.readBallotResponseRows_(sheet).map((r) => r.ranks), [[2, '', 1], [1, '', 2]]);
  assert.deepEqual(SM._snapOrphanColumns_(SM._readBallot_(sheet)), []);
}

// ---------------------------------------------------------------------------
// Write applier (rcballot-8i5.4): _applyBallotOps_ and the uniqueness helpers.
// ---------------------------------------------------------------------------

/** Runs fn and returns the GAS write / getValues calls it made on sheet. */
function countCalls(sheet, fn) {
  sheet._writeCalls = 0;
  sheet._getValuesCalls = 0;
  const result = fn();
  return { writes: sheet._writeCalls, reads: sheet._getValuesCalls, result: result };
}

/** update merges untouched cells (in-between candidate/orphan columns) from the snapshot; append inserts before the next marker. */
function testApplyBallotOpsUpdateMergesAndAppendMakesRoom() {
  const { SM, ss } = loadBallotModel();
  // Results 10-12, Responses 13 (header 14, V1 15, V2 16), Candidates 17 (header 18, Alice 19, Bob 20).
  const sheet = buildOrderedBallot(ss, 'Ops', ['results', 'responses', 'candidates']);
  sheet.getRange(14, 7).setValue('Carol'); // orphan rank column
  sheet.getRange(15, 7).setValue(4);
  sheet.getRange(15, 8).setValue('note'); // unlabelled cell
  const candidatesBefore = sectionSnapshot(sheet, 'candidates');

  let snap = SM._readBallot_(sheet);
  let c = countCalls(sheet, () => SM._applyBallotOps_(sheet, snap, [{ section: 'responses', update: [{ row: 15, cells: { alice: 3, Carol: 9 } }] }]));
  assert.deepEqual([c.writes, c.reads], [1, 0], 'update: one setValues, no read');
  assert.deepEqual(sheet.getRange(15, 5, 1, 4).getValues()[0], [3, 2, 9, 'note']); // Bob (col 6) kept
  assert.deepEqual(SM._snapResponses_(c.result).map((r) => r.ranks), [{ Alice: 3, Bob: 2 }, { Alice: 2, Bob: 1 }]);

  snap = c.result;
  c = countCalls(sheet, () => SM._applyBallotOps_(sheet, snap, [{ section: 'responses', append: [{ Date: 'd3', Name: 'V3', Weight: 1, Comment: '', Alice: 1, Bob: 2 }] }]));
  assert.deepEqual([c.writes, c.reads], [2, 1], 'append into a full section: insert + setValues, one re-read');
  assert.deepEqual(sheet.getRange(17, 1, 1, 6).getValues()[0], ['d3', 'V3', 1, '', 1, 2]);
  assert.equal(sheet.getRange(18, 1).getValue(), '[Candidates]');
  assert.deepEqual(sectionSnapshot(sheet, 'candidates'), candidatesBefore);
  assert.deepEqual(SM._snapResponses_(c.result).map((r) => r.name), ['V1', 'V2', 'V3']);
  assert.deepEqual(SM._readBallot_(sheet).tables, c.result.tables, 'returned snapshot matches the sheet');
}

/** ensureColumns in the same set as a Candidates append: header cell placed after the row insert; existing labels (any case) untouched. */
function testApplyBallotOpsEnsureColumnsWithAppend() {
  const { SM, ss } = loadBallotModel();
  const sheet = buildOrderedBallot(ss, 'Ens', ['results', 'candidates', 'responses']);
  const snap = SM._readBallot_(sheet);
  const c = countCalls(sheet, () => SM._applyBallotOps_(sheet, snap, [
    { section: 'candidates', append: [{ Name: 'Dan', Details: 'd' }] },
    { section: 'responses', ensureColumns: ['alice', 'Dan'] },
  ]));
  assert.equal(c.writes, 3, 'insert + candidate row + header cell');
  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Alice', 'Bob', 'Dan']);
  const header = SM._readBallotLayout_(sheet).sections.responses.headerRow;
  assert.deepEqual(sheet.getRange(header, 5, 1, 4).getValues()[0], ['Alice', 'Bob', 'Dan', '']);
  assert.deepEqual(SM._snapOrphanColumns_(c.result), []);
}

/** renameColumn finds the header cell by OLD label; ranks stay with the candidate. Fixed / duplicate targets refuse. */
function testApplyBallotOpsRenameColumn() {
  const { SM, ss, sandbox } = loadBallotModel();
  const sheet = buildOrderedBallot(ss, 'Ren', ['results', 'candidates', 'responses']); // Alice row 15, Responses header 18
  const snap = SM._readBallot_(sheet);
  SM._applyBallotOps_(sheet, snap, [
    { section: 'candidates', update: [{ row: 15, cells: { Name: 'Alicia' } }] },
    { section: 'responses', renameColumn: { from: 'ALICE', to: 'Alicia' } },
  ]);
  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Alicia', 'Bob']);
  assert.deepEqual(SM.readBallotResponseRows_(sheet).map((r) => r.ranks), [[1, 2], [2, 1]]);

  [[{ from: 'Alicia', to: 'bob' }, 'duplicateColumn'], [{ from: 'Weight', to: 'W' }, 'fixedLabel']].forEach(([rename, reason]) => {
    const cur = SM._readBallot_(sheet);
    const c = countCalls(sheet, () => assert.throws(() => SM._applyBallotOps_(sheet, cur, [
      { section: 'candidates', update: [{ row: 15, cells: { Details: 'changed' } }] },
      { section: 'responses', renameColumn: rename },
    ]), /^Error: Ballot sheet layout is damaged — /));
    assert.equal(c.writes, 0, reason + ': nothing written');
    assert.equal(logLinesWith(sandbox, '"reason":"' + reason + '"').length, 1, reason);
  });
}

/** deleteRows, then an update addressed in the ORIGINAL snapshot's rows: mapped across the delete. */
function testApplyBallotOpsDeleteRowsMapsLaterRows() {
  const { SM, ss } = loadBallotModel();
  const sheet = buildOrderedBallot(ss, 'Del', ['results', 'candidates', 'responses']); // Alice 15, Bob 16; V2 row 20
  const snap = SM._readBallot_(sheet);
  const c = countCalls(sheet, () => SM._applyBallotOps_(sheet, snap, [
    { section: 'candidates', deleteRows: [15] },
    { section: 'responses', update: [{ row: 20, cells: { Comment: 'moved' } }] },
  ]));
  assert.deepEqual([c.writes, c.reads], [2, 1]);
  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Bob']);
  assert.deepEqual(SM.readBallotResponseRows_(sheet).map((r) => r.name + ':' + r.comment), ['V1:', 'V2:moved']);
  assert.deepEqual(SM._snapOrphanColumns_(c.result).map((o) => o.label), ['Alice']);
  assert.throws(() => SM._applyBallotOps_(sheet, c.result, [{ section: 'candidates', deleteRows: [13] }]), /outside ballot section candidates/);
}

/** A table op on a header missing a fixed label refuses the whole set; unknown labels are a plain throw. */
function testApplyBallotOpsRefusals() {
  const { SM, ss, sandbox } = loadBallotModel();
  const sheet = buildOrderedBallot(ss, 'Bad', ['results', 'candidates', 'responses']);
  sheet.getRange(18, 3).setValue(''); // Responses header loses 'Weight'
  const snap = SM._readBallot_(sheet);
  let c = countCalls(sheet, () => assert.throws(() => SM._applyBallotOps_(sheet, snap, [
    { section: 'candidates', append: [{ Name: 'Dan' }] },
    { section: 'responses', ensureColumns: ['Dan'] },
  ]), /^Error: Ballot sheet layout is damaged — the \[Responses\] header is missing: Weight\.$/));
  assert.equal(c.writes, 0);
  const lines = logLinesWith(sandbox, 'ballot.write.refused');
  assert.equal(lines.length, 1);
  assert.ok(lines[0].indexOf('"section":"responses","reason":"missingLabel","labels":["Weight"]') !== -1, lines[0]);

  c = countCalls(sheet, () => assert.throws(() => SM._applyBallotOps_(sheet, snap, [
    { section: 'candidates', append: [{ Name: 'Dan', Colour: 'red' }] },
  ]), /^Error: Unknown column "Colour" in ballot section candidates\.$/));
  assert.equal(c.writes, 0);
  assert.throws(() => SM._applyBallotOps_(sheet, snap, [{ section: 'results', append: [] }]), /Unsupported ballot op/);
}

/** AC2 helpers: trim/case-insensitive; fixed Responses labels and marker words are taken; the renamed row is exempt. */
function testUniqueCandidateNameHelpers() {
  const { SM, ss } = loadBallotModel();
  const snap = SM._readBallot_(buildOrderedBallot(ss, 'Uniq', ['results', 'candidates', 'responses'])); // Alice row 15
  const msg = /^Error: A candidate with that name already exists\.$/;
  [' alice ', 'BOB', 'weight', 'Name', 'date', 'Comment', '[Results]', 'items'].forEach((n) => assert.throws(() => SM._assertUniqueCandidateName_(snap, n), msg, n));
  SM._assertUniqueCandidateName_(snap, 'Carol');
  SM._assertUniqueCandidateName_(snap, 'ALICE', 15); // renaming Alice's own row
  assert.throws(() => SM._assertUniqueCandidateName_(snap, 'Bob', 15), msg);
  assert.throws(() => SM._assertUniqueCandidateNames_(['A', ' a']), msg);
  assert.throws(() => SM._assertUniqueCandidateNames_(['A', 'Weight']), msg);
  SM._assertUniqueCandidateNames_(['A', 'B']);
}

/** AC7 (results): a rewrite is one setValues plus at most one insert/delete — no clearContent; healthy sheet: one repair read. */
function testWriteBallotResultsCallCounts() {
  const { SM, ss } = loadBallotModel();
  const sheet = buildOrderedBallot(ss, 'ResCount', ['results', 'candidates', 'responses']); // 2 output rows
  const others = ['candidates', 'responses'].map((k) => sectionSnapshot(sheet, k));
  [[[['a'], ['b']], 1, 1], [[['a'], ['b'], ['c'], ['d']], 2, 2], [[['only']], 2, 2], [[], 1, 2]].forEach(([rows, writes, reads]) => {
    const c = countCalls(sheet, () => SM.writeBallotResults_(sheet, rows));
    assert.deepEqual([c.writes, c.reads], [writes, reads], rows.length + ' rows');
    assert.deepEqual(sectionSnapshot(sheet, 'results'), [['[Results]']].concat(rows.map((r) => [''].concat(r))));
    assert.deepEqual(['candidates', 'responses'].map((k) => sectionSnapshot(sheet, k)), others);
  });
  // Results as the bottom section grows past the last row without an insert.
  const t = loadBallotModel();
  const bottom = buildOrderedBallot(t.ss, 'ResBottom', ['candidates', 'responses', 'results']);
  const c = countCalls(bottom, () => t.SM.writeBallotResults_(bottom, [['1'], ['2'], ['3']]));
  assert.deepEqual([c.writes, c.reads], [1, 1]);
  assert.deepEqual(sectionSnapshot(bottom, 'results'), [['[Results]'], ['', '1'], ['', '2'], ['', '3']]);
}

/** AC5: a new config key is written when a marker sits on row 1; AC7: existing keys cost one setValues per contiguous block. */
function testWriteBallotConfigMarkerOnRowOneAndCallCounts() {
  const { SM, ss } = loadBallotModel();
  const sheet = writeGrid(ss.insertSheet('Ballot-Row1'), [
    ['[Results]'], ['[Candidates]'], ['', 'Name', 'Details', 'Link Text', 'Link URL'], ['', 'Alice'],
    ['[Responses]'], ['Date', 'Name', 'Weight', 'Comment', 'Alice'], ['d', 'V1', 1, '', 1],
  ]);
  let c = countCalls(sheet, () => SM.writeBallotConfig_(sheet, { Title: 'T', 'Accept-New': 'TRUE' }));
  assert.deepEqual([c.writes, c.reads], [2, 2], 'insert + one setValues; repair read + re-read');
  assert.deepEqual(sheet.getRange(1, 1, 3, 2).getValues(), [['Title', 'T'], ['Accept-New', 'TRUE'], ['[Results]', '']]);
  assert.deepEqual(SM.readBallotConfig_(sheet), { Title: 'T', 'Accept-New': 'TRUE' });
  assert.deepEqual(SM.readBallotResponseRows_(sheet).map((r) => r.name + ':' + r.ranks.join(',')), ['V1:1']);

  const t = loadBallotModel();
  const healthy = buildOrderedBallot(t.ss, 'CfgCount', ['results', 'candidates', 'responses']);
  c = countCalls(healthy, () => t.SM.writeBallotConfig_(healthy, { Title: 'x', Description: 'y' }));
  assert.deepEqual([c.writes, c.reads], [1, 1], 'adjacent keys: one setValues');
  c = countCalls(healthy, () => t.SM.writeBallotConfig_(healthy, { Title: 'x2', Footer: 'f2' }));
  assert.deepEqual([c.writes, c.reads], [2, 1], 'rows 1 and 4: two blocks');
  assert.deepEqual(healthy.getRange(1, 1, 4, 2).getValues(), [['Title', 'x2'], ['Description', 'y'], ['Instructions', 'I'], ['Footer', 'f2']]);
}

/** AC7 (submit): a new vote is 1 setValues at the bottom, 2 (insert + setValues) when a row must be made; a re-vote is 1. */
function testSubmitBallotResponseCallCounts() {
  const { SM, ss } = loadBallotModel();
  buildOrderedBallot(ss, 'Sub1', ['results', 'candidates', 'responses']);
  const sheet = ss.getSheetByName('Ballot-Sub1');
  let c = countCalls(sheet, () => SM.submitBallotResponse_('Sub1', 'V3', ['Bob', 'Alice'], 'hi'));
  assert.deepEqual([c.writes, c.reads], [1, 1], 'new vote, Responses at the bottom');
  c = countCalls(sheet, () => SM.submitBallotResponse_('Sub1', 'v3', ['Alice', 'Bob'], 'again'));
  assert.deepEqual([c.writes, c.reads], [1, 1], 're-vote');
  assert.deepEqual(SM.readBallotResponseRows_(sheet).map((r) => r.name + ':' + r.ranks.join(',')), ['V1:1,2', 'V2:2,1', 'v3:1,2']);

  const t = loadBallotModel();
  const mid = buildOrderedBallot(t.ss, 'Sub2', ['results', 'responses', 'candidates']);
  const others = ['results', 'candidates'].map((k) => sectionSnapshot(mid, k));
  c = countCalls(mid, () => t.SM.submitBallotResponse_('Sub2', 'V3', ['Bob', 'Alice'], ''));
  assert.deepEqual([c.writes, c.reads], [2, 2], 'new vote needing an inserted row');
  assert.deepEqual(['results', 'candidates'].map((k) => sectionSnapshot(mid, k)), others);
  assert.deepEqual(t.SM.readBallotResponseRows_(mid).map((r) => r.name), ['V1', 'V2', 'V3']);
}

/** AC2/AC4/AC7 (add candidate): one-row append + header cell in one op set; duplicates and a damaged header write nothing. */
function testAddBallotCandidateOneRowAppendDuplicatesAndAtomicity() {
  const { SM, ss } = loadBallotModel();
  const sheet = buildOrderedBallot(ss, 'Add1', ['results', 'candidates', 'responses']);
  const results = sectionSnapshot(sheet, 'results');
  let c = countCalls(sheet, () => SM.addBallotCandidateForId_('Add1', 'Carol', 'about carol'));
  assert.deepEqual([c.writes, c.reads], [3, 2], 'insert + candidate row + header cell');
  assert.deepEqual(c.result, { candidate: 'Carol' });
  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Alice', 'Bob', 'Carol']);
  assert.equal(SM.readBallotCandidateDetails_(sheet)[2].details, 'about carol');
  assert.deepEqual(SM.readBallotResponseRows_(sheet).map((r) => r.ranks), [[1, 2, ''], [2, 1, '']]);
  assert.deepEqual(sectionSnapshot(sheet, 'results'), results);

  // Duplicates (case/space-insensitive) are refused on both add paths, with nothing written.
  c = countCalls(sheet, () => {
    assert.throws(() => SM.addBallotCandidateForId_('Add1', ' carol ', ''), /^Error: A candidate with that name already exists\.$/);
    assert.throws(() => SM.addBallotCandidateForAdmin_('Add1', 'ALICE', '', '', ''), /^Error: A candidate with that name already exists\.$/);
    assert.throws(() => SM.addBallotCandidateForId_('Add1', 'Weight', ''), /^Error: A candidate with that name already exists\.$/);
  });
  assert.equal(c.writes, 0);
  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Alice', 'Bob', 'Carol']);

  // Admin path: appends and highlights.
  SM.addBallotCandidateForAdmin_('Add1', 'Dan', 'd', 'lt', 'http://x');
  assert.deepEqual(SM.readBallotCandidateDetails_(sheet)[3], { name: 'Dan', details: 'd', linkText: 'lt', linkUrl: 'http://x' });

  // AC4: a damaged Responses header refuses the whole op set: no table row, no header column.
  const t = loadBallotModel();
  const bad = buildOrderedBallot(t.ss, 'Add2', ['results', 'candidates', 'responses']);
  bad.getRange(18, 3).setValue('');
  const before = [sectionSnapshot(bad, 'candidates'), sectionSnapshot(bad, 'responses')];
  c = countCalls(bad, () => assert.throws(() => t.SM.addBallotCandidateForId_('Add2', 'Eve', ''),
    /^Error: Ballot sheet layout is damaged — the \[Responses\] header is missing: Weight\.$/));
  assert.equal(c.writes, 0);
  assert.deepEqual([sectionSnapshot(bad, 'candidates'), sectionSnapshot(bad, 'responses')], before);

  // Accept-New off: refused with its own message.
  const u = loadBallotModel();
  const closed = buildOrderedBallot(u.ss, 'Add3', ['results', 'candidates', 'responses']);
  closed.getRange(6, 2).setValue('FALSE');
  assert.throws(() => u.SM.addBallotCandidateForId_('Add3', 'Eve', ''), /^Error: This ballot is not accepting new items\.$/);
}

function run() {
  testAddBallotCandidateOneRowAppendDuplicatesAndAtomicity();
  testSubmitBallotResponseCallCounts();
  testWriteBallotConfigMarkerOnRowOneAndCallCounts();
  testWriteBallotResultsCallCounts();
  testApplyBallotOpsUpdateMergesAndAppendMakesRoom();
  testApplyBallotOpsEnsureColumnsWithAppend();
  testApplyBallotOpsRenameColumn();
  testApplyBallotOpsDeleteRowsMapsLaterRows();
  testApplyBallotOpsRefusals();
  testUniqueCandidateNameHelpers();
  testCreateNewBallotCreatesSkeleton();
  testListBallotIdsAndFindBallotSheet();
  testWriteThenReadConfigRoundTrips();
  testWriteBallotConfigSelfHealsMissingAdminOnlyNotesRowOnOldSheet();
  testWriteBallotConfigInsertsRowWhenNoBlankSpacerExists();
  testLegacyInfoKeyMigratesToAdminOnlyNotes();
  testCandidatesAndSubmitResponseRoundTrip();
  testCountUniqueBallotRespondents();
  testFindRespondentsWithNewCandidates();
  testGetBallotForRespondentIncludesItemDetails();
  testAddBallotCandidateTracksDetailsPositionAligned();
  testSaveBallotCandidatesRenamesAndUpdatesDetailsWithoutDisturbingResponses();
  testSaveBallotCandidatesRejectsCountMismatch();
  testSaveBallotCandidatesRejectsEmptyName();
  testAddBallotCandidateForAdminAndSaveBallotCandidatesForId();
  testAddBallotCandidateForIdAcceptsRespondentDetails();
  testEnsureCandidatesSectionSelfHealsOldSheetAndBackfillsOnAdd();
  testMigratesExistingOldColumnCandidatesSectionOnRepair();
  testPlainTextMarkersAreNotMatchedOrRewritten();
  testWriteBallotResultsUsesColumnBAndClearsColumnA();
  testWriteBallotResultsClearsPreexistingColumnAText();
  testMarkerNamedCandidatesAreRejected();
  testPrePopulatedCandidateNotInResponsesIsReconciledIn();
  testReadBallotCandidateDetailsIgnoresBlankRow();
  testReconcileCompactsGapCausedByWiderCandidatesTable();
  testMissingResponsesSectionSelfHealsAndReturnsCandidates();
  testMissingResponsesSectionWithColumnAContentIsNotRepaired();
  testCandidateInColumnAIsLoggedAndIgnored();
  testCacheRefreshLogsDiscoverySummary();
  testWithBallotLockIsReentrantAndReleasesOnThrow();
  testModelEntryPointsWriteUnderLock();
  testSelfRepairWritesRunUnderLock();
  testNewBallotHasNoFrozenRows();
  testHighlightSectionMarkersUnfreezesStaleFourteenRowFreeze();
  testDuplicateIdThrows();
  testInvalidIdThrows();
  testAllMarkerOrdersReadTheSameData();
  testDuplicateMarkerUsesFirstOccurrenceAndIsLogged();
  testResponsesSectionEndsAtNextMarker();
  testSectionReadersReadTheSheetOnce();
  testNewVoteLandsInsideResponsesWhenNotLast();
  testNewVoteFillsBlankRowBeforeNextMarker();
  testNewVoteNeverOverwritesUnnamedRowWithContent();
  testAddBallotCandidateAfterRepairWritesResponsesHeader();
  testNoGetLastRowPlusOneWrites();
  testEachMissingMarkerIsRestoredAndLogged();
  testMissingResponsesMarkerWithOrphanedRowsIsNotRepaired();
  testWriteBallotResultsRefusesOnDamagedLayout();
  testWriteBallotConfigInsertsAboveFirstMarkerInAnyOrder();
  testSectionWritesLeaveOtherSectionsUnchangedInAllOrders();
  testParseBallotNeedsNoGasGlobals();
  testSnapshotMatchesOldReadersInAllMarkerOrders();
  testSnapshotHandSortedCandidatesKeepRanksByName();
  testSnapshotHandDeletedCandidateLeavesOrphanColumn();
  testSnapshotOrphanRankColumnAndCandidateWithoutColumn();
  testSnapshotReportsMissingFixedLabels();
  testSnapshotSkipsBlankKeyRowsMidSection();
  testSnapshotMapsLegacyConfigKey();
  testReadBallotIsOneGetValuesAndWritesNothing();
  testHeaderOnMarkerRowRepairRestoresLegacyVotes();
  testWriteEntryRepairsHeaderOnMarkerRow();
  testReadEntryPointsNeverWrite();
  testHeaderlessCandidatesRepairedOnRepairPath();
  testHandSortedCandidatesKeepRanksByNameThroughReadsAndWrites();
  testHandDeletedCandidateOrphanNeverAdopted();
  console.log('test_ballot_model: all tests passed');
}

run();
