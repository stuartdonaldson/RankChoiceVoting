'use strict';

const assert = require('assert');
const vm = require('vm');
const fs = require('fs');
const path = require('path');

const { createFakeSpreadsheet, createFakeGasGlobals } = require('./fakeGas');

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

  // Admin-Only-Notes reused the existing blank spacer row (6) — Results/Responses/
  // header stay put.
  assert.equal(sheet.getRange(6, 1).getValue(), 'Admin-Only-Notes');
  assert.equal(sheet.getRange(7, 1).getValue(), '[Results]');
  assert.equal(sheet.getRange(9, 1).getValue(), '[Responses]');
  assert.deepEqual(sheet.getRange(10, 1, 1, 4).getValues()[0], ['Date', 'Name', 'Weight', 'Comment']);
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

  // A new row was inserted at row 5 (just above Results), shifting Results/Responses/
  // header down by one row each.
  assert.equal(sheet.getRange(5, 1).getValue(), 'Admin-Only-Notes');
  assert.equal(sheet.getRange(6, 1).getValue(), '[Results]');
  assert.equal(sheet.getRange(8, 1).getValue(), '[Responses]');
  assert.deepEqual(sheet.getRange(9, 1, 1, 4).getValues()[0], ['Date', 'Name', 'Weight', 'Comment']);
}

/**
 * Confirms the old "Info" config key is recognized as a legacy alias for
 * "Admin-Only-Notes" and migrated to the new key name in place the first time
 * readBallotConfig_ scans it — covers a sheet that hasn't yet been touched since
 * the Info->Admin-Only-Notes rename.
 */
function testLegacyInfoKeyMigratesToAdminOnlyNotes() {
  const { SM, ss } = loadBallotModel();
  const sheet = ss.insertSheet('Ballot-LegacyInfo');
  sheet.getRange(1, 1, 6, 2).setValues([
    ['Title', 'T'], ['Description', 'D'], ['Footer', 'F'],
    ['Contact', 'C'], ['Accept-New', 'TRUE'], ['Info', 'old notes'],
  ]);
  sheet.getRange(8, 1).setValue('[Results]');

  const config = SM.readBallotConfig_(sheet);
  assert.equal(config['Admin-Only-Notes'], 'old notes');
  assert.equal(config.Info, undefined);
  assert.equal(sheet.getRange(6, 1).getValue(), 'Admin-Only-Notes'); // migrated in place
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
  // ranks are positional per candidate column: Alice=3, Bob=1, Carol=2
  assert.deepEqual(rows[0].ranks, [3, 1, 2]);

  // Re-submitting for the same voter (case-insensitive) overwrites the row rather than
  // appending a new one, and preserves their existing weight.
  SM.submitBallotResponse_('Candidates', 'voter one', ['Alice', 'Bob', 'Carol'], 'updated');
  const rows2 = SM.readBallotResponseRows_(sheet);
  assert.equal(rows2.length, 1);
  assert.equal(rows2[0].comment, 'updated');
  assert.deepEqual(rows2[0].ranks, [1, 2, 3]);
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
}

/**
 * findRespondentsWithNewCandidates_ flags a respondent whose last submission
 * predates a candidate added afterward — detected via the blank rank cell left
 * under that candidate's newly-appended Responses column (_reconcileCandidatesWithResponses_
 * never backfills old rows). A respondent who re-submits after the new candidate
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

  // Rename is purely positional (column index unchanged) — the voter's ranks,
  // recorded against column position, must survive untouched.
  const rows = SM.readBallotResponseRows_(sheet);
  assert.equal(rows.length, 1);
  assert.deepEqual(rows[0].ranks, [2, 1]); // Alice(now Alicia)=2, Bob(now Robert)=1
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
 * markers only, candidates already typed directly into the Responses header) to
 * confirm _ensureCandidatesSection_ inserts an empty Candidates section on first
 * touch, and that adding a further candidate correctly backfills the pre-existing
 * ones instead of misaligning the new candidate's details against the wrong column.
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

  assert.deepEqual(SM.readBallotCandidateDetails_(sheet), []); // no Candidates rows yet, but no throw
  // _ensureCandidatesSection_ inserts the Candidates marker+header directly above the
  // (now shifted) Responses marker: old Responses(10) -> 12, Candidates marker lands
  // at 10, written decorated.
  assert.equal(sheet.getRange(10, 1).getValue(), '[Candidates]');
  assert.equal(sheet.getRange(11, 1).getValue(), ''); // col A reserved, blank
  assert.deepEqual(sheet.getRange(11, 2, 1, 2).getValues()[0], ['Name', 'Details']);
  assert.equal(sheet.getRange(12, 1).getValue(), '[Responses]');

  SM.addBallotCandidate_(sheet, 'Carol', 'newest candidate');

  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Alice', 'Bob', 'Carol']);
  assert.deepEqual(SM.readBallotCandidateDetails_(sheet), [
    { name: 'Alice', details: '', linkText: '', linkUrl: '' },
    { name: 'Bob', details: '', linkText: '', linkUrl: '' },
    { name: 'Carol', details: 'newest candidate', linkText: '', linkUrl: '' },
  ]);
}

/**
 * Reproduces the live bug found on Ballot-ABC-2026-7: a Candidates section that DOES
 * exist (created before the column-A-reservation fix shipped), with Name/Details
 * still sitting in columns A/B instead of B/C. Before _migrateCandidatesColumnsIfNeeded_
 * existed, readBallotCandidateDetails_ would silently misread column B (the old
 * Details value) as the new Name column and an always-empty column C as Details —
 * returning every candidate with blank details instead of throwing, which is why this
 * went unnoticed until a user reported "I don't see any candidate details."
 */
function testMigratesExistingOldColumnCandidatesSectionOnRead() {
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
 * reads candidates — readBallotCandidates_ reconciles the Responses header to match
 * on every read, appending the missing column(s) without disturbing existing ones.
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

  // readBallotCandidates_ reconciles: Bob's column gets appended, Alice's existing
  // column/position (and Voter One's recorded rank) are left untouched.
  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Alice', 'Bob']);
  const rows = SM.readBallotResponseRows_(sheet);
  assert.equal(rows.length, 1);
  assert.deepEqual(rows[0].ranks, [1, '']); // Alice=1 (unchanged), Bob unranked

  // getBallotForRespondent_ (the RPC payload the ballot page renders from) also
  // sees the reconciled, complete candidate list.
  const result = SM.getBallotForRespondent_('PrePopulated', '');
  assert.deepEqual(result.candidates, ['Alice', 'Bob']);
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
}

/**
 * Regression test for a real bug seen on the SIT sheet: pre-populating the Candidates
 * table directly (skipping addBallotCandidate_) before any Responses columns exist meant
 * sheet.getLastColumn() reflected the Candidates table's own fixed-width header (through
 * its Link URL column) rather than the Responses header's candidate-column count, so the
 * old _reconcileCandidatesWithResponses_ appended candidate names one column to the right
 * of where submitBallotResponse_ actually writes ranks — leaving a blank gap column and
 * every candidate's name shifted off of its vote data.
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
 * re-created after the last row, so hand-pasted candidates are found instead of silently
 * dropped.
 */
function testMissingResponsesSectionSelfHealsAndReturnsCandidates() {
  const { SM, ss, sandbox } = loadBallotModel();
  const sheet = createBallotWithoutResponsesSection(SM, ss, 'NoResponses');
  sheet.getRange(13, 2, 2, 2).setValues([['Alice', 'first'], ['Bob', 'second']]);

  assert.deepEqual(SM.readBallotCandidateDetails_(sheet).map((it) => it.name), ['Alice', 'Bob']);
  assert.equal(sheet.getRange(15, 1).getValue(), '[Responses]');
  assert.deepEqual(sheet.getRange(16, 1, 1, 4).getValues()[0], ['Date', 'Name', 'Weight', 'Comment']);
  assert.deepEqual(SM.readBallotCandidates_(sheet), ['Alice', 'Bob']);
  assert.equal(logLinesWith(sandbox, 'ballot.repair.responsesSection').length, 1);
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
  assert.equal(sheet.getLastRow(), 13);
  assert.equal(logLinesWith(sandbox, 'ballot.parse.responsesMarkerMissing').length > 0, true);
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

/** Self-repair writes in read paths run inside the lock. */
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
  t.SM.readBallotConfig_(sheet);
  assertAllWritesLocked(log, 'readBallotConfig_ legacy key');

  // Candidate typed into the Candidates table only: Responses header reconcile.
  t = loadBallotModel();
  sheet = t.SM.createNewBallot_(t.ss, 'RepairF');
  t.SM.writeBallotCandidateDetails_(sheet, [{ name: 'Alice', details: '' }]);
  log = spyWrites(sheet, t.sandbox.LockService._state);
  t.SM.readBallotCandidates_(sheet);
  assertAllWritesLocked(log, '_reconcileCandidatesWithResponses_');

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
  const rows = [
    ['Title', 'T-' + id], ['Description', 'D'], ['Instructions', 'I'], ['Footer', 'F'],
    ['Contact', 'C'], ['Accept-New', 'TRUE'], ['Add-Instructions', 'AI'], ['Admin-Only-Notes', 'N'], [],
  ];
  order.forEach((key) => { sectionRows(key, opts && opts.blankRows).forEach((r) => rows.push(r)); });
  rows.forEach((r, i) => r.forEach((v, c) => { if (v !== '' && v != null) sheet.getRange(i + 1, c + 1).setValue(v); }));
  return sheet;
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

/** kix.4 AC5: each section reader scans column A exactly once on a healthy sheet. */
function testSectionReadersScanColumnAOnce() {
  const { SM, ss } = loadBallotModel();
  const sheet = buildOrderedBallot(ss, 'OneScan', ['responses', 'results', 'candidates']);
  let scans = 0;
  const origGetRange = sheet.getRange.bind(sheet);
  sheet.getRange = function (row, col, numRows, numCols) {
    if (row === 1 && col === 1 && numCols === 1 && numRows > 1) scans++;
    return origGetRange(row, col, numRows, numCols);
  };
  ['readBallotCandidateDetails_', 'readBallotCandidates_', 'readBallotResponseRows_'].forEach((fn) => {
    scans = 0;
    SM[fn](sheet);
    assert.equal(scans, 1, fn + ' column-A scans');
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

/** kix.6 AC1: each marker, removed alone, is restored per its rule and logged. */
function testEachMissingMarkerIsRestoredAndLogged() {
  // [Candidates] marker deleted: an empty section goes in directly above [Responses].
  let t = loadBallotModel();
  let sheet = buildOrderedBallot(t.ss, 'NoCand', ['results', 'candidates', 'responses']);
  const responsesRow = rowOf(sheet, '[Responses]');
  sheet.getRange(rowOf(sheet, '[Candidates]'), 1).setValue('');
  t.SM.readBallotCandidates_(sheet);
  assert.equal(sheet.getRange(responsesRow, 1).getValue(), '[Candidates]');
  assert.deepEqual(sheet.getRange(responsesRow + 1, 2, 1, 2).getValues()[0], ['Name', 'Details']);
  assert.equal(sheet.getRange(responsesRow + 2, 1).getValue(), '[Responses]');
  let lines = logLinesWith(t.sandbox, 'ballot.repair.candidatesSection');
  assert.equal(lines.length, 1);
  assert.equal(lines[0].indexOf('"row":' + responsesRow) !== -1, true, lines[0]);

  // [Candidates] and [Responses] both absent: Candidates is appended at the end, then
  // Responses below it (column A under the new Candidates header is blank).
  t = loadBallotModel();
  sheet = buildOrderedBallot(t.ss, 'OnlyResults', ['results']);
  const last = sheet.getLastRow();
  t.SM.readBallotCandidateDetails_(sheet);
  assert.equal(sheet.getRange(last + 1, 1).getValue(), '[Candidates]');
  assert.equal(sheet.getRange(last + 3, 1).getValue(), '[Responses]');
  assert.equal(logLinesWith(t.sandbox, 'ballot.repair.candidatesSection').length, 1);
  assert.equal(logLinesWith(t.sandbox, 'ballot.repair.responsesSection').length, 1);

  // [Responses] section gone, Results is the bottom section with column A blank below its
  // marker: Responses is appended after the last row.
  t = loadBallotModel();
  sheet = buildOrderedBallot(t.ss, 'NoResp', ['candidates', 'results']);
  const lastRow = sheet.getLastRow();
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
  let t = loadBallotModel();
  let sheet = buildOrderedBallot(t.ss, 'LostCand', ['results', 'candidates', 'responses']);
  sheet.getRange(rowOf(sheet, '[Candidates]'), 1).setValue('');
  t.SM._ensureBallotSections_(sheet);
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

function run() {
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
  testMigratesExistingOldColumnCandidatesSectionOnRead();
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
  testSectionReadersScanColumnAOnce();
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
  console.log('test_ballot_model: all tests passed');
}

run();
