'use strict';

/**
 * test_ballot_cache.js
 *
 * Covers BallotCache.js's script-properties cache (config + Candidates table per ballot)
 * and its three "kept fresh" paths:
 *   1. Admin/ballot RPC writes (webAdmin.js's _adminWriteConfig_/adminAddCandidate/
 *      adminSaveCandidate, webBallot.js's addBallotTopic via BallotModel.js's
 *      addBallotCandidateForId_) refresh the cache immediately after writing.
 *   2. onEdit(e) (Triggers.js — a simple trigger) refreshes the cache after a manual
 *      sheet edit that bypasses every RPC.
 *   3. A cache miss transparently falls back to reading (and re-populating) the sheet.
 */

const assert = require('assert');
const vm = require('vm');
const fs = require('fs');
const path = require('path');

const { createFakeSpreadsheet, createFakeGasGlobals } = require('./fakeGas');
const { prodAbc2026_7Rows, sitAbc2026_08Rows, writeGrid } = require('./ballotFixtures');

/**
 * Loads BallotCache.js, BallotModel.js, onOpen.js (for _getWebAppUrl_), webAdmin.js,
 * webBallot.js, and Triggers.js into one vm sandbox wired to a fresh FakeSpreadsheet —
 * mirroring how GAS merges every script/*.js file into one global namespace. The returned
 * `app` (the sandbox/vm context) is itself that namespace — every top-level function in
 * these files (exported via module.exports or not) is reachable as `app.<name>`, same as
 * calling it unqualified from another file would be in the real GAS runtime.
 */
function loadApp() {
  const ss = createFakeSpreadsheet();
  const sandbox = Object.assign({ module: { exports: {} } }, createFakeGasGlobals(ss));
  vm.createContext(sandbox);

  const files = ['BallotCache.js', 'BallotModel.js', 'onOpen.js', 'webAdmin.js', 'webBallot.js', 'Triggers.js', 'processRCV.js', 'processCondorcet.js'];
  files.forEach(function (file) {
    sandbox.module = { exports: {} }; // discarded — this test reads functions off `sandbox` itself
    const src = fs.readFileSync(path.join(__dirname, '..', 'script', file), 'utf8');
    vm.runInContext(src, sandbox, { filename: file });
  });

  return { app: sandbox, ss: ss };
}

function scriptProps(app) {
  return app.PropertiesService.getScriptProperties();
}

function testCacheMissBuildsFromSheetAndPopulatesCache() {
  const { app, ss } = loadApp();
  app.createNewBallot_(ss, 'Cached1');
  const sheet = app.findBallotSheet_(ss, 'Cached1');
  app.writeBallotConfig_(sheet, { Title: 'My Ballot' });

  assert.equal(scriptProps(app).getProperty('ballotCache:Cached1'), null);

  const cached = app.getCachedBallotData_('Cached1');
  assert.equal(cached.config.Title, 'My Ballot');
  assert.ok(scriptProps(app).getProperty('ballotCache:Cached1')); // now populated
}

function testGetCachedBallotDataReturnsNullForUnknownBallot() {
  const { app } = loadApp();
  assert.equal(app.getCachedBallotData_('DoesNotExist'), null);
}

function testGetBallotForRespondentServesFromCacheWithoutTouchingSheetWhenNoName() {
  const { app, ss } = loadApp();
  app.createNewBallot_(ss, 'NoNameCache');
  const sheet = app.findBallotSheet_(ss, 'NoNameCache');
  app.writeBallotConfig_(sheet, { Title: 'Cache Me', Description: 'Desc' });
  app.addBallotCandidate_(sheet, 'Alice', 'A note');

  // Warm the cache once (this call is allowed to touch the sheet).
  app.getBallotForRespondent_('NoNameCache', '');

  // Now break the sheet lookup so any further sheet access would throw — a true cache
  // hit must not call findBallotSheet_/SpreadsheetApp again.
  const originalFind = app.findBallotSheet_;
  app.findBallotSheet_ = function () { throw new Error('sheet should not be looked up on a cache hit'); };

  const result = app.getBallotForRespondent_('NoNameCache', '');
  assert.equal(result.title, 'Cache Me');
  assert.equal(result.description, 'Desc');
  assert.deepEqual(result.candidates, ['Alice']);
  assert.equal(result.itemDetails.Alice.details, 'A note');

  app.findBallotSheet_ = originalFind; // restore for anything else in this test run
}

function testAdminWriteConfigRefreshesCache() {
  const { app, ss } = loadApp();
  app.createNewBallot_(ss, 'AdminWrite');

  const first = app.getAdminEditData('AdminWrite');
  assert.equal(first.title, '[TODO: ballot title shown to respondents — delete this bracket and edit the example] F3 Active Book Club — Next Book Selection');

  app.adminSaveTitle('AdminWrite', 'Updated Title');

  const cached = app._readBallotCacheRaw_('AdminWrite');
  assert.equal(cached.config.Title, 'Updated Title');

  const second = app.getAdminEditData('AdminWrite');
  assert.equal(second.title, 'Updated Title');
}

function testAdminAddCandidateRefreshesCache() {
  const { app, ss } = loadApp();
  app.createNewBallot_(ss, 'AdminAdd');
  app.adminAddCandidate('AdminAdd', 'Bob', 'Bob details');

  const cached = app._readBallotCacheRaw_('AdminAdd');
  assert.equal(cached.candidates.length, 1);
  assert.equal(cached.candidates[0].name, 'Bob');
  assert.equal(cached.candidates[0].details, 'Bob details');
}

function testAddBallotTopicRefreshesCache() {
  const { app, ss } = loadApp();
  const sheet = app.createNewBallot_(ss, 'RespondentAdd');
  app.writeBallotConfig_(sheet, { 'Accept-New': 'TRUE' });
  // Warm the cache first so the write path is refreshing an existing entry, not just
  // populating one for the first time.
  app.getCachedBallotData_('RespondentAdd', sheet);

  app.addBallotTopic('RespondentAdd', 'Write-in Candidate', '');

  const cached = app._readBallotCacheRaw_('RespondentAdd');
  assert.ok(cached.candidates.some(function (c) { return c.name === 'Write-in Candidate'; }));
}

function testOnEditRefreshesCacheAfterManualSheetEdit() {
  const { app, ss } = loadApp();
  app.createNewBallot_(ss, 'ManualEdit');
  const sheet = app.findBallotSheet_(ss, 'ManualEdit');
  app.getCachedBallotData_('ManualEdit', sheet); // warm the cache

  // Simulate a human typing directly into the Title cell (row 1, col 2) in Sheets —
  // bypasses every RPC, so only onEdit(e) can pick this up.
  const range = sheet.getRange(1, 2);
  range.setValue('Hand-Edited Title');

  app.onEdit({ range: range });

  const raw = app.PropertiesService.getScriptProperties().getProperty('ballotCache:ManualEdit');
  const parsed = JSON.parse(raw);
  assert.equal(parsed.config.Title, 'Hand-Edited Title');
}

function testOnEditIgnoresNonBallotSheetsAndMissingEvent() {
  const { app, ss } = loadApp();
  const other = ss.insertSheet('NotABallot');
  const range = other.getRange(1, 1);
  range.setValue('irrelevant');

  // Should not throw, and should not create any cache entry for a non-Ballot- sheet.
  app.onEdit({ range: range });
  app.onEdit(null);
  app.onEdit({});

  assert.equal(app.PropertiesService.getScriptProperties().getProperty('ballotCache:NotABallot'), null);
}

/** 8i5.3 AC4: the admin list costs one getValues per ballot (one snapshot serves every field) and writes nothing. */
function testGetAdminListDataIsOneReadPerBallot() {
  const { app, ss } = loadApp();
  const sheet = app.createNewBallot_(ss, 'ListCache');
  app.writeBallotConfig_(sheet, { Title: 'Listed Ballot' });
  app.addBallotCandidate_(sheet, 'Cand A');
  app.addBallotCandidate_(sheet, 'Cand B');
  app.submitBallotResponse_('ListCache', 'V1', ['Cand A'], '');
  const legacy = writeGrid(ss.insertSheet('Ballot-ABC-2026-7'), prodAbc2026_7Rows());
  const sheets = [sheet, legacy];
  sheets.forEach(function (sh) { sh._getValuesCalls = 0; sh._writeCalls = 0; });

  const list = app.getAdminListData_();
  sheets.forEach(function (sh) {
    assert.equal(sh._getValuesCalls, 1, sh.getName() + ' getValues');
    assert.equal(sh._writeCalls, 0, sh.getName() + ' writes');
  });
  const entry = list.find(function (b) { return b.id === 'ListCache'; });
  assert.equal(entry.title, 'Listed Ballot');
  assert.equal(entry.candidateCount, 2);
  assert.equal(entry.responseCount, 1);
  assert.equal(entry.uniqueRespondentCount, 1);
  assert.deepEqual(JSON.parse(JSON.stringify(entry.staleRespondents)), ['V1']); // left Cand B blank
}

/**
 * kix.1: the admin entry points that mutate the sheet do their reads+writes inside one
 * script-lock hold. The fake lock is non-re-entrant, so adminSaveCandidate's nested
 * saveBallotCandidatesForId_ also proves the lock is re-entrant in practice.
 */
function spyAdminWrites(sheet, lockState) {
  const log = { writes: 0, unlocked: [] };
  const origGetRange = sheet.getRange.bind(sheet);
  sheet.getRange = function () {
    const range = origGetRange.apply(null, arguments);
    ['setValue', 'setValues', 'clearContent'].forEach(function (m) {
      const orig = range[m].bind(range);
      range[m] = function () { log.writes++; if (!lockState.held) log.unlocked.push(m); return orig.apply(null, arguments); };
    });
    return range;
  };
  ['insertRowsBefore', 'deleteRows'].forEach(function (m) {
    const orig = sheet[m].bind(sheet);
    sheet[m] = function () { log.writes++; if (!lockState.held) log.unlocked.push(m); return orig.apply(null, arguments); };
  });
  return log;
}

function testAdminMutatingEntryPointsRunUnderLock() {
  const { app, ss } = loadApp();
  const state = app.LockService._state;
  const sheet = app.createNewBallot_(ss, 'AdminLock');
  app.addBallotCandidateForAdmin_('AdminLock', 'Alice', '', '', '');
  app.addBallotCandidateForAdmin_('AdminLock', 'Bob', '', '', '');
  app.submitBallotResponse_('AdminLock', 'V1', ['Alice', 'Bob'], '');
  app.submitBallotResponse_('AdminLock', 'V2', ['Bob', 'Alice'], '');
  const log = spyAdminWrites(sheet, state);

  app.runBallotAnalysis_(sheet);
  assert.ok(log.writes > 0);
  assert.deepEqual(log.unlocked, [], 'runBallotAnalysis_ wrote outside the lock');

  log.writes = 0;
  app._adminWriteConfig_('AdminLock', { Title: 'Locked' });
  assert.ok(log.writes > 0);
  assert.deepEqual(log.unlocked, [], '_adminWriteConfig_ wrote outside the lock');

  log.writes = 0;
  app.adminSaveCandidate('AdminLock', 0, 'Alice', 'note', '', '');
  assert.ok(log.writes > 0);
  assert.deepEqual(log.unlocked, [], 'adminSaveCandidate wrote outside the lock');
  assert.equal(state.held, false);
}

function logLines(app, tag) {
  return app.Logger._lines.filter(function (l) { return String(l).indexOf(tag) !== -1; });
}

/**
 * rcballot-8i5.7 AC3 + admin-load repair: on the legacy PROD Ballot-ABC-2026-7 layout,
 * analysis refuses (logged, nothing written) instead of reporting "no responses"; the
 * respondent page logs the damage; opening the admin edit page repairs it, after which
 * analysis runs over all 5 votes.
 */
function testLegacyResponsesLayoutRefusedThenRepairedOnAdminLoad() {
  const { app, ss } = loadApp();
  const sheet = writeGrid(ss.insertSheet('Ballot-ABC-2026-7'), prodAbc2026_7Rows());
  const before = JSON.stringify(sheet.getDataRange().getValues());

  const refused = app.getAdminAnalysisData_('ABC-2026-7');
  assert.ok(/no "Name" column/.test(refused.error), refused.error);
  assert.equal(JSON.stringify(sheet.getDataRange().getValues()), before, 'refusal wrote to the sheet');
  const lines = logLines(app, 'ballot.analysis.refused');
  assert.equal(lines.length, 1);
  assert.ok(lines[0].indexOf('"reason":"responsesMissingName"') !== -1, lines[0]);

  app.getBallotForRespondent_('ABC-2026-7', 'Pogo');
  assert.equal(logLines(app, 'ballot.parse.responsesMissingName').length, 1);

  const edit = app.getAdminEditData('ABC-2026-7');
  assert.equal(edit.candidates.length, 5);
  assert.equal(logLines(app, 'ballot.repair.headerOnMarkerRow').length, 1);

  const results = app.getAdminAnalysisData_('ABC-2026-7');
  assert.equal(results.error, undefined, results.error);
  assert.equal(results.rcv.summary.reduce(function (n, r) { return n + (Number(r[2]) || 0); }, 0) >= 5, true);
}

/** 8i5.3 AC2: onEdit only reads — even on a damaged sheet a repair would change. */
function testOnEditNeverWrites() {
  [prodAbc2026_7Rows, sitAbc2026_08Rows].forEach(function (rows, i) {
    const { app, ss } = loadApp();
    const sheet = writeGrid(ss.insertSheet('Ballot-Edit' + i), rows());
    sheet._writeCalls = 0;
    app.onEdit({ range: sheet.getRange(1, 2) });
    assert.equal(sheet._writeCalls, 0, 'onEdit wrote');
    assert.equal(app.LockService._state.waitCount, 0);
    assert.ok(app._readBallotCacheRaw_('Edit' + i)); // but did refresh the cache
  });
}

/** 8i5.3 AC3: admin edit-page load repairs a damaged sheet (missing [Responses]) and logs it; the cache follows. */
function testAdminEditLoadRepairsMissingResponses() {
  const { app, ss } = loadApp();
  const sheet = app.createNewBallot_(ss, 'NoResp');
  sheet.deleteRows(13, 2); // [Responses] marker + header lost
  sheet.getRange(13, 2, 2, 2).setValues([['Alice', 'a'], ['Bob', 'b']]); // hand-pasted candidates

  assert.deepEqual(app.getCachedBallotData_('NoResp').candidates.length, 0); // reads: nothing, and no repair
  assert.equal(app.findBallotSheet_(ss, 'NoResp').getRange(15, 1).getValue(), '');

  const edit = app.getAdminEditData('NoResp');
  assert.equal(sheet.getRange(15, 1).getValue(), '[Responses]');
  assert.deepEqual(edit.candidates.map(function (c) { return c.name; }), ['Alice', 'Bob']);
  assert.equal(logLines(app, 'ballot.repair.responsesSection').length, 1);
  assert.equal(logLines(app, 'ballot.repair.rankColumns').length, 1);
  assert.equal(app._readBallotCacheRaw_('NoResp').candidates.length, 2); // repair refreshed the cache
}

/** SIT Ballot-ABC-2026-08 through the real entry points: no candidates until the admin edit page loads, then all 8. */
function testHeaderlessCandidatesShownAfterAdminLoad() {
  const { app, ss } = loadApp();
  writeGrid(ss.insertSheet('Ballot-ABC-2026-08'), sitAbc2026_08Rows());
  assert.equal(app.getBallotForRespondent_('ABC-2026-08', '').candidates.length, 0);
  const edit = app.getAdminEditData('ABC-2026-08');
  assert.equal(edit.candidates.length, 8);
  assert.deepEqual(JSON.parse(JSON.stringify(edit.orphanColumns)), []);
  assert.equal(app.getBallotForRespondent_('ABC-2026-08', '').candidates.length, 8); // cache refreshed by the repair
  const pogo = app.getBallotForRespondent_('ABC-2026-08', 'Pogo');
  assert.equal(pogo.candidates[0], 'The Boys in the Boat'); // unranked ('' -> 0) sorts first, as before
  assert.equal(pogo.candidates[1], 'The Good Life'); // Pogo's 1st choice
  const results = app.getAdminAnalysisData_('ABC-2026-08');
  assert.equal(results.error, undefined, results.error);
  assert.equal(results.rcv.winner, 'The Good Life');
}

/** 8i5.3 AC5/AC6: a hand-deleted candidate's column is an orphan: shown on the edit page, excluded from analysis; ranks follow names. */
function testOrphanColumnShownOnEditPageAndExcludedFromAnalysis() {
  const { app, ss } = loadApp();
  const sheet = app.createNewBallot_(ss, 'Orphan');
  ['Alice', 'Bob', 'Carol'].forEach(function (n) { app.addBallotCandidateForAdmin_('Orphan', n, '', '', ''); });
  app.submitBallotResponse_('Orphan', 'V1', ['Alice', 'Carol', 'Bob'], '');
  app.submitBallotResponse_('Orphan', 'V2', ['Alice', 'Bob', 'Carol'], '');
  app.submitBallotResponse_('Orphan', 'V3', ['Carol', 'Bob', 'Alice'], '');
  // Hand-delete Alice's Candidates row (she would have won); her Responses column stays.
  const layout = app._readBallotLayout_(sheet);
  sheet.deleteRows(layout.sections.candidates.firstDataRow, 1);

  const edit = app.getAdminEditData('Orphan');
  assert.deepEqual(edit.candidates.map(function (c) { return c.name; }), ['Bob', 'Carol']);
  assert.deepEqual(JSON.parse(JSON.stringify(edit.orphanColumns)), [{ label: 'Alice', col: 5 }]);

  const results = app.getAdminAnalysisData_('Orphan');
  assert.equal(results.error, undefined, results.error);
  const counted = results.rcv.summary.slice(1).map(function (r) { return r[0]; });
  assert.equal(counted.indexOf('Alice'), -1, 'orphan counted: ' + JSON.stringify(counted));
  // Without Alice, by name: V1 Carol>Bob, V2 Bob>Carol, V3 Carol>Bob — Carol wins 2-1.
  assert.equal(results.rcv.winner, 'Carol');
  assert.equal(results.condorcet.condorcet.winner, 'Carol');
}

function run() {
  testCacheMissBuildsFromSheetAndPopulatesCache();
  testGetCachedBallotDataReturnsNullForUnknownBallot();
  testGetBallotForRespondentServesFromCacheWithoutTouchingSheetWhenNoName();
  testAdminWriteConfigRefreshesCache();
  testAdminAddCandidateRefreshesCache();
  testAddBallotTopicRefreshesCache();
  testOnEditRefreshesCacheAfterManualSheetEdit();
  testOnEditIgnoresNonBallotSheetsAndMissingEvent();
  testGetAdminListDataIsOneReadPerBallot();
  testAdminMutatingEntryPointsRunUnderLock();
  testLegacyResponsesLayoutRefusedThenRepairedOnAdminLoad();
  testOnEditNeverWrites();
  testAdminEditLoadRepairsMissingResponses();
  testHeaderlessCandidatesShownAfterAdminLoad();
  testOrphanColumnShownOnEditPageAndExcludedFromAnalysis();
  console.log('test_ballot_cache: all tests passed');
}

run();
