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

  const files = ['BallotCache.js', 'BallotModel.js', 'onOpen.js', 'webAdmin.js', 'webBallot.js', 'Triggers.js'];
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
  assert.equal(first.title, '[TODO: ballot title shown to respondents]');

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

function testGetAdminListDataUsesCacheForTitleAndCandidateCount() {
  const { app, ss } = loadApp();
  const sheet = app.createNewBallot_(ss, 'ListCache');
  app.writeBallotConfig_(sheet, { Title: 'Listed Ballot' });
  app.addBallotCandidate_(sheet, 'Cand A');
  app.refreshBallotCache_('ListCache', sheet);

  const list = app.getAdminListData_();
  const entry = list.find(function (b) { return b.id === 'ListCache'; });
  assert.equal(entry.title, 'Listed Ballot');
  assert.equal(entry.candidateCount, 1);
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
  testGetAdminListDataUsesCacheForTitleAndCandidateCount();
  console.log('test_ballot_cache: all tests passed');
}

run();
