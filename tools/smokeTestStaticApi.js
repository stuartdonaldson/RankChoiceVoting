#!/usr/bin/env node
/**
 * smokeTestStaticApi.js — integration smoke test against a live SIT deployment, exercising the
 * static-pages migration specifically (rcballot-z4p): the doPost ?cmd=api JSON-RPC bridge
 * (script/ApiBridge.js) that static-pages/src/index.html now calls instead of
 * google.script.run, plus a check that the published static bundle is actually live on
 * GitHub Pages and pointed at this deployment.
 *
 * Unlike tools/smokeTest.js (which drives the secret-gated cmd=admin ops API), everything here
 * goes through cmd=api — unauthenticated, same as a real browser hitting the static page would
 * use — covering the full ballot lifecycle a respondent + admin would actually exercise:
 *   1. createBallotForAdmin_ a uniquely-named test ballot.
 *   2. getAdminListData_ — assert the new ballot appears.
 *   3. getAdminEditData — assert the default field skeleton.
 *   4. adminSaveTitle / adminSaveDescription / adminSaveAddSettings — edit landing-page fields.
 *   5. adminAddCandidate x3 — build a real candidate field (also seeds a center-squeeze-shaped
 *      RCV race so the finish-order computation has something non-trivial to chew on).
 *   6. adminSaveCandidate — edit one candidate's details/link.
 *   7. getBallotConfig / getBallotForName — the voting page's own two initial calls.
 *   8. submitBallotRanking x3 — three ballots with different orderings.
 *   9. addBallotTopic — a respondent-added candidate mid-race.
 *  10. getAdminAnalysisData_ — assert RCV + all four Condorcet methods return a sane shape.
 *  11. Fetch the live static page from GitHub Pages and assert it's stamped with this
 *      deployment's exec URL (proves the publish actually landed, not just committed).
 *  12. Cleanup: deleteSheet via the existing secret-gated cmd=admin ops API (best-effort).
 *
 * Requires a live SIT deployment (sitDeploymentId in local.settings.json — see
 * `npm run deploy:sit`). Prints PASS/FAIL per step; exits non-zero on any failure.
 *
 * Usage:
 *   node tools/smokeTestStaticApi.js
 */

'use strict';

const https = require('https');
const { post, loadSettings, saveSetting_, generateSecret_, ENV_MAP, UNGATED_ACTIONS } = require('./callWebapp.js');

const ENV = 'sit';

function apiUrl_(deploymentId) {
  return `https://script.google.com/macros/s/${deploymentId}/exec?cmd=api`;
}
function adminUrl_(deploymentId) {
  return `https://script.google.com/macros/s/${deploymentId}/exec?cmd=admin`;
}

/** Calls a whitelisted RPC (script/ApiBridge.js's _API_WHITELIST_) via cmd=api. */
function callApi_(deploymentId, name, ...args) {
  return post(apiUrl_(deploymentId), { name, args });
}

const results = [];
function record_(step, ok, detail) {
  results.push({ step, ok, detail });
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${step}${detail ? ' — ' + detail : ''}`);
}

function get_(url) {
  return new Promise((resolve, reject) => {
    https.get(url, res => {
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => resolve({ statusCode: res.statusCode, body: Buffer.concat(chunks).toString('utf8') }));
      res.on('error', reject);
    }).on('error', reject);
  });
}

async function ensureAdminSecret_(deploymentId, settings) {
  const { adminSecretKey } = ENV_MAP[ENV];
  if (settings[adminSecretKey]) return settings[adminSecretKey];

  const secret = generateSecret_();
  const result = await post(adminUrl_(deploymentId), { action: 'bootstrapSecret', secret });
  if (result && result.ok) {
    saveSetting_(adminSecretKey, secret);
    return secret;
  }
  return null; // already bootstrapped by someone else and we don't have it locally — cleanup step will just skip
}

async function main() {
  const settings = loadSettings();
  const { deploymentIdKey } = ENV_MAP[ENV];
  const deploymentId = settings[deploymentIdKey];

  if (!deploymentId || deploymentId.startsWith('<')) {
    console.error(`❌  ${deploymentIdKey} is not set in local.settings.json.`);
    console.error('    Run: npm run deploy:sit');
    process.exit(1);
  }

  const id = 'ApiSmokeTest' + Date.now();
  const sheetName = 'Ballot-' + id;
  const candidateA = 'Alpha';
  const candidateB = 'Bravo';
  const candidateC = 'Charlie';

  try {
    // 1. createBallotForAdmin_
    const created = await callApi_(deploymentId, 'createBallotForAdmin_', id);
    record_('createBallotForAdmin_', !!(created && created.result && created.result.ok), JSON.stringify(created));
    if (!created || !created.result || !created.result.ok) throw new Error('create failed: ' + JSON.stringify(created));

    // 2. getAdminListData_
    const listed = await callApi_(deploymentId, 'getAdminListData_');
    const found = !!(listed && listed.result && listed.result.some(b => b.id === id));
    record_('getAdminListData_ finds ' + id, found, found ? '' : JSON.stringify(listed));
    if (!found) throw new Error(id + ' not found in getAdminListData_ result');

    // 3. getAdminEditData default skeleton
    const editData1 = await callApi_(deploymentId, 'getAdminEditData', id);
    const skeletonOk = !!(editData1 && editData1.result && !editData1.result.error && Array.isArray(editData1.result.candidates));
    record_('getAdminEditData returns skeleton', skeletonOk, skeletonOk ? '' : JSON.stringify(editData1));
    if (!skeletonOk) throw new Error('getAdminEditData did not return a usable skeleton');

    // 4. Landing-page field edits
    const titleSave = await callApi_(deploymentId, 'adminSaveTitle', id, 'API Smoke Test Ballot');
    record_('adminSaveTitle', !!(titleSave && titleSave.result && titleSave.result.ok), JSON.stringify(titleSave));

    const descSave = await callApi_(deploymentId, 'adminSaveDescription', id, 'Seeded by smokeTestStaticApi.js');
    record_('adminSaveDescription', !!(descSave && descSave.result && descSave.result.ok), JSON.stringify(descSave));

    const addSettingsSave = await callApi_(deploymentId, 'adminSaveAddSettings', id, true, 'Add your own option!');
    record_('adminSaveAddSettings', !!(addSettingsSave && addSettingsSave.result && addSettingsSave.result.ok), JSON.stringify(addSettingsSave));

    // 5. Candidates
    const addA = await callApi_(deploymentId, 'adminAddCandidate', id, candidateA, 'First candidate', '', '');
    record_('adminAddCandidate ' + candidateA, !!(addA && addA.result && addA.result.candidate === candidateA), JSON.stringify(addA));
    const addB = await callApi_(deploymentId, 'adminAddCandidate', id, candidateB, 'Second candidate', '', '');
    record_('adminAddCandidate ' + candidateB, !!(addB && addB.result && addB.result.candidate === candidateB), JSON.stringify(addB));
    const addC = await callApi_(deploymentId, 'adminAddCandidate', id, candidateC, 'Third candidate', '', '');
    record_('adminAddCandidate ' + candidateC, !!(addC && addC.result && addC.result.candidate === candidateC), JSON.stringify(addC));

    // 6. Edit one candidate (index 0 = Alpha)
    const editCand = await callApi_(deploymentId, 'adminSaveCandidate', id, 0, candidateA, 'Edited details', 'More info', 'https://example.com/alpha');
    const editCandOk = !!(editCand && editCand.result && editCand.result.ok && editCand.result.details === 'Edited details');
    record_('adminSaveCandidate edits ' + candidateA, editCandOk, JSON.stringify(editCand));

    // 7. Voting page's own two initial calls
    const config = await callApi_(deploymentId, 'getBallotConfig', id);
    const configOk = !!(config && config.result && config.result.title === 'API Smoke Test Ballot');
    record_('getBallotConfig reflects saved title', configOk, JSON.stringify(config));

    const forVoter1 = await callApi_(deploymentId, 'getBallotForName', id, 'Voter One');
    const voter1Ok = !!(forVoter1 && forVoter1.result && forVoter1.result.candidates && forVoter1.result.candidates.length === 3);
    record_('getBallotForName returns 3 candidates', voter1Ok, JSON.stringify(forVoter1));

    // 8. Submit three ballots — deliberately not a Condorcet cycle, just three plausible orderings.
    const submit1 = await callApi_(deploymentId, 'submitBallotRanking', id, 'Voter One', [candidateA, candidateB, candidateC], 'nice ballot');
    record_('submitBallotRanking Voter One', !!(submit1 && submit1.result && submit1.result.ok), JSON.stringify(submit1));
    const submit2 = await callApi_(deploymentId, 'submitBallotRanking', id, 'Voter Two', [candidateB, candidateC, candidateA], '');
    record_('submitBallotRanking Voter Two', !!(submit2 && submit2.result && submit2.result.ok), JSON.stringify(submit2));
    const submit3 = await callApi_(deploymentId, 'submitBallotRanking', id, 'Voter Three', [candidateB, candidateA, candidateC], '');
    record_('submitBallotRanking Voter Three', !!(submit3 && submit3.result && submit3.result.ok), JSON.stringify(submit3));

    // 9. A respondent adding their own candidate mid-race
    const addTopic = await callApi_(deploymentId, 'addBallotTopic', id, 'Delta', 'added by a respondent');
    record_('addBallotTopic (respondent add)', !!(addTopic && addTopic.result && addTopic.result.candidate === 'Delta'), JSON.stringify(addTopic));

    // 10. Analysis
    const analysis = await callApi_(deploymentId, 'getAdminAnalysisData_', id);
    const a = analysis && analysis.result;
    const analysisOk = !!(a && !a.error && a.rcv && (a.rcv.winner || a.rcv.tie) &&
      a.condorcet && a.condorcet.condorcet && a.condorcet.schulze && a.condorcet.rankedPairs && a.condorcet.minimax &&
      a.finishOrders && a.finishOrders.condorcet);
    record_('getAdminAnalysisData_ returns full result shape', analysisOk, analysisOk ? ('winner=' + (a.rcv.winner || 'tie:' + a.rcv.tie)) : JSON.stringify(analysis));

    // 11. Live static page check — confirms the published bundle is actually serving and
    // pointed at THIS deployment, not just committed to the hosting repo.
    const expectedExecUrl = `https://script.google.com/macros/s/${deploymentId}/exec`;
    const staticUrl = 'https://f3go30.github.io/static-pages/ballot/sit/';
    try {
      const page = await get_(staticUrl);
      const live = page.statusCode === 200 && page.body.includes(expectedExecUrl);
      record_('live static page (' + staticUrl + ') stamped with this deployment', live,
        live ? '' : `statusCode=${page.statusCode}, execUrlFound=${page.body.includes(expectedExecUrl)}`);
    } catch (err) {
      record_('live static page fetch', false, err.message + ' (GitHub Pages propagation lag? retry in ~1 min)');
    }
  } finally {
    // 12. Cleanup — best-effort, secret-gated cmd=admin (separate from cmd=api above).
    const secret = await ensureAdminSecret_(deploymentId, loadSettings());
    if (secret) {
      const del = await post(adminUrl_(deploymentId), { action: 'deleteSheet', sheetName, adminSecret: secret });
      record_('deleteSheet cleanup', !!(del && del.ok), JSON.stringify(del));
    } else {
      record_('deleteSheet cleanup', false, 'no admin secret available locally — clean up ' + sheetName + ' by hand');
    }
  }

  const allPassed = results.every(r => r.ok);
  console.log('\n' + '='.repeat(60));
  console.log(allPassed ? 'STATIC API SMOKE TEST: ALL PASSED' : 'STATIC API SMOKE TEST: FAILED');
  console.log('='.repeat(60));
  if (!allPassed) process.exitCode = 1;
}

if (require.main === module) {
  main().catch(err => {
    console.error('❌ smokeTestStaticApi crashed:', err.message);
    process.exitCode = 1;
  });
}

module.exports = { apiUrl_, adminUrl_, callApi_ };
