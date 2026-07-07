#!/usr/bin/env node
/**
 * smokeTest.js — integration smoke test against a live SIT deployment.
 *
 * Proves the create-survey flow works end to end through the real web app
 * (not just the pure SurveyModel logic under test/test_survey_model.js):
 *   1. Bootstrap sitAdminSecret if not already saved locally.
 *   2. createSurvey a uniquely-named test survey.
 *   3. listSheets / getSheet — assert the Survey-<id> sheet exists with the
 *      expected config skeleton (Title/Description/Instructions/Footer/Contact/
 *      Accept-New/Add-Instructions, Results marker, Responses header).
 *   4. setSheet — seed one response row.
 *   5. getSheet — verify the seeded row round-trips.
 *   6. deleteSheet — clean up.
 *
 * Requires a live SIT deployment (sitDeploymentId set in local.settings.json
 * — see tools/manage-deployments.js --deploy-sit). Prints PASS/FAIL for each
 * step and a final summary; exits non-zero on any failure.
 *
 * Usage:
 *   node tools/smokeTest.js
 */

'use strict';

const { post, loadSettings, saveSetting_, generateSecret_, ENV_MAP, UNGATED_ACTIONS } = require('./callWebapp.js');

const ENV = 'sit';

function url_(deploymentId) {
  return `https://script.google.com/macros/s/${deploymentId}/exec?cmd=admin`;
}

function payload_(action, extra, adminSecret) {
  if (UNGATED_ACTIONS.has(action)) return { action, ...extra };
  return { action, adminSecret, ...extra };
}

const results = [];
function record_(step, ok, detail) {
  results.push({ step, ok, detail });
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${step}${detail ? ' — ' + detail : ''}`);
}

async function ensureAdminSecret_(deploymentId, settings) {
  const { adminSecretKey } = ENV_MAP[ENV];
  if (settings[adminSecretKey]) {
    return settings[adminSecretKey];
  }

  console.log(`→ ${adminSecretKey} not set locally — bootstrapping...`);
  const secret = generateSecret_();
  const result = await post(url_(deploymentId), payload_('bootstrapSecret', { secret }));

  if (result && result.ok) {
    saveSetting_(adminSecretKey, secret);
    record_('bootstrapSecret', true, 'saved new secret to local.settings.json');
    return secret;
  }

  record_('bootstrapSecret', false, JSON.stringify(result));
  console.error(`
❌  Could not bootstrap an admin secret (server said: ${result && result.error}).
    If the server already has ADMIN_SHARED_SECRET set (from a prior bootstrap),
    you must supply the matching secret by hand:
      1. Open the Apps Script editor > Project Settings > Script Properties.
      2. Copy the ADMIN_SHARED_SECRET value.
      3. Add it to local.settings.json as "${adminSecretKey}".
    Then re-run: node tools/smokeTest.js
`);
  return null;
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

  const adminSecret = await ensureAdminSecret_(deploymentId, settings);
  if (!adminSecret) {
    process.exitCode = 1;
    return;
  }

  const id = 'SmokeTest' + Date.now();
  const sheetName = 'Survey-' + id;

  try {
    // 1. createSurvey
    const createResult = await post(url_(deploymentId), payload_('createSurvey', { id }, adminSecret));
    record_('createSurvey', !!(createResult && createResult.ok && createResult.sheetName === sheetName),
      JSON.stringify(createResult));
    if (!createResult || !createResult.ok) {
      throw new Error('createSurvey failed: ' + JSON.stringify(createResult));
    }

    // 2. listSheets — assert the new sheet is present
    const listResult = await post(url_(deploymentId), payload_('listSheets', {}, adminSecret));
    const found = !!(listResult && listResult.ok && listResult.sheets.some(s => s.name === sheetName));
    record_('listSheets finds ' + sheetName, found, JSON.stringify(listResult && listResult.sheets));
    if (!found) throw new Error(sheetName + ' not found in listSheets result');

    // 3. getSheet — assert the config skeleton
    const getResult1 = await post(url_(deploymentId), payload_('getSheet', { sheetName }, adminSecret));
    const csv1 = (getResult1 && getResult1.csv) || '';
    const hasConfigSkeleton = csv1.includes('Title') && csv1.includes('Results') && csv1.includes('Responses') &&
      csv1.includes('Date\tName\tWeight\tComment');
    record_('getSheet shows config skeleton', hasConfigSkeleton, hasConfigSkeleton ? '' : csv1.slice(0, 300));
    if (!hasConfigSkeleton) throw new Error('config skeleton missing from getSheet csv');

    // 4. setSheet — seed one response row directly below the Responses header (row 14 -> row 15).
    const seededComment = 'seeded by smokeTest ' + id;
    const setResult = await post(url_(deploymentId), payload_('setSheet', {
      sheetName,
      row: 15,
      col: 1,
      rows: [[new Date().toISOString(), 'Smoke Voter', 1, seededComment]],
    }, adminSecret));
    record_('setSheet seeds response row', !!(setResult && setResult.ok), JSON.stringify(setResult));
    if (!setResult || !setResult.ok) throw new Error('setSheet failed: ' + JSON.stringify(setResult));

    // 5. getSheet — verify the seeded row round-trips
    const getResult2 = await post(url_(deploymentId), payload_('getSheet', { sheetName }, adminSecret));
    const csv2 = (getResult2 && getResult2.csv) || '';
    const roundTripped = csv2.includes(seededComment);
    record_('getSheet round-trips seeded row', roundTripped, roundTripped ? '' : csv2.slice(0, 500));
    if (!roundTripped) throw new Error('seeded comment not found in getSheet csv');
  } finally {
    // 6. deleteSheet — cleanup, best-effort (don't mask earlier failures with a cleanup failure)
    const deleteResult = await post(url_(deploymentId), payload_('deleteSheet', { sheetName }, adminSecret));
    record_('deleteSheet cleanup', !!(deleteResult && deleteResult.ok), JSON.stringify(deleteResult));
  }

  const allPassed = results.every(r => r.ok);
  console.log('\n' + '='.repeat(60));
  console.log(allPassed ? 'SMOKE TEST: ALL PASSED' : 'SMOKE TEST: FAILED');
  console.log('='.repeat(60));
  if (!allPassed) process.exitCode = 1;
}

if (require.main === module) {
  main().catch(err => {
    console.error('❌ smokeTest crashed:', err.message);
    process.exitCode = 1;
  });
}

module.exports = { url_, payload_ };
