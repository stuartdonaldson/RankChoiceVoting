'use strict';

const assert = require('assert');

// Pure functions exported for testing. main() is not called on require
// because the module checks require.main === module before calling main().
const { parseArgs_, buildPayload_, ENV_MAP, UNGATED_ACTIONS } = require('../tools/callWebapp.js');

// --- parseArgs_ ---

function testParseArgsDefaults() {
  const r = parseArgs_(['node', 'callWebapp.js', 'listSheets']);
  assert.equal(r.action, 'listSheets');
  assert.equal(r.env, 'sit');
  assert.deepEqual(r.extraBody, {});
}

function testParseArgsAllFlags() {
  const r = parseArgs_([
    'node', 'callWebapp.js', 'createBallot',
    '--env', 'prod',
    '--body', '{"id":"SmokeTest1"}',
  ]);
  assert.equal(r.action, 'createBallot');
  assert.equal(r.env, 'prod');
  assert.deepEqual(r.extraBody, { id: 'SmokeTest1' });
}

function testParseArgsBodyMerged() {
  const r = parseArgs_([
    'node', 'callWebapp.js', 'setScriptProperties',
    '--body', '{"properties":{"AXIOM_DATASET":"ballot"}}',
  ]);
  assert.deepEqual(r.extraBody, { properties: { AXIOM_DATASET: 'ballot' } });
}

function testParseArgsFlagsBeforeAction() {
  // Flags placed before the action must not have their values mistaken for the action.
  const r = parseArgs_([
    'node', 'callWebapp.js',
    '--env', 'sit',
    'createBallot',
    '--body', '{"id":"SmokeTest1"}',
  ]);
  assert.equal(r.action, 'createBallot');
  assert.equal(r.env, 'sit');
  assert.deepEqual(r.extraBody, { id: 'SmokeTest1' });
}

// --- buildPayload_ ---

function testBuildPayloadInjectsSecret() {
  const p = buildPayload_('listSheets', {}, 'secret99');
  assert.deepEqual(p, { action: 'listSheets', adminSecret: 'secret99' });
}

function testBuildPayloadMergesExtraBody() {
  const p = buildPayload_('createBallot', { id: 'SmokeTest1' }, 's3cr3t');
  assert.deepEqual(p, { action: 'createBallot', adminSecret: 's3cr3t', id: 'SmokeTest1' });
}

function testBuildPayloadBootstrapSecretNoSecretField() {
  const p = buildPayload_('bootstrapSecret', { secret: 'abc123' }, 'ignored');
  assert.deepEqual(p, { action: 'bootstrapSecret', secret: 'abc123' });
  assert.ok(!('adminSecret' in p));
}

function testBuildPayloadSetWebappUrlNoSecretField() {
  const p = buildPayload_('setWebappUrl', {}, 'ignored');
  assert.deepEqual(p, { action: 'setWebappUrl' });
  assert.ok(!('adminSecret' in p));
}

function testEnvMapShape() {
  assert.deepEqual(ENV_MAP.sit, { deploymentIdKey: 'sitDeploymentId', adminSecretKey: 'sitAdminSecret' });
  assert.deepEqual(ENV_MAP.prod, { deploymentIdKey: 'prodDeploymentId', adminSecretKey: 'prodAdminSecret' });
}

function testUngatedActionsSet() {
  assert.ok(UNGATED_ACTIONS.has('bootstrapSecret'));
  assert.ok(UNGATED_ACTIONS.has('setWebappUrl'));
  assert.ok(!UNGATED_ACTIONS.has('createBallot'));
}

function run() {
  testParseArgsDefaults();
  testParseArgsAllFlags();
  testParseArgsBodyMerged();
  testParseArgsFlagsBeforeAction();
  testBuildPayloadInjectsSecret();
  testBuildPayloadMergesExtraBody();
  testBuildPayloadBootstrapSecretNoSecretField();
  testBuildPayloadSetWebappUrlNoSecretField();
  testEnvMapShape();
  testUngatedActionsSet();
  console.log('test_callwebapp: all tests passed');
}

run();
