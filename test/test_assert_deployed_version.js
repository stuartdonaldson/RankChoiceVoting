'use strict';

const assert = require('assert');

// RECOMMENDATION.md §3.2 (gas-deploy Stage 1b/1c): assertDeployedVersion_ polls cmd=version
// until the webapp reports the exact version *and* target just stamped, or times out. Every
// path here uses an injected fake postFn/sleep — no real network call, no real wall-clock wait.
// Ported from F3Go30's test/test_assert_deployed_version.js; Stage 2 collapses both into
// packages/gas-deploy's lib/verify.js coverage.
const { assertDeployedVersion_ } = require('../tools/manage-deployments.js');

async function testMatchesOnFirstPoll() {
  const calls = [];
  const postFn = async (url, body) => {
    calls.push({ url, body });
    return { ok: true, version: '0.1.6.2', target: 'SIT', deploymentId: 'AKfycbwRCVTEST' };
  };
  const result = await assertDeployedVersion_('AKfycbwRCVTEST', '0.1.6.2', 'SIT', { postFn, log: () => {} });
  assert.deepEqual(result, { ok: true, attempts: 1, version: '0.1.6.2', target: 'SIT', deploymentId: 'AKfycbwRCVTEST' });
  assert.equal(calls.length, 1);
  assert.equal(calls[0].url, 'https://script.google.com/macros/s/AKfycbwRCVTEST/exec?cmd=version');
  // No secret anywhere in the request — cmd=version is ungated by design (§3.2).
  assert.deepEqual(calls[0].body, { action: 'version' });
}

async function testSucceedsAfterEdgePropagationDelay() {
  // The ~5s edge race (#9): the first poll still sees the previous version, the second sees the
  // new one. This is the propagation tolerance the AC requires.
  let attempt = 0;
  const postFn = async () => {
    attempt++;
    if (attempt === 1) return { ok: true, version: '0.1.6.1', target: 'SIT' };
    return { ok: true, version: '0.1.6.2', target: 'SIT' };
  };
  const sleeps = [];
  const sleep = async (ms) => { sleeps.push(ms); };
  const result = await assertDeployedVersion_('AKfycbwRCVTEST', '0.1.6.2', 'SIT', { postFn, sleep, log: () => {} });
  assert.equal(result.attempts, 2);
  assert.deepEqual(sleeps, [5000]);
}

async function testVersionMismatchEventuallyTimesOut() {
  const postFn = async () => ({ ok: true, version: '0.1.6.1', target: 'SIT' });
  let now = 0;
  const realNow = Date.now;
  Date.now = () => now;
  const sleep = async (ms) => { now += ms; };
  try {
    await assert.rejects(
      () => assertDeployedVersion_('AKfycbwRCVTEST', '0.1.6.2', 'SIT', { postFn, sleep, log: () => {}, intervalSec: 5, timeoutSec: 12 }),
      (err) => {
        assert.match(err.message, /timed out/);
        assert.match(err.message, /expected version=0\.1\.6\.2 target=SIT/);
        assert.match(err.message, /last seen version=0\.1\.6\.1 target=SIT/);
        return true;
      }
    );
  } finally {
    Date.now = realNow;
  }
}

async function testTargetMismatchEventuallyTimesOut() {
  // Wrong-environment deploy: version matches but target doesn't. RCV has three targets
  // (SIT/PROD/NUUC) sharing one version counter, so this check matters more here than in a
  // two-target project — a PROD-stamped build serving on SIT would otherwise pass silently.
  const postFn = async () => ({ ok: true, version: '0.1.6.2', target: 'PROD' });
  let now = 0;
  const realNow = Date.now;
  Date.now = () => now;
  const sleep = async (ms) => { now += ms; };
  try {
    await assert.rejects(
      () => assertDeployedVersion_('AKfycbwRCVTEST', '0.1.6.2', 'SIT', { postFn, sleep, log: () => {}, intervalSec: 5, timeoutSec: 12 }),
      (err) => {
        assert.match(err.message, /timed out/);
        assert.match(err.message, /last seen version=0\.1\.6\.2 target=PROD/);
        return true;
      }
    );
  } finally {
    Date.now = realNow;
  }
}

async function testUnreachableResponseTreatedAsMissAndCanStillTimeOut() {
  // post() rejects on a non-JSON/redirect-race response rather than returning an object — that
  // must count as a miss inside the poll loop, not escape it.
  const postFn = async () => { throw new Error('Non-JSON response'); };
  let now = 0;
  const realNow = Date.now;
  Date.now = () => now;
  const sleep = async (ms) => { now += ms; };
  try {
    await assert.rejects(
      () => assertDeployedVersion_('AKfycbwRCVTEST', '0.1.6.2', 'SIT', { postFn, sleep, log: () => {}, intervalSec: 5, timeoutSec: 6 }),
      (err) => {
        assert.match(err.message, /timed out/);
        assert.match(err.message, /last seen \(no response\)/);
        return true;
      }
    );
  } finally {
    Date.now = realNow;
  }
}

async function run() {
  await testMatchesOnFirstPoll();
  await testSucceedsAfterEdgePropagationDelay();
  await testVersionMismatchEventuallyTimesOut();
  await testTargetMismatchEventuallyTimesOut();
  await testUnreachableResponseTreatedAsMissAndCanStillTimeOut();
  console.log('test_assert_deployed_version: all tests passed');
}

run().catch(err => {
  console.error(err);
  process.exit(1);
});
