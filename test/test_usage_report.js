'use strict';

const assert = require('node:assert/strict');
const { summarize_, render_, parseArgs_ } = require('../tools/usage-report.js');

// rcballot-6mq: `pnpm run usage` — who used the app, doing what, when, grouped by env.
// The summariser is pure over Axiom `matches`, so this suite never touches the network.

const TZ = 'America/Los_Angeles';
let seq = 0;
function ev(time, data) {
  seq += 1;
  return { _time: time, _rowId: 'r' + seq, data: Object.assign({ side: 'gas' }, data) };
}
function api(time, target, fn, args, extra) {
  return ev(time, Object.assign({ name: 'doPost', cmd: 'api', target, body: { name: fn, args } }, extra || {}));
}

const matches = [
  // NUUC: a real voter who signs in with Google, then submits three times in a row.
  api('2026-11-21T04:48:39Z', 'NUUC', 'getBallotConfig', ['BeTheChurch']),
  api('2026-11-21T04:49:50Z', 'NUUC', 'recordSignin', ['BeTheChurch'], { email: 'voter@example.com' }),
  api('2026-11-21T04:50:02Z', 'NUUC', 'getBallotForName', ['BeTheChurch', 'M..n'], { email: 'voter@example.com', who: 'M..n' }),
  api('2026-11-21T04:50:21Z', 'NUUC', 'submitBallotRanking', ['BeTheChurch', 'M..n', ['A', 'B'], ''], { email: 'voter@example.com', who: 'M..n' }),
  api('2026-11-21T04:50:34Z', 'NUUC', 'submitBallotRanking', ['BeTheChurch', 'M..n', ['B', 'A'], ''], { email: 'voter@example.com', who: 'M..n' }),
  api('2026-11-21T04:50:43Z', 'NUUC', 'submitBallotRanking', ['BeTheChurch', 'M..n', ['A', 'B'], ''], { email: 'voter@example.com', who: 'M..n' }),
  // NUUC: name-only voter, never signs in. A later event pairs the same name with no email.
  api('2026-11-21T02:19:49Z', 'NUUC', 'getBallotForName', ['BeTheChurch', 'j..e'], { who: 'j..e' }),
  // NUUC: noise that must not appear as activity.
  ev('2026-11-21T04:50:02Z', { name: 'ballot.cache.refresh', target: 'NUUC' }),
  // NUUC: an anonymous admin results view that fails authorisation.
  api('2026-11-20T23:09:37Z', 'NUUC', 'getAdminAnalysisData_', ['BeTheChurch'], { who: 'admin' }),
  ev('2026-11-20T23:09:38Z', { name: 'getAdminAnalysisData_', target: 'NUUC', who: 'admin', message: 'Not authorized: sign in with Google to view results.' }),

  // SIT: a deploy (setWebappUrl + its version-verify posts) and a smoke test run.
  ev('2026-11-20T17:49:12Z', { name: 'doPost', target: 'SIT', who: 'admin' }),
  ev('2026-11-20T17:49:12Z', { name: 'handleAdminPost_.setWebappUrl', target: 'SIT', who: 'admin' }),
  ev('2026-11-20T17:49:58Z', { name: 'doPost', target: 'SIT' }),
  api('2026-11-20T17:51:08Z', 'SIT', 'submitBallotRanking', ['ApiSmokeTest1795', 'V..r O..e', ['X'], 'nice'], { who: 'V..r O..e' }),
  // SIT: an admin edit carrying an email.
  api('2026-11-21T03:04:26Z', 'SIT', 'adminSaveDescription', ['TEST', 'text'], { email: 'admin@example.org', who: 'admin' }),
];

const s = summarize_(matches, { tz: TZ });

// (2) grouped by env, in SIT, PROD, NUUC order; envs without events are omitted.
assert.deepEqual(s.envs.map(e => e.env), ['SIT', 'NUUC']);
const nuuc = s.envs.find(e => e.env === 'NUUC');
const sit = s.envs.find(e => e.env === 'SIT');

// (3) identities: email wins; a name seen with an email merges under it.
const ids = nuuc.people.map(p => p.id).sort();
assert.deepEqual(ids, ['admin (no email)', 'anonymous', 'j..e', 'voter@example.com'].sort());
const voter = nuuc.people.find(p => p.id === 'voter@example.com');
assert.deepEqual(voter.names, ['M..n']);
assert.deepEqual(voter.ballots, ['BeTheChurch']);
assert.equal(voter.counts.submitted, 3);
assert.equal(voter.counts['signed in'], 1);
assert.equal(voter.first, '11/20 20:49:50');
assert.equal(voter.last, '11/20 20:50:43');

// (3) smoke-test ballots are attributed to 'smoke test', not to the fake voter name.
assert.ok(sit.people.some(p => p.id === 'smoke test'));
assert.ok(!sit.people.some(p => p.id === 'V..r O..e'));
assert.equal(sit.people.find(p => p.id === 'admin@example.org').counts['admin edit'], 1);

// (4) timeline: cache refresh and deploy noise excluded; consecutive repeats collapsed.
assert.ok(!nuuc.timeline.some(t => /cache/.test(t.action)));
const submits = nuuc.timeline.filter(t => t.action === 'submitted');
assert.equal(submits.length, 1);
assert.equal(submits[0].times, 3);
assert.deepEqual(sit.deploys, ['11/20 09:49:12']);
assert.ok(!sit.timeline.some(t => t.id === 'anonymous'), 'version-verify posts are not activity');

// (4) a smoke run is one timeline line however many calls it made.
const smokeMatches = matches.concat([
  api('2026-11-20T17:51:11Z', 'SIT', 'submitBallotRanking', ['ApiSmokeTest1795', 'V..r T..o', ['X'], ''], { who: 'V..r T..o' }),
  api('2026-11-20T17:51:12Z', 'SIT', 'getAdminListData_', [], { who: 'admin' }),
  api('2026-11-20T17:51:18Z', 'SIT', 'addBallotTopic', ['ApiSmokeTest1795', 'Delta']),
]);
const smokeLines = summarize_(smokeMatches, { tz: TZ }).envs.find(e => e.env === 'SIT')
  .timeline.filter(t => t.id === 'smoke test');
assert.equal(smokeLines.length, 1);
assert.equal(smokeLines[0].times, 3);

// (5) problems
assert.equal(nuuc.problems.length, 1);
assert.match(nuuc.problems[0].message, /Not authorized/);

// render_ produces one section per env with the people and timeline headings.
const text = render_(s, { since: '24h', tz: TZ, env: null });
assert.match(text, /== SIT ==/);
assert.match(text, /== NUUC ==/);
assert.match(text, /voter@example\.com \(M\.\.n\)/);
assert.match(text, /submitted ×3/);
assert.match(text, /smoke test run \(1 calls\)/);
assert.match(text, /Deploys \(1\): 11\/20 09:49/);

// (1) arg parsing: --hours / --days map to --since; --env is normalised to upper case.
assert.deepEqual(
  (({ since, env }) => ({ since, env }))(parseArgs_(['node', 'x', '--days', '2', '--env', 'nuuc'])),
  { since: '2d', env: 'NUUC' });
assert.equal(parseArgs_(['node', 'x', '--hours', '6']).since, '6h');
assert.equal(parseArgs_(['node', 'x']).since, '24h');
assert.throws(() => parseArgs_(['node', 'x', '--env', 'bogus']), /--env/);

console.log('test_usage_report: all assertions passed');
