'use strict';

const assert = require('node:assert/strict');

const {
  buildAxiomRows_,
  maskPiiForLog_,
  maskNameForLog_,
  maskRecipientListForLog_,
} = require('../script/GasLogger.js');

const entries = [
  { ts: '2026-06-20T09:03:18.000Z', tag: 'handleAdminPost_.createBallot', data: { id: 'Test123' }, execId: 'exec-1' },
  { ts: '2026-06-20T09:05:18.000Z', tag: 'handleAdminPost_.error', data: { warning: 'sheet not found' }, execId: 'exec-2', runId: 'gaslogger-test' },
];

const rows = buildAxiomRows_(entries, '0.0.1');

assert.equal(rows.length, 2);

assert.equal(rows[0]._time, '2026-06-20T09:03:18.000Z');
assert.equal(rows[0].name, 'handleAdminPost_.createBallot');
assert.equal(rows[0].side, 'gas');
assert.equal(rows[0].version, '0.0.1');
assert.equal(rows[0].id, 'Test123');
assert.equal(rows[0].execId, 'exec-1');
assert.equal('runId' in rows[0], false);

assert.equal(rows[1].execId, 'exec-2');
assert.equal(rows[1].runId, 'gaslogger-test');
assert.equal(rows[1].warning, 'sheet not found');

// target defaults to 'unknown' when an entry predates the version/target stamp.
assert.equal(rows[0].target, 'unknown');

// An entry already stamped with its own version/target (by GasLogger.log()) wins over
// the fallback passed to buildAxiomRows_ — SIT vs PROD stay distinguishable even in a
// shared dataset.
const stampedRows = buildAxiomRows_(
  [{ ts: '2026-06-24T00:00:00.000Z', tag: 'x', data: {}, version: '0.0.2', target: 'SIT' }],
  '9.9.9'
);
assert.equal(stampedRows[0].version, '0.0.2');
assert.equal(stampedRows[0].target, 'SIT');

// maskPiiForLog_ — names: first/last character kept, middle collapsed to '...'.
assert.equal(maskPiiForLog_('Little John'), 'L...n');
assert.equal(maskPiiForLog_('Jo'), 'J...o');
assert.equal(maskPiiForLog_('J'), 'J');
assert.equal(maskPiiForLog_(''), '');
assert.equal(maskPiiForLog_(null), '');

// maskPiiForLog_ — emails: only the local part is masked, domain stays fully visible.
assert.equal(maskPiiForLog_('stuart.donaldson@gmail.com'), 's...n@gmail.com');
assert.equal(maskPiiForLog_('a@b.com'), 'a@b.com');

// maskRecipientListForLog_ — plain comma-separated addresses.
assert.equal(
  maskRecipientListForLog_('stuart.donaldson@gmail.com,a@b.com'),
  's...n@gmail.com,a@b.com'
);

// maskRecipientListForLog_ — 'Display Name <email>' form, both parts masked.
assert.equal(
  maskRecipientListForLog_('Little John <stuart.donaldson@gmail.com>'),
  'L...n <s...n@gmail.com>'
);

assert.equal(maskRecipientListForLog_(''), '');
assert.equal(maskRecipientListForLog_(null), '');

// maskNameForLog_ — per-word first/last letter, middle collapsed to '..'; used where a
// masked-but-recognizable hint (e.g. in ballot action logs) is more useful than a single blob.
assert.equal(maskNameForLog_('Stuart Donaldson'), 'S..t D..n');
assert.equal(maskNameForLog_('Jo'), 'Jo');
assert.equal(maskNameForLog_('J'), 'J');
assert.equal(maskNameForLog_(''), '');
assert.equal(maskNameForLog_(null), '');

// scrubForLog_ / actor: request payloads are logged, so secrets must not survive and identity rides along.
const { scrubForLog_ } = require('../script/GasLogger.js');
assert.deepEqual(scrubForLog_({ adminSecret: 's', properties: { A: '1' }, n: 1, s: 'x'.repeat(250) }),
  { adminSecret: '[redacted]', properties: ['A'], n: 1, s: 'x'.repeat(200) + '...(+50)' });
assert.equal(buildAxiomRows_([{ ts: 't', tag: 'x', data: {}, actor: 'L..e J..n' }])[0].actor, 'L..e J..n');
assert.equal('actor' in buildAxiomRows_([{ ts: 't', tag: 'x', data: {} }])[0], false);

// rcballot-8un: the caller's email rides on every entry when known, back-filled at verification.
{
  const { GasLogger } = require('../script/GasLogger.js');
  global.Logger = { log() {} };
  global.Utilities = { getUuid: () => 'u-1' };
  global.PropertiesService = { getScriptProperties: () => ({ getProperty: () => null }) };
  GasLogger.disable(); GasLogger.enable();
  GasLogger.init('t');
  GasLogger.log('doPost', { n: 1 });
  assert.equal('email' in GasLogger._entries[0], false); // AC 4: none known yet
  GasLogger.setEmail(' Someone@Example.org ');
  GasLogger.log('after', {});
  assert.equal(GasLogger._entries[0].email, 'someone@example.org'); // back-filled
  assert.equal(GasLogger._entries[1].email, 'someone@example.org');
  const r = buildAxiomRows_(GasLogger._entries);
  assert.equal(r[0].email, 'someone@example.org');
  assert.equal('email' in buildAxiomRows_([{ ts: 't', tag: 'x', data: {} }])[0], false);
  // Session fallback, and a throwing Session never breaks init.
  global.Session = { getActiveUser: () => ({ getEmail: () => 'Deployer@Example.org' }) };
  GasLogger.init('t2');
  GasLogger.log('x', {});
  assert.equal(GasLogger._entries[0].email, 'deployer@example.org');
  global.Session = { getActiveUser: () => { throw new Error('no'); } };
  GasLogger.init('t3');
  GasLogger.log('x', {});
  assert.equal('email' in GasLogger._entries[0], false);
  delete global.Session;
  GasLogger._entries = [];
}

console.log('test_gas_logger.js: PASS');
