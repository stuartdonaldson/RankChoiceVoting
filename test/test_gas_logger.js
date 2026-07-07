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

console.log('test_gas_logger.js: PASS');
