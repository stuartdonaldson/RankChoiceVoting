'use strict';

/**
 * rcballot-dx3.2 / rcballot-dx3.3 — NUUC-Dispatch sign-in and ballot admin authorization.
 *
 * Loads every script/*.js file the ?cmd=api path touches into one vm sandbox (as GAS merges
 * them into one global namespace) over fakeGas, plus a node:crypto HMAC shim for
 * Utilities.computeHmacSha256Signature (same shim as GAS-Core libs/LibIdentity/test). Tokens are
 * minted here exactly as NUUC-Dispatch issues them. Requests go through ApiBridge.js's
 * handleApiPost_ with the real envelope, so the verify-once-per-request seam is exercised too.
 *
 * Covers: verification seam (valid, forged signature, wrong aud, expired, missing key ->
 * disabled), Responses Email/Picture/Google-Sub columns + self-heal of an existing ballot,
 * resume by sub across a renamed typed name, Signins upsert, the admin rule matrix across every
 * gated RPC (unclaimed/claimed x anonymous/non-admin/ballot-admin/super-admin), add-myself, the
 * lockout guard, PROD staying open, and that no email/name/assertion reaches a log entry.
 */

const assert = require('assert');
const crypto = require('crypto');
const vm = require('vm');
const fs = require('fs');
const path = require('path');

const { createFakeSpreadsheet, createFakeGasGlobals } = require('./fakeGas');

const SECRET = 'rcballot-test-secret-value-0123456789abcdef';
const KID = 'ASSERTION_KEY_RCBALLOT_TEST_1';
const NOW = Math.floor(Date.now() / 1000);

function toSignedBytes(buf) { return Array.from(buf, (b) => (b > 127 ? b - 256 : b)); }
function toBuffer(bytes) { return Buffer.from(bytes.map((b) => (b < 0 ? b + 256 : b))); }

function loadApp(opts) {
  opts = opts || {};
  const ss = createFakeSpreadsheet();
  const g = createFakeGasGlobals(ss);
  const logs = [];
  Object.assign(g.Utilities, {
    computeHmacSha256Signature(value, key) {
      return toSignedBytes(crypto.createHmac('sha256', Buffer.from(key, 'utf8')).update(Buffer.from(value, 'utf8')).digest());
    },
    base64EncodeWebSafe(data) {
      const buf = typeof data === 'string' ? Buffer.from(data, 'utf8') : toBuffer(data);
      return buf.toString('base64').replace(/\+/g, '-').replace(/\//g, '_');
    },
    base64DecodeWebSafe(str) {
      if (!/^[A-Za-z0-9\-_=]*$/.test(str)) throw new Error('bad web-safe base64');
      return toSignedBytes(Buffer.from(str.replace(/-/g, '+').replace(/_/g, '/'), 'base64'));
    },
    newBlob(bytes) { const buf = toBuffer(bytes); return { getDataAsString: () => buf.toString('utf8') }; },
  });
  const sandbox = Object.assign({ module: { exports: {} } }, g, {
    GasLogger: {
      log(tag, data) { logs.push({ tag, data }); },
      logError(tag, err, data) { logs.push({ tag, data, message: String(err && err.message) }); },
      flush() {},
    },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/TEST/exec' }) },
    HtmlService: { createHtmlOutput: (h) => ({ _html: h }) },
    jsonOutput_: (obj) => ({ _json: obj }),
    BUILD_INFO: { env: opts.target || 'SIT' },
    APP_DEPLOY_TARGET: opts.target || 'SIT',
  });
  vm.createContext(sandbox);
  ['BallotCache.js', 'BallotModel.js', 'onOpen.js', 'libIdentity.js', 'Identity.js', 'webAdmin.js', 'webBallot.js', 'ApiBridge.js'].forEach((file) => {
    sandbox.module = { exports: {} };
    vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'script', file), 'utf8'), sandbox, { filename: file });
  });
  if (opts.key !== false) g.PropertiesService.getScriptProperties().setProperty(opts.kid || KID, SECRET);
  return { app: sandbox, ss, logs };
}

function b64u(obj) {
  return Buffer.from(JSON.stringify(obj)).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}
function mint(claims, o) {
  o = o || {};
  const header = b64u({ alg: 'HS256', typ: 'JWT', kid: o.kid || KID });
  const body = b64u({
    iss: 'nuuc-dispatch', aud: o.aud || 'rcballot-test', iat: NOW - 10, exp: o.exp || NOW + 3600,
    sub: claims.sub, email: claims.email, email_verified: claims.email_verified !== false,
  });
  const sig = crypto.createHmac('sha256', Buffer.from(o.secret || SECRET, 'utf8')).update(`${header}.${body}`).digest('base64')
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  return `${header}.${body}.${sig}`;
}

const ALICE = { sub: '1001', email: 'alice@example.com' };     // ballot admin
const BOB = { sub: '1002', email: 'bob@example.com' };         // signed in, not an admin
const ROOT = { sub: '1003', email: 'root@example.com' };       // super-admin (Config!Admins)

function call(app, name, args, who, profile) {
  const body = { name, args: args || [] };
  if (who) body.assertion = typeof who === 'string' ? who : mint(who);
  if (profile) body.profile = profile;
  return app.handleApiPost_({ postData: { contents: JSON.stringify(body) } })._json;
}

function newBallot(t, id, cands) {
  const sheet = t.app.createNewBallot_(t.ss, id);
  (cands || ['Alpha', 'Beta']).forEach((c) => t.app.addBallotCandidateForAdmin_(id, c));
  return sheet;
}

function setSuperAdmins(t, list) {
  const sheet = t.ss.getSheetByName('Config') || t.ss.insertSheet('Config');
  sheet.getRange(1, 1, 2, 2).setValues([['Key', 'Value'], ['Admins', list]]);
}

function responsesHeader(sheet) {
  const values = sheet.getDataRange ? sheet.getDataRange().getValues() : null;
  const r = values.findIndex((row) => String(row[0]).trim() === '[Responses]');
  return values[r + 1].map(String).filter(Boolean);
}

const tests = [];
function test(name, fn) { tests.push([name, fn]); }

// ---------------------------------------------------------------- verification seam
test('valid assertion verifies; reports ok with the verified email', () => {
  const t = loadApp();
  assert.equal(t.app.rcvIdentityEnabled_(), true);
  const v = t.app.rcvVerifyAssertion_(mint(ALICE));
  assert.equal(v.ok, true);
  assert.equal(v.email, ALICE.email);
  assert.equal(v.sub, ALICE.sub);
});

test('forged signature, wrong aud, expired, unverified email -> rejected by reason code', () => {
  const t = loadApp();
  assert.equal(t.app.rcvVerifyAssertion_(mint(ALICE, { secret: 'not-the-key-xxxxxxxxxxxxxxxx' })).reason, 'bad_signature');
  assert.equal(t.app.rcvVerifyAssertion_(mint(ALICE, { aud: 'rcballot' })).reason, 'bad_aud');
  assert.equal(t.app.rcvVerifyAssertion_(mint(ALICE, { exp: NOW - 5 })).reason, 'expired');
  assert.equal(t.app.rcvVerifyAssertion_(mint(Object.assign({}, ALICE, { email_verified: false }))).reason, 'email_unverified');
  assert.equal(t.app.rcvVerifyAssertion_('garbage').reason, 'malformed');
});

test('missing key Script Property -> not enabled: assertion ignored, no auth field', () => {
  const t = loadApp({ key: false });
  assert.equal(t.app.rcvIdentityEnabled_(), false);
  assert.equal(t.app.rcvVerifyAssertion_(mint(ALICE)).reason, 'disabled');
  newBallot(t, 'b1');
  const res = call(t.app, 'getBallotConfig', ['b1'], ALICE);
  assert.ok(res.result && !res.error);
  assert.ok(!('auth' in res));
});

test('PROD target has no audience -> disabled even with a key present', () => {
  const t = loadApp({ target: 'PROD' });
  assert.equal(t.app.rcvAssertionAud_(), '');
  assert.equal(t.app.rcvIdentityEnabled_(), false);
});

test('NUUC target uses aud rcballot and its own key', () => {
  const t = loadApp({ target: 'NUUC', kid: 'ASSERTION_KEY_RCBALLOT_1' });
  assert.equal(t.app.rcvAssertionAud_(), 'rcballot');
  assert.equal(t.app.rcvVerifyAssertion_(mint(ALICE, { aud: 'rcballot', kid: 'ASSERTION_KEY_RCBALLOT_1' })).ok, true);
  assert.equal(t.app.rcvVerifyAssertion_(mint(ALICE)).ok, false); // a SIT token is not a NUUC token
});

test('envelope: forged assertion -> anonymous vote, auth.status rejected, reason logged only', () => {
  const t = loadApp();
  const sheet = newBallot(t, 'b1');
  const forged = mint(ALICE, { secret: 'forged-forged-forged-forged-xx' });
  const res = call(t.app, 'submitBallotRanking', ['b1', 'Alice Typed', ['Beta', 'Alpha'], ''], forged);
  assert.deepEqual(res.result, { ok: true });
  assert.deepEqual(res.auth, { status: 'rejected', reason: 'bad_signature' });
  const rows = t.app.readBallotResponseRows_(sheet);
  assert.equal(rows[0].email, '');
  assert.equal(rows[0].sub, '');
  assert.ok(t.logs.some((l) => l.tag === 'identity.assertion.rejected' && l.data.reason === 'bad_signature'));
  assert.equal(t.ss.getSheetByName('Signins'), null);
});

// ---------------------------------------------------------------- Responses columns, resume, Signins
test('signed-in submit writes Email/Picture/Google-Sub (self-heals an existing ballot) and Signins', () => {
  const t = loadApp();
  const sheet = newBallot(t, 'b1');
  // An existing anonymous vote: the ballot has no identity columns yet.
  call(t.app, 'submitBallotRanking', ['b1', 'Anon', ['Alpha', 'Beta'], '']);
  assert.ok(!responsesHeader(sheet).includes('Email'));

  const pic = 'https://lh3.googleusercontent.com/a/abc=s96-c';
  const res = call(t.app, 'submitBallotRanking', ['b1', 'Alice A', ['Beta', 'Alpha'], 'hi'], ALICE, { name: 'Alice Example', picture: pic });
  assert.deepEqual(res.result, { ok: true });
  assert.deepEqual(res.auth, { status: 'ok', email: ALICE.email });
  const header = responsesHeader(sheet);
  ['Email', 'Picture', 'Google-Sub'].forEach((c) => assert.ok(header.includes(c), c));
  const rows = t.app.readBallotResponseRows_(sheet);
  assert.equal(rows.length, 2);
  const a = rows.find((r) => r.sub === ALICE.sub);
  assert.equal(a.email, ALICE.email);
  assert.equal(a.picture, pic);
  assert.equal(a.name, 'Alice A');
  // Identity columns are not candidates and not orphan rank columns.
  assert.deepEqual(t.app.readBallotCandidates_(sheet), ['Alpha', 'Beta']);
  assert.deepEqual(t.app._readBallotView_(sheet).orphans, []);

  const signins = t.ss.getSheetByName('Signins');
  const s = signins.getRange(1, 1, signins.getLastRow(), 7).getValues();
  assert.deepEqual(s[0], ['Email', 'Name', 'Picture', 'Sub', 'First Seen', 'Last Seen', 'Ballots']);
  assert.equal(s.length, 2);
  assert.deepEqual([s[1][0], s[1][1], s[1][2], s[1][3], s[1][6]], [ALICE.email, 'Alice Example', pic, ALICE.sub, 'b1']);
});

test('untrusted picture URL is dropped; formula-like display name is neutralised', () => {
  const t = loadApp();
  newBallot(t, 'b1');
  call(t.app, 'recordSignin', ['b1'], ALICE, { name: '=IMPORTXML("x")', picture: 'https://evil.example.com/p.png' });
  const s = t.ss.getSheetByName('Signins').getRange(2, 1, 1, 7).getValues()[0];
  assert.equal(s[2], '');
  assert.equal(s[1].charAt(0), "'");
});

test('Signins upserts by sub: one row, First Seen kept, Ballots unioned', () => {
  const t = loadApp();
  newBallot(t, 'b1');
  newBallot(t, 'b2');
  assert.deepEqual(call(t.app, 'recordSignin', ['b1'], ALICE).result, { signedIn: true });
  const first = t.ss.getSheetByName('Signins').getRange(2, 5).getValue();
  call(t.app, 'recordSignin', ['b2'], ALICE);
  call(t.app, 'recordSignin', ['b1'], ALICE);
  call(t.app, 'recordSignin', [''], BOB);
  const sheet = t.ss.getSheetByName('Signins');
  assert.equal(sheet.getLastRow(), 3);
  const row = sheet.getRange(2, 1, 1, 7).getValues()[0];
  assert.equal(row[4], first);
  assert.equal(row[6], 'b1, b2');
  assert.deepEqual(call(t.app, 'recordSignin', ['b1']).result, { signedIn: false }); // anonymous: no-op
});

test('resume by sub across a renamed typed name; anonymous name match never takes a signed-in row', () => {
  const t = loadApp();
  const sheet = newBallot(t, 'b1');
  call(t.app, 'submitBallotRanking', ['b1', 'Alice', ['Beta', 'Alpha'], 'first'], ALICE);
  const resumed = call(t.app, 'getBallotForName', ['b1', 'Ally Renamed'], ALICE).result;
  assert.deepEqual(resumed.candidates, ['Beta', 'Alpha']);
  assert.equal(resumed.comment, 'first');
  call(t.app, 'submitBallotRanking', ['b1', 'Ally Renamed', ['Alpha', 'Beta'], 'second'], ALICE);
  let rows = t.app.readBallotResponseRows_(sheet);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].name, 'Ally Renamed');

  // An anonymous voter typing the same name neither sees nor overwrites Alice's row.
  const anon = call(t.app, 'getBallotForName', ['b1', 'Ally Renamed']).result;
  assert.equal(anon.comment, '');
  call(t.app, 'submitBallotRanking', ['b1', 'Ally Renamed', ['Beta', 'Alpha'], 'anon'], null);
  rows = t.app.readBallotResponseRows_(sheet);
  assert.equal(rows.length, 2);
  assert.equal(t.app.countUniqueBallotRespondents_(sheet), 2);
});

test('a signed-in voter adopts their earlier anonymous row with the same name', () => {
  const t = loadApp();
  const sheet = newBallot(t, 'b1');
  call(t.app, 'submitBallotRanking', ['b1', 'Bob', ['Alpha', 'Beta'], 'anon']);
  call(t.app, 'submitBallotRanking', ['b1', 'Bob', ['Beta', 'Alpha'], 'signed'], BOB);
  const rows = t.app.readBallotResponseRows_(sheet);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].sub, BOB.sub);
});

test('identity column names are reserved as candidate names', () => {
  const t = loadApp();
  newBallot(t, 'b1');
  assert.throws(() => t.app.addBallotCandidateForAdmin_('b1', 'Email'), /already exists/);
  assert.throws(() => t.app.addBallotCandidateForAdmin_('b1', 'google-sub'), /already exists/);
});

// ---------------------------------------------------------------- admin rule matrix
// Every gated RPC with arguments valid enough to run when allowed. A refusal must be the
// 'Not authorized' error; an allowed call must not be.
function gatedCalls(id) {
  return [
    ['getAdminEditData', [id]],
    ['getAdminAnalysisData_', [id]],
    ['adminSaveTitle', [id, 'T']],
    ['adminSaveDescription', [id, 'D']],
    ['adminSaveInstructions', [id, 'I']],
    ['adminSaveFooter', [id, 'F']],
    ['adminSaveContact', [id, 'C']],
    ['adminSaveAdminOnlyNotes', [id, 'N']],
    ['adminSaveDescriptionLink', [id, 'L', 'https://example.com']],
    ['adminSaveFeedbackPrompt', [id, 'P']],
    ['adminSaveShowDetailsPreview', [id, true]],
    ['adminSaveAddSettings', [id, true, 'x']],
    ['adminAddCandidate', [id, 'Gamma' + Math.random().toString(36).slice(2, 6)]],
    ['adminSaveCandidate', [id, 0, 'Alpha', 'details', '', '']],
    ['adminReorderCandidates', [id, ['Beta', 'Alpha']]],
    ['adminLinkOrphanColumn', [id, 'NoSuchColumn', 'Alpha']],
    ['adminDiscardOrphanColumn', [id, 'NoSuchColumn']],
    ['adminDeleteCandidate', [id, 'NoSuchCandidate']],
    ['adminClearResponses', [id, id]],
    ['adminSaveAdmins', [id, ALICE.email, true]],
    ['adminClaimBallot', [id]],
    ['createBallotForAdmin_', [id]],
  ];
}

test('every gated RPC is in the ApiBridge whitelist', () => {
  const t = loadApp();
  gatedCalls('x').forEach(([name]) => assert.equal(typeof t.app._API_WHITELIST_[name], 'function', name));
});

const NOT_AUTH = /^Not authorized/;
function outcome(res) { return res.error && NOT_AUTH.test(res.error) ? 'refused' : 'allowed'; }

test('claimed ballot: anonymous and non-admin refused on every admin RPC; ballot admin and super-admin allowed', () => {
  const cases = [
    ['anonymous', null, 'refused'],
    ['non-admin', BOB, 'refused'],
    ['ballot-admin', ALICE, 'allowed'],
    ['super-admin', ROOT, 'allowed'],
  ];
  cases.forEach(([label, who, expected]) => {
    gatedCalls('b1').forEach(([name, args]) => {
      const t = loadApp();
      newBallot(t, 'b1');
      setSuperAdmins(t, ROOT.email);
      t.app._adminWriteConfigUnchecked_('b1', { Admins: ALICE.email });
      const res = call(t.app, name, args, who);
      assert.equal(outcome(res), expected, `${label} ${name}: ${res.error || 'ok'}`);
    });
  });
});

test('unclaimed ballot: every admin RPC open to anonymous, non-admin and admins', () => {
  [null, BOB, ROOT].forEach((who) => {
    gatedCalls('b1').forEach(([name, args]) => {
      if (name === 'adminClaimBallot' && !who) return; // needs a signed-in caller (tested below)
      const t = loadApp();
      newBallot(t, 'b1');
      setSuperAdmins(t, ROOT.email);
      const res = call(t.app, name, args, who);
      assert.equal(outcome(res), 'allowed', `${who ? who.email : 'anonymous'} ${name}: ${res.error || 'ok'}`);
    });
  });
});

test('not Dispatch-enabled (no key; PROD): claimed ballot stays open to anonymous', () => {
  [{ key: false }, { target: 'PROD' }].forEach((opts) => {
    gatedCalls('b1').forEach(([name, args]) => {
      if (name === 'adminSaveAdmins' || name === 'adminClaimBallot') return; // refused: sign-in off
      const t = loadApp(opts);
      newBallot(t, 'b1');
      t.app._adminWriteConfigUnchecked_('b1', { Admins: ALICE.email });
      assert.equal(outcome(call(t.app, name, args, null)), 'allowed', name);
    });
    const t = loadApp(opts);
    newBallot(t, 'b1');
    assert.match(call(t.app, 'adminSaveAdmins', ['b1', ALICE.email, true], ALICE).error, /not enabled/);
  });
});

test('Make me an admin: signed-in user claims an unclaimed ballot; anonymous cannot', () => {
  const t = loadApp();
  newBallot(t, 'b1');
  assert.match(call(t.app, 'adminClaimBallot', ['b1']).error, /Sign in/);
  assert.deepEqual(call(t.app, 'adminClaimBallot', ['b1'], BOB).result, { ok: true, admins: [BOB.email] });
  assert.match(call(t.app, 'getAdminEditData', ['b1']).error, NOT_AUTH);
  const data = call(t.app, 'getAdminEditData', ['b1'], BOB).result;
  assert.deepEqual(data.admins, [BOB.email]);
  assert.equal(data.access.claimed, true);
  assert.equal(data.access.isBallotAdmin, true);
});

test('lockout guard: removing yourself needs confirm; super-admin needs none; blank list unclaims', () => {
  const t = loadApp();
  newBallot(t, 'b1');
  setSuperAdmins(t, ROOT.email);
  t.app._adminWriteConfigUnchecked_('b1', { Admins: ALICE.email });
  assert.match(call(t.app, 'adminSaveAdmins', ['b1', BOB.email], ALICE).error, /^lockout_confirm:/);
  assert.match(call(t.app, 'adminSaveAdmins', ['b1', 'not-an-email'], ALICE).error, /email address/);
  assert.deepEqual(call(t.app, 'adminSaveAdmins', ['b1', 'Alice@Example.com, ' + BOB.email], ALICE).result.admins, [ALICE.email, BOB.email]);
  assert.deepEqual(call(t.app, 'adminSaveAdmins', ['b1', BOB.email], ROOT).result.admins, [BOB.email]);
  assert.match(call(t.app, 'adminSaveTitle', ['b1', 'x'], ALICE).error, NOT_AUTH);
  assert.equal(outcome(call(t.app, 'adminSaveAdmins', ['b1', ALICE.email, true], ROOT)), 'allowed'); // super-admin recovers
  call(t.app, 'adminSaveAdmins', ['b1', ''], ALICE);
  assert.equal(outcome(call(t.app, 'adminSaveTitle', ['b1', 'x'])), 'allowed');
});

test('admin list: claimed ballot hides notes and respondent names from non-admins', () => {
  const t = loadApp();
  newBallot(t, 'b1');
  t.app._adminWriteConfigUnchecked_('b1', { Admins: ALICE.email, 'Admin-Only-Notes': 'secret notes' });
  const anon = call(t.app, 'getAdminListData_', []).result[0];
  assert.equal(anon.claimed, true);
  assert.equal(anon.canAdmin, false);
  assert.equal(anon.adminNotes, '');
  const mine = call(t.app, 'getAdminListData_', [], ALICE).result[0];
  assert.equal(mine.canAdmin, true);
  assert.equal(mine.adminNotes, 'secret notes');
});

test('no email, name, picture or assertion reaches a log entry', () => {
  const t = loadApp();
  newBallot(t, 'b1');
  const token = mint(ALICE);
  t.app._adminWriteConfigUnchecked_('b1', { Admins: ALICE.email });
  call(t.app, 'submitBallotRanking', ['b1', 'Alice', ['Alpha', 'Beta'], ''], token, { name: 'Alice Example', picture: 'https://lh3.googleusercontent.com/a/x' });
  call(t.app, 'adminSaveTitle', ['b1', 'x'], BOB);
  call(t.app, 'adminSaveAdmins', ['b1', BOB.email], ALICE);
  call(t.app, 'getAdminEditData', ['b1'], mint(ALICE, { exp: NOW - 1 }));
  const text = JSON.stringify(t.logs);
  [ALICE.email, BOB.email, 'Alice Example', 'googleusercontent', token.split('.')[2]].forEach((s) => assert.ok(!text.includes(s), s));
  assert.ok(t.logs.some((l) => l.tag === 'identity.admin.refused' && l.data.reason === 'not_admin'));
});

test('request log scrubbing redacts the envelope and an admins list', () => {
  const { scrubForLog_ } = require('../script/GasLogger.js');
  const out = scrubForLog_({ name: 'x', assertion: 'a.b.c', profile: { name: 'N' } });
  assert.equal(out.assertion, '[redacted]');
  assert.equal(out.profile, '[redacted]');
});

// ---------------------------------------------------------------- static build (placeholders)
test('prod build stamps no sign-in; sit/nuuc stamp their Dispatch audience', () => {
  const { stampSource_ } = require('gas-static/lib/stamp.js');
  const staticPages = require('../tools/static-pages.js');
  const src = fs.readFileSync(path.join(__dirname, '..', 'static-pages', 'src', 'index.html'), 'utf8');
  const token = 'var STATIC_DISPATCH_ = null;';
  assert.ok(src.includes(token), 'placeholder present in source');
  // The page loads GIS only from script (withGis_), never via a static <script src> tag.
  assert.ok(!/<script[^>]+gsi\/client/.test(src));
  const stampFor = (env) => stampSource_(src, { [token]: `var STATIC_DISPATCH_ = ${JSON.stringify(staticPages.dispatchFor_(env))};` });
  assert.ok(stampFor('prod').includes('var STATIC_DISPATCH_ = null;'));
  assert.equal(staticPages.dispatchFor_('prod'), null);
  assert.equal(staticPages.dispatchFor_('sit').aud, 'rcballot-test');
  assert.equal(staticPages.dispatchFor_('nuuc').aud, 'rcballot');
  // The page treats any empty field as sign-in off: the sign-in bars start hidden and
  // renderSignIn_ returns before showing them unless DISPATCH_ has url, clientId and aud.
  assert.ok(/var DISPATCH_ = \(STATIC_DISPATCH_ && STATIC_DISPATCH_\.url && STATIC_DISPATCH_\.clientId && STATIC_DISPATCH_\.aud\)/.test(src));
  assert.ok(/id="ballotSigninBar" class="signin-bar hidden"/.test(src));
  assert.ok(/id="adminSigninBar" class="signin-bar hidden"/.test(src));
});

let failed = 0;
tests.forEach(([name, fn]) => {
  try { fn(); } catch (err) { failed++; console.error('FAIL ' + name + '\n  ' + (err && err.stack || err)); }
});
if (failed) { console.error(`test_identity_auth.js: ${failed} FAILED`); process.exit(1); }
console.log(`test_identity_auth.js: PASS (${tests.length} tests)`);
