#!/usr/bin/env node
/**
 * Read-only live comparison: old ballot readers vs the snapshot reader (rcballot-8i5.2).
 *
 * For each env, lists the Ballot-* sheets (admin listSheets), fetches each (admin getSheet),
 * loads it into the in-memory fake Sheets API and diffs the old readers' view against
 * script/BallotModel.js's snapshot view. Only the two read actions are ever sent; nothing is
 * written to a live sheet. The old readers may "repair" the FAKE sheet -- that is reported as
 * mutatedByOldReaders (the live sheet is damaged / would be repaired on read).
 *
 * Since rcballot-8i5.3 the readers in script/ ARE the snapshot, so the "old" side is loaded
 * from git (--old-rev, default c9f060a: the last commit with the positional readers), and the
 * new side is the current code AFTER the admin-load repair (_repairBallotSheet_) has run on
 * the fake sheet — i.e. what the respondent page / analysis will show once the admin edit
 * page has been opened. The repair steps that fired are listed per sheet.
 *
 * Usage: node tools/compare-ballot-readers.js [--env sit|prod|nuuc|all] [--old-rev <git rev>]
 */
'use strict';

const vm = require('vm');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const { createFakeSpreadsheet, createFakeGasGlobals } = require(path.join(ROOT, 'test', 'fakeGas'));
const ENVS = ['sit', 'prod', 'nuuc'];

const readWorking = (f) => fs.readFileSync(path.join(ROOT, 'script', f), 'utf8');
const readAtRev = (rev) => (f) => require('child_process').execFileSync('git', ['show', rev + ':script/' + f], { cwd: ROOT, encoding: 'utf8' });

/** Same loader shape as test/test_ballot_model.js loadBallotModel (kept self-contained). */
function loadBallotModel(read) {
  read = read || readWorking;
  const ss = createFakeSpreadsheet();
  const sandbox = Object.assign({ module: { exports: {} } }, createFakeGasGlobals(ss));
  vm.createContext(sandbox);
  vm.runInContext(read('BallotCache.js'), sandbox, { filename: 'BallotCache.js' });
  const cacheExports = sandbox.module.exports;
  sandbox.module = { exports: {} };
  vm.runInContext(read('BallotModel.js'), sandbox, { filename: 'BallotModel.js' });
  return { SM: Object.assign({}, cacheExports, sandbox.module.exports), ss, sandbox };
}

const plain = (x) => JSON.parse(JSON.stringify(x === undefined ? null : x));
const gridOf = (sheet) => plain(sheet.getDataRange().getValues());

function oldView(SM, sheet) {
  const names = SM.readBallotCandidates_(sheet);
  return {
    config: SM.readBallotConfig_(sheet),
    details: SM.readBallotCandidateDetails_(sheet),
    candidates: names,
    responses: SM.readBallotResponseRows_(sheet).map((r) => {
      const ranks = {};
      names.forEach((n, i) => { ranks[n] = r.ranks[i]; });
      return { row: r.row, date: r.date, name: r.name, weight: r.weight, comment: r.comment, ranks };
    }),
  };
}

function snapView(SM, snap) {
  const config = {};
  Object.keys(snap.config.byKey).forEach((k) => { config[k] = snap.config.byKey[k].value; });
  const cands = SM._snapCandidates_(snap);
  return {
    config,
    details: cands.map((c) => ({ name: c.name, details: c.details, linkText: c.linkText, linkUrl: c.linkUrl })),
    candidates: cands.map((c) => c.name),
    responses: SM._snapResponses_(snap).map((r) => ({ row: r.row, date: r.date, name: r.name, weight: r.weight, comment: r.comment, ranks: r.ranks })),
  };
}

const J = (x) => JSON.stringify(x);
const short = (x) => { const s = J(x); return s === undefined ? 'undefined' : (s.length > 90 ? s.slice(0, 87) + '...' : s); };

function diffViews(a, b) {
  const out = [];
  // config: key -> value (compared as strings; the CSV transport stringifies cells)
  const keys = new Set(Object.keys(a.config).concat(Object.keys(b.config)));
  keys.forEach((k) => { if (J(a.config[k]) !== J(b.config[k])) out.push(`config.${k}: old=${short(a.config[k])} new=${short(b.config[k])}`); });
  if (J(a.candidates) !== J(b.candidates)) out.push(`candidates: old=${short(a.candidates)} new=${short(b.candidates)}`);
  const dn = new Set(a.details.concat(b.details).map((d) => d.name));
  dn.forEach((n) => {
    const x = a.details.find((d) => d.name === n), y = b.details.find((d) => d.name === n);
    if (J(x) !== J(y)) out.push(`details[${n}]: old=${short(x)} new=${short(y)}`);
  });
  if (a.responses.length !== b.responses.length) out.push(`responses.count: old=${a.responses.length} new=${b.responses.length}`);
  // Keyed by respondent name (+ occurrence), not row: a repair may insert rows above Responses.
  const byRow = (v) => { const m = {}; v.responses.forEach((r) => { let k = r.name; while (m[k]) k += "'"; m[k] = r; }); return m; };
  const ar = byRow(a), br = byRow(b);
  Object.keys(ar).forEach((row) => {
    const x = ar[row], y = br[row];
    if (!y) { out.push(`response ${row}: missing in new`); return; }
    ['date', 'name', 'weight', 'comment'].forEach((f) => { if (J(x[f]) !== J(y[f])) out.push(`row ${row} ${f}: old=${short(x[f])} new=${short(y[f])}`); });
    const rk = new Set(Object.keys(x.ranks).concat(Object.keys(y.ranks)));
    rk.forEach((n) => { if (J(x.ranks[n]) !== J(y.ranks[n])) out.push(`row ${row} (${x.name}) rank[${n}]: old=${short(x.ranks[n])} new=${short(y.ranks[n])}`); });
  });
  Object.keys(br).forEach((row) => { if (!ar[row]) out.push(`response ${row}: missing in old`); });
  return out;
}

function fakeSheet(ss, name, rows) {
  const sheet = ss.insertSheet(name);
  rows.forEach((r, i) => r.forEach((v, j) => { if (v !== '' && v != null) sheet.getRange(i + 1, j + 1).setValue(v); }));
  return sheet;
}

function compareSheet(name, rows, oldRev) {
  // Old side: the positional readers at oldRev, on their own fake copy.
  const O = loadBallotModel(readAtRev(oldRev));
  const oldSheet = fakeSheet(O.ss, name, rows);
  const before = gridOf(oldSheet);
  const res = { name };
  let old;
  try { old = oldView(O.SM, oldSheet); } catch (e) { res.oldError = e.message; }
  res.mutated = J(gridOf(oldSheet)) !== J(before);
  if (res.mutated) res.mutatedNote = `rows ${before.length}->${gridOf(oldSheet).length}`;

  // New side: current code, after the admin-load repair.
  const { SM, ss, sandbox } = loadBallotModel();
  const sheet = fakeSheet(ss, name, rows);
  SM._repairBallotSheet_(sheet);
  res.repairs = sandbox.Logger._lines.map((l) => String(l).split(' ')[0]).filter((t) => t.indexOf('ballot.repair.') === 0);
  const snap = SM._readBallot_(sheet);
  res.problems = snap.problems.map((p) => p.code + (p.label ? `(${p.label})` : ''));
  res.problemsRaw = plain(snap.problems);
  res.orphans = plain(SM._snapOrphanColumns_(snap));
  res.missing = plain(SM._snapMissingLabels_(snap));
  res.nCands = SM._snapCandidates_(snap).length;
  res.nVotes = SM._snapResponses_(snap).length;
  if (old) res.diffs = diffViews(plain(old), plain(snapView(SM, snap)));
  return res;
}

async function main() {
  const i = process.argv.indexOf('--env');
  const sel = i > -1 ? process.argv[i + 1] : 'all';
  const r = process.argv.indexOf('--old-rev');
  const oldRev = r > -1 ? process.argv[r + 1] : 'c9f060a';
  if (sel !== 'all' && !ENVS.includes(sel)) { console.error('usage: --env sit|prod|nuuc|all'); process.exit(2); }
  const { post } = require(path.join(ROOT, 'tools', 'callWebapp.js'));
  const s = require(path.join(ROOT, 'local.settings.json'));
  let total = 0, equal = 0, mutatedN = 0;
  for (const env of sel === 'all' ? ENVS : [sel]) {
    const url = `https://script.google.com/macros/s/${s[env + 'DeploymentId']}/exec?cmd=admin`;
    const L = await post(url, { action: 'listSheets', adminSecret: s[env + 'AdminSecret'] });
    if (!L || !L.ok) { console.log(`## ${env}: listSheets failed: ${L && L.error}`); continue; }
    const sheets = L.sheets.filter((x) => x.name.startsWith('Ballot-'));
    console.log(`## ${env}: ${sheets.length} Ballot-* sheets`);
    for (const sh of sheets) {
      const r = await post(url, { action: 'getSheet', adminSecret: s[env + 'AdminSecret'], sheetName: sh.name });
      if (!r || !r.ok) { console.log(`${env} ${sh.name}: getSheet failed: ${r && r.error}`); continue; }
      const rows = r.csv.split('\n').map((l) => l.split('\t'));
      let res;
      try { res = compareSheet(sh.name, rows, oldRev); } catch (e) { console.log(`${env} ${sh.name}: ERROR ${e.message}`); continue; }
      total++;
      const eq = !res.oldError && res.diffs.length === 0;
      if (eq) equal++;
      if (res.mutated) mutatedN++;
      console.log(`${env} ${sh.name}: ${eq ? 'EQUAL' : 'DIFF'}  cands=${res.nCands} votes=${res.nVotes}  mutatedByOldReaders: ${res.mutated ? 'yes (' + res.mutatedNote + ')' : 'no'}  repairs: ${res.repairs.length ? res.repairs.join(',') : 'none'}`);
      if (res.oldError) console.log(`    old readers threw: ${res.oldError}`);
      res.diffs && res.diffs.slice(0, 12).forEach((d) => console.log('    - ' + d));
      if (res.diffs && res.diffs.length > 12) console.log(`    ... ${res.diffs.length - 12} more`);
      if (res.problems.length) console.log('    problems: ' + J(res.problemsRaw));
      if (res.orphans.length) console.log('    orphans: ' + J(res.orphans));
      if (res.missing.candidates && res.missing.candidates.length || res.missing.responses && res.missing.responses.length) console.log('    missingLabels: ' + J(res.missing));
    }
  }
  console.log(`\nSUMMARY: ${total} sheets, ${equal} EQUAL, ${total - equal} DIFF, ${mutatedN} mutatedByOldReaders`);
}

main().catch((e) => { console.error(e.message); process.exit(1); });
