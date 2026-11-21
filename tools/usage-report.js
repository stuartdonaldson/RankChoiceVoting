#!/usr/bin/env node
/**
 * usage-report.js — Who used the app, doing what, and when, from the Axiom log.
 *
 * Usage:
 *   node tools/usage-report.js [--since D | --hours N | --days N] [--env sit|prod|nuuc]
 *                              [--tz ZONE] [--file PATH]
 *
 *   pnpm run usage                       # all envs, last 24h, grouped by env
 *   pnpm run usage -- --days 3 --env nuuc
 *   node tools/usage-report.js --file dump.json   # summarise a query-axiom.js --raw dump
 *
 * Identity per event: the logged email when there is one, otherwise the entered
 * voter name, otherwise 'admin (no email)' or 'anonymous'. A name ever logged alongside an
 * email is reported under that email. Smoke-test ballots are attributed to 'smoke test'.
 * Events logged before rcballot-1j4 carry masked names ("S..t"); later ones carry names as entered.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const { query_, parseDuration_ } = require('./query-axiom.js');

const ENV_ORDER = ['SIT', 'PROD', 'NUUC'];
const SMOKE_RE = /SmokeTest/i;
const QUERY_LIMIT = 50000;
const COLLAPSE_MS = 5 * 60 * 1000;

// API function name → reported action. Anything starting `admin` is an admin edit.
const ACTIONS = {
  getBallotConfig: 'opened ballot',
  getBallotForName: 'loaded own ballot',
  recordSignin: 'signed in',
  submitBallotRanking: 'submitted',
  addBallotTopic: 'added topic',
  getAdminAnalysisData_: 'viewed results',
  getAdminListData_: 'admin list',
  getAdminEditData: 'admin view',
  createBallotForAdmin_: 'admin edit',
};

function actionFor_(fn) {
  if (ACTIONS[fn]) return ACTIONS[fn];
  if (/^admin/.test(fn)) return 'admin edit';
  return fn;
}

function fmtTime_(iso, tz) {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: tz, month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23',
  }).formatToParts(new Date(iso));
  const p = Object.fromEntries(parts.map(x => [x.type, x.value]));
  return `${p.month}/${p.day} ${p.hour}:${p.minute}:${p.second}`;
}

function isName_(who) {
  return !!who && who !== 'admin';
}

/** Turn one raw Axiom row into a user action, or null when it is not user activity. */
function classify_(m) {
  const d = m.data || {};
  const body = d.body || {};
  const env = d.target || '?';
  const base = { env, time: m._time, email: d.email || null, who: d.who || d.actor || null };

  if (d.name === 'handleAdminPost_.setWebappUrl') return Object.assign(base, { kind: 'deploy' });
  if (d.name === 'doPost' && body.name) {
    const args = Array.isArray(body.args) ? body.args : [];
    const ballot = typeof args[0] === 'string' && args[0] ? args[0] : null;
    return Object.assign(base, { kind: 'action', action: actionFor_(body.name), ballot });
  }
  if (d.name !== 'doPost' && (d.message || d.reason) && d.name !== 'identity.signin.required') {
    return Object.assign(base, { kind: 'problem', source: d.name, message: d.message || d.reason });
  }
  return null; // doPost carriers, version-verify posts, cache refreshes, internal markers
}

function identityOf_(r, nameToEmail) {
  if (r.ballot && SMOKE_RE.test(r.ballot)) return 'smoke test';
  if (r.email) return r.email;
  if (isName_(r.who)) return nameToEmail.get(r.who) || r.who;
  return r.who === 'admin' ? 'admin (no email)' : 'anonymous';
}

/**
 * Pure summary over Axiom matches.
 * @returns {{envs: Array<{env, people, timeline, deploys, problems}>}}
 */
function summarize_(matches, { tz }) {
  const rows = matches.map(classify_).filter(Boolean)
    .sort((a, b) => (a.time < b.time ? -1 : a.time > b.time ? 1 : 0));

  const nameToEmail = new Map();
  rows.forEach(r => { if (r.email && isName_(r.who)) nameToEmail.set(r.who, r.email); });

  const envs = [...new Set(rows.map(r => r.env))]
    .sort((a, b) => {
      const ia = ENV_ORDER.indexOf(a), ib = ENV_ORDER.indexOf(b);
      return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib) || a.localeCompare(b);
    });

  return {
    envs: envs.map(env => {
      const mine = rows.filter(r => r.env === env);
      const people = new Map();
      const timeline = [];

      mine.filter(r => r.kind === 'action').forEach(r => {
        const id = identityOf_(r, nameToEmail);
        const t = fmtTime_(r.time, tz);
        let p = people.get(id);
        if (!p) {
          p = { id, names: new Set(), ballots: new Set(), counts: {}, first: t, last: t };
          people.set(id, p);
        }
        p.last = t;
        if (isName_(r.who) && id !== r.who && id !== 'smoke test') p.names.add(r.who);
        if (r.ballot) p.ballots.add(r.ballot);
        p.counts[r.action] = (p.counts[r.action] || 0) + 1;

        if (id === 'smoke test') {
          // One line per smoke run (its ballot), however many calls it made.
          const run = timeline.find(x => x.id === id && x.ballot === r.ballot);
          if (run) { run.times += 1; run.until = t; }
          else timeline.push({ time: t, until: t, lastMs: 0, id, action: 'smoke test run', ballot: r.ballot, times: 1 });
          return;
        }
        const prev = timeline[timeline.length - 1];
        if (prev && prev.id === id && prev.action === r.action && prev.ballot === r.ballot &&
            Date.parse(r.time) - prev.lastMs <= COLLAPSE_MS) {
          prev.times += 1;
          prev.lastMs = Date.parse(r.time);
          prev.until = t;
        } else {
          timeline.push({ time: t, until: t, lastMs: Date.parse(r.time), id, action: r.action, ballot: r.ballot, times: 1 });
        }
      });

      return {
        env,
        people: [...people.values()].map(p => Object.assign(p, {
          names: [...p.names], ballots: [...p.ballots],
        })),
        timeline: timeline.map(({ lastMs, ...rest }) => rest),
        deploys: mine.filter(r => r.kind === 'deploy').map(r => fmtTime_(r.time, tz)),
        problems: mine.filter(r => r.kind === 'problem').map(r => ({
          time: fmtTime_(r.time, tz), source: r.source,
          message: String(r.message).length > 120 ? String(r.message).slice(0, 117) + '...' : String(r.message),
          id: identityOf_(r, nameToEmail),
        })),
      };
    }),
  };
}

function label_(p) {
  return p.names.length ? `${p.id} (${p.names.join(', ')})` : p.id;
}

function render_(summary, { since, tz, env }) {
  const out = [];
  out.push(`Usage report — last ${since}, ${env || 'all envs'}, times in ${tz}`);
  if (!summary.envs.length) {
    out.push('', 'No activity.');
    return out.join('\n');
  }
  summary.envs.forEach(e => {
    out.push('', `== ${e.env} ==`);
    const people = e.people.slice().sort((a, b) => (a.first < b.first ? -1 : 1));
    out.push('', `People (${people.length})`);
    people.forEach(p => {
      const counts = Object.entries(p.counts).map(([k, v]) => `${k} ${v}`).join(', ');
      const range = p.first === p.last ? p.first : `${p.first} – ${p.last}`;
      out.push(`  ${label_(p)}`);
      out.push(`      ${range}${p.ballots.length ? '  ballots: ' + p.ballots.join(', ') : ''}`);
      out.push(`      ${counts}`);
    });
    out.push('', 'Timeline');
    e.timeline.forEach(t => {
      const when = t.times > 1 && t.until !== t.time ? `${t.time}–${t.until.slice(6)}` : t.time;
      const what = t.action === 'smoke test run' ? `${t.action} (${t.times} calls)`
        : t.times > 1 ? `${t.action} ×${t.times}` : t.action;
      out.push(`  ${when.padEnd(24)} ${t.id.padEnd(28)} ${what}${t.ballot ? '  [' + t.ballot + ']' : ''}`);
    });
    if (e.deploys.length) {
      out.push('', `Deploys (${e.deploys.length}): ${e.deploys.map(d => d.slice(0, 11)).join(', ')}`);
    }
    if (e.problems.length) {
      out.push('', `Problems (${e.problems.length})`);
      e.problems.forEach(p => out.push(`  ${p.time}  ${p.source}: ${p.message}`));
    }
  });
  return out.join('\n');
}

function parseArgs_(argv) {
  const args = argv.slice(2);
  const opts = {
    since: '24h', env: null, file: null, help: false,
    tz: Intl.DateTimeFormat().resolvedOptions().timeZone,
  };
  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a === '--') continue; // pnpm 11 forwards the separator
    else if (a === '--help' || a === '-h') opts.help = true;
    else if (a === '--since') opts.since = args[++i];
    else if (a === '--hours') opts.since = `${parseInt(args[++i], 10)}h`;
    else if (a === '--days') opts.since = `${parseInt(args[++i], 10)}d`;
    else if (a === '--env') opts.env = String(args[++i] || '').toUpperCase();
    else if (a === '--tz') opts.tz = args[++i];
    else if (a === '--file') opts.file = args[++i];
    else throw new Error(`Unknown argument '${a}' (see --help)`);
  }
  if (opts.env && !ENV_ORDER.includes(opts.env)) {
    throw new Error(`--env must be one of ${ENV_ORDER.join(', ').toLowerCase()}`);
  }
  parseDuration_(opts.since); // validates
  return opts;
}

function printHelp_() {
  console.log(`usage-report.js — who used the app, doing what, when (from Axiom).

Usage:
  node tools/usage-report.js [--since D | --hours N | --days N] [--env sit|prod|nuuc]
                             [--tz ZONE] [--file PATH]

Options:
  --since D     Lookback, e.g. 30m, 6h, 3d (default 24h)
  --hours N     Same as --since Nh
  --days N      Same as --since Nd
  --env E       Only this env; omitted = all envs, grouped
  --tz ZONE     IANA time zone for display (default: system zone)
  --file PATH   Summarise a saved 'query-axiom.js --raw' dump instead of querying
  --help, -h    Show this help
`);
}

async function main() {
  const opts = parseArgs_(process.argv);
  if (opts.help) { printHelp_(); return; }

  let matches;
  if (opts.file) {
    matches = JSON.parse(fs.readFileSync(opts.file, 'utf8')).matches || [];
    // A dump may span more than the requested window or envs: apply both locally.
    const cutoff = Date.now() - parseDuration_(opts.since);
    matches = matches.filter(m => Date.parse(m._time) >= cutoff);
  } else {
    const settings = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'local.settings.json'), 'utf8'));
    if (!settings.axiomDataset || !settings.axiomQueryToken) {
      throw new Error('axiomDataset / axiomQueryToken not set in local.settings.json');
    }
    const result = await query_(settings.axiomDataset, settings.axiomQueryToken, {
      // No server-side env filter: identity merging needs name/email pairings from every env.
      since: opts.since, limit: QUERY_LIMIT,
    });
    matches = result.matches || [];
    if (matches.length >= QUERY_LIMIT) {
      console.error(`WARNING: hit the ${QUERY_LIMIT}-event limit; the oldest part of the window is missing.`);
    }
  }
  const summary = summarize_(matches, { tz: opts.tz });
  if (opts.env) summary.envs = summary.envs.filter(e => e.env === opts.env);
  console.log(render_(summary, opts));
}

if (require.main === module) {
  main().catch(err => {
    console.error('ERROR:', err.message);
    process.exit(1);
  });
}

module.exports = { summarize_, render_, parseArgs_, classify_ };
