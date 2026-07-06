#!/usr/bin/env node
/**
 * query-axiom.js — Query the Axiom dataset without re-deriving the APL
 * request shape and auth every time.
 *
 * Ported from F3Go30's tools/query_axiom.py, kept single-language (Node) for
 * this project. Reads axiomDataset + axiomQueryToken from local.settings.json
 * (same settings file every other tool in this repo uses). Never prints the
 * token.
 *
 * Usage:
 *   node tools/query-axiom.js [--limit N] [--since DURATION] [--name SUBSTR]
 *                              [--where APL_EXPR] [--raw [PATH]]
 *
 * Examples:
 *   node tools/query-axiom.js                          # last 200 events, last 24h
 *   node tools/query-axiom.js --limit 50 --since 2h
 *   node tools/query-axiom.js --name handleAdminPost
 *   node tools/query-axiom.js --where "data.action == 'createSurvey'"
 *   node tools/query-axiom.js --raw /tmp/axiom_dump.json
 *
 * DURATION accepts <N>s / <N>m / <N>h / <N>d (e.g. 30m, 2h, 1d). Default: 24h.
 */

'use strict';

const https = require('https');
const fs    = require('fs');
const path  = require('path');

const ROOT          = path.join(__dirname, '..');
const SETTINGS_PATH = path.join(ROOT, 'local.settings.json');
const DURATION_RE    = /^(\d+)([smhd])$/;

function loadSettings_() {
  if (!fs.existsSync(SETTINGS_PATH)) {
    console.error(`ERROR: local.settings.json not found at ${SETTINGS_PATH}`);
    process.exit(1);
  }
  return JSON.parse(fs.readFileSync(SETTINGS_PATH, 'utf8'));
}

function parseDuration_(spec) {
  const m = DURATION_RE.exec(String(spec || '').trim());
  if (!m) throw new Error(`Bad --since value '${spec}', expected e.g. 30m, 2h, 1d`);
  const n = parseInt(m[1], 10);
  const unitMs = { s: 1000, m: 60000, h: 3600000, d: 86400000 }[m[2]];
  return n * unitMs;
}

function buildApl_(dataset, { limit, name, where }) {
  const filters = [];
  if (name) filters.push(`name contains '${name}'`);
  if (where) filters.push(where);
  // `where` must precede `order by`/`limit` in the pipeline, otherwise filtering
  // would apply after the top-N cut and could return fewer than `limit` rows.
  let apl = `['${dataset}']`;
  for (const f of filters) apl += ` | where ${f}`;
  apl += ` | order by _time desc | limit ${limit}`;
  return apl;
}

function isoNoMillis_(date) {
  return date.toISOString().replace(/\.\d+Z$/, 'Z');
}

function query_(dataset, token, opts) {
  const now = new Date();
  const start = new Date(now.getTime() - parseDuration_(opts.since));
  const apl = buildApl_(dataset, opts);
  const body = JSON.stringify({
    apl,
    startTime: isoNoMillis_(start),
    endTime: isoNoMillis_(now),
  });

  return new Promise((resolve, reject) => {
    const req = https.request(
      'https://api.axiom.co/v1/datasets/_apl?format=legacy',
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
          'Content-Length': Buffer.byteLength(body),
        },
      },
      res => {
        const chunks = [];
        res.on('data', c => chunks.push(c));
        res.on('end', () => {
          const text = Buffer.concat(chunks).toString('utf8');
          if (res.statusCode >= 300) {
            return reject(new Error(`Axiom query failed (${res.statusCode}): ${text.slice(0, 500)}`));
          }
          try {
            resolve(JSON.parse(text));
          } catch (e) {
            reject(new Error(`Axiom returned non-JSON response: ${text.slice(0, 200)}`));
          }
        });
      }
    );
    req.setTimeout(20000, () => { req.destroy(); reject(new Error('Request timeout')); });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

/**
 * True for null/''/{} and for objects whose values are all themselves empty — catches
 * fields like a stray DBG:{a:null,b:null} that schema-backfill adds to every row once
 * any event has ever set it.
 */
function isEmpty_(v) {
  if (v === null || v === undefined || v === '') return true;
  if (typeof v === 'object' && !Array.isArray(v)) {
    return Object.values(v).every(isEmpty_);
  }
  if (typeof v === 'object' && Array.isArray(v)) return v.length === 0;
  return false;
}

function printTable_(matches) {
  matches.forEach(m => {
    const data = m.data || {};
    const nonnull = {};
    Object.keys(data).forEach(k => { if (!isEmpty_(data[k])) nonnull[k] = data[k]; });
    const side = nonnull.side || '?'; delete nonnull.side;
    const name = nonnull.name || '?'; delete nonnull.name;
    delete nonnull.version;
    delete nonnull.caller;
    const detail = Object.keys(nonnull).map(k => `${k}=${JSON.stringify(nonnull[k])}`).join(' ');
    console.log(`${m._time}  ${String(side).padEnd(6)} ${String(name).padEnd(32)} ${detail}`);
  });
}

function parseArgs_(argv) {
  const args = argv.slice(2);
  const opts = { limit: 200, since: '24h', name: null, where: null, raw: undefined, help: false };

  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a === '--help' || a === '-h') { opts.help = true; }
    else if (a === '--limit' || a === '-n') { opts.limit = parseInt(args[++i], 10); }
    else if (a === '--since') { opts.since = args[++i]; }
    else if (a === '--name') { opts.name = args[++i]; }
    else if (a === '--where') { opts.where = args[++i]; }
    else if (a === '--raw') {
      // Optional path argument — only consume the next token if it isn't itself a flag.
      if (args[i + 1] && !args[i + 1].startsWith('--')) { opts.raw = args[++i]; }
      else { opts.raw = '-'; }
    }
  }
  return opts;
}

function printHelp_() {
  console.log(`query-axiom.js — Query the Axiom dataset for this project.

Usage:
  node tools/query-axiom.js [--limit N] [--since DURATION] [--name SUBSTR]
                             [--where APL_EXPR] [--raw [PATH]]

Options:
  --limit, -n N   Max events to return (default: 200)
  --since D       How far back to look, e.g. 30m, 2h, 1d (default: 24h)
  --name SUBSTR   Filter to event names containing this substring
  --where EXPR    Raw APL 'where' expression, e.g. "data.action == 'createSurvey'"
  --raw [PATH]    Dump full JSON response (to PATH, or stdout if no PATH given)
  --help, -h      Show this help
`);
}

async function main() {
  const opts = parseArgs_(process.argv);
  if (opts.help) { printHelp_(); return; }

  const settings = loadSettings_();
  const dataset = settings.axiomDataset;
  const token = settings.axiomQueryToken;
  if (!dataset || !token) {
    console.error('ERROR: axiomDataset / axiomQueryToken not set in local.settings.json');
    process.exitCode = 1;
    return;
  }

  const result = await query_(dataset, token, opts);
  const matches = result.matches || [];

  if (opts.raw !== undefined) {
    const text = JSON.stringify(result, null, 2);
    if (opts.raw === '-') {
      console.log(text);
    } else {
      fs.writeFileSync(opts.raw, text, 'utf8');
      console.error(`Wrote ${matches.length} events to ${opts.raw}`);
    }
    return;
  }

  console.log(`${matches.length} events, ${opts.since} lookback, dataset=${dataset}`);
  printTable_(matches);
}

if (require.main === module) {
  main().catch(err => {
    console.error('ERROR:', err.message);
    process.exit(1);
  });
}

module.exports = { parseArgs_, parseDuration_, buildApl_, isEmpty_, query_ };
