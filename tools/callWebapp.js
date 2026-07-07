#!/usr/bin/env node
/**
 * RankChoiceVoting web app caller — general `?cmd=admin` API client.
 *
 * Adapted from F3Go30's tools/callWebapp.js, trimmed to this project's single
 * cmd=admin endpoint (no --cmd switching — this project only has an admin
 * doPost). Used by tools/manage-deployments.js right after a PROD deploy to
 * stamp the WEBAPP_URL script property with the freshly-deployed exec URL
 * (onOpen.js's "About" dialog reads it), and by tools/smokeTest.js to drive
 * integration tests against a live SIT deployment.
 *
 * Usage:
 *   node tools/callWebapp.js <action> [--env sit|prod|nuuc] [--body '{"key":"val"}']
 *
 * The admin secret (sitAdminSecret/prodAdminSecret/nuucAdminSecret in
 * local.settings.json) is injected into the POST body automatically for every
 * action EXCEPT bootstrapSecret and setWebappUrl, which are ungated on the
 * server side.
 *
 * Examples:
 *   node tools/callWebapp.js setWebappUrl --env prod
 *   node tools/callWebapp.js bootstrapSecret --env sit
 *   node tools/callWebapp.js listSheets --env sit
 *   node tools/callWebapp.js getSheet --body '{"sheetName":"Survey-Test123"}'
 *   node tools/callWebapp.js createSurvey --body '{"id":"SmokeTest1"}'
 */

'use strict';

const https   = require('https');
const crypto  = require('crypto');
const fs      = require('fs');
const path    = require('path');

const ROOT          = path.join(__dirname, '..');
const SETTINGS_PATH = path.join(ROOT, 'local.settings.json');

// Actions the server handles BEFORE the admin-secret gate — never send a
// secret we may not have yet (bootstrapSecret is how we obtain one), and
// setWebappUrl is intentionally ungated so a fresh deploy can stamp its own URL.
const UNGATED_ACTIONS = new Set(['bootstrapSecret', 'setWebappUrl']);

const ENV_MAP = {
  sit:  { deploymentIdKey: 'sitDeploymentId',  adminSecretKey: 'sitAdminSecret'  },
  prod: { deploymentIdKey: 'prodDeploymentId', adminSecretKey: 'prodAdminSecret' },
  nuuc: { deploymentIdKey: 'nuucDeploymentId', adminSecretKey: 'nuucAdminSecret' },
};

// Flags that consume the following argv slot as their value — used to skip both when
// scanning for the action token, so `--env sit` before the action doesn't leave "sit"
// mistaken for it.
const VALUE_FLAGS = new Set(['--env', '--body']);

function parseArgs_(argv) {
  const args = argv.slice(2);
  let action;
  for (let i = 0; i < args.length; i++) {
    if (VALUE_FLAGS.has(args[i])) {
      i++;
      continue;
    }
    if (!args[i].startsWith('--')) {
      action = args[i];
      break;
    }
  }
  if (!action) {
    console.error('Usage: callWebapp.js <action> [--env sit|prod|nuuc] [--body \'{"key":"val"}\']');
    process.exit(1);
  }

  const envIdx = args.indexOf('--env');
  const env = envIdx !== -1 ? args[envIdx + 1] : 'sit';
  if (!ENV_MAP[env]) {
    console.error(`❌  Unknown env "${env}". Use sit, prod, or nuuc.`);
    process.exit(1);
  }

  const bodyIdx = args.indexOf('--body');
  let extraBody = {};
  if (bodyIdx !== -1) {
    try {
      extraBody = JSON.parse(args[bodyIdx + 1]);
    } catch {
      console.error('❌  --body must be valid JSON.');
      process.exit(1);
    }
  }

  return { action, env, extraBody };
}

function buildPayload_(action, extraBody, adminSecret) {
  if (UNGATED_ACTIONS.has(action)) {
    return { action, ...extraBody };
  }
  return { action, adminSecret, ...extraBody };
}

function loadSettings() {
  if (!fs.existsSync(SETTINGS_PATH)) {
    console.error(`❌  local.settings.json not found at ${SETTINGS_PATH}`);
    process.exit(1);
  }
  return JSON.parse(fs.readFileSync(SETTINGS_PATH, 'utf8'));
}

function saveSetting_(key, value) {
  const settings = JSON.parse(fs.readFileSync(SETTINGS_PATH, 'utf8'));
  settings[key] = value;
  fs.writeFileSync(SETTINGS_PATH, JSON.stringify(settings, null, 2) + '\n', 'utf8');
}

/** Generates a >=24-char hex secret suitable for ADMIN_SHARED_SECRET. */
function generateSecret_() {
  return crypto.randomBytes(16).toString('hex'); // 32 hex chars
}

// POST to the GAS web app. GAS responds with a 302 redirect to a GET-only
// echo endpoint — follow as GET, never pin the method through the redirect.
function post(url, body) {
  return new Promise((resolve, reject) => {
    const bodyStr = JSON.stringify(body);
    const parsed  = new URL(url);
    const req = https.request(
      {
        hostname: parsed.hostname,
        path:     parsed.pathname + parsed.search,
        method:   'POST',
        headers:  {
          'Content-Type':   'text/plain',
          'Content-Length': Buffer.byteLength(bodyStr),
        },
      },
      res => {
        if (res.statusCode === 301 || res.statusCode === 302) {
          res.resume();
          return get(res.headers['location']).then(resolve, reject);
        }
        collectBody(res).then(resolve, reject);
      }
    );
    req.setTimeout(60000, () => { req.destroy(); reject(new Error('Request timeout')); });
    req.on('error', reject);
    req.write(bodyStr);
    req.end();
  });
}

function get(url) {
  return new Promise((resolve, reject) => {
    const req = https.get(url, res => {
      if (res.statusCode === 301 || res.statusCode === 302) {
        res.resume();
        return get(res.headers['location']).then(resolve, reject);
      }
      collectBody(res).then(resolve, reject);
    });
    req.setTimeout(60000, () => { req.destroy(); reject(new Error('Request timeout')); });
    req.on('error', reject);
  });
}

function collectBody(res) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    res.on('data', c => chunks.push(c));
    res.on('end', () => {
      const text = Buffer.concat(chunks).toString('utf8');
      try { resolve(JSON.parse(text)); } catch { resolve(text); }
    });
    res.on('error', reject);
  });
}

async function main() {
  const { action, env, extraBody } = parseArgs_(process.argv);
  const settings = loadSettings();
  const { deploymentIdKey, adminSecretKey } = ENV_MAP[env];
  const deploymentId = settings[deploymentIdKey];

  if (!deploymentId || deploymentId.startsWith('<')) {
    console.error(`❌  ${deploymentIdKey} is not set in local.settings.json.`);
    console.error('    Run the deploy script for this environment first.');
    process.exit(1);
  }

  const url = `https://script.google.com/macros/s/${deploymentId}/exec?cmd=admin`;

  // Convenience: `bootstrapSecret` with no --body generates a fresh secret,
  // bootstraps it on the server, and saves it locally on success.
  let body = extraBody;
  let generatedSecret = null;
  if (action === 'bootstrapSecret' && Object.keys(extraBody).length === 0) {
    generatedSecret = generateSecret_();
    body = { secret: generatedSecret };
  }

  const adminSecret = settings[adminSecretKey];
  const payload = buildPayload_(action, body, adminSecret);

  console.error(`→ ${env.toUpperCase()}  ${action}`);

  const result = await post(url, payload);
  console.log(JSON.stringify(result, null, 2));

  if (action === 'bootstrapSecret' && generatedSecret && result && result.ok) {
    saveSetting_(adminSecretKey, generatedSecret);
    console.error(`💾 ${adminSecretKey} saved to local.settings.json`);
  }

  if (result && result.ok === false) process.exit(1);
}

if (require.main === module) {
  main().catch(err => {
    console.error('❌', err.message);
    process.exit(1);
  });
}

module.exports = { parseArgs_, buildPayload_, post, loadSettings, saveSetting_, generateSecret_, ENV_MAP, UNGATED_ACTIONS };
