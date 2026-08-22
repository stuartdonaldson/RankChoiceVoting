#!/usr/bin/env node
/**
 * RankChoiceVoting web app caller — a thin wrapper over gas-deploy's lib/webapp.js.
 *
 * URL resolution, secret injection, the POST→GET redirect and the non-JSON-response diagnostic
 * all live in the package (RECOMMENDATION.md §3.3). What stays here is this project's own
 * vocabulary: which envs exist, which settings keys hold their secrets, which actions the server
 * answers before its secret gate, and the bootstrapSecret convenience.
 *
 * Usage:
 *   node tools/callWebapp.js <action> [--env sit|prod|nuuc] [--body '{"key":"val"}']
 *
 * Examples:
 *   node tools/callWebapp.js setWebappUrl --env prod
 *   node tools/callWebapp.js bootstrapSecret --env sit
 *   node tools/callWebapp.js getSheet --body '{"sheetName":"Ballot-Test123"}'
 */

'use strict';

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const { webapp } = require('gas-deploy');
const callWebappCli = require('gas-deploy/bin/call-webapp.js');

const ROOT = path.join(__dirname, '..');
const SETTINGS_PATH = path.join(ROOT, 'local.settings.json');

/**
 * Handled by the server BEFORE the admin-secret gate: bootstrapSecret is how a secret is
 * obtained in the first place, and setWebappUrl only stores the running deployment's own exec
 * URL, which a fresh project must be able to do before any secret exists.
 */
const UNGATED_ACTIONS = new Set(['bootstrapSecret', 'setWebappUrl']);

const ENV_MAP = {
  sit:  { deploymentIdKey: 'sitDeploymentId',  secretKey: 'sitAdminSecret',  scriptIdKey: 'sitScriptId'  },
  prod: { deploymentIdKey: 'prodDeploymentId', secretKey: 'prodAdminSecret', scriptIdKey: 'prodScriptId' },
  nuuc: { deploymentIdKey: 'nuucDeploymentId', secretKey: 'nuucAdminSecret', scriptIdKey: 'nuucScriptId', authKey: 'nuucAuth' },
};

const config = {
  root: ROOT,
  envMap: ENV_MAP,
  authField: 'adminSecret',
  ungatedActions: [...UNGATED_ACTIONS],
};

function loadSettings() {
  if (!fs.existsSync(SETTINGS_PATH)) {
    console.error(`❌  local.settings.json not found at ${SETTINGS_PATH}`);
    process.exit(1);
  }
  return JSON.parse(fs.readFileSync(SETTINGS_PATH, 'utf8'));
}

function saveSetting_(key, value) {
  const settings = loadSettings();
  settings[key] = value;
  fs.writeFileSync(SETTINGS_PATH, JSON.stringify(settings, null, 2) + '\n', 'utf8');
}

/** Generates a 32-hex-char secret suitable for ADMIN_SHARED_SECRET. */
function generateSecret_() {
  return crypto.randomBytes(16).toString('hex');
}

async function main() {
  const { action, env, extraBody } = callWebappCli.parseArgs(process.argv);
  const envKey = callWebappCli.normalizeEnv(env, ENV_MAP);

  // Convenience: `bootstrapSecret` with no --body generates a fresh secret, bootstraps it, and
  // records it locally on success. The generated value is never echoed to stdout.
  let generated = null;
  let argv = process.argv;
  if (action === 'bootstrapSecret' && Object.keys(extraBody).length === 0) {
    generated = generateSecret_();
    argv = [...process.argv, '--body', JSON.stringify({ secret: generated })];
  }

  const result = await callWebappCli.run(config, argv);

  if (generated && result && result.ok) {
    saveSetting_(ENV_MAP[envKey].secretKey, generated);
    console.error(`💾 ${ENV_MAP[envKey].secretKey} saved to local.settings.json`);
  }
  return result;
}

if (require.main === module) {
  main().catch(err => {
    console.error('❌', err.message);
    process.exit(1);
  });
}

module.exports = { post: webapp.post, loadSettings, saveSetting_, generateSecret_, ENV_MAP, UNGATED_ACTIONS };
