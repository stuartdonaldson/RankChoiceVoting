#!/usr/bin/env node
/**
 * RankChoiceVoting deployment config. The pipeline itself — auth, stamping, push,
 * named-deployment update, hook sequencing, deploy verification, summary — lives in the shared
 * `gas-deploy` package; this file is only what is specific to this project.
 *
 *   pnpm run deploy:sit | deploy:prod | deploy:nuuc
 *   node tools/manage-deployments.js --summary --env sit    # read-only, deploys nothing
 *
 * See the gas-deploy README and best-practices/gas-deployment/RECOMMENDATION.md for internals.
 */

'use strict';

const path = require('path');
const { execSync } = require('child_process');
const { runCli, constStamper } = require('gas-deploy');

const ROOT = path.join(__dirname, '..');

// Static entry point per deploy label. Must stay in sync with script/ApiBridge.js's
// _staticPagesBaseUrl_ — the GAS-side copy, read at runtime from APP_DEPLOY_TARGET.
const STATIC_ENTRY_BASE_URL = {
  SIT:  'https://f3go30.github.io/static-pages/ballot/sit/',
  PROD: 'https://f3go30.github.io/static-pages/ballot/prod/',
  NUUC: 'https://nuuc-it.github.io/Static/pub/ballot/',
};

const call = (args) => execSync(`node tools/callWebapp.js ${args}`, { stdio: 'inherit', cwd: ROOT });

const config = {
  root: ROOT,
  stamper: constStamper({ file: 'script/version.js' }),
  describeDeployment: (version) => `v${version} RCV`,

  // NUUC runs under a separate Google account, so it names its own authKey; sit/prod take the
  // package's default. Each target is its own script project, hence its own deploymentIdKey.
  targets: {
    sit:  { scriptIdKey: 'sitScriptId',  label: 'SIT',  emoji: '🧪', counter: 'build',   deploymentIdKey: 'sitDeploymentId',  sheetIdKey: 'sitSheetId'  },
    prod: { scriptIdKey: 'prodScriptId', label: 'PROD', emoji: '🚀', counter: 'version', deploymentIdKey: 'prodDeploymentId', sheetIdKey: 'prodSheetId' },
    nuuc: { scriptIdKey: 'nuucScriptId', label: 'NUUC', emoji: '⛪', counter: 'version', deploymentIdKey: 'nuucDeploymentId', sheetIdKey: 'nuucSheetId', authKey: 'nuucAuth' },
  },

  // --summary only: reads the stamped file so a live-vs-local divergence can be flagged (someone
  // deployed from elsewhere, or a deploy half-failed). Display-only — the deploy path never
  // reads this back, package.json stays the source of truth.
  readLocalVersion: () => {
    const src = require('fs').readFileSync(path.join(ROOT, 'script/version.js'), 'utf8');
    const grab = (name) => (src.match(new RegExp(`const ${name}\\s*=\\s*'([^']+)'`)) || [])[1];
    return { version: grab('APP_VERSION'), now: grab('APP_VERSION_DATE') };
  },

  extraRows: ({ label }) => [
    { label: 'Static page', value: STATIC_ENTRY_BASE_URL[label], missing: '(static hosting not configured for this target)' },
  ],

  postDeploy: [
    // Every deploy, not just PROD: each env is its own script project with its own WEBAPP_URL
    // property, and onOpen.js's menu reads it. Without it _getWebAppUrl_() falls back to
    // ScriptApp.getService().getUrl(), which only resolves inside a live web app request.
    { name: 'Stamp WEBAPP_URL', run: ({ targetKey }) => call(`setWebappUrl --env ${targetKey}`),
      retryCommand: 'node tools/callWebapp.js setWebappUrl --env <env>' },

    // Bootstrapped before this deployment's URL is ever shared — bootstrapSecret is reachable by
    // anyone on an ANYONE_ANONYMOUS deployment until a secret is set. Re-running on a
    // bootstrapped target exits non-zero ("already_bootstrapped"), which is why it is not required.
    { name: 'Bootstrap ADMIN_SHARED_SECRET (no-op if already set)',
      run: ({ targetKey }) => call(`bootstrapSecret --env ${targetKey}`),
      retryCommand: 'node tools/callWebapp.js bootstrapSecret --env <env>' },

    // The static front end shares this deploy's version/build counter, so it ships in the same
    // deploy and the two can never drift. --skip-bump: the pipeline already bumped. Required —
    // a stale static bundle against new GAS code is a broken app, not a warning.
    { name: 'Publish static pages', required: true,
      run: ({ targetKey }) => execSync(`node ${path.join(__dirname, 'publish-static-pages.js')} --env ${targetKey} --skip-bump`,
        { stdio: 'inherit', cwd: ROOT }) },
  ],
};

if (require.main === module) {
  runCli(config).catch(err => {
    if (err && (err.name === 'ExitPromptError' || (err.message || '').includes('force closed'))) return console.log('\n❌ Cancelled.');
    console.error('❌ Error:', err.message);
    process.exit(1);
  });
}

module.exports = { config, STATIC_ENTRY_BASE_URL };
