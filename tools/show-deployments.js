#!/usr/bin/env node
/**
 * Every env's URLs and the version each half is actually serving, in one read-only pass.
 *
 *   pnpm run deployments
 *
 * Unlike `manage-deployments.js --summary --env <env>`, this never shells out to clasp or rewrites
 * .clasp.json: the webapp URL comes from the `*DeploymentId` cache the deploy writes back to
 * local.settings.json, and the versions come from the two routes that report them with no secret —
 * the backend's `?cmd=version` and the static page's `version.json`.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const { queryLiveVersion, webapp } = require('gas-deploy');
const { config } = require('./manage-deployments');
const { liveUrl_ } = require('./static-pages');

const ROOT = path.join(__dirname, '..');

/**
 * One record per env in gas-project.json order. Never throws for a single env's failure: an
 * unreachable backend or page is recorded as `error`, so the other envs still print.
 */
async function collect_({ project, settings, targets, liveUrl, fetchBackend, fetchStatic }) {
  return Promise.all(Object.entries(project.envs).map(async ([env, def]) => {
    const target = targets[env] || { label: env.toUpperCase(), deploymentIdKey: `${env}DeploymentId` };
    const deploymentId = settings[target.deploymentIdKey] || null;
    const pageUrl = liveUrl(env);

    const [backend, page] = await Promise.all([
      deploymentId
        ? fetchBackend(deploymentId).then(r => r || { error: 'cmd=version did not answer' })
        : null,
      fetchStatic(`${pageUrl}version.json`).then(
        r => (r && r.version ? r : { error: 'version.json missing or malformed' }),
        err => ({ error: err.message }),
      ),
    ]);

    const warnings = [];
    if (backend && backend.target && backend.target !== target.label) {
      warnings.push(`backend reports target ${backend.target}, expected ${target.label}`);
    }
    if (page && page.env && page.env !== env) {
      warnings.push(`static page reports env ${page.env}, expected ${env}`);
    }
    if (backend && backend.version && page && page.version && backend.version !== page.version) {
      warnings.push(`backend v${backend.version} but static page v${page.version}`);
    }

    return {
      env, label: target.label, emoji: target.emoji || '📦', deploymentIdKey: target.deploymentIdKey,
      spreadsheet: def.sheetId ? `https://docs.google.com/spreadsheets/d/${def.sheetId}/edit` : null,
      adminPage: `${pageUrl}?cmd=admin`,
      staticPage: pageUrl,
      webapp: deploymentId ? webapp.execUrl(deploymentId) : null,
      backend, page, warnings,
    };
  }));
}

function line_(label, value) {
  return `   ${(label + ':').padEnd(14)}${value}`;
}

function version_(r) {
  if (!r) return null;
  if (r.error) return `(unreachable: ${r.error})`;
  return `v${r.version} (${r.target || r.env})`;
}

function render_(records) {
  const out = [];
  for (const r of records) {
    out.push(`\n${r.emoji}  ${r.label}`);
    out.push(line_('Spreadsheet', r.spreadsheet || '(sheetId not set in gas-project.json)'));
    out.push(line_('Admin page', r.adminPage));
    out.push(line_('Webapp', r.webapp || `(${r.deploymentIdKey} not in local.settings.json — deploy once to cache it)`));
    out.push(line_('Backend', version_(r.backend) || '(no deployment ID)'));
    out.push(line_('Static page', version_(r.page)));
    for (const w of r.warnings) out.push(`   ⚠️  ${w}`);
  }
  out.push('');
  return out.join('\n');
}

async function main() {
  const project = JSON.parse(fs.readFileSync(path.join(ROOT, 'gas-project.json'), 'utf8'));
  const settings = JSON.parse(fs.readFileSync(path.join(ROOT, 'local.settings.json'), 'utf8'));
  const records = await collect_({
    project, settings, targets: config.targets, liveUrl: liveUrl_,
    fetchBackend: (id) => queryLiveVersion(id),
    fetchStatic: (url) => webapp.get(`${url}?cachebust=${Date.now()}`, { timeoutMs: 30000 }),
  });
  console.log(render_(records));
}

if (require.main === module) {
  main().catch((err) => { console.error('❌ Error:', err.message); process.exit(1); });
}

module.exports = { collect_, render_ };
