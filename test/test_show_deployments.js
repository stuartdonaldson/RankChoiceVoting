'use strict';

const assert = require('node:assert/strict');
const { collect_, render_ } = require('../tools/show-deployments.js');

// rcballot-q5r: `pnpm run deployments` — one read-only listing of every env's URLs and the
// version each half (GAS backend, static page) is actually serving. Fetchers are injected so
// this suite never touches the network.

const project = {
  envs: {
    sit:  { scriptId: 'S-SIT',  sheetId: 'SHEET-SIT' },
    prod: { scriptId: 'S-PROD', sheetId: 'SHEET-PROD' },
    nuuc: { scriptId: 'S-NUUC', sheetId: 'SHEET-NUUC' },
  },
};
const targets = {
  sit:  { label: 'SIT',  emoji: '🧪', deploymentIdKey: 'sitDeploymentId' },
  prod: { label: 'PROD', emoji: '🚀', deploymentIdKey: 'prodDeploymentId' },
  nuuc: { label: 'NUUC', emoji: '⛪', deploymentIdKey: 'nuucDeploymentId' },
};
const liveUrl = (env) => `https://pages.example/${env}/`;

async function run_(settings, backend, page) {
  const records = await collect_({
    project, settings, targets, liveUrl,
    fetchBackend: async (id) => backend[id] === undefined ? null : backend[id],
    fetchStatic: async (url) => {
      if (page[url] instanceof Error) throw page[url];
      return page[url];
    },
  });
  return { records, text: render_(records) };
}

(async () => {
  // AC 1, 2 — every env, in gas-project.json order, with every URL and both versions.
  {
    const { records, text } = await run_(
      { sitDeploymentId: 'D-SIT', prodDeploymentId: 'D-PROD', nuucDeploymentId: 'D-NUUC' },
      {
        'D-SIT':  { version: '0.1.7.3', target: 'SIT' },
        'D-PROD': { version: '0.1.7',   target: 'PROD' },
        'D-NUUC': { version: '0.1.8',   target: 'NUUC' },
      },
      {
        'https://pages.example/sit/version.json':  { version: '0.1.7.3', env: 'sit' },
        'https://pages.example/prod/version.json': { version: '0.1.7',   env: 'prod' },
        'https://pages.example/nuuc/version.json': { version: '0.1.8',   env: 'nuuc' },
      },
    );
    assert.deepEqual(records.map(r => r.env), ['sit', 'prod', 'nuuc']);
    const nuuc = records[2];
    assert.equal(nuuc.spreadsheet, 'https://docs.google.com/spreadsheets/d/SHEET-NUUC/edit');
    assert.equal(nuuc.adminPage, 'https://pages.example/nuuc/?cmd=admin');
    assert.equal(nuuc.webapp, 'https://script.google.com/macros/s/D-NUUC/exec');
    assert.deepEqual(nuuc.warnings, []);
    assert.ok(text.indexOf('SIT') < text.indexOf('PROD') && text.indexOf('PROD') < text.indexOf('NUUC'));
    assert.match(text, /Spreadsheet:\s+https:\/\/docs\.google\.com\/spreadsheets\/d\/SHEET-PROD\/edit/);
    assert.match(text, /Admin page:\s+https:\/\/pages\.example\/prod\/\?cmd=admin/);
    assert.match(text, /Backend:\s+v0\.1\.8 \(NUUC\)/);
    assert.match(text, /Static page:\s+v0\.1\.7\.3 \(sit\)/);
    assert.doesNotMatch(text, /⚠️/);
  }

  // AC 3 — missing deploymentId, unreachable backend, unreachable page: explained, never thrown,
  // and the other envs still print.
  {
    const { records, text } = await run_(
      { prodDeploymentId: 'D-PROD', nuucDeploymentId: 'D-NUUC' },
      { 'D-NUUC': { version: '0.1.8', target: 'NUUC' } },
      {
        'https://pages.example/prod/version.json': new Error('Request timeout'),
        'https://pages.example/nuuc/version.json': { version: '0.1.8', env: 'nuuc' },
      },
    );
    assert.equal(records.length, 3);
    assert.equal(records[0].webapp, null);
    assert.match(text, /Webapp:\s+\(sitDeploymentId not in local\.settings\.json — deploy once to cache it\)/);
    assert.match(text, /Backend:\s+\(unreachable: cmd=version did not answer\)/);
    assert.match(text, /Static page:\s+\(unreachable: Request timeout\)/);
    assert.match(text, /Backend:\s+v0\.1\.8 \(NUUC\)/);
  }

  // AC 4 — a backend or page serving another env's build is flagged, as is backend/page skew.
  {
    const { records } = await run_(
      { sitDeploymentId: 'D-SIT', prodDeploymentId: 'D-PROD', nuucDeploymentId: 'D-NUUC' },
      {
        'D-SIT':  { version: '0.1.7.3', target: 'PROD' },
        'D-PROD': { version: '0.1.7',   target: 'PROD' },
        'D-NUUC': { version: '0.1.8',   target: 'NUUC' },
      },
      {
        'https://pages.example/sit/version.json':  { version: '0.1.7.3', env: 'sit' },
        'https://pages.example/prod/version.json': { version: '0.1.6',   env: 'prod' },
        'https://pages.example/nuuc/version.json': { version: '0.1.8',   env: 'sit' },
      },
    );
    assert.match(records[0].warnings.join('\n'), /backend reports target PROD/);
    assert.match(records[1].warnings.join('\n'), /backend v0\.1\.7 but static page v0\.1\.6/);
    assert.match(records[2].warnings.join('\n'), /static page reports env sit/);
  }

  console.log('test_show_deployments: all assertions passed');
})().catch((err) => { console.error(err); process.exit(1); });
