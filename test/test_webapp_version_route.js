'use strict';

const assert = require('node:assert/strict');

// RECOMMENDATION.md §3.2 (gas-deploy Stage 1b/1c): cmd=version must return the stamped build
// identity with no secret required, on both GET and POST, so it works on an ANYONE_ANONYMOUS
// deployment and before ADMIN_SHARED_SECRET is bootstrapped. Ported from F3Go30's
// test/test_webapp_version_route.js — Stage 2 folds both into packages/gas-deploy's coverage.

global.GasLogger = { log: function () {}, logError: function () {}, run: function (name, fn) { return fn(); } };
global.ContentService = {
  MimeType: { JSON: 'application/json' },
  createTextOutput: function (text) {
    return { _text: text, setMimeType: function () { return this; } };
  },
};
global.HtmlService = {
  createHtmlOutput: function (html) { return { _html: html }; },
};

// The stamped build identity — version.js's globals. In the real GAS runtime version.js and
// WebApp.js are concatenated into one global scope; under Node they must be stubbed.
global.APP_VERSION = '0.1.6.1';
global.APP_VERSION_DATE = '2026-08-21T22:52:10.111Z';
global.APP_DEPLOY_TARGET = 'SIT';

global.PropertiesService = {
  getScriptProperties: function () { return { getProperty: function () { return null; }, setProperty: function () {} }; },
};
global.ScriptApp = {
  getService: function () { return { getUrl: function () { return 'https://script.google.com/macros/s/AKfycbwRCVTEST/exec'; } }; },
};
// Cross-file globals WebApp.js's other branches reach for (ApiBridge.js / webBallot.js) — never
// invoked by the cmd=version path, but stubbed so require() and the router are safe under Node.
global._renderStaticRedirect_ = function () { return { _text: '{}', setMimeType: function () { return this; } }; };
global._handleBallot = function () { return { _text: '{}', setMimeType: function () { return this; } }; };
global.handleApiPost_ = function () { return { _text: '{}', setMimeType: function () { return this; } }; };

const { handleVersionRequest_, extractDeploymentIdFromUrl_, doGet, doPost } = require('../script/WebApp.js');

function readJson_(output) {
  return JSON.parse(output._text);
}

(function testExtractDeploymentIdFromUrl() {
  assert.equal(
    extractDeploymentIdFromUrl_('https://script.google.com/macros/s/AKfycbwRCVTEST/exec'),
    'AKfycbwRCVTEST'
  );
  assert.equal(extractDeploymentIdFromUrl_(''), null);
  assert.equal(extractDeploymentIdFromUrl_(null), null);
  assert.equal(extractDeploymentIdFromUrl_('not a webapp url'), null);
})();

(function testHandleVersionRequestReturnsStampedIdentity() {
  const body = readJson_(handleVersionRequest_());
  assert.deepEqual(body, {
    ok: true,
    version: '0.1.6.1',
    versionDate: '2026-08-21T22:52:10.111Z',
    target: 'SIT',
    deploymentId: 'AKfycbwRCVTEST',
  });
})();

(function testDoGetRoutesCmdVersionNoSecret() {
  const body = readJson_(doGet({ parameter: { cmd: 'version' } }));
  assert.equal(body.ok, true);
  assert.equal(body.version, '0.1.6.1');
  assert.ok(!('adminSecret' in body));
})();

(function testDoPostRoutesCmdVersionNoSecret() {
  // No adminSecret in the payload at all — cmd=version must not require one, and must be
  // routed ahead of the cmd=admin branch that would otherwise gate it.
  const body = readJson_(doPost({ parameter: { cmd: 'version' }, postData: { type: 'text/plain', length: 2, contents: '{}' } }));
  assert.equal(body.ok, true);
  assert.equal(body.target, 'SIT');
  assert.equal(body.deploymentId, 'AKfycbwRCVTEST');
})();

console.log('test_webapp_version_route: all tests passed');
