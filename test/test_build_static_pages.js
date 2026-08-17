/**
 * Unit tests for tools/build-static-pages.js — pure, filesystem-free functions only
 * (stampSource_, execUrlForEnv_, versionStringFor). Adapted from F3Go30's
 * test/test_build_static_pages.js.
 */
const assert = require('assert');
const { stampSource_, execUrlForEnv_, versionStringFor } = require('../tools/build-static-pages.js');

let failures = 0;
function test(name, fn) {
  try {
    fn();
    console.log(`  ok - ${name}`);
  } catch (err) {
    failures++;
    console.error(`  FAIL - ${name}`);
    console.error(`    ${err.message}`);
  }
}

console.log('build-static-pages.js');

test('stampSource_ replaces both placeholders', () => {
  const src = [
    '<script>',
    'var STATIC_BUILD_VERSION_ = null;',
    'var STATIC_WEBAPP_URL_ = null;',
    '</script>',
  ].join('\n');
  const out = stampSource_(src, { versionString: '0.1.4.2', webAppUrl: 'https://script.google.com/macros/s/ABC/exec' });
  assert.ok(out.includes('var STATIC_BUILD_VERSION_ = "0.1.4.2";'));
  assert.ok(out.includes('var STATIC_WEBAPP_URL_ = "https://script.google.com/macros/s/ABC/exec";'));
  assert.ok(!out.includes('= null;'));
});

test('stampSource_ throws when the version placeholder is missing', () => {
  const src = 'var STATIC_WEBAPP_URL_ = null;';
  assert.throws(() => stampSource_(src, { versionString: '1.0.0', webAppUrl: 'https://x/exec' }), /STATIC_BUILD_VERSION_/);
});

test('stampSource_ throws when the webapp placeholder is missing', () => {
  const src = 'var STATIC_BUILD_VERSION_ = null;';
  assert.throws(() => stampSource_(src, { versionString: '1.0.0', webAppUrl: 'https://x/exec' }), /STATIC_WEBAPP_URL_/);
});

test('versionStringFor: sit is version.build, prod/nuuc are bare version', () => {
  const pkg = { version: '0.1.4', build: 7 };
  assert.strictEqual(versionStringFor('sit', pkg), '0.1.4.7');
  assert.strictEqual(versionStringFor('prod', pkg), '0.1.4');
  assert.strictEqual(versionStringFor('nuuc', pkg), '0.1.4');
});

test('versionStringFor: sit defaults build to 0 when unset', () => {
  assert.strictEqual(versionStringFor('sit', { version: '0.1.4' }), '0.1.4.0');
});

test('execUrlForEnv_ builds the /exec URL from each env\'s deploymentId key', () => {
  const settings = { sitDeploymentId: 'S1', prodDeploymentId: 'P1', nuucDeploymentId: 'N1' };
  assert.strictEqual(execUrlForEnv_('sit', settings), 'https://script.google.com/macros/s/S1/exec');
  assert.strictEqual(execUrlForEnv_('prod', settings), 'https://script.google.com/macros/s/P1/exec');
  assert.strictEqual(execUrlForEnv_('nuuc', settings), 'https://script.google.com/macros/s/N1/exec');
});

test('execUrlForEnv_ throws when the deployment id is missing', () => {
  assert.throws(() => execUrlForEnv_('sit', {}), /sitDeploymentId/);
});

test('execUrlForEnv_ throws for an unknown env', () => {
  assert.throws(() => execUrlForEnv_('staging', { sitDeploymentId: 'S1' }), /unknown env/);
});

if (failures > 0) {
  console.error(`\n${failures} failure(s).`);
  process.exit(1);
} else {
  console.log('\nAll build-static-pages.js tests passed.');
}
