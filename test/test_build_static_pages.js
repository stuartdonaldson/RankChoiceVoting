/**
 * Unit tests for tools/build-static-pages.js — pure, filesystem-free functions only
 * (stampSource_, execUrlForEnv_, versionStringFor). Adapted from F3Go30's
 * test/test_build_static_pages.js.
 */
const assert = require('assert');
const { stampSource_, execUrlForEnv_, versionStringFor, devContactFromVersionJs_ } = require('../tools/build-static-pages.js');

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

// Minimal fixture carrying all five placeholders stampSource_ requires (version, webapp, dev
// contact, theme attribute, theme fonts comment) — mirrors the real static-pages/src/index.html
// shape closely enough to exercise the replacement logic without dragging in the whole file.
function fixtureSrc_() {
  return [
    '<html data-theme="STATIC_THEME_">',
    '<head><!-- STATIC_THEME_FONTS_ --></head>',
    '<script>',
    'var STATIC_BUILD_VERSION_ = null;',
    'var STATIC_WEBAPP_URL_ = null;',
    'var STATIC_DEV_CONTACT_ = null;',
    '</script>',
    '</html>',
  ].join('\n');
}

test('stampSource_ replaces all placeholders', () => {
  const out = stampSource_(fixtureSrc_(), {
    versionString: '0.1.4.2',
    webAppUrl: 'https://script.google.com/macros/s/ABC/exec',
    theme: 'f3',
    devContact: 'dev@example.com',
  });
  assert.ok(out.includes('var STATIC_BUILD_VERSION_ = "0.1.4.2";'));
  assert.ok(out.includes('var STATIC_WEBAPP_URL_ = "https://script.google.com/macros/s/ABC/exec";'));
  assert.ok(out.includes('var STATIC_DEV_CONTACT_ = "dev@example.com";'));
  assert.ok(out.includes('data-theme="f3"'));
  assert.ok(out.includes('fonts.googleapis.com'));
  assert.ok(!out.includes('= null;'));
  assert.ok(!out.includes('STATIC_THEME_"'));
});

test('stampSource_ leaves data-theme empty and drops fonts when theme is falsy (legacy)', () => {
  const out = stampSource_(fixtureSrc_(), {
    versionString: '0.1.4',
    webAppUrl: 'https://script.google.com/macros/s/ABC/exec',
    theme: '',
    devContact: '',
  });
  assert.ok(out.includes('data-theme=""'));
  assert.ok(!out.includes('fonts.googleapis.com'));
  assert.ok(out.includes('var STATIC_DEV_CONTACT_ = "";'));
});

test('stampSource_ throws when the version placeholder is missing', () => {
  const src = fixtureSrc_().replace('var STATIC_BUILD_VERSION_ = null;', '');
  assert.throws(() => stampSource_(src, { versionString: '1.0.0', webAppUrl: 'https://x/exec', theme: '' }), /STATIC_BUILD_VERSION_/);
});

test('stampSource_ throws when the webapp placeholder is missing', () => {
  const src = fixtureSrc_().replace('var STATIC_WEBAPP_URL_ = null;', '');
  assert.throws(() => stampSource_(src, { versionString: '1.0.0', webAppUrl: 'https://x/exec', theme: '' }), /STATIC_WEBAPP_URL_/);
});

test('stampSource_ throws when the dev contact placeholder is missing', () => {
  const src = fixtureSrc_().replace('var STATIC_DEV_CONTACT_ = null;', '');
  assert.throws(() => stampSource_(src, { versionString: '1.0.0', webAppUrl: 'https://x/exec', theme: '' }), /STATIC_DEV_CONTACT_/);
});

test('stampSource_ throws when the theme attribute placeholder is missing', () => {
  const src = fixtureSrc_().replace('data-theme="STATIC_THEME_"', '');
  assert.throws(() => stampSource_(src, { versionString: '1.0.0', webAppUrl: 'https://x/exec', theme: '' }), /STATIC_THEME_/);
});

test('stampSource_ throws when the theme fonts placeholder is missing', () => {
  const src = fixtureSrc_().replace('<!-- STATIC_THEME_FONTS_ -->', '');
  assert.throws(() => stampSource_(src, { versionString: '1.0.0', webAppUrl: 'https://x/exec', theme: '' }), /STATIC_THEME_FONTS_/);
});

test('devContactFromVersionJs_ extracts APP_CONTACT', () => {
  const src = "const APP_VERSION = '0.1.0';\nconst APP_CONTACT       = 'stuart.donaldson@gmail.com';\n";
  assert.strictEqual(devContactFromVersionJs_(src), 'stuart.donaldson@gmail.com');
});

test('devContactFromVersionJs_ returns empty string when APP_CONTACT is absent', () => {
  assert.strictEqual(devContactFromVersionJs_("const APP_VERSION = '0.1.0';\n"), '');
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
