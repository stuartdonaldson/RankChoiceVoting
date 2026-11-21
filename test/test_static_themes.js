'use strict';

// rcballot-xz0: each env's design language is its own CSS module under static-pages/src/themes/,
// and the stamped page links only that env's module. Pure config checks — no build, no network.

const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { THEME, themeCssLink_ } = require('../tools/static-pages.js');

const ROOT = path.join(__dirname, '..');
const project = JSON.parse(fs.readFileSync(path.join(ROOT, 'gas-project.json'), 'utf8'));
const indexHtml = fs.readFileSync(path.join(ROOT, 'static-pages/src/index.html'), 'utf8');

// AC 1 — the agreed env -> design language assignment.
assert.deepEqual(THEME, { sit: 'northlake', prod: 'f3', nuuc: 'nuuc' });

// AC 2, 6 — every declared env has a theme module, scoped to its own data-theme, and the stamped
// link names exactly that module with a per-release cache-buster.
for (const env of Object.keys(project.envs)) {
  const theme = THEME[env];
  assert.ok(theme, `env ${env} has no THEME entry`);
  const css = fs.readFileSync(path.join(ROOT, 'static-pages/src/themes', `${theme}.css`), 'utf8');
  assert.match(css, new RegExp(`html\\[data-theme="${theme}"\\]`), `${theme}.css is not scoped to its data-theme`);
  assert.equal(
    themeCssLink_({ env, version: '1.2.3.4' }),
    `<link rel="stylesheet" href="themes/${theme}.css?v=1.2.3.4">`,
  );
}

// A theme module never styles another theme's page: no selector names a foreign data-theme.
for (const f of fs.readdirSync(path.join(ROOT, 'static-pages/src/themes'))) {
  const own = path.basename(f, '.css');
  const css = fs.readFileSync(path.join(ROOT, 'static-pages/src/themes', f), 'utf8');
  for (const m of css.matchAll(/data-theme="([^"]+)"/g)) assert.equal(m[1], own, `${f} names data-theme="${m[1]}"`);
}

// Unknown env fails the build rather than publishing an unstyled page.
assert.throws(() => themeCssLink_({ env: 'nope', version: '1' }), /no THEME entry/);

// The page carries the placeholder the link is stamped into, after the base <style> block so a
// module's component rules come later in the cascade.
assert.ok(indexHtml.indexOf('</style>\n<!-- STATIC_THEME_CSS_ -->') > 0, 'theme link placeholder missing or misplaced');
// Theme tokens live only in the modules now.
assert.doesNotMatch(indexHtml, /^\s*html\[data-theme="\w+"\]\s*\{/m);

console.log('test_static_themes: OK');
