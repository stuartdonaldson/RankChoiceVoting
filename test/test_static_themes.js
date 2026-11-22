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
// rcballot-gk3: Communications approved the palette, so nuuc moved to northlake (themes/nuuc.css retired).
assert.deepEqual(THEME, { sit: 'northlake', prod: 'f3', nuuc: 'northlake' });

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

// --- rcballot-71a: Northlake v0.3 header tools and light/dark scheme -------------------------

const themesDir = path.join(ROOT, 'static-pages/src/themes');
const moduleCss = (name) => fs.readFileSync(path.join(themesDir, `${name}.css`), 'utf8');

// Only the northlake module opts into the header tools / scheme switch; f3 never does, so
// their pages get no data-scheme and no header JS.
for (const f of fs.readdirSync(themesDir)) {
  const optedIn = /--app-shell:\s*1\s*;/.test(fs.readFileSync(path.join(themesDir, f), 'utf8'));
  assert.equal(optedIn, f === 'northlake.css', `${f}: --app-shell opt-in must be northlake-only`);
}
// The dark scheme is a scoped block that also sets color-scheme (§4.2).
assert.match(moduleCss('northlake'), /html\[data-theme="northlake"\]\[data-scheme="dark"\]\s*\{\s*color-scheme:\s*dark;/);

// The <head> script sits right after the theme <link> (so the opt-in is readable and it runs
// before first paint), and is the only place that resolves the preference.
const headScript = /<!-- STATIC_THEME_CSS_ -->\n<script>([\s\S]*?)<\/script>\n<\/head>/.exec(indexHtml);
assert.ok(headScript, 'scheme script must directly follow the STATIC_THEME_CSS_ placeholder in <head>');

// Behaviour of that script, run against a minimal fake DOM.
const vm = require('node:vm');
function runHead({ optIn, stored, storageThrows, systemDark }) {
  const attrs = {};
  const storage = {
    getItem() { if (storageThrows) throw new Error('blocked'); return stored; },
    setItem(k, v) { if (storageThrows) throw new Error('blocked'); storage.saved = [k, v]; },
  };
  const listeners = [];
  const media = { matches: !!systemDark, addEventListener: (t, fn) => listeners.push(fn) };
  const win = { matchMedia: () => media };
  const ctx = {
    window: win,
    localStorage: storage,
    document: { documentElement: { setAttribute: (k, v) => { attrs[k] = v; } } },
    getComputedStyle: () => ({ getPropertyValue: (n) => (n === '--app-shell' && optIn ? ' 1' : '') }),
  };
  vm.runInNewContext(headScript[1], ctx);
  return { attrs, api: win.appScheme_, storage, media, listeners };
}
// No opt-in (f3): nothing is set and no API is exposed.
{
  const r = runHead({ optIn: false, systemDark: true });
  assert.deepEqual(r.attrs, {});
  assert.equal(r.api, undefined);
}
// Default is System, following the OS; a change of OS preference re-applies while on System.
{
  const r = runHead({ optIn: true, stored: null, systemDark: true });
  assert.deepEqual(r.attrs, { 'data-scheme-pref': 'system', 'data-scheme': 'dark' });
  r.media.matches = false; r.listeners.forEach((fn) => fn());
  assert.equal(r.attrs['data-scheme'], 'light');
}
// A stored explicit choice wins over the OS, persists under the §14 key, and ignores junk.
{
  const r = runHead({ optIn: true, stored: 'light', systemDark: true });
  assert.equal(r.attrs['data-scheme'], 'light');
  r.api.set('dark');
  assert.deepEqual(r.storage.saved, ['northlake-theme', 'dark']);
  assert.equal(r.attrs['data-scheme'], 'dark');
  r.api.set('purple');
  assert.equal(r.api.get(), 'dark');
  assert.equal(runHead({ optIn: true, stored: 'purple' }).attrs['data-scheme-pref'], 'system');
}
// Storage blocked: falls back to System, and an in-session choice still survives an OS change.
{
  const r = runHead({ optIn: true, storageThrows: true, systemDark: false });
  assert.equal(r.attrs['data-scheme'], 'light');
  r.api.set('dark');
  r.media.matches = false; r.listeners.forEach((fn) => fn());
  assert.equal(r.attrs['data-scheme'], 'dark');
}

// The header markup stays hidden by default (only a theme module shows it), and the header JS
// and the sign-in relocation are gated on the opt-in the head script reports.
assert.match(indexHtml, /^\s*\.app-brand \{ display: none; \}/m);
assert.match(indexHtml, /var APP_SHELL_ = !!window\.appScheme_;/);
assert.match(indexHtml, /if \(!DISPATCH_\) return;\n\s*if \(APP_SHELL_\) \{ renderShellSignIn_\(container, intro\); return; \}/);
assert.match(indexHtml, /function initAppShell_\(\) \{\n\s*if \(!APP_SHELL_\) return;/);

// Every asset the page references is committed (never read from Drive at build time, §2.5).
for (const m of indexHtml.matchAll(/src="(assets\/[^"]+)"/g)) {
  assert.ok(fs.existsSync(path.join(ROOT, 'static-pages/src', m[1])), `${m[1]} referenced but not in static-pages/src`);
}

// --- rcballot-719: 02 §2.3 token values and contrast, both schemes ------------------------------

const nlCss = moduleCss('northlake');
const lightBlock = /html\[data-theme="northlake"\]\s*\{([\s\S]*?)\n\}/.exec(nlCss)[1];
const darkBlock = /html\[data-theme="northlake"\]\[data-scheme="dark"\]\s*\{([\s\S]*?)\n\}/.exec(nlCss)[1];
function tokens(block, base) {
  const raw = { ...base };
  for (const m of block.matchAll(/(--nl-[\w-]+):\s*([^;]+);/g)) raw[m[1]] = m[2].trim();
  const t = {};
  const res = (k) => { const v = /^var\((--nl-[\w-]+)\)$/.exec(raw[k]); return v ? res(v[1]) : raw[k]; };
  for (const k of Object.keys(raw)) t[k] = res(k);
  t.raw = raw;
  return t;
}
const light = tokens(lightBlock, {});
const dark = tokens(darkBlock, light.raw);
const lum = (hex) => {
  const c = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255)
    .map((v) => (v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4));
  return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2];
};
const ratio = (a, b) => { const [x, y] = [lum(a), lum(b)].sort((p, q) => q - p); return (x + 0.05) / (y + 0.05); };
const atLeast = (scheme, t, fg, bg, min) => {
  const r = ratio(t[fg], t[bg]);
  assert.ok(r >= min, `${scheme}: ${fg} on ${bg} is ${r.toFixed(2)}:1, needs ${min}`);
};

// Exact 02 §2.3 values for the tokens this bead added or changed.
const expected = {
  light: {
    '--nl-border-strong': '#7f8c92', '--nl-disabled-bg': '#d5dde0', '--nl-disabled-text': '#4a565e',
    '--nl-success-bg': '#edf6e6', '--nl-success-text': '#2e6b1f', '--nl-warning-bg': '#fff4e0', '--nl-warning-text': '#8a4b00',
    '--nl-error-bg': '#fdecea', '--nl-error-text': '#b3261e', '--nl-info-bg': '#e9f3fb', '--nl-info-text': '#1663a5',
    '--nl-radius-control': '8px', '--nl-radius-card': '12px', '--nl-radius-dialog': '16px',
  },
  dark: {
    '--nl-border-strong': '#7f939f', '--nl-disabled-bg': '#33444f', '--nl-disabled-text': '#b5c6d0',
    '--nl-success-bg': '#1f3a27', '--nl-success-text': '#a5d68a', '--nl-warning-bg': '#3d2e14', '--nl-warning-text': '#ffc46b',
    '--nl-error-bg': '#4a2220', '--nl-error-text': '#ffb4ab', '--nl-info-bg': '#183b53', '--nl-info-text': '#b9dfff',
    '--nl-radius-control': '8px', '--nl-radius-card': '12px', '--nl-radius-dialog': '16px',
  },
};
for (const [scheme, t] of [['light', light], ['dark', dark]]) {
  for (const [k, v] of Object.entries(expected[scheme])) assert.equal(t[k], v, `${scheme} ${k}`);
  // Control borders: >= 3:1 against both surface and page background (WCAG 1.4.11).
  atLeast(scheme, t, '--nl-border-strong', '--nl-surface', 3);
  atLeast(scheme, t, '--nl-border-strong', '--nl-background', 3);
  // Disabled controls keep legible text; status text is legible on its own tint and on the surface.
  atLeast(scheme, t, '--nl-disabled-text', '--nl-disabled-bg', 4.5);
  for (const s of ['success', 'warning', 'error', 'info']) {
    atLeast(scheme, t, `--nl-${s}-text`, `--nl-${s}-bg`, 4.5);
    atLeast(scheme, t, `--nl-${s}-text`, '--nl-surface', 4.5);
  }
  // Core text pairs on every surface a rule actually uses (muted never on the subtle surface).
  for (const bg of ['--nl-surface', '--nl-background']) {
    for (const fg of ['--nl-text', '--nl-text-muted', '--nl-heading', '--nl-primary']) atLeast(scheme, t, fg, bg, 4.5);
  }
  atLeast(scheme, t, '--nl-on-primary', '--nl-primary', 4.5);
  atLeast(scheme, t, '--nl-on-primary', '--nl-primary-hover', 4.5);
  atLeast(scheme, t, '--nl-text', '--nl-surface-subtle', 4.5);
  atLeast(scheme, t, '--nl-primary', '--nl-background', 3); // focus ring, non-text
}
// Spec's known failing pair: light muted text on the muted surface is below 4.5:1 ...
assert.ok(ratio(light['--nl-text-muted'], light['--nl-surface-subtle']) < 4.5);
// ... so no rule may consume --nl-surface-subtle (02 §2.3: never muted text on the muted surface).
assert.doesNotMatch(nlCss.replace(/\/\*[\s\S]*?\*\//g, '').replace(/--nl-surface-subtle:[^;]+;/g, ''), /surface-subtle/);
// The control-border token is what form fields actually use.
assert.match(nlCss, /input\[type="text"\], input\[type="url"\], textarea\)\s*\{[^}]*border-color: var\(--nl-border-strong\)/);

console.log('test_static_themes: OK');
