#!/usr/bin/env node
/**
 * Builds static-pages/src/index.html into per-environment, publishable copies under
 * static-pages/dist/<env>/index.html. Adapted from F3Go30's tools/build-static-pages.js.
 *
 * Three environments, matching the GAS deploy targets (tools/manage-deployments.js):
 *   sit  -> static-pages/dist/sit/index.html
 *   prod -> static-pages/dist/prod/index.html
 *   nuuc -> static-pages/dist/nuuc/index.html
 *
 * This build step stamps two placeholders in the source:
 *
 *   1. STATIC_BUILD_VERSION_ (`var STATIC_BUILD_VERSION_ = null;`) — a version string in the
 *      same shape script/version.js gets stamped with (manage-deployments.js's stampVersion):
 *      "<version>.<build>" for sit, bare "<version>" for prod/nuuc (both bump the semver patch
 *      and reset build to 0, same as PROD's own stamping). This is only ever a fast, offline
 *      first-paint value — it's not currently reconciled against a live version check, unlike
 *      F3Go30's equivalent — so it does not need to match whichever version is actually live on
 *      that environment's deployment at build time.
 *   2. STATIC_WEBAPP_URL_ (`var STATIC_WEBAPP_URL_ = null;`) — the env's GAS webapp /exec URL,
 *      derived from the deployment ID in local.settings.json (sitDeploymentId/prodDeploymentId/
 *      nuucDeploymentId). This is the baked default backend the page calls when opened WITHOUT
 *      a ?webapp= query param. A ?webapp= param still overrides it when present. Left null in
 *      the unbuilt source so static-pages/src/index.html also runs directly for manual/local
 *      testing (with ?webapp= supplied explicitly).
 *
 * Usage:
 *   node tools/build-static-pages.js [--env sit|prod|nuuc|all]   (default: all)
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const SRC_PATH = path.join(ROOT, 'static-pages', 'src', 'index.html');
const DIST_ROOT = path.join(ROOT, 'static-pages', 'dist');
const PKG_PATH = path.join(ROOT, 'package.json');
const SETTINGS_PATH = path.join(ROOT, 'local.settings.json');
const VERSION_JS_PATH = path.join(ROOT, 'script', 'version.js');

const VERSION_PLACEHOLDER = 'var STATIC_BUILD_VERSION_ = null;';
const WEBAPP_PLACEHOLDER = 'var STATIC_WEBAPP_URL_ = null;';
const DEV_CONTACT_PLACEHOLDER = 'var STATIC_DEV_CONTACT_ = null;';
const THEME_ATTR_PLACEHOLDER = 'data-theme="STATIC_THEME_"';
const THEME_FONTS_PLACEHOLDER = '<!-- STATIC_THEME_FONTS_ -->';

// env -> local.settings.json key holding that env's deployment ID (mirrors
// tools/manage-deployments.js's TARGETS deploymentIdKey mapping).
const DEPLOYMENT_ID_KEY = { sit: 'sitDeploymentId', prod: 'prodDeploymentId', nuuc: 'nuucDeploymentId' };

// env -> design language (CSS custom-property theme, see static-pages/src/index.html's :root /
// html[data-theme] blocks). sit and prod carry the F3 tactical design language (docs/
// f3-dl-idea.md); nuuc carries the Northlake UU design language (docs/nuuc-dl-idea.md). An empty
// theme name would leave html's data-theme attribute blank, matching no html[data-theme="..."]
// selector and falling through to the :root (legacy) values — not used by any env today.
const THEME = { sit: 'f3', prod: 'f3', nuuc: 'nuuc' };

// Google Fonts stylesheet for themes that need webfonts beyond system defaults. Themes not
// listed here get no font <link> stamped.
const THEME_FONTS_HTML = {
  f3: '<link rel="preconnect" href="https://fonts.googleapis.com">\n' +
    '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>\n' +
    '<link href="https://fonts.googleapis.com/css2?family=Saira+Stencil+One&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;700&display=swap" rel="stylesheet">',
  nuuc: '<link rel="preconnect" href="https://fonts.googleapis.com">\n' +
    '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>\n' +
    '<link href="https://fonts.googleapis.com/css2?family=Lora:wght@400;600;700&family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">',
};

function versionStringFor(env, pkg) {
  if (env === 'sit') return `${pkg.version}.${pkg.build || 0}`;
  return String(pkg.version);
}

// Resolves the env's GAS webapp /exec base URL from its deployment ID. Throws (rather than
// stamping an empty/partial URL) when the ID is absent, so a misconfigured settings file fails
// the build loudly instead of shipping a dead default backend.
function execUrlForEnv_(env, settings) {
  const key = DEPLOYMENT_ID_KEY[env];
  if (!key) throw new Error(`build-static-pages: unknown env '${env}'`);
  const deploymentId = settings[key];
  if (!deploymentId) {
    throw new Error(`build-static-pages: ${key} is not set in local.settings.json (needed to bake the ${env} webapp URL)`);
  }
  return `https://script.google.com/macros/s/${deploymentId}/exec`;
}

// Pulls APP_CONTACT out of script/version.js (the same "who to contact" address already shown
// in the GAS-side About dialog — see onOpen.js's showAbout()) so the static page's own
// troubleshooting messages (callApi_'s network-failure copy) can point a stuck respondent
// somewhere real instead of a dead end. Regex, not require() — version.js is a GAS script file
// (bare consts, no module.exports), not a requirable Node module.
function devContactFromVersionJs_(versionJsSrc) {
  const m = /^const APP_CONTACT\s*=\s*'([^']*)'/m.exec(versionJsSrc);
  return m ? m[1] : '';
}

// Pure string transform: swaps every source placeholder for its stamped value. Kept
// filesystem-free so it is unit-testable (test/test_build_static_pages.js); buildOne wires it
// to the real files. Throws if any placeholder is missing, guarding against a source edit that
// silently drops one.
function stampSource_(src, { versionString, webAppUrl, theme, devContact }) {
  if (!src.includes(VERSION_PLACEHOLDER)) {
    throw new Error(`static-pages/src/index.html: expected placeholder not found: ${VERSION_PLACEHOLDER}`);
  }
  if (!src.includes(WEBAPP_PLACEHOLDER)) {
    throw new Error(`static-pages/src/index.html: expected placeholder not found: ${WEBAPP_PLACEHOLDER}`);
  }
  if (!src.includes(DEV_CONTACT_PLACEHOLDER)) {
    throw new Error(`static-pages/src/index.html: expected placeholder not found: ${DEV_CONTACT_PLACEHOLDER}`);
  }
  if (!src.includes(THEME_ATTR_PLACEHOLDER)) {
    throw new Error(`static-pages/src/index.html: expected placeholder not found: ${THEME_ATTR_PLACEHOLDER}`);
  }
  if (!src.includes(THEME_FONTS_PLACEHOLDER)) {
    throw new Error(`static-pages/src/index.html: expected placeholder not found: ${THEME_FONTS_PLACEHOLDER}`);
  }
  return src
    .replace(VERSION_PLACEHOLDER, `var STATIC_BUILD_VERSION_ = ${JSON.stringify(versionString)};`)
    .replace(WEBAPP_PLACEHOLDER, `var STATIC_WEBAPP_URL_ = ${JSON.stringify(webAppUrl)};`)
    .replace(DEV_CONTACT_PLACEHOLDER, `var STATIC_DEV_CONTACT_ = ${JSON.stringify(devContact || '')};`)
    .replace(THEME_ATTR_PLACEHOLDER, `data-theme="${theme || ''}"`)
    .replace(THEME_FONTS_PLACEHOLDER, THEME_FONTS_HTML[theme] || '');
}

function loadSettings_() {
  if (!fs.existsSync(SETTINGS_PATH)) {
    throw new Error('local.settings.json not found at project root — copy local.settings.json.example and populate the deployment IDs.');
  }
  return JSON.parse(fs.readFileSync(SETTINGS_PATH, 'utf8'));
}

function buildOne(env) {
  const pkg = JSON.parse(fs.readFileSync(PKG_PATH, 'utf8'));
  const settings = loadSettings_();
  const versionString = versionStringFor(env, pkg);
  const webAppUrl = execUrlForEnv_(env, settings);
  const theme = THEME[env] || '';
  const devContact = devContactFromVersionJs_(fs.readFileSync(VERSION_JS_PATH, 'utf8'));
  const src = fs.readFileSync(SRC_PATH, 'utf8');
  const out = stampSource_(src, { versionString, webAppUrl, theme, devContact });
  const outDir = path.join(DIST_ROOT, env);
  fs.mkdirSync(outDir, { recursive: true });
  fs.writeFileSync(path.join(outDir, 'index.html'), out, 'utf8');
  // Small companion file, mirroring F3Go30's convention — not currently read back by this
  // project, but cheap to carry forward for a future "confirm the live static build" check.
  fs.writeFileSync(path.join(outDir, 'version.json'), JSON.stringify({ version: versionString }), 'utf8');
  console.log(`built static-pages/dist/${env}/index.html (v${versionString}, theme '${theme || 'legacy'}', webapp ${webAppUrl})`);
}

function main() {
  const args = process.argv.slice(2);
  const envIdx = args.indexOf('--env');
  const env = envIdx !== -1 ? args[envIdx + 1] : 'all';
  const envs = env === 'all' ? ['sit', 'prod', 'nuuc'] : [env];
  envs.forEach(buildOne);
}

if (require.main === module) {
  main();
}

module.exports = { stampSource_, execUrlForEnv_, versionStringFor, devContactFromVersionJs_, buildOne };
