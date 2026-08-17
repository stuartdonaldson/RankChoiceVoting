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

const VERSION_PLACEHOLDER = 'var STATIC_BUILD_VERSION_ = null;';
const WEBAPP_PLACEHOLDER = 'var STATIC_WEBAPP_URL_ = null;';

// env -> local.settings.json key holding that env's deployment ID (mirrors
// tools/manage-deployments.js's TARGETS deploymentIdKey mapping).
const DEPLOYMENT_ID_KEY = { sit: 'sitDeploymentId', prod: 'prodDeploymentId', nuuc: 'nuucDeploymentId' };

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

// Pure string transform: swaps both source placeholders for their stamped values. Kept
// filesystem-free so it is unit-testable (test/test_build_static_pages.js); buildOne wires it
// to the real files. Throws if either placeholder is missing, guarding against a source edit
// that silently drops one.
function stampSource_(src, { versionString, webAppUrl }) {
  if (!src.includes(VERSION_PLACEHOLDER)) {
    throw new Error(`static-pages/src/index.html: expected placeholder not found: ${VERSION_PLACEHOLDER}`);
  }
  if (!src.includes(WEBAPP_PLACEHOLDER)) {
    throw new Error(`static-pages/src/index.html: expected placeholder not found: ${WEBAPP_PLACEHOLDER}`);
  }
  return src
    .replace(VERSION_PLACEHOLDER, `var STATIC_BUILD_VERSION_ = ${JSON.stringify(versionString)};`)
    .replace(WEBAPP_PLACEHOLDER, `var STATIC_WEBAPP_URL_ = ${JSON.stringify(webAppUrl)};`);
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
  const src = fs.readFileSync(SRC_PATH, 'utf8');
  const out = stampSource_(src, { versionString, webAppUrl });
  const outDir = path.join(DIST_ROOT, env);
  fs.mkdirSync(outDir, { recursive: true });
  fs.writeFileSync(path.join(outDir, 'index.html'), out, 'utf8');
  // Small companion file, mirroring F3Go30's convention — not currently read back by this
  // project, but cheap to carry forward for a future "confirm the live static build" check.
  fs.writeFileSync(path.join(outDir, 'version.json'), JSON.stringify({ version: versionString }), 'utf8');
  console.log(`built static-pages/dist/${env}/index.html (v${versionString}, webapp ${webAppUrl})`);
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

module.exports = { stampSource_, execUrlForEnv_, versionStringFor, buildOne };
