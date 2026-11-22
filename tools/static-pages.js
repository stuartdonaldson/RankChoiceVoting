#!/usr/bin/env node
/**
 * RankChoiceVoting static front-end config. The pipeline itself — placeholder stamping, the
 * env-agreement assertion, per-env dist/, the scoped cross-repo publish, the PUBLISHERS.md
 * ownership guard and the published-build verification — lives in GAS-Core's shared `gas-static`
 * package. See that package's README. This file is only what is specific to this project.
 *
 * It replaces tools/build-static-pages.js and tools/publish-static-pages.js, both deleted: they
 * were this project's hand-written copy of those steps, and had already diverged from F3Go30's and
 * GActionSheet's copies of the same three steps.
 *
 * Never invoked on its own in normal use — tools/manage-deployments.js chains
 * build -> publish -> assertPublishedBuild as required post-deploy steps of
 * `pnpm run deploy:sit` / `:prod` / `:nuuc`, so the published page and the deployment it calls can
 * never diverge.
 *
 *   node -e "require('./tools/static-pages.js').build('sit')"     # build only, no publish
 *
 * SOURCE OF THE BACKEND URL
 *   `script/version.js` — the server-side build identity manage-deployments.js regenerates on every
 *   deploy, carrying the /exec URL of the deployment that deploy just landed in. This replaces the
 *   old build's `https://script.google.com/macros/s/<sitDeploymentId>/exec` construction from
 *   local.settings.json, which was reconciled against nothing and could name a deployment this
 *   deploy never landed in — GAS-Core adr/0001. Building an env whose declared `deployTarget`
 *   disagrees with the stamped `env` now throws before anything is written.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const { runStatic, readBuildInfo_ } = require('gas-static');

const ROOT = path.join(__dirname, '..');
const VERSION_JS = 'script/version.js';

// Identifiers of external resources — the script projects, the spreadsheets, and the static hosts
// this project publishes to — are committed project truth (GAS-Core adr/0004). Before this file
// existed the host repo, its Pages URL and each env's folder had THREE copies: manage-deployments'
// STATIC_ENTRY_BASE_URL, publish-static-pages' DEST_MAP, and script/ApiBridge.js's
// _staticPagesBaseUrl_ — the last of which is GAS-side and cannot require() the other two.
const project = JSON.parse(fs.readFileSync(path.join(ROOT, 'gas-project.json'), 'utf8'));

// env -> the APP_DEPLOY_TARGET / BUILD_INFO.env value that env's deploy stamps. These are the
// deploy labels from manage-deployments.js's targets, and gas-static refuses to build an env whose
// declared deployTarget disagrees with what is currently stamped.
const DEPLOY_TARGET = { sit: 'SIT', prod: 'PROD', nuuc: 'NUUC' };

// env -> design language. Each design language is its own CSS module,
// static-pages/src/themes/<theme>.css, copied to every env's dist/ as an asset; the stamped page
// links only its env's module and sets data-theme to match. prod carries the F3 tactical language
// (docs/f3-dl-idea.md); sit and nuuc use the Northlake Digital Design System Option A (DevStandard
// ui-framework/ 03; rcballot-gk3 moved nuuc off the retired nuuc theme). A page-content choice, so
// it stays here rather than in gas-project.json, and is expressed through gas-static's generic
// `placeholders` map — the package deliberately does not know about themes.
const THEME = { sit: 'northlake', prod: 'f3', nuuc: 'northlake' };
const THEMES_DIR = 'static-pages/src/themes';

/**
 * The <link> to env's theme module. `?v=<version>` makes each release fetch its own stylesheet
 * rather than a cached one from the previous publish, since assertPublishedBuild only verifies
 * index.html. Throws when the env has no theme or its module is missing, so a typo fails the
 * build instead of publishing an unstyled page.
 *
 * @param {{env:string, version:string}} ctx
 * @return {string}
 */
function themeCssLink_(ctx) {
  const theme = THEME[ctx.env];
  if (!theme) throw new Error(`tools/static-pages.js: no THEME entry for env '${ctx.env}'`);
  const file = path.join(ROOT, THEMES_DIR, `${theme}.css`);
  if (!fs.existsSync(file)) throw new Error(`${THEMES_DIR}/${theme}.css not found (THEME.${ctx.env})`);
  return `<link rel="stylesheet" href="themes/${theme}.css?v=${encodeURIComponent(ctx.version)}">`;
}

// Google Fonts stylesheet for themes needing webfonts beyond system defaults. A theme not listed
// here gets no font <link> stamped.
const THEME_FONTS_HTML = {
  f3: '<link rel="preconnect" href="https://fonts.googleapis.com">\n' +
    '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>\n' +
    '<link href="https://fonts.googleapis.com/css2?family=Saira+Stencil+One&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;700&display=swap" rel="stylesheet">',
  northlake: '<link rel="preconnect" href="https://fonts.googleapis.com">\n' +
    '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>\n' +
    '<link href="https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600;700&display=swap" rel="stylesheet">',
};

/**
 * APP_CONTACT out of script/version.js — the same "who to contact" address the GAS-side About
 * dialog shows (onOpen.js's showAbout()), so the static page's own troubleshooting copy
 * (callApi_'s network-failure message) points a stuck respondent somewhere real.
 *
 * Regex rather than require(): version.js is a GAS script file (bare consts, no module.exports).
 * This is NOT a second copy of the package's BUILD_INFO reader — APP_CONTACT is hand-maintained
 * source, not part of the generated BUILD_INFO literal, so `readBuildInfo_` cannot supply it.
 */
function devContact_() {
  const src = fs.readFileSync(path.join(ROOT, VERSION_JS), 'utf8');
  const m = /^const APP_CONTACT\s*=\s*'([^']*)'/m.exec(src);
  return m ? m[1] : '';
}

/**
 * NUUC-Dispatch sign-in settings for env (rcballot-dx3), from gas-project.json
 * envs.<env>.dispatch — committed, non-secret project truth (GAS-Core adr/0004). An env without a
 * `dispatch` block (prod) stamps null, which the page reads as "sign-in off"; so does any empty
 * field (the page checks), which is how an env builds before its Dispatch target is registered.
 *
 * @param {string} env
 * @return {?{url:string, clientId:string, aud:string}}
 */
function dispatchFor_(env) {
  const d = project.envs[env] && project.envs[env].dispatch;
  if (!d) return null;
  return { url: String(d.url || ''), clientId: String(d.clientId || ''), aud: String(d.aud || '') };
}

/** Reads one field of the generated BUILD_INFO literal with the same code the pipeline reads it. */
function buildInfoField_(name) {
  const value = readBuildInfo_(path.join(ROOT, VERSION_JS))[name];
  if (!value) throw new Error(`${VERSION_JS} has no ${name} — deploy first.`);
  return value;
}

/** Composes gas-static's per-env config from the committed identifiers. */
function envsFromProject_() {
  const out = {};
  for (const [env, def] of Object.entries(project.envs)) {
    const host = project.staticHosts[def.static.host];
    if (!host) throw new Error(`gas-project.json: env '${env}' names unknown staticHosts entry '${def.static.host}'`);
    out[env] = {
      deployTarget: DEPLOY_TARGET[env],
      repoKey: host.repoKey,
      dest: def.static.dest,
      label: DEPLOY_TARGET[env],
    };
  }
  return out;
}

const envs = envsFromProject_();

/**
 * The live base URL is COMPOSED from the host's Pages URL and this env's folder, never declared
 * separately — so the URL the deploy summary prints and the URL `assertPublishedBuild` polls can
 * never drift from the folder actually published to. Both halves are registered in the host repo's
 * PUBLISHERS.md, which is what authorises the publish in the first place (GAS-Core adr/0003).
 */
function liveUrl_(env) {
  return project.staticHosts[project.envs[env].static.host].pagesUrl + envs[env].dest + '/';
}

module.exports = runStatic({
  root: ROOT,
  // Must match this project's entry in BOTH host repos' PUBLISHERS.md ownership maps
  // (f3go30/static-pages for sit/prod, nuuc-it/Static for nuuc): gas-static refuses a publish to a
  // folder registered to another project.
  projectName: 'RankChoiceVoting',
  srcDir: 'static-pages/src',
  distDir: 'static-pages/dist',
  stampedPages: ['index.html'],
  webappUrl: { from: 'buildInfo', file: VERSION_JS, envField: 'env' },

  // gas-static stamps STATIC_BUILD_VERSION_ and STATIC_WEBAPP_URL_ itself. The three below are
  // this project's own, carried through the generic `placeholders` map with NO package change —
  // gas-static's README §Provenance records that as the reason RCV's theme/fonts/contact stamping
  // was deliberately not ported into the package. `placeholders` is raw-token substitution, so
  // each value is the whole replacement text, not just the right-hand side.
  placeholders: {
    'var STATIC_DEV_CONTACT_ = null;': () =>
      `var STATIC_DEV_CONTACT_ = ${JSON.stringify(devContact_())};`,
    'data-theme="STATIC_THEME_"': (ctx) => `data-theme="${THEME[ctx.env] || ''}"`,
    '<!-- STATIC_THEME_FONTS_ -->': (ctx) => THEME_FONTS_HTML[THEME[ctx.env]] || '',
    '<!-- STATIC_THEME_CSS_ -->': themeCssLink_,
    'var STATIC_DISPATCH_ = null;': (ctx) => `var STATIC_DISPATCH_ = ${JSON.stringify(dispatchFor_(ctx.env))};`,
  },

  envs,
  liveUrl: liveUrl_,
  commitMessage: ({ env }) => `Publish RankChoiceVoting static pages v${buildInfoField_('version')} (${env})`,
});

module.exports.liveUrl_ = liveUrl_;
module.exports.envsFromProject_ = envsFromProject_;
module.exports.dispatchFor_ = dispatchFor_;
module.exports.themeCssLink_ = themeCssLink_;
module.exports.THEME = THEME;
