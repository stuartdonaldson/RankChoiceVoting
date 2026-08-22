#!/usr/bin/env node
/**
 * RankChoiceVoting Deployment Manager
 *
 * Adapted from F3Go30's tools/manage-deployments.js. Manages three environments —
 * SIT (dev scriptId), PROD (prod scriptId), and NUUC (a separate live deployment under its
 * own Google account) — and owns .clasp.json, version stamping, and named-deployment updates.
 *
 * Run via npm scripts:
 *   npm run deploy:sit    # bump build only + stamp SIT  + clasp push -f to sitScriptId
 *   npm run deploy:prod   # bump patch      + stamp PROD + clasp push -f to prodScriptId  (alias: npm run push)
 *   npm run deploy:nuuc   # bump patch      + stamp NUUC + clasp push -f to nuucScriptId
 *
 * package.json carries two counters: "version" (semver, PROD/NUUC-facing) and "build" (a plain
 * integer, SIT-facing). A SIT deploy leaves "version" untouched and bumps "build" instead
 * (unless --skip-bump), so repeated SIT deploys between PROD releases don't burn through patch
 * numbers; the SIT-stamped APP_VERSION is `${version}.${build}` (e.g. "0.0.1.7"). A PROD or
 * NUUC deploy bumps the patch segment of "version" (unless --skip-bump) and *always* resets
 * "build" to 0.
 *
 * Direct invocation:
 *   node tools/manage-deployments.js --deploy-sit
 *   node tools/manage-deployments.js --deploy-prod
 *   node tools/manage-deployments.js --deploy-nuuc
 *
 * Prerequisites:
 *   - local.settings.json at project root with sitScriptId, prodScriptId, and nuucScriptId
 *     populated, plus claspAuth pointing at this project's clasp credential file for SIT/PROD
 *     and nuucAuth pointing at the credential file for NUUC's separate Google account. No
 *     deployment ID fields needed — each script project must carry exactly one active named
 *     (Web app) deployment, looked up fresh via `clasp deployments` on every deploy (see
 *     findActiveDeploymentId_).
 *   - clasp authenticated into claspAuth and nuucAuth (clasp_config_auth=<file> clasp login)
 *   - @inquirer/prompts installed (npm install)
 */

const { execSync }  = require('child_process');
const fs            = require('fs');
const os            = require('os');
const path          = require('path');

// callWebapp.js's post() is this project's only HTTP client for the deployed webapp;
// assertDeployedVersion_ below polls cmd=version through it rather than adding a second one
// (RECOMMENDATION.md §3.3).
const { post: postWebapp_ } = require('./callWebapp.js');

const ROOT          = path.join(__dirname, '..');
const SETTINGS_PATH = path.join(ROOT, 'local.settings.json');
const CLASP_PATH    = path.join(ROOT, '.clasp.json');
const VERSION_PATH  = path.join(ROOT, 'script', 'version.js');
const PKG_PATH      = path.join(ROOT, 'package.json');

// clasp reads its credential file from the `clasp_config_auth` env var (lower-case, exact
// match — see @google/clasp's commands/program.js). CLASP_CONFIG is not a real clasp variable;
// setting it is a no-op that silently falls back to the default ~/.clasprc.json.
function expandHome_(p) {
  return p && p.startsWith('~') ? path.join(os.homedir(), p.slice(1)) : p;
}

function resolveClaspAuthPath_(settings, claspAuthKey) {
  const claspAuth = settings[claspAuthKey || 'claspAuth'];
  if (!claspAuth) {
    console.error(`❌  ${claspAuthKey || 'claspAuth'} is not set in local.settings.json`);
    process.exit(1);
  }
  return expandHome_(claspAuth);
}

// NUUC deploys under a separate Google account from SIT/PROD, hence its own claspAuthKey
// (nuucAuth) instead of the shared claspAuth used by sit/prod.
const TARGETS = {
  sit:  { scriptIdKey: 'sitScriptId',  label: 'SIT',  emoji: '🧪', deploymentIdKey: 'sitDeploymentId',  claspAuthKey: 'claspAuth', sheetIdKey: 'sitSheetId'  },
  prod: { scriptIdKey: 'prodScriptId', label: 'PROD', emoji: '🚀', deploymentIdKey: 'prodDeploymentId', claspAuthKey: 'claspAuth', sheetIdKey: 'prodSheetId' },
  nuuc: { scriptIdKey: 'nuucScriptId', label: 'NUUC', emoji: '⛪', deploymentIdKey: 'nuucDeploymentId', claspAuthKey: 'nuucAuth',  sheetIdKey: 'nuucSheetId' },
};

// Static front-end entry point per deploy label — must stay in sync with script/ApiBridge.js's
// _staticPagesBaseUrl_ (that's the GAS-side copy of this same map, read at runtime from
// APP_DEPLOY_TARGET; this one is the node-side copy, keyed the same way, used only to print a
// clickable link at the end of a deploy).
const STATIC_ENTRY_BASE_URL = {
  SIT:  'https://f3go30.github.io/static-pages/ballot/sit/',
  PROD: 'https://f3go30.github.io/static-pages/ballot/prod/',
  NUUC: 'https://nuuc-it.github.io/Static/pub/ballot/',
};

// ─────────────────────────────────────────────────────────────────────────
// Settings
// ─────────────────────────────────────────────────────────────────────────

function loadSettings() {
  if (!fs.existsSync(SETTINGS_PATH)) {
    console.error(`❌  local.settings.json not found at ${SETTINGS_PATH}`);
    console.error('    Copy local.settings.json.example and populate the ID fields.');
    process.exit(1);
  }
  return JSON.parse(fs.readFileSync(SETTINGS_PATH, 'utf8'));
}

// ─────────────────────────────────────────────────────────────────────────
// .clasp.json
// ─────────────────────────────────────────────────────────────────────────

function writeClasp(scriptId) {
  fs.writeFileSync(CLASP_PATH, JSON.stringify({ scriptId, rootDir: 'script' }, null, 2) + '\n', 'utf8');
}

// ─────────────────────────────────────────────────────────────────────────
// version.js stamping
// ─────────────────────────────────────────────────────────────────────────

function stampVersion(label, options = {}) {
  const versionPath = options.versionPath || VERSION_PATH;
  const now = options.now || new Date().toISOString();
  const version = options.versionOverride
    || JSON.parse(fs.readFileSync(options.pkgPath || PKG_PATH, 'utf8')).version
    || '0.0.0';

  let src = fs.readFileSync(versionPath, 'utf8');

  src = replaceConst(src, 'APP_VERSION',       `'${version}'`);
  src = replaceConst(src, 'APP_VERSION_DATE',  `'${now}'`);
  src = replaceConst(src, 'APP_DEPLOY_TARGET', `'${label}'`);

  fs.writeFileSync(versionPath, src, 'utf8');
  console.log(`📝 version.js stamped: v${version}  ${now}  ${label}`);

  return { version, now, label };
}

/** Increments the patch segment of package.json's semver "version". PROD path only. */
function bumpPatchVersion_(pkgPath) {
  const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf8'));
  const parts = String(pkg.version || '0.0.0').split('.');
  const patch = (parseInt(parts[2], 10) || 0) + 1;
  const newVersion = `${parts[0]}.${parts[1]}.${patch}`;

  pkg.version = newVersion;
  fs.writeFileSync(pkgPath, JSON.stringify(pkg, null, 2) + '\n', 'utf8');

  return newVersion;
}

/** Increments package.json's integer "build". SIT path only. */
function bumpBuildNumber_(pkgPath) {
  const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf8'));
  const build = (parseInt(pkg.build, 10) || 0) + 1;

  pkg.build = build;
  fs.writeFileSync(pkgPath, JSON.stringify(pkg, null, 2) + '\n', 'utf8');

  return build;
}

/** Resets package.json's "build" counter to 0. Called unconditionally by the PROD path. */
function resetBuildNumber_(pkgPath) {
  const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf8'));
  pkg.build = 0;
  fs.writeFileSync(pkgPath, JSON.stringify(pkg, null, 2) + '\n', 'utf8');
}

/**
 * Replace the value of a `const NAME = <value>;` line.
 * If the const doesn't exist, it is appended before the trailing blank line/EOF.
 */
function replaceConst(src, name, value) {
  const re = new RegExp(`^(const ${name}\\s*=\\s*)([^;]+)(;)`, 'm');
  if (re.test(src)) {
    return src.replace(re, `$1${value}$3`);
  }
  return src.trimEnd() + `\nconst ${name.padEnd(18)} = ${value};\n`;
}

// ─────────────────────────────────────────────────────────────────────────
// Named deployment lookup
// ─────────────────────────────────────────────────────────────────────────

/**
 * Each script project is expected to carry exactly one active named deployment (excluding the
 * @HEAD test-deployment clasp always lists). Rather than storing its ID (which goes stale the
 * moment a deployment is recreated), look it up fresh from `clasp deployments` every time — run
 * after .clasp.json has been written for the target scriptId.
 */
function findActiveDeploymentId_(claspEnv) {
  const output = execSync('clasp deployments', { cwd: ROOT, env: claspEnv }).toString();
  const deploymentLines = output
    .split('\n')
    .map(line => line.trim())
    .filter(line => line.startsWith('-') && !line.includes('@HEAD'));

  if (deploymentLines.length === 0) {
    throw new Error('No active (non-@HEAD) deployment found — create a Web app deployment via the script editor first.');
  }
  if (deploymentLines.length > 1) {
    throw new Error(`Expected exactly one active deployment, found ${deploymentLines.length}:\n${deploymentLines.join('\n')}`);
  }

  const match = deploymentLines[0].match(/^-\s*(\S+)/);
  if (!match) {
    throw new Error(`Could not parse deployment ID from: ${deploymentLines[0]}`);
  }
  return match[1];
}

function saveDeploymentId_(targetKey, deploymentId) {
  const settings = JSON.parse(fs.readFileSync(SETTINGS_PATH, 'utf8'));
  settings[TARGETS[targetKey].deploymentIdKey] = deploymentId;
  fs.writeFileSync(SETTINGS_PATH, JSON.stringify(settings, null, 2) + '\n', 'utf8');
}

// ─────────────────────────────────────────────────────────────────────────
// Deploy verification (RECOMMENDATION.md §3.2 — assert the version actually serving)
// ─────────────────────────────────────────────────────────────────────────

function sleep_(ms) { return new Promise((resolve) => setTimeout(resolve, ms)); }

/**
 * Polls the deployment's cmd=version route (script/WebApp.js's handleVersionRequest_) until it
 * reports the exact version *and* target just stamped, or the timeout expires. This is what
 * turns "clasp deploy exited 0" into "the webapp is actually serving what we just pushed" — the
 * gap RECOMMENDATION.md §3.2 calls #13. The target half is what catches a deploy landing in the
 * wrong environment, which matters here because SIT, PROD and NUUC share one version counter:
 * a PROD-stamped build served from the SIT deployment would otherwise look identical.
 *
 * Dependency-injected (postFn/sleep) so the match/mismatch/timeout paths are unit-testable
 * without a real network call or a real wait — see test/test_assert_deployed_version.js.
 */
async function assertDeployedVersion_(deploymentId, expectedVersion, expectedTarget, options = {}) {
  const { postFn = postWebapp_, intervalSec = 5, timeoutSec = 60, sleep = sleep_, log = () => {} } = options;
  const url = `https://script.google.com/macros/s/${deploymentId}/exec?cmd=version`;
  const startedAt = Date.now();
  let attempt = 0;
  let lastResult = null;

  for (;;) {
    attempt++;
    try {
      lastResult = await postFn(url, { action: 'version' });
    } catch (err) {
      log(`  attempt ${attempt}: request failed (${err.message})`);
      lastResult = null;
    }

    if (lastResult && lastResult.ok && lastResult.version === expectedVersion && lastResult.target === expectedTarget) {
      return { ok: true, attempts: attempt, version: lastResult.version, target: lastResult.target, deploymentId: lastResult.deploymentId };
    }

    const seen = lastResult && typeof lastResult === 'object'
      ? `version=${lastResult.version || '(none)'} target=${lastResult.target || '(none)'}`
      : '(no response)';
    log(`  attempt ${attempt}: expected version=${expectedVersion} target=${expectedTarget}, got ${seen}`);

    if (Date.now() - startedAt + intervalSec * 1000 > timeoutSec * 1000) {
      throw new Error(
        `assertDeployedVersion_ timed out after ${attempt} attempts (${timeoutSec}s): ` +
        `expected version=${expectedVersion} target=${expectedTarget}, last seen ${seen}`
      );
    }
    await sleep(intervalSec * 1000);
  }
}

// ─────────────────────────────────────────────────────────────────────────
// Deploy
// ─────────────────────────────────────────────────────────────────────────

async function deploy(targetKey, options = {}) {
  const { scriptIdKey, label, emoji, claspAuthKey } = TARGETS[targetKey];
  const settings = loadSettings();
  const scriptId = settings[scriptIdKey];

  if (!scriptId || scriptId.startsWith('<')) {
    console.error(`❌  ${scriptIdKey} is not set in local.settings.json`);
    process.exit(1);
  }

  const claspAuthPath = resolveClaspAuthPath_(settings, claspAuthKey);
  const claspEnv = { ...process.env, clasp_config_auth: claspAuthPath };

  console.log(`\n${emoji}  Deploying to ${label} (${scriptId.slice(0, 12)}…)\n`);

  writeClasp(scriptId);
  console.log(`✅ .clasp.json written (rootDir: script, scriptId: ${scriptId.slice(0, 12)}…)`);

  // SIT bumps the build counter and leaves "version" alone; PROD and NUUC (both stable,
  // customer-facing deployments) bump the patch version and always reset build to 0.
  let version;
  if (targetKey === 'sit') {
    if (!options.skipBump) {
      const build = bumpBuildNumber_(PKG_PATH);
      console.log(`🔢 build number bumped to ${build}`);
    }
    const pkg = JSON.parse(fs.readFileSync(PKG_PATH, 'utf8'));
    version = `${pkg.version}.${pkg.build || 0}`;
  } else {
    if (!options.skipBump) {
      const bumped = bumpPatchVersion_(PKG_PATH);
      console.log(`🔢 package.json version bumped to v${bumped}`);
    }
    resetBuildNumber_(PKG_PATH);
    console.log('🔢 build counter reset to 0');
    version = JSON.parse(fs.readFileSync(PKG_PATH, 'utf8')).version;
  }

  stampVersion(label, { versionOverride: version });

  console.log(`\n🚀 Running: clasp push -f  (clasp_config_auth=${claspAuthPath})\n`);
  execSync('clasp push -f', { stdio: 'inherit', cwd: ROOT, env: claspEnv });
  console.log(`\n✅ ${label} push complete.`);

  console.log(`\n🔎 Looking up active deployment for ${label}…\n`);
  const deploymentId = findActiveDeploymentId_(claspEnv);
  console.log(`\n🌐 Updating named deployment ${deploymentId.slice(0, 12)}…\n`);
  // Captured (not 'inherit') so the revision number clasp reports (e.g. "...@8.") can be parsed
  // out for the deploy summary below; still echoed to stdout so nothing is lost from the console.
  const deployOutput = execSync(
    `clasp deploy --deploymentId ${deploymentId} --description "v${version} RCV"`,
    { cwd: ROOT, env: claspEnv }
  ).toString();
  process.stdout.write(deployOutput);
  const revisionMatch = deployOutput.match(/@(\d+)\b/);
  const revision = revisionMatch ? revisionMatch[1] : null;
  console.log(`\n✅ ${label} named deployment updated.`);

  saveDeploymentId_(targetKey, deploymentId);
  console.log(`💾 ${TARGETS[targetKey].deploymentIdKey} saved to local.settings.json`);

  // Stamp WEBAPP_URL on every deploy (not just PROD) — SIT and PROD are separate script
  // projects, each with their own WEBAPP_URL script property. onOpen.js's "Open Ballot Admin
  // Page" menu item and showAbout() read this property; without it (e.g. right after a SIT
  // deploy that's never been stamped), _getWebAppUrl_() falls back to
  // ScriptApp.getService().getUrl(), which only resolves correctly from inside an actual
  // running web app request — not from the spreadsheet-bound editor/menu context — so the
  // menu ends up with an empty or stale URL until this runs.
  console.log(`\n🔗 Setting WEBAPP_URL script property on ${label}…`);
  execSync(`node tools/callWebapp.js setWebappUrl --env ${targetKey}`, { stdio: 'inherit', cwd: ROOT });

  // Bootstrap ADMIN_SHARED_SECRET immediately, before this deployment's URL is ever shared —
  // bootstrapSecret is reachable by anyone on an ANYONE_ANONYMOUS deployment until a secret is
  // set, so closing that window here (rather than leaving it to a manual follow-up step) is
  // what keeps the race narrow. bootstrapAdminSecret_ refuses to overwrite an existing secret,
  // so re-running deploy on an already-bootstrapped target is a harmless no-op (exit code 1,
  // "already_bootstrapped" — not treated as a deploy failure).
  console.log(`\n🔐 Bootstrapping ADMIN_SHARED_SECRET on ${label} (no-op if already set)…`);
  try {
    execSync(`node tools/callWebapp.js bootstrapSecret --env ${targetKey}`, { stdio: 'inherit', cwd: ROOT });
  } catch (err) {
    console.log(`ℹ️  ${label} admin secret already bootstrapped (or bootstrap failed) — continuing.`);
  }

  // The static front end (static-pages/) shares this deploy's version/build counter
  // (build-static-pages.js's versionStringFor) — publish it as part of the same deploy rather
  // than as a separate step, so GAS and the static UI it serves never drift out of sync.
  // --skip-bump because deploy() already bumped/reset the counter above; publish-static-pages.js
  // just builds+publishes whatever package.json now says.
  console.log(`\n📄 Publishing static pages (${targetKey})…\n`);
  execSync(
    `node ${path.join(__dirname, 'publish-static-pages.js')} --env ${targetKey} --skip-bump`,
    { stdio: 'inherit', cwd: ROOT }
  );

  // Deploy verification (RECOMMENDATION.md §3.2) — the mandatory last step before the summary.
  // On mismatch the deploy fails loudly (non-zero exit, expected-vs-actual) but still prints the
  // summary, so the operator can see what *is* deployed rather than being left with an error.
  console.log(`\n🔍 Verifying ${label} is actually serving v${version}…`);
  let verified;
  try {
    verified = await assertDeployedVersion_(deploymentId, version, label, { log: console.log });
    console.log(`✅ ${label} verified — serving v${verified.version} (target ${verified.target})`);
  } catch (err) {
    console.error(`\n❌ Deploy verification failed: ${err.message}`);
    printDeploySummary_(targetKey, { version, revision, deploymentId, settings });
    process.exitCode = 1;
    return;
  }

  // Report what the server confirmed, not what was stamped locally (§3.2).
  printDeploySummary_(targetKey, { version: verified.version, revision, deploymentId, settings });
}

/** Prints the post-deploy summary: static entry point, spreadsheet, revision, and stamped version. */
function printDeploySummary_(targetKey, { version, revision, deploymentId, settings }) {
  const { label, emoji, sheetIdKey } = TARGETS[targetKey];
  const sheetId = settings[sheetIdKey];
  const staticUrl = STATIC_ENTRY_BASE_URL[label] || '(static hosting not configured for this target)';
  const sheetUrl = sheetId ? `https://docs.google.com/spreadsheets/d/${sheetId}/edit` : '(sheetId not set in local.settings.json)';

  console.log(`\n${emoji}  ${label} deploy summary`);
  console.log(`   Version:     v${version}`);
  console.log(`   Revision:    ${revision ? '@' + revision : '(could not be parsed from clasp deploy output)'}`);
  console.log(`   Static page: ${staticUrl}`);
  console.log(`   Spreadsheet: ${sheetUrl}`);
  console.log(`   Webapp:      https://script.google.com/macros/s/${deploymentId}/exec\n`);
}

// ─────────────────────────────────────────────────────────────────────────
// Interactive menu (no-flag invocation)
// ─────────────────────────────────────────────────────────────────────────

async function interactiveMenu() {
  let select;
  try {
    ({ select } = require('@inquirer/prompts'));
  } catch {
    console.error('❌  @inquirer/prompts is not installed. Run: npm install');
    process.exit(1);
  }

  const action = await select({
    message: 'Deploy target:',
    choices: [
      { name: '🧪 SIT   — push to sitScriptId (SIT stamp)',    value: 'sit'  },
      { name: '🚀 PROD  — push to prodScriptId (PROD stamp)',  value: 'prod' },
      { name: '⛪ NUUC  — push to nuucScriptId (NUUC stamp)',  value: 'nuuc' },
      { name: '❌ Exit',                                        value: 'exit' },
    ],
  });

  if (action !== 'exit') await deploy(action);
}

// ─────────────────────────────────────────────────────────────────────────
// Entry point
// ─────────────────────────────────────────────────────────────────────────

async function main() {
  const args = process.argv.slice(2);
  const options = { skipBump: args.includes('--skip-bump') };

  if (args.includes('--deploy-sit'))  return await deploy('sit', options);
  if (args.includes('--deploy-prod')) return await deploy('prod', options);
  if (args.includes('--deploy-nuuc')) return await deploy('nuuc', options);

  await interactiveMenu();
}

if (require.main === module) {
  main().catch(err => {
    if (err && (err.name === 'ExitPromptError' || err.message?.includes('force closed'))) {
      console.log('\n❌ Cancelled.');
      return;
    }
    console.error('❌ Error:', err.message);
    process.exit(1);
  });
}

module.exports = {
  replaceConst,
  stampVersion,
  bumpPatchVersion_,
  bumpBuildNumber_,
  resetBuildNumber_,
  printDeploySummary_,
  assertDeployedVersion_,
  TARGETS,
  STATIC_ENTRY_BASE_URL,
};
