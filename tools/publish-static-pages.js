#!/usr/bin/env node
/**
 * Publishes the built static-pages/dist/<env>/ output to the sibling static-hosting repo that
 * env's GitHub Pages site actually serves. Adapted from F3Go30's tools/publish-static-pages.js,
 * generalized because RankChoiceVoting's three targets don't all publish to the same repo or
 * the same layout within it (see ApiBridge.js's _staticPagesBaseUrl_ for the resulting URLs):
 *
 *   sit/prod -> ../F3Static (local.settings.json's staticPagesRepoPath, f3go30/static-pages) —
 *               namespaced under ballot/<env>/ alongside F3Go30's own dist/<env>/, since this
 *               repo already hosts one product's static front end and we're adding a second
 *               rather than standing up a whole new repo for it.
 *   nuuc     -> ../Static (local.settings.json's nuucStaticPagesRepoPath, nuuc-it/Static) — a
 *               single pub/ballot/ folder (no sit/prod split; NUUC is one environment),
 *               matching that repo's existing multi-app convention (pub/AS, pub/AS-sit, ...).
 *
 * Normally invoked automatically, once per target, as the last step of
 * tools/manage-deployments.js's deploy() (npm run deploy:sit / :prod / :nuuc) — with
 * --skip-bump, since deploy() already bumped/reset the counter for this push. Direct invocation
 * exists only for recovery — e.g. retrying a publish whose git push failed after a deploy()
 * already ran (pass --skip-bump then too).
 *
 * Runs tools/build-static-pages.js first (so dist/ is never stale), then copies the requested
 * env's folder into its destination repo, and commits + pushes from that repo.
 *
 * Usage:
 *   node tools/publish-static-pages.js [--env sit|prod|nuuc|all] [--skip-bump]
 */
const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const { bumpBuildNumber_ } = require('./manage-deployments.js');

const ROOT = path.resolve(__dirname, '..');
const DIST_ROOT = path.join(ROOT, 'static-pages', 'dist');
const SETTINGS_PATH = path.join(ROOT, 'local.settings.json');
const PKG_PATH = path.join(ROOT, 'package.json');

// env -> {repoSettingKey, relDest} — repoSettingKey names the local.settings.json key holding
// that env's destination repo checkout path; relDest is where within that repo this env's
// build lands (see file header for why sit/prod and nuuc differ).
const DEST_MAP = {
  sit: { repoSettingKey: 'staticPagesRepoPath', relDest: path.join('ballot', 'sit') },
  prod: { repoSettingKey: 'staticPagesRepoPath', relDest: path.join('ballot', 'prod') },
  nuuc: { repoSettingKey: 'nuucStaticPagesRepoPath', relDest: path.join('pub', 'ballot') },
};

function loadSettings() {
  if (!fs.existsSync(SETTINGS_PATH)) {
    console.error(`❌  local.settings.json not found at ${SETTINGS_PATH}`);
    process.exit(1);
  }
  return JSON.parse(fs.readFileSync(SETTINGS_PATH, 'utf8'));
}

function copyDir(src, dest) {
  fs.rmSync(dest, { recursive: true, force: true });
  fs.mkdirSync(dest, { recursive: true });
  for (const entry of fs.readdirSync(src, { withFileTypes: true })) {
    const srcPath = path.join(src, entry.name);
    const destPath = path.join(dest, entry.name);
    if (entry.isDirectory()) {
      copyDir(srcPath, destPath);
    } else {
      fs.copyFileSync(srcPath, destPath);
    }
  }
}

function publishOne(env, settings) {
  const { repoSettingKey, relDest } = DEST_MAP[env];
  const repoPathSetting = settings[repoSettingKey];
  if (!repoPathSetting) {
    console.error(`❌  ${repoSettingKey} is not set in local.settings.json`);
    process.exit(1);
  }
  const repoRoot = path.resolve(ROOT, repoPathSetting);
  if (!fs.existsSync(path.join(repoRoot, '.git'))) {
    console.error(`❌  ${repoRoot} does not look like a git checkout (no .git found)`);
    process.exit(1);
  }

  const src = path.join(DIST_ROOT, env);
  const dest = path.join(repoRoot, relDest);
  if (!fs.existsSync(src)) {
    console.error(`❌  ${src} not found — build did not produce it`);
    process.exit(1);
  }
  copyDir(src, dest);
  console.log(`📦 copied static-pages/dist/${env} -> ${path.relative(ROOT, dest)}`);

  // Scoped to just this env's relDest — an unscoped `git add`/status would let one env's
  // publish commit and push whatever happened to be sitting modified elsewhere in that repo
  // (e.g. a half-finished publish for a different app sharing the same static host).
  const status = execFileSync('git', ['status', '--porcelain', '--', relDest], { cwd: repoRoot }).toString().trim();
  if (!status) {
    console.log(`✅ ${path.relative(ROOT, dest)} already matches build output — nothing to publish.`);
    return;
  }

  const pkg = JSON.parse(fs.readFileSync(PKG_PATH, 'utf8'));
  execFileSync('git', ['add', relDest], { cwd: repoRoot, stdio: 'inherit' });
  const message = `Publish RankChoiceVoting static pages v${pkg.version}.${pkg.build || 0} (${env})`;
  execFileSync('git', ['commit', '-m', message], { cwd: repoRoot, stdio: 'inherit' });
  execFileSync('git', ['push'], { cwd: repoRoot, stdio: 'inherit' });
  console.log(`🚀 Published ${env} and pushed (${repoRoot}).`);
}

function main() {
  const args = process.argv.slice(2);
  const envIdx = args.indexOf('--env');
  const env = envIdx !== -1 ? args[envIdx + 1] : 'all';
  const envs = env === 'all' ? ['sit', 'prod', 'nuuc'] : [env];

  const settings = loadSettings();

  if (!args.includes('--skip-bump')) {
    const build = bumpBuildNumber_(PKG_PATH);
    console.log(`🔢 build number bumped to ${build}`);
  }

  // Build only the env(s) being published — building `all` during a single-env publish would
  // require every OTHER env's deployment ID to be configured too, and fail the publish if it
  // isn't (see build-static-pages.js's execUrlForEnv_).
  console.log('🔨 Building static pages...');
  execFileSync('node', [path.join(__dirname, 'build-static-pages.js'), '--env', env], { cwd: ROOT, stdio: 'inherit' });

  envs.forEach((e) => publishOne(e, settings));
}

if (require.main === module) {
  main();
}

module.exports = { copyDir, DEST_MAP };
