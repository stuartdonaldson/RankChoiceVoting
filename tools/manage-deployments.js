#!/usr/bin/env node
/**
 * RankChoiceVoting deployment config. The pipeline itself lives in the shared `gas-deploy`
 * package (see its README, and best-practices/gas-deployment/RECOMMENDATION.md); this file is
 * only what is specific to this project.
 *
 *   pnpm run deploy:sit | deploy:prod | deploy:nuuc
 *   node tools/manage-deployments.js --summary --env sit    # read-only, deploys nothing
 *   pnpm run deployments                                   # every env's URLs + live versions, no clasp
 *
 * Each env is its own Apps Script project. Their IDs and spreadsheet IDs are committed project
 * truth in gas-project.json (GAS-Core adr/0004); local.settings.json carries only clasp auth,
 * admin secrets, API tokens, the static-host checkout paths and the deploymentId cache.
 */

'use strict';

const path = require('path');
const { execSync, execFileSync } = require('child_process');
const { runCli, constStamper, buildInfoStamper } = require('gas-deploy');
const ROOT = path.join(__dirname, '..');

/**
 * The static front end, published as part of the deploy rather than beside it.
 *
 * The hook chain (build -> publish -> assertPublishedBuild, each required, the publish chained)
 * and the summary row are gas-static's, not this project's — see that package's README §"Chaining
 * off a deploy". This used to be one hand-rolled `Publish static pages` hook shelling out to
 * tools/publish-static-pages.js, which read nothing back: the published page was never verified
 * against the deployment it was built for.
 *
 * `envs` in tools/static-pages.js are keyed by these target keys, so no envFor mapping is needed.
 */
const staticPipeline = require('./static-pages');

const call = (args) => execSync(`node tools/callWebapp.js ${args}`, { stdio: 'inherit', cwd: ROOT });

/**
 * Stamps script/version.js twice, in one pass, because two different readers need two shapes out
 * of the same file:
 *
 *   - the bare `const APP_VERSION / APP_VERSION_DATE / APP_DEPLOY_TARGET` globals, read across the
 *     GAS-concatenated scope (onOpen.js's About dialog, ApiBridge.js, WebApp.js's cmd=version);
 *   - a generated `BUILD_INFO` literal, which is the ONLY thing gas-static reads when it builds
 *     the static page.
 *
 * BUILD_INFO.webappUrl is the /exec URL of the deployment THIS run just resolved, which is why
 * `resolveBeforeStamp` is on below: without it the URL is not knowable until after the stamp and
 * the published page would point at the previous deploy. That round trip is the whole point of
 * GAS-Core adr/0001 — the old build constructed the URL from a deployment ID in
 * local.settings.json, which was reconciled against nothing.
 *
 * BUILD_INFO.staticUrl is the live page URL for this env, composed by tools/static-pages.js from
 * gas-project.json. Stamping it is what lets script/ApiBridge.js's _staticPagesBaseUrl_ stop
 * carrying its own hardcoded third copy of the same three URLs.
 */
function stampAll_(ctx) {
  constStamper({ file: 'script/version.js' })(ctx);
  return buildInfoStamper({
    file: 'script/version.js',
    // The GAS runtime and gas-static read these by name: BUILD_INFO.env is what gas-static asserts
    // against each env's declared deployTarget, and .webappUrl is the backend the page calls.
    fields: { date: 'buildDate', target: 'env', webAppUrl: 'webappUrl' },
    extraFields: (c) => ({ staticUrl: staticPipeline.liveUrl_(c.targetKey) }),
  })(ctx);
}

const config = {
  root: ROOT,
  stamper: stampAll_,
  // BUILD_INFO.webappUrl is not knowable until the deployment has been resolved, so the one
  // read-only `clasp deployments` call moves ahead of the stamp.
  resolveBeforeStamp: true,
  describeDeployment: (version) => `v${version} RCV`,

  // NUUC runs under a separate Google account, hence its own authKey; sit/prod take the default.
  // No scriptIdKey/sheetIdKey: those are declared in gas-project.json now, and leaving the legacy
  // *Key indirection declared beside it would be the second mechanism adr/0004 argues against.
  targets: {
    sit:  { label: 'SIT',  emoji: '🧪', counter: 'build',   deploymentIdKey: 'sitDeploymentId'  },
    prod: { label: 'PROD', emoji: '🚀', counter: 'version', deploymentIdKey: 'prodDeploymentId' },
    nuuc: { label: 'NUUC', emoji: '⛪', counter: 'version', deploymentIdKey: 'nuucDeploymentId', authKey: 'nuucAuth' },
  },

  // --summary only, display-only: lets a live-vs-local divergence be flagged. The deploy path
  // never reads this back — package.json stays the source of truth.
  readLocalVersion: () => {
    const src = require('fs').readFileSync(path.join(ROOT, 'script/version.js'), 'utf8');
    const g = (n) => (src.match(new RegExp(`const ${n}\\s*=\\s*'([^']+)'`)) || [])[1];
    return { version: g('APP_VERSION'), now: g('APP_VERSION_DATE') };
  },

  // Puts the published page's URL in the deploy summary block — on the verification-failure path
  // too, which is when "which page is actually live?" matters most.
  extraRows: staticPipeline.summaryRows(),

  postDeploy: [
    // Every deploy, not just PROD: each env is its own script project with its own WEBAPP_URL,
    // and onOpen.js's menu reads it — the ScriptApp.getService().getUrl() fallback only resolves
    // inside a live web app request, not from the spreadsheet-bound editor.
    { name: 'Stamp WEBAPP_URL', run: ({ targetKey }) => call(`setWebappUrl --env ${targetKey}`),
      retryCommand: 'node tools/callWebapp.js setWebappUrl --env <env>' },

    // Bootstrapped before this deployment's URL is ever shared: bootstrapSecret is reachable by
    // anyone on an ANYONE_ANONYMOUS deployment until a secret is set. Re-running exits non-zero
    // ("already_bootstrapped"), which is why this hook is not required.
    { name: 'Bootstrap ADMIN_SHARED_SECRET (no-op if already set)',
      run: ({ targetKey }) => call(`bootstrapSecret --env ${targetKey}`),
      retryCommand: 'node tools/callWebapp.js bootstrapSecret --env <env>' },

    // build -> publish -> assertPublishedBuild, all three required. The page ships in the same
    // deploy as the backend it calls, and the last step is the one that can tell a published page
    // from a merely propagated one.
    ...staticPipeline.deployHooks(),

    // Last, so only a deploy that verified end to end gets a release tag. That tag is what the
    // next release deploy diffs against to decide whether it may reuse this version.
    { name: 'Tag release (release targets only, no-op if tagged)',
      run: ({ target, version }) => {
        if (target.counter !== 'version') return;
        const git = gitRunner_();
        if (tagExists_(git, `v${version}`)) return console.log(`   v${version} already tagged`);
        git(['tag', '-a', `v${version}`, '-m', `Release v${version}`]);
        console.log(`   tagged v${version} (local; pushed with the next git push --follow-tags)`);
      },
      retryCommand: 'git tag -a v<version> -m "Release v<version>"' },
  ],
};

/**
 * Release-version reuse (rcballot-j56). PROD and NUUC are two deployments of one release, so a
 * back-to-back deploy to both must report one version — gas-deploy's `counter: 'version'` bumps
 * the patch on every deploy, which made identical code read 0.1.7 on PROD and 0.1.8 on NUUC.
 *
 * Prototyped here ahead of a `counter: 'release'` mode in gas-deploy itself: the decision is made
 * before runCli and expressed through its existing `--skip-bump` flag, so no package change is
 * needed yet. Only the explicit `--deploy-<key>` path is covered; the interactive menu still bumps.
 *
 * The files the deploy itself writes are excluded from both the clean-tree check and the diff,
 * as is .beads/ (bd's export churns on every issue edit and never ships).
 */
const STAMPED_EXCLUDES = [':!script/version.js', ':!package.json', ':!.beads'];

const gitRunner_ = () => (args) => execFileSync('git', args, { cwd: ROOT, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });

function tagExists_(git, tag) {
  try { git(['rev-parse', '-q', '--verify', `refs/tags/${tag}`]); return true; } catch { return false; }
}

/** package.json minus the two fields the deploy stamps, so a dependency pin change still counts. */
function pkgSansCounters_(pkg) {
  const { version, build, ...rest } = pkg;
  return JSON.stringify(rest);
}

/** Throws if anything but the stamped files is uncommitted: a reused version must name a commit. */
function assertCleanTree_(git) {
  const dirty = git(['status', '--porcelain', '--', '.', ...STAMPED_EXCLUDES]).trim();
  if (dirty) {
    throw new Error(`release deploy refused: uncommitted changes outside the stamped files:\n${dirty}\nCommit or stash them first.`);
  }
}

/** { skipBump, reason } for a release target at package.json `version`. */
function decideBump_({ version, currentPkg, git }) {
  const tag = `v${version}`;
  if (!tagExists_(git, tag)) return { skipBump: false, reason: `no tag ${tag}` };

  const changed = git(['diff', '--name-only', tag, '--', '.', ...STAMPED_EXCLUDES]).split('\n').filter(Boolean);
  if (changed.length) {
    const shown = changed.slice(0, 5).join(', ') + (changed.length > 5 ? ', …' : '');
    return { skipBump: false, reason: `${changed.length} file(s) changed since ${tag}: ${shown}` };
  }

  const tagged = JSON.parse(git(['show', `${tag}:package.json`]));
  if (pkgSansCounters_(tagged) !== pkgSansCounters_(currentPkg)) {
    return { skipBump: false, reason: `package.json changed since ${tag} (beyond version/build)` };
  }
  return { skipBump: true, reason: `no code change since ${tag}` };
}

/** argv for runCli: a release target gains --skip-bump when decideBump_ says the version holds. */
function releaseArgv_(argv, { targets, currentPkg, git, log }) {
  const targetKey = Object.keys(targets).find(k => argv.includes(`--deploy-${k}`));
  if (!targetKey || targets[targetKey].counter !== 'version') return argv;

  assertCleanTree_(git);
  if (argv.includes('--skip-bump')) {
    log(`🔢 --skip-bump given: deploying v${currentPkg.version} as is`);
    return argv;
  }
  const { skipBump, reason } = decideBump_({ version: currentPkg.version, currentPkg, git });
  log(skipBump ? `🔢 reusing v${currentPkg.version} — ${reason}` : `🔢 bumping patch — ${reason}`);
  return skipBump ? [...argv, '--skip-bump'] : argv;
}

if (require.main === module) {
  let argv;
  try {
    argv = releaseArgv_(process.argv, {
      targets: config.targets,
      currentPkg: JSON.parse(require('fs').readFileSync(path.join(ROOT, 'package.json'), 'utf8')),
      git: gitRunner_(),
      log: console.log,
    });
  } catch (err) {
    console.error('❌ Error:', err.message);
    process.exit(1);
  }
  runCli(config, argv).catch(err => {
    if (err && (err.name === 'ExitPromptError' || (err.message || '').includes('force closed'))) return console.log('\n❌ Cancelled.');
    console.error('❌ Error:', err.message);
    process.exit(1);
  });
}

module.exports = { config, stampAll_, decideBump_, assertCleanTree_, releaseArgv_, STAMPED_EXCLUDES };
