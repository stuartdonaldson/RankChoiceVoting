'use strict';

const assert = require('node:assert/strict');
const { decideBump_, assertCleanTree_, releaseArgv_, STAMPED_EXCLUDES } = require('../tools/manage-deployments.js');

// rcballot-j56: a release deploy (PROD, NUUC) reuses the package.json version when nothing but
// the stamped files has changed since the v<version> tag, so back-to-back deploys of one release
// report one version. git is injected so this suite never shells out.

const PKG = { name: 'rankchoicevoting', version: '0.1.8', build: 0, dependencies: { 'gas-deploy': 'x' } };

/** A fake git: `responses` maps the joined argv to stdout, or to an Error to throw. */
function fakeGit_(responses) {
  const calls = [];
  const git = (args) => {
    const key = args.join(' ');
    calls.push(key);
    if (!(key in responses)) throw new Error(`unexpected git call: ${key}`);
    const r = responses[key];
    if (r instanceof Error) throw r;
    return r;
  };
  git.calls = calls;
  return git;
}

const TAG = 'refs/tags/v0.1.8';
const verify = `rev-parse -q --verify ${TAG}`;
const diff = ['diff', '--name-only', 'v0.1.8', '--', '.', ...STAMPED_EXCLUDES].join(' ');
const show = 'show v0.1.8:package.json';
const status = ['status', '--porcelain', '--', '.', ...STAMPED_EXCLUDES].join(' ');

// AC 1 — tag present, no source change, package.json differs only in version/build: reuse.
{
  const git = fakeGit_({
    [verify]: 'abc123\n',
    [diff]: '',
    [show]: JSON.stringify({ ...PKG, version: '0.1.7', build: 5 }),
  });
  const d = decideBump_({ version: '0.1.8', currentPkg: PKG, git });
  assert.equal(d.skipBump, true);
  assert.match(d.reason, /no code change since v0\.1\.8/);
}

// AC 1 — no tag: bump.
{
  const git = fakeGit_({ [verify]: new Error('exit 1') });
  const d = decideBump_({ version: '0.1.8', currentPkg: PKG, git });
  assert.equal(d.skipBump, false);
  assert.match(d.reason, /no tag v0\.1\.8/);
}

// AC 1 — source changed since the tag: bump, and say which files.
{
  const git = fakeGit_({ [verify]: 'abc\n', [diff]: 'script/WebApp.js\ntools/static-pages.js\n' });
  const d = decideBump_({ version: '0.1.8', currentPkg: PKG, git });
  assert.equal(d.skipBump, false);
  assert.match(d.reason, /2 file\(s\) changed since v0\.1\.8/);
  assert.match(d.reason, /script\/WebApp\.js/);
}

// AC 1 — package.json changed beyond version/build (a dependency pin moved): bump.
{
  const git = fakeGit_({
    [verify]: 'abc\n',
    [diff]: '',
    [show]: JSON.stringify({ ...PKG, dependencies: { 'gas-deploy': 'older' } }),
  });
  const d = decideBump_({ version: '0.1.8', currentPkg: PKG, git });
  assert.equal(d.skipBump, false);
  assert.match(d.reason, /package\.json changed since v0\.1\.8/);
}

// AC 2 — clean tree passes; a change outside the stamped files throws, naming it.
{
  assertCleanTree_(fakeGit_({ [status]: '' }));
  assert.throws(
    () => assertCleanTree_(fakeGit_({ [status]: ' M script/WebApp.js\n?? scratch.txt\n' })),
    /uncommitted changes[\s\S]*script\/WebApp\.js[\s\S]*scratch\.txt/
  );
}

// AC 1, 3, 5 — argv rewriting: release targets gain --skip-bump only on reuse; SIT and the
// interactive menu (no --deploy-* flag) are passed through untouched without touching git.
{
  const targets = {
    sit:  { counter: 'build' },
    prod: { counter: 'version' },
    nuuc: { counter: 'version' },
  };
  const logs = [];
  const log = (m) => logs.push(m);
  const reuse = () => fakeGit_({ [status]: '', [verify]: 'a\n', [diff]: '', [show]: JSON.stringify(PKG) });

  const base = ['node', 'manage-deployments.js'];
  assert.deepEqual(releaseArgv_([...base, '--deploy-prod'], { targets, currentPkg: PKG, git: reuse(), log }),
    [...base, '--deploy-prod', '--skip-bump']);
  assert.ok(logs.some(m => /reusing v0\.1\.8/.test(m)), 'decision is logged');

  const bumpGit = fakeGit_({ [status]: '', [verify]: new Error('no tag') });
  assert.deepEqual(releaseArgv_([...base, '--deploy-nuuc'], { targets, currentPkg: PKG, git: bumpGit, log }),
    [...base, '--deploy-nuuc']);
  assert.ok(logs.some(m => /bumping patch/.test(m)));

  const untouched = fakeGit_({});
  assert.deepEqual(releaseArgv_([...base, '--deploy-sit'], { targets, currentPkg: PKG, git: untouched, log }),
    [...base, '--deploy-sit']);
  assert.deepEqual(releaseArgv_(base, { targets, currentPkg: PKG, git: untouched, log }), base);
  assert.deepEqual(untouched.calls, []);

  // An explicit --skip-bump is the operator's call: honoured, not duplicated, still clean-checked.
  const explicit = fakeGit_({ [status]: '' });
  assert.deepEqual(releaseArgv_([...base, '--deploy-prod', '--skip-bump'], { targets, currentPkg: PKG, git: explicit, log }),
    [...base, '--deploy-prod', '--skip-bump']);
}

console.log('test_release_version: all assertions passed');
