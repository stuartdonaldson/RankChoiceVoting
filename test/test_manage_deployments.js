'use strict';

const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');

const { replaceConst, stampVersion, bumpPatchVersion_, bumpBuildNumber_, resetBuildNumber_ } = require('../tools/manage-deployments');

function withTempDir(fn) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rcv-deploy-test-'));
  try {
    fn(dir);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

function testReplaceConstAppendsWhenMissing() {
  const src = "const APP_VERSION = '1.0.0';\n";
  const out = replaceConst(src, 'APP_DEPLOY_TARGET', "'SIT'");
  assert.ok(out.includes('const APP_DEPLOY_TARGET'));
  assert.ok(out.includes("'SIT'"));
}

function testStampVersionUpdatesAllFields() {
  withTempDir((dir) => {
    const pkgPath = path.join(dir, 'package.json');
    const versionPath = path.join(dir, 'version.js');

    fs.writeFileSync(pkgPath, JSON.stringify({ version: '0.0.1' }), 'utf8');
    fs.writeFileSync(
      versionPath,
      [
        "const APP_VERSION = '0.0.0';",
        "const APP_VERSION_DATE = '2000-01-01T00:00:00.000Z';",
        "const APP_DEPLOY_TARGET = 'SIT';",
        '',
      ].join('\n'),
      'utf8'
    );

    const targets = ['SIT', 'PROD'];

    for (const target of targets) {
      stampVersion(target, {
        pkgPath,
        versionPath,
        now: '2026-07-05T12:34:56.000Z',
      });

      const out = fs.readFileSync(versionPath, 'utf8');
      assert.ok(out.includes("const APP_VERSION = '0.0.1';"));
      assert.ok(out.includes("const APP_VERSION_DATE = '2026-07-05T12:34:56.000Z';"));
      assert.ok(out.includes(`const APP_DEPLOY_TARGET = '${target}';`));
    }
  });
}

function testStampVersionUsesVersionOverride() {
  withTempDir((dir) => {
    const pkgPath = path.join(dir, 'package.json');
    const versionPath = path.join(dir, 'version.js');

    fs.writeFileSync(pkgPath, JSON.stringify({ version: '0.0.1', build: 4 }), 'utf8');
    fs.writeFileSync(versionPath, "const APP_VERSION = '0.0.0';\nconst APP_VERSION_DATE = '';\nconst APP_DEPLOY_TARGET = '';\n", 'utf8');

    const { version } = stampVersion('SIT', {
      pkgPath,
      versionPath,
      now: '2026-07-05T12:34:56.000Z',
      versionOverride: '0.0.1.4',
    });

    assert.equal(version, '0.0.1.4');
    const out = fs.readFileSync(versionPath, 'utf8');
    assert.ok(out.includes("const APP_VERSION = '0.0.1.4';"));
  });
}

function testBumpPatchVersionIncrementsPatchOnly() {
  withTempDir((dir) => {
    const pkgPath = path.join(dir, 'package.json');
    fs.writeFileSync(pkgPath, JSON.stringify({ name: 'rankchoicevoting', version: '0.0.1' }, null, 2) + '\n', 'utf8');

    const newVersion = bumpPatchVersion_(pkgPath);

    assert.equal(newVersion, '0.0.2');
    const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf8'));
    assert.equal(pkg.version, '0.0.2');
    assert.equal(pkg.name, 'rankchoicevoting'); // other fields untouched
  });
}

function testBumpPatchVersionIsIdempotentAcrossCalls() {
  withTempDir((dir) => {
    const pkgPath = path.join(dir, 'package.json');
    fs.writeFileSync(pkgPath, JSON.stringify({ version: '0.0.1' }), 'utf8');

    bumpPatchVersion_(pkgPath);
    bumpPatchVersion_(pkgPath);
    const newVersion = bumpPatchVersion_(pkgPath);

    assert.equal(newVersion, '0.0.4');
  });
}

function testBumpPatchVersionDoesNotTouchBuild() {
  withTempDir((dir) => {
    const pkgPath = path.join(dir, 'package.json');
    fs.writeFileSync(pkgPath, JSON.stringify({ version: '0.0.1', build: 7 }), 'utf8');

    bumpPatchVersion_(pkgPath);

    const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf8'));
    assert.equal(pkg.build, 7);
  });
}

function testBumpBuildNumberIncrementsFromZeroWhenMissing() {
  withTempDir((dir) => {
    const pkgPath = path.join(dir, 'package.json');
    fs.writeFileSync(pkgPath, JSON.stringify({ version: '0.0.1' }), 'utf8');

    const build = bumpBuildNumber_(pkgPath);

    assert.equal(build, 1);
    const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf8'));
    assert.equal(pkg.build, 1);
    assert.equal(pkg.version, '0.0.1'); // version untouched by a SIT build bump
  });
}

function testBumpBuildNumberIsIdempotentAcrossCalls() {
  withTempDir((dir) => {
    const pkgPath = path.join(dir, 'package.json');
    fs.writeFileSync(pkgPath, JSON.stringify({ version: '0.0.1', build: 0 }), 'utf8');

    bumpBuildNumber_(pkgPath);
    bumpBuildNumber_(pkgPath);
    const build = bumpBuildNumber_(pkgPath);

    assert.equal(build, 3);
  });
}

function testResetBuildNumberZeroesExistingCount() {
  withTempDir((dir) => {
    const pkgPath = path.join(dir, 'package.json');
    fs.writeFileSync(pkgPath, JSON.stringify({ version: '0.0.1', build: 12 }), 'utf8');

    resetBuildNumber_(pkgPath);

    const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf8'));
    assert.equal(pkg.build, 0);
    assert.equal(pkg.version, '0.0.1'); // version untouched by a build reset
  });
}

function run() {
  testReplaceConstAppendsWhenMissing();
  testStampVersionUpdatesAllFields();
  testStampVersionUsesVersionOverride();
  testBumpPatchVersionIncrementsPatchOnly();
  testBumpPatchVersionIsIdempotentAcrossCalls();
  testBumpPatchVersionDoesNotTouchBuild();
  testBumpBuildNumberIncrementsFromZeroWhenMissing();
  testBumpBuildNumberIsIdempotentAcrossCalls();
  testResetBuildNumberZeroesExistingCount();
  console.log('test_manage_deployments: all tests passed');
}

run();
