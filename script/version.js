/**
 * version.js — build/deploy stamp.
 *
 * These consts are rewritten by tools/manage-deployments.js on every deploy:
 *   APP_VERSION        — semver ("prod"/"nuuc") or `${version}.${build}` ("sit")
 *   APP_VERSION_DATE   — ISO timestamp of the deploy
 *   APP_DEPLOY_TARGET  — 'SIT', 'PROD', or 'NUUC'
 * Do not hand-edit; commit whatever the tool last stamped.
 */
const APP_VERSION       = '0.1.18.1';
const APP_VERSION_DATE  = '2026-11-20T05:42:42.197Z';
const APP_DEPLOY_TARGET = 'SIT';
const APP_AUTHOR        = 'Stuart Donaldson';
const APP_CONTACT       = 'stuart.donaldson@gmail.com';
const BUILD_INFO = {
  "version": "0.1.18.1",
  "buildDate": "2026-11-20T05:42:42.197Z",
  "env": "SIT",
  "webappUrl": "https://script.google.com/macros/s/AKfycbwRGVywtwcP9zAS2HvOJDlgBOa7t_H6l98yKBhR4fWzacDRvAg62fd5HFdhQ97C2Ef7uA/exec",
  "staticUrl": "https://f3go30.github.io/static-pages/ballot/sit/"
};
