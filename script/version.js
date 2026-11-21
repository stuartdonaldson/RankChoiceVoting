/**
 * version.js — build/deploy stamp.
 *
 * These consts are rewritten by tools/manage-deployments.js on every deploy:
 *   APP_VERSION        — semver ("prod"/"nuuc") or `${version}.${build}` ("sit")
 *   APP_VERSION_DATE   — ISO timestamp of the deploy
 *   APP_DEPLOY_TARGET  — 'SIT', 'PROD', or 'NUUC'
 * Do not hand-edit; commit whatever the tool last stamped.
 */
const APP_VERSION       = '0.1.21';
const APP_VERSION_DATE  = '2026-11-21T02:44:11.750Z';
const APP_DEPLOY_TARGET = 'PROD';
const APP_AUTHOR        = 'Stuart Donaldson';
const APP_CONTACT       = 'stuart.donaldson@gmail.com';
const BUILD_INFO = {
  "version": "0.1.21",
  "buildDate": "2026-11-21T02:44:11.750Z",
  "env": "PROD",
  "webappUrl": "https://script.google.com/macros/s/AKfycbyOgscq50xGNorLGpQlNFbNcTGozG3FrxWOBimtYoK3vhuETaufoh6-RLWKvWCnnNj2/exec",
  "staticUrl": "https://f3go30.github.io/static-pages/ballot/prod/"
};
