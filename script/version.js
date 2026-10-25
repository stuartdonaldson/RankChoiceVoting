/**
 * version.js — build/deploy stamp.
 *
 * These consts are rewritten by tools/manage-deployments.js on every deploy:
 *   APP_VERSION        — semver ("prod"/"nuuc") or `${version}.${build}` ("sit")
 *   APP_VERSION_DATE   — ISO timestamp of the deploy
 *   APP_DEPLOY_TARGET  — 'SIT', 'PROD', or 'NUUC'
 * Do not hand-edit; commit whatever the tool last stamped.
 */
const APP_VERSION       = '0.1.8';
const APP_VERSION_DATE  = '2026-10-25T15:50:09.116Z';
const APP_DEPLOY_TARGET = 'NUUC';
const APP_AUTHOR        = 'Stuart Donaldson';
const APP_CONTACT       = 'stuart.donaldson@gmail.com';
const BUILD_INFO = {
  "version": "0.1.8",
  "buildDate": "2026-10-25T15:50:09.116Z",
  "env": "NUUC",
  "webappUrl": "https://script.google.com/macros/s/AKfycbyCC3jrN0_wVTivKkjC2COfeyqaopvaPlpKUEUdu18D9UsYa7llEXnemGRIzwAGXhw4fA/exec",
  "staticUrl": "https://nuuc-it.github.io/Static/pub/ballot/"
};
