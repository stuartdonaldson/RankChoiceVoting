/**
 * version.js — build/deploy stamp.
 *
 * These consts are rewritten by tools/manage-deployments.js on every deploy:
 *   APP_VERSION        — semver ("prod"/"nuuc") or `${version}.${build}` ("sit")
 *   APP_VERSION_DATE   — ISO timestamp of the deploy
 *   APP_DEPLOY_TARGET  — 'SIT', 'PROD', or 'NUUC'
 * Do not hand-edit; commit whatever the tool last stamped.
 */
const APP_VERSION       = '0.1.19.1';
const APP_VERSION_DATE  = '2026-11-20T23:18:40.039Z';
const APP_DEPLOY_TARGET = 'SIT';
const APP_AUTHOR        = 'Stuart Donaldson';
const APP_CONTACT       = 'stuart.donaldson@gmail.com';
const BUILD_INFO = {
  "version": "0.1.19.1",
  "buildDate": "2026-11-20T23:18:40.039Z",
  "env": "SIT",
  "webappUrl": "https://script.google.com/macros/s/AKfycbzKpwFfgRqHx1v3A0pvVAbUNAMab_A8rqOf0jqgRP5DPTmNNPU6_qelWHMG1QPkXbhP/exec",
  "staticUrl": "https://nuuc-it.github.io/Static/pub/ballot-sit/"
};
