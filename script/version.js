/**
 * version.js — build/deploy stamp.
 *
 * These consts are rewritten by tools/manage-deployments.js on every deploy:
 *   APP_VERSION        — semver ("prod") or `${version}.${build}` ("sit")
 *   APP_VERSION_DATE   — ISO timestamp of the deploy
 *   APP_DEPLOY_TARGET  — 'SIT' or 'PROD'
 * Do not hand-edit; commit whatever the tool last stamped.
 */
const APP_VERSION       = '0.0.1.23';
const APP_VERSION_DATE  = '2026-07-06T20:27:40.425Z';
const APP_DEPLOY_TARGET = 'SIT';
const APP_AUTHOR        = 'Stuart Donaldson';
const APP_CONTACT       = 'stuart.donaldson@gmail.com';
