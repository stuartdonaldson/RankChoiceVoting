/**
 * version.js — build/deploy stamp.
 *
 * These consts are rewritten by tools/manage-deployments.js on every deploy:
 *   APP_VERSION        — semver ("prod"/"nuuc") or `${version}.${build}` ("sit")
 *   APP_VERSION_DATE   — ISO timestamp of the deploy
 *   APP_DEPLOY_TARGET  — 'SIT', 'PROD', or 'NUUC'
 * Do not hand-edit; commit whatever the tool last stamped.
 */
const APP_VERSION       = '0.0.3';
const APP_VERSION_DATE  = '2026-07-07T01:11:02.896Z';
const APP_DEPLOY_TARGET = 'NUUC';
const APP_AUTHOR        = 'Stuart Donaldson';
const APP_CONTACT       = 'stuart.donaldson@gmail.com';
