/**
 * BallotCache.js
 *
 * Script-properties cache of each ballot's "static" content — config (Title,
 * Description, Instructions, Footer, Contact, Accept-New, Add-Instructions,
 * Admin-Only-Notes) and the Candidates table (name/details/linkText/linkUrl) — so a
 * respondent or admin page load that only needs this content never has to open the
 * spreadsheet at all. Opening/reading the bound sheet is the dominant cost of every
 * RPC in this app, which is what makes the static-pages front end feel slow.
 *
 * Response rows (who voted, their ranking) are deliberately NOT cached here — they
 * change on every submission, so they're still read straight from the sheet.
 *
 * Kept fresh three ways:
 *   1. Every write path that touches a ballot's config or candidates rebuilds that
 *      ballot's cache entry immediately after writing — see refreshBallotCache_'s
 *      callers in BallotModel.js/webAdmin.js/webBallot.js.
 *   2. onEdit(e) (Triggers.js — a simple trigger, so no installable-trigger setup is
 *      needed) rebuilds the cache whenever someone hand-edits a Ballot-<id> sheet
 *      directly in Sheets, so a manual edit never leaves the cache stale.
 *   3. A cache miss (key never written, or Properties lost it) transparently falls
 *      back to reading the sheet and populating the cache — see getCachedBallotData_.
 *
 * Script Properties has a 9KB-per-value / 500KB-total quota (see
 * https://developers.google.com/apps-script/guides/properties#quotas) — comfortably
 * enough for this app's ballot sizes, but an unusually large candidate list/details
 * blob could exceed the per-value limit; refreshBallotCache_ catches that (and any
 * other Properties failure) and leaves the ballot uncached rather than throwing, so
 * every read path still works correctly via the sheet fallback — just not sped up.
 */

var _BALLOT_CACHE_KEY_PREFIX_ = 'ballotCache:';

/**
 * @param {string} id
 * @return {string}
 */
function _ballotCacheKey_(id) {
  return _BALLOT_CACHE_KEY_PREFIX_ + String(id || '').trim();
}

/**
 * @param {string} id
 * @return {{config:Object, candidates:Array<{name:string,details:string,linkText:string,linkUrl:string}>, updatedAt:string}|null}
 */
function _readBallotCacheRaw_(id) {
  var raw = PropertiesService.getScriptProperties().getProperty(_ballotCacheKey_(id));
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (err) {
    return null; // corrupt/legacy value — treat as a miss, next rebuild overwrites it
  }
}

/**
 * Reads config + candidates fresh from `sheet` and writes them into the cache,
 * overwriting whatever was there. Call this right after any write to a ballot's
 * config or Candidates table so the cache never lags behind what was just saved.
 *
 * @param {string} id
 * @param {Sheet} sheet
 * @return {{config:Object, candidates:Array, updatedAt:string}}
 */
function refreshBallotCache_(id, sheet) {
  // readBallotCandidates_ runs the Candidates<->Responses-header reconciliation (appending
  // any Responses column a Candidates-table entry is still missing) — read it first so the
  // detail rows read right after reflect a fully reconciled sheet. Its return value (plain
  // names, Responses-header order) also backfills any candidate that predates the
  // Candidates table with a blank-detail entry, matching what getAdminEditData/
  // getBallotForRespondent_ used to backfill inline before this cache existed.
  var candidateNames = readBallotCandidates_(sheet);
  var candidateRows = readBallotCandidateDetails_(sheet);
  while (candidateRows.length < candidateNames.length) {
    candidateRows.push({ name: candidateNames[candidateRows.length], details: '', linkText: '', linkUrl: '' });
  }

  var data = {
    config: readBallotConfig_(sheet),
    candidates: candidateRows,
    updatedAt: new Date().toISOString()
  };
  try {
    PropertiesService.getScriptProperties().setProperty(_ballotCacheKey_(id), JSON.stringify(data));
  } catch (err) {
    // Over quota, or some other Properties failure — leave the ballot uncached. Every
    // read path falls back to the sheet on a miss, so correctness is unaffected; this
    // ballot just doesn't get the speedup until a future write succeeds in caching it.
    Logger.log('Could not cache ballot "' + id + '": ' + err);
  }
  return data;
}

/**
 * Returns cached config+candidates for a ballot, transparently rebuilding (and
 * re-caching) from the sheet on a miss. Pass `sheet` when the caller already has it
 * open (e.g. right after a write, or once it's already been looked up for another
 * reason) to avoid a redundant lookup; omit it to let this function look the sheet up
 * itself — SpreadsheetApp.getActiveSpreadsheet() is only touched when the cache
 * doesn't already have the answer, which is the whole point of this module.
 *
 * @param {string} id
 * @param {Sheet=} sheet
 * @return {{config:Object, candidates:Array}|null} null if the ballot doesn't exist
 *   (only possible when `sheet` isn't passed and no sheet is found for `id`).
 */
function getCachedBallotData_(id, sheet) {
  var cached = _readBallotCacheRaw_(id);
  if (cached) return cached;
  if (!sheet) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    sheet = findBallotSheet_(ss, id);
    if (!sheet) return null;
  }
  return refreshBallotCache_(id, sheet);
}

/**
 * Clears one ballot's cache entry. Not currently wired to any caller (there's no
 * ballot-delete feature yet) — kept for completeness/future use, e.g. if a delete
 * path is ever added.
 *
 * @param {string} id
 */
function invalidateBallotCache_(id) {
  PropertiesService.getScriptProperties().deleteProperty(_ballotCacheKey_(id));
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    _ballotCacheKey_: _ballotCacheKey_,
    _readBallotCacheRaw_: _readBallotCacheRaw_,
    refreshBallotCache_: refreshBallotCache_,
    getCachedBallotData_: getCachedBallotData_,
    invalidateBallotCache_: invalidateBallotCache_
  };
}
