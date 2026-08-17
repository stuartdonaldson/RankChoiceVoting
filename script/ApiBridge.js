/**
 * ApiBridge.js
 *
 * Generic JSON-RPC bridge for the static front end (static-pages/src/index.html), reached via
 * doPost ?cmd=api (WebApp.js). Mirrors google.script.run's call shape ({name, args}) so porting
 * the old HtmlService pages (webAdminEditPage.html, webBallotPage.html) to fetch() was a
 * mechanical swap of
 *   google.script.run.withSuccessHandler(f).withFailureHandler(g)[name](...args)
 * for
 *   callApi_(name, ...args).then(f, g)
 * — see static-pages/src/index.html's callApi_ — without rewriting any of the underlying
 * admin/ballot RPC functions below (webAdmin.js, webBallot.js).
 *
 * Same trust boundary as the pages it replaces: this deployment is ANYONE_ANONYMOUS, so anyone
 * with the ballot/admin link could already reach these through the old server-rendered pages —
 * exposing them as a JSON endpoint doesn't grant access anyone didn't already have.
 *
 * Add to this whitelist only when the static client needs a new RPC — do not widen it to "any
 * global function", which would turn every helper in the project into an accidental public API.
 */
var _API_WHITELIST_ = {
  // Ballot voting page (webBallot.js)
  getBallotConfig: getBallotConfig,
  getBallotForName: getBallotForName,
  addBallotTopic: addBallotTopic,
  submitBallotRanking: submitBallotRanking,
  // Admin list/create (webAdmin.js)
  getAdminListData_: getAdminListData_,
  createBallotForAdmin_: createBallotForAdmin_,
  // Admin edit page (webAdmin.js)
  getAdminEditData: getAdminEditData,
  adminSaveTitle: adminSaveTitle,
  adminSaveDescription: adminSaveDescription,
  adminSaveInstructions: adminSaveInstructions,
  adminSaveFooter: adminSaveFooter,
  adminSaveContact: adminSaveContact,
  adminSaveAdminOnlyNotes: adminSaveAdminOnlyNotes,
  adminSaveAddSettings: adminSaveAddSettings,
  adminSaveCandidate: adminSaveCandidate,
  adminAddCandidate: adminAddCandidate,
  // Admin analysis (webAdmin.js)
  getAdminAnalysisData_: getAdminAnalysisData_
};

/**
 * doPost ?cmd=api handler. Body: {"name": "<whitelisted fn>", "args": [...]}, sent with
 * Content-Type: text/plain (not application/json) so the browser treats it as a CORS "simple
 * request" and skips a preflight OPTIONS — Apps Script web apps don't implement OPTIONS, so a
 * real application/json POST from a cross-origin static page would otherwise fail outright
 * (the same workaround F3Go30's static front end uses). GAS still reads e.postData.contents as
 * the raw JSON string regardless of the declared Content-Type. Response shape mirrors
 * google.script.run: {result: ...} on success, {error: "message"} on failure — see
 * static-pages/src/index.html's callApi_.
 *
 * @param {Object} e doPost event.
 * @return {TextOutput}
 */
function handleApiPost_(e) {
  var payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput_({ error: 'invalid_json' });
  }

  var fn = _API_WHITELIST_[payload.name];
  if (typeof fn !== 'function') {
    return jsonOutput_({ error: 'unknown_rpc: ' + payload.name });
  }

  try {
    var result = fn.apply(null, payload.args || []);
    return jsonOutput_({ result: result });
  } catch (err) {
    GasLogger.logError('handleApiPost_.error', err, { name: payload.name });
    return jsonOutput_({ error: String(err && err.message ? err.message : err) });
  }
}

/**
 * Where the static front end (this deployment's UI, formerly rendered server-side via
 * HtmlService — see WebApp.js doGet) is published for the currently deployed target. Each
 * target's static content is hosted on a *sibling* GitHub Pages repo, not this one — published
 * by tools/publish-static-pages.js as the last step of every tools/manage-deployments.js deploy:
 *   SIT/PROD  -> ../F3Static (f3go30/static-pages), which already hosts F3Go30's own static
 *                front end at dist/{sit,prod}/ — RankChoiceVoting's build is namespaced
 *                alongside it at ballot/{sit,prod}/ rather than getting its own repo, following
 *                the same multi-app-per-static-host convention nuuc-it/Static already uses
 *                (pub/AS, pub/AS-sit, ...).
 *   NUUC      -> ../Static (nuuc-it/Static), at pub/ballot/ (a single folder — NUUC is one
 *                environment, unlike SIT/PROD, so there's no sit/prod split to make here).
 * Returns '' for an unrecognized/unset APP_DEPLOY_TARGET so callers can render an explicit
 * "not configured" message instead of a broken link.
 *
 * @return {string}
 */
function _staticPagesBaseUrl_() {
  var target = (typeof APP_DEPLOY_TARGET !== 'undefined' && APP_DEPLOY_TARGET) || '';
  if (target === 'SIT') return 'https://f3go30.github.io/static-pages/ballot/sit/';
  if (target === 'PROD') return 'https://f3go30.github.io/static-pages/ballot/prod/';
  if (target === 'NUUC') return 'https://nuuc-it.github.io/Static/pub/ballot/';
  return '';
}

/**
 * The one-tap "taking you to..." page GAS now serves for every ?cmd=admin / ?cmd=ballot
 * arrival — the actual UI has moved to the static front end (_staticPagesBaseUrl_), which calls
 * this deployment's ?cmd=api JSON endpoint as its backend. doGet's own query string is
 * forwarded as-is, so a legacy bookmarked/shared link (e.g. ?cmd=ballot&id=X) still lands on
 * the right view once tapped through.
 *
 * NOT an auto-redirect: GAS wraps doGet HTML in a sandboxed iframe
 * (allow-top-navigation-by-user-activation), so a scripted top-level navigation on load has no
 * user activation and Chrome silently refuses it — this has to be a genuine, explicit tap
 * (target="_top", absolute href), the same pattern already used for cross-frame form actions
 * elsewhere in this app (see webAdmin.js's _renderCreateForm_ comment).
 *
 * @param {string} queryString e.queryString, forwarded verbatim.
 * @param {string} label shown above the tap-through link.
 * @return {HtmlOutput}
 */
function _renderStaticRedirect_(queryString, label) {
  var base = _staticPagesBaseUrl_();
  if (!base) {
    return HtmlService.createHtmlOutput(
      '<p>Static hosting is not configured for this deployment (APP_DEPLOY_TARGET="' +
      _escapeHtml_((typeof APP_DEPLOY_TARGET !== 'undefined' && APP_DEPLOY_TARGET) || '') + '").</p>'
    ).setTitle('RankChoiceVoting');
  }
  var staticUrl = base + (queryString ? '?' + queryString : '');
  return HtmlService.createHtmlOutput(
    '<body style="font-family:Arial,sans-serif;max-width:480px;margin:60px auto;padding:0 16px;text-align:center;">' +
    '<p>' + _escapeHtml_(label) + '</p>' +
    '<p><a target="_top" href="' + _escapeHtml_(staticUrl) + '" ' +
    'style="display:inline-block;padding:12px 20px;background:#1a73e8;color:#fff;' +
    'border-radius:6px;text-decoration:none;font-family:Arial,sans-serif;">Tap here to continue &rarr;</a></p>' +
    '</body>'
  ).setTitle('RankChoiceVoting')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    _staticPagesBaseUrl_: _staticPagesBaseUrl_
  };
}
