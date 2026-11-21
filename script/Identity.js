/**
 * Identity.js — optional Google sign-in via NUUC-Dispatch, and ballot admin authorization
 * (rcballot-dx3.2, rcballot-dx3.3). Pattern: ../NForms/src/Security.js.
 *
 * FLOW: the static page signs in with Google Identity Services, sends the GIS credential to
 * NUUC-Dispatch (verify_identity), and gets back an HS256 signed assertion for this deployment's
 * audience. Every ?cmd=api call then carries that assertion in the request ENVELOPE (beside
 * name/args — never as an RPC argument); ApiBridge.js's handleApiPost_ calls rcvBeginRequest_ once,
 * which verifies it with the vendored GAS-Core LibIdentity (libIdentity.js, failure 'deny') and
 * holds the result for THIS request only. RPCs read it with rcvRequestIdentity_().
 *
 * ENABLEMENT: a deployment is Dispatch-enabled only when its stamped target maps to an audience
 * (SIT -> rcballot-test, NUUC -> rcballot; PROD -> none) AND that audience's key Script Property
 * exists. Not enabled: assertions are ignored, nobody is signed in, and every admin RPC stays open
 * exactly as before. Enabled but no/invalid assertion: the caller is anonymous.
 *
 * ADMIN RULE (rcvRequireBallotAdmin_, called at the top of every admin RPC in webAdmin.js — at
 * the function, not only in the ApiBridge whitelist, so a google.script.run call from any
 * HtmlService page this deployment serves cannot skip it): a ballot whose [Config] 'Admins' row
 * is blank is open to anyone, anonymous included. Otherwise the caller's verified email must be
 * in that row or in the spreadsheet-level 'Config' tab's 'Admins' key (super-admins), else the
 * RPC throws 'Not authorized …'. Fails closed: a read error while checking refuses.
 *
 * RESULTS (AC rev 2): getAdminAnalysisData_ needs a verified signed-in caller (rcvRequireSignedIn_)
 * but not an admin.
 *
 * PII: never log an email, name, picture URL or assertion. Logs carry reason codes only.
 *
 * Every top-level name ends in '_' so google.script.run cannot reach it.
 */

var RCV_ASSERTION_TARGETS_ = Object.freeze({
  'rcballot-test': 'ASSERTION_KEY_RCBALLOT_TEST_1',
  'rcballot': 'ASSERTION_KEY_RCBALLOT_1'
});
var RCV_AUD_BY_TARGET_ = Object.freeze({ SIT: 'rcballot-test', NUUC: 'rcballot' });

/** Spreadsheet-level tabs (not ballot sheets — listBallotIds_ only lists Ballot-*). */
var RCV_CONFIG_SHEET_ = 'Config';
var RCV_SIGNINS_SHEET_ = 'Signins';
var RCV_SIGNINS_HEADER_ = ['Email', 'Name', 'Picture', 'Sub', 'First Seen', 'Last Seen', 'Ballots'];
var RCV_ADMINS_KEY_ = 'Admins';

var RCV_NOT_AUTHORIZED_ = 'Not authorized: this ballot is managed by its listed admins. ' +
  'Sign in with a Google account that is one of this ballot\'s admins.';

// Per-request auth state, set by rcvBeginRequest_. GAS runs one request per execution, so a
// module-level var is request-scoped; rcvBeginRequest_ resets it first so nothing carries over.
var rcvRequest_ = null;

/** This deployment's assertion audience from its stamped target, or '' (PROD: not enabled). */
function rcvAssertionAud_() {
  var target = '';
  if (typeof BUILD_INFO !== 'undefined' && BUILD_INFO && BUILD_INFO.env) target = BUILD_INFO.env;
  else if (typeof APP_DEPLOY_TARGET !== 'undefined') target = APP_DEPLOY_TARGET;
  var aud = RCV_AUD_BY_TARGET_[String(target || '').toUpperCase()];
  return aud || '';
}

/** @return {boolean} true when this deployment has an audience AND its key Script Property. */
function rcvIdentityEnabled_() {
  var aud = rcvAssertionAud_();
  if (!aud) return false;
  try {
    return !!PropertiesService.getScriptProperties().getProperty(RCV_ASSERTION_TARGETS_[aud]);
  } catch (err) {
    return false;
  }
}

/**
 * Verifies one assertion. Disabled deployments never verify.
 *
 * @param {*} assertion
 * @return {{ok:true, email:string, sub:string}|{ok:false, reason:string}} reason is a short code.
 */
function rcvVerifyAssertion_(assertion) {
  if (!rcvIdentityEnabled_()) return { ok: false, reason: 'disabled' };
  if (!assertion) return { ok: false, reason: 'none' };
  var r = LibIdentity_verify_(assertion, rcvAssertionAud_(), { targets: RCV_ASSERTION_TARGETS_, failure: 'deny' });
  if (!r.valid) return { ok: false, reason: r.reason };
  if (r.email_verified !== true) return { ok: false, reason: 'email_unverified' };
  var email = String(r.email || '').trim().toLowerCase();
  if (!email) return { ok: false, reason: 'missing_identity' };
  return { ok: true, email: email, sub: String(r.sub) };
}

/**
 * GIS profile picture URL, accepted only from Google's own image host. UNVERIFIED: it comes from
 * the client-decoded GIS credential, not the assertion — display-only.
 */
function rcvSafePicture_(url) {
  var u = String(url || '').trim();
  return u.length <= 1000 && /^https:\/\/[a-z0-9-]+\.googleusercontent\.com\/[^\s"'<>]*$/i.test(u) ? u : '';
}

/** Text safe to write to a cell: trimmed, capped, never read as a formula. */
function rcvSheetText_(value, max) {
  var t = String(value == null ? '' : value).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max || 200);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

/**
 * Starts a ?cmd=api request: verifies the envelope's assertion ONCE and records the result.
 *
 * @param {{assertion:(string|undefined), profile:(Object|undefined)}} envelope parsed POST body
 * @return {?{status:string, reason:(string|undefined), email:(string|undefined)}} what the client
 *   is told about its assertion — null when it sent none or the deployment is not enabled.
 */
function rcvBeginRequest_(envelope) {
  rcvRequest_ = { enabled: rcvIdentityEnabled_(), identity: null };
  var assertion = envelope && envelope.assertion;
  if (!rcvRequest_.enabled || !assertion) return null;
  var v = rcvVerifyAssertion_(assertion);
  if (!v.ok) {
    // Treated as anonymous; logged by reason code only (rcballot-dx3.2 AC3).
    GasLogger.log('identity.assertion.rejected', { reason: v.reason });
    return { status: 'rejected', reason: v.reason };
  }
  var profile = (envelope && envelope.profile) || {};
  if (typeof GasLogger !== 'undefined' && GasLogger.setEmail) GasLogger.setEmail(v.email); // rcballot-8un
  rcvRequest_.identity = {
    email: v.email,
    sub: v.sub,
    // Display-only, UNVERIFIED (GIS payload decoded client-side): never used for authorization.
    picture: rcvSafePicture_(profile.picture),
    displayName: rcvSheetText_(profile.name, 100)
  };
  return { status: 'ok', email: v.email };
}

/** Clears the per-request state (end of handleApiPost_). */
function rcvEndRequest_() { rcvRequest_ = null; }

/** @return {?{email:string, sub:string, picture:string, displayName:string}} this request's verified caller. */
function rcvRequestIdentity_() {
  return rcvRequest_ && rcvRequest_.identity ? rcvRequest_.identity : null;
}

/** Enablement for this request (computed once by rcvBeginRequest_; recomputed outside a request). */
function rcvRequestEnabled_() {
  return rcvRequest_ ? rcvRequest_.enabled : rcvIdentityEnabled_();
}

/** Comma/semicolon/whitespace-separated emails -> unique, lowercased, in order. */
function rcvParseEmailList_(text) {
  var seen = {};
  var out = [];
  String(text == null ? '' : text).split(/[\s,;]+/).forEach(function (e) {
    var k = e.trim().toLowerCase();
    if (k && !seen[k]) { seen[k] = true; out.push(k); }
  });
  return out;
}

function rcvLooksLikeEmail_(e) { return /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(e); }

/** Super-admins: the 'Admins' key (col A) of the spreadsheet-level Config tab (value col B). */
function rcvSuperAdmins_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RCV_CONFIG_SHEET_);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  var values = sheet.getRange(1, 1, lastRow, 2).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim().toLowerCase() === RCV_ADMINS_KEY_.toLowerCase()) {
      return rcvParseEmailList_(values[i][1]);
    }
  }
  return [];
}

/** A ballot's own admins ([Config] 'Admins' row), [] when unclaimed or the ballot is absent. */
function rcvBallotAdmins_(id) {
  var sheet = findBallotSheet_(SpreadsheetApp.getActiveSpreadsheet(), id);
  if (!sheet) return [];
  return rcvParseEmailList_(_readBallotView_(sheet).config[RCV_ADMINS_KEY_]);
}

/**
 * The admin decision for ballot id and this request's caller. Pure over its inputs' reads.
 *
 * @param {string} id
 * @return {{enabled:boolean, claimed:boolean, allowed:boolean, isSuperAdmin:boolean,
 *   isBallotAdmin:boolean, admins:Array<string>, email:string}}
 */
function rcvBallotAccess_(id) {
  var who = rcvRequestIdentity_();
  var email = who ? who.email : '';
  if (!rcvRequestEnabled_()) {
    return { enabled: false, claimed: false, allowed: true, isSuperAdmin: false, isBallotAdmin: false, admins: [], email: '' };
  }
  var admins = rcvBallotAdmins_(id);
  var supers = email ? rcvSuperAdmins_() : [];
  var isBallotAdmin = !!email && admins.indexOf(email) !== -1;
  var isSuperAdmin = !!email && supers.indexOf(email) !== -1;
  return {
    enabled: true,
    claimed: admins.length > 0,
    allowed: admins.length === 0 || isBallotAdmin || isSuperAdmin,
    isSuperAdmin: isSuperAdmin,
    isBallotAdmin: isBallotAdmin,
    admins: admins,
    email: email
  };
}

/**
 * Throws unless this request's caller may administer ballot id (see the file header rule).
 * Fails closed: any error while deciding refuses.
 *
 * @param {string} id
 * @return {Object} rcvBallotAccess_'s result when allowed.
 */
function rcvRequireBallotAdmin_(id) {
  var access;
  try {
    access = rcvBallotAccess_(id);
  } catch (err) {
    GasLogger.log('identity.admin.refused', { reason: 'check_error' });
    throw new Error(RCV_NOT_AUTHORIZED_);
  }
  if (!access.allowed) {
    GasLogger.log('identity.admin.refused', { reason: rcvRequestIdentity_() ? 'not_admin' : 'anonymous' });
    throw new Error(RCV_NOT_AUTHORIZED_);
  }
  return access;
}

/**
 * Throws message unless this request has a verified signed-in caller. Open (no-op) on a
 * deployment without sign-in, like the admin rule.
 *
 * @param {string} message refusal text (starts 'Not authorized')
 */
function rcvRequireSignedIn_(message) {
  if (!rcvRequestEnabled_() || rcvRequestIdentity_()) return;
  GasLogger.log('identity.signin.required', { reason: 'anonymous' });
  throw new Error(message);
}

/**
 * Upserts the caller into the spreadsheet-level Signins tab (created on first use), keyed by
 * verified Sub: Email (verified), Name and Picture (display-only, unverified — kept from the last
 * sign-in that sent one), First Seen, Last Seen, Ballots (ids, comma-separated, union). No-op when
 * nobody is signed in.
 *
 * @param {string=} ballotId the ballot the sign-in happened on, if any
 * @return {boolean} true when a row was written
 */
function rcvRecordSignin_(ballotId) {
  var who = rcvRequestIdentity_();
  if (!who) return false;
  return _withBallotLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(RCV_SIGNINS_SHEET_);
    if (!sheet) {
      sheet = ss.insertSheet(RCV_SIGNINS_SHEET_);
      sheet.getRange(1, 1, 1, RCV_SIGNINS_HEADER_.length).setValues([RCV_SIGNINS_HEADER_]);
    }
    var width = RCV_SIGNINS_HEADER_.length;
    var lastRow = Math.max(sheet.getLastRow(), 1);
    var values = sheet.getRange(1, 1, lastRow, width).getValues();
    var row = -1;
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][3] || '').trim() === who.sub) { row = i + 1; break; }
    }
    var now = new Date();
    var prev = row === -1 ? ['', '', '', '', now, now, ''] : values[row - 1];
    var ballots = String(prev[6] || '').split(',').map(function (b) { return b.trim(); }).filter(Boolean);
    var bid = rcvSheetText_(ballotId, 100);
    if (bid && ballots.indexOf(bid) === -1) ballots.push(bid);
    var out = [
      who.email,
      who.displayName || prev[1] || '',
      who.picture || prev[2] || '',
      who.sub,
      prev[4] || now,
      now,
      ballots.join(', ')
    ];
    if (row === -1) row = lastRow + 1;
    sheet.getRange(row, 1, 1, width).setValues([out]);
    return true;
  });
}

/** Writes a ballot's Admins row (through the ballot write path), keeping the cache in sync. */
function rcvWriteBallotAdmins_(id, admins) {
  return _adminWriteConfigUnchecked_(id, { Admins: admins.join(', ') });
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    rcvAssertionAud_: rcvAssertionAud_,
    rcvIdentityEnabled_: rcvIdentityEnabled_,
    rcvVerifyAssertion_: rcvVerifyAssertion_,
    rcvBeginRequest_: rcvBeginRequest_,
    rcvRequestIdentity_: rcvRequestIdentity_,
    rcvParseEmailList_: rcvParseEmailList_,
    rcvSafePicture_: rcvSafePicture_
  };
}
