/**
 * libIdentity.js -- canonical verifier for brokered signed identity assertions.
 *
 * Contract: docs/interfaces/signed-identity-assertion.md (GAS-Core, estate-wide).
 * Wire format: JWT-compatible HS256, base64url(header).base64url(payload).base64url(hmac).
 * The Script Property named by the pinned kid IS the HMAC key (UTF-8, never base64-decoded).
 *
 * Scope: VERIFY side only. Issuing stays with the broker (NUUC-Dispatch Assertion_issue).
 * This is NOT the direct per-app GIS/tokeninfo verifier (PracticeMix Identity.js).
 *
 * Usage (flat-copied into the consumer's script dir; see GAS-Core README):
 *
 *   var r = LibIdentity_verify_(token, 'gactionsheet', {
 *     targets: { gactionsheet: 'ASSERTION_KEY_GACTIONSHEET_1' }, // aud -> kid registry
 *     failure: 'deny'                                           // REQUIRED: 'deny' | 'throw'
 *   });
 *   if (!r.valid) { ... }  // r.reason is a short code, never token or secret material
 *
 * Failure posture is a DECLARED option, not a convention each caller must invert:
 *   'deny'  -> every verification failure returns {valid:false, reason}; never throws.
 *   'throw' -> every verification failure throws an Error (name 'LibIdentityError',
 *              .reason set). Success is the only normal return.
 * There is no posture that admits an unverified token. Omitting or misspelling
 * `failure`, or omitting `targets`, is a configuration error and always throws
 * (the call is denied by exception; it can never fall through to "valid").
 *
 * Checks, in order (all fail closed): input is a string; three segments; header and
 * payload parse as JSON objects; alg === 'HS256' and kid present; expectedAud is
 * registered in targets (own-property only); header.kid === targets[expectedAud]
 * (kid pin, checked BEFORE any Script Property read); secret exists; constant-time
 * HMAC compare; iss; aud; exp is a number and not past; sub and email are
 * non-empty strings.
 *
 * Secrets and tokens are never logged or placed in reasons.
 *
 * RankChoiceVoting deviation from the vendored GAS-Core copy (same as NForms nforms-5dg.2; rcballot-dx3.2): every top-level
 * function name ends in '_' (LibIdentity_verify -> LibIdentity_verify_, _libIdentityX -> libIdentityX_).
 * This web app is ANYONE_ANONYMOUS and any HtmlService page it serves exposes, via google.script.run, every top-level function whose name
 * does not end in '_'; the vendored names let an anonymous visitor call the verifier with their own
 * targets and probe Script Properties. Re-apply this rename when re-vendoring.
 */

var LIBIDENTITY_DEFAULT_ISSUER = 'nuuc-dispatch';

/**
 * @param {string} jwt
 * @param {string} expectedAud the verifier's own registered identifier
 * @param {{targets: Object<string,string>, failure: ('deny'|'throw'), issuer: (string|undefined),
 *          nowSeconds: (number|undefined)}} options
 * @returns {{valid: true, sub: string, email: string, email_verified: boolean} |
 *           {valid: false, reason: string}}
 */
function LibIdentity_verify_(jwt, expectedAud, options) {
  var posture = options && options.failure;
  if (posture !== 'deny' && posture !== 'throw') {
    throw libIdentityError_('bad_config', "options.failure must be declared as 'deny' or 'throw'");
  }
  if (!options.targets || typeof options.targets !== 'object') {
    throw libIdentityError_('bad_config', 'options.targets (aud -> kid registry) is required');
  }

  var result = libIdentityCheck_(jwt, expectedAud, options);
  if (!result.valid && posture === 'throw') {
    throw libIdentityError_(result.reason, 'assertion rejected: ' + result.reason);
  }
  return result;
}

function libIdentityError_(reason, message) {
  var e = new Error(message);
  e.name = 'LibIdentityError';
  e.reason = reason;
  return e;
}

function libIdentityCheck_(jwt, expectedAud, options) {
  function invalid(reason) { return { valid: false, reason: reason }; }

  try {
    if (typeof jwt !== 'string') return invalid('malformed');
    var parts = jwt.split('.');
    if (parts.length !== 3) return invalid('malformed');

    var header, payload;
    try {
      header = JSON.parse(libIdentityB64urlToString_(parts[0]));
      payload = JSON.parse(libIdentityB64urlToString_(parts[1]));
    } catch (e) {
      return invalid('malformed');
    }
    if (!header || typeof header !== 'object' || !payload || typeof payload !== 'object') {
      return invalid('malformed');
    }

    if (header.alg !== 'HS256' || !header.kid) return invalid('bad_alg');

    // Pin the kid before any Script Property read: the token never selects its own key.
    var expectedKid = (typeof expectedAud === 'string' &&
        Object.prototype.hasOwnProperty.call(options.targets, expectedAud))
      ? options.targets[expectedAud]
      : null;
    if (!expectedKid || typeof expectedKid !== 'string') return invalid('unknown_aud');
    if (header.kid !== expectedKid) return invalid('bad_kid');

    var secret = PropertiesService.getScriptProperties().getProperty(expectedKid);
    if (!secret) return invalid('unknown_kid');

    var expectedSig = libIdentityB64urlFromBytes_(
      Utilities.computeHmacSha256Signature(parts[0] + '.' + parts[1], secret)
    );
    if (!libIdentityConstantTimeEquals_(expectedSig, parts[2])) return invalid('bad_signature');

    var issuer = options.issuer || LIBIDENTITY_DEFAULT_ISSUER;
    if (payload.iss !== issuer) return invalid('bad_iss');
    if (payload.aud !== expectedAud) return invalid('bad_aud');

    var now = options.nowSeconds || Math.floor(Date.now() / 1000);
    if (typeof payload.exp !== 'number' || now > payload.exp) return invalid('expired');

    if (typeof payload.sub !== 'string' || !payload.sub ||
        typeof payload.email !== 'string' || !payload.email) {
      return invalid('missing_identity');
    }

    return {
      valid: true,
      sub: payload.sub,
      email: payload.email,
      email_verified: payload.email_verified === true
    };
  } catch (e) {
    // Any unexpected runtime failure (e.g. HMAC service error) denies.
    return invalid('verify_error');
  }
}

function libIdentityB64urlFromBytes_(bytes) {
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

function libIdentityB64urlToString_(segment) {
  return Utilities.newBlob(Utilities.base64DecodeWebSafe(segment)).getDataAsString();
}

/** Length-then-XOR compare so signature checks don't short-circuit. */
function libIdentityConstantTimeEquals_(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string' || a.length !== b.length) return false;
  var diff = 0;
  for (var i = 0; i < a.length; i++) {
    diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return diff === 0;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { LibIdentity_verify_: LibIdentity_verify_ };
}
