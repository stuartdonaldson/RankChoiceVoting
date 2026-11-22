/**
 * webAdmin.js
 *
 * Admin RPCs for the static front end (static-pages/src/index.html), reached via ApiBridge.js's
 * doPost ?cmd=api whitelist: getAdminListData_, createBallotForAdmin_, getAdminAnalysisData_,
 * getAdminEditData, adminSave*, adminAddCandidate. Also reachable from the spreadsheet's
 * "Voting and Ballot Tools" menu (onOpen.js).
 *
 * rcballot-719: the legacy HtmlService admin pages (_handleAdmin, its HTML renderers, and
 * webAdminEditPage.html) were retired; doGet redirects cmd=admin to the static page.
 */

/**
 * JSON data source for the static admin list view (static-pages/src/index.html), reached via
 * ApiBridge.js's ?cmd=api dispatcher. Same fields/logic _renderAdminList_ used to render
 * server-side as HTML — ballotUrl/editUrl/analyzeUrl are deliberately NOT included here; the
 * static client builds those itself (same-origin relative links off its own location), so this
 * function only returns data, not navigation.
 *
 * @return {Array<{id:string, title:string, candidateCount:number, responseCount:number,
 *   uniqueRespondentCount:number, adminNotes:string, staleRespondents:Array<string>}>}
 */
function getAdminListData_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ids = listBallotIds_(ss);
  // rcballot-dx3.3: the list shows every ballot, but a claimed ballot's admin-only fields
  // (notes, respondent names) go only to its admins / super-admins. Edit and Analysis enforce
  // the rule server-side on their own RPCs.
  var enabled = rcvRequestEnabled_();
  var who = rcvRequestIdentity_();
  var supers = enabled && who ? rcvSuperAdmins_() : [];
  var isSuper = !!who && supers.indexOf(who.email) !== -1;
  return ids.map(function (id) {
    // One snapshot read per ballot card serves every field (rcballot-8i5.3); no write.
    var view = _readBallotView_(findBallotSheet_(ss, id));
    var latest = _latestResponseByRespondent_(view.rows);
    var admins = enabled ? rcvParseEmailList_(view.config.Admins) : [];
    var allowed = !admins.length || isSuper || (!!who && admins.indexOf(who.email) !== -1);
    return {
      id: id,
      title: view.config.Title || id,
      candidateCount: view.names.length,
      responseCount: view.rows.length,
      uniqueRespondentCount: latest.length,
      adminNotes: allowed ? String(view.config['Admin-Only-Notes'] || '').trim() : '',
      staleRespondents: allowed ? _staleRespondents_(latest) : [],
      claimed: admins.length > 0,
      canAdmin: allowed
    };
  });
}

/**
 * RPC: creates a new ballot from the static list view's "Create New Ballot" form (replaces the
 * old GET-form-submit-to-doGet flow — see _renderCreateForm_, now unreachable from doGet). On
 * success, the client switches straight to the edit view with the "just created" banner.
 *
 * @param {string} id
 * @return {{ok:boolean, id:string}} throws (surfaced as {error} by handleApiPost_) on an
 *   invalid/duplicate id, same as createNewBallot_.
 */
function createBallotForAdmin_(id) {
  rcvRequireBallotAdmin_(id); // rcballot-dx3.3: server-side admin rule (Identity.js)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  createNewBallot_(ss, id);
  return { ok: true, id: id };
}

/**
 * JSON data source for the static admin analysis view. Runs the same analysis
 * _renderAdminAnalysis_ used to render as an HTML string (and still writes the Results section
 * back to the sheet, same as before) but returns the raw result for the client to render.
 *
 * @param {string} id
 * @return {Object} {rcv, rcvFinishOrder, condorcet, finishOrders, sheetName, comments} or {error}.
 */
function getAdminAnalysisData_(id) {
  // rcballot-dx3.3 AC rev 2: results need a verified signed-in caller, but not an admin.
  rcvRequireSignedIn_('Not authorized: sign in with Google to view results.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findBallotSheet_(ss, id);
  if (!sheet) return { error: 'no_ballot' };

  var results = runBallotAnalysis_(sheet);
  if (results.error) return { error: results.error };

  return {
    rcv: results.rcv,
    rcvFinishOrder: results.rcvFinishOrder,
    condorcet: results.condorcet,
    finishOrders: results.finishOrders,
    sheetName: getBallotSheetName_(id),
    comments: _feedbackComments_(_readBallotView_(sheet).rows)
  };
}

/**
 * Non-blank feedback comments from a view's response rows, newest first (rcballot-5i2).
 * Rows with weight 0 are kept but flagged counted:false. Never log the result (PII).
 *
 * @param {Array<{row:number, date:*, name:string, weight:number, comment:string}>} rows
 * @return {Array<{name:string, date:string, comment:string, counted:boolean}>} date is an
 *   ISO string when the cell is a date, else the cell's text ('' when absent).
 */
function _feedbackComments_(rows) {
  var items = [];
  rows.forEach(function (r) {
    var text = String(r.comment || '');
    if (!text.trim()) return;
    var ms = NaN;
    var date = '';
    if (Object.prototype.toString.call(r.date) === '[object Date]') {
      ms = r.date.getTime();
      date = isNaN(ms) ? '' : r.date.toISOString();
    } else if (r.date !== '' && r.date != null) {
      date = String(r.date);
      ms = Date.parse(date);
    }
    items.push({ ms: isNaN(ms) ? -Infinity : ms, row: r.row,
      out: { name: r.name, date: date, comment: text, counted: r.weight > 0 } });
  });
  items.sort(function (a, b) { return (b.ms - a.ms) || (b.row - a.row) || 0; });
  return items.map(function (i) { return i.out; });
}

/**
 * RPC: called once by webAdminEditPage.html on load (after reading the ballot id from
 * google.script.url.getLocation) to fetch everything the edit-mode ballot preview needs
 * to render. Field names match what the client displays/edits, not the sheet's config
 * key spelling — see BallotModel.js for the underlying "Title"/"Accept-New"/etc. keys.
 *
 * @param {string} id
 * @return {Object} or {error:'no_ballot'} if the id doesn't match a ballot sheet.
 */
function getAdminEditData(id) {
  var access = rcvRequireBallotAdmin_(id); // rcballot-dx3.3: server-side admin rule (Identity.js)
  // Admin edit-page load is one of the two places sheet repairs run (the other is the
  // write paths — rcballot-8i5.7/.3): a healthy sheet costs one read and no lock, and that
  // same snapshot serves the page (a repair also rebuilds the cache). Orphan rank columns
  // — Responses columns matching no candidate, e.g. after a hand-deleted Candidates row —
  // are returned for display; they are never counted by analysis.
  //
  // The cache is rebuilt from that snapshot on every load, not just after a repair
  // (rcballot-91f): a sheet copied in from another spreadsheet and renamed over a ballot
  // fires neither onEdit nor an RPC refresh, and nameless ballot loads never read the sheet,
  // so opening the ballot here is the admin's way to resync it. One Properties write; no
  // extra sheet read.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findBallotSheet_(ss, id);
  if (!sheet) return { error: 'no_ballot' };
  var rep = _repairBallotSheet_(sheet);
  var view = _ballotViewFromSnapshot_(rep.snap);
  refreshBallotCache_(id, sheet, view);
  // Repair could not fix it: tell the operator rather than let the first vote fail.
  var layoutError = '';
  if (_responsesMissingName_(rep.snap)) {
    layoutError = 'The Responses header row on this ballot sheet has no "Name" column, so votes cannot be saved. ' +
      'Open the sheet, insert a new row directly under [Responses] (do not overwrite any vote) and put Date | Name | Weight | Comment in it, followed by the candidate names that head the rank columns.';
    _logBallotEvent_('ballot.layout.unrepaired', { sheet: sheet.getName(), reason: 'responsesMissingName' });
  }

  var config = view.config;
  // Per-candidate vote count (responses with a rank in the candidate's column) so the delete
  // confirm can say how many votes go unlinked (rcballot-cck.4).
  var responses = _snapResponses_(rep.snap);
  var candidateRows = view.candidates.map(function (c) {
    return { name: c.name, details: c.details, linkText: c.linkText, linkUrl: c.linkUrl, votes: _countNonBlank_(responses, 'ranks', c.name) };
  });

  var baseUrl = _getWebAppUrl_();
  return {
    id: id,
    title: config.Title || '',
    description: config.Description || '',
    instructions: config.Instructions || '',
    footer: config.Footer || '',
    contact: config.Contact || '',
    acceptNew: _ballotAcceptsNew_(config),
    addInstructions: config['Add-Instructions'] || '',
    feedbackPrompt: config['Feedback-Prompt'] || '',
    defaultFeedbackPrompt: _DEFAULT_FEEDBACK_PROMPT_,
    thankYouMessage: config['Thank-You-Message'] || '',
    defaultThankYouMessage: _DEFAULT_THANK_YOU_MESSAGE_,
    showDetailsPreview: _ballotFlag_(config, 'Show-Details-Preview'),
    descriptionLinkText: config['Description-Link-Text'] || '',
    descriptionLinkUrl: config['Description-Link-URL'] || '',
    adminOnlyNotes: config['Admin-Only-Notes'] || '',
    // rcballot-dx3.3: only reaches a caller the admin rule allowed. enabled=false (PROD, or no
    // assertion key) means sign-in is off here and the page shows no admin controls.
    admins: access.admins,
    access: { enabled: access.enabled, claimed: access.claimed, signedInEmail: access.email,
      isSuperAdmin: access.isSuperAdmin, isBallotAdmin: access.isBallotAdmin },
    candidates: candidateRows,
    layoutError: layoutError,
    orphanColumns: view.orphans.filter(function (o) { return _orphanHasRanks_(rep.snap, o.label); })
      .map(function (o) { return { label: o.label, col: o.col, votes: _countNonBlank_(responses, 'orphanRanks', o.label) }; }),
    ballotUrl: baseUrl + '?cmd=ballot&id=' + encodeURIComponent(id),
    adminListUrl: baseUrl + '?cmd=admin',
    appVersion: (typeof APP_VERSION !== 'undefined' && APP_VERSION) || '',
    appDeployTarget: (typeof APP_DEPLOY_TARGET !== 'undefined' && APP_DEPLOY_TARGET) || ''
  };
}

/**
 * @param {string} id
 * @param {Object} updates passed straight through to writeBallotConfig_
 * @return {{ok:boolean}}
 */
function _adminWriteConfig_(id, updates) {
  rcvRequireBallotAdmin_(id); // rcballot-dx3.3: every adminSave* goes through here
  return _adminWriteConfigUnchecked_(id, updates);
}

/** _adminWriteConfig_ minus the admin check — callers must have checked already. */
function _adminWriteConfigUnchecked_(id, updates) {
  return _withBallotLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findBallotSheet_(ss, id);
    if (!sheet) throw new Error('No ballot found for id "' + id + '".');
    var snap = writeBallotConfig_(sheet, updates); // repairs first (_withBallotWrite_)
    refreshBallotCache_(id, sheet, _ballotViewFromSnapshot_(snap)); // keep the cache (BallotCache.js) in sync with this write
    return { ok: true };
  });
}

/**
 * RPC (rcballot-dx3.3, AC rev 2): replaces ballot id's Admins list. Same admin rule as every
 * other save. Refuses on a deployment without sign-in, and refuses a non-empty list from an
 * anonymous caller (unclaimed ballots included). A non-super-admin's list ALWAYS keeps their own
 * verified email (added if missing), so nobody can lock themselves out; super-admins may save
 * any list, including one without themselves. A blank list (from a super-admin, or an anonymous
 * caller on an unclaimed ballot) makes the ballot unclaimed (open to anyone) again; a signed-in
 * non-super-admin's blank list becomes just themselves. Extra arguments (the retired
 * confirmLockout flag) are ignored.
 *
 * @param {string} id
 * @param {string} value comma-separated emails
 * @return {{ok:boolean, admins:Array<string>}}
 */
function adminSaveAdmins(id, value) {
  var access = rcvRequireBallotAdmin_(id);
  if (!access.enabled) throw new Error('Sign-in is not enabled on this deployment, so ballot admins cannot be set.');
  var admins = rcvParseEmailList_(value);
  var bad = admins.filter(function (e) { return !rcvLooksLikeEmail_(e); });
  if (bad.length) throw new Error('Every admin entry must be an email address (' + bad.length + ' is not).');
  // Claiming a ballot needs a verified identity, even on an unclaimed one: otherwise an anonymous
  // visitor could hand any ballot to a stranger's address and lock everyone else out.
  if (admins.length && !access.email) {
    GasLogger.log('identity.admin.refused', { reason: 'anonymous_set_admins' });
    throw new Error('Not authorized: sign in with Google to set ballot admins.');
  }
  if (access.email && !access.isSuperAdmin && admins.indexOf(access.email) === -1) admins.push(access.email);
  rcvWriteBallotAdmins_(id, admins);
  return { ok: true, admins: admins };
}

/**
 * RPC (rcballot-dx3.3) "Make me an admin": adds the signed-in caller's verified email to ballot
 * id's Admins. Allowed wherever the caller may already administer the ballot — on an unclaimed
 * ballot, that is anyone signed in.
 *
 * @param {string} id
 * @return {{ok:boolean, admins:Array<string>}}
 */
function adminClaimBallot(id) {
  var access = rcvRequireBallotAdmin_(id);
  if (!access.enabled) throw new Error('Sign-in is not enabled on this deployment.');
  if (!access.email) throw new Error('Sign in with Google first.');
  var admins = access.admins.slice();
  if (admins.indexOf(access.email) === -1) admins.push(access.email);
  rcvWriteBallotAdmins_(id, admins);
  return { ok: true, admins: admins };
}

/** RPC: saves the Title field's pencil-icon editor. @return {{ok:boolean}} */
function adminSaveTitle(id, value) { return _adminWriteConfig_(id, { Title: String(value || '') }); }

/** RPC: saves the landing-page Description field's pencil-icon editor. @return {{ok:boolean}} */
function adminSaveDescription(id, value) { return _adminWriteConfig_(id, { Description: String(value || '') }); }

/** RPC: saves the ballot-page Instructions field's pencil-icon editor. @return {{ok:boolean}} */
function adminSaveInstructions(id, value) { return _adminWriteConfig_(id, { Instructions: String(value || '') }); }

/** RPC: saves the Footer field's pencil-icon editor. @return {{ok:boolean}} */
function adminSaveFooter(id, value) { return _adminWriteConfig_(id, { Footer: String(value || '') }); }

/** RPC: saves the Contact field's pencil-icon editor. @return {{ok:boolean}} */
function adminSaveContact(id, value) { return _adminWriteConfig_(id, { Contact: String(value || '') }); }

/**
 * RPC: saves the optional Description link (Link Text + Link URL) together. A non-blank
 * URL must be http(s); a blank URL clears the link (and the text is kept as typed).
 *
 * @param {string} id
 * @param {string} text
 * @param {string} url
 * @return {{ok:boolean}}
 */
function adminSaveDescriptionLink(id, text, url) {
  var u = String(url || '').trim();
  if (u && !_isHttpUrl_(u)) throw new Error('Link URL must start with http:// or https:// and contain no spaces.');
  return _adminWriteConfig_(id, { 'Description-Link-Text': String(text || '').trim(), 'Description-Link-URL': u });
}

/** RPC: saves the Feedback-Prompt (blank = built-in default). @return {{ok:boolean}} */
function adminSaveFeedbackPrompt(id, value) { return _adminWriteConfig_(id, { 'Feedback-Prompt': String(value || '').trim() }); }

/** RPC: saves the Thank-You-Message shown after Submit (blank = built-in default; rcballot-o3i). @return {{ok:boolean}} */
function adminSaveThankYouMessage(id, value) { return _adminWriteConfig_(id, { 'Thank-You-Message': String(value || '').trim() }); }

/** RPC: saves the Show-Details-Preview yes/no toggle. @return {{ok:boolean}} */
function adminSaveShowDetailsPreview(id, on) { return _adminWriteConfig_(id, { 'Show-Details-Preview': on ? 'TRUE' : 'FALSE' }); }

/** RPC: saves the Admin-Only Notes field's pencil-icon editor. @return {{ok:boolean}} */
function adminSaveAdminOnlyNotes(id, value) { return _adminWriteConfig_(id, { 'Admin-Only-Notes': String(value || '') }); }

/**
 * RPC: saves the combined "Adding New Candidates" editor (Accept-New checkbox +
 * Add-Instructions text) in one call, since they're edited together as one field group.
 *
 * @param {string} id
 * @param {boolean} acceptNew
 * @param {string} addInstructions
 * @return {{ok:boolean}}
 */
function adminSaveAddSettings(id, acceptNew, addInstructions) {
  return _adminWriteConfig_(id, {
    'Accept-New': acceptNew ? 'TRUE' : 'FALSE',
    'Add-Instructions': String(addInstructions || '')
  });
}

/**
 * RPC: saves one existing candidate's name/details/link from its pencil-icon editor.
 * Updates just that candidate's Candidates row (found by index into the Candidates-table
 * order getAdminEditData returned) and, on a rename, relabels its Responses rank column
 * (found by the OLD name) so recorded ranks stay with the candidate. Names are unique
 * case-insensitively (a case-only rename of the same candidate is allowed).
 *
 * @param {string} id
 * @param {number} index
 * @param {string} name
 * @param {string} details
 * @param {string=} linkText
 * @param {string=} linkUrl
 * @return {{ok:boolean, name:string, details:string, linkText:string, linkUrl:string}}
 */
function adminSaveCandidate(id, index, name, details, linkText, linkUrl) {
  rcvRequireBallotAdmin_(id); // rcballot-dx3.3: server-side admin rule (Identity.js)
  name = String(name || '').trim();
  details = String(details || '');
  linkText = String(linkText || '');
  linkUrl = String(linkUrl || '');
  return _withBallotLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findBallotSheet_(ss, id);
    if (!sheet) throw new Error('No ballot found for id "' + id + '".');
    var after = _withBallotWrite_(sheet, function (snap) {
      var c = _snapCandidates_(snap)[index];
      if (!c) throw new Error('Invalid candidate index.');
      if (!name) throw new Error('Candidate name cannot be empty.');
      _assertNotMarkerName_(name);
      _assertUniqueCandidateName_(snap, name, c.row);
      var ops = [{ section: 'candidates', update: [{ row: c.row, cells: { Name: name, Details: details, 'Link Text': linkText, 'Link URL': linkUrl } }] }];
      if (c.rankLabel !== null && c.rankLabel !== name) {
        ops.push({ section: 'responses', renameColumn: { from: c.rankLabel, to: name } });
      }
      return _applyBallotOps_(sheet, snap, ops);
    });
    refreshBallotCache_(id, sheet, _ballotViewFromSnapshot_(after)); // keep the cache (BallotCache.js) in sync with this write
    return { ok: true, name: name, details: details, linkText: linkText, linkUrl: linkUrl };
  });
}

/**
 * RPC: deletes one candidate (rcballot-8i5.6). Removes its Candidates row only; the
 * Responses rank column is kept and becomes an orphan (listed by getAdminEditData, never
 * counted by analysis), so the votes can be re-linked later (adminLinkOrphanColumn).
 *
 * @param {string} id
 * @param {string} name candidate to delete (trim/case-insensitive).
 * @return {{ok:boolean, name:string}}
 */
function adminDeleteCandidate(id, name) {
  rcvRequireBallotAdmin_(id); // rcballot-dx3.3: server-side admin rule (Identity.js)
  name = String(name || '').trim();
  return _withBallotLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findBallotSheet_(ss, id);
    if (!sheet) throw new Error('No ballot found for id "' + id + '".');
    var after = _withBallotWrite_(sheet, function (snap) {
      var key = name.toLowerCase();
      var found = null;
      _snapCandidates_(snap).forEach(function (c) { if (!found && c.name.toLowerCase() === key) found = c; });
      if (!found) throw new Error('Candidate not found.');
      return _applyBallotOps_(sheet, snap, [{ section: 'candidates', deleteRows: [found.row] }]);
    });
    refreshBallotCache_(id, sheet, _ballotViewFromSnapshot_(after));
    return { ok: true, name: name };
  });
}

/**
 * RPC: reorders the candidates (rcballot-8i5.6). Rewrites the Candidates rows' content in
 * the given order; the Responses columns are not touched (ranks are keyed by name).
 *
 * @param {string} id
 * @param {Array<string>} order every current candidate name exactly once, in the new order.
 * @return {{ok:boolean, order:Array<string>}} the candidate names in their new order.
 */
function adminReorderCandidates(id, order) {
  rcvRequireBallotAdmin_(id); // rcballot-dx3.3: server-side admin rule (Identity.js)
  return _withBallotLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findBallotSheet_(ss, id);
    if (!sheet) throw new Error('No ballot found for id "' + id + '".');
    var result = [];
    var after = _withBallotWrite_(sheet, function (snap) {
      var cands = _snapCandidates_(snap);
      if (!Array.isArray(order) || order.length !== cands.length) {
        throw new Error('Reorder must list every candidate exactly once.');
      }
      var byKey = {};
      cands.forEach(function (c) { byKey[c.name.toLowerCase()] = c; });
      var used = {};
      var updates = [];
      order.forEach(function (n, i) {
        var key = String(n || '').trim().toLowerCase();
        var c = byKey[key];
        if (!c || used[key]) throw new Error('Reorder must list every candidate exactly once.');
        used[key] = true;
        result.push(c.name);
        if (c.row !== cands[i].row) {
          updates.push({ row: cands[i].row, cells: { Name: c.name, Details: c.details, 'Link Text': c.linkText, 'Link URL': c.linkUrl } });
        }
      });
      if (!updates.length) return snap;
      return _applyBallotOps_(sheet, snap, [{ section: 'candidates', update: updates }]);
    });
    refreshBallotCache_(id, sheet, _ballotViewFromSnapshot_(after));
    return { ok: true, order: result };
  });
}

/** @return {number} responses whose `field` map (ranks / orphanRanks) has a non-blank value under key. */
function _countNonBlank_(responses, field, key) {
  return responses.filter(function (r) {
    var v = r[field][key];
    return v !== '' && v !== null && v !== undefined;
  }).length;
}

/** @return {boolean} true when any response has a non-blank rank under the candidate's column. */
function _columnHasRanks_(snap, candidateName) {
  return _snapResponses_(snap).some(function (r) {
    var v = r.ranks[candidateName];
    return v !== '' && v !== null && v !== undefined;
  });
}

/** @return {boolean} true when any response has a non-blank value under the orphan column label. */
function _orphanHasRanks_(snap, label) {
  return _snapResponses_(snap).some(function (r) {
    var v = r.orphanRanks[label];
    return v !== '' && v !== null && v !== undefined;
  });
}

/**
 * RPC: links an orphan Responses rank column to an existing candidate (rcballot-8i5.6) by
 * renaming the column header to the candidate's name, so the column's votes count for that
 * candidate. Refused when the candidate's own rank column already holds votes (columns are
 * never merged) or the label is not currently an orphan. A blank own column (the repair
 * backfills one for every candidate) is swapped out: it takes the orphan's old label, so one
 * empty orphan column is left behind (not listed: getAdminEditData hides vote-less orphans).
 *
 * @param {string} id
 * @param {string} label orphan column header, as listed in getAdminEditData's orphanColumns.
 * @param {string} candidateName
 * @return {{ok:boolean, label:string, candidate:string}}
 */
function adminLinkOrphanColumn(id, label, candidateName) {
  rcvRequireBallotAdmin_(id); // rcballot-dx3.3: server-side admin rule (Identity.js)
  label = String(label || '').trim();
  candidateName = String(candidateName || '').trim();
  return _withBallotLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findBallotSheet_(ss, id);
    if (!sheet) throw new Error('No ballot found for id "' + id + '".');
    var linked = '';
    var after = _withBallotWrite_(sheet, function (snap) {
      var orphan = null;
      _snapOrphanColumns_(snap).forEach(function (o) { if (!orphan && o.label.toLowerCase() === label.toLowerCase()) orphan = o; });
      if (!orphan) throw new Error('Orphan column not found.');
      var cand = null;
      _snapCandidates_(snap).forEach(function (c) { if (!cand && c.name.toLowerCase() === candidateName.toLowerCase()) cand = c; });
      if (!cand) throw new Error('Candidate not found.');
      var ops = [];
      if (cand.rankLabel !== null) {
        // The repair that precedes every write backfills an empty column for a candidate with
        // none, so "no column of its own" means "no votes in it". Columns are never merged.
        if (_columnHasRanks_(snap, cand.name)) {
          throw new Error('"' + cand.name + '" already has votes in its own rank column; columns are not merged.');
        }
        // Park the empty column under a temp label so the orphan's column can take the name,
        // then give the empty column the orphan's old label (temp labels never reach the sheet).
        ops.push({ section: 'responses', renameColumn: { from: cand.rankLabel, to: '\u0001link' } });
        ops.push({ section: 'responses', renameColumn: { from: orphan.label, to: cand.name } });
        ops.push({ section: 'responses', renameColumn: { from: '\u0001link', to: orphan.label } });
      } else {
        ops.push({ section: 'responses', renameColumn: { from: orphan.label, to: cand.name } });
      }
      linked = cand.name;
      return _applyBallotOps_(sheet, snap, ops);
    });
    refreshBallotCache_(id, sheet, _ballotViewFromSnapshot_(after));
    return { ok: true, label: label, candidate: linked };
  });
}

/**
 * RPC: discards an unlinked (orphan) Responses column (rcballot-cck.4): deletes the column
 * from the sheet, with the votes in it. Each voter's other rankings are untouched. Refused
 * when the label is not currently an orphan (a candidate's own column is never deletable here).
 *
 * @param {string} id
 * @param {string} label orphan column header, as listed in getAdminEditData's orphanColumns.
 * @return {{ok:boolean, label:string, votes:number}} votes = how many responses had a rank in it.
 */
function adminDiscardOrphanColumn(id, label) {
  rcvRequireBallotAdmin_(id); // rcballot-dx3.3: server-side admin rule (Identity.js)
  label = String(label || '').trim();
  return _withBallotLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findBallotSheet_(ss, id);
    if (!sheet) throw new Error('No ballot found for id "' + id + '".');
    var votes = 0;
    var discarded = '';
    var after = _withBallotWrite_(sheet, function (snap) {
      var orphan = null;
      _snapOrphanColumns_(snap).forEach(function (o) { if (!orphan && o.label.toLowerCase() === label.toLowerCase()) orphan = o; });
      if (!orphan) throw new Error('Orphan column not found.');
      votes = _countNonBlank_(_snapResponses_(snap), 'orphanRanks', orphan.label);
      discarded = orphan.label;
      return _applyBallotOps_(sheet, snap, [{ section: 'responses', deleteColumn: orphan.label }]);
    });
    refreshBallotCache_(id, sheet, _ballotViewFromSnapshot_(after));
    _logBallotEvent_('ballot.orphan.discarded', { sheet: sheet.getName(), label: discarded, votes: votes });
    return { ok: true, label: discarded, votes: votes };
  });
}

/**
 * RPC: clears every response on a ballot (rcballot-cck.5). Under the ballot lock: copies the
 * whole sheet to a "Backup-<id>-<yyyyMMdd-HHmm>" sheet (not listed as a ballot — see
 * BALLOT_BACKUP_PREFIX), then deletes all Responses data rows (marker + header stay) and
 * drops orphan columns. Config and Candidates are kept. The typed confirmation must equal the
 * ballot id exactly; it is re-checked here, not only in the page.
 *
 * @param {string} id
 * @param {string} confirmId what the admin typed; must equal id.
 * @return {{ok:boolean, backupSheet:string, responsesCleared:number, orphansDropped:number}}
 */
function adminClearResponses(id, confirmId) {
  rcvRequireBallotAdmin_(id); // rcballot-dx3.3: server-side admin rule (Identity.js)
  return _withBallotLock_(function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = findBallotSheet_(ss, id);
    if (!sheet) throw new Error('No ballot found for id "' + id + '".');
    if (String(confirmId || '').trim() !== String(id)) throw new Error('Confirmation does not match the ballot id.');
    var cleared = 0;
    var dropped = 0;
    var backupName = '';
    var after = _withBallotWrite_(sheet, function (snap) {
      var t = snap.tables.responses;
      if (!t || t.headerRow === null) throw new Error('This ballot has no Responses section to clear.');
      // Backup first; if the copy fails nothing has been cleared.
      backupName = _backupSheetName_(ss, id, new Date());
      sheet.copyTo(ss).setName(backupName);
      var rows = [];
      for (var r = t.headerRow + 1; r <= t.rowsUsed; r++) rows.push(r);
      cleared = _snapResponses_(snap).length;
      var ops = [];
      if (rows.length) ops.push({ section: 'responses', deleteRows: rows });
      _snapOrphanColumns_(snap).forEach(function (o) { ops.push({ section: 'responses', deleteColumn: o.label }); dropped++; });
      return ops.length ? _applyBallotOps_(sheet, snap, ops) : snap;
    });
    refreshBallotCache_(id, sheet, _ballotViewFromSnapshot_(after));
    _logBallotEvent_('ballot.responses.cleared', { sheet: sheet.getName(), backupSheet: backupName, responses: cleared, orphanColumns: dropped });
    return { ok: true, backupSheet: backupName, responsesCleared: cleared, orphansDropped: dropped };
  });
}

/**
 * RPC: adds a new candidate from the "+ Add Candidate" panel.
 *
 * @param {string} id
 * @param {string} name
 * @param {string=} details
 * @param {string=} linkText
 * @param {string=} linkUrl
 * @return {{candidate:string}}
 */
function adminAddCandidate(id, name, details, linkText, linkUrl) {
  rcvRequireBallotAdmin_(id); // rcballot-dx3.3: server-side admin rule (Identity.js)
  return addBallotCandidateForAdmin_(id, name, details, linkText, linkUrl);
}

/**
 * Runs RCV + all four Condorcet methods against a ballot's current
 * candidates/responses, writes a summary into the sheet's Results section,
 * and returns the raw results for HTML rendering.
 *
 * @param {Sheet} sheet a Ballot-<name> sheet
 * @return {{rcv:Object, condorcet:Object}|{error:string}}
 */
function runBallotAnalysis_(sheet) {
  // Read + compute + Results rewrite as one unit, so a concurrent submit can't append a
  // response row while the Results section's rows are being inserted/deleted.
  return _withBallotLock_(function () {
    return _runBallotAnalysisLocked_(sheet);
  });
}

/**
 * Body of runBallotAnalysis_; caller holds the ballot lock.
 *
 * @param {Sheet} sheet
 * @return {{rcv:Object, condorcet:Object}|{error:string}}
 */
function _runBallotAnalysisLocked_(sheet) {
  // A Responses header with no 'Name' label keys no votes: refuse rather than report a
  // ballot with "no responses" (rcballot-8i5.7). The admin edit-page load repairs the known
  // legacy cause (labels on the [Responses] marker row).
  var view = _readBallotView_(sheet);
  if (_responsesMissingName_(view.snap)) {
    _logBallotEvent_('ballot.analysis.refused', { sheet: sheet.getName(), reason: 'responsesMissingName' });
    return { error: 'Ballot sheet layout is damaged: the Responses header has no "Name" column. ' +
      'Open this ballot\'s edit page (which repairs the known cause), then run the analysis again.' };
  }
  // Candidates in Candidates-table order; each ballot's ranks are aligned to it BY NAME, and
  // orphan rank columns (no matching candidate) are not part of any ballot (rcballot-8i5.3).
  var candidateNames = view.names;
  // Deduped to one ballot per respondent (case-insensitive name, latest wins) — a
  // respondent who re-voted must only count once, using their most recent ranking.
  var responseRows = _latestResponseByRespondent_(view.rows);

  if (candidateNames.length < 2) {
    return { error: 'Ballot needs at least two candidates before it can be analyzed.' };
  }
  if (responseRows.length === 0) {
    return { error: 'Ballot has no responses yet.' };
  }

  // A weight of 0 (or negative, read as 0) excludes the voter from analysis (rcballot-8lq).
  var ballots = responseRows.filter(function (r) { return r.weight > 0; }).map(function (r) {
    return { voterName: r.name, ranks: r.ranks, weight: r.weight };
  });
  if (ballots.length === 0) {
    return { error: 'Ballot has no responses with a weight above 0.' };
  }

  var rcv = runRankedChoiceVoting(candidateNames, ballots, null);
  var rcvFinishOrder = computeRCVFinishOrder(candidateNames, ballots);

  var candidates = candidateNames.map(function (name) { return { name: name }; });
  var condorcet = {
    condorcet: findCondorcetWinner(ballots, candidateNames),
    schulze: findSchulzeWinner(ballots, candidateNames),
    rankedPairs: findRankedPairsWinner(ballots, candidateNames),
    minimax: findMinimaxWinner(ballots, candidateNames)
  };
  var finishOrders = {
    condorcet: computeCondorcetFinishOrder(findCondorcetWinner, ballots, candidateNames),
    schulze: computeCondorcetFinishOrder(findSchulzeWinner, ballots, candidateNames),
    rankedPairs: computeCondorcetFinishOrder(findRankedPairsWinner, ballots, candidateNames),
    minimax: computeCondorcetFinishOrder(findMinimaxWinner, ballots, candidateNames)
  };

  writeBallotResults_(sheet, _buildResultsRows_(rcv, condorcet, rcvFinishOrder, finishOrders));

  return { rcv: rcv, condorcet: condorcet, rcvFinishOrder: rcvFinishOrder, finishOrders: finishOrders };
}

/**
 * Renders a finish-order array (as returned by computeRCVFinishOrder /
 * computeCondorcetFinishOrder) as a single string, e.g. "1. C  2. B  3. A",
 * with ties for a place shown as "2. (tie) B / A".
 *
 * @param {Array<{place:number, names:Array<string>}>} order
 * @return {string}
 */
function _formatFinishOrder_(order) {
  return order.map(function (entry) {
    var label = entry.names.length > 1 ? '(tie) ' + entry.names.join(' / ') : entry.names[0];
    return entry.place + '. ' + label;
  }).join('  ');
}

/**
 * @param {Object} rcv result of runRankedChoiceVoting
 * @param {Object} condorcet {condorcet, schulze, rankedPairs, minimax}
 * @param {Array<{place:number, names:Array<string>}>} rcvFinishOrder
 * @param {Object} finishOrders {condorcet, schulze, rankedPairs, minimax} finish orders
 * @return {Array<Array>} rows to write into the Results section.
 */
function _buildResultsRows_(rcv, condorcet, rcvFinishOrder, finishOrders) {
  var rows = [];
  rows.push(['Analysis run: ' + new Date().toISOString()]);
  rows.push(['Method note: every Finish order below (RCV and each Condorcet method) is ' +
    'computed by winner-removal - find the winner, remove them, and re-run on who\'s ' +
    'left for each place - not from a single run\'s elimination order or ranked-candidates ' +
    'score (see rcballot-14e).']);
  rows.push(['']);
  rows.push(['Ranked Choice Voting (RCV)']);
  rows.push([rcv.winner ? 'Winner: ' + rcv.winner : 'Tie: ' + (rcv.tie || []).join(', ')]);
  if (rcvFinishOrder) {
    rows.push(['Finish order (winner-removal, not the final-round runner-up)', _formatFinishOrder_(rcvFinishOrder)]);
  }
  rcv.summary.forEach(function (row) { rows.push(row); });
  rows.push(['']);
  rows.push(['Condorcet Analysis', 'Winner']);
  rows.push(['Basic Condorcet', condorcet.condorcet.winner || 'None (cycle)']);
  rows.push(['Schulze', condorcet.schulze.winner || 'None (cycle)']);
  rows.push(['Ranked Pairs', condorcet.rankedPairs.winner ||
    (condorcet.rankedPairs.tie && condorcet.rankedPairs.tie.length
      ? 'Tie: ' + condorcet.rankedPairs.tie.join(', ')
      : 'None (cycle)')]);
  rows.push(['Minimax', condorcet.minimax.winner ||
    'Tie: ' + (condorcet.minimax.tiedForMinScore || []).join(', ')]);

  if (finishOrders) {
    rows.push(['']);
    rows.push(['Condorcet Analysis', 'Finish order (winner-removal)']);
    rows.push(['Basic Condorcet', _formatFinishOrder_(finishOrders.condorcet)]);
    rows.push(['Schulze', _formatFinishOrder_(finishOrders.schulze)]);
    rows.push(['Ranked Pairs', _formatFinishOrder_(finishOrders.rankedPairs)]);
    rows.push(['Minimax', _formatFinishOrder_(finishOrders.minimax)]);
  }

  ['condorcet', 'schulze', 'rankedPairs', 'minimax'].forEach(function (key) {
    rows.push(['']);
    rows.push([key + ' — ranked candidates (heuristic display score, not finish order)']);
    rows.push(['Rank', 'Candidate', 'Score']);
    condorcet[key].rankedCandidates.forEach(function (rc, i) {
      rows.push([i + 1, rc.candidate, rc.score]);
    });
  });

  return rows;
}

/**
 * @param {*} value
 * @return {string}
 */
function _escapeHtml_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    _buildResultsRows_: _buildResultsRows_,
    _escapeHtml_: _escapeHtml_
  };
}
