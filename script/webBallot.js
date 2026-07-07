/**
 * webBallot.js
 *
 * doGet ?cmd=ballot&id=<name> — self-contained ballot webapp. Respondents
 * enter their name, drag-reorder the candidate list sourced from the
 * "Ballot-<name>" sheet, and submit a ranking that is written back as a row
 * in that sheet's Responses section. Re-entering a name already used pulls
 * in that respondent's last ranking for re-ordering. Branding (title,
 * description, footer, contact) is pulled from the sheet's own config rows
 * — see BallotModel.js for the sheet layout.
 *
 * If no id is given, or the id doesn't match a Ballot-<name> sheet, the page
 * renders generic branding explaining that no ballot is available.
 *
 * Entry point: _handleBallot(e), wired from WebApp.js doGet (cmd === 'ballot').
 * All other functions here are either internal helpers or RPCs called from
 * webBallotPage.html via google.script.run.
 */

/**
 * ?cmd=ballot&id=<id>                                        — render the page
 * ?cmd=ballot&id=<id>&action=data&name=<name>                — read ranking/candidates
 * ?cmd=ballot&id=<id>&action=add&phrase=<phrase>              — append a candidate
 * ?cmd=ballot&id=<id>&action=submit&name=<name>&order=<json>  — save a ranking
 *
 * @param {Object} e doGet event.
 * @return {HtmlOutput|TextOutput}
 */
function _handleBallot(e) {
  var params = (e && e.parameter) || {};
  var action = params.action || 'page';
  var id = params.id || '';

  GasLogger.log('webapp.ballot', { action: action, id: id, name: maskNameForLog_(params.name) });
  GasLogger.flush();

  if (action === 'page') {
    // GAS wraps webapp output in an outer frame that only reliably honors a
    // mobile viewport when set via addMetaTag — a plain <meta viewport> tag
    // inside the page's own HTML is not enough; without this, phones render
    // the page as a zoomed-out desktop layout.
    var title = id ? _ballotTitleOrDefault_(id) : 'RankChoiceVoting Ballot';
    return HtmlService.createHtmlOutputFromFile('webBallotPage')
      .setTitle(title)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
  }

  try {
    switch (action) {
      case 'config':
        return _ballotJsonResponse(getBallotForRespondent_(id, '') || { error: 'no_ballot' });
      case 'data':
        return _ballotJsonResponse(getBallotForRespondent_(id, params.name) || { error: 'no_ballot' });
      case 'add':
        return _ballotJsonResponse(addBallotCandidateForId_(id, params.phrase, params.details));
      case 'submit':
        return _ballotJsonResponse(submitBallotResponse_(id, params.name, JSON.parse(params.order || '[]'), params.comment));
      default:
        return _ballotJsonResponse({ error: 'Unknown ballot action: ' + action });
    }
  } catch (err) {
    return _ballotJsonResponse({ error: String(err && err.message ? err.message : err) });
  }
}

/**
 * @param {string} id
 * @return {string}
 */
function _ballotTitleOrDefault_(id) {
  var config = getBallotForRespondent_(id, '');
  return (config && config.title) || 'RankChoiceVoting Ballot';
}

/**
 * Wraps a plain object as a JSON TextOutput.
 *
 * @param {Object} obj
 * @return {TextOutput}
 */
function _ballotJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * RPC: called from webBallotPage.html once with no name to fetch branding
 * for the current id before the respondent has typed their name.
 *
 * @param {string} id
 * @return {Object}
 */
function getBallotConfig(id) {
  return getBallotForRespondent_(id, '') || { error: 'no_ballot' };
}

/**
 * RPC: returns the candidate list and saved comment for the ballot, ordered
 * by the named respondent's previous ranking if one exists, else in sheet
 * order.
 *
 * @param {string} id
 * @param {string} name
 * @return {Object}
 */
function getBallotForName(id, name) {
  var result = getBallotForRespondent_(id, name);
  if (!result) throw new Error('No ballot found for id "' + id + '".');
  return result;
}

/**
 * RPC: appends a new candidate to the shared list, with an optional Details note.
 * The client places the new item at the top of the respondent's own ranking —
 * this call only adds the column; ranking position is a client-side concern,
 * persisted on their next submit.
 *
 * @param {string} id
 * @param {string} phrase
 * @param {string=} details
 * @return {{candidate:string}}
 */
function addBallotTopic(id, phrase, details) {
  return addBallotCandidateForId_(id, phrase, details);
}

/**
 * RPC: saves a respondent's ranking and comment.
 *
 * @param {string} id
 * @param {string} name
 * @param {Array<string>} orderedCandidates
 * @param {string=} comment
 * @return {{ok:boolean}}
 */
function submitBallotRanking(id, name, orderedCandidates, comment) {
  return submitBallotResponse_(id, name, orderedCandidates, comment);
}
