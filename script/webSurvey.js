/**
 * webSurvey.js
 *
 * doGet ?cmd=survey&id=<name> — self-contained survey webapp. Respondents
 * enter their name, drag-reorder the candidate list sourced from the
 * "Survey-<name>" sheet, and submit a ranking that is written back as a row
 * in that sheet's Responses section. Re-entering a name already used pulls
 * in that respondent's last ranking for re-ordering. Branding (title,
 * description, footer, contact) is pulled from the sheet's own config rows
 * — see SurveyModel.js for the sheet layout.
 *
 * If no id is given, or the id doesn't match a Survey-<name> sheet, the page
 * renders generic branding explaining that no survey is available.
 *
 * Entry point: _handleSurvey(e), wired from WebApp.js doGet (cmd === 'survey').
 * All other functions here are either internal helpers or RPCs called from
 * webSurveyPage.html via google.script.run.
 */

/**
 * ?cmd=survey&id=<id>                                        — render the page
 * ?cmd=survey&id=<id>&action=data&name=<name>                — read ranking/candidates
 * ?cmd=survey&id=<id>&action=add&phrase=<phrase>              — append a candidate
 * ?cmd=survey&id=<id>&action=submit&name=<name>&order=<json>  — save a ranking
 *
 * @param {Object} e doGet event.
 * @return {HtmlOutput|TextOutput}
 */
function _handleSurvey(e) {
  var params = (e && e.parameter) || {};
  var action = params.action || 'page';
  var id = params.id || '';

  GasLogger.log('webapp.survey', { action: action, id: id, name: maskNameForLog_(params.name) });
  GasLogger.flush();

  if (action === 'page') {
    // GAS wraps webapp output in an outer frame that only reliably honors a
    // mobile viewport when set via addMetaTag — a plain <meta viewport> tag
    // inside the page's own HTML is not enough; without this, phones render
    // the page as a zoomed-out desktop layout.
    var title = id ? _surveyTitleOrDefault_(id) : 'RankChoiceVoting Survey';
    return HtmlService.createHtmlOutputFromFile('webSurveyPage')
      .setTitle(title)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
  }

  try {
    switch (action) {
      case 'config':
        return _surveyJsonResponse(getSurveyForRespondent_(id, '') || { error: 'no_survey' });
      case 'data':
        return _surveyJsonResponse(getSurveyForRespondent_(id, params.name) || { error: 'no_survey' });
      case 'add':
        return _surveyJsonResponse(addSurveyCandidateForId_(id, params.phrase, params.details));
      case 'submit':
        return _surveyJsonResponse(submitSurveyResponse_(id, params.name, JSON.parse(params.order || '[]'), params.comment));
      default:
        return _surveyJsonResponse({ error: 'Unknown survey action: ' + action });
    }
  } catch (err) {
    return _surveyJsonResponse({ error: String(err && err.message ? err.message : err) });
  }
}

/**
 * @param {string} id
 * @return {string}
 */
function _surveyTitleOrDefault_(id) {
  var config = getSurveyForRespondent_(id, '');
  return (config && config.title) || 'RankChoiceVoting Survey';
}

/**
 * Wraps a plain object as a JSON TextOutput.
 *
 * @param {Object} obj
 * @return {TextOutput}
 */
function _surveyJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * RPC: called from webSurveyPage.html once with no name to fetch branding
 * for the current id before the respondent has typed their name.
 *
 * @param {string} id
 * @return {Object}
 */
function getSurveyConfig(id) {
  return getSurveyForRespondent_(id, '') || { error: 'no_survey' };
}

/**
 * RPC: returns the candidate list and saved comment for the survey, ordered
 * by the named respondent's previous ranking if one exists, else in sheet
 * order.
 *
 * @param {string} id
 * @param {string} name
 * @return {Object}
 */
function getSurveyForName(id, name) {
  var result = getSurveyForRespondent_(id, name);
  if (!result) throw new Error('No survey found for id "' + id + '".');
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
function addSurveyTopic(id, phrase, details) {
  return addSurveyCandidateForId_(id, phrase, details);
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
function submitSurveyRanking(id, name, orderedCandidates, comment) {
  return submitSurveyResponse_(id, name, orderedCandidates, comment);
}
