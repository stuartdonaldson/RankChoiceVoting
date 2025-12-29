/**
 * Requirement:
 * Processes ranked choice voting (RCV) results from Google Form responses, redistributing votes each round,
 * eliminating the candidate with the fewest first-place votes, and logging each stage.
 * Handles tie-breaking using second-choice and least last-place votes, and logs all actions to the processing sheet.
 * 
 * Form Response Interface:
 * - Column 1: Voter's name (not used in processing).
 * - Columns 2 and onward: Each column represents a candidate and contains the voter's ranking for that candidate.
 *   Example: "1" in a column means the voter ranked that candidate 1st, "2" means the voter ranked that candidate 2nd, etc.

 * @returns {Array<Object>} An array of the top vote-getting candidate(s) after all calculated tie-breakers.
 * If there is more than one value, a tie could not be broken using second-choice and least-favorite votes.
*/
var processingSheet = null;

function processRankedChoiceVotes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var responseSheet = ss.getSheetByName("Candidate Responses");
  var candidateSheet = ss.getSheetByName("Candidates");
  processingSheet = ss.getSheetByName("RCV Processing");

  if (!responseSheet || !candidateSheet || !processingSheet) {
    Logger.log("Error: One or more required sheets not found.");
    return;
  }

  processingSheet.clear();
  processingSheet
    .getRange(1, 1, 1, 3)
    .setValues([["Round", "Eliminated Candidate", "Explanation"]]);

  var candidateNames = candidateSheet
    .getRange(2, 1, candidateSheet.getLastRow() - 1, 1)
    .getValues()
    .flat();

  // Raw response sheet is TimeStamp, Voter Name, Candidate 1 rank, Candidate 2 rank, ...
  if (candidateNames.length === 0) {
    Logger.log("Error: No candidates found in the Candidates sheet.");
    return;
  }
  if (responseSheet.getLastRow() < 2) {
    Logger.log("Error: No responses found in the Candidate Responses sheet.");
    return;
  }
  // Get all responses, starting from the second row (to skip headers)
  // and the second column (to skip the timestamp).
  // The range is from the second row, second column to the last row and
  // the last column (which is candidates.length + 1, because the first column is time stamp and the second column is voter's name).
  // Note: The first row is the header, so we start from the second row.

  var rawResponses = responseSheet
    .getRange(2, 2, responseSheet.getLastRow() - 1, candidateNames.length + 1)
    .getValues();
  // If there are duplicate voters, take only the latest voter responses
  var uniqueResponses = {};
  rawResponses.forEach((response, index) => {
    var voterName = response[0]; // Assuming the voter's name is in the first column
    if (
      !uniqueResponses[voterName] ||
      uniqueResponses[voterName].index < index
    ) {
      uniqueResponses[voterName] = { response, index };
    }
  });
  // After deduplication
  rawResponses = Object.values(uniqueResponses).map((entry) => entry.response);
  // Transform to { voterName, ranks }

  var allBallots = rawResponses.map((response) => ({
    voterName: response[0],
    ranks: response.slice(1),
  }));
  // Compress the rankings for each response
  allBallots = allBallots.map(({ voterName, ranks }) => {
    // Remove blanks, sort, and reassign compressed ranks
    let ranked = [];
    for (let i = 0; i < ranks.length; i++) {
      if (ranks[i] !== "") {
        ranked.push({ idx: i, rank: parseInt(ranks[i], 10) });
      }
    }
    ranked.sort((a, b) => a.rank - b.rank);
    let compressedRanks = Array(ranks.length).fill("");
    ranked.forEach((entry, i) => {
      compressedRanks[entry.idx] = i + 1;
    });
    return { voterName, ranks: compressedRanks };
  });

  // ranks are now compressed, e.g. [1, 2, 3, "", 5] becomes [1, 2, 3, "", 4]
  // with positions of candidates preserved, but ranks compressed

  // Initialize candidate status
  // allCandidates is an array of objects, each object has the following properties:
  // index: the index of the candidate in the candidates array is the index in the ranked array
  var allCandidates = candidateNames.map((name) => ({
    name,
    votes: 0,
    eliminated: false,
    reason: null,
  }));

  var roundNumber = 1;

  /*
   * For each voter's ballot:
   *   - Collect their ranked choices (ignoring blanks).
   *   - Sort the choices by rank (so 1st choice comes first).
   *   - Find the highest-ranked candidate on their ballot who is still in the race (not eliminated).
   *   - Give that candidate one vote.
   *   - Stop after the first valid (non-eliminated) candidate is found.
   */
  function countVotes(ballots) {
    allCandidates.forEach((r) => (r.votes = 0)); // set votes to 0 for each candidate

    ballots.forEach((ballot) => {
      var ranked = [];
      // Collect ranks and their corresponding candidate status, ranks is ordered by candidate index
      for (var c_i = 0; c_i < ballot.ranks.length; c_i++) {
        var rank = ballot.ranks[c_i];
        if (!isNaN(rank)) {
          ranked.push({ candidate: allCandidates[c_i], rank });
        }
      }
      ranked.sort((a, b) => a.rank - b.rank);
      for (var j = 0; j < ranked.length; j++) {
        var r = ranked[j].candidate;
        if (r && !r.eliminated) {
          r.votes++;
          break;
        }
      }
    });
  }
  /**
   * Returns a table (array of arrays) summarizing candidate status.
   * Columns: Candidate, Status, Votes, Elimination Round, Elimination Reason
   * @param {Array<Object>} allCandidates - Array of candidate status objects.
   * @returns {Array<Array>} Table with header and one row per candidate.
   */
  function getCandidateSummaryTable(allCandidates, allBallots) {

    const table = [
      [
        "Candidate",
        "Status",
        "Votes",
        "Elimination Round",
        "Elimination Reason",
      ],
    ];
    // Sort: active first by votes desc, then eliminated by round asc
    const sorted = allCandidates.slice().sort((a, b) => {
      if (a.eliminated && b.eliminated) {
        return b.eliminated.round - a.eliminated.round;
      }
      if (a.eliminated) return 1;
      if (b.eliminated) return -1;
      return b.votes - a.votes;
    });
    sorted.forEach((c) => {
      let status;
      if (c.eliminated) {
        status = "Eliminated";
      } else if (c.votes > (allBallots.length / 2)) {
        status = "Winner";
      } else if (c.votes === 0) {
        status = "No Votes";
      } else {
        status = "Active";
      }
      table.push([
        c.name,
        status,
        c.eliminated ? c.eliminated.votes : c.votes,
        c.eliminated ? c.eliminated.round : "",
        c.eliminated ? c.eliminated.reason : "",
      ]);
    });
    return table;
  }

  function getEffectiveBallots() {
    return allBallots.map((response) => {
      var ranked = [];
      for (var i = 0; i < response.ranks.length; i++) {
        var rank = response.ranks[i];
        if (!isNaN(rank)) {
          ranked.push({ candidate: allCandidates[i], rank });
        }
      }
      ranked = ranked
        .sort((a, b) => a.rank - b.rank)
        .filter((entry) => !entry.candidate.eliminated);

      return ranked.map((entry, index) => ({
        name: entry.candidate.name,
        newRank: index + 1,
      }));
    });
  }

  function logBallots(title, ballotData) {
    logProcess(["", title]);
    var activeCandidates = allCandidates
      .filter((r) => !r.eliminated)
      .map((r) => r.name);

    var header = ["Voter"].concat(activeCandidates);
    logProcess(header);

    ballotData.forEach((voterData, idx) => {
      var row = [`Voter ${idx + 1}`];
      activeCandidates.forEach((cand) => {
        var found = voterData.find((v) => v.name === cand);
        row.push(found ? found.newRank : "");
      });
      logProcess(row);
    });
  }

  function logProgress(roundNumber, candidate) {
    var explanation = candidate.eliminated
      ? `Eliminated - ${candidate.eliminated.reason}, votes are redistributed`
      : `Wins with more than 50% of the votes`;

    logProcess([roundNumber, candidate.name, explanation]);
  }

  var eliminated = null;
  while (true) {
    logBallots(
      roundNumber === 1
        ? "Initial ballots"
        : `Redistributed ballots in round ${roundNumber}`,
      getEffectiveBallots()
    );

    countVotes(allBallots);
  
    // Log the candidate summary table to the processing sheet
    logProcess(["", "Candidate Summary"]);
    getCandidateSummaryTable( allCandidates, allBallots ).forEach((row) => {
      logProcess(row);
    });

    // Count only non-exhausted ballots
    var activeBallots = allBallots.filter((ballot) => {
      // At least one ranked candidate is not eliminated
      return ballot.ranks.some(
        (rank, i) => !isNaN(rank) && !allCandidates[i].eliminated
      );
    });
    var totalVotes = activeBallots.length;
    var winner = allCandidates.find(
      (r) => !r.eliminated && r.votes > totalVotes / 2
    );

    if (winner) {
      logProgress(roundNumber, winner);
      Logger.log(`Winner: ${winner.name}`);
      return {
        winner: winner.name,
        tie: null,
        summary: getCandidateSummaryTable(allCandidates, allBallots),
      };
     }

    var remaining = allCandidates.filter((r) => !r.eliminated);
    var minVotes = Math.min(...remaining.map((r) => r.votes));
    var toEliminate = remaining.filter((r) => r.votes === minVotes);

    if (toEliminate.length === 1) {
      // Only one candidate has the fewest votes, eliminate them
      eliminated = toEliminate[0];
      eliminated.eliminated = {
        round: roundNumber,
        reason: "fewest votes",
        votes: minVotes,
      };
      logProgress(roundNumber, eliminated);
      roundNumber++;
      continue;
    }
    // More than one candidate has the fewest votes, we have a tie
    // eliminate one of the candidates with the fewest first place votes (standard in RCV)
    var { eliminate, reason } = breakTie(
      toEliminate,
      allBallots,
      allCandidates
    );

    if (!eliminate) {
      // Announce tie and stop processing - tieNames is an array of names
      var tieNames = toEliminate.map((c) => c.name);
      logProcess([
        "",
        "",
        `Tie remains after all tie-breakers: ${tieNames.join(
          ", "
        )}. Manual review required for runoff.`,
      ]);
      return {
        winner: null,
        tie: tieNames,
        summary: getCandidateSummaryTable(allCandidates, allBallots),
      };
     }

    eliminate.eliminated = {
      round: roundNumber,
      reason: reason,
      votes: minVotes,
    };
    logProgress(roundNumber, eliminate, getEffectiveBallots());
    roundNumber++;
  }
}

/**
 * Attempts to break a tie among candidates with the fewest votes using a series of tie-breaking strategies:
 * 1. Second choice: Eliminates the candidate with the fewest second-choice votes among tied candidates.
 * 2. Least last-place: If still tied, eliminates the candidate with the fewest last-place votes among tied candidates.
 * 3. If still unresolved, reports the tie for manual resolution.
 * 
 * Logs the tie-breaking process and results to the processing sheet if provided.
 * 
 * @param {Array<Object>} candidates - Array of candidate status objects currently tied for elimination.
 * @param {Array<Object>} ballots - Array of deduplicated and compressed voter responses.
 * @param {Array<Object>} allCandidates - Array of all candidate status objects.
 * @returns {Object|null} The candidate object to eliminate, or null if the tie could not be resolved.
 */
function breakTie(candidates, ballots, allCandidates) {
  let tieMsg = "";
  let tiedCandidates = candidates.slice();

  let leastSecondElim = countTieWinners(
    tiedCandidates,
    "leastSecond",
    ballots,
    allCandidates
  );
  if (leastSecondElim.length === 1) {
    return { eliminate: leastSecondElim[0], reason: "fewest second choice votes" };
  } else {
    // not broken by second choice, logprocess the leastSecondwinners
    tieMsg = "Second choice failed, tie remains: " + leastSecondElim.map((c) => c.name).join(", ");
    logProcess3("", "", tieMsg);
    tiedCandidates = leastSecondElim;
  }

  let mostLastElim = countTieWinners(
    tiedCandidates,
    "mostLastPlace",
    ballots,
    allCandidates
  );
  if (mostLastElim.length === 1) {
    return { eliminate: mostLastElim[0], reason: "most last-place votes" };
  } else {
    // not broken by least last place, logprocess the mostLastPlacecount
    tieMsg = "Most last place failed, tie remains: " +  mostLastElim.map((c) => c.name).join(", ");
    logProcess3("", "", tieMsg);
    tiedCandidates = mostLastElim;
  }

  // 3. Still tied: report tie for manual resolution
  logProcess3("", "", "Manual review required");
  return { eliminate: null, reason: "unable to eliminate a last place tie between: " + tiedCandidates.map((c) => c.name).join(", ")};
}
/**
 * Determines which candidates remain tied after applying a tie-breaker criterion.
 * Depending on the mode, it counts either second-choice or last-place votes for each tied candidate,
 * and returns the subset of candidates who have the maximum (for second choice) or minimum (for least last place) count.
 *
 * @param {Array<Object>} candidates - Array of candidate status objects currently tied.
 * @param {string} mode - The tie-breaking mode: "leastSecond" or "mostLastPlace".
 * @param {Array<Object>} ballots - Array of all deduplicated and compressed voter ballots.
 * @param {Array<Object>} allCandidates - Array of all candidate status objects.
 * @returns {Array<Object>} Array of candidate objects who remain tied after applying the tie-breaker.
 */
// Helper to count which candidates are still tied after a tie-breaker
function countTieWinners(candidates, mode, ballots, allCandidates) {
  function countByRank(rankPosition) {
    // rankPosition is the rank to count, e.g. 2 for second choice, -1 for last place

    // initialize counts[candidateName] = 0 for each candidate
    let counts = {};
    candidates.forEach((c) => (counts[c.name] = 0));

    ballots.forEach((ballot) => {

      // ranked is an ordered array of candidates by their ranks
      let ranked = ballot.ranks
        .map((rank, i) => ({
          candidate: allCandidates[i],
          rank: rank,
        }))
        .filter((e) => !isNaN(e.rank))
        .sort((a, b) => a.rank - b.rank);

      if (ranked.length > Math.abs(rankPosition) - 1) {
        let idx = rankPosition === -1 ? ranked.length - 1 : rankPosition - 1;
        let chosen = ranked[idx];
        if (chosen && counts.hasOwnProperty(chosen.candidate.name)) {
          counts[chosen.candidate.name]++;
        }
      }
    });
    logProcess3(
      "",
      "",
      `By rank ${rankPosition === -1 ? "Last" : rankPosition}: ` +
        Object.entries(counts)
          .map(([k, v]) => `${k}: ${v}`)
          .join(", ")
    );
    return counts;
  }

  let counts, targetValue, filterFn;
  if (mode === "leastSecond") {
    counts = countByRank(2);
    // now set filter to look for the lowest count and return those candidates
    targetValue = Math.min(...Object.values(counts));
    filterFn = (c) => counts[c.name] === targetValue;
  } else if (mode === "mostLastPlace") {
    counts = countByRank(-1);
    targetValue = Math.max(...Object.values(counts));
    filterFn = (c) => counts[c.name] === targetValue;
  } else {
    return candidates; // No tie-breaker applied, return all candidates
  }
  return candidates.filter(filterFn);
}

/**
 * Appends a row to the processing sheet for logging purposes.
 * @param {Array} rowArr - Array of values to log as a row.
 */
function logProcess(rowArr) {
  if (processingSheet) {
    processingSheet.appendRow(rowArr);
  }
}
function logProcess3(a,b,message) {
  if (processingSheet) {
    processingSheet.appendRow([a, b, message]);
  }
}

