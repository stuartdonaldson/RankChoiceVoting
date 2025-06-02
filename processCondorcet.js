function processCondorcetVoting() {
  // Load candidates and ballots using helper functions
  var candidates = loadCandidates();
  if (candidates.length === 0) {
    Logger.log("No candidates found.");
    return;
  }
  var candidateNames = candidates.map(function (c) {
    return c.name;
  });

  var ballots = loadBallots();
  if (ballots.length === 0) {
    Logger.log("No ballots found.");
    return;
  }

  // Basic Condorcet
  var result = findCondorcetWinner(ballots, candidateNames);
  Logger.log("=== Basic Condorcet ===");
  if (result.winner) {
    Logger.log("Condorcet winner: " + result.winner);
  } else {
    Logger.log("No Condorcet winner (cycle detected).");
  }
  Logger.log("Pairwise matrix:");
  Logger.log(result.matrix);

  // Schulze Method
  var schulze = findSchulzeWinner(ballots, candidateNames);
  Logger.log("=== Schulze Method ===");
  if (schulze.winner) {
    Logger.log("Schulze winner: " + schulze.winner);
  } else {
    Logger.log("No Schulze winner (cycle detected).");
  }
  Logger.log("Pairwise matrix:");
  Logger.log(schulze.matrix);
  Logger.log("Strongest paths:");
  Logger.log(schulze.paths);

  // Ranked Pairs
  var rankedPairs = findRankedPairsWinner(ballots, candidateNames);
  Logger.log("=== Ranked Pairs (Tideman) ===");
  if (rankedPairs.winner) {
    Logger.log("Ranked Pairs winner: " + rankedPairs.winner);
  } else {
    Logger.log("No Ranked Pairs winner (cycle detected).");
  }
  Logger.log("Pairwise matrix:");
  Logger.log(rankedPairs.matrix);
  Logger.log("Locked pairs:");
  Logger.log(rankedPairs.locked);

  // Minimax
  var minimax = findMinimaxWinner(ballots, candidateNames);
  Logger.log("=== Minimax (Simpson) ===");
  if (minimax.winner) {
    Logger.log("Minimax winner: " + minimax.winner);
  } else {
    Logger.log("No Minimax winner (tie or cycle detected).");
  }
  Logger.log("Pairwise matrix:");
  Logger.log(minimax.matrix);
  Logger.log("Minimax scores:");
  Logger.log(minimax.scores);

  // return array describing all results
  return {
    condorcet: result,
    schulze: schulze,
    rankedPairs: rankedPairs,
    minimax: minimax,
  };
}

/**
 * Build the pairwise preference matrix.
 * Returns a 2D array where matrix[i][j] is the number of ballots preferring i over j.
 */
function buildPairwiseMatrix(ballots, candidateNames) {
  const n = candidateNames.length;
  const matrix = Array.from({ length: n }, () => Array(n).fill(0));
  ballots.forEach(({ ranks }) => {
    for (let i = 0; i < n; i++) {
      const rankI = ranks[i];
      if (isNaN(rankI)) continue;
      for (let j = 0; j < n; j++) {
        if (i === j) continue;
        const rankJ = ranks[j];
        if (!isNaN(rankJ) && rankI < rankJ) matrix[i][j]++;
      }
    }
  });
  return matrix;
}

/**
 * Find the Condorcet winner (basic).
 *
 * Pairwise Comparison:
 *   - For every pair of candidates (A and B), count how many voters prefer A over B and how many prefer B over A.
 *
 * Check for a Condorcet Winner:
 *   - A candidate is a Condorcet winner if they beat every other candidate in these head-to-head matchups
 *     (i.e., more voters prefer them over each opponent than vice versa).
 *
 * Result:
 *   - If such a candidate exists, they are declared the winner.
 *   - If no candidate beats all others (due to cycles, e.g., A beats B, B beats C, C beats A), 
 *     there is no Condorcet winner.
 */
function findCondorcetWinner(ballots, candidateNames) {
  const matrix = buildPairwiseMatrix(ballots, candidateNames);
  const n = candidateNames.length;
  for (let i = 0; i < n; i++) {
    let isWinner = true;
    for (let j = 0; j < n; j++) {
      if (i !== j && matrix[i][j] <= matrix[j][i]) {
        isWinner = false;
        break;
      }
    }
    if (isWinner) {
      return { winner: candidateNames[i], matrix };
    }
  }
  return { winner: null, matrix };
}

/**
 * Schulze method (strongest paths)
 */
function findSchulzeWinner(ballots, candidateNames) {
  const n = candidateNames.length;
  const d = buildPairwiseMatrix(ballots, candidateNames);

  // Compute strongest paths
  const p = Array.from({ length: n }, () => Array(n).fill(0));
  for (let i = 0; i < n; i++)
    for (let j = 0; j < n; j++)
      if (i !== j) p[i][j] = d[i][j] > d[j][i] ? d[i][j] : 0;

  for (let i = 0; i < n; i++)
    for (let j = 0; j < n; j++)
      if (i !== j)
        for (let k = 0; k < n; k++)
          if (i !== k && j !== k)
            p[j][k] = Math.max(p[j][k], Math.min(p[j][i], p[i][k]));

  // Find winner
  for (let i = 0; i < n; i++) {
    let wins = true;
    for (let j = 0; j < n; j++) {
      if (i !== j && p[i][j] <= p[j][i]) {
        wins = false;
        break;
      }
    }
    if (wins) return { winner: candidateNames[i], matrix: d, paths: p };
  }
  return { winner: null, matrix: d, paths: p };
}

/**
 * Ranked Pairs (Tideman)
 */
function findRankedPairsWinner(ballots, candidateNames) {
  const n = candidateNames.length;
  const d = buildPairwiseMatrix(ballots, candidateNames);

  // List all pairs and sort by strength of victory
  let pairs = [];
  for (let i = 0; i < n; i++)
    for (let j = 0; j < n; j++)
      if (i !== j && d[i][j] > d[j][i])
        pairs.push({ winner: i, loser: j, margin: d[i][j] - d[j][i] });
  pairs.sort((a, b) => b.margin - a.margin);

  // Lock pairs without creating cycles
  const locked = Array.from({ length: n }, () => Array(n).fill(false));
  function createsCycle(winner, loser, visited = []) {
    if (winner === loser) return true;
    visited[loser] = true;
    for (let i = 0; i < n; i++) {
      if (locked[loser][i] && (!visited[i]) && createsCycle(winner, i, visited.slice())) {
        return true;
      }
    }
    return false;
  }
  pairs.forEach(({ winner, loser }) => {
    if (!createsCycle(winner, loser)) locked[winner][loser] = true;
  });

  // Find source of the graph (no arrows pointing to them)
  for (let i = 0; i < n; i++) {
    let isSource = true;
    for (let j = 0; j < n; j++) {
      if (locked[j][i]) {
        isSource = false;
        break;
      }
    }
    if (isSource) return { winner: candidateNames[i], matrix: d, locked: locked };
  }
  return { winner: null, matrix: d, locked: locked };
}

/**
 * Minimax (Simpson)
 */
function findMinimaxWinner(ballots, candidateNames) {
  const n = candidateNames.length;
  const d = buildPairwiseMatrix(ballots, candidateNames);

  // For each candidate, find their worst pairwise defeat
  const scores = [];
  for (let i = 0; i < n; i++) {
    let worst = 0;
    for (let j = 0; j < n; j++) {
      if (i !== j) {
        worst = Math.max(worst, d[j][i]);
      }
    }
    scores.push(worst);
  }

  // Winner is candidate with the smallest worst defeat
  let minScore = Math.min(...scores);
  let winnerIdx = scores.indexOf(minScore);
  if (scores.filter(s => s === minScore).length > 1) {
    return { winner: null, matrix: d, scores: scores };
  }
  return { winner: candidateNames[winnerIdx], matrix: d, scores: scores };
}