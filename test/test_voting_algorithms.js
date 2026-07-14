'use strict';

const assert = require('assert');
const vm = require('vm');
const fs = require('fs');
const path = require('path');

const { createFakeGasGlobals } = require('./fakeGas');

/**
 * Loads processRCV.js and processCondorcet.js into one shared vm context.
 * In real Apps Script both files share a single global namespace (all
 * script/*.js get concatenated), so processCondorcet.js's use of the
 * _isRankedValue() helper defined in processRCV.js only works here if both
 * run in the same context.
 */
function loadVotingAlgorithms() {
  const sandbox = Object.assign({ module: { exports: {} } }, createFakeGasGlobals(null));
  vm.createContext(sandbox);

  const rcvSrc = fs.readFileSync(path.join(__dirname, '..', 'script', 'processRCV.js'), 'utf8');
  vm.runInContext(rcvSrc, sandbox, { filename: 'processRCV.js' });
  const RCV = sandbox.module.exports;

  sandbox.module.exports = {};
  const condorcetSrc = fs.readFileSync(path.join(__dirname, '..', 'script', 'processCondorcet.js'), 'utf8');
  vm.runInContext(condorcetSrc, sandbox, { filename: 'processCondorcet.js' });
  const Condorcet = sandbox.module.exports;

  return { RCV, Condorcet };
}

function makeBallots(candidateNames, rows) {
  return rows.map((ranks, i) => ({ voterName: `Voter ${i + 1}`, ranks, weight: 1 }));
}

// Minimal stand-in for a Sheet: just enough for logProcess()'s appendRow() calls
// so tests can inspect which round/candidate/explanation got logged.
function makeRecordingSheet() {
  const rows = [];
  return {
    rows,
    appendRow(row) { rows.push(row); },
  };
}

// ---------------------------------------------------------------------------
// RCV: basic outcomes
// ---------------------------------------------------------------------------

function testMajorityWinnerRound1() {
  const { RCV } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B', 'C'];
  const ballots = makeBallots(candidateNames, [
    [1, 2, 3],
    [1, 3, 2],
    [1, 2, 3],
    [2, 1, 3],
    [3, 2, 1],
  ]);

  const result = RCV.runRankedChoiceVoting(candidateNames, ballots, null);
  assert.equal(result.winner, 'A');
  assert.equal(result.tie, null);
}

function testEliminationCascadeRedistribution() {
  const { RCV } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B', 'C'];
  const ballots = makeBallots(candidateNames, [
    [1, 2, ''],
    [1, 2, ''],
    ['', 1, ''],
    ['', 1, ''],
    ['', 1, ''],
    ['', '', 1],
    ['', '', 1],
    ['', '', 1],
    ['', '', 1],
  ]);
  // Round 1: A=2, B=3, C=4, total=9, majority>4.5 -> no winner, A (fewest) eliminated.
  // Round 2: A's 2 ballots redistribute to B -> B=5, C=4, total=9 -> B wins.
  const result = RCV.runRankedChoiceVoting(candidateNames, ballots, null);
  assert.equal(result.winner, 'B');

  const aRow = result.summary.find((row) => row[0] === 'A');
  assert.equal(aRow[1], 'Eliminated');
  assert.equal(aRow[3], 1); // eliminated in round 1
}

function testExhaustedBallotsUseActiveWeightForMajority() {
  const { RCV } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B', 'C'];
  const ballots = makeBallots(candidateNames, [
    [1, '', ''],
    [1, '', ''],
    ['', 1, ''],
    ['', 1, ''],
    ['', 1, ''],
    ['', '', 1],
    ['', '', 1],
    ['', '', 1],
    ['', '', 1],
  ]);
  // Round 1: A=2 (exhausted-only ballots), B=3, C=4, total=9, majority 4.5 -> no winner.
  // A eliminated (fewest). A's ballots have no other rank -> exhausted in round 2.
  // Round 2: B=3, C=4; if majority correctly uses ACTIVE ballots (7, not 9),
  // threshold is 3.5 and C=4 wins immediately in round 2.
  const sheet = makeRecordingSheet();
  const result = RCV.runRankedChoiceVoting(candidateNames, ballots, sheet);

  assert.equal(result.winner, 'C');
  const winsInRound2 = sheet.rows.some(
    (row) => row[0] === 2 && row[1] === 'C' && String(row[2]).indexOf('Wins') === 0
  );
  assert.ok(winsInRound2, 'expected C to be declared winner in round 2 using active-ballot majority');
}

function testWeightedBallotsCountInFirstChoice() {
  const { RCV } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B'];
  const ballots = [
    { voterName: 'v1', ranks: [1, 2], weight: 5 },
    { voterName: 'v2', ranks: [2, 1], weight: 1 },
  ];
  // First-choice counting must use ballot weight, not ballot count: A=5, B=1, total=6.
  const result = RCV.runRankedChoiceVoting(candidateNames, ballots, null);
  assert.equal(result.winner, 'A');
}

// Repro-style coverage: raw ranks with gaps (1,3,7) must compress to (1,2,3)
// and drive redistribution identically to already-contiguous ranks.
function testNonContiguousRanksCompressCorrectly() {
  const { RCV } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B', 'C'];
  const ballots = makeBallots(candidateNames, [
    [1, 9, ''],
    [1, 9, ''],
    ['', 1, ''],
    ['', 1, ''],
    ['', 1, ''],
    ['', '', 1],
    ['', '', 1],
    ['', '', 1],
    ['', '', 1],
  ]);
  // Same shape as testEliminationCascadeRedistribution but B's rank is raw "9"
  // instead of "2" - compression must still treat it as B's second choice.
  const result = RCV.runRankedChoiceVoting(candidateNames, ballots, null);
  assert.equal(result.winner, 'B');
}

function testDuplicateRankValuesOnOneBallot() {
  const { RCV } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B', 'C'];
  const ballots = makeBallots(candidateNames, [
    [1, 1, 2], // A and B both marked rank 1 on the same ballot
  ]);
  // Must not throw; stable sort keeps the first-listed duplicate (A) as the
  // effective top choice.
  const result = RCV.runRankedChoiceVoting(candidateNames, ballots, null);
  assert.equal(result.winner, 'A');
}

function testSingleBallotMajority() {
  const { RCV } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B'];
  const ballots = [{ voterName: 'v1', ranks: [1, ''], weight: 1 }];
  const result = RCV.runRankedChoiceVoting(candidateNames, ballots, null);
  assert.equal(result.winner, 'A');
}

function testAllBlankBallotIsUnresolvableTie() {
  const { RCV } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B'];
  const ballots = [{ voterName: 'v1', ranks: ['', ''], weight: 1 }];
  const result = RCV.runRankedChoiceVoting(candidateNames, ballots, null);
  assert.equal(result.winner, null);
  assert.deepEqual(result.tie.slice().sort(), ['A', 'B']);
}

function testTwoCandidateDeadTie() {
  const { RCV } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B'];
  const ballots = makeBallots(candidateNames, [
    [1, 2],
    [2, 1],
  ]);
  const result = RCV.runRankedChoiceVoting(candidateNames, ballots, null);
  assert.equal(result.winner, null);
  assert.deepEqual(result.tie.slice().sort(), ['A', 'B']);
}

// ---------------------------------------------------------------------------
// RCV: tie-break ladder (fewest second-choice -> most last-place -> unresolved)
// ---------------------------------------------------------------------------

function testTieBrokenByFewestSecondChoice() {
  const { RCV } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B', 'C', 'D'];
  const rows = [
    { A: 1, B: 2, C: 3, D: 4 }, // A first, B second
    { A: 2, B: 1, C: 3, D: 4 }, // B first, A second
    { A: 2, B: 3, C: 1, D: 4 }, // C first, A second
    { A: 2, B: 3, C: 1, D: 4 }, // C first, A second
    { A: 2, B: 3, C: 4, D: 1 }, // D first, A second
    { A: 2, B: 3, C: 4, D: 1 }, // D first, A second
  ];
  const ballots = rows.map((r, i) => ({
    voterName: `v${i + 1}`,
    ranks: candidateNames.map((name) => r[name]),
    weight: 1,
  }));

  // First-choice: A=1, B=1, C=2, D=2 -> tie for fewest is A,B.
  // Second-choice tally among A,B across all ballots: A=5 (v2..v6), B=1 (v1).
  // B has fewest second-choice votes -> B is eliminated.
  const result = RCV.runRankedChoiceVoting(candidateNames, ballots, null);
  const bRow = result.summary.find((row) => row[0] === 'B');
  assert.equal(bRow[1], 'Eliminated');
  assert.equal(bRow[4], 'fewest second choice votes');
}

function testTieBrokenByMostLastPlace() {
  const { RCV } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B', 'C', 'D'];
  const rows = [
    { A: 1, B: 2, C: 3, D: 4 }, // A first; B second; D last
    { B: 1, A: 2, C: 3, D: 4 }, // B first; A second; D last
    { C: 1, D: 2, A: 3, B: 4 }, // C first; B last
    { C: 1, D: 2, A: 3, B: 4 }, // C first; B last
    { D: 1, C: 2, A: 3, B: 4 }, // D first; B last
    { D: 1, C: 2, A: 3, B: 4 }, // D first; B last
  ];
  const ballots = rows.map((r, i) => ({
    voterName: `v${i + 1}`,
    ranks: candidateNames.map((name) => r[name]),
    weight: 1,
  }));

  // First-choice: A=1, B=1, C=2, D=2 -> tie for fewest is A,B.
  // Second-choice tally among A,B: row1->B, row2->A, rows3/4->D(skip), rows5/6->C(skip)
  //   => A=1, B=1 -> still tied.
  // Last-place tally among A,B: row1/2 last=D(skip); rows3-6 last=B (4 times); A never last
  //   => A=0, B=4 -> B has the MOST last-place votes -> B eliminated.
  const result = RCV.runRankedChoiceVoting(candidateNames, ballots, null);
  const bRow = result.summary.find((row) => row[0] === 'B');
  assert.equal(bRow[1], 'Eliminated');
  assert.equal(bRow[4], 'most last-place votes');
}

function testUnresolvableTieReturnsNullWinnerWithTieList() {
  const { RCV } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B', 'C', 'D'];
  const rows = [
    { A: 1, B: 2, C: 3, D: 4 }, // A first, B second, D last
    { B: 1, A: 2, C: 3, D: 4 }, // B first, A second, D last
    { C: 1, A: 2, B: 3, D: 4 }, // C first, A second, D last
    { C: 1, B: 2, A: 3, D: 4 }, // C first, B second, D last
    { D: 1, A: 2, B: 3, C: 4 }, // D first, A second, C last
    { D: 1, B: 2, A: 3, C: 4 }, // D first, B second, C last
  ];
  const ballots = rows.map((r, i) => ({
    voterName: `v${i + 1}`,
    ranks: candidateNames.map((name) => r[name]),
    weight: 1,
  }));

  // First-choice: A=1, B=1, C=2, D=2 -> tie for fewest is A,B.
  // Second-choice among A,B: row1->B, row2->A, row3->A, row4->B, row5->A, row6->B => A=3,B=3 (tied).
  // Last-place among A,B: never A or B (always C or D) => A=0,B=0 (tied).
  // Neither tie-breaker resolves it -> manual review, winner:null.
  const result = RCV.runRankedChoiceVoting(candidateNames, ballots, null);
  assert.equal(result.winner, null);
  assert.deepEqual(result.tie.slice().sort(), ['A', 'B']);
}

function testSummaryTableConsistentWithWinner() {
  const { RCV } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B', 'C'];
  const ballots = makeBallots(candidateNames, [
    [1, 2, ''],
    [1, 2, ''],
    ['', 1, ''],
    ['', 1, ''],
    ['', 1, ''],
    ['', '', 1],
    ['', '', 1],
    ['', '', 1],
    ['', '', 1],
  ]);
  const result = RCV.runRankedChoiceVoting(candidateNames, ballots, null);
  assert.equal(result.winner, 'B');

  const winnerRow = result.summary.find((row) => row[0] === result.winner);
  assert.equal(winnerRow[1], 'Winner');

  const otherRows = result.summary.filter(
    (row) => row[0] !== 'Candidate' && row[0] !== result.winner
  );
  otherRows.forEach((row) => {
    assert.notEqual(row[1], 'Winner');
  });
}

// ---------------------------------------------------------------------------
// RCV: bug repros (rcballot-35e, rcballot-c6n)
// ---------------------------------------------------------------------------

// Repro from rcballot-35e: 3 ballots unanimously rank A over B, nobody ranks C.
// A blank rank cell must not be treated as better than an actual rank.
function testUnrankedCandidateCannotWinRCV() {
  const { RCV } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B', 'C'];
  const ballots = makeBallots(candidateNames, [
    [1, 2, ''],
    [1, 2, ''],
    [1, 2, ''],
  ]);

  const result = RCV.runRankedChoiceVoting(candidateNames, ballots, null);
  assert.equal(result.winner, 'A');
}

// Partial ballots: a voter who only ranks some candidates shouldn't have blanks
// misread as top-ranked choices in later rounds/tie-breaks either.
function testPartialBallotsDoNotCorruptRedistribution() {
  const { RCV } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B', 'C', 'D'];
  const ballots = makeBallots(candidateNames, [
    [1, 2, '', ''],
    [1, '', '', ''],
    ['', 1, 2, ''],
    ['', '', 1, ''],
    ['', '', '', ''], // fully blank ballot, exhausted from round 1
  ]);

  const result = RCV.runRankedChoiceVoting(candidateNames, ballots, null);
  // D is never ranked by anyone and must never win or be favored.
  assert.notEqual(result.winner, 'D');
}

// Repro from rcballot-c6n: tie-break must count weighted votes, not raw ballot
// counts, when comparing second-choice votes among tied candidates.
function testTieBreakUsesBallotWeightNotBallotCount() {
  const { RCV } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B', 'C'];
  const ballots = [
    { voterName: 'v1', ranks: [3, 2, 1], weight: 3 },
    { voterName: 'v2', ranks: [2, 3, 1], weight: 1 },
    { voterName: 'v3', ranks: [2, 3, 1], weight: 1 },
    { voterName: 'v4', ranks: [1, 2, 3], weight: 3 },
    { voterName: 'v5', ranks: [2, 1, 3], weight: 3 },
  ];

  const result = RCV.runRankedChoiceVoting(candidateNames, ballots, null);
  const aRow = result.summary.find((row) => row[0] === 'A');
  assert.equal(aRow[1], 'Eliminated');
  assert.equal(aRow[4], 'fewest second choice votes');
}

// ---------------------------------------------------------------------------
// Condorcet family: pairwise matrix
// ---------------------------------------------------------------------------

function testUnrankedCandidateCannotWinCondorcetMethods() {
  const { Condorcet } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B', 'C'];
  const ballots = makeBallots(candidateNames, [
    [1, 2, ''],
    [1, 2, ''],
    [1, 2, ''],
  ]);

  assert.equal(Condorcet.findCondorcetWinner(ballots, candidateNames).winner, 'A');
  assert.equal(Condorcet.findSchulzeWinner(ballots, candidateNames).winner, 'A');
  assert.equal(Condorcet.findRankedPairsWinner(ballots, candidateNames).winner, 'A');
  assert.equal(Condorcet.findMinimaxWinner(ballots, candidateNames).winner, 'A');
}

// Ranked beats unranked; unranked vs unranked contributes nothing to either side.
function testPairwiseMatrixSemantics() {
  const { Condorcet } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B', 'C'];
  const ballots = makeBallots(candidateNames, [
    [1, 2, ''], // A>B, A beats C (unranked), B beats C (unranked)
  ]);

  const matrix = Condorcet.buildPairwiseMatrix(ballots, candidateNames);
  // A beats B (rank 1 < rank 2)
  assert.equal(matrix[0][1], 1);
  assert.equal(matrix[1][0], 0);
  // A (ranked) beats C (unranked)
  assert.equal(matrix[0][2], 1);
  assert.equal(matrix[2][0], 0);
  // B (ranked) beats C (unranked)
  assert.equal(matrix[1][2], 1);
  assert.equal(matrix[2][1], 0);
}

function testUnrankedVsUnrankedAddsNoPreference() {
  const { Condorcet } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B'];
  const ballots = makeBallots(candidateNames, [
    ['', ''], // nobody ranked either candidate
  ]);

  const matrix = Condorcet.buildPairwiseMatrix(ballots, candidateNames);
  assert.equal(matrix[0][1], 0);
  assert.equal(matrix[1][0], 0);
}

function testPairwiseMatrixRespectsWeight() {
  const { Condorcet } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B'];
  const ballots = [
    { voterName: 'v1', ranks: [1, 2], weight: 4 },
    { voterName: 'v2', ranks: [2, 1], weight: 1 },
  ];

  const matrix = Condorcet.buildPairwiseMatrix(ballots, candidateNames);
  assert.equal(matrix[0][1], 4); // A preferred with weight 4
  assert.equal(matrix[1][0], 1); // B preferred with weight 1
}

// ---------------------------------------------------------------------------
// Condorcet family: clear winner, 2-candidate races, textbook divergence
// ---------------------------------------------------------------------------

function testTwoCandidatesOnly() {
  const { Condorcet } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B'];
  const ballots = [
    { voterName: 'v1', ranks: [1, 2], weight: 1 },
    { voterName: 'v2', ranks: [1, 2], weight: 1 },
    { voterName: 'v3', ranks: [1, 2], weight: 1 },
    { voterName: 'v4', ranks: [2, 1], weight: 1 },
  ];

  assert.equal(Condorcet.findCondorcetWinner(ballots, candidateNames).winner, 'A');
  assert.equal(Condorcet.findSchulzeWinner(ballots, candidateNames).winner, 'A');
  assert.equal(Condorcet.findMinimaxWinner(ballots, candidateNames).winner, 'A');
  const rp = Condorcet.findRankedPairsWinner(ballots, candidateNames);
  assert.equal(rp.winner, 'A');
  assert.equal(rp.tie, null);
}

// Classic "Tennessee capital" textbook example: Memphis has the most
// first-choice (plurality) votes, but Nashville is the true Condorcet
// winner - it beats every other city head-to-head. All four Condorcet-family
// methods must agree on Nashville, diverging from the plurality leader.
function testTennesseeCapitalExampleSchulzeDiffersFromPlurality() {
  const { Condorcet } = loadVotingAlgorithms();
  const candidateNames = ['Memphis', 'Nashville', 'Chattanooga', 'Knoxville'];
  const ballots = [
    { voterName: 'g1', ranks: [1, 2, 3, 4], weight: 42 }, // Memphis>Nashville>Chattanooga>Knoxville
    { voterName: 'g2', ranks: [4, 1, 2, 3], weight: 26 }, // Nashville>Chattanooga>Knoxville>Memphis
    { voterName: 'g3', ranks: [4, 3, 1, 2], weight: 15 }, // Chattanooga>Knoxville>Nashville>Memphis
    { voterName: 'g4', ranks: [4, 3, 2, 1], weight: 17 }, // Knoxville>Chattanooga>Nashville>Memphis
  ];

  // Plurality (first-choice) leader is Memphis with 42 of 100 votes.
  const pluralityCounts = {};
  candidateNames.forEach((name) => (pluralityCounts[name] = 0));
  ballots.forEach((b) => {
    b.ranks.forEach((rank, i) => {
      if (rank === 1) pluralityCounts[candidateNames[i]] += b.weight;
    });
  });
  const pluralityLeader = Object.entries(pluralityCounts).sort((a, b) => b[1] - a[1])[0][0];
  assert.equal(pluralityLeader, 'Memphis');

  assert.equal(Condorcet.findCondorcetWinner(ballots, candidateNames).winner, 'Nashville');
  assert.equal(Condorcet.findSchulzeWinner(ballots, candidateNames).winner, 'Nashville');
  assert.equal(Condorcet.findMinimaxWinner(ballots, candidateNames).winner, 'Nashville');
  assert.equal(Condorcet.findRankedPairsWinner(ballots, candidateNames).winner, 'Nashville');
}

// Pairwise exact tie between two candidates (A and B split 1-1 head-to-head,
// both beat C decisively). Also exercises minimax's tied-worst-defeat case:
// A's worst defeat (by B) equals B's worst defeat (by A), so minimax must
// report no winner rather than arbitrarily picking one.
function testPairwiseExactTieAcrossMethods() {
  const { Condorcet } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B', 'C'];
  const ballots = makeBallots(candidateNames, [
    [1, 2, 3], // A>B>C
    [2, 1, 3], // B>A>C
  ]);

  assert.equal(Condorcet.findCondorcetWinner(ballots, candidateNames).winner, null);
  assert.equal(Condorcet.findSchulzeWinner(ballots, candidateNames).winner, null);
  assert.equal(Condorcet.findMinimaxWinner(ballots, candidateNames).winner, null);

  const rp = Condorcet.findRankedPairsWinner(ballots, candidateNames);
  assert.equal(rp.winner, null);
  assert.deepEqual(rp.tie.slice().sort(), ['A', 'B']);
}

// ---------------------------------------------------------------------------
// Ranked Pairs bug repros (rcballot-vi6)
// ---------------------------------------------------------------------------

// Repro from rcballot-vi6: a perfect 3-cycle with equal margins on every
// locked pair used to hand Ranked Pairs a winner purely because of
// candidate-enumeration order (the A->B edge got locked first). Every
// candidate is symmetric here, so no winner should be reported.
function testRankedPairsPerfectCycleReportsNoWinner() {
  const { Condorcet } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B', 'C'];
  const ballots = makeBallots(candidateNames, [
    [1, 2, 3], // A>B>C
    [3, 1, 2], // B>C>A
    [2, 3, 1], // C>A>B
  ]);

  const result = Condorcet.findRankedPairsWinner(ballots, candidateNames);
  assert.equal(result.winner, null);
  assert.ok(Array.isArray(result.tie) && result.tie.length > 1);

  // Other Condorcet methods correctly see no winner for the same cycle
  // (basic Condorcet, Schulze, and minimax all lack a clear winner in a
  // perfectly symmetric 3-cycle).
  assert.equal(Condorcet.findCondorcetWinner(ballots, candidateNames).winner, null);
  assert.equal(Condorcet.findSchulzeWinner(ballots, candidateNames).winner, null);
  assert.equal(Condorcet.findMinimaxWinner(ballots, candidateNames).winner, null);
}

// Repro from rcballot-vi6: a dead pairwise tie between A and B (each beats
// the other on exactly one ballot) locks a graph with two sources. The old
// code silently returned the lower-indexed source; it must instead report
// both candidates as tied.
function testRankedPairsDeadTieReportsBothSources() {
  const { Condorcet } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B', 'C'];
  const ballots = makeBallots(candidateNames, [
    [1, 2, 3], // A>B>C
    [2, 1, 3], // B>A>C
  ]);

  const result = Condorcet.findRankedPairsWinner(ballots, candidateNames);
  assert.equal(result.winner, null);
  assert.deepEqual(result.tie.slice().sort(), ['A', 'B']);
}

// A genuine, unambiguous Ranked Pairs winner (no ties in play) must still
// report a winner with tie:null, i.e. the fix must not turn every result
// into a tie.
function testRankedPairsUnambiguousWinnerStillReported() {
  const { Condorcet } = loadVotingAlgorithms();
  const candidateNames = ['A', 'B', 'C'];
  const ballots = makeBallots(candidateNames, [
    [1, 2, 3],
    [1, 2, 3],
    [1, 3, 2],
  ]);

  const result = Condorcet.findRankedPairsWinner(ballots, candidateNames);
  assert.equal(result.winner, 'A');
  assert.equal(result.tie, null);
}

function run() {
  // RCV core outcomes
  testMajorityWinnerRound1();
  testEliminationCascadeRedistribution();
  testExhaustedBallotsUseActiveWeightForMajority();
  testWeightedBallotsCountInFirstChoice();
  testNonContiguousRanksCompressCorrectly();
  testDuplicateRankValuesOnOneBallot();
  testSingleBallotMajority();
  testAllBlankBallotIsUnresolvableTie();
  testTwoCandidateDeadTie();

  // RCV tie-break ladder
  testTieBrokenByFewestSecondChoice();
  testTieBrokenByMostLastPlace();
  testUnresolvableTieReturnsNullWinnerWithTieList();
  testSummaryTableConsistentWithWinner();

  // RCV bug repros
  testUnrankedCandidateCannotWinRCV();
  testPartialBallotsDoNotCorruptRedistribution();
  testTieBreakUsesBallotWeightNotBallotCount();

  // Condorcet family
  testUnrankedCandidateCannotWinCondorcetMethods();
  testPairwiseMatrixSemantics();
  testUnrankedVsUnrankedAddsNoPreference();
  testPairwiseMatrixRespectsWeight();
  testTwoCandidatesOnly();
  testTennesseeCapitalExampleSchulzeDiffersFromPlurality();
  testPairwiseExactTieAcrossMethods();

  // Ranked Pairs bug repros
  testRankedPairsPerfectCycleReportsNoWinner();
  testRankedPairsDeadTieReportsBothSources();
  testRankedPairsUnambiguousWinnerStillReported();

  console.log('test_voting_algorithms: all tests passed');
}

run();
