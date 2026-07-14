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

function run() {
  testUnrankedCandidateCannotWinRCV();
  testUnrankedCandidateCannotWinCondorcetMethods();
  testPairwiseMatrixSemantics();
  testUnrankedVsUnrankedAddsNoPreference();
  testPartialBallotsDoNotCorruptRedistribution();
  testTieBreakUsesBallotWeightNotBallotCount();
  console.log('test_voting_algorithms: all tests passed');
}

run();
