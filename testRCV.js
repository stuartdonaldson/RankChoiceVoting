function runRCVTestSuite() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var candidateSheet = ss.getSheetByName("Candidates");
  var responseSheet = ss.getSheetByName("Candidate Responses");
  var processingSheet = ss.getSheetByName("RCV Processing");
  var testResultsSheet =
    ss.getSheetByName("RCV Test Results") || ss.insertSheet("RCV Test Results");

  testResultsSheet.clear();
  testResultsSheet.appendRow([
    "Test Name",
    "Expected Winner",
    "Actual Winner",
    "Pass/Fail",
    "Details",
  ]);

  var testCases = [
    {
      name: "Test 1: Simple Majority Winner (No Elimination)",
      candidates: ["Alice", "Bob", "Carol"],
      votes: [
        [1, 2, 3], // Alice, Bob, Carol
        [1, 3, 2], // Alice, Carol, Bob
        [1, 2, 3], // Alice, Bob, Carol
        [1, 3, 2], // Alice, Carol, Bob
        [2, 1, 3], // Bob, Alice, Carol
        [3, 2, 1], // Carol, Bob, Alice
      ],
      expectedWinner: "Alice", // Alice gets 4/6 = 66% > 50%
      description:
        "Alice should win with majority (4 out of 6 votes) in first round",
    },

    {
      name: "Test 2: Single Elimination - Clear Winner After Redistribution",
      candidates: ["Alice", "Bob", "Carol"],
      votes: [
        [1, 2, 3], // Alice, Bob, Carol
        [1, 2, 3], // Alice, Bob, Carol
        [2, 1, 3], // Bob, Alice, Carol
        [2, 1, 3], // Bob, Alice, Carol
        [3, 1, 2], // Carol, Alice, Bob
        [3, 2, 1], // Carol, Bob, Alice
      ],
      expectedWinner: "Bob",
      description:
        "Round 1: eliminate Carol with 1 vote, Round 2: Bob wins with 4 votes after redistribution",
    },

    {
      name: "Test 3: Multiple Elimination Rounds",
      candidates: ["Alice", "Bob", "Carol", "Dave"],
      votes: [
        [1, 2, 3, 4], // Alice, Bob, Carol, Dave
        [2, 1, 3, 4], // Bob, Alice, Carol, Dave
        [3, 2, 1, 4], // Carol, Bob, Alice, Dave
        [4, 3, 2, 1], // Dave, Carol, Bob, Alice
        [4, 1, 2, 3], // Dave, Alice, Bob, Carol
        [1, 4, 3, 2], // Alice, Dave, Carol, Bob
      ],
      expectedWinner: "Carol",
      description:
        "Multiple rounds of elimination should eventually lead to Carol winning",
    },

    {
      name: "Test 4: All Candidates Tied",
      candidates: ["Alice", "Bob", "Carol"],
      votes: [
        [1, 2, 3], // Alice, Bob, Carol
        [2, 1, 3], // Bob, Alice, Carol
        [3, 1, 2], // Carol, Alice, Bob
        [1, 3, 2], // Alice, Carol, Bob
        [2, 3, 1], // Bob, Carol, Alice
        [3, 2, 1], // Carol, Bob, Alice
      ],
      expectedWinner: "Alice,Bob,Carol",
      description:
        "All tied at 2 votes each for 1st, 2nd and last, no eliminations.",
    },

    {
      name: "Test 7: Complex Redistribution yet no winner.",
      candidates: ["Alice", "Bob", "Carol", "Dave"],
      votes: [
        [1, 2, 3, 4], // Alice, Bob, Carol, Dave
        [1, 3, 2, 4], // Alice, Carol, Bob, Dave
        [2, 1, 4, 3], // Bob, Alice, Dave, Carol
        [3, 4, 1, 2], // Carol, Dave, Alice, Bob
        [4, 2, 1, 3], // Dave, Bob, Alice, Carol
        [4, 3, 2, 1], // Dave, Carol, Bob, Alice
        [2, 4, 3, 1], // Bob, Dave, Carol, Alice
        [3, 1, 4, 2], // Carol, Alice, Dave, Bob
      ],
      expectedWinner: "Alice,Bob,Carol,Dave",
      description:
        "Complex redistribution, yet all havae 2 votes at each rank, no winner.",
    },

    {
      name: "Test 8: Carl nad 3/5 votes and wins",
      candidates: ["Alice", "Bob", "Carol"],
      votes: [
        [1, 2, 3], // Alice, Bob, Carol
        [1, 2, 3], // Alice, Bob, Carol
        [3, 2, 1], // Carol, Bob, Alice
        [3, 2, 1], // Carol, Bob, Alice
        [2, 3, 1], // Bob, Carol, Alice
      ],
      expectedWinner: "Carol",
      description:
        "Carol should win with 3 out of 5 votes in the first round.",
    },

    {
      name: "Test 9: Single Voter Decisive",
      candidates: ["Bob", "Alice", "Carol"],
      votes: [
        [1, 2, 3], 
        [2, 1, 3], 
        [3, 1, 2], 
      ],
      expectedWinner: "Alice",
      description:
        "With only 3 voters, Alice should win on first round",
    },

    {
      name: "Test 10: Large Field Elimination",
      candidates: ["Alice", "Bob", "Carol", "Dave", "Eve"],
      votes: [
        [1, 2, 3, 4, 5], // Alice, Bob, Carol, Dave, Eve
        [1, 3, 2, 5, 4], // Alice, Carol, Bob, Eve, Dave
        [2, 1, 4, 3, 5], // Bob, Alice, Dave, Carol, Eve
        [3, 2, 1, 5, 4], // Carol, Bob, Alice, Eve, Dave
        [5, 4, 3, 2, 1], // Eve, Dave, Carol, Bob, Alice
        [4, 5, 1, 2, 3], // Dave, Eve, Alice, Bob, Carol
        [1, 4, 5, 3, 2], // Alice, Dave, Eve, Carol, Bob
        [2, 5, 4, 1, 3], // Bob, Eve, Dave, Alice, Carol
      ],
      expectedWinner: "Alice",
      description: "Large field should eliminate down to Alice as winner",
    },
  ];

  testCases.forEach(function (test) {
    console.log(`\n=== Running ${test.name} ===`);
    console.log(`Description: ${test.description}`);

    // Setup candidates
    candidateSheet.clear();
    candidateSheet.appendRow(["Candidate"]);
    test.candidates.forEach((c) => candidateSheet.appendRow([c]));

    // Setup votes
    responseSheet.clear();
    var headers = ["Timestamp", "Name"].concat(test.candidates);
    responseSheet.appendRow(headers);
    test.votes.forEach((vote, i) => {
      var row = [
        Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone(),
          "yyyy-MM-dd HH:mm:ss"
        ),
        "TestVoter" + (i + 1),
      ].concat(vote);
      responseSheet.appendRow(row);
    });

    // Run the RCV process
    processingSheet.clear();
    var winners = processRankedChoiceVotes();

    // Determine results
    var actualWinner;
    if (!winners || winners.length === 0) {
      actualWinner = "No winner";
      console.log("Result: No winner could be determined");
    } else if (winners.length === 1) {
      actualWinner = winners[0];
      console.log(`Result: ${actualWinner} wins`);
    } else {
      actualWinner = winners.sort().join(",");
      console.log(`Result: Tie between ${actualWinner}`);
    }

    // Check if test passed
    var expectedWinnerSorted = test.expectedWinner.split(",").sort().join(",");
    var actualWinnerSorted = actualWinner.split(",").sort().join(",");
    var pass = actualWinnerSorted === expectedWinnerSorted ? "PASS" : "FAIL";

    // Log results
    var details =
      pass === "PASS"
        ? test.description
        : `Expected: ${test.expectedWinner}, Got: ${actualWinner}. ${test.description}`;

    testResultsSheet.appendRow([
      test.name,
      test.expectedWinner,
      actualWinner,
      pass,
      details,
    ]);

    console.log(`Test Result: ${pass}`);
    if (pass === "FAIL") {
      console.log(`  Expected: ${test.expectedWinner}`);
      console.log(`  Actual: ${actualWinner}`);
    }
  });

  console.log("\n=== RCV Test Suite Complete ===");
  console.log("Check 'RCV Test Results' sheet for detailed outcomes.");
  Logger.log("RCV test suite complete. See 'RCV Test Results' for outcomes.");
}
