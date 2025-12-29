# Rank Choice Voting (RCV) Google Apps Script Project

This project provides a complete solution for running Ranked Choice Voting (RCV) and multiple Condorcet voting methods using Google Sheets and Google Forms. It includes scripts to generate a ranked-choice voting form, process the results, and display step-by-step elimination rounds and Condorcet analyses.

## Features

- **Google Sheets Add-on Menu:**  
  Adds a custom "Voting and Ballot Tools" menu to your Google Sheet for creating forms and processing results.

- **Automated Google Form Creation:**  
  Generates a Google Form for voters to rank candidates. Candidate names and descriptions are automatically pulled from your sheet and displayed in the form to help voters make informed choices.

- **RCV Processing:**  
  Processes form responses using the Ranked Choice Voting method, including:
  - Multi-round elimination with vote redistribution
  - Automatic tie-breakers (second-choice and last-place votes)
  - Comprehensive candidate status tracking
  - Visual HTML results table with elimination details

- **Condorcet Methods:**  
  Supports multiple Condorcet voting methods with integrated UI dialogs, including:
  - **Basic Condorcet (pairwise comparison)** — Identifies a candidate who beats every other candidate in head-to-head matchups. If no such candidate exists due to a cycle (A beats B, B beats C, C beats A), no winner is declared.
  - **Schulze Method** — Uses the "strongest path" algorithm to resolve cycles by finding the candidate with the strongest chain of victories. Particularly robust for handling complex voting patterns.
  - **Ranked Pairs (Tideman)** — Locks in pairwise victories in order of strength (largest margin first) while avoiding cycles. Creates a definitive ranking by building an acyclic graph of victories.
  - **Minimax (Simpson)** — Selects the candidate whose worst pairwise defeat is the smallest. Minimizes maximum opposition, making it resistant to strategic voting.
  
  Results are displayed in a formatted dialog for easy comparison across methods.

- **Step-by-Step Results:**  
  Outputs detailed elimination rounds to a dedicated "RCV Processing" sheet for transparency.

- **Dialog Results:**  
  Shows the winner(s) or tie status in a user-friendly dialog, including:
  - Winner announcement or tie notification
  - HTML-formatted candidate summary table
  - Status, vote counts, and elimination details for all candidates
  - Works for both RCV and Condorcet analyses
  
## Setup

1. **Clone or copy this repository to your local machine.**

2. **Prepare your Google Sheet:**
   - Create a sheet named `Candidates` with the following columns:
     - **Column 1:** Candidate Name
     - **Column 2:** Candidate Description (including book title, if applicable) - This will be displayed in the form to help voters make informed choices
   - Create a sheet named `RCV Processing` for detailed elimination round logs
   - Optionally, add more sheets as needed.

3. **Deploy scripts:**
   - Use [clasp](https://github.com/google/clasp) to push the scripts to your Apps Script project, or copy them into the Apps Script editor attached to your Google Sheet.

4. **Open your Google Sheet.**
   - You should see a new menu: **Voting and Ballot Tools**.

## Usage

1. **Create the RCV Form:**
   - From the **Voting and Ballot Tools** menu, select **Create Ballot Form**.
   - A dialog will appear with a link to the generated Google Form.
   - The form will include candidate descriptions to help voters make informed choices.
   - Share this link with your voters.

2. **Collect Votes:**
   - Voters fill out the form, ranking all candidates.

3. **Run RCV Analysis:**
   - Once voting is complete, select **Run RCV Analysis** from the menu.
   - A dialog will display:
     - The winner (or tie status if unresolved)
     - A comprehensive candidate summary table showing:
       - Current status (Winner, Active, Eliminated, No Votes)
       - Vote counts
       - Elimination round (if applicable)
       - Reason for elimination
   - For step-by-step ballot redistribution details, view the **RCV Processing** sheet.

4. **Run Condorcet Analysis:**
   - Select **Run Condorcet Analysis** from the menu to see results for all supported Condorcet methods.
   - Results are displayed in an easy-to-read dialog format showing winners for each method:
     - Basic Condorcet (pairwise comparison)
     - Schulze Method
     - Ranked Pairs (Tideman)
     - Minimax (Simpson)
   - Useful for comparing different voting method outcomes on the same ballot data.

## Files

- `onOpen.js` — Adds the custom menu and handles UI dialogs for RCV and Condorcet results.
- `createRCVform.js` — Builds the Google Form from your candidate list, including helper functions `loadCandidates()` and `loadBallots()` for data loading.
- `processRCV.js` — Processes the RCV results with comprehensive tie-breakers, elimination logic, and candidate status tracking.
- `processCondorcet.js` — Implements and analyzes four Condorcet voting methods.
- `cleanResponses.js` — Utility functions for cleaning up old forms and response sheets.
- `testRCV.js` — Contains test cases for validating the RCV logic.
- `.claspignore` — Specifies files/folders to ignore when pushing with clasp.
- `.gitignore` — Specifies files/folders to ignore in Git.
- `README.md` — This file.

## Customization

- **Candidate Descriptions:**  
  Edit the `Candidates` sheet to update candidate names and descriptions. Descriptions are displayed in the voting form to help voters make informed decisions.

- **Tie-breaker Logic:**  
  The script uses second-choice votes (fewest) and last-place votes (most) as sequential tie-breakers. You can modify the `breakTie()` function in `processRCV.js` to adjust this logic.

- **Form Appearance:**  
  Modify the form title, description, and layout in `createRCVform.js` to customize the voting experience.

## Requirements

- Google Sheets
- Google Apps Script (via the Script Editor or clasp)
- Google Forms (created automatically by the script)

## License

MIT License

---

**Questions or issues?**  
Open an issue or submit a pull request!

---

**Note:**  
This project was developed with the assistance of GitHub Copilot AI.

&copy; Stuart Donaldson (F3 Little John - Puget Sound region)