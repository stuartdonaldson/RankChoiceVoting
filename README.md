# Rank Choice Voting (RCV) Google Apps Script Project

This project provides a complete solution for running Ranked Choice Voting (RCV) elections using Google Sheets and Google Forms. It includes scripts to generate a ranked-choice voting form, process the results, and display step-by-step elimination rounds.

## Features

- **Google Sheets Add-on Menu:**  
  Adds a custom menu to your Google Sheet for creating an RCV form and processing results.

- **Automated Google Form Creation:**  
  Generates a Google Form for voters to rank candidates, with candidate descriptions pulled from your sheet.

- **RCV Processing:**  
  Processes form responses using the Ranked Choice Voting method, including tie-breakers and multi-round elimination.

- **Step-by-Step Results:**  
  Outputs detailed elimination rounds to a dedicated "RCV Processing" sheet for transparency.

- **Dialog Results:**  
  Shows the winner(s) or tie status in a dialog, with instructions for reviewing the elimination process.

## Setup

1. **Clone or copy this repository to your local machine.**

2. **Prepare your Google Sheet:**
   - Create a sheet named `Candidates` with the following columns:
     - **Column 1:** Candidate Name
     - **Column 2:** Candidate Description (including book title, if applicable)
   - Optionally, add more sheets as needed.

3. **Deploy scripts:**
   - Use [clasp](https://github.com/google/clasp) to push the scripts to your Apps Script project, or copy them into the Apps Script editor attached to your Google Sheet.

4. **Open your Google Sheet.**
   - You should see a new menu: **Ranked Choice Voting**.

## Usage

1. **Create the RCV Form:**
   - From the **Ranked Choice Voting** menu, select **Create RCV Form**.
   - A dialog will appear with a link to the generated Google Form. Share this link with your voters.

2. **Collect Votes:**
   - Voters fill out the form, ranking all candidates.

3. **Process RCV Data:**
   - Once voting is complete, select **Process RCV Data** from the menu.
   - A dialog will display the winner or tie status.
   - For detailed elimination rounds, view the **RCV Processing** sheet.

## Files

- `onOpen.js` — Adds the custom menu and handles UI dialogs.
- `createRCVform.js` — Builds the Google Form from your candidate list.
- `processRCV.js` — Processes the RCV results, including tie-breakers and elimination logic.
- `testRCV.js` — Contains test cases for validating the RCV logic.
- `.claspignore` — Specifies files/folders to ignore when pushing with clasp.
- `.gitignore` — Specifies files/folders to ignore in Git.
- `README.md` — This file.

## Customization

- **Candidate Descriptions:**  
  Edit the `Candidates` sheet to update candidate names and descriptions.

- **Tie-breaker Logic:**  
  The script uses second-choice and last-place votes as tie-breakers. You can modify `processRCV.js` to adjust this logic.

## Requirements

- Google Sheets
- Google Apps Script (via the Script Editor or clasp)
- Google Forms (created automatically by the script)

## License

MIT License

---

**Questions or issues?**  
Open an issue or submit a pull request!