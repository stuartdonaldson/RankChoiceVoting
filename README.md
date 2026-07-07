# Rank Choice Voting (RCV) Google Apps Script Web App

A multi-survey Ranked Choice Voting (RCV) and Condorcet analysis tool built on Google Sheets and Google Apps Script. Everything — creating a survey, collecting votes, and running analysis — happens through one deployed web app; there is no separate Google Form.

## How it works, in one picture

One Google Spreadsheet holds any number of surveys. Each survey is its own sheet, named `Survey-<id>`, laid out top to bottom in four sections:

```
Row 1..8   Config          Title, Description, Instructions, Footer, Contact,
                            Accept-New, Add-Instructions, Admin-Only-Notes
Row 10     [Results]        marker
Row 11+                     analysis output — overwritten each time analysis runs
Row M      [Candidates]     marker
Row M+1                     header: Name | Details
Row M+2+                     one row per candidate
Row N      [Responses]      marker
Row N+1                     header: Date | Name | Weight | Comment | <candidate columns...>
Row N+2+                     one row per respondent
```

- **Candidates** is the source of truth for who's running. The Responses section's candidate columns are a derived, position-aligned mirror that vote rows key their ranks off of — it's kept in sync automatically whenever candidates are read.
- **Responses** holds one row per respondent, keyed by name (case-insensitive). **Voting again with the same name overwrites that respondent's existing row** rather than adding a second one — only their most recent ranking ever counts. This is also why the admin survey list shows both a raw response-row count and a smaller (or equal) unique-respondent count: they only diverge if someone's name was typed inconsistently, or a sheet was hand-edited to include a genuine duplicate. Analysis always dedupes the same way before computing results.
- **Results** is overwritten every time you run analysis on that survey — it's a snapshot of the last run, not a log.
- Column A is reserved for section markers/config keys only — it is never used for candidate names or respondent data, so a candidate can safely be named anything (including "Responses" or "Results") without corrupting the sheet layout.

## The web app

Deployed once as a Google Apps Script Web App, the same URL serves three views:

| URL | Purpose |
|---|---|
| `<url>` or `<url>?cmd=admin` | Survey list — create a survey, and for each existing one: view it, edit it, or run analysis |
| `<url>?cmd=admin&action=edit&id=<id>` | Edit a survey: a live preview styled exactly like the real ballot, with a pencil icon on every editable field. Each field saves immediately (no page-wide Save) |
| `<url>?cmd=survey&id=<id>` | The respondent-facing ballot — what you share with voters |

From the spreadsheet itself, **Voting and Ballot Tools > Open Survey Admin Page** opens the admin list, and **About** shows the deployed web app URL plus a direct link to every survey.

### Creating and configuring a survey

1. From the admin page, create a new survey with an id (used in the URL and the `Survey-<id>` sheet name).
2. Open its edit page and use the pencil icons to set:
   - **Title / Description** — shown on the landing page before a respondent enters their name.
   - **Instructions** — shown above the ranking list on the ballot page itself (replaces Description there).
   - **Footer / Contact** — shown at the bottom of the ballot page.
   - **Candidates** — add candidates and an optional admin-only "Details" note shown to respondents.
   - **Accept new candidates from respondents** — if on, respondents can add a candidate themselves from the ballot page; **Add-Instructions** is the text shown above that button.
   - **Admin-Only Notes** — free-text notes for your own reference (purpose, audience, scheduling); never shown to respondents.
3. Share the `?cmd=survey&id=<id>` link with voters.

### Voting

A respondent enters their name, drag-reorders the candidate list (most preferred on top), optionally adds a candidate (if enabled) or leaves feedback, and submits. Returning with the same name loads their previous ranking for review or change — submitting again replaces it.

If a candidate is added after someone has already voted, the admin survey list flags that respondent (they may want to come back and rank the addition) — their existing ranking is otherwise left as-is.

### Running analysis

From the admin list, **Run Analysis** computes and displays:

- **Ranked Choice Voting (RCV)** — multi-round elimination with vote redistribution, automatic tie-breakers, and a full candidate status/round summary.
- **Condorcet methods** — Basic Condorcet (pairwise), Schulze (strongest path), Ranked Pairs (Tideman), and Minimax (Simpson), for comparing outcomes across methods when there's no unambiguous pairwise winner.

Each run also overwrites the survey sheet's own `[Results]` section with the same summary, so the sheet always reflects the last analysis you ran.

## Files

- `script/WebApp.js` — `doGet`/`doPost` router (`cmd=survey`, `cmd=admin`).
- `script/SurveyModel.js` — the sheet-layout model described above; every other file reads/writes surveys through this module.
- `script/webAdmin.js` — admin survey list, create-survey form, and analysis view.
- `script/webAdminEditPage.html` — the live-preview survey editor.
- `script/webSurvey.js` / `script/webSurveyPage.html` — the respondent-facing ballot page and its RPCs.
- `script/processRCV.js` — RCV elimination logic and tie-breakers.
- `script/processCondorcet.js` — the four Condorcet methods.
- `script/onOpen.js` — the spreadsheet's custom menu and About dialog.
- `script/GasLogger.js` — structured logging (see its header for setup).
- `script/version.js` — stamped automatically by `tools/manage-deployments.js` on every deploy.
- `test/` — Node-based unit tests that exercise the sheet model against a fake Sheets API (`test/fakeGas.js`), so core logic is tested without needing a live Apps Script deployment.
- `tools/` — deployment (`manage-deployments.js`) and smoke-test (`smokeTest.js`, `callWebapp.js`) scripts.

## Setup

1. **Clone this repository.**
2. **Install [clasp](https://github.com/google/clasp)** and authenticate it against your Google account.
3. Push `script/` to an Apps Script project bound to your target Google Sheet: `clasp push`.
4. **Deploy as a Web App** (Deploy > New deployment > Web app), or use `tools/manage-deployments.js` if you're using this project's deployment pipeline — it also stamps `script/version.js` and the `WEBAPP_URL` script property used to build in-app links.
5. Open the spreadsheet — you should see the **Voting and Ballot Tools** menu.

## Requirements

- Google Sheets
- Google Apps Script (via `clasp` or the Script Editor)

## License

MIT License

---

**Questions or issues?** Open an issue or submit a pull request!

&copy; Stuart Donaldson (F3 Little John - Puget Sound region)
