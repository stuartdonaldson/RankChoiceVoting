'use strict';

/**
 * ballotFixtures.js — sheet grids copied from live ballots, for regression tests.
 * Each export returns a fresh plain 2D array (grid[0] = row 1; rows may be ragged/empty).
 */

/**
 * PROD Ballot-ABC-2026-7 (rcballot-8i5.7): legacy Responses layout. Rows 59-72 are the live
 * sheet's, verbatim; the config and Results rows above them are stand-ins. The [Responses]
 * marker row (66) carries Name|Weight|Comment in cols B-D; the header row below (67) holds
 * Date + candidate names with B-D blank.
 */
function prodAbc2026_7Rows() {
  const rows = [
    ['Title', 'ABC July'], ['Description', 'D'], ['Instructions', 'I'], ['Footer', 'F'],
    ['Contact', 'C'], ['Accept-New', 'FALSE'], ['Add-Instructions', ''], ['Admin-Only-Notes', ''], [],
    ['[Results]'], ['', 'Analysis run: earlier'],
  ];
  while (rows.length < 58) rows.push([]);
  const names = ['The Death of Ivan Ilyich', 'Man’s Search for Meaning', 'The Foreign Correspondent',
    'Admiral McRaven Commencement Speech', 'Blood in the Machine'];
  rows.push(['[Candidates]']); // 59
  rows.push(['', 'Name', 'Details', 'Link Text', 'Link URL']); // 60
  names.forEach((n) => rows.push(['', n, 'About ' + n])); // 61-65
  rows.push(['[Responses]', 'Name', 'Weight', 'Comment']); // 66
  rows.push(['Date', '', '', ''].concat(names)); // 67
  rows.push(['Sat Jul 11 2026', 'Little John', 1, '', 2, 1, 3, 4, 5]); // 68
  rows.push(['Sat Jul 11 2026', 'Crazy Ivan', 1, '', 1, 2, 5, 3, 4]); // 69
  rows.push(['Sun Jul 12 2026', 'Paul Cooley', 1, '', 1, 2, 4, 3, 5]); // 70
  rows.push(['Sun Jul 12 2026', 'Glory', 1, '', 1, 3, 2, 5, 4]); // 71
  rows.push(['Mon Jul 13 2026', 'Pogo', 1, '', 1, 2, 3, 4, 5]); // 72
  return rows;
}

/** Writes a fixture grid into a FakeSheet (non-empty cells only, so getLastRow matches Sheets). */
function writeGrid(sheet, rows) {
  rows.forEach((r, i) => r.forEach((v, c) => { if (v !== '' && v != null) sheet.getRange(i + 1, c + 1).setValue(v); }));
  return sheet;
}

module.exports = { prodAbc2026_7Rows, writeGrid };
