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

/**
 * SIT Ballot-ABC-2026-08 (rcballot-8i5.3): the Candidates section has NO header row — the
 * row right below [Candidates] (78) is the first candidate — so the snapshot finds no
 * 'Name' label and every Responses rank column is an orphan. Rows 77-92 follow the live
 * sheet (details shortened); config/Results above are stand-ins. "The Boys in the Boat"
 * (row 78) has no Responses column: the old positional reader took its row for a header.
 */
function sitAbc2026_08Rows() {
  const rows = [
    ['Title', 'ABC August'], ['Description', 'D'], ['Instructions', 'I'], ['Footer', 'F'],
    ['Contact', 'C'], ['Accept-New', 'false'], ['Add-Instructions', ''], ['Admin-Only-Notes', ''], [],
    ['[Results]'], ['', 'Analysis run: earlier'],
  ];
  while (rows.length < 76) rows.push([]);
  const ranked = ['Endurance', 'The Righteous Mind', 'Hansel and Gretel + other original fairy tales', 'Joseph',
    'The Feather Thief', 'The Good Life', 'Who Moved My Cheese?'];
  rows.push(['[Candidates]']); // 77
  ['The Boys in the Boat'].concat(ranked).forEach((n) => rows.push(['', n, n + ' - details', 'Amazon', 'https://www.amazon.com/' + encodeURIComponent(n)])); // 78-85
  rows.push(['[Responses]']); // 86
  rows.push(['Date', 'Name', 'Weight', 'Comment'].concat(ranked)); // 87
  rows.push(['Wed Aug 12 2026', 'Little John', 1, '', 1, 2, 6, 7, 5, 4, 3]); // 88
  rows.push(['Wed Aug 12 2026', 'Glory', 1, '', 4, 3, 2, 5, 7, 1, 6]); // 89
  rows.push(['Sat Aug 15 2026', 'MooseJaw', 1, '', 5, 3, 6, 2, 4, 1, 7]); // 90
  rows.push(['Sat Aug 15 2026', 'Crazy Ivan', 1, '', 2, 6, 4, 5, 3, 7, 1]); // 91
  rows.push(['Sun Aug 16 2026', 'Pogo', 1, '', 2, 7, 5, 6, 4, 1, 3]); // 92
  return rows;
}

/** Writes a fixture grid into a FakeSheet (non-empty cells only, so getLastRow matches Sheets). */
function writeGrid(sheet, rows) {
  rows.forEach((r, i) => r.forEach((v, c) => { if (v !== '' && v != null) sheet.getRange(i + 1, c + 1).setValue(v); }));
  return sheet;
}

module.exports = { prodAbc2026_7Rows, sitAbc2026_08Rows, writeGrid };
