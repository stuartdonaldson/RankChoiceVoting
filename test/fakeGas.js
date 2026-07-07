'use strict';

/**
 * fakeGas.js — reusable in-memory shim for the small slice of the Google Apps
 * Script API this project's script/*.js files depend on (SpreadsheetApp,
 * PropertiesService, Logger, Utilities, LockService). Lets script/BallotModel.js
 * (and friends) run unmodified under plain Node via vm.runInContext.
 *
 * A FakeSheet stores cell values in a flat `row_col` map plus tracked
 * lastRow/lastCol counters — deliberately not a dense 2D array, so sparse
 * writes (e.g. a single setValue far down the sheet) don't allocate empty
 * rows.
 */

function cellKey_(row, col) {
  return row + '_' + col;
}

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows == null ? 1 : numRows;
    this.numCols = numCols == null ? 1 : numCols;
  }

  getValue() {
    return this.sheet._getCell(this.row, this.col);
  }

  setValue(value) {
    this.sheet._setCell(this.row, this.col, value);
    return this;
  }

  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowOut = [];
      for (let c = 0; c < this.numCols; c++) {
        rowOut.push(this.sheet._getCell(this.row + r, this.col + c));
      }
      out.push(rowOut);
    }
    return out;
  }

  setValues(values) {
    values.forEach((rowValues, r) => {
      rowValues.forEach((value, c) => {
        this.sheet._setCell(this.row + r, this.col + c, value);
      });
    });
    return this;
  }

  clearContent() {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) {
        this.sheet._setCell(this.row + r, this.col + c, '');
      }
    }
    return this;
  }

  // Formatting is a no-op store (tests only need these to be chainable and not throw);
  // nothing in this project's logic branches on cell formatting.
  setBackground(color) { this._background = color; return this; }
  setFontWeight(weight) { this._fontWeight = weight; return this; }
}

class FakeSheet {
  constructor(name) {
    this._name = name;
    this._cells = new Map();
    this._lastRow = 0;
    this._lastCol = 0;
    this._frozenRows = 0;
  }

  _getCell(row, col) {
    const v = this._cells.get(cellKey_(row, col));
    return v === undefined ? '' : v;
  }

  _setCell(row, col, value) {
    this._cells.set(cellKey_(row, col), value);
    if (row > this._lastRow) this._lastRow = row;
    if (col > this._lastCol) this._lastCol = col;
  }

  getName() { return this._name; }
  setName(name) { this._name = name; return this; }
  isSheetHidden() { return false; }
  getIndex() {
    return this._spreadsheet ? this._spreadsheet.getSheets().indexOf(this) + 1 : 1;
  }

  getRange(row, col, numRows, numCols) {
    return new FakeRange(this, row, col, numRows, numCols);
  }

  getLastRow() { return this._lastRow; }
  getLastColumn() { return this._lastCol; }

  getDataRange() {
    return this.getRange(1, 1, Math.max(this._lastRow, 1), Math.max(this._lastCol, 1));
  }

  setFrozenRows(n) { this._frozenRows = n; return this; }
  getFrozenRows() { return this._frozenRows; }

  /** Shifts every row after `afterRow` down by `howMany`, growing lastRow. */
  insertRowsAfter(afterRow, howMany) {
    const moved = new Map();
    for (const [key, value] of this._cells.entries()) {
      const [rowStr, colStr] = key.split('_');
      const row = parseInt(rowStr, 10);
      const col = parseInt(colStr, 10);
      const newRow = row > afterRow ? row + howMany : row;
      moved.set(cellKey_(newRow, col), value);
    }
    this._cells = moved;
    if (this._lastRow > afterRow) this._lastRow += howMany;
    return this;
  }

  /** Shifts row `beforeRow` and everything below it down by `howMany`. */
  insertRowsBefore(beforeRow, howMany) {
    return this.insertRowsAfter(beforeRow - 1, howMany);
  }

  /** Removes rows [startRow, startRow+howMany-1] and shifts rows below up. */
  deleteRows(startRow, howMany) {
    const endRow = startRow + howMany - 1;
    const moved = new Map();
    for (const [key, value] of this._cells.entries()) {
      const [rowStr, colStr] = key.split('_');
      const row = parseInt(rowStr, 10);
      const col = parseInt(colStr, 10);
      if (row >= startRow && row <= endRow) continue; // deleted
      const newRow = row > endRow ? row - howMany : row;
      moved.set(cellKey_(newRow, col), value);
    }
    this._cells = moved;
    if (this._lastRow >= startRow) this._lastRow = Math.max(0, this._lastRow - howMany);
    return this;
  }
}

class FakeSpreadsheet {
  constructor() {
    this._sheets = [];
  }

  insertSheet(name) {
    if (this.getSheetByName(name)) {
      throw new Error('A sheet with the name "' + name + '" already exists. Please enter another name.');
    }
    const sheet = new FakeSheet(name);
    sheet._spreadsheet = this;
    this._sheets.push(sheet);
    return sheet;
  }

  getSheetByName(name) {
    return this._sheets.find(s => s.getName() === name) || null;
  }

  getSheets() {
    return this._sheets.slice();
  }

  deleteSheet(sheet) {
    const idx = this._sheets.indexOf(sheet);
    if (idx !== -1) this._sheets.splice(idx, 1);
  }
}

function createFakeSpreadsheet() {
  return new FakeSpreadsheet();
}

/** In-memory Script Properties store — one Map per call, mirrors PropertiesService semantics. */
function createFakePropertiesService() {
  const store = new Map();
  const properties = {
    getProperty(key) {
      return store.has(key) ? store.get(key) : null;
    },
    setProperty(key, value) {
      store.set(key, String(value));
      return properties;
    },
    setProperties(obj) {
      Object.keys(obj || {}).forEach(k => store.set(k, String(obj[k])));
      return properties;
    },
    deleteProperty(key) {
      store.delete(key);
      return properties;
    },
  };
  return {
    getScriptProperties: () => properties,
    _store: store,
  };
}

let _uuidCounter = 0;
function createFakeUtilities() {
  return {
    getUuid() {
      _uuidCounter += 1;
      return 'fake-uuid-' + _uuidCounter;
    },
  };
}

function createFakeLockService() {
  return {
    getScriptLock() {
      return {
        waitLock() {},
        releaseLock() {},
      };
    },
  };
}

function createFakeLogger() {
  const lines = [];
  return {
    log(msg) { lines.push(msg); },
    _lines: lines,
  };
}

/**
 * @param {FakeSpreadsheet=} ss - Spreadsheet returned by SpreadsheetApp.getActiveSpreadsheet().
 *   Pass the same instance the test itself holds so both sides see the same sheets.
 * @return {Object} globals to spread into a vm sandbox: SpreadsheetApp, PropertiesService,
 *   Logger, Utilities, LockService.
 */
function createFakeGasGlobals(ss) {
  const activeSpreadsheet = ss || createFakeSpreadsheet();
  return {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => activeSpreadsheet,
    },
    PropertiesService: createFakePropertiesService(),
    Logger: createFakeLogger(),
    Utilities: createFakeUtilities(),
    LockService: createFakeLockService(),
  };
}

module.exports = {
  FakeRange,
  FakeSheet,
  FakeSpreadsheet,
  createFakeSpreadsheet,
  createFakePropertiesService,
  createFakeUtilities,
  createFakeLockService,
  createFakeLogger,
  createFakeGasGlobals,
};
