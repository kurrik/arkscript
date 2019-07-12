/**
 *  Utilities for storing and retrieving data from the sheet.
 */

/**
 * Returns a range covering all data in a given row.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The spreadsheet to return a row from.
 * @param {GoogleAppsScript.Integer} row The 1-indexed row to return a range for.
 * @return {GoogleAppsScript.Spreadsheet.Range} The range associated with the row.
 */
function getRowRange(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    row: GoogleAppsScript.Integer
  ) : GoogleAppsScript.Spreadsheet.Range {
  return sheet.getRange(row, 1, 1, sheet.getLastColumn());
}

interface ArkSheetsRow {
  sheet: string,
  row: number,
  columns: { [key: string]: ArkCell; }
}
function getRow(
    sheet: GoogleAppsScript.Spreadsheet.Sheet, 
    row: GoogleAppsScript.Integer
  ) : ArkSheetsRow {
  const cols = getHeaderCols_(sheet);
  const frozenColumnCount = sheet.getFrozenColumns();
  const range = getRowRange(sheet, row);
  const data = range.getValues()[0];
  const dataFormulas = range.getFormulas()[0];
  const item = {
    'sheet': sheet.getName(),
    'row': row,
    'columns': {},
  };
  for (var i = 0; i < cols.length; i++) {
    var cell = getCellSchema_(data[i], null, dataFormulas[i]);
    if (i < frozenColumnCount) {
      cell.frozen = true;
    }
    item.columns[cols[i]] = cell;
  }
  return item;
}

function getColumnRange(sheet, col) {
  return sheet.getRange(1, col, sheet.getLastRow(), 1);
}

function getColumn(sheet, col) {
  const rows = getHeaderRows_(sheet);
  const frozenRowCount = sheet.getFrozenRows();
  const range = getColumnRange(sheet, col);
  const data = range.getValues();
  const dataFormulas = range.getFormulas();
  const item = {
    'sheet': sheet.getName(),
    'col': col,
    'rows': {},
  };
  for (var i = 0; i < rows.length; i++) {
    var cell = getCellSchema_(data[i][0], null, dataFormulas[i][0]);
    if (i < frozenRowCount) {
      cell.frozen = true;
    }
    item.rows[rows[i]] = cell;
  }
  return item;
}

function getColumnByHeader(sheet, colHeader) {
  const cols = getHeaderCols_(sheet);
  const colIndex = cols.indexOf(colHeader);
  if (colIndex > -1) {
    return getColumn(sheet, colIndex + 1);
  }
  return {};
}

function writeSparseColumn(sheet, col, data) {
  const map = getHeaderRowsMap_(sheet);
  for (var key in data.rows) {
    if (map.hasOwnProperty(key)) {
      var rowIndex = map[key];
      var cell = data.rows[key];
      var cellRange = sheet.getRange(rowIndex + 1, col);
      writeRange_(cellRange, cell);
    }
  }
}

function writeSparseRow(sheet, row, data) {
  const map = getHeaderColsMap_(sheet);
  for (var key in data.columns) {
    if (map.hasOwnProperty(key)) {
      var colIndex = map[key];
      var cell = data.columns[key];
      var cellRange = sheet.getRange(row, colIndex + 1);
      writeRange_(cellRange, cell);
    }
  }
}

function writeSheetValues(sheet, data) {
  sheet.clear();
  const range = sheet.getRange(1, 1, data.length, data[0].length);
  range.setValues(data);
}