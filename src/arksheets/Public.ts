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

/**
 * Returns a representation of data in a row.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The spreadsheet to return a row from.
 * @param {GoogleAppsScript.Integer} row The 1-indexed row to return a range for.
 * @return {ArkSheetsRow} Representation of the row data where cells are indexed by column header.
 */
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

/**
 * Returns a range covering all data in a given column.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The spreadsheet to return a column from.
 * @param {GoogleAppsScript.Integer} col The 1-indexed column to return a range for.
 * @return {GoogleAppsScript.Spreadsheet.Range} The range associated with the column.
 */
function getColumnRange(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  col: GoogleAppsScript.Integer
): GoogleAppsScript.Spreadsheet.Range {
  return sheet.getRange(1, col, sheet.getLastRow(), 1);
}

/**
 * Returns a representation of data in a column.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The spreadsheet to return a column from.
 * @param {GoogleAppsScript.Integer} col The 1-indexed column to return a range for.
 * @return {ArkSheetsColumn} Representation of the column data where cells are indexed by row header.
 */
function getColumn(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  col: GoogleAppsScript.Integer
): ArkSheetsColumn {
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

/**
 * Returns a representation of data in the column with the specified header value.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The spreadsheet to return a column from.
 * @param {string} colHeader The header text of the column to return data for.
 * @return {ArkSheetsColumn} Representation of the column data where cells are indexed by row header.
 */
function getColumnByHeader(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  colHeader: string
): ArkSheetsColumn {
  const cols = getHeaderCols_(sheet);
  const colIndex = cols.indexOf(colHeader);
  if (colIndex > -1) {
    return getColumn(sheet, colIndex + 1);
  }
  return {
    'sheet': sheet.getName(),
    'col': -1,
    'rows': {},
  };
}

/**
 * Writes data indexed by row header to a specific column.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The spreadsheet to write to.
 * @param {GoogleAppsScript.Integer} col The 1-indexed column to write to.
 * @param {KeyedData} rows A map of string keys (row headers) to cell data to write.
 */
function writeSparseColumn(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  col: GoogleAppsScript.Integer,
  rows: KeyedData
): void {
  const map = getHeaderRowsMap_(sheet);
  for (var key in rows) {
    if (map.hasOwnProperty(key)) {
      var rowIndex = map[key];
      var cell = rows[key];
      var cellRange = sheet.getRange(rowIndex + 1, col);
      writeRange_(cellRange, cell);
    }
  }
}

/**
 * Writes data indexed by column header to a specific row.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The spreadsheet to write to.
 * @param {GoogleAppsScript.Integer} row The 1-indexed row to write to.
 * @param {KeyedData} columns A map of string keys (column headers) to cell data to write.
 */
function writeSparseRow(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: GoogleAppsScript.Integer,
  columns: KeyedData
): void {
  const map = getHeaderColsMap_(sheet);
  for (var key in columns) {
    if (map.hasOwnProperty(key)) {
      var colIndex = map[key];
      var cell = columns[key];
      var cellRange = sheet.getRange(row, colIndex + 1);
      writeRange_(cellRange, cell);
    }
  }
}

/**
 * Clears a spreadsheet and replaces all of its values with the supplied data.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The spreadsheet to write to.
 * @param {any[][]} data Data to write.  This is a 2D array of values.
 */
function writeSheetValues(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  data: any[][]
): void {
  sheet.clear();
  const range = sheet.getRange(1, 1, data.length, data[0].length);
  range.setValues(data);
}