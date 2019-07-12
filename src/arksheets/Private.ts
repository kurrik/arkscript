/**
 *  Private helper methods for the Sheets.gs library.
 */

// Returns an array of lowercased column headers in sheet order.
// E.g. [ "project", "4/8/2019", "4/15/2019" ] etc.
function getHeaderCols_(sheet) {
  const cols = [];
  const range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const values = range.getDisplayValues();
  for (var i = 0; i < values[0].length; i++) {
    var value = values[0][i];
    if (value) {
      cols.push(value.toString().toLowerCase());
    } else {
      cols.push("");
    }
  }
  return cols;
}

// Returns an array of lowercased row headers in sheet order.
// E.g. [ "project", "engineers", "run", "ooo" ] etc.
function getHeaderRows_(sheet) {
  const rows = [];
  const range = sheet.getRange(1, 1, sheet.getLastRow(), 1);
  const values = range.getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    var value = values[i][0];
    if (value) {
      rows.push(value.toString().toLowerCase());
    } else {
      rows.push("")
    }
  }
  return rows;
}

// Returns a map of lowercased column header -> row index.
// E.g. { "project" -> 0, "type" -> 1, "4/8/2019" -> 2 } etc.
function getHeaderColsMap_(sheet) {
  const cols = getHeaderCols_(sheet);
  const map = {};
  for (var i = 0; i < cols.length; i++) {
    map[cols[i]] = i;
  }
  return map;
}

// Returns a map of lowercased row header -> row index.
// E.g. { "project" -> 0, "engineers" -> 1, "run" -> 2, "pq" -> 3 } etc.
function getHeaderRowsMap_(sheet) {
  const rows = getHeaderRows_(sheet);
  const map = {};
  for (var i = 0; i < rows.length; i++) {
    map[rows[i]] = i;
  }
  return map;
}

interface ArkCell {
  value: any;
  frozen: boolean;
  hidden: boolean;
  comment: string;
  formula: string;
}

// Note, values not guaranteed to be strings.
// Per https://developers.google.com/apps-script/reference/spreadsheet/range#getvalues
// "The values may be of type Number, Boolean, Date, or String, depending on the value of the cell."
function getCellSchema_(value: any, comment: string, formula: string): ArkCell {
  return {
    'value': value,
    'frozen': false,
    'hidden': false,
    'comment': comment,
    'formula': formula,
  }
}

function writeRange_(range, data) {
  if (data.formula && data.formula !== "") {
    range.setFormula(data.formula.toString());
  } else {
    range.setValue(data.value.toString());
  }
  if (data.comment && data.comment !== "") {
    range.setNote(data.comment);
  }
}
