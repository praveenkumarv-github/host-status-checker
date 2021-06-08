// convert all formulas to values in every sheet of the Google Sheet
function formulasToValuesGlobal() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(function(sheet) {
    var range = sheet.getDataRange();
    range.copyValuesToRange(sheet, 1, range.getLastColumn(), 1, range.getLastRow());
  });
};