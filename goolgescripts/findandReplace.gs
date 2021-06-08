function FindandReplace(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var lastRow = sheet.getLastRow()
  var lastColumn = sheet.getLastColumn()
  var range = sheet.getRange(1, 1, lastRow, lastColumn)
  var to_replace = "Blank";
  var replace_with = "";
  var data  = range.getValues();
 
    var oldValue="";
    var newValue="";
    var cellsChanged = 0;
 
    for (var r=0; r<data.length; r++) {
      for (var i=0; i<data[r].length; i++) {
        oldValue = data[r][i];
        newValue = data[r][i].toString().replace(to_replace, replace_with);
        if (oldValue!=newValue)
        {
          cellsChanged++;
          data[r][i] = newValue;
        }
      }
    }
    range.setValues(data);
}