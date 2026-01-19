function copyValuesAndFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var mainsheet = sheets[0]; // first sheet is considerd master/main
  var noOfSheets = sheets.length;
  var fromRange = mainsheet.getRange("A2:B");  // get columns a and b from first sheet

  // only for total column
  var formulas = mainsheet.getRange("B3:B").getFormulas();  

  for (var i = 1; i < noOfSheets; i++) {
    // make sure there are enough columns in target sheet
    for (var j = 0; sheets[i].getMaxRows() < mainsheet.getMaxRows(); j++) {
      sheets[i].appendRow([" "]);
    }

    // make sure there are not to many columns in target sheet
    if(sheets[i].getMaxRows() > mainsheet.getMaxRows()) {
      sheets[i].deleteRows(mainsheet.getMaxRows()+1,sheets[i].getMaxRows()-mainsheet.getMaxRows());
    }

    // make size of column to fit data
    sheets[i].setColumnWidth(1, mainsheet.getColumnWidth(1));
    sheets[i].setColumnWidth(2, mainsheet.getColumnWidth(2));

    // Reset row heights to same as mainsheet
    for (var k = 1; k < sheets[i].getMaxRows(); k++) {
      sheets[i].setRowHeight(k, mainsheet.getRowHeight(k)); 
    }

    copyRangeToRange(fromRange, sheets[i].getRange("A2:B"));
    sheets[i].getRange("B3:B").setFormulas(formulas);
  }
}

function copyRangeToRange(fromRange, toRange) {
  var values = fromRange.getValues();
  var fonts = fromRange.getFontFamilies();
  var colors = fromRange.getFontColors();
  var sizes = fromRange.getFontSizes();
  
  // Använd endast ett anrop för att kopiera alla värden
  toRange.setValues(values);
  toRange.setFontFamilies(fonts);
  toRange.setFontColors(colors);
  toRange.setFontSizes(sizes);
}

function setAlternatingColours(sheet) {
  for(var i = 1; i<sheet.getMaxRows();i=i+2){
    sheet.getRange(i, 1, 1, sheet.getMaxColumns()).setBackground("#c2f0c2"); 
    sheet.getRange(i+1, 1, 1, sheet.getMaxColumns()).setBackground("white");  
  }
}

function DeleteRowsFrom(sheet, row) {
  const row_last = sheet.getMaxRows();
  if (row <= row_last) {
    sheet.deleteRows(row, row_last - row + 1);
  }
}

function TrimGroupName(name) {
  var name = name.replace(/grupp /g,"").replace(/Grupp /g,"").replace(/\"/g,"");
  return name;
}
