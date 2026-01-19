function createTemplate(destDoc, srcSheet) {
    srcSheet
        .copyTo(destDoc)
        .setName("template");
    templateSheet = destDoc.getSheets()[1];

    var srcRowCount = srcSheet.getMaxRows();

    templateSheet.deleteColumns(5, templateSheet.getMaxColumns() - 5);
    templateSheet.setName("template");

    // set labels/descriptions column
    var data = srcSheet
        .getRange(1, 1, srcRowCount, 1)
        .getValues();
    templateSheet
        .getRange(2, 1, srcRowCount, 1)
        .setValues(data)
        .setFontSize(15).setFontWeight("bold")
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}

function prepareTotalSheet(totalSheet, srcSheet0) {
  totalSheet
    .getRange( 4, totalSheet.getLastColumn() + 1, srcSheet0.getMaxRows() - 3, 1 )
    .setValues( srcSheet0.getRange("A4:A").getDisplayValues() )
    .setFontSize(15)
    .setFontWeight("bold")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  totalSheet
    .setColumnWidth(1, 400);
}

function finalizeTotalsSheet(totalSheet, srcSheetsLength) {
  console.log("finalizeTotalsSheet()");

  // remove Extra Rows And Columns
  {
    var lastRow = totalSheet.getLastRow();
    var lastColumn = totalSheet.getLastColumn();

    // Remove rows after the last data row
    totalSheet.deleteRows(lastRow + 1, totalSheet.getMaxRows() - lastRow);

    // Remove columns after the last data column
    totalSheet.deleteColumns(lastColumn + 1, totalSheet.getMaxColumns() - lastColumn);
  }

  setAlternatingColours(totalSheet);

  totalSheet.getRange(2, 2, 1, 1)
    .setValue("Total")
    .setFontSize(15)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  var nrows = totalSheet.getLastRow();
  var ncols = totalSheet.getLastColumn();
  var first_summable_row = 5;
  for (var i = first_summable_row; i < nrows; i++) {
    var data = totalSheet.getRange(i, 3, 1, ncols - 2).getValues()[0];
    var row_total = 0;

    for (var j = 0; j < data.length; j++) {
      row_total += Number(data[j]);
    }

    totalSheet.getRange(i, 2, 1, 1)
      .setValue(row_total)
      .setFontSize(15)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
  }

  for(var i = 0; i < srcSheetsLength; i++) {
    var col = i + 2;
    totalSheet
      .autoResizeColumn(col);
    var autosizedWidth = totalSheet.getColumnWidth(col);
    totalSheet
      .setColumnWidth(col, autosizedWidth * 1.2);
  }
}

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
