function createPickingListVeggies() {
  console.log("createPickingListVeggies()");

  // TODO: impl.
}

function createPickingListVeggies() {
  console.log("createPickingListVeggies()");

  var srcSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var destDoc = SpreadsheetApp.create("Picking_list_" + Date());

  createTemplate(destDoc, srcSheets[0]);
  templateSheet = destDoc.getSheets()[1];

  // prepare picking sheet
  pickingSheet = destDoc.getSheets()[0];
  prepareTotalSheet(pickingSheet, srcSheets[0]);

  // pull data from each page into the picking list
  for(var i = 0; i < srcSheets.length; i++) {
    console.log("Behandlar ark: " + srcSheets[i].getName());
    updatePickingList(pickingSheet, srcSheets[i]);
  }

  // auto-resize totalssheet column
  finalizeTotalsSheet(pickingSheet, srcSheets.length);

  // clean up the template sheet
  destDoc.deleteSheet(templateSheet);
}

function updatePickingList(sheet, srcSheet) {
  console.log("updatePickingList()");

  // make space for the "total-totals" at column 2
  total_cols = sheet.getLastColumn();
  col = total_cols + 2;
  if (total_cols > 2) {
    col = total_cols + 1;
  }

  sheet
    .getRange( 1, col, srcSheet.getMaxRows(), 1 )
    .setValues( srcSheet.getRange("B:B").getDisplayValues() )
    .setFontSize(15).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

  var column_name = srcSheet.getName().replace(/grupp /g,"").replace(/Grupp /g,"").replace(/\"/g,"");
  sheet
    .getRange( 2, col, 1, 1)
    .setValue( column_name )
    .setFontSize(15)
    .setFontWeight("bold")

  sheet
    .autoResizeColumn(col);
}
