function createPrintableSheets() {
  console.log("createPrintableSheets()");

  var srcSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var destDoc = SpreadsheetApp.create("Printable_sheets_" + Date());

  createTemplate(destDoc, srcSheets[0]);
  template = destDoc.getSheets()[1];

  for(var i = 0; i < srcSheets.length; i++) {
    console.log("Behandlar ark: " + srcSheets[i].getName());

    // TODO: potentially inline into createMemberPages
    createGroupTotalsPage(destDoc, srcSheets[i]);
    createMemberPages(destDoc, srcSheets[i], template); 
  }

  destDoc.deleteSheet(template);
}

function createGroupTotalsPage(destDoc, srcSheet) {
  console.log("createGroupTotalsPage()");

  currentSheetName = "Total order " + srcSheet.getName();
  var destSheet = srcSheet
    .copyTo(destDoc)
    .setName(currentSheetName);

  // Add values of total orders
  destSheet
    .getRange("A:B")
    .setValues( srcSheet.getRange("A:B").getDisplayValues() )
    .setFontSize(15).setFontWeight("bold");

  setAlternatingColours(destSheet);
  destSheet.setFrozenColumns(0);
  destSheet.deleteColumns(3, destSheet.getMaxColumns() - 2);
  destSheet.autoResizeRows(1, destSheet.getMaxRows());

  destSheet.autoResizeColumns(2, 1);
  destSheet
    .getRange(1, 2, destSheet.getMaxRows(), 1) // col 2, the totals values
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
}

function createMemberPages(destDoc, srcSheet, template) {
  var srcRowCount = srcSheet.getMaxRows();

  var veggies = srcSheet
    .getRange("A:A")
    .getValues();
  template
    .getRange(2, 1, srcSheet.getMaxRows(), 1)
    .setValues(veggies)
    .setFontSize(15).setFontWeight("bold")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  template.autoResizeRows(1, template.getMaxRows());

  var numberOfSheets = Math.ceil((srcSheet.getMaxColumns()-2)/2/2);
  for (var i = 0; i < numberOfSheets; i++) {
    var currentSheetName = srcSheet.getName() + " " + (i+1) + " of " + numberOfSheets;
    Logger.log(currentSheetName);
  
    var destSheet = template
      .copyTo(destDoc)
      .setName(currentSheetName); // duplicate template sheet and set name
    var data = srcSheet
      .getRange(1, 3 + 4*i, srcRowCount, 4)
      .getValues();

    destSheet
      .getRange(2, 2, srcRowCount, 4)
      .setValues(data)
      .setFontSize(15).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
    destSheet.getRange(1, 1)
      .setValue(currentSheetName)
      .setFontSize(15).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

    destSheet
      .autoResizeColumns(2, 1);
    destSheet
      .autoResizeColumns(4, 1);
  }
}

function createTemplate(destDoc, srcSheet) {
    srcSheet
        .copyTo(destDoc)
        .setName("template");
    template = destDoc.getSheets()[1];

    var srcRowCount = srcSheet.getMaxRows();

    template.deleteColumns(5, template.getMaxColumns() - 5);
    template.setName("template");

    // set labels/descriptions column
    var data = srcSheet
        .getRange(1, 1, srcRowCount, 1)
        .getValues();
    template
        .getRange(2, 1, srcRowCount, 1)
        .setValues(data)
        .setFontSize(15).setFontWeight("bold")
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}
