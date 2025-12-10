function onOpen() {
  var ui = SpreadsheetApp.getUi();

  // Create menu items
  ui.createMenu('Administrative functions')
      .addItem('Copy all items to all sheets', 'copyValuesAndFormatting')
      .addItem('Create Printable Sheets', 'createPrintableSheetsV2')
      .addItem('Create Picking List', 'createPickingListV2')
      .addToUi();
}

function onEdit() {
  //This function ensures that new sheets created byother users than the document owner will be deleted. It is called on row 17 on edit

  var newSheetName = /^Blad[\d]+$/
  var ssdoc = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ssdoc.getSheets();
  
  // is the change made by the owner ?
  if (Session.getActiveUser().getEmail() == ssdoc.getOwner().getEmail()) {
    return;
  }
  // if not the owner, delete all unauthorised sheets
  for (var i = 0; i < sheets.length; i++) {
    if (newSheetName.test(sheets[i].getName())) {
      ssdoc.deleteSheet(sheets[i]);
    }
  }
}

// jg-250728: used by createPickingList() and createPrintable()
function setAlternatingColours(sheet) {
  for(var i = 1; i<sheet.getMaxRows();i=i+2){
    sheet.getRange(i, 1, 1, sheet.getMaxColumns()).setBackground("#c2f0c2"); 
    sheet.getRange(i+1, 1, 1, sheet.getMaxColumns()).setBackground("white");  
  }
}

// jg-250728: Unused
// Adjust number of rows in target sheet
function adjustSheetRowCount(sheet, desiredRowCount) {
  var currentRowCount = sheet.getMaxRows();
  if (currentRowCount < desiredRowCount) {
    sheet.insertRowsAfter(currentRowCount, desiredRowCount - currentRowCount);
  }
  else if (currentRowCount > desiredRowCount) {
    sheet.deleteRows(desiredRowCount + 1, currentRowCount - desiredRowCount);
  }
}
