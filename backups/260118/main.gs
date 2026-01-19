function onOpen() {
  var ui = SpreadsheetApp.getUi();

  // Create menu items
  ui.createMenu('Administrative functions')
      .addItem('Copy all items to all sheets', 'copyValuesAndFormatting')
      .addItem('Create Family Sheets', 'CreatePrintableSheets')
      .addItem('Create Picking List (Veggies)', 'CreatePickingListVeggies')
      .addItem('Create Picking List (Dairy & Meats)', 'CreatePickingListDairyMeat')
      .addToUi();
}

function onEdit() {
  // This function ensures that new sheets created byother users 
  // than the document owner will be deleted.

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

