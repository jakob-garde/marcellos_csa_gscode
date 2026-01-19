function CreatePickingListVeggies() {
  console.log("CreatePickingListVeggies()");

  var source = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var destDoc = SpreadsheetApp.create("Picklist_Veggies_" + Date());

  // prepare picking sheet
  pick_sheet = destDoc.getSheets()[0];

  row_start = 4
  //row_cnt = source[0].getMaxRows() - 3;
  // TODO: find row "Mejeri"
  row_cnt = 17 - 3;
  CreatePickingList(pick_sheet, source, row_start, row_cnt);
}

function CreatePickingListDairyMeat() {
  console.log("CreatePickingListDairyMeat()");

  var source = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var destDoc = SpreadsheetApp.create("Picklist_Meat_" + Date());

  // prepare picking sheet
  pick_sheet = destDoc.getSheets()[0];

  // TODO: find row "Mejeri"
  row_start = 18
  row_cnt = source[0].getMaxRows() - 3;
  CreatePickingList(pick_sheet, source, row_start, row_cnt);
}

function CreatePickingList(pick_sheet, source, row_start, row_cnt) {
  // prepare picking list columns and row-categories
  col_start = pick_sheet.getLastColumn() + 1;
  col_cnt = 1;

  src_col_start = 1;

  pick_sheet
    .getRange(4, col_start, row_cnt, col_cnt)
    .setValues(source[0]
                  .getRange(row_start, src_col_start, row_cnt, col_cnt)
                  .getDisplayValues()
              )
    .setFontSize(15)
    .setFontWeight("bold")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  pick_sheet
    .setColumnWidth(1, 400);


  // pull data from each page into the picking list
  for (var i = 0; i < source.length; i++) {
    console.log("Behandlar ark: " + source[i].getName());

    UpdatePickingList(pick_sheet, source[i], row_start, row_cnt);
  }

  // auto-resize totalssheet column
  finalizePickingList(pick_sheet, source.length);
}

// copies values from source totals column
function UpdatePickingList(sheet, souce, row_start, row_cnt) {
  console.log("UpdatePickingList()");

  // leaves space for the "total-totals" at column 2
  total_cols = sheet.getLastColumn();
  col_start = total_cols + 2;
  if (total_cols > 2) {
    col_start = total_cols + 1;
  }

  // copies from row_start to row_end
  sheet
    .getRange( 4, col_start, row_cnt, 1 )
    .setValues( souce
                  //.getRange("B:B")
                  //.getRange(1, 2, sheet.getMaxRows(), 1)
                  .getRange( row_start, 2, row_cnt, 1)
                  .getDisplayValues()
              )
    .setFontSize(15).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

  // trim column title
  var column_name = souce.getName().replace(/grupp /g,"").replace(/Grupp /g,"").replace(/\"/g,"");
  sheet
    .getRange( 2, col_start, 1, 1)
    .setValue( column_name )
    .setFontSize(15)
    .setFontWeight("bold")

  sheet
    .autoResizeColumn(col_start);
}

// sums rows to create the Total column
function finalizePickingList(sheet, group_cnt) {
  console.log("finalizePickingList()");

  // remove Extra Rows And Columns
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  // Remove rows after the last data row
  sheet.deleteRows(lastRow + 1, sheet.getMaxRows() - lastRow);
  // Remove columns after the last data column
  sheet.deleteColumns(lastColumn + 1, sheet.getMaxColumns() - lastColumn);

  setAlternatingColours(sheet);

  sheet.getRange(2, 2, 1, 1)
    .setValue("Total")
    .setFontSize(15)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  var nrows = sheet.getLastRow();
  var ncols = sheet.getLastColumn();
  var first_summable_row = 5;
  for (var i = first_summable_row; i < nrows; i++) {
    var data = sheet.getRange(i, 3, 1, ncols - 2).getValues()[0];
    var row_total = 0;

    for (var j = 0; j < data.length; j++) {
      row_total += Number(data[j]);
    }

    sheet.getRange(i, 2, 1, 1)
      .setValue(row_total)
      .setFontSize(15)
      .setFontWeight("bold")
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center");
  }

  for (var i = 0; i < group_cnt; i++) {
    var col = i + 2;
    sheet
      .autoResizeColumn(col);
    var autosizedWidth = sheet.getColumnWidth(col);
    sheet
      .setColumnWidth(col, autosizedWidth * 1.2);
  }
}
