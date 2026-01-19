function CreatePrintableSheets() {
  console.log("CreatePrintableSheets()");

  var srcSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var dest_doc = SpreadsheetApp.create("Printable_sheets_" + Date());

  template = CreateTemplate(dest_doc, srcSheets[0]);

  //for(var i = 0; i < srcSheets.length; i++) {
  for(var i = 0; i < 1; i++) {
    console.log("Behandlar ark: " + srcSheets[i].getName());

    src_sheet = srcSheets[i];

    row_start = 3;
    row_end = 17;
    CreateGroupTotalsPage(dest_doc, src_sheet, "Veggies ", row_start, row_end);

    row_start = 18;
    row_end = src_sheet.getMaxRows();
    CreateGroupTotalsPage(dest_doc, src_sheet, "Meat ", row_start, row_end);

    //CreateMemberPages(dest_doc, srcSheets[i], template); 
  }

  dest_doc.deleteSheet(template);
}

function CreateGroupTotalsPage(dest_doc, src_sheet, name_prefix, row_start, row_end) {
  console.log("CreateGroupTotalsPage()");

  name = name_prefix + src_sheet.getName();
  var dest_sheet = src_sheet
    .copyTo(dest_doc)
    .setName(name);

  // Add values of total orders
  row_cnt = row_end - row_start + 1;
  row_first_nonhdr = 3;
  dest_sheet
    .getRange(row_first_nonhdr, 1, row_cnt, 2)
    .setValues( src_sheet
                  .getRange(row_start, 1, row_cnt, 2)
                  .getDisplayValues()
              )
    .setFontSize(15).setFontWeight("bold");

  // trim rows below row_start + row_cnt
  DeleteRowsFrom(dest_sheet, row_cnt + row_first_nonhdr);

  setAlternatingColours(dest_sheet);
  dest_sheet.setFrozenColumns(0);
  dest_sheet.deleteColumns(3, dest_sheet.getMaxColumns() - 2);
  dest_sheet.autoResizeRows(1, dest_sheet.getMaxRows());

  // format totals values
  dest_sheet.autoResizeColumns(2, 1);
  dest_sheet
    //.getRange(1, 2, dest_sheet.getMaxRows(), 1) // col 2, the totals values
    .getRange(1, 2, row_cnt, 1)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
}

function CreateMemberPages(dest_doc, src_sheet, template) {
  console.log("CreateMemberPages()");

  var src_row_cnt = src_sheet.getMaxRows();

  var veggies = src_sheet
    .getRange(1, 1, src_sheet.getMaxRows(), 1)
    .getValues();
  template
    .getRange(2, 1, src_sheet.getMaxRows(), 1)
    .setValues(veggies)
    .setFontSize(15).setFontWeight("bold")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  template.autoResizeRows(1, template.getMaxRows());

  var sheet_cnt = Math.ceil((src_sheet.getMaxColumns()-2)/2/2);
  for (var i = 0; i < sheet_cnt; i++) {
    var name = src_sheet.getName() + " " + (i+1) + " of " + sheet_cnt;
    Logger.log(name);
  
    var dest_sheet = template
      .copyTo(dest_doc)
      .setName(name); // duplicate template sheet and set name
    var data = src_sheet
      .getRange(1, 3 + 4*i, src_row_cnt, 4)
      .getValues();

    dest_sheet
      .getRange(2, 2, src_row_cnt, 4)
      .setValues(data)
      .setFontSize(15).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
    dest_sheet.getRange(1, 1)
      .setValue(name)
      .setFontSize(15).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

    dest_sheet
      .autoResizeColumns(2, 1);
    dest_sheet
      .autoResizeColumns(4, 1);
  }
}

function CreateTemplate(dest_doc, src_sheet) {
  src_sheet
    .copyTo(dest_doc)
    .setName("template");
  template = dest_doc.getSheets()[1];

  var src_row_cnt = src_sheet.getMaxRows();

  template.deleteColumns(5, template.getMaxColumns() - 5);
  template.setName("template");

  // set labels/descriptions column
  var data = src_sheet
    .getRange(1, 1, src_row_cnt, 1)
    .getValues();
  template
    .getRange(2, 1, src_row_cnt, 1)
    .setValues(data)
    .setFontSize(15).setFontWeight("bold")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  return template;
}
