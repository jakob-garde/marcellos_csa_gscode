function CreatePrintableSheets() {
  console.log("CreatePrintableSheets()");

  var src_sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var dest_doc = SpreadsheetApp.create("Printable_sheets_" + Date());

  template = CreateTemplate(dest_doc, src_sheets[0]);

  const row_mejeri = FindRow(src_sheets[0], "Mejeri");
  console.log("mejeri row: ", row_mejeri);

  for(var i = 0; i < src_sheets.length; i++) {
    console.log("Behandlar ark: " + src_sheets[i].getName());

    src_sheet = src_sheets[i];

    row_start = 3;
    row_end = row_mejeri;
    CreateGroupTotalsPage(dest_doc, src_sheet, "Veggies ", row_start, row_end);

    row_start = row_mejeri;
    row_end = src_sheet.getMaxRows() + 1;
    CreateGroupTotalsPage(dest_doc, src_sheet, "Meat ", row_start, row_end);

    row_start = 3;
    row_end = row_mejeri;
    CreateMemberPages(dest_doc, src_sheet, template, "Vegg_", row_start, row_end);

    row_start = row_mejeri;
    row_end = src_sheet.getMaxRows() + 1;
    CreateMemberPages(dest_doc, src_sheet, template, "Meat_", row_start, row_end);
  }

  dest_doc.deleteSheet(template);
}

function CreateMemberPages(dest_doc, src_sheet, template, name_prefix, row_start, row_end) {
  console.log("CreateMemberPages()");

  const row_cnt = row_end - row_start;

  // set group-specific header titles (specific instructions for individual groups may exist)
  template
    .getRange(2, 1, 2, 1)
    .setValues(src_sheet
                .getRange(1, 1, 2, 1)
                .getValues()
              )
    .setFontSize(15).setFontWeight("bold")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  template
    .autoResizeRows(1, template.getMaxRows());

  // set group-specific item titles
  template
    .getRange(4, 1, row_cnt, 1)
    .setValues(src_sheet
                .getRange(row_start, 1, row_cnt, 1)
                .getValues()
              )
    .setFontSize(15).setFontWeight("bold")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  template
    .autoResizeRows(1, template.getMaxRows());

  var group_cnt = Math.ceil((src_sheet.getMaxColumns()-2)/2/2);
  for (var i = 0; i < group_cnt; i++) {
  //for (var i = 0; i < 1; i++) {
    var name = name_prefix + TrimGroupName(src_sheet.getName()) + " " + (i+1) + " of " + group_cnt;
    Logger.log(name);

    var dest_sheet = template
      .copyTo(dest_doc)
      .setName(name); // duplicate template sheet and set name

    dest_sheet.deleteRows(3, 1);

    // set family names
    dest_sheet
      .getRange(2, 2, 1, 4)
      .setValues(src_sheet
                    .getRange(2, 3 + 4*i, 1, 4)
                    .getDisplayValues()
      )
      .setFontSize(15).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);

    // set family data
    dest_sheet
      .getRange(3, 2, row_cnt, 4)
      .setValues(src_sheet
                    .getRange(row_start, 3 + 4*i, row_cnt, 4)
                    .getValues() // TODO: why are we not using getDisplayValues() here?
      )
      .setFontSize(15).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);

    DeleteRowsFrom(dest_sheet, row_cnt + 3);

    // set the group name as cell(1,1)
    dest_sheet.getRange(1, 1)
      .setValue(name)
      .setFontSize(15).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

    dest_sheet
      .autoResizeColumns(2, 1);
    dest_sheet
      .autoResizeColumns(4, 1);
  }
}

function CreateGroupTotalsPage(dest_doc, src_sheet, name_prefix, row_start, row_end) {
  console.log("CreateGroupTotalsPage()");

  name = name_prefix + TrimGroupName(src_sheet.getName());
  var dest_sheet = src_sheet
    .copyTo(dest_doc)
    .setName(name);

  // Add values of total orders
  const row_cnt = row_end - row_start;
  const row_first_nonhdr = 3;
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
    .getRange(1, 2, row_cnt, 1)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
}

function CreateTemplate(dest_doc, src_sheet) {
  src_sheet
    .copyTo(dest_doc)
    .setName("template");
  dest_doc.deleteSheet(dest_doc.getSheets()[0]);
  template = dest_doc.getSheets()[0];
  template.setName("template");

  template
    .deleteColumns(2, 1);
  template
    .deleteColumns(5, template.getMaxColumns() - 5);
  template
    .getRange(1, 2, template.getMaxRows(), 4)
    .clearContent();

  return template;
}
