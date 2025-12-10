function createPrintableSheetsV2() {
    console.log("createPrintableSheetsV2()");

    var srcSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var destDoc = SpreadsheetApp.create("Printable_sheets_" + Date());

    createTemplate(destDoc, srcSheets[0]);
    templateSheet = destDoc.getSheets()[1];

    // prepare totals sheet
    totalSheet = destDoc.getSheets()[0];
    prepareTotalSheet(totalSheet, srcSheets[0]);

    for(var i = 0; i < srcSheets.length; i++) {
        console.log("Behandlar ark: " + srcSheets[i].getName());
        updateTotalsSheet(totalSheet, srcSheets[i]); 
        createGroupTotalsPage(destDoc, srcSheets[i]);
        createMemberPages(destDoc, srcSheets[i], templateSheet); 
    }

    // auto-resize totalssheet column
    finalizeTotalsSheet(totalSheet, srcSheets.length);

    // clean up the template sheet
    destDoc.deleteSheet(templateSheet);
}

function createPickingListV2() {
    console.log("createPickingListV2()");

    var srcSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var destDoc = SpreadsheetApp.create("Picking_list_" + Date());

    createTemplate(destDoc, srcSheets[0]);
    templateSheet = destDoc.getSheets()[1];

    // prepare totals sheet
    totalSheet = destDoc.getSheets()[0];
    prepareTotalSheet(totalSheet, srcSheets[0]);

    for(var i = 0; i < srcSheets.length; i++) {
        console.log("Behandlar ark: " + srcSheets[i].getName());
        updateTotalsSheet(totalSheet, srcSheets[i]);

        // no group pages, only the shared totals page == picking list
    }

    // auto-resize totalssheet column
    finalizeTotalsSheet(totalSheet, srcSheets.length);

    // clean up the template sheet
    destDoc.deleteSheet(templateSheet);
}

function removeExtraRowsAndColumns(sheet) {
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();

    // Remove rows after the last data row
    sheet.deleteRows(lastRow + 1, sheet.getMaxRows() - lastRow);

    // Remove columns after the last data column
    sheet.deleteColumns(lastColumn + 1, sheet.getMaxColumns() - lastColumn);
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
    removeExtraRowsAndColumns(totalSheet);
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

function updateTotalsSheet(totalSheet, srcSheet) {
    // make space for the "total-totals" at column 2
    total_cols = totalSheet.getLastColumn();
    col = total_cols + 2;
    if (total_cols > 2) {
        col = total_cols + 1;
    }

    totalSheet
      .getRange( 1, col, srcSheet.getMaxRows(), 1 )
      .setValues( srcSheet.getRange("B:B").getDisplayValues() )
      .setFontSize(15).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

    var column_name = srcSheet.getName().replace(/grupp /g,"").replace(/Grupp /g,"").replace(/\"/g,"");
    totalSheet
      .getRange( 2, col, 1, 1)
      .setValue( column_name )
      .setFontSize(15)
      .setFontWeight("bold")

    totalSheet
        .autoResizeColumn(col);
}

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

function createGroupTotalsPage(destDoc, srcSheet) {
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

function createMemberPages(destDoc, srcSheet, templateSheet) {
    var srcRowCount = srcSheet.getMaxRows();
    
    var veggies = srcSheet
        .getRange("A:A")
        .getValues();
    templateSheet
        .getRange(2, 1, srcSheet.getMaxRows(), 1)
        .setValues(veggies)
        .setFontSize(15).setFontWeight("bold")
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    templateSheet.autoResizeRows(1, templateSheet.getMaxRows());

    var numberOfSheets = Math.ceil((srcSheet.getMaxColumns()-2)/2/2);
    for (var i = 0; i < numberOfSheets; i++) {
        var currentSheetName = srcSheet.getName() + " " + (i+1) + " of " + numberOfSheets;
        Logger.log(currentSheetName);
      
        var destSheet = templateSheet
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
