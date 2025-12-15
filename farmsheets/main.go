package main

import (
	"context"
	"fmt"
	"log"

	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

func drive_ex() {
	ctx := context.Background()
	scopes := []string{
		"https://spreadsheets.google.com/feeds",
		"https://www.googleapis.com/auth/spreadsheets",
		"https://www.googleapis.com/auth/drive.file",
		"https://www.googleapis.com/auth/drive",
	}
	creds := "./farm-sheets-94e924dcabb8.json"

	svc, err := drive.NewService(ctx, option.WithCredentialsFile(creds), option.WithScopes(scopes...))
	if err != nil {
		log.Fatalf("Unable to create Drive client: %v", err)
	}

	// files list object
	r, err := svc.Files.List().
		PageSize(10).
		Fields("nextPageToken, files(id, name)").
		Do()
	if err != nil {
		log.Fatalf("Unable to retrieve files: %v", err)
	}

	// print files
	items := r.Files
	for _, f := range items {
		fmt.Printf("%s (%s)\n", f.Name, f.Id)
	}
}

func write_sheet_ex(svc *sheets.Service, spredsheet_id string) {
	writeRange := "Sheet1!A1:D2"
	values := [][]interface{}{{"Stone", "Fire", "Coffee", "Smoke"}, {111, 222, 333, 444}}

	body := &sheets.ValueRange{
		Values: values,
	}

	_, err := svc.Spreadsheets.Values.Update(spredsheet_id, writeRange, body).
		ValueInputOption("RAW").
		Do()
	if err != nil {
		log.Fatalf("unable to update data: %v", err)
	}

	fmt.Println("Sheet updated successfully.")
}

func main() {

	ctx := context.Background()
	creds := "./farm-sheets-94e924dcabb8.json"

	svc, err := sheets.NewService(ctx, option.WithCredentialsFile(creds))
	if err != nil {
		log.Fatalf("unable to create Sheets client: %v", err)
	}

	//spreadsheetID := "15fK71g_KNd52QEZwrJ2i9MKsJSQrZ4azhDiHRrnkl0s" // test document
	spreadsheetID := "1C9PIXa_Tm1eP0lHVF3073Nvo3NNYnerdqgqVLQmmMUw" // dev lista
	readRange := "A:D"

	// retrieve values
	resp, err := svc.Spreadsheets.Values.Get(spreadsheetID, readRange).Do()
	if err != nil {
		log.Fatalf("unable to read data: %v", err)
	}
	// output rows
	for _, row := range resp.Values {
		fmt.Println(row)
	}

	// NOTE:
	//  we know there is a spreadsheetID and a separate sheetID for eaach sheet in the doc.
	//  There is always one "default" sheet though, which we are getting
	//sheet_array := svc.Spreadsheets.Sheets
	//fmt.Print((sheet_array))

	ss, err := svc.Spreadsheets.Get(spreadsheetID).Do()
	if err != nil {
		log.Fatalf("unable to read data: %v", err)
	}

	// iterate sheets
	for _, sheet := range ss.Sheets {
		fmt.Printf("%s (%s)\n", sheet.Properties.Title, sheet.Properties.SheetType)

		// here we get all of the valuse of the sheet
		// it takes another server call
		// THUS: I want to batch call the server for all the sheet values once and for all and
		// store them in a slice (maybe slc = []sheets.ValueRange)

		vals, err := svc.Spreadsheets.Values.Get(spreadsheetID, sheet.Properties.Title).Do()
		if err != nil {
			fmt.Println("could not get sheet values")
		} else {
			fmt.Println(sheet.Properties.Title)
		}

		// print length of the outer array (rows/lines)
		fmt.Println(len(vals.Values))

		// these are some of the row lengths
		fmt.Println(len(vals.Values[0]))
		fmt.Println(len(vals.Values[1]))
		fmt.Println(len(vals.Values[2]))
		fmt.Println(len(vals.Values[3]))
		fmt.Println(len(vals.Values[4]))
		fmt.Println(len(vals.Values[5]))
		fmt.Println(len(vals.Values[6]))
		fmt.Println(len(vals.Values[7]))

		// print a few valuse
		fmt.Println(vals.Values[7][0])
		fmt.Println(vals.Values[7][1])
	}
}
