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

func iterate_sheets_ex(svc *sheets.Service, spredsheet_id string) {
	ss, err := svc.Spreadsheets.Get(spredsheet_id).Do()
	if err != nil {
		log.Fatalf("unable to read data: %v", err)
	}

	for _, sheet := range ss.Sheets {
		fmt.Printf("%s (%s)\n", sheet.Properties.Title, sheet.Properties.SheetType)

		vals, err := svc.Spreadsheets.Values.Get(ss.SpreadsheetId, sheet.Properties.Title).Do()
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
