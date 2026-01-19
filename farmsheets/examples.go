package main

import (
	"context"
	"encoding/json"
	"fmt"
	"log"
	"net/http"
	"os"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

func getClient(ctx context.Context, config *oauth2.Config) *http.Client {
	tokFile := "token.json"
	tok, err := tokenFromFile(tokFile)
	if err != nil {
		tok = getTokenFromWeb(config)
		saveToken(tokFile, tok)
	}
	return config.Client(ctx, tok)
}

func getTokenFromWeb(config *oauth2.Config) *oauth2.Token {
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	fmt.Println("Visit this URL and paste the code:", authURL)

	var code string
	fmt.Scan(&code)

	tok, err := config.Exchange(context.Background(), code)
	if err != nil {
		log.Fatal(err)
	}
	return tok
}

func tokenFromFile(file string) (*oauth2.Token, error) {
	f, err := os.Open(file)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	tok := &oauth2.Token{}
	err = json.NewDecoder(f).Decode(tok)
	return tok, err
}

func saveToken(path string, tok *oauth2.Token) {
	fmt.Println("Saving token to", path)
	f, err := os.Create(path)
	if err != nil {
		log.Fatal(err)
	}
	defer f.Close()
	json.NewEncoder(f).Encode(tok)
}

func DriveExample() {
	ctx := context.Background()
	/*
		scopes := []string{
			drive.DriveScope,
			sheets.SpreadsheetsScope,
			"https://spreadsheets.google.com/feeds",
			//"https://www.googleapis.com/auth/spreadsheets",
			//"https://www.googleapis.com/auth/drive.file",
			//"https://www.googleapis.com/auth/drive",
		}
	*/

	//creds := "./client_secret_500250491618-5a6i6ul1iipe47fefcif0gcb6aq6mhqg.apps.googleusercontent.com.json"
	//creds := "./farm-sheets-94e924dcabb8.json"

	creds, err := os.ReadFile("credentials_oauth.json")
	if err != nil {
		log.Fatal(err)
	}
	/*
		svc, err := drive.NewService(ctx,
			option.WithCredentialsFile(creds),
			option.WithScopes(scopes...))
		if err != nil {
			log.Fatalf("Unable to create Drive client: %v", err)
		}
	*/

	config, _ := google.ConfigFromJSON(
		creds,
		drive.DriveScope,
		sheets.SpreadsheetsScope,
	)
	client := getClient(ctx, config)
	svc, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		log.Fatal(err)
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

	/*
		file := &drive.File{
			Name:     "service_sheet_create",
			MimeType: "application/vnd.google-apps.spreadsheet",
			Parents:  []string{"1NT32_gi8ix5AeGHFHSoHHeR6vKQGbewr"},
		}

		created, err := svc.Files.Create(file).SupportsAllDrives(true).Do()
		if err != nil {
			fmt.Println(err)
		} else {
			fmt.Println("Created: ", created.Name)
		}
	*/
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

func PrintValuesDims_ex(titles []string, values []*sheets.ValueRange) {
	for idx, sheet_vals := range values {
		fmt.Println()
		fmt.Println()
		fmt.Println("Sheet title: ", titles[idx])
		fmt.Println("Sheet values: ")
		fmt.Println("Sheet rows: ", len(sheet_vals.Values))

		// iterate rows
		for idx, row := range sheet_vals.Values {
			fmt.Println(idx, len(row))
		}
	}
}

func DownloadSpreadSheetDataSequentially_ex(spredsheet_id string, svc *sheets.Service) (titles []string, values []*sheets.ValueRange) {
	fmt.Println("Getting sheet data ...")

	ss, err := svc.Spreadsheets.Get(spredsheet_id).Do()
	if err != nil {
		log.Fatalf("unable to read data: %v", err)
	}

	titles = make([]string, 0, 100)
	values = make([](*sheets.ValueRange), 0, 100)

	for _, sheet := range ss.Sheets {
		fmt.Printf("%s (%s)\n", sheet.Properties.Title, sheet.Properties.SheetType)

		vals, err := svc.Spreadsheets.Values.Get(ss.SpreadsheetId, sheet.Properties.Title).Do()
		values = append(values, vals)
		titles = append(titles, sheet.Properties.Title)

		if err != nil {
			fmt.Println("could not get sheet values")
			os.Exit(1)
		}
	}

	fmt.Println("Download complete")
	return
}
