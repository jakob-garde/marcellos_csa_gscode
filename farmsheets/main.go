package main

import (
	"context"
	"encoding/json"
	"fmt"
	"log"
	"os"
	"sync"

	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

func DownloadSpreadSheetData(spredsheet_id string, svc *sheets.Service) (titles []string, values []*sheets.ValueRange) {
	fmt.Println("Getting sheet data ...")

	spread, err := svc.Spreadsheets.Get(spredsheet_id).Do()
	if err != nil {
		log.Fatalf("Unable to read data: %v", err)
	}

	titles = make([]string, len(spread.Sheets))
	values = make([](*sheets.ValueRange), len(spread.Sheets))

	var wg sync.WaitGroup
	for idx, sheet := range spread.Sheets {
		wg.Go(func() {
			vals, err := svc.Spreadsheets.Values.Get(spread.SpreadsheetId, sheet.Properties.Title).Do()
			if err != nil {
				fmt.Println("could not get sheet values")
				os.Exit(1)
			}

			values[idx] = vals
			titles[idx] = sheet.Properties.Title
		})
	}
	wg.Wait()
	fmt.Println("Download complete")

	for idx := range spread.Sheets {
		fmt.Println(idx, titles[idx])
	}
	return
}

func SaveValuesCache(titles []string, values []*sheets.ValueRange) {
	ftitles, err := os.Create("titles.json")
	if err != nil {
		return
	}
	fvalues, err := os.Create("values.json")
	if err != nil {
		return
	}

	json.NewEncoder(ftitles).Encode(titles)
	json.NewEncoder(fvalues).Encode(values)

	defer ftitles.Close()
	defer fvalues.Close()
}

func LoadValuesCache() (titles []string, values []*sheets.ValueRange) {
	f, err := os.Open("titles.json")
	if err != nil {
		return
	}
	err = json.NewDecoder(f).Decode(&titles)
	if err != nil {
		fmt.Println("Could not load titles.json")
		os.Exit(1)
	}

	f, err = os.Open("values.json")
	if err != nil {
		return
	}
	err = json.NewDecoder(f).Decode(&values)
	if err != nil {
		fmt.Println("Could not load values.json")
		os.Exit(1)
	}

	return
}

func PrintAllData(titles []string, values []*sheets.ValueRange) {
	for idx_sheet, sheet := range values {

		sheet_vals := sheet
		sheet_title := titles[idx_sheet]

		fmt.Println("Sheet title: ", sheet_title)

		for idx_row, row := range sheet_vals.Values {
			fmt.Println(idx_row, len(row))
			for idx_col, cell := range row {

				// cell is now idx_row, idx_col data

				fmt.Print(idx_col, ": ", cell, " ")
			}
		}
	}
}

func GetCategories(values []*sheets.ValueRange, from_idx int) (categories []interface{}) {
	category_sheet := values[0]

	for idx_row, row := range category_sheet.Values {
		fmt.Println(idx_row, len(row))

		if idx_row >= from_idx && len(row) > 0 {
			categories = append(categories, row[0])
		}
	}
	return categories
}

func main() {
	do_download := false

	if len(os.Args) > 1 && (os.Args[1] == "--download" || os.Args[1] == "-d") {
		do_download = true
	} else if len(os.Args) > 1 && (os.Args[1] == "--help" || os.Args[1] == "-h") {
		fmt.Println("Use --download to cache sheet data")
		return
	}

	if do_download {

		ctx := context.Background()
		creds := "./farm-sheets-94e924dcabb8.json"
		svc, err := sheets.NewService(ctx, option.WithCredentialsFile(creds))
		if err != nil {
			log.Fatalf("unable to create Sheets client: %v", err)
		}

		spredsheet_id := "1C9PIXa_Tm1eP0lHVF3073Nvo3NNYnerdqgqVLQmmMUw" // dev lista
		titles, values := DownloadSpreadSheetData(spredsheet_id, svc)

		SaveValuesCache(titles, values)

	} else {

		titles, values := LoadValuesCache()
		PrintAllData(titles, values)
		cats := GetCategories(values, 2)

		for _, cat := range cats {
			fmt.Println(cat)
		}
	}
}
