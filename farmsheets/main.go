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
		log.Fatalf("unable to read data: %v", err)
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

	// print the results to check order is preserved
	for idx := range spread.Sheets {
		fmt.Println(idx, titles[idx])
	}

	fmt.Println("Download complete")
	return
}

func DownloadSpreadSheetData_seq(spredsheet_id string, svc *sheets.Service) (titles []string, values []*sheets.ValueRange) {
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
		fmt.Println("could not load titles")
		os.Exit(1)
	}

	f, err = os.Open("values.json")
	if err != nil {
		return
	}
	err = json.NewDecoder(f).Decode(&values)
	if err != nil {
		fmt.Println("could not load values")
		os.Exit(1)
	}

	return
}

func PrintValuesDims(titles []string, values []*sheets.ValueRange) {
	for idx, vals := range values {
		fmt.Println()
		fmt.Println()
		fmt.Println("Sheet title: ", titles[idx])
		fmt.Println("Sheet values: ")
		fmt.Println("Sheet rows: ", len(vals.Values))

		// iterate rows
		for idx, row := range vals.Values {
			fmt.Println(idx, len(row))
		}
	}
}

func main() {
	// TODO: get as a cmd line arg
	do_download := true

	if do_download {
		ctx := context.Background()
		creds := "./farm-sheets-94e924dcabb8.json"
		svc, err := sheets.NewService(ctx, option.WithCredentialsFile(creds))
		if err != nil {
			log.Fatalf("unable to create Sheets client: %v", err)
		}

		//spredsheet_id := "15fK71g_KNd52QEZwrJ2i9MKsJSQrZ4azhDiHRrnkl0s" // test document
		spredsheet_id := "1C9PIXa_Tm1eP0lHVF3073Nvo3NNYnerdqgqVLQmmMUw" // dev lista

		//titles, values := DownloadSpreadSheetData_seq(spredsheet_id, svc)
		titles, values := DownloadSpreadSheetData(spredsheet_id, svc)

		SaveValuesCache(titles, values)

	} else {
		titles, values := LoadValuesCache()
		PrintValuesDims(titles, values)
	}
}
