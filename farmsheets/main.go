package main

import (
	"context"
	"encoding/json"
	"fmt"
	"log"
	"os"
	"strconv"
	"strings"
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

func LoadValuesCache() (groups []string, values []*sheets.ValueRange) {
	f, err := os.Open("titles.json")
	if err != nil {
		return
	}
	err = json.NewDecoder(f).Decode(&groups)
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

func PrintAllData(groups []string, values []*sheets.ValueRange) {
	for idx_sheet, sheet := range values {

		sheet_vals := sheet
		sheet_title := groups[idx_sheet]

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

func GetCategories(values []*sheets.ValueRange, from_idx int) (categories []string) {
	category_sheet := values[0]

	for idx_row, row := range category_sheet.Values {
		fmt.Println(idx_row, len(row))

		if idx_row >= from_idx && len(row) > 0 {
			strval := fmt.Sprint(row[0])
			categories = append(categories, strval)
		}
	}
	return categories
}

func CreatePickingListData(categories []string, groups []string, values []*sheets.ValueRange) [][]string {
	row_offset := 2
	col_offsset := 2
	row_cnt := len(categories) + row_offset
	col_cnt := len(groups) + col_offsset

	fmt.Println("Picking List row_cnt: ", row_cnt)

	dest := make([][]string, row_cnt)
	for i := range dest {
		row := make([]string, col_cnt)
		dest[i] = row

		// Title row / groups
		switch {
		case i == 0:
			{
				for j := range row {
					switch {
					case j == 1:
						{
							row[j] = "Total"
						}
					case j > 1:
						{
							row[j] = groups[j-col_offsset]
						}
					}
				}
			}
		case i >= row_offset:
			{
				var sum int
				sum = 0

				for j := range row {
					switch {
					case j == 0: // categories column
						{
							if i >= col_offsset {
								row[j] = categories[i-row_offset]
							}
						}
					case j == 1:
						{
							// Total column - nothing right now
						}
					case j >= 2:
						{
							sheet := values[j-col_offsset]
							sheet_row := sheet.Values[i-row_offset+4]
							if len(sheet_row) > 1 {

								sheet_cell := sheet_row[1]
								cell := sheet_cell

								row[j] = fmt.Sprint(cell)
								intval, err := strconv.Atoi(row[j])
								if err != nil {
									// empty cell or someone put a string
									fmt.Println("Error: Could not parse value", i, j)
								} else {
									sum += intval
								}
							}
						}
					}
				}

				// record Total
				sum_str := strconv.Itoa(sum)
				row[1] = sum_str
			}
		}
	}
	return dest
}

func main() {
	do_download := false
	do_create := true

	if len(os.Args) > 1 && (os.Args[1] == "--download" || os.Args[1] == "-d") {
		do_download = true
	} else if len(os.Args) > 1 && (os.Args[1] == "--create") {
		do_create = true
	} else if len(os.Args) > 1 && (os.Args[1] == "--help" || os.Args[1] == "-h") {
		fmt.Println("Use --download to cache sheet data")
		return
	}

	if do_create {
		DriveExample()

		/*
			ctx := context.Background()
			creds := "./farm-sheets-94e924dcabb8.json"
			svc, err := sheets.NewService(ctx, option.WithCredentialsFile(creds))
			if err != nil {
				log.Fatalf("unable to create Sheets client: %v", err)
			}

			spreadsheet := &sheets.Spreadsheet{
				Properties: &sheets.SpreadsheetProperties{
					Title:   "New Spreadsheet",
				},
			}

			resp, err := svc.Spreadsheets.Create(spreadsheet).Do()
			if err != nil {
				log.Fatal(err)
			}

			fmt.Println("Spreadsheet ID:", resp.SpreadsheetId)
		*/
	} else if do_download {
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

		groups, values := LoadValuesCache()

		// trim the group names
		for idx := range groups {
			groupname := groups[idx]
			groupname = strings.TrimPrefix(groupname, "grupp")
			groupname = strings.Trim(groupname, " \"")
			groups[idx] = groupname
		}

		// test functions
		PrintAllData(groups, values)
		first_category_row_idx := 4
		cats := GetCategories(values, first_category_row_idx)

		for _, cat := range cats {
			fmt.Println(cat)
		}

		// actual work functions
		pickings := CreatePickingListData(cats, groups, values)

		fmt.Println()
		fmt.Println()
		fmt.Println("Picking List:")
		fmt.Println()
		for _, row := range pickings {
			fmt.Print("[  ")
			row_cnt := len(row)
			for idx, cell := range row {

				fmt.Print(cell)

				if idx < row_cnt-1 {
					fmt.Print(", ")
				}
			}
			fmt.Print("  ]\n")
		}
	}
}
