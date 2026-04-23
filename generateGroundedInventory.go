package excelutils

import (
	"context"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"slices"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

type AuctionInfo struct {
	AuctionName string
}

type AuctionProvider interface {
	CheckAuctionExists(ctx context.Context, city, state string) (bool, error)
	GetAuction(ctx context.Context, city, state string) (AuctionInfo, error)
}

func GenerateGrounded(csvRecords [][]string, savePath string, ctx context.Context, cfg AuctionProvider) error {

	//create absolute filepaths
	home, err := os.UserHomeDir()
	if err != nil {
		log.Fatal("$HOME environment variable error: ", err)
	}

	templateFilePath := filepath.Join(home, "workspace", "github.com", "jamjallred", "sf_server_utils", "assets", "grounded_template.xlsx")
	templateCopyPath := filepath.Join(home, "workspace", "github.com", "jamjallred", "sf_server_utils", "assets", "grounded_template_copy.xlsx")

	if err = copyTemplate(templateFilePath, templateCopyPath); err != nil {
		fmt.Println("error copying template: ", err)
		return err
	}

	dst, err := excelize.OpenFile(templateCopyPath)
	if err != nil {
		fmt.Println("error opening new excel file: ", err)
		return err
	}
	defer dst.Close()

	fmt.Println("made it here") // TESTING LINE ``````````````````````````````````

	if err := generateGroundedSheet(dst, csvRecords, savePath, ctx, cfg); err != nil {
		fmt.Println("error generating sheet: ", err)
		return err
	}

	return nil

}

func generateGroundedSheet(dst *excelize.File, csv [][]string, savePath string, ctx context.Context, cfg AuctionProvider) error {

	// newHeaders := []string{"State", "City", "Year", "Make", "Model", "Trim", "Drivetrain", "VIN", "Mileage", "Color", "But It Now Price", "Vehicle Link"}
	//dst.SetCellValue(sheetName, "G1", "Drive")                    // rename "Body Type" to "Drive"
	colIndices := []int{11, 10, 1, 2, 3, 4, 13, 0, 5, 8, 7, 12} // column order from source sheet

	rowNum := 0 // scope to outside loop block to use for style formatting
	sheetName := dst.GetSheetName(0)
	fmt.Println("Generating grounded sheet...")

	for i, row := range csv[1:] { // skipping header row

		rowData := make([]any, 0, len(colIndices))
		state := row[11]
		city := row[10]

		auction_exists, err := cfg.CheckAuctionExists(ctx, city, state)
		if err != nil {
			log.Printf("Error accessing database table city_auction_map: %v", err)
			auction_exists = false
		}

		if auction_exists {
			auction, err := cfg.GetAuction(ctx, city, state)
			if err != nil {
				log.Printf("Error accessing database table city_auction_map: %v", err)
			}
			row[10] = fmt.Sprintf("%s (%s)", city, auction.AuctionName)
		}

		for _, colIdx := range colIndices[0:] {
			rowData = append(rowData, row[colIdx])
		}

		rowNum = i + 2 // A2 is first row after header
		rowRef := fmt.Sprintf("A%v", rowNum)
		if err := dst.SetSheetRow(sheetName, rowRef, &rowData); err != nil {
			fmt.Println("error setting row:", err)
			return err
		}

	}

	//sort row data by State, City, Yr, Make, then Model
	fmt.Println("Sorting sheet...")
	csv = nil // free up memory
	dstRows, err := dst.GetRows(sheetName)
	if err != nil {
		fmt.Println(err)
		return err
	}

	slices.SortFunc(dstRows[1:], func(a, b []string) int {
		if n := strings.Compare(a[0], b[0]); n != 0 { // compare state
			return n
		}
		if n := strings.Compare(a[1], b[1]); n != 0 { // compare city
			return n
		}
		if n := strings.Compare(a[2], b[2]); n != 0 { // compare year
			return n
		}
		if n := strings.Compare(a[3], b[3]); n != 0 { // compare make
			return n
		}
		return strings.Compare(a[4], b[4]) // compare model
	})

	// write sorted data into sheet
	for i, row := range dstRows[1:] { // skip header row
		rowNum = i + 2
		rowRef := fmt.Sprintf("A%v", rowNum)
		year, _ := strconv.Atoi(strings.TrimSpace(row[2]))
		miles, _ := strconv.Atoi(strings.TrimSpace(row[8]))
		price, _ := strconv.Atoi(strings.TrimSpace(strings.ReplaceAll(strings.ReplaceAll(row[10], "$", ""), ",", "")))
		rowData := make([]any, 0, len(row))
		rowData = append(rowData, row[0], row[1], year, row[3], row[4], row[5], row[6], row[7], miles, row[9], price, row[11])

		if err = dst.SetSheetRow(sheetName, rowRef, &rowData); err != nil {
			fmt.Println("error setting sorted row:", err)
			return err
		}
	}

	// convert Yr, Miles, Price, MSRP to Number style, skip header row
	fmt.Println("Converting Number Cells...")
	decimalPlaces := 0
	numID, err := dst.NewStyle(&excelize.Style{
		NumFmt: 1, // number format with no decimals
		Font: &excelize.Font{
			Family: "Calibri",
			Size:   10,
		},
	})
	if err != nil {
		fmt.Println("error creating style:", err)
		return err
	}
	currencyID, err := dst.NewStyle(&excelize.Style{
		NumFmt:        177, // US Dollar format
		DecimalPlaces: &decimalPlaces,
		Font: &excelize.Font{
			Family: "Calibri",
			Size:   10,
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#92D050"},
			Pattern: 1,
		},
	})
	if err != nil {
		fmt.Println("error creating style:", err)
		return err
	}
	dst.SetCellStyle(sheetName, "C2", fmt.Sprintf("C%v", rowNum), numID)
	dst.SetCellStyle(sheetName, "I2", fmt.Sprintf("I%v", rowNum), numID)
	dst.SetCellStyle(sheetName, "K2", fmt.Sprintf("K%v", rowNum), currencyID)

	//save file
	fmt.Println("Saving file...")
	if err = dst.SaveAs(savePath); err != nil {
		log.Fatalf("error saving file: %v", err)
	}

	fmt.Println("Sheet generated successfully.")

	return nil

}
