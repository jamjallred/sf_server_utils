package excelutils

import (
	"fmt"
	"log"
	"os"
	"slices"
	"strings"

	"github.com/xuri/excelize/v2"
)

func GenerateFinalReport(csvRecords [][]string, xlsxRecords *excelize.File, savePath string) error {

	fmt.Println("made it here") // TESTING LINE ``````````````````````````````````

	if err := generateFinalReportSheet(xlsxRecords, csvRecords, savePath); err != nil {
		fmt.Println("error generating sheet: ", err)
		return err
	}

	return nil

}

func generateFinalReportSheet(dst *excelize.File, csv [][]string, savePath string) error {

	sheetName := dst.GetSheetName(0)
	fmt.Println("Generating final results sheet...")

	dstRows, err := dst.GetRows(sheetName)
	if err != nil {
		fmt.Printf("error retrieving rows: %v\n", err)
		return err
	}

	// delete rows where the final column is false (unclaimed)
	for i, row := range dstRows {
		fmt.Println("removing row") // DEBUG LINE, REMOVE LATER
		if strings.EqualFold(row[20], "false") {
			dst.RemoveRow(sheetName, i+1)
		}
	}

	dst.InsertCols(sheetName, "V", 1) // insert extra column for final
	// set header cell style for new column
	styleID, err := dst.GetCellStyle(sheetName, "A1")
	if err != nil {
		fmt.Println(err)
		return err
	}
	err = dst.SetCellStyle(sheetName, "V1", "V1", styleID)
	if err != nil {
		fmt.Println(err)
		return err
	}
	err = dst.SetCellValue(sheetName, "V1", "FINAL")
	if err != nil {
		fmt.Println(err)
		return err
	}
	// end new column styles

	dstRows, err = dst.GetRows(sheetName) // repopulate dstRows from culled sheet w/ new column
	if err != nil {
		fmt.Printf("error retrieving rows (2): %v\n", err)
		return err
	}

	repNames := []string{os.Getenv("EMPLOYEE_1"), os.Getenv("EMPLOYEE_2"), os.Getenv("EMPLOYEE_3"), os.Getenv("EMPLOYEE_4")}

	for i, row := range dstRows {
		fmt.Println("checking row")
		rowNum := i + 1
		for _, csvRow := range csv {
			if row[7] != csvRow[0] {
				continue
			}
			repName := csvRow[15]
			if slices.Contains(repNames, csvRow[15]) {
				dst.SetCellValue(sheetName, fmt.Sprintf("V%v", rowNum), repName)
			} else {
				dst.SetCellValue(sheetName, fmt.Sprintf("V%v", rowNum), "MISSED")
			}
		}
	}

	//save file
	fmt.Println("Saving file...")
	if err := dst.SaveAs(savePath); err != nil {
		log.Fatalf("error saving file: %v", err)
	}

	fmt.Println("Sheet generated successfully.")

	return nil

}
