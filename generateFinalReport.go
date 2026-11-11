package excelutils

import (
	"fmt"
	"slices"
	"strings"

	"github.com/xuri/excelize/v2"
)

const (
	dibsClaimedCol = 20 // column U: "false" when nobody claimed the car
	dibsVINCol     = 7  // column H
	resultsVINCol  = 0
	resultsRepCol  = 15
)

// GenerateFinalReport edits the dibs workbook in place: it drops unclaimed
// rows and adds a FINAL column (V) with the rep who closed each car, or
// "MISSED" when someone outside repNames did. csvRecords are the auction
// results, header row included.
func GenerateFinalReport(csvRecords [][]string, dibs *excelize.File, repNames []string) error {
	sheet := dibs.GetSheetName(0)

	rows, err := dibs.GetRows(sheet)
	if err != nil {
		return fmt.Errorf("reading dibs sheet: %w", err)
	}

	// Remove unclaimed rows bottom-up so earlier row numbers stay valid.
	for i := len(rows) - 1; i >= 1; i-- {
		if strings.EqualFold(strings.TrimSpace(cell(rows[i], dibsClaimedCol)), "false") {
			if err := dibs.RemoveRow(sheet, i+1); err != nil {
				return fmt.Errorf("removing row %d: %w", i+1, err)
			}
		}
	}

	if err := dibs.InsertCols(sheet, "V", 1); err != nil {
		return fmt.Errorf("inserting FINAL column: %w", err)
	}
	if styleID, err := dibs.GetCellStyle(sheet, "A1"); err == nil {
		_ = dibs.SetCellStyle(sheet, "V1", "V1", styleID)
	}
	if err := dibs.SetCellValue(sheet, "V1", "FINAL"); err != nil {
		return err
	}

	// VIN -> rep from the results sheet. Later rows win, as before.
	repByVIN := map[string]string{}
	for i, rec := range csvRecords {
		if i == 0 {
			continue // header
		}
		if vin := strings.TrimSpace(cell(rec, resultsVINCol)); vin != "" {
			repByVIN[vin] = strings.TrimSpace(cell(rec, resultsRepCol))
		}
	}

	rows, err = dibs.GetRows(sheet)
	if err != nil {
		return fmt.Errorf("re-reading dibs sheet: %w", err)
	}
	for i := 1; i < len(rows); i++ {
		rep, ok := repByVIN[strings.TrimSpace(cell(rows[i], dibsVINCol))]
		if !ok {
			continue
		}
		value := "MISSED"
		if rep != "" && slices.Contains(repNames, rep) {
			value = rep
		}
		if err := dibs.SetCellValue(sheet, fmt.Sprintf("V%d", i+1), value); err != nil {
			return err
		}
	}
	return nil
}
