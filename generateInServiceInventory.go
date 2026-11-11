package excelutils

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Source columns for the nationwide sheet, after the State and City that come
// from the airport code (column F): Yr, Make, Model, Trim, Drive, VIN, Color,
// Miles, Price, MSRP.
const (
	inServiceCodeCol = 5
	inServiceVINCol  = 8
	inServiceSkip    = 3 // date, blank line, header
)

var inServiceColumns = []int{9, 10, 11, 12, 18, 8, 17, 13, 20, 19}

// InServiceResult is the generated nationwide sheet plus bookkeeping.
type InServiceResult struct {
	File         *excelize.File
	Rows         int // cars written
	Duplicates   int // repeated VINs skipped
	UnknownCodes int // distinct airport codes not in the map (rows skipped)
}

// GenerateInService builds the nationwide sheet from an in-service inventory
// export. The caller must Close result.File.
func GenerateInService(src *excelize.File) (*InServiceResult, error) {
	codes, err := LoadAirportMap()
	if err != nil {
		return nil, err
	}

	srcRows, err := src.GetRows(src.GetSheetName(0))
	if err != nil {
		return nil, fmt.Errorf("reading upload: %w", err)
	}
	if len(srcRows) <= inServiceSkip {
		return nil, fmt.Errorf("the sheet has no data rows (expected a date line, a blank line and a header first)")
	}

	dst, err := excelize.OpenFile(assetPath("nationwide_template.xlsx"))
	if err != nil {
		return nil, fmt.Errorf("opening nationwide template: %w", err)
	}

	res := &InServiceResult{File: dst}
	seenVIN := map[string]bool{}
	seenCode := map[string]bool{}
	var newCodes [][]string
	var rows [][]string

	for _, in := range srcRows[inServiceSkip:] {
		vin := strings.TrimSpace(cell(in, inServiceVINCol))
		if vin != "" && seenVIN[vin] {
			res.Duplicates++
			continue
		}
		seenVIN[vin] = true

		code := cell(in, inServiceCodeCol)
		place, ok := codes[code]
		if !ok {
			if !seenCode[code] {
				seenCode[code] = true
				newCodes = append(newCodes, []string{code, cell(in, 4), cell(in, 3), cell(in, 2)})
			}
			continue
		}

		row := append([]string{place.State, place.City}, pick(in, inServiceColumns)...)
		if len(strings.TrimSpace(row[11])) <= 6 { // MSRP must look like a real price
			row[11] = "n/a"
		}
		rows = append(rows, row)
	}
	res.UnknownCodes = len(newCodes)

	if err := appendNewCodes(newCodes); err != nil {
		dst.Close()
		return nil, fmt.Errorf("saving unknown airport codes: %w", err)
	}

	if err := writeInServiceSheet(dst, rows); err != nil {
		dst.Close()
		return nil, err
	}
	res.Rows = len(rows)
	return res, nil
}

func writeInServiceSheet(dst *excelize.File, rows [][]string) error {
	sheet := dst.GetSheetName(0)
	if err := dst.SetCellValue(sheet, "G1", "Drive"); err != nil { // template says "Body Type"
		return err
	}

	sortInventory(rows)
	for i, r := range rows {
		values := []any{
			r[0], r[1], numberOrText(r[2]), r[3], r[4], r[5], r[6], r[7], r[8],
			numberOrText(r[9]), numberOrText(r[10]), numberOrText(r[11]),
		}
		if err := dst.SetSheetRow(sheet, fmt.Sprintf("A%d", i+2), &values); err != nil {
			return fmt.Errorf("writing row %d: %w", i+2, err)
		}
	}

	styles, err := newStyles(dst, "")
	if err != nil {
		return fmt.Errorf("creating styles: %w", err)
	}
	last := len(rows) + 1
	if err := styleColumns(dst, sheet, "C", "C", last, styles.number); err != nil {
		return err
	}
	if err := styleColumns(dst, sheet, "J", "J", last, styles.number); err != nil {
		return err
	}
	return styleColumns(dst, sheet, "K", "L", last, styles.currency)
}

// Generate reads an in-service export from xlsxPath and saves the nationwide
// sheet to savePath. Kept for command-line callers.
func Generate(xlsxPath, savePath string) error {
	src, err := excelize.OpenFile(xlsxPath)
	if err != nil {
		return fmt.Errorf("opening %s: %w", xlsxPath, err)
	}
	defer src.Close()

	res, err := GenerateInService(src)
	if err != nil {
		return err
	}
	defer res.File.Close()
	return res.File.SaveAs(savePath)
}
