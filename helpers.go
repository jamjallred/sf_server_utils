package excelutils

import (
	"os"
	"path/filepath"
	"slices"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

// AssetsDir returns the directory holding the Excel templates and the airport
// code list. Set SF_ASSETS_DIR to override it; otherwise it falls back to the
// assets folder inside this repo's usual checkout location.
func AssetsDir() string {
	if dir := os.Getenv("SF_ASSETS_DIR"); dir != "" {
		return dir
	}
	home, err := os.UserHomeDir()
	if err != nil {
		return "assets"
	}
	return filepath.Join(home, "workspace", "github.com", "jamjallred", "sf_server_utils", "assets")
}

func assetPath(name string) string {
	return filepath.Join(AssetsDir(), name)
}

// cell returns row[i], or "" when the row is too short. Spreadsheet rows
// returned by excelize drop trailing empty cells, so indexing directly panics.
func cell(row []string, i int) string {
	if i < 0 || i >= len(row) {
		return ""
	}
	return row[i]
}

// pick builds a new row from the given source column indexes.
func pick(row []string, cols []int) []string {
	out := make([]string, len(cols))
	for i, c := range cols {
		out[i] = cell(row, c)
	}
	return out
}

// ParseMoney turns values like "$12,345" or " 9800 " into an int.
func ParseMoney(s string) (int, bool) {
	s = strings.NewReplacer("$", "", ",", "").Replace(strings.TrimSpace(s))
	if s == "" {
		return 0, false
	}
	if n, err := strconv.Atoi(s); err == nil {
		return n, true
	}
	if f, err := strconv.ParseFloat(s, 64); err == nil {
		return int(f), true
	}
	return 0, false
}

// numberOrText returns the parsed number when s is numeric, otherwise the
// original text, so values like "n/a" survive instead of becoming 0.
func numberOrText(s string) any {
	if n, ok := ParseMoney(s); ok {
		return n
	}
	return s
}

// sortInventory sorts rows by state, city, year, make, then model, which are
// the first five columns of every generated sheet.
func sortInventory(rows [][]string) {
	slices.SortStableFunc(rows, func(a, b []string) int {
		for i := 0; i < 5; i++ {
			if n := strings.Compare(cell(a, i), cell(b, i)); n != 0 {
				return n
			}
		}
		return 0
	})
}

type sheetStyles struct {
	number   int
	currency int
}

// newStyles registers the number and US dollar styles used by the reports.
// fill is an optional background colour for currency cells (e.g. "#92D050").
func newStyles(f *excelize.File, fill string) (sheetStyles, error) {
	font := &excelize.Font{Family: "Calibri", Size: 10}
	number, err := f.NewStyle(&excelize.Style{NumFmt: 1, Font: font})
	if err != nil {
		return sheetStyles{}, err
	}
	decimals := 0
	currencyStyle := &excelize.Style{NumFmt: 177, DecimalPlaces: &decimals, Font: font}
	if fill != "" {
		currencyStyle.Fill = excelize.Fill{Type: "pattern", Color: []string{fill}, Pattern: 1}
	}
	currency, err := f.NewStyle(currencyStyle)
	if err != nil {
		return sheetStyles{}, err
	}
	return sheetStyles{number: number, currency: currency}, nil
}

// styleColumns applies a style to rows 2..lastRow of the columns from..to.
func styleColumns(f *excelize.File, sheet, from, to string, lastRow, style int) error {
	if lastRow < 2 {
		return nil
	}
	return f.SetCellStyle(sheet, from+"2", to+strconv.Itoa(lastRow), style)
}
