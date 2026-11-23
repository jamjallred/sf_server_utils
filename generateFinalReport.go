package excelutils

import (
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Fallback column indexes, used when a header can't be found.
const (
	dibsClaimedCol = 20 // column U on the old dibs sheet: "false" when nobody claimed the car
	dibsVINCol     = 7  // column H
	resultsVINCol  = 0
	resultsRepCol  = 22 // column W, "Avis Rep Name"
)

// resultsRepHeaders are the headers the auction results use for the rep's name.
var resultsRepHeaders = []string{"Avis Rep Name", "Rep Name"}

// nicknames maps a formal first name to the one the team uses, so "Michael
// Allred" in the auction results matches "Mike Allred". Keys and values are
// lower case.
var nicknames = map[string]string{
	"michael": "mike",
}

// GenerateFinalReport edits the dibs workbook in place: it drops unclaimed
// rows (old dibs sheets only) and adds a FINAL column with the full name of the
// rep who closed each car, or "MISSED" when someone outside repNames did.
// csvRecords are the auction results, header row included.
//
// Names are matched loosely: case, spacing and nicknames (Michael = Mike) are
// ignored, so "MICHAEL  ALLRED" in the results counts as "Mike Allred".
func GenerateFinalReport(csvRecords [][]string, dibs *excelize.File, repNames []string) error {
	sheet := dibs.GetSheetName(0)

	rows, err := dibs.GetRows(sheet)
	if err != nil {
		return fmt.Errorf("reading dibs sheet: %w", err)
	}
	if len(rows) == 0 {
		return fmt.Errorf("the dibs sheet is empty")
	}
	header := rows[0]

	team := map[string]string{} // nameKey -> full name as configured
	firstNames := map[string]bool{}
	for _, n := range repNames {
		if k := nameKey(n); k != "" {
			team[k] = strings.Join(strings.Fields(n), " ")
			firstNames[strings.Fields(k)[0]] = true
		}
	}

	vinCol := headerIndex(header, "vin")
	if vinCol < 0 {
		vinCol = dibsVINCol
	}
	// The All Dibs export has a column per person (Mike, Stan, …) and only
	// lists claimed cars. The old dibs sheet has a Claimed column (U) instead.
	claimedCol := headerIndex(header, "claimed")
	if claimedCol < 0 && !hasNameColumns(header, firstNames) {
		claimedCol = dibsClaimedCol
	}

	if claimedCol >= 0 {
		// Remove unclaimed rows bottom-up so earlier row numbers stay valid.
		for i := len(rows) - 1; i >= 1; i-- {
			if strings.EqualFold(strings.TrimSpace(cell(rows[i], claimedCol)), "false") {
				if err := dibs.RemoveRow(sheet, i+1); err != nil {
					return fmt.Errorf("removing row %d: %w", i+1, err)
				}
			}
		}
	}

	// FINAL goes right after Claimed (column V on the old sheet), otherwise
	// after the last column.
	finalIdx := len(header)
	if claimedCol >= 0 {
		finalIdx = claimedCol + 1
	}
	finalCol, err := excelize.ColumnNumberToName(finalIdx + 1)
	if err != nil {
		return err
	}
	if finalIdx < len(header) {
		if err := dibs.InsertCols(sheet, finalCol, 1); err != nil {
			return fmt.Errorf("inserting FINAL column: %w", err)
		}
	}
	if styleID, err := dibs.GetCellStyle(sheet, "A1"); err == nil {
		_ = dibs.SetCellStyle(sheet, finalCol+"1", finalCol+"1", styleID)
	}
	if err := dibs.SetCellValue(sheet, finalCol+"1", "FINAL"); err != nil {
		return err
	}

	// VIN -> rep from the results sheet. Later rows win, as before.
	repCol := repColumn(csvRecords, team)
	resVINCol := resultsVINCol
	if len(csvRecords) > 0 {
		if c := headerIndex(csvRecords[0], "vin"); c >= 0 {
			resVINCol = c
		}
	}
	repByVIN := map[string]string{}
	for i, rec := range csvRecords {
		if i == 0 {
			continue // header
		}
		if vin := vinKey(cell(rec, resVINCol)); vin != "" {
			repByVIN[vin] = cell(rec, repCol)
		}
	}

	rows, err = dibs.GetRows(sheet)
	if err != nil {
		return fmt.Errorf("re-reading dibs sheet: %w", err)
	}
	for i := 1; i < len(rows); i++ {
		rep, ok := repByVIN[vinKey(cell(rows[i], vinCol))]
		if !ok {
			continue
		}
		value := "MISSED"
		if full, ok := team[nameKey(rep)]; ok {
			value = full
		}
		if err := dibs.SetCellValue(sheet, fmt.Sprintf("%s%d", finalCol, i+1), value); err != nil {
			return err
		}
	}
	return nil
}

// nameKey normalises a person's name for matching: lower case, single spaces,
// invisible characters dropped, and the first name mapped through nicknames.
func nameKey(s string) string {
	s = strings.NewReplacer("\u200b", "", "\ufeff", "").Replace(strings.ToLower(s))
	f := strings.Fields(s) // also splits on non-breaking spaces
	if len(f) == 0 {
		return ""
	}
	if n, ok := nicknames[f[0]]; ok {
		f[0] = n
	}
	return strings.Join(f, " ")
}

func vinKey(s string) string {
	return strings.ToUpper(strings.TrimSpace(strings.TrimPrefix(s, "\ufeff")))
}

// headerIndex finds a column by its header text, ignoring case and spacing.
func headerIndex(header []string, name string) int {
	for i, h := range header {
		if strings.EqualFold(strings.TrimSpace(strings.TrimPrefix(h, "\ufeff")), name) {
			return i
		}
	}
	return -1
}

// hasNameColumns reports whether any header is a team member's first name.
func hasNameColumns(header []string, firstNames map[string]bool) bool {
	for _, h := range header {
		if firstNames[nameKey(h)] {
			return true
		}
	}
	return false
}

// repColumn returns the results column holding the rep's name: the "Avis Rep
// Name" header when there is one, otherwise the column with the most team
// members' names, otherwise column W.
func repColumn(records [][]string, team map[string]string) int {
	if len(records) > 0 {
		for _, h := range resultsRepHeaders {
			if c := headerIndex(records[0], h); c >= 0 {
				return c
			}
		}
	}
	var counts []int
	for i, rec := range records {
		if i == 0 {
			continue
		}
		for c, v := range rec {
			if _, ok := team[nameKey(v)]; ok {
				for len(counts) <= c {
					counts = append(counts, 0)
				}
				counts[c]++
			}
		}
	}
	at := func(c int) int {
		if c < len(counts) {
			return counts[c]
		}
		return 0
	}
	best := resultsRepCol
	for c := range counts {
		if at(c) > at(best) {
			best = c
		}
	}
	return best
}
