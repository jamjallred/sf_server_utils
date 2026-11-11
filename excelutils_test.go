package excelutils

import (
	"fmt"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestParseMoney(t *testing.T) {
	cases := map[string]struct {
		want int
		ok   bool
	}{
		"$12,345":  {12345, true},
		" 9800 ":   {9800, true},
		"41250.00": {41250, true},
		"n/a":      {0, false},
		"":         {0, false},
	}
	for in, c := range cases {
		got, ok := ParseMoney(in)
		if got != c.want || ok != c.ok {
			t.Errorf("ParseMoney(%q) = %d, %v; want %d, %v", in, got, ok, c.want, c.ok)
		}
	}
}

func TestSortInventory(t *testing.T) {
	rows := [][]string{
		{"TX", "Austin", "2022", "Ford", "F-150"},
		{"CA", "San Jose", "2021", "Honda", "Civic"},
		{"TX", "Austin", "2020", "Toyota", "Camry"},
		{"CA", "Fresno"}, // short row must not panic
	}
	sortInventory(rows)
	want := []string{"Fresno", "San Jose", "Austin", "Austin"}
	for i, w := range want {
		if rows[i][1] != w {
			t.Fatalf("row %d city = %q, want %q (rows: %v)", i, rows[i][1], w, rows)
		}
	}
	if rows[2][2] != "2020" {
		t.Errorf("expected 2020 before 2022 within Austin, got %v", rows[2])
	}
}

// Rows 2, 3 and 4 are unclaimed in a row; the old top-down delete skipped
// every other one and removed claimed rows instead.
func TestGenerateFinalReportRemovesUnclaimedRows(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()
	sheet := f.GetSheetName(0)

	set := func(r int, vin, claimed string) {
		f.SetCellValue(sheet, fmt.Sprintf("H%d", r), vin)
		f.SetCellValue(sheet, fmt.Sprintf("U%d", r), claimed)
	}
	f.SetCellValue(sheet, "A1", "State")
	set(1, "VIN", "Claimed")
	set(2, "V-A", "false")
	set(3, "V-B", "FALSE")
	set(4, "V-C", "false")
	set(5, "V-D", "true")
	set(6, "V-E", "true")

	results := [][]string{
		make([]string, 16), // header
		append([]string{"V-D"}, append(make([]string, 14), "Alice")...),
		append([]string{"V-E"}, append(make([]string, 14), "Outsider")...),
	}

	if err := GenerateFinalReport(results, f, []string{"Alice", "Bob"}); err != nil {
		t.Fatal(err)
	}

	rows, _ := f.GetRows(sheet)
	if len(rows) != 3 {
		t.Fatalf("got %d rows, want header + 2 claimed rows: %v", len(rows), rows)
	}
	checks := []struct{ cell, want string }{
		{"V1", "FINAL"},
		{"H2", "V-D"}, {"V2", "Alice"},
		{"H3", "V-E"}, {"V3", "MISSED"},
	}
	for _, c := range checks {
		if got, _ := f.GetCellValue(sheet, c.cell); got != c.want {
			t.Errorf("%s = %q, want %q", c.cell, got, c.want)
		}
	}
}
