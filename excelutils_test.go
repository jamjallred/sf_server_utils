package excelutils

import (
	"fmt"
	"strings"
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

// The All Dibs export: first-name dibs columns, no Claimed column, and reps in
// the results written as full names (sometimes "Michael" for "Mike").
func TestGenerateFinalReportAllDibsLayout(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()
	sheet := f.GetSheetName(0)

	header := []any{"State", "City", "Yr", "Make", "Model", "Trim", "Drive", "Vin", "Color", "Miles", "Price", "MSRP",
		"2nd Wave", "Mike", "Stan", "Jaxon", "Gregg", "Mike Notes", "Stan Notes", "Jaxon Notes", "Gregg Notes"}
	f.SetSheetRow(sheet, "A1", &header)
	vins := []string{"V-1", "V-2", "V-3", "V-4", "V-5", "V-6", "V-7", "V-8"}
	for i, v := range vins {
		f.SetCellValue(sheet, fmt.Sprintf("H%d", i+2), v)
		// A note that says "false" must not remove the row.
		f.SetCellValue(sheet, fmt.Sprintf("U%d", i+2), "false")
	}

	rep := func(vin, name string) []string {
		return append([]string{vin}, append(make([]string, 14), name)...)
	}
	results := [][]string{
		make([]string, 16),
		rep("V-1", "Jaxon Coppinger"),
		rep("V-2", "Stan Coppinger"),
		rep("v-3 ", "Gregg Wessels"), // VIN case and spacing differ
		rep("V-4", "Mike Allred"),
		rep("V-5", "Michael Allred"),
		rep("V-6", "  MICHAEL\u00a0 allred "),
		rep("V-7", "Mike Smith"), // same first name, not on the team
		rep("V-8", ""),
	}
	team := []string{"Mike Allred", "Stan Coppinger", "Jaxon Coppinger", "Gregg Wessels"}
	if err := GenerateFinalReport(results, f, team); err != nil {
		t.Fatal(err)
	}

	rows, _ := f.GetRows(sheet)
	if len(rows) != 9 {
		t.Fatalf("got %d rows, want all 9 kept", len(rows))
	}
	want := map[string]string{
		"V1": "FINAL", "U1": "Gregg Notes",
		"V2": "Jaxon Coppinger", "V3": "Stan Coppinger", "V4": "Gregg Wessels",
		"V5": "Mike Allred", "V6": "Mike Allred", "V7": "Mike Allred",
		"V8": "MISSED", "V9": "MISSED",
	}
	for ref, w := range want {
		if got, _ := f.GetCellValue(sheet, ref); got != w {
			t.Errorf("%s = %q, want %q", ref, got, w)
		}
	}
}

// When the results export moves the rep column, it is found by its names.
func TestGenerateFinalReportFindsMovedRepColumn(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()
	sheet := f.GetSheetName(0)
	f.SetSheetRow(sheet, "A1", &[]any{"State", "City", "Yr", "Make", "Model", "Trim", "Drive", "Vin", "Mike", "Stan"})
	f.SetCellValue(sheet, "H2", "V-1")
	f.SetCellValue(sheet, "H3", "V-2")
	results := [][]string{
		{"VIN", "x", "Buyer"},
		{"V-1", "", "Stan Coppinger"},
		{"V-2", "", "Someone Else"},
	}
	if err := GenerateFinalReport(results, f, []string{"Mike Allred", "Stan Coppinger"}); err != nil {
		t.Fatal(err)
	}
	for ref, w := range map[string]string{"K1": "FINAL", "K2": "Stan Coppinger", "K3": "MISSED"} {
		if got, _ := f.GetCellValue(sheet, ref); got != w {
			t.Errorf("%s = %q, want %q", ref, got, w)
		}
	}
}

// The real auction export: VIN first, the rep under "Avis Rep Name" (column
// W), and column P is "Image Count".
func TestGenerateFinalReportAuctionExportHeaders(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()
	sheet := f.GetSheetName(0)
	f.SetSheetRow(sheet, "A1", &[]any{"State", "City", "Yr", "Make", "Model", "Trim", "Drive", "Vin", "Mike", "Gregg"})
	f.SetCellValue(sheet, "H2", "KNDNB5K32S6519491")
	f.SetCellValue(sheet, "H3", "W1KAF4GB6SR282330")

	header := strings.Split("\ufeffVIN,Year,Make,Model,Trim,Drivetrain,MVA,Auction End Date,Auction Start Date,Location,Hits,Buy Now,Opening Price,Days Remaining,Tenant Badge,Image Count,Service Fee,Sale Fee,Seller,Status of Vehicle,Reservation Date,Reservation Time,Avis Rep Name,Invoice Date,Invoice Status", ",")
	row := func(vin, rep string) []string {
		r := make([]string, len(header))
		r[0], r[15], r[22] = vin, "1", rep
		return r
	}
	results := [][]string{header, row("KNDNB5K32S6519491", "Michael Allred"), row("W1KAF4GB6SR282330", "Gregg Wessels")}
	if err := GenerateFinalReport(results, f, []string{"Mike Allred", "Gregg Wessels"}); err != nil {
		t.Fatal(err)
	}
	for ref, w := range map[string]string{"K1": "FINAL", "K2": "Mike Allred", "K3": "Gregg Wessels"} {
		if got, _ := f.GetCellValue(sheet, ref); got != w {
			t.Errorf("%s = %q, want %q", ref, got, w)
		}
	}
}
