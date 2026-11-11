package excelutils

import (
	"context"
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
)

// AuctionLookup finds the auction that serves a city. found is false (with a
// nil error) when the city has no auction mapped.
type AuctionLookup interface {
	LookupAuction(ctx context.Context, city, state string) (name string, found bool, err error)
}

// Source CSV columns, in output order:
// State, City, Year, Make, Model, Trim, Drivetrain, VIN, Mileage, Color,
// Buy It Now Price, Vehicle Link.
var groundedColumns = []int{11, 10, 1, 2, 3, 4, 13, 0, 5, 8, 7, 12}

// GenerateGrounded builds the grounded inventory workbook from the uploaded
// CSV records (header row included). The caller owns the returned file and
// must Close it after writing or saving it.
func GenerateGrounded(ctx context.Context, csvRecords [][]string, auctions AuctionLookup) (*excelize.File, error) {
	if len(csvRecords) < 2 {
		return nil, fmt.Errorf("the CSV has no data rows")
	}

	dst, err := excelize.OpenFile(assetPath("grounded_template.xlsx"))
	if err != nil {
		return nil, fmt.Errorf("opening grounded template: %w", err)
	}

	if err := fillGroundedSheet(ctx, dst, csvRecords[1:], auctions); err != nil {
		dst.Close()
		return nil, err
	}
	return dst, nil
}

func fillGroundedSheet(ctx context.Context, dst *excelize.File, records [][]string, auctions AuctionLookup) error {
	sheet := dst.GetSheetName(0)

	// Cache lookups: the same city appears on many rows.
	type key struct{ city, state string }
	auctionCache := map[key]string{}

	rows := make([][]string, 0, len(records))
	for _, rec := range records {
		row := pick(rec, groundedColumns)
		state, city := row[0], row[1]

		k := key{city, state}
		name, cached := auctionCache[k]
		if !cached && auctions != nil {
			n, found, err := auctions.LookupAuction(ctx, city, state)
			if err != nil {
				log.Printf("auction lookup failed for %s, %s: %v", city, state, err)
			} else if found {
				name = n
			}
			auctionCache[k] = name
		}
		if name != "" {
			row[1] = fmt.Sprintf("%s (%s)", city, name)
		}
		rows = append(rows, row)
	}

	sortInventory(rows)

	for i, row := range rows {
		values := []any{
			row[0], row[1], numberOrText(row[2]), row[3], row[4], row[5], row[6], row[7],
			numberOrText(row[8]), row[9], numberOrText(row[10]), row[11],
		}
		if err := dst.SetSheetRow(sheet, fmt.Sprintf("A%d", i+2), &values); err != nil {
			return fmt.Errorf("writing row %d: %w", i+2, err)
		}
	}

	styles, err := newStyles(dst, "#92D050")
	if err != nil {
		return fmt.Errorf("creating styles: %w", err)
	}
	last := len(rows) + 1
	for _, s := range []struct {
		from, to string
		style    int
	}{
		{"C", "C", styles.number},   // year
		{"I", "I", styles.number},   // mileage
		{"K", "K", styles.currency}, // price
	} {
		if err := styleColumns(dst, sheet, s.from, s.to, last, s.style); err != nil {
			return fmt.Errorf("styling columns: %w", err)
		}
	}
	return nil
}
