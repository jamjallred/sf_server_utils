package excelutils

import (
	"encoding/csv"
	"encoding/gob"
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

type CityState struct {
	City  string
	State string
}

type RentalDistrictZone struct {
	District string
	Rental   string
	Zone     string
}

const (
	airportCodesFile = "Airport_Codes.xlsx"
	airportMapFile   = "airport_code_map.gob"
	newCodesFile     = "airport_codes_to_update.csv"
)

// CreateAirportMap rebuilds the cached airport code map from
// Airport_Codes.xlsx. Kept for existing callers; see RebuildAirportMap.
func CreateAirportMap() error {
	_, err := RebuildAirportMap()
	return err
}

// RebuildAirportMap reads Airport_Codes.xlsx, writes the gob cache and
// returns the number of codes loaded.
func RebuildAirportMap() (int, error) {
	m, err := readAirportCodes()
	if err != nil {
		return 0, err
	}

	f, err := os.Create(assetPath(airportMapFile))
	if err != nil {
		return 0, fmt.Errorf("creating airport map cache: %w", err)
	}
	defer f.Close()

	if err := gob.NewEncoder(f).Encode(m); err != nil {
		return 0, fmt.Errorf("encoding airport map: %w", err)
	}
	log.Printf("airport code map rebuilt with %d codes", len(m))
	return len(m), nil
}

// LoadAirportMap returns the airport code map, rebuilding the cache when it
// is missing or older than Airport_Codes.xlsx.
func LoadAirportMap() (map[string]CityState, error) {
	cacheInfo, cacheErr := os.Stat(assetPath(airportMapFile))
	srcInfo, srcErr := os.Stat(assetPath(airportCodesFile))

	stale := cacheErr != nil || (srcErr == nil && srcInfo.ModTime().After(cacheInfo.ModTime()))
	if stale {
		if _, err := RebuildAirportMap(); err != nil {
			return nil, err
		}
	}

	f, err := os.Open(assetPath(airportMapFile))
	if err != nil {
		return nil, fmt.Errorf("opening airport map cache: %w", err)
	}
	defer f.Close()

	m := map[string]CityState{}
	if err := gob.NewDecoder(f).Decode(&m); err != nil {
		return nil, fmt.Errorf("decoding airport map: %w", err)
	}
	return m, nil
}

func readAirportCodes() (map[string]CityState, error) {
	path := assetPath(airportCodesFile)
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf("opening %s: %w", path, err)
	}
	defer f.Close()

	rows, err := f.GetRows(f.GetSheetName(0))
	if err != nil {
		return nil, fmt.Errorf("reading airport codes: %w", err)
	}

	m := make(map[string]CityState, len(rows))
	for i, row := range rows {
		code := strings.TrimSpace(cell(row, 1))
		if i == 0 || code == "" {
			continue // header or blank line
		}
		m[code] = CityState{City: cell(row, 2), State: cell(row, 3)}
	}
	return m, nil
}

// appendNewCodes appends unknown airport codes to airport_codes_to_update.csv
// so they can be added to Airport_Codes.xlsx later.
func appendNewCodes(codes [][]string) error {
	if len(codes) == 0 {
		return nil
	}
	path := assetPath(newCodesFile)
	_, statErr := os.Stat(path)
	isNew := os.IsNotExist(statErr)

	f, err := os.OpenFile(path, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0o644)
	if err != nil {
		return err
	}
	defer f.Close()

	w := csv.NewWriter(f)
	if isNew {
		if err := w.Write([]string{"Airport Code", "Rental Desc", "District Desc", "Rental Zone Desc"}); err != nil {
			return err
		}
	}
	if err := w.WriteAll(codes); err != nil {
		return err
	}
	log.Printf("saved %d unknown airport codes to %s", len(codes), path)
	return w.Error()
}
