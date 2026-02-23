package excelutils

import (
	"encoding/csv"
	"encoding/gob"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"slices"
	"strconv"
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

func Generate(xlsxPath, savePath string) error {

	//create absolute filepaths
	home, err := os.UserHomeDir()
	if err != nil {
		log.Fatal("$HOME environment variable error: ", err)
	}
	mapFilePath := filepath.Join(home, "workspace", "github.com", "jamjallred", "sf_server_utils", "assets", "airport_code_map.gob")
	templateFilePath := filepath.Join(home, "workspace", "github.com", "jamjallred", "sf_server_utils", "assets", "nationwide_template.xlsx")
	templateCopyPath := filepath.Join(home, "workspace", "github.com", "jamjallred", "sf_server_utils", "assets", "nationwide_template_copy.xlsx")

	// load airport map file
	if _, err := os.Stat(mapFilePath); err != nil {
		fmt.Println("No map found. Creating airport map...")
		createAirportMap()
	}

	airport_code_map := make(map[string]CityState)
	mapfile, err := os.Open(mapFilePath)
	if err != nil {
		fmt.Println("error opening map file:", err)
		return err
	}
	defer mapfile.Close()

	decoder := gob.NewDecoder(mapfile)
	if err = decoder.Decode(&airport_code_map); err != nil {
		fmt.Println("error decoding map: ", err)
		return err
	}

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

	if _, err := os.Stat(xlsxPath); err != nil {
		log.Fatal(err)
	}

	src, err := excelize.OpenFile(xlsxPath)
	if err != nil {
		fmt.Println("error opening file: ", err)
		return err
	}
	defer src.Close()

	if err := generateSheet(dst, src, airport_code_map, savePath); err != nil {
		fmt.Println("error generating sheet: ", err)
		return err
	}

	return nil

}

func generateSheet(dst, src *excelize.File, airport_code_map map[string]CityState, savePath string) error {

	// newHeaders := []string{"State", "City", "Yr", "Make", "Model", "Trim", "Drive", "Vin", "Color", "Miles", "Price", "MSRP", "Notes", "Notes2"}
	dst.SetCellValue("Sheet1", "G1", "Drive")                    // rename "Body Type" to "Drive"
	colIndices := []int{5, 9, 10, 11, 12, 18, 8, 17, 13, 20, 19} // column order from source sheet

	// Widen MSRP and Price Columns to avoid ###### overflow nonsense
	src.SetColWidth("Sheet1", "T", "U", 100)

	srcRows, err := src.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return err
	}

	adjust := 0 // to adjust for skipped rows due to missing airport codes
	rowNum := 0 // scope to outside loop block to use for style formatting

	fmt.Println("Generating sheet...")

	vin_list := make(map[string]struct{}) // to check for vin uniqueness
	var new_codes [][]string              // new codes to add to airport_code_map
	// saved to new csv file "airport_codes_to_update.csv"

	for i, row := range srcRows[3:] { // skipping date, empty line, header rows

		// check for vin uniqueness
		if _, ok := vin_list[row[8]]; ok {
			continue // This is a dupe in the sheet, skip over it
		} else {
			vin_list[row[8]] = struct{}{} // This is a new vin, add it to the list and process it
		}

		val, ok := airport_code_map[row[colIndices[0]]] // airport_code_map is [string]CityState
		if !ok {                                        // code, rental desc, district desc, rental zone desc
			new_codes = append(new_codes, []string{row[5], row[4], row[3], row[2]})
			adjust += 1
			continue
		}
		rowData := make([]interface{}, 0, len(colIndices))
		rowData = append(rowData, val.State, val.City)

		for _, colIdx := range colIndices[1:] {
			rowData = append(rowData, row[colIdx])
		}

		rowNum = i + 2 - adjust // A2 is first row after header
		rowRef := fmt.Sprintf("A%v", rowNum)
		if err = dst.SetSheetRow("Sheet1", rowRef, &rowData); err != nil {
			fmt.Println("error setting row:", err)
			return err
		}

		// check if MSRP is sane, if not, set to n/a
		msrpStr, err := dst.GetCellValue("Sheet1", fmt.Sprintf("L%v", rowNum))
		if err != nil {
			fmt.Println("error getting MSRP cell value:", err)
			return err
		}

		msrpStr = strings.TrimSpace(msrpStr)
		if len(msrpStr) <= 6 {
			dst.SetCellValue("Sheet1", fmt.Sprintf("L%v", rowNum), "n/a")
		}

	}

	// save new_codes by appending to the end of a csv file
	updatePath := "./assets/airport_codes_to_update.csv"
	if err := save_new_codes(new_codes, updatePath); err != nil {
		return err
	}

	//sort row data by State, City, Yr, Make, then Model
	fmt.Println("Sorting sheet...")
	srcRows = nil // free up memory
	dstRows, err := dst.GetRows("Sheet1")
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
		miles, _ := strconv.Atoi(strings.TrimSpace(row[9]))
		price, _ := strconv.Atoi(strings.TrimSpace(strings.ReplaceAll(strings.ReplaceAll(row[10], "$", ""), ",", "")))
		msrp, _ := strconv.Atoi(strings.TrimSpace(strings.ReplaceAll(strings.ReplaceAll(row[11], "$", ""), ",", "")))
		rowData := make([]interface{}, 0, len(row))
		rowData = append(rowData, row[0], row[1], year, row[3], row[4], row[5], row[6], row[7], row[8], miles, price, msrp)

		if err = dst.SetSheetRow("Sheet1", rowRef, &rowData); err != nil {
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
	})
	if err != nil {
		fmt.Println("error creating style:", err)
		return err
	}
	dst.SetCellStyle("Sheet1", "C2", fmt.Sprintf("C%v", rowNum), numID)
	dst.SetCellStyle("Sheet1", "J2", fmt.Sprintf("J%v", rowNum), numID)
	dst.SetCellStyle("Sheet1", "K2", fmt.Sprintf("L%v", rowNum), currencyID)

	// CONSIDER UPDATING THIS TO NOT SAVE NEW FILE TO DISK (MIGHT BE GOOD IDEA TO SAVE THO IDK)
	//save file
	fmt.Println("Saving file...")
	if err = dst.SaveAs(savePath); err != nil {
		log.Fatalf("error saving file: %v", err)
	}

	fmt.Println("Sheet generated successfully.")

	return nil

}

func createAirportMap() error {

	f, err := excelize.OpenFile("assets/Airport_Codes.xlsx")
	if err != nil {
		fmt.Println(err)
		return err
	}
	defer f.Close()

	mapfile, err := os.Create("assets/airport_code_map.gob")
	if err != nil {
		fmt.Println("error creating map file:", err)
		return err
	}
	defer mapfile.Close()

	airport_code_map := make(map[string]CityState)

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		fmt.Println(err)
		return err
	}

	for i, row := range rows {
		if i == 0 {
			continue // skip header
		}
		airport_code_map[row[1]] = CityState{City: row[2], State: row[3]}
	}

	encoder := gob.NewEncoder(mapfile)
	if err = encoder.Encode(airport_code_map); err != nil {
		fmt.Println("error encoding map:", err)
		return err
	}

	fmt.Println("Airport code map created successfully.")

	return nil

}

func copyTemplate(templatePath, copyPath string) error {

	template, err := os.Open(templatePath)
	if err != nil {
		fmt.Println("error opening template file:", err)
		return err
	}
	defer template.Close()

	newFile, err := os.Create(copyPath)
	if err != nil {
		fmt.Println("error creating new file:", err)
		return err
	}
	defer newFile.Close()

	_, err = io.Copy(newFile, template)
	if err != nil {
		fmt.Println("error copying template to new file:", err)
		return err
	}

	err = newFile.Sync()
	if err != nil {
		fmt.Println("error syncing file to disk:", err)
		return err
	}

	return nil

}

func save_new_codes(new_codes [][]string, filePath string) error {

	isNew := false
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		isNew = true
	}

	f, err := os.OpenFile(filePath, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
	if err != nil {
		return err
	}
	defer f.Close()

	w := csv.NewWriter(f)
	defer w.Flush()

	// if new, write header line
	header := []string{"Airport Code", "Rental Desc", "District Desc", "Rental Zone Desc"}
	if isNew {
		if err := w.Write(header); err != nil {
			return err
		}
	}

	err = w.WriteAll(new_codes)
	if err != nil {
		return err
	}

	return w.Error()

}
