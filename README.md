# sf_server_utils

Shared spreadsheet utilities (package `excelutils`) used by `sf_car_database`.

## What's in here

| Function | Input | Output |
| --- | --- | --- |
| `GenerateGrounded(ctx, csvRecords, auctions)` | Grounded inventory CSV | Grounded sheet built from `grounded_template.xlsx`, sorted, with auction names added to cities |
| `GenerateInService(src)` | In-service inventory `.xlsx` export | Nationwide sheet built from `nationwide_template.xlsx`; unknown airport codes are appended to `airport_codes_to_update.csv` |
| `GenerateFinalReport(csvRecords, dibs, repNames)` | Auction results CSV + dibs `.xlsx` | Dibs workbook with unclaimed rows removed and a `FINAL` column |
| `RebuildAirportMap()` / `LoadAirportMap()` | `Airport_Codes.xlsx` | Cached `airport_code_map.gob` |

The generators return an `*excelize.File` and never write to disk on their own. The caller decides whether to stream the file, save it, or both, and must `Close()` it.

## Assets

Templates and the airport code list are read from `SF_ASSETS_DIR`. If that isn't set, they come from `~/workspace/github.com/jamjallred/sf_server_utils/assets`. The folder needs:

- `grounded_template.xlsx`
- `nationwide_template.xlsx`
- `Airport_Codes.xlsx`

`airport_code_map.gob` is rebuilt automatically when it's missing or older than `Airport_Codes.xlsx`.

## Tests

```sh
go test ./...
```
