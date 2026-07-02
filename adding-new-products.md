---
inclusion: manual
---

# Adding New Products to the CW Pipeline

## When a New Product Code is Added

In the legacy system, adding a new product (e.g., `5009565`) required:
1. Creating a new VBA macro in `OpenFile.xlsm`
2. Adding a line to the `AUndoReportsCW` master macro
3. Updating the Access UNION ALL query with a new `CW_XXX` table

## In the Python Automation

The script automatically picks up all `S_All_6_*.txt` files via glob, but **only processes products listed in `VALID_CW_PRODUCTS`**. To add a new product:

1. **Add the 3-letter product code** to the `VALID_CW_PRODUCTS` set in `cw_automation.py`:
   ```python
   VALID_CW_PRODUCTS: set[str] = {
       "ABC", "ABM", ..., "NEW",  # ← add here
   }
   ```
2. That's it. The glob pattern handles file discovery automatically.

## Why the Product Filter Exists

The `S_All_6_*.txt` feed files can contain products that aren't part of the CW promo system (e.g., NAS = national spots, BEC, LIV). Without filtering, these inflate the output with irrelevant data. The `VALID_CW_PRODUCTS` set mirrors exactly which tables were included in the Access UNION ALL query.

## If a New Reference Table Entry is Needed

If a new station, market, or group owner appears, update the relevant CSV:
- **New station**: Add row to `T_CW_Stations_COPY.csv`
- **New market**: Add row to `T_RAW_Market_List_CW.csv`
- **New group owner/station mapping**: Add row to `T_Group_Owners_LCL.csv`
- **New daypart rule**: Add rows to `T_RAW_Daypart.csv` (one per DOW per timezone per network)

## Validation
After updating, run the script with a known date range and compare per-product row counts against previous output to confirm no regressions.
