---
inclusion: auto
---

# Data Flow Reference — CW Pipeline

This documents the exact data transformations replicated from the legacy Access/VBA system.

## Input → Output Chain

```
SFTP Server (Nielsen)
       │  Step 0 (optional): Download latest S_All_6_*.txt
       ▼
S_All_6_<product>.txt  (raw tab-delimited from vendor)
       │  Step 1-3: Rename columns in-memory + filter to VALID_CW_PRODUCTS + UNION ALL
       │  (no intermediate KTSpot files needed)
       ▼
Q_UNION_ALL  (single DataFrame, valid CW products only)
       │
       ├──  Step 4a: T_ALL_RAW1  (date-filtered + daypart + group owners)
       │
       └──  Step 4b: T_CW_NewPromos_Weekly  (date-filtered + exclude 2-6 AM)
                 │  Step 4c
                 ▼
            T_CW_NewPromos  (+ Market_List_CW + Group_Owners + Weekbreak + Daypart)
                 │  Step 4d
                 ▼
            T_CW_Station_NewPromos_Weekly  (LEFT JOIN on CW_Stations master)
                 │  Step 5
                 ▼
            CW_Weekof_OUTPUT.xlsx  → Sheet2  (43 columns)
```

## Product Filtering
After union, only rows where `Product` is in `VALID_CW_PRODUCTS` are kept. This replicates the Access UNION ALL which only includes specific product tables (CW_CWM, CW_WLD, etc.). Products like NAS, BEC, LIV that appear in S_All_6 files but aren't part of the CW promo system are excluded.

## Column Rename Map (Step 1)
The raw `S_All_6` files have vendor column names (34 columns). These are mapped to the 38-column KTSpot format:
- `"Market Code"` → `"MarketCode"`
- `"Market Rank"` → `"MktRank"`
- `"Media Type"` → `"Medium"`
- `"Week Of"` → `"WeekOf"`
- `"Type of Demographic"` → `"DemoOrder"`
- `"Demographic"` → `"demo"`
- `"Buy Time Period"` → `"BuyFm-To"`
- `"Buy Prg Name"` → `"BuyPgmName"`
- `"Spot Length"` → `"LengthOfSpot"`
- `"ISCI in Buy"` → `"BuyCommercial"`
- `"Buy Dayprt"` → `"BuyDayPart"`
- `"Air Date"` → `"AirDate"`
- `"Air Day"` → `"AirDay"`
- `"Air Time"` → `"AirTime"`
- `"ISCI Length"` → `"DurationCml"`
- `"Air ISCI"` → `"AirISCI"`
- `"Cmml Title"` → `"ISCI/ADID Title"`
- `"Aired Prg Name"` → `"AirProgram"`
- `"Air Detected Event"` → `"AirDetectedEvent"`

Missing columns (ActRatingDisply, NationalMktCount, NationalMktMin, ActImpressionDisplay, ReportFlag) are added as empty strings.

## Date Parsing
Raw dates use `M/D/YYYY` format (single-digit month/day, e.g., `3/9/2026`). Polars requires zero-padded format, so dates are padded before parsing:
```python
.str.replace(r"^(\d)/", "0$1/")
.str.replace(r"/(\d)/", "/0$1/")
.str.to_date(format="%m/%d/%Y", strict=False)
```

## Daypart Lookup Logic
The original Access `DLookUp` expression:
```
DLookUp("[Daypart]", "T_RAW_Daypart",
  "[Ntwrk]='" & Network & "' AND [TimeZone]='" & TZ & "'
   And [DOW]='" & AirDay & "'
   And [StartTime]<= #AirTime# And [Endtime]>= #AirTime#")
```

Python equivalent: JOIN on `Network=Ntwrk`, `TimeZone`, `AirDay=DOW`, then filter where `StartSec <= AirTimeSec <= EndSec`. All times are converted to integer seconds-since-midnight for fast comparison.

## Key Filters
- **Date range**: `AirDate BETWEEN start_date AND end_date`
- **Time exclusion** (Step 4b): `AirTime NOT BETWEEN 02:00:00 AND 05:59:59` (seconds 7200–21599)
- **Product filter**: Only `VALID_CW_PRODUCTS` (48 codes from Access UNION ALL)

## Derived Columns
| Column | Formula |
|--------|---------|
| `GroupOwners` | `IIf(Group_Owners IS NULL, "Unassigned", "") + Group_Owners` |
| `Avg_Day` | `AirCount / 7` |
| `Mkt_Rk` | `Market + " " + MktRank` |
| `Rk_Mkt` | `MktRank + "-" + Market` (step 4c) or `Right("00"+MktRank, 3) + " -" + Market` (step 4a) |
| `Designation` | `Mid(AirISCI, 12, 1)` — character at 0-indexed position 11 |
| `M-F length` | `DurationCml` if `DayNum = 1` else `0` |
| `TIMP` | `ActImpression * 1` |
| `Length in Min` | `DurationCml / 60` |
| `TRP` | `ActRatings * 1` |
| `Translation` | From `T_CW_Product_Translation.csv` LEFT JOIN on Product (optional, blank if file not present) |

## Final Output Columns (43, in order)
```
AirTime, CW_Rank, CW_Market, CW_Station, CW_Network_Affiliation, CW_Owner,
Agency, Product, ActRatings, DEMO, AirProgram, ActImpression, AirISCI,
ISCI/ADID Title, Daypart, DurationCml, AirDate, AirCount, RtgSource, WeekOf,
Designation, AirDay, Translation, Weekbreak, DayNum, DemoOrder, DAY_HH,
MKT_Type, STA_Type, M-F length, AD25-54EST_RTG_, HHEST_RTG_, AD25-54EST_IMP,
HH_UE, HHEST_IMP, AD25-54UE, HHEST_RTG, AD25-54EST_RTG, HH_RTG_F,
AD25-54_RTG_F, TIMP, Length in Min, TRP
```
