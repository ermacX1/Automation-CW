---
inclusion: auto
---

# CW Weekly Report Automation — Project Overview

## Purpose
This project automates the CW (TV network promo tracking) weekly report pipeline that previously ran through a combination of Excel VBA macros and an MS Access database.

## Technology Stack
- **Python 3.14** (venv at `./venv`)
- **Polars** — high-performance DataFrame library (replaces pandas and Access SQL)
- **openpyxl** — Excel read/write for the final output workbook
- **paramiko** — SFTP connectivity to download feed files from remote server
- **cryptography** — PPK key parsing for SFTP auth
- Entry point: `cw_automation.py`

## Data Pipeline (6 Steps)

0. **(Optional) SFTP Download** — Connect to Nielsen SFTP server, download latest `S_All_6_*.txt` files (triggered with `--download` flag)
1. **Step 1-3** — Read all `S_All_6_<product>.txt` tab-delimited feeds directly, rename columns in-memory (replicating VBA macros), skip header-only files, filter to valid CW products only, union all into one DataFrame (`Q_UNION_ALL`)
2. **Step 4a-d** — Filter by date range, join reference tables (Group Owners, Market List, Stations, Weekbreak), compute Daypart via time-range lookup, exclude 2:00–5:59 AM window, derive calculated columns → final `T_CW_Station_NewPromos_Weekly` table
3. **Step 5** — Write results to Sheet2 of the output Excel workbook (43 columns matching legacy format)

## SFTP Configuration
| Setting | Value |
|---------|-------|
| Host | `10.207.176.203` |
| Protocol | SFTP (SSH) |
| Key File | `C:\DATA FEED\nts-upload.ppk` (PuTTY PPK v2 RSA, unencrypted) |
| Remote Path | `/us-east-1-nlsn-watch-adintel-nts-client-sftp-prod/NTS/cw_csv/Detection` |

## Key Reference Files (CSVs)
| File | Purpose |
|------|---------|
| `T_RAW_Daypart.csv` | Daypart lookup by Network + TimeZone + Day-of-Week + time range |
| `T_Group_Owners_LCL.csv` | Station → Group Owner mapping |
| `T_RAW_Market_List_CW.csv` | MarketCode → TimeZone + Market Name |
| `T_CW_Stations_COPY.csv` | CW station master list (Rank, Market, Network, Owner, Type) |
| `Weekbreak.csv` | Day-of-week → Weekday/Weekend + DayNum flag |

## Valid CW Products
Only these product codes (from the Access UNION ALL query) are processed:
```
ABC, ABM, ABW, ACC, ALA, AVP, BAN, BBO, CHO, CNR, CRE, CRI,
CWM, CWS, FAM, GBC, GST, HAR, HLL, IAM, JOA, LOT, MAS, MVE,
NOR, PAC, PBA, PBR, PEN, POL, SAD, SCR, SOC, SPC, SPL, SUL,
SUN, TFA, TFK, TRV, TWL, WFA, WHO, WLD, WRA, WRG, WXT, 61S
```

## Running the Script
```bash
# Basic run (local files already present)
venv\Scripts\python.exe cw_automation.py --start YYYY-MM-DD --end YYYY-MM-DD

# Download latest from SFTP first, then process
venv\Scripts\python.exe cw_automation.py --start YYYY-MM-DD --end YYYY-MM-DD --download

# Point at a custom feed directory
venv\Scripts\python.exe cw_automation.py --start YYYY-MM-DD --end YYYY-MM-DD --data-feed-dir "C:\path\to\feeds"
```

## Output
`CW_Weekof_OUTPUT.xlsx` — Sheet2 contains the final T_CW_Station_NewPromos_Weekly data (43 columns).
