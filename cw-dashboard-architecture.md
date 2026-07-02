---
inclusion: auto
---

# CW Dashboard Architecture & Logic Documentation

## Overview

The **KeepingTrac Promo Report Dashboard** is a React-based single-page application that visualizes CW Network television promotional ad performance data. It is deployed on an AWS EC2 instance with **Redis** as the primary data cache and **S3** as a durable backup store.

## Data Pipeline

### Source Data Flow
1. **Raw Input**: Tab-delimited `.txt` files (`S_All_6_XXXXXXX.txt`) are generated per product estimate
2. **Excel Macros** (`OpenFile.xlsm`): Reformat raw files into `KTSpot_UnitDtls_AllColumns_XXXXXXX.txt` adding proper headers
3. **Access DB** (`HELEN1.mdb`): Imports formatted files; a UNION ALL query aggregates data across all product tables (CW_CWM, CW_WHO, CW_PEN, CW_MAS, CW_ALA, CW_SPL, CW_SPC, CW_CRE, CW_SOC, etc.)
4. **Python Automation** (`cw_automation.py`): Processes the DB output into the final `CW_Weekof_OUTPUT.xlsx` / `.csv`
5. **Dashboard Ingestion**: The CSV is loaded into Redis and served via the Flask API

### CSV Schema (43 columns)
Key columns: `AirTime`, `CW_Rank`, `CW_Market`, `CW_Station`, `CW_Network_Affiliation`, `CW_Owner`, `Agency`, `Product`, `ActRatings`, `DEMO`, `AirProgram`, `ActImpression`, `AirISCI`, `ISCI/ADID Title`, `Daypart`, `DurationCml`, `AirDate`, `AirCount`, `RtgSource`, `WeekOf`, `Designation`, `AirDay`, `Translation`, `Weekbreak`, `DayNum`, `DemoOrder`, `DAY_HH`, `MKT_Type`, `STA_Type`, `M-F length`, `TRP`, `Length in Min`

## Dashboard Logic

### Frontend (React 18 + Tailwind CSS + Babel)
- **Data Source**: Fetches JSON data from `/api/data` endpoint (backed by Redis)
- **Theming**: 5 themes (Classic/Bright/Vibrant/Pastel/Neon) via CSS variables
- **Filtering**: 9 dropdown filters + date range picker
  - Affiliation, Demographic, Station Type, Owner, Daypart, Air ISCI, Market, Product (Show Code), Station
- **Pivot Table**: Groups data by `CW_Rank | CW_Market | CW_Station | CW_Network_Affiliation`
  - Columns pivot on `Translation` (Summary tab) or `Designation` (Designation tab)
  - Metrics: Aired (count), Length-Min, TRP
  - De-duplication: Uses spot signature (`Station_Date_Time_Product`) to avoid multi-demo double-counting for Aired and Length
- **Export**: Downloads filtered data as `.xlsx` via SheetJS

### Backend (Flask + Redis + S3)
- **Redis Cache**: All CSV data stored as JSON in Redis key `cw:dashboard:data`
- **S3 Backup**: CSV uploaded to S3 bucket for durability
- **API Endpoints**:
  - `GET /` — Serves the dashboard HTML
  - `GET /api/data` — Returns JSON data from Redis
  - `POST /api/upload` — Accepts new CSV, updates Redis + S3
  - `GET /api/health` — Health check

## Deployment Architecture

```
┌─────────────────────────────────────────────────┐
│                   AWS EC2 Instance               │
│  ┌───────────┐  ┌──────────┐  ┌──────────────┐  │
│  │   Nginx   │→ │  Gunicorn│→ │  Flask App   │  │
│  │  (proxy)  │  │  (WSGI)  │  │  (backend)   │  │
│  └───────────┘  └──────────┘  └──────┬───────┘  │
│                                       │          │
│  ┌───────────────────────────────────┐│          │
│  │         Redis Server              ││          │
│  │  Key: cw:dashboard:data (JSON)    │◄┘         │
│  └───────────────────────────────────┘           │
└──────────────────────────────────────────────────┘
                        │
                        ▼
              ┌──────────────────┐
              │   AWS S3 Bucket  │
              │  cw-dashboard-   │
              │  data-backup/    │
              └──────────────────┘
```

## Key Design Decisions
- **Redis as primary store**: Sub-millisecond reads for ~71K rows of JSON data
- **S3 for durability**: Cold backup; Redis can be rehydrated from S3 on restart
- **Single HTML file**: Dashboard is self-contained (React + data) for simplicity
- **Server-side data injection**: Flask renders the dashboard template with data from Redis injected as `window.__INJECTED_DATA__`

## File References
- #[[file:cw_automation.py]] — Python automation pipeline
- #[[file:CW_Weekof_OUTPUT_06.01-07.csv]] — Current output data
- #[[file:Data Flow - CW.txt]] — Original data flow documentation
