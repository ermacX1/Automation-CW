"""
CW Weekly Report Automation  —  Polars edition
================================================
Replicates the full Access/Excel macro pipeline described in Data Flow - CW.txt

Pipeline steps:
  0.   (Optional) Download latest S_All_6_*.txt from SFTP
  1-3. Read all S_All_6_<product>.txt raw feed files from DATA_FEED_DIR,
       rename columns (replicating VBA macros), union all → Q_UNION_ALL
  4a.  Filter to date range, join Group_Owners + Market_List, derive Daypart → T_ALL_RAW1
  4b.  Build T_CW_NewPromos_Weekly  (filter air-time, exclude 2:00–5:59 AM)
  4c.  Build T_CW_NewPromos  (join Market_List_CW, Group_Owners, Weekbreak)
  4d.  Build T_CW_Station_NewPromos_Weekly  (join CW_Stations)
  5.   Write Sheet2 of the output workbook

Usage:
  python cw_automation.py --start 2026-06-01 --end 2026-06-07
  python cw_automation.py --start 2026-06-01 --end 2026-06-07 --download

All file paths live in the CONFIG section below.
"""

import os
import sys
import glob
import shutil
import stat
import base64
import struct
import io
import argparse
import logging
from datetime import datetime, time

import polars as pl
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import paramiko
from cryptography.hazmat.primitives.asymmetric.rsa import (
    rsa_crt_dmp1, rsa_crt_dmq1, RSAPrivateNumbers, RSAPublicNumbers,
)
from cryptography.hazmat.primitives.serialization import Encoding, PrivateFormat, NoEncryption
from cryptography.hazmat.backends import default_backend

# ──────────────────────────────────────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────────────────────────────────────
BASE_DIR      = r"C:\Users\arbarman2401\Documents\Automation\CW"
DATA_FEED_DIR = r"C:\Users\arbarman2401\Documents\Automation\CW\All_Files_S_2Week_April_26"   # where S_All_6_*.txt files live

# SFTP Configuration
SFTP_HOST     = "10.207.176.203"
SFTP_PORT     = 22
SFTP_USERNAME = "nts_upload"
SFTP_KEY_PATH = r"C:\DATA FEED\nts-upload.ppk"
SFTP_REMOTE_DIR = "/us-east-1-nlsn-watch-adintel-nts-client-sftp-prod/NTS/cw_csv/Detections"

DAYPART_CSV      = os.path.join(BASE_DIR, "T_RAW_Daypart.csv")
GROUP_OWNERS_CSV = os.path.join(BASE_DIR, "T_Group_Owners_LCL.csv")
MARKET_LIST_CSV  = os.path.join(BASE_DIR, "T_RAW_Market_List_CW.csv")
STATIONS_CSV             = os.path.join(BASE_DIR, "T_CW_Stations_COPY.csv")
PRODUCT_TRANSLATION_CSV  = os.path.join(BASE_DIR, "T_CW_Product_Translation.csv")
WEEKBREAK_CSV            = os.path.join(BASE_DIR, "Weekbreak.csv")

OUTPUT_XLSX   = os.path.join(BASE_DIR, "CW_Weekof_OUTPUT.xlsx")
OUTPUT_CSV    = os.path.join(BASE_DIR, "CW_Weekof_OUTPUT.csv")  # will be overridden with date-stamped name
TEMPLATE_XLSX = os.path.join(BASE_DIR, "CW_Weekof-03.09-15.26.xlsx")

# ──────────────────────────────────────────────────────────────────────────────
# Column rename map: S_All_6 raw headers  →  KTSpot / Access naming
# ──────────────────────────────────────────────────────────────────────────────
S_ALL_RENAME: dict[str, str] = {
    "Agency":               "Agency",
    "Advertiser":           "Advertiser",
    "Product":              "Product",
    "Estimate":             "Estimate",
    "Market Code":          "MarketCode",
    "Market Rank":          "MktRank",
    "Market":               "Market",
    "Network":              "Network",
    "Station":              "Station",
    "Media Type":           "Medium",
    "Week Of":              "WeekOf",
    "Type of Demographic":  "DemoOrder",
    "Demographic":          "demo",
    "Buy Rotation":         "BuyRotation",
    "Buy Time Period":      "BuyFm-To",
    "Buy Prg Name":         "BuyPgmName",
    "Spot Length":          "LengthOfSpot",
    "ISCI in Buy":          "BuyCommercial",
    "Buy Dayprt":           "BuyDayPart",
    "Air Date":             "AirDate",
    "Air Day":              "AirDay",
    "Air Time":             "AirTime",
    "ISCI Length":          "DurationCml",
    "Air ISCI":             "AirISCI",
    "Cmml Title":           "ISCI/ADID Title",
    "Aired Prg Name":       "AirProgram",
    "Est Ratings":          "EstRatings",
    "Act Ratings":          "ActRatings",
    "Act Impression":       "ActImpression",
    "Rtg Source":           "RtgSource",
    "Buy Count":            "BuyCount",
    "Air Count":            "AirCount",
    "Air Detected Event":   "AirDetectedEvent",
    # already correctly named in KTSpot files
    "ActRatingDisply":      "ActRatingDisply",
    "NationalMktCount":     "NationalMktCount",
    "NationalMktMin":       "NationalMktMin",
    "ActImpressionDisplay": "ActImpressionDisplay",
    "ReportFlag":           "ReportFlag",
}

KTSPOT_COLS: list[str] = [
    "Agency", "Advertiser", "Product", "Estimate",
    "MarketCode", "MktRank", "Market", "Network", "Station", "Medium",
    "WeekOf", "DemoOrder", "demo", "BuyRotation", "BuyFm-To", "BuyPgmName",
    "LengthOfSpot", "BuyCommercial", "BuyDayPart",
    "AirDate", "AirDay", "AirTime", "DurationCml",
    "AirISCI", "ISCI/ADID Title", "AirProgram",
    "EstRatings", "ActRatings", "ActImpression",
    "RtgSource", "BuyCount", "AirCount", "AirDetectedEvent",
    "ActRatingDisply", "NationalMktCount", "NationalMktMin",
    "ActImpressionDisplay", "ReportFlag",
]

# ──────────────────────────────────────────────────────────────────────────────
# Valid CW product codes (from Access UNION ALL query tables)
# Only data matching these products is included in the pipeline.
# ──────────────────────────────────────────────────────────────────────────────
VALID_CW_PRODUCTS: set[str] = {
    "ABC", "ABM", "ABW", "ACC", "ALA", "AVP", "BAN", "BBO",
    "CHO", "CNR", "CRE", "CRI", "CWM", "CWS", "FAM", "GBC",
    "GST", "HAR", "HLL", "IAM", "JOA", "LOT", "MAS", "MVE",
    "NOR", "PAC", "PBA", "PBR", "PEN", "POL", "SAD", "SCR",
    "SOC", "SPC", "SPL", "SUL", "SUN", "TFA", "TFK", "TRV",
    "TWL", "WFA", "WHO", "WLD", "WRA", "WRG", "WXT", "61S",
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════════════
# STEP 0 — SFTP Download
# ══════════════════════════════════════════════════════════════════════════════

def _parse_ppk_v2(ppk_path: str) -> paramiko.RSAKey:
    """Parse a PuTTY PPK v2 unencrypted RSA key → paramiko RSAKey."""
    with open(ppk_path, 'r') as f:
        lines = f.readlines()

    key_type = None
    encryption = None
    public_b64: list[str] = []
    private_b64: list[str] = []

    i = 0
    while i < len(lines):
        line = lines[i].rstrip('\n\r')
        if line.startswith('PuTTY-User-Key-File-2:'):
            key_type = line.split(':', 1)[1].strip()
        elif line.startswith('Encryption:'):
            encryption = line.split(':', 1)[1].strip()
        elif line.startswith('Public-Lines:'):
            count = int(line.split(':', 1)[1].strip())
            for _ in range(count):
                i += 1
                public_b64.append(lines[i].strip())
        elif line.startswith('Private-Lines:'):
            count = int(line.split(':', 1)[1].strip())
            for _ in range(count):
                i += 1
                private_b64.append(lines[i].strip())
        i += 1

    if encryption != 'none':
        raise ValueError("Only unencrypted PPK files are supported")
    if key_type != 'ssh-rsa':
        raise ValueError(f"Expected ssh-rsa key, got {key_type}")

    pub_data = base64.b64decode(''.join(public_b64))
    priv_data = base64.b64decode(''.join(private_b64))

    def _read_mpint(data: bytes, offset: int) -> tuple[int, int]:
        length = struct.unpack('>I', data[offset:offset+4])[0]
        value = int.from_bytes(data[offset+4:offset+4+length], 'big')
        return value, offset + 4 + length

    def _read_string(data: bytes, offset: int) -> tuple[bytes, int]:
        length = struct.unpack('>I', data[offset:offset+4])[0]
        value = data[offset+4:offset+4+length]
        return value, offset + 4 + length

    # Public: "ssh-rsa" + e + n
    offset = 0
    _, offset = _read_string(pub_data, offset)
    e, offset = _read_mpint(pub_data, offset)
    n, offset = _read_mpint(pub_data, offset)

    # Private: d + p + q + iqmp
    offset = 0
    d, offset = _read_mpint(priv_data, offset)
    p, offset = _read_mpint(priv_data, offset)
    q, offset = _read_mpint(priv_data, offset)
    iqmp_val, offset = _read_mpint(priv_data, offset)

    dmp1 = rsa_crt_dmp1(d, p)
    dmq1 = rsa_crt_dmq1(d, q)
    pub_numbers = RSAPublicNumbers(e, n)
    priv_numbers = RSAPrivateNumbers(p, q, d, dmp1, dmq1, iqmp_val, pub_numbers)
    private_key = priv_numbers.private_key(default_backend())

    pem_bytes = private_key.private_bytes(Encoding.PEM, PrivateFormat.OpenSSH, NoEncryption())
    return paramiko.RSAKey.from_private_key(io.StringIO(pem_bytes.decode()))


def step0_download_feeds(
    local_dir: str,
    host: str = SFTP_HOST,
    port: int = SFTP_PORT,
    username: str = SFTP_USERNAME,
    key_path: str = SFTP_KEY_PATH,
    remote_dir: str = SFTP_REMOTE_DIR,
    max_workers: int = 8,
) -> None:
    """
    Connect to the SFTP server and download the latest S_All_6_*.txt files
    using multithreaded parallel downloads.

    Logic:
      - List all S_All_6_*.txt on the server (latest batch within 24h of newest)
      - Compare local file size to remote: skip if identical, else re-download
      - Download in parallel using a thread pool (one SFTP connection per thread)
      - After download, delete any local file that has exactly 1 data row (header-only)
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed

    log.info("=== STEP 0: Downloading latest feeds from SFTP ===")
    log.info("  Host: %s:%d", host, port)
    log.info("  Remote dir: %s", remote_dir)
    log.info("  Local dir: %s", local_dir)

    # Load PPK key
    pkey = _parse_ppk_v2(key_path)
    log.info("  Key loaded: %d-bit RSA from %s", pkey.get_bits(), os.path.basename(key_path))

    # Use a control connection to list files
    transport = paramiko.Transport((host, port))
    transport.connect(username=username, pkey=pkey)
    sftp = paramiko.SFTPClient.from_transport(transport)

    try:
        # List remote files
        remote_files = sftp.listdir_attr(remote_dir)
        s_all_files = [
            f for f in remote_files
            if f.filename.startswith("S_All_6_") and f.filename.endswith(".txt")
        ]
        log.info("  Found %d S_All_6_*.txt files on server", len(s_all_files))

        if not s_all_files:
            log.warning("  No S_All_6 files found in remote directory")
            return

        # Sort by modification time descending
        s_all_files.sort(key=lambda f: f.st_mtime or 0, reverse=True)

        # Latest batch: files within 24h of the newest
        latest_mtime = s_all_files[0].st_mtime
        cutoff_mtime = latest_mtime - 86400
        latest_batch = [f for f in s_all_files if (f.st_mtime or 0) >= cutoff_mtime]
        log.info("  Latest batch: %d files (within 24h of newest)", len(latest_batch))

        # Ensure local directory exists
        os.makedirs(local_dir, exist_ok=True)

        # Determine which files need downloading (size mismatch or missing)
        to_download: list[tuple[str, str, int]] = []  # (remote_path, local_path, size)
        skipped = 0
        for remote_file in latest_batch:
            remote_path = f"{remote_dir}/{remote_file.filename}"
            local_path = os.path.join(local_dir, remote_file.filename)
            remote_size = remote_file.st_size or 0

            if os.path.exists(local_path):
                local_size = os.path.getsize(local_path)
                if local_size == remote_size:
                    skipped += 1
                    continue
                # Size mismatch → will re-download (replace local)

            to_download.append((remote_path, local_path, remote_size))

        log.info("  To download: %d, Already up-to-date: %d", len(to_download), skipped)

    finally:
        sftp.close()
        transport.close()

    if not to_download:
        log.info("  All files are current — nothing to download")
        _cleanup_single_row_files(local_dir)
        return

    # ── Parallel download using thread pool ──────────────────────────────────
    def _download_one(item: tuple[str, str, int]) -> str:
        """Download a single file using its own SFTP connection."""
        r_path, l_path, r_size = item
        t = paramiko.Transport((host, port))
        t.connect(username=username, pkey=pkey)
        s = paramiko.SFTPClient.from_transport(t)
        try:
            s.get(r_path, l_path)
            return os.path.basename(l_path)
        finally:
            s.close()
            t.close()

    downloaded = 0
    errors = 0
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = {pool.submit(_download_one, item): item for item in to_download}
        for future in as_completed(futures):
            item = futures[future]
            filename = os.path.basename(item[1])
            try:
                future.result()
                downloaded += 1
                log.info("  ✓ %s (%d bytes)", filename, item[2])
            except Exception as e:
                errors += 1
                log.error("  ✗ %s FAILED: %s", filename, str(e))

    log.info("  Downloaded: %d, Errors: %d, Skipped: %d", downloaded, errors, skipped)

    # ── Post-download: remove header-only files (exactly 1 row) ──────────────
    _cleanup_single_row_files(local_dir)


def _cleanup_single_row_files(local_dir: str) -> None:
    """Delete S_All_6_*.txt files that have only a header row (no data)."""
    removed = 0
    for path in glob.glob(os.path.join(local_dir, "S_All_6_*.txt")):
        # Quick heuristic: files ≤ 500 bytes are almost certainly header-only
        if os.path.getsize(path) <= 500:
            try:
                with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                    lines = f.readlines()
                # Header + at most 1 empty/whitespace line = no real data
                data_lines = [l for l in lines[1:] if l.strip()]
                if len(data_lines) <= 1:
                    os.remove(path)
                    removed += 1
            except Exception:
                pass
    if removed:
        log.info("  Cleaned up %d header-only file(s)", removed)


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _read_tsv(path: str) -> pl.DataFrame:
    """Read a tab-separated text file into a Polars DataFrame (all strings)."""
    return pl.read_csv(
        path,
        separator="\t",
        infer_schema=False,
        ignore_errors=True,
        truncate_ragged_lines=True,
    )


def _ensure_cols(df: pl.DataFrame, cols: list[str]) -> pl.DataFrame:
    """Add any missing columns as empty-string literals."""
    for col in cols:
        if col not in df.columns:
            df = df.with_columns(pl.lit("").alias(col))
    return df.select(cols)


def _parse_time_col(series: pl.Series) -> pl.Series:
    """
    Parse a string time column (HH:MM:SS or hh:MM:SS AM/PM) to
    total seconds since midnight as Int64 — used for range comparisons.
    Returns null for unparseable values.
    """
    def _to_seconds(s: str | None) -> int | None:
        if not s:
            return None
        s = s.strip()
        for fmt in ("%H:%M:%S", "%I:%M:%S %p", "%H:%M"):
            try:
                t = datetime.strptime(s, fmt).time()
                return t.hour * 3600 + t.minute * 60 + t.second
            except ValueError:
                continue
        return None

    return series.map_elements(_to_seconds, return_dtype=pl.Int64)


def _load_daypart_table(path: str) -> pl.DataFrame:
    """
    Load T_RAW_Daypart.csv and pre-convert times to seconds-since-midnight
    so range comparisons are fast integer comparisons.
    """
    log.info("Loading daypart table: %s", os.path.basename(path))
    dp = pl.read_csv(path, infer_schema=False)
    dp = dp.rename({c: c.strip() for c in dp.columns})

    dp = dp.with_columns([
        pl.col("DOW").str.strip_chars(),
        pl.col("TimeZone").str.strip_chars(),
        pl.col("Ntwrk").str.strip_chars(),
        _parse_time_col(dp["StartTime"]).alias("StartSec"),
        _parse_time_col(dp["EndTime"]).alias("EndSec"),
    ])
    log.info("  Daypart table: %d rows", len(dp))
    return dp


def _apply_daypart(df: pl.DataFrame, dp: pl.DataFrame) -> pl.DataFrame:
    """
    Vectorised daypart lookup.
    Merges df with dp on Network+TimeZone+AirDay, filters by time range,
    picks the first match per original row, adds a 'Daypart' column.
    """
    # Convert AirTime to seconds
    working = df.with_columns(
        _parse_time_col(df["AirTime"]).alias("_AirSec"),
        pl.Series("_row_id", range(len(df)), dtype=pl.Int64),
    )

    # Join on the three key columns
    joined = working.join(
        dp.select(["Ntwrk", "TimeZone", "DOW", "StartSec", "EndSec", "Daypart"]),
        left_on=["Network", "TimeZone", "AirDay"],
        right_on=["Ntwrk", "TimeZone", "DOW"],
        how="left",
    )

    # Keep only rows where AirTime falls inside [StartSec, EndSec]
    matched = joined.filter(
        pl.col("_AirSec").is_not_null()
        & pl.col("StartSec").is_not_null()
        & (pl.col("_AirSec") >= pl.col("StartSec"))
        & (pl.col("_AirSec") <= pl.col("EndSec"))
    )

    # One match per original row (first one wins)
    best = matched.unique(subset=["_row_id"], keep="first").select(
        ["_row_id", "Daypart"]
    )

    # Left-join back so every original row gets a Daypart (null → "")
    result = working.join(best, on="_row_id", how="left").with_columns(
        pl.col("Daypart").fill_null("").alias("Daypart")
    ).drop(["_AirSec", "_row_id"])

    return result


# ══════════════════════════════════════════════════════════════════════════════
# STEP 1-3 — Read S_All_6 raw feeds → rename columns → union (Q_UNION_ALL)
# ══════════════════════════════════════════════════════════════════════════════

def step1_3_read_and_union_raw_feeds(data_feed_dir: str) -> pl.DataFrame:
    """
    Replicate the VBA macros + Access UNION ALL in a single step:
      - Read every S_All_6_<product>.txt (tab-delimited raw feed)
      - Rename columns to KTSpot / Access naming convention
      - Add any missing columns (ActRatingDisply, NationalMktCount, etc.)
      - Skip files that are header-only (no data rows)
      - Union all into one DataFrame  →  Q_UNION_ALL

    This replaces both the VBA macro step (column rename + save as KTSpot)
    and the Access UNION ALL query across per-product tables.
    """
    log.info("=== STEP 1-3: Reading S_All_6 raw feeds → column rename → UNION ALL ===")

    raw_files = glob.glob(os.path.join(data_feed_dir, "S_All_6_*.txt"))
    if not raw_files:
        raise FileNotFoundError(
            f"No S_All_6_*.txt files found in {data_feed_dir}"
        )

    # Only process files with actual data (>400 bytes = not header-only)
    raw_files = [f for f in raw_files if os.path.getsize(f) > 400]
    log.info("  Found %d S_All_6 files with data", len(raw_files))

    frames: list[pl.DataFrame] = []
    for raw_path in sorted(raw_files):
        basename = os.path.basename(raw_path)
        try:
            df = _read_tsv(raw_path)
            # Strip whitespace from column names
            df = df.rename({c: c.strip() for c in df.columns})
            # Apply the column rename (VBA macro equivalent)
            rename_map = {c: S_ALL_RENAME[c] for c in df.columns if c in S_ALL_RENAME}
            if rename_map:
                df = df.rename(rename_map)
            # Ensure all KTSpot columns exist (add missing ones as empty)
            df = _ensure_cols(df, KTSPOT_COLS)
            # Skip if no actual data rows
            if len(df) == 0:
                continue
            frames.append(df)
            log.info("  %s: %d rows", basename, len(df))
        except Exception as e:
            log.warning("  %s: SKIPPED (%s)", basename, str(e))
            continue

    if not frames:
        raise ValueError("No valid data found in any S_All_6 files")

    union_df = pl.concat(frames, rechunk=True)
    log.info("  Q_UNION_ALL (raw): %d rows from %d file(s)", len(union_df), len(frames))

    # Filter to only valid CW products (replicates the Access UNION ALL
    # which only includes specific product tables like CW_CWM, CW_WLD, etc.)
    union_df = union_df.filter(
        pl.col("Product").is_in(VALID_CW_PRODUCTS)
    )
    log.info("  Q_UNION_ALL (filtered to valid CW products): %d rows", len(union_df))

    return union_df


# ══════════════════════════════════════════════════════════════════════════════
# STEP 4a — T_ALL_RAW1
# ══════════════════════════════════════════════════════════════════════════════

def step4a_t_all_raw1(
    union_df: pl.DataFrame,
    start_date: datetime,
    end_date: datetime,
    dp: pl.DataFrame,
    group_owners: pl.DataFrame,
    market_list: pl.DataFrame,
) -> pl.DataFrame:
    """
    Replicates Access query (a):
      - Date-filter AirDate
      - LEFT JOIN Group_Owners on Station = OnO_Stations
      - LEFT JOIN Market_List  on MarketCode  (gets TimeZone)
      - DLookUp Daypart
      - Derived columns: GroupOwners, Avg_Day, Mkt_Rk, Rk_Mkt …
    """
    log.info("=== STEP 4a: Building T_ALL_RAW1 ===")

    df = union_df.with_columns(
        pl.col("AirDate").str.strip_chars()
          .str.replace(r"^(\d)/", "0$1/")
          .str.replace(r"/(\d)/", "/0$1/")
          .str.to_date(format="%m/%d/%Y", strict=False)
    ).filter(
        pl.col("AirDate").is_between(start_date.date(), end_date.date())
    )
    log.info("  After date filter: %d rows", len(df))

    # Cast numeric columns
    for col in ["MktRank", "AirCount", "ActRatings", "ActImpression", "DurationCml"]:
        df = df.with_columns(
            pl.col(col).cast(pl.Float64, strict=False)
        )

    # LEFT JOIN Group_Owners
    go = group_owners.select(
        pl.col("OnO_Stations").str.strip_chars(),
        pl.col("Group_Owners").str.strip_chars(),
    ).unique(subset=["OnO_Stations"])
    df = df.join(go, left_on="Station", right_on="OnO_Stations", how="left")

    # LEFT JOIN Market_List for TimeZone
    tz_cols = ["MarketCode"] + (["TimeZone"] if "TimeZone" in market_list.columns else [])
    ml = market_list.select(tz_cols).unique(subset=["MarketCode"])
    df = df.join(ml, on="MarketCode", how="left")

    if "TimeZone" not in df.columns:
        df = df.with_columns(pl.lit("").alias("TimeZone"))

    # Daypart lookup
    df = _apply_daypart(df, dp)

    # Derived columns
    df = df.with_columns([
        pl.when(pl.col("Group_Owners").is_null())
          .then(pl.lit("Unassigned"))
          .otherwise(pl.lit(""))
          .alias("TEST"),
    ]).with_columns(
        (pl.col("TEST") + pl.col("Group_Owners").fill_null("")).alias("GroupOwners"),
        (pl.col("AirCount").fill_null(0) / 7).alias("Avg_Day"),
        (pl.col("Market").fill_null("") + pl.lit(" ") +
         pl.col("MktRank").fill_null(0).cast(pl.Int64).cast(pl.Utf8)).alias("Mkt_Rk"),
        (pl.lit("00") +
         pl.col("MktRank").fill_null(0).cast(pl.Int64).cast(pl.Utf8)).alias("MktRank00"),
    ).with_columns(
        pl.col("MktRank00").str.slice(-3).alias("MktRankRight")
    ).with_columns(
        (pl.col("MktRankRight") + pl.lit(" -") +
         pl.col("Market").fill_null("")).alias("Rk_Mkt")
    )

    log.info("  T_ALL_RAW1: %d rows", len(df))
    return df


# ══════════════════════════════════════════════════════════════════════════════
# STEP 4b — T_CW_NewPromos_Weekly
# ══════════════════════════════════════════════════════════════════════════════

def step4b_t_cw_newpromos_weekly(
    union_df: pl.DataFrame,
    start_date: datetime,
    end_date: datetime,
) -> pl.DataFrame:
    """
    Replicates Access query (b):
      - AirDate between CWStartDate and CWEndDate
      - AirTime NOT between 02:00:00 and 05:59:59
    """
    log.info("=== STEP 4b: Building T_CW_NewPromos_Weekly ===")

    df = union_df.with_columns(
        pl.col("AirDate").str.strip_chars()
          .str.replace(r"^(\d)/", "0$1/")
          .str.replace(r"/(\d)/", "/0$1/")
          .str.to_date(format="%m/%d/%Y", strict=False)
    ).filter(
        pl.col("AirDate").is_between(start_date.date(), end_date.date())
    )

    # Exclude AirTime 02:00:00 – 05:59:59  (7200 – 21599 seconds)
    df = df.with_columns(
        _parse_time_col(df["AirTime"]).alias("_AirSec")
    ).filter(
        pl.col("_AirSec").is_null()
        | ~pl.col("_AirSec").is_between(7200, 21599)
    ).drop("_AirSec")

    log.info("  T_CW_NewPromos_Weekly: %d rows", len(df))
    return df


# ══════════════════════════════════════════════════════════════════════════════
# STEP 4c — T_CW_NewPromos
# ══════════════════════════════════════════════════════════════════════════════

def step4c_t_cw_newpromos(
    weekly_df: pl.DataFrame,
    dp: pl.DataFrame,
    market_list_cw: pl.DataFrame,
    group_owners: pl.DataFrame,
    weekbreak: pl.DataFrame,
) -> pl.DataFrame:
    """
    Replicates Access query (c):
      - LEFT JOIN Market_List_CW   on MarketCode   (TimeZone)
      - LEFT JOIN Group_Owners     on Station = OnO_Stations
      - LEFT JOIN Weekbreak        on AirDay = Day
      - DLookUp Daypart
      - Derived: GroupOwners, TEST, Rk_Mkt, Designation, Avg_Day
    """
    log.info("=== STEP 4c: Building T_CW_NewPromos ===")

    df = weekly_df.clone()

    for col in ["MktRank", "AirCount", "ActRatings", "ActImpression", "DurationCml"]:
        df = df.with_columns(pl.col(col).cast(pl.Float64, strict=False))

    # TimeZone from Market_List_CW
    if "TimeZone" in market_list_cw.columns:
        ml = market_list_cw.select(
            pl.col("MarketCode").str.strip_chars(),
            pl.col("TimeZone").str.strip_chars(),
        ).unique(subset=["MarketCode"])
        df = df.join(ml, on="MarketCode", how="left")

    if "TimeZone" not in df.columns:
        df = df.with_columns(pl.lit("").alias("TimeZone"))

    # Group_Owners
    go = group_owners.select(
        pl.col("OnO_Stations").str.strip_chars(),
        pl.col("Group_Owners").str.strip_chars(),
    ).unique(subset=["OnO_Stations"])
    df = df.join(go, left_on="Station", right_on="OnO_Stations", how="left")

    # Weekbreak on AirDay = Day
    wb = weekbreak.select(
        pl.col("Day").str.strip_chars(),
        pl.col("Weekbreak").str.strip_chars(),
        pl.col("DayNum").str.strip_chars(),
    ).unique(subset=["Day"])
    df = df.join(wb, left_on="AirDay", right_on="Day", how="left")

    # Daypart lookup
    df = _apply_daypart(df, dp)

    # Derived columns
    df = df.with_columns([
        pl.when(pl.col("Group_Owners").is_null())
          .then(pl.lit("Unassigned"))
          .otherwise(pl.lit(""))
          .alias("TEST"),
    ]).with_columns(
        (pl.col("TEST") + pl.col("Group_Owners").fill_null("")).alias("GroupOwners"),
        (pl.col("MktRank").fill_null(0).cast(pl.Int64).cast(pl.Utf8)
         + pl.lit("-") + pl.col("Market").fill_null("")).alias("Rk_Mkt"),
        # Designation = Mid(AirISCI, 12, 1)  →  0-indexed char at position 11
        pl.col("AirISCI").map_elements(
            lambda x: str(x)[11] if x and len(str(x)) > 11 else "",
            return_dtype=pl.Utf8,
        ).alias("Designation"),
        (pl.col("AirCount").fill_null(0) / 7).alias("Avg_Day"),
    )

    log.info("  T_CW_NewPromos: %d rows", len(df))
    return df


# ══════════════════════════════════════════════════════════════════════════════
# STEP 4d — T_CW_Station_NewPromos_Weekly
# ══════════════════════════════════════════════════════════════════════════════

def step4d_t_cw_station_newpromos_weekly(
    newpromos: pl.DataFrame,
    stations: pl.DataFrame,
    product_translation: pl.DataFrame | None = None,
) -> pl.DataFrame:
    """
    Replicates Access query (d):
      T_CW_Stations_COPY LEFT JOIN T_CW_NewPromos on CW_Station = Station
      + zero-fill columns + TRP / TIMP / Length-in-Min calculations.
    """
    log.info("=== STEP 4d: Building T_CW_Station_NewPromos_Weekly ===")

    st = stations.select([
        "CW_Rank", "CW_Market", "CW_Station",
        "CW_Network_Affiliation", "CW_Owner",
        "MKT_Type", "STA_Type",
    ])

    df = st.join(newpromos, left_on="CW_Station", right_on="Station", how="left")

    # LEFT JOIN Product_Translation if available (adds Translation column)
    if product_translation is not None and "Product" in df.columns:
        pt = product_translation.select(
            pl.col("Product").str.strip_chars(),
            pl.col("Translation").str.strip_chars(),
        ).unique(subset=["Product"])
        df = df.join(pt, on="Product", how="left")
    elif "Translation" not in df.columns:
        df = df.with_columns(pl.lit("").alias("Translation"))

    # Numeric casts
    for col in ["DurationCml", "AirCount", "ActRatings", "ActImpression"]:
        if col in df.columns:
            df = df.with_columns(pl.col(col).cast(pl.Float64, strict=False).fill_null(0))

    day_num = (
        pl.col("DayNum").cast(pl.Int64, strict=False).fill_null(0)
        if "DayNum" in df.columns
        else pl.lit(0)
    )
    df = df.with_columns(day_num.alias("DayNum"))

    # Zero-fill placeholder columns
    zero_cols = [
        "DAY_HH", "AD25-54EST_RTG_", "HHEST_RTG_", "AD25-54EST_IMP",
        "HH_UE", "HHEST_IMP", "AD25-54UE", "HHEST_RTG", "AD25-54EST_RTG",
        "HH_RTG_F", "AD25-54_RTG_F",
    ]
    df = df.with_columns([pl.lit(0.0).alias(c) for c in zero_cols])

    # Calculated columns
    df = df.with_columns([
        pl.when(pl.col("DayNum") == 1)
          .then(pl.col("DurationCml"))
          .otherwise(0.0)
          .alias("M-F length"),
        (pl.col("ActImpression") * 1.0).alias("TIMP"),
        (pl.col("DurationCml") / 60.0).alias("Length in Min"),
        (pl.col("ActRatings") * 1.0).alias("TRP"),
    ])

    # Select only the final output columns matching Sheet2 layout
    output_cols = [
        "AirTime", "CW_Rank", "CW_Market", "CW_Station",
        "CW_Network_Affiliation", "CW_Owner", "Agency", "Product",
        "ActRatings", "demo", "AirProgram", "ActImpression",
        "AirISCI", "ISCI/ADID Title", "Daypart", "DurationCml",
        "AirDate", "AirCount", "RtgSource", "WeekOf",
        "Designation", "AirDay", "Translation", "Weekbreak", "DayNum",
        "DemoOrder", "DAY_HH", "MKT_Type", "STA_Type",
        "M-F length", "AD25-54EST_RTG_", "HHEST_RTG_",
        "AD25-54EST_IMP", "HH_UE", "HHEST_IMP", "AD25-54UE",
        "HHEST_RTG", "AD25-54EST_RTG", "HH_RTG_F", "AD25-54_RTG_F",
        "TIMP", "Length in Min", "TRP",
    ]
    # Only select columns that exist in the dataframe
    available_cols = [c for c in output_cols if c in df.columns]
    df = df.select(available_cols)

    # Rename 'demo' → 'DEMO' to match expected output header
    if "demo" in df.columns:
        df = df.rename({"demo": "DEMO"})

    log.info("  T_CW_Station_NewPromos_Weekly: %d rows, %d columns", len(df), len(df.columns))
    return df


# ══════════════════════════════════════════════════════════════════════════════
# STEP 5 — Write Excel
# ══════════════════════════════════════════════════════════════════════════════

def step5_write_csv(final_df: pl.DataFrame, output_path: str) -> None:
    """Write T_CW_Station_NewPromos_Weekly as CSV (raw data for dashboard template)."""
    log.info("=== STEP 5: Writing output CSV ===")

    # Format AirDate as date-only string (no time component)
    if "AirDate" in final_df.columns:
        final_df = final_df.with_columns(
            pl.col("AirDate").cast(pl.Date, strict=False)
              .dt.strftime("%m/%d/%Y")
              .alias("AirDate")
        )

    final_df.write_csv(output_path)
    log.info("  Saved: %s  (%d data rows)", output_path, len(final_df))


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="CW Weekly Report Automation (Polars edition)"
    )
    p.add_argument("--start", required=True,
                   help="Report week start date  YYYY-MM-DD")
    p.add_argument("--end",   required=True,
                   help="Report week end date    YYYY-MM-DD")
    p.add_argument("--data-feed-dir", default=DATA_FEED_DIR,
                   help=f"Folder with S_All_6_*.txt files  [default: {DATA_FEED_DIR}]")
    p.add_argument("--output", default=None,
                   help="Output CSV path (auto-generated from dates if omitted)")
    p.add_argument("--download", action="store_true",
                   help="Download latest S_All_6 files from SFTP before processing")
    return p.parse_args()


def main() -> None:
    args       = parse_args()
    start_date = datetime.strptime(args.start, "%Y-%m-%d")
    end_date   = datetime.strptime(args.end,   "%Y-%m-%d")
    feed_dir   = args.data_feed_dir

    # Build date-stamped output filename: CW_Weekof_OUTPUT_MM.DD-DD.csv
    if args.output:
        output_path = args.output
    else:
        start_mm = start_date.strftime("%m")
        start_dd = start_date.strftime("%d")
        end_dd   = end_date.strftime("%d")
        output_name = f"CW_Weekof_OUTPUT_{start_mm}.{start_dd}-{end_dd}.csv"
        output_path = os.path.join(BASE_DIR, output_name)

    log.info("CW Automation (Polars)  |  Week: %s → %s", args.start, args.end)
    log.info("Data feed dir : %s", feed_dir)
    log.info("Output file   : %s", output_path)

    # ── Step 0: Download from SFTP (if requested) ────────────────────────────
    if args.download:
        step0_download_feeds(feed_dir)

    # ── Load reference tables ────────────────────────────────────────────────
    log.info("Loading reference tables ...")
    dp           = _load_daypart_table(DAYPART_CSV)
    group_owners = pl.read_csv(GROUP_OWNERS_CSV, infer_schema=False)
    group_owners = group_owners.rename({c: c.strip() for c in group_owners.columns})
    market_list  = pl.read_csv(MARKET_LIST_CSV,  infer_schema=False)
    market_list  = market_list.rename({c: c.strip() for c in market_list.columns})
    stations     = pl.read_csv(STATIONS_CSV,      infer_schema=False)
    stations     = stations.rename({c: c.strip() for c in stations.columns})
    weekbreak    = pl.read_csv(WEEKBREAK_CSV,     infer_schema=False)
    weekbreak    = weekbreak.rename({c: c.strip() for c in weekbreak.columns})

    # Optional: Product Translation table
    product_translation = None
    if os.path.exists(PRODUCT_TRANSLATION_CSV):
        product_translation = pl.read_csv(PRODUCT_TRANSLATION_CSV, infer_schema=False)
        product_translation = product_translation.rename(
            {c: c.strip() for c in product_translation.columns}
        )
        log.info("  Loaded product translation table: %d rows", len(product_translation))
    else:
        log.info("  Product translation file not found — skipping (Translation column will be blank)")

    # ── Step 1-3: Read S_All_6 files, rename columns, union ──────────────────
    union_df = step1_3_read_and_union_raw_feeds(feed_dir)

    # ── Step 4a ──────────────────────────────────────────────────────────────
    _t_all_raw1 = step4a_t_all_raw1(
        union_df, start_date, end_date, dp, group_owners, market_list
    )

    # ── Step 4b ──────────────────────────────────────────────────────────────
    t_weekly = step4b_t_cw_newpromos_weekly(union_df, start_date, end_date)

    # ── Step 4c ──────────────────────────────────────────────────────────────
    t_newpromos = step4c_t_cw_newpromos(
        t_weekly, dp, market_list, group_owners, weekbreak
    )

    # ── Step 4d ──────────────────────────────────────────────────────────────
    final_df = step4d_t_cw_station_newpromos_weekly(
        t_newpromos, stations, product_translation
    )

    # ── Step 5 ───────────────────────────────────────────────────────────────
    step5_write_csv(final_df, output_path)

    log.info("✓ Done.  Output: %s", output_path)


if __name__ == "__main__":
    main()
