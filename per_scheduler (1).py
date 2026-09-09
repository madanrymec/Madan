"""
Purchase Registry Scheduler
============================
Fetches the current-month purchase register from the SAP OData service
(exported as .xlsx), transforms it, and refreshes it into the MySQL table
`purchase_registry` using a DELETE-then-INSERT strategy scoped to the
current calendar month (based on Post_Date).

Design notes / assumptions (read before deploying)
----------------------------------------------------
1. COLUMN ORDER, NOT COLUMN NAME, IS THE SOURCE OF TRUTH.
   The source Excel has duplicate header labels (e.g. "Amount" appears
   ~20 times, "Currency" twice, "Fiscal Year" twice). There is no reliable
   way to map by header text. Instead, COLUMN_MAP below encodes the exact
   left-to-right order of columns as given in your mapping sheet, and we
   read the workbook positionally (by column index), ignoring header text
   entirely. If the SAP report layout is ever changed (a column added,
   removed, or reordered) this mapping MUST be updated, or data will be
   silently misplaced. A column-count guard (EXPECTED_COLUMN_COUNT) aborts
   the run instead of inserting misaligned data if the shape changes.

2. Financial year fix: SAP gives "2026-2026"; we normalise to "2026-2027"
   (end_year = start_year + 1) before insert.

3. Trailing-minus decimals: SAP text-exports negatives as "1234.56-".
   parse_sap_decimal() handles trailing minus, leading minus, thousands
   separators, blanks, and SAP's synonym for zero ("*").

4. Delete/insert scope: every run computes the current calendar month
   window from Post_Date ([first_of_month, first_of_next_month)) at
   execution time -- not a fixed/hardcoded month. Once a month rolls
   over, the window shifts automatically, so prior months already
   committed are never touched -> "month end preservation" is automatic,
   no special month-end code path needed.

5. Safety: delete + insert happen inside a single DB transaction so a
   failed insert rolls back the delete (no data loss window). Deadlock
   handling retries the whole transaction a few times on MySQL error 1213
   / 1205. A file-based + in-process lock stops overlapping runs if a
   previous cycle is still running when the next 4-hour tick fires.
"""

import io
import os
import sys
import time
import logging
import threading
from datetime import datetime, date
from decimal import Decimal, InvalidOperation
from logging.handlers import RotatingFileHandler

import requests
import pandas as pd
import pymysql
import pymysql.cursors
from pymysql.err import Error as MySQLError
from dbutils.pooled_db import PooledDB
from flask import Flask, jsonify, request
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.interval import IntervalTrigger

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------

API_URL = os.environ.get(
    "PURREG_API_URL",
    "https://kecapi.kirloskar-electric.com/sap/opu/odata/sap/ZPURREG_SRV/ZPURREGSet?$format=xlsx",
)
API_USERNAME = os.environ.get("PURREG_API_USERNAME", "corpodata")
API_PASSWORD = os.environ.get("PURREG_API_PASSWORD", "Kec12345")
API_TIMEOUT_SECONDS = int(os.environ.get("PURREG_API_TIMEOUT", "180"))
API_MAX_RETRIES = int(os.environ.get("PURREG_API_MAX_RETRIES", "3"))
API_RETRY_BACKOFF_SECONDS = int(os.environ.get("PURREG_API_RETRY_BACKOFF", "15"))

DB_CONFIG = {
    "host": os.environ.get("DB_HOST", "localhost"),
    "user": os.environ.get("DB_USER", "root"),
    "password": os.environ.get("DB_PASSWORD", ""),
    "database": os.environ.get("DB_NAME", "dashboard"),
}
TABLE_NAME = os.environ.get("DB_TABLE", "purchase_registry")

APP_HOST = os.environ.get("APP_HOST", "192.7.200.48")
APP_PORT = int(os.environ.get("APP_PORT", "5100"))
JOB_INTERVAL_HOURS = int(os.environ.get("JOB_INTERVAL_HOURS", "4"))

INSERT_BATCH_SIZE = int(os.environ.get("INSERT_BATCH_SIZE", "1000"))
DB_TX_MAX_RETRIES = int(os.environ.get("DB_TX_MAX_RETRIES", "3"))
DB_TX_RETRY_DELAY_SECONDS = 3

LOG_DIR = os.environ.get("PURREG_LOG_DIR", os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs"))
os.makedirs(LOG_DIR, exist_ok=True)

# --------------------------------------------------------------------------
# LOGGING
# --------------------------------------------------------------------------

logger = logging.getLogger("purchase_scheduler")
logger.setLevel(logging.INFO)

_file_handler = RotatingFileHandler(
    os.path.join(LOG_DIR, "purchase_scheduler.log"), maxBytes=10 * 1024 * 1024, backupCount=10
)
_console_handler = logging.StreamHandler(sys.stdout)
_fmt = logging.Formatter("%(asctime)s | %(levelname)-8s | %(message)s")
_file_handler.setFormatter(_fmt)
_console_handler.setFormatter(_fmt)
logger.addHandler(_file_handler)
logger.addHandler(_console_handler)

# --------------------------------------------------------------------------
# COLUMN MAPPING (positional, excel order == this list's order)
# type codes: str | text | date | dec3 | dec2 | fin_year
# --------------------------------------------------------------------------

COLUMN_MAP = [
    ("As_on_Date", "date"),
    ("po_number", "str"),
    ("item", "str"),
    ("po_type", "str"),
    ("Currency", "str"),
    ("IR_No", "str"),
    ("GR_No", "str"),
    ("Year", "str"),
    ("T", "str"),
    ("Event", "str"),
    ("St", "str"),
    ("Desc_Status", "str"),
    ("Desc_Block", "str"),
    ("Doc_Date", "date"),
    ("Post_Date", "date"),
    ("Ven_Inv_No", "str"),
    ("Vendor_cd", "str"),
    ("Vendor_Name", "str"),
    ("Payment_terms_Desc", "str"),
    ("Inco_terms_Desc", "str"),
    ("House_No_Street", "str"),
    ("Vat_Reg_No", "str"),
    ("Crcy", "str"),
    ("BusA", "str"),
    ("ValA", "str"),
    ("HeaderText", "text"),
    ("Revwith", "str"),
    ("DocumentNo", "str"),
    ("Year1", "str"),
    ("Material", "str"),
    ("Short_Text", "text"),
    ("Mat_Group", "str"),
    ("Val_Type", "str"),
    ("Tx", "str"),
    ("Description", "str"),
    ("Quantity", "dec3"),
    ("OUn", "str"),
    ("ValCl1", "str"),
    ("ValCl2", "str"),
    ("Amount", "dec2"),
    ("Base_Amount", "dec2"),
    ("Custom_Duty", "dec2"),
    ("Cess_Custom", "dec2"),
    ("SP_Custom", "dec2"),
    ("CVD_Amount", "dec2"),
    ("Cess_CVD", "dec2"),
    ("Others", "dec2"),
    ("Add_Duty", "dec2"),
    ("Land_Charges", "dec2"),
    ("Freight", "dec2"),
    ("Excise", "dec2"),
    ("Ed_Cess", "dec2"),
    ("C_S_T", "dec2"),
    ("VAT", "dec2"),
    ("VAT_ptg", "dec2"),
    ("H_S_CESS", "dec2"),
    ("OTHERS_2", "dec2"),
    ("Del_costs", "dec2"),
    ("GL_Amount", "dec2"),
    ("G_L_Acct", "str"),
    ("Price_Diff", "dec2"),
    ("W_tax_amnt", "dec2"),
    ("Total", "dec2"),
    ("fin_Year", "fin_year"),
    ("Month_year", "str"),
]

EXPECTED_COLUMN_COUNT = len(COLUMN_MAP)
SQL_COLUMNS = [c[0] for c in COLUMN_MAP]

# varchar(N) limits taken from the CREATE TABLE, used to defensively
# truncate (with a warning) instead of crashing the whole batch on a
# "Data too long for column" error.
MAX_LENGTHS = {
    "po_number": 30, "item": 10, "po_type": 10, "Currency": 10, "IR_No": 20,
    "GR_No": 20, "Year": 20, "T": 20, "Event": 20, "St": 20,
    "Desc_Status": 50, "Desc_Block": 50, "Ven_Inv_No": 55, "Vendor_cd": 20,
    "Vendor_Name": 100, "Payment_terms_Desc": 100, "Inco_terms_Desc": 100,
    "House_No_Street": 200, "Vat_Reg_No": 100, "Crcy": 10, "BusA": 50,
    "ValA": 10, "Revwith": 50, "DocumentNo": 50, "Year1": 10, "Material": 100,
    "Mat_Group": 20, "Val_Type": 20, "Tx": 10, "Description": 100, "OUn": 10,
    "ValCl1": 200, "ValCl2": 200, "G_L_Acct": 10, "fin_Year": 20,
    "Month_year": 30,
}

# --------------------------------------------------------------------------
# DB CONNECTION POOL
# --------------------------------------------------------------------------

_pool = PooledDB(
    creator=pymysql,
    maxconnections=5,
    mincached=1,
    blocking=True,
    ping=1,               # ping=1 -> check/reconnect stale connections on each checkout
    autocommit=False,
    host=DB_CONFIG["host"],
    user=DB_CONFIG["user"],
    password=DB_CONFIG["password"],
    database=DB_CONFIG["database"],
    charset="utf8mb4",
    cursorclass=pymysql.cursors.Cursor,
)


def get_connection():
    return _pool.connection()


# --------------------------------------------------------------------------
# HELPER / PARSING FUNCTIONS
# --------------------------------------------------------------------------

def _is_blank(value) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    s = str(value).strip()
    return s == "" or s.lower() in ("nan", "none", "nat")


def parse_sap_decimal(value):
    """Handle SAP's trailing-minus negatives ('1234.56-'), leading minus,
    thousands separators, blank/'*' zero markers. Returns Decimal or None."""
    if _is_blank(value):
        return None
    if isinstance(value, (int, float, Decimal)):
        try:
            return Decimal(str(value))
        except InvalidOperation:
            return None

    s = str(value).strip()
    if s in ("*", "-"):
        return None

    negative = False
    if s.endswith("-"):
        negative = True
        s = s[:-1].strip()
    if s.startswith("-"):
        negative = True
        s = s[1:].strip()
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1].strip()

    s = s.replace(",", "").replace(" ", "")
    if s == "":
        return None

    try:
        val = Decimal(s)
    except InvalidOperation:
        logger.warning("Could not parse decimal value %r -> storing NULL", value)
        return None

    return -abs(val) if negative else val


def parse_sap_date(value):
    """Return a python date, or None. Handles native datetime/Timestamp
    (normal case once pandas parses the xlsx cell), plus common SAP
    string date formats as a fallback."""
    if _is_blank(value):
        return None
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.date()

    s = str(value).strip()
    for fmt in ("%d.%m.%Y", "%Y%m%d", "%d-%m-%Y", "%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    try:
        parsed = pd.to_datetime(s, errors="raise")
        return parsed.date()
    except Exception:
        logger.warning("Could not parse date value %r -> storing NULL", value)
        return None


def parse_str(value, column=None):
    if _is_blank(value):
        return None
    s = str(value).strip()
    # pandas turns "0012345" into an int/float if the column looked numeric;
    # guard against the common "123.0" artifact on what should be a code.
    if s.endswith(".0") and s.replace(".0", "").isdigit():
        s = s[:-2]
    max_len = MAX_LENGTHS.get(column)
    if max_len and len(s) > max_len:
        logger.warning("Truncating %s (len %d > %d): %r", column, len(s), max_len, s)
        s = s[:max_len]
    return s


def fix_financial_year(value):
    """'2026-2026' -> '2026-2027'. Leaves already-correct values untouched
    and passes through anything that doesn't match the expected pattern
    (logged as a warning) rather than guessing."""
    if _is_blank(value):
        return None
    s = str(value).strip()
    parts = s.split("-")
    if len(parts) == 2 and parts[0].strip().isdigit() and parts[1].strip().isdigit():
        start = int(parts[0].strip())
        end = int(parts[1].strip())
        if end == start:
            return f"{start}-{start + 1}"
        # already looks like a correct 'YYYY-YYYY+1' style value
        return f"{start}-{end}"
    logger.warning("Unexpected fin_Year format %r -> storing as-is", value)
    return s[:20]


def convert_cell(value, dtype, column):
    if dtype == "date":
        return parse_sap_date(value)
    if dtype in ("dec2", "dec3"):
        return parse_sap_decimal(value)
    if dtype == "fin_year":
        return fix_financial_year(value)
    if dtype == "text":
        return None if _is_blank(value) else str(value).strip()
    # default: str
    return parse_str(value, column)


# --------------------------------------------------------------------------
# FETCH
# --------------------------------------------------------------------------

def fetch_api_dataframe() -> pd.DataFrame:
    """Downloads the xlsx export and returns a raw DataFrame with NO header
    row consumed (header=None) so column access is purely positional."""
    last_exc = None
    for attempt in range(1, API_MAX_RETRIES + 1):
        try:
            logger.info("Fetching purchase register from API (attempt %d/%d)", attempt, API_MAX_RETRIES)
            resp = requests.get(
                API_URL,
                auth=(API_USERNAME, API_PASSWORD),
                timeout=API_TIMEOUT_SECONDS,
            )
            resp.raise_for_status()
            df = pd.read_excel(io.BytesIO(resp.content), header=None, dtype=object)
            # first row is the (duplicate-laden, unreliable) header row - drop it
            df = df.iloc[1:].reset_index(drop=True)
            logger.info("Fetched %d data rows, %d columns", len(df), df.shape[1])
            return df
        except Exception as exc:  # noqa: BLE001 - want to retry on any transient issue
            last_exc = exc
            logger.error("API fetch attempt %d failed: %s", attempt, exc)
            if attempt < API_MAX_RETRIES:
                time.sleep(API_RETRY_BACKOFF_SECONDS)
    raise RuntimeError(f"Failed to fetch API data after {API_MAX_RETRIES} attempts") from last_exc


def transform_dataframe(df: pd.DataFrame):
    """Validates column count and converts every row into a dict keyed by
    SQL column name, using COLUMN_MAP positionally. Rows with a hard parse
    problem are skipped (logged), not silently corrupted."""
    actual_cols = df.shape[1]
    if actual_cols != EXPECTED_COLUMN_COUNT:
        raise ValueError(
            f"Column count mismatch: API returned {actual_cols} columns, "
            f"expected {EXPECTED_COLUMN_COUNT}. Aborting to avoid misaligned "
            f"data -- update COLUMN_MAP if the SAP report layout changed."
        )

    records = []
    skipped = 0
    for _, row in df.iterrows():
        try:
            record = {}
            for idx, (sql_col, dtype) in enumerate(COLUMN_MAP):
                record[sql_col] = convert_cell(row.iloc[idx], dtype, sql_col)
            if record.get("Post_Date") is None:
                # Post_Date drives the delete/insert window and is a core
                # business key -- a row with no posting date can't be
                # placed safely, so it's skipped rather than guessed at.
                skipped += 1
                continue
            records.append(record)
        except Exception as exc:  # noqa: BLE001
            skipped += 1
            logger.warning("Skipping unparsable row: %s", exc)

    logger.info("Transformed %d rows (%d skipped)", len(records), skipped)
    return records


# --------------------------------------------------------------------------
# CURRENT MONTH WINDOW
# --------------------------------------------------------------------------

def current_month_window():
    today = date.today()
    start = date(today.year, today.month, 1)
    if today.month == 12:
        end = date(today.year + 1, 1, 1)
    else:
        end = date(today.year, today.month + 1, 1)
    return start, end


# --------------------------------------------------------------------------
# DB WRITE (delete current-month window, insert fresh rows) - one transaction
# --------------------------------------------------------------------------

INSERT_SQL = (
    f"INSERT INTO `{TABLE_NAME}` (`" + "`,`".join(SQL_COLUMNS) + "`) "
    f"VALUES (" + ",".join(["%s"] * len(SQL_COLUMNS)) + ")"
)

DELETE_SQL = f"DELETE FROM `{TABLE_NAME}` WHERE `Post_Date` >= %s AND `Post_Date` < %s"


def replace_current_month(records):
    if not records:
        logger.info("No records to insert -- skipping DB write entirely (delete also skipped).")
        return {"deleted": 0, "inserted": 0}

    window_start, window_end = current_month_window()
    logger.info("Refreshing window Post_Date in [%s, %s)", window_start, window_end)

    last_exc = None
    for attempt in range(1, DB_TX_MAX_RETRIES + 1):
        conn = get_connection()
        try:
            conn.begin()  # explicit start of transaction (autocommit=False already, but be explicit)
            cursor = conn.cursor()

            cursor.execute(DELETE_SQL, (window_start, window_end))
            deleted = cursor.rowcount
            logger.info("Deleted %d existing rows for the current month window", deleted)

            inserted = 0
            rows_as_tuples = [tuple(r[col] for col in SQL_COLUMNS) for r in records]
            for i in range(0, len(rows_as_tuples), INSERT_BATCH_SIZE):
                batch = rows_as_tuples[i:i + INSERT_BATCH_SIZE]
                cursor.executemany(INSERT_SQL, batch)
                inserted += cursor.rowcount if cursor.rowcount != -1 else len(batch)

            conn.commit()
            cursor.close()
            logger.info("Committed: deleted=%d inserted=%d", deleted, inserted)
            return {"deleted": deleted, "inserted": inserted}

        except MySQLError as exc:
            conn.rollback()
            last_exc = exc
            errno = exc.args[0] if exc.args else None
            # 1213 = deadlock, 1205 = lock wait timeout -> retry the whole tx
            if errno in (1213, 1205) and attempt < DB_TX_MAX_RETRIES:
                logger.warning(
                    "Transient DB error (errno %s) on attempt %d/%d, retrying: %s",
                    errno, attempt, DB_TX_MAX_RETRIES, exc,
                )
                time.sleep(DB_TX_RETRY_DELAY_SECONDS)
                continue
            logger.error("DB write failed (errno %s): %s", errno, exc)
            raise
        finally:
            conn.close()

    raise RuntimeError(f"DB transaction failed after {DB_TX_MAX_RETRIES} attempts") from last_exc


# --------------------------------------------------------------------------
# JOB ORCHESTRATION + OVERLAP LOCK
# --------------------------------------------------------------------------

_job_lock = threading.Lock()
_last_run_info = {"status": "never_run"}


def run_sync_job():
    if not _job_lock.acquire(blocking=False):
        logger.warning("Previous sync still running -- skipping this scheduled tick.")
        return {"status": "skipped_overlap"}

    started_at = datetime.now()
    try:
        logger.info("=== Sync job started at %s ===", started_at)
        df = fetch_api_dataframe()
        records = transform_dataframe(df)
        result = replace_current_month(records)
        _last_run_info.update({
            "status": "success",
            "started_at": started_at.isoformat(),
            "finished_at": datetime.now().isoformat(),
            "fetched_rows": len(df),
            **result,
        })
        logger.info("=== Sync job finished successfully ===")
        return _last_run_info
    except Exception as exc:  # noqa: BLE001
        logger.exception("Sync job failed: %s", exc)
        _last_run_info.update({
            "status": "error",
            "started_at": started_at.isoformat(),
            "finished_at": datetime.now().isoformat(),
            "error": str(exc),
        })
        return _last_run_info
    finally:
        _job_lock.release()


# --------------------------------------------------------------------------
# FLASK APP + SCHEDULER
# --------------------------------------------------------------------------

app = Flask(__name__)
scheduler = BackgroundScheduler(timezone="Asia/Kolkata")


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "last_run": _last_run_info})


@app.route("/trigger", methods=["POST"])
def trigger():
    """Manual on-demand run. Runs inline (blocking) and returns the result."""
    result = run_sync_job()
    code = 200 if result.get("status") == "success" else 500
    return jsonify(result), code


def start_scheduler():
    scheduler.add_job(
        run_sync_job,
        trigger=IntervalTrigger(hours=JOB_INTERVAL_HOURS),
        id="purchase_registry_sync",
        max_instances=1,
        coalesce=True,
        misfire_grace_time=3600,
        next_run_time=datetime.now(),  # run once immediately on startup
    )
    scheduler.start()
    logger.info("Scheduler started: every %d hour(s)", JOB_INTERVAL_HOURS)


if __name__ == "__main__":
    start_scheduler()
    try:
        # For production use a WSGI server (waitress/gunicorn) instead of
        # the Flask dev server -- see run_waitress.py.
        app.run(host=APP_HOST, port=APP_PORT, debug=False, use_reloader=False)
    finally:
        scheduler.shutdown(wait=False)
