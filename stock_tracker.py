"""
Automated NSE Stock Price Tracker
----------------------------------
Tracks:
  - JSW Steel (JSWSTEEL.NS)   - Quantity: 152
  - Coal India (COALINDIA.NS) - Quantity: 239
  - Eternal (ETERNAL.NS)       - Quantity: 313

Fetches latest trading session closing prices, updates Sheet 1 (Tracker)
and updates Sheet 4 (Dashboard) while safely preserving formatting,
formulas, charts, and historical data without duplicate dates.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import shutil
import socket
import sys
import urllib.parse
import urllib.request
from copy import copy
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

import warnings
import openpyxl

# Suppress harmless openpyxl extension warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

try:
    import yfinance as yf
except ImportError:
    yf = None

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("StockTracker")

# Stock configuration
POSITIONS = {
    "JSWSTEEL.NS": {
        "name": "JSW Steel",
        "price_col": "B",
        "value_col": "C",
        "quantity": 152,
    },
    "COALINDIA.NS": {
        "name": "Coal India",
        "price_col": "D",
        "value_col": "E",
        "quantity": 239,
    },
    "ETERNAL.NS": {
        "name": "Eternal",
        "price_col": "F",
        "value_col": "G",
        "quantity": 313,
    },
}

DEFAULT_WORKBOOK_NAME = "automated_stock_price_tracker.xlsx"


def check_internet_connection(host: str = "query1.finance.yahoo.com", port: int = 443, timeout: float = 3.0) -> bool:
    """Quick check for network reachability."""
    try:
        socket.setdefaulttimeout(timeout)
        socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect((host, port))
        return True
    except (socket.error, OSError):
        return False


def fetch_prices_yfinance(symbol: str) -> Dict[date, float]:
    """Fetch daily closing prices for a symbol using yfinance."""
    if yf is None:
        raise RuntimeError("yfinance package is not installed.")
    
    ticker = yf.Ticker(symbol)
    df = ticker.history(period="1mo")
    if df is None or df.empty:
        raise RuntimeError(f"No market data returned by yfinance for '{symbol}'.")

    prices: Dict[date, float] = {}
    for idx, row in df.iterrows():
        close_val = row.get("Close")
        if close_val is not None and not (isinstance(close_val, float) and (close_val != close_val)):
            session_date = idx.date() if hasattr(idx, "date") else idx
            prices[session_date] = float(close_val)

    if not prices:
        raise RuntimeError(f"No valid closing prices found in yfinance history for '{symbol}'.")
    return prices


def fetch_prices_direct(symbol: str, range_: str = "1mo") -> Dict[date, float]:
    """Direct HTTP fallback to Yahoo Finance public chart API with timezone offset."""
    url = (
        "https://query1.finance.yahoo.com/v8/finance/chart/"
        + urllib.parse.quote(symbol, safe="")
        + f"?range={range_}&interval=1d&events=history"
    )
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) StockTracker/1.0"},
    )
    with urllib.request.urlopen(req, timeout=20) as resp:
        data = json.load(resp)

    result = data.get("chart", {}).get("result")
    if not result:
        error_info = data.get("chart", {}).get("error", {})
        raise RuntimeError(f"Yahoo Finance API error for '{symbol}': {error_info}")

    meta = result[0].get("meta", {})
    gmtoffset = meta.get("gmtoffset", 19800)  # Default to IST (UTC +5:30 = 19800 sec)
    timestamps = result[0].get("timestamp", [])
    quotes = result[0].get("indicators", {}).get("quote", [{}])[0]
    closes = quotes.get("close", [])

    out: Dict[date, float] = {}
    for ts, close in zip(timestamps, closes):
        if close is not None:
            session_date = datetime.fromtimestamp(ts + gmtoffset, tz=timezone.utc).date()
            out[session_date] = float(close)

    if not out:
        raise RuntimeError(f"Direct API returned empty closing prices for '{symbol}'.")
    return out



import time
import random
try:
    import pandas as pd
    from jugaad_data.nse import stock_df
except ImportError:
    pass

def fetch_symbol_history(symbol: str) -> Dict[date, float]:
    """Fetch history using jugaad-data with fallback to yfinance."""
    clean_symbol = symbol.replace(".NS", "")
    max_retries = 3
    
    # Attempt jugaad-data first
    for attempt in range(max_retries):
        try:
            from datetime import timedelta
            to_d = datetime.now().date()
            from_d = to_d - timedelta(days=30)
            df = stock_df(symbol=clean_symbol, from_date=from_d, to_date=to_d, series="EQ")
            if not df.empty:
                df['DATE'] = pd.to_datetime(df['DATE']) + pd.Timedelta(hours=5, minutes=30)
                prices = {}
                for idx, row in df.iterrows():
                    d = row['DATE'].date()
                    prices[d] = float(row['CLOSE'])
                if prices:
                    return prices
        except Exception as exc:
            pass
        time.sleep(random.uniform(2, 5))
            
    logger.warning(f"jugaad-data fetch failed for {symbol}. Trying yfinance fallback...")
    
    # Attempt yfinance first
    if yf is not None:
        try:
            return fetch_prices_yfinance(symbol)
        except Exception as exc:
            logger.warning(f"yfinance fetch failed for {symbol}: {exc}. Trying direct API fallback...")

    # Fallback to direct HTTP
    try:
        return fetch_prices_direct(symbol)
    except Exception as exc:
        raise RuntimeError(f"Unable to fetch market data for {symbol}: {exc}") from exc


def normalize_date(value: Any) -> Optional[date]:
    """Parse various Excel date representations into a standard date object."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    if "T" in text:
        text = text.split("T", 1)[0]
    for fmt in ("%Y-%m-%d", "%d-%b-%y", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(text[:10] if fmt in ("%Y-%m-%d", "%Y/%m/%d") else text, fmt).date()
        except ValueError:
            pass
    return None


def find_row_for_date(tracker_ws: openpyxl.worksheet.worksheet.Worksheet, target_date: date) -> Tuple[Optional[int], int]:
    """
    Search Sheet 1 - Tracker for an existing row matching target_date.
    Returns:
        (existing_row_index_or_None, last_populated_data_row_index)
    """
    existing_row = None
    last_row = 1  # Header is at row 1

    for row_idx in range(2, 1001):
        cell_val = tracker_ws.cell(row=row_idx, column=1).value
        if cell_val in (None, ""):
            continue
        last_row = row_idx
        row_date = normalize_date(cell_val)
        if row_date == target_date:
            existing_row = row_idx

    return existing_row, last_row


def copy_row_formatting(ws: openpyxl.worksheet.worksheet.Worksheet, src_row: int, dst_row: int, max_col: int = 7):
    """Safely copy cell styles (font, border, fill, number_format, alignment) from src_row to dst_row."""
    for col_idx in range(1, max_col + 1):
        src_cell = ws.cell(row=src_row, column=col_idx)
        dst_cell = ws.cell(row=dst_row, column=col_idx)
        if src_cell.has_style:
            dst_cell.font = copy(src_cell.font)
            dst_cell.border = copy(src_cell.border)
            dst_cell.fill = copy(src_cell.fill)
            dst_cell.number_format = copy(src_cell.number_format)
            dst_cell.protection = copy(src_cell.protection)
            dst_cell.alignment = copy(src_cell.alignment)


def update_workbook(
    workbook_path: Path,
    target_date: date,
    prices: Dict[str, float],
    dry_run: bool = False,
) -> Tuple[int, bool, Dict[str, float]]:
    """
    Update the workbook with latest stock prices and formulas.
    Returns:
        (row_num, is_new_entry, calculated_values)
    """
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook file does not exist: {workbook_path}")

    # Load workbook preserving formulas and charts
    wb = openpyxl.load_workbook(workbook_path, data_only=False)

    if "Sheet 1 - Tracker" not in wb.sheetnames:
        raise ValueError("Workbook missing required sheet: 'Sheet 1 - Tracker'")
    if "Sheet 4 - Dashboard" not in wb.sheetnames:
        raise ValueError("Workbook missing required sheet: 'Sheet 4 - Dashboard'")

    tracker = wb["Sheet 1 - Tracker"]
    dashboard = wb["Sheet 4 - Dashboard"]

    existing_row, last_row = find_row_for_date(tracker, target_date)
    is_new = existing_row is None
    target_row = (last_row + 1) if is_new else existing_row

    if is_new:
        logger.info(f"Adding new entry for date {target_date.isoformat()} at row {target_row}.")
        # Copy formatting from preceding data row to preserve style
        copy_row_formatting(tracker, src_row=last_row, dst_row=target_row, max_col=7)
    else:
        logger.info(f"Found existing entry for date {target_date.isoformat()} at row {target_row}. Updating prices...")

    # Set Date in Column A as a datetime object
    dt_val = datetime.combine(target_date, datetime.min.time())
    tracker.cell(row=target_row, column=1, value=dt_val)

    # Calculate portfolio values
    calculated_values: Dict[str, float] = {}
    total_val = 0.0

    for symbol, conf in POSITIONS.items():
        price = round(float(prices[symbol]), 2)
        qty = conf["quantity"]
        val = round(price * qty, 2)
        calculated_values[symbol] = val
        total_val += val

        price_col = conf["price_col"]
        val_col = conf["value_col"]

        # Map column letter to 1-based index
        col_price_idx = openpyxl.utils.column_index_from_string(price_col)
        col_val_idx = openpyxl.utils.column_index_from_string(val_col)

        # Set Price
        tracker.cell(row=target_row, column=col_price_idx, value=price)
        # Set Formula e.g. =B13*152
        formula_str = f"={price_col}{target_row}*{qty}"
        tracker.cell(row=target_row, column=col_val_idx, value=formula_str)

    calculated_values["TOTAL"] = round(total_val, 2)

    # Expand Excel Table range if table exists
    if "TrackerTable" in tracker.tables:
        table = tracker.tables["TrackerTable"]
        max_table_row = max(last_row, target_row)
        new_table_ref = f"A1:G{max_table_row}"
        table.ref = new_table_ref
        logger.info(f"Updated TrackerTable range to {new_table_ref}")

    # Update Dashboard Line Chart ranges to include the new row
    if getattr(dashboard, "_charts", None):
        chart = dashboard._charts[0]
        max_data_row = max(last_row, target_row)
        for idx, (sym, conf) in enumerate(POSITIONS.items()):
            val_col = conf["value_col"]
            if idx < len(chart.series):
                s = chart.series[idx]
                if getattr(s, "val", None) and getattr(s.val, "numRef", None):
                    s.val.numRef.f = f"'Sheet 1 - Tracker'!${val_col}$2:${val_col}${max_data_row}"
                if getattr(s, "cat", None) and getattr(s.cat, "strRef", None):
                    s.cat.strRef.f = f"'Sheet 1 - Tracker'!$A$2:$A${max_data_row}"
        logger.info(f"Updated Dashboard trend chart range to row {max_data_row}")

    if dry_run:
        logger.info("[DRY RUN] Workbook was not saved to disk.")
        return target_row, is_new, calculated_values

    # Create safety backup before saving
    backup_dir = workbook_path.parent / "backups"
    backup_dir.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = backup_dir / f"{workbook_path.stem}_{timestamp}.xlsx"
    shutil.copy2(workbook_path, backup_path)
    logger.info(f"Safety backup created at: {backup_path.name}")

    # Save workbook atomically using temporary file
    temp_save_path = workbook_path.with_name(f".tmp_{workbook_path.name}")
    wb.save(temp_save_path)
    wb.close()

    # Replace original file with temporary file
    shutil.move(temp_save_path, workbook_path)
    logger.info(f"Successfully saved updated workbook: {workbook_path.name}")

    # Synchronize interactive HTML dashboard
    update_html_dashboard(workbook_path)

    return target_row, is_new, calculated_values


def update_html_dashboard(workbook_path: Path) -> Optional[Path]:
    """Synchronize portfolio_dashboard.html with the latest Excel tracker data."""
    try:
        wb = openpyxl.load_workbook(workbook_path, data_only=True)
        if "Sheet 1 - Tracker" not in wb.sheetnames:
            return None
        ws = wb["Sheet 1 - Tracker"]
        sessions = []
        for r in range(2, ws.max_row + 1):
            d = ws.cell(r, 1).value
            if not d:
                continue
            d_str = d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else str(d)[:10]
            jsw_c = float(ws.cell(r, 2).value or 0)
            jsw_v = float(ws.cell(r, 3).value or round(jsw_c * 152, 2))
            coal_c = float(ws.cell(r, 4).value or 0)
            coal_v = float(ws.cell(r, 5).value or round(coal_c * 239, 2))
            et_c = float(ws.cell(r, 6).value or 0)
            et_v = float(ws.cell(r, 7).value or round(et_c * 313, 2))
            tot = round(jsw_v + coal_v + et_v, 2)
            sessions.append({
                "date": d_str,
                "jsw_close": round(jsw_c, 2),
                "jsw_val": round(jsw_v, 2),
                "coal_close": round(coal_c, 2),
                "coal_val": round(coal_v, 2),
                "eternal_close": round(et_c, 2),
                "eternal_val": round(et_v, 2),
                "total_val": tot,
            })

        html_path = workbook_path.with_name("portfolio_dashboard.html")
        if html_path.exists() and sessions:
            import re
            content = html_path.read_text(encoding="utf-8")
            new_sessions_js = f"const SESSIONS = {json.dumps(sessions, indent=6)};"
            updated_content = re.sub(r"const SESSIONS = \[[\s\S]*?\];", new_sessions_js, content)
            html_path.write_text(updated_content, encoding="utf-8")
            logger.info(f"Updated HTML dashboard: {html_path.name}")

            # Also sync to Downloads directory if separate
            downloads_dir = Path.home() / "Downloads"
            if downloads_dir.exists() and downloads_dir.resolve() != html_path.parent.resolve():
                shutil.copy2(html_path, downloads_dir / html_path.name)
            return html_path
    except Exception as exc:
        logger.warning(f"Could not update HTML dashboard: {exc}")
    return None


def run_tracker(workbook_path_str: Optional[str] = None, dry_run: bool = False) -> int:
    """Main execution routine."""
    logger.info("=" * 65)
    logger.info("Starting Automated NSE Stock Price Tracker")
    logger.info("=" * 65)

    # 1. Determine workbook path
    script_dir = Path(__file__).resolve().parent
    if workbook_path_str:
        wb_path = Path(workbook_path_str).resolve()
    else:
        wb_path = script_dir / DEFAULT_WORKBOOK_NAME

    if not wb_path.exists():
        logger.error(f"[ERROR] Target workbook does not exist: {wb_path}")
        return 1

    # 2. Check Internet Connectivity
    if not check_internet_connection():
        logger.error(
            "[ERROR] Internet connection is unavailable. "
            "Cannot connect to market data servers. Please check your network and retry."
        )
        return 2

    # 3. Fetch Stock Data for all 3 symbols
    histories: Dict[str, Dict[date, float]] = {}
    for symbol, conf in POSITIONS.items():
        name = conf["name"]
        logger.info(f"Fetching market data for {name} ({symbol})...")
        try:
            hist = fetch_symbol_history(symbol)
            histories[symbol] = hist
            logger.info(f"  -> Successfully fetched {len(hist)} sessions for {symbol}.")
        except Exception as exc:
            logger.error(f"[ERROR] Failed to fetch market data for {name} ({symbol}): {exc}")
            return 3

    # 4. Find latest common NSE trading day
    common_dates = set.intersection(*(set(h.keys()) for h in histories.values()))
    if not common_dates:
        logger.error("[ERROR] No common trading date found across all 3 stocks.")
        return 4

    latest_date = max(common_dates)
    today = datetime.now().date()
    weekday_name = latest_date.strftime("%A")

    if latest_date < today:
        logger.info(
            f"[INFO] Notice: Today is {today.isoformat()} ({datetime.now().strftime('%A')}). "
            f"NSE market is closed (weekend/holiday/pre-market). "
            f"Using latest available trading session: {latest_date.isoformat()} ({weekday_name})."
        )
    else:
        logger.info(f"[INFO] Latest trading session date: {latest_date.isoformat()} ({weekday_name})")

    # 5. Extract latest prices and validate
    prices: Dict[str, float] = {}
    for symbol, conf in POSITIONS.items():
        p = histories[symbol].get(latest_date)
        if p is None or p <= 0:
            logger.error(f"[ERROR] Missing or invalid closing price for {symbol} on {latest_date.isoformat()}.")
            return 5
        prices[symbol] = p

    # Log fetched prices
    print("\n" + "-" * 65)
    print(f"LATEST NSE CLOSING PRICES ({latest_date.isoformat()} - {weekday_name}):")
    print("-" * 65)
    for sym, conf in POSITIONS.items():
        p = prices[sym]
        qty = conf["quantity"]
        val = p * qty
        print(f"  * {conf['name']:<12} ({sym:<12}) : Rs {p:>9.2f}  x  {qty:>3} shares  =  Rs {val:>11.2f}")
    print("-" * 65 + "\n")

    # 6. Update Workbook
    try:
        row_num, is_new, calculated_values = update_workbook(wb_path, latest_date, prices, dry_run=dry_run)
    except Exception as exc:
        logger.error(f"[ERROR] Failed to update workbook: {exc}", exc_info=True)
        return 6

    # 7. Print summary
    status_str = "NEW ROW ADDED" if is_new else "EXISTING ROW UPDATED"
    print("=" * 65)
    print("UPDATE SUMMARY:")
    print(f"  Workbook       : {wb_path.name}")
    print(f"  Web Dashboard  : portfolio_dashboard.html")
    print(f"  Trading Date   : {latest_date.isoformat()} ({weekday_name})")
    print(f"  Status         : {status_str} (Row {row_num})")
    print(f"  JSW Value      : Rs {calculated_values['JSWSTEEL.NS']:,.2f}")
    print(f"  Coal Value     : Rs {calculated_values['COALINDIA.NS']:,.2f}")
    print(f"  Eternal Value  : Rs {calculated_values['ETERNAL.NS']:,.2f}")
    print(f"  Total Portfolio: Rs {calculated_values['TOTAL']:,.2f}")
    print("=" * 65)
    print("Tracker & Dashboard successfully synchronized!\n")
    return 0


def main():
    parser = argparse.ArgumentParser(description="Automated NSE Stock Price Tracker")
    parser.add_argument(
        "--file",
        "-f",
        help="Path to automated_stock_price_tracker.xlsx (defaults to script directory)",
        default=None,
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Fetch prices and display updates without modifying the Excel file",
    )
    args = parser.parse_args()
    sys.exit(run_tracker(workbook_path_str=args.file, dry_run=args.dry_run))


if __name__ == "__main__":
    main()
