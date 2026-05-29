#!/usr/bin/env python3
"""
Paint Google Sheet rows by Status — yellow for 'Sent intro',
teal for 'Replied'. Idempotent — re-running just re-applies.

Usage:
    python paint_sheet.py             # color the sheet
    python paint_sheet.py --tab Vlad  # specify tab
"""
import argparse
import os
import sys
from pathlib import Path

import gspread
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials

ROOT = Path(__file__).parent
load_dotenv(ROOT / ".env")

# Soft pastel colors (RGB 0-1 floats for gspread)
COLOR_MAP = {
    "Sent intro":         {"red": 1.000, "green": 0.953, "blue": 0.804},  # soft yellow #FFF3CD
    "Sent program info":  {"red": 1.000, "green": 0.800, "blue": 0.600},  # orange #FFCC99 — escalation
    "Replied":            {"red": 0.820, "green": 0.925, "blue": 0.925},  # teal/cyan #D1ECF1
    "Wants money for posts": {"red": 0.769, "green": 0.871, "blue": 0.949},  # blue #C4DEF2
    "Live collab":        {"red": 0.851, "green": 0.945, "blue": 0.812},  # green #D9F1CF
    "Dead":               {"red": 0.949, "green": 0.831, "blue": 0.831},  # red #F2D4D4
    "Do not contact":     {"red": 0.949, "green": 0.831, "blue": 0.831},  # red #F2D4D4
}
DEFAULT_COLOR = {"red": 1.0, "green": 1.0, "blue": 1.0}  # white


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--tab", default="Vlad")
    args = parser.parse_args()

    sheet_id = os.environ.get("GSHEETS_ID")
    if not sheet_id:
        sys.exit("ERROR: GSHEETS_ID missing in .env")
    creds_path = Path(os.environ.get("GSHEETS_CREDS", Path.home() / ".config/timestripe-gsheets.json"))
    creds = Credentials.from_service_account_file(str(creds_path),
        scopes=["https://www.googleapis.com/auth/spreadsheets"])
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(sheet_id)
    ws = sh.worksheet(args.tab)

    all_rows = ws.get_all_values()
    # Find header row (containing "Name")
    header_row_idx = None
    for i, row in enumerate(all_rows):
        if row and row[0].strip().lower() == "name":
            header_row_idx = i + 1  # 1-based
            break
    if not header_row_idx:
        sys.exit("ERROR: could not locate header row with 'Name' in col A")
    headers = all_rows[header_row_idx - 1]
    status_col = headers.index("Status") + 1  # 1-based

    # Last column letter for full-row coloring
    last_col = len(headers)
    from gspread.utils import rowcol_to_a1
    last_col_letter = rowcol_to_a1(1, last_col).rstrip("1")

    # Build batch of formatting requests grouped by color
    by_color = {}
    for i, row in enumerate(all_rows, start=1):
        if i <= header_row_idx:
            continue
        if not row or not row[0].strip():
            continue
        status = row[status_col - 1].strip() if len(row) >= status_col else ""
        color = COLOR_MAP.get(status, DEFAULT_COLOR)
        key = (color["red"], color["green"], color["blue"])
        by_color.setdefault(key, []).append(i)

    print(f"Rows to paint by status:")
    for color_key, row_indices in by_color.items():
        # Find status name for label
        label = next((s for s, c in COLOR_MAP.items()
                      if (c["red"], c["green"], c["blue"]) == color_key), "Default/None")
        print(f"  {label:<25} {len(row_indices)} rows")

    # Apply formatting using batch_update via Sheets API
    requests = []
    spreadsheet_id = ws.spreadsheet.id
    sheet_id_internal = ws._properties["sheetId"]

    for color_key, row_indices in by_color.items():
        color = {"red": color_key[0], "green": color_key[1], "blue": color_key[2]}
        for row_idx in row_indices:
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id_internal,
                        "startRowIndex": row_idx - 1,
                        "endRowIndex": row_idx,
                        "startColumnIndex": 0,
                        "endColumnIndex": last_col,
                    },
                    "cell": {"userEnteredFormat": {"backgroundColor": color}},
                    "fields": "userEnteredFormat.backgroundColor",
                }
            })

    if requests:
        # Send in chunks to avoid request size limit
        BATCH = 200
        for i in range(0, len(requests), BATCH):
            ws.spreadsheet.batch_update({"requests": requests[i:i + BATCH]})
        print(f"\n✓ Painted {len(requests)} rows in '{args.tab}'")
    else:
        print("Nothing to paint.")


if __name__ == "__main__":
    main()
