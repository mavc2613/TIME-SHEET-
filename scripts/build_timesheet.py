#!/usr/bin/env python3
"""
Builds output/timesheet.xlsx from data/punch_records.csv.

Run this after new rows are appended to the CSV:
    python3 scripts/build_timesheet.py
"""
import csv
import datetime as dt
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
CSV_PATH = ROOT / "data" / "punch_records.csv"
OUT_PATH = ROOT / "output" / "timesheet.xlsx"

HEADERS = ["Employee ID", "Date", "Time Out", "Time In", "Notes"]
FONT_NAME = "Arial"
BLANK = "-"


def parse_date(value):
    value = (value or "").strip()
    if not value:
        return None
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return value  # leave as raw text if unparseable, so nothing silently disappears


def parse_time(value):
    value = (value or "").strip()
    if not value:
        return None
    for fmt in ("%H:%M", "%H:%M:%S", "%I:%M %p", "%I:%M:%S %p"):
        try:
            return dt.datetime.strptime(value, fmt).time()
        except ValueError:
            continue
    return value


def main():
    with open(CSV_PATH, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    wb = Workbook()
    ws = wb.active
    ws.title = "Punch Records"

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name=FONT_NAME, bold=True, color="FFFFFF")
    for col, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r, rec in enumerate(rows, start=2):
        date_val = parse_date(rec.get("date", "")) or BLANK
        time_out_val = parse_time(rec.get("time_out", "")) or BLANK
        time_in_val = parse_time(rec.get("time_in", "")) or BLANK
        notes_val = rec.get("notes", "").strip() or BLANK
        values = [
            rec.get("employee_id", "") or BLANK,
            date_val,
            time_out_val,
            time_in_val,
            notes_val,
        ]
        for c, val in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.font = Font(name=FONT_NAME)
            if c == 2 and isinstance(val, dt.date):
                cell.number_format = "yyyy-mm-dd"
            if c in (3, 4) and isinstance(val, dt.time):
                cell.number_format = "HH:MM"

    widths = [14, 12, 10, 10, 40]
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    ws.freeze_panes = "A2"

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT_PATH)
    print(f"Wrote {len(rows)} rows to {OUT_PATH}")


if __name__ == "__main__":
    main()
