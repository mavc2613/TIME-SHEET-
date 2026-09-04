# Timesheet Automation

Turns photos/screenshots of punch-machine output into a clean Excel timesheet.
Images are read directly by Claude (built-in vision) — no external OCR API or
credentials needed.

## Workflow

1. Upload punch-machine images in the chat (in batches, however many you have).
2. Claude reads each image and extracts one row per punch:
   `employee_id, employee_name, date, time_in, time_out`, plus the source
   filename for traceability.
3. Extracted rows are appended to `data/punch_records.csv` (the master record —
   nothing is overwritten between batches).
4. `scripts/build_timesheet.py` regenerates `output/timesheet.xlsx` from the
   full CSV.

## Folders

- `punch_images/` — optional: drop image files here instead of pasting them
  in chat, if you'd rather commit them to the repo.
- `data/punch_records.csv` — master data store, accumulates across all batches.
- `output/timesheet.xlsx` — the generated timesheet (regenerated each run,
  never hand-edited).
- `scripts/build_timesheet.py` — rebuilds the Excel file from the CSV.

## Regenerating the Excel file manually

```bash
pip install -r requirements.txt
python3 scripts/build_timesheet.py
```
