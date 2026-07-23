#!/usr/bin/env python3
"""Populate a built workbook with fixture data from sample_data/*.csv, for
manual testing in Excel. Never used to seed real client data.

Run:
    python scripts/seed_sample_data.py dist/AnvayaFilmsCRM.xlsm
"""
from __future__ import annotations

import csv
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

REPO_ROOT = Path(__file__).resolve().parent.parent
SAMPLE_DATA_DIR = REPO_ROOT / "sample_data"

# sample_data/<file>.csv -> sheet name it seeds
CSV_TO_SHEET = {
    "leads.csv": "02_CRM_Leads",
    "bookings.csv": "05_Bookings",
    "payments.csv": "12_Payments",
}


def coerce(value: str):
    """CSV cells are always strings; write them back as real Excel types
    so formulas that do arithmetic/date-math on seeded rows (e.g.
    05_Bookings row 2's BalanceAmount/DeliveryDeadline formulas against
    the seeded BK-0001 row) actually compute instead of silently treating
    the cell as text."""
    try:
        return int(value)
    except ValueError:
        pass
    try:
        return float(value)
    except ValueError:
        pass
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError:
        pass
    return value


def seed_sheet(ws, csv_path: Path) -> int:
    with csv_path.open(newline="") as f:
        reader = csv.reader(f)
        header = next(reader)

        sheet_header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        if header != sheet_header:
            raise ValueError(
                f"{csv_path.name} header {header} doesn't match "
                f"{ws.title} sheet header {sheet_header} — "
                f"update schema/*.yaml or the CSV to match."
            )

        row_count = 0
        for row_idx, row in enumerate(reader, start=2):
            for col_idx, value in enumerate(row, start=1):
                if value == "":
                    # Leave blank rather than overwrite -- row 2 on some
                    # sheets (e.g. 05_Bookings) carries a live formula in
                    # formula-type columns (see build_workbook.py
                    # FORMULA_TEMPLATES); an empty CSV cell there means
                    # "don't touch this column", not "clear it".
                    continue
                ws.cell(row=row_idx, column=col_idx, value=coerce(value))
            row_count += 1
        return row_count


def main() -> int:
    if len(sys.argv) != 2:
        print("usage: seed_sample_data.py <path-to-xlsm>", file=sys.stderr)
        return 1

    workbook_path = Path(sys.argv[1])
    if not workbook_path.exists():
        print(f"{workbook_path} does not exist — build it first with build_workbook.py", file=sys.stderr)
        return 1

    wb = load_workbook(workbook_path, keep_vba=workbook_path.suffix == ".xlsm")

    for csv_name, sheet_name in CSV_TO_SHEET.items():
        csv_path = SAMPLE_DATA_DIR / csv_name
        if not csv_path.exists():
            print(f"skip: {csv_path} not found")
            continue
        if sheet_name not in wb.sheetnames:
            print(f"skip: sheet '{sheet_name}' not in {workbook_path}")
            continue
        n = seed_sheet(wb[sheet_name], csv_path)
        print(f"seeded {n} row(s) into {sheet_name} from {csv_name}")

    wb.save(workbook_path)
    print(f"saved {workbook_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
