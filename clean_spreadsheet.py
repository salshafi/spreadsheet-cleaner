"""Spreadsheet Cleaner: turn a messy CSV/Excel file into a clean Excel report.

Usage:
    python clean_spreadsheet.py messy_sample.xlsx
    python clean_spreadsheet.py input.csv -o cleaned.xlsx
"""
import argparse
import logging
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger("cleaner")

# Date formats we expect to see, tried in this order.
# Slash dates are read as month/day/year.
DATE_FORMATS = ["%m/%d/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d", "%b %d %Y"]


def load(path: Path) -> pd.DataFrame:
    """Read a CSV or Excel file into a DataFrame."""
    # stop early with a clear message if the path is wrong
    if not path.exists():
        sys.exit(f"File not found: {path}")
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path)
    if path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(path)
    sys.exit("Unsupported file type. Use .csv or .xlsx")


def parse_date(value):
    """Try each known format; return NaT if none of them match."""
    if pd.isna(value):
        return pd.NaT
    text = str(value).strip()
    for fmt in DATE_FORMATS:
        try:
            return pd.to_datetime(text, format=fmt)
        except ValueError:
            continue  # wrong format, try the next one
    return pd.NaT  # e.g. "not available"


def clean(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Return the cleaned DataFrame and a small summary of what changed."""
    stats = {"rows_in": len(df)}
    df = df.copy()  # don't touch the original

    # strip stray spaces from the column headers
    df.columns = [str(c).strip() for c in df.columns]

    # trim text and collapse double spaces ("emma  Wilson" -> "emma Wilson")
    for col in df.select_dtypes(include=["object", "string"]).columns:
        df[col] = df[col].astype("string").str.strip().str.replace(r"\s+", " ", regex=True)

    # column-specific fixes, only if the column exists
    if "Customer Name" in df:
        df["Customer Name"] = df["Customer Name"].str.title()
    if "Email" in df:
        df["Email"] = df["Email"].str.lower()
    if "City" in df:
        df["City"] = df["City"].str.title()
    if "Status" in df:
        df["Status"] = df["Status"].str.title()
    if "Amount" in df:
        # drop $ and commas, then turn into numbers (bad values become NaN)
        df["Amount"] = pd.to_numeric(
            df["Amount"].astype("string").str.replace(r"[$,]", "", regex=True),
            errors="coerce",
        )
    if "Order Date" in df:
        df["Order Date"] = df["Order Date"].map(parse_date)

    # do this after cleaning, so "BOB SMITH" and "Bob Smith" count as the same row
    before = len(df)
    df = df.drop_duplicates().reset_index(drop=True)
    stats["duplicates_removed"] = before - len(df)

    # flag missing values rather than guessing what they should be
    stats["missing_cells"] = int(df.isna().sum().sum())
    stats["rows_out"] = len(df)
    return df, stats


def save_excel(df: pd.DataFrame, out: Path) -> None:
    """Write a formatted Excel report, with missing cells highlighted."""
    with pd.ExcelWriter(out, engine="openpyxl", datetime_format="YYYY-MM-DD") as writer:
        df.to_excel(writer, index=False, sheet_name="Cleaned")
        ws = writer.sheets["Cleaned"]

        header_fill = PatternFill("solid", fgColor="1F4E78")
        missing_fill = PatternFill("solid", fgColor="FFE699")

        # style the header row
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")

        # highlight empty cells, show dates without a time, format money
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                # pandas can write blanks as "" instead of None, so check both
                if cell.value is None or cell.value == "":
                    cell.value = "MISSING"
                    cell.fill = missing_fill
                elif isinstance(cell.value, datetime):
                    cell.number_format = "YYYY-MM-DD"
                elif isinstance(cell.value, float):
                    cell.number_format = "#,##0.00"

        # size each column to fit its longest value
        for i, col in enumerate(df.columns, start=1):
            texts = df[col].astype(object).where(df[col].notna(), "MISSING")
            width = max(len(str(col)), *(len(str(v)) for v in texts)) + 3
            ws.column_dimensions[get_column_letter(i)].width = width

        ws.freeze_panes = "A2"  # keep the header visible when scrolling


def main() -> None:
    parser = argparse.ArgumentParser(description="Clean a messy CSV/Excel file.")
    parser.add_argument("input", type=Path, help="path to .csv or .xlsx file")
    parser.add_argument("-o", "--output", type=Path, help="output .xlsx path")
    args = parser.parse_args()

    # default output name: <input>_cleaned.xlsx next to the input file
    out = args.output or args.input.with_name(f"{args.input.stem}_cleaned.xlsx")

    df = load(args.input)
    cleaned, stats = clean(df)
    save_excel(cleaned, out)

    log.info("Rows in: %d", stats["rows_in"])
    log.info("Duplicates removed: %d", stats["duplicates_removed"])
    log.info("Missing cells flagged: %d", stats["missing_cells"])
    log.info("Rows out: %d", stats["rows_out"])
    log.info("Saved: %s", out)


if __name__ == "__main__":
    main()
