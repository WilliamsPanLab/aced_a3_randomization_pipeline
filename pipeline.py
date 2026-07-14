"""Pipeline for generating report for ACE-D Aim 3 randomization.

Reads two input CSVs, cleans/reformats each, and writes them to a single
Excel workbook with one tab per CSV.
"""

import argparse
import csv
from pathlib import Path

import pandas as pd


def load_csv(path: Path) -> pd.DataFrame:
    """Load a CSV that may have ragged rows (stray trailing commas, or
    rows shorter/longer than the header) without pandas' strict rectangular
    parser rejecting it.
    """
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = [row for row in csv.reader(f) if any(cell.strip() for cell in row)]

    header = [cell.strip() for cell in rows[0]]
    width = max(len(header), max((len(row) for row in rows[1:]), default=0))
    header += [f"col_{i}" for i in range(len(header), width)]

    data = [row + [None] * (width - len(row)) for row in rows[1:]]
    df = pd.DataFrame(data, columns=header)

    for col in df.columns:
        df[col] = df[col].map(lambda v: v.strip() if isinstance(v, str) else v)
        df[col] = df[col].replace({"": None, "NA": None})
        numeric = pd.to_numeric(df[col], errors="coerce")
        if numeric.notna().equals(df[col].notna()):
            df[col] = numeric

    return df


WN_SHEET_NAME = "WebNeuro"
EC_SHEET_NAME = "EtCere"


def write_excel(sheets: dict[str, pd.DataFrame], output_path: Path) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("wn_csv", type=Path, help="Path to WebNeuro input CSV")
    parser.add_argument("ec_csv", type=Path, help="Path to EtCere input CSV")
    parser.add_argument("output", type=Path, help="Path to output .xlsx file")
    args = parser.parse_args()

    df1 = load_csv(args.wn_csv)
    df2 = load_csv(args.ec_csv)

    write_excel({WN_SHEET_NAME: df1, EC_SHEET_NAME: df2}, args.output)
    print(f"Wrote {args.output} with tabs: {WN_SHEET_NAME!r}, {EC_SHEET_NAME!r}")


if __name__ == "__main__":
    main()
