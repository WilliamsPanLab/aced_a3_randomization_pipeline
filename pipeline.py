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

# Identifying/demographic columns, in the order they should appear first.
WN_IDENTIFYING_COLUMNS = ["ID", "Session", "Age", "Gender", "TestDate"]

# Raw WebNeuro test variables, grouped and ordered by the WebNeuro test
# battery's administration order.
WN_RAW_VARIABLE_ORDER = [
    # Motor Tapping
    "tdomnk", "tdomsdk",
    # Choice Reaction Time
    "chlrrtav",
    # Verbal Memory
    "ctmrec1", "ctmrec2", "ctmrec3", "ctmsco13",
    # Emotion Identification
    "getcpA", "getcpD", "getcpF", "getcpH", "getcpN", "getcpS",
    "getcrtA", "getcrtD", "getcrtF", "getcrtH", "getcrtN", "getcrtS",
    "gettrtA", "gettrtD", "gettrtF", "gettrtH", "gettrtN", "gettrtS",
    # Digit Span (Forward)
    "digitot", "digitsp",
    # Verbal Interference (Words/Colors)
    "vi_difrt", "vcrtne2", "vi_sco2", "vcrtne", "vi_sco1",
    # Switching of Attention (Part 1/2)
    "esoadur1", "esoaerr1", "scavr0t1", "esoadur2", "esoaerr2", "scavr0t2",
    # GoNo-Go
    "g2avrtk", "g2errk", "g2fnk", "g2fpk", "g2sdrtk",
    # Delayed Memory
    "ctmrec4",
    # Emotion Priming
    "dgtcnA", "dgtcnD", "dgtcnF", "dgtcnH", "dgtcnS",
    "dgtcrtA", "dgtcrtD", "dgtcrtF", "dgtcrtH", "dgtcrtN", "dgtcrtS",
    # N-Back Continuous Performance Test
    "wmacck", "wmfnk", "wmfpk", "wmrtk",
    # Maze
    "emzcompk", "emzerrk", "emzinitk", "emzoverk", "emztrlsk",
]

# Normed counterparts follow the same order as their raw variables.
WN_NORMED_VARIABLE_ORDER = [f"{name}_norm" for name in WN_RAW_VARIABLE_ORDER]


def order_wn_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Reorder WebNeuro columns: identifying/demographic info first, then
    raw test variables, then normed test variables, each ordered by the
    WebNeuro test battery's administration order (testorder).
    """
    known = set(WN_RAW_VARIABLE_ORDER) | set(WN_NORMED_VARIABLE_ORDER)
    identifying = [c for c in WN_IDENTIFYING_COLUMNS if c in df.columns]
    identifying += [c for c in df.columns if c not in known and c not in identifying]

    ordered = (
        identifying
        + [c for c in WN_RAW_VARIABLE_ORDER if c in df.columns]
        + [c for c in WN_NORMED_VARIABLE_ORDER if c in df.columns]
    )
    return df[ordered]


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

    df1 = order_wn_columns(load_csv(args.wn_csv)).tail(1)
    df2 = load_csv(args.ec_csv)

    write_excel({WN_SHEET_NAME: df1, EC_SHEET_NAME: df2}, args.output)
    print(f"Wrote {args.output} with tabs: {WN_SHEET_NAME!r}, {EC_SHEET_NAME!r}")


if __name__ == "__main__":
    main()
