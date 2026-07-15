"""Pipeline for generating report for ACE-D Aim 3 randomization.

Reads two input CSVs, cleans/reformats each, and writes them to a single
Excel workbook with one tab per CSV.
"""

import argparse
import csv
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RichTextProperties


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

# WebNeuro test variables, grouped by test and ordered by the WebNeuro test
# battery's administration order. Each group's color (for the norm-score
# chart) is chosen so it's never adjacent, in this sequence, to a
# similar-looking neighbor.
WN_TEST_GROUPS = [
    ("Motor Tapping", ["tdomnk", "tdomsdk"], "0E76E4"),
    ("Choice Reaction Time", ["chlrrtav"], "8A4100"),
    ("Verbal Memory", ["ctmrec1", "ctmrec2", "ctmrec3", "ctmsco13"], "009C61"),
    ("Emotion Identification", [
        "getcpA", "getcpD", "getcpF", "getcpH", "getcpN", "getcpS",
        "getcrtA", "getcrtD", "getcrtF", "getcrtH", "getcrtN", "getcrtS",
        "gettrtA", "gettrtD", "gettrtF", "gettrtH", "gettrtN", "gettrtS",
    ], "C900E9"),
    ("Digit Span (Forward)", ["digitot", "digitsp"], "BE6900"),
    ("Verbal Interference", ["vi_difrt", "vcrtne2", "vi_sco2", "vcrtne", "vi_sco1"], "0095AC"),
    ("Switching of Attention", ["esoadur1", "esoaerr1", "scavr0t1", "esoadur2", "esoaerr2", "scavr0t2"], "008600"),
    ("GoNo-Go", ["g2avrtk", "g2errk", "g2fnk", "g2fpk", "g2sdrtk"], "4B32B4"),
    ("Delayed Memory", ["ctmrec4"], "E62B35"),
    ("Emotion Priming", [
        "dgtcnA", "dgtcnD", "dgtcnF", "dgtcnH", "dgtcnS",
        "dgtcrtA", "dgtcrtD", "dgtcrtF", "dgtcrtH", "dgtcrtN", "dgtcrtS",
    ], "23690D"),
    ("N-Back Continuous Performance Test", ["wmacck", "wmfnk", "wmfpk", "wmrtk"], "C84E81"),
    ("Maze", ["emzcompk", "emzerrk", "emzinitk", "emzoverk", "emztrlsk"], "DC4400"),
]

WN_RAW_VARIABLE_ORDER = [name for _, names, _ in WN_TEST_GROUPS for name in names]

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


# One color per session line (older session first).
WN_SESSION_COLORS = ["1F77B4", "D62728"]

CHART_DATA_SHEET_NAME = "_NormChartData"
CHART_SHEET_NAME = "Norm Score Chart"


def add_wn_norm_score_chart(output_path: Path, wn_df: pd.DataFrame) -> None:
    """Add a line chart of the WebNeuro normed scores to the workbook: one
    colored line per session, plotted against every normed test variable.
    """
    categories = [c for c in WN_NORMED_VARIABLE_ORDER if c in wn_df.columns]
    if not categories:
        return
    n_cols = len(categories)

    wb = load_workbook(output_path)

    data_ws = wb.create_sheet(CHART_DATA_SHEET_NAME)
    data_ws.sheet_state = "hidden"
    data_ws.append([None, *categories])
    for row_idx in range(len(wn_df)):
        session = wn_df.iloc[row_idx].get("Session", row_idx + 1)
        values = [wn_df.iloc[row_idx][c] for c in categories]
        data_ws.append([f"Session {session}", *values])
    n_lines = len(wn_df)

    cats_ref = Reference(data_ws, min_col=2, max_col=n_cols + 1, min_row=1, max_row=1)
    data_ref = Reference(data_ws, min_col=1, max_col=n_cols + 1, min_row=2, max_row=n_lines + 1)

    chart = LineChart()
    chart.height = 14
    chart.width = 32

    chart.x_axis.delete = False
    chart.x_axis.axPos = "b"
    # Normed scores straddle 0, so the default "autoZero" crossing would draw
    # the category axis (and its labels) wherever y=0 falls in the plot,
    # i.e. mid-chart. Pin it to the value axis's minimum so it stays at the
    # bottom regardless of the data's sign.
    chart.x_axis.crosses = "min"
    chart.x_axis.txPr = RichText(
        bodyPr=RichTextProperties(rot=-2700000, vert="horz"),
        # r=[] is required: Paragraph defaults to a single empty text run,
        # which OOXML renders as literal (blank) label text, hiding the
        # auto-generated category labels this txPr is meant to style.
        p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=700)), endParaRPr=CharacterProperties(sz=700), r=[])],
    )

    chart.y_axis.delete = False
    # openpyxl defaults tick marks to explicitly-none, so they must be set
    # explicitly to get visible ticks and labels.
    chart.y_axis.majorTickMark = "out"
    chart.y_axis.tickLblPos = "nextTo"

    chart.add_data(data_ref, titles_from_data=True, from_rows=True)
    chart.set_categories(cats_ref)
    for series, color in zip(chart.series, WN_SESSION_COLORS):
        series.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=color, w=19050))
        series.marker = Marker(symbol="none")
        series.smooth = False

    chart.legend.position = "b"
    # overlay=False reserves dedicated space for the legend below the plot
    # area, so it doesn't sit on top of the rotated x-axis labels.
    chart.legend.overlay = False

    chart_ws = wb.create_sheet(CHART_SHEET_NAME)
    chart_ws.add_chart(chart, "A1")

    wb.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("wn_csv", type=Path, help="Path to WebNeuro input CSV")
    parser.add_argument("ec_csv", type=Path, help="Path to EtCere input CSV")
    parser.add_argument("output", type=Path, help="Path to output .xlsx file")
    args = parser.parse_args()

    df1 = order_wn_columns(load_csv(args.wn_csv)).tail(2)
    df2 = load_csv(args.ec_csv)

    write_excel({WN_SHEET_NAME: df1, EC_SHEET_NAME: df2}, args.output)
    add_wn_norm_score_chart(args.output, df1)
    print(f"Wrote {args.output} with tabs: {WN_SHEET_NAME!r}, {EC_SHEET_NAME!r}")


if __name__ == "__main__":
    main()
