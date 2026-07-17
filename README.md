Pipeline for generating report for ACE-D Aim 3 randomization.

## Setup

Local:

```
pip install -r requirements.txt
```

Docker (no local Python needed):

```
docker build -t rand-pipeline .
```

## Usage

Local:

```
python pipeline.py <wn_csv> <ec_csv> <output.xlsx> [--subset-wn-tests]
```

`--subset-wn-tests` restricts the WebNeuro tab/chart to the tests used by
the composite scores plus Switching of Attention 1 (see below); omit it to
keep every WebNeuro test.

Docker — mount your current directory to `/workspace` so input/output files
are visible on both sides:

macOS/Linux (bash/zsh):

```
docker run --rm -v "$(pwd):/workspace" rand-pipeline python pipeline.py <wn_csv> <ec_csv> <output.xlsx>
```

Windows PowerShell:

```
docker run --rm -v "${PWD}:/workspace" rand-pipeline python pipeline.py <wn_csv> <ec_csv> <output.xlsx>
```

Windows cmd.exe:

```
docker run --rm -v "%cd%:/workspace" rand-pipeline python pipeline.py <wn_csv> <ec_csv> <output.xlsx>
```

Note: on Windows, use PowerShell or cmd.exe for this command, not Git Bash —
Git Bash's `$(pwd)` produces a POSIX-style path that Docker Desktop on
Windows doesn't translate correctly, so the mount silently points to the
wrong location.

Reads two CSVs (tolerating ragged rows — stray trailing commas, rows shorter
or longer than the header), strips whitespace/quoting from headers and
values, converts numeric-looking columns, and writes each CSV to its own tab
in a single Excel workbook (`wn_csv` → "WebNeuro" tab, `ec_csv` → "EtCere" tab,
fixed).

The WebNeuro tab's columns are reordered: identifying/demographic info
(`ID`, `Session`, `Age`, `Gender`, `TestDate`) first, then raw test
variables, then normed (`_norm`) test variables — each group ordered by the
WebNeuro test battery's administration order. The variable order is
hardcoded in `pipeline.py` (`WN_RAW_VARIABLE_ORDER`).

With `--subset-wn-tests`, only the tests used by the composite scores, plus
Switching of Attention 1, are kept — Digit Span (Forward), Stroop Word,
Stroop Color, Switching of Attention 1/2, GoNo-Go, and Maze
(`WN_REPORTED_TESTS` in `pipeline.py`). Every other WebNeuro test's
raw/normed variables are dropped from the tab and chart.

The first two rows of `wn_csv` are kept (a WebNeuro export can contain one row
per session; the WebNeuro tab reflects the first two sessions — screening and
baseline).

The WebNeuro tab also gets 6 composite score columns appended at the end —
`maze_composite`, `gng_composite`, `stroopw_composite`, `stroopc_composite`,
`swoa_composite`, and `digit_composite` — each the row-wise mean of a subset
of that test's normed variables (`WN_COMPOSITE_GROUPS` in `pipeline.py`).

Only two rows of `ec_csv` are kept: the "Referenced <participant>" row (row
10, `EC_REFERENCED_ROW` in `pipeline.py` — its label includes the
participant's name, so it's picked by row position rather than a label
match) and the QC metric rows (Signal-to-Noise Ratio, Critical Motion
Control — matched by label prefix, `EC_QC_LABEL_PREFIXES`). The rest of the
EtCere export's fixed report template (raw/global scores, healthy-norm
stats, etc.) is dropped.

The workbook also gets a "Norm Score Chart" tab: a line chart of the normed
scores for both sessions, one dashed line per session (`WN_SESSION_DASH_STYLES`
in `pipeline.py`; both lines are the same color and told apart by dash
pattern instead), labeled "Screening" and "Baseline" (`WN_SESSION_LABELS` in
`pipeline.py`). Each variable's column is shaded with a background band
colored by the WebNeuro test it belongs to (`WN_TEST_GROUPS`). The y-axis
has value ticks, and the x-axis shows every variable name (rotated 45° to
fit all 64). The legend (bottom) lists the two session lines.

## Test data

`data/` is git-ignored and excluded from the built Docker image (see
`.dockerignore`) — it's only meant to hold local participant CSVs for manual
testing, never to be committed or shipped in an image.
