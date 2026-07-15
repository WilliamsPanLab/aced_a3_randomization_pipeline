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
python pipeline.py <wn_csv> <ec_csv> <output.xlsx>
```

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

The last two rows of `wn_csv` are kept (a WebNeuro export can contain one row
per session; the WebNeuro tab reflects the two most recent sessions).

The workbook also gets a "Norm Score Chart" tab: a line chart of the normed
scores for both sessions, one colored line per session (`WN_SESSION_COLORS`
in `pipeline.py`). The y-axis is titled "Normed Score" with value ticks, and
the x-axis shows every variable name (rotated 45° to fit all 64). The legend
(bottom) lists the two session lines.

## Test data

`data/` is git-ignored and excluded from the built Docker image (see
`.dockerignore`) — it's only meant to hold local participant CSVs for manual
testing, never to be committed or shipped in an image.
