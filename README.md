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

## Test data

`data/` is git-ignored and excluded from the built Docker image (see
`.dockerignore`) — it's only meant to hold local participant CSVs for manual
testing, never to be committed or shipped in an image.
