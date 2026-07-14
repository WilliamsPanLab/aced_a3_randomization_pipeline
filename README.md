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

Or open this folder in VS Code and choose "Dev Containers: Reopen in Container"
(uses the same Dockerfile, plus the Claude Code CLI feature and a live mount
of the repo — including the git-ignored `data/` folder — for local testing).

## Usage

Local:

```
python pipeline.py <wn_csv> <ec_csv> <output.xlsx>
```

Docker:

```
docker run --rm -v "${PWD}:/workspace" rand-pipeline python pipeline.py <wn_csv> <ec_csv> <output.xlsx>
```

Reads two CSVs (tolerating ragged rows — stray trailing commas, rows shorter
or longer than the header), strips whitespace/quoting from headers and
values, converts numeric-looking columns, and writes each CSV to its own tab
in a single Excel workbook (`wn_csv` → "WebNeuro" tab, `ec_csv` → "EtCere" tab,
fixed).

## Test data

`data/` is git-ignored and excluded from the built Docker image (see
`.dockerignore`) — it's only meant to hold local participant CSVs for manual
testing, never to be committed or shipped in an image.
