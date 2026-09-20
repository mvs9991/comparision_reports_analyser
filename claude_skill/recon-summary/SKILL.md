---
name: recon-summary
description: Summarise a folder of recon mismatch CSV exports into the Excel summary workbook (Overview, Errors, Across Tables, per-table sheets) plus an HTML page of the same findings, then report what needs attention. Use when the user gives a recon run folder (one <table>_recon subfolder per table) and wants the recon / mismatch / comparison summary report.
argument-hint: <run-folder> [output-folder] [--recon-table "schema.{table}"] [--examples N] [--max-combos N]
---

# Recon summary

Runs the Recon Mismatch Analyzer on a run folder. The output is exactly TWO
files - the Excel workbook and the HTML page. Do not create any other report
files, and do not modify the tool's scripts or any input file.

Arguments given: `$ARGUMENTS`

## 1. Locate the tool

Tool folder (written here by `install_skill.py`): `{{TOOL_FOLDER}}`

It must contain `recon_analyzer.py` and `recon_html_report.py`. If the line above
still shows a `{{...}}` placeholder, or that folder no longer holds those files
(the tool was moved), check the current working directory; if not there either,
ask the user where the tool folder is and suggest re-running
`python install_skill.py` from it.

Python to use: `<tool folder>\.venv\Scripts\python.exe` if it exists, otherwise
`python` (or `py` if `python` is not found). Needs openpyxl and python-dateutil
(see `requirements.txt`); if either is missing, say so and point at STEPS.txt
Part 1 Step 4 rather than installing anything unasked.

## 2. Check the input

- No run folder given: ask for it.
- The folder must CONTAIN table subfolders (`<table>_recon\<table>_recon_mismatch.csv`).
  If the user pointed at a single table folder or a CSV, say so and suggest its
  parent folder instead - do not guess.
- `<run>` below is the run folder's own name (e.g. `20260907`). Outputs are
  always named `report_summary_<run>` - do not invent other names.
- The second positional argument, if given, is the output FOLDER. Turn it into
  `-o "<output-folder>\report_summary_<run>.xlsx"`. If the user gives a path
  ending in `.xlsx`, pass it as `-o` unchanged. With neither, omit `-o`: the
  default is `<tool folder>\reports\<run>\report_summary_<run>.xlsx`.
- Pass `--recon-table`, `--examples`, `--max-combos`, `--assessments` through
  unchanged. Add `--csv` or `--json` only if the user explicitly asks for those files.

## 3. Run it (one command)

```
"<python>" "<tool folder>\recon_html_report.py" "<run-folder>" --summary [-o "<output-folder>\report_summary_<run>.xlsx"] [other options]
```

Quote every path. One run: the CSVs are read once, the workbook is written, and
the HTML page is built from the same results. Result:

| File | What it is |
|---|---|
| `report_summary_<run>.xlsx` | the workbook: Overview, Errors, Across Tables, one sheet per table, What the terms mean |
| `report_summary_<run>.html` | the same findings as a web page - opens offline in any browser |

A large run can take minutes (STEPS.txt: ~1 minute for 200 tables) - use a long
timeout. If it fails, show the error line and match it against STEPS.txt
Part 4. If the workbook was saved under a timestamped name because the original
was open in Excel, tell the user the new name the tool printed.

## 4. Report back

Use the `SUMMARY` block the command printed (from `--summary`) - do not open the
workbook or re-analyse anything. Reply briefly, in this order, using only what
it says:

1. The two file paths - the HTML page first, then the xlsx.
2. One line of totals: tables analysed, need manual check, not analysed,
   analysed with warnings.
3. **Not analysed** tables with their reason (these are on the Errors sheet) -
   never describe them as clean.
4. **Manual check** tables: name, worst risk, and the reasons (if there are many
   tables, show the first 10 and say how many more).
5. **Warnings**, if any, one line per table.
6. The top 3 cross-table patterns, e.g. "date_format_diff in 17 tables" - these
   are usually one fix.

Keep the tool's own labels (REAL VALUE DIFF, CANNOT VERIFY - check source, ...).
Do not call any table fine, safe or low risk beyond what the report states.

Then open the HTML page for the user (`Start-Process "<html path>"` in PowerShell,
or `start "" "<html path>"` in cmd) unless they asked not to.
