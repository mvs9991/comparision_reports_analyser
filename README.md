# Recon Mismatch Analyzer

Turn thousands of raw "recon mismatch" rows into a clear answer to one question:
**how exactly does the new table differ from the old one, and does it matter?**

When you migrate or re-ingest a table, a row-by-row reconciliation tells you *that*
values differ - not *why* or *how badly*. This tool reads those mismatch CSV exports and
classifies every differing value pair (`s2` = live/source, `s3` = new table) into a named
pattern with a risk level, then summarises the result per column, per table and across tables.

- Trailing spaces, `2024-01-05` vs `20240105`, `10.00` vs `10.0` - recognised and labelled as format noise.
- Rounding, truncation, NULL vs empty string, changed currency, genuinely different values - surfaced as real differences.
- Anything it cannot prove either way (`01/05/2024` vs `2024-01-05`, `50%` vs `50`, `12,34`) is labelled
  **CANNOT VERIFY** rather than waved through.

Every statement is computed by exact string, decimal or date comparison. The tool never guesses at causes.

## What you get

One run produces exactly two files, named after the run folder:

| File | Contents |
|---|---|
| `report_summary_<run>.xlsx` | Overview, Errors, Across Tables, one sheet per table, and a glossary |
| `report_summary_<run>.html` | The same findings as an offline, single-file web page: filters, charts, highlighted differences, one-click lookup queries, print-friendly |

## Quick start

Requires Python 3.8+ (tested on 3.9).

```bash
pip install -r requirements.txt
python test_patterns.py                                  # self-test: should end "all 79 cases produced the expected pattern"
python test_end_to_end.py                                # whole-run self-test: should end "all end-to-end checks passed"
python recon_html_report.py sample_input/20260907        # try it on the bundled sample
```

Reports are written to `reports/20260907/`. Open the `.html` in any browser and the `.xlsx` in Excel.

Analyse your own export:

```bash
python recon_html_report.py "<run-folder>" --summary
python recon_html_report.py "<run-folder>" -o "<output-dir>/report_summary_<run>.xlsx"
```

## Input format

Point the tool at a **run folder** that contains one subfolder per table:

```
<run-folder>/
    orders_recon/
        orders_recon_mismatch.csv      <- the only file read
    customers_recon/
        customers_recon_mismatch.csv
```

Each mismatch CSV has this header:

```
match_type, mismatch_columns, <primary key columns...>, <col>__s2, <col>__s3, ...
```

- primary-key columns sit between `mismatch_columns` and the first `__s2` column;
- `mismatch_columns` lists plain column names separated by commas (`balance,open_date`);
- see [sample_input/](sample_input/) for working examples.

A table folder with no readable mismatch CSV never stops the run - it is listed as **NOT ANALYSED** on the
Errors sheet. An empty result is never presented as a clean table.

## Risk levels

| Label | Meaning |
|---|---|
| `REAL VALUE DIFF` | The value itself differs. Needs a decision. |
| `CANNOT VERIFY - check source` | The export cannot prove match or difference. Check the source tables. |
| `TEXT FIELD - read manually` | Free-text wording changed. A person must read it. |
| `FORMAT/TYPE - check joins` | Proven same value, written differently - can still break exact joins. |
| `WHITESPACE/CASE ONLY` | Only spacing or capitalisation differs. |

The workbook's "What the terms mean" sheet documents all 41 patterns and the exact test behind each.

## Options

| Option | Effect |
|---|---|
| `-o <file>.xlsx` | Where the workbook goes; the HTML is written beside it |
| `--summary` | Print what needs attention (unanalysed tables, manual checks, warnings, top patterns) |
| `--recon-table "schema.{table}"` | Table name used in the generated `SELECT` lookup queries |
| `--examples N` | Examples kept per pattern (default 3) |
| `--max-combos N` | Column combinations listed per sheet (default 500, `0` = all) |
| `--csv` | Also write three CSV summaries next to the page |
| `--json <file>` | Also keep the raw JSON digest |

`recon_analyzer.py` on its own is the engine and writes the Excel workbook only.

## Use it from Claude Code (optional)

```bash
python install_skill.py          # installs the /recon-summary skill for your user
```

Restart Claude Code, then run `/recon-summary <run-folder>`. It runs the tool, opens the page and reports
what needs attention using the report's own labels.

## Performance

Streaming reader with bounded memory (stays under ~500 MB): about a minute for 200 tables / ~430,000 rows,
about four minutes for a single 3,000,000-row (1 GB) file. Building the HTML adds about a second.

## Project layout

```
recon_html_report.py   main command: Excel + HTML in one run
recon_analyzer.py      analysis engine (also a stand-alone Excel-only CLI)
test_patterns.py       coverage matrix for every comparison rule
test_end_to_end.py     whole-run checks: broken inputs, workbook vs page, CLI handling
install_skill.py       installs the Claude Code skill
claude_skill/          source of the /recon-summary skill
sample_input/          two small example tables in the expected format
STEPS.txt              detailed operator manual (setup, offline install, troubleshooting)
CLAUDE.md, HANDOFF.md  orientation for coding agents and new maintainers
```

## Contributing

Run `python test_patterns.py` and `python test_end_to_end.py` after any change to the scripts. To add a comparison pattern, follow the
recipe in [CLAUDE.md](CLAUDE.md); for background and design rationale see [HANDOFF.md](HANDOFF.md).
