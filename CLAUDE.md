# CLAUDE.md

Guidance for Claude Code (and any other coding agent) working in this repository.
For the full story of the project see [HANDOFF.md](HANDOFF.md); for end-user
instructions see [README.md](README.md) and [STEPS.txt](STEPS.txt).

## What this project is

**Recon Mismatch Analyzer** - a small Python tool that reads "recon" mismatch CSV
exports (a row-by-row comparison of two copies of the same table, `s2` = live/source,
`s3` = the new parallel table being validated) and explains, for every mismatched
column, *how* the two values differ: whitespace, date format, rounding, truncation,
NULL handling, real value change, and so on.

Output is exactly two files per run, named `report_summary_<run-folder-name>`:
an Excel workbook (`.xlsx`) and a self-contained HTML page (`.html`) with the same numbers.

## Commands

```
pip install -r requirements.txt                 # openpyxl, python-dateutil (+2 transitive), Python 3.8+
python test_patterns.py                         # MUST end with: all 72 cases produced the expected pattern
python recon_html_report.py sample_input\20260907            # xlsx + html -> reports\20260907\
python recon_html_report.py "<run-folder>" --summary         # also print what needs attention
python recon_html_report.py "<run-folder>" -o "<out-dir>\report_summary_<run>.xlsx"
python recon_analyzer.py "<run-folder>" -o out.xlsx          # engine alone: Excel only, no HTML
python install_skill.py                                      # install the /recon-summary Claude Code skill
```

There is no build step, linter config or CI. `test_patterns.py` is the only test suite;
run it after **any** change to `recon_analyzer.py`. The sample data in `sample_input/`
must keep producing a clean run (Errors sheet says "(none) - no problems").

## Layout

| Path | Role |
|---|---|
| `recon_analyzer.py` | The engine (~2,240 lines): CSV streaming, pair classification, risk tiers, Excel writer, JSON digest. Also a CLI. |
| `recon_html_report.py` | The normal entry point. Runs `recon_analyzer.py` as a subprocess, reads its JSON digest, renders the HTML page. Imports `recon_analyzer` (read-only) for descriptions/glossary. Unknown flags are passed through to the analyzer. |
| `test_patterns.py` | Coverage matrix: `(type, column, s2, s3, expected_pattern)` cases + tier and observation guards. |
| `install_skill.py` | Copies `claude_skill/recon-summary/` into `~/.claude/skills/` (or `./.claude/skills/` with `--project`) and substitutes the `{{TOOL_FOLDER}}` placeholder. |
| `claude_skill/recon-summary/SKILL.md` | Source of the `/recon-summary` slash command. Edit here, never the installed copy. |
| `sample_input/20260907/` | Two tiny example tables in the exact input format. |
| `reports/` | Generated output. **Gitignored** - never commit. |
| `STEPS.txt` | Long-form operator manual (Windows cmd oriented). Keep in sync when flags or output change. |

## Data flow

```
<run>/<table>/<table>_mismatch.csv --(recon_analyzer.analyze_table, streaming)--> per-table result dicts
   -> build_workbook()  -> report_summary_<run>.xlsx
   -> build_digest()    -> temp JSON --(recon_html_report.build_view/render_html)--> report_summary_<run>.html
```

Input CSV: header is `match_type, mismatch_columns, <primary-key cols...>, <col>__s2, <col>__s3, ...`.
Primary-key columns are inferred as everything between `mismatch_columns` and the first `__s2`
column. `mismatch_columns` holds plain column names (`balance,open_date`), not `balance__s2`.
Only that one CSV per table folder is read; a table folder is any subdirectory of the run folder.
File discovery is any `*mismatch*.csv` (prefers `<folder>_mismatch.csv`); the `<x>_recon` folder
naming in the docs is a convention of the source system, not a requirement.

## Design rules - do not break these

1. **Never reassure without proof.** A pattern that says two values mean the same thing
   (`date_format_diff`, `numeric_scale_diff`, ...) is returned only after exact
   `Decimal`/date arithmetic proves equality. When it cannot be proven the result is a
   `CANNOT VERIFY` pattern (`date_ambiguous_order`, `ambiguous_decimal_comma`,
   `percent_sign_diff`, `no_actual_diff`, ...). Nothing is labelled "low risk" or "fine".
2. **Never guess causes.** Output states measured facts only. `value_diff` is the honest
   fallback when no rule explains a difference.
3. **Never lose a problem silently.** Missing/unreadable/wrong-header tables are listed as
   `NOT ANALYSED` (Errors sheet); skipped or malformed rows become warnings. An empty row is
   never a clean table.
4. **Bounded memory.** CSVs are streamed; sizes of caches and trackers are capped
   (`SAMPLE_ROWS`, `CACHE_MAX`, `PK_TRACK_MAX` in `recon_analyzer.py`). Keep it that way.
5. **The workbook is the record.** The HTML page is a view of the same digest; page and
   workbook must always show identical numbers. Analysis happens once.
6. **Pattern descriptions and glossary live in `recon_analyzer.py`** (`PATTERN_INFO`, `GLOSSARY`)
   and are imported by the HTML script - never duplicate wording there.

## Recipe: adding or changing a comparison pattern

1. Add the detection to `classify_pair()` (order matters: the first matching test wins, so put
   specific rules before general ones; the fallback is `value_diff`).
2. Register it in `PATTERN_INFO` with a risk tier (`R_REAL`, `R_UNSURE`, `R_TEXT`, `R_CHECK`, `R_LOW`)
   and a plain-language description.
3. If it is a genuine *value* change (not a representation change), add it to `SHAPE_EXCLUDE`.
4. Add a glossary row in `GLOSSARY` (meaning, example, exact test).
5. Add cases to `CASES` in `test_patterns.py`, including a "must not be reassuring" case if the
   pattern could ever be mistaken for harmless; add it to `MUST_NOT_BE_REASSURING` if its tier must
   stay REAL or CANNOT VERIFY.
6. Run `python test_patterns.py` and the sample run. Update the case count quoted in
   `STEPS.txt` (Part 1 Step 5) and here if it changes (currently 72).

## Conventions

- Python standard library plus `openpyxl` and `python-dateutil`; do not add dependencies without a strong reason.
- Match the existing style: long explanatory comments that say *why*, plain functions, no classes,
  f-strings, error messages written for a non-programmer operator.
- User-facing text uses the report's own labels verbatim (`REAL VALUE DIFF`,
  `CANNOT VERIFY - check source`, `TEXT FIELD - read manually`, `FORMAT/TYPE - check joins`,
  `WHITESPACE/CASE ONLY`). Don't rename them: the skill, docs and glossary all quote them.
- The HTML page must stay a single offline file (no external scripts, fonts or network calls).
- Docs and examples target Windows (cmd/PowerShell paths); the Python itself is portable.

## Gotchas

- `recon_analyzer.py` on its own writes only Excel and defaults to `recon_summary.xlsx` in the
  current directory; `recon_html_report.py` defaults to `reports\<run>\report_summary_<run>.xlsx`.
- If the target `.xlsx` is open in Excel the workbook is saved under a timestamped fallback name;
  the HTML keeps the requested name.
- The `.html` embeds the absolute path of the run folder it was built from (in its data block).
  Check that path before sharing a page outside your organisation.
- Counts are **sample** counts: the source export keeps up to ~20 rows per distinct
  `mismatch_columns` combination, so "Records" is not a production total.
- The module docstring at the top of `recon_analyzer.py` shows an older `<tablename>_mismatch.csv`
  layout; the code is more permissive (see "Data flow" above).
- Do not commit generated output (`reports/`, `__pycache__/`, `.venv/`, `wheels/`) or the installed
  skill copy (`.claude/`). See `.gitignore`.
- Keep the repository free of personal data: no real names, e-mail addresses or machine-specific
  absolute paths in code, docs, sample data or commit metadata.
