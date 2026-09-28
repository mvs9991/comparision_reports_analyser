# Handoff

A briefing for whoever picks this project up next - a person, Claude Code, or any other agent.
Read this first, then [CLAUDE.md](CLAUDE.md) (working rules), then skim [STEPS.txt](STEPS.txt)
(operator manual). Last verified: `python test_patterns.py` -> `all 79 cases produced the expected pattern`
and `python test_end_to_end.py` -> `all end-to-end checks passed`, on Python 3.9.6 and 3.12.10.

## 1. The problem being solved

A team validates a data migration by running a **recon** (reconciliation) between two copies of
each table:

- **s2** - the live / source table (what production has today).
- **s3** - the new parallel table that was ingested and is being validated.

The recon tool emits, per table, a `*_mismatch.csv` listing every record whose primary key exists on
both sides but whose values differ in one or more columns. Reading thousands of such rows by hand is
impractical, and most differences are noise (trailing spaces, `2024-01-05` vs `20240105`,
`10.00` vs `10.0`). The few that matter (rounding, truncation, NULL vs empty, genuinely different
values) get lost.

This tool classifies **every differing value pair** into a named pattern with a risk tier, then rolls
the results up per column, per table and across tables ("17 tables show the same date-format
change - one fix, not 17").

## 2. What exists today (state: complete and working)

| Deliverable | Status |
|---|---|
| Pair classifier with 41 named patterns (plus dynamic `punct_removed_in_s3:<ch>` / `punct_added_in_s3:<ch>`), exact decimal/date arithmetic | Done, covered by 79 test cases |
| End-to-end self-test (`test_end_to_end.py`): broken inputs, workbook vs page, CLI checks, page escaping | Done |
| Streaming CSV reader (constant memory; tested to ~3M rows / 1 GB per table per STEPS.txt) | Done |
| Excel workbook: Overview, Errors, Across Tables, one sheet per table, glossary | Done |
| Self-contained HTML report (filters, charts, highlighted diffs, "Copy SELECT" buttons, print-friendly) | Done |
| Claude Code skill `/recon-summary <run-folder>` + installer | Done |
| Operator manual (`STEPS.txt`), README, this handoff, CLAUDE.md | Done |

There is no open bug list. No CI, no packaging (`setup.py`/`pyproject.toml`) and no licence file -
add them if the project is to be distributed formally.

## 3. Architecture in one page

```
recon_html_report.py  (entry point)
  |-- subprocess -> recon_analyzer.py <run-folder> -o <xlsx> --json <temp digest>
  |                    find_mismatch_csv -> open_table (streaming) -> analyze_table
  |                    -> classify_pair (per value pair, cached)
  |                    -> build_workbook (openpyxl)   -> report_summary_<run>.xlsx
  |                    -> build_digest                -> temp JSON
  |-- build_view(digest) -> render_html(view) -> report_summary_<run>.html
  |-- temp digest deleted unless --json was given
```

Key functions in `recon_analyzer.py`:

- `classify_pair(col, s2, s3)` - the heart. Ordered rule cascade, first match wins: identical ->
  null handling -> whitespace (edge / internal) -> case -> leading zeros -> infinity tokens ->
  boolean on a flag-named column -> epoch vs formatted date -> dates and numbers (which is tried
  first depends on whether the column name / values look like dates) -> plain boolean ->
  typographic (smart quotes, dashes) -> punctuation added/removed -> encoding loss -> unicode fold
  -> separator format -> truncation / containment -> fallback `value_diff`.
  Returns `(pattern, delta)`. Text-field and boolean-column overrides are applied afterwards in
  `analyze_table()`.
- `PATTERN_INFO` / `pattern_info()` - pattern -> (risk tier, plain-language description).
  Dynamic patterns `punct_removed_in_s3:<ch>` / `punct_added_in_s3:<ch>` are handled in code.
- `analyze_table()` - streams one CSV, groups by `mismatch_columns` combination, keeps counters
  and a few examples per pattern, tracks duplicate primary keys and "unlisted" differences (columns
  that differ but were not named in `mismatch_columns`), and calls `manual_check_reasons()`.
- `build_workbook()` + `write_*_sheet()` - Excel output. `GLOSSARY` documents every label and the
  exact test behind it.
- `main()` - CLI: discovers table folders, isolates per-table failures so one bad table never
  stops the run.

`recon_html_report.py`: `build_view()` reshapes the digest (sorting, cross-table rollup),
`TEMPLATE` is one big HTML/JS/CSS string, `render_html()` injects the JSON, `print_summary()`
produces the `--summary` console block the skill reads.

## 4. Risk tiers (the vocabulary everything else uses)

| Label | Meaning |
|---|---|
| `REAL VALUE DIFF` | The value itself differs (rounding, truncation, NULL vs empty, currency changed, ...). Needs a decision. |
| `CANNOT VERIFY - check source` | The CSV cannot prove match or difference (ambiguous `01/05/2024`, `12,34`, `50%`, export lost the difference). Deliberately never called fine. |
| `TEXT FIELD - read manually` | Free-text wording changed; a human must read it. |
| `FORMAT/TYPE - check joins` | Proven same value, written differently. Breaks exact joins / type-sensitive logic. |
| `WHITESPACE/CASE ONLY` | Only spacing or capitalisation differs. Still not called "low risk". |

A table is flagged **Manual check needed** when something measurable looks odd even though every
check ran (unexplained `value_diff`, unverifiable values, mixed real+format changes, a column empty
in s3, duplicate keys, malformed lines, differences the recon did not list, ...). See
`manual_check_reasons()`.

## 5. Principles that shaped the code (why it looks the way it does)

- **Honesty over convenience.** A false "looks fine" in a migration sign-off is the worst outcome, so
  every reassuring verdict needs proof and every unprovable case gets its own explicit label. The
  `NO FALSE OK` and `PROVEN SAME` blocks in `test_patterns.py` record past false reassurances.
- **Fail visibly, per table.** Bad header, missing CSV, unreadable folder, stray CSVs in the run
  folder -> a `NOT ANALYSED` row on the Errors sheet, never a crash and never a blank that looks clean.
- **Scale.** Streaming, bounded caches, bounded duplicate-key tracking (with a warning when the
  bound is hit), Excel cell/sheet limits respected (`safe_cell`, `--max-combos`, chunked "Where" cells).
- **Operator-friendly.** Written for people who are not programmers: no config files, one command,
  errors phrased as instructions, works offline, works on locked-down office laptops (see the
  proxy / offline-wheel install options in STEPS.txt).

## 6. How to verify your work

1. `python test_patterns.py` - must end with `all 79 cases produced the expected pattern`;
   `python test_end_to_end.py` - must end with `all end-to-end checks passed`.
2. `python recon_html_report.py sample_input\20260907 --summary` - expect 2 tables analysed, the
   Errors sheet reading "(none) - no problems", and both output files under `reports\20260907\`.
3. Open the `.html` in a browser (no server needed) and the `.xlsx` in Excel; numbers must match.
4. If you touched the skill, run `python install_skill.py`, restart Claude Code, run
   `/recon-summary sample_input\20260907`.

## 7. Known limitations / quirks

- Windows-first documentation (cmd/PowerShell). The Python code is portable but has only been
  exercised on Windows.
- Counts reflect the **sampled** recon export (up to ~20 rows per distinct combination), not
  production totals - stated in the glossary and digest.
- Only `*mismatch*.csv` is read; the `added`/`removed`/`matched` exports are ignored by design.
- HTML shows at most 200 combinations per table (`MAX_COMBOS_IN_PAGE`); the workbook shows 500 by
  default (`--max-combos 0` for all).
- The generated HTML embeds the absolute path of the analysed run folder. Review before sharing externally.
- The opening docstring of `recon_analyzer.py` describes an older input naming; the code accepts any
  `*mismatch*.csv` inside each table folder.
- Free-text/boolean detection is heuristic (column-name patterns + value shape); see `is_text_field`
  and `is_bool_field`. Tune there if a column is mis-typed.
- The skill (`SKILL.md`) hard-codes the tool folder at install time; re-run `install_skill.py` after
  moving the folder.

## 8. Ideas not yet done

- `pyproject.toml` / pip-installable console scripts; a licence file.
- CI (GitHub Actions) running both test scripts on Windows + Linux.
- Convert both test scripts to `pytest` (keeping the table-driven `CASES`).
- Split `recon_analyzer.py` (classifier / analysis / Excel writer) and move the HTML template into its
  own file - they are independent and the module is large.
- Redact or relativise the run-folder path embedded in the HTML.
- Cross-platform quick-start notes (macOS/Linux shell equivalents).

## 9. Repository hygiene

- Commit only source, docs, `sample_input/`, the skill source and `requirements.txt`.
- `reports/`, `__pycache__/`, `.venv/`, `wheels/` and `.claude/` are gitignored generated/local
  artefacts.
- Keep the repo free of personal data (names, e-mail addresses, machine-specific paths) - including
  commit author metadata; use a neutral git identity for this repository.
