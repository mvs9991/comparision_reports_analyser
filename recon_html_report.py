"""Recon Mismatch Analyzer - Excel workbook + HTML page in one run.

Output of a run: TWO files, side by side, named after the run folder:
    report_summary_<folder>.xlsx   the workbook, with all its sheets
    report_summary_<folder>.html   the same findings as a web page (opens offline)

How: recon_analyzer.py reads the mismatch CSVs once and writes the workbook
plus a JSON digest; the page is built from that digest, which is temporary
and removed afterwards. Nothing is analysed twice. The pattern descriptions
and glossary are imported from recon_analyzer.py, so they match the workbook
word for word.

       python recon_html_report.py "D:\\recon_exports\\20260907"
           -> reports\\20260907\\report_summary_20260907.xlsx + .html

       python recon_html_report.py "D:\\recon_exports\\20260907" -o "D:\\recon_reports\\report_summary_20260907.xlsx"
           -> D:\\recon_reports\\report_summary_20260907.xlsx + .html

Any other recon_analyzer.py option is passed straight through, e.g.
    --recon-table "myschema.{table}"   --examples 5   --max-combos 0

Optional extras (nothing extra is written unless asked):
    --summary       print what needs attention to the console
    --csv           also write three CSV summaries beside the page
    --json <file>   also keep the JSON digest

A page can also be built from a digest kept earlier, without re-analysing:
       python recon_html_report.py "D:\\recon_reports\\digest.json"
"""

import argparse
import csv
import json
import os
import shutil
import subprocess
import sys
import tempfile
from collections import Counter, defaultdict
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import recon_analyzer as ra  # noqa: E402  (read-only: descriptions, risk order, glossary)

# risk label -> short id used by the page for colours and filters
TIER_ID = {
    ra.R_REAL: "real",
    ra.R_UNSURE: "unsure",
    ra.R_TEXT: "text",
    ra.R_CHECK: "check",
    ra.R_LOW: "low",
}
TIERS = [
    {"id": TIER_ID[label], "label": label}
    for label in sorted(TIER_ID, key=lambda l: ra.RISK_ORDER[l])
]
TIER_LABEL = {t["id"]: t["label"] for t in TIERS}

# a table with thousands of combinations would make the page slow to open;
# the workbook remains the complete record
MAX_COMBOS_IN_PAGE = 200


def tier_of(risk):
    # an unknown label is treated as the worst tier, the same as worst_risk() does
    return TIER_ID.get(risk, "real")


def build_view(digest, recon_table="{table}"):
    tables = []
    across = defaultdict(lambda: {"tables": set(), "columns": set(), "records": 0, "example": None})

    for t in digest.get("tables", []):
        pk_cols = t.get("pk_cols", [])
        tier_counts = Counter()
        columns = []
        for c in t.get("columns", []):
            tier_counts[tier_of(c["risk"])] += 1
            patterns = []
            for p, n in sorted(c["patterns"].items(), key=lambda kv: (ra.RISK_ORDER.get(ra.pattern_risk(kv[0]), 0), -kv[1])):
                risk, desc = ra.pattern_info(p)
                patterns.append({"name": p, "n": n, "tier": tier_of(risk), "desc": desc})
                g = across[p]
                g["tables"].add(t["table"])
                g["columns"].add(f"{t['table']}.{c['column']}")
                g["records"] += n
            examples = []
            for e in c.get("examples", []):
                pk = e.get("pk", {})
                examples.append({
                    "pattern": e["pattern"], "s2": e["s2"], "s3": e["s3"],
                    "pk": pk,
                    "sql": ra.lookup_query(recon_table, t["table"], list(pk.keys()), list(pk.values())),
                })
                g = across[e["pattern"]]
                if g["example"] is None:
                    g["example"] = [e["s2"], e["s3"]]
            columns.append({
                "column": c["column"],
                "field_type": c.get("field_type", "data"),
                "records": c["records"],
                "tier": tier_of(c["risk"]),
                "risk": c["risk"],
                "patterns": patterns,
                "observations": c.get("observations", []),
                "examples": examples,
            })

        combos = t.get("combos", [])
        worst = ra.worst_risk([c["risk"] for c in t.get("columns", [])])
        tables.append({
            "table": t["table"],
            "sample_rows": t.get("sample_rows", 0),
            "distinct_combos": t.get("distinct_combos", len(combos)),
            "pk_cols": pk_cols,
            "headline": t.get("headline", ""),
            "notes": t.get("notes", []),
            "unlisted": t.get("unlisted_diffs", {}),
            "reasons": t.get("manual_check_reasons", []),
            "worst": tier_of(worst),
            "tier_counts": dict(tier_counts),
            "columns": columns,
            "combos": [
                {"combo": cb["combo"], "records": cb["records"], "tier": tier_of(cb["risk"]),
                 "per_column": cb.get("per_column", {})}
                for cb in combos[:MAX_COMBOS_IN_PAGE]
            ],
            "combos_hidden": max(0, len(combos) - MAX_COMBOS_IN_PAGE),
        })

    # the Overview sheet's order: manual check first, then most real / unverifiable columns
    tier_rank = {tr["id"]: i for i, tr in enumerate(TIERS)}
    tables.sort(key=lambda x: (not x["reasons"], -x["tier_counts"].get("real", 0),
                               -x["tier_counts"].get("unsure", 0), -x["sample_rows"], x["table"]))

    across_rows = sorted(
        (
            {
                "pattern": p, "tier": tier_of(ra.pattern_risk(p)), "desc": ra.pattern_desc(p),
                "tables": len(g["tables"]), "columns": len(g["columns"]), "records": g["records"],
                "where": sorted(g["columns"], key=ra.natural_key), "example": g["example"],
            }
            for p, g in across.items()
        ),
        key=lambda r: (tier_rank[r["tier"]], -r["tables"], -r["records"], r["pattern"]),
    )

    glossary = []
    for term, meaning, example, test in ra.GLOSSARY:
        if not meaning:
            glossary.append({"section": term, "rows": []})
        elif glossary:
            glossary[-1]["rows"].append({"term": term, "meaning": meaning, "example": example, "test": test})

    return {
        "run_folder": digest.get("run_folder", ""),
        "generated_at": digest.get("generated_at", ""),
        "page_built_at": datetime.now().isoformat(timespec="seconds"),
        "sampling_note": digest.get("sampling_note", ""),
        "tiers": TIERS,
        "tables": tables,
        "errors": digest.get("errors", []),
        "across": across_rows,
        "glossary": glossary,
        "max_combos_in_page": MAX_COMBOS_IN_PAGE,
    }


def _csv_cell(value):
    if value is None or isinstance(value, (int, float)):
        return value
    text = str(value)
    # Excel runs a CSV cell starting with = + - @ as a formula; a plain number is fine
    if text[:1] in ("=", "@") or (text[:1] in ("+", "-") and not _is_number(text)):
        text = " " + text
    return text


def _is_number(text):
    try:
        float(text)
        return True
    except ValueError:
        return False


def _write_csv(path, header, rows):
    # utf-8-sig so Excel shows accented and non-English characters correctly
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(header)
        for row in rows:
            w.writerow([_csv_cell(v) for v in row])


def write_csvs(view, base_path):
    """Three flat CSV summaries, the same facts as the page and the workbook."""
    tier_labels = [t["label"] for t in TIERS]
    overview, columns = [], []
    for e in view["errors"]:
        overview.append([e["table"], "NOT ANALYSED", "YES", "", "", "", ""] + [""] * len(TIERS)
                        + ["", "", "", e["error"]])
    for t in view["tables"]:
        overview.append(
            [t["table"], "analysed", "YES" if t["reasons"] else "no",
             TIER_LABEL[t["worst"]] if t["columns"] else "",
             t["sample_rows"], t["distinct_combos"], len(t["columns"])]
            + [t["tier_counts"].get(tr["id"], 0) for tr in TIERS]
            + ["; ".join(t["reasons"]), "; ".join(t["notes"]) or "none", t["headline"], ""]
        )
        for c in t["columns"]:
            ex = c["examples"][0] if c["examples"] else None
            columns.append([
                t["table"], c["column"], c["risk"], c["records"], c["field_type"],
                "; ".join(f"{p['name']}={p['n']}" for p in c["patterns"]),
                "; ".join(c["observations"]),
                ex["pattern"] if ex else "", ex["s2"] if ex else "", ex["s3"] if ex else "",
                ", ".join(f"{k}={v}" for k, v in ex["pk"].items()) if ex else "",
                ex["sql"] if ex else "",
            ])

    paths = {
        "overview": base_path + "_overview.csv",
        "columns": base_path + "_columns.csv",
        "across": base_path + "_across_tables.csv",
    }
    _write_csv(paths["overview"],
               ["Table", "Status", "Manual check needed?", "Worst risk", "Sample rows", "Combos",
                "Mismatched columns"] + tier_labels
               + ["Why a manual check", "Warnings", "Headline", "Not analysed because"],
               overview)
    _write_csv(paths["columns"],
               ["Table", "Column", "Risk", "Records", "Field type", "Patterns", "Observations",
                "Example pattern", "Example s2", "Example s3", "Example key", "Lookup query"],
               columns)
    _write_csv(paths["across"],
               ["Risk", "Pattern", "Tables affected", "Columns affected", "Records",
                "Example s2", "Example s3", "Where (table.column)", "What it means"],
               [[TIER_LABEL[r["tier"]], r["pattern"], r["tables"], r["columns"], r["records"],
                 r["example"][0] if r["example"] else "", r["example"][1] if r["example"] else "",
                 "\n".join(r["where"]), r["desc"]]
                for r in view["across"]])
    return paths


def render_html(view, title):
    data = json.dumps(view, ensure_ascii=False)
    # the data sits inside a <script> element; "</" must not close it early
    data = data.replace("</", "<\\/")
    return (TEMPLATE
            .replace("__TITLE__", _html_escape(title))
            .replace("/*__DATA__*/null", data))


def _html_escape(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


def write_html(digest_path, html_path, recon_table="{table}", title=None):
    with open(digest_path, encoding="utf-8") as fh:
        digest = json.load(fh)
    view = build_view(digest, recon_table)
    run_name = os.path.basename(os.path.normpath(view["run_folder"])) or "recon run"
    html = render_html(view, title or f"Recon summary - {run_name}")
    out_dir = os.path.dirname(os.path.abspath(html_path))
    os.makedirs(out_dir, exist_ok=True)
    with open(html_path, "w", encoding="utf-8") as fh:
        fh.write(html)
    return view


def html_page_csv_base(html_path):
    return os.path.splitext(os.path.abspath(html_path))[0]


def main():
    ap = argparse.ArgumentParser(
        description="Build an HTML dashboard from a recon run folder or a recon_analyzer.py JSON digest",
        epilog="Options not listed here are passed through to recon_analyzer.py (run-folder mode only).",
    )
    ap.add_argument("input", help="A run folder (runs recon_analyzer.py first) or a digest .json file")
    ap.add_argument("-o", "--output",
                    help="Run folder: the .xlsx path, exactly as for recon_analyzer.py - the .html "
                         "is written beside it with the same name "
                         "(default: reports\\<folder>\\report_summary_<folder>.xlsx beside this script). "
                         "Digest: the .html path (default: beside the .json)")
    ap.add_argument("--recon-table", default="{table}",
                    help="Table name template for lookup queries, as in recon_analyzer.py")
    ap.add_argument("--title", help="Page title")
    ap.add_argument("--csv", action="store_true",
                    help="Also write three CSV summaries (overview, columns, across tables) beside the page")
    ap.add_argument("--json", dest="json_path",
                    help="Run folder: also keep the JSON digest at this path (normally it is temporary)")
    ap.add_argument("--summary", action="store_true",
                    help="Print what needs attention (not analysed, manual checks, warnings, top patterns)")
    args, passthrough = ap.parse_known_args()
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="backslashreplace")
        except (AttributeError, ValueError):
            pass

    src = args.input
    temp_dir = None
    try:
        if os.path.isdir(src):
            run_name = os.path.basename(os.path.abspath(src))
            xlsx = args.output or os.path.join(HERE, "reports", run_name, f"report_summary_{run_name}.xlsx")
            if not xlsx.lower().endswith(".xlsx"):
                xlsx += ".xlsx"
            html_path = os.path.splitext(os.path.abspath(xlsx))[0] + ".html"
            if args.json_path:
                digest_path = args.json_path
            else:
                # the digest only carries data from the analyzer to the page; it is not
                # a report, so it lives in a temporary folder and is removed afterwards
                temp_dir = tempfile.mkdtemp(prefix="recon_report_")
                digest_path = os.path.join(temp_dir, "digest.json")
            run_analyzer(src, xlsx, digest_path, args.recon_table, passthrough,
                         hide_digest_line=temp_dir is not None)
        elif os.path.isfile(src):
            if passthrough:
                ap.error(f"unrecognised arguments: {' '.join(passthrough)}")
            digest_path = src
            html_path = args.output or os.path.splitext(src)[0] + ".html"
        else:
            sys.exit(f"ERROR: {src} is neither a folder nor a file")

        if not html_path.lower().endswith((".html", ".htm")):
            html_path += ".html"
        try:
            view = write_html(digest_path, html_path, args.recon_table, args.title)
        except (OSError, ValueError, KeyError) as exc:
            sys.exit(f"ERROR: could not build the HTML page from {digest_path}: {exc}")
    finally:
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)
    print(f"HTML report: {os.path.abspath(html_path)}")

    if args.csv:
        try:
            for path in write_csvs(view, html_page_csv_base(html_path)).values():
                print(f"CSV summary: {path}")
        except OSError as exc:
            # the page is already written; a CSV that is open in Excel must not mask that
            print(f"WARNING: HTML saved but a CSV summary could not be written "
                  f"(is it open in Excel?): {exc}", file=sys.stderr)

    if args.summary:
        print_summary(view)
    else:
        manual = sum(1 for t in view["tables"] if t["reasons"])
        print(f"  {len(view['tables'])} table(s) analysed, {manual} need a manual check, "
              f"{len(view['errors'])} not analysed")


def run_analyzer(src, xlsx, digest_path, recon_table, passthrough, hide_digest_line):
    cmd = [sys.executable, os.path.join(HERE, "recon_analyzer.py"), src,
           "-o", xlsx, "--json", digest_path, "--recon-table", recon_table] + passthrough
    # UTF-8 with escapes: a table name or header the console codepage cannot print
    # must not stop the run when output is captured (a log file, Claude Code)
    env = dict(os.environ, PYTHONIOENCODING="utf-8:backslashreplace")
    # the digest is how we know the analyzer finished; never show a stale one
    before = os.path.getmtime(digest_path) if os.path.exists(digest_path) else None
    proc = subprocess.Popen(cmd, env=env, stdout=subprocess.PIPE)
    for raw in proc.stdout:
        line = raw.decode("utf-8", "backslashreplace").rstrip("\r\n")
        if hide_digest_line and line.startswith("JSON digest:"):
            continue  # a temporary file the user never needs to open
        print(line, flush=True)
    rc = proc.wait()
    if rc != 0:
        sys.exit(rc)
    if not os.path.exists(digest_path) or os.path.getmtime(digest_path) == before:
        sys.exit("ERROR: recon_analyzer.py finished but wrote no JSON digest - see the messages above")


def print_summary(view, max_reasons=3):
    tables, errors = view["tables"], view["errors"]
    manual = [t for t in tables if t["reasons"]]
    warned = [t for t in tables if t["notes"]]
    print("\nSUMMARY")
    print(f"  {len(tables)} table(s) analysed, {len(manual)} need a manual check, "
          f"{len(errors)} not analysed, {len(warned)} analysed with warnings")
    if errors:
        print("\n  NOT ANALYSED (nothing in these was compared):")
        for e in errors:
            print(f"    - {e['table']}: {e['error']}")
    if manual:
        print("\n  MANUAL CHECK NEEDED:")
        for t in manual:
            shown = "; ".join(t["reasons"][:max_reasons])
            more = f" (+{len(t['reasons']) - max_reasons} more)" if len(t["reasons"]) > max_reasons else ""
            print(f"    - {t['table']} [worst: {TIER_LABEL[t['worst']]}]: {shown}{more}")
    if warned:
        print("\n  ANALYSED WITH WARNINGS (not everything could be checked):")
        for t in warned:
            for note in t["notes"][:max_reasons]:
                print(f"    - {t['table']}: {note}")
            if len(t["notes"]) > max_reasons:
                print(f"    - {t['table']}: (+{len(t['notes']) - max_reasons} more - see the Errors sheet)")
    if view["across"]:
        print("\n  TOP PATTERNS ACROSS TABLES:")
        for r in view["across"][:5]:
            print(f"    - {TIER_LABEL[r['tier']]}: {r['pattern']} in {r['tables']} table(s), "
                  f"{r['columns']} column(s), {r['records']} record(s)")


TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__</title>
<style>
:root {
  color-scheme: light;
  --page: #f9f9f7; --surface: #fcfcfb; --surface-2: #f3f2ee;
  --ink: #0b0b0b; --ink-2: #52514e; --muted: #75736d;
  --grid: #e1e0d9; --axis: #c3c2b7; --ring: rgba(11,11,11,0.10);
  --accent: #2a78d6; --accent-wash: rgba(42,120,214,0.10);
  --mark: rgba(208,59,59,0.16); --mark-ink: #8f1f1f;
  --t-real: #d03b3b; --t-unsure: #ec835a; --t-text: #eda100; --t-check: #2a78d6; --t-low: #86b6ef;
  --shadow: 0 1px 2px rgba(0,0,0,.04), 0 4px 16px rgba(0,0,0,.04);
}
@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) {
    color-scheme: dark;
    --page: #0d0d0d; --surface: #1a1a19; --surface-2: #232321;
    --ink: #ffffff; --ink-2: #c3c2b7; --muted: #9a988f;
    --grid: #2c2c2a; --axis: #383835; --ring: rgba(255,255,255,0.10);
    --accent: #3987e5; --accent-wash: rgba(57,135,229,0.16);
    --mark: rgba(230,103,103,0.22); --mark-ink: #ffb4b4;
    --t-real: #e66767; --t-unsure: #ec835a; --t-text: #c98500; --t-check: #3987e5; --t-low: #184f95;
    --shadow: none;
  }
}
:root[data-theme="dark"] {
  color-scheme: dark;
  --page: #0d0d0d; --surface: #1a1a19; --surface-2: #232321;
  --ink: #ffffff; --ink-2: #c3c2b7; --muted: #9a988f;
  --grid: #2c2c2a; --axis: #383835; --ring: rgba(255,255,255,0.10);
  --accent: #3987e5; --accent-wash: rgba(57,135,229,0.16);
  --mark: rgba(230,103,103,0.22); --mark-ink: #ffb4b4;
  --t-real: #e66767; --t-unsure: #ec835a; --t-text: #c98500; --t-check: #3987e5; --t-low: #184f95;
  --shadow: none;
}
* { box-sizing: border-box; }
html { scroll-behavior: smooth; scroll-padding-top: 72px; }
body { margin: 0; background: var(--page); color: var(--ink);
  font: 14px/1.5 system-ui, -apple-system, "Segoe UI", Roboto, sans-serif; }
.wrap { max-width: 1240px; margin: 0 auto; padding-inline: 20px; padding-block: 28px 64px; }
h1 { font-size: 26px; line-height: 1.2; margin: 0 0 6px; letter-spacing: -0.01em; }
h2 { font-size: 18px; margin: 0 0 4px; }
h3 { font-size: 14px; margin: 18px 0 8px; color: var(--ink-2); text-transform: uppercase; letter-spacing: .04em; }
p { margin: 0; }
a { color: var(--accent); }
code, .mono { font-family: ui-monospace, "Cascadia Mono", Consolas, monospace; font-size: 12.5px; }
.sub { color: var(--ink-2); }
.muted { color: var(--muted); }
.eyebrow { font-size: 12px; font-weight: 600; letter-spacing: .06em; text-transform: uppercase; color: var(--accent); }
.meta { display: flex; flex-wrap: wrap; gap: 6px 20px; margin-top: 10px; color: var(--ink-2); font-size: 13px; }
.meta b { color: var(--ink); font-weight: 600; overflow-wrap: anywhere; }
.meta span, header p, .attn-item > div, .tname, .section-head > div { min-width: 0; overflow-wrap: anywhere; }
.card { background: var(--surface); border: 1px solid var(--ring); border-radius: 12px; box-shadow: var(--shadow); padding: 18px 20px; }
.section { margin-top: 28px; }
.section-head { display: flex; align-items: baseline; justify-content: space-between; gap: 12px; flex-wrap: wrap; margin-bottom: 12px; }

/* stat tiles */
.tiles { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 12px; margin-top: 22px; }
.tile { padding: 14px 16px; }
.tile .k { font-size: 12.5px; color: var(--ink-2); }
.tile .v { font-size: 30px; font-weight: 650; line-height: 1.15; margin-top: 4px; }
.tile .h { font-size: 12px; color: var(--muted); margin-top: 2px; }
.tile.alert { border-color: color-mix(in srgb, var(--t-real) 45%, transparent); }
.tile.alert .v { color: var(--t-real); }

/* risk pills and swatches */
.sw { display: inline-block; width: 10px; height: 10px; border-radius: 3px; flex: none; }
.pill { display: inline-flex; align-items: center; gap: 6px; padding: 2px 9px 2px 7px; border-radius: 999px;
  font-size: 12px; font-weight: 600; white-space: nowrap; background: var(--surface-2); color: var(--ink); border: 1px solid var(--ring); }
.t-real .sw, .sw.t-real { background: var(--t-real); }
.t-unsure .sw, .sw.t-unsure { background: var(--t-unsure); }
.t-text .sw, .sw.t-text { background: var(--t-text); }
.t-check .sw, .sw.t-check { background: var(--t-check); }
.t-low .sw, .sw.t-low { background: var(--t-low); }
.legend { display: flex; flex-wrap: wrap; gap: 6px 16px; font-size: 12.5px; color: var(--ink-2); }
.legend span { display: inline-flex; align-items: center; gap: 6px; }

/* part-to-whole strip */
.strip { display: flex; gap: 2px; height: 26px; margin: 14px 0 10px; }
.strip > div { border-radius: 4px; min-width: 4px; cursor: default; }
.bar-real { background: var(--t-real); } .bar-unsure { background: var(--t-unsure); }
.bar-text { background: var(--t-text); } .bar-check { background: var(--t-check); } .bar-low { background: var(--t-low); }
.strip-values { display: flex; flex-wrap: wrap; gap: 6px 22px; }
.strip-values div { display: flex; align-items: center; gap: 7px; font-size: 13px; color: var(--ink-2); }
.strip-values b { color: var(--ink); font-size: 15px; }

/* horizontal bar rows */
.grid2 { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr); gap: 16px; }
@media (max-width: 900px) { .grid2 { grid-template-columns: 1fr; } }
.hbars { margin-top: 12px; display: grid; gap: 7px; }
.hrow { display: grid; grid-template-columns: minmax(90px, 36%) 1fr 64px; align-items: center; gap: 10px;
  padding: 3px 6px; margin: 0 -6px; border-radius: 6px; cursor: pointer; text-decoration: none; color: inherit; }
.hrow:hover, .hrow:focus-visible { background: var(--accent-wash); outline: none; }
.hrow .name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 13px; }
.hrow .track { display: flex; gap: 2px; height: 14px; }
.hrow .track > div { height: 100%; border-radius: 0 4px 4px 0; }
.hrow .track > div:not(:last-child) { border-radius: 0; }
.hrow .track > div:first-child { border-radius: 0; }
.hrow .track > div:only-child, .hrow .track > div:last-child { border-radius: 0 4px 4px 0; }
.hrow .num { text-align: right; font-variant-numeric: tabular-nums; font-size: 12.5px; color: var(--ink-2); }
.hrow .flag { color: var(--t-real); font-weight: 700; margin-right: 4px; }
.axis-note { font-size: 12px; color: var(--muted); margin-top: 10px; }
.more { margin-top: 10px; }

button, .btn { font: inherit; font-size: 13px; color: var(--ink); background: var(--surface); border: 1px solid var(--ring);
  border-radius: 8px; padding: 5px 11px; cursor: pointer; }
button:hover { background: var(--surface-2); }
button:focus-visible, input:focus-visible { outline: 2px solid var(--accent); outline-offset: 1px; }

/* attention list */
.attn { display: grid; gap: 10px; }
.attn-item { display: grid; grid-template-columns: auto 1fr; gap: 4px 12px; padding: 12px 14px; border-radius: 10px;
  background: var(--surface); border: 1px solid var(--ring); border-left: 4px solid var(--t-unsure); }
.attn-item.err { border-left-color: var(--t-real); }
.attn-item .tag { font-size: 11px; font-weight: 700; letter-spacing: .05em; text-transform: uppercase; color: var(--ink-2); padding-top: 2px; }
.attn-item ul { margin: 4px 0 0; padding-left: 18px; color: var(--ink-2); }
.attn-item a { font-weight: 600; }

/* filter bar */
.filters { position: sticky; top: 0; z-index: 5; background: color-mix(in srgb, var(--page) 88%, transparent);
  backdrop-filter: blur(8px); -webkit-backdrop-filter: blur(8px); margin: 28px -20px 0; padding: 10px 20px;
  border-bottom: 1px solid var(--ring); display: flex; flex-wrap: wrap; gap: 8px 14px; align-items: center; }
.filters input[type=search] { font: inherit; font-size: 13px; padding: 6px 10px; border-radius: 8px; border: 1px solid var(--ring);
  background: var(--surface); color: var(--ink); width: 240px; max-width: 100%; }
.chk { display: inline-flex; align-items: center; gap: 6px; font-size: 13px; color: var(--ink-2); cursor: pointer; user-select: none; }
.chk input { accent-color: var(--accent); margin: 0; }
.filters .count { margin-left: auto; font-size: 12.5px; color: var(--muted); }

/* table sections */
details.tbl { background: var(--surface); border: 1px solid var(--ring); border-radius: 12px; margin-top: 12px; box-shadow: var(--shadow); }
details.tbl > summary { list-style: none; cursor: pointer; padding: 14px 18px; display: grid;
  grid-template-columns: 1fr auto; gap: 6px 16px; align-items: center; }
details.tbl > summary::-webkit-details-marker { display: none; }
details.tbl > summary:focus-visible { outline: 2px solid var(--accent); border-radius: 12px; }
.tname { font-weight: 650; font-size: 15px; display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
.tname .chev { display: inline-block; transition: transform .15s; color: var(--muted); }
details[open] .chev { transform: rotate(90deg); }
.manual { font-size: 11px; font-weight: 700; letter-spacing: .04em; padding: 2px 8px; border-radius: 999px;
  background: color-mix(in srgb, var(--t-text) 22%, transparent); color: var(--ink); }
.tstats { display: flex; gap: 14px; color: var(--ink-2); font-size: 12.5px; white-space: nowrap; }
.tstats b { color: var(--ink); }
.mini { display: flex; gap: 2px; height: 8px; width: 160px; grid-column: 1 / -1; }
.mini > div { border-radius: 2px; }
.tbody { padding: 0 18px 18px; border-top: 1px solid var(--grid); }
.headline { margin-top: 14px; color: var(--ink-2); }
.reasons { margin: 10px 0 0; padding: 10px 14px 10px 30px; border-radius: 8px; background: color-mix(in srgb, var(--t-text) 12%, transparent); }
.notes { margin: 10px 0 0; padding: 10px 14px 10px 30px; border-radius: 8px; background: var(--surface-2); color: var(--ink-2); }

.tablewrap { overflow-x: auto; border: 1px solid var(--grid); border-radius: 8px; }
table.data { border-collapse: collapse; width: 100%; font-size: 13px; }
table.data th { text-align: left; font-weight: 600; font-size: 12px; color: var(--ink-2); background: var(--surface-2);
  padding: 8px 10px; border-bottom: 1px solid var(--grid); white-space: nowrap; }
table.data td { padding: 9px 10px; border-bottom: 1px solid var(--grid); vertical-align: top; }
table.data tr:last-child td { border-bottom: 0; }
table.data td.n { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
.chip { display: inline-flex; align-items: center; gap: 5px; font-family: ui-monospace, "Cascadia Mono", Consolas, monospace;
  font-size: 12px; padding: 1px 7px; border-radius: 6px; background: var(--surface-2); border: 1px solid var(--ring);
  margin: 0 4px 4px 0; cursor: help; white-space: nowrap; }
.chip i { font-style: normal; color: var(--muted); }
.obs { margin: 0; padding-left: 16px; color: var(--ink-2); font-size: 12.5px; }
.ex { display: grid; gap: 6px; min-width: 300px; }
.ex-row { display: grid; grid-template-columns: auto 1fr; gap: 2px 8px; padding: 6px 8px; border-radius: 6px; background: var(--surface-2); }
.ex-row .lbl { font-size: 11px; font-weight: 700; color: var(--muted); padding-top: 1px; }
.ex-row .val { font-family: ui-monospace, "Cascadia Mono", Consolas, monospace; font-size: 12.5px; word-break: break-all; }
.ex-row mark { background: var(--mark); color: var(--mark-ink); border-radius: 3px; padding: 0 1px; }
.ex-row .ws { color: var(--mark-ink); opacity: .8; }
.ex-row .empty { font-style: italic; color: var(--muted); font-family: system-ui, sans-serif; }
.ex-foot { grid-column: 1 / -1; display: flex; align-items: center; gap: 8px; flex-wrap: wrap; font-size: 11.5px; color: var(--muted); margin-top: 2px; }
.ex-foot button { font-size: 11.5px; padding: 1px 7px; border-radius: 6px; }

/* tooltip */
#tip { position: fixed; z-index: 50; pointer-events: none; max-width: 360px; padding: 8px 10px; border-radius: 8px;
  background: var(--ink); color: var(--page); font-size: 12.5px; line-height: 1.4; box-shadow: 0 6px 24px rgba(0,0,0,.2);
  opacity: 0; transition: opacity .08s; }
#tip.on { opacity: 1; }
#tip b { font-weight: 650; }

.gloss-sec { margin-top: 18px; }
.empty-state { padding: 26px; text-align: center; color: var(--muted); }
.ok { color: var(--ink-2); }
footer { margin-top: 40px; font-size: 12px; color: var(--muted); }

@media (max-width: 640px) {
  .wrap { padding-inline: 16px; }
  .filters { margin-inline: -16px; padding-inline: 16px; }
  details.tbl > summary { grid-template-columns: 1fr; }
  .tstats { flex-wrap: wrap; white-space: normal; }
  .hrow { grid-template-columns: minmax(80px, 40%) 1fr 44px; }
  h1 { font-size: 22px; }
}
@media print {
  .filters, .more, button { display: none !important; }
  body { background: #fff; }
  .card, details.tbl { box-shadow: none; break-inside: avoid; }
}
</style>
</head>
<body>
<div class="wrap" id="app"></div>
<div id="tip" role="tooltip"></div>
<noscript><p style="padding:20px">This report needs JavaScript to display. The same content is in the Excel workbook.</p></noscript>
<script id="data" type="application/json">/*__DATA__*/null</script>
<script>
(function () {
  "use strict";
  var V = JSON.parse(document.getElementById("data").textContent);
  var TIERS = V.tiers, TIER = {};
  TIERS.forEach(function (t) { TIER[t.id] = t; });
  var app = document.getElementById("app");

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function fmt(n) { return Number(n || 0).toLocaleString(); }
  function plural(n, w) { return fmt(n) + " " + w + (n === 1 ? "" : "s"); }
  function slug(s) { return "t-" + String(s).replace(/[^A-Za-z0-9_-]/g, "_"); }
  function pill(tier) { return '<span class="pill t-' + tier + '"><span class="sw"></span>' + esc(TIER[tier].label) + "</span>"; }
  function sw(tier) { return '<span class="sw t-' + tier + '"></span>'; }

  // show exactly where two values differ; spaces inside the difference are made visible
  function showWs(s) { return esc(s).replace(/ /g, '<span class="ws">\u00b7</span>').replace(/\t/g, '<span class="ws">\u2192</span>'); }
  // control characters would otherwise render as nothing at all; show them as \x07,
  // the same way the workbook does
  function showCtl(s) {
    return String(s).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]/g, function (c) {
      return "\\x" + ("0" + c.charCodeAt(0).toString(16)).slice(-2);
    });
  }
  function diffPair(a, b) {
    a = showCtl(a == null ? "" : a); b = showCtl(b == null ? "" : b);
    if (a === "" || b === "") return [a === "" ? '<span class="empty">(empty)</span>' : "<mark>" + showWs(a) + "</mark>",
                                      b === "" ? '<span class="empty">(empty)</span>' : "<mark>" + showWs(b) + "</mark>"];
    // character-level diff (longest common subsequence) for normal-length values,
    // so 1,234.50 vs 1234.5 marks only the comma and the trailing zero
    if (a.length * b.length <= 90000) {
      var n = a.length, m = b.length, L = [], i, j;
      for (i = 0; i <= n; i++) { L.push(new Uint16Array(m + 1)); }
      for (i = n - 1; i >= 0; i--) for (j = m - 1; j >= 0; j--)
        L[i][j] = a[i] === b[j] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
      var ka = [], kb = []; i = 0; j = 0;
      while (i < n || j < m) {
        if (i < n && j < m && a[i] === b[j]) { ka.push(1); kb.push(1); i++; j++; }
        else if (j >= m || (i < n && L[i + 1][j] >= L[i][j + 1])) { ka.push(0); i++; }
        else { kb.push(0); j++; }
      }
      function paint(x, keep) {
        var out = "", run = "";
        for (var k = 0; k < x.length; k++) {
          if (keep[k]) { if (run) { out += "<mark>" + showWs(run) + "</mark>"; run = ""; } out += esc(x[k]); }
          else run += x[k];
        }
        return out + (run ? "<mark>" + showWs(run) + "</mark>" : "");
      }
      return [paint(a, ka), paint(b, kb)];
    }
    var p = 0, max = Math.min(a.length, b.length);
    while (p < max && a[p] === b[p]) p++;
    var s = 0;
    while (s < max - p && a[a.length - 1 - s] === b[b.length - 1 - s]) s++;
    function part(x) {
      if (x === "") return '<span class="empty">(empty)</span>';
      var mid = x.slice(p, x.length - s);
      return esc(x.slice(0, p)) + (mid ? "<mark>" + showWs(mid) + "</mark>" : "") + esc(x.slice(x.length - s));
    }
    return [part(a), part(b)];
  }

  var tables = V.tables, errors = V.errors;
  // readable anchors (#t-orders_recon); a suffix keeps two similar names apart
  var usedIds = {};
  tables.forEach(function (t) {
    var id = slug(t.table), n = 2;
    while (usedIds[id]) id = slug(t.table) + "-" + n++;
    usedIds[id] = 1; t.id = id;
  });
  var manual = tables.filter(function (t) { return t.reasons.length; });
  var colTotals = {}; TIERS.forEach(function (t) { colTotals[t.id] = 0; });
  var totalCols = 0, totalRows = 0;
  tables.forEach(function (t) {
    totalRows += t.sample_rows;
    t.columns.forEach(function (c) { colTotals[c.tier]++; totalCols++; });
  });
  var realTables = tables.filter(function (t) { return (t.tier_counts.real || 0) > 0; }).length;

  var runName = (V.run_folder.split(/[\\/]/).filter(Boolean).pop()) || "recon run";
  var h = [];

  // ---------- header ----------
  h.push('<header><div class="eyebrow">Recon mismatch summary</div><h1>' + esc(runName) + "</h1>");
  h.push('<div class="meta"><span>Run folder <b class="mono">' + esc(V.run_folder) + "</b></span>" +
    "<span>Analysed <b>" + esc(V.generated_at.replace("T", " ")) + "</b></span></div>");
  if (V.sampling_note) h.push('<p class="muted" style="margin-top:6px;font-size:12.5px">Note: ' + esc(V.sampling_note) + ". Record counts are sample counts, not production totals.</p>");
  h.push("</header>");

  // ---------- stat tiles ----------
  h.push('<div class="tiles">');
  h.push(tile("Tables analysed", fmt(tables.length), plural(totalRows, "sample row")));
  h.push(tile("Need a manual check", fmt(manual.length), "of " + fmt(tables.length) + " tables", manual.length > 0));
  h.push(tile("Not analysed", fmt(errors.length), errors.length ? "never read these as clean" : "every table was read", errors.length > 0));
  h.push(tile("Tables with real value diffs", fmt(realTables), plural(colTotals.real, "column") + " affected", realTables > 0));
  h.push(tile("Mismatched columns", fmt(totalCols), "across all tables"));
  h.push("</div>");
  function tile(k, v, hint, alert) {
    return '<div class="card tile' + (alert ? " alert" : "") + '"><div class="k">' + esc(k) + '</div><div class="v">' + v + '</div><div class="h">' + esc(hint) + "</div></div>";
  }

  // ---------- risk mix ----------
  h.push('<section class="section card"><div class="section-head"><div><h2>Mismatched columns by risk</h2>' +
    '<p class="sub">Each column is ranked by its worst record. Read left to right: the most urgent is first.</p></div></div>');
  if (totalCols) {
    h.push('<div class="strip" role="img" aria-label="Columns by risk">');
    TIERS.forEach(function (t) {
      var n = colTotals[t.id]; if (!n) return;
      h.push('<div class="bar-' + t.id + '" style="flex:' + n + '" data-tip="' + esc("<b>" + esc(t.label) + "</b><br>" + plural(n, "column") + " (" + Math.round(100 * n / totalCols) + "%)") + '"></div>');
    });
    h.push('</div><div class="strip-values">');
    TIERS.forEach(function (t) {
      h.push("<div>" + sw(t.id) + "<b>" + fmt(colTotals[t.id]) + "</b>" + esc(t.label) + "</div>");
    });
    h.push("</div>");
  } else {
    h.push('<p class="empty-state">No mismatched columns were found in any analysed table.</p>');
  }
  h.push("</section>");

  // ---------- attention list ----------
  if (errors.length || manual.length) {
    h.push('<section class="section"><div class="section-head"><div><h2>Needs your attention</h2>' +
      '<p class="sub">Not analysed first, then tables where something measurable looks odd.</p></div></div><div class="attn">');
    errors.forEach(function (e) {
      h.push('<div class="attn-item err"><span class="tag">Not analysed</span><div><b>' + esc(e.table) + "</b><div class=\"sub\">" + esc(e.error) + "</div></div></div>");
    });
    manual.forEach(function (t) {
      h.push('<div class="attn-item"><span class="tag">Manual check</span><div><a href="#' + t.id + '" data-open="' + t.id + '">' + esc(t.table) + "</a> " + pill(t.worst) +
        "<ul>" + t.reasons.map(function (r) { return "<li>" + esc(r) + "</li>"; }).join("") + "</ul></div></div>");
    });
    h.push("</div></section>");
  }

  // ---------- two charts ----------
  h.push('<div class="grid2 section">');
  h.push('<section class="card"><h2>Tables, worst first</h2><p class="sub">Mismatched columns per table, split by risk. Click a row to open it.</p>' +
    '<div class="legend" style="margin-top:10px">' + TIERS.map(function (t) { return "<span>" + sw(t.id) + esc(t.label) + "</span>"; }).join("") + "</div>" +
    '<div class="hbars" id="tableBars"></div><div class="axis-note"><span class="flag" style="color:var(--t-real);font-weight:700">!</span> = manual check needed. Number = mismatched columns.</div>' +
    '<button class="more" id="moreTables" hidden></button></section>');
  h.push('<section class="card"><h2>Patterns across tables</h2><p class="sub">The same behaviour in many tables is usually one fix, not many.</p>' +
    '<div class="legend" style="margin-top:10px">' + TIERS.map(function (t) { return "<span>" + sw(t.id) + esc(t.label) + "</span>"; }).join("") + "</div>" +
    '<div class="hbars" id="patternBars"></div><div class="axis-note">Bar = sample records showing the pattern. Label = tables affected.</div>' +
    '<button class="more" id="morePatterns" hidden></button></section>');
  h.push("</div>");

  // ---------- filters + per-table detail ----------
  h.push('<div class="filters" role="search">' +
    '<input type="search" id="q" placeholder="Filter tables or columns\u2026" aria-label="Filter tables or columns">' +
    '<label class="chk"><input type="checkbox" id="onlyManual"> Manual check only</label>' +
    TIERS.map(function (t) { return '<label class="chk"><input type="checkbox" class="tierf" value="' + t.id + '" checked>' + sw(t.id) + esc(t.label) + "</label>"; }).join("") +
    '<button id="expandAll">Expand all</button><button id="collapseAll">Collapse all</button>' +
    '<span class="count" id="count"></span></div>');
  h.push('<section class="section" id="tablesSec"><h2>Table detail</h2><p class="sub">Columns worst first, with examples. The changed part of each value is highlighted; \u00b7 marks a space.</p><div id="tableList"></div></section>');

  // ---------- across-tables table view ----------
  h.push('<section class="section"><div class="section-head"><div><h2>Across tables \u2014 full list</h2><p class="sub">Every pattern seen in the run. Hover a pattern name for its meaning.</p></div></div>' +
    '<div class="tablewrap card" style="padding:0"><table class="data"><thead><tr><th>Risk</th><th>Pattern</th><th>Tables</th><th>Columns</th><th>Records</th><th>Example (s2 \u2192 s3)</th><th>Where</th></tr></thead><tbody>');
  V.across.forEach(function (r) {
    var ex = r.example ? diffPair(r.example[0], r.example[1]) : null;
    // every table.column, one per line; a long list folds so the table stays scannable
    var whereLines = r.where.map(function (w) { return "<div>" + esc(w) + "</div>"; }).join("");
    var where = r.where.length > 15
      ? "<details><summary>" + fmt(r.where.length) + " columns \u2014 show all</summary>" + whereLines + "</details>"
      : whereLines;
    h.push("<tr><td>" + pill(r.tier) + '</td><td><span class="chip" data-tip="' + esc(esc(r.desc)) + '">' + esc(r.pattern) + '</span></td><td class="n">' + fmt(r.tables) +
      '</td><td class="n">' + fmt(r.columns) + '</td><td class="n">' + fmt(r.records) + "</td><td>" +
      (ex ? '<span class="mono">' + ex[0] + ' <span class="muted">\u2192</span> ' + ex[1] + "</span>" : "") + '</td><td class="muted" style="font-size:12.5px">' + where + "</td></tr>");
  });
  if (!V.across.length) h.push('<tr><td colspan="7" class="empty-state">No patterns found.</td></tr>');
  h.push("</tbody></table></div></section>");

  // ---------- glossary ----------
  h.push('<section class="section"><details class="tbl"><summary><div class="tname"><span class="chev">\u25b6</span>What the terms mean</div></summary><div class="tbody">');
  V.glossary.forEach(function (g) {
    h.push('<div class="gloss-sec"><h3>' + esc(g.section) + '</h3><div class="tablewrap"><table class="data"><thead><tr><th>Term</th><th>What it means</th><th>Example</th><th>How it is decided</th></tr></thead><tbody>');
    g.rows.forEach(function (r) {
      h.push('<tr><td><code>' + esc(r.term) + "</code></td><td>" + esc(r.meaning) + '</td><td class="mono">' + esc(r.example) + '</td><td class="muted">' + esc(r.test) + "</td></tr>");
    });
    h.push("</tbody></table></div></div>");
  });
  h.push("</div></details></section>");
  h.push("<footer>Generated by recon_html_report.py from the recon_analyzer.py JSON digest \u00b7 page built " + esc(V.page_built_at.replace("T", " ")) +
    ". The Excel workbook from the same run is the complete record.</footer>");

  app.innerHTML = h.join("");

  // ---------- per-table detail rendering ----------
  function tableSection(t) {
    var s = [];
    var total = t.columns.length;
    s.push('<details class="tbl" id="' + t.id + '" data-name="' + esc(t.table.toLowerCase()) + '"><summary>');
    s.push('<div class="tname"><span class="chev">\u25b6</span>' + esc(t.table) + " " + pill(t.worst) + (t.reasons.length ? '<span class="manual">MANUAL CHECK</span>' : "") + "</div>");
    s.push('<div class="tstats"><span><b>' + fmt(t.sample_rows) + "</b> rows</span><span><b>" + fmt(t.distinct_combos) + "</b> combos</span><span><b>" + fmt(total) + "</b> columns</span></div>");
    if (total) {
      s.push('<div class="mini" aria-hidden="true">');
      TIERS.forEach(function (tr) { var n = t.tier_counts[tr.id] || 0; if (n) s.push('<div class="bar-' + tr.id + '" style="flex:' + n + '"></div>'); });
      s.push("</div>");
    }
    s.push('</summary><div class="tbody"></div></details>');
    return s.join("");
  }

  function tableBody(t) {
    var s = [];
    s.push('<p class="headline">' + esc(t.headline) + "</p>");
    if (t.pk_cols.length) s.push('<p class="muted" style="margin-top:4px;font-size:12.5px">Primary key: <code>' + t.pk_cols.map(esc).join(", ") + "</code></p>");
    if (t.reasons.length) s.push('<ul class="reasons"><b style="margin-left:-16px">Why a manual check is needed</b>' + t.reasons.map(function (r) { return "<li>" + esc(r) + "</li>"; }).join("") + "</ul>");
    if (t.notes.length) s.push('<ul class="notes"><b style="margin-left:-16px">Warnings</b>' + t.notes.map(function (r) { return "<li>" + esc(r) + "</li>"; }).join("") + "</ul>");

    s.push('<h3>Columns</h3><div class="tablewrap"><table class="data"><thead><tr><th>Column</th><th>Risk</th><th>Records</th><th>Patterns</th><th>Observations</th><th>Examples (s2 / s3)</th></tr></thead><tbody>');
    t.columns.forEach(function (c) {
      s.push('<tr data-col="' + esc(c.column.toLowerCase()) + '" data-tier="' + c.tier + '"><td><b class="mono">' + esc(c.column) + "</b>" + (c.field_type === "text" ? '<div class="muted" style="font-size:11.5px">free text</div>' : "") + "</td>");
      s.push("<td>" + pill(c.tier) + '</td><td class="n">' + fmt(c.records) + "</td><td>");
      c.patterns.forEach(function (p) {
        s.push('<span class="chip t-' + p.tier + '" data-tip="' + esc("<b>" + esc(TIER[p.tier].label) + "</b><br>" + esc(p.desc)) + '"><span class="sw"></span>' + esc(p.name) + " <i>\u00d7" + fmt(p.n) + "</i></span>");
      });
      s.push('</td><td><ul class="obs">' + c.observations.map(function (o) { return "<li>" + esc(o) + "</li>"; }).join("") + '</ul></td><td><div class="ex">');
      c.examples.forEach(function (e) {
        var d = diffPair(e.s2, e.s3);
        var pk = Object.keys(e.pk).map(function (k) { return esc(k) + "=" + esc(e.pk[k]); }).join(", ");
        s.push('<div class="ex-row"><span class="lbl">s2</span><span class="val">' + d[0] + '</span><span class="lbl">s3</span><span class="val">' + d[1] + "</span>" +
          '<div class="ex-foot"><code>' + esc(e.pattern) + "</code>" + (pk ? "<span>" + pk + "</span>" : "") +
          (e.sql ? '<button data-sql="' + esc(e.sql) + '" data-tip="' + esc(esc(e.sql)) + '">Copy SELECT</button>' : "") + "</div></div>");
      });
      s.push("</div></td></tr>");
    });
    if (!t.columns.length) s.push('<tr><td colspan="6" class="empty-state">No mismatched columns.</td></tr>');
    s.push("</tbody></table></div>");

    if (t.combos.length) {
      s.push("<h3>Breakdown by mismatch_columns combination</h3><div class=\"tablewrap\"><table class=\"data\"><thead><tr><th>Combination</th><th>Risk</th><th>Records</th><th>Patterns per column</th></tr></thead><tbody>");
      t.combos.forEach(function (cb) {
        var per = Object.keys(cb.per_column).map(function (col) {
          var pats = cb.per_column[col];
          return '<div><code>' + esc(col) + '</code> <span class="muted">' + Object.keys(pats).map(function (p) { return esc(p) + " \u00d7" + fmt(pats[p]); }).join(", ") + "</span></div>";
        }).join("");
        s.push('<tr><td class="mono">' + esc(cb.combo) + "</td><td>" + pill(cb.tier) + '</td><td class="n">' + fmt(cb.records) + '</td><td style="font-size:12.5px">' + per + "</td></tr>");
      });
      s.push("</tbody></table></div>");
      if (t.combos_hidden) s.push('<p class="muted" style="margin-top:8px">+' + fmt(t.combos_hidden) + " more combinations are listed in the Excel workbook.</p>");
    }
    return s.join("");
  }

  var list = document.getElementById("tableList");
  list.innerHTML = tables.length ? tables.map(tableSection).join("") : '<div class="card empty-state">No tables were analysed.</div>';
  // bodies are built on first open so a 200-table run opens instantly
  var byId = {};
  tables.forEach(function (t) { byId[t.id] = t; });
  list.addEventListener("toggle", function (ev) {
    var d = ev.target;
    if (d.open && !d.dataset.built) { d.querySelector(".tbody").innerHTML = tableBody(byId[d.id]); d.dataset.built = "1"; applyRowFilter(d); }
  }, true);

  // ---------- bar charts ----------
  function tableBars(limit) {
    var rows = visibleTables().slice(0, limit);
    var max = 1; tables.forEach(function (t) { max = Math.max(max, t.columns.length); });
    document.getElementById("tableBars").innerHTML = rows.map(function (t) {
      var segs = TIERS.map(function (tr) {
        var n = t.tier_counts[tr.id] || 0;
        return n ? '<div class="bar-' + tr.id + '" style="width:calc(' + (100 * n / max) + '% - 2px)"></div>' : "";
      }).join("");
      var tip = "<b>" + esc(t.table) + "</b><br>" + TIERS.filter(function (tr) { return t.tier_counts[tr.id]; }).map(function (tr) { return esc(tr.label) + ": " + t.tier_counts[tr.id]; }).join("<br>") +
        (t.reasons.length ? "<br><br>Manual check: " + t.reasons.length + " reason(s)" : "");
      return '<a class="hrow" href="#' + t.id + '" data-open="' + t.id + '" data-tip="' + esc(tip) + '"><span class="name">' +
        (t.reasons.length ? '<span class="flag">!</span>' : "") + esc(t.table) + '</span><span class="track">' + segs + '</span><span class="num">' + fmt(t.columns.length) + "</span></a>";
    }).join("") || '<p class="muted">No tables match the filters.</p>';
    var btn = document.getElementById("moreTables"), n = visibleTables().length;
    btn.hidden = n <= 15;
    btn.textContent = limit >= n ? "Show top 15" : "Show all " + fmt(n) + " tables";
  }
  var tableLimit = 15;
  document.getElementById("moreTables").onclick = function () { tableLimit = tableLimit >= visibleTables().length ? 15 : 1e9; tableBars(tableLimit); };

  var patLimit = 15;
  function patternBars() {
    var rows = V.across.slice(0, patLimit), max = 1;
    V.across.forEach(function (r) { max = Math.max(max, r.records); });
    document.getElementById("patternBars").innerHTML = rows.map(function (r) {
      var tip = "<b>" + esc(r.pattern) + "</b> \u00b7 " + esc(TIER[r.tier].label) + "<br>" + esc(r.desc) + "<br><br>" + plural(r.records, "record") + " in " + plural(r.columns, "column") + " of " + plural(r.tables, "table");
      return '<div class="hrow" tabindex="0" style="cursor:default" data-tip="' + esc(tip) + '"><span class="name mono">' + esc(r.pattern) + '</span><span class="track"><div class="bar-' + r.tier + '" style="width:' + (100 * r.records / max) + '%;border-radius:0 4px 4px 0"></div></span><span class="num">' + plural(r.tables, "table") + "</span></div>";
    }).join("") || '<p class="muted">No patterns found.</p>';
    var btn = document.getElementById("morePatterns");
    btn.hidden = V.across.length <= 15;
    btn.textContent = patLimit >= V.across.length ? "Show top 15" : "Show all " + fmt(V.across.length) + " patterns";
  }
  document.getElementById("morePatterns").onclick = function () { patLimit = patLimit >= V.across.length ? 15 : 1e9; patternBars(); };

  // ---------- filters ----------
  var q = document.getElementById("q"), onlyManual = document.getElementById("onlyManual");
  function activeTiers() { var a = {}; document.querySelectorAll(".tierf").forEach(function (c) { if (c.checked) a[c.value] = 1; }); return a; }
  function tableMatches(t) {
    var term = q.value.trim().toLowerCase(), tiers = activeTiers();
    if (onlyManual.checked && !t.reasons.length) return false;
    var cols = t.columns.filter(function (c) { return tiers[c.tier]; });
    if (t.columns.length && !cols.length) return false;
    if (!term) return true;
    return t.table.toLowerCase().indexOf(term) >= 0 || cols.some(function (c) { return c.column.toLowerCase().indexOf(term) >= 0; });
  }
  function visibleTables() { return tables.filter(tableMatches); }
  function applyRowFilter(d) {
    var term = q.value.trim().toLowerCase(), tiers = activeTiers(), t = byId[d.id];
    var nameHit = !term || t.table.toLowerCase().indexOf(term) >= 0;
    d.querySelectorAll("tr[data-tier]").forEach(function (tr) {
      tr.hidden = !tiers[tr.dataset.tier] || (!nameHit && tr.dataset.col.indexOf(term) < 0);
    });
  }
  function applyFilters() {
    var shown = 0;
    tables.forEach(function (t) {
      var d = document.getElementById(t.id), ok = tableMatches(t);
      d.hidden = !ok; if (ok) shown++;
      if (d.dataset.built) applyRowFilter(d);
    });
    document.getElementById("count").textContent = "Showing " + fmt(shown) + " of " + plural(tables.length, "table");
    tableBars(tableLimit);
  }
  q.addEventListener("input", applyFilters);
  onlyManual.addEventListener("change", applyFilters);
  document.querySelectorAll(".tierf").forEach(function (c) { c.addEventListener("change", applyFilters); });
  document.getElementById("expandAll").onclick = function () { list.querySelectorAll("details.tbl:not([hidden])").forEach(function (d) { d.open = true; }); };
  document.getElementById("collapseAll").onclick = function () { list.querySelectorAll("details.tbl").forEach(function (d) { d.open = false; }); };
  // table bodies are built on open, so a print or PDF would otherwise show empty sections
  window.addEventListener("beforeprint", function () {
    list.querySelectorAll("details.tbl:not([hidden])").forEach(function (d) {
      d.open = true;
      if (!d.dataset.built) { d.querySelector(".tbody").innerHTML = tableBody(byId[d.id]); d.dataset.built = "1"; applyRowFilter(d); }
    });
  });

  // links from the attention list and the bar chart open the table they point at
  document.addEventListener("click", function (ev) {
    var a = ev.target.closest("[data-open]");
    if (a) { var d = document.getElementById(a.dataset.open); if (d) { d.hidden = false; d.open = true; } }
    var b = ev.target.closest("button[data-sql]");
    if (b) copy(b.dataset.sql, b);
  });
  function copy(text, btn) {
    function done() { var o = btn.textContent; btn.textContent = "Copied"; setTimeout(function () { btn.textContent = o; }, 1200); }
    function fallback() {
      var ta = document.createElement("textarea"); ta.value = text; ta.style.position = "fixed"; ta.style.opacity = "0";
      document.body.appendChild(ta); ta.select();
      try { document.execCommand("copy"); done(); } catch (e) { window.prompt("Copy this query:", text); }
      document.body.removeChild(ta);
    }
    if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(text).then(done, fallback); else fallback();
  }

  // ---------- tooltip (hover and keyboard focus) ----------
  var tip = document.getElementById("tip"), tipFor = null;
  function showTip(el, x, y) {
    if (tipFor !== el) { tip.innerHTML = el.getAttribute("data-tip"); tip.classList.add("on"); tipFor = el; }
    var r = tip.getBoundingClientRect(), pad = 12;
    var left = Math.min(window.innerWidth - r.width - 8, Math.max(8, x + pad));
    var top = y + pad + r.height > window.innerHeight ? y - r.height - pad : y + pad;
    tip.style.left = left + "px"; tip.style.top = Math.max(8, top) + "px";
  }
  function hideTip() { tip.classList.remove("on"); tipFor = null; }
  document.addEventListener("mousemove", function (ev) {
    var el = ev.target.closest && ev.target.closest("[data-tip]");
    if (el) showTip(el, ev.clientX, ev.clientY); else if (tipFor) hideTip();
  });
  document.addEventListener("focusin", function (ev) {
    var el = ev.target.closest && ev.target.closest("[data-tip]");
    if (el) { var r = el.getBoundingClientRect(); showTip(el, r.left, r.bottom); } else hideTip();
  });
  document.addEventListener("scroll", hideTip, true);

  applyFilters();
  patternBars();
  if (location.hash) { var d0 = document.getElementById(location.hash.slice(1)); if (d0 && d0.tagName === "DETAILS") d0.open = true; }
})();
</script>
</body>
</html>
"""


if __name__ == "__main__":
    main()
