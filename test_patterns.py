"""Coverage matrix: every (column, s2, s3) with the pattern it must produce.

Run:  python test_patterns.py      (exit code 0 = all pass)
"""
import io
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from recon_analyzer import classify_pair, column_observations, pattern_risk  # noqa: E402

CASES = [
    # ---- INTEGER ----
    ("INTEGER", "acct_no", "000123", "123", "leading_zero_diff"),
    ("INTEGER", "qty", "1,000", "1000", "numeric_thousands_sep_diff"),
    ("INTEGER", "cnt", "5", "5.0", "numeric_scale_diff"),
    ("INTEGER", "cnt", "+5", "5", "numeric_sign_format_diff"),
    ("INTEGER", "amt_owed", "(500)", "-500", "accounting_negative_diff"),
    ("INTEGER", "qty", "5", "7", "numeric_value_diff"),
    ("INTEGER", "big", "1000000", "1E+06", "numeric_scientific_notation"),
    ("INTEGER", "amount", "500-", "-500", "numeric_sign_format_diff"),

    # ---- DECIMAL ----
    ("DECIMAL", "balance", "10.00", "10.0", "numeric_scale_diff"),
    ("DECIMAL", "balance", "1,234.50", "1234.50", "numeric_thousands_sep_diff"),
    ("DECIMAL", "price", "99.456789", "99.46", "numeric_rounding"),
    ("DECIMAL", "price", "100.00", "150.00", "numeric_value_diff"),
    ("DECIMAL", "fee", "$10.00", "10.00", "currency_symbol_diff"),
    ("DECIMAL", "amount", "500.00-", "500.0-", "numeric_scale_diff"),
    ("DECIMAL", "rate", "10.00%", "10.0%", "numeric_scale_diff"),

    # ---- FLOAT / DOUBLE ----
    ("FLOAT", "rate", "0.3", "0.30000000000000004", "float_precision_noise"),
    ("FLOAT", "rate", "2.675", "2.6749999999999998", "float_precision_noise"),
    ("FLOAT", "val", "1.0", "1", "numeric_scale_diff"),
    ("FLOAT", "rate", "0.1", "0.10000000000000001", "float_precision_noise"),
    ("DOUBLE", "val", "Infinity", "", "missing_in_s3"),
    ("DOUBLE", "val", "inf", "1.0", "infinity_value"),
    ("DOUBLE", "val", "NaN", "", "null_representation_diff"),

    # ---- BOOLEAN ----
    ("BOOLEAN", "active_flag", "Y", "true", "boolean_format_diff"),
    ("BOOLEAN", "active_flag", "Y", "N", "boolean_value_diff"),
    ("BOOLEAN", "is_closed", "1", "0", "boolean_value_diff"),
    ("BOOLEAN", "is_closed", "1", "true", "boolean_format_diff"),
    ("BOOLEAN", "deleted_ind", "T", "F", "boolean_value_diff"),
    ("BOOLEAN", "status_yn", "yes", "Y", "boolean_format_diff"),

    # ---- DATE / TIMESTAMP ----
    ("DATE", "open_date", "2024-01-05", "20240105", "date_format_diff"),
    ("DATE", "open_date", "2024-02-10", "2024-02-10 00:00:00", "date_format_diff"),
    ("DATE", "txn_ts", "2024-01-02T08:30:00", "2024-01-02 08:30:00", "date_format_diff"),
    ("DATE", "txn_ts", "2024-01-01 00:00:00.000", "2024-01-01 00:00:00", "date_format_diff"),
    ("DATE", "close_date", "2024-03-01", "2024-03-02", "date_value_diff"),
    ("TIMESTAMP", "txn_ts", "2024-01-01T00:00:00Z", "2024-01-01 00:00:00", "timezone_marker_diff"),
    ("TIMESTAMP", "txn_ts", "1704067200", "2024-01-01 00:00:00", "epoch_vs_formatted_date"),
    ("TIMESTAMP", "txn_ts", "1704067200000", "2024-01-01 00:00:00", "epoch_vs_formatted_date"),

    # ---- TEXT ----
    ("TEXT", "name", " BOB ", "BOB", "whitespace_only"),
    ("TEXT", "name", "jane doe", "JANE DOE", "case_difference"),
    ("TEXT", "name", "A  B", "A B", "whitespace_internal"),
    ("TEXT", "name", "JOSÉ", "JOSE", "unicode_fold_diff"),
    ("TEXT", "name", "MÜLLER", "M?LLER", "encoding_loss"),
    ("TEXT", "name", "CHRISTOPHER ALEXANDER", "CHRISTOPHER", "truncated_in_s3"),
    ("TEXT", "note", "it’s fine", "it's fine", "typographic_diff"),
    ("TEXT", "note", "a—b", "a-b", "typographic_diff"),
    ("TEXT", "status", "CLOSED", "CANCELLED", "value_diff"),
    ("TEXT", "ssn", "123-45-6789", "123456789", "punct_removed_in_s3:-"),
    ("TEXT", "path", "C:\\data\\in", "C:/data/in", "separator_format_diff"),
    ("TEXT", "acct", '12"2', "122", 'punct_removed_in_s3:"'),

    # ---- NULL HANDLING ----
    ("NULL", "close_date", "NULL", "", "null_representation_diff"),
    ("NULL", "close_date", "2024-06-01", "", "missing_in_s3"),
    ("NULL", "close_date", "", "2024-06-01", "missing_in_s2"),
    ("NULL", "note", "\\N", "", "null_representation_diff"),

    # ---- MUST NEVER BE CALLED HARMLESS (each was once wrongly reassuring) ----
    ("NO FALSE OK", "amount", "12345678901234567.89", "12345678901234567.88", "numeric_value_diff"),
    ("NO FALSE OK", "amount", "1234567890.12", "1234567890.13", "numeric_value_diff"),
    ("NO FALSE OK", "rate", "50%", "50", "percent_sign_diff"),
    ("NO FALSE OK", "fee", "$10.00", "€10.00", "currency_symbol_changed"),
    ("NO FALSE OK", "code", "12,34", "1234", "ambiguous_decimal_comma"),
    ("NO FALSE OK", "amt", "1,5", "1.5", "ambiguous_decimal_comma"),
    ("NO FALSE OK", "code", "v1.10", "v11.0", "value_diff"),
    ("NO FALSE OK", "acct", "12-345", "123-45", "value_diff"),
    ("NO FALSE OK", "name", "A中B", "AB", "value_diff"),
    ("NO FALSE OK", "open_date", "01/05/2024", "2024-01-05", "date_ambiguous_order"),
    ("NO FALSE OK", "open_date", "2024", "2024-01-01", "date_incomplete"),
    ("NO FALSE OK", "txn_ts", "2024-01-01 00:00:00.123456789", "2024-01-01 00:00:00.123456", "date_value_diff"),
    ("NO FALSE OK", "txn_ts", "1704067200", "2024-01-01 00:00:00.900", "value_diff"),
    ("NO FALSE OK", "code", "1-23", "12-3", "value_diff"),
    ("NO FALSE OK", "amt", "10", "10", "no_actual_diff"),
    ("NO FALSE OK", "rate", "99.456789012345678", "99.46", "numeric_rounding"),
    ("NO FALSE OK", "amount", "500-", "500", "numeric_value_diff"),
    ("NO FALSE OK", "amount", "1,234.50-", "1234.50", "numeric_value_diff"),
    ("NO FALSE OK", "amount", "12-", "12", "numeric_value_diff"),
    ("NO FALSE OK", "qty", "1", "true", "value_diff"),

    # ---- MUST STILL BE RECOGNISED WHEN GENUINELY EQUAL ----
    ("PROVEN SAME", "open_date", "13/05/2024", "2024-05-13", "date_format_diff"),
    ("PROVEN SAME", "open_date", "05/05/2024", "2024-05-05", "date_format_diff"),
    ("PROVEN SAME", "qty", "1,234", "1234", "numeric_thousands_sep_diff"),
    ("PROVEN SAME", "price", "99.456789", "99.460", "numeric_rounding"),
    ("PROVEN SAME", "price", "12.34", "12", "numeric_rounding"),
    ("PROVEN SAME", "price", "100.00", "150.00", "numeric_value_diff"),
    ("PROVEN SAME", "open_date", "01/05/2024", "03/07/2024", "date_value_diff"),
]

# the risk tier each of these patterns must carry; a regression that moves one of
# them into a gentler tier would re-introduce a false reassurance
MUST_NOT_BE_REASSURING = {
    "no_actual_diff", "percent_sign_diff", "ambiguous_decimal_comma",
    "date_ambiguous_order", "date_incomplete", "column_not_in_file",
    "currency_symbol_changed", "numeric_rounding", "null_representation_diff",
    "currency_symbol_diff",
}

def run():
    fails = []
    cur = None
    for typ, col, s2, s3, expected in CASES:
        if typ != cur:
            print(f"\n--- {typ} ---")
            cur = typ
        try:
            got, _ = classify_pair(col, s2, s3)
        except Exception as exc:
            got = f"CRASH:{type(exc).__name__}:{exc}"
        ok = got == expected
        if not ok:
            fails.append((typ, col, s2, s3, expected, got))
        print(f"  {'ok ' if ok else 'FAIL'} {s2!r:>28} -> {s3!r:<26} {got}"
              + ("" if ok else f"   (expected {expected})"))

    for pattern in sorted(MUST_NOT_BE_REASSURING):
        tier = pattern_risk(pattern)
        if tier not in ("REAL VALUE DIFF", "CANNOT VERIFY - check source"):
            fails.append(("TIER", pattern, "", "", "REAL or CANNOT VERIFY", tier))

    # a column mixing a free-text change with a format difference must never be
    # summarised as "every one proven to be the same value"
    for mix in ({"text_field_diff": 1, "separator_format_diff": 1},
                {"value_diff": 1, "whitespace_only": 1},
                {"date_ambiguous_order": 1, "date_format_diff": 1}):
        obs = column_observations(mix, {}, sum(mix.values()))
        if any("proven to be the same" in o for o in obs):
            fails.append(("OBSERVATION", "+".join(mix), "", "", "no 'proven same' claim", "; ".join(obs)))

    print()
    print("=" * 78)
    if fails:
        print(f"{len(fails)} of {len(CASES)} FAILED:")
        for typ, col, s2, s3, exp, got in fails:
            print(f"  [{typ}] {col}: {s2!r} -> {s3!r}  expected {exp}, got {got}")
    else:
        print(f"all {len(CASES)} cases produced the expected pattern")
    return 1 if fails else 0


if __name__ == "__main__":
    # only when run directly: rewrapping stdout on import swallowed the output
    # of any script that imported CASES from here
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    sys.exit(run())
