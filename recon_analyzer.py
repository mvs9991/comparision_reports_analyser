"""
Recon Mismatch Analyzer
=======================
Reads recon output (one subfolder per table, each containing
`<tablename>_mismatch.csv`) and produces an Excel workbook that describes,
for every mismatched column, *what* differs between s2 (live/source) and
s3 (newly-ingested parallel table).

Every statement in the output is computed by exact string/numeric/date
comparison. Nothing is inferred about *why* a difference exists.

Usage:
    python recon_analyzer.py <input_folder> [-o recon_summary.xlsx]
                             [--json digest.json] [--assessments notes.json]

Expected input layout:
    <input_folder>/<tablename>/<tablename>_mismatch.csv

Expected CSV layout (comma-delimited, double-quoted):
    match_type, mismatch_columns, <pk cols...>, <col>__s2, <col>__s3, ...

`mismatch_columns` holds the comma-separated non-PK column names that
differ for that record, so only those pairs are compared.
"""

import argparse
import csv
import json
import os
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime, timedelta, timezone
from decimal import ROUND_DOWN, ROUND_HALF_EVEN, ROUND_HALF_UP, Decimal, InvalidOperation
from functools import lru_cache
from itertools import chain, islice

from dateutil import parser as dtparser
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Value comparison
# ---------------------------------------------------------------------------

NULL_TOKENS = {"", "null", "none", "nan", "\\n", "n/a", "na", "nil", "(null)", "<null>"}
BOOL_MAP = {
    "y": True, "n": False, "yes": True, "no": False,
    "true": True, "false": False, "t": True, "f": False,
    "1": True, "0": False,
}
INF_TOKENS = {"inf", "+inf", "-inf", "infinity", "+infinity", "-infinity"}
CURRENCY_CHARS = "$€£¥₹"
FLAG_NAME = re.compile(
    r"((^|_)(is|has)_|_flag$|_flg$|^flag_|_ind$|_yn$|_bool$|_boolean$)", re.IGNORECASE
)

# smart quotes, dashes and non-breaking spaces against their ASCII equivalents
TYPO_MAP = str.maketrans({
    "‘": "'", "’": "'", "‚": "'", "‛": "'",
    "“": '"', "”": '"', "„": '"',
    "–": "-", "—": "-", "―": "-", "−": "-",
    " ": " ", "•": "*",
})
PUNCT_CANDIDATES = ['"', "'", ",", "-", "_", "/", "\\", ".", " ", "(", ")", "#", "&", "*", ":"]

DATEISH = re.compile(r"\d{2,4}[-/.]\d{1,2}[-/.]\d{1,4}|\d{1,2}:\d{2}")
DATE_NAME = re.compile(r"(date|time|stamp|_dt$|^dt_)", re.IGNORECASE)
TEXT_NAME_PATTERN = re.compile(
    r"(desc|description|comment|note|remark|narrative|summary|reason|"
    r"address|text|title|detail|message|instruction)",
    re.IGNORECASE,
)


# U+FFFD is what an undecodable byte becomes; the rest are control characters.
# Values carrying these are kept as-is in the output and counted, so the junk
# stays visible instead of being cleaned away behind your back.
CHARSET_JUNK = re.compile(r"[\000-\010\013\014\016-\037\177�]")


def has_charset_junk(s):
    # printable ASCII can hold neither a control character nor U+FFFD, and it is
    # nearly every value, so the regex only runs on the rare remainder
    if s.isascii() and s.isprintable():
        return False
    return CHARSET_JUNK.search(s) is not None


def is_null_like(s):
    return s.strip().lower() in NULL_TOKENS


# \W is Unicode-aware, so accented letters survive; an ASCII-only class would
# delete them and make an encoding-corrupted value look like a separator change
NON_ALNUM = re.compile(r"[\W_]+", re.UNICODE)


def _alnum(s):
    return NON_ALNUM.sub("", s)


def _fold(s):
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")


def _has_nonascii(s):
    return not s.isascii()


def _typo(s):
    return s.translate(TYPO_MAP)


PLAIN_NUMBER = re.compile(r"^[-+]?\d+(?:\.\d+)?$")
# commas are only read as thousands separators when they group correctly;
# '12,34' could be a decimal comma, so it is not treated as a number at all
GROUPED_NUMBER = re.compile(r"^[-+]?\d{1,3}(?:,\d{3})+(?:\.\d+)?$")
COMMA_NUMBER = re.compile(r"^[-+]?\d+(?:,\d+)+(?:\.\d+)?$")
HAS_DIGIT = re.compile(r"\d")
MONTH_WORD = re.compile(
    r"\b(jan(uary)?|feb(ruary)?|mar(ch)?|apr(il)?|may|june?|july?|aug(ust)?|"
    r"sep(t(ember)?)?|oct(ober)?|nov(ember)?|dec(ember)?)\b",
    re.IGNORECASE,
)
# a leading dd/MM or MM/dd whose first two fields could each be day or month
AMBIGUOUS_DAY_MONTH = re.compile(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})\b")
FRACTION_OF_SECOND = re.compile(r":\d{2}[.,](\d+)")
EPOCH = datetime(1970, 1, 1)

# the handful of layouts that cover almost every exported date; parsing these
# directly is roughly fifty times cheaper than handing them to dateutil
FAST_DATE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})$")
FAST_COMPACT = re.compile(r"^(\d{4})(\d{2})(\d{2})$")
FAST_DATETIME = re.compile(
    r"^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?(?:\.(\d{1,6}))?$"
)

# dateutil fills any missing part of a date from its 'default'. Parsing twice
# with two unrelated defaults exposes a value that is missing parts: the two
# results disagree. With a single default, '2024' silently became today's
# month and day, so the verdict depended on the day the report was run.
_DEFAULT_A = datetime(1901, 2, 3)
_DEFAULT_B = datetime(1902, 7, 8)


def fast_parse_dt(s):
    """Parse the common complete layouts without dateutil. None means 'not one of these'."""
    m = FAST_DATE.match(s)
    if m:
        try:
            return datetime(int(m[1]), int(m[2]), int(m[3]))
        except ValueError:
            return None
    m = FAST_DATETIME.match(s)
    if m:
        try:
            return datetime(
                int(m[1]), int(m[2]), int(m[3]), int(m[4]), int(m[5]),
                int(m[6] or 0), int((m[7] or "0").ljust(6, "0")),
            )
        except ValueError:
            return None
    m = FAST_COMPACT.match(s)
    if m:
        month, day = int(m[2]), int(m[3])
        if 1 <= month <= 12 and 1 <= day <= 31:
            try:
                return datetime(int(m[1]), month, day)
            except ValueError:
                return None
    return None


def date_plausible(s):
    """Cheap gate before dateutil, which is slowest when it is about to fail.

    Prose like 'Customer requested address change ref 5' gets fully tokenised
    before being rejected, so free-text columns otherwise dominate the runtime.
    A dozen letters still allows '15 September 2024'.
    """
    if len(s) > 40 or not HAS_DIGIT.search(s):
        return False
    letters = 0
    for ch in s:
        if ch.isalpha():
            letters += 1
            if letters > 12:
                return False
    return True


def date_readings(s):
    """Every (reading_with_default_A, reading_with_default_B) the text could mean.

    One entry normally; two when day and month genuinely cannot be told apart.
    The two datetimes in an entry differ when the text is missing date parts.
    Raises ValueError when the text is not a date.
    """
    fast = fast_parse_dt(s)
    if fast is not None:
        return [(fast, fast)]
    if not date_plausible(s):
        raise ValueError("not date-like")

    m = AMBIGUOUS_DAY_MONTH.match(s)
    if m and int(m[1]) <= 12 and int(m[2]) <= 12 and int(m[1]) != int(m[2]):
        orders = (False, True)
    else:
        orders = (False,)  # dateutil itself swaps an unambiguous one like 13/05

    readings = []
    for dayfirst in orders:
        readings.append((
            dtparser.parse(s, default=_DEFAULT_A, dayfirst=dayfirst),
            dtparser.parse(s, default=_DEFAULT_B, dayfirst=dayfirst),
        ))
    return readings


def _fraction_digits(s):
    m = FRACTION_OF_SECOND.search(s)
    return m[1].rstrip("0") if m else ""


def _to_dec(s):
    """Exact decimal value of the text, or ValueError.

    Decimal, not float: a float holds about 15 significant digits, so two
    different 17-digit amounts could compare equal and be reported as the same
    number written differently.
    """
    if PLAIN_NUMBER.match(s):
        return Decimal(s)
    core = s.strip()
    for sym in CURRENCY_CHARS:
        core = core.replace(sym, "")
    core = core.replace("%", "").strip()
    negative = False
    if core.startswith("(") and core.endswith(")"):
        core = core[1:-1].strip()
        negative = True
    if "," in core:
        if not GROUPED_NUMBER.match(core):
            raise ValueError("comma is not valid thousands grouping")
        core = core.replace(",", "")
    if core in ("", "-", "+", "."):
        raise ValueError("not numeric")
    try:
        val = Decimal(core)
    except InvalidOperation:
        raise ValueError("not numeric")
    if not val.is_finite():
        raise ValueError("not finite")
    return -val if negative else val


def _significant_digits(s):
    mantissa = re.split(r"[eE]", s, 1)[0]
    return len(re.sub(r"\D", "", mantissa).lstrip("0"))


def _currency_symbols(s):
    return {c for c in s if c in CURRENCY_CHARS}


def _has_currency(s):
    return bool(_currency_symbols(s))


def _only_deletions_of(longer, shorter, ch):
    """True when `shorter` is `longer` with some `ch` characters deleted and nothing else changed.

    Equal text after removing every `ch` is not enough: 'v1.10' and 'v11.0'
    both become 'v110', yet the dot moved and the value is different.
    """
    j = 0
    for c in longer:
        if j < len(shorter) and c == shorter[j]:
            j += 1
        elif c != ch:
            return False
    return j == len(shorter)


def _foldable(s):
    """Every non-ASCII character has an ASCII base letter (É -> E).

    Without this, a character with no ASCII form (中) is simply dropped by the
    fold and a lost character would be reported as 'accents flattened'.
    """
    for c in s:
        if ord(c) > 127 and not unicodedata.normalize("NFKD", c).encode("ascii", "ignore"):
            return False
    return True


def _compare_readings(ra, rb):
    """Compare one reading of each side. Returns (kind, delta_seconds)."""
    a1, a2 = ra
    b1, b2 = rb
    complete = a1 == a2 and b1 == b2
    a_aware, b_aware = a1.tzinfo is not None, b1.tzinfo is not None

    if a_aware != b_aware:
        # aware and naive cannot be subtracted; compare the wall-clock parts
        wa1, wa2, wb1, wb2 = (d.replace(tzinfo=None) for d in (a1, a2, b1, b2))
        if wa1 == wb1 and wa2 == wb2:
            return "marker", None
        return ("differ", (wb1 - wa1).total_seconds()) if complete else ("incomplete", None)

    if a1 == b1 and a2 == b2:
        if a_aware and a1.utcoffset() != b1.utcoffset():
            return "offset", None
        return "equal", None
    return ("differ", (b1 - a1).total_seconds()) if complete else ("incomplete", None)


@lru_cache(maxsize=4096)
def _column_is_dateish(col):
    return DATE_NAME.search(col) is not None


def classify_pair(col, s2, s3):
    """Return (pattern, delta) describing how s3 differs from s2.

    A pattern that says two values mean the same thing is only returned when
    that equality has been proven. Anything that cannot be proven either way
    gets a CANNOT VERIFY pattern rather than a reassuring one.
    """
    a = "" if s2 is None else str(s2)
    b = "" if s3 is None else str(s3)

    if a == b:
        return "no_actual_diff", None

    a_null, b_null = is_null_like(a), is_null_like(b)
    if a_null and b_null:
        return "null_representation_diff", None
    if a_null != b_null:
        return ("missing_in_s3", None) if b_null else ("missing_in_s2", None)

    if a.strip() == b.strip():
        return "whitespace_only", None
    if " ".join(a.split()) == " ".join(b.split()):
        return "whitespace_internal", None
    if a.strip().lower() == b.strip().lower():
        return "case_difference", None

    a_s, b_s = a.strip(), b.strip()

    if a_s.lstrip("0") == b_s.lstrip("0") and a_s.lstrip("0"):
        return "leading_zero_diff", None

    def try_epoch():
        """One side an epoch integer, the other a formatted date for exactly that instant."""
        for raw, other_raw in ((a_s, b_s), (b_s, a_s)):
            if not re.fullmatch(r"\d{10}|\d{13}", raw):
                continue
            try:
                readings = date_readings(other_raw)
            except (ValueError, TypeError, OverflowError):
                continue
            if len(readings) != 1 or readings[0][0] != readings[0][1]:
                continue
            other = readings[0][0]
            if other.tzinfo is not None:
                other = other.astimezone(timezone.utc).replace(tzinfo=None)
            as_dt = EPOCH + (timedelta(milliseconds=int(raw)) if len(raw) == 13
                             else timedelta(seconds=int(raw)))
            if other == as_dt:  # exact: 0.9 seconds apart is not the same instant
                return "epoch_vs_formatted_date", None
        return None

    def try_date():
        try:
            ra_list, rb_list = date_readings(a_s), date_readings(b_s)
        except (ValueError, TypeError, OverflowError):
            return None

        outcomes = [_compare_readings(ra, rb) for ra in ra_list for rb in rb_list]
        kinds = {k for k, _ in outcomes}

        if len(ra_list) > 1 or len(rb_list) > 1:
            # day/month order unknown: only report what EVERY reading agrees on
            if kinds == {"differ"}:
                return "date_value_diff", None
            if len(kinds) == 1 and kinds <= {"equal", "offset", "marker"}:
                kind = next(iter(kinds))
            else:
                return "date_ambiguous_order", None
            delta = None
        else:
            kind, delta = outcomes[0]

        if kind == "differ":
            return "date_value_diff", delta
        if kind == "incomplete":
            return "date_incomplete", None

        # equal to the microsecond, but a longer fraction (nanoseconds) may still differ
        if _fraction_digits(a_s) != _fraction_digits(b_s):
            return "date_value_diff", None
        if kind == "marker":
            return "timezone_marker_diff", None
        if kind == "offset":
            return "timezone_representation_diff", None
        return "date_format_diff", None

    def numeric_format_pattern():
        """Pin down *how* two numerically identical values are written differently.

        Scale is checked before separators: a changed decimal-place count is the
        signal that the column's precision or type changed, which is what breaks
        downstream joins and type-sensitive comparisons.
        """
        if ("e" in a_s.lower()) != ("e" in b_s.lower()):
            return "numeric_scientific_notation", None
        if _has_currency(a_s) != _has_currency(b_s):
            return "currency_symbol_diff", None
        if (a_s.startswith("(") and a_s.endswith(")")) != (b_s.startswith("(") and b_s.endswith(")")):
            return "accounting_negative_diff", None

        def decimals(s):
            s = s.replace(",", "")
            return len(s.split(".")[1]) if "." in s else 0

        da, db = decimals(a_s), decimals(b_s)
        if da != db:
            return "numeric_scale_diff", float(db - da)
        if ("," in a_s) != ("," in b_s):
            return "numeric_thousands_sep_diff", None
        if a_s.startswith("+") != b_s.startswith("+"):
            return "numeric_sign_format_diff", None
        return "numeric_format_diff", None

    def try_num():
        # '12,34' beside '1234': that comma may be a decimal comma (12.34), so
        # neither 'same number' nor 'comma removed, rest identical' is proven
        for x, y in ((a_s, b_s), (b_s, a_s)):
            if COMMA_NUMBER.match(x) and not GROUPED_NUMBER.match(x):
                try:
                    _to_dec(y)
                    return "ambiguous_decimal_comma", None
                except ValueError:
                    pass

        try:
            na, nb = _to_dec(a_s), _to_dec(b_s)
        except ValueError:
            return None

        # '50%' next to '50' could mean 50 or 0.5 - no way to prove which
        if ("%" in a_s) != ("%" in b_s):
            return "percent_sign_diff", None
        sym_a, sym_b = _currency_symbols(a_s), _currency_symbols(b_s)
        if sym_a and sym_b and sym_a != sym_b:
            return "currency_symbol_changed", None

        if na == nb:
            return numeric_format_pattern()

        # Float round-trip noise has a specific shape: one side is the short value,
        # the other is that value written out to double precision with trailing
        # junk digits (0.3 -> 0.30000000000000004). Rounding the long side back to
        # the short side's decimal places must give the short side exactly.
        # Two values with the SAME number of decimal places that differ
        # (12345678901234567.89 -> .88) are never noise, however large.
        dp_a = max(0, -na.as_tuple().exponent)
        dp_b = max(0, -nb.as_tuple().exponent)
        if dp_a != dp_b:
            (short_v, short_dp), (long_v, long_s) = (
                ((na, dp_a), (nb, b_s)) if dp_a < dp_b else ((nb, dp_b), (na, a_s))
            )
            try:
                back = long_v.quantize(Decimal(1).scaleb(-short_dp), rounding=ROUND_HALF_EVEN)
            except InvalidOperation:
                back = None
            if (back == short_v
                    and _significant_digits(long_s) >= 15
                    and abs(na - nb) <= max(abs(na), abs(nb)) * Decimal("1e-15")):
                return "float_precision_noise", float(nb - na)

        # rounding s2 can only land on s3 at s3's own number of decimal places,
        # so that single precision is tested rather than every one from 0 to 6
        try:
            places_a = max(0, -na.normalize().as_tuple().exponent)
            places_b = max(0, -nb.normalize().as_tuple().exponent)
            if places_b < places_a:
                step = Decimal(1).scaleb(-places_b)
                for mode in (ROUND_HALF_UP, ROUND_HALF_EVEN, ROUND_DOWN):
                    if na.quantize(step, rounding=mode) == nb:
                        return "numeric_rounding", float(places_b)
        except InvalidOperation:
            pass
        return "numeric_value_diff", float(nb - na)

    if a_s.lower() in INF_TOKENS or b_s.lower() in INF_TOKENS:
        return "infinity_value", None

    ba, bb = BOOL_MAP.get(a_s.lower()), BOOL_MAP.get(b_s.lower())
    both_bool = ba is not None and bb is not None

    # on a column named like a flag, 1/0 is a boolean rather than arithmetic
    if both_bool and FLAG_NAME.search(col):
        return ("boolean_format_diff", None) if ba == bb else ("boolean_value_diff", None)

    epoch = try_epoch()
    if epoch:
        return epoch

    # dates are only attempted on values that look like dates, so that a code
    # such as '1-23' is never re-read as 23 January
    date_first = _column_is_dateish(col) or bool(DATEISH.search(a_s) or DATEISH.search(b_s))
    # a month name only makes a date when digits are present too ('15 Sep 2024')
    date_like = date_first or bool(
        (HAS_DIGIT.search(a_s) and MONTH_WORD.search(a_s))
        or (HAS_DIGIT.search(b_s) and MONTH_WORD.search(b_s))
    )
    if date_first:
        result = try_date() or try_num()
    elif date_like:
        result = try_num() or try_date()
    else:
        result = try_num()
    if result:
        return result

    if both_bool:
        return ("boolean_format_diff", None) if ba == bb else ("boolean_value_diff", None)

    if _typo(a) == _typo(b):
        return "typographic_diff", None

    for ch in PUNCT_CANDIDATES:
        if ch not in a and ch not in b:
            continue
        count_a, count_b = a.count(ch), b.count(ch)
        if count_a == count_b or a.replace(ch, "") != b.replace(ch, ""):
            continue
        label = "space" if ch == " " else ch
        if count_a > count_b and _only_deletions_of(a, b, ch):
            return f"punct_removed_in_s3:{label}", None
        if count_b > count_a and _only_deletions_of(b, a, ch):
            return f"punct_added_in_s3:{label}", None

    # character-set checks run before the separator check: a corrupted value
    # must not be reported as a harmless punctuation difference
    if _has_nonascii(a) and ("�" in b or "?" in b) and not _has_nonascii(b):
        return "encoding_loss", None
    if _fold(a) and _fold(a) == _fold(b) and _foldable(a) and _foldable(b):
        return "unicode_fold_diff", None

    # same letters and digits in the same groups; a separator that moved
    # ('12-345' vs '123-45') changes the grouping and is not a format difference
    tokens_a = [t for t in NON_ALNUM.split(a) if t]
    tokens_b = [t for t in NON_ALNUM.split(b) if t]
    if tokens_a and "".join(tokens_a) == "".join(tokens_b) and (
            tokens_a == tokens_b or len(tokens_a) == 1 or len(tokens_b) == 1):
        return "separator_format_diff", None

    if b and a.startswith(b):
        return "truncated_in_s3", float(len(b))
    if a and b.startswith(a):
        return "truncated_in_s2", float(len(a))
    if b and b in a:
        return "partial_value_diff", None
    if a and a in b:
        return "partial_value_diff", None

    return "value_diff", None


# ---------------------------------------------------------------------------
# Pattern metadata: plain-language description + risk tier
# ---------------------------------------------------------------------------

R_REAL = "REAL VALUE DIFF"
R_UNSURE = "CANNOT VERIFY - check source"
R_TEXT = "TEXT FIELD - read manually"
R_CHECK = "FORMAT/TYPE - check joins"
# a statement of fact, not a judgement: trailing spaces and case still break
# exact-match joins, so nothing here is labelled low risk
R_LOW = "WHITESPACE/CASE ONLY"

RISK_ORDER = {R_REAL: 0, R_UNSURE: 1, R_TEXT: 2, R_CHECK: 3, R_LOW: 4}

PATTERN_INFO = {
    "no_actual_diff": (R_UNSURE, "the recon flagged this column but the exported s2 and s3 text is identical - the real difference was lost in the export (commonly NULL vs empty string, which both export as blank, or a precision/type difference the CSV cannot show). Check the source tables"),
    "column_not_in_file": (R_UNSURE, "mismatch_columns names this column but the file has no __s2/__s3 pair for it - these values were NOT compared"),
    "date_ambiguous_order": (R_UNSURE, "day/month order cannot be determined from the text (e.g. 01/05/2024); the values match under one reading and not the other, so equality is not proven"),
    "date_incomplete": (R_UNSURE, "a side is missing part of the date or time (e.g. year only) - equality cannot be proven either way"),
    "ambiguous_decimal_comma": (R_UNSURE, "a comma that is not valid thousands grouping (e.g. 12,34) - it may be a decimal comma, so it cannot be proven whether the numbers match"),
    "percent_sign_diff": (R_UNSURE, 'NOT MATCHED - a % sign is on one side only (50% vs 50). It cannot be told whether 50% here means 50 or 0.5, so these cannot be treated as matched'),
    "currency_symbol_changed": (R_REAL, "the currency symbol changed between s2 and s3 (e.g. $ -> EUR) - the amount may now be in a different currency"),
    "whitespace_only": (R_LOW, "leading/trailing whitespace differs; trimmed values identical"),
    "whitespace_internal": (R_LOW, "internal spacing differs; collapsed values identical"),
    "case_difference": (R_LOW, "letter case differs only"),

    "float_precision_noise": (R_CHECK, 'DECIMAL VALUE MISMATCH - the numerical values match; only the decimal digits differ (the column is held as float/double, e.g. 0.3 vs 0.30000000000000004)'),
    "currency_symbol_diff": (R_UNSURE, 'NOT MATCHED - a currency symbol is on one side only ($10.00 vs 10.00). The digits agree, but the currency is not confirmed on both sides, so these cannot be treated as matched'),
    "accounting_negative_diff": (R_CHECK, "negative written as (n) on one side and -n on the other; same number"),
    "numeric_sign_format_diff": (R_CHECK, "explicit + sign on one side only; same number"),
    "timezone_marker_diff": (R_CHECK, 'DATE FORMAT ISSUE - the clock time is the same, but one side carries a timezone and the other does not; conversions downstream may not agree'),
    "timezone_representation_diff": (R_CHECK, 'DATE FORMAT ISSUE - the same instant, written at a different UTC offset'),
    "epoch_vs_formatted_date": (R_CHECK, 'DATE FORMAT ISSUE - the date/time is the same; one side stores it as an epoch number, the other as a formatted date'),
    "typographic_diff": (R_CHECK, "smart quotes / long dashes / non-breaking spaces on one side, plain ASCII equivalents on the other"),
    "infinity_value": (R_REAL, "one or both sides hold an infinite or overflow value"),

    "numeric_scale_diff": (R_CHECK, "DECIMAL VALUE MISMATCH - the numerical values match; only the number of decimal places differs (e.g. 10.00 vs 10.0), which usually means the column's type or precision changed"),
    "numeric_thousands_sep_diff": (R_CHECK, "thousands separator on one side only; same number, different stored string"),
    "numeric_scientific_notation": (R_CHECK, "one side written in scientific notation (e.g. 1E+06); same number, different stored string"),
    "numeric_format_diff": (R_CHECK, "numerically equal but written differently; the stored string differs"),
    "date_format_diff": (R_CHECK, 'DATE FORMAT ISSUE - the date/time is the same; only the format differs (e.g. 2024-02-10 vs 20240210, or date vs timestamp)'),
    "boolean_format_diff": (R_CHECK, "same boolean written differently (Y/N vs true/false vs 1/0); the stored string differs"),
    "leading_zero_diff": (R_CHECK, "leading zeros differ; digits otherwise identical - changes the stored string"),
    "separator_format_diff": (R_CHECK, "same alphanumerics, different separators/punctuation"),
    "null_representation_diff": (R_REAL, "one side is an empty string, the other a NULL/placeholder token (e.g. '' vs 'NULL') - these behave differently in joins, aggregations, IS NULL checks and downstream filters"),
    "unicode_fold_diff": (R_CHECK, "the same letters, but accented on one side and plain on the other (e.g. JOSE with and without the accent)"),

    "numeric_rounding": (R_REAL, "s3 is a ROUNDED version of s2 - real precision was lost. Rounding is not an acceptable migration difference: treat this as a data difference, not a formatting one"),
    "numeric_value_diff": (R_REAL, "numeric values genuinely differ"),
    "date_value_diff": (R_REAL, "date/time values genuinely differ"),
    "boolean_value_diff": (R_REAL, "boolean values are opposite"),
    "missing_in_s3": (R_REAL, "s2 has a value, s3 is empty/null - data missing in s3"),
    "missing_in_s2": (R_REAL, "s3 has a value, s2 is empty/null - extra data in s3"),
    "truncated_in_s3": (R_REAL, "s3 value is a cut-off prefix of s2 - data lost"),
    "truncated_in_s2": (R_REAL, "s2 value is a cut-off prefix of s3 - s3 holds more data"),
    "partial_value_diff": (R_REAL, "one side's value is contained inside the other"),
    "encoding_loss": (R_REAL, "non-ASCII characters replaced or lost in s3 - encoding problem"),
    "value_diff": (R_REAL, "values differ with no recognisable formatting pattern"),
    "text_field_diff": (R_TEXT, "free-text/description field - wording differs, needs a human read"),
}


def pattern_info(pattern):
    if pattern in PATTERN_INFO:
        return PATTERN_INFO[pattern]
    if pattern.startswith("punct_removed_in_s3:"):
        ch = pattern.split(":", 1)[1]
        return R_CHECK, f"the character {ch} present in s2 is absent in s3; rest of the value identical"
    if pattern.startswith("punct_added_in_s3:"):
        ch = pattern.split(":", 1)[1]
        return R_CHECK, f"s3 contains an extra {ch} character not present in s2; rest of the value identical"
    return R_REAL, "unclassified difference"


def pattern_risk(pattern):
    return pattern_info(pattern)[0]


def pattern_desc(pattern):
    return pattern_info(pattern)[1]


def worst_risk(risks):
    # nothing measured is not the same as nothing wrong
    return min(risks, key=lambda r: RISK_ORDER.get(r, 0)) if risks else R_UNSURE


# Digit runs become '#' so the layout itself is the signature. Glossed only
# where the meaning is unambiguous - ##/##/#### could be dd/MM or MM/dd and is
# left as a mask rather than guessed at.
SHAPE_GLOSS = {
    "####-##-##": "yyyy-MM-dd",
    "########": "yyyyMMdd",
    "####-##-## ##:##:##": "yyyy-MM-dd HH:mm:ss",
    "####-##-##T##:##:##": "yyyy-MM-ddTHH:mm:ss",
    "####-##-## ##:##": "yyyy-MM-dd HH:mm",
    "####/##/##": "yyyy/MM/dd",
    "##:##:##": "HH:mm:ss",
}

# a format signature only says something for representation changes; for a
# genuine value change it is noise that fragments the cross-table rollup
SHAPE_EXCLUDE = {
    "no_actual_diff", "whitespace_only", "whitespace_internal", "case_difference",
    "column_not_in_file", "date_ambiguous_order", "date_incomplete", "ambiguous_decimal_comma",
    "currency_symbol_changed",
    "numeric_value_diff", "numeric_rounding", "date_value_diff", "boolean_value_diff",
    "missing_in_s3", "missing_in_s2", "truncated_in_s3", "truncated_in_s2",
    "partial_value_diff", "encoding_loss", "value_diff", "text_field_diff",
    "typographic_diff", "infinity_value",
}

NUMERIC_SHAPE_PATTERNS = {
    "float_precision_noise", "currency_symbol_diff", "percent_sign_diff",
    "accounting_negative_diff", "numeric_sign_format_diff",
}


def wants_shape(pattern):
    return pattern not in SHAPE_EXCLUDE


def numeric_shape(s):
    """Decimal places (plus separator/notation), not digit count.

    '##.##' and '####.##' are the same story - two decimal places - so a mask
    would split one finding across rows by how big the numbers happen to be.
    """
    if "e" in s.lower():
        return "scientific"
    plain = s.replace(",", "")
    places = len(plain.split(".")[1]) if "." in plain else 0
    return f"{places}dp" + (" +sep" if "," in s else "")


def value_shape(s, pattern=""):
    s = "" if s is None else str(s).strip()
    if not s:
        return "(empty)"
    if s.lower() in NULL_TOKENS:
        return s  # a sentinel token, not somebody's data
    if pattern.startswith("numeric_") or pattern in NUMERIC_SHAPE_PATTERNS:
        return numeric_shape(s)
    # without digits the mask would be the literal value: that would fragment the
    # rollup and copy real data into a summary sheet
    if not HAS_DIGIT.search(s):
        return "(text)"
    mask = re.sub(r"\d", "#", s)
    if len(mask) > 40:
        return "(long)"
    return SHAPE_GLOSS.get(mask, mask)


# ---------------------------------------------------------------------------
# Text-field detection (column name pattern OR value shape)
# ---------------------------------------------------------------------------

def is_text_field(colname, sample_values):
    if TEXT_NAME_PATTERN.search(colname):
        return True
    vals = [str(v) for v in sample_values if v and str(v).strip()]
    if not vals:
        return False
    avg_len = sum(len(v) for v in vals) / len(vals)
    avg_words = sum(len(v.split()) for v in vals) / len(vals)
    unique_ratio = len(set(vals)) / len(vals)
    return avg_len > 25 and avg_words > 3 and unique_ratio > 0.5


TEXT_OVERRIDE_PATTERNS = {"value_diff", "partial_value_diff"}

# on a boolean column these arithmetic verdicts are the wrong reading of 1 vs 0
BOOL_OVERRIDE_PATTERNS = {
    "numeric_value_diff", "numeric_scale_diff", "numeric_format_diff",
    "numeric_sign_format_diff",
}


def is_bool_field(colname, sample_values):
    if FLAG_NAME.search(colname):
        return True
    vals = [str(v).strip().lower() for v in sample_values if str(v).strip()]
    if len(vals) < 2 or not all(v in BOOL_MAP for v in vals):
        return False
    # a column holding only 0 and 1 is indistinguishable from a real number, so
    # that case needs the column name to confirm it
    return any(v not in ("0", "1") for v in vals)


# ---------------------------------------------------------------------------
# CSV discovery / loading
# ---------------------------------------------------------------------------

def find_mismatch_csv(table_folder):
    """Return (path or None, note or None).

    Prefers <folder>_mismatch.csv / <folder>_mismatched.csv. When a folder holds
    several mismatch files, the choice is reported instead of made silently.
    """
    hits = sorted(
        n for n in os.listdir(table_folder)
        if n.lower().endswith(".csv") and "mismatch" in n.lower()
    )
    if not hits:
        return None, None
    folder = os.path.basename(os.path.normpath(table_folder)).lower()
    exact = [n for n in hits if n.lower() in (f"{folder}_mismatch.csv", f"{folder}_mismatched.csv")]
    chosen = exact[0] if exact else hits[0]
    note = None
    if len(hits) > 1:
        note = (f"{len(hits)} mismatch CSV files are in this folder ({', '.join(hits)}); "
                f"ONLY {chosen} was analysed")
    return os.path.join(table_folder, chosen), note


def split_mismatch_columns(raw):
    if raw is None:
        return []
    s = str(raw).strip().strip("[]")
    if not s:
        return []
    parts = [p.strip().strip('"').strip("'") for p in s.split(",")]
    return [p for p in parts if p]


def open_table(csv_path):
    """Read the header, then hand back a generator over the remaining rows.

    Streaming keeps memory flat no matter how many rows the file holds, so a
    full recon export (millions of rows) costs the same as a 20-row sample.
    Decoding uses errors="replace": an invalid byte becomes U+FFFD, which is
    kept in the output and counted, rather than aborting or being guessed at.
    """
    notes = []
    stats = {"nul_lines": 0, "short_rows": 0, "long_rows": 0, "blank_lines": 0}
    fh = open(csv_path, newline="", encoding="utf-8-sig", errors="replace")

    def lines():
        # csv.reader refuses a NUL outright; swapping it for the U+FFFD marker
        # keeps the row readable AND keeps the junk visible to the charset check
        for line in fh:
            if "\x00" in line:
                stats["nul_lines"] += 1
                line = line.replace("\x00", "�")
            yield line

    reader = csv.reader(lines())

    try:
        header = next(reader)
    except StopIteration:
        fh.close()
        raise ValueError("file is empty - no header row")

    cols = [c.strip() for c in header]
    lower = [c.lower() for c in cols]

    dupes = [c for c, n in Counter(cols).items() if n > 1]
    if dupes:
        notes.append(
            f"the header repeats column name(s) {dupes[:5]} - only the first "
            f"occurrence of each is used"
        )

    if "mismatch_columns" not in lower:
        fh.close()
        raise ValueError(f"no 'mismatch_columns' column found; header was: {cols[:8]}")

    mm_idx = lower.index("mismatch_columns")
    mt_idx = lower.index("match_type") if "match_type" in lower else None

    s2_idx = [i for i, c in enumerate(lower) if c.endswith("__s2")]
    s3_by_name = {lower[i]: i for i in range(len(lower)) if lower[i].endswith("__s3")}

    # resolved case-insensitively: headers may use __S2/__S3, and the names
    # inside mismatch_columns may not match the header's casing either
    pair_map = {}
    for i in s2_idx:
        base = cols[i][:-4]
        j = s3_by_name.get(f"{base}__s3".lower())
        if j is not None and base.lower() not in pair_map:
            pair_map[base.lower()] = (base, i, j)

    first_s2 = min(s2_idx, default=len(cols))
    pk_idx = list(range(mm_idx + 1, first_s2))
    pk_cols = [cols[i] for i in pk_idx]

    width = len(cols)

    def rows():
        try:
            for row in reader:
                # a completely empty line (commonly a blank line at the end of the
                # export) holds no record; counting it would report a clean file
                # as malformed and send it for a manual check
                if not row:
                    stats["blank_lines"] += 1
                    continue
                # a short row would otherwise raise IndexError downstream; both
                # cases are counted so a malformed file is reported, not absorbed
                if len(row) < width:
                    stats["short_rows"] += 1
                    row = row + [""] * (width - len(row))
                elif len(row) > width:
                    stats["long_rows"] += 1
                yield row
        finally:
            fh.close()

    return {
        "cols": cols, "mm_idx": mm_idx, "mt_idx": mt_idx,
        "pk_idx": pk_idx, "pk_cols": pk_cols,
        "pair_map": pair_map, "notes": notes, "rows": rows(), "stats": stats,
    }


# ---------------------------------------------------------------------------
# Aggregate observations (facts computed across a column's records)
# ---------------------------------------------------------------------------

def fmt_delta(pattern, value):
    if value is None:
        return None
    if pattern.startswith("date_"):
        return str(timedelta(seconds=value))
    if pattern == "numeric_rounding":
        return f"{int(value)} decimal places"
    if pattern == "numeric_scale_diff":
        n = int(value)
        word = "place" if abs(n) == 1 else "places"
        return f"{abs(n)} {'fewer' if n < 0 else 'more'} decimal {word} in s3"
    if pattern.startswith("truncated_"):
        return f"{int(value)} characters"
    return f"{value:+g}"


def _plural(n, word="record"):
    return f"{n} {word}" + ("" if n == 1 else "s")


def column_observations(patterns, delta_state, total):
    obs = []

    # consistency is checked per pattern group, not only when the column has a
    # single pattern - a constant offset inside one group is the most useful
    # fact available and would otherwise be hidden by unrelated patterns
    for p, n in sorted(patterns.items(), key=lambda kv: -kv[1]):
        state = delta_state.get(p)
        if n < 2 or not state or state[0] is None or not state[1]:
            continue
        shown = fmt_delta(p, state[0])
        if p == "numeric_rounding":
            obs.append(f"s3 rounded to {shown} in all {_plural(n)}")
        elif p == "numeric_scale_diff":
            obs.append(f"{shown} in all {_plural(n)} - consistent precision/type change")
        elif p == "truncated_in_s3":
            obs.append(f"s3 cut off at {shown} in all {_plural(n)}")
        elif p in ("date_value_diff", "numeric_value_diff"):
            obs.append(f"all {_plural(n)} with a real difference differ by exactly {shown} (s3 - s2)")

    if len(patterns) == 1:
        obs.append(f"single consistent pattern across {_plural(total)}")
    else:
        real = sum(n for p, n in patterns.items() if pattern_risk(p) == R_REAL)
        unsure = sum(n for p, n in patterns.items() if pattern_risk(p) == R_UNSURE)
        # a free-text change is not proven equal either, so it must not fall
        # through to the "every one proven to be the same value" line below
        text = sum(n for p, n in patterns.items() if pattern_risk(p) == R_TEXT)
        if real:
            obs.append(f"MIXED: {real} of {_plural(total)} show a real value difference")
        if unsure:
            obs.append(f"{unsure} of {_plural(total)} CANNOT be verified from the file")
        if text:
            obs.append(f"MIXED: {text} of {_plural(total)} are free-text wording changes that need reading")
        if not real and not unsure and not text:
            obs.append(f"MIXED formatting patterns across {_plural(total)} (every one proven to be the same value)")

    if patterns.get("missing_in_s3") == total:
        obs.append(f"s3 empty in every one of the {_plural(total)}")
    if patterns.get("missing_in_s2") == total:
        obs.append(f"s2 empty in every one of the {_plural(total)}")
    return obs


# ---------------------------------------------------------------------------
# Per-table analysis
# ---------------------------------------------------------------------------

BLANK_MARKER = "[mismatch_columns blank - every column compared]"

# a date written differently is flagged for a person to look at, even though the
# date itself is proven to be the same
DATE_FORMAT_PATTERNS = {
    "date_format_diff", "epoch_vs_formatted_date",
    "timezone_marker_diff", "timezone_representation_diff",
}


def _names(cols, limit=4):
    shown = ", ".join(cols[:limit])
    return shown + (f" (+{len(cols) - limit} more)" if len(cols) > limit else "")


def manual_check_reasons(columns_summary, total_rows, unlisted, blank_rows_identical,
                         dupes, dup_info, pk_cols, skipped_match_type, stats):
    """Everything measurable that a person should look at, even though the checks ran.

    A real value difference on its own is an expected finding, not an oddity, so
    it does not appear here; it is already shown in the headline. These are the
    things that suggest the comparison, the export or the data is behaving
    strangely.
    """
    reasons = []

    def cols_where(test):
        return [c["column"] for c in columns_summary if test(c)]

    if total_rows == 0:
        reasons.append("the mismatch file has no rows at all")
    elif not columns_summary:
        reasons.append("rows exist but no column in them could be compared")

    unexplained = cols_where(lambda c: "value_diff" in c["patterns"])
    if unexplained:
        reasons.append(f"differences that no rule could explain in {_names(unexplained)}")

    unsure = cols_where(lambda c: any(pattern_risk(p) == R_UNSURE for p in c["patterns"]))
    if unsure:
        reasons.append(f"values the file cannot prove match or differ in {_names(unsure)}")

    text = cols_where(lambda c: "text_field_diff" in c["patterns"])
    if text:
        reasons.append(f"free-text wording changed in {_names(text)} - needs reading")

    mixed = cols_where(lambda c: (
        any(pattern_risk(p) == R_REAL for p in c["patterns"])
        and any(pattern_risk(p) in (R_CHECK, R_LOW) for p in c["patterns"])
    ))
    if mixed:
        reasons.append(
            f"inconsistent behaviour - real changes mixed with format-only differences - in {_names(mixed)}"
        )

    date_format = cols_where(lambda c: any(p in DATE_FORMAT_PATTERNS for p in c["patterns"]))
    if date_format:
        reasons.append(f"date format issue (the dates are the same, the format differs) in {_names(date_format)}")

    offset = cols_where(lambda c: any("differ by exactly" in o for o in c["observations"]))
    if offset:
        reasons.append(f"every differing record shifted by the same amount in {_names(offset)}")

    emptied = cols_where(lambda c: c["records"] > 1 and c["patterns"].get("missing_in_s3") == c["records"])
    if emptied:
        reasons.append(f"s3 is empty for every record in {_names(emptied)} - column may not be loaded")

    cut = cols_where(lambda c: "truncated_in_s3" in c["patterns"] or "encoding_loss" in c["patterns"])
    if cut:
        reasons.append(f"values cut off or characters lost in s3 in {_names(cut)}")

    junk = cols_where(lambda c: c.get("charset_junk", {}).get("s2") or c.get("charset_junk", {}).get("s3"))
    if junk:
        reasons.append(f"invalid / control characters in {_names(junk)}")

    if unlisted:
        reasons.append(f"the recon did not flag differences in {_names(sorted(unlisted))}")
    if blank_rows_identical:
        reasons.append(f"{_plural(blank_rows_identical, 'row')} flagged as mismatch with nothing visibly different")
    if dupes:
        reasons.append(f"duplicate primary keys ({len(dupes)} key value(s))")
    if not pk_cols:
        reasons.append("no primary key columns found, so duplicate keys could not be checked")
    elif not dup_info["complete"]:
        reasons.append("the duplicate-key check stopped early and is incomplete")
    if skipped_match_type:
        reasons.append(f"{sum(skipped_match_type.values())} row(s) skipped because match_type was not mismatch")
    if stats["short_rows"] or stats["long_rows"]:
        reasons.append("malformed lines in the file (wrong number of fields)")
    return reasons


def _worst_first(item):
    """Sort key for (pattern, value) items: highest risk first, then by name."""
    return (RISK_ORDER.get(pattern_risk(item[0]), 0), item[0])


SAMPLE_ROWS = 500          # buffered up front to judge column types
CACHE_MAX = 200_000        # bounded so a huge file cannot grow it without limit
PK_TRACK_MAX = 3_000_000   # beyond this the duplicate-PK check stops and says so


def analyze_table(table_name, csv_path, examples_per_group=3):
    loaded = open_table(csv_path)
    mm_idx = loaded["mm_idx"]
    mt_idx = loaded["mt_idx"]
    pk_idx = loaded["pk_idx"]
    pk_cols = loaded["pk_cols"]
    pair_map = loaded["pair_map"]
    row_iter = loaded["rows"]
    notes = list(loaded["notes"])

    buffered = list(islice(row_iter, SAMPLE_ROWS))

    # decide from the sample whether match_type actually carries 'mismatch'; if
    # it never does, filtering would silently discard the whole file
    filter_rows = False
    if mt_idx is not None:
        if any("mismatch" in r[mt_idx].lower() for r in buffered):
            filter_rows = True
        elif buffered:
            seen = sorted({r[mt_idx] for r in buffered})[:5]
            notes.append(
                f"match_type never held 'mismatch' in the first {len(buffered)} rows "
                f"(saw: {seen}); every row was analysed instead of filtering"
            )

    text_flags, bool_flags = {}, {}
    for base, i2, i3 in pair_map.values():
        s2_sample = [r[i2] for r in buffered]
        text_flags[base] = is_text_field(base, s2_sample)
        bool_flags[base] = is_bool_field(base, s2_sample + [r[i3] for r in buffered])

    col_patterns = defaultdict(Counter)
    col_delta_state = defaultdict(dict)   # col -> pattern -> [first, all_same, n]
    col_examples = defaultdict(lambda: defaultdict(list))
    col_shapes = defaultdict(lambda: defaultdict(Counter))
    col_records = Counter()
    col_charset = defaultdict(lambda: [0, 0])   # col -> [s2 hits, s3 hits]

    pk_seen = set()
    pk_dupes = Counter()
    pk_tracking = bool(pk_cols)

    combos = defaultdict(lambda: {
        "rows": 0,
        "pks": [],
        "patterns": defaultdict(Counter),
        "examples": defaultdict(dict),   # col -> pattern -> first (s2, s3)
    })

    cache = {}
    missing_pairs = Counter()
    total_rows = 0
    skipped_match_type = Counter()
    blank_rows = 0
    blank_rows_identical = 0
    unlisted = {}                      # col -> [rows, (pk, s2, s3) of the first]
    pairs_list = list(pair_map.values())
    pairs_with_lower = [(base, i2, i3, base.lower()) for base, i2, i3 in pairs_list]

    for tup in chain(buffered, row_iter):
        if filter_rows and "mismatch" not in tup[mt_idx].lower():
            skipped_match_type[tup[mt_idx]] += 1
            continue
        total_rows += 1
        row_pk = tuple(tup[i] for i in pk_idx)

        if pk_tracking:
            joined = "\x1e".join(row_pk)
            if joined in pk_seen:
                pk_dupes[row_pk] += 1
            elif len(pk_seen) >= PK_TRACK_MAX:
                pk_tracking = False
                notes.append(
                    f"duplicate-primary-key check stopped after {PK_TRACK_MAX:,} "
                    f"distinct keys to bound memory; later rows were not checked"
                )
            else:
                pk_seen.add(joined)

        flagged = split_mismatch_columns(tup[mm_idx])
        if flagged:
            flagged_lower = {f.lower() for f in flagged}
            # names are canonicalised to the header's spelling so that the same
            # column referred to in two different cases collapses into one entry
            resolved, key_parts = [], []
            for raw in flagged:
                pair = pair_map.get(raw.lower())
                if pair:
                    if pair not in resolved:
                        resolved.append(pair)
                        key_parts.append(pair[0])
                else:
                    key_parts.append(raw)
                    missing_pairs[raw] += 1
            key = tuple(sorted(set(key_parts)))

            # the recon's own list can be incomplete: any other column whose text
            # differs on this row is a difference nobody was told about
            for base, i2, i3, base_lower in pairs_with_lower:
                if tup[i2] != tup[i3] and base_lower not in flagged_lower:
                    seen_unlisted = unlisted.get(base)
                    if seen_unlisted is None:
                        unlisted[base] = [1, (row_pk, tup[i2], tup[i3])]
                    else:
                        seen_unlisted[0] += 1
        else:
            # nothing says which columns differ, so every pair is compared rather
            # than the row being skipped
            blank_rows += 1
            resolved = [p for p in pairs_list if tup[p[1]] != tup[p[2]]]
            if not resolved:
                blank_rows_identical += 1
                continue
            key = (BLANK_MARKER,) + tuple(sorted(p[0] for p in resolved))

        centry = combos[key]
        centry["rows"] += 1
        if len(centry["pks"]) < examples_per_group:
            centry["pks"].append(row_pk)

        for col, i2, i3 in resolved:
            s2v, s3v = tup[i2], tup[i3]

            if has_charset_junk(s2v):
                col_charset[col][0] += 1
            if has_charset_junk(s3v):
                col_charset[col][1] += 1

            ck = (col, s2v, s3v)
            if ck in cache:
                pattern, delta = cache[ck]
            else:
                pattern, delta = classify_pair(col, s2v, s3v)
                if text_flags.get(col) and pattern in TEXT_OVERRIDE_PATTERNS:
                    pattern = "text_field_diff"
                elif bool_flags.get(col) and pattern in BOOL_OVERRIDE_PATTERNS:
                    ba = BOOL_MAP.get(str(s2v).strip().lower())
                    bb = BOOL_MAP.get(str(s3v).strip().lower())
                    if ba is not None and bb is not None:
                        pattern = "boolean_format_diff" if ba == bb else "boolean_value_diff"
                        delta = None
                if len(cache) >= CACHE_MAX:
                    cache.clear()
                cache[ck] = (pattern, delta)

            col_records[col] += 1
            col_patterns[col][pattern] += 1

            # only "are they all the same" matters, so keep a running check
            # rather than a list that would grow with the row count
            state = col_delta_state[col].get(pattern)
            if state is None:
                col_delta_state[col][pattern] = [delta, True, 1]
            else:
                state[2] += 1
                if state[1] and state[0] != delta:
                    state[1] = False
            if wants_shape(pattern):
                col_shapes[col][pattern][
                    (value_shape(s2v, pattern), value_shape(s3v, pattern))
                ] += 1
            if len(col_examples[col][pattern]) < examples_per_group:
                col_examples[col][pattern].append((s2v, s3v, row_pk))

            centry["patterns"][col][pattern] += 1
            if pattern not in centry["examples"][col]:
                centry["examples"][col][pattern] = (s2v, s3v)

    # the column summary shows 2 examples per pattern (6 in total) by default;
    # raising --examples above its default of 3 raises both, as STEPS.txt promises
    per_pattern = 2 if examples_per_group <= 3 else examples_per_group
    max_column_examples = max(6, per_pattern * 3)

    columns_summary = []
    for col, patterns in col_patterns.items():
        total = col_records[col]
        risk = worst_risk([pattern_risk(p) for p in patterns])
        dominant, dom_n = patterns.most_common(1)[0]
        examples = []
        for p, exs in sorted(col_examples[col].items(), key=_worst_first):
            for s2v, s3v, pkv in exs[:per_pattern]:
                examples.append((p, s2v, s3v, pkv))
        shapes = {}
        for p, counter in col_shapes[col].items():
            if counter:
                shapes[p] = counter.most_common(1)[0][0]

        junk_s2, junk_s3 = col_charset.get(col, (0, 0))
        if junk_s2 or junk_s3:
            sides = []
            if junk_s2:
                sides.append(f"s2 side in {_plural(junk_s2)}")
            if junk_s3:
                sides.append(f"s3 side in {_plural(junk_s3)}")
            observations = column_observations(patterns, col_delta_state[col], total)
            observations.insert(0, (
                "NOT VALID UTF-8 / control characters observed on the "
                + " and ".join(sides)
                + " - the characters are left untouched in the examples "
                  "(shown escaped, e.g. \\x07 or the U+FFFD marker)"
            ))
        else:
            observations = column_observations(patterns, col_delta_state[col], total)

        columns_summary.append({
            "column": col,
            "is_text_field": text_flags.get(col, False),
            "records": total,
            "risk": risk,
            "charset_junk": {"s2": junk_s2, "s3": junk_s3},
            "dominant_pattern": dominant,
            "dominant_count": dom_n,
            "patterns": dict(patterns),
            "shapes": shapes,
            "descriptions": [pattern_desc(p) for p in patterns],
            "observations": observations,
            "examples": examples[:max_column_examples],
        })

    for name, n in sorted(missing_pairs.items()):
        columns_summary.append({
            "column": name,
            "is_text_field": False,
            "records": n,
            "risk": R_UNSURE,
            "charset_junk": {"s2": 0, "s3": 0},
            "dominant_pattern": "column_not_in_file",
            "dominant_count": n,
            "patterns": {"column_not_in_file": n},
            "shapes": {},
            "descriptions": [pattern_desc("column_not_in_file")],
            "observations": [
                f"named in mismatch_columns on {_plural(n, 'row')}, but the file has no "
                f"{name}__s2 / {name}__s3 columns - these values were NOT compared"
            ],
            "examples": [],
        })

    columns_summary.sort(key=lambda c: (RISK_ORDER.get(c["risk"], 0), -c["records"], c["column"]))

    combos_summary = []
    for key, entry in combos.items():
        col_bits = []
        risks = []
        for col in key:
            pats = entry["patterns"].get(col, Counter())
            total = sum(pats.values())
            if not total:
                continue
            risks.append(worst_risk([pattern_risk(p) for p in pats]))
            if len(pats) == 1:
                p = next(iter(pats))
                col_bits.append({
                    "column": col, "text": f"{col}: {p} ({total}/{total})",
                    "patterns": dict(pats),
                })
            else:
                parts = ", ".join(f"{p}={n}" for p, n in pats.most_common())
                col_bits.append({
                    "column": col, "text": f"{col}: MIXED [{parts}] (of {total})",
                    "patterns": dict(pats),
                })
        combos_summary.append({
            "combo": ", ".join(key),
            "column_count": len([c for c in key if c != BLANK_MARKER]),
            "rows": entry["rows"],
            "risk": worst_risk(risks),
            "column_detail": col_bits,
            "text_fields": [c for c in key if text_flags.get(c)],
            "examples": {
                c: [v for _, v in sorted(entry["examples"][c].items(), key=_worst_first)][:examples_per_group]
                for c in key if entry["examples"].get(c)
            },
            "pks": entry["pks"],
        })

    combos_summary.sort(key=lambda c: (RISK_ORDER.get(c["risk"], 0), -c["rows"], c["combo"]))

    # the same PK on more than one row means the recon join fanned out (or the
    # source held duplicates) - the comparison itself is suspect, not just the data
    dupes = pk_dupes
    dup_info = {
        "checked": bool(pk_cols),
        "complete": bool(pk_cols) and pk_tracking,
        "pk_values_duplicated": len(dupes),
        # each extra sighting beyond the first, plus the original row
        "rows_involved": sum(dupes.values()) + len(dupes),
        "examples": [
            ", ".join(f"{c}={v}" for c, v in zip(pk_cols, pk)) + f"  (x{n + 1})"
            for pk, n in sorted(dupes.items(), key=lambda kv: -kv[1])[:3]
        ],
    }

    if loaded["stats"]["nul_lines"]:
        notes.append(
            f"NUL bytes found on {_plural(loaded['stats']['nul_lines'], 'line')} - "
            f"each was replaced with the U+FFFD junk marker so the row could still "
            f"be read; the affected columns are flagged below"
        )

    stats = loaded["stats"]
    if stats["short_rows"]:
        notes.append(
            f"MALFORMED FILE: {_plural(stats['short_rows'], 'row')} had fewer fields than the "
            f"header; the missing fields were read as blank, so some 'missing' verdicts in this "
            f"sheet may really be broken lines"
        )
    if stats["long_rows"]:
        notes.append(
            f"MALFORMED FILE: {_plural(stats['long_rows'], 'row')} had more fields than the "
            f"header; the extra fields were NOT compared"
        )
    if skipped_match_type:
        shown = ", ".join(f"'{v}' x{n}" for v, n in skipped_match_type.most_common(5))
        notes.append(
            f"{sum(skipped_match_type.values())} row(s) were NOT analysed because match_type "
            f"did not say mismatch: {shown}"
        )
    if blank_rows:
        notes.append(
            f"{_plural(blank_rows, 'row')} had a blank mismatch_columns, so every column was "
            f"compared for them (shown under '{BLANK_MARKER}')"
        )
    if blank_rows_identical:
        notes.append(
            f"CANNOT VERIFY: {_plural(blank_rows_identical, 'row')} had a blank mismatch_columns "
            f"AND identical s2/s3 text in every column - whatever the recon saw is not visible "
            f"in this file"
        )
    for name, (n, (pkv, s2v, s3v)) in sorted(unlisted.items(), key=lambda kv: -kv[1][0])[:20]:
        pk_text = ", ".join(f"{c}={v}" for c, v in zip(pk_cols, pkv)) or "row"
        notes.append(
            f"RECON DID NOT FLAG: '{name}' differs on {_plural(n, 'row')} where mismatch_columns "
            f"did not name it (e.g. {pk_text}: {s2v!r} -> {s3v!r})"
        )
    if len(unlisted) > 20:
        notes.append(f"... and {len(unlisted) - 20} more column(s) the recon did not flag")

    risk_counts = Counter(c["risk"] for c in columns_summary)
    if total_rows == 0:
        notes.append("this file contains no mismatch rows")
    elif not columns_summary:
        notes.append(
            "rows were present but mismatch_columns named no column that has a "
            "matching __s2/__s3 pair - check the column naming in this file"
        )
    if dupes:
        notes.append(
            f"DUPLICATE PRIMARY KEYS: {len(dupes)} PK value(s) appear on more than one row "
            f"({dup_info['rows_involved']} rows involved). Examples: {'; '.join(dup_info['examples'])}"
        )
    if missing_pairs:
        notes.append(
            f"CANNOT VERIFY: {len(missing_pairs)} column(s) named in mismatch_columns have no "
            f"__s2/__s3 pair in the file: {sorted(missing_pairs)[:5]}"
        )

    headline = (
        f"{risk_counts.get(R_REAL, 0)} of {len(columns_summary)} mismatched columns show real value "
        f"differences; {risk_counts.get(R_UNSURE, 0)} cannot be verified from the file; "
        f"{risk_counts.get(R_CHECK, 0)} same value, different format/type (check downstream joins); "
        f"{risk_counts.get(R_LOW, 0)} whitespace/case only; {risk_counts.get(R_TEXT, 0)} text fields"
    )
    if unlisted:
        headline += f"; RECON DID NOT FLAG differences in {len(unlisted)} further column(s)"
    if blank_rows_identical:
        headline += f"; {_plural(blank_rows_identical, 'row')} cannot be verified at all"

    needs_attention = bool(
        risk_counts.get(R_REAL) or risk_counts.get(R_UNSURE) or risk_counts.get(R_TEXT)
        or unlisted or blank_rows_identical or dupes or skipped_match_type
        or stats["short_rows"] or stats["long_rows"] or not dup_info["complete"] and pk_cols
    )

    manual_reasons = manual_check_reasons(
        columns_summary, total_rows, unlisted, blank_rows_identical, dupes,
        dup_info, pk_cols, skipped_match_type, stats,
    )

    return {
        "table_name": table_name,
        "csv_path": csv_path,
        "manual_check_reasons": manual_reasons,
        "total_sample_rows": total_rows,
        "distinct_combos": len(combos_summary),
        "pk_cols": pk_cols,
        "columns": columns_summary,
        "combos": combos_summary,
        "risk_counts": dict(risk_counts),
        "pk_duplicates": dup_info,
        "notes": notes,
        "unlisted": {name: n for name, (n, _) in unlisted.items()},
        "needs_attention": needs_attention,
        "headline": headline,
    }


# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------

# Excel rejects these outright; one stray byte in an exported text field would
# otherwise abort the save after every table had already been analysed
ILLEGAL_CELL_CHARS = re.compile(r"[\000-\010\013\014\016-\037]")
MAX_CELL_CHARS = 32000  # Excel's hard ceiling is 32767


def safe_cell(value):
    if value is None or isinstance(value, (int, float)):
        return value
    # escaped rather than deleted: Excel refuses the raw byte, but replacing it
    # silently would hide where the junk actually is
    text = ILLEGAL_CELL_CHARS.sub(lambda m: "\\x%02x" % ord(m.group()), str(value))
    if len(text) > MAX_CELL_CHARS:
        text = text[:MAX_CELL_CHARS] + "  ...[truncated]"
    # openpyxl turns a leading '=' into a formula; keep it as text
    if text.startswith("="):
        text = " " + text
    return text


def arow(ws, values):
    """Append a row, sanitising every cell first."""
    ws.append([safe_cell(v) for v in values])
    return ws.max_row


HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
SECTION_FONT = Font(bold=True, size=11, color="1F4E78")
RISK_FILL = {
    R_REAL: PatternFill("solid", fgColor="F4B183"),
    R_UNSURE: PatternFill("solid", fgColor="D9C3E9"),
    R_TEXT: PatternFill("solid", fgColor="FFE699"),
    R_CHECK: PatternFill("solid", fgColor="DDEBF7"),
    R_LOW: PatternFill("solid", fgColor="C6E0B4"),
}
WRAP = Alignment(wrap_text=True, vertical="top")


def sanitize_sheet_name(name, used):
    safe = re.sub(r"[\[\]:*?/\\]", "_", name)[:31] or "sheet"
    base, i = safe, 1
    while safe.lower() in used:
        suffix = f"_{i}"
        safe = base[: 31 - len(suffix)] + suffix
        i += 1
    used.add(safe.lower())
    return safe


def fmt_examples(examples):
    return "\n".join(f"{s2!r} -> {s3!r}" for s2, s3 in examples)


def lookup_query(table_template, table_name, pk_cols, pk_vals):
    """SELECT for one sample record, for pasting straight into a SQL client."""
    if not pk_cols:
        return ""
    table = table_template.format(table=table_name)
    conds = " AND ".join(
        # values can legitimately contain a single quote, which would otherwise
        # produce broken SQL when pasted
        f"{col}='{str(val).replace(chr(39), chr(39) * 2)}'"
        for col, val in zip(pk_cols, pk_vals)
    )
    return f"SELECT * FROM {table} WHERE {conds};"


def style_header_row(ws, row_idx, ncols):
    for i in range(1, ncols + 1):
        cell = ws.cell(row=row_idx, column=i)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = WRAP


def write_table_sheet(ws, r, recon_table_template="{table}", max_combos=500):
    arow(ws, [f"Table: {r['table_name']}"])
    ws["A1"].font = Font(bold=True, size=14)
    arow(ws, [f"Source file: {r['csv_path']}"])
    arow(ws, [
        f"Sampled mismatch records in this file: {r['total_sample_rows']} "
        f"(export keeps up to 20 rows per distinct mismatch_columns combination, "
        f"so production mismatch volume is higher than this)"
    ])
    arow(ws, [f"Distinct mismatch_columns combinations: {r['distinct_combos']}"])
    arow(ws, [f"Primary key columns: {', '.join(r['pk_cols']) or '(none detected)'}"])
    arow(ws, [f"Headline: {r['headline']}"])
    if r.get("manual_check_reasons"):
        arow(ws, ["MANUAL CHECK NEEDED: " + "; ".join(r["manual_check_reasons"])])
        _mark_manual_cell = ws.cell(row=ws.max_row, column=1)
        _mark_manual_cell.fill = MANUAL_FILL
        _mark_manual_cell.font = MANUAL_FONT
    for note in r["notes"]:
        arow(ws, [f"Note: {note}"])
    arow(ws, [])

    arow(ws, ["COLUMN-LEVEL SUMMARY (every mismatched column in this table)"])
    ws.cell(row=ws.max_row, column=1).font = SECTION_FONT
    header = [
        "Risk", "Column", "Field Type", "Records", "What differs (s2 vs s3)",
        "Pattern breakdown", "Consistency observations",
        "Examples  s2 -> s3  (repr: quotes/spaces visible)",
        "Lookup query per example (same order)",
    ]
    arow(ws, header)
    style_header_row(ws, ws.max_row, len(header))

    for c in r["columns"]:
        patterns = ", ".join(f"{p}={n}" for p, n in sorted(c["patterns"].items(), key=lambda kv: -kv[1]))
        arow(ws, [
            c["risk"],
            c["column"],
            "text field" if c["is_text_field"] else "data field",
            c["records"],
            " | ".join(dict.fromkeys(c["descriptions"])),
            patterns,
            "; ".join(c["observations"]),
            "\n".join(f"[{p}] {s2!r} -> {s3!r}" for p, s2, s3, _ in c["examples"]),
            "\n".join(
                lookup_query(recon_table_template, r["table_name"], r["pk_cols"], pkv)
                for _, _, _, pkv in c["examples"]
            ),
        ])
        row_idx = ws.max_row
        fill = RISK_FILL.get(c["risk"])
        if fill:
            ws.cell(row=row_idx, column=1).fill = fill
        for i in range(1, len(header) + 1):
            ws.cell(row=row_idx, column=i).alignment = WRAP

    arow(ws, [])
    arow(ws, ["BREAKDOWN BY DISTINCT mismatch_columns COMBINATION"])
    ws.cell(row=ws.max_row, column=1).font = SECTION_FONT
    header2 = [
        "Risk", "mismatch_columns combination", "# Cols", "Records",
        "Per-column pattern", "Text fields", "Examples  s2 -> s3", "Sample PK values",
    ]
    arow(ws, header2)
    style_header_row(ws, ws.max_row, len(header2))

    shown_combos = r["combos"][:max_combos] if max_combos > 0 else r["combos"]
    for combo in shown_combos:
        examples_str = "\n".join(
            f"{col}: " + fmt_examples(exs) for col, exs in combo["examples"].items()
        )
        pk_str = "\n".join(
            ", ".join(f"{pk}={v}" for pk, v in zip(r["pk_cols"], pkv))
            for pkv in combo["pks"]
        )
        arow(ws, [
            combo["risk"],
            combo["combo"],
            combo["column_count"],
            combo["rows"],
            "; ".join(b["text"] for b in combo["column_detail"]),
            ", ".join(combo["text_fields"]) or "-",
            examples_str,
            pk_str,
        ])
        row_idx = ws.max_row
        fill = RISK_FILL.get(combo["risk"])
        if fill:
            ws.cell(row=row_idx, column=1).fill = fill
        for i in range(1, len(header2) + 1):
            ws.cell(row=row_idx, column=i).alignment = WRAP

    hidden = len(r["combos"]) - len(shown_combos)
    if hidden > 0:
        arow(ws, [
            "", f"... {hidden} further combination(s) not shown "
                f"(--max-combos is {max_combos}); run with --max-combos 0 to list them all, "
                f"or see the JSON digest if --json was given",
        ])

    for i, w in enumerate([22, 26, 12, 10, 46, 34, 40, 42, 60], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def auto_assessment(r, max_cols=8):
    """Plain-language per-column rundown built only from measured facts.

    Used for every table; a hand-written note passed via --assessments
    replaces it for that table.
    """
    bits = []
    for c in r["columns"][:max_cols]:
        desc = " / ".join(dict.fromkeys(c["descriptions"]))
        bit = f"{c['column']} ({_plural(c['records'])}, {c['risk']}): {desc}"
        if c["observations"]:
            bit += f" [{c['observations'][0]}]"
        bits.append(bit)
    remaining = len(r["columns"]) - max_cols
    if remaining > 0:
        bits.append(f"(+{remaining} more column(s) - see the {r['table_name']} sheet)")
    return "  |  ".join(bits)


MANUAL_FILL = PatternFill("solid", fgColor="FFFF00")
MANUAL_FONT = Font(bold=True, color="C00000")


def _mark_manual(ws, row_idx):
    """Bright yellow, bold red table name: this one needs a person to look at it."""
    for col in (1, 2):
        ws.cell(row=row_idx, column=col).fill = MANUAL_FILL
        ws.cell(row=row_idx, column=col).font = MANUAL_FONT


def write_overview_sheet(ws, results, assessments):
    header = [
        "Table", "Manual check needed?", "Records", "Combos", "Mismatched Columns",
        "PK duplicates", "Headline", "Warnings (read these)", "Assessment",
    ]

    arow(ws, header)
    style_header_row(ws, 1, len(header))

    # a table that could not be analysed goes first: an empty row must never be
    # mistaken for a clean table
    for r in sorted((r for r in results if r.get("error")), key=lambda r: r["table_name"]):
        arow(ws, [
            r["table_name"], "YES - nothing in this table was compared", "", "", "", "not checked",
            "NOT ANALYSED - nothing in this table was compared",
            f"{r['error']} (see Errors sheet)", "",
        ])
        _mark_manual(ws, ws.max_row)
        ws.cell(row=ws.max_row, column=7).fill = RISK_FILL[R_REAL]
        for i in range(1, len(header) + 1):
            ws.cell(row=ws.max_row, column=i).alignment = WRAP

    ordered = sorted(
        [r for r in results if not r.get("error")],
        key=lambda r: (-int(bool(r.get("manual_check_reasons"))),
                       -r["risk_counts"].get(R_REAL, 0), -r["risk_counts"].get(R_UNSURE, 0),
                       -int(r.get("needs_attention", False)), -r["total_sample_rows"],
                       r["table_name"]),
    )
    for r in ordered:
        dup = r.get("pk_duplicates", {})
        if not dup.get("checked"):
            dup_text = "not checked - no primary key columns found"
        elif dup.get("pk_values_duplicated"):
            dup_text = f"{dup['pk_values_duplicated']} PK value(s) on {dup['rows_involved']} rows"
            if not dup.get("complete"):
                dup_text += " (check stopped early - incomplete)"
        elif not dup.get("complete"):
            dup_text = "none found, but the check stopped early - incomplete"
        else:
            dup_text = "none"
        reasons = r.get("manual_check_reasons") or []
        arow(ws, [
            r["table_name"],
            ("YES:\n- " + "\n- ".join(reasons)) if reasons else "no",
            r["total_sample_rows"],
            r["distinct_combos"],
            len(r["columns"]),
            dup_text,
            r["headline"],
            "\n".join(r["notes"]) if r["notes"] else "none",
            assessments.get(r["table_name"]) or auto_assessment(r),
        ])
        row_idx = ws.max_row
        if reasons:
            _mark_manual(ws, row_idx)
        elif r["risk_counts"].get(R_REAL, 0):
            ws.cell(row=row_idx, column=1).fill = RISK_FILL[R_REAL]
        elif r.get("needs_attention"):
            ws.cell(row=row_idx, column=1).fill = RISK_FILL[R_UNSURE]
        if r["risk_counts"].get(R_REAL, 0):
            ws.cell(row=row_idx, column=7).fill = RISK_FILL[R_REAL]
        if dup.get("pk_values_duplicated") or (dup.get("checked") and not dup.get("complete")):
            ws.cell(row=row_idx, column=6).fill = RISK_FILL[R_REAL]
        if r["notes"]:
            ws.cell(row=row_idx, column=8).fill = RISK_FILL[R_UNSURE]
        for i in range(1, len(header) + 1):
            ws.cell(row=ws.max_row, column=i).alignment = WRAP

    for i, w in enumerate([28, 44, 10, 10, 16, 24, 60, 60, 100], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # only the header row is frozen, and nothing in it is merged: a merged cell
    # or an oversized row inside a frozen pane makes Excel redraw and scroll wrongly
    ws.freeze_panes = "A2"

    overall = assessments.get("_overall") if assessments else None
    if overall:
        arow(ws, [])
        arow(ws, ["Overall assessment"])
        ws.cell(row=ws.max_row, column=1).font = Font(bold=True, size=12)
        # left unwrapped so it spills across the empty cells beside it rather
        # than forming one very tall row
        arow(ws, [overall])


def write_across_tables_sheet(ws, results):
    """One row per (pattern, s2 format, s3 format) seen anywhere in the run.

    Turns "every table writes dates differently" into a single line naming how
    many tables and columns share that exact behaviour.
    """
    groups = defaultdict(lambda: {
        "tables": set(), "columns": set(), "records": 0, "example": None,
    })

    for r in results:
        if r.get("error"):
            continue
        for c in r["columns"]:
            for pattern, n in c["patterns"].items():
                s2_shape, s3_shape = c.get("shapes", {}).get(pattern, ("-", "-"))
                g = groups[(pattern, s2_shape, s3_shape)]
                g["tables"].add(r["table_name"])
                g["columns"].add(f"{r['table_name']}.{c['column']}")
                g["records"] += n
                if g["example"] is None:
                    for p, s2v, s3v, _ in c["examples"]:
                        if p == pattern:
                            g["example"] = f"{s2v!r} -> {s3v!r}"
                            break

    header = [
        "Risk", "Pattern", "s2 format", "s3 format", "Tables affected",
        "Columns affected", "Records", "Example", "Where (table.column)",
    ]
    arow(ws, header)
    style_header_row(ws, 1, len(header))

    ordered = sorted(
        groups.items(),
        key=lambda kv: (RISK_ORDER.get(pattern_risk(kv[0][0]), 0), -len(kv[1]["tables"]), kv[0]),
    )
    for (pattern, s2_shape, s3_shape), g in ordered:
        # every table.column, one per line, so filtering on a pattern shows all of
        # them. A cell holds ~32,000 characters; a longer list carries on in
        # continuation rows that repeat the risk/pattern/formats, so a filter on
        # any of those still returns the whole list
        chunks = _where_chunks(sorted(g["columns"], key=natural_key))
        for part, where in enumerate(chunks):
            first = part == 0
            arow(ws, [
                pattern_risk(pattern), pattern, s2_shape, s3_shape,
                len(g["tables"]) if first else "", len(g["columns"]) if first else "",
                g["records"] if first else "",
                (g["example"] or "") if first else f"(continued, part {part + 1} of {len(chunks)})",
                where,
            ])
            fill = RISK_FILL.get(pattern_risk(pattern))
            if fill:
                ws.cell(row=ws.max_row, column=1).fill = fill
            for i in range(1, len(header) + 1):
                ws.cell(row=ws.max_row, column=i).alignment = WRAP

    for i, w in enumerate([22, 28, 22, 22, 14, 16, 10, 34, 60], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    # filter buttons on the header, e.g. Pattern = date_format_diff
    ws.auto_filter.ref = f"A1:{get_column_letter(len(header))}{ws.max_row}"


WHERE_CELL_CHARS = 30000   # under Excel's 32,767 and safe_cell's own cut-off


def natural_key(s):
    """orders2 before orders10: digit runs compare as numbers, not as text."""
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", s)]


def _where_chunks(names):
    """Newline-joined names, split into as few cells as fit Excel's size limit."""
    chunks, current, size = [], [], 0
    for name in names:
        if current and size + len(name) + 1 > WHERE_CELL_CHARS:
            chunks.append("\n".join(current))
            current, size = [], 0
        current.append(name)
        size += len(name) + 1
    chunks.append("\n".join(current))
    return chunks


GLOSSARY = [

    ('BASICS', '', '', ''),
    ('s2', 'Your live / source table - what is in production today (Spark 2.x, NiFi 1.9, etc.)', '', ''),
    ('s3', 'The new parallel table you ingested and are validating (the _elcr side)', '', ''),
    ('One row of input', "One record whose primary key was found on BOTH sides, but at least one column's value differed", '', ''),
    ('Records', 'How many rows involved this column or combination. With the sampled export this is NOT your production total - it keeps up to 20 rows per distinct mismatch_columns combination', '', ''),
    ('Combos', "How many distinct sets of columns mismatched. 'only balance' and 'balance + open_date' are two different combinations", '', ''),

    ('RISK LABELS', '', '', ''),
    ('REAL VALUE DIFF', 'The actual value is different. s3 does not hold the same data as s2. This needs a decision from you', '100.00 -> 150.00', 'assigned when any pattern found on the column is a proven value change'),
    ('CANNOT VERIFY - check source', 'The file does not contain enough to prove the values match OR differ. The report will not call these fine - look at the source tables', '01/05/2024 -> 2024-01-05', 'assigned whenever a verdict would depend on something the file cannot show'),
    ('FORMAT/TYPE - check joins', 'The value MEANS the same but is WRITTEN differently. Fine arithmetically; NOT fine if anything downstream joins on it, compares it as text, or depends on its type', '10.00 -> 10.0', 'assigned only when the two values are PROVEN equal (exact decimal arithmetic, unambiguous dates) and only the stored text differs'),
    ('WHITESPACE/CASE ONLY', 'Only spacing or capitalisation differs. Not called low risk: trailing spaces and case still break exact-match joins', "'BOB ' -> 'BOB'", 'assigned when only strip() or lower() is needed to make the two equal'),
    ('TEXT FIELD - read manually', 'A free-text column whose wording changed. No rule can judge this - a person has to read it', '', 'column name matches desc/comment/note/remark/... OR its values average over 25 chars, over 3 words, and are mostly unique'),

    ('WHAT THE PATTERN NAMES MEAN', '', '', ''),
    ('value_diff', 'Values differ and NO rule could explain how. The script measured a difference but will not guess at it. These are the ones worth your attention', 'CLOSED -> CANCELLED', 'reached only after EVERY other test failed. Nothing is asserted about it except that the values differ'),
    ('column_not_in_file', 'CANNOT VERIFY. mismatch_columns names a column, but the file has no __s2/__s3 pair for it, so nothing was compared', '', 'the name in mismatch_columns matches no header column ending __s2 with a matching __s3'),
    ('date_ambiguous_order', 'CANNOT VERIFY. Day and month could be either way round, and the values match under one reading but not the other', '01/05/2024 -> 2024-01-05', 'the first two fields are both 12 or less and differ; the verdict is reported only if both readings agree'),
    ('date_incomplete', 'CANNOT VERIFY. One side is missing part of the date (e.g. year only), so equality cannot be proven', '2024 -> 2024-01-01', 'the text parses to different results when the missing parts are filled two different ways'),
    ('ambiguous_decimal_comma', 'CANNOT VERIFY. A comma that is not valid thousands grouping may be a decimal comma (12,34 = 12.34?)', '12,34 -> 1234', 'the commas are not in 3-digit groups, and the other side is a number'),
    ('percent_sign_diff', 'NOT MATCHED - CANNOT VERIFY. A % sign on one side only - 50% might mean 50 or 0.5, so it is not treated as a match', '50% -> 50', 'both sides are numbers and exactly one carries %'),
    ('currency_symbol_changed', 'REAL. The currency symbol itself changed - the amount may be in a different currency', '$10.00 -> EUR 10.00', 'both sides carry a currency symbol and the symbols differ'),
    ('numeric_value_diff', 'Genuinely different numbers', '5 -> 7', 'Decimal(s2) != Decimal(s3) exactly, and it is neither rounding nor float noise'),
    ('numeric_scale_diff', "DECIMAL VALUE MISMATCH. The numerical values match; only the number of decimal places differs. Usually means the column's type or precision changed", '10.00 -> 10.0', 'Decimal(s2) == Decimal(s3) exactly AND the number of decimal places differs'),
    ('numeric_rounding', 's3 has been rounded - real precision was lost', '99.456789 -> 99.46', 'Decimal(s2) rounded (half-up, half-even or truncated) to 0..6 places equals Decimal(s3)'),
    ('float_precision_noise', 'DECIMAL VALUE MISMATCH. The numerical values match; only the decimal digits differ, because the column is held as float/double', '0.3 -> 0.30000000000000004', "the sides have different decimal places, rounding the longer back to the shorter's places gives the shorter exactly, the longer has 15+ significant digits, and they differ by at most 1 part in 10^15. Same decimal places that differ is NEVER noise"),
    ('numeric_thousands_sep_diff', 'Thousands separator on one side only', '1,000 -> 1000', "Decimal(s2) == Decimal(s3) exactly AND (',' in s2) != (',' in s3); commas must be valid 3-digit grouping"),
    ('numeric_scientific_notation', 'One side written in scientific notation', '1000000 -> 1E+06', "Decimal(s2) == Decimal(s3) exactly AND ('e' in s2.lower()) != ('e' in s3.lower())"),
    ('numeric_sign_format_diff', 'An explicit + sign on one side only', '+5 -> 5', "Decimal(s2) == Decimal(s3) exactly AND s2.startswith('+') != s3.startswith('+')"),
    ('accounting_negative_diff', 'Negative written as (n) on one side, -n on the other', '(500) -> -500', 'numbers equal once (n) is read as -n, AND only one side is wrapped in parentheses'),
    ('currency_symbol_diff', 'NOT MATCHED - CANNOT VERIFY. A currency symbol is on one side only. The digits agree, but the currency is not confirmed on both sides, so it is not treated as a match', '$10.00 -> 10.00', 'the digits are equal once symbols are removed, AND only one side carries a currency symbol'),
    ('leading_zero_diff', 'Leading zeros lost or gained. The digits match but the stored text does not', '000123 -> 123', "s2.strip().lstrip('0') == s3.strip().lstrip('0')"),
    ('date_format_diff', "DATE FORMAT ISSUE. The date/time is the same; only the format differs. Also raises 'Manual check needed?' on the Overview", '2024-02-10 -> 20240210', 'both parse as dates and are equal under every possible reading: day/month order is unambiguous, no date part is missing, and any fraction of a second matches digit for digit'),
    ('date_value_diff', 'Genuinely different dates or times', '2024-03-01 -> 2024-03-02', 'both parse as date/time but the instants differ; the gap is measured and reported'),
    ('timezone_marker_diff', "DATE FORMAT ISSUE. Same clock time, but one side carries a timezone and the other does not. Also raises 'Manual check needed?'", '...T00:00:00Z -> ...00:00:00', 'one parsed side has a timezone and the other does not, while the wall-clock parts are equal'),
    ('timezone_representation_diff', "DATE FORMAT ISSUE. The same instant written at a different UTC offset. Also raises 'Manual check needed?'", '', 'both sides timezone-aware, instants equal, UTC offsets different'),
    ('epoch_vs_formatted_date', "DATE FORMAT ISSUE. The date/time is the same; one side is an epoch number, the other a readable date. Also raises 'Manual check needed?'", '1704067200 -> 2024-01-01', "one side is exactly 10 or 13 digits; read as epoch seconds/millis (UTC) it equals the other side's complete date EXACTLY"),
    ('boolean_format_diff', 'Same true/false meaning, written differently', 'Y -> true', 'both sides map into {y,n,yes,no,true,false,t,f,1,0} AND map to the SAME boolean'),
    ('boolean_value_diff', 'Opposite true/false values', 'Y -> N', 'both sides map into that set AND map to OPPOSITE booleans'),
    ('missing_in_s3', 's2 has a value and s3 is empty - data missing on the new side', "2024-06-01 -> ''", 's3 is empty or a null token while s2 is not'),
    ('missing_in_s2', 's3 has a value and s2 is empty - extra data on the new side', "'' -> 2024-06-01", 's2 is empty or a null token while s3 is not'),
    ('null_representation_diff', "Both sides are 'nothing' but written differently. These behave DIFFERENTLY in joins, aggregations and IS NULL checks", "NULL -> ''", "BOTH sides are in {'', null, none, nan, n/a, na, nil, (null)} but the text differs"),
    ('truncated_in_s3', 'The s3 value is the start of the s2 value, cut short - content was lost', 'CHRISTOPHER ALEXANDER -> CHRISTOPHER', 's2.startswith(s3) and s3 is the shorter of the two'),
    ('truncated_in_s2', 'The reverse: s3 holds more text than s2', '', 's3.startswith(s2) and s2 is the shorter of the two'),
    ('partial_value_diff', "One side's value sits inside the other", '', 'one value appears inside the other, but not as its prefix'),
    ('encoding_loss', 'Accented or non-English characters were replaced or lost - an encoding problem', 'MULLER (with umlaut) -> M?LLER', "s2 holds non-ASCII characters, s3 holds '?' or the replacement char, and s3 is pure ASCII"),
    ('unicode_fold_diff', 'The same letters, accented on one side and plain on the other', 'JOSE (accented) -> JOSE', 'folding accents to plain letters makes them equal, AND every non-English character has a plain-letter base - a character with none (e.g. Chinese) counts as lost, not folded'),
    ('typographic_diff', 'Curly quotes, long dashes or non-breaking spaces versus their plain keyboard equivalents', "it's (curly) -> it's (straight)", 'equal after mapping curly quotes, long dashes and non-breaking spaces to their ASCII forms'),
    ('separator_format_diff', 'Same letters and digits, different punctuation between them', 'C:\\data\\in -> C:/data/in', 'the letters and digits are the same AND sit in the same groups (or one side has no separators at all); a separator that moved, like 12-345 vs 123-45, does not qualify'),
    ('punct_removed_in_s3:X', 'The character X is in s2 but gone from s3; everything else is identical', '12"2 -> 122  (X is ")', 's3 is exactly s2 with some X characters deleted and nothing else changed or moved'),
    ('punct_added_in_s3:X', 'The reverse - s3 has an extra X', '', 's2 is exactly s3 with some X characters deleted and nothing else changed or moved'),
    ('whitespace_only', 'Only leading/trailing spaces differ', "'BOB ' -> 'BOB'", 's2.strip() == s3.strip()'),
    ('whitespace_internal', 'Only spacing inside the value differs', "'A  B' -> 'A B'", "' '.join(s2.split()) == ' '.join(s3.split())"),
    ('case_difference', 'Only upper/lower case differs', 'jane -> JANE', 's2.strip().lower() == s3.strip().lower()'),
    ('infinity_value', 'One or both sides hold an infinite / overflow value', '', 'either side is in {inf, +inf, -inf, infinity, +infinity, -infinity}'),
    ('text_field_diff', 'A free-text column where the wording changed', '', 'the column was judged free-text AND the difference had no other explanation'),
    ('no_actual_diff', 'CANNOT VERIFY. The recon flagged this column, but the exported s2 and s3 text is identical - the real difference was lost in the export. Most often NULL vs empty string (both export as blank), or a precision/type difference the CSV cannot show. Check the source tables', '', 's2 == s3 exactly, yet mismatch_columns named this column'),

    ('READING THE COLUMNS', '', '', ''),
    ('Pattern breakdown', "Each pattern and how many records showed it. 'numeric_scale_diff=5' means 5 records had that", '', ''),
    ('MIXED [...]', 'This column showed MORE THAN ONE kind of difference across its records. Read the list - a real value change can sit alongside formatting differences', '', ''),
    ('Consistency observations', 'Facts measured across all the records: a constant offset, a uniform rounding, whether the behaviour is the same every time', '', ''),
    ('Lookup query per example', 'A ready SELECT for each example, in the same order as the examples beside it. Paste it in to pull that exact record', '', ''),
    ('Warnings (read these)', "On the Overview: everything the tool could not fully check for that table - skipped rows, malformed lines, columns the recon did not flag, duplicate keys. 'none' means nothing was skipped", '', ''),
    ('Manual check needed?', "On the Overview. YES (table name in yellow) means every check ran, but something measurable looks odd and a person should look before signing off: differences no rule explains, date format issues (even though the dates match), values that cannot be verified, free-text changes, real and format-only changes mixed in one column, a constant shift, a column empty in s3, cut-off values, invalid characters, differences the recon did not flag, duplicate or uncheckable keys, skipped or malformed rows, or several mismatch files. The reasons are listed in the cell. A table with only ordinary real value differences says 'no' here - those are already in the headline", '', 'rule-based: any one of the conditions listed triggers it'),
    ('RECON DID NOT FLAG', 'A column whose s2 and s3 text differ on a row where mismatch_columns did not name it. Your recon may be missing a difference - or deliberately normalising before it compares', '', 'every __s2/__s3 pair on every row is compared, not just the listed ones'),
    ('[mismatch_columns blank - every column compared]', 'Rows whose mismatch_columns was empty. Every column is compared for them instead of the row being skipped', '', ''),
    ('NOT ANALYSED', 'On the Overview: a table or file nothing was compared for (missing file, unreadable folder, bad header). Never read an empty row as clean', '', ''),
    ('Risk ranking', 'A column is ranked by its WORST record, not its typical one. One real value change among twenty whitespace differences still ranks REAL', '', 'max() over the risk of every pattern seen on that column'),

    ('FORMAT NOTATION (Across Tables sheet)', '', '', ''),
    ('####-##-##', 'Each # is one digit - this is the LAYOUT of the value, not the value itself', '2024-01-05', ''),
    ('yyyy-MM-dd', 'A friendly name, used only where the layout is unambiguous', '', ''),
    ('##/##/####', 'Left as digits on purpose: day-month order cannot be told apart from the value alone, so it is not guessed', '01/05/2024', ''),
    ('2dp / 1dp / 0dp', "Number of decimal places. Numbers are grouped this way so '10.00 -> 10.0' and '9999.00 -> 9999.0' count as one finding", '', ''),
    ('+sep', 'A thousands separator is present', '1,234.50', ''),
    ('scientific', 'Written in scientific notation', '1E+06', ''),
    ('(text)', 'The value holds no digits. Collapsed deliberately so real data is not copied into a summary sheet', '', ''),
    ('(empty)', 'Blank value', '', ''),
    ('-', 'Not applicable - for a genuine value change, the format is not the story', '', ''),
]


def write_glossary_sheet(ws):
    arow(ws, ["Term", "What it means", "Example", "How it is decided (the exact test applied)"])
    style_header_row(ws, 1, 4)
    for term, meaning, example, test in GLOSSARY:
        arow(ws, [term, meaning, example, test])
        row = ws.max_row
        if not meaning:  # section heading
            ws.cell(row=row, column=1).font = Font(bold=True, size=12, color="1F4E78")
        for i in range(1, 5):
            ws.cell(row=row, column=i).alignment = WRAP
    for i, w in enumerate([30, 76, 30, 80], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


def write_errors_sheet(ws, results):
    """Every problem met while reading the run, in one list.

    NOT ANALYSED: nothing in that table was compared (no mismatch CSV, unreadable
    folder, bad header, CSVs loose in the run folder). ANALYSED WITH WARNING: the
    table was compared, but something could not be fully checked - the same
    warnings as the Overview's "Warnings (read these)" column, one per row.
    """
    header = ["Table / folder", "Status", "Problem", "Path"]
    arow(ws, header)
    style_header_row(ws, 1, len(header))

    rows = []
    for r in sorted((r for r in results if r.get("error")), key=lambda r: r["table_name"]):
        rows.append((r["table_name"], "NOT ANALYSED", r["error"], r.get("csv_path", ""), RISK_FILL[R_REAL]))
    for r in sorted((r for r in results if not r.get("error")), key=lambda r: r["table_name"]):
        for note in r["notes"]:
            rows.append((r["table_name"], "ANALYSED WITH WARNING", note, r["csv_path"], RISK_FILL[R_UNSURE]))

    if not rows:
        arow(ws, ["(none)", "no problems",
                  "Every table folder had a readable mismatch CSV, and nothing was skipped "
                  "or left unchecked", ""])
    for table, status, problem, path, fill in rows:
        arow(ws, [table, status, problem, path])
        ws.cell(row=ws.max_row, column=2).fill = fill
        for i in range(1, len(header) + 1):
            ws.cell(row=ws.max_row, column=i).alignment = WRAP

    for i, w in enumerate([28, 24, 90, 60], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


def build_workbook(results, output_path, assessments=None, recon_table_template="{table}",
                   max_combos=500):
    assessments = assessments or {}
    wb = Workbook()
    # openpyxl emits an empty <workbookProtection/> by default, which some Excel
    # builds treat as a protected workbook and then block copying cells
    wb.security = None
    ws_over = wb.active
    ws_over.title = "Overview"
    write_overview_sheet(ws_over, results, assessments)

    # always present and right after Overview, so "were there any problems?" has
    # one fixed place to look - an absent sheet is easy to mistake for a missed one
    write_errors_sheet(wb.create_sheet("Errors"), results)

    write_across_tables_sheet(wb.create_sheet("Across Tables"), results)

    used = {"overview", "errors", "across tables"}
    ordered = sorted(
        [r for r in results if not r.get("error")],
        key=lambda r: (-r["risk_counts"].get(R_REAL, 0), -r["risk_counts"].get(R_UNSURE, 0),
                       r["table_name"]),
    )
    for r in ordered:
        ws = wb.create_sheet(sanitize_sheet_name(r["table_name"], used))
        write_table_sheet(ws, r, recon_table_template, max_combos)

    write_glossary_sheet(wb.create_sheet("What the terms mean"))

    # the workbook is rebuilt from scratch every run, so saving replaces the
    # file rather than adding to it. If it is open in Excel the write is
    # refused, and the analysis must not be thrown away for that
    try:
        wb.save(output_path)
        return output_path
    except OSError as exc:
        first_error = exc

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base, ext = os.path.splitext(output_path)
    name = os.path.basename(base)
    # same folder under a new name (the usual case: the file is open in Excel),
    # then the folder the command was run from, so the work is never discarded
    candidates = [
        f"{base}_{stamp}{ext or '.xlsx'}",
        os.path.join(os.getcwd(), f"{name}_{stamp}{ext or '.xlsx'}"),
    ]
    for fallback in candidates:
        try:
            wb.save(fallback)
        except OSError:
            continue
        print(
            f"\nCould not write {output_path} ({first_error.__class__.__name__}: {first_error}).\n"
            f"If it is open in Excel, close it next time. Saved here instead:\n  {fallback}",
            file=sys.stderr,
        )
        return fallback
    raise OSError(
        f"could not write {output_path} ({first_error}) or any fallback location: "
        + ", ".join(candidates)
    )


# ---------------------------------------------------------------------------
# JSON digest (compact input for a review pass)
# ---------------------------------------------------------------------------

def build_digest(results, input_folder):
    return {
        "run_folder": os.path.abspath(input_folder),
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "sampling_note": "each table's CSV holds up to 20 sample rows per distinct mismatch_columns combination",
        "tables": [
            {
                "table": r["table_name"],
                "sample_rows": r["total_sample_rows"],
                "distinct_combos": r["distinct_combos"],
                "pk_cols": r["pk_cols"],
                "headline": r["headline"],
                "notes": r["notes"],
                "unlisted_diffs": r.get("unlisted", {}),
                "manual_check_reasons": r.get("manual_check_reasons", []),
                "columns": [
                    {
                        "column": c["column"],
                        "field_type": "text" if c["is_text_field"] else "data",
                        "records": c["records"],
                        "risk": c["risk"],
                        "patterns": c["patterns"],
                        "observations": c["observations"],
                        "examples": [
                            {
                                "pattern": p, "s2": s2, "s3": s3,
                                "pk": dict(zip(r["pk_cols"], pkv)),
                            }
                            for p, s2, s3, pkv in c["examples"]
                        ],
                    }
                    for c in r["columns"]
                ],
                "combos": [
                    {
                        "combo": c["combo"],
                        "records": c["rows"],
                        "risk": c["risk"],
                        "per_column": {b["column"]: b["patterns"] for b in c["column_detail"]},
                    }
                    for c in r["combos"]
                ],
            }
            for r in results if not r.get("error")
        ],
        "errors": [
            {"table": r["table_name"], "error": r["error"]}
            for r in results if r.get("error")
        ],
    }


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Summarize recon mismatch CSVs into Excel")
    ap.add_argument("input_folder", help="Run folder containing one subfolder per table")
    ap.add_argument("-o", "--output", default="recon_summary.xlsx", help="Output .xlsx path")
    ap.add_argument("--json", dest="json_path", help="Also write a compact JSON digest here")
    ap.add_argument("--assessments", help="JSON file of {table_name: note, _overall: note} to merge in")
    ap.add_argument("--examples", type=int, default=3,
                    help="Examples kept per pattern/group (default 3, minimum 1)")
    ap.add_argument("--max-combos", type=int, default=500,
                    help="Most mismatch_columns combinations written per sheet "
                         "(default 500, 0 for no limit). Keeps the workbook openable "
                         "when a table has thousands of combinations")
    ap.add_argument(
        "--recon-table", default="{table}",
        help="Table name template for the generated lookup queries, where {table} is "
             "the folder name. e.g. 'myschema.{table}' or '{table}_recon' (default '{table}')",
    )
    args = ap.parse_args()

    if not os.path.isdir(args.input_folder):
        sys.exit(f"ERROR: {args.input_folder} is not a directory")
    args.examples = max(1, args.examples)

    if not args.output.lower().endswith(".xlsx"):
        args.output += ".xlsx"
    # create output folders BEFORE the analysis: a bad path must fail in the
    # first second, not after every table has been read
    for target in filter(None, (args.output, args.json_path)):
        out_dir = os.path.dirname(os.path.abspath(target))
        try:
            os.makedirs(out_dir, exist_ok=True)
        except OSError as exc:
            sys.exit(f"ERROR: cannot create the output folder {out_dir}: {exc}")

    try:
        subfolders = sorted(
            d for d in os.listdir(args.input_folder)
            if os.path.isdir(os.path.join(args.input_folder, d))
        )
    except OSError as exc:
        sys.exit(f"ERROR: cannot read {args.input_folder}: {exc}")
    if not subfolders:
        sys.exit(f"ERROR: no table subfolders found in {args.input_folder}")

    stray = sorted(
        f for f in os.listdir(args.input_folder)
        if f.lower().endswith(".csv") and os.path.isfile(os.path.join(args.input_folder, f))
    )

    assessments = {}
    if args.assessments:
        try:
            with open(args.assessments, encoding="utf-8") as fh:
                assessments = json.load(fh)
        except (OSError, ValueError) as exc:
            sys.exit(f"ERROR: could not read --assessments {args.assessments}: {exc}")

    results = []
    if stray:
        results.append({
            "table_name": "(files directly in the input folder)",
            "csv_path": args.input_folder,
            "error": f"{len(stray)} CSV file(s) sit directly in the input folder and were NOT "
                     f"analysed - only table subfolders are read: {', '.join(stray[:10])}",
        })
        print(f"WARNING: {len(stray)} CSV file(s) directly in the input folder were not analysed",
              file=sys.stderr)

    for i, table_name in enumerate(subfolders, start=1):
        folder = os.path.join(args.input_folder, table_name)
        try:
            csv_path, pick_note = find_mismatch_csv(folder)
        except OSError as exc:
            # an unreadable subfolder must not take down the rest of the run
            results.append({
                "table_name": table_name, "csv_path": folder,
                "error": f"folder could not be listed: {exc.__class__.__name__}: {exc}",
            })
            print(f"[{i}/{len(subfolders)}] ERROR {table_name}: {exc}", file=sys.stderr)
            continue

        if not csv_path:
            try:
                present = sorted(f for f in os.listdir(folder) if f.lower().endswith(".csv"))
            except OSError:
                present = []
            results.append({
                "table_name": table_name,
                "csv_path": folder,
                "error": "no *mismatch*.csv file in this folder. CSVs present: "
                         + (", ".join(present) if present else "(none)"),
            })
            print(f"[{i}/{len(subfolders)}] MISSING {table_name}: no *mismatch*.csv")
            continue
        try:
            r = analyze_table(table_name, csv_path, examples_per_group=args.examples)
            if pick_note:
                r["notes"].insert(0, pick_note)
                r["needs_attention"] = True
                r["manual_check_reasons"].insert(0, "several mismatch files in the folder - only one was analysed")
            results.append(r)
            print(
                f"[{i}/{len(subfolders)}] {table_name}: {r['total_sample_rows']} rows, "
                f"{r['distinct_combos']} combos, {len(r['columns'])} cols, "
                f"{r['risk_counts'].get(R_REAL, 0)} real-diff, "
                f"{r['risk_counts'].get(R_UNSURE, 0)} cannot-verify"
                + (f", {len(r['notes'])} warning(s)" if r["notes"] else "")
                + ("  <-- MANUAL CHECK" if r["manual_check_reasons"] else "")
            )
        except Exception as exc:
            results.append({"table_name": table_name, "csv_path": csv_path, "error": str(exc)})
            print(f"[{i}/{len(subfolders)}] ERROR {table_name}: {exc}", file=sys.stderr)

    if not results:
        sys.exit("ERROR: nothing to summarize")

    try:
        written = build_workbook(results, args.output, assessments, args.recon_table,
                                 args.max_combos)
    except OSError as exc:
        sys.exit(f"ERROR: the analysis finished but the report could not be saved: {exc}")
    print(f"\nWorkbook: {written}")

    if args.json_path:
        try:
            with open(args.json_path, "w", encoding="utf-8") as fh:
                json.dump(build_digest(results, args.input_folder), fh,
                          indent=1, ensure_ascii=False)
            print(f"JSON digest: {args.json_path}")
        except (OSError, ValueError) as exc:
            # the workbook is already written; a digest failure must not mask that
            print(f"WARNING: workbook saved but JSON digest failed: {exc}", file=sys.stderr)


if __name__ == "__main__":
    main()
