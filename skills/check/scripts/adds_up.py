#!/usr/bin/env python3
"""Adds Up: pull the text, tables and formulas out of documents and list
arithmetic leads for a reviewer to confirm.

Usage:
    python3 adds_up.py extract FILE [FILE ...] [--json] [--all-rows] [--max-rows N]
    python3 adds_up.py grim --n 12 --percent 37.5
    python3 adds_up.py grim --n 25 --mean 3.48

Supported files: .docx .pptx .xlsx .xlsm .csv .tsv .md .txt .pdf

Standard library only. The script reads the files named on the command line
and prints to standard output. It writes nothing, changes nothing and makes
no network requests. PDF text needs pdfplumber or pypdf if one is installed;
without them the script says so and the reviewer reads the PDF directly.

Everything listed under "Leads" is a candidate, not a finding. Each one has
to be confirmed against the document before it is reported.
"""

import argparse
import collections
import csv
import datetime
import io
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from decimal import Decimal, InvalidOperation

# --------------------------------------------------------------------------
# Number parsing
# --------------------------------------------------------------------------

GROUP_CHARS = "\u00a0\u202f\u2009 '\u2019"     # no-break, narrow and thin spaces, apostrophes
CURRENCY_CHARS = "€$£¥₺₹₽₩"
MINUS_CHARS = "-\u2212\u2013"                     # hyphen, minus sign, en dash
CURRENCY_CODES = ("EUR|USD|GBP|CHF|TRY|JPY|CAD|AUD|NZD|SEK|NOK|DKK|PLN|CZK|HUF|RON|INR|CNY|RMB|HKD|SGD|"
                  "KRW|BRL|MXN|ZAR|AED|SAR|ILS|RUB|UAH|TL")

DATE_CELL_RES = [
    re.compile(r"^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}$"),
    re.compile(r"^(?:19|20)\d{2}[-/]\d{1,2}([-/]\d{1,2})?$"),
    re.compile(r"^(?:19|20)\d{2}\.\d{1,2}\.\d{1,2}$"),
    re.compile(r"^\d{1,2}:\d{2}(:\d{2})?$"),
    re.compile(r"^\d{1,2}[./]\d{1,2}\.$"),
]

CELL_RE = re.compile(
    r"^(\d[\d.," + re.escape(GROUP_CHARS) + r"]*?)\s*"
    r"(%|‰|[" + re.escape(CURRENCY_CHARS) + r"]|(?:[^\W\d_]|[°µ])[^\s\d]{0,7})?$"
)


class Num(object):
    """A parsed number: exact value, displayed decimals, unit and format style."""

    __slots__ = ("value", "decimals", "pct", "unit", "raw", "style", "yearlike")

    def __init__(self, value, decimals, pct, unit, raw, style, yearlike):
        self.value = value
        self.decimals = decimals
        self.pct = pct
        self.unit = unit
        self.raw = raw
        self.style = style
        self.yearlike = yearlike

    @property
    def step(self):
        """Size of one unit in the last displayed digit."""
        return Decimal(1).scaleb(-self.decimals)


def _convert_core(core, conv):
    """Turn a digit string with separators into (Decimal, decimals, style).

    style is 'us' (decimal point), 'eu' (decimal comma), 'amb' (one separator
    followed by exactly three digits, read according to conv) or 'plain'.
    """
    style = "plain"
    if re.search("[" + re.escape(GROUP_CHARS) + "]", core):
        pat = r"\d{1,3}(?:[" + re.escape(GROUP_CHARS) + r"]\d{3})+(?:[.,]\d+)?"
        if not re.fullmatch(pat, core):
            return None
        core = re.sub("[" + re.escape(GROUP_CHARS) + "]", "", core)
        if "," in core:
            style = "eu"
            core = core.replace(",", ".")
        elif "." in core:
            style = "us"
    else:
        has_dot = "." in core
        has_comma = "," in core
        if has_dot and has_comma:
            last = max(core.rfind("."), core.rfind(","))
            dec = core[last]
            grp = "," if dec == "." else "."
            int_part, frac = core[:last], core[last + 1:]
            if dec in int_part or not frac.isdigit():
                return None
            if not re.fullmatch(r"\d{1,3}(?:" + re.escape(grp) + r"\d{3})+", int_part):
                return None
            style = "us" if dec == "." else "eu"
            core = int_part.replace(grp, "") + "." + frac
        elif has_dot or has_comma:
            sep = "." if has_dot else ","
            natural = "us" if has_dot else "eu"      # style if sep is the decimal mark
            other = "eu" if has_dot else "us"        # style if sep groups thousands
            parts = core.split(sep)
            if any(not p.isdigit() for p in parts):
                return None
            if len(parts) > 2:
                if not (1 <= len(parts[0]) <= 3 and all(len(p) == 3 for p in parts[1:])):
                    return None
                if parts[0].startswith("0"):
                    return None
                style = other
                core = "".join(parts)
            else:
                a, b = parts
                if len(b) == 3 and 1 <= len(a) <= 3 and not a.startswith("0"):
                    style = "amb"
                    as_decimal = (conv == natural)
                    core = (a + "." + b) if as_decimal else (a + b)
                else:
                    style = natural
                    core = a + "." + b
    try:
        value = Decimal(core)
    except InvalidOperation:
        return None
    decimals = len(core.split(".")[1]) if "." in core else 0
    return value, decimals, style


def parse_cell(text, conv="us"):
    """Parse one table cell or token into a Num, or return None."""
    if text is None:
        return None
    s = str(text).strip()
    if not s or len(s) > 40:
        return None
    for rx in DATE_CELL_RES:
        if rx.match(s):
            return None
    raw = s
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()
    unit = ""
    m = re.match(r"^([" + re.escape(CURRENCY_CHARS) + r"]|(?:" + CURRENCY_CODES + r")(?=[\s\d+\-]))\s*", s)
    if m:
        unit = m.group(1)
        s = s[m.end():]
    m = re.match(r"^([+" + re.escape(MINUS_CHARS) + r"])\s*", s)
    if m:
        if m.group(1) != "+":
            neg = not neg
        s = s[m.end():]
    m = re.match(r"^([" + re.escape(CURRENCY_CHARS) + r"])\s*", s)
    if m:
        unit = unit or m.group(1)
        s = s[m.end():]
    m = CELL_RE.match(s)
    if not m:
        return None
    core = m.group(1).strip()
    if core[-1] in ".,":
        return None
    tail = m.group(2) or ""
    pct = tail == "%"
    if tail and not pct:
        unit = unit or tail
    conv_res = _convert_core(core, conv)
    if conv_res is None:
        return None
    value, decimals, style = conv_res
    if neg:
        value = -value
    yearlike = (
        style == "plain" and not pct and not unit and not neg
        and len(core) == 4 and 1900 <= int(core) <= 2100
    )
    return Num(value, decimals, pct, unit, raw, style, yearlike)


FMT = {"conv": "us", "group": False}
STATS = collections.Counter()     # what the rules looked at in the current file, and how much agreed


def has_grouping(raw):
    return bool(re.search(r"\d[.," + re.escape(GROUP_CHARS) + r"]\d{3}(?!\d)", raw or ""))


def fmt(value, decimals=None):
    """Format a Decimal without exponent, in the number format of the file being read."""
    if decimals is not None:
        q = Decimal(1).scaleb(-decimals)
        value = value.quantize(q)
    s = format(value, "f")
    if "." in s and decimals is None:
        s = s.rstrip("0").rstrip(".")
    if s in ("-0", ""):
        s = "0"
    sign = "-" if s.startswith("-") else ""
    whole, _, frac = s.lstrip("-").partition(".")
    dec_mark, grp_mark = (",", ".") if FMT["conv"] == "eu" else (".", ",")
    if FMT["group"] and len(whole) > 3:
        whole = re.sub(r"(?<=\d)(?=(\d{3})+$)", grp_mark, whole)
    return sign + whole + (dec_mark + frac if frac else "")


# --------------------------------------------------------------------------
# Leads
# --------------------------------------------------------------------------

def lead(where, kind, message, strength="normal"):
    return {"where": where, "kind": kind, "strength": strength, "message": message}


def classify(stated, computed, n_parts, part_step, stated_step):
    """Compare a stated value with the sum of its displayed parts.

    'exact'    the figures agree
    'rounding' a gap that rounded parts commonly produce
    'stretch'  possible only if nearly every part was rounded the same way
    'off'      more than rounding can produce
    """
    diff = abs(stated - computed)
    if diff <= Decimal("1e-8") * max(Decimal(1), abs(stated)):
        return "exact"
    worst = part_step * n_parts / 2 + stated_step / 2
    if diff > worst:
        return "off"
    # Each rounded figure is off by up to half a step, spread evenly: variance step^2 / 12.
    spread = ((part_step * part_step * n_parts + stated_step * stated_step) / 12).sqrt()
    return "rounding" if diff <= spread * Decimal("2.5") else "stretch"


def close(stated, computed, step, slack="0.5"):
    """True if computed rounds to stated at the displayed precision."""
    return abs(stated - computed) <= step * Decimal(slack) + Decimal("1e-12")


# --------------------------------------------------------------------------
# Table model and checks
# --------------------------------------------------------------------------

SUBTOTAL_RE = re.compile(
    r"\b(sub[\s-]?totals?|zwischensummen?|teilsummen?|ara\s+toplam|sous[\s-]?total|subtotale)\b",
    re.I,
)
TOTAL_RE = re.compile(
    r"\b(grand\s+total|totals?|sum|overall|gesamt\w*|summe|insgesamt|toplam|genel\s+toplam|"
    r"somme|suma|totale|totaal|total\s+général)\b",
    re.I,
)
RATE_RE = re.compile(r"(?<![\d.,])(\d{1,2}(?:[.,]\d{1,2})?)\s?%")
AVG_RE = re.compile(
    r"(\b(average|mean|avg\.?|durchschnitt\w*|mittelwert|ortalama|moyenne|promedio)\b|Ø)",
    re.I,
)
# A label that is nothing but "Total" (possibly "Total net", "Summe inkl. MwSt.", "Grand total (EUR)").
# "Total revenue" or "Average revenue per unit" name a quantity and need not be a sum of their neighbours.
BARE_TOTAL_RE = re.compile(
    r"^\W*(?:grand\s+|sub[\s-]?|ara\s+|genel\s+|sous[\s-]?|zwischen|teil|gesamt)?"
    r"(?:totals?|sum|summe|gesamt\w*|insgesamt|toplam|somme|suma|totale|totaal|overall)"
    r"(?:\s+(?:net\w*|gross|brutto|amount|betrag|due|tutar|general|g\u00e9n\u00e9ral|"
    r"(?:excl|incl|exkl|inkl)\w*\.?.*|\(.*\)|[A-Z]{3}))?\W*$", re.I)
BARE_AVG_RE = re.compile(
    r"^\W*(?:average|mean|avg\.?|durchschnitt|mittelwert|ortalama|moyenne|promedio|\u00d8)(?:\s*\(.*\))?\W*$", re.I)
MEMO_RE = re.compile(
    r"^\W*(of\s+which|thereof|davon|darunter|dav\.|including|incl\.|inkl\.|dont|bunun|memo)\b", re.I)
SKIP_HEADER_RE = re.compile(
    r"^\s*(id|no\.?|nr\.?|#|pos\.?|item\s*no\.?|date|datum|tarih|code|kod|sku|ref\.?)\s*$", re.I)
SMALL_INT_HEADER_RE = re.compile(
    r"^\s*(year|jahr|yıl|week|kw|month|monat|ay|age|alter|yaş|day|tag|gün)\s*$", re.I)


class Table(object):
    def __init__(self, name, rows, where="", row_ids=None, col_ids=None, meta=None, joiner=" ", stored=False):
        width = max((len(r) for r in rows), default=0)
        self.rows = [list(r) + [""] * (width - len(r)) for r in rows]
        self.name = name
        self.where = where
        self.width = width
        self.row_ids = row_ids or ["R%d" % (i + 1) for i in range(len(rows))]
        self.col_ids = col_ids or ["C%d" % (i + 1) for i in range(width)]
        self.meta = meta or {}          # (r, c) -> dict, used for spreadsheets
        self.joiner = joiner            # " " gives "R2 C3", "" gives "B7"
        self.stored = stored            # values as stored, not as displayed (0.30 arrives as 0.3)

    def ref(self, r, c):
        if self.joiner == "":
            return "%s%s%s" % (self.name, self.col_ids[c], self.row_ids[r])
        return "%s %s %s" % (self.name, self.row_ids[r], self.col_ids[c])

    def rowref(self, r):
        if self.joiner == "":
            return "%srow %s" % (self.name, self.row_ids[r])
        return "%s %s" % (self.name, self.row_ids[r])

    def span(self, rs):
        """Describe a set of row indexes compactly: 'R2 to R5', 'rows 2 to 12, 14, 17 to 30'."""
        if not rs:
            return "no rows"
        runs, start, prev = [], rs[0], rs[0]
        for r in rs[1:]:
            if r != prev + 1:
                runs.append((start, prev))
                start = r
            prev = r
        runs.append((start, prev))
        parts = [self.row_ids[x] if x == y else "%s to %s" % (self.row_ids[x], self.row_ids[y]) for x, y in runs]
        if len(parts) > 6:
            parts = parts[:5] + ["... %s" % parts[-1]]
        return ("rows " if self.joiner == "" else "") + ", ".join(parts)


def _row_label(row, nums_row):
    for c, cell in enumerate(row):
        if cell and cell.strip() and nums_row[c] is None:
            return cell.strip()
    return ""


def check_table(t, conv):
    """Return a list of leads for one table."""
    leads = []
    n_rows = len(t.rows)
    if n_rows < 2 or t.width < 1:
        return leads
    nums = [[parse_cell(cell, conv) for cell in row] for row in t.rows]

    # Columns that mostly hold numbers below the first row.
    numeric_cols = []
    for c in range(t.width):
        filled = [r for r in range(1, n_rows) if t.rows[r][c].strip()]
        numeric = [r for r in filled if nums[r][c] is not None]
        if numeric and len(numeric) * 2 >= len(filled):
            numeric_cols.append(c)
    if not numeric_cols:
        return leads

    # Header row: the first row holds labels or years, not data.
    head_cells = [nums[0][c] for c in numeric_cols]
    data_like = [n for n in head_cells if n is not None and not n.yearlike]
    has_header = len(data_like) * 2 < len(head_cells) or len(data_like) == 0
    first = 1 if has_header else 0
    headers = [t.rows[0][c].strip() if has_header else "" for c in range(t.width)]

    labels = [_row_label(t.rows[r], nums[r]) for r in range(n_rows)]
    kind = {}
    for r in range(first, n_rows):
        lab = labels[r]
        if MEMO_RE.search(lab):
            kind[r] = "memo"
        elif SUBTOTAL_RE.search(lab):
            kind[r] = "subtotal"
        elif TOTAL_RE.search(lab):
            kind[r] = "total"
        elif AVG_RE.search(lab):
            kind[r] = "avg" if BARE_AVG_RE.match(lab) else "memo"
    rate_rows = {}
    for r in range(first, n_rows):
        m = RATE_RE.search(labels[r])
        if m:
            rate_rows[r] = Decimal(m.group(1).replace(",", "."))
    detail_rows = [r for r in range(first, n_rows) if r not in kind]

    pct_cols = set()
    year_cols = set()
    for c in numeric_cols:
        vals = [nums[r][c] for r in range(first, n_rows) if nums[r][c] is not None]
        if vals and sum(1 for v in vals if v.pct) * 2 > len(vals):
            pct_cols.add(c)
        if "%" in headers[c]:
            pct_cols.add(c)
        if vals and all(v.yearlike for v in vals):
            year_cols.add(c)
        if SKIP_HEADER_RE.match(headers[c] or ""):
            year_cols.add(c)
        if SMALL_INT_HEADER_RE.match(headers[c] or "") and vals and all(
                v.decimals == 0 and (v.yearlike or abs(v.value) <= 200) for v in vals):
            year_cols.add(c)     # a column of years, weeks or ages, not of amounts
    total_cols = set(c for c in numeric_cols if has_header and TOTAL_RE.search(headers[c]))
    avg_cols = set(c for c in numeric_cols if has_header and AVG_RE.search(headers[c]))

    reported = set()
    usable = [c for c in numeric_cols if c not in year_cols]
    pending = []     # total-row leads that may turn out to be the knock-on of another lead
    fixes = {}       # (row, col) -> the value another lead says the cell should hold

    def text_cause(cells, diff):
        """Name a number stored as text that explains why a formula total is too low."""
        hits = [(rc, nums[rc[0]][rc[1]]) for rc in cells if t.meta.get(rc, {}).get("text") and nums[rc[0]][rc[1]]]
        if hits and diff < 0 and same(sum((n.value for _, n in hits), Decimal(0)), -diff):
            return (" The gap equals %s, held as text, which SUM skips. The formula is right and the text "
                    "cell is the cause." % ", ".join("%s (%s)" % (t.ref(*rc), n.raw) for rc, n in hits))
        return ""

    def pstep(parts):
        """Step of the parts: finest for stored values, coarsest for displayed ones."""
        steps = [p.step for p in parts]
        return min(steps) if t.stored else max(steps)

    def gap_words(state, diff, n_parts, subject="The stated figure", ref="its parts give"):
        """(strength, sentence) for a figure that differs from what its parts give."""
        gap = "%s %s than %s" % (fmt(abs(diff)), "higher" if diff > 0 else "lower", ref)
        if state == "rounding":
            return "weak", ("%s is %s. Rounded parts commonly produce a gap this small; "
                            "it is an error only if the figures are exact." % (subject, gap))
        if state == "stretch":
            return "normal", ("%s is %s. Rounding could produce this only if nearly all %d parts "
                              "were rounded the same way, which is unlikely." % (subject, gap, n_parts))
        return "normal", "%s is %s, more than rounding can produce." % (subject, gap)

    def col_name(c):
        return ' ("%s")' % headers[c] if headers[c] else ""

    def typed_note(r, c):
        m = t.meta.get((r, c))
        if not m:
            return ""
        if m.get("formula"):
            return " The cell holds a formula: %s." % m["formula"]
        return " The cell holds a typed number, not a formula."

    # ---- rows derived from a rate named in their label ----------------
    for r, rate in sorted(rate_rows.items()):
        for c in numeric_cols:
            stated = nums[r][c]
            if c in year_cols or c in pct_cols or stated is None or stated.pct:
                continue
            base, base_desc = None, ""
            for x in reversed(range(first, r)):
                if kind.get(x) in ("total", "subtotal") and x not in rate_rows and nums[x][c] is not None:
                    base, base_desc = nums[x][c].value, '"%s" (%s)' % (labels[x], nums[x][c].raw)
                    break
            if base is None:
                rs = [x for x in range(first, r) if x not in kind and x not in rate_rows and nums[x][c] is not None]
                if len(rs) >= 2:
                    base = sum((nums[x][c].value for x in rs), Decimal(0))
                    base_desc = "the sum of %s (%s)" % (t.span(rs), fmt(base))
            if base is None:
                continue
            part = base * rate / 100
            options = [("%s%% of the base" % fmt(rate), part), ("base plus %s%%" % fmt(rate), base + part),
                       ("base minus %s%%" % fmt(rate), base - part),
                       ("the %s%% contained in a gross base" % fmt(rate), base - base / (1 + rate / 100)),
                       ("a gross base less %s%%" % fmt(rate), base / (1 + rate / 100))]
            STATS["calcs"] += 1
            if any(close(abs(stated.value), abs(v), stated.step, "1.0") for _, v in options):
                STATS["calcs_ok"] += 1
                reported.add((r, c))
                continue
            # The stated base may itself be the faulty figure: try the detail rows it stands for.
            alt = ""
            rs = [x for x in range(first, r) if x not in kind and x not in rate_rows and nums[x][c] is not None]
            if len(rs) >= 2:
                true_base = sum((nums[x][c].value for x in rs), Decimal(0))
                if true_base != base:
                    tp = true_base * rate / 100
                    for d, v in (("%s%% of it" % fmt(rate), tp), ("it plus %s%%" % fmt(rate), true_base + tp),
                                 ("it minus %s%%" % fmt(rate), true_base - tp)):
                        if close(abs(stated.value), abs(v), stated.step, "1.0"):
                            alt = (" It does match the sum of the detail rows %s (%s): %s = %s. So this row was "
                                   "worked out from the detail rows and the stated base is the figure to check."
                                   % (t.span(rs), fmt(true_base, stated.decimals), d, fmt(v, stated.decimals)))
            leads.append(lead(
                t.ref(r, c), "rate-row",
                'Row "%s" states %s. With %s as the base: %s. None matches.%s' % (
                    labels[r], stated.raw, base_desc,
                    "; ".join("%s = %s" % (d, fmt(v, stated.decimals)) for d, v in options[:3]), alt),
                "weak" if (alt or r not in kind) else "normal"))
            reported.add((r, c))

    # ---- which columns are the sum of their neighbours -----------------
    sumcols = {}     # column -> (other columns, "before"/"after", per-row results, rows that hold)
    for c in usable:
        declared = c in total_cols
        if c in pct_cols or c in avg_cols:
            continue
        left = [x for x in usable if x < c and x not in total_cols and x not in avg_cols and x not in pct_cols]
        right = [x for x in usable if x > c and x not in total_cols and x not in avg_cols and x not in pct_cols]
        others = left if (left or not declared) else right
        if len(others) < 2:
            continue
        res = []
        for r in range(first, n_rows):
            if nums[r][c] is None or kind.get(r) in ("avg", "memo") or r in rate_rows:
                continue
            parts = [nums[r][x] for x in others if nums[r][x] is not None]
            if len(parts) < 2:
                continue
            total = sum((p.value for p in parts), Decimal(0))
            state = classify(nums[r][c].value, total, len(parts), pstep(parts), nums[r][c].step)
            prod_ok = False
            if len(parts) == 2:
                prod_ok = close(nums[r][c].value, parts[0].value * parts[1].value, nums[r][c].step)
            res.append((r, total, state, prod_ok, len(parts)))
        sum_holds = sum(1 for x in res if x[2] in ("exact", "rounding"))
        prod_holds = sum(1 for x in res if x[3])
        if prod_holds > sum_holds:
            continue    # product column, handled by the relation check
        if len(res) >= 2 and sum_holds == 0:
            continue    # the column is not a sum of its neighbours
        exact_holds = sum(1 for x in res if x[2] == "exact")
        if not declared and not (exact_holds >= 2 and exact_holds * 2 >= len(res)):
            continue    # no header says "total" and the pattern is not there
        sumcols[c] = (others, "before" if left else "after", res, sum_holds)

    def interior(rows_, c):
        """Sum of the detail cells that a total cell in a sum column stands for."""
        cells = [nums[x][o] for x in rows_ for o in sumcols[c][0] if nums[x][o] is not None]
        if not cells:
            return None
        return sum((p.value for p in cells), Decimal(0))

    def same(a, b):
        return abs(a - b) <= Decimal("1e-8") * max(Decimal(1), abs(a))

    # ---- total and subtotal rows -------------------------------------
    for r in sorted(kind):
        if kind[r] not in ("total", "subtotal") or r in rate_rows:
            continue
        checks = []
        for c in numeric_cols:
            if c in year_cols or nums[r][c] is None:
                continue
            stated = nums[r][c]
            above = [x for x in range(first, r)]
            below = [x for x in range(r + 1, n_rows)]
            cands = []
            # A: every detail row above
            a_rows = [x for x in above if x not in kind and nums[x][c] is not None]
            if a_rows:
                cands.append(("all detail rows above", a_rows))
            # B: detail rows since the previous total-like row
            b_rows = []
            for x in reversed(above):
                if x in kind and kind[x] in ("total", "subtotal"):
                    break
                if x not in kind and nums[x][c] is not None:
                    b_rows.append(x)
            b_rows.reverse()
            if b_rows and b_rows != a_rows:
                cands.append(("detail rows since the previous total", b_rows))
            # C: subtotal rows above
            c_rows = []
            for x in reversed(above):
                if kind.get(x) == "total":
                    break
                if kind.get(x) == "subtotal" and nums[x][c] is not None:
                    c_rows.append(x)
            c_rows.reverse()
            if c_rows and kind[r] == "total":
                cands.append(("the subtotals", c_rows))
            # D: a bare "Total" stated first, with its breakdown underneath
            if not a_rows and BARE_TOTAL_RE.match(labels[r]):
                d_rows = []
                for x in below:
                    if x in kind and kind[x] in ("total", "subtotal"):
                        break
                    if x not in kind and nums[x][c] is not None:
                        d_rows.append(x)
                if len(d_rows) >= 2:
                    cands.append(("the rows below", d_rows))
            cands = [cd for cd in cands if len(cd[1]) >= 2 or cd[0] == "the subtotals"]
            if not cands:
                continue
            best = None
            for desc, rs in cands:
                total = sum((nums[x][c].value for x in rs), Decimal(0))
                step = pstep([nums[x][c] for x in rs])
                state = classify(stated.value, total, len(rs), step, stated.step)
                rank = {"exact": 0, "rounding": 1, "stretch": 2, "off": 3}[state]
                key = (rank, abs(stated.value - total))
                if best is None or key < best[0]:
                    best = (key, desc, rs, total, state)
            _, desc, rs, total, state = best
            if c in pct_cols:
                near100 = abs(stated.value - 100) <= Decimal("0.5") or abs(total - 100) <= 1
                if not near100:
                    continue
            checks.append((c, stated, desc, rs, total, state))
        n_checked = len([x for x in checks if x[0] not in pct_cols])
        n_hold = len([x for x in checks if x[0] not in pct_cols and x[5] in ("exact", "rounding")])
        for c, stated, desc, rs, total, state in checks:
            STATS["totals"] += 1
            if state == "exact":
                STATS["totals_ok"] += 1
                continue
            diff = stated.value - total
            dec = max(nums[x][c].decimals for x in rs)
            unit = "%" if c in pct_cols else ""
            inner_note = ""
            if c in sumcols and desc != "the subtotals":
                inner = interior(rs, c)
                if inner is not None and same(stated.value, inner):
                    # The corner cell agrees with the detail cells; the totals beside it are what is off.
                    STATS["totals_ok"] += 1
                    reported.add((r, c))
                    continue
                if inner is not None:
                    inner_note = " The detail cells of the table sum to %s." % fmt(inner, dec)
            strength, tail = gap_words(state, diff, len(rs))
            if state == "rounding":
                # Small whole numbers are usually counts, and counts are exact.
                counts = stated.decimals == 0 and all(
                    nums[x][c].decimals == 0 and abs(nums[x][c].value) < 1000 for x in rs)
                if counts and c not in pct_cols:
                    strength = "normal"
                    tail = ("The stated figure is %s %s than its parts give. An error if these are exact "
                            "counts; rounding explains it only if the figures are rounded." % (
                                fmt(abs(diff)), "higher" if diff > 0 else "lower"))
            elif not (n_hold >= 1 or n_checked <= 1 or c in pct_cols):
                strength = "weak"   # no column of this row behaves like a sum
            if n_hold == 0 and not BARE_TOTAL_RE.match(labels[r]) and c not in pct_cols:
                # "Total revenue" under "Units" and "Price" names a quantity; it is not their sum.
                if len(rs) == 2 and close(stated.value, nums[rs[0]][c].value * nums[rs[1]][c].value, stated.step, "1.0"):
                    STATS["totals_ok"] += 1
                    reported.add((r, c))
                    continue
                strength = "weak"
            if desc == "the subtotals" and len(rs) == 1 and total != 0:
                # One subtotal and a larger total: tax, shipping or a discount usually sits between.
                leads.append(lead(
                    t.ref(r, c), "total-row",
                    '"%s"%s states %s, which is %s (%s%%) %s the subtotal in %s (%s). Check that tax, '
                    "a discount or another line accounts for exactly that gap." % (
                        labels[r], col_name(c), stated.raw, fmt(abs(diff), stated.decimals),
                        fmt(abs(diff) / abs(total) * 100, 2), "above" if diff > 0 else "below",
                        t.span(rs), nums[rs[0]][c].raw),
                    "weak"))
                reported.add((r, c))
                continue
            leads.append(lead(
                t.ref(r, c), "total-row",
                '"%s"%s states %s. %s (%s) sum to %s%s. %s%s%s%s' % (
                    labels[r], col_name(c), stated.raw, desc.capitalize(), t.span(rs),
                    fmt(total, dec), unit, tail, inner_note, typed_note(r, c),
                    text_cause([(x, c) for x in rs], diff)),
                strength))
            pending.append((leads[-1], c, rs, stated, total, dec, unit))
            reported.add((r, c))

    # ---- unlabelled last row that behaves like a total ---------------
    if not kind and len(detail_rows) >= 4:
        r = detail_rows[-1]
        above = detail_rows[:-1]
        res = []
        for c in numeric_cols:
            if c in year_cols or c in pct_cols or nums[r][c] is None:
                continue
            rs = [x for x in above if nums[x][c] is not None]
            if len(rs) < 3:
                continue
            total = sum((nums[x][c].value for x in rs), Decimal(0))
            step = pstep([nums[x][c] for x in rs])
            res.append((c, rs, total, classify(nums[r][c].value, total, len(rs), step, nums[r][c].step)))
        holds = [x for x in res if x[3] == "exact"]
        if len(holds) >= 2 and len(holds) * 2 > len(res):
            STATS["totals"] += len(res)
            STATS["totals_ok"] += len(holds)
            for c, rs, total, state in res:
                if state == "exact":
                    continue
                strength, tail = gap_words(state, nums[r][c].value - total, len(rs))
                leads.append(lead(
                    t.ref(r, c), "total-row",
                    'Last row "%s" equals the sum of the rows above in %d other column(s), but%s '
                    "states %s where %s sum to %s. %s" % (
                        labels[r], len(holds), col_name(c), nums[r][c].raw, t.span(rs),
                        fmt(total), tail),
                    strength))
                reported.add((r, c))

    # ---- average rows -------------------------------------------------
    for r in sorted(kind):
        if kind[r] != "avg":
            continue
        for c in numeric_cols:
            if c in year_cols or nums[r][c] is None:
                continue
            rs = []
            for x in reversed(range(first, r)):
                if x in kind:
                    break
                if nums[x][c] is not None:
                    rs.append(x)
            rs.reverse()
            if len(rs) < 2:
                continue
            mean = sum((nums[x][c].value for x in rs), Decimal(0)) / len(rs)
            stated = nums[r][c]
            STATS["calcs"] += 1
            if close(stated.value, mean, stated.step):
                STATS["calcs_ok"] += 1
                continue
            leads.append(lead(
                t.ref(r, c), "average-row",
                '"%s"%s states %s. The plain mean of %s is %s. A weighted mean would differ; '
                "check which one is meant." % (
                    labels[r], col_name(c), stated.raw, t.span(rs),
                    fmt(mean, stated.decimals + 2)),
                "normal"))
            reported.add((r, c))

    # ---- total columns ------------------------------------------------
    for c in sorted(sumcols):
        others, side, res, sum_holds = sumcols[c]
        for r, total, state, prod_ok, n_parts in res:
            if (r, c) in reported:
                continue
            STATS["totals"] += 1
            if state == "exact":
                STATS["totals_ok"] += 1
                continue
            inner_note = ""
            if kind.get(r) in ("total", "subtotal"):
                rs = [x for x in range(first, r) if x not in kind and x not in rate_rows]
                inner = interior(rs, c) if rs else None
                if inner is not None and same(nums[r][c].value, inner):
                    STATS["totals_ok"] += 1
                    reported.add((r, c))
                    continue
                if inner is not None:
                    inner_note = (" The detail cells of the table sum to %s, so the totals around this cell "
                                  "do not agree with each other; find the faulty total before changing this one."
                                  % fmt(inner))
            strength, tail = gap_words(state, nums[r][c].value - total, n_parts)
            if state != "rounding" and sum_holds == 0:
                strength = "weak"
            leads.append(lead(
                t.ref(r, c), "total-column",
                'Row "%s": column "%s" states %s. The %d cells %s it sum to %s. %s%s%s%s' % (
                    labels[r], headers[c] or t.col_ids[c], nums[r][c].raw, n_parts,
                    side, fmt(total), tail, inner_note, typed_note(r, c),
                    text_cause([(r, o) for o in others], nums[r][c].value - total)),
                strength))
            if r not in kind:
                fixes[(r, c)] = total
            reported.add((r, c))

    # ---- share-of-total columns --------------------------------------
    total_row = None
    for r in sorted(kind):
        if kind[r] == "total":
            total_row = r
    for k in sorted(pct_cols):
        for i in usable:
            if i == k or i in pct_cols:
                continue
            rs = [r for r in detail_rows if nums[r][k] is not None and nums[r][i] is not None]
            if len(rs) < 3:
                continue
            base = None
            if total_row is not None and nums[total_row][i] is not None:
                base = nums[total_row][i].value
            if base is None:
                base = sum((nums[r][i].value for r in rs), Decimal(0))
            if base == 0:
                continue
            holds, breaks = [], []
            for r in rs:
                share = nums[r][i].value / base * 100
                if close(nums[r][k].value, share, nums[r][k].step, "1.0"):
                    holds.append(r)
                else:
                    breaks.append((r, share))
            if len(holds) >= 2 and len(holds) > len(breaks):
                STATS["calcs"] += len(holds) + len(breaks)
                STATS["calcs_ok"] += len(holds)
                for r, share in breaks:
                    if (r, k) in reported:
                        continue
                    leads.append(lead(
                        t.ref(r, k), "share",
                        'Row "%s": %s states %s. %s / %s gives %s%%. The share matches in %d other row(s).' % (
                            labels[r], col_name(k).strip() or "share column", nums[r][k].raw,
                            nums[r][i].raw, fmt(base), fmt(share, nums[r][k].decimals + 1), len(holds)),
                        "normal"))
                    fixes[(r, k)] = share.quantize(nums[r][k].step, rounding="ROUND_HALF_UP")
                    reported.add((r, k))

    # ---- percent columns that should reach 100 ------------------------
    for c in sorted(pct_cols):
        if any(kind.get(r) in ("total", "subtotal") and nums[r][c] is not None for r in kind):
            continue
        rs = [r for r in detail_rows if nums[r][c] is not None]
        if len(rs) < 3:
            continue
        total = sum((nums[r][c].value for r in rs), Decimal(0))
        if total == 100 or not (Decimal(95) <= total <= Decimal(105)):
            continue
        state = classify(Decimal(100), total, len(rs), pstep([nums[r][c] for r in rs]), Decimal(0))
        strength, tail = gap_words(state, total - 100, len(rs), "The sum", "100%")
        leads.append(lead(
            "%s %s" % (t.name, t.col_ids[c]) if t.joiner else "%scolumn %s" % (t.name, t.col_ids[c]),
            "percent-sum",
            "Percentages%s over %s sum to %s%%, not 100%%. %s This matters only if they are shares of one whole." % (
                col_name(c), t.span(rs), fmt(total), tail),
            strength))

    # ---- fractions that should reach 1 (shares stored as 0.36) ---------
    for c in numeric_cols:
        if c in pct_cols or c in year_cols:
            continue
        rs = [r for r in detail_rows if nums[r][c] is not None]
        if len(rs) < 3 or len(rs) != len(detail_rows) or kind:
            continue
        vals = [nums[r][c] for r in rs]
        if not all(0 < v.value < 1 for v in vals):
            continue
        total = sum((v.value for v in vals), Decimal(0))
        if total == 1 or not (Decimal("0.95") <= total <= Decimal("1.05")):
            continue
        state = classify(Decimal(1), total, len(vals), pstep(vals), Decimal(0))
        strength, tail = gap_words(state, total - 1, len(vals), "The sum", "1")
        leads.append(lead(
            "%s %s" % (t.name, t.col_ids[c]) if t.joiner else "%scolumn %s" % (t.name, t.col_ids[c]),
            "percent-sum",
            "Shares%s over %s sum to %s (%s%%), not 1 (100%%). %s This matters only if they are shares of one whole." % (
                col_name(c), t.span(rs), fmt(total), fmt(total * 100), tail),
            strength))

    # ---- row relations that hold in most rows and break in a few ------
    rel_cols = [c for c in usable if c not in avg_cols][:10]
    found = {}
    seen_rel = {}    # set of three columns -> (rows that hold, rows that break), counted once
    rel_rows = [r for r in range(first, n_rows) if kind.get(r) not in ("avg", "memo") and r not in rate_rows]
    sample = rel_rows[:40]

    def rel_value(rel, a, b):
        try:
            if rel == "sum":
                return a + b
            if rel == "diff":
                return a - b
            if rel == "prod":
                return a * b
            if rel == "ratio%":
                return a / b * 100
            if rel == "ratio":
                return a / b
            return (b - a) / a * 100
        except (InvalidOperation, ZeroDivisionError):
            return None

    def rel_scan(rel, i, j, k, rows):
        ok, bad, nontrivial = [], [], 0
        for r in rows:
            if r in kind and rel not in ("sum", "diff"):
                continue
            a, b, z = nums[r][i], nums[r][j], nums[r][k]
            if a is None or b is None or z is None:
                continue
            v = rel_value(rel, a.value, b.value)
            if v is None:
                continue
            if close(z.value, v, z.step, "0.5" if rel in ("sum", "diff") else "1.0"):
                ok.append(r)
                if a.value != 0 and b.value != 0 and not (rel == "prod" and (a.value == 1 or b.value == 1)):
                    nontrivial += 1
            else:
                bad.append((r, v))
        return ok, bad, nontrivial

    if len(rel_cols) >= 3:
        for k in rel_cols:
            for i in rel_cols:
                for j in rel_cols:
                    if len({i, j, k}) < 3:
                        continue
                    for rel in ("sum", "diff", "prod", "ratio%", "ratio", "change%"):
                        if rel in ("sum", "prod") and i > j:
                            continue
                        if len(rel_rows) > len(sample) and len(rel_scan(rel, i, j, k, sample)[0]) < 2:
                            continue
                        rows_ok, rows_bad, nontrivial = rel_scan(rel, i, j, k, rel_rows)
                        need = 2
                        if rel in ("ratio%", "ratio", "change%"):
                            # whole-number results match by accident far more easily
                            fine = any(nums[r][k].decimals > 0 or nums[r][k].pct for r in rows_ok)
                            need = 2 if fine else 3
                        if len(rows_ok) >= need and nontrivial >= need and len(rows_ok) > len(rows_bad):
                            trio = frozenset((i, j, k))
                            if trio not in seen_rel or len(rows_ok) > seen_rel[trio][0]:
                                seen_rel[trio] = (len(rows_ok), len(rows_bad))
                        if len(rows_ok) >= need and nontrivial >= need and rows_bad and len(rows_ok) > len(rows_bad):
                            # "a + b = c" and "a x b = c" are the plain way to state a relation;
                            # their rearrangements (c - a = b, c / a = b) describe the same break.
                            plain = 1 if rel in ("sum", "prod") else 0
                            for r, v in rows_bad:
                                key = (r, frozenset((i, j, k)))
                                score = (plain, len(rows_ok), -len(rows_bad))
                                if key not in found or score > found[key][0]:
                                    found[key] = (score, rel, i, j, k, v, len(rows_ok))
    for n_ok, n_bad in seen_rel.values():
        STATS["calcs"] += n_ok + n_bad
        STATS["calcs_ok"] += n_ok
    words = {
        "sum": "%s + %s", "diff": "%s - %s", "prod": "%s x %s",
        "ratio%": "%s / %s x 100", "ratio": "%s / %s", "change%": "change from %s to %s in %%",
    }
    for (r, _cols), (score, rel, i, j, k, v, n_ok) in sorted(
            found.items(), key=lambda kv: (kv[0][0], kv[1][4])):
        if (r, k) in reported or (r, i) in reported or (r, j) in reported:
            continue
        z = nums[r][k]

        def nm(c):
            return headers[c] or t.col_ids[c]
        formula = words[rel] % (nm(i), nm(j))
        leads.append(lead(
            t.ref(r, k), "row-relation",
            'Row "%s": "%s" states %s. In %d other row(s) this column equals %s; here that gives %s '
            "(from %s and %s).%s" % (
                labels[r], nm(k), z.raw, n_ok, formula,
                fmt(v, z.decimals + (0 if rel in ("sum", "diff", "prod") else 1)),
                nums[r][i].raw, nums[r][j].raw, typed_note(r, k)),
            "normal"))
        fixes[(r, k)] = v
        reported.add((r, k))

    # ---- totals that are only the knock-on of another lead -------------
    for ld, c, rs, stated, total, dec, unit in pending:
        hit = [(x, c) for x in rs if (x, c) in fixes]
        if not hit:
            continue
        adjusted = total + sum((fixes[rc] - nums[rc[0]][rc[1]].value for rc in hit), Decimal(0))
        step = pstep([nums[x][c] for x in rs])
        if classify(stated.value, adjusted, len(rs), step, stated.step) in ("exact", "rounding"):
            ld["strength"] = "weak"
            ld["message"] += (" Knock-on: with the value the lead on %s points to, the parts sum to %s%s, which "
                              "agrees. This total is then correct and should not be changed." % (
                                  ", ".join(t.ref(*rc) for rc in hit), fmt(adjusted, dec), unit))

    # ---- mixed units in one column -----------------------------------
    for c in numeric_cols:
        units = {}
        for r in detail_rows:
            n = nums[r][c]
            if n is not None and n.unit:
                units.setdefault(n.unit.lower().rstrip("."), []).append(r)
        if len(units) >= 2:
            desc = ", ".join("%s in %s" % (u, t.span(rs)) for u, rs in sorted(units.items()))
            leads.append(lead(
                "%s %s" % (t.name, t.col_ids[c]) if t.joiner else "%scolumn %s" % (t.name, t.col_ids[c]),
                "mixed-units",
                "Column%s mixes units: %s. Totals and comparisons across these rows need one unit." % (
                    col_name(c), desc),
                "normal"))
    return leads


# --------------------------------------------------------------------------
# Prose checks
# --------------------------------------------------------------------------

MONTHS = {}
for _i, _names in enumerate([
    "january jan januar jänner ocak janvier enero",
    "february feb februar şubat février febrero",
    "march mar märz mart mars marzo",
    "april apr nisan avril abril",
    "may mai mayıs mayo",
    "june jun juni haziran juin junio",
    "july jul juli temmuz juillet julio",
    "august aug ağustos août agosto",
    "september sep sept eylül septembre septiembre",
    "october oct oktober okt ekim octobre octubre",
    "november nov kasım novembre noviembre",
    "december dec dezember dez aralık décembre diciembre",
], 1):
    for _n in _names.split():
        MONTHS[_n] = _i

WEEKDAYS = {}
for _i, _names in enumerate([
    "monday mon montag pazartesi lundi lunes",
    "tuesday tue tues dienstag salı mardi martes",
    "wednesday wed mittwoch çarşamba mercredi miércoles",
    "thursday thu thur thurs donnerstag perşembe jeudi jueves",
    "friday fri freitag cuma vendredi viernes",
    "saturday sat samstag sonnabend cumartesi samedi sábado",
    "sunday sun sonntag pazar dimanche domingo",
]):
    for _n in _names.split():
        WEEKDAYS[_n] = _i

WEEKDAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
_MONTH_ALT = "|".join(sorted((re.escape(m) for m in MONTHS), key=len, reverse=True))
_WD_ALT = "|".join(sorted((re.escape(w) for w in WEEKDAYS), key=len, reverse=True))

DATE_EXPRS = [
    ("dmy_name", r"(?P<d>\d{1,2})(?:st|nd|rd|th|\.)?\s+(?:of\s+)?(?P<mon>" + _MONTH_ALT + r")\.?,?\s+(?P<y>\d{4})"),
    ("mdy_name", r"(?P<mon>" + _MONTH_ALT + r")\.?\s+(?P<d>\d{1,2})(?:st|nd|rd|th)?,?\s+(?P<y>\d{4})"),
    ("dmy_dot", r"(?P<d>\d{1,2})\.(?P<m>\d{1,2})\.(?P<y>\d{4})"),
    ("iso", r"(?P<y>\d{4})-(?P<m>\d{2})-(?P<d>\d{2})"),
    ("slash", r"(?P<a>\d{1,2})/(?P<b>\d{1,2})/(?P<y>\d{4})"),
]


def _dates_from(kind, m):
    """Return (list of possible dates, readable text) for a date match."""
    g = m.groupdict()
    y = int(g["y"])
    out = []
    try:
        if kind in ("dmy_name", "mdy_name"):
            out.append(datetime.date(y, MONTHS[g["mon"].lower()], int(g["d"])))
        elif kind in ("dmy_dot", "iso"):
            out.append(datetime.date(y, int(g["m"]), int(g["d"])))
    except ValueError:
        pass
    if kind == "slash":
        for d, mo in ((g["a"], g["b"]), (g["b"], g["a"])):
            try:
                out.append(datetime.date(y, int(mo), int(d)))
            except ValueError:
                pass
    return out


_PCT = r"(\d{1,3}(?:[.,]\d+)?)\s?%"
POINTS_RES = [
    re.compile(r"\bfrom\s+" + _PCT + r"\s+to\s+" + _PCT + r"[^.;\n]{0,40}?\b(?:increase|rise|growth|gain|jump|"
               r"decrease|drop|fall|decline|reduction|improvement|change)\s+of\s+" + _PCT +
               r"(?!\s*-?\s?(?:points?|pts|pp\b))", re.I),
    re.compile(r"\bvon\s+" + _PCT + r"\s+auf\s+" + _PCT + r"[^.;\n]{0,40}?\b(?:Anstieg|Zuwachs|Steigerung|Wachstum|"
               r"R\u00fcckgang|Abnahme|Plus|Minus|Ver\u00e4nderung|Verbesserung)\s+(?:von|um)\s+" + _PCT +
               r"(?!\s*-?\s?Punkt)", re.I),
]
SPAN_LINK_RE = re.compile(r"^\s*(?:to|until|till|through|and|bis|bis\s+zum|bis\s+einschlie\u00dflich|und|"
                          r"[-\u2013\u2014]|ile|ve)(?:\s+the)?\s*$", re.I)
SPAN_UNIT_RE = re.compile(r"(?<![\d.,])(\d{1,4})[\s-]?(weeks?|wochen?|hafta|days?|tage?n?|g\u00fcn|"
                          r"months?|monate?n?|ay)\b", re.I)


def _dec(text):
    return Decimal(text.replace(",", "."))


def check_spans(where, text):
    """Durations stated next to a start and an end date: '1 March 2026 to 31 May 2026, 14 weeks'."""
    leads = []
    found = []
    for kind, expr in DATE_EXPRS:
        if kind == "slash":
            continue
        for m in re.finditer(r"(?<![\d./-])" + expr + r"(?![\d/-])", text, re.I):
            dates = _dates_from(kind, m)
            if dates:
                found.append((m.start(), m.end(), dates[0], m.group(0)))
    found.sort()
    for (s1, e1, d1, t1), (s2, e2, d2, t2) in zip(found, found[1:]):
        if s2 < e1 or not SPAN_LINK_RE.match(text[e1:s2]):
            continue
        if d2 < d1:
            leads.append(lead(where, "duration", '"%s ... %s": the end date is before the start date.' % (t1, t2)))
            continue
        m = SPAN_UNIT_RE.search(text[e2:e2 + 60])
        if not m:
            continue
        n, unit = int(m.group(1)), m.group(2).lower()
        days = (d2 - d1).days
        STATS["spans"] += 1
        if unit.startswith(("week", "woche", "hafta")):
            ok = min(abs(n * 7 - days), abs(n * 7 - (days + 1))) <= 3
            actual = "%d days (%d counting both end days), which is %s weeks" % (
                days, days + 1, fmt_plain((Decimal(days) / 7).quantize(Decimal("0.1"))))
        elif unit.startswith(("day", "tag", "g")):
            ok = n in (days, days + 1)
            actual = "%d days, or %d counting both the first and the last day" % (days, days + 1)
        else:
            whole = (d2.year - d1.year) * 12 + d2.month - d1.month
            ok = n in (whole, whole + 1)
            actual = "%d days, about %d months" % (days, whole + (1 if d2.day >= d1.day else 0))
        if ok:
            STATS["spans_ok"] += 1
            continue
        leads.append(lead(where, "duration", '"%s ... %s ... %s": from %s to %s is %s.' % (
            t1, t2, m.group(0), d1.isoformat(), d2.isoformat(), actual)))
    return leads


def check_prose(lines):
    """lines: list of (where, text). Return leads on dates, durations and percentages."""
    leads = []
    for where, text in lines:
        if not text:
            continue
        leads.extend(check_spans(where, text))
        for rx in POINTS_RES:
            for m in rx.finditer(text):
                a, b, n = _dec(m.group(1)), _dec(m.group(2)), _dec(m.group(3))
                if a == 0:
                    continue
                points = abs(b - a)
                relative = abs(b - a) / a * 100
                step = Decimal(1).scaleb(-(len(m.group(3).split(",")[-1].split(".")[-1])
                                           if re.search(r"[.,]", m.group(3)) else 0))
                STATS["points"] += 1
                if close(n, relative, step, "1.0"):
                    STATS["points_ok"] += 1
                    continue
                if close(n, points, step):
                    leads.append(lead(where, "percentage-points",
                                      '"%s": %s%% to %s%% is %s percentage points. As a percentage change it is %s%%. '
                                      "Written as a plain percent, the figure reads as the relative change." % (
                                          m.group(0).strip(), m.group(1), m.group(2), fmt_plain(points),
                                          fmt_plain(relative.quantize(Decimal("0.1"))))))
                else:
                    leads.append(lead(where, "percentage-points",
                                      '"%s": %s%% to %s%% is %s percentage points, or %s%% as a percentage change. '
                                      "The stated %s%% is neither." % (
                                          m.group(0).strip(), m.group(1), m.group(2), fmt_plain(points),
                                          fmt_plain(relative.quantize(Decimal("0.1"))), m.group(3))))
        seen = set()
        for kind, expr in DATE_EXPRS:
            before = re.compile(r"\b(?P<wd>" + _WD_ALT + r")\b\.?,?\s+(?:the\s+|den\s+|dem\s+|le\s+|el\s+)?" + expr, re.I)
            after = re.compile(expr + r"\s*,?\s*\(?(?P<wd>" + _WD_ALT + r")\b", re.I)
            plain = re.compile(r"(?<![\d./-])" + expr + r"(?![\d/-])", re.I)
            for rx in (before, after):
                for m in rx.finditer(text):
                    if m.span() in seen:
                        continue
                    dates = _dates_from(kind, m)
                    if not dates:
                        continue
                    seen.add(m.span())
                    wd = WEEKDAYS[m.group("wd").lower()]
                    STATS["weekdays"] += 1
                    if any(d.weekday() == wd for d in dates):
                        STATS["weekdays_ok"] += 1
                        continue
                    actual = " or ".join(sorted(set(
                        "%s (%s)" % (WEEKDAY_NAMES[d.weekday()], d.isoformat()) for d in dates)))
                    leads.append(lead(
                        where, "weekday",
                        '"%s": that date is a %s, not a %s.' % (
                            m.group(0).strip(), actual, WEEKDAY_NAMES[wd])))
            for m in plain.finditer(text):
                if kind == "slash":
                    continue
                if not _dates_from(kind, m):
                    g = m.groupdict()
                    if kind == "dmy_dot":
                        try:
                            datetime.date(int(g["y"]), int(g["d"]), int(g["m"]))
                            continue   # valid as month.day.year
                        except ValueError:
                            pass
                    leads.append(lead(where, "date", '"%s" is not a calendar date.' % m.group(0).strip()))

        pats = [
            (r"(?<![\d.,/])(?P<n>\d{1,7})\s*(?:of|out\s+of|von|/)\s*(?P<N>\d{1,7})(?![\d/.,]\d)"
             r"[^.%\n()]{0,40}?\(\s*(?P<p>\d{1,3}(?:[.,]\d+)?)\s*%\s*\)"),
            (r"(?P<p>\d{1,3}(?:[.,]\d+)?)\s*%\s*\(\s*(?:n\s*=\s*)?(?P<n>\d{1,7})\s*/\s*(?P<N>\d{1,7})\s*\)"),
        ]
        done = set()
        for pat in pats:
            for m in re.finditer(pat, text, re.I):
                if m.start() in done:
                    continue
                done.add(m.start())
                n, big = int(m.group("n")), int(m.group("N"))
                if big == 0 or n > big:
                    continue
                ptxt = m.group("p").replace(",", ".")
                stated = Decimal(ptxt)
                dec = len(ptxt.split(".")[1]) if "." in ptxt else 0
                actual = Decimal(n) * 100 / Decimal(big)
                step = Decimal(1).scaleb(-dec)
                STATS["nofn"] += 1
                if close(stated, actual, step):
                    STATS["nofn_ok"] += 1
                    continue
                note = ""
                if abs(stated - actual) < step:
                    note = " The stated figure looks truncated instead of rounded."
                leads.append(lead(
                    where, "n-of-N",
                    '"%s": %d / %d = %s%%, not %s%%.%s' % (
                        m.group(0).strip(), n, big, fmt(actual, dec + 1), m.group("p"), note)))
    return leads


STRONG_EU = re.compile(
    r"(?<![\d.,])(?:\d{1,3}(?:\.\d{3})+,\d+|\d+,\d+\s?%|[€$£]\s?\d+,\d{2}(?!\d)|"
    r"\d+,\d{2}\s?(?:€|EUR|USD|GBP|CHF|TL|₺))")
STRONG_US = re.compile(
    r"(?<![\d.,])(?:\d{1,3}(?:,\d{3})+\.\d+|\d+\.\d+\s?%|[€$£]\s?\d+\.\d{2}(?!\d)|"
    r"\d+\.\d{2}\s?(?:€|EUR|USD|GBP|CHF|TL|₺))")


def number_format_evidence(lines, tables):
    eu, us = [], []
    for where, text in lines:
        for m in STRONG_EU.finditer(text or ""):
            eu.append((where, m.group(0)))
        for m in STRONG_US.finditer(text or ""):
            us.append((where, m.group(0)))
    for t in tables:
        for r, row in enumerate(t.rows):
            for c, cell in enumerate(row):
                n = parse_cell(cell, "us")
                if n is None:
                    continue
                if n.style == "eu":
                    eu.append((t.ref(r, c), n.raw))
                elif n.style == "us":
                    us.append((t.ref(r, c), n.raw))
    return eu, us


# --------------------------------------------------------------------------
# Office Open XML helpers
# --------------------------------------------------------------------------

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def _xml(z, name):
    try:
        return ET.fromstring(z.read(name))
    except KeyError:
        return None


def _rels(z, part):
    """Relationships of a part: id -> (type, resolved target path)."""
    folder, base = os.path.split(part)
    root = _xml(z, "%s/_rels/%s.rels" % (folder, base) if folder else "_rels/%s.rels" % base)
    out = {}
    if root is None:
        return out
    for rel in root.iter(REL + "Relationship"):
        if rel.get("TargetMode") == "External":
            continue
        target = rel.get("Target", "")
        if target.startswith("/"):
            path = target.lstrip("/")
        else:
            path = os.path.normpath(os.path.join(folder, target)).replace("\\", "/")
        out[rel.get("Id")] = (rel.get("Type", "").rsplit("/", 1)[-1], path)
    return out


def chart_table(z, part, name, where):
    """Read the cached data of a chart part as a Table."""
    root = _xml(z, part)
    if root is None:
        return None
    title = "".join(x.text or "" for x in root.iter(C + "title") for x in x.iter(A + "t")).strip()
    ctype = "chart"
    plot = next(root.iter(C + "plotArea"), None)
    if plot is not None:
        for el in plot:
            tag = el.tag.split("}")[-1]
            if tag.endswith("Chart"):
                ctype = re.sub(r"(?<=[a-z])(?=[A-Z0-9])", " ", tag[:-5]).lower() + " chart"
                break
    series = []
    cats = []
    for ser in root.iter(C + "ser"):
        tx = ser.find(C + "tx")
        sname = ""
        if tx is not None:
            sname = "".join(v.text or "" for v in tx.iter(C + "v")).strip()
        cat = ser.find(C + "cat")
        if cat is None:
            cat = ser.find(C + "xVal")
        val = ser.find(C + "val")
        if val is None:
            val = ser.find(C + "yVal")
        cmap, vmap = {}, {}
        if cat is not None:
            for pt in cat.iter(C + "pt"):
                v = pt.find(C + "v")
                cmap[int(pt.get("idx", "0"))] = (v.text or "") if v is not None else ""
        if val is not None:
            for pt in val.iter(C + "pt"):
                v = pt.find(C + "v")
                vmap[int(pt.get("idx", "0"))] = (v.text or "") if v is not None else ""
        if len(cmap) > len(cats):
            cats = [cmap.get(i, "") for i in range(max(cmap) + 1)] if cmap else []
        series.append((sname or "Series %d" % (len(series) + 1), vmap))
    if not series:
        return None
    n = max([len(cats)] + [max(v) + 1 if v else 0 for _, v in series])
    rows = [["(category)"] + [s for s, _ in series]]
    for i in range(n):
        row = [cats[i] if i < len(cats) else str(i + 1)]
        for _, vmap in series:
            row.append(_tidy_float(vmap.get(i, "")))
        rows.append(row)
    return Table(name, rows, stored=True,
                 where="%s, %s%s, values as stored in the chart (the workbook embedded behind it is not read)" % (
                     where, ctype, ' titled "%s"' % title if title else ""))


def fmt_plain(d):
    """Decimal to text with a decimal point and no grouping, as spreadsheets store it."""
    s = format(d, "f")
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    return s if s not in ("-0", "") else "0"


def _tidy_float(s):
    """Show a stored float without binary noise such as 0.30000000000000004."""
    if s is None or s == "":
        return ""
    try:
        d = Decimal(s)
    except InvalidOperation:
        return s
    if d == d.to_integral_value() and abs(d) < Decimal("1e15"):
        return fmt_plain(d.to_integral_value())
    try:
        return fmt_plain(Decimal(format(float(s), ".12g")))
    except (ValueError, InvalidOperation):
        return s


# --------------------------------------------------------------------------
# Extractors. Each returns a dict:
#   {"blocks": [("text", where, text) | ("table", Table)], "notes": [...], "leads": [...]}
# --------------------------------------------------------------------------

def _w_text(p):
    out = []
    for el in p.iter():
        if el.tag == W + "t":
            out.append(el.text or "")
        elif el.tag == W + "tab":
            out.append("\t")
        elif el.tag in (W + "br", W + "cr"):
            out.append(" ")
    return "".join(out).strip()


def _w_table(tbl):
    rows = []
    for tr in tbl.findall(W + "tr"):
        row = []
        for tc in tr.findall(W + "tc"):
            texts = [_w_text(p) for p in tc.iter(W + "p")]
            row.append(" ".join(x for x in texts if x))
            span = tc.find(W + "tcPr/" + W + "gridSpan")
            if span is not None:
                try:
                    row.extend([""] * (int(span.get(W + "val", "1")) - 1))
                except ValueError:
                    pass
        rows.append(row)
    return rows


def extract_docx(path):
    blocks, notes = [], []
    with zipfile.ZipFile(path) as z:
        root = _xml(z, "word/document.xml")
        if root is None:
            raise ValueError("not a Word document")
        body = root.find(W + "body")
        counters = {"p": 0, "t": 0}
        last_p = [""]
        last_head = [""]

        def walk(parent):
            for el in parent:
                if el.tag == W + "p":
                    text = _w_text(el)
                    if not text:
                        continue
                    counters["p"] += 1
                    style = el.find(W + "pPr/" + W + "pStyle")
                    sval = style.get(W + "val", "") if style is not None else ""
                    head = ""
                    if re.match(r"(?i)(heading|berschrift|überschrift|title|titel|başlık)", sval):
                        head = "(%s) " % sval
                        last_head[0] = text
                    where = "P%d" % counters["p"]
                    last_p[0] = where
                    blocks.append(("text", where, head + text))
                elif el.tag == W + "tbl":
                    counters["t"] += 1
                    name = "T%d" % counters["t"]
                    place = "after %s" % last_p[0] if last_p[0] else "at the start"
                    if last_head[0]:
                        place += ', under the heading "%s"' % last_head[0][:80]
                    blocks.append(("table", Table(name, _w_table(el), where=place)))
                elif el.tag in (W + "sdt", W + "sdtContent", W + "ins", W + "smartTag", W + "customXml"):
                    walk(el)
        walk(body)
        foot = _xml(z, "word/footnotes.xml")
        if foot is not None:
            for fn in foot.findall(W + "footnote"):
                if fn.get(W + "type"):
                    continue
                text = " ".join(x for x in (_w_text(p) for p in fn.iter(W + "p")) if x)
                if text:
                    blocks.append(("text", "FN%s" % fn.get(W + "id", "?"), "(footnote) " + text))
        n = 0
        for name in sorted(z.namelist()):
            if re.match(r"word/charts/chart\d+\.xml$", name):
                n += 1
                t = chart_table(z, name, "CH%d" % n, "embedded")
                if t is not None:
                    blocks.append(("table", t))
        seen_hf = set()
        n_hf = 0
        for name in sorted(z.namelist()):
            m = re.match(r"word/(header|footer)\d*\.xml$", name)
            if not m:
                continue
            part = _xml(z, name)
            if part is None:
                continue
            for para in part.iter(W + "p"):
                text = _w_text(para)
                if text and text not in seen_hf and not text.isdigit():
                    seen_hf.add(text)
                    n_hf += 1
                    blocks.append(("text", "%s%d" % ("HDR" if m.group(1) == "header" else "FTR", n_hf),
                                   "(page %s) %s" % (m.group(1), text)))
        names = z.namelist()
        n_img = sum(1 for x in names if x.startswith("word/media/"))
        comments = _xml(z, "word/comments.xml")
        n_com = len(comments.findall(W + "comment")) if comments is not None else 0
        n_del = sum(1 for _ in root.iter(W + "del"))
        read = ["%d paragraph(s)" % counters["p"], "%d table(s)" % counters["t"]]
        if n:
            read.append("the stored data of %d chart(s)" % n)
        read.append("footnotes, page headers and footers, and text boxes where present")
        skipped = []
        if n_img:
            skipped.append("%d image(s), so figures shown only inside a picture are missing" % n_img)
        if n_com:
            skipped.append("%d comment(s)" % n_com)
        if n_del:
            skipped.append("%d tracked deletion(s), the text is read as if all changes were accepted" % n_del)
        notes.append("Read: %s. Not read: %s." % (
            ", ".join(read), "; ".join(skipped) if skipped else "nothing, the file has no images, comments or tracked deletions"))
    return {"blocks": blocks, "notes": notes, "leads": []}


def _a_paragraphs(tx):
    out = []
    for p in tx.iter(A + "p"):
        parts = []
        for el in p.iter():
            if el.tag == A + "t":
                parts.append(el.text or "")
            elif el.tag == A + "br":
                parts.append(" ")
        text = "".join(parts).strip()
        if text:
            out.append(text)
    return out


def extract_pptx(path):
    blocks, notes = [], []
    with zipfile.ZipFile(path) as z:
        pres = _xml(z, "ppt/presentation.xml")
        if pres is None:
            raise ValueError("not a PowerPoint file")
        rels = _rels(z, "ppt/presentation.xml")
        order = []
        pics = []
        lst = pres.find(P + "sldIdLst")
        if lst is not None:
            for s in lst:
                rid = s.get(R + "id")
                if rid in rels:
                    order.append(rels[rid][1])
        for idx, part in enumerate(order, 1):
            root = _xml(z, part)
            if root is None:
                continue
            srels = _rels(z, part)
            tag = "S%d" % idx
            counts = {"text": 0, "table": 0, "chart": 0}

            def walk(parent):
                for el in parent:
                    if el.tag == P + "sp":
                        tx = el.find(P + "txBody")
                        if tx is None:
                            continue
                        paras = _a_paragraphs(tx)
                        if not paras:
                            continue
                        ph = el.find(P + "nvSpPr/" + P + "nvPr/" + P + "ph")
                        ptype = ph.get("type", "") if ph is not None else ""
                        if ptype in ("sldNum", "dt", "ftr"):
                            continue
                        if ptype in ("title", "ctrTitle"):
                            blocks.append(("text", tag + " title", " ".join(paras)))
                        else:
                            for para in paras:
                                counts["text"] += 1
                                blocks.append(("text", "%s text %d" % (tag, counts["text"]), para))
                    elif el.tag == P + "graphicFrame":
                        tbl = next(el.iter(A + "tbl"), None)
                        if tbl is not None:
                            counts["table"] += 1
                            rows = []
                            for tr in tbl.findall(A + "tr"):
                                row = []
                                for tc in tr.findall(A + "tc"):
                                    tx = tc.find(A + "txBody")
                                    row.append(" ".join(_a_paragraphs(tx)) if tx is not None else "")
                                rows.append(row)
                            blocks.append(("table", Table("%s table %d" % (tag, counts["table"]), rows,
                                                          where="slide %d" % idx)))
                        ch = next(el.iter(C + "chart"), None)
                        if ch is not None and ch.get(R + "id") in srels:
                            counts["chart"] += 1
                            t = chart_table(z, srels[ch.get(R + "id")][1],
                                            "%s chart %d" % (tag, counts["chart"]), "slide %d" % idx)
                            if t is not None:
                                blocks.append(("table", t))
                    elif el.tag == P + "grpSp":
                        walk(el)
                    elif el.tag == P + "pic":
                        counts["pic"] = counts.get("pic", 0) + 1
            tree = root.find(P + "cSld/" + P + "spTree")
            if tree is not None:
                walk(tree)
            if counts.get("pic"):
                pics.append("slide %d (%d)" % (idx, counts["pic"]))
            for rid, (rtype, target) in srels.items():
                if rtype == "notesSlide":
                    nroot = _xml(z, target)
                    if nroot is None:
                        continue
                    paras = []
                    for sp in nroot.iter(P + "sp"):
                        ph = sp.find(P + "nvSpPr/" + P + "nvPr/" + P + "ph")
                        if ph is not None and ph.get("type") in ("sldNum", "sldImg", "dt", "hdr", "ftr"):
                            continue
                        tx = sp.find(P + "txBody")
                        if tx is not None:
                            paras.extend(_a_paragraphs(tx))
                    if paras:
                        blocks.append(("text", tag + " notes", " ".join(paras)))
        notes.append("Read: %d slide(s) with their text, tables, speaker notes and the stored data of their charts. "
                     "Not read: %s." % (
                         len(order),
                         "pictures on " + ", ".join(pics) + ", so figures shown only inside a picture are missing"
                         if pics else "nothing, the deck has no pictures"))
    return {"blocks": blocks, "notes": notes, "leads": []}


# ---- spreadsheets ---------------------------------------------------------

def col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def num_to_col(n):
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


REF_RE = re.compile(r"(?<![A-Za-z0-9_.!'])(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(!])")
BUILTIN_DATE_FMTS = set(list(range(14, 23)) + list(range(27, 37)) + list(range(45, 48)) + list(range(50, 59)))


def to_r1c1(formula, row, col):
    """Rewrite A1 references relative to the cell, so copied formulas compare equal."""
    parts = re.split(r'("[^"]*")', formula)

    def repl(m):
        c = col_to_num(m.group(2))
        r = int(m.group(4))
        cs = "C%d" % c if m.group(1) else "C[%d]" % (c - col)
        rs = "R%d" % r if m.group(3) else "R[%d]" % (r - row)
        return rs + cs
    return "".join(p if p.startswith('"') else REF_RE.sub(repl, p) for p in parts)


def _fmt_kind(code):
    code = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", code or "")
    if "%" in code:
        return "percent"
    low = code.lower()
    if re.search(r"[dy]", low) or re.search(r"h+:m+|m+:s+", low) or re.search(r"mmm", low):
        return "date"
    return ""


BUILTIN_FORMATS = {1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%",
                   37: "#,##0", 38: "#,##0", 39: "#,##0.00", 40: "#,##0.00"}


def _fmt_shape(code):
    """(decimals, grouping) a number format displays, or None for General and unusual codes."""
    code = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.|_.|\*.', "", (code or "").split(";")[0])
    m = re.search(r"[0#?][0#?,]*(?:\.([0#?]+))?", code)
    if not m or "E" in code.upper() or "/" in code:
        return None
    return (len(m.group(1)) if m.group(1) else 0, "," in m.group(0).split(".")[0])


def shown_as(raw, shape, percent):
    """The text a cell displays for a stored number, in plain English notation."""
    try:
        d = Decimal(format(float(raw) * (100 if percent else 1), ".12g"))
    except (ValueError, InvalidOperation):
        return ""
    d = d.quantize(Decimal(1).scaleb(-shape[0]), rounding="ROUND_HALF_UP")
    txt = format(d, "f")
    if shape[1]:
        whole, _, frac = txt.lstrip("-").partition(".")
        whole = re.sub(r"(?<=\d)(?=(\d{3})+$)", ",", whole)
        txt = ("-" if txt.startswith("-") else "") + whole + ("." + frac if frac else "")
    return txt + ("%" if percent else "")


def from_r1c1(text, row, col):
    """Turn a relative formula pattern back into A1 notation for the cell at (row, col)."""
    def repl(m):
        r = row + int(m.group(1)) if m.group(1) is not None else int(m.group(2))
        c = col + int(m.group(3)) if m.group(3) is not None else int(m.group(4))
        if r < 1 or c < 1:
            return "#REF!"
        return "%s%s%s%d" % ("" if m.group(3) is not None else "$", num_to_col(c),
                             "" if m.group(1) is not None else "$", r)
    parts = re.split(r'("[^"]*")', text)
    rx = re.compile(r"R(?:\[(-?\d+)\]|(\d+))C(?:\[(-?\d+)\]|(\d+))")
    return "".join(q if q.startswith('"') else rx.sub(repl, q) for q in parts)


def extract_xlsx(path, conv_hint="us"):
    blocks, notes, leads = [], [], []
    with zipfile.ZipFile(path) as z:
        wb = _xml(z, "xl/workbook.xml")
        if wb is None:
            raise ValueError("not an Excel workbook")
        rels = _rels(z, "xl/workbook.xml")
        date1904 = False
        pr = wb.find(S + "workbookPr")
        if pr is not None and pr.get("date1904") in ("1", "true"):
            date1904 = True
        shared = []
        sst = _xml(z, "xl/sharedStrings.xml")
        if sst is not None:
            for si in sst.findall(S + "si"):
                parts = []
                for child in si:
                    if child.tag == S + "t":
                        parts.append(child.text or "")
                    elif child.tag == S + "r":
                        parts.extend(t.text or "" for t in child.iter(S + "t"))
                shared.append("".join(parts))
        style_kind = []
        style_shape = []
        styles = _xml(z, "xl/styles.xml")
        if styles is not None:
            custom = {}
            nf = styles.find(S + "numFmts")
            if nf is not None:
                for f in nf:
                    custom[int(f.get("numFmtId", "0"))] = f.get("formatCode", "")
            xfs = styles.find(S + "cellXfs")
            if xfs is not None:
                for xf in xfs:
                    fid = int(xf.get("numFmtId", "0"))
                    if fid in custom:
                        style_kind.append(_fmt_kind(custom[fid]))
                    elif fid in (9, 10):
                        style_kind.append("percent")
                    elif fid in BUILTIN_DATE_FMTS:
                        style_kind.append("date")
                    else:
                        style_kind.append("")
                    code = custom.get(fid, BUILTIN_FORMATS.get(fid, ""))
                    style_shape.append(_fmt_shape(code) if style_kind[-1] != "date" else None)
        sheets = wb.find(S + "sheets")
        uncalculated = []
        hidden_sheets = []
        shared_copies = [0]    # cells that repeat a shared formula without carrying its text
        all_formulas = []      # (sheet name, formula text)
        for sh in (sheets if sheets is not None else []):
            name = sh.get("name", "Sheet")
            rid = sh.get(R + "id")
            if rid not in rels:
                continue
            if sh.get("state") in ("hidden", "veryHidden"):
                hidden_sheets.append(name)
            root = _xml(z, rels[rid][1])
            if root is None or root.find(S + "sheetData") is None:
                continue
            prefix = "%s!" % name
            cells = {}          # (row, col) -> dict
            shared_forms = {}   # si -> (r1c1, master ref)
            hidden_rows = []
            hidden_cols = []
            cols_el = root.find(S + "cols")
            if cols_el is not None:
                for ce in cols_el:
                    if ce.get("hidden") in ("1", "true"):
                        a, b = int(ce.get("min", "1")), int(ce.get("max", "1"))
                        hidden_cols.append(num_to_col(a) if a == b else "%s:%s" % (num_to_col(a), num_to_col(min(b, 16384))))
            for row in root.find(S + "sheetData"):
                rnum = int(row.get("r", "0") or 0)
                if row.get("hidden") in ("1", "true"):
                    hidden_rows.append(rnum)
                for cnum_seq, c in enumerate(row, 1):
                    ref = c.get("r")
                    if ref:
                        m = re.match(r"([A-Z]+)(\d+)", ref)
                        cnum, rnum2 = col_to_num(m.group(1)), int(m.group(2))
                    else:
                        cnum, rnum2 = cnum_seq, rnum
                    ctype = c.get("t", "n")
                    f = c.find(S + "f")
                    v = c.find(S + "v")
                    raw = v.text if v is not None and v.text is not None else ""
                    info = {"formula": None, "r1c1": None, "kind": "empty", "text": "", "err": False, "shown": ""}
                    if f is not None:
                        ftxt = f.text or ""
                        if f.get("t") == "shared":
                            si = f.get("si")
                            if ftxt:
                                shared_forms[si] = (to_r1c1(ftxt, rnum2, cnum), "%s%d" % (num_to_col(cnum), rnum2))
                                info["formula"] = "=" + ftxt
                                info["r1c1"] = shared_forms[si][0]
                            elif si in shared_forms:
                                info["formula"] = "(same formula as %s, filled down or across)" % shared_forms[si][1]
                                info["r1c1"] = shared_forms[si][0]
                            else:
                                info["formula"] = "(shared formula)"
                        else:
                            info["formula"] = "=" + ftxt
                            info["r1c1"] = to_r1c1(ftxt, rnum2, cnum)
                        if ftxt:
                            all_formulas.append((name, ftxt))
                        else:
                            shared_copies[0] += 1
                    if ctype == "s":
                        try:
                            info["text"] = shared[int(raw)]
                        except (ValueError, IndexError):
                            info["text"] = ""
                        info["kind"] = "string"
                    elif ctype == "inlineStr":
                        info["text"] = "".join(t.text or "" for t in c.iter(S + "t"))
                        info["kind"] = "string"
                    elif ctype == "str":
                        info["text"] = raw
                        info["kind"] = "string"
                    elif ctype == "b":
                        info["text"] = "TRUE" if raw == "1" else "FALSE"
                        info["kind"] = "bool"
                    elif ctype == "e":
                        info["text"] = raw
                        info["kind"] = "error"
                        info["err"] = True
                    elif raw != "":
                        sidx = int(c.get("s", "0") or 0)
                        kind = style_kind[sidx] if sidx < len(style_kind) else ""
                        info["kind"] = "number"
                        if kind == "percent":
                            try:
                                info["text"] = fmt_plain(Decimal(format(float(raw) * 100, ".12g"))) + "%"
                            except (ValueError, InvalidOperation):
                                info["text"] = raw
                        elif kind == "date":
                            try:
                                base = datetime.datetime(1904, 1, 1) if date1904 else datetime.datetime(1899, 12, 30)
                                dt = base + datetime.timedelta(days=float(raw))
                                info["text"] = dt.date().isoformat() if dt.time() == datetime.time(0) else dt.isoformat(" ", "minutes")
                                info["kind"] = "date"
                            except (ValueError, OverflowError):
                                info["text"] = _tidy_float(raw)
                        else:
                            info["text"] = _tidy_float(raw)
                        shape = style_shape[sidx] if sidx < len(style_shape) else None
                        if shape is not None and info["kind"] == "number":
                            shown = shown_as(raw, shape, kind == "percent")
                            if shown and shown.replace(",", "") != info["text"]:
                                info["shown"] = shown
                    elif f is not None:
                        info["kind"] = "uncalculated"
                    if info["kind"] != "empty" or info["formula"]:
                        cells[(rnum2, cnum)] = info
            if not cells:
                continue
            if hidden_rows:
                names = []
                for hr in hidden_rows[:8]:
                    texts = [cells[k]["text"] for k in sorted(cells) if k[0] == hr and cells[k]["kind"] == "string"]
                    if texts:
                        names.append('row %d "%s"' % (hr, texts[0][:40]))
                notes.append('Sheet "%s" has hidden rows: %s%s. Hidden rows still count in every SUM range that '
                             "spans them, so the visible rows will not add up to the totals shown." % (
                                 name, _compress(hidden_rows), " (%s)" % ", ".join(names) if names else ""))
            if hidden_cols:
                notes.append('Sheet "%s" has hidden columns: %s.' % (name, ", ".join(hidden_cols)))

            def ref_of(r, c):
                return "%s%s%d" % (prefix, num_to_col(c), r)

            def is_num(info):
                return info is not None and info["kind"] == "number"

            # Error values and formulas never calculated
            for (r, c), info in sorted(cells.items()):
                if info["err"]:
                    leads.append(lead(ref_of(r, c), "error-value",
                                      "Cell shows %s%s." % (info["text"], " from " + info["formula"] if info["formula"] else "")))
                if info["kind"] == "uncalculated":
                    uncalculated.append(ref_of(r, c))

            # Numbers stored as text next to real numbers
            for (r, c), info in sorted(cells.items()):
                if info["kind"] != "string" or info["formula"]:
                    continue
                n = parse_cell(info["text"], conv_hint)
                if n is None or n.yearlike:
                    continue
                if is_num(cells.get((r - 1, c))) or is_num(cells.get((r + 1, c))):
                    leads.append(lead(ref_of(r, c), "number-as-text",
                                      'Cell holds "%s" as text, between numeric cells. SUM and similar functions skip text.'
                                      % info["text"]))

            # SUM ranges that stop short
            sum_rx = re.compile(r"(?<![A-Za-z0-9_.!])SUM\(\s*\$?([A-Z]{1,3})\$?(\d+)\s*:\s*\$?([A-Z]{1,3})\$?(\d+)\s*\)", re.I)
            by_row = {}
            explained = set()
            for (r, c), info in sorted(cells.items()):
                ftxt = info["formula"] or ""
                if not ftxt.startswith("="):
                    continue
                for m in sum_rx.finditer(ftxt):
                    c1, r1, c2, r2 = col_to_num(m.group(1).upper()), int(m.group(2)), col_to_num(m.group(3).upper()), int(m.group(4))
                    if c1 == c2 == c and r > r2:
                        gap = [x for x in range(r2 + 1, r) if is_num(cells.get((x, c)))]
                        if gap:
                            explained.add((r, c))
                            cl = num_to_col(c)
                            where_gap = ("%s%d (%s)" % (cl, gap[0], cells[(gap[0], c)]["text"]) if len(gap) == 1 else
                                         "%d cells between %s%d and %s%d" % (len(gap), cl, gap[0], cl, gap[-1]))
                            leads.append(lead(ref_of(r, c), "sum-range",
                                              "%s stops at row %d, but %s %s left out." % (
                                                  ftxt, r2, where_gap,
                                                  "holds a number and is" if len(gap) == 1 else "hold numbers and are")))
                        above = cells.get((r1 - 1, c))
                        if is_num(above) and not above["formula"]:
                            leads.append(lead(ref_of(r, c), "sum-range",
                                              "%s starts at row %d, but %s%d directly above holds the number %s and is left out." % (
                                                  ftxt, r1, num_to_col(c), r1 - 1, above["text"]), "weak"))
                        if re.fullmatch(r"=\s*SUM\([^()]*\)\s*", ftxt, re.I):
                            by_row.setdefault(r, []).append((c, r1, r2))
                    elif r1 == r2 == r and c > c2:
                        gap = [x for x in range(c2 + 1, c) if is_num(cells.get((r, x)))]
                        if gap:
                            explained.add((r, c))
                            where_gap = ("%s%d (%s)" % (num_to_col(gap[0]), r, cells[(r, gap[0])]["text"])
                                         if len(gap) == 1 else
                                         "%d cells between %s%d and %s%d" % (
                                             len(gap), num_to_col(gap[0]), r, num_to_col(gap[-1]), r))
                            leads.append(lead(ref_of(r, c), "sum-range",
                                              "%s stops at column %s, but %s %s left out." % (
                                                  ftxt, num_to_col(c2), where_gap,
                                                  "holds a number and is" if len(gap) == 1 else "hold numbers and are")))
            for r, items in by_row.items():
                if len(items) < 3:
                    continue
                spans = {}
                for c, r1, r2 in items:
                    spans.setdefault((r1, r2), []).append(c)
                common = max(spans.items(), key=lambda kv: len(kv[1]))
                if len(common[1]) * 2 > len(items):
                    for span, cs in spans.items():
                        if span == common[0]:
                            continue
                        for c in cs:
                            if (r, c) in explained:
                                continue
                            explained.add((r, c))
                            leads.append(lead(ref_of(r, c), "sum-range",
                                              "SUM covers rows %d to %d, while %d sibling totals in row %d cover rows %d to %d." % (
                                                  span[0], span[1], len(common[1]), r, common[0][0], common[0][1])))

            # Typed constant between identical formulas, and formulas that break a pattern
            for (r, c), info in sorted(cells.items()):
                for (dr, dc, axis) in ((1, 0, "column"), (0, 1, "row")):
                    prev = cells.get((r - dr, c - dc))
                    nxt = cells.get((r + dr, c + dc))
                    if not prev or not nxt or not prev["r1c1"] or prev["r1c1"] != nxt["r1c1"]:
                        continue
                    if info["formula"] is None and info["kind"] == "number":
                        leads.append(lead(ref_of(r, c), "typed-over-formula",
                                          "Typed number %s sits between two cells that share one formula pattern in its %s. "
                                          "By that pattern this cell would hold =%s. A typed number does not update "
                                          "with its inputs." % (info["text"], axis, from_r1c1(prev["r1c1"], r, c))))
                        break
                    if info["r1c1"] and info["r1c1"] != prev["r1c1"] and (r, c) not in explained:
                        leads.append(lead(ref_of(r, c), "formula-pattern",
                                          "Formula %s differs from its neighbours in the same %s. By their pattern "
                                          "this cell would hold =%s." % (
                                              info["formula"], axis, from_r1c1(prev["r1c1"], r, c)),
                                          "weak" if axis == "row" else "normal"))
                        break

            # A typed number in a column that is otherwise calculated
            flagged = set(ld["where"] for ld in leads)
            by_col = {}
            for (r, c), info in cells.items():
                if info["kind"] in ("number", "uncalculated"):
                    by_col.setdefault(c, []).append((r, info))
            for c, items in by_col.items():
                forms = [x for x in items if x[1]["formula"]]
                typed = [x for x in items if not x[1]["formula"]]
                if len(forms) >= 3 and typed and len(typed) * 5 <= len(items):
                    for r, info in typed:
                        if ref_of(r, c) in flagged:
                            continue
                        label = next((cells[k]["text"] for k in sorted(cells)
                                      if k[0] == r and cells[k]["kind"] == "string"), "")
                        leads.append(lead(ref_of(r, c), "typed-among-formulas",
                                          '%s is a typed number (%s) in a column where %d other cells are formulas. '
                                          "If it can be counted or calculated from the workbook, recalculate it." % (
                                              '"%s"' % label if label else "This cell", info["text"], len(forms)),
                                          "weak"))

            # A label that names the base of a ratio, against the cell the formula divides by
            for (r, c), info in sorted(cells.items()):
                ftxt = info["formula"] or ""
                m = re.fullmatch(r"=\s*\(?[^()]*\)?\s*/\s*\$?([A-Z]{1,3})\$?(\d+)\s*", ftxt)
                if not m:
                    continue
                label = next((cells[k]["text"] for k in sorted(cells) if k[0] == r and cells[k]["kind"] == "string"), "")
                lm = re.search(r"(?:%\s*(?:of|von|vom|des|der)|\bper|\bpro|\bje)\s+([^\W\d_]{3,})", label, re.I)
                if not lm:
                    continue
                stem = lm.group(1).lower()[:4]
                dr, dc = int(m.group(2)), col_to_num(m.group(1))
                dlabel = next((cells[k]["text"] for k in sorted(cells)
                               if k[0] == dr and cells[k]["kind"] == "string"), "")
                if not dlabel or stem in dlabel.lower():
                    continue
                other = [cells[k]["text"] for k in sorted(cells)
                         if k[1] < dc and k[0] != r and cells[k]["kind"] == "string" and stem in cells[k]["text"].lower()]
                if other:
                    leads.append(lead(ref_of(r, c), "ratio-base",
                                      'The label "%s" names "%s" as the base, but %s divides by %s%d, the row labelled '
                                      '"%s". A row labelled "%s" exists.' % (
                                          label, lm.group(1), ftxt, m.group(1), dr, dlabel, other[0])))

            # Value grid, split into blocks at fully empty rows
            rows_used = sorted(set(r for r, _ in cells))
            cols_used = sorted(set(c for _, c in cells))
            cmin, cmax = cols_used[0], cols_used[-1]
            block, prev_r = [], None
            groups = []
            for r in rows_used:
                if prev_r is not None and r != prev_r + 1 and block:
                    groups.append(block)
                    block = []
                block.append(r)
                prev_r = r
            if block:
                groups.append(block)
            for g in groups:
                rows, meta = [], {}
                for ri, r in enumerate(g):
                    row = []
                    for ci, c in enumerate(range(cmin, cmax + 1)):
                        info = cells.get((r, c))
                        row.append(info["text"] if info else "")
                        if info and info["kind"] in ("number", "uncalculated", "error", "date", "string", "bool"):
                            meta[(ri, ci)] = {"formula": info["formula"], "shown": info["shown"],
                                              "text": info["kind"] == "string" and not info["formula"]
                                              and parse_cell(info["text"], conv_hint) is not None}
                    rows.append(row)
                t = Table(prefix, rows, where='sheet "%s", rows %d to %d' % (name, g[0], g[-1]),
                          row_ids=[str(r) for r in g],
                          col_ids=[num_to_col(c) for c in range(cmin, cmax + 1)],
                          meta=meta, joiner="", stored=True)
                blocks.append(("table", t))
        n_sheets = len(list(sheets)) if sheets is not None else 0
        notes.append("Read: %d sheet(s)%s, %d formula(s)%s." % (
            n_sheets, ", none hidden" if not hidden_sheets else ", %d of them hidden" % len(hidden_sheets),
            len(all_formulas) + shared_copies[0],
            "" if uncalculated else ", all with stored results"))
        if not any("hidden rows" in x or "hidden columns" in x for x in notes):
            notes.append("No hidden rows or columns.")
        for hs in hidden_sheets:
            users = sorted(set(sn for sn, ftxt in all_formulas
                               if sn != hs and (hs + "!" in ftxt or "'" + hs + "'!" in ftxt)))
            if users:
                notes.append('Sheet "%s" is hidden and formulas on %s read from it. Its figures feed the visible '
                             "sheets and were checked like the others." % (hs, ", ".join('"%s"' % u for u in users)))
            else:
                notes.append('Sheet "%s" is hidden. No formula on another sheet refers to it.' % hs)
        if uncalculated:
            leads.insert(0, lead(
                "whole file", "uncalculated",
                "%d formula cell(s) have no stored result (%s%s). The file was saved by a program that does not "
                "calculate, so formula values are unknown here and totals built on them cannot be checked. "
                "Open it in a spreadsheet program, save it and run the check again." % (
                    len(uncalculated), ", ".join(uncalculated[:5]), ", ..." if len(uncalculated) > 5 else "")))
    return {"blocks": blocks, "notes": notes, "leads": leads}


def _compress(nums):
    nums = sorted(nums)
    out, start, prev = [], nums[0], nums[0]
    for n in nums[1:]:
        if n == prev + 1:
            prev = n
            continue
        out.append(str(start) if start == prev else "%d-%d" % (start, prev))
        start = prev = n
    out.append(str(start) if start == prev else "%d-%d" % (start, prev))
    return ", ".join(out)


# ---- plain text, Markdown, CSV, PDF ---------------------------------------

def _read_text(path):
    with open(path, "rb") as fh:
        data = fh.read()
    for enc in ("utf-8-sig", "utf-16", "cp1252", "latin-1"):
        try:
            return data.decode(enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    return data.decode("utf-8", "replace")


def _blocks_from_lines(lines, prefix=""):
    """Turn text lines into text blocks and Markdown pipe tables."""
    blocks = []
    i, tcount = 0, 0
    n = len(lines)
    while i < n:
        line = lines[i]
        if line.count("|") >= 2 and i + 1 < n and re.match(r"^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)+\|?\s*$", lines[i + 1]):
            j = i
            rows = []
            while j < n and lines[j].count("|") >= 1 and lines[j].strip():
                if j != i + 1:
                    cells = lines[j].strip()
                    cells = cells[1:] if cells.startswith("|") else cells
                    cells = cells[:-1] if cells.endswith("|") else cells
                    rows.append([re.sub(r"[*_`]", "", c).strip() for c in cells.split("|")])
                j += 1
            tcount += 1
            blocks.append(("table", Table("%sT%d" % (prefix, tcount), rows,
                                          where="%slines %d to %d" % (prefix, i + 1, j))))
            i = j
            continue
        if line.strip():
            blocks.append(("text", "%sL%d" % (prefix, i + 1), line.strip()))
        i += 1
    return blocks


def extract_text(path):
    return {"blocks": _blocks_from_lines(_read_text(path).splitlines()), "notes": [], "leads": []}


def extract_csv(path):
    text = _read_text(path)
    sample = text[:8192]
    delim = None
    if path.lower().endswith(".tsv"):
        delim = "\t"
    else:
        try:
            delim = csv.Sniffer().sniff(sample, delimiters=";,\t|").delimiter
        except csv.Error:
            first = sample.splitlines()[0] if sample.splitlines() else ""
            delim = max(";,\t|", key=first.count)
    rows = [row for row in csv.reader(io.StringIO(text), delimiter=delim)]
    rows = [[c.strip() for c in row] for row in rows if any(c.strip() for c in row)]
    t = Table("", rows, where="the whole file",
              row_ids=[str(i + 1) for i in range(len(rows))],
              col_ids=[num_to_col(i + 1) for i in range(max((len(r) for r in rows), default=0))],
              joiner="")
    return {"blocks": [("table", t)], "notes": ["Delimiter read as %r." % delim], "leads": []}


NUMISH_RE = re.compile(r"^(?:[" + re.escape(CURRENCY_CHARS) + r"%]|[-\u2013\u2014]|n/?a|" + CURRENCY_CODES + r"|\(|\))$")


def infer_tables(words, prefix, where):
    """Find column-aligned number blocks in positioned words (PDF pages without ruled tables).

    words: dicts with text, x0, x1, top. Returns a list of Table objects.
    """
    lines = []
    for w in sorted(words, key=lambda w: (round(w["top"]), w["x0"])):
        if lines and abs(lines[-1][0]["top"] - w["top"]) <= 3:
            lines[-1].append(w)
        else:
            lines.append([w])
    rows = []      # (label words, tail words) or None for a line that is not a table row
    for line in lines:
        line.sort(key=lambda w: w["x0"])
        cut = len(line)
        while cut > 0:
            txt = line[cut - 1]["text"]
            if parse_cell(txt) is not None or NUMISH_RE.match(txt):
                cut -= 1
            else:
                break
        tail = line[cut:]
        n_nums = sum(1 for w in tail if parse_cell(w["text"]) is not None)
        rows.append((line[:cut], tail) if n_nums >= 1 else None)
    tables = []
    i = 0
    while i < len(rows):
        if rows[i] is None:
            i += 1
            continue
        j = i
        while j < len(rows) and rows[j] is not None:
            j += 1
        block = list(range(i, j))
        if len(block) >= 3 and sum(1 for k in block if len(rows[k][1]) >= 2) >= 2:
            spans = sorted((w["x0"], w["x1"]) for k in block for w in rows[k][1]
                           if parse_cell(w["text"]) is not None)
            bands = []
            for x0, x1 in spans:
                if bands and x0 <= bands[-1][1] + 1:
                    bands[-1][1] = max(bands[-1][1], x1)
                else:
                    bands.append([x0, x1])
            counts = [sum(1 for k in block if any(w["x0"] <= b[1] + 1 and w["x1"] >= b[0] - 1 for w in rows[k][1]))
                      for b in bands]
            while len(bands) > 1 and counts[0] * 2 < len(block) and counts[1] > counts[0]:
                bands.pop(0)      # a number that belongs to the row label
                counts.pop(0)

            def band_of(w):
                best, dist = None, None
                for bi, b in enumerate(bands):
                    d = 0 if (w["x0"] <= b[1] + 1 and w["x1"] >= b[0] - 1) else min(abs(w["x0"] - b[1]), abs(b[0] - w["x1"]))
                    if dist is None or d < dist:
                        best, dist = bi, d
                return best, dist
            grid = []
            head = None
            if i > 0 and rows[i - 1] is None and lines[i - 1]:
                cells = [[] for _ in range(len(bands) + 1)]
                for w in lines[i - 1]:
                    bi, d = band_of(w)
                    if w["x1"] < bands[0][0] - 2 and d > 0:
                        cells[0].append(w["text"])
                    else:
                        cells[bi + 1].append(w["text"])
                if sum(1 for c in cells[1:] if c) * 2 >= len(bands):
                    head = [" ".join(c) for c in cells]
            if head:
                grid.append(head)
            for k in block:
                label, tail = rows[k]
                cells = [[] for _ in range(len(bands) + 1)]
                cells[0] = [w["text"] for w in label]
                for w in tail:
                    bi, d = band_of(w)
                    if w["x1"] < bands[0][0] - 2 and d > 0:
                        cells[0].append(w["text"])
                    else:
                        cells[bi + 1].append(w["text"])
                grid.append([" ".join(c) for c in cells])
            tables.append(Table(
                "%s%d" % (prefix, len(tables) + 1), grid,
                where="%s, text lines %d to %d, columns inferred from the position of the numbers"
                      % (where, (i if head else i + 1), j)))
        i = j
    return tables


def extract_pdf(path):
    blocks, notes = [], []
    try:
        import pdfplumber  # optional
    except Exception:
        pdfplumber = None
    if pdfplumber is not None:
        inferred = False
        with pdfplumber.open(path) as pdf:
            for pno, page in enumerate(pdf.pages, 1):
                text = page.extract_text() or ""
                if not text.strip():
                    notes.append("Page %d has no text layer (scanned or image only)." % pno)
                for b in _blocks_from_lines(text.splitlines(), "p%d " % pno):
                    blocks.append(b)
                ruled = []
                try:
                    for tb in page.find_tables():
                        rows = tb.extract()
                        clean = [[(c or "").replace("\n", " ").strip() for c in row] for row in rows if row]
                        if len(clean) >= 2:
                            ruled.append((tb.bbox, clean))
                except Exception:
                    ruled = []
                for tno, (bbox, clean) in enumerate(ruled, 1):
                    blocks.append(("table", Table("p%d T%d" % (pno, tno), clean, where="page %d" % pno)))
                try:
                    words = page.extract_words()
                except Exception:
                    words = []
                loose = [w for w in words if not any(
                    b[0] - 1 <= w["x0"] and w["x1"] <= b[2] + 1 and b[1] - 1 <= w["top"] <= b[3] + 1
                    for b, _ in ruled)]
                for tb in infer_tables(loose, "p%d %s" % (pno, "I"), "page %d" % pno):
                    blocks.append(("table", tb))
                    inferred = True
        if inferred:
            notes.append("Tables named I1, I2, ... have no ruling lines in the PDF. Their columns were inferred from "
                         "where the numbers sit, so check a row against the page before reporting a lead from them.")
        return {"blocks": blocks, "notes": notes, "leads": []}
    try:
        import pypdf  # optional
    except Exception:
        pypdf = None
    if pypdf is not None:
        reader = pypdf.PdfReader(path)
        for pno, page in enumerate(reader.pages, 1):
            text = page.extract_text() or ""
            if not text.strip():
                notes.append("Page %d has no text layer (scanned or image only)." % pno)
            blocks.extend(_blocks_from_lines(text.splitlines(), "p%d " % pno))
        notes.append("Tables were read as plain lines (pypdf). Check table arithmetic from the text.")
        return {"blocks": blocks, "notes": notes, "leads": []}
    notes.append("No PDF text library is installed (pdfplumber or pypdf). Read this PDF directly and "
                 "apply the checks by hand, or convert it to text first.")
    return {"blocks": blocks, "notes": notes, "leads": []}


EXTRACTORS = {
    ".docx": extract_docx, ".docm": extract_docx,
    ".pptx": extract_pptx, ".pptm": extract_pptx,
    ".xlsx": extract_xlsx, ".xlsm": extract_xlsx,
    ".csv": extract_csv, ".tsv": extract_csv,
    ".md": extract_text, ".markdown": extract_text, ".txt": extract_text, ".text": extract_text,
    ".pdf": extract_pdf,
}


# --------------------------------------------------------------------------
# Running an extraction
# --------------------------------------------------------------------------

AMBIGUOUS_RE = re.compile(r"(?<![\d.,])[1-9]\d{0,2}[.,]\d{3}(?![\d.,])")


def held_summary():
    """One line on what the rules looked at and how much of it agreed."""
    parts = []
    for key, name in (("totals", "totals"), ("calcs", "row calculations"),
                      ("nofn", "counts with a percentage"), ("weekdays", "weekdays"),
                      ("spans", "durations"), ("points", "percentage-point statements")):
        if STATS[key]:
            parts.append("%s: %d of %d agree" % (name, STATS[key + "_ok"], STATS[key]))
    if not parts:
        return "The rules found nothing they could test in this file. Everything has to be checked by reading."
    return ("Tested by rule: " + "; ".join(parts) + ". Anything else was not tested by rule, for example "
            "figures in sentences, growth rates and text against tables.")


def process(path):
    ext = os.path.splitext(path)[1].lower()
    result = {"file": os.path.basename(path), "type": ext.lstrip("."), "blocks": [],
              "notes": [], "leads": [], "error": None}
    if ext not in EXTRACTORS:
        result["error"] = ("Unsupported file type %s. Supported: %s. Old .doc, .xls and .ppt files "
                           "must be saved in the current format first." % (ext or "(none)", " ".join(sorted(EXTRACTORS))))
        return result
    try:
        data = EXTRACTORS[ext](path)
    except zipfile.BadZipFile:
        result["error"] = "The file is not a valid Office file (it may be password protected or damaged)."
        return result
    except Exception as exc:   # report and carry on with the other files
        result["error"] = "Could not read the file: %s: %s" % (type(exc).__name__, exc)
        return result
    STATS.clear()
    lines = [(b[1], b[2]) for b in data["blocks"] if b[0] == "text"]
    tables = [b[1] for b in data["blocks"] if b[0] == "table"]
    eu, us = number_format_evidence(lines, tables)
    conv = "eu" if len(eu) > len(us) else "us"
    FMT["conv"] = conv
    FMT["group"] = any(has_grouping(cell) and parse_cell(cell, conv) is not None
                       for t in tables for row in t.rows for cell in row)
    leads = list(data["leads"])
    untested = []
    for t in tables:
        before = sum(STATS.values())
        leads.extend(check_table(t, conv))
        if sum(STATS.values()) == before and len(t.rows) >= 2:
            untested.append((t.name.rstrip("!") if t.joiner == "" else t.name) or "the table")
    leads.extend(check_prose(lines))
    for t in tables:
        cell_lines = []
        for r, row in enumerate(t.rows):
            for c, cell in enumerate(row):
                if cell and len(cell) > 8 and parse_cell(cell, conv) is None:
                    cell_lines.append((t.ref(r, c), cell))
        leads.extend(check_prose(cell_lines))
    notes = list(data["notes"])
    stored_only = bool(tables) and all(t.stored for t in tables) and not lines
    if eu and us and not stored_only:
        major, minor = (eu, us) if len(eu) >= len(us) else (us, eu)
        names = ("a decimal comma", "a decimal point")
        if major is us:
            names = (names[1], names[0])
        leads.append(lead(
            "whole file", "number-format",
            "Two number formats in one file: %d figure(s) are written with %s and %d with %s, for example %s. "
            "Confirm each is intended." % (
                len(major), names[0], len(minor), names[1],
                "; ".join('%s "%s"' % (w, x) for w, x in minor[:4])),
            "weak"))
    if not stored_only:
        amb = [(w, m.group(0)) for w, text in lines for m in AMBIGUOUS_RE.finditer(text or "")]
        for t in tables:
            if t.stored:
                continue
            for r, row in enumerate(t.rows):
                for c, cell in enumerate(row):
                    n = parse_cell(cell, conv)
                    if n is not None and n.style == "amb":
                        amb.append((t.ref(r, c), n.raw))
        mark = "decimal comma" if conv == "eu" else "decimal point"
        if amb and eu and us:
            notes.append("%d figure(s) such as %s can be read two ways, and the file mixes both number formats. "
                         "They were read with a %s, the more frequent one. Check each against its context." % (
                             len(amb), "; ".join('%s "%s"' % a for a in amb[:3]), mark))
        elif amb and not (eu or us):
            notes.append("%d figure(s) such as %s can be read two ways and nothing in the file settles it. "
                         "They were read with a %s (so 1,234 is one thousand two hundred and thirty-four). "
                         "Say so in the report if a finding depends on it." % (
                             len(amb), "; ".join('%s "%s"' % a for a in amb[:3]), mark))
    # Leads in reading order, so that leads on the same cell or table sit together.
    order = {}
    for bi, b in enumerate(data["blocks"]):
        if b[0] == "text":
            order.setdefault(b[1], (bi, 0, 0))
            continue
        t = b[1]
        for r in range(len(t.rows)):
            for c in range(t.width):
                order.setdefault(t.ref(r, c), (bi, r, c))
        for c in range(t.width):
            key = "%s %s" % (t.name, t.col_ids[c]) if t.joiner else "%scolumn %s" % (t.name, t.col_ids[c])
            order.setdefault(key, (bi, len(t.rows), c))
    last = (len(data["blocks"]) + 1, 0, 0)
    first = (-1, 0, 0)
    leads.sort(key=lambda ld: first if ld["kind"] == "uncalculated" else order.get(ld["where"], last))
    held = held_summary()
    untested = [u for u in untested if not any(ld["where"].startswith(u) for ld in leads)]
    if untested:
        names = sorted(set(untested), key=untested.index)
        held += " No rule applied to %s%s: check %s by reading." % (
            ", ".join(names[:10]), " and %d more" % (len(names) - 10) if len(names) > 10 else "",
            "it" if len(names) == 1 else "them")
    lead_rows = {}
    for ld in leads:
        pos = order.get(ld["where"])
        if pos:
            lead_rows.setdefault(pos[0], set()).add(pos[1])
    result.update(blocks=data["blocks"], notes=notes, leads=leads, convention=conv, held=held, lead_rows=lead_rows)
    return result


def render(results, max_rows, all_rows=False):
    out = []
    out.append("# Adds Up extraction")
    out.append("")
    out.append("Leads are candidates produced by rule. Confirm each one against the content before reporting it,")
    out.append("and read the content yourself for everything the rules cannot see.")
    for idx, res in enumerate(results, 1):
        out.append("")
        out.append("## File %d: %s" % (idx, res["file"]))
        if res["error"]:
            out.append("")
            out.append("ERROR: " + res["error"])
            continue
        out.append("")
        out.append("### Content")
        for bi, b in enumerate(res["blocks"]):
            if b[0] == "text":
                out.append("[%s] %s" % (b[1], b[2]))
                continue
            t = b[1]
            label = t.name.rstrip("!") if t.joiner == "" and t.name else (t.name or "table")
            out.append("")
            out.append("[%s] table, %d rows x %d columns, %s" % (label, len(t.rows), t.width, t.where))
            n = len(t.rows)
            if all_rows or n <= max_rows:
                keep = list(range(n))
            else:
                # Long table: the top, the bottom and every row a lead points to, with its neighbours.
                wanted = set(range(min(n, 12))) | set(range(max(0, n - 6), n))
                for r in res.get("lead_rows", {}).get(bi, ()):
                    wanted.update(x for x in (r - 1, r, r + 1) if 0 <= x < n)
                keep = sorted(wanted)
            if t.joiner != "":
                out.append("  %-5s %s" % ("", " | ".join(t.col_ids)))
            prev = -1
            for r in keep:
                if r != prev + 1:
                    out.append("  ... %d row(s) not printed (%s to %s). They were checked by the rules." % (
                        r - prev - 1, t.row_ids[prev + 1], t.row_ids[r - 1]))
                prev = r
                row = t.rows[r]
                if t.joiner == "":
                    parts = []
                    for c, cell in enumerate(row):
                        m = t.meta.get((r, c), {})
                        if not cell and not m.get("formula"):
                            continue
                        s = "%s%s: %s" % (t.col_ids[c], t.row_ids[r], cell)
                        if m.get("shown"):
                            s += " (shown as %s)" % m["shown"]
                        if m.get("formula"):
                            s += "  {%s}" % m["formula"]
                        parts.append(s)
                    out.append("  " + " | ".join(parts))
                else:
                    out.append("  %-5s %s" % (t.row_ids[r], " | ".join(row)))
            if len(keep) < n:
                out.append("  (%d of %d rows printed. Run again with --all-rows to print every row.)" % (len(keep), n))
            out.append("")
        out.append("")
        out.append("### Leads (%d), in reading order" % len(res["leads"]))
        if not res["leads"]:
            out.append("None from the automatic rules.")
        for n, ld in enumerate(res["leads"], 1):
            out.append("%d. [%s] (%s%s) %s" % (
                n, ld["where"], ld["kind"], ", weak" if ld["strength"] == "weak" else "", ld["message"]))
        out.append("")
        out.append("### What the rules covered")
        out.append(res.get("held", ""))
        if res["notes"]:
            out.append("")
            out.append("### Notes")
            for note in res["notes"]:
                out.append("- " + note)
    return "\n".join(out) + "\n"


def to_json(results):
    out = []
    for res in results:
        item = {k: res.get(k) for k in ("file", "type", "error", "notes", "leads", "convention", "held")}
        item["content"] = []
        for b in res["blocks"]:
            if b[0] == "text":
                item["content"].append({"kind": "text", "where": b[1], "text": b[2]})
            else:
                t = b[1]
                item["content"].append({
                    "kind": "table", "name": t.name, "where": t.where, "rows": t.rows,
                    "row_ids": t.row_ids, "col_ids": t.col_ids,
                    "formulas": {t.ref(r, c): m["formula"] for (r, c), m in sorted(t.meta.items()) if m.get("formula")},
                })
        out.append(item)
    return json.dumps(out, ensure_ascii=False, indent=1)


# --------------------------------------------------------------------------
# GRIM: can this percentage or mean come from whole numbers?
# --------------------------------------------------------------------------

def _round_half_up(d, decimals):
    return d.quantize(Decimal(1).scaleb(-decimals), rounding="ROUND_HALF_UP")


def _round_half_even(d, decimals):
    return d.quantize(Decimal(1).scaleb(-decimals), rounding="ROUND_HALF_EVEN")


def grim(n, value_text, percent):
    """Test whether a percentage or mean can come from whole numbers.

    Returns (counts that display as the stated value, nearest achievable values, decimals).
    """
    txt = value_text.replace(",", ".").rstrip("%").strip()
    stated = Decimal(txt)
    decimals = len(txt.split(".")[1]) if "." in txt else 0
    scale = Decimal(100) if percent else Decimal(1)
    half = Decimal(1).scaleb(-decimals) / 2
    low = int(((stated - half) * n / scale).to_integral_value(rounding="ROUND_FLOOR")) - 1
    high = int(((stated + half) * n / scale).to_integral_value(rounding="ROUND_CEILING")) + 1
    matches, near = [], []
    for k in range(max(0, low), high + 1):
        exact = Decimal(k) * scale / Decimal(n)
        if stated in (_round_half_up(exact, decimals), _round_half_even(exact, decimals)):
            matches.append(k)
        near.append((k, _round_half_up(exact, decimals)))
    near.sort(key=lambda kv: (abs(kv[1] - stated), kv[0]))
    return matches, near[:3], decimals


def cmd_grim(args):
    if (args.percent is None) == (args.mean is None):
        print("Give exactly one of --percent or --mean.")
        return 2
    is_pct = args.percent is not None
    value = args.percent if is_pct else args.mean
    matches, near, decimals = grim(args.n, value, is_pct)
    span = ""
    if matches:
        span = str(matches[0]) if len(matches) == 1 else "%d to %d" % (matches[0], matches[-1])
    if is_pct:
        what = "%s%% of n=%d" % (value.rstrip("%"), args.n)
        if matches:
            print("CONSISTENT: %s corresponds to a whole count (%s of %d)." % (what, span, args.n))
            if len(matches) > 1:
                print("Several counts display as this percentage, so the test cannot tell them apart. "
                      "If the count is stated, divide it by n instead.")
        else:
            print("NOT POSSIBLE: %s does not correspond to a whole count. Nearest: %s." % (
                what, "; ".join("%d/%d = %s%%" % (k, args.n, v) for k, v in near)))
            print("Check first that n is the right denominator and that the figure is a simple proportion.")
    else:
        what = "mean %s with n=%d" % (value, args.n)
        if matches:
            print("CONSISTENT: %s can come from whole-number values (total %s)." % (what, span))
        else:
            print("NOT POSSIBLE for whole-number data: %s. Nearest: %s." % (
                what, "; ".join("total %d gives %s" % (k, v) for k, v in near)))
            print("This only applies if every single value is a whole number (counts, scale points). "
                  "Check also that n is right.")
    return 0


# --------------------------------------------------------------------------
# Command line
# --------------------------------------------------------------------------

def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
    ap = argparse.ArgumentParser(description="Adds Up: extract documents and list arithmetic leads.")
    sub = ap.add_subparsers(dest="cmd")
    ex = sub.add_parser("extract", help="extract text, tables and formulas and list leads")
    ex.add_argument("files", nargs="+", help="files to read: docx, pptx, xlsx, xlsm, csv, tsv, md, txt, pdf")
    ex.add_argument("--json", action="store_true", help="print JSON instead of text")
    ex.add_argument("--max-rows", type=int, default=60,
                    help="tables longer than this are printed in part: top, bottom and rows with leads (default 60)")
    ex.add_argument("--all-rows", action="store_true", help="print every row of every table")
    gr = sub.add_parser("grim", help="test whether a percentage or mean can come from whole numbers")
    gr.add_argument("--n", type=int, required=True, help="sample size")
    gr.add_argument("--percent", help="reported percentage of n, written as in the document, e.g. 37.5")
    gr.add_argument("--mean", help="reported mean of whole-number values, written as in the document, e.g. 3.48")
    args = ap.parse_args(argv)
    if args.cmd == "grim":
        if args.n <= 0:
            print("n must be a positive whole number.")
            return 2
        return cmd_grim(args)
    if args.cmd != "extract":
        ap.print_help()
        return 2
    results = []
    for path in args.files:
        if not os.path.isfile(path):
            results.append({"file": path, "type": "", "blocks": [], "notes": [], "leads": [],
                            "error": "File not found."})
            continue
        results.append(process(path))
    sys.stdout.write(to_json(results) if args.json else render(results, args.max_rows, args.all_rows))
    return 0


if __name__ == "__main__":
    sys.exit(main())
