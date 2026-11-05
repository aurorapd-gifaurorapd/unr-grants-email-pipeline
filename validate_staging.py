#!/usr/bin/env python3
"""
validate_staging.py
====================
Stage 4.5 (pre-build) - validate a validated_*.xlsx staging sheet before
build_html.py runs on it. Catches data problems that would either break
build_html.py's assumptions or produce a broken/incorrect email.

Input: a validated_*.xlsx staging file.
Output: a table of errors/warnings printed to the terminal.

Exit 0 if clean (no errors). Exit 1 if any FAIL-level errors were found.
WARN-level issues are printed but do not affect the exit code.

Usage:
    python3 validate_staging.py validated_Grant_Opportunities_Staging.xlsx
"""
import sys, re, glob, argparse, datetime
from pathlib import Path
try:
    import openpyxl
except ImportError:
    print("ERROR: pip install openpyxl"); sys.exit(1)

from build_html import SECTION_ORDER, TYPE_LABELS, ENCODING_MAP, norm as html_norm, looks_like_abbr, find_duplicates
from grants_common import normalize_deadline, canonical_type

# Windows terminals default to a codepage that can't encode arbitrary
# characters we print (e.g. the offending char in the U+00FF check below).
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="backslashreplace")

# Same 16-column layout merge_and_assess.py writes and build_html.py reads.
COL = {
    1:"email",2:"eligflag",3:"type",4:"rec",5:"loi",6:"agency",7:"abbr",
    8:"title",9:"summary",10:"deadline",11:"l1url",12:"l1text",13:"l2url",
    14:"l2text",15:"elignotes",16:"recordid"}

HTML_ENTITY_RE = re.compile(r"&(?:[a-zA-Z]+|#\d+|#x[0-9a-fA-F]+);")
DATE_FORMATS = ("%B %d, %Y", "%b %d, %Y", "%Y-%m-%d", "%m/%d/%Y")

def resolve_input(raw_path):
    p = Path(raw_path)
    if p.exists():
        return p
    matches = glob.glob(raw_path)
    if len(matches) == 1:
        return Path(matches[0])
    if len(matches) > 1:
        raise SystemExit(f"Multiple files match '{raw_path}': {matches}. Specify one.")
    raise SystemExit(f"Input file not found: {raw_path}")

def read_rows(path):
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    rows = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or not any(row):
            continue
        r = {COL[j + 1]: (row[j] if j < len(row) else "") for j in range(16)}
        r["_rownum"] = i
        rows.append(r)
    return wb, ws, rows

def clean_summary_newlines(ws):
    """SILENT: strip newlines out of Summary cells in place.
    Returns True if any cell was changed (caller saves the workbook)."""
    modified = False
    for row in ws.iter_rows(min_row=2):
        cell = row[8]  # column I = Summary
        if isinstance(cell.value, str) and ("\n" in cell.value or "\r" in cell.value):
            cell.value = re.sub(r"[\r\n]+", " ", cell.value).strip()
            modified = True
    return modified

def norm_title(t):
    t = html_norm(str(t or ""))
    return re.sub(r"\s+", " ", t).strip().lower()

def parse_loose_date(s):
    s = (s or "").strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None

def check_row(r):
    fails, warns = [], []
    rn = r["_rownum"]
    email = str(r.get("email") or "").strip().lower()
    abbr = str(r.get("abbr") or "").strip()
    ttype = str(r.get("type") or "").strip()
    title = str(r.get("title") or "").strip()
    summary = str(r.get("summary") or "")
    deadline = r.get("deadline")
    l1url = str(r.get("l1url") or "").strip()
    l2url = str(r.get("l2url") or "").strip()

    if abbr and abbr not in SECTION_ORDER and abbr != "NEW" and not looks_like_abbr(abbr):
        warns.append((rn, "Abbreviation", f"'{abbr}' will go in the foundations (NEW) section"))
    elif abbr and abbr not in SECTION_ORDER and abbr != "NEW":
        warns.append((rn, "Abbreviation", f"'{abbr}' will get its own section -- if it belongs under "
                      f"an agency (e.g. DARPA under DOD), change it"))

    if ttype and canonical_type(ttype) not in TYPE_LABELS and canonical_type(ttype) != "Grant":
        fails.append((rn, "Type", f"'{ttype}' is not a type -- use Grant, DCL, RFI, Event, "
                      f"Limited Submission, or Policy"))

    for label, url in (("Link1 URL", l1url), ("Link2 URL", l2url)):
        if not url:
            continue
        if url.lower().startswith("http://"):
            fails.append((rn, label, f"insecure link '{url}' (use https://)"))

    # An Excel date cell is fine: build_html renders it as "Month DD, YYYY".
    deadline_is_datetime = isinstance(deadline, (datetime.date, datetime.datetime))
    deadline_str = "" if deadline_is_datetime else str(deadline or "").strip()
    if not deadline_is_datetime and deadline_str:
        parsed = parse_loose_date(deadline_str)
        if parsed:
            months_out = (parsed - datetime.date.today()).days / 30.44
            if months_out > 12:
                warns.append((rn, "Deadline", f"'{deadline_str}' is more than 12 months out"))
        # Same shape rules build_html.py hard-stops on, reported here first so
        # they can be fixed in the workbook instead of at build time. Only
        # Email?=Y rows can reach the email, so only those are failures.
        if deadline_str:
            _, dl_problems = normalize_deadline(deadline, ttype)
            for p in dl_problems:
                (fails if email == "y" else warns).append((rn, "Deadline", p))

    if email == "y":
        missing = []
        if not title: missing.append("Title")
        if not l1url: missing.append("Link1 URL")
        if not deadline_is_datetime and not deadline_str: missing.append("Deadline")
        if missing:
            fails.append((rn, "Email?=Y", f"missing required field(s): {', '.join(missing)}"))

    if not summary.strip():
        warns.append((rn, "Summary", "empty summary"))
    else:
        if HTML_ENTITY_RE.search(summary):
            warns.append((rn, "Summary", "HTML entity found in summary"))
        for ch in summary:
            if ord(ch) > 0xFF and ch not in ENCODING_MAP:
                fails.append((rn, "Summary", f"character U+{ord(ch):04X} ('{ch}') above U+00FF not in ENCODING_MAP"))

    return fails, warns

def find_duplicate_titles(rows):
    by_title = {}
    for r in rows:
        title = str(r.get("title") or "").strip()
        if not title:
            continue
        by_title.setdefault(norm_title(title), []).append(r["_rownum"])
    fails = []
    for rownums in by_title.values():
        if len(rownums) < 2:
            continue
        for rn in rownums:
            others = ", ".join(str(x) for x in rownums if x != rn)
            fails.append((rn, "Title", f"duplicate normalized title (also on row(s) {others})"))
    return fails

def print_table(rows, headers):
    widths = [len(h) for h in headers]
    str_rows = [[str(c) for c in row] for row in rows]
    for row in str_rows:
        for i, c in enumerate(row):
            widths[i] = max(widths[i], len(c))
    def fmt_row(cols):
        return "  ".join(c.ljust(widths[i]) for i, c in enumerate(cols))
    print(fmt_row(headers))
    print("  ".join("-" * w for w in widths))
    for row in str_rows:
        print(fmt_row(row))

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("infile", help="validated_*.xlsx staging file to validate")
    args = ap.parse_args()
    path = resolve_input(args.infile)

    wb, ws, rows = read_rows(path)
    if clean_summary_newlines(ws):
        wb.save(path)

    if not rows:
        print("No data rows found.")
        sys.exit(0)

    all_fails, all_warns = [], []
    for r in rows:
        fails, warns = check_row(r)
        all_fails.extend(fails)
        all_warns.extend(warns)
    all_fails.extend(find_duplicate_titles(rows))
    # Same check build_html.py stops on, reported here first with row numbers.
    for rs in find_duplicates([r for r in rows
                               if str(r.get("email") or "").strip().lower() == "y"]).values():
        for r in rs:
            others = ", ".join(str(o["_rownum"]) for o in rs if o is not r)
            all_fails.append((r["_rownum"], "Record ID",
                              f"same opportunity as row(s) {others} -- delete one"))

    all_fails.sort(key=lambda x: x[0])
    all_warns.sort(key=lambda x: x[0])

    if all_fails:
        print("ERRORS:")
        print_table(all_fails, ["Row", "Field", "Message"])
    if all_warns:
        if all_fails:
            print()
        print("WARNINGS:")
        print_table(all_warns, ["Row", "Field", "Message"])

    print()
    print(f"{len(all_fails)} error(s), {len(all_warns)} warning(s)")
    sys.exit(1 if all_fails else 0)

if __name__ == "__main__":
    main()
