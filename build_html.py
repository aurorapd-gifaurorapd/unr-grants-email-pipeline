#!/usr/bin/env python3
"""
build_html.py
=============
Stage 5: read a VALIDATED staging sheet and emit the weekly grant email HTML,
exactly to the Grants Email SKILL v0.9 spec (colors, fonts, anchors, CRLF,
latin-1 encoding, DCL/RFI/EVENT labels).

Input: an .xlsx whose filename starts with 'validated_' (the human has set
Email?=Y on the rows to include and finalized Summary/Deadline/Type).

Only rows with Email? == 'Y' (case-insensitive) are included. Records are grouped
into agency sections in a fixed federal-first order, then 'New Foundation R&D'
(abbrev NEW) last. Within a section, original sheet order is preserved.

Usage:
    python3 build_html.py validated_Grant_Opportunities_Staging.xlsx \
            --email-date 2026-06-15 --out Grant_Opportunities_06-15-2026.html

If --email-date is omitted, defaults to the Monday of the current week (the
week the email is being built/sent), not the following Monday.
"""
import sys, re, argparse, datetime, difflib
from pathlib import Path
try:
    import openpyxl
except ImportError:
    print("ERROR: pip install openpyxl"); sys.exit(1)

from grants_common import (  # single source of truth -- do not re-copy here
    ENCODING_MAP, TYPE_LABELS, TYPE_DEADLINE_LABELS, canonical_type,
    normalize_deadline, public_sam_url)

# ── SKILL constants ──────────────────────────────────────────────────────────
NAVY = "#041E42"; LINK = "#4F5254"; TITLECOL = "#000847"; REDLABEL = "#CC0000"
# Section order. The first six are a deliberate editorial choice, not alphabetical
# and not federal-first: DOD, DOE, then UNR's own internal section, then NIH, NSF,
# USDA. Everything else federal follows in the order below, and NEW is always last.
SECTION_ORDER = ["DOD","DOE","UNR","NIH","NSF","USDA",
                 "HHS","NASA","DOI","EPA","ED",
                 "NEH","NEA","NOAA","USGS","DHS","DOT","SBA","DOJ","DOS","DOC",
                 "DOL","HUD","CDC","FDA"]

# Internal UNR opportunities get their own section rather than being swept into
# "New Foundation R&D". merge_and_assess.lookup_agency() now tags these UNR at
# merge time, but a workbook exported before that change still carries "NEW", so
# recognize them by funder name here too and let either path work.
UNR_ABBR = "UNR"
UNR_RE = re.compile(r"university of nevada,?\s+reno|\bUNR\b", re.I)


def looks_like_abbr(a):
    """Short and all caps, like NSF or DARPA -- not a spelled-out funder name."""
    return bool(re.fullmatch(r"[A-Z][A-Z0-9&-]{1,9}", a or ""))


def section_abbr(rec):
    """Section key for a record: its abbrev, except that an unclassified record
    from UNR is promoted out of NEW into its own section."""
    abbr = (str(rec.get("abbr") or "").strip()) or "NEW"
    # A funder name typed where an abbreviation belongs ("Simons Foundation")
    # is a foundation, not a new federal section.
    if abbr not in SECTION_ORDER and not looks_like_abbr(abbr):
        abbr = "NEW"
    if abbr in ("NEW", "") and UNR_RE.search(str(rec.get("agency") or "")):
        return UNR_ABBR
    return abbr

def norm(t):
    if not t: return ""
    for k,v in ENCODING_MAP.items(): t = t.replace(k,v)
    return t.strip()

def anchor_for(abbr, full):
    base = (abbr or full or "sec").lower()
    return re.sub(r"[^a-z0-9]+","", base) or "sec"

def fmt_email_date(d):       # 'Jun. 15, 2026'
    return d.strftime("%b. ") + d.strftime("%d").lstrip("0") + d.strftime(", %Y")

def current_week_monday():
    """Monday of the current week (on/before today) — the email header shows
    the week the email is being built/sent, not the following week."""
    t = datetime.date.today()
    return t - datetime.timedelta(days=t.weekday())

def parse_date_arg(s):
    for fmt in ("%Y-%m-%d","%m/%d/%Y","%B %d, %Y","%b %d, %Y"):
        try: return datetime.datetime.strptime(s, fmt).date()
        except ValueError: continue
    raise SystemExit(f"Bad --email-date: {s}")

# ── Read validated sheet ─────────────────────────────────────────────────────
COL = {  # 1-indexed -> name (matches merge_and_assess HEADERS)
    1:"email",2:"eligflag",3:"type",4:"rec",5:"loi",6:"agency",7:"abbr",
    8:"title",9:"summary",10:"deadline",11:"l1url",12:"l1text",13:"l2url",
    14:"l2text",15:"elignotes",16:"recordid"}

def read_rows(path):
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    out = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not any(row): continue
        r = {COL[i+1]: (row[i] if i < len(row) else "") for i in range(16)}
        if str(r.get("email","")).strip().lower() == "y":
            out.append(r)
    return out

# ── Duplicate gate ───────────────────────────────────────────────────────────
# Last checkpoint before an email ships. merge_and_assess.py dedups records
# against each other at collection time, but Scout-agent-sourced records
# (RFI/DCL/Event/Policy items entered outside the four Python collectors) are
# written straight into the staging/validated workbook and never pass through
# that dedup, so the same opportunity can reach this file twice under
# different titles. Matching on the solicitation-number digits alone (not
# title wording, which is more prone to coincidental overlap) is a strong,
# low-false-positive signal that two rows are the same underlying
# opportunity -- so this hard-stops the build rather than just warning.
def normalized_identifier(s):
    """Alphanumeric key for a solicitation number. Comparing digits alone
    collapsed 'RFA-NS-27-002' and 'RFA-AI-27-002' onto the same key and failed
    the build on two distinct NIH solicitations, so the letters have to count.

    Returns "" unless the source has >=4 digits: this key is also derived from
    the link TEXT as a fallback, and a generic label like 'NIH Highlighted
    Topic' would otherwise normalize to a perfectly stable 'nihhighlightedtopic'
    and flag every pair of rows sharing that label as the same opportunity
    (it did -- two unrelated NIH topics in the 08-26 workbook). A real
    solicitation number always carries digits; a label like that never does."""
    s = s or ""
    if sum(c.isdigit() for c in s) < 4:
        return ""
    return re.sub(r"[^a-z0-9]", "", s.lower())

def _same_opportunity(a, b):
    """Rows sharing a solicitation number are only duplicates if they also
    share a link or a near-identical title. NSF 26-522 is one solicitation
    covering three separate programs (Astro-Core, Astro-Innovations, ATI),
    each with its own title and page -- those are three opportunities."""
    ua = str(a.get("l1url") or "").strip().rstrip("/").lower()
    ub = str(b.get("l1url") or "").strip().rstrip("/").lower()
    if ua and ua == ub:
        return True
    ta = re.sub(r"[^a-z0-9 ]", "", str(a.get("title") or "").lower()).strip()
    tb = re.sub(r"[^a-z0-9 ]", "", str(b.get("title") or "").lower()).strip()
    return bool(ta) and difflib.SequenceMatcher(None, ta, tb).ratio() >= 0.85


def find_duplicates(rows):
    """{solicitation key: [rows]} for rows that are the same opportunity."""
    by_num = {}
    for r in rows:
        num = (normalized_identifier(str(r.get("recordid") or "")) or
               normalized_identifier(str(r.get("l1text") or "")))
        if len(num) >= 4:
            by_num.setdefault(num, []).append(r)
    dups = {}
    for num, rs in by_num.items():
        hit = [r for r in rs if any(_same_opportunity(r, o) for o in rs if o is not r)]
        if hit:
            dups[num] = hit
    return dups


def check_no_duplicates(rows):
    dups = find_duplicates(rows)
    if not dups:
        return
    lines = ["Duplicate opportunity detected among Email?=Y rows -- fix the validated "
             "file before building (see 'Scout-agent-sourced records' in Known Open "
             "Issues):"]
    for num, rs in dups.items():
        titles = "; ".join(f"“{str(r.get('title') or '').strip()}” ({r.get('l1text')})" for r in rs)
        lines.append(f"  solicitation #{num}: {titles}")
    raise SystemExit("\n".join(lines))

# ── Deadline gate ────────────────────────────────────────────────────────────
# The other last checkpoint before an email ships. validate_staging.py checks
# the STAGING sheet, but every deadline defect that actually reached an inbox
# was introduced AFTER that -- during the summary-rewrite pass, which writes
# free text into the Deadline cell. Nothing validated the final workbook, so
# 'Septebmer 30, 2026', '2026-08-24T13:00:00' and 'n/a' all shipped. Same
# hard-stop treatment as the duplicate gate: a bad date in front of ~200
# researchers is worse than a failed build.
def check_deadlines(rows):
    problems = []
    for r in rows:
        text, probs = normalize_deadline(r.get("deadline"), r.get("type"))
        for p in probs:
            problems.append((str(r.get("title") or "(untitled)").strip(), p))
    if not problems:
        return
    lines = ["Deadline problems among Email?=Y rows -- fix the validated file "
             "before building:"]
    for title, p in problems:
        lines.append(f'  "{title[:60]}"')
        lines.append(f"      {p}")
    lines.append("")
    lines.append("Accepted shapes: a single date (rendered 'Month DD, YYYY'); "
                 "'proposals accepted at any time'; 'Archived - no current "
                 "deadline'; a multi-step/recurring form containing real dates; "
                 "or, for an Event, a date with a Pacific clock time.")
    raise SystemExit("\n".join(lines))

# ── HTML assembly ────────────────────────────────────────────────────────────
def grant_block(r):
    ttype = canonical_type(str(r.get("type") or "").strip())
    label = TYPE_LABELS.get(ttype, "")
    label_span = f'<span style="color:{REDLABEL};">{label}</span>' if label else ""
    title = norm(str(r.get("title") or ""))
    l1url = public_sam_url(str(r.get("l1url") or "").strip())
    l1text = norm(str(r.get("l1text") or "")) or "Link"
    # Deadline shape is settled by grants_common.normalize_deadline() and
    # enforced by check_deadlines() before we get here, so this only has to
    # take the normalized text.
    deadline, _ = normalize_deadline(r.get("deadline"), ttype)
    deadline = norm(deadline)
    summary = norm(str(r.get("summary") or ""))
    deadline_label = TYPE_DEADLINE_LABELS.get(ttype, "Application Deadline:")
    # optional Link2
    l2url = public_sam_url(str(r.get("l2url") or "").strip())
    l2text = norm(str(r.get("l2text") or ""))
    link_line = f'Link: <a href="{l1url}" style="color:{LINK}; font-weight:normal;"> {l1text}</a>'
    if l2url:
        link_line += f', <a href="{l2url}" style="color:{LINK}; font-weight:normal;"> {l2text or "More information"}</a>'
    return (
f'<!-- Grant Opportunity -->\n'
f'<tr><td><div><strong style="color:{NAVY}; font-size:12pt;">{label_span}{title}</strong></div></td></tr>\n'
f'<tr><td><div><font style="color:{NAVY}; font-size:9.75pt;">{link_line}\n</font></div></td></tr>\n'
f'<tr><td><div style="color:{NAVY}; font-size:9.75pt;">{deadline_label} {deadline}</div></td></tr>\n'
f'<tr><td><div><strong style="color:{NAVY}; font-size:9.75pt;">Summary:</strong></div></td></tr>\n'
f'<tr><td><div style="color:{NAVY}; font-size:9.75pt;">{summary}</div></td></tr>\n'
f'<tr><td>&nbsp;</td></tr>\n')

def build(rows, email_date):
    # group by section
    groups = {}
    for r in rows:
        groups.setdefault(section_abbr(r), []).append(r)
    # order: known federal abbrevs in SECTION_ORDER, any other federal, NEW last
    ordered = [a for a in SECTION_ORDER if a in groups]
    others = [a for a in groups if a not in SECTION_ORDER and a != "NEW"]
    ordered += sorted(others)
    if "NEW" in groups: ordered.append("NEW")

    def section_name(abbr, sample):
        if abbr == "NEW": return "New Foundation R&D Funding Opportunities"
        if abbr == UNR_ABBR: return "University of Nevada, Reno"
        return norm(str(sample.get("agency") or abbr))

    H = []
    H.append('<html lang="en">\n<head>\n<meta charset="utf-8">\n'
             '<meta content="width=device-width, initial-scale=1.0" name="viewport">\n'
             '<title>Grant Opportunities</title>\n</head>\n<body>\n'
             '<a id="top" name="top"></a>\n'
             '<table style="width:100%; align-content:center;">\n<tbody>\n')
    # title block
    H.append('<!-- TITLE AND SUBTITLE SECTION -->\n'
             '<tr>\n<td>\n<center>\n'
             f'<h1><font color="{TITLECOL}" face="Arial" size="6">Grants &amp; Sponsor Opportunities</font></h1>\n'
             '</center>\n</td>\n</tr>\n<tr>\n<td>\n<center>\n'
             f'<font color="{TITLECOL}" face="Arial" size="4">for the week starting {fmt_email_date(email_date)}</font>\n'
             '</center>\n</td>\n</tr>\n<tr><td>&nbsp;</td></tr>\n')
    # contents
    H.append(f'<tr><td><div style="color:{NAVY}; font-size:18pt; font-weight:normal;">Contents:</div></td></tr>\n')
    for abbr in ordered:
        nm = section_name(abbr, groups[abbr][0])
        H.append(f'<tr><td><div><a href="#{anchor_for(abbr,nm)}" style="color:{LINK}; font-weight:normal;">{nm}</a></div></td></tr>\n')
    # sections
    for abbr in ordered:
        nm = section_name(abbr, groups[abbr][0])
        anc = anchor_for(abbr, nm)
        H.append('<!-- START AGENCY SECTION -->\n<!-- Agency heading -->\n'
                 '<tr><td><div>&nbsp;</div></td></tr>\n'
                 f'<tr><td><div><h2 style="color:{NAVY}; font-size:18pt; font-weight:normal;">'
                 f'<a name="{anc}" style="text-decoration:none; color:inherit;">{nm}</a></h2></div></td></tr>\n'
                 '<tr><td>&nbsp;</td></tr>\n')
        for r in groups[abbr]:
            H.append(grant_block(r))
        H.append('<!-- Return to top -->\n'
                 f'<tr><td><div><a href="#top" style="color:{LINK}; font-weight:normal;">Return to top</a></div></td></tr>\n'
                 '<!-- END AGENCY SECTION -->\n')
    # footer
    H.append('<!-- DO NOT EDIT BELOW THIS -->\n<!-- Credits to sources -->\n'
             f'<tr><td><div style="color:{NAVY}; font-size:9.75pt;">For questions, contact the '
             f'<a href="mailto:researchdevelopment@unr.edu" style="color:{LINK}; font-weight:normal;">'
             'Research and Proposal Development Services</a>.</div></td></tr>\n'
             '<tr><td><div>&nbsp;</div></td></tr>\n<tr><td><div>&nbsp;</div></td></tr>\n'
             '  </tbody>\n</table>\n</body>\n</html>\n')
    html = "".join(H)
    for k,v in ENCODING_MAP.items(): html = html.replace(k,v)
    # Defensive: anything still above U+00FF here would blow up the latin-1
    # write in main(). ENCODING_MAP covers known cases (smart quotes, dashes,
    # odd spaces); this is the backstop for whatever wasn't anticipated --
    # a weekly email should never fail to build over one stray character.
    stray = sorted({c for c in html if ord(c) > 0xFF})
    if stray:
        print("WARNING: stripped unmapped character(s) above U+00FF: "
              + ", ".join(f"U+{ord(c):04X}" for c in stray), file=sys.stderr)
        html = "".join(c for c in html if ord(c) <= 0xFF)
    html = html.replace("\r\n","\n").replace("\r","\n").replace("\n","\r\n")
    return html

def resolve_final_input(raw_path):
    """Always build from the post-rewrite validated_*_final.xlsx file.

    Accepts a WEEK_DIR, a validated_*.xlsx, or a validated_*_final.xlsx and
    resolves to the _final file in every case, since building from the
    pre-rewrite validated file (before the Step 4 summary rewrite ran) is a
    known failure mode this pipeline must avoid.
    """
    p = Path(raw_path)
    if p.is_dir():
        candidates = sorted(p.glob("validated_*_final.xlsx"),
                             key=lambda f: f.stat().st_mtime, reverse=True)
        if not candidates:
            raise SystemExit(f"No validated_*_final.xlsx found in {p}. "
                              f"Run the Step 4 summary rewrite first.")
        if len(candidates) > 1:
            print(f"Note: multiple _final files in {p}; using newest: {candidates[0].name}")
        return str(candidates[0])
    if p.name.endswith("_final.xlsx"):
        if not p.exists():
            raise SystemExit(f"Input file not found: {p}")
        return str(p)
    final_candidate = p.with_name(p.stem + "_final" + p.suffix)
    if final_candidate.exists():
        print(f"Note: building from {final_candidate.name} (final overrides {p.name}).")
        return str(final_candidate)
    if not p.exists():
        raise SystemExit(f"Input file not found: {p}")
    raise SystemExit(f"No {final_candidate.name} found next to {p.name}. "
                      f"Run the Step 4 summary rewrite first — build_html.py "
                      f"never builds from the pre-rewrite validated file.")

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("infile", help="WEEK_DIR, validated_*.xlsx, or validated_*_final.xlsx "
                                    "-- always resolves to the _final.xlsx file")
    ap.add_argument("--email-date", default=None)
    ap.add_argument("--out", default=None)
    args = ap.parse_args()
    infile = resolve_final_input(args.infile)
    ed = parse_date_arg(args.email_date) if args.email_date else current_week_monday()
    rows = read_rows(infile)
    if not rows:
        raise SystemExit("No rows with Email?=Y found. Set Email?=Y on rows to include.")
    check_no_duplicates(rows)
    check_deadlines(rows)
    html = build(rows, ed)
    out = args.out or f"Grant_Opportunities_{ed.strftime('%m-%d-%Y')}.html"
    with open(out, "w", encoding="latin-1", errors="strict") as f:
        f.write(html)
    # quick verify
    raw = open(out, "rb").read()
    print(f"Wrote {out}")
    print(f"  records: {len(rows)} | sections: {html.count('START AGENCY SECTION')} | "
          f"CRLF: {raw.count(b(chr(13)+chr(10)))} | lone LF: {raw.count(b(chr(10)))-raw.count(b(chr(13)+chr(10)))}")

def b(s): return s.encode("latin-1")

if __name__ == "__main__":
    main()
