#!/usr/bin/env python3
"""
grants_common.py
================
Shared helpers and the canonical record schema used by ALL collectors
(grants_collect / pivot_collect / nih_rss_collect / nsf_rss_collect) and by
merge_and_assess.py.

Every collector emits a list of dicts in CANONICAL_FIELDS shape and writes JSON.
merge_and_assess.py reads any mix of those JSON files.

Keeping this in one place means the schema can't drift between collectors.
"""
import re, datetime

# ── Canonical record schema ──────────────────────────────────────────────────
# Every collector must emit dicts with (at least) these keys. Missing -> None/"".
CANONICAL_FIELDS = [
    "source",            # 'federal' | 'pivot' | 'nih_guide' | 'nsf_rss'
    "opp_id",            # internal id (grants.gov numeric id, pivot id, guide id)
    "opp_number",        # public number (PA-..., PAR-..., NSF 26-506, etc.)  -> dedup key
    "title",
    "agency_code",       # grants.gov agency code if known (e.g. 'NSF', 'USDA-NIFA')
    "agency_name",       # raw funder/agency name
    "open_date",         # posting/publication date, ISO yyyy-mm-dd if known
    "post_date",         # explicit post date (federal) ISO; used for recency
    "close_date",        # best single deadline, human string 'Month DD, YYYY' or raw
    "close_date_raw",    # original unparsed deadline text (for multi-date review)
    "status",            # 'posted' etc.
    "url",               # primary link (grants.gov detail, agency page, funder page)
    "pivot_url",         # secondary link (Pivot-RP) if any
    "description",       # abstract / synopsis text (NOT email-ready; Claude rewrites)
    "eligibility_desc",  # eligibility prose
    "career_stage",      # pivot career stage text if any
    "activity_location", # pivot activity location if any
    "award_amount",      # raw amount text if any
    "cost_sharing",      # 'Yes'/'No'/'' if known
    "type_hint",         # 'GRANT' | 'RFI' | 'DCL' | 'EVENT' | 'NOTICE' | '' (from title)
    "needs_review",      # list[str] of flags the collector already found
    "summary",           # '' – Claude fills this during the drafting pass
]

# ── Encoding normalization (same map the SKILL uses for HTML) ────────────────
ENCODING_MAP = {
    "‘": "'", "’": "'", "“": '"', "”": '"',
    "–": "-", "—": "-", "‑": "-", "…": "...",
    "•": "-", " ": " ",
    " ": " ",  # narrow no-break space
    " ": " ",  # thin space
    " ": " ",  # figure space
    "​": "",   # zero-width space
}
def normalize_encoding(text):
    if not text:
        return text or ""
    for k, v in ENCODING_MAP.items():
        text = text.replace(k, v)
    return re.sub(r"[ \t]+", " ", text).strip()

# ── Staging-sheet Type vocabulary (single source of truth) ───────────────────
# The values a human can put in the "Type" column, shared by:
#   build_review_html.py  -- TYPE_OPTIONS populates the Type <select> dropdown
#   merge_and_assess.py   -- same list becomes the Excel data-validation dropdown
#   build_html.py         -- TYPE_LABELS gives the red-prefix email label (or no
#                             label for Grant), TYPE_DEADLINE_LABELS gives the
#                             date-line label ("Application Deadline:" unless
#                             overridden, e.g. Event -> "Date:")
# Previously each file kept its own copy and they drifted: "Limited Submission"
# was a valid dropdown choice in two files but had no entry in TYPE_LABELS, so
# it silently rendered with no red label in the email (the 08-13 DEPSCoR
# record had to smuggle "LIMITED SUBMISSION." into the title as a workaround).
# "Policy" formalizes what several comments already called Scout-agent-sourced
# "Policy items" (see P2-2 in PIPELINE_IMPROVEMENTS.md) as a real, selectable
# type instead of an undocumented convention.
#
# NOTE: this is a different vocabulary from detect_type()'s GRANT/RFI/DCL/
# EVENT/NOTICE below, which classifies raw collector input, not sheet values.
TYPE_INFO = [
    # (code,                email label or None,          deadline-line label)
    ("Grant",               None,                          "Application Deadline:"),
    ("DCL",                 "Dear Colleague Letter ",      "Application Deadline:"),
    ("RFI",                 "Request for Information ",    "Application Deadline:"),
    ("Event",               "Event ",                      "Date:"),
    ("Limited Submission",  "Limited Submission ",         "Application Deadline:"),
    ("Policy",              "Policy ",                     "Effective Date:"),
]
TYPE_OPTIONS = [""] + [code for code, _, _ in TYPE_INFO]
TYPE_LABELS = {code: label for code, label, _ in TYPE_INFO if label}
TYPE_DEADLINE_LABELS = {code: dl for code, _, dl in TYPE_INFO}
_TYPE_LOOKUP = {code.lower(): code for code, _, _ in TYPE_INFO}

def canonical_type(t):
    """Case/whitespace-insensitive lookup into TYPE_INFO. Returns the
    canonical code (e.g. 'event' -> 'Event'), or the stripped input unchanged
    if it isn't a recognized type, so unknown types still fail validation
    rather than being silently coerced into something valid."""
    key = (t or "").strip()
    return _TYPE_LOOKUP.get(key.lower(), key)

# ── Type detection from title/heading (shared by every collector) ────────────
def detect_type(title, number=""):
    """Return one of GRANT/RFI/DCL/EVENT/NOTICE based on title + number prefix."""
    t = (title or "").lower()
    n = (number or "").upper()
    if "dear colleague" in t or "dcl" in t.split():
        return "DCL"
    if "request for information" in t or re.search(r"\brfi\b", t) or n.startswith("NOT-") and "information" in t:
        return "RFI"
    if any(k in t for k in ["webinar", "workshop", "seminar", "information session",
                            "info session", "office hours", "conference", "symposium"]):
        return "EVENT"
    if n.startswith("NOT-"):
        # NIH guide notice that isn't an RFI/DCL -> generic notice (usually skipped)
        return "NOTICE"
    return "GRANT"

# ── Deadline parsing – handles the Pivot 'DD Mon YYYY - ...' multi-date mess ──
TODAY = datetime.date.today()
_MONTHS = "(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
_DATE_PATTERNS = [
    (re.compile(rf"\b(\d{{1,2}})\s+({_MONTHS})[a-z]*\s+(\d{{4}})\b"), "%d %b %Y"),  # 10 Jul 2026
    (re.compile(r"\b([A-Z][a-z]+)\s+(\d{1,2}),\s*(\d{4})\b"), "%B %d %Y"),          # July 10, 2026
    (re.compile(r"\b(\d{4})-(\d{2})-(\d{2})\b"), "%Y-%m-%d"),                       # 2026-07-10
    (re.compile(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b"), "%m/%d/%Y"),                   # 07/10/2026
]
def _try_parse(token, fmt):
    try:
        if fmt == "%d %b %Y":
            d, mon, y = token
            return datetime.datetime.strptime(f"{int(d):02d} {mon[:3]} {y}", "%d %b %Y").date()
        if fmt == "%B %d %Y":
            mon, d, y = token
            return datetime.datetime.strptime(f"{mon} {int(d):02d} {y}", "%B %d %Y").date()
        if fmt == "%Y-%m-%d":
            y, m, d = token
            return datetime.date(int(y), int(m), int(d))
        if fmt == "%m/%d/%Y":
            m, d, y = token
            return datetime.date(int(y), int(m), int(d))
    except (ValueError, IndexError):
        return None

def extract_dates(raw):
    """Return sorted list of unique date objects found in a messy string."""
    if not raw:
        return []
    found = []
    for rx, fmt in _DATE_PATTERNS:
        for m in rx.findall(raw):
            d = _try_parse(m, fmt)
            if d:
                found.append(d)
    return sorted(set(found))

def assess_deadline(raw):
    """
    Return (display_str, status, days_out, multi_flag).
    status in {'ok','soon','passed','rolling','unknown'}.
    Picks the FIRST date that is today-or-later. Flags if >1 date present.
    """
    raw = (raw or "").strip()
    low = raw.lower()
    if not raw:
        return ("rolling / not specified", "rolling", None, None)
    if any(k in low for k in ["rolling", "accepted anytime", "any time", "anytime",
                              "no deadline", "continuous", "open until filled"]):
        # still try to find a concrete date too; rolling wins only if no date
        dates = extract_dates(raw)
        if not dates:
            return ("proposals accepted at any time", "rolling", None, None)
    dates = extract_dates(raw)
    if not dates:
        return (raw[:60], "unknown", None, None)
    future = [d for d in dates if (d - TODAY).days >= 0]
    multi = f"multiple dates found: {', '.join(d.strftime('%b %d %Y') for d in dates)}" if len(dates) > 1 else None
    if not future:
        # all dates passed
        latest = dates[-1]
        return (latest.strftime("%B %d, %Y"), "passed", (latest - TODAY).days, multi)
    d = future[0]
    days = (d - TODAY).days
    disp = d.strftime("%B %d, %Y")
    status = "soon" if days <= 14 else "ok"
    return (disp, status, days, multi)

# ── Email-ready deadline normalization (the final gate) ──────────────────────
# assess_deadline() above runs at COLLECTION time and is deliberately lenient:
# when it can't parse anything it returns (raw[:60], "unknown") so a human can
# sort it out during review. normalize_deadline() is the opposite end -- it runs
# at BUILD time, after the human/Claude rewrite pass, and its job is to be
# strict about what may reach a recipient's inbox.
#
# Why this exists: the June 22 email rendered 45/45 deadlines as a uniform
# "Month DD, YYYY". Once the Type vocabulary added per-type date labels
# (TYPE_DEADLINE_LABELS above), the Deadline cell effectively became free text
# and the rewrite pass started writing prose into it. Shipped since:
#     "Septebmer 30, 2026"        <- typo, 2026-09-01 email
#     "2026-08-24T13:00:00"       <- ISO timestamp, 2026-08-17 email
#     "n/a"                       <- placeholder, 2026-08-24 email
#     "October 19-23, 2026"       <- a range under "Application Deadline:"
#     "Date: proposals accepted at any time"   <- an Event with no date
# build_html.py stores this as a string, so the isinstance(datetime) guard it
# used to rely on never fired for any of them.
#
# A Deadline cell is valid only if it is one of these shapes:
#   S1  a single date            -> normalized to "Month DD, YYYY"
#   S2  the rolling canon        -> ROLLING_CANON
#   S3  the archived canon       -> ARCHIVED_CANON
#   S4  an Event date, with any clock time labelled Pacific
#   S5  a compound/recurring form carrying >=1 parseable date, or matching
#       ALLOWED_NONDATE_RE (multi-step, target dates, LOI dates, ...)
# Anything else is returned unchanged alongside a non-empty problem list, and
# build_html.py hard-stops rather than shipping it.

_MONTH_FULL = ["january", "february", "march", "april", "may", "june", "july",
               "august", "september", "october", "november", "december"]
_MONTH_VOCAB = set(_MONTH_FULL) | {m[:3] for m in _MONTH_FULL} | {"sept"}

ROLLING_CANON = "proposals accepted at any time"
ARCHIVED_CANON = "Archived - no current deadline"

_ROLLING_RE = re.compile(
    r"rolling|accepted\s+any\s?time|any\s?time|no\s+deadline|continuous|"
    r"open\s+until\s+filled|not\s+specified", re.I)
_ARCHIVED_RE = re.compile(r"\barchived\b|\bexpired\b|\bno longer\b", re.I)
_PLACEHOLDER_RE = re.compile(r"^(n/?a|tbd|tba|none|null|unknown|\?+|-+|\.+)$", re.I)
_ISO_DT_RE = re.compile(r"\b\d{4}-\d{2}-\d{2}[T ]\d{1,2}:\d{2}")
_ISO_D_RE = re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})$")
# 24-hour clock: a time not qualified by am/pm. "13:00" yes, "1:00 PM" no --
# and in a range ("1:00-2:30 PM") the am/pm on the far end qualifies both, so
# the lookahead has to see past the second time before deciding.
_CLOCK24_RE = re.compile(
    r"\b(?:[01]?\d|2[0-3]):[0-5]\d\b"
    r"(?!\s*(?:[APap]\.?[Mm]|(?:[-–—]|\s+to\s+)\s*\d{1,2}:[0-5]\d\s*[APap]\.?[Mm]))")
_AMPM_RE = re.compile(r"\d\s*(?::\d\d)?\s*[APap]\.?[Mm]\b")
_PACIFIC_RE = re.compile(r"\bP[DS]?T\b|\bPacific\b")
# "Month DD, YYYY" and "DD Month YYYY" -- the two shapes whose month word we
# can safely spell-check without guessing at arbitrary prose.
_MONTHISH_RE = re.compile(r"\b([A-Za-z]{3,12})\s+(\d{1,2})(?:st|nd|rd|th)?\s*,\s*(\d{4})\b")
_DAYFIRST_RE = re.compile(r"\b(\d{1,2})\s+([A-Za-z]{3,12})\s+(\d{4})\b")
# Legitimate non-single-date phrasings the rewrite pass is allowed to use.
ALLOWED_NONDATE_RE = re.compile(
    r"recurring|target\s+date|letter\s+of\s+intent|\bLOI\b|step[\s-]?[12]|"
    r"see\s+(?:the\s+)?(?:solicitation|program\s+page|announcement)|"
    r"varies|multiple\s+deadlines|full\s+proposal", re.I)


def _levenshtein(a, b):
    if a == b:
        return 0
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def _month_typos(s):
    """Flag a word sitting in a date's month slot that isn't a month but is
    close to one. Catches 'Septebmer 30, 2026' (distance 2 from 'september')
    without firing on 'Step 1' or 'Proposals:' (nowhere near any month)."""
    out = []
    for rx, widx in ((_MONTHISH_RE, 1), (_DAYFIRST_RE, 2)):
        for m in rx.finditer(s):
            w = m.group(widx)
            if w.lower() in _MONTH_VOCAB:
                continue
            near = min(_MONTH_FULL, key=lambda mo: _levenshtein(w.lower(), mo))
            if _levenshtein(w.lower(), near) <= 3:
                out.append(f"{w!r} is not a month -- did you mean {near.title()!r}?")
    return out


_SAM_WORKSPACE_RE = re.compile(r"sam\.gov/workspace/contract/opp/", re.I)


def public_sam_url(url):
    """The SAM.gov API's uiLink points at the signed-in workspace view, which
    readers without a SAM account can't open. The public page is /opp/<id>/view."""
    return _SAM_WORKSPACE_RE.sub("sam.gov/opp/", url or "")


def normalize_deadline(value, ttype="Grant"):
    """Return (display_text, problems).

    display_text is what the email should show. A non-empty problems list means
    the cell needs a human before the email ships -- callers should treat it as
    a build-stopping error, not a warning, because every item on the list above
    reached a real inbox as a warning-free pass-through.
    """
    ttype = canonical_type(ttype or "Grant")

    # A cell re-typed as a real Excel date by a manual edit or a pandas
    # round-trip. str() would give '2026-09-17 00:00:00'.
    if isinstance(value, (datetime.date, datetime.datetime)):
        return value.strftime("%B %d, %Y"), []

    s = normalize_encoding(str(value or ""))
    if not s:
        return "", ["deadline is empty"]

    if _PLACEHOLDER_RE.match(s):
        return s, [f"{s!r} is a placeholder, not a deadline -- use a real date, "
                   f"{ROLLING_CANON!r}, or {ARCHIVED_CANON!r}"]

    if _ISO_DT_RE.search(s):
        return s, [f"{s!r} is a raw ISO timestamp -- write a readable date"
                   + (", and a Pacific clock time" if ttype == "Event" else "")]

    m = _ISO_D_RE.match(s)
    if m:
        try:
            d = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            return d.strftime("%B %d, %Y"), []
        except ValueError:
            return s, [f"{s!r} is not a valid date"]

    problems = _month_typos(s)
    dates = extract_dates(s)

    # A non-Event deadline with a cutoff time ("January 27, 2027, 12:00 p.m. ET"):
    # the email shows plain dates, so drop the time rather than stopping the build.
    if (ttype != "Event" and not problems and len(dates) == 1
            and (_AMPM_RE.search(s) or _CLOCK24_RE.search(s))
            and not ALLOWED_NONDATE_RE.search(s)):
        return dates[0].strftime("%B %d, %Y"), []

    # S2/S3: phrasing-only cells, canonicalized so the batch reads consistently.
    if not dates and not problems:
        if _ARCHIVED_RE.search(s):
            return (s if ttype == "Event" else ARCHIVED_CANON,
                    ["an Event needs an actual date, not an archived marker"]
                    if ttype == "Event" else [])
        if _ROLLING_RE.search(s):
            if ttype == "Event":
                return s, ["an Event needs an actual date, not a rolling phrase"]
            return ROLLING_CANON, []

    # S1: exactly one date and nothing else of substance -> canonical form.
    if len(dates) == 1 and not ALLOWED_NONDATE_RE.search(s) and not _AMPM_RE.search(s):
        residue = _MONTHISH_RE.sub("", _DAYFIRST_RE.sub("", s))
        residue = re.sub(r"[\s,.;:()-]|\b\d{4}\b", "", residue)
        if not residue:
            s = dates[0].strftime("%B %d, %Y")

    # Ordinary deadlines carry no clock time; Events keep theirs but must say PT.
    if _CLOCK24_RE.search(s):
        problems.append(f"{s!r} carries a 24-hour clock time -- "
                        + ("convert it to a Pacific am/pm time" if ttype == "Event"
                           else "an ordinary deadline should be a plain date"))
    elif ttype != "Event" and _AMPM_RE.search(s) and not _PACIFIC_RE.search(s):
        problems.append(f"{s!r} carries a clock time with no time zone")
    if ttype == "Event" and _AMPM_RE.search(s) and not _PACIFIC_RE.search(s):
        problems.append(f"{s!r} is an Event time not stated in Pacific -- "
                        "convert to PT and keep the source zone in parentheses")

    # S5: everything else must still contain a real date or a known phrasing.
    if not dates and not problems and not ALLOWED_NONDATE_RE.search(s):
        problems.append(f"no recognizable date in {s!r}")

    return s, problems


def iso(d):
    return d.strftime("%Y-%m-%d") if isinstance(d, (datetime.date, datetime.datetime)) else (d or "")

def blank_record():
    return {k: ("" if k not in ("needs_review",) else []) for k in CANONICAL_FIELDS}
