#!/usr/bin/env python3
"""
sam_collect.py  (v3)
====================
Federal collector for the gap grants_collect.py (Grants.gov) can't cover:
agencies that post funding calls as SAM.gov CONTRACT opportunities (BAAs,
and RFIs filed as Sources Sought) instead of, or in addition to, Grants.gov
assistance-listing FOAs.

WHAT CHANGED IN v3 (2026-08-06) -- read this before touching anything
---------------------------------------------------------------------
v2 wrote EVERY hit to the output file. score_hit() computed a relevance score
but main() re-sorted by _days_out and never applied a threshold, so the score
only decided which notice got the (expensive) detail fetch. A real run on
2026-08-06 produced 68 records of which 60 were DOE facilities procurement
(substation transformers, roof replacement, trash collection, propane).
Measured precision: ~12%.

v3 fixes four things:

1. PSC RELEVANCE GATE (the actual fix). The v2 record never captured
   classificationCode. Product Service Code separates R&D from procurement
   cleanly -- the entire PSC "A" series is Research & Development. Verified
   against two live notices from the same NSF office on the same day:
       BAA: Establishing Consortia to Advance Energy-Water Security -> AJ11
       Commercial Business Intelligence Data Feed Services          -> DA10
   NAICS does NOT work for this: that same BAA carries NAICS 813920
   ("Professional Organizations"), so a 541715 R&D filter would have DROPPED
   the exact record we wanted. Do not reintroduce a NAICS filter.
   ClassificationCode is column 19 of SAM's own ContractOpportunitiesFullCSV
   extract, confirming it is a first-class field on these notices.

2. PTYPE CODES WERE WRONG. v2 sent ptype=["o","p","g"] with a comment saying
   g = Sources Sought. It isn't. g = Sale of Surplus Property. The real codes:
       o = Solicitation          p = Presolicitation
       k = Combined Synopsis/Solicitation
       r = Sources Sought        s = Special Notice
       a = Award Notice (deliberately excluded -- already awarded)
   v2 therefore returned zero Sources Sought and zero Combined Synopsis
   notices. Many BAAs post as k; BAA amendments and industry days post as s.

3. ISO DEADLINE PARSE. SAM returns '2026-08-10T12:00:00-04:00'.
   grants_common.assess_deadline() cannot parse that -- its ISO pattern is
   \\b(\\d{4})-(\\d{2})-(\\d{2})\\b and the trailing \\b fails against the 'T'.
   Every SAM record was coming back status='unknown', days=None. That silently
   disabled the lead-time filter (the guard is `days is not None and ...`),
   which is the ONLY reason the Energy-Water BAA survived v2 at all. Fixing
   the parse without also revisiting MIN_LEAD_DAYS would have DROPPED it.
   See MIN_LEAD_DAYS below.

4. RAW RESPONSE CACHING. The key is capped at ~10 calls/day, which makes
   iterative development impossible. Every search response is now written to
   the cache dir, and --from-cache replays them at zero API cost. Tune filters
   offline; spend calls only on fresh data.

WHY NOT THE BULK CSV
--------------------
SAM.gov publishes ContractOpportunitiesFullCSV.csv (Data Services > Contract
Opportunities > datagov) with no API key and no rate limit. It was evaluated
on 2026-08-06 and rejected for now: three separate downloads returned a
well-formed 685-byte header row with ZERO data rows. The download link is also
a javascript:void(0) handler that mints a 45-minute AWS signed URL per click,
so it cannot be scripted without a browser step. Worth re-testing later -- if
SAM fixes the empty-file problem it eliminates the API key entirely, which
matters for backup staff. The column names are recorded in CSV_COLUMNS below
so a future CSV-based collector can be written without re-deriving them.

SAM.gov has no structured eligibility field for contract opportunities
(unlike Grants.gov's applicantEligibilityDesc). Eligibility language, if
present at all, lives in the free-text notice description and/or attached
BAA documents (PDF/DOCX). This script pulls both for the shortlisted
candidates and extracts any sentence containing eligibility-flavored
keywords into eligibility_desc for human review -- it does NOT make an
eligibility determination itself.

Reads the key from the SAM_API_KEY environment variable. Never hardcode,
print, or log the key value.

Usage:
    python3 sam_collect.py --out sam_opportunities.json
    python3 sam_collect.py --no-detail                 # search only, 7 calls
    python3 sam_collect.py --from-cache 2026-08-06     # ZERO calls, replay
    python3 sam_collect.py --keep-all                  # disable PSC gate (debug)
"""
import os, re, sys, json, time, glob, argparse, datetime
import urllib.request, urllib.parse, urllib.error
from io import BytesIO
from grants_common import (normalize_encoding, detect_type, assess_deadline,
                           blank_record, iso, public_sam_url)

API_KEY = os.environ.get("SAM_API_KEY")
SEARCH_URL = "https://api.sam.gov/opportunities/v2/search"

# ── Org codes ────────────────────────────────────────────────────────────────
# CAVEAT: SAM's agency-hierarchy filtering has been reported unreliable for
# years. A 0-hit result from an org code may mean the code is wrong, the org
# posted nothing, OR the hierarchy filter silently failed. Treat a persistent
# 0 as "unknown", not "broken code". Confirmed live 2026-07-27 unless noted.
ORGS = [
    ("8900",   "Department of Energy"),
    ("4900",   "National Science Foundation"),
    ("HR0011", "DARPA"),
    ("N00014", "Office of Naval Research"),
    ("W911NF", "Army Research Office"),   # routes through ACC-APG
    ("HM0476", "NGA"),
    ("FA9550", "AFOSR"),                  # UNCONFIRMED -- 0 hits in a 30-day test window
]

# ── Notice types ─────────────────────────────────────────────────────────────
# 'a' (Award Notice) deliberately excluded: already awarded, nothing to apply to.
# Adding ptype values costs nothing extra -- still 1 search call per org.
PTYPES = ["o", "p", "k", "r", "s"]

POSTED_WITHIN_DAYS = 14

# ── POLICY KNOB -- read this ─────────────────────────────────────────────────
# 0 = keep tight-deadline notices and FLAG them. 10 = drop anything closing
# within 10 days (the grants_collect.py convention).
# Set to 0 deliberately: BAAs are amended constantly (the Energy-Water BAA was
# on Amendment 3 with an extended white-paper date), and a 4-day BAA deadline
# is still worth surfacing to a PI with a flag. Change to 10 if the weekly
# email should never carry a sub-10-day contract deadline.
MIN_LEAD_DAYS = 0

SEARCH_LIMIT = 1000          # API max. Same 1 call as limit=100, 10x the headroom.
DEFAULT_MAX_DETAIL = 1
DESCRIPTION_MAX_CHARS = 4000  # raw material for the later Claude rewrite pass,
                               # not an email-ready summary -- capped so one
                               # multi-page BAA attachment doesn't blow up the cell
CACHE_DIR = "sam_cache"

# ── PSC relevance gate ───────────────────────────────────────────────────────
# The whole PSC "A" series is Research & Development (AA agriculture,
# AC defense, AJ general science & technology, AN space, AZ other R&D...).
# Everything the 2026-08-06 noise consisted of sits outside it: Y/Z construction
# and maintenance, S utilities and housekeeping, 7A IT, 44/54/55 supplies,
# D IT services.
PSC_RESEARCH_PREFIX = "A"

# Titles that name a research funding vehicle explicitly. These BYPASS the PSC
# gate (kept even with a non-A PSC) because a mis-coded BAA is a costly miss,
# and they are flagged so the reviewer knows why they survived. Keep this list
# TIGHT -- do not add generic words like "research" or "program", which is how
# v2's score_hit() ended up scoring facilities work as relevant.
BAA_TITLE_RE = re.compile(
    r"\bs?baa\b|broad agency announcement|research announcement|"
    r"\bbroad agency\b|program solicitation|funding opportunity|"
    r"\bnra\b|research opportunity",
    re.I,
)

# ── NAICS secondary signal (v4, 2026-08-14) ─────────────────────────────────
# PSC "A" alone still isn't enough: DOE/SLAC posts physical lab-equipment
# purchases and materials-licensing deals under R&D-series PSC codes. Reviewed
# against a real 25-hit batch on 2026-08-14 -- 7 of the 10 actual junk records
# shared a manufacturing/mining-sector NAICS despite an "A" PSC: Lift Truck,
# "FDG 150 Floating Ion", "PXle Module" (NAICS 3333, machinery mfg), High
# Definition Oscilloscope, PHOTEK PHOTODETECTOR (334515, instrument mfg), a
# "TECHNOLOGY LICENSING OPPORTUNITY" notice (325211, plastics mfg), and
# "E-Seal Technology...Leaky Well Remediation" (213112, oil/gas support).
# Real BAAs kept fine in the same batch: PENG, RAAPTR, PINPOINT, NGA BIG-ST,
# ManTech sBAA -- all NAICS 5417xx (Professional/Scientific/Technical
# Services). This is a DENYLIST, not the reverse: v2 already learned the hard
# way (see WHAT CHANGED IN v3 above) that requiring a specific "good" NAICS
# drops real BAAs with odd codes (the Energy-Water BAA carried 813920,
# "Professional Organizations" -- sector 81, not manufacturing/goods, so it
# still survives this denylist). Only deny sectors that are essentially never
# a university funding opportunity: mining/oil&gas, utilities, construction,
# manufacturing, wholesale/retail trade, transportation.
NAICS_NON_RESEARCH_SECTORS = {
    "21", "22", "23",              # mining/oil & gas, utilities, construction
    "31", "32", "33",              # manufacturing
    "42", "44", "45", "48", "49",  # wholesale/retail trade, transportation
}

# Titles that are categorically not something to apply/propose to, regardless
# of PSC/NAICS. "Notice of Intent to Award" and "TECHNOLOGY LICENSING
# OPPORTUNITY" were already known (see P1-5 in PIPELINE_IMPROVEMENTS.md).
# "Proposers Day"/"Industry Day" confirmed unwanted in the same 2026-08-14
# batch -- informational events pitched at industry contractors, not an
# opportunity in themselves. Unconditional drop: unlike the NAICS denylist,
# no BAA-title override, because these patterns describe the notice's own
# nature, not a proxy that could misfire on a real solicitation.
JUNK_TITLE_RE = re.compile(
    r"^\s*notice of intent to award|technology licensing opportunity|"
    r"\bproposers?\s+day\b|\bindustry\s+day\b",
    re.I,
)

# Column names from SAM's ContractOpportunitiesFullCSV.csv extract, captured
# 2026-08-06. Recorded for a future CSV-based collector; unused at runtime.
CSV_COLUMNS = [
    "NoticeId", "Title", "Sol#", "Department/Ind.Agency", "CGAC", "Sub-Tier",
    "FPDS Code", "Office", "AAC Code", "PostedDate", "Type", "BaseType",
    "ArchiveType", "ArchiveDate", "SetASideCode", "SetASide", "ResponseDeadLine",
    "NaicsCode", "ClassificationCode", "PopStreetAddress", "PopCity", "PopState",
    "PopZip", "PopCountry", "Active", "AwardNumber", "AwardDate", "Award$",
    "Awardee", "PrimaryContactTitle", "PrimaryContactFullname",
    "PrimaryContactEmail", "PrimaryContactPhone", "PrimaryContactFax",
    "SecondaryContactTitle", "SecondaryContactFullname", "SecondaryContactEmail",
    "SecondaryContactPhone", "SecondaryContactFax", "OrganizationType", "State",
    "City", "ZipCode", "CountryCode", "AdditionalInfoLink", "Link", "Description",
]

# ── Eligibility keyword sets (unchanged from v2) ─────────────────────────────
HIGHER_ED_KEYWORDS = [
    "institution of higher education", "institutions of higher education",
    "accredited college", "accredited university", "college or university",
    "colleges and universities", "academic institution", "academic institutions",
    "university", "universities", "research institution",
]
SET_ASIDE_KEYWORDS = [
    "small business set-aside", "small business set aside", "set-aside for small business",
    "total small business", "8(a) set-aside", "8(a) business development",
    "hubzone set-aside", "sdvosb set-aside", "service-disabled veteran-owned set-aside",
    "woman-owned small business set-aside", "wosb set-aside",
]
ELIGIBILITY_KEYWORDS = [
    "eligib", "may submit", "may propose", "open to", "restricted to",
    "small business", "8(a)", "hubzone", "sdvosb", "woman-owned", "women-owned",
    "u.s. citizen", "domestic entity", "ffrdc", "federally funded research and development",
    "any responsible source", "non-us", "non-u.s.", "foreign entit", "nonprofit", "non-profit",
]

_calls_made = 0


def _mask(url):
    return re.sub(r"([?&]api_key=)[^&]+", r"\1REDACTED", url)


# ── Deadline parsing ─────────────────────────────────────────────────────────
ISO_DT_RE = re.compile(r"^(\d{4}-\d{2}-\d{2})(?:[T ]|$)")


def parse_sam_deadline(raw):
    """SAM returns ISO-8601 with timezone, e.g. '2026-08-10T12:00:00-04:00'.
    grants_common.assess_deadline() cannot parse that (its ISO pattern ends in
    \\b, which fails against the 'T'), so it silently returned 'unknown' for
    every SAM record. Strip to the date part here, then hand a clean string to
    the shared helper so status semantics stay identical across all collectors.
    Returns the same 4-tuple as assess_deadline()."""
    raw = (raw or "").strip()
    if not raw:
        return assess_deadline("")
    m = ISO_DT_RE.match(raw)
    if m:
        return assess_deadline(m.group(1))
    return assess_deadline(raw)


# ── Relevance gate ───────────────────────────────────────────────────────────
def relevance(hit):
    """Return (keep: bool, reason: str, flags: list[str]).

    Decision order:
      1. Title matches a known non-opportunity pattern -> drop, unconditional
      2. PSC in the R&D 'A' series:
           a. NAICS sector is manufacturing/goods/mining AND title doesn't
              name a BAA/research vehicle -> drop (likely equipment/materials
              procurement misclassified under an R&D PSC)
           b. otherwise                                 -> keep, clean
      3. Title names a BAA/research vehicle             -> keep, flagged (PSC disagrees)
      4. otherwise                                       -> drop, logged to sam_skipped.json
    """
    psc = (hit.get("classificationCode") or "").strip().upper()
    naics = (hit.get("naicsCode") or "").strip()
    naics_sector = naics[:2]
    title = hit.get("title") or ""
    flags = []

    if JUNK_TITLE_RE.search(title):
        return False, f"title matches known non-opportunity pattern (PSC {psc or 'none'})", flags

    is_baa_title = bool(BAA_TITLE_RE.search(title))

    if psc.startswith(PSC_RESEARCH_PREFIX):
        if naics_sector in NAICS_NON_RESEARCH_SECTORS and not is_baa_title:
            return False, (f"PSC {psc} is R&D series but NAICS {naics} (sector {naics_sector}) "
                            f"is manufacturing/goods/mining -- likely procurement misclassified "
                            f"under an R&D PSC"), flags
        return True, f"PSC {psc} (R&D series)", flags

    if is_baa_title:
        flags.append(
            f"KEPT ON TITLE SIGNAL: names a research funding vehicle but PSC is "
            f"'{psc or '(none)'}', not the R&D 'A' series -- verify this is a real "
            f"funding opportunity and not a procurement using similar wording"
        )
        return True, f"title signal (PSC {psc or 'none'})", flags

    return False, f"PSC {psc or '(none)'} not R&D and no BAA title signal", flags


def parse_search_payload(raw, org_code):
    """Parse and validate SAM search responses for a single org."""
    if not raw:
        return {"status": "error", "message": f"{org_code}: empty response", "hits": [], "total_records": 0}
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as exc:
        return {"status": "error", "message": f"{org_code}: malformed JSON ({exc})", "hits": [], "total_records": 0}
    if not isinstance(data, dict):
        return {"status": "error", "message": f"{org_code}: unexpected payload type {type(data).__name__}", "hits": [], "total_records": 0}
    hits = data.get("opportunitiesData") or []
    if not isinstance(hits, list):
        return {"status": "error", "message": f"{org_code}: opportunitiesData was not a list", "hits": [], "total_records": 0}
    total_records = data.get("totalRecords", len(hits))
    if total_records and not hits:
        return {"status": "warning", "message": f"{org_code}: totalRecords={total_records} but 0 hits returned", "hits": [], "total_records": total_records}
    return {"status": "ok", "message": "", "hits": hits, "total_records": total_records}


def dedupe_hits(hits):
    """Remove repeated notices while preserving the first-seen version."""
    seen, deduped = set(), []
    for hit in hits:
        if not isinstance(hit, dict):
            continue
        notice_id = (hit.get("noticeId") or hit.get("solicitationNumber") or "").strip()
        key = notice_id.lower() if notice_id else f"{hit.get('title','')}::{hit.get('uiLink','')}"
        if key in seen:
            continue
        seen.add(key)
        deduped.append(hit)
    return deduped


def _get(url, label, binary=False):
    """Single authenticated GET. Increments the call counter. No retry on
    HTTP errors (a 429/403 won't fix itself and each retry is another call);
    only transient network errors get one retry."""
    global _calls_made
    if not API_KEY:
        print("ERROR: SAM_API_KEY environment variable is not set.", file=sys.stderr)
        sys.exit(1)
    sep = "&" if "?" in url else "?"
    full = url if "api_key=" in url else f"{url}{sep}api_key={API_KEY}"
    _calls_made += 1
    print(f"  [call {_calls_made}] {label}: {_mask(full)}", file=sys.stderr)
    try:
        with urllib.request.urlopen(full, timeout=30) as r:
            return r.read() if binary else r.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        body = e.read().decode(errors="replace")[:500]
        print(f"    HTTP {e.code} on {label}: {body}", file=sys.stderr)
        return None
    except urllib.error.URLError as e:
        print(f"    network error on {label}: {e} -- one retry", file=sys.stderr)
        time.sleep(3)
        try:
            with urllib.request.urlopen(full, timeout=30) as r:
                return r.read() if binary else r.read().decode("utf-8", errors="replace")
        except (urllib.error.HTTPError, urllib.error.URLError) as e2:
            print(f"    retry failed on {label}: {e2}", file=sys.stderr)
            return None


def cache_path(cache_dir, stamp, org_code):
    return os.path.join(cache_dir, f"sam_raw_{stamp}_{org_code}.json")


def search_org(org_code, posted_from, posted_to, cache_dir, stamp):
    """One search call per org. Raw response is cached to disk before parsing
    so filter tuning can be replayed later at zero API cost."""
    q = {
        "organizationCode": org_code,
        "postedFrom": posted_from,
        "postedTo": posted_to,
        "limit": SEARCH_LIMIT,
        "api_key": API_KEY,
    }
    url = f"{SEARCH_URL}?{urllib.parse.urlencode(q)}"
    url += "&" + "&".join(f"ptype={p}" for p in PTYPES)
    raw = _get(url, f"search org={org_code}")
    if raw is None:
        return []

    os.makedirs(cache_dir, exist_ok=True)
    try:
        with open(cache_path(cache_dir, stamp, org_code), "w", encoding="utf-8") as f:
            f.write(raw)
    except OSError as e:
        print(f"    WARNING: could not cache raw response: {e}", file=sys.stderr)

    return _hits_from_raw(raw, org_code)


def load_cached_org(cache_dir, stamp, org_code):
    path = cache_path(cache_dir, stamp, org_code)
    if not os.path.exists(path):
        print(f"    no cache file for {org_code} at {path}", file=sys.stderr)
        return []
    with open(path, encoding="utf-8") as f:
        return _hits_from_raw(f.read(), org_code)


def _hits_from_raw(raw, org_code):
    payload = parse_search_payload(raw, org_code)
    if payload["status"] == "error":
        print(f"    {payload['message']}", file=sys.stderr)
        return []
    if payload["status"] == "warning":
        print(f"    {payload['message']}", file=sys.stderr)
    hits = dedupe_hits(payload["hits"])
    if payload["total_records"] > len(payload["hits"]):
        print(f"    WARNING: {org_code} totalRecords={payload['total_records']} but only "
              f"{len(payload['hits'])} returned -- results TRUNCATED, records were lost. "
              f"Narrow the date window or add offset paging.", file=sys.stderr)
    return hits


def strip_html(text):
    return normalize_encoding(re.sub(r"<[^>]+>", " ", text or ""))


def extract_eligibility_snippets(text, max_chars=1500):
    """Return (snippet_text, higher_ed_verdict). The verdict is a flag for a
    human to check, not an eligibility determination."""
    sentences = [s.strip() for s in re.split(r"(?<=[.!?])\s+", text or "") if len(s.strip()) > 15]
    higher_ed_hits = [s for s in sentences if any(k in s.lower() for k in HIGHER_ED_KEYWORDS)]
    set_aside_hits = [s for s in sentences if any(k in s.lower() for k in SET_ASIDE_KEYWORDS)]
    other_hits = [s for s in sentences
                  if any(k in s.lower() for k in ELIGIBILITY_KEYWORDS)
                  and s not in higher_ed_hits and s not in set_aside_hits]
    if set_aside_hits:
        verdict = "SMALL-BUSINESS SET-ASIDE LANGUAGE FOUND -- likely excludes higher-ed, verify"
    elif higher_ed_hits:
        verdict = "HIGHER-ED LANGUAGE FOUND -- verify scope/restrictions before assuming eligible"
    else:
        verdict = ""
    snippet = " ".join(higher_ed_hits + set_aside_hits + other_hits)[:max_chars]
    return snippet, verdict


def extract_pdf_text(data):
    import pdfplumber
    text = []
    with pdfplumber.open(BytesIO(data)) as pdf:
        for page in pdf.pages[:30]:
            t = page.extract_text()
            if t:
                text.append(t)
    return "\n".join(text)


def extract_docx_text(data):
    import docx
    return "\n".join(p.text for p in docx.Document(BytesIO(data)).paragraphs)


def fetch_detail_text(hit, max_attachments):
    """Fetch description + up to max_attachments attachments for one hit.
    NOTE: the search response's 'description' field is a URL, not text -- each
    one is a separate authenticated call. Returns (description_text, elig_text,
    notes).

    Previously this fetched the full text just to mine it for eligibility
    keywords and then discarded it -- description_text (rec["description"])
    stayed "" for every record, which is why Summary was empty downstream even
    for the one record per run that got a detail fetch at all. Now the fetched
    text (capped at DESCRIPTION_MAX_CHARS) is returned as raw material for the
    Claude summary-rewrite pass, same as every other collector's description
    field."""
    chunks, notes = [], []

    desc_url = hit.get("description") or ""
    if desc_url:
        raw = _get(desc_url, f"description {hit.get('noticeId','?')}")
        if raw:
            try:
                body = json.loads(raw).get("description", raw)
            except json.JSONDecodeError:
                body = raw
            if isinstance(body, str) and body.strip():
                chunks.append(strip_html(body))
            else:
                notes.append("description payload empty or not text")
        else:
            notes.append("description fetch failed")
    else:
        notes.append("no description URL available")

    links = hit.get("resourceLinks") or []
    if len(links) > max_attachments:
        notes.append(f"{len(links) - max_attachments} attachment(s) skipped (call budget)")
    for link in links[:max_attachments]:
        data = _get(link, f"attachment {hit.get('noticeId','?')}", binary=True)
        if not data:
            notes.append("attachment fetch failed")
            continue
        try:
            if data[:4] == b"%PDF":
                text = extract_pdf_text(data)
                if text.strip():
                    chunks.append(text)
                else:
                    notes.append("attachment PDF parsed but produced no text")
            elif data[:2] == b"PK":
                text = extract_docx_text(data)
                if text.strip():
                    chunks.append(text)
                else:
                    notes.append("attachment DOCX parsed but produced no text")
            else:
                notes.append("attachment not PDF/DOCX -- skipped parsing")
        except Exception as e:
            notes.append(f"attachment parse error: {e}")

    full_text = "\n".join(c for c in chunks if c)
    elig, verdict = extract_eligibility_snippets(full_text)
    if verdict:
        notes.append(verdict)
    elif full_text:
        notes.append("fetched description/attachment text but no higher-ed or eligibility keywords found -- verify manually")

    description = full_text[:DESCRIPTION_MAX_CHARS]
    if len(full_text) > DESCRIPTION_MAX_CHARS:
        notes.append(f"description truncated to {DESCRIPTION_MAX_CHARS} of {len(full_text)} fetched chars")
    return description, elig, notes


def normalize(hit, agency_name, keep_reason, relevance_flags):
    rec = blank_record()
    notice_id = hit.get("noticeId") or ""
    title = hit.get("title") or ""
    close_raw = hit.get("responseDeadLine") or hit.get("responseDate") or ""
    disp, status, days, multi = parse_sam_deadline(close_raw)

    psc = (hit.get("classificationCode") or "").strip().upper()
    naics = (hit.get("naicsCode") or "").strip()
    set_aside_code = (hit.get("typeOfSetAside") or "").strip()
    set_aside_desc = (hit.get("typeOfSetAsideDescription") or "").strip()

    rec.update({
        "source": "sam_baa",
        "opp_id": notice_id,
        "opp_number": hit.get("solicitationNumber") or notice_id,
        "title": normalize_encoding(title),
        "agency_code": hit.get("fullParentPathCode") or "",
        "agency_name": agency_name,
        "open_date": iso(None),
        "post_date": hit.get("postedDate") or "",
        "close_date": disp,
        "close_date_raw": close_raw,
        "status": hit.get("type") or hit.get("baseType") or "posted",
        "url": public_sam_url(hit.get("uiLink") or (f"https://sam.gov/opp/{notice_id}/view" if notice_id else "")),
        "description": "",
        "eligibility_desc": "",
        "award_amount": "",
        "cost_sharing": "",
        "type_hint": detect_type(title, hit.get("solicitationNumber") or ""),
    })

    # Provenance fields -- non-canonical, prefixed so merge_and_assess ignores
    # them but they survive into the JSON for auditing why a record was kept.
    rec["sam_psc"] = psc
    rec["sam_naics"] = naics
    rec["sam_set_aside"] = set_aside_desc or set_aside_code
    rec["sam_sub_tier"] = hit.get("subTier") or ""
    rec["sam_office"] = hit.get("office") or ""
    rec["sam_keep_reason"] = keep_reason

    rec["needs_review"].append(
        "SAM.gov contract opportunity -- no structured eligibility field, verify against source docs")
    rec["needs_review"].extend(relevance_flags)

    if set_aside_code or set_aside_desc:
        rec["needs_review"].append(
            f"SET-ASIDE: {set_aside_desc or set_aside_code} -- UNR is not a small business "
            f"concern for federal contracting purposes; likely ineligible, verify before including")
    if multi:
        rec["needs_review"].append(multi)
    if status == "soon":
        rec["needs_review"].append(f"deadline tight ({days}d) -- BAAs are frequently amended/extended, verify current date")
    if status in ("rolling", "unknown"):
        rec["needs_review"].append("no firm deadline - verify (rolling/unknown)")

    rec["_dl_status"] = status
    rec["_days_out"] = days
    return rec


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="sam_opportunities.json")
    ap.add_argument("--skipped-out", default="sam_skipped.json")
    ap.add_argument("--cache-dir", default=CACHE_DIR)
    ap.add_argument("--from-cache", default="",
                    help="replay cached raw responses for a date (YYYY-MM-DD). ZERO API calls.")
    ap.add_argument("--keep-all", action="store_true",
                    help="disable the PSC relevance gate (debugging only -- restores v2 noise)")
    ap.add_argument("--no-detail", action="store_true", help="search only, skip description/attachment fetch")
    ap.add_argument("--max-detail", type=int, default=DEFAULT_MAX_DETAIL)
    ap.add_argument("--max-attachments", type=int, default=1)
    ap.add_argument("--detail-ids", default="")
    args = ap.parse_args()

    from_cache = args.from_cache.strip()
    if not from_cache and not API_KEY:
        print("ERROR: SAM_API_KEY environment variable is not set.", file=sys.stderr)
        sys.exit(1)

    today = datetime.date.today()
    stamp = from_cache or today.strftime("%Y-%m-%d")
    posted_from = (today - datetime.timedelta(days=POSTED_WITHIN_DAYS)).strftime("%m/%d/%Y")
    posted_to = today.strftime("%m/%d/%Y")

    if from_cache:
        print(f"SAM.gov: REPLAYING cache {args.cache_dir} for {stamp} -- 0 API calls", file=sys.stderr)
    else:
        print(f"SAM.gov: {len(ORGS)} orgs, ptype {PTYPES}, window {posted_from}-{posted_to}", file=sys.stderr)
        print(f"Budget: {len(ORGS)} search call(s) + up to {args.max_detail} detail fetch(es).", file=sys.stderr)

    all_hits, agency_summaries = [], []
    for code, name in ORGS:
        hits = (load_cached_org(args.cache_dir, stamp, code) if from_cache
                else search_org(code, posted_from, posted_to, args.cache_dir, stamp))
        print(f"  {name} ({code}): {len(hits)} hits", file=sys.stderr)
        agency_summaries.append((name, len(hits)))
        for h in hits:
            all_hits.append((h, name))

    # ── Relevance gate ──
    relevant, skipped = [], []
    for hit, name in all_hits:
        keep, reason, flags = relevance(hit)
        if keep or args.keep_all:
            relevant.append((hit, name, reason if keep else "keep-all override", flags))
        else:
            skipped.append({
                "title": hit.get("title", ""),
                "agency": name,
                "psc": (hit.get("classificationCode") or ""),
                "naics": (hit.get("naicsCode") or ""),
                "type": hit.get("type", ""),
                "url": hit.get("uiLink", ""),
                "reason": reason,
            })
    print(f"\nRelevance gate: kept {len(relevant)} of {len(all_hits)} "
          f"({len(skipped)} dropped -> {args.skipped_out})", file=sys.stderr)

    records = [normalize(h, n, r, f) for h, n, r, f in relevant]

    kept, dropped_leadtime = [], 0
    for rec in records:
        status, days = rec.get("_dl_status"), rec.get("_days_out")
        if status == "passed" or (days is not None and days < MIN_LEAD_DAYS):
            dropped_leadtime += 1
            continue
        kept.append(rec)

    forced_ids = {s.strip() for s in args.detail_ids.split(",") if s.strip()}
    kept.sort(key=lambda r: (r.get("opp_id") not in forced_ids,
                             r.get("_days_out") if r.get("_days_out") is not None else 9999))

    if not args.no_detail and not from_cache:
        hit_by_id = {h.get("noticeId"): h for h, _, _, _ in relevant}
        for rec in kept[:args.max_detail]:
            hit = hit_by_id.get(rec["opp_id"])
            if not hit:
                continue
            print(f"Fetching detail for {rec['opp_number']} ({rec['title'][:60]})...", file=sys.stderr)
            description, elig, notes = fetch_detail_text(hit, max_attachments=args.max_attachments)
            rec["description"] = description
            rec["eligibility_desc"] = elig
            rec["needs_review"].extend(notes)
            time.sleep(0.5)
    elif from_cache:
        print("--from-cache: skipping detail fetch (would cost API calls).", file=sys.stderr)

    for rec in kept:
        rec.pop("_dl_status", None)
        rec.pop("_days_out", None)

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(kept, f, indent=2, ensure_ascii=False)
    with open(args.skipped_out, "w", encoding="utf-8") as f:
        json.dump(skipped, f, indent=2, ensure_ascii=False)

    print(f"\nAgency hit counts: {', '.join(f'{n}: {c}' for n, c in agency_summaries)}", file=sys.stderr)
    print(f"Kept {len(kept)} | relevance-dropped: {len(skipped)} | lead-time/passed: {dropped_leadtime}", file=sys.stderr)
    print(f"Total API calls this run: {_calls_made}", file=sys.stderr)
    print(f"Wrote -> {args.out}  (skips -> {args.skipped_out})", file=sys.stderr)
    if not from_cache:
        print(f"Raw responses cached in {args.cache_dir}/ -- replay with "
              f"--from-cache {stamp} at zero API cost.", file=sys.stderr)


if __name__ == "__main__":
    main()
