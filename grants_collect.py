#!/usr/bin/env python3
"""
grants_collect.py  (v2)
=======================
Federal collector. Hits the public Grants.gov REST API (no key required):

    search2          -> opportunities matching the saved filter
    fetchOpportunity -> full detail (description, eligibility, award, dates)

RUN ON YOUR LAPTOP (needs outbound HTTPS to api.grants.gov).

    python3 grants_collect.py --out federal_opportunities.json
    python3 grants_collect.py --no-detail        # fast preview, search only
"""
import json, os, re, urllib.request, urllib.error, datetime, time, argparse, sys
from grants_common import (normalize_encoding, detect_type, assess_deadline,
                           blank_record, iso)

API = "https://api.grants.gov/v1/api"

# ── Saved filter ─────────────────────────────────────────────────────────────
# Grants.gov requires exact sub-agency codes here; a parent code like "DOD"
# matches nothing. Any research arm not listed is invisible to this collector
# no matter how well it scores on eligibility or keyword -- that is how the
# Air Force YIP (FA955026S0003, DOD-AFOSR) was missed in Sept 2026.
#
# Army note: there is no "DOD-ARO" or "DOD-ARL" code. Both the Army Research
# Office and DEVCOM Army Research Laboratory post under DOD-AMC (Army Materiel
# Command) with W911NF-* numbers -- e.g. W911NF25S0007 (ARO Staff Research
# Program), W911NF-23-S-0001 (ARL BAA), W911NF26S0095 (DURIP). DEVCOM/Natick
# solicitations come through DOD-AMC-ACCAPGN. Use those codes, not ARO/ARL.
AGENCIES = [
    "USDA-NIFA","USDA-NRCS","DOD-DARPA-BTO","DOD-DARPA-I2O","DOD-DARPA-IPTO",
    "DOD-DARPA-DSO","DOD-AMRAA","DOD-ONR-NRL","DOD-ONR","DOD-AFOSR","DOD-AFRL",
    "DOD-DTRA","DOD-AMC","DOD-AMC-ACCAPGN","ED","HUD","USDOJ-OJP-NIJ","DOL-OPE",
    "EPA","NEA","NEH","NSF","SBA","HHS-NIH11","NASA",
    # "DOE", "DOI" and "DOT" used to sit here and matched *nothing*: verified
    # 2026-09-09 against all 1,003 live postings, each returned 0 hits as a
    # filter while their sub-agency codes returned plenty. DOE, DOI and DOT
    # funding was therefore invisible to this collector from the initial
    # commit onward -- that is why DE-FOA-0003647 (ASPECT, DOE-GFO, posted
    # 09/04/2026) had to be added to the email by hand.
    "DOE-GFO","DOE-ID","DOE-NETL","DOE-ARPAE","DOE-01",
    "DOI-FWS","DOI-BLM","DOI-USGS1","DOI-BOR","DOI-NPS",
    "DOT-FHWA","DOT-FTA","DOT-NHTSA","DOT-FRA","DOT-FMCSA",
    # FAA codes ("DOT-FAA-FAA ARG", "DOT-FAA-FAA COE-AJFE", ...) contain a
    # literal space and return 0 hits when passed as a filter even though the
    # postings exist. Left out deliberately; revisit if FAA coverage matters.
]
ELIGIBILITIES = ["06", "07", "25", "99"]
KEYWORD = "research"

# A 7-day window assumed this collector runs every week without fail. It does
# not: a skipped or late week silently loses everything posted in the gap, with
# nothing in the output to say so. The five Institute of Education Sciences
# NOFOs posted 08/06/2026 (ED-GRANTS-080626-001..005, all closing 10/01/2026)
# were never collected for exactly this reason -- the next run was 08/20, two
# weeks later, so they aged out of a 7-day window while still open and
# applicable for another seven weeks.
#
# The window therefore widens well past the run cadence, and a ledger records
# what has already been emitted so the wider window does not re-surface the
# same opportunity every week. This mirrors nsf_rss_collect.py, which solved
# the same problem for NSF's bursty feed. An opportunity returns only if
# Grants.gov re-posts it (its postDate changes), so a genuine re-issue still
# reaches review.
POSTED_WITHIN_DAYS    = 30
MIN_LEAD_DAYS         = 10
SEEN_LEDGER           = "federal_seen.json"
LEDGER_RETENTION_DAYS = 365


def post(path, payload, retries=5):
    url = f"{API}/{path}"
    data = json.dumps(payload).encode()
    for attempt in range(retries):
        try:
            req = urllib.request.Request(
                url, data=data,
                headers={"Content-Type": "application/json"}, method="POST")
            with urllib.request.urlopen(req, timeout=30) as r:
                return json.loads(r.read().decode())
        except (urllib.error.HTTPError, urllib.error.URLError, ConnectionResetError) as e:
            if attempt == retries - 1:
                raise
            wait = 2.0 * (attempt + 1)
            print(f"  [retry {attempt+1}/{retries-1}] {e} — waiting {wait:.0f}s", file=sys.stderr)
            time.sleep(wait)
    return None


def load_ledger(path):
    """{key: {'first_seen': iso, 'last_post': iso, 'title': str}}. A missing or
    unreadable ledger is not fatal -- a bad ledger should cost a week of
    duplicate review rows, never the whole collection run."""
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError) as e:
        print(f"  WARNING: ignoring unreadable ledger {path}: {e}", file=sys.stderr)
        return {}


def save_ledger(path, ledger, today):
    cutoff = today - datetime.timedelta(days=LEDGER_RETENTION_DAYS)
    pruned = {k: v for k, v in ledger.items()
              if (v.get("last_post") or v.get("first_seen") or "") >= cutoff.isoformat()}
    with open(path, "w", encoding="utf-8") as f:
        json.dump(dict(sorted(pruned.items())), f, indent=2, ensure_ascii=False)
    return len(ledger) - len(pruned)


def ledger_key(hit):
    """Prefer the public number; fall back to the grants.gov id so a record with
    no number still gets tracked instead of colliding on the empty string."""
    return (hit.get("number") or "").strip() or "id:%s" % hit.get("id")


def parse_ggov_date(s):
    if not s:
        return None
    try:
        return datetime.datetime.strptime(s.strip(), "%m/%d/%Y").date()
    except ValueError:
        return None


def search_opportunities():
    hits, start = [], 0
    while True:
        payload = {
            "rows": 100, "startRecordNum": start, "keyword": KEYWORD,
            "oppStatuses": "posted",
            "eligibilities": "|".join(ELIGIBILITIES),
            "agencies": "|".join(AGENCIES),
        }
        out = post("search2", payload)
        data = out.get("data", {})
        batch = data.get("oppHits", [])
        hits.extend(batch)
        total = data.get("hitCount", 0)
        start += len(batch)
        if start >= total or not batch:
            break
    return hits


def fetch_detail(opp_id):
    out = post("fetchOpportunity", {"opportunityId": opp_id})
    return out.get("data", {}) if out else {}


_ARCHIVED_DESC_RE = re.compile(r"archiv|not\s+accept|no\s+longer\s+accept", re.I)


def archived_reason(detail, today):
    """search2 can report an opportunity as 'posted' (and so 'new') when the
    agency is actually archiving it -- e.g. NSF 24-594, posted 09/30/2026 with
    archiveDate 10/01/2026 and 'Funding Opportunity is Archived, Proposals are
    not accepted'. Only the fetchOpportunity detail says so. Returns a reason
    string if the record is archived or archiving, else None."""
    syn = detail.get("synopsis") or {}
    desc = syn.get("responseDateDesc") or ""
    if _ARCHIVED_DESC_RE.search(desc):
        return f"responseDateDesc: {desc.strip()[:80]}"
    raw = (syn.get("archiveDate") or "").strip()
    try:
        archive = datetime.datetime.strptime(raw[:12], "%b %d, %Y").date() if raw else None
    except ValueError:
        archive = None
    if archive and (archive - today).days < MIN_LEAD_DAYS:
        return f"archiveDate {archive.isoformat()}"
    return None


def pull_synopsis(detail):
    syn = detail.get("synopsis") or {}
    desc = (syn.get("synopsisDesc") or detail.get("synopsisDesc")
            or syn.get("description") or "")
    elig = (syn.get("applicantEligibilityDesc") or syn.get("applicantTypes")
            or detail.get("applicantEligibilityDesc") or "")
    if isinstance(elig, list):
        elig = "; ".join(str(e.get("description", e)) if isinstance(e, dict) else str(e)
                         for e in elig)
    close = (syn.get("responseDate") or detail.get("closeDate") or "")
    cost = syn.get("costSharing")
    cost = "Yes" if cost in (True, "true", "Yes", "Y") else ("No" if cost in (False, "false", "No", "N") else "")
    ceiling = syn.get("awardCeiling") or ""
    floor = syn.get("awardFloor") or ""
    amt = ""
    if ceiling or floor:
        amt = f"floor {floor}, ceiling {ceiling}".strip()
    return desc, elig, close, cost, amt


def normalize(hit, detail):
    rec = blank_record()
    desc, elig, close_detail, cost, amt = pull_synopsis(detail)
    opp_id = hit.get("id")
    number = hit.get("number") or ""
    close_raw = hit.get("closeDate") or close_detail or ""
    disp, status, days, multi = assess_deadline(close_raw)
    rec.update({
        "source": "federal",
        "opp_id": opp_id,
        "opp_number": number,
        "title": normalize_encoding(hit.get("title") or ""),
        "agency_code": hit.get("agencyCode") or "",
        "agency_name": hit.get("agencyName") or "",
        "open_date": iso(parse_ggov_date(hit.get("openDate"))),
        "post_date": iso(parse_ggov_date(hit.get("postDate") or hit.get("openDate"))),
        "close_date": disp,
        "close_date_raw": close_raw,
        "status": hit.get("oppStatus") or "posted",
        "url": f"https://grants.gov/search-results-detail/{opp_id}",
        "description": normalize_encoding(desc),
        "eligibility_desc": normalize_encoding(elig),
        "award_amount": amt,
        "cost_sharing": cost,
        "type_hint": detect_type(hit.get("title"), number),
    })
    if multi:
        rec["needs_review"].append(multi)
    if status == "soon":
        rec["needs_review"].append(f"deadline tight ({days}d)")
    if not desc:
        rec["needs_review"].append("no description from API - verify/fetch manually")
    rec["_dl_status"] = status
    rec["_days_out"] = days
    rec["_post_date_obj"] = parse_ggov_date(hit.get("postDate") or hit.get("openDate"))
    return rec


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="federal_opportunities.json")
    ap.add_argument("--no-detail", action="store_true")
    ap.add_argument("--days", type=int, default=POSTED_WITHIN_DAYS,
                    help=f"posting window in days (default {POSTED_WITHIN_DAYS}); "
                         "deliberately wider than the weekly cadence so a skipped "
                         "week does not lose opportunities")
    ap.add_argument("--seen", default=SEEN_LEDGER,
                    help=f"ledger of already-emitted opportunities (default "
                         f"{SEEN_LEDGER}); keep it committed so weekly runs do not "
                         "repeat themselves")
    ap.add_argument("--ignore-seen", action="store_true",
                    help="emit matching opportunities even if the ledger has them "
                         "(does not write the ledger)")
    args = ap.parse_args()

    print(f"Grants.gov: {len(AGENCIES)} agencies, elig {ELIGIBILITIES}, "
          f"keyword '{KEYWORD}'...", file=sys.stderr)
    hits = search_opportunities()
    print(f"  {len(hits)} total posted hits returned.", file=sys.stderr)

    today = datetime.date.today()
    ledger = {} if args.ignore_seen else load_ledger(args.seen)
    if ledger:
        print(f"  ledger: {len(ledger)} opportunity(ies) already emitted.", file=sys.stderr)
    kept, dropped_old, dropped_leadtime, dropped_seen, reposted = [], 0, 0, 0, 0
    dropped_archived = 0
    for i, h in enumerate(hits, 1):
        # Both of these are decided from the search hit alone. Filtering here
        # rather than after fetch_detail() is what makes the wider window
        # affordable: 30 days over ~660 hits would otherwise mean ~660 detail
        # calls per run to throw nearly all of them away.
        pd = parse_ggov_date(h.get("postDate") or h.get("openDate"))
        if pd is None or (today - pd).days > args.days:
            dropped_old += 1
            continue

        key = ledger_key(h)
        prior = ledger.get(key)
        if prior:
            if prior.get("last_post") == iso(pd):
                dropped_seen += 1
                continue
            # Same opportunity, new postDate: Grants.gov re-posted it, so it goes
            # back through review rather than being silently suppressed.
            reposted += 1
            print(f"  re-posted {key:22} {prior.get('last_post')} -> {iso(pd)}",
                  file=sys.stderr)

        detail = {} if args.no_detail else fetch_detail(h["id"])
        if not args.no_detail:
            time.sleep(0.25)
        why = archived_reason(detail, today)
        if why:
            dropped_archived += 1
            print(f"  archived {h.get('number') or h.get('id')}: {why}", file=sys.stderr)
            continue
        rec = normalize(h, detail)
        rec.pop("_post_date_obj", None)

        status = rec.pop("_dl_status", None)
        days = rec.pop("_days_out", None)
        if status == "passed" or (days is not None and days < MIN_LEAD_DAYS):
            dropped_leadtime += 1
            continue
        if status in ("rolling", "unknown"):
            rec["needs_review"].append("no firm deadline - verify (rolling/unknown)")

        kept.append(rec)
        ledger[key] = {"first_seen": (prior or {}).get("first_seen") or iso(today),
                       "last_post": iso(pd),
                       "title": (rec.get("title") or "")[:120]}
        print(f"  [{i}] kept {rec['opp_number']} ({rec['close_date']})", file=sys.stderr)

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(kept, f, indent=2, ensure_ascii=False)

    if args.ignore_seen:
        print("\n--ignore-seen: ledger NOT updated.", file=sys.stderr)
    else:
        expired = save_ledger(args.seen, ledger, today)
        print(f"\nLedger -> {args.seen} ({len(ledger) - expired} tracked"
              + (f", {expired} pruned past {LEDGER_RETENTION_DAYS}d" if expired else "")
              + ")", file=sys.stderr)
    print(f"Kept {len(kept)} | window {args.days}d | dropped not-recent: {dropped_old} "
          f"| already emitted: {dropped_seen} | lead-time/passed: {dropped_leadtime} "
          f"| archived: {dropped_archived}"
          + (f" | re-posted: {reposted}" if reposted else ""), file=sys.stderr)
    print(f"Wrote -> {args.out}", file=sys.stderr)


if __name__ == "__main__":
    main()
