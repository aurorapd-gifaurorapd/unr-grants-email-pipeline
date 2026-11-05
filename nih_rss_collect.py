#!/usr/bin/env python3
"""
nih_rss_collect.py  (v2 -- NIH Guide search, not RSS)
=====================================================
Collects NIH Guide notices (RFIs, webinars, policy notices, other notices)
and emits canonical records for merge_and_assess.py. NIH funding
opportunities (RFA/PAR/PA) are skipped by default because the Grants.gov
federal pull already has them; use --keep-nofos to include them.

Why v2: the RSS feed (newsfeed/fundingopps.xml) only holds the last few days
of postings. When the email skipped a week (10/22 -> 11/05) the feed had
already dropped two weeks of webinars and policy notices. v2 asks the NIH
Guide search (search.grants.nih.gov/guide/api/data) for a date range instead,
and keeps a seen-list (nih_seen.json) so a wide window never re-sends a
notice that already went out. The filename is unchanged so the documented
weekly command still works.

    python nih_rss_collect.py --out Grant_MM_DD_YYYY\\nih_opportunities.json

Options:
    --days N        look back N days (default 30; the seen-list stops repeats)
    --ignore-seen   don't skip or record seen notices (testing)
    --seed          record every notice in the window as already sent, write
                    no records (one-time setup, or after a manual catch-up)
    --keep-nofos    also emit RFA/PAR/PA funding opportunities
"""
import argparse, datetime, html, json, os, re, sys, time, urllib.error, urllib.parse, urllib.request
from grants_common import normalize_encoding, detect_type, assess_deadline, blank_record, iso

API = "https://search.grants.nih.gov/guide/api/data"
NOTICE_URL = "https://grants.nih.gov/grants/guide/notice-files/{}"
SEEN_LEDGER = "nih_seen.json"
POSTED_WITHIN_DAYS = 30
LEDGER_RETENTION_DAYS = 365
PER_PAGE = 100
UA = {"User-Agent": "UNR-grants-collector/1.0"}

RESP_RE = re.compile(r"Response Date:?\s*([A-Z][a-z]+ \d{1,2}, \d{4})", re.I)

# Administrative notices that change nothing a faculty reader acts on. They
# are logged to the skipped file, not silently dropped.
ADMIN_RE = re.compile(
    r"\bparticipation\b|\bwithdraw|\bcorrection\b|\brescind|"
    r"grants administration information|\bnot participating\b|"
    r"\breinstatement\b|\bexpiration\b", re.I)


def stop_run(problem, steps):
    print("\n" + "=" * 70, file=sys.stderr)
    print(f"STOPPED: {problem}", file=sys.stderr)
    print("No output file was written.", file=sys.stderr)
    print("\nWhat to do:", file=sys.stderr)
    for s in steps:
        print(f"  - {s}", file=sys.stderr)
    print("=" * 70, file=sys.stderr)
    sys.exit(2)


def http_get(url, timeout=60):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def search(date_from, date_to, types):
    """All Guide records released in [date_from, date_to], paged."""
    hits, start = [], 0
    rng = f"{date_from:%m%d%Y}-{date_to:%m%d%Y}"
    while True:
        q = {"perpage": PER_PAGE, "sort": "reldate:desc", "from": start, "type": types,
             "parentic": "all", "primaryic": "all", "activitycodes": "all", "doctype": "all",
             "parentfoa": "all", "daterange": rng, "clinicaltrials": "all", "fields": "all",
             "spons": "true", "query": ""}
        try:
            data = json.loads(http_get(f"{API}?{urllib.parse.urlencode(q)}"))
        except (urllib.error.URLError, ValueError) as e:
            stop_run(f"The NIH Guide search did not answer ({e}).",
                     ["Check your internet connection and run nih_rss_collect.py again.",
                      "If it keeps failing, NIH's search may be down; try again later today."])
        block = (data.get("data") or {}).get("hits") or {}
        page = block.get("hits") or []
        total = block.get("total")
        total = total.get("value") if isinstance(total, dict) else total
        hits.extend(h.get("_source") or {} for h in page)
        start += len(page)
        if not page or (total is not None and start >= int(total)):
            return hits


def notice_details(filename):
    """(response_date_text, purpose_text) from the notice page. Never fatal."""
    try:
        raw = http_get(NOTICE_URL.format(filename), timeout=30).decode("utf-8", "replace")
    except (urllib.error.URLError, OSError):
        return "", ""
    text = re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", raw)))
    m = RESP_RE.search(text)
    purpose = ""
    p = re.search(r"\bPurpose\b(.*?)(?:\bInquiries\b|$)", text)
    if p:
        purpose = p.group(1).strip()[:1500]
    return (m.group(1) if m else ""), purpose


_DATE = r"(?:(?:Mon|Tues|Wednes|Thurs|Fri|Satur|Sun)day,?\s+)?([A-Z][a-z]+ \d{1,2}, \d{4})"
_DATE_LABEL_RE = re.compile(r"\bDate:\s*" + _DATE)
_EVENT_SENT_RE = re.compile(r"[^.]*\b(?:webinar|session|office hours?|will (?:be )?(?:held|host))\b[^.]*")
_NOT_EVENT_RE = re.compile(r"no later than|\bby\b|deadline|submit|due", re.I)


def event_date(text):
    """The event's own date from a webinar notice, or "" if it isn't clear.
    Taking the first date in the text is wrong: NOT-OD-27-010 lists the
    question deadline (Nov 17) before the webinar (Nov 19)."""
    text = text or ""
    m = _DATE_LABEL_RE.search(text)
    if m:
        return m.group(1)
    for sent in _EVENT_SENT_RE.findall(text):
        if _NOT_EVENT_RE.search(sent):
            continue
        m = re.search(_DATE, sent)
        if m:
            return m.group(1)
    return ""


def clean_purpose(s):
    s = re.sub(r"<[^>]+>", " ", s or "")
    s = re.sub(r"^p(?=[A-Z])", "", s.strip())        # API strips tags but leaves the 'p'
    return normalize_encoding(re.sub(r"\s+", " ", s).strip())


def type_for(subject, title, number):
    subj = (subject or "").lower()
    if "request for information" in subj:
        return "RFI"
    if "webinar" in title.lower() or "office hours" in title.lower():
        return "EVENT"
    if "policy notice" in subj:
        return "POLICY"
    return detect_type(title, number)


def load_ledger(path):
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
    cutoff = (today - datetime.timedelta(days=LEDGER_RETENTION_DAYS)).isoformat()
    pruned = {k: v for k, v in ledger.items() if (v.get("released") or v.get("first_seen") or "") >= cutoff}
    with open(path, "w", encoding="utf-8") as f:
        json.dump(dict(sorted(pruned.items())), f, indent=2, ensure_ascii=False)
    return len(ledger) - len(pruned)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="nih_opportunities.json")
    ap.add_argument("--skipped-out", default=None,
                    help="where to log administrative notices that were skipped "
                         "(default: nih_skipped.json next to --out)")
    ap.add_argument("--days", type=int, default=POSTED_WITHIN_DAYS)
    ap.add_argument("--seen", default=SEEN_LEDGER)
    ap.add_argument("--ignore-seen", action="store_true")
    ap.add_argument("--seed", action="store_true")
    ap.add_argument("--keep-nofos", action="store_true")
    args = ap.parse_args()
    skipped_out = args.skipped_out or os.path.join(os.path.dirname(args.out) or ".", "nih_skipped.json")

    today = datetime.date.today()
    date_from = today - datetime.timedelta(days=args.days)
    types = "notices,active,activenosis" if args.keep_nofos else "notices"
    print(f"Searching NIH Guide {date_from:%b %d} - {today:%b %d, %Y} ({args.days} days)...", file=sys.stderr)
    docs = search(date_from, today, types)
    print(f"  {len(docs)} Guide records in window.", file=sys.stderr)

    ledger = {} if args.ignore_seen else load_ledger(args.seen)
    if not args.ignore_seen and not args.seed and not ledger:
        stop_run(f"No NIH seen-list found ({args.seen}).",
                 ["Without it, a 30-day search would re-send notices that already went out.",
                  "Run once:  python nih_rss_collect.py --seed",
                  "That marks everything up to today as already sent. Then run the normal command."])

    if args.seed:
        for d in docs:
            n = d.get("docnum") or ""
            if n:
                ledger.setdefault(n, {"first_seen": iso(today), "released": (d.get("reldate") or "")[:10],
                                      "title": (d.get("title") or "")[:120], "seeded": True})
        save_ledger(args.seen, ledger, today)
        print(f"Seeded {args.seen}: {len(ledger)} notices marked as already sent. No records written.",
              file=sys.stderr)
        return

    kept, skipped, already = [], [], 0
    for d in docs:
        number = (d.get("docnum") or "").strip()
        title = normalize_encoding(re.sub(r"\s+", " ", html.unescape(d.get("title") or "")).strip())
        released = (d.get("reldate") or "")[:10]
        if number in ledger:
            already += 1
            continue
        subject = d.get("noticeSubject") or ""
        if number.startswith("NOT-") and ADMIN_RE.search(title):
            skipped.append({"docnum": number, "title": title, "released": released, "reason": "administrative notice"})
            continue

        resp, page_purpose = notice_details(d.get("filename") or f"{number}.html")
        time.sleep(0.3)                                   # be polite to grants.nih.gov
        ttype = type_for(subject, title, number)
        if not resp and ttype == "EVENT":
            resp = event_date(page_purpose)
        if resp:
            disp, status, _days, _multi = assess_deadline(resp)
        else:
            disp, status = "rolling / not specified", "unknown"
        if status == "passed":
            skipped.append({"docnum": number, "title": title, "released": released, "reason": f"date passed ({disp})"})
            continue

        rec = blank_record()
        rec.update({
            "source": "nih_guide",
            "opp_id": number,
            "opp_number": number,
            "title": title,
            "agency_code": "NIH",
            "agency_name": "National Institutes of Health",
            "post_date": released,
            "open_date": released,
            "close_date": disp,
            "close_date_raw": resp,
            "status": "posted",
            "url": NOTICE_URL.format(d.get("filename") or f"{number}.html"),
            "description": clean_purpose(page_purpose or d.get("purpose")),
            "type_hint": ttype,
        })
        if subject:
            rec["needs_review"].append(f"NIH {subject.lower()}")
        if ttype == "EVENT":
            rec["needs_review"].append("webinar - put the date and Pacific time in Deadline")
        elif not resp:
            rec["needs_review"].append("no response date found - verify")
        kept.append(rec)
        ledger[number] = {"first_seen": iso(today), "released": released, "title": title[:120]}
        print(f"  kept {ttype:6} {number:16} {title[:55]}", file=sys.stderr)

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(kept, f, indent=2, ensure_ascii=False)
    with open(skipped_out, "w", encoding="utf-8") as f:
        json.dump(skipped, f, indent=2, ensure_ascii=False)

    if args.ignore_seen:
        print("\n--ignore-seen: seen-list NOT updated.", file=sys.stderr)
    else:
        pruned = save_ledger(args.seen, ledger, today)
        print(f"\nSeen-list -> {args.seen} ({len(ledger) - pruned} tracked)", file=sys.stderr)
    print(f"Kept {len(kept)} | window {args.days}d | already sent: {already} | "
          f"skipped (admin or passed): {len(skipped)} -> {skipped_out}", file=sys.stderr)
    print(f"Wrote -> {args.out}", file=sys.stderr)


if __name__ == "__main__":
    main()
