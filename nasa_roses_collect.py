#!/usr/bin/env python3
"""
nasa_roses_collect.py
=====================
Collects new NASA ROSES opportunities and announcements from the ROSES blog
and emits canonical records for merge_and_assess.py.

Why: ROSES is one big NASA solicitation (NNH25ZDA001N for ROSES-2025) that
adds program elements all year through numbered amendments. Grants.gov lists
the umbrella solicitation once, so none of the other collectors see a new
element. The 10/22/2026 email had no NASA items at all, although Amendments
77-80 (four new opportunities) were posted in that window.

Where the data comes from: the ROSES blog posts are WordPress "topic" pages
whose parent is the year's ROSES page (roses-2025, ...). The site's public
WordPress API returns them with their post date, title and full text, so no
HTML scraping is needed. The ROSES year is found automatically, so ROSES-2027
(expected by January 2027) needs no code change.

NSPIRES is NOT read: it refuses scripted requests. The blog posts state the
due dates in fixed wording ("Mandatory Notices of Intent (NOIs) are due by
...", "Step-1 proposals are due ..."), and link to the NSPIRES page, which is
used as the email link when the post gives one.

    python nasa_roses_collect.py --out Grant_MM_DD_YYYY\\nasa_opportunities.json

Options:
    --days N        look back N days (default 30; the seen-list stops repeats)
    --ignore-seen   don't skip or record seen posts (testing)
    --seed          record every post in the window as already sent, write no
                    records (one-time setup)
    --mark-sent IDS record only these post ids as already sent (comma list),
                    e.g. posts added to the email by hand
"""
import argparse, datetime, html, json, os, re, sys, urllib.error, urllib.parse, urllib.request
from grants_common import normalize_encoding, assess_deadline, blank_record, iso

API = "https://science.nasa.gov/wp-json/wp/v2/topic"
ROSES_HOME_ID = 60348            # parent page of roses-2025, roses-2024, ...
SEEN_LEDGER = "nasa_seen.json"
POSTED_WITHIN_DAYS = 30
LEDGER_RETENTION_DAYS = 365
YEARS_TO_WATCH = 2               # the new ROSES year and the one before it overlap for months
UA = {"User-Agent": "UNR-grants-collector/1.0"}

ELEMENT_RE = re.compile(r"\b([A-G]\.\d{1,2})\b")
SOLNUM_RE = re.compile(r"solNum=(NNH\d{2}[A-Z0-9]+-[A-Z0-9]+)")
DATE = r"([A-Z][a-z]+\.?\s*\d{1,2},\s*\d{4})"
# "Mandatory Notices of Intent (NOIs) are due by November 16, 2026"
# "Step-1 proposals are due by November 20, 2026, and Step-2 proposals are due February 5, 2027"
# "Neither NOIs nor Step-1 Proposals are requested. Proposals are due February 3, 2027"
DUE_RE = re.compile(
    r"(mandatory\s+notices?\s+of\s+intent(?:\s*\(NOIs?\))?|required\s+notices?\s+of\s+intent(?:\s*\(NOIs?\))?|"
    r"notices?\s+of\s+intent(?:\s*\(NOIs?\))?|NOIs?|step-1\s+proposals?|step-2\s+proposals?|proposals?)"
    r"\s+(?:are|is)\s+(?:now\s+)?due\s+(?:by\s+|on\s+)?(?:[A-Z][a-z]+day,?\s+)?" + DATE,
    re.I)
# "The due date for Phase-1 Proposals (submitted via ARK RPS) for all three is January 28, 2027"
DUE2_RE = re.compile(
    r"due dates?\s+for\s+((?:phase|step)-\d\s+proposals?|notices?\s+of\s+intent|proposals?)"
    r"(?:\s*\([^)]*\))?(?:\s+for\s+[^.]*?)?\s+(?:is|are)\s+(?:[A-Z][a-z]+day,?\s+)?" + DATE, re.I)
SKIP_TITLE_RE = re.compile(r"\bselect(?:s|ions?|ed)\b", re.I)
NEW_RE = re.compile(r"new opportunit|second opportunit|new program element", re.I)
DATE_CHANGE_RE = re.compile(r"due date|delayed|postponed|extended|deadline", re.I)


def stop_run(problem, steps):
    print("\n" + "=" * 70, file=sys.stderr)
    print(f"STOPPED: {problem}", file=sys.stderr)
    print("No output file was written.", file=sys.stderr)
    print("\nWhat to do:", file=sys.stderr)
    for s in steps:
        print(f"  - {s}", file=sys.stderr)
    print("=" * 70, file=sys.stderr)
    sys.exit(2)


def api_get(params):
    url = f"{API}?{urllib.parse.urlencode(params)}"
    try:
        with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=60) as r:
            return json.loads(r.read())
    except (urllib.error.URLError, ValueError) as e:
        stop_run(f"NASA's science.nasa.gov site did not answer ({e}).",
                 ["Check your internet connection and run nasa_roses_collect.py again.",
                  "If it keeps failing, the site may be down; try again later today."])


def roses_years():
    """[(year, page_id)] for the newest ROSES year pages, newest first."""
    pages = api_get({"parent": ROSES_HOME_ID, "per_page": 100, "_fields": "id,slug"})
    years = sorted(((int(m.group(1)), p["id"]) for p in pages
                    for m in [re.fullmatch(r"roses-(\d{4})", p.get("slug", ""))] if m), reverse=True)
    if not years:
        stop_run("Couldn't find the ROSES year pages on science.nasa.gov.",
                 ["NASA may have reorganized the ROSES pages. Tell whoever maintains the pipeline.",
                  "Meanwhile, check https://science.nasa.gov/researchers/solicitations/ by hand."])
    return years[:YEARS_TO_WATCH]


def posts_since(page_ids, since):
    out, page = [], 1
    while True:
        batch = api_get({"parent": ",".join(str(i) for i in page_ids), "after": f"{since}T00:00:00",
                         "per_page": 100, "page": page, "orderby": "date", "order": "desc",
                         "_fields": "id,date,link,title,content,parent"})
        out.extend(batch)
        if len(batch) < 100:
            return out
        page += 1


def body_text(rendered):
    t = re.sub(r"<[^>]+>", " ", rendered or "")
    t = re.sub(r"\s+", " ", html.unescape(t)).strip()
    # The rendered page starts with the ROSES side menu and ends with
    # "Explore More" / "Keep Exploring" teasers for unrelated articles.
    if "Scientific Information Policy" in t:
        t = t.split("Scientific Information Policy", 1)[1]
    t = re.split(r"\b(?:Explore More|Keep Exploring|Discover More Topics)\b", t)[0]
    return normalize_encoding(t.strip("  -"))


def parse_date(s):
    s = re.sub(r"\.", "", s)
    s = re.sub(r"([A-Za-z])(\d)", r"\1 \2", s)          # "October8, 2026"
    for fmt in ("%B %d, %Y", "%b %d, %Y"):
        try:
            return datetime.datetime.strptime(re.sub(r"\s+", " ", s), fmt).date()
        except ValueError:
            pass
    return None


def due_dates(text):
    """[(label, date, required)] in the order stated."""
    found = []
    matches = sorted(list(DUE_RE.finditer(text)) + list(DUE2_RE.finditer(text)), key=lambda m: m.start())
    for m in matches:
        label = re.sub(r"\s+", " ", m.group(1)).strip()
        d = parse_date(m.group(2))
        if not d:
            continue
        low = label.lower()
        if "step-1" in low:
            kind, required = "Step-1 proposals", True
        elif "step-2" in low:
            kind, required = "Step-2 proposals", True
        elif "phase-1" in low:
            kind, required = "Phase-1 proposals", True
        elif "intent" in low or low.startswith("noi"):
            required = "mandatory" in low or "required" in low
            kind = "Mandatory NOIs" if required else "NOIs (optional)"
        else:
            kind, required = "Proposals", True
        if (kind, d) not in [(k, x) for k, x, _ in found]:
            found.append((kind, d, required))
    return found


def fmt(d):
    return d.strftime("%B %d, %Y")


def classify(title, text, element):
    if NEW_RE.search(title) or NEW_RE.search(text[:2000]) and element:
        return "new"
    if element and DATE_CHANGE_RE.search(title):
        return "date change"
    if element:
        return "change"
    return "announcement"


def clean_title(title, element, kind):
    t = re.sub(r"^Amendment\s+\d+\s*:\s*", "", title).strip()
    if kind == "new":
        t = re.sub(r"^New Opportunity\s*:\s*", "", t, flags=re.I).strip()
        if element and element not in t:
            t = f"{element} {t}"
    return t


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
    pruned = {k: v for k, v in ledger.items() if (v.get("posted") or v.get("first_seen") or "") >= cutoff}
    with open(path, "w", encoding="utf-8") as f:
        json.dump(dict(sorted(pruned.items())), f, indent=2, ensure_ascii=False)
    return len(ledger) - len(pruned)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="nasa_opportunities.json")
    ap.add_argument("--skipped-out", default=None,
                    help="where to log posts that were skipped (default: nasa_skipped.json next to --out)")
    ap.add_argument("--days", type=int, default=POSTED_WITHIN_DAYS)
    ap.add_argument("--seen", default=SEEN_LEDGER)
    ap.add_argument("--ignore-seen", action="store_true")
    ap.add_argument("--seed", action="store_true")
    ap.add_argument("--mark-sent", default="",
                    help="comma list of post ids to record as already sent (e.g. items added to the email by "
                         "hand); writes no records")
    args = ap.parse_args()
    skipped_out = args.skipped_out or os.path.join(os.path.dirname(args.out) or ".", "nasa_skipped.json")

    today = datetime.date.today()
    since = today - datetime.timedelta(days=args.days)
    years = roses_years()
    print(f"ROSES years watched: {', '.join(str(y) for y, _ in years)}", file=sys.stderr)
    posts = posts_since([pid for _, pid in years], since)
    print(f"  {len(posts)} ROSES blog posts since {since:%b %d, %Y}.", file=sys.stderr)

    ledger = {} if args.ignore_seen else load_ledger(args.seen)
    if args.mark_sent:
        ids = {i.strip() for i in args.mark_sent.split(",") if i.strip()}
        by_id = {str(p["id"]): p for p in posts}
        missing = ids - set(by_id)
        if missing:
            stop_run(f"Post id(s) not found in the last {args.days} days: {', '.join(sorted(missing))}",
                     ["Check the ids (the number in the post's ?p= short link), or widen --days."])
        for i in sorted(ids):
            p = by_id[i]
            ledger[i] = {"first_seen": iso(today), "posted": p["date"][:10],
                         "title": html.unescape(p["title"]["rendered"])[:120], "marked_sent": True}
            print(f"  marked sent: {i} {ledger[i]['title'][:70]}", file=sys.stderr)
        save_ledger(args.seen, ledger, today)
        print(f"Saved {args.seen}: {len(ledger)} posts tracked. No records written.", file=sys.stderr)
        return
    if not args.ignore_seen and not args.seed and not ledger:
        stop_run(f"No NASA seen-list found ({args.seen}).",
                 ["Without it, a 30-day search would re-send posts that already went out.",
                  "Run once:  python nasa_roses_collect.py --seed",
                  "That marks every post up to today as already sent. Then run the normal command."])

    year_of = {pid: y for y, pid in years}
    if args.seed:
        for p in posts:
            ledger.setdefault(str(p["id"]), {"first_seen": iso(today), "posted": p["date"][:10],
                                             "title": html.unescape(p["title"]["rendered"])[:120], "seeded": True})
        save_ledger(args.seen, ledger, today)
        print(f"Seeded {args.seen}: {len(ledger)} posts marked as already sent. No records written.", file=sys.stderr)
        return

    kept, skipped, already = [], [], 0
    for p in posts:
        key = str(p["id"])
        title = normalize_encoding(html.unescape(p["title"]["rendered"]).strip())
        posted = p["date"][:10]
        if key in ledger:
            already += 1
            continue
        if SKIP_TITLE_RE.search(title):
            skipped.append({"post": key, "title": title, "posted": posted, "reason": "award selections, not an opportunity"})
            continue

        rendered = p["content"]["rendered"]
        text = body_text(rendered)
        em = ELEMENT_RE.search(title) or ELEMENT_RE.search(text[:600])
        element = em.group(1) if em else ""
        kind = classify(title, text, element)
        sols = list(dict.fromkeys(SOLNUM_RE.findall(html.unescape(rendered))))
        dates = due_dates(text)
        future = [(k, d, r) for k, d, r in dates if d > today]     # due today is too late for the email
        if dates and not future:
            skipped.append({"post": key, "title": title, "posted": posted,
                            "reason": f"every date has passed ({fmt(max(d for _k, d, _r in dates))})"})
            continue

        # Deadline rule: the earliest REQUIRED future date (Step-1, mandatory
        # NOI, or the proposal itself). Optional NOIs only go in the sequence.
        required = [d for _k, d, r in future if r]
        deadline = min(required) if required else (min(d for _k, d, _r in future) if future else None)
        sequence = "; ".join(f"{k} due {fmt(d)}" for k, d, _r in dates)

        yy = str(year_of.get(p.get("parent"), today.year))[-2:]
        number = sols[0] if (kind == "new" and len(sols) == 1) else (
            f"ROSES-{yy} {element}" if kind == "new" and element else "")
        rec = blank_record()
        rec.update({
            "source": "nasa_roses",
            "opp_id": f"nasa-roses-{key}",
            "opp_number": number,
            "title": clean_title(title, element, kind),
            "agency_code": "NASA",
            "agency_name": "National Aeronautics and Space Administration",
            "post_date": posted,
            "open_date": posted,
            "close_date": fmt(deadline) if deadline else "",
            "close_date_raw": fmt(deadline) if deadline else "",
            "status": "posted",
            "url": (f"https://nspires.nasaprs.com/external/solicitations/summary.do?solNum={sols[0]}"
                    if len(sols) == 1 else p["link"]),
            "description": text[:2000],
            "type_hint": "GRANT" if kind in ("new", "date change") else "POLICY",
        })
        if kind != "new" and len(sols) == 1:
            rec["opp_number"] = ""
        rec["needs_review"].append(f"NASA ROSES {kind}" + (f" ({element})" if element else ""))
        if sequence:
            rec["needs_review"].append("dates: " + sequence)
        if kind in ("change", "date change"):
            rec["needs_review"].append("change to an existing ROSES program; include only if it matters to UNR")
        if not deadline:
            rec["needs_review"].append("no due date in the post - for a Policy item use the date it takes effect")
        if rec["url"] != p["link"]:
            rec["needs_review"].append(f"ROSES blog post: {p['link']}")
        kept.append(rec)
        ledger[key] = {"first_seen": iso(today), "posted": posted, "title": title[:120]}
        print(f"  kept {kind:12} {rec['opp_number'] or '-':22} {rec['title'][:50]}", file=sys.stderr)

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
          f"skipped: {len(skipped)} -> {skipped_out}", file=sys.stderr)
    print(f"Wrote -> {args.out}", file=sys.stderr)


if __name__ == "__main__":
    main()
