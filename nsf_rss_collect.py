#!/usr/bin/env python3
"""
nsf_rss_collect.py
==================
Reads the NSF funding RSS feed and emits canonical records.
Discovery-only: surfaces new NSF items with blank deadline flagged for review.

RUN ON YOUR LAPTOP (needs outbound HTTPS to www.nsf.gov).

    python3 nsf_rss_collect.py --out nsf_opportunities.json
"""
import urllib.request, json, argparse, sys, os, datetime, re
import xml.etree.ElementTree as ET
from email.utils import parsedate_to_datetime
from grants_common import (normalize_encoding, detect_type, assess_deadline,
                           blank_record, iso)

FEED = "https://www.nsf.gov/rss/rss_www_funding.xml"

# NSF publishes this feed in bursts, not continuously. Measured 2026-09-03, the
# whole 20-item feed spanned ages 9-27 days -- 16 of those items landed in a
# single two-day burst on Aug 13-14. A 7-day window (the original value)
# therefore returns nothing on most weeks: 17 records on 08-17, then 2, then 0,
# then 0, with a healthy feed and zero parse failures the entire time. The
# collector wasn't broken; the window was narrower than the publishing cadence.
#
# A wider window alone would re-surface the same solicitations every week,
# because nothing dedups across weekly runs (merge_and_assess.py dedups against
# the staging workbook, and each week starts a fresh one). So the window widens
# AND a small ledger records what has already been emitted. An item returns only
# if NSF genuinely re-publishes it -- i.e. its pubDate changes -- so a real
# solicitation update still reaches review.
POSTED_WITHIN_DAYS = 30
SEEN_LEDGER = "nsf_seen.json"
LEDGER_RETENTION_DAYS = 365

NUM_RE = re.compile(r"\b((?:NSF|PD)\s?\d{2}-\d{3,4}[A-Z]?)\b")
DATE_RE = re.compile(r"\b([A-Z][a-z]+ \d{1,2}, \d{4})\b")


def fetch_feed(url):
    req = urllib.request.Request(url, headers={"User-Agent": "UNR-grants-collector/1.0"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read()


def parse_pubdate(s):
    try:
        return parsedate_to_datetime(s).date()
    except (TypeError, ValueError):
        return None


def load_ledger(path):
    """{key: {'first_seen': iso, 'last_pub': iso, 'title': str}}. Missing or
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
              if (v.get("last_pub") or v.get("first_seen") or "") >= cutoff.isoformat()}
    with open(path, "w", encoding="utf-8") as f:
        json.dump(dict(sorted(pruned.items())), f, indent=2, ensure_ascii=False)
    return len(ledger) - len(pruned)


def ledger_key(number, link):
    return (number or link or "").strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="nsf_opportunities.json")
    ap.add_argument("--feed", default=FEED)
    ap.add_argument("--days", type=int, default=POSTED_WITHIN_DAYS,
                    help=f"posting window in days (default {POSTED_WITHIN_DAYS}); "
                         "NSF posts in bursts, so this is deliberately wide")
    ap.add_argument("--seen", default=SEEN_LEDGER,
                    help=f"ledger of already-emitted items (default {SEEN_LEDGER}); "
                         "keep it committed so weekly runs don't repeat themselves")
    ap.add_argument("--ignore-seen", action="store_true",
                    help="emit matching items even if the ledger has them (does "
                         "not write the ledger)")
    args = ap.parse_args()

    print("Fetching NSF funding feed...", file=sys.stderr)
    raw = fetch_feed(args.feed)
    root = ET.fromstring(raw)
    items = root.findall(".//item")
    print(f"  {len(items)} items in feed.", file=sys.stderr)

    today = datetime.date.today()
    ledger = {} if args.ignore_seen else load_ledger(args.seen)
    if ledger:
        print(f"  ledger: {len(ledger)} item(s) already emitted.", file=sys.stderr)
    kept, dropped_old, dropped_seen, republished = [], 0, 0, 0
    for it in items:
        title = normalize_encoding((it.findtext("title") or ""))
        link = (it.findtext("link") or "").strip()
        desc = normalize_encoding((it.findtext("description") or ""))
        pub = parse_pubdate(it.findtext("pubDate"))

        if pub is None or (today - pub).days > args.days:
            dropped_old += 1
            continue

        num_m = NUM_RE.search(title) or NUM_RE.search(desc)
        number = num_m.group(1).replace("  ", " ") if num_m else ""

        key = ledger_key(number, link)
        prior = ledger.get(key)
        if prior:
            if prior.get("last_pub") == iso(pub):
                dropped_seen += 1
                continue
            # Same solicitation, new pubDate: NSF re-published it, so it goes
            # back through review rather than being silently suppressed.
            republished += 1
            print(f"  re-published {number or '(no number)':14} "
                  f"{prior.get('last_pub')} -> {iso(pub)}", file=sys.stderr)

        dm = DATE_RE.search(desc)
        if dm:
            disp, status, days, multi = assess_deadline(dm.group(1))
            review = [] if status in ("ok", "soon") else ["verify deadline (parsed from description)"]
        else:
            disp, status, days, multi = ("", "unknown", None, None)
            review = ["NO deadline in NSF feed - resolve during review (rolling pd## -> exclude)"]

        rec = blank_record()
        rec.update({
            "source": "nsf_rss",
            "opp_id": number or link,
            "opp_number": number,
            "title": title,
            "agency_code": "NSF",
            "agency_name": "National Science Foundation",
            "post_date": iso(pub),
            "open_date": iso(pub),
            "close_date": disp,
            "close_date_raw": dm.group(1) if dm else "",
            "status": "posted",
            "url": link,
            "description": desc,
            "type_hint": detect_type(title, number),
            "needs_review": review,
        })
        kept.append(rec)
        ledger[key] = {"first_seen": (prior or {}).get("first_seen") or iso(today),
                       "last_pub": iso(pub),
                       "title": title[:120]}
        print(f"  kept {number or '(no number)':14} {title[:50]}", file=sys.stderr)

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(kept, f, indent=2, ensure_ascii=False)

    if args.ignore_seen:
        print("\n--ignore-seen: ledger NOT updated.", file=sys.stderr)
    else:
        expired = save_ledger(args.seen, ledger, today)
        print(f"\nLedger -> {args.seen} ({len(ledger) - expired} tracked"
              + (f", {expired} pruned past {LEDGER_RETENTION_DAYS}d" if expired else "")
              + ")", file=sys.stderr)
    print(f"Kept {len(kept)} | window {args.days}d | dropped not-recent: "
          f"{dropped_old} | already emitted: {dropped_seen}"
          + (f" | re-published: {republished}" if republished else ""), file=sys.stderr)
    print("NOTE: NSF deadlines need resolving during review (feed has none).", file=sys.stderr)
    print(f"Wrote -> {args.out}", file=sys.stderr)


if __name__ == "__main__":
    main()
