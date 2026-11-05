#!/usr/bin/env python3
"""
nih_rss_collect.py
==================
Reads the NIH Guide for Grants & Contracts RSS feed and emits canonical records.
Keeps only notice-type items (RFI / DCL / NOTICE) by default; use --keep-nofos
to also keep NOFOs (which overlap the Grants.gov federal pull).

RUN ON YOUR LAPTOP (needs outbound HTTPS to grants.nih.gov).

    python3 nih_rss_collect.py --out nih_opportunities.json
"""
import urllib.request, json, argparse, sys, datetime, re
import xml.etree.ElementTree as ET
from email.utils import parsedate_to_datetime
from grants_common import (normalize_encoding, detect_type, assess_deadline,
                           blank_record, iso)

FEED = "https://grants.nih.gov/grants/guide/newsfeed/fundingopps.xml"
POSTED_WITHIN_DAYS = 7

NUM_RE = re.compile(r"\b((?:NOT|PAR|PA|RFA|PAS|RFI)-[A-Z0-9]{1,3}-\d{2}-\d{2,4}[A-Z0-9]*)\b")
RESP_RE = re.compile(r"(?:response date|responses?\s+due|due date)\s*[:\-]?\s*"
                     r"([A-Z][a-z]+ \d{1,2}, \d{4}|\d{1,2}/\d{1,2}/\d{4})", re.I)


def fetch_feed(url):
    req = urllib.request.Request(url, headers={"User-Agent": "UNR-grants-collector/1.0"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read()


def parse_pubdate(s):
    try:
        return parsedate_to_datetime(s).date()
    except (TypeError, ValueError):
        return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="nih_opportunities.json")
    ap.add_argument("--feed", default=FEED)
    ap.add_argument("--keep-nofos", action="store_true")
    args = ap.parse_args()

    print(f"Fetching NIH Guide feed...", file=sys.stderr)
    raw = fetch_feed(args.feed)
    root = ET.fromstring(raw)
    items = root.findall(".//item")
    print(f"  {len(items)} items in feed.", file=sys.stderr)

    today = datetime.date.today()
    kept, dropped_old, dropped_nofo = [], 0, 0
    for it in items:
        title = normalize_encoding((it.findtext("title") or ""))
        link = (it.findtext("link") or "").strip()
        desc = normalize_encoding((it.findtext("description") or ""))
        pub = parse_pubdate(it.findtext("pubDate"))

        if pub is None or (today - pub).days > POSTED_WITHIN_DAYS:
            dropped_old += 1
            continue

        num_m = NUM_RE.search(title) or NUM_RE.search(desc)
        number = num_m.group(1) if num_m else ""
        ttype = detect_type(title, number)

        if ttype in ("GRANT",) and not args.keep_nofos:
            dropped_nofo += 1
            continue

        resp = RESP_RE.search(desc) or RESP_RE.search(title)
        if resp:
            disp, status, days, multi = assess_deadline(resp.group(1))
        else:
            disp, status, days, multi = ("rolling / not specified", "unknown", None, None)

        rec = blank_record()
        rec.update({
            "source": "nih_guide",
            "opp_id": number or link,
            "opp_number": number,
            "title": title,
            "agency_code": "NIH",
            "agency_name": "National Institutes of Health",
            "post_date": iso(pub),
            "open_date": iso(pub),
            "close_date": disp,
            "close_date_raw": resp.group(1) if resp else "",
            "status": "posted",
            "url": link,
            "description": desc,
            "type_hint": ttype,
        })
        if ttype == "NOTICE":
            rec["needs_review"].append("NIH notice - confirm if substantive (skip admin/expiration notices)")
        if status == "unknown":
            rec["needs_review"].append("no response date found - verify")
        kept.append(rec)
        print(f"  kept {ttype:6} {number or '(no number)':18} {title[:45]}", file=sys.stderr)

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(kept, f, indent=2, ensure_ascii=False)
    print(f"\nKept {len(kept)} | dropped not-recent: {dropped_old} | "
          f"dropped NOFOs (in federal pull): {dropped_nofo}", file=sys.stderr)
    print(f"Wrote -> {args.out}", file=sys.stderr)


if __name__ == "__main__":
    main()
