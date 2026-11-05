#!/usr/bin/env python3
"""
pivot_collect.py  (v2)
======================
Reads a Pivot-RP CSV export, applies hard categorical filters, and emits
canonical records for merge_and_assess.py.

Usage:
    python3 pivot_collect.py opps_export.csv --out pivot_opportunities.json
"""
import csv, json, re, argparse, sys
from grants_common import (normalize_encoding, detect_type, assess_deadline,
                           blank_record)

# ── Geography ────────────────────────────────────────────────────────────────
US_TOKENS = ["united states", "usa", "u.s.", "america", "north america", "unrestricted"]
US_STATES = ["alabama","alaska","arizona","arkansas","california","colorado",
    "connecticut","delaware","florida","georgia","hawaii","idaho","illinois",
    "indiana","iowa","kansas","kentucky","louisiana","maine","maryland",
    "massachusetts","michigan","minnesota","mississippi","missouri","montana",
    "nebraska","nevada","new hampshire","new jersey","new mexico","new york",
    "north carolina","north dakota","ohio","oklahoma","oregon","pennsylvania",
    "rhode island","south carolina","south dakota","tennessee","texas","utah",
    "vermont","virginia","washington","west virginia","wisconsin","wyoming"]
def is_us_ish(value):
    t = (value or "").strip().lower()
    if not t or "unspecified" in t or "not specified" in t:
        return True
    if any(tok in t for tok in US_TOKENS):
        return True
    if any(st in t for st in US_STATES):
        return True
    return False

def geography_ok(activity, inst):
    if is_us_ish(activity) or is_us_ish(inst):
        return True, None
    return False, f"non-US: activity='{(activity or '').strip()[:30]}' inst='{(inst or '').strip()[:30]}'"

# ── Citizenship ──────────────────────────────────────────────────────────────
def citizenship_ok(value):
    t = (value or "").strip().lower()
    if not t:
        return True, None
    if any(tok in t for tok in ["unrestricted", "united states", "north america", "u.s.", "usa"]):
        return True, None
    return False, f"citizenship excludes US: '{value.strip()[:40]}'"

# ── Career stage ─────────────────────────────────────────────────────────────
FACULTY_STAGES = ["early career", "emerging in field", "mid-career",
                  "established in field", "not applicable"]
STUDENT_STAGES = ["undergraduate student", "graduate student", "pre-doc",
                  "predoctoral", "high school"]
def career_ok(value):
    t = (value or "").strip().lower()
    if not t:
        return True, None
    has_fac = any(s in t for s in FACULTY_STAGES)
    has_stu = any(s in t for s in STUDENT_STAGES)
    if has_fac:
        return True, None
    if has_stu:
        return False, f"student-only (no faculty stage): '{value.strip()[:40]}'"
    return True, None

# ── Language ─────────────────────────────────────────────────────────────────
NON_EN_MARKERS = re.compile(
    r"\b(de la|de l'|des|du|pour|avec|dans|sur|och|for|della|delle|degli|"
    r"para|con|fur|und|a|appel|dossiers|manifestion)\b", re.IGNORECASE)
ACCENT_RE = re.compile(r"[àâäéèêëïîôùûüçöå]", re.IGNORECASE)
def english_ok(title, abstract):
    sample = f"{title} {abstract}"
    if ACCENT_RE.search(title) and NON_EN_MARKERS.search(sample):
        return False, f"appears non-English: '{title.strip()[:40]}'"
    if ACCENT_RE.search(title) and len(ACCENT_RE.findall(title)) >= 3:
        return False, f"appears non-English: '{title.strip()[:40]}'"
    return True, None

# ── Soft flags ───────────────────────────────────────────────────────────────
RESTRICT_PATTERNS = [
    (r"minority[- ]serving",          "POTENTIAL INELIGIBILITY: MSI-restricted"),
    (r"\bhbcu\b|historically black",  "POTENTIAL INELIGIBILITY: HBCU-restricted"),
    (r"hispanic[- ]serving|\bhsi\b",  "POTENTIAL INELIGIBILITY: HSI-restricted"),
    (r"tribal college|\btcu\b",       "POTENTIAL INELIGIBILITY: TCU-restricted"),
    (r"members? only|membership required", "POTENTIAL INELIGIBILITY: members-only"),
    (r"prior grantee|current grantee", "POTENTIAL INELIGIBILITY: prior-grantees-only"),
]
STUDENT_TITLE_RE = re.compile(r"\b(student|dissertation|trainee|undergraduate|predoctoral|pre-doctoral)\b", re.I)
def soft_flags(title, elig, funding_type):
    flags = []
    t = (elig or "").lower()
    for pat, msg in RESTRICT_PATTERNS:
        if re.search(pat, t):
            flags.append(msg)
    if "epscor" in t:
        flags.append("ADVANTAGE: EPSCoR-relevant (Nevada qualifies)")
    if STUDENT_TITLE_RE.search(title or ""):
        flags.append("TITLE hints student/trainee - verify faculty eligibility")
    ft = (funding_type or "").lower()
    if ft and "research: project grants" not in ft:
        flags.append(f"funding type not research-grant: '{(funding_type or '').strip()[:50]}'")
    return flags


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("infile")
    ap.add_argument("--out", default="pivot_opportunities.json")
    ap.add_argument("--skipped-out", default="pivot_skipped.json")
    args = ap.parse_args()

    with open(args.infile, encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    records, skipped, flagged = [], [], 0
    for r in rows:
        title = normalize_encoding(r.get("Title", ""))
        if not title:
            continue
        funder   = normalize_encoding(r.get("Funder", ""))
        abstract = normalize_encoding(r.get("Abstract", "")) or normalize_encoding(r.get("AI Generated Summary", ""))
        elig     = normalize_encoding(r.get("Eligibility", ""))
        career   = normalize_encoding(r.get("Career Stage", ""))
        activity = r.get("Activity Location", "")
        inst_loc = r.get("Applicant/Institution Location", "")
        citizen  = r.get("Citizenship", "")
        ftype    = normalize_encoding(r.get("Funding Type", ""))
        deadline_raw = normalize_encoding(r.get("Deadline", ""))

        reasons = []
        for ok, why in (geography_ok(activity, inst_loc),
                        citizenship_ok(citizen),
                        career_ok(career),
                        english_ok(title, abstract)):
            if not ok:
                reasons.append(why)
        if reasons:
            skipped.append({"title": title, "funder": funder,
                            "career_stage": career,
                            "activity_location": (activity or "").strip(),
                            "institution_location": (inst_loc or "").strip(),
                            "citizenship": (citizen or "").strip(),
                            "reasons": reasons})
            continue

        disp, status, days, multi = assess_deadline(deadline_raw)
        review = soft_flags(title, elig, ftype)
        if multi:
            review.append(multi)
        if status == "passed":
            review.append(f"deadline PASSED ({disp}) - verify")
        elif status == "soon":
            review.append(f"deadline tight ({days}d)")
        elif status in ("rolling", "unknown"):
            review.append("no firm deadline - verify (rolling/unknown)")

        rec = blank_record()
        rec.update({
            "source": "pivot",
            "opp_id": (r.get("Opportunity ID") or "").strip(),
            "opp_number": "",
            "title": title,
            "agency_name": funder,
            "close_date": disp,
            "close_date_raw": deadline_raw,
            "status": "posted",
            "url": normalize_encoding(r.get("More Information", "")) or normalize_encoding(r.get("Link to Pivot-RP", "")),
            "pivot_url": normalize_encoding(r.get("Link to Pivot-RP", "")),
            "description": abstract,
            "eligibility_desc": elig,
            "career_stage": career,
            "activity_location": (activity or "").strip(),
            "award_amount": normalize_encoding(r.get("Amount", "")),
            "type_hint": detect_type(title),
            "needs_review": review,
        })
        records.append(rec)
        if review:
            flagged += 1

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2, ensure_ascii=False)
    with open(args.skipped_out, "w", encoding="utf-8") as f:
        json.dump(skipped, f, indent=2, ensure_ascii=False)

    print(f"Kept {len(records)} -> {args.out}", file=sys.stderr)
    print(f"Skipped {len(skipped)} (logged -> {args.skipped_out}):", file=sys.stderr)
    for s in skipped:
        print(f"  - {s['title'][:46]:48} | {'; '.join(s['reasons'])}", file=sys.stderr)
    print(f"{flagged} of kept need review.", file=sys.stderr)


if __name__ == "__main__":
    main()
