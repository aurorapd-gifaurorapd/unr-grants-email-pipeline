#!/usr/bin/env python3
"""
merge_and_assess.py  (v2)
=========================
Stage 3: read all collector JSONs, MERGE, DEDUPLICATE, assess eligibility +
deadline, pre-fill a transparent recommendation, and write the Excel staging
sheet for human review.

Usage:
    python3 merge_and_assess.py federal_opportunities.json pivot_opportunities.json \
            nih_opportunities.json nsf_opportunities.json
    python3 merge_and_assess.py *.json --out Grant_Opportunities_Staging.xlsx
"""
import sys, os, json, re, argparse
from datetime import datetime
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:
    print("ERROR: openpyxl not installed. Run: pip install openpyxl")
    sys.exit(1)

from grants_common import assess_deadline, TYPE_INFO
from build_html import SECTION_ORDER

TODAY = datetime.today().date()
SOON_THRESHOLD_DAYS = 14

# ── Agency rollup ────────────────────────────────────────────────────────────
AGENCY_MAP = [
    ("national institutes of health","National Institutes of Health","NIH"),
    ("national science foundation","National Science Foundation","NSF"),
    ("centers for disease control","Centers for Disease Control and Prevention","CDC"),
    ("national endowment for the humanities","National Endowment for the Humanities","NEH"),
    ("national endowment for the arts","National Endowment for the Arts","NEA"),
    ("national institute of food and agriculture","United States Department of Agriculture","USDA"),
    ("department of energy","United States Department of Energy","DOE"),
    ("national aeronautics and space administration","National Aeronautics and Space Administration","NASA"),
    ("department of defense","United States Department of Defense","DOD"),
    # DoD sub-agencies sam_collect.py's ORGS list searches under their own names
    # (not "Department of Defense"), so none of these matched any pattern above
    # and fell through to the "NEW" fallback -- confirmed 2026-08-06 when a
    # DARPA record instead landed with a hand-set "DARPA" abbreviation and its
    # own standalone email section rather than rolling into DOD. sam_collect.py's
    # fullParentPathCode isn't a reliable match key (varies by SAM.gov hierarchy
    # response), so this matches on the literal agency_name strings ORGS uses
    # ("DARPA", "Office of Naval Research", "Army Research Office", "NGA",
    # "AFOSR") -- keep this list in sync with sam_collect.py's ORGS.
    ("darpa","United States Department of Defense","DOD"),
    ("defense advanced research projects agency","United States Department of Defense","DOD"),
    ("office of naval research","United States Department of Defense","DOD"),
    ("army research office","United States Department of Defense","DOD"),
    ("nga","United States Department of Defense","DOD"),
    ("afosr","United States Department of Defense","DOD"),
    ("environmental protection agency","Environmental Protection Agency","EPA"),
    ("department of education","United States Department of Education","ED"),
    ("national oceanic and atmospheric","National Oceanic and Atmospheric Administration","NOAA"),
    ("geological survey","United States Geological Survey","USGS"),
    ("department of homeland security","Department of Homeland Security","DHS"),
    ("department of transportation","United States Department of Transportation","DOT"),
    ("small business administration","Small Business Administration","SBA"),
    ("department of the interior","United States Department of the Interior","DOI"),
    ("health resources and services","United States Department of Health and Human Services","HHS"),
    ("office of the assistant secretary for health","United States Department of Health and Human Services","HHS"),
    ("office of the national coordinator","United States Department of Health and Human Services","HHS"),
    ("advanced research projects agency for health","United States Department of Health and Human Services","HHS"),
    ("arpa-h","United States Department of Health and Human Services","HHS"),
    ("bureau of land management","United States Department of the Interior","DOI"),
    ("department of health and human services","United States Department of Health and Human Services","HHS"),
    ("national institute of justice","United States Department of Justice","DOJ"),
    ("department of justice","United States Department of Justice","DOJ"),
    ("department of state","United States Department of State","DOS"),
    ("department of commerce","United States Department of Commerce","DOC"),
    ("department of labor","United States Department of Labor","DOL"),
    ("department of housing","United States Department of Housing and Urban Development","HUD"),
    # Internal UNR opportunities (Scout-sourced COBRE/pilot calls and the like) get
    # their own abbrev so build_html.py can give them a dedicated email section
    # instead of filing them under "New Foundation R&D".
    ("university of nevada, reno","University of Nevada, Reno","UNR"),
]
CODE_ROLLUP = {
    "NSF":"NSF","USDA":"USDA","EPA":"EPA","NEH":"NEH","NEA":"NEA","ED":"ED",
    "NASA":"NASA","DOE":"DOE","DOD":"DOD","HHS":"HHS","NIH":"NIH","DOI":"DOI",
    "DOT":"DOT","DOC":"DOC","DOL":"DOL","HUD":"HUD","DOJ":"DOJ","SBA":"SBA",
    "USDOJ":"DOJ","USDA-NIFA":"USDA","USDA-NRCS":"USDA",
    "HHS-NIH11":"NIH",
}
ABBR_TO_FULL = {abbr: full for _, full, abbr in AGENCY_MAP}
# Word-boundary, not substring: "national science foundation" must not match
# inside "US-Israel Binational Science Foundation" (the "national" in
# "binational" has no boundary before it, so \b correctly excludes it).
_AGENCY_PATTERNS = [(re.compile(r"\b" + re.escape(pat) + r"\b"), full, abbr)
                     for pat, full, abbr in AGENCY_MAP]

def lookup_agency(agency_code, agency_name):
    if agency_code:
        if agency_code in CODE_ROLLUP:
            abbr = CODE_ROLLUP[agency_code]
            return ABBR_TO_FULL.get(abbr, agency_name or abbr), abbr
        prefix = agency_code.split("-")[0]
        if prefix in CODE_ROLLUP:
            abbr = CODE_ROLLUP[prefix]
            return ABBR_TO_FULL.get(abbr, agency_name or abbr), abbr
    name = (agency_name or "").lower()
    for rx, full, abbr in _AGENCY_PATTERNS:
        if rx.search(name):
            return full, abbr
    return (agency_name or agency_code or "UNKNOWN").strip()[:200], "NEW"

# ── Eligibility tiering ──────────────────────────────────────────────────────
STANDARD_PHRASES = [
    "higher education institutions","institutions of higher education",
    "institution of higher education","public/state controlled institutions",
    "public and state controlled institutions","academic institution",
    "accredited institution of higher education","nonprofit","non-profit",
    "universities","colleges and universities","unrestricted",
]
RESTRICT_PATTERNS = [
    (r"minority[- ]serving","MSI-restricted"),
    (r"\bhbcu\b|historically black","HBCU-restricted"),
    (r"hispanic[- ]serving|\bhsi\b","HSI-restricted"),
    (r"tribal college|\btcu\b","TCU-restricted"),
    (r"individual applicants only|individuals only","Individuals only"),
    (r"members? only|membership required","Members only"),
    (r"prior grantee|active award recipient|current grantees","Prior grantees only"),
    (r"previously received|already received|must have received.{0,60}(dod|dow|federal|government) (fund|grant|award|contract)","Prior funded-only"),
    (r"students? only|undergraduate only|graduate students? only|pre-?doc","Student-only (not faculty)"),
]
INSTITUTION_DROP = re.compile(
    r"\bhbcu\b|historically black college|minority[- ]serving institution|"
    r"\bmurep\b|minority university research|hispanic[- ]serving institution|\bhsi\b program|"
    r"tribal college[s]? (research|extension|program)",
    re.I
)

SELECTIVE_NOTES = [
    (r"only one application|one application per|limited submission","Limited submission"),
    (r"early stage investigator|\besi\b","ESI"),
    (r"early[- ]career","Early career"),
    (r"\bepscor\b","EPSCoR"),
    (r"postdoctoral|postdoc","Postdoctoral"),
]
def assess_eligibility(elig_text, applicant_type=""):
    t = ((elig_text or "") + " " + (applicant_type or "")).lower()
    restrictions = [label for pat, label in RESTRICT_PATTERNS if re.search(pat, t)]
    notes = [label for pat, label in SELECTIVE_NOTES if re.search(pat, t)]
    has_standard = any(p in t for p in STANDARD_PHRASES)
    if restrictions and not has_standard:
        tier = "restricted"
    elif has_standard:
        tier = "standard"
    elif not t.strip():
        tier = "empty"
    else:
        tier = "likely_open"
    return tier, notes, restrictions

def eligibility_label(tier, notes, restrictions):
    if tier == "restricted":
        return "Restricted: " + "; ".join(restrictions or ["see notes"])
    if tier == "standard":
        return ("Standard + notes: " + "; ".join(notes)) if notes else "Standard (Higher Ed, Nonprofits, Gov)"
    if tier == "likely_open":
        return "Likely open (foundation - verify)" + (": " + "; ".join(notes) if notes else "")
    if tier == "empty":
        return "No eligibility text - verify"
    return "UNKNOWN - verify eligibility"

# ── Recommendation reason string ─────────────────────────────────────────────
def recommend(tier, restrictions, dl_status, days_out, notes, source):
    parts = []
    if tier == "restricted":
        parts.append("FLAG: " + "; ".join(restrictions or ["restricted"]))
    elif tier in ("empty",):
        parts.append("CHECK eligibility (no text)")
    elif tier == "likely_open":
        parts.append("Likely eligible (verify)")
    else:
        parts.append("Higher-ed eligible")
    if any("EPSCoR" in n for n in notes):
        parts.append("EPSCoR advantage (Nevada)")
    if dl_status == "soon":
        parts.append(f"deadline tight ({days_out}d)")
    elif dl_status in ("rolling", "unknown"):
        parts.append("verify deadline")
    elif days_out is not None:
        parts.append(f"deadline ~{days_out}d out")
    if any("Limited submission" in n for n in notes):
        parts.append("limited submission")
    parts.append({"federal":"Federal","pivot":"Foundation/non-federal",
                  "nih_guide":"NIH notice","nsf_rss":"NSF (verify deadline)"}.get(source, source))
    return "; ".join(parts)

# ── Dedup signatures ─────────────────────────────────────────────────────────
def norm_number(n):
    return re.sub(r"[^a-z0-9]", "", (n or "").lower())
_STOP = {"the","a","an","of","for","and","in","to","on","program","grant",
         "grants","award","awards","research","fund","funds","fellowship"}
def title_sig(title, funder=None):
    words = re.findall(r"[a-z0-9]+", (title or "").lower())
    sig_words = [w for w in words if w not in _STOP][:4]
    sig = "".join(sig_words)
    if len(sig) < 6:
        return "FULL:" + "".join(words)
    return sig
SOURCE_RANK = {"federal":0, "nih_guide":1, "nsf_rss":2, "pivot":3}

# ── Excel ────────────────────────────────────────────────────────────────────
HEADERS = ["Email?","Eligibility Flag","Type","Recommendation","LOI Alert",
    "Funding Agency","Abbreviation","Title","Summary","Deadline",
    "Link1 URL","Link1 Text","Link2 URL","Link2 Text","Eligibility Notes","Record ID"]
HEADER_FILL = PatternFill("solid", fgColor="FF041E42")
FLAG_FILL   = PatternFill("solid", fgColor="FFFFF2CC")

def build_row(rec):
    source = rec.get("source", "federal")
    full_name, abbr = lookup_agency(rec.get("agency_code"), rec.get("agency_name"))
    if abbr != "NEW" and abbr not in SECTION_ORDER:
        # AGENCY_MAP/CODE_ROLLUP should only ever produce an abbr that's either
        # "NEW" or already in build_html.SECTION_ORDER. If this fires, one of
        # those tables drifted out of sync -- a silently misrouted section
        # (e.g. the 2026-08-13 DARPA incident) is worse than a hard stop here.
        raise SystemExit(
            f"Agency routing bug: lookup_agency() produced abbr '{abbr}' for "
            f"agency_name='{rec.get('agency_name')}' agency_code='{rec.get('agency_code')}', "
            f"which is not in build_html.SECTION_ORDER and not 'NEW'. "
            f"Fix AGENCY_MAP/CODE_ROLLUP or add '{abbr}' to SECTION_ORDER.")
    tier, notes, restrictions = assess_eligibility(rec.get("eligibility_desc"), rec.get("applicant_type",""))
    elig_label = eligibility_label(tier, notes, restrictions)
    disp, dl_status, days, multi = assess_deadline(rec.get("close_date_raw") or rec.get("close_date"))
    rec_reason = recommend(tier, restrictions, dl_status, days, notes, source)
    record_id = rec.get("opp_number") or rec.get("opp_id") or ""
    link1 = rec.get("url","")
    link1_text = rec.get("opp_number") or full_name
    link2 = ""
    link2_text = ""
    # NOTICE/GRANT map to blank (the Type column's default). POLICY only ever
    # comes from the grant-email-scout skill -- none of the four collectors'
    # detect_type() produces it -- so it was missing here until Scout needed it.
    type_pref = {"RFI":"RFI","DCL":"DCL","EVENT":"Event","POLICY":"Policy",
                 "NOTICE":"","GRANT":""}.get(rec.get("type_hint",""), "")
    review_join = "; ".join(rec.get("needs_review", []))
    row = ["", "Y" if tier == "restricted" else "", type_pref, rec_reason, "",
           full_name, abbr, rec.get("title",""),
           rec.get("summary") or rec.get("description",""), disp,
           link1, link1_text, link2, link2_text,
           (elig_label + ((" || " + review_join) if review_join else "")), record_id]
    flag = (tier == "restricted") or (dl_status == "soon") or (tier in ("empty",)) or bool(rec.get("needs_review"))
    return row, flag, dl_status, record_id

def load_existing(path):
    if not os.path.exists(path):
        return set(), set(), None
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    ids, sigs = set(), set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[15]:
            ids.add(norm_number(str(row[15])))
        if row and row[7]:
            sigs.add(title_sig(str(row[7]), str(row[5]) if len(row) > 5 else ""))
    return ids, sigs, wb

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("infiles", nargs="+")
    ap.add_argument("--out", default="Grant_Opportunities_Staging.xlsx")
    args = ap.parse_args()

    existing_ids, existing_sigs, wb = load_existing(args.out)
    if wb is None:
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Grant Opportunities"
        ws.append(HEADERS)
        for c, _ in enumerate(HEADERS, 1):
            cell = ws.cell(row=1, column=c)
            cell.font = Font(bold=True, color="FFFFFFFF"); cell.fill = HEADER_FILL
        print(f"Created new staging file: {args.out}")
    else:
        ws = wb.active
        print(f"Opened existing staging: {args.out} ({len(existing_ids)} prior records)")

    records = []
    for path in args.infiles:
        if not os.path.exists(path):
            print(f"  (skip missing {path})"); continue
        with open(path, encoding="utf-8") as f:
            recs = json.load(f)
        guess = ("pivot" if "pivot" in path.lower() else
                 "nih_guide" if "nih" in path.lower() else
                 "nsf_rss" if "nsf" in path.lower() else "federal")
        for r in recs:
            r.setdefault("source", guess)
        records.extend(recs)
        print(f"  loaded {len(recs)} from {os.path.basename(path)}")

    records.sort(key=lambda r: SOURCE_RANK.get(r.get("source","pivot"), 9))
    seen_num, seen_sig = {}, {}
    deduped, dup_in_run = [], 0
    for r in records:
        # An NIH notice is its own document: key it on its NOT- number, not on
        # the parent NOFO number that its title happens to cite.
        notice = re.search(r"notice-files/(NOT-[A-Za-z0-9-]+?)\.html?", r.get("url") or "")
        if notice:
            r["opp_number"] = notice.group(1)
        num = norm_number(r.get("opp_number"))
        sig = title_sig(r.get("title"), r.get("agency_name"))
        if num and num in seen_num:
            dup_in_run += 1; continue
        # A title match only counts as a duplicate when at least one record has
        # no opportunity number. Two different numbers are two opportunities.
        if sig in seen_sig and not (num and seen_sig[sig]):
            dup_in_run += 1; continue
        if num: seen_num[num] = True
        seen_sig.setdefault(sig, num)
        deduped.append(r)

    # Title matches here are only against rows from a prior merge; matches
    # within this run were already settled by the loop above.
    prior_sigs = set(existing_sigs)
    added = dropped_passed = skipped_dup = dropped_inst = flagged = 0
    drop_log, flag_log = [], []
    for rec in deduped:
        _inst_check = " ".join([
            rec.get("title", ""),
            rec.get("eligibility_desc", ""),
            rec.get("description", "")[:800],
        ])
        if INSTITUTION_DROP.search(_inst_check):
            dropped_inst += 1
            print(f"  [institution-drop] {rec.get('title','')[:65]}", file=sys.stderr)
            continue
        row, flag, dl_status, rid = build_row(rec)
        if dl_status == "passed":
            dropped_passed += 1; drop_log.append(f"  {rec.get('title','')[:50]} | {row[9]}"); continue
        nrid, nsig = norm_number(rid), title_sig(row[7], row[5])
        if (nrid and nrid in existing_ids) or (nsig in prior_sigs):
            skipped_dup += 1; continue
        ws.append(row)
        if nrid: existing_ids.add(nrid)
        existing_sigs.add(nsig)
        added += 1
        if flag:
            flagged += 1
            for c in range(1, 17):
                ws.cell(row=ws.max_row, column=c).fill = FLAG_FILL
            flag_log.append(f"  {row[7][:50]} | {row[3]}")

    widths = {"A":7,"B":13,"C":10,"D":40,"E":12,"F":28,"G":7,"H":40,
              "I":50,"J":18,"K":30,"L":18,"M":30,"N":10,"O":40,"P":24}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    for col in ("D","H","I","O"):
        for cell in ws[col]:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    dv_type = DataValidation(type="list",
                              formula1='"' + ",".join(code for code, _, _ in TYPE_INFO) + '"',
                              allow_blank=True)
    dv_email = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    ws.add_data_validation(dv_type); ws.add_data_validation(dv_email)
    if ws.max_row >= 2:
        dv_type.add(f"C2:C{ws.max_row}")
        dv_email.add(f"A2:A{ws.max_row}")
    ws.freeze_panes = "A2"
    wb.save(args.out)

    print(f"\n{'='*60}")
    print(f"DONE - added: {added} | in-run dups: {dup_in_run} | "
          f"cross-week dups: {skipped_dup} | passed-deadline dropped: {dropped_passed} | "
          f"institution-drop (HBCU/MSI/TCU): {dropped_inst} | flagged: {flagged}")
    print(f"Saved: {args.out}")
    if drop_log:
        print("\nDROPPED (deadline passed):"); [print(l) for l in drop_log]
    if flag_log:
        print("\nFLAGGED (review):"); [print(l) for l in flag_log]

if __name__ == "__main__":
    main()
