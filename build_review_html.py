#!/usr/bin/env python3
"""
build_review_html.py
=====================
Stage 3.5 (new): read the merge_and_assess.py staging .xlsx and generate a
single self-contained HTML review tool — no server required, works offline
except for one CDN script (SheetJS) used only when you click "Export xlsx".

WHY THIS EXISTS
Reviewing 15-30 multi-paragraph summaries in Excel cells is painful (wrap text,
tiny row heights, no filtering by flag status). This gives you a scrollable
card view: search, filter by section / flagged / needs-AI-help, inline-edit
Title / Type / Deadline / Summary, check "Include in email" per record, then
export a validated_*.xlsx in the EXACT 16-column layout build_html.py expects
— plus columns 17 (Q, "Needs AI Summary Help") and 18 (R, "AI Rewrite Notes")
that build_html.py ignores, so they're safe passthrough for your own tracking.
Re-running the script against an already-exported file reads columns 17/18
back in, so notes you've already left aren't wiped out by a re-run.

EXPORT ONLY INCLUDES CHECKED ROWS: rows not marked "Include in email" are
dropped entirely from the exported file, not just marked blank. This is
purely to shrink the file handed to a downstream Copilot/Claude rewrite
agent — build_html.py already ignores Email?!=Y rows regardless, so this
does not change what ends up in the final email either way. It does mean
the exported file is no longer a record of what was reviewed-and-excluded
that week — if you want that audit trail, keep the original staging file.

WHAT IT DOES NOT DO
- Does not call any AI. The "Needs AI Summary Help" checkbox is just a flag
  for YOU (or a downstream Copilot/Claude step) — it does not rewrite anything.
- Does not touch Link1/Link2 URLs or the Recommendation/LOI/Eligibility Notes/
  Record ID columns — those pass through unchanged. build_html.py doesn't
  consume Recommendation/LOI/Eligibility Notes/Record ID anyway (verified
  against build_html.py's read_rows()/grant_block()), so editing them here
  would only affect your own reference copy, not the final email.
- Does not decode HTML entities (&rsquo; etc.) in the exported data — only in
  the on-screen display, so what you export is exactly what was in the source
  file unless you edit a field yourself.

HEURISTIC for pre-checking "Needs AI Summary Help" (you can override any row):
  - Summary is empty, OR
  - Summary still contains raw HTML tags (<p>, <div>, etc.) or entities
    (&nbsp;, &rsquo;, ...) -- sign it hasn't been through the rewrite pass, OR
  - Summary is under ~30 words -- likely a truncated/raw abstract fragment.

Usage:
    python3 build_review_html.py Grant_Opportunities_Staging.xlsx
    python3 build_review_html.py Grant_Opportunities_Staging.xlsx --out review.html
"""
import sys, re, json, argparse, html, datetime
try:
    import openpyxl
except ImportError:
    print("ERROR: pip install openpyxl"); sys.exit(1)

from grants_common import TYPE_OPTIONS  # single source of truth -- do not re-copy here

HEADERS = ["Email?","Eligibility Flag","Type","Recommendation","LOI Alert",
    "Funding Agency","Abbreviation","Title","Summary","Deadline",
    "Link1 URL","Link1 Text","Link2 URL","Link2 Text","Eligibility Notes","Record ID"]
COL_KEYS = ["email","eligflag","type","rec","loi","agency","abbr","title",
            "summary","deadline","l1url","l1text","l2url","l2text","elignotes","recordid"]

HTML_TAG_RE = re.compile(r"<[a-zA-Z/][^>]*>")
ENTITY_RE = re.compile(r"&[a-zA-Z#0-9]+;")

# ── Cross-row duplicate detection ───────────────────────────────────────────
# merge_and_assess.py dedups records against each other using this same
# title_sig() logic (first 4 significant title words) plus a normalized
# opp-number match -- but only for records that actually pass through it.
# Scout-agent-sourced records (RFI/DCL/Event/Policy items entered outside the
# four Python collectors) are written straight into the staging/validated
# workbook and never see that dedup pass, so the same opportunity can show up
# twice: once collected normally, once Scout-agent-sourced under a different
# title. This mirrors that logic here so the review tool catches it too.
_DUP_STOP = {"the","a","an","of","for","and","in","to","on","program","grant",
             "grants","award","awards","research","fund","funds","fellowship"}

def title_sig(title):
    words = re.findall(r"[a-z0-9]+", (title or "").lower())
    sig_words = [w for w in words if w not in _DUP_STOP][:4]
    sig = "".join(sig_words)
    if len(sig) < 6:
        return "FULL:" + "".join(words)
    return sig

def digits_only(s):
    return re.sub(r"\D", "", s or "")

def dup_signature(rec):
    """(title-signature, solicitation-number-digits) for cross-row matching.
    The number comes from Record ID if present, else the visible Link1 Text
    (e.g. "NSF 26-513"), since Scout-agent rows often lack a Record ID but do
    show a solicitation number in the link text."""
    num = digits_only(rec["recordid"]) or digits_only(rec["l1text"])
    return title_sig(rec["title"]), (num if len(num) >= 4 else "")


def needs_ai_help(summary):
    s = (summary or "").strip()
    if not s:
        return True
    if HTML_TAG_RE.search(s) or ENTITY_RE.search(s):
        return True
    word_count = len(re.findall(r"\S+", re.sub(HTML_TAG_RE, " ", s)))
    if word_count < 30:
        return True
    return False


# abbrs build_html.py recognizes as federal (SECTION_ORDER) -- anything else
# that isn't "NEW" gets its own standalone email section instead of folding
# into an existing one. Scout-agent-sourced records are the usual way a wrong
# one slips in, since they skip merge_and_assess.py's lookup_agency() (which
# would otherwise roll the record up correctly) and get a hand-typed
# Abbreviation instead. Confirmed twice 2026-08-06: ACLS got its own section
# with abbr="ACLS" instead of "NEW" (it's a foundation), and a DARPA record
# got its own section with abbr="DARPA" instead of "DOD" (it's a DoD
# sub-agency, not a foundation) -- so the fix isn't always "NEW"; it depends
# on whether the funder is actually federal.
KNOWN_FEDERAL_ABBRS = {"HHS","NIH","NSF","USDA","DOE","DOD","NASA","DOI","EPA","ED",
    "NEH","NEA","NOAA","USGS","DHS","DOT","SBA","DOJ","DOS","DOC","DOL","HUD","CDC","FDA"}


def row_quality_notes(rec):
    """Extra QA flags surfaced in the UI that aren't in the original columns."""
    notes = []
    if rec["type"].strip().lower() == "event" and "any time" in rec["deadline"].lower():
        notes.append("Type=Event but deadline reads as rolling/notice-style — verify this isn't mislabeled.")
    if HTML_TAG_RE.search(rec["summary"] or "") or ENTITY_RE.search(rec["summary"] or ""):
        notes.append("Summary still contains raw HTML tags/entities — likely not yet rewritten.")
    if not rec["deadline"].strip():
        notes.append("No deadline value.")
    abbr = rec["abbr"].strip().upper()
    if abbr and abbr not in KNOWN_FEDERAL_ABBRS and abbr != "NEW":
        notes.append(f"Abbreviation {rec['abbr']!r} isn't a recognized federal code and isn't "
                      f"'NEW' — it will get its own standalone email section instead of folding "
                      f"into an existing one. Verify: if this is a foundation/non-federal funder "
                      f"it should be 'NEW'; if it's a federal sub-agency (e.g. a DoD or HHS "
                      f"component) it should be that parent department's code instead.")
    return notes


FLAG_FILL_RGB = "FFFFF2CC"  # exact fill color merge_and_assess.py applies to flagged rows


def read_rows(path):
    """Loads both cell values (data_only) and a second pass over the styled
    workbook to read the actual fill color merge_and_assess.py applied — this
    is the ground truth for 'flagged', not a text-match heuristic. A row is
    only truly flagged if merge_and_assess.py's own logic (restricted
    eligibility, empty eligibility text, tight/soon deadline, or a collector
    needs_review flag) set it, and that logic is captured in the cell fill,
    not in whether the word 'verify' happens to appear in a notes column."""
    wb_vals = openpyxl.load_workbook(path, data_only=True)
    ws_vals = wb_vals.active
    wb_style = openpyxl.load_workbook(path, data_only=False)
    ws_style = wb_style.active

    out = []
    for i, row in enumerate(ws_vals.iter_rows(min_row=2, values_only=True), start=2):
        if not row or not any(row):
            continue
        rec = {}
        for j in range(min(16, len(row))):
            val = row[j]
            key = COL_KEYS[j]
            if key == "deadline" and isinstance(val, (datetime.date, datetime.datetime)):
                # Same defensive handling as build_html.py: a cell re-typed as
                # a real Excel date shows as '2026-09-17 00:00:00' if str()'d
                # directly. Format it properly instead.
                rec[key] = val.strftime("%B %d, %Y")
            else:
                rec[key] = "" if val is None else str(val)
        for k in COL_KEYS:
            rec.setdefault(k, "")
        # merge_and_assess.py leaves Type blank on purpose for ordinary grants
        # (only RFI/DCL/Event get an explicit label), and build_html.py's
        # TYPE_LABELS treats blank and "Grant" identically. Defaulting blank
        # to "Grant" here is a display fill, not a behavior change.
        if not rec["type"].strip():
            rec["type"] = "Grant"
        rec["_row"] = i
        rec["_email_checked"] = rec["email"].strip().upper() == "Y"
        rec["_qa_notes"] = row_quality_notes(rec)

        # column 17 (index 16) = prior "Needs AI Summary Help" value, if this
        # file was previously exported by this tool. Respect it over the
        # heuristic so re-runs don't silently undo a manual override.
        prior_aihelp = str(row[16]).strip() if len(row) > 16 and row[16] is not None else ""
        if prior_aihelp:
            rec["_needs_ai_help"] = prior_aihelp.upper() == "Y"
        else:
            rec["_needs_ai_help"] = needs_ai_help(rec["summary"])

        # column 18 (index 17) = prior AI rewrite notes, if present.
        rec["ainotes"] = str(row[17]).strip() if len(row) > 17 and row[17] not in (None, "") else ""

        fill = ws_style.cell(row=i, column=1).fill
        fg = fill.fgColor.rgb if fill and fill.fgColor else None
        rec["_flagged"] = (fg == FLAG_FILL_RGB) or bool(rec["_qa_notes"])
        out.append(rec)

    # Cross-row duplicate detection (see dup_signature() above) -- needs the
    # full list, so it runs as a second pass after every row is loaded.
    by_tsig, by_num = {}, {}
    for rec in out:
        tsig, num = dup_signature(rec)
        rec["_dupsig"] = (tsig, num)
        by_tsig.setdefault(tsig, []).append(rec)
        if num:
            by_num.setdefault(num, []).append(rec)
    for rec in out:
        tsig, num = rec["_dupsig"]
        others = [r for r in by_tsig.get(tsig, []) if r is not rec]
        for r in by_num.get(num, []):
            if r is not rec and r not in others:
                others.append(r)
        if others:
            names = "; ".join(f"row {r['_row']} “{r['title'][:50]}”" for r in others)
            rec["_qa_notes"].append(f"Possible duplicate of {names} — verify before including.")
            rec["_flagged"] = True

    return out


PAGE_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Grants Email — Review</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root {
    --navy: #041E42; --link: #4F5254; --flag: #FFF2CC; --ai: #E8F0FE;
    --border: #D8DCE1; --bg: #F7F8FA;
  }
  * { box-sizing: border-box; }
  body { font-family: -apple-system, Segoe UI, Arial, sans-serif; margin: 0; background: var(--bg); color: #1a1a1a; }
  header { background: var(--navy); color: #fff; padding: 16px 24px; position: sticky; top: 0; z-index: 10; }
  header h1 { margin: 0 0 4px 0; font-size: 20px; }
  header .sub { font-size: 13px; opacity: 0.85; }
  .toolbar { background: #fff; border-bottom: 1px solid var(--border); padding: 12px 24px;
    display: flex; flex-wrap: wrap; gap: 10px; align-items: center; position: sticky; top: 68px; z-index: 9; }
  .toolbar input[type=text] { padding: 7px 10px; border: 1px solid var(--border); border-radius: 5px; width: 240px; }
  .toolbar select { padding: 7px 8px; border: 1px solid var(--border); border-radius: 5px; }
  .toolbar label.chk { font-size: 13px; display: flex; align-items: center; gap: 4px; }
  .toolbar .spacer { flex: 1; }
  .toolbar button { padding: 8px 14px; border: none; border-radius: 5px; cursor: pointer; font-size: 13px; font-weight: 600; }
  #exportBtn { background: var(--navy); color: #fff; }
  #exportCsvBtn { background: #fff; color: var(--navy); border: 1px solid var(--navy) !important; }
  .stats { font-size: 12.5px; color: #555; padding: 8px 24px 0 24px; }
  #cards { padding: 16px 24px 60px 24px; display: flex; flex-direction: column; gap: 12px; max-width: 1100px; margin: 0 auto; }
  .card { background: #fff; border: 1px solid var(--border); border-radius: 8px; padding: 14px 18px; }
  .card.flagged { border-left: 5px solid #D9A400; background: linear-gradient(to right, var(--flag) 0%, #fff 2%); }
  .card.excluded { opacity: 0.55; }
  .card-top { display: flex; gap: 12px; align-items: flex-start; }
  .card-top .incl { display: flex; align-items: center; gap: 6px; padding-top: 3px; }
  .card-top .incl input { width: 18px; height: 18px; }
  .card-title { flex: 1; }
  .card-title .t { font-weight: 700; color: var(--navy); font-size: 15px; }
  .title-input { width: 100%; font-weight: 700; color: var(--navy); font-size: 15px; border: 1px solid transparent;
    border-radius: 4px; padding: 2px 4px; background: transparent; font-family: inherit; }
  .title-input:hover, .title-input:focus { border-color: var(--border); background: #fff; outline: none; }
  .meta { font-size: 12px; color: #666; margin-top: 2px; }
  .meta span.pill { display: inline-block; background: #EEF1F5; border-radius: 10px; padding: 1px 8px; margin-right: 6px; }
  .qa { margin-top: 6px; font-size: 12px; color: #8a5a00; background: var(--flag); border-radius: 4px; padding: 5px 8px; }
  .row2 { display: grid; grid-template-columns: 140px 1fr 160px; gap: 10px; margin-top: 10px; align-items: start; }
  .row2 label { font-size: 11px; font-weight: 700; text-transform: uppercase; color: #888; display: block; margin-bottom: 3px; }
  select, .deadline-input { width: 100%; padding: 6px 8px; border: 1px solid var(--border); border-radius: 4px; font-size: 13px; }
  textarea.summary { width: 100%; min-height: 90px; padding: 8px; border: 1px solid var(--border); border-radius: 4px;
    font-size: 13px; font-family: inherit; resize: vertical; }
  .ai-help { display: flex; align-items: center; gap: 6px; font-size: 12.5px; background: var(--ai); border-radius: 5px;
    padding: 8px; margin-top: 8px; }
  .ai-notes-wrap { background: var(--ai); border-radius: 0 0 5px 5px; margin-top: -1px; padding: 0 8px 8px 8px; }
  textarea.ai-notes { width: 100%; min-height: 44px; padding: 6px 8px; border: 1px solid var(--border);
    border-radius: 4px; font-size: 12.5px; font-family: inherit; resize: vertical; }
  .links { margin-top: 8px; font-size: 12.5px; }
  .links a { color: var(--link); margin-right: 12px; }
  details.ref { margin-top: 8px; font-size: 12px; }
  details.ref summary { cursor: pointer; color: #555; }
  details.ref .refbox { background: #FAFAFA; border: 1px solid var(--border); border-radius: 4px; padding: 8px; margin-top: 4px; }
  .empty { text-align: center; color: #888; padding: 40px; }
</style>
</head>
<body>
<header>
  <h1>Grants &amp; Sponsor Opportunities — Weekly Review</h1>
  <div class="sub">Source file: __SOURCE_NAME__ &nbsp;|&nbsp; Generated for review — edits here do not touch the original .xlsx until you export</div>
</header>
<div class="toolbar">
  <input type="text" id="searchBox" placeholder="Search title, agency, summary...">
  <select id="sectionFilter"><option value="">All sections</option></select>
  <label class="chk"><input type="checkbox" id="flaggedOnly"> Flagged only</label>
  <label class="chk"><input type="checkbox" id="aiOnly"> Needs AI help only</label>
  <label class="chk"><input type="checkbox" id="includedOnly"> Included only</label>
  <select id="sortBy">
    <option value="row">Original order</option>
    <option value="agency">Sort: Agency</option>
    <option value="deadline">Sort: Deadline (as text)</option>
    <option value="title">Sort: Title</option>
  </select>
  <div class="spacer"></div>
  <button id="exportBtn">Export validated_*.xlsx</button>
  <button id="exportCsvBtn">Export CSV (backup)</button>
</div>
<div class="stats" id="statsLine"></div>
<div id="cards"></div>

<script>
const DATA = __DATA_JSON__;
const TYPE_OPTIONS = __TYPE_OPTIONS_JSON__;
const SOURCE_NAME = __SOURCE_NAME_JSON__;

function decodeEntities(str) {
  const el = document.createElement('textarea');
  el.innerHTML = str || '';
  return el.value;
}
function escAttr(str) {
  return (str || '').replace(/"/g, '&quot;');
}

// populate section filter
const sections = [...new Set(DATA.map(r => r.abbr || '(none)'))].sort();
const sectionSel = document.getElementById('sectionFilter');
sections.forEach(s => {
  const o = document.createElement('option'); o.value = s; o.textContent = s;
  sectionSel.appendChild(o);
});

function typeOptionsHtml(current) {
  return TYPE_OPTIONS.map(t => `<option value="${escAttr(t)}" ${t===current?'selected':''}>${t || '(blank)'}</option>`).join('');
}

function cardHtml(rec) {
  const flaggedClass = rec._flagged ? 'flagged' : '';
  const exclClass = rec._email_checked ? '' : 'excluded';
  const qaHtml = (rec._qa_notes && rec._qa_notes.length)
    ? `<div class="qa">${rec._qa_notes.map(n => '⚠ ' + n).join('<br>')}</div>` : '';
  const link1 = rec.l1url ? `<a href="${escAttr(rec.l1url)}" target="_blank">${decodeEntities(rec.l1text) || 'Link 1'}</a>` : '';
  const link2 = rec.l2url ? `<a href="${escAttr(rec.l2url)}" target="_blank">${decodeEntities(rec.l2text) || 'Link 2'}</a>` : '';
  return `
  <div class="card ${flaggedClass} ${exclClass}" data-row="${rec._row}">
    <div class="card-top">
      <div class="incl">
        <input type="checkbox" class="f-email" ${rec._email_checked ? 'checked' : ''}>
      </div>
      <div class="card-title">
        <input type="text" class="title-input f-title" value="${escAttr(rec.title)}">
        <div class="meta">
          <span class="pill">${decodeEntities(rec.agency) || rec.abbr}</span>
          <span class="pill">${rec.abbr || ''}</span>
          ${rec.eligflag ? '<span class="pill" style="background:#FDECEA;color:#B3261E;">Elig Flag</span>' : ''}
        </div>
        ${qaHtml}
      </div>
    </div>
    <div class="row2">
      <div>
        <label>Type</label>
        <select class="f-type">${typeOptionsHtml(rec.type)}</select>
      </div>
      <div>
        <label>Deadline</label>
        <input type="text" class="deadline-input f-deadline" value="${escAttr(rec.deadline)}">
      </div>
      <div>
        <label>Record ID</label>
        <div style="font-size:12px;color:#666;padding-top:6px;">${decodeEntities(rec.recordid) || '—'}</div>
      </div>
    </div>
    <div style="margin-top:10px;">
      <label style="font-size:11px;font-weight:700;text-transform:uppercase;color:#888;">Summary</label>
      <textarea class="summary f-summary">${rec.summary}</textarea>
    </div>
    <label class="ai-help">
      <input type="checkbox" class="f-aihelp" ${rec._needs_ai_help ? 'checked' : ''}>
      Needs AI help rewriting this summary
    </label>
    <div class="ai-notes-wrap" style="display:${rec._needs_ai_help ? 'block' : 'none'};">
      <label style="font-size:11px;font-weight:700;text-transform:uppercase;color:#888;">Notes for the rewrite (e.g. "add funding amount", "strip HTML tags")</label>
      <textarea class="ai-notes f-ainotes" placeholder="Optional — leave specific instructions here">${rec.ainotes || ''}</textarea>
    </div>
    <div class="links">${link1} ${link2}</div>
    <details class="ref">
      <summary>Reference: recommendation / eligibility notes / LOI alert</summary>
      <div class="refbox">
        <div><strong>Recommendation:</strong> ${decodeEntities(rec.rec) || '—'}</div>
        <div><strong>Eligibility Notes:</strong> ${decodeEntities(rec.elignotes) || '—'}</div>
        <div><strong>LOI Alert:</strong> ${decodeEntities(rec.loi) || '—'}</div>
      </div>
    </details>
  </div>`;
}

function render() {
  const q = document.getElementById('searchBox').value.trim().toLowerCase();
  const sect = sectionSel.value;
  const flaggedOnly = document.getElementById('flaggedOnly').checked;
  const aiOnly = document.getElementById('aiOnly').checked;
  const includedOnly = document.getElementById('includedOnly').checked;
  const sortBy = document.getElementById('sortBy').value;

  let rows = DATA.filter(r => {
    if (sect && (r.abbr || '(none)') !== sect) return false;
    if (flaggedOnly && !r._flagged) return false;
    if (aiOnly && !r._needs_ai_help) return false;
    if (includedOnly && !r._email_checked) return false;
    if (q) {
      const hay = (r.title + ' ' + r.agency + ' ' + r.summary).toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });

  if (sortBy === 'agency') rows = rows.slice().sort((a,b) => (a.agency||'').localeCompare(b.agency||''));
  else if (sortBy === 'title') rows = rows.slice().sort((a,b) => (a.title||'').localeCompare(b.title||''));
  else if (sortBy === 'deadline') rows = rows.slice().sort((a,b) => (a.deadline||'').localeCompare(b.deadline||''));

  const container = document.getElementById('cards');
  container.innerHTML = rows.length ? rows.map(cardHtml).join('') : '<div class="empty">No records match the current filters.</div>';

  // wire up inputs to sync back to DATA by row id
  container.querySelectorAll('.card').forEach(card => {
    const row = parseInt(card.dataset.row, 10);
    const rec = DATA.find(r => r._row === row);
    card.querySelector('.f-email').addEventListener('change', e => {
      rec._email_checked = e.target.checked;
      card.classList.toggle('excluded', !rec._email_checked);
      updateStats();
    });
    card.querySelector('.f-title').addEventListener('input', e => { rec.title = e.target.value; });
    card.querySelector('.f-type').addEventListener('change', e => { rec.type = e.target.value; });
    card.querySelector('.f-deadline').addEventListener('input', e => { rec.deadline = e.target.value; });
    card.querySelector('.f-summary').addEventListener('input', e => { rec.summary = e.target.value; });
    card.querySelector('.f-aihelp').addEventListener('change', e => {
      rec._needs_ai_help = e.target.checked;
      card.querySelector('.ai-notes-wrap').style.display = e.target.checked ? 'block' : 'none';
      updateStats();
    });
    card.querySelector('.f-ainotes').addEventListener('input', e => { rec.ainotes = e.target.value; });
  });
  updateStats();
}

function updateStats() {
  const total = DATA.length;
  const included = DATA.filter(r => r._email_checked).length;
  const flagged = DATA.filter(r => r._flagged).length;
  const aiHelp = DATA.filter(r => r._needs_ai_help).length;
  document.getElementById('statsLine').textContent =
    `${total} total records | ${included} marked Include | ${flagged} flagged for review | ${aiHelp} need AI summary help`;
}

['searchBox','sectionFilter','sortBy'].forEach(id => {
  document.getElementById(id).addEventListener('input', render);
});
['flaggedOnly','aiOnly','includedOnly'].forEach(id => {
  document.getElementById(id).addEventListener('change', render);
});

function buildExportRows() {
  // exact 16-col order + 17th col (Q) for AI-help tracking, safely ignored by build_html.py
  const headers = ["Email?","Eligibility Flag","Type","Recommendation","LOI Alert",
    "Funding Agency","Abbreviation","Title","Summary","Deadline",
    "Link1 URL","Link1 Text","Link2 URL","Link2 Text","Eligibility Notes","Record ID",
    "Needs AI Summary Help","AI Rewrite Notes"];
  const rows = [headers];
  // Only rows marked "Include in email" are exported — this is purely to
  // shrink the file for downstream agent reliability. build_html.py already
  // ignores Email?!=Y rows regardless, so this does not change what ends up
  // in the actual email either way.
  DATA.filter(r => r._email_checked).forEach(r => {
    rows.push([
      "Y",
      r.eligflag || "",
      r.type || "",
      r.rec || "",
      r.loi || "",
      r.agency || "",
      r.abbr || "",
      r.title || "",
      r.summary || "",
      r.deadline || "",
      r.l1url || "",
      r.l1text || "",
      r.l2url || "",
      r.l2text || "",
      r.elignotes || "",
      r.recordid || "",
      r._needs_ai_help ? "Y" : "",
      r.ainotes || "",
    ]);
  });
  return rows;
}

document.getElementById('exportBtn').addEventListener('click', () => {
  const rows = buildExportRows();
  if (rows.length <= 1) {
    alert('No rows are checked "Include in email" — nothing to export. Check at least one row first.');
    return;
  }
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [7,13,10,40,12,28,7,40,50,18,30,18,30,10,40,24,18,40].map(w => ({wch:w}));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Grant Opportunities");
  const stamp = new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, `validated_${SOURCE_NAME.replace(/\.xlsx$/i,'')}_${stamp}.xlsx`);
});

document.getElementById('exportCsvBtn').addEventListener('click', () => {
  const rows = buildExportRows();
  const csv = rows.map(r => r.map(v => {
    const s = String(v).replace(/"/g, '""');
    return /[",\n]/.test(s) ? `"${s}"` : s;
  }).join(',')).join('\r\n');
  const blob = new Blob([csv], {type: 'text/csv'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  const stamp = new Date().toISOString().slice(0,10);
  a.download = `validated_${SOURCE_NAME.replace(/\.xlsx$/i,'')}_${stamp}.csv`;
  a.click();
});

render();
</script>
</body>
</html>
"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("infile")
    ap.add_argument("--out", default=None)
    args = ap.parse_args()

    rows = read_rows(args.infile)
    if not rows:
        raise SystemExit("No data rows found in the staging file.")

    out = args.out or "review.html"
    page = PAGE_TEMPLATE
    page = page.replace("__SOURCE_NAME__", html.escape(args.infile.split("/")[-1]))
    page = page.replace("__DATA_JSON__", json.dumps(rows, ensure_ascii=False))
    page = page.replace("__TYPE_OPTIONS_JSON__", json.dumps(TYPE_OPTIONS))
    page = page.replace("__SOURCE_NAME_JSON__", json.dumps(args.infile.split("/")[-1].replace(".xlsx", "")))

    with open(out, "w", encoding="utf-8") as f:
        f.write(page)

    flagged = sum(1 for r in rows if r["_flagged"])
    ai_help = sum(1 for r in rows if r["_needs_ai_help"])
    print(f"Wrote {out}")
    print(f"  {len(rows)} records | {flagged} flagged for review | {ai_help} pre-marked as needing AI summary help")


if __name__ == "__main__":
    main()
