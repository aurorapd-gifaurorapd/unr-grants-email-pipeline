#!/usr/bin/env python3
"""
run_week.py
===========
One command per step of the weekly grants email, for running from PowerShell.

It runs the same scripts you would run by hand (grants_collect.py,
merge_and_assess.py, build_html.py, ...). Those scripts are not changed, so
if run_week.py ever misbehaves you can go back to the commands in
"HOW TO RUN - Weekly Grants Email.docx".

It never deletes anything. When a step would replace a file, the old copy is
moved into a "backup" folder inside the week folder first.

    py run_week.py start          make this week's folder, work out the lookback
    py run_week.py check          check the Pivot export (add --sam to test the SAM key, 1 call)
    py run_week.py collect        run every collector once (asks before SAM.gov)
    py run_week.py merge          make a fresh staging file and open the review page
    py run_week.py finalize       bring in the review export, make the _final copy
    py run_week.py add FILE       add non-grant rows from FILE to _final
    py run_week.py check-final    run every check the email build will run
    py run_week.py build          build the email and open it
    py run_week.py status         what's done and what to do next

Every command takes --week Grant_MM_DD_YYYY to work on a folder other than
the newest one. Progress is saved in week.json inside the week folder, so you
can close PowerShell and pick up later.
"""
import argparse, csv, datetime, glob, json, os, re, shutil, statistics, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
ARCHIVE = os.path.join(os.path.dirname(HERE), "Grants Email weekly folder archive")
DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")
STAGING = "Grant_Opportunities_Staging.xlsx"
SEEN_FILES = ["federal_seen.json", "nsf_seen.json", "nih_seen.json"]
LOOKBACK_MARGIN = 3          # extra days on top of the gap since the last email

# Collector defaults (the windows built into each script).
DEFAULT_DAYS = {"federal": 30, "nsf": 30, "nih": 30, "sam": 14}

FOLDER_RE = re.compile(r"^Grant_(\d{2})_(\d{2})_(\d{2}|\d{4})$")
STEPS = ["start", "check", "collect", "merge", "finalize", "check-final", "build"]


# ── small helpers ────────────────────────────────────────────────────────────
def say(msg=""):
    print(msg, flush=True)


def stop(problem, steps=()):
    say("")
    say("=" * 70)
    say(f"STOPPED: {problem}")
    if steps:
        say("\nWhat to do:")
        for s in steps:
            say(f"  - {s}")
    say("=" * 70)
    sys.exit(2)


def folder_date(name):
    m = FOLDER_RE.match(name)
    if not m:
        return None
    mm, dd, yy = m.groups()
    year = int(yy) + (2000 if len(yy) == 2 else 0)
    try:
        return datetime.date(year, int(mm), int(dd))
    except ValueError:
        return None


def week_folders():
    """(date, path) for every Grant_* week folder here and in the archive, oldest first."""
    found = []
    for base in (HERE, ARCHIVE):
        if not os.path.isdir(base):
            continue
        for name in os.listdir(base):
            p = os.path.join(base, name)
            d = folder_date(name)
            if d and os.path.isdir(p):
                found.append((d, p))
    return sorted(found)


def sent_email_in(path):
    return bool(glob.glob(os.path.join(path, "Grant_Opportunities*.html")))


def resolve_week(arg):
    if arg:
        p = arg if os.path.isabs(arg) else os.path.join(HERE, arg)
        if not os.path.isdir(p):
            stop(f"Week folder not found: {arg}")
        return p
    here = [(d, p) for d, p in week_folders() if os.path.dirname(p) == HERE]
    if not here:
        stop("No Grant_MM_DD_YYYY folder in the pipeline folder.",
             ["Run:  py run_week.py start"])
    return here[-1][1]


def state_path(week):
    return os.path.join(week, "week.json")


def load_state(week):
    try:
        with open(state_path(week), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {"steps": {}, "collectors": {}}


def save_state(week, st):
    with open(state_path(week), "w", encoding="utf-8") as f:
        json.dump(st, f, indent=2)


def mark(week, st, step, **info):
    st.setdefault("steps", {})[step] = {"done": now(), **info}
    save_state(week, st)


def now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M")


def backup(week, path):
    """Move path into WEEK/backup/ with a time stamp. Returns the new path."""
    if not os.path.exists(path):
        return None
    bdir = os.path.join(week, "backup")
    os.makedirs(bdir, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(path))
    dest = os.path.join(bdir, f"{stem}_{datetime.datetime.now():%Y%m%d_%H%M%S}{ext}")
    shutil.move(path, dest)
    return dest


def run(args, env=None):
    """Run one pipeline script from the pipeline folder, output shown live."""
    e = dict(os.environ, PYTHONIOENCODING="utf-8", **(env or {}))
    say(f"\n> py {' '.join(os.path.relpath(a, HERE) if os.path.isabs(a) else a for a in args)}")
    return subprocess.run([sys.executable, *args], cwd=HERE, env=e).returncode


def open_file(path):
    try:
        os.startfile(path)                      # Windows
    except AttributeError:
        subprocess.run(["open" if sys.platform == "darwin" else "xdg-open", path])
    except OSError as e:
        say(f"  (could not open it automatically: {e})")


def count_json(path):
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        return len(data) if isinstance(data, list) else None
    except (OSError, ValueError):
        return None


def pivot_rows(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = [r for r in csv.reader(f) if any(c.strip() for c in r)]
    return rows[0] if rows else [], max(len(rows) - 1, 0)


def sam_key():
    """SAM_API_KEY from this shell, or from the saved user setting (setx only
    reaches new windows, so a window opened before setx won't have it)."""
    k = os.environ.get("SAM_API_KEY")
    if k:
        return k
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as key:
            return winreg.QueryValueEx(key, "SAM_API_KEY")[0]
    except (ImportError, OSError):
        return None


def final_files(week):
    return sorted(glob.glob(os.path.join(week, "validated_*_final.xlsx")), key=os.path.getmtime)


# ── start ────────────────────────────────────────────────────────────────────
def cmd_start(a):
    today = datetime.date.today()
    name = f"Grant_{today:%m_%d_%Y}"
    week = os.path.join(HERE, a.week or name)
    made = not os.path.isdir(week)
    os.makedirs(week, exist_ok=True)

    sent = [(d, p) for d, p in week_folders()
            if os.path.normcase(p) != os.path.normcase(week) and d <= today and sent_email_in(p)]
    st = load_state(week)
    say(f"Week folder: {os.path.relpath(week, HERE)}  ({'created' if made else 'already there'})")
    if sent:
        last_d, last_p = sent[-1]
        gap = (today - last_d).days
        days = max(DEFAULT_DAYS["federal"], gap + LOOKBACK_MARGIN)
        say(f"Last email: {last_d:%B %d, %Y}  ({os.path.relpath(last_p, os.path.dirname(HERE))})")
        say(f"Gap: {gap} days.  Federal, NSF and NIH will look back {days} days.")
        if gap + LOOKBACK_MARGIN > DEFAULT_DAYS["sam"]:
            say(f"NOTE: SAM.gov only looks back {DEFAULT_DAYS['sam']} days, which is less than the gap "
                f"plus margin ({gap + LOOKBACK_MARGIN}). SAM items posted before "
                f"{today - datetime.timedelta(days=DEFAULT_DAYS['sam']):%B %d} won't be found.")
        st.update(last_email=last_d.isoformat(), gap_days=gap, lookback_days=days)
    else:
        say("Couldn't find a previous week folder with a built email; using the default 30-day lookback.")
        st.update(last_email=None, gap_days=None, lookback_days=DEFAULT_DAYS["federal"])
    mark(week, st, "start")

    if not os.path.exists(os.path.join(week, "opps_export.csv")):
        say(f"\nNext: export from Pivot and save it as opps_export.csv in {os.path.relpath(week, HERE)},")
        say("then run:  py run_week.py check")
    else:
        say("\nNext:  py run_week.py check")


# ── check ────────────────────────────────────────────────────────────────────
def cmd_check(a):
    week = resolve_week(a.week)
    st = load_state(week)
    problems = []

    for mod in ("openpyxl", "pdfplumber", "docx"):
        try:
            __import__(mod)
        except ImportError:
            problems.append(f"Python package missing: {mod}  (fix: py -m pip install -r requirements.txt)")

    csv_path = os.path.join(week, "opps_export.csv")
    if not os.path.exists(csv_path):
        others = glob.glob(os.path.join(week, "*.csv"))
        problems.append("No opps_export.csv in the week folder"
                        + (f" (found {', '.join(os.path.basename(o) for o in others)}; rename it)" if others else ""))
    else:
        head, n = pivot_rows(csv_path)
        missing = [c for c in ("Opportunity ID", "Title", "Deadline", "Funder") if c not in head]
        if missing:
            problems.append(f"opps_export.csv is missing columns {missing}; is it a Pivot opportunity export?")
        past = []
        for d, p in week_folders()[::-1]:
            f = os.path.join(p, "opps_export.csv")
            if os.path.normcase(p) != os.path.normcase(week) and os.path.exists(f):
                past.append(pivot_rows(f)[1])
            if len(past) == 6:
                break
        typical = int(statistics.median(past)) if past else None
        say(f"Pivot export: {n} rows" + (f"  (recent weeks: about {typical})" if typical else ""))
        if typical and n < typical * 0.5:
            problems.append(f"The Pivot export has {n} rows; recent weeks had about {typical}. "
                            "It may be a partial export. Re-export from Pivot, or say so and carry on.")
        st["pivot_rows"] = n

    key = sam_key()
    if not key:
        say("SAM.gov key: not set. SAM.gov will be skipped unless you set it "
            '(PowerShell: setx SAM_API_KEY "your-key", then open a new window).')
    elif a.sam:
        say("SAM.gov key: testing with one search call...")
        say("  " + sam_probe(key))
    else:
        say("SAM.gov key: set (not tested; add --sam to test it with 1 of today's ~10 calls).")

    say("Seen-lists: " + ", ".join(f"{s} {'ok' if os.path.exists(os.path.join(HERE, s)) else 'MISSING'}"
                                    for s in SEEN_FILES))
    for s in SEEN_FILES:
        if not os.path.exists(os.path.join(HERE, s)):
            problems.append(f"{s} is missing from the pipeline folder. Don't run collect until it's restored "
                            "(look in the last week folder's seen_backup).")

    if problems:
        say("\nPROBLEMS:")
        for p in problems:
            say(f"  - {p}")
        save_state(week, st)
        sys.exit(1)
    mark(week, st, "check")
    say("\nAll checks passed.  Next:  py run_week.py collect")


def sam_probe(key):
    import urllib.request, urllib.error, urllib.parse
    today = datetime.date.today()
    q = {"postedFrom": f"{today - datetime.timedelta(days=7):%m/%d/%Y}", "postedTo": f"{today:%m/%d/%Y}",
         "limit": 1, "api_key": key}
    try:
        with urllib.request.urlopen("https://api.sam.gov/opportunities/v2/search?" + urllib.parse.urlencode(q),
                                    timeout=30) as r:
            r.read()
        return "accepted."
    except urllib.error.HTTPError as e:
        body = e.read().decode(errors="replace")
        if e.code in (401, 403) or "API_KEY_INVALID" in body:
            return "REJECTED (expired or wrong). Get a new key at sam.gov and setx SAM_API_KEY."
        if e.code == 429 or "OVER_RATE_LIMIT" in body:
            return "today's call limit is used up. Run collect tomorrow, or skip SAM.gov this week."
        return f"HTTP {e.code}; try again later."
    except urllib.error.URLError as e:
        return f"no answer from SAM.gov ({e.reason})."


# ── collect ──────────────────────────────────────────────────────────────────
def collectors(week, days):
    w = lambda f: os.path.join(week, f)
    return [
        ("federal", "federal_opportunities.json",
         ["grants_collect.py", "--out", w("federal_opportunities.json"), "--days", str(days)]),
        ("pivot", "pivot_opportunities.json",
         ["pivot_collect.py", w("opps_export.csv"), "--out", w("pivot_opportunities.json"),
          "--skipped-out", w("pivot_skipped.json")]),
        ("nih", "nih_opportunities.json",
         ["nih_rss_collect.py", "--out", w("nih_opportunities.json"), "--days", str(days)]),
        ("nsf", "nsf_opportunities.json",
         ["nsf_rss_collect.py", "--out", w("nsf_opportunities.json"), "--days", str(days)]),
        ("sam", "sam_opportunities.json",
         ["sam_collect.py", "--out", w("sam_opportunities.json"), "--skipped-out", w("sam_skipped.json")]),
    ]


def cmd_collect(a):
    week = resolve_week(a.week)
    st = load_state(week)
    if "start" not in st.get("steps", {}):
        stop("This week hasn't been started.", ["Run:  py run_week.py start"])
    days = a.days or st.get("lookback_days") or DEFAULT_DAYS["federal"]
    only = set(a.only.split(",")) if a.only else None
    redo = set(a.redo.split(",")) if a.redo else set()
    plan = [c for c in collectors(week, days) if not only or c[0] in only]

    # Saving the seen-lists first means a bad run can be undone by copying them back.
    if not a.dry_run:
        sdir = os.path.join(week, "seen_backup")
        os.makedirs(sdir, exist_ok=True)
        for s in SEEN_FILES:
            src = os.path.join(HERE, s)
            if os.path.exists(src) and not os.path.exists(os.path.join(sdir, s)):
                shutil.copy2(src, sdir)
        say(f"Seen-lists saved to {os.path.relpath(sdir, HERE)} (only the first time this week).")

    results = st.setdefault("collectors", {})
    for name, outfile, args in plan:
        out = os.path.join(week, outfile)
        say(f"\n--- {name} ---")
        # A second run the same week returns 0 (the seen-list already has
        # everything) and would overwrite the real output with an empty file.
        if os.path.exists(out) and name not in redo:
            say(f"Already collected ({count_json(out)} records in {outfile}). Skipping.")
            say(f"  To run it again anyway: py run_week.py collect --only {name} --redo {name}"
                + ("  (it will likely return 0 because of the seen-list)" if name in ("federal", "nsf", "nih") else ""))
            continue
        if name == "pivot" and not os.path.exists(os.path.join(week, "opps_export.csv")):
            say("No opps_export.csv; skipping Pivot.")
            results[name] = {"ran": now(), "ok": False, "note": "no export"}
            continue
        env = None
        if name == "sam":
            key = sam_key()
            if not key:
                say("SAM_API_KEY isn't set; skipping SAM.gov this week.")
                results[name] = {"ran": now(), "ok": False, "note": "no key"}
                continue
            gap = st.get("gap_days")
            if gap and gap + LOOKBACK_MARGIN > DEFAULT_DAYS["sam"]:
                say(f"NOTE: SAM.gov looks back only {DEFAULT_DAYS['sam']} days; the gap since the last email is {gap}.")
            if not a.yes and not a.dry_run:
                ans = input("SAM.gov uses up to 9 of your ~10 calls a day. Run it now? [y/N] ").strip().lower()
                if ans not in ("y", "yes"):
                    say("Skipped SAM.gov.")
                    results[name] = {"ran": now(), "ok": False, "note": "skipped by you"}
                    continue
            env = {"SAM_API_KEY": key}
        if a.dry_run:
            say("(dry run) would run: py " + " ".join(os.path.relpath(x, HERE) if os.path.isabs(x) else x for x in args))
            continue
        if name in redo:
            moved = backup(week, out)
            if moved:
                say(f"Old output moved to {os.path.relpath(moved, HERE)}")
        code = run(args, env)
        n = count_json(out)
        ok = code == 0 and n is not None
        results[name] = {"ran": now(), "ok": ok, "records": n, "exit": code}
        save_state(week, st)
        say(f"{name}: {'OK' if ok else 'FAILED'}" + (f", {n} records" if n is not None else ""))
        if ok and n == 0:
            say(f"  WARNING: {name} returned 0 records. Look at the messages above before merging.")

    if a.dry_run:
        return
    say("\nSummary:")
    for name, _o, _a in collectors(week, days):
        r = results.get(name)
        if r:
            say(f"  {name:8} {'OK    ' if r.get('ok') else 'NOT OK'} {r.get('records', '')} {r.get('note', '')}")
        elif os.path.exists(os.path.join(week, _o)):
            say(f"  {name:8} OK     {count_json(os.path.join(week, _o))} (collected earlier)")
    mark(week, st, "collect", lookback_days=days)
    say("\nNext:  py run_week.py merge")


# ── merge ────────────────────────────────────────────────────────────────────
def cmd_merge(a):
    week = resolve_week(a.week)
    st = load_state(week)
    inputs = sorted(glob.glob(os.path.join(week, "*_opportunities.json")))
    if not inputs:
        stop("No collector outputs (*_opportunities.json) in the week folder.", ["Run:  py run_week.py collect"])
    staging = os.path.join(week, STAGING)
    # merge_and_assess.py appends to an existing staging file, so a re-merge
    # would double every row. Start fresh, keeping the old one in backup/.
    moved = backup(week, staging)
    if moved:
        say(f"Old staging file moved to {os.path.relpath(moved, HERE)}")
    say(f"Merging {len(inputs)} files: " + ", ".join(os.path.basename(i) for i in inputs))
    if run(["merge_and_assess.py", *inputs, "--out", staging]) != 0 or not os.path.exists(staging):
        stop("The merge failed; see the messages above.")
    review = os.path.join(week, "review.html")
    if os.path.exists(review):
        backup(week, review)
    if run(["build_review_html.py", staging, "--out", review]) != 0:
        stop("Building the review page failed; see the messages above.")
    mark(week, st, "merge", inputs=[os.path.basename(i) for i in inputs])
    say(f"\nReview page: {os.path.relpath(review, HERE)}")
    if not a.no_open:
        open_file(review)
    say("Review it in your browser and export. Then run:  py run_week.py finalize")


# ── finalize ─────────────────────────────────────────────────────────────────
def cmd_finalize(a):
    week = resolve_week(a.week)
    st = load_state(week)
    if a.file:
        export = a.file
    else:
        cands = [p for d in (DOWNLOADS, week) for p in glob.glob(os.path.join(d, "validated_*.xlsx"))
                 if not p.endswith("_final.xlsx") and not os.path.basename(p).startswith("~$")]
        merged = st.get("steps", {}).get("merge", {}).get("done")
        if not cands:
            stop("Couldn't find the review export (validated_*.xlsx) in Downloads or the week folder.",
                 ["Export from the review page, or give the path:  py run_week.py finalize --file PATH"])
        export = max(cands, key=os.path.getmtime)
        when = datetime.datetime.fromtimestamp(os.path.getmtime(export)).strftime("%Y-%m-%d %H:%M")
        say(f"Newest export: {export}  (saved {when})")
        if merged and when < merged:
            stop("That export is older than this week's merge, so it's probably last week's.",
                 ["Export again from this week's review page, or pass it with --file PATH."])
    if not os.path.exists(export):
        stop(f"File not found: {export}")

    dest = os.path.join(week, os.path.basename(export))
    if os.path.normcase(os.path.abspath(export)) != os.path.normcase(dest):
        if os.path.exists(dest):
            backup(week, dest)
        shutil.copy2(export, dest)               # copy, not move: Downloads keeps its copy
        say(f"Copied into the week folder: {os.path.basename(dest)}")
    final = dest[:-5] + "_final.xlsx"
    if os.path.exists(final):
        if not a.replace:
            stop(f"{os.path.basename(final)} already exists, and may have edits in it.",
                 ["To keep working on it, skip this step.",
                  "To start it over from the export: py run_week.py finalize --replace (the old one goes to backup/)."])
        say(f"Old _final moved to {os.path.relpath(backup(week, final), HERE)}")
    shutil.copy2(dest, final)
    mark(week, st, "finalize", export=os.path.basename(dest), final=os.path.basename(final))
    say(f"Made {os.path.basename(final)}. Every edit from here on goes into this file.")
    say("\nNext: non-grant additions (py run_week.py add FILE) and summary rewrites, then:  py run_week.py check-final")


# ── add ──────────────────────────────────────────────────────────────────────
def cmd_add(a):
    import openpyxl
    week = resolve_week(a.week)
    st = load_state(week)
    finals = final_files(week)
    if not finals:
        stop("No validated_*_final.xlsx in the week folder.", ["Run:  py run_week.py finalize"])
    final = finals[-1]
    if not os.path.exists(a.file):
        stop(f"File not found: {a.file}")

    src = openpyxl.load_workbook(a.file).worksheets[0]
    head = [str(c.value or "").strip() for c in src[1]]
    if head[:3] != ["Email?", "Eligibility Flag", "Type"] or "Record ID" not in head:
        stop(f"{os.path.basename(a.file)} doesn't have the 16 staging columns (Email?, Eligibility Flag, Type, ...).")
    new_rows = [list(r[:16]) for r in src.iter_rows(min_row=2, values_only=True) if r and any(r)]

    wb = openpyxl.load_workbook(final)
    ws = wb.worksheets[0]
    fhead = [str(c.value or "").strip() for c in ws[1]]
    if fhead[:16] != head[:16]:
        stop("The columns in the two files don't line up; add these rows by hand.")
    have_ids = {str(r[15]).strip().lower() for r in ws.iter_rows(min_row=2, values_only=True) if r[15]}
    have_titles = {str(r[7]).strip().lower() for r in ws.iter_rows(min_row=2, values_only=True) if r[7]}

    bdir = os.path.join(week, "backup")
    os.makedirs(bdir, exist_ok=True)
    saved = shutil.copy2(final, os.path.join(
        bdir, f"{os.path.basename(final)[:-5]}_before_add_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx"))
    added = 0
    for r in new_rows:
        rid, title = str(r[15] or "").strip().lower(), str(r[7] or "").strip().lower()
        if (rid and rid in have_ids) or (title and title in have_titles):
            say(f"  already in _final, not added: {r[7]}")
            continue
        if not str(r[0] or "").strip():
            r[0] = "Y"                            # rows added to _final are meant for the email
        ws.append(r + [None] * (len(fhead) - 16))
        have_ids.add(rid); have_titles.add(title)
        added += 1
        say(f"  added: [{r[2]}] {r[7]}")
    wb.save(final)
    say(f"\nAdded {added} row(s) to {os.path.basename(final)} (Email? set to Y where it was blank).")
    say(f"Copy before adding: {os.path.relpath(saved, HERE)}")
    st.setdefault("additions", []).append({"file": os.path.basename(a.file), "added": added, "when": now()})
    save_state(week, st)
    cmd_check_final(a)


# ── check-final ──────────────────────────────────────────────────────────────
def cmd_check_final(a):
    week = resolve_week(a.week)
    st = load_state(week)
    finals = final_files(week)
    if not finals:
        stop("No validated_*_final.xlsx in the week folder.", ["Run:  py run_week.py finalize"])
    final = finals[-1]
    say(f"\nChecking {os.path.basename(final)}")
    code = run(["validate_staging.py", final])

    # The two checks build_html.py stops on, run here so they show up before the build.
    sys.path.insert(0, HERE)
    import build_html
    rows = build_html.read_rows(final)
    gate = []
    for check in (build_html.check_no_duplicates, build_html.check_deadlines):
        try:
            check(rows)
        except SystemExit as e:
            gate.append(str(e))
    say(f"\nEmail build checks on {len(rows)} rows marked Email?=Y:")
    for g in gate:
        say(g)
    if not gate:
        say("  duplicates: none   deadlines: all readable")
    if code != 0 or gate:
        say("\nFix the problems above in the _final file, then run check-final again.")
        sys.exit(1)
    mark(week, st, "check-final", rows=len(rows))
    say("\nReady.  Next:  py run_week.py build")


# ── build ────────────────────────────────────────────────────────────────────
def cmd_build(a):
    week = resolve_week(a.week)
    st = load_state(week)
    if a.email_date:
        ed = datetime.datetime.strptime(a.email_date, "%Y-%m-%d").date()
    else:
        t = datetime.date.today()
        ed = t - datetime.timedelta(days=t.weekday())            # Monday of this week, as build_html does
    out = os.path.join(week, f"Grant_Opportunities_{ed:%m-%d-%Y}.html")
    if os.path.exists(out):
        say(f"Old email moved to {os.path.relpath(backup(week, out), HERE)}")
    if run(["build_html.py", week, "--email-date", ed.isoformat(), "--out", out]) != 0:
        stop("The build stopped; see the messages above.", ["Run:  py run_week.py check-final  to see every problem."])
    mark(week, st, "build", email=os.path.basename(out), email_date=ed.isoformat())
    say(f"\nEmail: {os.path.relpath(out, HERE)}")
    if not a.no_open:
        open_file(out)
    say("Proofread it in the browser (titles, deadlines, links), then paste it into Emma.")


# ── status ───────────────────────────────────────────────────────────────────
NEXT_HINT = {
    "start": "py run_week.py start",
    "check": "py run_week.py check",
    "collect": "py run_week.py collect",
    "merge": "py run_week.py merge",
    "finalize": "review the page, export, then: py run_week.py finalize",
    "check-final": "additions and summary rewrites into _final, then: py run_week.py check-final",
    "build": "py run_week.py build",
}


def cmd_status(a):
    week = resolve_week(a.week)
    st = load_state(week)
    steps = st.get("steps", {})
    say(f"Week folder: {os.path.relpath(week, HERE)}")
    if st.get("last_email"):
        say(f"Last email: {st['last_email']}  (gap {st.get('gap_days')} days, lookback {st.get('lookback_days')})")
    for s in STEPS:
        info = steps.get(s)
        say(f"  [{'x' if info else ' '}] {s:12} {info['done'] if info else ''}")
    if st.get("collectors"):
        say("Collectors: " + ", ".join(f"{k} {'ok' if v.get('ok') else 'not ok'}"
                                         + (f" ({v['records']})" if v.get("records") is not None else "")
                                         for k, v in st["collectors"].items()))
    last = max((i for i, s in enumerate(STEPS) if s in steps), default=-1)
    nxt = STEPS[last + 1] if last + 1 < len(STEPS) else None
    say(f"\nNext: {NEXT_HINT[nxt]}" if nxt else "\nAll steps done. Move the week folder to the archive after sending.")


# ── main ─────────────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(description="Weekly grants email, one step at a time.",
                                 formatter_class=argparse.RawDescriptionHelpFormatter, epilog=__doc__)
    sub = ap.add_subparsers(dest="cmd", required=True)

    def add(name, fn, help_):
        p = sub.add_parser(name, help=help_)
        p.add_argument("--week", help="week folder (default: the newest Grant_* folder here)")
        p.set_defaults(fn=fn)
        return p

    add("start", cmd_start, "make this week's folder and work out the lookback")
    p = add("check", cmd_check, "check the Pivot export, packages and seen-lists")
    p.add_argument("--sam", action="store_true", help="also test the SAM key (uses 1 call)")
    p = add("collect", cmd_collect, "run every collector once")
    p.add_argument("--days", type=int, help="lookback for federal, NSF and NIH (default: from start)")
    p.add_argument("--only", help="comma list: federal,pivot,nih,nsf,sam")
    p.add_argument("--redo", help="comma list of collectors to run again even if already collected")
    p.add_argument("--yes", action="store_true", help="don't ask before SAM.gov")
    p.add_argument("--dry-run", action="store_true", help="show the commands without running them")
    p = add("merge", cmd_merge, "make a fresh staging file and the review page")
    p.add_argument("--no-open", action="store_true")
    p = add("finalize", cmd_finalize, "bring in the review export and make _final")
    p.add_argument("--file", help="the export, if it isn't the newest validated_*.xlsx in Downloads")
    p.add_argument("--replace", action="store_true", help="start _final over (old one goes to backup/)")
    p = add("add", cmd_add, "add non-grant rows to _final")
    p.add_argument("file", help="xlsx with the 16 staging columns, e.g. non_grant_additions_YYYY-MM-DD.xlsx")
    add("check-final", cmd_check_final, "run every check the email build runs")
    p = add("build", cmd_build, "build the email and open it")
    p.add_argument("--email-date", help="YYYY-MM-DD (default: Monday of this week)")
    p.add_argument("--no-open", action="store_true")
    add("status", cmd_status, "what's done and what's next")

    a = ap.parse_args()
    a.fn(a)


if __name__ == "__main__":
    main()
