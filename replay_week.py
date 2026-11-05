#!/usr/bin/env python3
"""
replay_week.py
==============
Re-runs the merge on a past week's saved collector outputs and compares the
result with that week's staging file. Use it to test any change to the merge,
dedup or deadline logic before trusting it on a live Thursday.

Nothing in the week folder is changed: the replay writes to a temporary file.
No collector is run, so no API calls are made and no seen-list is touched.

    python replay_week.py Grant_11_05_2026
    python replay_week.py Grant_10_22_2026 --against "Grant_10_22_2026\\Grant_Opportunities_Staging (1).xlsx"
    python replay_week.py Grant_11_05_2026 --merge-script "OLD Versions\\scripts replaced 2026-11-05\\merge_and_assess.py"

--merge-script runs a different copy of merge_and_assess.py (for example an
older version) so you can see exactly what a code change does.

Reading the comparison:
- Rows "only in original" with no Record ID are usually rows added by hand to
  that week's staging file (non-grant additions). The collectors never had them.
- Deadlines on records with several dates (mostly Pivot) are chosen as the first
  date on or after *today*, so replaying an old week can shift them. To compare
  two versions of the code fairly, replay both on the same day with
  --merge-script, rather than comparing against an old staging file.
"""
import argparse, glob, os, subprocess, sys, tempfile
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))


def load(path):
    ws = openpyxl.load_workbook(path).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    head = [str(c) for c in rows[0]]
    out = {}
    for r in rows[1:]:
        rec = dict(zip(head, r))
        key = rec.get("Record ID") or rec.get("Title")
        if key:
            out[str(key)] = rec
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("week", help="week folder, e.g. Grant_11_05_2026")
    ap.add_argument("--against", help="staging file to compare with (default: the week's Grant_Opportunities_Staging*.xlsx)")
    ap.add_argument("--merge-script", default=os.path.join(HERE, "merge_and_assess.py"))
    ap.add_argument("--keep", help="also save the replayed staging file to this path")
    args = ap.parse_args()

    inputs = sorted(glob.glob(os.path.join(args.week, "*_opportunities.json")))
    if not inputs:
        sys.exit(f"No *_opportunities.json files in {args.week}")
    against = args.against
    if not against:
        cands = sorted(glob.glob(os.path.join(args.week, "Grant_Opportunities_Staging*.xlsx")))
        against = cands[0] if cands else None

    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "replay.xlsx")
        env = dict(os.environ, PYTHONPATH=HERE + os.pathsep + os.environ.get("PYTHONPATH", ""))
        res = subprocess.run([sys.executable, args.merge_script, *inputs, "--out", out],
                             capture_output=True, text=True, encoding="utf-8", errors="replace", env=env, cwd=HERE)
        done = [l for l in res.stdout.splitlines() if l.startswith("DONE")]
        print(f"Replayed {len(inputs)} collector files with {os.path.relpath(args.merge_script, HERE)}")
        print("  " + (done[0] if done else "merge produced no DONE line"))
        if res.returncode != 0 or not os.path.exists(out):
            print(res.stdout[-2000:], res.stderr[-2000:])
            sys.exit("Merge failed; see output above.")
        new = load(out)
        if args.keep:
            import shutil
            shutil.copy(out, args.keep)
            print(f"  saved replay -> {args.keep}")

    if not against:
        print(f"  {len(new)} rows (no staging file to compare with)")
        return
    old = load(against)
    only_new = sorted(set(new) - set(old))
    only_old = sorted(set(old) - set(new))
    print(f"\nCompared with {against}")
    print(f"  replay rows: {len(new)}   original rows: {len(old)}")
    print(f"  only in replay ({len(only_new)}):")
    for k in only_new:
        print(f"    + {k} | {str(new[k].get('Title'))[:70]}")
    print(f"  only in original ({len(only_old)}):")
    for k in only_old:
        print(f"    - {k} | {str(old[k].get('Title'))[:70]}")
    changed = 0
    for k in sorted(set(new) & set(old)):
        for col in ("Deadline", "Title", "Type", "Abbreviation"):
            if str(new[k].get(col) or "") != str(old[k].get(col) or ""):
                if changed == 0:
                    print("  changed cells:")
                changed += 1
                print(f"    ~ {k} {col}: {old[k].get(col)!r} -> {new[k].get(col)!r}")
    if not (only_new or only_old or changed):
        print("  identical on Record ID, Title, Deadline, Type and Abbreviation")


if __name__ == "__main__":
    main()
