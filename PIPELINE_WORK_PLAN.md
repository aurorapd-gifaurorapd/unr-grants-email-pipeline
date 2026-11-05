# Grants Email Pipeline: Work Plan

Written November 5, 2026, after the 11/05 run. Builds on the backlog in
`OLD Versions/weekly notes for improvements/PIPELINE_IMPROVEMENTS.md`.

## The problem this plan solves

The 11/05 run went well in Claude Code, but only because Claude did a lot of
work by hand that the scripts don't do. Someone running the documented
PowerShell steps would have:

- **Sent an email with no SAM.gov items.** The SAM key had expired, every call
  returned `API_KEY_INVALID`, and the script reported "0 hits", which looks normal.
- **Missed two weeks of NIH notices.** The NIH feed only reaches back about three
  days, so a 14-day gap loses webinars and policy notices.
- **Lost 18 real opportunities**, including both Great Basin agreements. The
  duplicate filter matched on the first four title words and silently dropped
  them, and the federal seen-list meant they would never come back.
- **Used a near-empty Pivot export** (5 rows instead of the usual ~180) without
  any warning.
- **Lost edits made to staging after the review export.** The SCALED concept
  paper deadline and the AFOSR white paper date had to be re-applied by hand.
- **Shipped raw SAM dates and placeholders** (`2026-12-04T14:00:00`, `n/a`, a
  blank deadline) until the problem check caught them.
- **Copied and pasted 10 additions and 24 summary rewrites by hand.**

**The goal: anyone can run the full week from PowerShell with a few commands, get
the same result Claude Code got on 11/05, and get a clear message whenever
something needs a human.** Claude Code should make the week faster, not be the
only way it works.

## Principles

- **Fail loudly.** A failed search must stop with a message, never report 0.
- **Every command ends with "Next step:"** and the exact command to run.
- **No hand-typed variables.** The week folder is found automatically, not set with `$w`.
- **Scripts back up before they overwrite**, and never silently append to old output.
- **Anything Claude did by hand twice becomes a script.**
- **Same rules everywhere.** The deadline rule, the duplicate rule and the
  allowed types live in `grants_common.py`, and the skills and docs point to them.

## Where the old backlog stands

| Item | Status |
|---|---|
| P0-0 Set up git | **Open** (the folder is not a git repo) |
| P0-1 `validate_staging.py` | Done |
| P0-2 `ENCODING_MAP` (U+202F) | **Check**: no U+202F entry found in the code |
| P0-3 Agency routing guard | Done (merge stops on an unknown section) |
| P1-1 One Type vocabulary | Done (`TYPE_INFO` in `grants_common.py`) |
| P1-2 Public SAM URLs | Done |
| P1-3 Newlines in summaries | Done (validator strips them) |
| P1-4 Monday header date | Done (`current_week_monday`) |
| P1-5 SAM PSC gate | Done; 74 of 88 dropped on 11/05, worth a spot-check |
| P1-6 SAM in the merge command | Done |
| P2-1 NASA ROSES collector | **Open**: Phase 3 below |
| P2-2 Federal Register / policy watch | **Open**: Phase 6 below |
| P2-3 Drop the NSF RSS collector | Open |
| P3 Strip `utm_*` tracking tags | **Open** (needed by hand on 11/05) |

Note: the backlog file lives in `OLD Versions`. Moving it to the pipeline folder,
next to this plan, would make it easier to find.

---

## Phase 0: Safety net (one session)

1. **Git.** Run `git init` with the `.gitignore` from backlog P0-0, then commit. Box
   sometimes struggles to sync `.git` folders, so also push to a private GitHub
   repo, which gives you a backup that doesn't depend on Box.
2. **Root `CLAUDE.md`.** `Claude Instructions/README.md` points to `../CLAUDE.md`,
   which doesn't exist. Add one at the pipeline root that covers the weekly steps,
   the rules below, and "never run collectors twice in a day".
3. **Replay harness.** Add `replay_week.py <week folder>`, which re-runs the merge
   on a past week's saved collector outputs and compares the result with what was
   sent. Every later change gets tested against the 10/22 and 11/05 data before
   it's trusted.

**Done when** git has a baseline commit and `replay_week.py Grant_11_05_2026` reproduces the 63-row staging file.

## Phase 1: One command for the week (the biggest PowerShell win)

Create `run_week.py` with these subcommands:

| Command | What it does |
|---|---|
| `py run_week.py start` | Creates `Grant_MM_DD_YYYY`, finds the last sent email, and works out the lookback (days since then, plus a margin) |
| `py run_week.py check` | Preflight: tests the SAM key with one call, checks the Pivot export (header, row count compared with recent weeks, date range), and confirms dependencies |
| `py run_week.py collect` | Runs every collector with the computed lookback, saves the seen-lists into the week folder first, and refuses to re-run a collector that already ran today |
| `py run_week.py merge` | Backs up and replaces staging (never appends), then builds and opens the review page |
| `py run_week.py finalize` | Finds the newest export in Downloads, moves it to the week folder, and makes the `_final` copy |
| `py run_week.py add <file>` | Appends rows from a non-grant additions file to `_final` and checks them |
| `py run_week.py check-final` | Runs `validate_staging.py` plus the build's duplicate and deadline checks |
| `py run_week.py build` | Builds the email HTML and opens it |
| `py run_week.py status` | Shows which steps are done and what to do next |

Progress is saved in `week.json` in the week folder, so you can close
PowerShell and pick up where you left off.

**Done when** Carrie can run start → build from PowerShell with no Claude and no `$w`.

## Phase 2: Collector fixes from 11/05

1. **SAM.gov (done November 5, 2026):**
   - A rejected key (401/403) or a used-up daily limit (429) now stops on the
     first call, with plain steps to fix it, and writes no output file.
   - SAM timestamps like `2026-12-04T14:00:00` now come through as `December 04, 2026`.
     The fix is in the shared date reader in `grants_common.py`, so every collector benefits.
2. **NIH (done November 5, 2026):** `nih_rss_collect.py` now uses the NIH Guide
   search over a 30-day window, with a new seen-list (`nih_seen.json`, seeded on
   11/05, so don't delete it). Details:
   - It reads each notice's page for its response date, and the event date for webinars.
   - It skips notices whose date has passed, and notices that only announce which
     NIH institutes take part in a program. Both are logged to `nih_skipped.json`.
   - It stops with a message if NIH's search doesn't answer.

   The weekly command is unchanged. Originals are in `OLD Versions/scripts replaced 2026-11-05`.
3. **Federal (Grants.gov):** find out why the Forest Service opportunity
   USDA-FS-2026-LSR (posted 10/28) was missed. Likely causes are the posted date on
   a version-20 listing, or an agency filter.
4. **Links:** strip `utm_*` tracking tags from every URL during collection (backlog P3).
5. **Encoding:** confirm U+202F and the other invisible spaces are handled (backlog P0-2).

**Done when** a run with a bad SAM key stops, and replaying 11/05 finds the NIH gap notices with no manual step.

## Phase 3: NASA ROSES collector (new priority)

**Why it works now:** the 08-13 backlog said NSPIRES blocks crawling. On 11/05
everything NASA-side read cleanly:

| Source | What it gives | 11/05 result |
|---|---|---|
| ROSES blog index and posts (`science.nasa.gov/researchers/solicitations/roses-2025/`) | What changed this week: new elements, amendments, announcements | Worked (Amendment 81, the ROSES-27 announcement) |
| NSPIRES summary page (`summary.do?solNum=…`) | Title, status, NOI / Step-1 / Step-2 dates | Worked, public, no login |
| Program element PDF (NSPIRES document link) | "Summary of Key Information": budget, number of awards, duration | Worked by direct download |

**Design for `nasa_roses_collect.py`:**
1. Read the ROSES blog index and keep posts newer than the ones in `nasa_seen.json`.
   Use the same seen-list pattern as the federal and NSF collectors, so a wider
   window doesn't bring old posts back.
2. Sort each post into one of these types:
   - **New opportunity** (a new program element, Grant)
   - **Due date change**
   - **Clarification or other amendment** (Notice)
   - **Program-wide announcement** (Policy, like the ROSES-27 announcement)
3. For new opportunities and date changes, read the NSPIRES page for the element's
   dates and status. Optionally read the element PDF for award size and duration.
4. Treat a required NOI or Step-1 as the earliest required step. It goes in the
   Deadline cell, with the full date sequence in the summary, the same rule used everywhere.
5. Find the ROSES year from the index, so the switch to ROSES-2027 (expected by
   January 2027) needs no code change.

**Done when** replaying 11/05 produces the Astrophysics Theory Program (Amendment 81) and the ROSES-27 notice, and replaying 10/22 produces no false new items.

## Phase 4: Duplicates and seen-lists

Step 1 is already done: the duplicate filter no longer drops records with different numbers.

1. **Flags instead of drops.** Records that might be duplicates are flagged in
   Eligibility Notes ("possible duplicate of …") rather than removed. Every drop
   is written to `dedup_log.csv` in the week folder with its reason.
2. **Re-merging into an existing staging file:** apply the same number-aware rule there.
3. **Seen-lists:** look at marking federal records "seen" only once they're
   actually sent, so a wrong drop can be recovered. Do this only if wrong drops
   still happen after items 1 and 2.
4. **Replay check:** run every change against the past week folders. The 10/22
   replay found 5 records lost under the old rule.

## Phase 5: Review to final file (where most of the hand work happened)

1. **One rule:** once the review page is exported, every edit goes into `_final`.
   `run_week.py finalize` makes the copy, and any later command that changes
   staging warns that the edit won't reach the email.
2. **Additions:** `run_week.py add` appends the non-grant additions file to `_final`,
   sets Email? to Y, and runs the checks. No more pasting rows.
3. **New problem checks**, each based on something that happened on 11/05:
   - Identical titles with different numbers: suggest adding the number (the Great Basin pair).
   - Summary text pasted twice (the NSF Super Intelligence Dear Colleague Letter).
   - A required LOI, white paper, pre-proposal or concept paper dated before the
     Deadline cell: warn (SCALED, the AFOSR center, the AFOSR-Taiwan program).
   - An Event with no Pacific time in its Deadline cell (the NIH R25 webinar).
   - `http://` links and ISO timestamps: offer to fix them automatically, and
     report what changed.
4. **Write the deadline rule down:** the Deadline cell holds the earliest
   *required* date, and the summary gives the full sequence in Pacific and Eastern
   time. Put this in `CLAUDE.md`, the how-to doc and both skills.

## Phase 6: Skills and finding additions

1. **Summary writer:** approve, then write (already in the backlog). The person
   approves rewrites, the skill writes only those cells, then checks its work.
2. **`fetch_sources.py <final file>`:** for every row needing a rewrite, save
   source text to `sources/<Record ID>.txt`. Sources are the Grants.gov API
   (description, award ceiling and floor, cost share), NIH notice pages, darpa.mil
   pages, and any PDFs in the week folder. This works the same from PowerShell
   (then upload to claude.ai) or Claude Code. Note that new NIH NOFO pages under
   `rfa-files/` returned 404 on 11/05, so for those use the Grants.gov API.
3. **`scout.py`** (a candidate list for non-grant additions, built from official sources):
   - **NIH Guide API:** notices in the date range.
   - **Federal Register API:** RFIs and notices from the funding agencies. This
     worked cleanly for the USGS notice, and covers backlog P2-2.
   - **NSF events pages and DARPA events pages.**
   - **The NASA collector** (Phase 3).

   It writes a short candidate table, and a person picks which to keep. Web search
   becomes a backup: on 11/05 it reported an NSF office-hours session that didn't exist.
4. **Keep each skill in two places, in sync:** the installed skill and the copy in
   `Claude Skills/` (and its `.zip` for claude.ai).

## Phase 7: Documentation and handoff

1. **Rewrite `HOW TO RUN`** around `run_week.py`. Add troubleshooting rows for: SAM
   key rejected, Pivot export too small, collectors already run today, and edits
   made after export.
2. **Reliability table** in `CLAUDE.md`: which sources can be read and how
   (Grants.gov API yes, Grants.gov pages no, SAM.gov pages no, NSPIRES yes, NIH
   Guide API yes, darpa.mil yes, Federal Register API yes).
3. **Weekly human checklist,** kept from the backlog: stale funder facts, buried
   eligibility rules, multi-stage deadlines, and cost-policy problems.

---

## Order and size

| # | Phase | Size | Why this order |
|---|---|---|---|
| 1 | Phase 0: safety net | 1 session | Everything after this needs rollback and replay |
| 2 | Phase 2: items 1–2 (SAM fails loudly, NIH API) | 1 session | Cheap, and stops silent data loss next week |
| 3 | Phase 1: `run_week.py` | 2–3 sessions | Biggest win for running from PowerShell |
| 4 | Phase 3: NASA ROSES collector | 2 sessions | Biggest coverage gap; ROSES-27 arrives by January |
| 5 | Phase 5: final-file flow and new checks | 1–2 sessions | Removes most copy and paste |
| 6 | Phase 4: duplicate flags, log, seen-list | 1–2 sessions | Needs the replay harness |
| 7 | Phase 6: skills, `fetch_sources.py`, `scout.py` | 2 sessions | Builds on the stable final-file flow |
| 8 | Phase 7: docs and handoff | 1 session | Last, so the docs describe the finished tools |
| – | Phase 2: items 3–5 | Fold into any session | Small, independent fixes |

**How to work it:** one item per Claude Code session, and look at the changes
before they're saved. Then replay the 10/22 and 11/05 weeks and confirm the
record and section counts match, or that every difference is explained.
