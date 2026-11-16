# Grants email pipeline: to-do list

Started November 16, 2026. The detailed designs are in
[PIPELINE_WORK_PLAN.md](PIPELINE_WORK_PLAN.md). This file is the short working
list, in the order to do things. Check items off as they're finished.

## A. Before the 11/19 email

- [x] **A1. Carry-forward items (plan 8b).** *Done 11/16.* `run_week.py repeat`
      saves a row to `carry_forward.xlsx` with an end date, and `run_week.py
      carry` adds the saved rows to the new week's `_final`. The NSF MRI
      (27-500) is saved through January 04, 2027, with its note in bold.
- [ ] **A4. On 11/19, run `py run_week.py carry` after `finalize`,** and confirm
      the MRI shows in the built email with the bold note.
- [ ] **A5. Add the `carry` step to the HOW TO RUN doc.**
- [ ] **A6. Look at one replay difference.** `replay_week.py Grant_11_12_2026`
      gives 37 rows where the original had 38. The missing one is the DARPA
      RFI, DARPA-SN-27-13. `run_week.py` doesn't touch the merge, so this
      isn't from the carry-forward change, but it should be explained.
- [x] **A2. Write down the 11/12 corrections** in `CLAUDE.md`: no clock times in
      the Deadline cell, numbers as link text, HRSA under HHS, and the shorter
      CESU summaries. *Done 11/16.*
- [ ] **A3. Copy the same corrections into the two skills** (additions and
      summary writer), and re-upload them to claude.ai. *Waiting on: question 1.*

## B. Before the first December email (12/03)

- [ ] **B1. Finish the standing list.** Needs answers to questions 6 to 9. Then
      remove the temporary Link Check column.
- [ ] **B2. Build the monthly Standing Funding Opportunities block (plan 8a):**
      `run_week.py standing`, the first-of-month check in `start`, and the
      grouped block at the top of `build_html.py`.

## C. One home for the rules (suggestion 1)

- [ ] **C1. One rules file.** Make a single short rules file the authority. The
      manual, the how-to doc, both skills and the Copilot text point to it
      instead of restating the rules.
- [ ] **C2. `run_week.py wrap-up`.** A last step each week that asks "what did
      you correct this week?" and records the answers in the rules file. It
      also commits the seen-lists, so the GitHub backup stays current.
- [ ] **C3. Tidy loose ends.**
      - The `Grant_10_22_26` folder. Needs question 4.
      - The nested skill folder and the deleted zips. Needs question 5.
      - Move the old backlog out of `OLD Versions` or retire it.

## D. Close the hand-work gap (suggestion 2)

- [ ] **D1. Summaries command.** Gather the source text for rows that need a
      rewrite, and apply approved rewrites to `_final` with a backup and a
      check (plan Phase 6, items 1 and 2).
- [ ] **D2. Candidate list for additions** from official feeds: the NIH Guide,
      the Federal Register, and the NSF and DARPA event pages (plan Phase 6,
      item 3, `scout.py`).
- [ ] **D3. The handoff test.** Someone other than Aurora runs start to build
      from PowerShell with no Claude. Needs question 2.
- [ ] **D4. Remaining checks from plan Phase 5:** a required early step dated
      before the Deadline cell, summary text pasted twice, `utm_` tags
      stripped at collection.

## E. Measure it (suggestion 3)

- [ ] **E1. Reject reasons.** The review page keeps every row, with a one-click
      reason for the ones left out (not eligible, off-topic, duplicate,
      expired).
- [ ] **E2. Spot-check the SAM filter.** It dropped 74 of 88 on 11/05.
- [ ] **E3. Monthly look at Emma's click report.** Needs question 3.

## Questions for Aurora

1. Is "no clock times in the Deadline cell" permanent, including for events?
   (The 11/12 event went out with the date only.)
2. Who else runs this? Is Carrie a real backup who needs the PowerShell route?
3. Do you have Emma's click reports, and has anyone looked at them?
4. Is `Grant_10_22_26` a leftover that can move to `OLD Versions`?
5. Was the nested `non-grant-email-additions/non-grant-email-additions/` skill
   folder (with the zips deleted) intentional?
6. Should NOAA rows in the standing list use NOAA or DOC?
7. Should CISE Future CoRe (25-543, deadline 02/04/2027) stay out of the
   standing list?
8. Which Candidates should be promoted to the Standing sheet?
9. Can you check the ROSES-2025 NSPIRES link and its expiration by hand?
10. Is January 04, 2027 the right end date for the MRI carry-forward?
11. Are the Copilot agents still in use, or can their instructions be retired?
