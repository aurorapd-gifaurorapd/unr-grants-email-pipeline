# CLAUDE.md: UNR weekly grants email pipeline

The full operating manual is [Claude Instructions/Grants Email Pipeline.md](Claude%20Instructions/Grants%20Email%20Pipeline.md).
The step-by-step human guide is `HOW TO RUN - Weekly Grants Email.docx`. The
improvement plan is [PIPELINE_WORK_PLAN.md](PIPELINE_WORK_PLAN.md). The rules below
are newer than the manual; where they disagree, these win.

## Weekly run, in order

1. The week folder is `Grant_MM_DD_YYYY` (the run date), containing the Pivot export `opps_export.csv`.
2. Run the collectors once each:
   - `grants_collect.py`
   - `nih_rss_collect.py`
   - `nsf_rss_collect.py`
   - `pivot_collect.py`
   - `sam_collect.py`
3. Merge **all** `*_opportunities.json` files in the week folder into
   `Grant_Opportunities_Staging.xlsx`. Delete the staging file first if you're
   re-merging, because the merge appends to an existing file.
4. Run `build_review_html.py`. **WAIT** while the human reviews and exports `validated_*.xlsx`.
5. Copy the export to `validated_…_final.xlsx`. Every edit from here on goes into
   `_final`, never into staging, because staging edits don't reach the email.
6. Non-grant additions (skill `non-grant-email-additions`) and summary rewrites
   (skill `grant-email-summary-writer`) go into `_final`.
7. Run `validate_staging.py <_final>` until it shows 0 errors, then `build_html.py <week folder>`.

## Hard rules learned on 11/05

- **Run collectors once a day.** `grants_collect.py`, `nsf_rss_collect.py` and
  `nih_rss_collect.py` update `federal_seen.json`, `nsf_seen.json` and
  `nih_seen.json`. A second run returns 0. Never delete or hand-edit these files.
- **Check that the lookback covers the gap since the last email.** The default
  windows are federal 30 days, NSF 30, NIH 30 and SAM 14. If more than 14 days
  have passed since the last email, say so before running `sam_collect.py`.
- **SAM.gov:** about 10 calls a day, and a full run uses up to 9. If the key is
  rejected or the limit is reached, the collector stops and writes nothing. Don't
  retry the same day. To check the key, make one search call with `limit=1`.
- **Pivot:** if `opps_export.csv` has far fewer rows than recent weeks (usually
  about 150–180), tell the human before going on. It may be a partial export.
- **Deadline rule:** the Deadline cell holds the earliest *required* date (a
  required LOI, white paper, pre-proposal, concept paper or State Forester
  submission). The summary ends with a sentence giving the full sequence, with
  times as `2:00 PM PT (5:00 PM ET)`. Events give the date plus Pacific time.
- **Allowed Types:** `Grant`, `DCL`, `RFI`, `Event` and `Policy`. Never use
  `Limited Submission` as a Type; mark that in the summary and Recommendation instead.
- **Two opportunities with identical titles but different numbers** (for example
  CESU agreements) are separate opportunities. Add the number to the title so the
  email shows them as distinct.
- **Strip `utm_*` tracking tags** from links before they go in the sheet.
- **Confirm before editing** a pipeline script. Save the original first in
  `OLD Versions/scripts replaced YYYY-MM-DD/`, then test with
  `python replay_week.py <week folder>`, which writes nothing to the week folder.

## Which sources can be read (as of 11/05)

| Source | How to read it |
|---|---|
| Grants.gov listing pages | Can't be read. Use the API: `POST https://api.grants.gov/v1/api/fetchOpportunity` with `{"opportunityId": N}` |
| SAM.gov pages | Can't be read. Use the PDF the human attaches, the collector's data, or the agency's page (for example darpa.mil) |
| NIH Guide notices | `grants.nih.gov/grants/guide/notice-files/<NOT-…>.html` works. New NOFO `rfa-files/` pages returned 404, so use the Grants.gov API for those |
| NIH Guide search | `search.grants.nih.gov/guide/api/data`, with `type=notices` and `daterange=MMDDYYYY-MMDDYYYY` |
| NSPIRES (NASA) | Summary pages (`summary.do?solNum=…`) are public. Program element PDFs download directly |
| NASA ROSES blog, nsf.gov, neh.gov, darpa.mil, science.osti.gov | Readable |
| Federal Register | `https://www.federalregister.gov/api/v1/documents/<doc number>.json` |
| simpler.grants.gov | Opportunity pages are readable; attachments come from `files.simpler.grants.gov` |
