# Claude Instructions

This folder contains Claude-ready workflow guidance for the weekly grants email pipeline.

## Files

- [CLAUDE.md](../CLAUDE.md) — repository-level Claude entrypoint for the weekly workflow.
- [summary-rewrite-workflow.md](summary-rewrite-workflow.md) — the step-by-step process for rewriting flagged summaries after review.

## Use

Use these files when working in Claude Code on this repository so the summary rewrite pass follows the same rules as the original Copilot workflow.

## Weekly rewrite add-ons

- Independently verify apparent duplicates in broad NSF or directorate-wide records against the official NSF pages. Treat related bundled programs as distinct when their titles, opportunity numbers, program pages, or scopes differ; flag only evidence of an actual duplicate.
- Check the official source during the rewrite pass for active versus archived status. An archived opportunity must not be described as an open call; correct its deadline/status wording when the validated workbook is misleading.
- Verify recurring deadlines and event dates/times against the source. Correct the Deadline cell when it is truncated, ambiguous, or missing the source time zone; convert event times to Pacific Time and retain the source zone in parentheses.
- Normalize SAM.gov opportunity links to public `https://sam.gov/opp/<notice-id>/view` URLs before final validation. Do not leave authenticated workspace URLs in the final workbook.
- Keep a short record of these non-routine checks and corrections so the workflow and skills can be updated when a new failure mode appears.
