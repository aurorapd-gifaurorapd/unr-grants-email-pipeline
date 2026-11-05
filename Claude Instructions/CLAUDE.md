# CLAUDE.md

This repository runs the UNR weekly grants email pipeline. Follow the detailed operating instructions in [Grants Email Pipeline.md](Grants%20Email%20Pipeline.md) for the full collection, review, rewrite, and build workflow.

## Working rules

- Treat the review export as the human gate. Do not decide which rows belong in the email yourself.
- Do not build final HTML from a workbook that did not come out of the review tool export.
- When the human has exported a validated workbook, evaluate every row's summary against the quality bar and rewrite whichever ones fall short — the "Needs AI Summary Help" column is a priority signal, not the scope boundary, since some rows (e.g. Scout-agent-sourced records) reach this stage without ever being flagged.
- Preserve every row that already meets the quality bar exactly as-is; write a new workbook with only the rewritten summaries changed.
- Never invent funding amounts, deadlines, eligibility facts, times, or other program details. If something is missing, omit it or mark it as unresolved during drafting, then resolve it through source research before finalizing.
- Convert any clock time in a rewritten summary (deadline cutoffs, event/webinar times) to Pacific Time.
- Use the skill named grant-email-summary-rewrite for the summary-writing pass.

## Typical flow

1. Run the collectors and create the staging workbook.
2. Generate the review HTML and wait for the human to export a validated workbook.
3. Use the grant-email-summary-rewrite skill to review every row and rewrite whichever summaries fall short of the quality bar.
4. Save the rewritten workbook as validated_<name>_final.xlsx.
5. Build the HTML email from that final workbook.

## Important constraints

- Do not overwrite the human's validated export.
- Do not change out-of-scope rows or other cells beyond the Summary column for the flagged rows.
- Keep the output prose specific, factual, and distinct from other rows in the same batch.
