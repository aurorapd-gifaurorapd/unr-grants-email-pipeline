# Summary rewrite workflow for Claude Code

Use this workflow after the human has exported a validated workbook from the review HTML tool.

## Scope

Process only the rows where column Q, "Needs AI Summary Help", is "Y". Treat every other row as out of scope and leave it unchanged.

## Rules

1. Read the row's current Summary, the AI Rewrite Notes, and the source context from the workbook.
2. Remove any HTML tags or entities from the existing summary text and use it only as raw notes.
3. Research any missing details requested by the notes, using the row's primary URL first, then the secondary URL, then a broader web search if needed.
4. Draft one paragraph of full sentences, normally 4-5 sentences when the source supports it.
5. Include only facts that are actually present or verified.
6. Preserve specific topic names, funder priorities, methods/data types, and career-stage restrictions when they are present.
7. Do not use bullets, labels, or formulaic templates.
8. Eliminate any temporary unresolved markers before the final output.
9. Make each final summary distinct from the others in the same batch, with different openings and sentence structure.
10. Write the updated workbook to a new file named validated_<name>_final.xlsx rather than overwriting the human's export.

## Output

- Update only the Summary column for the in-scope rows.
- Leave every out-of-scope row's Summary unchanged.
- Preserve all other cells and workbook structure.
- Report how many rows were updated and note any unresolved details that still could not be verified.
