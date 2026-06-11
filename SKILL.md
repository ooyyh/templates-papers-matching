---
name: templates-papers-matching
description: Create, repair, or audit template-based academic artifacts so the output matches the required template and submission rules. Use when a user provides or references templates, papers, reports, theses, coursework, submission formats, formatting requirements, or asks for matching/compliance checks across documents. Supports DOCX template matching with automated formatting comparison, and should be extended similarly for other template/paper formats when needed.
---

# Templates, Papers, Matching

## Non-Negotiable Rule

Treat the template and submission requirements as the source of truth. Do not rely on defaults, theme styles, or visual guesswork. Generate or edit the target artifact so corresponding elements inherit or explicitly match the template, then run a machine comparison or another concrete compliance check. Do not report completion until the check passes or you clearly state the remaining mismatch.

## Reference Map

- Read `references/workflow.md` when creating, repairing, or auditing a real submission.
- Read `references/reporting.md` when the user asks for a compliance report, audit memo, acceptance criteria, or a concise delivery summary.
- Use `scripts/compare_docx_template.py` for DOCX layout checks; patch or extend it when a new repeatable check is needed.

## General Workflow

1. Identify the template, target artifact, and submission rules.
   - If the template is `.doc`, convert it to `.docx` first with Word COM on Windows when available.
   - Keep the converted template as the comparison baseline.
   - Never overwrite the original template.

2. Inspect the template before editing.
   - Extract representative paragraph properties for: cover title, title/topic line, student/name fields, article title, author, affiliation, abstract, keywords, heading 1, heading 2, body after heading 1, body after heading 2, references heading, and reference entries.
   - Inspect section properties: page size, margins, header distance, footer distance.
   - Inspect actual run colors in `word/document.xml`; theme colors can reopen as blue.

3. Build the target from the template when possible.
   - Prefer opening the template `.docx`, clearing the body, and inserting content into that document.
   - Reuse template styles such as `Normal`, `Heading 1`, `Heading 2`, and any custom styles (`CM19`, `CM20`, etc.) instead of creating new styles.
   - If a paragraph type has direct formatting in the template, copy the actual paragraph properties and run properties.

4. Format by element role, not by guess.
   - Cover fields must use the same paragraph count/order, indentation, tabs, line spacing, and font treatment as the template.
   - Abstract and keywords must match label font, body font, hanging indent, alignment, and line spacing.
   - Heading levels must match style, spacing, indentation, line spacing, font, and color.
   - Body paragraphs after different heading levels may use different template styles; preserve those differences.
   - References must match heading spacing and entry hanging indent.

5. Validate.
   - Run `scripts/compare_docx_template.py` against the template and target.
   - Check the report for `format_failed []` or no mismatches.
   - Check document XML for non-black colors if the template requires black headings/body text.
   - Check content requirements separately, such as word count and reference count.
   - Save a JSON report when the result will be reused, attached, or compared across iterations.

6. Iterate until clean.
   - If validation shows mismatches, patch the generator or document and rerun validation.
   - Do not say the format is fixed based only on opening the file or eyeballing it.

## DOCX Script

Use the bundled checker:

```powershell
python C:\Users\OYeah\.codex\skills\templates-papers-matching\scripts\compare_docx_template.py `
  --template path\to\template.docx `
  --target path\to\target.docx `
  --check-colors `
  --summary `
  --report path\to\format-report.json
```

For templates whose matching positions are unusual, pass a JSON role mapping:

```json
{
  "cover_title": [8, 8],
  "topic": [10, 10],
  "article_title": [24, 23],
  "abstract": [28, 27],
  "h1": [31, 30],
  "h2": [35, 36],
  "ref_item": [53, 86]
}
```

Then run:

```powershell
python ...\compare_docx_template.py --template template.docx --target target.docx --mapping mapping.json --check-colors
```

The mapping values are `[template_paragraph_index, target_paragraph_index]`.

Useful options:

- `--dump-map`: print the auto-detected template/target paragraph map.
- `--roles h1,h2,ref_item`: compare only selected roles while diagnosing.
- `--include-run`: include first text run font checks.
- `--require-black-only`: fail if explicit non-black document colors remain.
- `--report report.json`: write the full machine-readable result.

## Completion Checklist

Before final response, confirm:

- The target was generated from, or explicitly matched to, the template.
- Section margins/page settings match.
- All mapped paragraph properties match: style, alignment, indentation, line spacing, spacing before/after.
- Font role properties match or intentionally inherit from the template.
- No unwanted blue/theme color remains in target document content.
- User-facing content requirements are satisfied.
- The final `.docx` path is provided.
