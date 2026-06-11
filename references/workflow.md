# DOCX Template Matching Workflow

Use this reference when the template has unusual structure or the default checker mapping does not fit.

## Task Intake

Collect these inputs before editing:

- Template path and target path.
- Submission rules that are not visible in the template, such as word count, reference count, naming rules, PDF export, or language requirements.
- Whether the target should preserve existing content exactly or allow content repair.
- Whether the deliverable needs a JSON report, a human-readable audit summary, or both.

If the user provides only a `.doc` template, convert it to `.docx` first and keep both files. Use the converted `.docx` only as the working baseline.

## Build Strategy

Prefer this order:

1. Open the template `.docx` as the starting document.
2. Clear body content while preserving section properties.
3. Insert target content into the template structure.
4. Apply existing template styles by name.
5. Copy direct paragraph/run properties only for roles that rely on direct formatting.

Avoid rebuilding a document from a blank default `Document()` unless no usable template document exists. Blank defaults often miss margins, section sizes, inherited styles, and theme settings.

## Extract Template Paragraph Indices

Run:

```powershell
python -c "from docx import Document; d=Document('template.docx'); [print(i, repr(p.text[:80]), p.style.name, p.alignment) for i,p in enumerate(d.paragraphs) if p.text.strip()]"
```

Identify the template paragraph index for each role:

- `cover_title`
- `topic`
- `student_id`
- `name`
- `major`
- `teacher`
- `college`
- `article_title`
- `author`
- `affiliation`
- `abstract`
- `keywords`
- `h1`
- `body_after_h1`
- `h2`
- `body_after_h2`
- `refs_heading`
- `ref_item`

Then identify the matching target paragraph indices after generation and pass a JSON mapping to `compare_docx_template.py`.

Use the checker to print the auto-detected mapping:

```powershell
python scripts\compare_docx_template.py --template template.docx --target target.docx --dump-map --summary
```

If the map is wrong, create a small JSON file with only the roles that need explicit indices.

## Inspect Full Properties

Use this snippet when a mismatch needs diagnosis:

```powershell
python -c "from docx import Document; d=Document('file.docx');
for i,p in enumerate(d.paragraphs):
    if p.text.strip():
        pf=p.paragraph_format
        print(i, repr(p.text[:60]), p.style.name, p.alignment, pf.left_indent, pf.first_line_indent, pf.line_spacing, pf.space_before, pf.space_after)"
```

## Diagnose Failures

Read failures in this order:

1. `section_failed`: page size, margins, header distance, and footer distance do not match.
2. `format_failed`: mapped paragraph styles, alignment, indentation, spacing, or line spacing differ.
3. `document_colors`: inspect explicit colors in `word/document.xml`.
4. `color_failed`: explicit non-black colors remain when `--require-black-only` is used.

Use `--roles` to isolate a problem:

```powershell
python scripts\compare_docx_template.py --template template.docx --target target.docx --roles h1,h2,body_after_h2 --include-run
```

Use `--report` for repeatable iteration:

```powershell
python scripts\compare_docx_template.py --template template.docx --target target.docx --check-colors --summary --report reports\format-report.json
```

## Important Details

- Use `Document(template_docx)` and clear the body when creating a new document.
- Preserve `w:sectPr` when clearing the body so page setup remains from the template.
- If Word opens headings in blue, inspect `word/document.xml` and `word/styles.xml`; remove theme/accent heading colors or explicitly set target runs to black.
- Compare paragraph properties first. Compare run properties only when template compliance requires exact font inheritance at run level.
- If the template uses custom styles, reuse them by name instead of recreating them.

## Final Delivery Pattern

Report the final document path, validation command, and pass/fail status. If any mismatch remains, list the exact role and property still differing instead of saying the document is “basically matched.”
