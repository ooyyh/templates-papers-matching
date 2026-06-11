# DOCX Template Matching Workflow

Use this reference when the template has unusual structure or the default checker mapping does not fit.

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

## Inspect Full Properties

Use this snippet when a mismatch needs diagnosis:

```powershell
python -c "from docx import Document; d=Document('file.docx');
for i,p in enumerate(d.paragraphs):
    if p.text.strip():
        pf=p.paragraph_format
        print(i, repr(p.text[:60]), p.style.name, p.alignment, pf.left_indent, pf.first_line_indent, pf.line_spacing, pf.space_before, pf.space_after)"
```

## Important Details

- Use `Document(template_docx)` and clear the body when creating a new document.
- Preserve `w:sectPr` when clearing the body so page setup remains from the template.
- If Word opens headings in blue, inspect `word/document.xml` and `word/styles.xml`; remove theme/accent heading colors or explicitly set target runs to black.
- Compare paragraph properties first. Compare run properties only when template compliance requires exact font inheritance at run level.
- If the template uses custom styles, reuse them by name instead of recreating them.
