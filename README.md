# Templates Papers Matching

A Codex skill for creating, repairing, and auditing academic papers or report documents against a required template.

![Workflow preview](assets/workflow-preview.png)

## What It Does

- Treats the provided template and submission rules as the source of truth.
- Helps generate or repair `.docx` artifacts so roles such as cover fields, abstracts, headings, body text, and references match the template.
- Provides a DOCX checker that compares section settings, paragraph formatting, optional run formatting, and explicit document colors.
- Produces repeatable JSON validation evidence for delivery or review.

## Repository Layout

- `SKILL.md`: Codex skill instructions and trigger metadata.
- `scripts/compare_docx_template.py`: DOCX template compliance checker.
- `references/workflow.md`: practical workflow for real template-matching tasks.
- `references/reporting.md`: concise compliance report guidance.
- `assets/workflow-preview.svg`: editable workflow preview image.
- `assets/workflow-preview.png`: rendered preview image.
- `agents/openai.yaml`: UI metadata for the skill.

## Quick Usage

Run a DOCX format check:

```powershell
python scripts\compare_docx_template.py `
  --template path\to\template.docx `
  --target path\to\target.docx `
  --check-colors `
  --summary `
  --report path\to\format-report.json
```

Print the auto-detected paragraph role mapping:

```powershell
python scripts\compare_docx_template.py `
  --template path\to\template.docx `
  --target path\to\target.docx `
  --dump-map `
  --summary
```

Compare only selected roles while diagnosing:

```powershell
python scripts\compare_docx_template.py `
  --template path\to\template.docx `
  --target path\to\target.docx `
  --roles h1,h2,ref_item `
  --include-run
```

## Custom Role Mapping

When automatic paragraph detection does not match the template structure, provide a JSON mapping:

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
python scripts\compare_docx_template.py `
  --template path\to\template.docx `
  --target path\to\target.docx `
  --mapping path\to\mapping.json `
  --check-colors `
  --summary
```

Mapping values are `[template_paragraph_index, target_paragraph_index]`.

## Validation

The skill itself can be validated with:

```powershell
python C:\Users\OYeah\.codex\skills\.system\skill-creator\scripts\quick_validate.py .
```

The checker requires Python and `python-docx`.
