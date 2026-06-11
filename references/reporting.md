# Compliance Reporting

Use this reference when the user asks for an audit, acceptance criteria, or a delivery note for a template-matched paper or report.

## Report Contents

Keep reports short and evidence-based:

- Source template path.
- Target artifact path.
- Submission rules checked outside the template.
- Validation command.
- Machine result path, if generated.
- Pass/fail summary for sections, mapped paragraph formatting, colors, and content requirements.
- Remaining mismatches with role names and properties.

Do not include broad claims such as “format is consistent” unless the machine check passed for the relevant roles.

## Suggested Markdown Shape

```markdown
# Template Compliance Report

## Inputs

- Template: `path/to/template.docx`
- Target: `path/to/target.docx`
- Rules: word count, reference count, export format, naming rules

## Checks

- Format check: `python scripts/compare_docx_template.py ...`
- Section settings: pass/fail
- Paragraph roles: pass/fail
- Explicit colors: pass/fail/not checked
- Content rules: pass/fail/not checked

## Result

Pass or fail in one sentence.

## Remaining Work

List exact mismatches, or write `None`.
```

## Acceptance Criteria

Use these criteria before final delivery:

- `section_failed` is empty, unless the user approved a difference.
- `format_failed` is empty for all required roles.
- `color_failed` is empty when color restrictions apply.
- Content checks are listed separately from format checks.
- The final response includes the document path and the validation evidence.
