# cafa-thesis-word-format

Reusable Codex Skill for privacy-safe CAFA undergraduate thesis Word formatting.

The required skill entry point is `SKILL.md`. The optional `agents/openai.yaml` file only improves UI presentation and is not required for the skill to trigger or run.

## What It Covers

- Basic Word formatting workflow.
- Green source annotation conversion into native Word page footnotes.
- Circled-number footnote markers.
- Small-five footnote text formatting.
- Existing image caption format checks.
- Confirmed repair passes for fragmented footnotes, missing page numbers, major chapter page breaks, and legacy floating-image layout issues.
- Existing reference-list format checks.
- Content-protection validation.
- Change and validation reports.

## Privacy Rule

Do not commit, push, or upload real thesis documents, generated reports, private files, personal identifiers, or school source-document copies. This repository is intended to contain only reusable skill instructions, sanitized references, and parameterized script templates.

## Usage

Install or copy this folder into a Codex skills directory, then ask Codex to use `cafa-thesis-word-format` for a CAFA undergraduate thesis Word formatting task.

The agent must plan first, wait for human confirmation, execute only the confirmed scope, and then report validation results.

## Repair Template

For follow-up repair passes, use the sanitized template:

```bash
python scripts/repair_layout_issues_template.py \
  --source-docx SOURCE.docx \
  --target-docx TARGET_REPAIRED.docx \
  --report-md REPAIR_REPORT.md \
  --validation-report-md VALIDATION_REPORT.md
```

The template is parameterized and should be reviewed before use on a real thesis. It must never be committed with real thesis documents, generated outputs, or private report excerpts.
