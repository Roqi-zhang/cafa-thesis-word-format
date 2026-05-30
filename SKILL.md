---
name: cafa-thesis-word-format
description: Use this skill when formatting Central Academy of Fine Arts undergraduate thesis Word documents, applying CAFA thesis formatting rules, converting green source annotations into Word native footnotes, enforcing circled-number superscript footnote markers, formatting image captions and references, and generating content-protection validation reports.
---

# CAFA Thesis Word Format

Use this skill for privacy-safe automation of CAFA undergraduate thesis `.docx` formatting. Treat `SKILL.md` as the required entry point. `agents/openai.yaml` is optional UI metadata only and must not be required for the skill to work.

## Required Collaboration Flow

Always follow this sequence:

1. Understand the task and inspect the named source document only as needed.
2. Produce a concrete plan that states source file, target file, allowed changes, forbidden changes, validation steps, and reports.
3. Wait for explicit human confirmation before modifying any Word document.
4. Execute only the confirmed scope, writing outputs to user-approved paths.
5. Provide verification results and manual checking standards.

Do not treat the initial request as execution approval when the task affects a thesis document. Ask for confirmation after planning.

## Allowed Changes

Allowed only after confirmation:

- Apply page, paragraph, font, spacing, and style adjustments required by the thesis format.
- Convert green source annotations already present in the body into Word native page footnotes.
- Remove only the green source annotation text that has been converted into footnotes.
- Preserve native `w:footnoteReference` and `word/footnotes.xml` structure.
- Use circled-number visible markers for body references and footnote-area markers.
- Format footnote text as small-five size, Songti for Chinese text, Times New Roman for Latin text and digits, black, regular weight, single line spacing.
- Normalize existing image caption formatting without inventing or rewriting caption content.
- Repair confirmed problem-fix passes such as fragmented native footnotes, missing page numbers, major chapter page breaks, and legacy floating image layouts when explicitly approved.
- Normalize existing reference-list formatting without adding, deleting, or completing bibliography entries.
- Generate a change report and a content-protection validation report.

## Forbidden Changes

Never do the following unless the user separately confirms a new plan:

- Overwrite the source document.
- Rewrite, polish, translate, summarize, or otherwise change ordinary body text.
- Invent authors, titles, publication data, URLs, page numbers, access dates, image sources, image authorship, or missing bibliography data.
- Put "missing information" placeholders into the Word body, footnotes, captions, or bibliography.
- Update table of contents, page numbers, fields, caption numbers, cross-references, or heading numbers.
- Delete bibliography entries or suspected template examples.
- Add bibliography entries from footnotes without a separately confirmed list.
- Move images, resize images, change wrapping, or renumber figures.
- Convert native page footnotes into endnotes or manual text lists.
- Upload or commit real thesis files, generated reports, private files, or personal identifiers.

## Reference Files

Load reference files only when needed:

- `references/cafa-thesis-format-standard.md`: detailed formatting rules.
- `references/execution-checklist.md`: planning, execution, and validation checklist.
- `references/privacy-safety-rules.md`: repository and GitHub privacy rules.

## Script Templates

The scripts in `scripts/` are parameterized templates. Before using one on a real document:

- Read the script and confirm the command-line arguments.
- Pass explicit source and target paths supplied or approved by the user.
- Do not hard-code private paths, thesis titles, author names, report content, or one-off note counts.
- Make a new output file rather than overwriting the source.
- Run validation after any document-modifying script.

Suggested use:

```bash
python scripts/process_docx_template.py --source-docx SOURCE.docx --target-docx TARGET.docx --report-md REPORT.md
python scripts/fix_footnote_style_template.py --source-docx TARGET.docx --target-docx TARGET_FOOTNOTES.docx --report-md FOOTNOTE_REPORT.md
python scripts/validate_docx_template.py --source-docx SOURCE.docx --target-docx TARGET_FOOTNOTES.docx --report-md VALIDATION_REPORT.md
python scripts/repair_layout_issues_template.py --source-docx SOURCE.docx --target-docx TARGET_REPAIRED.docx --report-md REPAIR_REPORT.md --validation-report-md VALIDATION_REPORT.md
```

Use `repair_layout_issues_template.py` only for a confirmed repair pass, not for the initial planning inspection. It is designed to copy the source to a target, safely merge clearly fragmented footnotes, add centered PAGE fields, set major headings to start on new pages, normalize figure captions, and prevent legacy floating images from overlapping text. Missing or uncertain source data must remain in reports only.

## Completion Report

Report these items after execution:

- Source file remained untouched.
- Output file path and report paths.
- Count of native footnotes and body footnote references.
- Count of remaining green text runs.
- Whether circled-number superscript markers are present.
- Whether footnote text formatting matches requirements.
- Whether ordinary body text protection passed.
- Image caption and reference-list changes made, if any.
- Items that require manual visual checking.
