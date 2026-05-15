# Execution Checklist

Use this checklist for every thesis-formatting task.

## Before Planning

- Identify the exact source document named by the user.
- Inspect only metadata and structure needed for planning.
- Confirm the task category: base formatting, green annotations, footnotes, captions, references, validation, or combined workflow.
- Check whether the request includes a target output path and report path.
- If the source file or task description is ambiguous, ask before planning execution.

## Plan Requirements

The plan must state:

- Source file.
- Target output file.
- Reports to be generated.
- Allowed changes.
- Forbidden changes.
- How green annotations will be detected and converted.
- How native footnotes will be preserved.
- How captions and references will be handled.
- Validation commands or checks.
- Manual review points.

Wait for explicit human confirmation before modifying a Word document.

## Execution Requirements

- Create a new output file.
- Never overwrite the source document.
- Use parameterized scripts or reviewed code.
- Keep missing information in reports only.
- Preserve native Word structures.
- Stop and report if parsing is unsafe or document structure is unexpected.

## Validation Checks

- Source file remains unchanged.
- Output package is a valid `.docx` zip.
- Required OOXML parts exist.
- Native footnote count matches body reference count.
- Footnote markers are circled numbers and body markers are superscript.
- Footnote text is 9 pt, Songti-equivalent for Chinese, Times New Roman-equivalent for Latin text and digits, black, regular, single spaced.
- Remaining green text count is expected.
- Ordinary body text protection passes.
- Captions and references were not rewritten.
- Reports list missing data and manual checks.

## Completion Message

Include:

- Output files created.
- Validation summary.
- Privacy summary.
- Manual checking checklist.
- Any limitations, skipped checks, or unresolved items.
