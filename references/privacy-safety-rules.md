# Privacy And Repository Safety Rules

This skill repository must contain only reusable instructions, sanitized references, parameterized scripts, and repository metadata.

## Never Include

- Real thesis `.docx`, `.doc`, `.pdf`, images, or generated reports.
- Original school submission files or copies of school source documents.
- Thesis body text, abstract text, footnote contents, bibliography contents, captions, or report excerpts from a real student document.
- Names, student IDs, advisors, departments, phone numbers, emails, private URLs, or private filesystem paths.
- Files from user working directories such as source folders, output folders, backup folders, generated reports, or cache folders.

## Script Rules

- Scripts must be templates that accept command-line parameters.
- Scripts must not hard-code real file names, source folders, output folders, backup folders, or one-off report names.
- Scripts must not include example data copied from a thesis or report.
- Scripts may contain generic validation phrases and generic placeholder names only.

## GitHub Rules

- Show the full file list before any commit, push, or upload.
- Confirm no ignored private files are staged.
- If a remote repository is missing or unknown, tell the user to create or provide it.
- Do not create, bind, push to, or upload to an unknown remote.

## Required Ignore Defaults

The repository `.gitignore` must exclude:

- Working source folders.
- Generated output folders.
- Backup folders.
- Word, PDF, image, text-report, log, and cache files.
- Office temporary files.
