# cafa-thesis-word-format

Reusable Codex Skill for privacy-safe CAFA undergraduate thesis Word formatting.

The required skill entry point is `SKILL.md`. The optional `agents/openai.yaml` file only improves UI presentation and is not required for the skill to trigger or run.

## What It Covers

- Basic Word formatting workflow.
- Green source annotation conversion into native Word page footnotes.
- Circled-number footnote markers.
- Small-five footnote text formatting.
- Existing image caption format checks.
- Existing reference-list format checks.
- Content-protection validation.
- Change and validation reports.

## Privacy Rule

Do not commit, push, or upload real thesis documents, generated reports, private files, personal identifiers, or school source-document copies. This repository is intended to contain only reusable skill instructions, sanitized references, and parameterized script templates.

## Usage

Install or copy this folder into a Codex skills directory, then ask Codex to use `cafa-thesis-word-format` for a CAFA undergraduate thesis Word formatting task.

The agent must plan first, wait for human confirmation, execute only the confirmed scope, and then report validation results.
