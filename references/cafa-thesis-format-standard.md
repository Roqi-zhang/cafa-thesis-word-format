# CAFA Thesis Word Formatting Standard

This reference summarizes reusable formatting rules for CAFA undergraduate thesis Word documents. It intentionally contains no thesis body text, report excerpts, private paths, personal identifiers, or school source-document quotations.

## Core Principles

- Modify formatting, not academic content.
- Work from a copy and never overwrite the source document.
- Keep all high-risk operations behind a confirmed plan.
- Use native Word structures whenever possible, especially for footnotes.
- Put missing or uncertain source information into reports only.
- Do not invent bibliography, image-source, author, date, URL, page, or access-date data.
- Generate both a change report and a content-protection validation report.

## Page And Body Formatting

- Use A4 portrait layout unless the confirmed task says otherwise.
- Apply the page margins required by the thesis standard or confirmed template.
- Preserve headers, footers, page numbers, table of contents, fields, heading numbers, and cross-references unless separately confirmed.
- Body Chinese text should use Songti-equivalent East Asian font settings.
- Body Latin text and digits should use Times New Roman-equivalent ASCII and hAnsi font settings.
- Body text is typically small-four size with fixed thesis line spacing, no extra before/after spacing, and first-line indent where appropriate.
- Prefer style-based formatting over broad direct formatting when the document structure supports it.

## Green Source Annotation To Footnote Conversion

- Convert only source annotations already marked as green text in the body.
- Process annotations in document order.
- Remove only the converted green annotation text and its annotation wrappers when confirmed.
- Insert a Word native footnote reference at the annotation position.
- Use only information already present in the green annotation as footnote text.
- Do not complete missing authors, publication data, URLs, page numbers, dates, or access dates.
- If an annotation cannot be parsed safely, report it and leave it for manual review.
- After conversion, validate that green source annotation runs are no longer present, unless unresolved items were intentionally left.

## Footnote Formatting

- Preserve native page footnotes in `word/footnotes.xml`.
- Preserve native `w:footnoteReference` links in `word/document.xml`.
- Body visible markers should be circled numbers and superscript.
- Footnote-area visible markers should be matching circled numbers followed by one space or a tab.
- Footnote text should use the Footnote Text style where available.
- Footnote Chinese text should use Songti-equivalent East Asian font settings.
- Footnote Latin text and digits should use Times New Roman-equivalent font settings.
- Footnote size should be small-five, represented as 9 pt or half-point value `18` in OOXML.
- Footnote text should be black, regular, left aligned, single spaced, with zero before/after spacing and no first-line indent.
- Do not change footnote content text except for the required marker formatting.

## Image Caption Rules

- Normalize formatting of existing image captions only.
- Do not create image sources, authorship, dates, materials, dimensions, or process notes.
- Do not renumber figures or repair numbering automatically.
- Do not move, resize, crop, or change wrapping for images unless separately confirmed.
- Report duplicated numbers, missing numbers, questionable order, and incomplete source information.

## Reference List Rules

- Normalize formatting of existing reference entries only.
- Do not delete, add, reorder, or complete entries automatically.
- Do not synchronize footnotes into the bibliography without a separately confirmed list.
- Use GB/T 7714-2015 as the reference standard for checking, but report missing data instead of fabricating it.
- Suspected template examples should be reported for manual review rather than automatically deleted.

## Validation Requirements

Validate at minimum:

- Source file still exists and was not overwritten.
- Output `.docx` opens as a valid zip package.
- `word/document.xml` exists.
- `word/footnotes.xml` exists when footnotes are expected.
- Body footnote reference count matches native footnote count.
- Remaining green annotation count is expected.
- Body ordinary text is protected.
- Footnote markers and footnote text formatting match the confirmed requirements.
- Image captions and references were changed only within the confirmed formatting scope.
- Manual visual checks are listed when rendering was not performed.
