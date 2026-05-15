#!/usr/bin/env python3
"""Parameterized template for native Word footnote style normalization."""

from __future__ import annotations

import argparse
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}

ET.register_namespace("w", NS["w"])


def qn(name: str) -> str:
    prefix, local = name.split(":", 1)
    return f"{{{NS[prefix]}}}{local}"


CIRCLED_MARKS = [
    "\u2460",
    "\u2461",
    "\u2462",
    "\u2463",
    "\u2464",
    "\u2465",
    "\u2466",
    "\u2467",
    "\u2468",
    "\u2469",
    "\u246a",
    "\u246b",
    "\u246c",
    "\u246d",
    "\u246e",
    "\u246f",
    "\u2470",
    "\u2471",
    "\u2472",
    "\u2473",
    "\u3251",
    "\u3252",
    "\u3253",
    "\u3254",
    "\u3255",
    "\u3256",
    "\u3257",
    "\u3258",
    "\u3259",
    "\u325a",
    "\u325b",
    "\u325c",
    "\u325d",
    "\u325e",
    "\u325f",
]


def ensure_child(parent: ET.Element, tag: str, first: bool = False) -> ET.Element:
    child = parent.find(tag, NS)
    if child is None:
        child = ET.Element(qn(tag))
        if first:
            parent.insert(0, child)
        else:
            parent.append(child)
    return child


def set_child(parent: ET.Element, tag: str, attrs: dict[str, str] | None = None) -> ET.Element:
    child = parent.find(tag, NS)
    if child is None:
        child = ET.SubElement(parent, qn(tag))
    for key, value in (attrs or {}).items():
        child.set(qn(key), value)
    return child


def remove_child(parent: ET.Element, tag: str) -> None:
    child = parent.find(tag, NS)
    if child is not None:
        parent.remove(child)


def set_common_run_format(rpr: ET.Element, superscript: bool) -> None:
    fonts = set_child(rpr, "w:rFonts")
    fonts.set(qn("w:eastAsia"), "Songti")
    fonts.set(qn("w:ascii"), "Times New Roman")
    fonts.set(qn("w:hAnsi"), "Times New Roman")
    fonts.set(qn("w:cs"), "Times New Roman")
    set_child(rpr, "w:sz", {"w:val": "18"})
    set_child(rpr, "w:szCs", {"w:val": "18"})
    set_child(rpr, "w:color", {"w:val": "000000"})
    for tag in ("w:b", "w:bCs", "w:i", "w:iCs"):
        remove_child(rpr, tag)
    if superscript:
        set_child(rpr, "w:vertAlign", {"w:val": "superscript"})
    else:
        remove_child(rpr, "w:vertAlign")


def circled_mark(index: int) -> str:
    if index < 1 or index > len(CIRCLED_MARKS):
        raise ValueError(f"No built-in circled marker for footnote index {index}.")
    return CIRCLED_MARKS[index - 1]


def replace_run_with_marker_reference(run: ET.Element, note_id: int, note_index: int) -> None:
    run.clear()
    rpr = ET.SubElement(run, qn("w:rPr"))
    style = ET.SubElement(rpr, qn("w:rStyle"))
    style.set(qn("w:val"), "FootnoteReference")
    set_common_run_format(rpr, superscript=True)
    ref = ET.SubElement(run, qn("w:footnoteReference"))
    ref.set(qn("w:id"), str(note_id))
    ref.set(qn("w:customMarkFollows"), "1")
    text = ET.SubElement(run, qn("w:t"))
    text.text = circled_mark(note_index)


def make_note_marker_run(note_index: int) -> ET.Element:
    run = ET.Element(qn("w:r"))
    rpr = ET.SubElement(run, qn("w:rPr"))
    set_common_run_format(rpr, superscript=False)
    text = ET.SubElement(run, qn("w:t"))
    text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text.text = circled_mark(note_index) + " "
    return run


def format_footnote_paragraph(para: ET.Element) -> None:
    ppr = ensure_child(para, "w:pPr", first=True)
    set_child(ppr, "w:pStyle", {"w:val": "FootnoteText"})
    set_child(ppr, "w:jc", {"w:val": "left"})
    spacing = set_child(ppr, "w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "240")
    spacing.set(qn("w:lineRule"), "auto")
    indent = set_child(ppr, "w:ind")
    indent.set(qn("w:left"), "0")
    indent.set(qn("w:firstLine"), "0")
    indent.set(qn("w:hanging"), "0")


def format_footnote_runs(para: ET.Element) -> None:
    for run in para.findall(".//w:r", NS):
        if run.find("w:footnoteRef", NS) is not None:
            continue
        rpr = ensure_child(run, "w:rPr", first=True)
        set_common_run_format(rpr, superscript=False)


def normalize_footnotes(document_root: ET.Element, footnotes_root: ET.Element, max_notes: int | None) -> list[str]:
    report: list[str] = []
    ordered_note_ids: list[int] = []
    for run in document_root.findall(".//w:r", NS):
        ref = run.find("w:footnoteReference", NS)
        if ref is None:
            continue
        raw_id = ref.get(qn("w:id"))
        if raw_id and raw_id.isdigit():
            ordered_note_ids.append(int(raw_id))

    if max_notes is not None and len(ordered_note_ids) > max_notes:
        raise ValueError("Document contains more footnotes than the configured marker limit.")

    note_order = {note_id: index + 1 for index, note_id in enumerate(ordered_note_ids)}
    for run in document_root.findall(".//w:r", NS):
        ref = run.find("w:footnoteReference", NS)
        if ref is None:
            continue
        note_id = int(ref.get(qn("w:id")))
        replace_run_with_marker_reference(run, note_id, note_order[note_id])

    for footnote in footnotes_root.findall("w:footnote", NS):
        raw_id = footnote.get(qn("w:id"))
        if not raw_id or not raw_id.isdigit():
            continue
        note_id = int(raw_id)
        if note_id not in note_order:
            continue
        para = footnote.find("w:p", NS)
        if para is None:
            continue
        format_footnote_paragraph(para)
        format_footnote_runs(para)
        children = list(para)
        for child in children:
            if child.tag == qn("w:r") and child.find("w:footnoteRef", NS) is not None:
                para.remove(child)
                break
        insert_at = 1 if para.find("w:pPr", NS) is not None else 0
        para.insert(insert_at, make_note_marker_run(note_order[note_id]))

    report.append(f"Body footnote references normalized: {len(ordered_note_ids)}")
    report.append(f"Native footnotes normalized: {len(note_order)}")
    return report


def rewrite_docx(source: Path, target: Path, replacements: dict[str, bytes]) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source, "r") as zin, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = replacements.get(item.filename)
            if data is None:
                data = zin.read(item.filename)
            zout.writestr(item, data)


def main() -> None:
    parser = argparse.ArgumentParser(description="Template footnote style normalization for DOCX files.")
    parser.add_argument("--source-docx", required=True, type=Path)
    parser.add_argument("--target-docx", required=True, type=Path)
    parser.add_argument("--report-md", type=Path)
    parser.add_argument("--max-notes", type=int, default=len(CIRCLED_MARKS))
    args = parser.parse_args()

    if args.source_docx.resolve() == args.target_docx.resolve():
        raise ValueError("Refusing to overwrite the source document.")

    with zipfile.ZipFile(args.source_docx, "r") as package:
        document_root = ET.fromstring(package.read("word/document.xml"))
        footnotes_root = ET.fromstring(package.read("word/footnotes.xml"))

    report = normalize_footnotes(document_root, footnotes_root, args.max_notes)
    rewrite_docx(
        args.source_docx,
        args.target_docx,
        {
            "word/document.xml": ET.tostring(document_root, encoding="utf-8", xml_declaration=True),
            "word/footnotes.xml": ET.tostring(footnotes_root, encoding="utf-8", xml_declaration=True),
        },
    )

    lines = ["# Footnote Style Report", "", *[f"- {line}" for line in report]]
    if args.report_md:
        args.report_md.parent.mkdir(parents=True, exist_ok=True)
        args.report_md.write_text("\n".join(lines) + "\n", encoding="utf-8")
    else:
        print("\n".join(lines))


if __name__ == "__main__":
    main()
