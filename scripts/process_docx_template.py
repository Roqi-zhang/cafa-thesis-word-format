#!/usr/bin/env python3
"""Parameterized template for CAFA thesis DOCX processing.

This template intentionally contains no private filenames, thesis titles,
student data, report excerpts, or hard-coded working directories.
Review and adapt the parsing policy before using it on a real document.
"""

from __future__ import annotations

import argparse
import copy
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


def qn(name: str) -> str:
    prefix, local = name.split(":", 1)
    return f"{{{NS[prefix]}}}{local}"


def text_of(element: ET.Element) -> str:
    parts: list[str] = []
    for node in element.iter():
        if node.tag == qn("w:t"):
            parts.append(node.text or "")
        elif node.tag == qn("w:tab"):
            parts.append("\t")
        elif node.tag == qn("w:br"):
            parts.append("\n")
    return "".join(parts)


def ensure_child(parent: ET.Element, tag: str, first: bool = False) -> ET.Element:
    child = parent.find(tag, NS)
    if child is None:
        child = ET.Element(qn(tag))
        if first:
            parent.insert(0, child)
        else:
            parent.append(child)
    return child


def set_child(parent: ET.Element, tag: str, attrs: dict[str, str]) -> ET.Element:
    child = parent.find(tag, NS)
    if child is None:
        child = ET.SubElement(parent, qn(tag))
    for key, value in attrs.items():
        child.set(qn(key), value)
    return child


def run_color(run: ET.Element) -> str:
    rpr = run.find("w:rPr", NS)
    color = rpr.find("w:color", NS) if rpr is not None else None
    return (color.get(qn("w:val")) if color is not None else "").upper()


def is_source_annotation_run(run: ET.Element, source_color: str) -> bool:
    return run_color(run) == source_color.upper()


def make_text_run(text: str, template_run: ET.Element | None = None) -> ET.Element:
    run = ET.Element(qn("w:r"))
    if template_run is not None:
        rpr = template_run.find("w:rPr", NS)
        if rpr is not None:
            cloned = copy.deepcopy(rpr)
            color = cloned.find("w:color", NS)
            if color is not None:
                cloned.remove(color)
            if len(cloned):
                run.append(cloned)
    t = ET.SubElement(run, qn("w:t"))
    if text[:1].isspace() or text[-1:].isspace():
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return run


def make_footnote_reference_run(note_id: int) -> ET.Element:
    run = ET.Element(qn("w:r"))
    rpr = ET.SubElement(run, qn("w:rPr"))
    style = ET.SubElement(rpr, qn("w:rStyle"))
    style.set(qn("w:val"), "FootnoteReference")
    ref = ET.SubElement(run, qn("w:footnoteReference"))
    ref.set(qn("w:id"), str(note_id))
    return run


def strip_annotation_wrappers(text: str) -> str:
    cleaned = text.replace("**", "").strip()
    bracket_pairs = [("(", ")"), ("[", "]"), ("\uff08", "\uff09")]
    for left, right in bracket_pairs:
        if cleaned.startswith(left) and cleaned.endswith(right):
            return cleaned[len(left) : -len(right)].strip()
    return cleaned


def extract_annotation_text(group_text: str) -> str | None:
    cleaned = strip_annotation_wrappers(group_text)
    cleaned = re.sub(r"[ \t]+", " ", cleaned)
    return cleaned or None


@dataclass
class ConversionResult:
    footnotes_created: int = 0
    green_runs_removed: int = 0
    unresolved_annotations: int = 0


def next_footnote_id(footnotes_root: ET.Element) -> int:
    ids: list[int] = []
    for footnote in footnotes_root.findall("w:footnote", NS):
        raw_id = footnote.get(qn("w:id"))
        if raw_id and raw_id.lstrip("-").isdigit():
            ids.append(int(raw_id))
    return max([0, *ids]) + 1


def append_footnote(footnotes_root: ET.Element, note_id: int, note_text: str) -> None:
    footnote = ET.SubElement(footnotes_root, qn("w:footnote"))
    footnote.set(qn("w:id"), str(note_id))
    para = ET.SubElement(footnote, qn("w:p"))
    ppr = ET.SubElement(para, qn("w:pPr"))
    style = ET.SubElement(ppr, qn("w:pStyle"))
    style.set(qn("w:val"), "FootnoteText")

    marker_run = ET.SubElement(para, qn("w:r"))
    marker_ref = ET.SubElement(marker_run, qn("w:footnoteRef"))
    marker_ref.set(qn("w:id"), str(note_id))

    text_run = ET.SubElement(para, qn("w:r"))
    text_node = ET.SubElement(text_run, qn("w:t"))
    text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = " " + note_text


def convert_green_annotations(
    document_root: ET.Element,
    footnotes_root: ET.Element,
    source_color: str,
) -> ConversionResult:
    result = ConversionResult()
    note_id = next_footnote_id(footnotes_root)

    for para in document_root.findall(".//w:p", NS):
        new_children: list[ET.Element] = []
        buffer: list[ET.Element] = []

        def flush_buffer() -> None:
            nonlocal note_id
            if not buffer:
                return
            group_text = "".join(text_of(run) for run in buffer)
            note_text = extract_annotation_text(group_text)
            if note_text:
                append_footnote(footnotes_root, note_id, note_text)
                new_children.append(make_footnote_reference_run(note_id))
                result.footnotes_created += 1
                result.green_runs_removed += len(buffer)
                note_id += 1
            else:
                result.unresolved_annotations += 1
                new_children.extend(buffer)
            buffer.clear()

        for child in list(para):
            if child.tag == qn("w:r") and is_source_annotation_run(child, source_color):
                buffer.append(child)
                continue
            flush_buffer()
            new_children.append(child)
        flush_buffer()

        if result.footnotes_created or result.unresolved_annotations:
            para[:] = new_children

    return result


def read_docx_parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path, "r") as package:
        return {name: package.read(name) for name in package.namelist()}


def write_docx_parts(source: Path, target: Path, replacements: dict[str, bytes]) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source, "r") as zin, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = replacements.get(item.filename)
            if data is None:
                data = zin.read(item.filename)
            zout.writestr(item, data)


def write_report(path: Path | None, lines: list[str]) -> None:
    if path is None:
        print("\n".join(lines))
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Template DOCX processing workflow for CAFA thesis formatting.")
    parser.add_argument("--source-docx", required=True, type=Path)
    parser.add_argument("--target-docx", required=True, type=Path)
    parser.add_argument("--report-md", type=Path)
    parser.add_argument("--source-color", default="00B050", help="OOXML RGB color for source annotations.")
    parser.add_argument("--convert-green-footnotes", action="store_true")
    args = parser.parse_args()

    if args.source_docx.resolve() == args.target_docx.resolve():
        raise ValueError("Refusing to overwrite the source document.")

    parts = read_docx_parts(args.source_docx)
    if "word/document.xml" not in parts:
        raise ValueError("DOCX is missing word/document.xml.")

    document_root = ET.fromstring(parts["word/document.xml"])
    replacements: dict[str, bytes] = {}
    report = ["# Processing Report", "", "- Source was copied before modification."]

    if args.convert_green_footnotes:
        if "word/footnotes.xml" not in parts:
            raise ValueError("Footnote conversion requires word/footnotes.xml. Add the part before using this template.")
        footnotes_root = ET.fromstring(parts["word/footnotes.xml"])
        result = convert_green_annotations(document_root, footnotes_root, args.source_color)
        replacements["word/document.xml"] = ET.tostring(document_root, encoding="utf-8", xml_declaration=True)
        replacements["word/footnotes.xml"] = ET.tostring(footnotes_root, encoding="utf-8", xml_declaration=True)
        report.extend(
            [
                f"- Footnotes created: {result.footnotes_created}",
                f"- Green runs removed: {result.green_runs_removed}",
                f"- Unresolved annotations: {result.unresolved_annotations}",
            ]
        )
    else:
        report.append("- Green annotation conversion was not requested.")

    if replacements:
        write_docx_parts(args.source_docx, args.target_docx, replacements)
    else:
        write_docx_parts(args.source_docx, args.target_docx, {})

    write_report(args.report_md, report)


if __name__ == "__main__":
    main()
