#!/usr/bin/env python3
"""Parameterized validation template for privacy-safe DOCX thesis workflows."""

from __future__ import annotations

import argparse
import hashlib
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


def qn(name: str) -> str:
    prefix, local = name.split(":", 1)
    return f"{{{NS[prefix]}}}{local}"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def text_of(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.iter(qn("w:t")))


def run_color(run: ET.Element) -> str:
    rpr = run.find("w:rPr", NS)
    color = rpr.find("w:color", NS) if rpr is not None else None
    return (color.get(qn("w:val")) if color is not None else "").upper()


def visible_body_text_without_footnote_refs(root: ET.Element) -> str:
    texts: list[str] = []
    for para in root.findall(".//w:p", NS):
        for run in para.findall("w:r", NS):
            if run.find("w:footnoteReference", NS) is not None:
                continue
            texts.append(text_of(run))
        texts.append("\n")
    return "".join(texts)


@dataclass
class ValidationSummary:
    source_exists: bool
    target_exists: bool
    source_hash: str | None
    target_hash: str | None
    package_valid: bool
    document_xml_present: bool
    footnotes_xml_present: bool
    body_footnote_refs: int
    native_footnotes: int
    green_runs: int
    body_text_changed: bool | None


def count_native_footnotes(footnotes_root: ET.Element | None) -> int:
    if footnotes_root is None:
        return 0
    count = 0
    for footnote in footnotes_root.findall("w:footnote", NS):
        raw_id = footnote.get(qn("w:id"))
        if raw_id and raw_id.isdigit() and int(raw_id) > 0:
            count += 1
    return count


def validate(source: Path | None, target: Path, source_color: str) -> ValidationSummary:
    source_hash = sha256(source) if source and source.exists() else None
    target_hash = sha256(target) if target.exists() else None
    package_valid = False
    document_xml_present = False
    footnotes_xml_present = False
    body_footnote_refs = 0
    native_footnotes = 0
    green_runs = 0
    body_text_changed: bool | None = None

    if target.exists():
        with zipfile.ZipFile(target, "r") as package:
            package_valid = package.testzip() is None
            names = set(package.namelist())
            document_xml_present = "word/document.xml" in names
            footnotes_xml_present = "word/footnotes.xml" in names
            document_root = ET.fromstring(package.read("word/document.xml")) if document_xml_present else None
            footnotes_root = ET.fromstring(package.read("word/footnotes.xml")) if footnotes_xml_present else None
            if document_root is not None:
                body_footnote_refs = len(document_root.findall(".//w:footnoteReference", NS))
                green_runs = sum(
                    1
                    for run in document_root.findall(".//w:r", NS)
                    if run_color(run) == source_color.upper()
                )
            native_footnotes = count_native_footnotes(footnotes_root)

    if source and source.exists() and target.exists():
        with zipfile.ZipFile(source, "r") as src_pkg, zipfile.ZipFile(target, "r") as tgt_pkg:
            if "word/document.xml" in src_pkg.namelist() and "word/document.xml" in tgt_pkg.namelist():
                src_root = ET.fromstring(src_pkg.read("word/document.xml"))
                tgt_root = ET.fromstring(tgt_pkg.read("word/document.xml"))
                body_text_changed = visible_body_text_without_footnote_refs(src_root) != visible_body_text_without_footnote_refs(tgt_root)

    return ValidationSummary(
        source_exists=bool(source and source.exists()),
        target_exists=target.exists(),
        source_hash=source_hash,
        target_hash=target_hash,
        package_valid=package_valid,
        document_xml_present=document_xml_present,
        footnotes_xml_present=footnotes_xml_present,
        body_footnote_refs=body_footnote_refs,
        native_footnotes=native_footnotes,
        green_runs=green_runs,
        body_text_changed=body_text_changed,
    )


def report_lines(summary: ValidationSummary) -> list[str]:
    return [
        "# DOCX Validation Report",
        "",
        f"- Source exists: {summary.source_exists}",
        f"- Target exists: {summary.target_exists}",
        f"- Source SHA-256: {summary.source_hash or 'not provided'}",
        f"- Target SHA-256: {summary.target_hash or 'not available'}",
        f"- Package valid: {summary.package_valid}",
        f"- word/document.xml present: {summary.document_xml_present}",
        f"- word/footnotes.xml present: {summary.footnotes_xml_present}",
        f"- Body footnote references: {summary.body_footnote_refs}",
        f"- Native footnotes: {summary.native_footnotes}",
        f"- Remaining green runs: {summary.green_runs}",
        f"- Visible body text changed: {summary.body_text_changed}",
        "",
        "Manual checks required:",
        "- Open the generated document in Word-compatible software.",
        "- Confirm page layout, footnote placement, and visual marker rendering.",
        "- Confirm captions and references were not rewritten.",
    ]


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate a generated thesis DOCX without hard-coded paths.")
    parser.add_argument("--target-docx", required=True, type=Path)
    parser.add_argument("--source-docx", type=Path)
    parser.add_argument("--report-md", type=Path)
    parser.add_argument("--source-color", default="00B050")
    args = parser.parse_args()

    summary = validate(args.source_docx, args.target_docx, args.source_color)
    lines = report_lines(summary)
    if args.report_md:
        args.report_md.parent.mkdir(parents=True, exist_ok=True)
        args.report_md.write_text("\n".join(lines) + "\n", encoding="utf-8")
    else:
        print("\n".join(lines))


if __name__ == "__main__":
    main()
